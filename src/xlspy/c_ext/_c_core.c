#define PY_SSIZE_T_CLEAN
#include <Python.h>
#include <structmember.h>
#include <datetime.h>
#include <math.h>
#include <pystrtod.h>

/* ===== Types cache (imported from Python) ===== */
static PyObject *decimal_type = NULL;
static PyObject *datetime_type = NULL;
static PyObject *date_type = NULL;
static PyObject *timedelta_type = NULL;

static int _ensure_types(void)
{
    if (decimal_type && datetime_type && date_type && timedelta_type)
        return 1;

    if (!decimal_type)
    {
        PyObject *mod = PyImport_ImportModule("decimal");
        if (!mod) { PyErr_Clear(); return 0; }
        decimal_type = PyObject_GetAttrString(mod, "Decimal");
        Py_DECREF(mod);
        if (!decimal_type) { PyErr_Clear(); return 0; }
    }

    if (!datetime_type || !date_type || !timedelta_type)
    {
        PyObject *mod = PyImport_ImportModule("datetime");
        if (!mod) { PyErr_Clear(); return 0; }
        if (!datetime_type)
        {
            datetime_type = PyObject_GetAttrString(mod, "datetime");
            if (!datetime_type) { Py_DECREF(mod); PyErr_Clear(); return 0; }
        }
        if (!date_type)
        {
            date_type = PyObject_GetAttrString(mod, "date");
            if (!date_type) { Py_DECREF(mod); PyErr_Clear(); return 0; }
        }
        if (!timedelta_type)
        {
            timedelta_type = PyObject_GetAttrString(mod, "timedelta");
            if (!timedelta_type) { Py_DECREF(mod); PyErr_Clear(); return 0; }
        }
        Py_DECREF(mod);
    }
    return 1;
}

/* ===== Helper: read variable-length integer (BIFF format) ===== */
static long _read_varint(const unsigned char *data, Py_ssize_t data_len, Py_ssize_t *pos)
{
    if (*pos >= data_len) return -1;
    long value = 0;
    int shift = 0;
    while (1)
    {
        unsigned char b = data[*pos]; (*pos)++;
        value |= (long)(b & 0x7F) << shift;
        if ((b & 0x80) == 0) break;
        shift += 7;
        if (shift > 28 || *pos >= data_len) return -1;
    }
    return value;
}

/* Assembled unsigned: with a 32-bit long (Windows), `(long)byte << 24`
   overflows for bytes >= 0x80, which is undefined behaviour. */
static unsigned long _le_uint32(const unsigned char *p)
{
    return (unsigned long)p[0] | ((unsigned long)p[1] << 8) |
           ((unsigned long)p[2] << 16) | ((unsigned long)p[3] << 24);
}

static long _read_uint32(const unsigned char *data, Py_ssize_t data_len, Py_ssize_t pos)
{
    if (pos + 4 > data_len) return -1;
    return (long)_le_uint32(data + pos);
}

static int _read_uint32_safe(const unsigned char *data, Py_ssize_t data_len, Py_ssize_t pos, long *out)
{
    if (pos + 4 > data_len) return 0;
    *out = (long)_le_uint32(data + pos);
    return 1;
}

#define XLSB_MAX_ROW_INDEX 1048575UL

/* Row index of a BrtRowHdr body, like the reference: a short body and row
   0xFFFFFFFF both read as -1 (on every platform, whatever the width of long).
   A row past Excel's last would be padded with billions of empty rows, so it
   raises ValueError and returns 0; otherwise returns 1. */
static int _read_row_index(const unsigned char *data, Py_ssize_t data_len, Py_ssize_t pos, long *out)
{
    *out = -1;
    if (pos + 4 > data_len) return 1;
    unsigned long raw = _le_uint32(data + pos);
    if (raw == 0xFFFFFFFFUL) return 1;
    if (raw > XLSB_MAX_ROW_INDEX)
    {
        PyErr_Format(PyExc_ValueError,
                     "XLSB row record has row index %lu, past the last row (%lu)",
                     raw, XLSB_MAX_ROW_INDEX);
        return 0;
    }
    *out = (long)raw;
    return 1;
}

static double _read_double(const unsigned char *data, Py_ssize_t data_len, Py_ssize_t pos)
{
    if (pos + 8 > data_len) return 0.0;
    unsigned long long bits = 0;
    for (int i = 0; i < 8; i++)
        bits |= (unsigned long long)data[pos + i] << (i * 8);
    double result;
    memcpy(&result, &bits, 8);
    return result;
}

static PyObject *_number_from_double(double value)
{
    /* Avoid an epsilon test here: it turns tiny, non-zero doubles into zero. */
    if (isfinite(value) && trunc(value) == value)
        return PyLong_FromDouble(value);
    return PyFloat_FromDouble(value);
}

static double _read_rk(const unsigned char *data, Py_ssize_t data_len, Py_ssize_t pos)
{
    if (pos + 4 > data_len) return 0.0;
    unsigned int raw = (unsigned int)data[pos] | ((unsigned int)data[pos+1] << 8) |
                       ((unsigned int)data[pos+2] << 16) | ((unsigned int)data[pos+3] << 24);
    double result;
    if (raw & 2)
    {
        int int_val = (int)(raw >> 2);
        if (int_val & 0x20000000)
            int_val |= (int)0xC0000000;
        result = (double)int_val;
    }
    else
    {
        unsigned long long mant = (unsigned long long)(raw & 0xFFFFFFFC) << 32;
        double dval;
        memcpy(&dval, &mant, 8);
        result = dval;
    }
    if (raw & 1)
        result /= 100.0;
    return result;
}

static PyObject *_utf16le_to_pystr(const unsigned char *data, Py_ssize_t data_len, Py_ssize_t pos, long char_len)
{
    if (char_len <= 0) return PyUnicode_FromStringAndSize("", 0);
    Py_ssize_t byte_len = char_len * 2;
    if (pos + byte_len > data_len) return PyUnicode_FromStringAndSize("", 0);
    int byteorder = -1; /* UTF-16LE */
    return PyUnicode_DecodeUTF16(
        (const char *)(data + pos), (Py_ssize_t)byte_len, "surrogatepass", &byteorder
    );
}

/* Date-conversion failure: keep a real error (OverflowError for
   out-of-range serials, ValueError for NaN, ...) pending so the caller
   aborts iteration exactly like the Python reference, which lets
   serial_to_datetime() raise.  A failure with no error set (paranoia)
   falls back to the numeric value.  Returns 1 when fatal. */
static int _date_failed(int *fatal, int *is_date)
{
    if (PyErr_Occurred()) { *fatal = 1; return 1; }
    PyErr_Clear();
    *is_date = 0;
    return 0;
}

/* ===== Read a single cell's value from a worksheet cell record ===== */
static PyObject *_parse_cell_value(const unsigned char *data, Py_ssize_t data_len,
                                    Py_ssize_t body_start, long rec_id, long rec_len,
                                    PyObject *shared_strings, Py_ssize_t ss_len,
                                    PyObject *styles_list, Py_ssize_t styles_len,
                                    PyObject *date_num_fmts, long xf_idx, int *has_cell,
                                    int date_1904, int *fatal, int *skip)
{
    PyObject *cell_val = NULL;
    int is_date = 0;
    *fatal = 0;
    *skip = 0;

    if (rec_id == 0x01 || rec_id == 0x03 || rec_id == 0x0B)
    {
        cell_val = Py_None;
    }
    else if (rec_id == 0x02)
    {
        /* The reference ignores number records shorter than 12 bytes. */
        if (rec_len < 12) { *skip = 1; return NULL; }
        double dval = _read_rk(data, data_len, body_start + 8);
        if (xf_idx > 0 && xf_idx < styles_len)
        {
            PyObject *style = PyList_GetItem(styles_list, xf_idx);
            if (style && PyDict_Check(style))
            {
                PyObject *nf = PyDict_GetItemString(style, "numFmtId");
                if (nf)
                    is_date = (PySet_Contains(date_num_fmts, nf) == 1);
            }
        }
        if (is_date)
        {
            _ensure_types();
            /* Excel's fictitious 1900-02-29: serials 1..59 read one day later. */
            double serial = dval;
            if (!date_1904 && serial >= 1.0 && serial < 60.0)
                serial += 1.0;
            PyObject *epoch = date_1904
                ? PyDateTime_FromDateAndTime(1904, 1, 1, 0, 0, 0, 0)
                : PyDateTime_FromDateAndTime(1899, 12, 30, 0, 0, 0, 0);
            if (!epoch) { if (_date_failed(fatal, &is_date)) return NULL; }
            else
            {
                PyObject *td_kw = Py_BuildValue("{s:d}", "days", serial);
                if (!td_kw) { Py_DECREF(epoch); if (_date_failed(fatal, &is_date)) return NULL; }
                else
                {
                    PyObject *noargs = PyTuple_New(0);
                    PyObject *td = PyObject_Call(timedelta_type, noargs, td_kw);
                    Py_DECREF(noargs);
                    Py_DECREF(td_kw);
                    if (!td) { Py_DECREF(epoch); if (_date_failed(fatal, &is_date)) return NULL; }
                    else
                    {
                        cell_val = PyNumber_Add(epoch, td);
                        Py_DECREF(td);
                        Py_DECREF(epoch);
                        if (!cell_val) { if (_date_failed(fatal, &is_date)) return NULL; }
                    }
                }
            }
        }
        if (!is_date)
        {
            cell_val = _number_from_double(dval);
        }
    }
    else if (rec_id == 0x04 || rec_id == 0x0A)
    {
        /* The reference ignores bool records shorter than 9 bytes, and
           only byte value 1 reads as True. */
        if (rec_len < 9) { *skip = 1; return NULL; }
        cell_val = (data[body_start + 8] == 1) ? Py_True : Py_False;
    }
    else if (rec_id == 0x05 || rec_id == 0x09)
    {
        /* The reference ignores float records shorter than 16 bytes. */
        if (rec_len < 16) { *skip = 1; return NULL; }
        double dval = _read_double(data, data_len, body_start + 8);
        if (xf_idx > 0 && xf_idx < styles_len)
        {
            PyObject *style = PyList_GetItem(styles_list, xf_idx);
            if (style && PyDict_Check(style))
            {
                PyObject *nf = PyDict_GetItemString(style, "numFmtId");
                if (nf)
                    is_date = (PySet_Contains(date_num_fmts, nf) == 1);
            }
        }
        if (is_date)
        {
            _ensure_types();
            /* Excel's fictitious 1900-02-29: serials 1..59 read one day later. */
            double serial = dval;
            if (!date_1904 && serial >= 1.0 && serial < 60.0)
                serial += 1.0;
            PyObject *epoch = date_1904
                ? PyDateTime_FromDateAndTime(1904, 1, 1, 0, 0, 0, 0)
                : PyDateTime_FromDateAndTime(1899, 12, 30, 0, 0, 0, 0);
            if (!epoch) { if (_date_failed(fatal, &is_date)) return NULL; }
            else
            {
                PyObject *td_kw = Py_BuildValue("{s:d}", "days", serial);
                if (!td_kw) { Py_DECREF(epoch); if (_date_failed(fatal, &is_date)) return NULL; }
                else
                {
                    PyObject *noargs = PyTuple_New(0);
                    PyObject *td = PyObject_Call(timedelta_type, noargs, td_kw);
                    Py_DECREF(noargs);
                    Py_DECREF(td_kw);
                    if (!td) { Py_DECREF(epoch); if (_date_failed(fatal, &is_date)) return NULL; }
                    else
                    {
                        cell_val = PyNumber_Add(epoch, td);
                        Py_DECREF(td);
                        Py_DECREF(epoch);
                        if (!cell_val) { if (_date_failed(fatal, &is_date)) return NULL; }
                    }
                }
            }
        }
        if (!is_date)
        {
            cell_val = _number_from_double(dval);
        }
    }
    else if (rec_id == 0x06 || rec_id == 0x08)
    {
        /* Mirror the reference: records shorter than 12 bytes, or whose
           claimed length overruns the record, are ignored (not stored as
           empty strings, and never clobbering an earlier duplicate). */
        if (rec_len < 12) { *skip = 1; return NULL; }
        unsigned long str_len =
            (unsigned long)_read_uint32(data, data_len, body_start + 8);
        if ((Py_ssize_t)str_len * 2 + 12 > (Py_ssize_t)rec_len)
        { *skip = 1; return NULL; }
        cell_val = _utf16le_to_pystr(data, data_len, body_start + 12,
                                     (long)str_len);
        if (!cell_val)
        {
            if (PyErr_Occurred()) { *fatal = 1; return NULL; }
            cell_val = PyUnicode_FromStringAndSize("", 0);
        }
    }
    else if (rec_id == 0x07)
    {
        /* Records shorter than 12 bytes are ignored by the reference. */
        if (rec_len < 12) { *skip = 1; return NULL; }
        long idx = _read_uint32(data, data_len, body_start + 8);
        if (idx >= 0 && idx < ss_len)
        {
            cell_val = PyList_GetItem(shared_strings, (Py_ssize_t)idx);
            if (cell_val) Py_INCREF(cell_val);
        }
        if (!cell_val) cell_val = Py_None;
    }

    if (!cell_val)
    {
        /* A real error (e.g. string decode failure) propagates like the
           reference; a plain miss (OOB shared-string index) becomes None. */
        if (PyErr_Occurred()) { *fatal = 1; return NULL; }
        cell_val = Py_None;
    }
    if (cell_val == Py_None || cell_val == Py_True || cell_val == Py_False)
        *has_cell = 0;
    else
        *has_cell = 1;
    return cell_val;
}

/* ===== Clean up cur_cells when a row cannot be built ===== */
static void _free_cur_cells(PyObject **cur_cells, Py_ssize_t cur_cells_len)
{
    if (!cur_cells) return;
    for (Py_ssize_t c = 0; c < cur_cells_len; c++)
        Py_XDECREF(cur_cells[c]);
    free(cur_cells);
}

/* ===== Existing writer functions ===== */

static double _get_excel_serial(PyObject *cell, int cell_kind)
{
    PyObject *ymd = PyObject_GetAttrString(cell, "year");
    if (!ymd) { PyErr_Clear(); return -1.0; }
    long year = PyLong_AsLong(ymd);
    Py_DECREF(ymd);
    if (year == -1 && PyErr_Occurred()) { PyErr_Clear(); return -1.0; }
    if (year < 1900 || year > 9999) return -1.0;

    PyObject *ord = PyObject_CallMethod(cell, "toordinal", NULL);
    if (!ord) { PyErr_Clear(); return -1.0; }
    long ord_val = PyLong_AsLong(ord);
    Py_DECREF(ord);
    if (ord_val == -1 && PyErr_Occurred()) { PyErr_Clear(); return -1.0; }

    long epoch_ord = 693594;
    /* Excel's fictitious 1900-02-29: dates before 1900-03-01 use the
       1899-12-31 epoch, so 1900-01-01 is stored as serial 1. */
    if (ord_val < 693655)
        epoch_ord = 693595;
    double serial = (double)(ord_val - epoch_ord);

    if (cell_kind == 1)
    {
        struct { const char *name; double val; } tp[4] = {
            {"hour", 0.0}, {"minute", 0.0}, {"second", 0.0}, {"microsecond", 0.0}
        };
        int ok = 1;
        for (int i = 0; i < 4; i++)
        {
            PyObject *a = PyObject_GetAttrString(cell, tp[i].name);
            if (a)
            {
                tp[i].val = (double)PyLong_AsLong(a);
                if (tp[i].val == -1 && PyErr_Occurred()) { PyErr_Clear(); ok = 0; }
                Py_DECREF(a);
            }
            else { PyErr_Clear(); ok = 0; }
            if (!ok) break;
        }
        if (ok)
        {
            double ts = tp[0].val * 3600.0 + tp[1].val * 60.0 + tp[2].val + tp[3].val / 1e6;
            serial += ts / 86400.0;
        }
    }
    return serial;
}

/* xlspy.Style, registered by set_style_type: the second item of a
   (value, Style) cell, which the writers treat like (value, format). */
static PyObject *g_style_type = NULL;

/* Kind of the second item of a 2-tuple cell: 1 format string, 2 Style,
   0 anything else (the tuple is then an ordinary value). */
static int _cell_overlay_kind(PyObject *item)
{
    if (PyUnicode_Check(item)) return 1;
    if (g_style_type && PyObject_TypeCheck(item, (PyTypeObject *)g_style_type)) return 2;
    return 0;
}

static PyObject *C_set_style_type(PyObject *self, PyObject *type)
{
    if (!PyType_Check(type))
    {
        PyErr_SetString(PyExc_TypeError, "set_style_type expects a class");
        return NULL;
    }
    Py_INCREF(type);
    Py_XSETREF(g_style_type, type);
    Py_RETURN_NONE;
}

static PyObject *C_calc_column_widths(PyObject *self, PyObject *args)
{
    PyObject *rows;
    int max_cols;

    if (!PyArg_ParseTuple(args, "Oi", &rows, &max_cols))
        return NULL;
    if (!PyList_Check(rows))
    {
        PyErr_SetString(PyExc_TypeError, "rows must be a list");
        return NULL;
    }

    _ensure_types();

    Py_ssize_t nrows = PyList_Size(rows);
    PyObject *result = PyList_New(max_cols);
    if (!result) return NULL;

    for (int i = 0; i < max_cols; i++)
    {
        double max_width = 0.0;
        for (Py_ssize_t r = 0; r < nrows; r++)
        {
            PyObject *row = PyList_GetItem(rows, r);
            if (!row || !PyList_Check(row)) continue;
            Py_ssize_t row_len = PyList_Size(row);
            if (i >= row_len) continue;

            PyObject *cell = PyList_GetItem(row, i);
            if (!cell || cell == Py_None) continue;

            /* Mirror the Python fallback: unwrap (value, format) tuples,
               then measure Decimal as float(cell). */
            PyObject *owned = NULL;
            if (PyTuple_Check(cell))
            {
                Py_ssize_t tlen = PyObject_Length(cell);
                if (tlen < 0) { PyErr_Clear(); continue; }
                if (tlen == 2)
                {
                    PyObject *fmt = PySequence_GetItem(cell, 1);
                    if (!fmt) { PyErr_Clear(); continue; }
                    int is_fmt = _cell_overlay_kind(fmt);
                    Py_DECREF(fmt);
                    if (is_fmt)
                    {
                        owned = PySequence_GetItem(cell, 0);
                        if (!owned) { PyErr_Clear(); continue; }
                        cell = owned;
                        if (cell == Py_None) { Py_DECREF(owned); continue; }
                    }
                }
            }
            if (decimal_type && PyObject_IsInstance(cell, decimal_type) == 1)
            {
                PyObject *f = PyNumber_Float(cell);
                Py_XDECREF(owned);
                owned = NULL;
                if (!f) { PyErr_Clear(); continue; }
                owned = f;
                cell = f;
            }

            double temp_width;
            int is_dt = datetime_type ? PyObject_IsInstance(cell, datetime_type) : 0;
            if (is_dt == 1)
                temp_width = 18.0;
            else
            {
                PyObject *str_cell = PyObject_Str(cell);
                if (!str_cell) { PyErr_Clear(); Py_XDECREF(owned); continue; }
                /* Count code points up to the first newline, as len() does in
                   the Python fallback (UTF-8 bytes would widen non-ASCII text). */
                Py_ssize_t slen = PyUnicode_GET_LENGTH(str_cell);
                int skind = PyUnicode_KIND(str_cell);
                const void *sdata = PyUnicode_DATA(str_cell);
                Py_ssize_t lenn = 0;
                while (lenn < slen && PyUnicode_READ(skind, sdata, lenn) != '\n') lenn++;
                temp_width = 1.25 * (double)lenn + 4.0;
                Py_DECREF(str_cell);
            }
            Py_XDECREF(owned);
            if (temp_width > max_width) max_width = temp_width;
        }
        if (max_width == 0.0) max_width = 4.0;
        if (max_width > 255.0) max_width = 255.0;
        PyObject *wv = PyFloat_FromDouble(max_width);
        if (!wv) { Py_DECREF(result); return NULL; }
        PyList_SET_ITEM(result, i, wv);
    }
    return result;
}

/* ===== Fast pre-scan: does a row contain any (value, format) tuple? =====
   Mirrors `_is_formatted_cell` (formats.py): a non-None cell that is a
   tuple of length 2 whose second item is a str.  Returns 1/0, or None when
   the row cannot be scanned here (caller falls back to the Python check,
   which raises the same errors for non-iterable rows). */
static PyObject *C_xlsb_row_has_formatted(PyObject *self, PyObject *args)
{
    PyObject *row;
    if (!PyArg_ParseTuple(args, "O", &row))
        return NULL;

    PyObject *seq = PySequence_Fast(row, "row must be a sequence");
    if (!seq) { PyErr_Clear(); Py_RETURN_NONE; }

    Py_ssize_t n = PySequence_Fast_GET_SIZE(seq);
    PyObject **items = PySequence_Fast_ITEMS(seq);
    int found = 0;
    for (Py_ssize_t i = 0; i < n; i++)
    {
        PyObject *cell = items[i];
        if (cell == Py_None)
            continue;
        if (PyTuple_Check(cell))
        {
            /* PyObject_Length honors __len__ overrides, exactly like len(). */
            Py_ssize_t tlen = PyObject_Length(cell);
            if (tlen < 0) { PyErr_Clear(); Py_DECREF(seq); Py_RETURN_NONE; }
            if (tlen == 2)
            {
                /* PySequence_GetItem honors __getitem__ overrides, like cell[1]. */
                PyObject *fmt = PySequence_GetItem(cell, 1);
                if (!fmt) { PyErr_Clear(); Py_DECREF(seq); Py_RETURN_NONE; }
                int is_fmt = _cell_overlay_kind(fmt);
                Py_DECREF(fmt);
                if (is_fmt) { found = 1; break; }
            }
        }
    }
    Py_DECREF(seq);
    if (found) Py_RETURN_TRUE;
    Py_RETURN_FALSE;
}

/* ===== XLSB shared-strings table encoder ===== */

static unsigned char *_write_varint_bytes(unsigned char *p, unsigned long value)
{
    while (1)
    {
        unsigned char b = (unsigned char)(value & 0x7F);
        value >>= 7;
        if (value) { *p++ = (unsigned char)(b | 0x80); }
        else { *p++ = b; break; }
    }
    return p;
}

static size_t _varint_len(unsigned long value)
{
    size_t n = 1;
    while (value >>= 7) n++;
    return n;
}

/* Validate surrogate pairing over the whole string, replicating the
   `str.encode('utf-16-le')` error behavior (which runs before any
   truncation).  CPython's UTF-16 encoder rejects every surrogate code
   point -- even a high+low pair written as two escapes -- so any
   0xD800..0xDFFF unit is an error at its own index.  Returns 0 when valid,
   -1 with *err_pos set to the offending index otherwise. */
static int _sst_check_surrogates(int kind, const void *data, Py_ssize_t slen,
                                 Py_ssize_t *err_pos)
{
    for (Py_ssize_t k = 0; k < slen; k++)
    {
        Py_UCS4 ch = PyUnicode_READ(kind, data, k);
        if (ch >= 0xD800 && ch <= 0xDFFF) { *err_pos = k; return -1; }
    }
    return 0;
}

static PyObject *_sst_surrogate_error(PyObject *s, Py_ssize_t pos)
{
    PyObject *exc = PyObject_CallFunction(
        PyExc_UnicodeEncodeError, "sOnns",
        "utf-16-le", s, pos, pos + 1, "surrogates not allowed");
    if (!exc) return NULL;
    PyErr_SetObject(PyExc_UnicodeEncodeError, exc);
    Py_DECREF(exc);
    return NULL;
}

/* Excel's per-cell text limit in UTF-16 code units, for XLSX and XLSB alike
   (see xlspy/_text.py, which owns the error message). */
#define _CELL_TEXT_LIMIT 32767

/* UTF-16 code units in one shared string (a non-BMP character is two).  The
   caller has already rejected lone surrogates and checked the cell limit. */
static Py_ssize_t _sst_units(int kind, const void *data, Py_ssize_t slen)
{
    Py_ssize_t ulen = 0;
    for (Py_ssize_t k = 0; k < slen; k++)
        ulen += (PyUnicode_READ(kind, data, k) >= 0x10000) ? 2 : 1;
    return ulen;
}

static unsigned char *_write_utf16_unit(unsigned char *p, unsigned int unit)
{
    *p++ = (unsigned char)(unit & 0xFF);
    *p++ = (unsigned char)((unit >> 8) & 0xFF);
    return p;
}

/* Encode every UTF-16LE unit of the string; _sst_units sized the record. */
static unsigned char *_write_sst_string(unsigned char *p, int kind,
                                        const void *data, Py_ssize_t slen)
{
    for (Py_ssize_t k = 0; k < slen; k++)
    {
        Py_UCS4 ch = PyUnicode_READ(kind, data, k);
        if (ch < 0x10000)
        {
            p = _write_utf16_unit(p, (unsigned int)ch);
        }
        else
        {
            ch -= 0x10000;
            p = _write_utf16_unit(p, 0xD800u | (unsigned int)(ch >> 10));
            p = _write_utf16_unit(p, 0xDC00u | (unsigned int)(ch & 0x3FF));
        }
    }
    return p;
}

/* Encode the complete sharedStrings.bin part (book + SST records).
   Two passes: first measure the exact size (single allocation, no
   realloc loop), then fill.  Returns bytes, or None on error. */
static PyObject *C_encode_xlsb_sst(PyObject *self, PyObject *args)
{
    PyObject *ss_list;
    unsigned int sst_all, sst_unique;

    if (!PyArg_ParseTuple(args, "OII", &ss_list, &sst_all, &sst_unique))
        return NULL;
    if (!PyList_Check(ss_list))
    {
        PyErr_SetString(PyExc_TypeError, "ss_list must be a list");
        return NULL;
    }

    Py_ssize_t n = PyList_GET_SIZE(ss_list);

    /* Pass 1: measure. */
    size_t total = 3 + 8;  /* BrtBeginSst header + counts */
    for (Py_ssize_t i = 0; i < n; i++)
    {
        PyObject *s = PyList_GET_ITEM(ss_list, i);
        if (!PyUnicode_Check(s))
        {
            PyErr_SetString(PyExc_TypeError, "SST entries must be str");
            return NULL;
        }
        Py_ssize_t slen = PyUnicode_GET_LENGTH(s);
        Py_ssize_t ulen = 0;
        if (slen > 0)
        {
            int kind = PyUnicode_KIND(s);
            const void *data = PyUnicode_DATA(s);
            Py_ssize_t err_pos = 0;
            if (_sst_check_surrogates(kind, data, slen, &err_pos) < 0)
                return _sst_surrogate_error(s, err_pos);
            ulen = _sst_units(kind, data, slen);
            if (ulen > _CELL_TEXT_LIMIT)
            {
                PyErr_Format(PyExc_ValueError,
                    "cell text exceeds Excel's %d-character limit (%zd UTF-16 code units)",
                    _CELL_TEXT_LIMIT, ulen);
                return NULL;
            }
        }
        unsigned long rec_len = 5UL + 2UL * (unsigned long)ulen;
        total += 1 + _varint_len(rec_len) + 1 + 4 + 2 * (size_t)ulen;
    }
    total += 3;  /* BrtEndSst */

    if (total > (size_t)PY_SSIZE_T_MAX)
    {
        PyErr_NoMemory();
        return NULL;
    }
    PyObject *result = PyBytes_FromStringAndSize(NULL, (Py_ssize_t)total);
    if (!result) return NULL;
    unsigned char *p = (unsigned char *)PyBytes_AS_STRING(result);

    /* Pass 2: fill. */
    *p++ = 0x9F; *p++ = 0x01; *p++ = 0x08;
    memcpy(p, &sst_all, 4); p += 4;
    memcpy(p, &sst_unique, 4); p += 4;
    for (Py_ssize_t i = 0; i < n; i++)
    {
        PyObject *s = PyList_GET_ITEM(ss_list, i);
        Py_ssize_t slen = PyUnicode_GET_LENGTH(s);
        Py_ssize_t ulen = 0;
        int kind = 0;
        const void *data = NULL;
        if (slen > 0)
        {
            kind = PyUnicode_KIND(s);
            data = PyUnicode_DATA(s);
            ulen = _sst_units(kind, data, slen);
        }
        unsigned long rec_len = 5UL + 2UL * (unsigned long)ulen;
        *p++ = 0x13;
        p = _write_varint_bytes(p, rec_len);
        *p++ = 0x00;
        {
            unsigned int ul = (unsigned int)ulen;
            memcpy(p, &ul, 4); p += 4;
        }
        if (ulen > 0)
            p = _write_sst_string(p, kind, data, slen);
    }
    *p++ = 0xA0; *p++ = 0x01; *p++ = 0x00;

    return result;
}

/* ===== XLSB shared-strings table reader =====
   Parses a sharedStrings.bin part into a list of strings.  Observable
   behavior mirrors `_fill_bin_shared_strings` + `BiffReader` exactly,
   quirks included:
   - every BrtSSTItem (0x13) appends its string in record order, EXCEPT
     the value of the very first parsed record, which the Python loop
     consumes for the unique-count lookup and discards (so a part that
     does not start with BrtBeginSst loses its first item);
   - BrtBeginSst (159) and unknown records are skipped (the unique-count
     preallocation is unobservable: missing slots are truncated away and
     surplus items are appended);
   - clean EOF (including a truncated varint or a record whose payload
     runs past the end) stops silently, like the short reads of the
     Python path;
   - a short BrtSSTItem/BrtBeginSst record raises IndexError, mirroring
     the out-of-range buffer access of the Python path;
   - when the claimed string length overruns the record, the available
     bytes are decoded (raising UnicodeDecodeError on odd length),
     mirroring the slice-then-decode of `_get_string`.
   All length arithmetic uses Py_ssize_t so hostile record lengths cannot
   overflow into out-of-bounds reads. */
static PyObject *C_read_xlsb_shared_strings(PyObject *self, PyObject *args)
{
    Py_buffer view;
    if (!PyArg_ParseTuple(args, "y*", &view))
        return NULL;
    const unsigned char *data = (const unsigned char *)view.buf;
    Py_ssize_t data_len = view.len;

    PyObject *result = PyList_New(0);
    if (!result) { PyBuffer_Release(&view); return NULL; }

    Py_ssize_t pos = 0;
    int first_record = 1;
    while (pos < data_len)
    {
        long rec_id = _read_varint(data, data_len, &pos);
        if (rec_id < 0) break;  /* truncated varint: silent stop */
        long rec_len = _read_varint(data, data_len, &pos);
        if (rec_len < 0) break;  /* truncated varint: silent stop */
        if ((Py_ssize_t)rec_len > data_len - pos)
            break;  /* payload runs past the end: silent stop */
        Py_ssize_t body = pos;
        pos += (Py_ssize_t)rec_len;
        /* The Python loop consumes the first parsed record for the
           unique-count lookup and discards its value -- even when that
           record is unknown or inconsistent. */
        int is_first = first_record;
        first_record = 0;

        if (rec_id == 159)
        {
            /* BrtBeginSst: the Python path reads the count dword. */
            if ((Py_ssize_t)rec_len < 5)
            {
                Py_DECREF(result);
                PyBuffer_Release(&view);
                PyErr_SetString(PyExc_IndexError, "truncated BrtBeginSst record");
                return NULL;
            }
            continue;
        }
        if (rec_id != 0x13)
            continue;
        /* BrtSSTItem: dword length at +1, UTF-16LE chars at +5. */
        if ((Py_ssize_t)rec_len < 5)
        {
            Py_DECREF(result);
            PyBuffer_Release(&view);
            PyErr_SetString(PyExc_IndexError, "truncated BrtSSTItem record");
            return NULL;
        }
        long slen_raw = _read_uint32(data, data_len, body + 1);
        /* Unsigned: hostile lengths (>= 2^31) must clamp to the available
           bytes like Python's slice, not go negative into PyUnicode_New. */
        unsigned long slen = (unsigned long)slen_raw;
        /* Mirror _get_string's slice-then-decode: only the bytes present
           in the record are decoded (odd length raises UnicodeDecodeError,
           like bytes.decode). */
        Py_ssize_t avail = (Py_ssize_t)rec_len - 5;
        if (avail < 0) avail = 0;
        Py_ssize_t want = (Py_ssize_t)slen * 2;
        Py_ssize_t have = (want <= avail) ? want : avail;
        int byteorder = -1;  /* UTF-16LE */
        PyObject *s = PyUnicode_DecodeUTF16(
            (const char *)(data + body + 5), have,
            "surrogatepass", &byteorder);
        if (!s)
        {
            Py_DECREF(result);
            PyBuffer_Release(&view);
            return NULL;
        }
        if (is_first)
        {
            /* First record's value is discarded by the Python loop. */
            Py_DECREF(s);
            continue;
        }
        if (PyList_Append(result, s) < 0)
        {
            Py_DECREF(s);
            Py_DECREF(result);
            PyBuffer_Release(&view);
            return NULL;
        }
        Py_DECREF(s);
    }

    PyBuffer_Release(&view);
    return result;
}

/* 1 when s fits in a cell.  Each code point is one or two units, so only
   text longer than half the limit needs counting, and only 4-byte-kind
   strings can hold the two-unit characters. */
static int _cell_text_fits(PyObject *s)
{
    Py_ssize_t slen = PyUnicode_GET_LENGTH(s);
    if (slen <= _CELL_TEXT_LIMIT / 2) return 1;
    if (slen > _CELL_TEXT_LIMIT) return 0;
    if (PyUnicode_KIND(s) != PyUnicode_4BYTE_KIND) return 1;
    const Py_UCS4 *data = PyUnicode_4BYTE_DATA(s);
    Py_ssize_t units = slen;
    for (Py_ssize_t k = 0; k < slen; k++)
        if (data[k] >= 0x10000) units++;
    return units <= _CELL_TEXT_LIMIT;
}

/* Shared strings in encoder.  Returns -2 (row fallback) for text over the
   cell limit, so the Python encoder raises the error. */
static int _write_str_cell(unsigned char **buf, size_t *buf_size, size_t *buf_cap,
                           PyObject *cell_str, PyObject *ss_dict, PyObject *ss_list,
                           unsigned int col_idx, unsigned int row_idx,
                           unsigned int style_override,
                           unsigned int *sst_uniq, unsigned int *sst_all)
{
    if (!_cell_text_fits(cell_str)) return -2;
    PyObject *idx_obj = PyDict_GetItemWithError(ss_dict, cell_str);
    if (!idx_obj && PyErr_Occurred()) return -1;  /* e.g. a str subclass whose hash raises */
    unsigned int string_index;
    unsigned int style_ref = style_override ? style_override : ((row_idx == 0) ? 3 : 0);

    if (idx_obj)
        string_index = (unsigned int)PyLong_AsLong(idx_obj);
    else
    {
        string_index = *sst_uniq;
        PyObject *si = PyLong_FromLong((long)string_index);
        if (!si) return -1;
        /* List before dict, as in _xlsx_sst_index: see the note there. */
        if (PyList_Append(ss_list, cell_str) < 0) { Py_DECREF(si); return -1; }
        if (PyDict_SetItem(ss_dict, cell_str, si) < 0) { Py_DECREF(si); return -1; }
        Py_DECREF(si);
        (*sst_uniq)++;
    }
    (*sst_all)++;

    size_t need = *buf_size + 14;
    if (need > *buf_cap)
    {
        *buf_cap = need + 4096;
        unsigned char *nb = (unsigned char *)PyMem_Realloc(*buf, *buf_cap);
        if (!nb) return -1;
        *buf = nb;
    }

    (*buf)[(*buf_size)++] = 0x07;
    (*buf)[(*buf_size)++] = 0x0C;
    memcpy(*buf + *buf_size, &col_idx, 4); (*buf_size) += 4;
    memcpy(*buf + *buf_size, &style_ref, 4); (*buf_size) += 4;
    memcpy(*buf + *buf_size, &string_index, 4); (*buf_size) += 4;
    return 0;
}

/* Inline text (BrtCellSt, record 0x06): the string is stored in the cell, so
   the shared-strings table does not grow.  The count is in UTF-16 code units
   (a non-BMP character is two), exactly as Excel reads it.  Returns -2 when
   the string needs the Python path, which owns the error: more than 32767
   units, or a lone surrogate. */
static int _write_inline_str_cell(unsigned char **buf, size_t *buf_size, size_t *buf_cap,
                                  PyObject *s, unsigned int col_idx, unsigned int style_ref)
{
    int kind = PyUnicode_KIND(s);
    const void *data = PyUnicode_DATA(s);
    Py_ssize_t slen = PyUnicode_GET_LENGTH(s);
    Py_ssize_t units = 0;

    for (Py_ssize_t k = 0; k < slen; k++)
    {
        Py_UCS4 ch = PyUnicode_READ(kind, data, k);
        if (ch >= 0xD800 && ch <= 0xDFFF) return -2;
        units += (ch >= 0x10000) ? 2 : 1;
    }
    if (units > _CELL_TEXT_LIMIT) return -2;

    size_t payload = 12 + 2 * (size_t)units;
    size_t need = *buf_size + 1 + _varint_len((unsigned long)payload) + payload;
    if (need > *buf_cap)
    {
        *buf_cap = need + 4096;
        unsigned char *nb = (unsigned char *)PyMem_Realloc(*buf, *buf_cap);
        if (!nb) return -1;
        *buf = nb;
    }

    unsigned char *p = *buf + *buf_size;
    *p++ = 0x06;
    p = _write_varint_bytes(p, (unsigned long)payload);
    unsigned int cch = (unsigned int)units;
    memcpy(p, &col_idx, 4); p += 4;
    memcpy(p, &style_ref, 4); p += 4;
    memcpy(p, &cch, 4); p += 4;
    for (Py_ssize_t k = 0; k < slen; k++)
    {
        Py_UCS4 ch = PyUnicode_READ(kind, data, k);
        if (ch < 0x10000)
            p = _write_utf16_unit(p, (unsigned int)ch);
        else
        {
            ch -= 0x10000;
            p = _write_utf16_unit(p, 0xD800u | (unsigned int)(ch >> 10));
            p = _write_utf16_unit(p, 0xDC00u | (unsigned int)(ch & 0x3FF));
        }
    }
    *buf_size = (size_t)(p - *buf);
    return 0;
}

/* One text cell in the storage the caller chose.  Same return contract as
   _write_inline_str_cell (-1 error, -2 row fallback, 0 written). */
static int _write_text_cell(unsigned char **buf, size_t *buf_size, size_t *buf_cap,
                            PyObject *cell_str, PyObject *ss_dict, PyObject *ss_list,
                            unsigned int col_idx, unsigned int row_idx, unsigned int style_override,
                            unsigned int *sst_uniq, unsigned int *sst_all, int inline_strings)
{
    if (!inline_strings)
        return _write_str_cell(buf, buf_size, buf_cap, cell_str, ss_dict, ss_list,
                               col_idx, row_idx, style_override, sst_uniq, sst_all);
    unsigned int style_ref = style_override ? style_override : ((row_idx == 0) ? 3 : 0);
    return _write_inline_str_cell(buf, buf_size, buf_cap, cell_str, col_idx, style_ref);
}

/* Encode one row.  Returns new bytes, Py_None when the row must go to the
   Python encoder, or NULL with an exception set.  Shared
   string state is read from and written back to *sst_uniq_io / *sst_all_io
   only when the row is encoded. */
static int _is_fmt_tuple(PyObject *cell, int *out);
static int _xlsx_fmt_xf(PyObject *cell, PyObject *fmt_xf, unsigned int *style,
                        PyObject **value);

static PyObject *_encode_xlsb_row_core(PyObject *row, PyObject *ss_dict, PyObject *ss_list,
                                       unsigned int *sst_uniq_io, unsigned int *sst_all_io,
                                       unsigned int row_idx, PyObject *style_ids, int inline_strings,
                                       PyObject *fmt_xf, PyObject *row_header)
{
    unsigned int sst_uniq = *sst_uniq_io;
    unsigned int sst_all = *sst_all_io;

    if (!PyList_Check(row)) { PyErr_SetString(PyExc_TypeError, "row must be a list"); return NULL; }
    if (!inline_strings)
    {
        if (!PyDict_Check(ss_dict)) { PyErr_SetString(PyExc_TypeError, "ss_dict must be a dict"); return NULL; }
        if (!PyList_Check(ss_list)) { PyErr_SetString(PyExc_TypeError, "ss_list must be a list"); return NULL; }
    }
    if (style_ids != Py_None && !PyList_Check(style_ids) && !PyTuple_Check(style_ids) && !PyLong_Check(style_ids)) { PyErr_SetString(PyExc_TypeError, "style_ids must be an integer, list, tuple, or None"); return NULL; }
    if (fmt_xf != Py_None && !PyDict_Check(fmt_xf)) { PyErr_SetString(PyExc_TypeError, "fmt_xf must be a dict or None"); return NULL; }
    if (row_header != Py_None && (!PyBytes_Check(row_header) || PyBytes_GET_SIZE(row_header) > 256)) { PyErr_SetString(PyExc_TypeError, "row_header must be a BrtRowHdr record (bytes) or None"); return NULL; }
    unsigned int row_style = 0;
    if (PyLong_Check(style_ids))
    {
        unsigned long style_value = PyLong_AsUnsignedLong(style_ids);
        if (PyErr_Occurred()) return NULL;
        row_style = (unsigned int)style_value;
    }

    _ensure_types();

    unsigned char *buf = NULL;
    size_t buf_size = 0;
    size_t buf_cap = 4096;
    int err = 0;

    buf = (unsigned char *)PyMem_Malloc(buf_cap);
    if (!buf) return PyErr_NoMemory();

    Py_ssize_t row_len = PyList_Size(row);
    Py_ssize_t style_len = PyList_Check(style_ids) ? PyList_GET_SIZE(style_ids)
                         : PyTuple_Check(style_ids) ? PyTuple_GET_SIZE(style_ids) : 0;
    Py_ssize_t style_pair_index = 0;
    unsigned int last_col = (row_len > 0) ? (unsigned int)(row_len - 1) : 0;

    if (row_header != Py_None)
    {
        /* A custom height, hidden flag or outline level: the caller's record. */
        memcpy(buf, PyBytes_AS_STRING(row_header), (size_t)PyBytes_GET_SIZE(row_header));
        buf_size = (size_t)PyBytes_GET_SIZE(row_header);
    }
    else
    {
        buf[buf_size++] = 0x00; buf[buf_size++] = 0x19;
        memcpy(buf + buf_size, &row_idx, 4); buf_size += 4;
        memset(buf + buf_size, 0, 4); buf_size += 4;
        unsigned int row_h = 0x12c;
        memcpy(buf + buf_size, &row_h, 4); buf_size += 4;
        buf[buf_size++] = 0x00; buf[buf_size++] = 0x01; buf[buf_size++] = 0x00; buf[buf_size++] = 0x00; buf[buf_size++] = 0x00;
        unsigned int r0 = 0;
        memcpy(buf + buf_size, &r0, 4); buf_size += 4;
        memcpy(buf + buf_size, &last_col, 4); buf_size += 4;
    }

    for (Py_ssize_t col_idx = 0; col_idx < row_len; col_idx++)
    {
        PyObject *cell = PyList_GetItem(row, col_idx);
        unsigned int cell_style = 0;
        if (style_ids != Py_None)
        {
            if (row_style)
                cell_style = row_style;
            else if (PyList_Check(style_ids) && col_idx < style_len)
            {
                PyObject *style_obj = PyList_GetItem(style_ids, col_idx);
                if (style_obj && style_obj != Py_None)
                {
                    unsigned long style_value = PyLong_AsUnsignedLong(style_obj);
                    if (PyErr_Occurred()) { PyErr_Clear(); style_value = 0; }
                    cell_style = (unsigned int)style_value;
                }
            }
            else if (PyTuple_Check(style_ids))
            {
                while (style_pair_index < style_len)
                {
                    PyObject *pair = PyTuple_GET_ITEM(style_ids, style_pair_index);
                    if (!PyTuple_Check(pair) || PyTuple_GET_SIZE(pair) != 2)
                    {
                        PyErr_SetString(PyExc_TypeError, "sparse style entries must be (column, style_id) pairs");
                        PyMem_Free(buf);
                        return NULL;
                    }
                    long style_col = PyLong_AsLong(PyTuple_GET_ITEM(pair, 0));
                    if (PyErr_Occurred()) { PyMem_Free(buf); return NULL; }
                    if (style_col < col_idx)
                    {
                        style_pair_index++;
                        continue;
                    }
                    if (style_col == col_idx)
                    {
                        unsigned long style_value = PyLong_AsUnsignedLong(PyTuple_GET_ITEM(pair, 1));
                        if (PyErr_Occurred()) { PyMem_Free(buf); return NULL; }
                        cell_style = (unsigned int)style_value;
                        style_pair_index++;
                    }
                    break;
                }
            }
        }
        if (cell && PyTuple_Check(cell))
        {
            /* (value, format): the xf the Python writer recorded in fmt_xf,
               else the Python encoder writes the row and records it. */
            int fmt = 0;
            _is_fmt_tuple(cell, &fmt);
            if (fmt)
            {
                PyObject *value = NULL;
                int xf_rc = PyTuple_CheckExact(cell)
                    ? _xlsx_fmt_xf(cell, fmt_xf, &cell_style, &value) : 0;
                if (xf_rc < 0) { err = 1; break; }
                if (xf_rc == 0) { PyMem_Free(buf); Py_RETURN_NONE; }
                /* The exact tuple keeps its item alive while the row is encoded. */
                Py_DECREF(value);
                cell = value;
            }
        }
        if (!cell || cell == Py_None)
        {
            if (cell_style)
            {
                size_t need = buf_size + 10;
                if (need > buf_cap)
                {
                    buf_cap = need + 4096;
                    unsigned char *nb = (unsigned char *)PyMem_Realloc(buf, buf_cap);
                    if (!nb) { err = 1; break; }
                    buf = nb;
                }
                unsigned int ci = (unsigned int)col_idx;
                buf[buf_size++] = 0x01; buf[buf_size++] = 0x08;
                memcpy(buf + buf_size, &ci, 4); buf_size += 4;
                memcpy(buf + buf_size, &cell_style, 4); buf_size += 4;
            }
            continue;
        }

        size_t need = buf_size + 32;
        if (need > buf_cap)
        {
            buf_cap = need + 4096;
            unsigned char *nb = (unsigned char *)PyMem_Realloc(buf, buf_cap);
            if (!nb) { err = 1; break; }
            buf = nb;
        }

        unsigned int ci = (unsigned int)col_idx;
        unsigned int zero = 0;

        if (PyUnicode_Check(cell))
        {
            int rc = _write_text_cell(&buf, &buf_size, &buf_cap, cell, ss_dict, ss_list, ci, row_idx, cell_style, &sst_uniq, &sst_all, inline_strings);
            if (rc == -2) { PyMem_Free(buf); Py_RETURN_NONE; }
            if (rc < 0) { err = 1; break; }
        }
        else if (cell == Py_True || cell == Py_False)
        {
            buf[buf_size++] = 0x04; buf[buf_size++] = 0x09;
            memcpy(buf + buf_size, &ci, 4); buf_size += 4;
            unsigned int style_ref = cell_style ? cell_style : zero;
            memcpy(buf + buf_size, &style_ref, 4); buf_size += 4;
            buf[buf_size++] = (cell == Py_True) ? 1 : 0;
        }
        else if (PyLong_Check(cell))
        {
            long long val = PyLong_AsLongLong(cell);
            if (val == -1 && PyErr_Occurred())
            {
                /* Beyond int64 range: store a double, exactly like the
                   Python path (float(cell) never raises for ints). */
                PyErr_Clear();
                PyObject *f = PyNumber_Float(cell);
                if (!f) { PyErr_Clear(); goto fallback; }
                double fval = PyFloat_AsDouble(f);
                Py_DECREF(f);
                buf[buf_size++] = 0x05; buf[buf_size++] = 0x10;
                memcpy(buf + buf_size, &ci, 4); buf_size += 4;
                unsigned int style_ref = cell_style ? cell_style : zero;
                memcpy(buf + buf_size, &style_ref, 4); buf_size += 4;
                memcpy(buf + buf_size, &fval, 8); buf_size += 8;
            }
            else if (val >= -536870912LL && val <= 536870911LL)
            {
                /* Shift in unsigned arithmetic: left-shifting a negative int is
                   undefined; the stored bit pattern is the same. */
                unsigned int rk_val = ((unsigned int)val << 2) | 2u;
                buf[buf_size++] = 0x02; buf[buf_size++] = 0x0C;
                memcpy(buf + buf_size, &ci, 4); buf_size += 4;
                unsigned int style_ref = cell_style ? cell_style : zero;
                memcpy(buf + buf_size, &style_ref, 4); buf_size += 4;
                memcpy(buf + buf_size, &rk_val, 4); buf_size += 4;
            }
            else
            {
                buf[buf_size++] = 0x05; buf[buf_size++] = 0x10;
                memcpy(buf + buf_size, &ci, 4); buf_size += 4;
                unsigned int style_ref = cell_style ? cell_style : zero;
                memcpy(buf + buf_size, &style_ref, 4); buf_size += 4;
                double fval = (double)val;
                memcpy(buf + buf_size, &fval, 8); buf_size += 8;
            }
        }
        else if (PyFloat_Check(cell))
        {
            buf[buf_size++] = 0x05; buf[buf_size++] = 0x10;
            memcpy(buf + buf_size, &ci, 4); buf_size += 4;
            unsigned int style_ref = cell_style ? cell_style : zero;
            memcpy(buf + buf_size, &style_ref, 4); buf_size += 4;
            double fval = PyFloat_AsDouble(cell);
            memcpy(buf + buf_size, &fval, 8); buf_size += 8;
        }
        else if (decimal_type && PyObject_IsInstance(cell, decimal_type) == 1)
        {
            PyObject *f = PyObject_CallMethod(cell, "__float__", NULL);
            if (!f) { PyErr_Clear(); goto fallback; }
            double fval = PyFloat_AsDouble(f); Py_DECREF(f);
            buf[buf_size++] = 0x05; buf[buf_size++] = 0x10;
            memcpy(buf + buf_size, &ci, 4); buf_size += 4;
            unsigned int style_ref = cell_style ? cell_style : zero;
            memcpy(buf + buf_size, &style_ref, 4); buf_size += 4;
            memcpy(buf + buf_size, &fval, 8); buf_size += 8;
        }
        else if (datetime_type && PyObject_IsInstance(cell, datetime_type) == 1)
        {
            double serial = _get_excel_serial(cell, 1);
            if (serial == -1.0) goto fallback;
            unsigned int sv = 1;
            buf[buf_size++] = 0x05; buf[buf_size++] = 0x10;
            memcpy(buf + buf_size, &ci, 4); buf_size += 4;
            unsigned int style_ref = cell_style ? cell_style : sv;
            memcpy(buf + buf_size, &style_ref, 4); buf_size += 4;
            memcpy(buf + buf_size, &serial, 8); buf_size += 8;
        }
        else if (date_type && PyObject_IsInstance(cell, date_type) == 1)
        {
            double serial = _get_excel_serial(cell, 2);
            if (serial == -1.0) goto fallback;
            unsigned int sv = 2;
            buf[buf_size++] = 0x05; buf[buf_size++] = 0x10;
            memcpy(buf + buf_size, &ci, 4); buf_size += 4;
            unsigned int style_ref = cell_style ? cell_style : sv;
            memcpy(buf + buf_size, &style_ref, 4); buf_size += 4;
            memcpy(buf + buf_size, &serial, 8); buf_size += 8;
        }
        else
        {
            /* Any other object goes to the Python encoder, which refuses
               Formula and Hyperlink cells and writes str() of the rest.  Such
               cells are rare, so the caller need not scan every row. */
            PyMem_Free(buf);
            Py_RETURN_NONE;
            fallback:  /* reached by goto: dates out of range, odd Decimals */
            {
                PyObject *str_cell = PyObject_Str(cell);
                /* str() failed: the Python encoder calls it again and raises. */
                if (!str_cell) { PyErr_Clear(); PyMem_Free(buf); Py_RETURN_NONE; }
                int rc = _write_text_cell(&buf, &buf_size, &buf_cap, str_cell, ss_dict, ss_list, ci, row_idx, cell_style, &sst_uniq, &sst_all, inline_strings);
                Py_DECREF(str_cell);
                if (rc == -2) { PyMem_Free(buf); Py_RETURN_NONE; }
                if (rc < 0) { err = 1; break; }
            }
        }
    }

    if (err)
    {
        PyMem_Free(buf);
        /* Keep a Python error (from the shared-string dict) over a generic one. */
        return PyErr_Occurred() ? NULL : PyErr_NoMemory();
    }

    PyObject *result_bytes = PyBytes_FromStringAndSize((const char *)buf, buf_size);
    PyMem_Free(buf);
    if (result_bytes)
    {
        *sst_uniq_io = sst_uniq;
        *sst_all_io = sst_all;
    }
    return result_bytes;
}

static int _xlsx_rollback(PyObject *ss_dict, PyObject *ss_list, Py_ssize_t len_before);

/* Shared-string row: returns (bytes, sst_uniq, sst_all), or None when the
   Python encoder must handle the row.  On None, the strings this call added
   to the table are removed again, so the caller's counters stay valid. */
static PyObject *C_encode_xlsb_row(PyObject *self, PyObject *args)
{
    PyObject *row;
    PyObject *ss_dict;
    PyObject *ss_list;
    unsigned int sst_uniq;
    unsigned int sst_all;
    unsigned int row_idx;
    PyObject *style_ids = Py_None;
    PyObject *fmt_xf = Py_None;
    PyObject *row_header = Py_None;

    if (!PyArg_ParseTuple(args, "OOOIII|OOO", &row, &ss_dict, &ss_list, &sst_uniq, &sst_all, &row_idx, &style_ids, &fmt_xf, &row_header))
        return NULL;
    if (!PyList_Check(ss_list)) { PyErr_SetString(PyExc_TypeError, "ss_list must be a list"); return NULL; }
    Py_ssize_t list_before = PyList_GET_SIZE(ss_list);

    PyObject *result_bytes = _encode_xlsb_row_core(row, ss_dict, ss_list, &sst_uniq, &sst_all,
                                                   row_idx, style_ids, 0, fmt_xf, row_header);
    if (!result_bytes || result_bytes == Py_None)
    {
        /* Fallback or error: undo the table entries this row added. */
        _xlsx_rollback(ss_dict, ss_list, list_before);
        return result_bytes;
    }

    /* "O" in Py_BuildValue creates a new (incref'd) reference, so the
       local one must be released -- otherwise every call leaks it. */
    PyObject *res = Py_BuildValue("OII", result_bytes, sst_uniq, sst_all);
    Py_DECREF(result_bytes);
    return res;
}

/* Inline-string row: returns the row bytes, or None when the Python encoder
   must handle the row.  Takes no shared-string state. */
static PyObject *C_encode_xlsb_row_inline(PyObject *self, PyObject *args)
{
    PyObject *row;
    unsigned int row_idx;
    PyObject *style_ids = Py_None;
    unsigned int unused_uniq = 0;
    unsigned int unused_all = 0;
    PyObject *fmt_xf = Py_None;
    PyObject *row_header = Py_None;

    if (!PyArg_ParseTuple(args, "OI|OOO", &row, &row_idx, &style_ids, &fmt_xf, &row_header))
        return NULL;

    return _encode_xlsb_row_core(row, NULL, NULL, &unused_uniq, &unused_all,
                                 row_idx, style_ids, 1, fmt_xf, row_header);
}

/* ===== XLSX row encoder =====
   Encodes one worksheet row (`<row><c/>...</row>`) for the plain (unstyled)
   XLSX path.  Mirrors `_write_worksheet_xml` cell-by-cell:
   - None -> `<c/>`; bool -> `<c t="b">`; numbers (int/float/Decimal, with
     int converted via float() first, exactly like the Python path) ->
     `<c><v>repr</v></c>` using the same shortest-repr routine as
     `repr()`; datetimes/dates -> date serials with `s="1"`/`s="2"`;
   - strings go through shared-strings dedup (or inline-escaped cells when
     shared strings are off), with `s="3"` on the header row, like
     `_write_string_cell`;
   - `float('nan')`/`inf` become the strings "NaN"/"\u221e"/"-\u221e"
     (Decimal NaN stays a numeric NaN, like the Python path, which only
     special-cases float instances);
   - out-of-range years become strings; any other object becomes
     `str(cell)`;
   - a row containing a `(value, format)` tuple returns None (no error) so
     the caller falls back to the Python row writer, which owns format
     registration.  Other unencodable rows (non-sequences, lone
     surrogates in inline mode, text over the cell limit, inline text that
     needs _xHHHH_ escapes) also fall back, letting the Python path raise
     or escape exactly as it does on its own.
   Returns (row_xml_str, sst_uniq, sst_all), None for row fallback, or
   NULL+error only on allocation failure. */

/* Growable UTF-8 byte buffer (doubling). */
typedef struct { unsigned char *data; size_t size; size_t cap; } _XBuf;

static int _xbuf_reserve(_XBuf *b, size_t extra)
{
    if (b->size + extra <= b->cap) return 1;
    size_t ncap = b->cap ? b->cap : 1024;
    while (ncap < b->size + extra) ncap *= 2;
    unsigned char *nb = (unsigned char *)PyMem_Realloc(b->data, ncap);
    if (!nb) return 0;
    b->data = nb;
    b->cap = ncap;
    return 1;
}

static int _xbuf_put(_XBuf *b, const void *src, size_t n)
{
    if (n == 0) return 1;
    if (!_xbuf_reserve(b, n)) return 0;
    memcpy(b->data + b->size, src, n);
    b->size += n;
    return 1;
}

static int _xbuf_byte(_XBuf *b, unsigned char c)
{
    return _xbuf_put(b, &c, 1);
}

static int _xbuf_cstr(_XBuf *b, const char *s)
{
    return _xbuf_put(b, s, strlen(s));
}

static int _xbuf_uint(_XBuf *b, unsigned int v)
{
    char tmp[16];
    int n = snprintf(tmp, sizeof(tmp), "%u", v);
    return (n > 0) ? _xbuf_put(b, tmp, (size_t)n) : 0;
}

/* Identical to repr(): the same routine CPython's float repr uses. */
static int _xbuf_double(_XBuf *b, double v)
{
    char *s = PyOS_double_to_string(v, 'r', 0, Py_DTSF_ADD_DOT_0, NULL);
    if (!s) return 0;
    int ok = _xbuf_cstr(b, s);
    PyMem_Free(s);
    return ok;
}

static int _is_hex_byte(unsigned char c)
{
    return (c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f');
}

/* XML-escape & < > over UTF-8 bytes (multibyte units never collide), copying
   runs of plain bytes at once.  Returns 1 written, 0 out of memory, or 2 when
   the text needs an _xHHHH_ escape (see xlspy/_text.py) -- a control
   character other than tab and line feed, U+FFFE/U+FFFF, or literal text of
   the form _xHHHH_ -- in which case the caller hands the row to Python. */
static int _xbuf_escaped(_XBuf *b, const char *s, Py_ssize_t n)
{
    Py_ssize_t run = 0;  /* start of the plain bytes not yet copied */
    for (Py_ssize_t i = 0; i < n; i++)
    {
        unsigned char c = (unsigned char)s[i];
        const char *entity;
        if (c == '&') entity = "&amp;";
        else if (c == '<') entity = "&lt;";
        else if (c == '>') entity = "&gt;";
        else
        {
            if (c < 0x20)
            {
                if (c != '\t' && c != '\n') return 2;
            }
            else if (c == '_')
            {
                if (i + 6 < n && s[i + 1] == 'x' &&
                    _is_hex_byte((unsigned char)s[i + 2]) && _is_hex_byte((unsigned char)s[i + 3]) &&
                    _is_hex_byte((unsigned char)s[i + 4]) && _is_hex_byte((unsigned char)s[i + 5]) &&
                    s[i + 6] == '_')
                    return 2;
            }
            else if (c == 0xEF)
            {
                if (i + 2 < n && (unsigned char)s[i + 1] == 0xBF &&
                    ((unsigned char)s[i + 2] == 0xBE || (unsigned char)s[i + 2] == 0xBF))
                    return 2;
            }
            continue;
        }
        if (!_xbuf_put(b, s + run, (size_t)(i - run))) return 0;
        if (!_xbuf_cstr(b, entity)) return 0;
        run = i + 1;
    }
    return _xbuf_put(b, s + run, (size_t)(n - run));
}

static int _long_attr(PyObject *obj, const char *name, long *out)
{
    PyObject *a = PyObject_GetAttrString(obj, name);
    if (!a) { PyErr_Clear(); return 0; }
    long v = PyLong_AsLong(a);
    Py_DECREF(a);
    if (v == -1 && PyErr_Occurred()) { PyErr_Clear(); return 0; }
    *out = v;
    return 1;
}

/* Howard Hinnant's days_from_civil (days since 1970-01-01). */
static long long _days_from_civil(long long y, long long m, long long d)
{
    long long yy = (m <= 2) ? y - 1 : y;
    const long long era = (yy >= 0 ? yy : yy - 399) / 400;
    const long long yoe = yy - era * 400;
    const long long mp = (m > 2) ? m - 3 : m + 9;
    const long long doy = (153 * mp + 2) / 5 + d - 1;
    const long long doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    return era * 146097 + doe - 719468;
}

/* Excel serial bit-identical to `datetime_to_serial` (which computes
   `(moment - epoch).total_seconds() / 86400.0` with exact integer
   microseconds, i.e. a single rounding per division).
   Returns 1 + serial; 2 = year out of range (caller stringifies);
   0 = row fallback (unreadable/defensive ranges). */
static int _xlsx_serial(PyObject *cell, int is_datetime, double *serial)
{
    long y, mo, d, hh = 0, mi = 0, ss = 0, us = 0;
    if (!_long_attr(cell, "year", &y) || !_long_attr(cell, "month", &mo) ||
        !_long_attr(cell, "day", &d))
        return 0;
    if (is_datetime &&
        (!_long_attr(cell, "hour", &hh) || !_long_attr(cell, "minute", &mi) ||
         !_long_attr(cell, "second", &ss) || !_long_attr(cell, "microsecond", &us)))
        return 0;
    if (y < 1900 || y > 9999)
        return 2;
    if (mo < 1 || mo > 12 || d < 1 || d > 31)
        return 0;
    if (is_datetime &&
        (hh < 0 || hh > 23 || mi < 0 || mi > 59 || ss < 0 || ss > 61 ||
         us < 0 || us > 999999))
        return 0;
    long long ord = _days_from_civil(y, mo, d) + 719163;
    long long epoch_ord = (y == 1900 && mo < 3) ? 693595 : 693594;
    long long whole = (ord - epoch_ord) * 86400LL
        + hh * 3600LL + (long long)mi * 60LL + ss;
    long long total_us = whole * 1000000LL + us;
    if (total_us > 9007199254740992LL || total_us < -9007199254740992LL)
        return 0;  /* beyond exact-double range: slow path stays correct */
    *serial = (double)total_us / 1000000.0 / 86400.0;
    return 1;
}

/* Shared-string index, mirroring `_write_string_cell` dedup. */
static int _xlsx_sst_index(PyObject *ss_dict, PyObject *ss_list, PyObject *s,
                           unsigned int *uniq, unsigned int *all,
                           unsigned int *out_idx)
{
    PyObject *idx = PyDict_GetItemWithError(ss_dict, s);
    if (!idx && PyErr_Occurred()) return 0;
    if (idx)
    {
        *out_idx = (unsigned int)PyLong_AsLong(idx);
    }
    else
    {
        *out_idx = *uniq;
        PyObject *si = PyLong_FromLong((long)*uniq);
        if (!si) return 0;
        /* List first: a failed insert then leaves only a list entry, which the
           rollback removes. A dict key without a list entry would survive it. */
        if (PyList_Append(ss_list, s) < 0) { Py_DECREF(si); return 0; }
        if (PyDict_SetItem(ss_dict, s, si) < 0) { Py_DECREF(si); return 0; }
        Py_DECREF(si);
        (*uniq)++;
    }
    (*all)++;
    return 1;
}

/* String cell, mirroring `_write_string_cell` (1 ok, 0 row fallback, -1 OOM).
   Text over the cell limit falls back, so the Python encoder raises. */
static int _xlsx_style_attr(_XBuf *b, unsigned int style_id)
{
    if (!style_id) return 1;
    return _xbuf_cstr(b, " s=\"") && _xbuf_uint(b, style_id) &&
           _xbuf_byte(b, '"');
}

/* "<c><v>" or "<c s=\"N\"><v>" for a numeric cell. */
static int _xlsx_open_value(_XBuf *b, unsigned int style_id)
{
    if (!style_id) return _xbuf_cstr(b, "<c><v>");
    return _xbuf_cstr(b, "<c s=\"") && _xbuf_uint(b, style_id) &&
           _xbuf_cstr(b, "\"><v>");
}

static int _xlsx_string_cell(_XBuf *b, PyObject *s, int is_header,
                             unsigned int style_id, int use_sst,
                             PyObject *ss_dict, PyObject *ss_list,
                             unsigned int *uniq, unsigned int *all)
{
    if (!_cell_text_fits(s)) return 0;
    if (use_sst)
    {
        unsigned int idx;
        if (!_xlsx_sst_index(ss_dict, ss_list, s, uniq, all, &idx))
            return -1;
        if (!_xbuf_cstr(b, "<c t=\"s\"")) return -1;
        if (style_id)
        {
            if (!_xlsx_style_attr(b, style_id)) return -1;
        }
        else if (is_header && !_xbuf_cstr(b, " s=\"3\"")) return -1;
        if (!_xbuf_cstr(b, "><v>")) return -1;
        if (!_xbuf_uint(b, idx)) return -1;
        if (!_xbuf_cstr(b, "</v></c>")) return -1;
        return 1;
    }
    Py_ssize_t n = 0;
    const char *utf8 = PyUnicode_AsUTF8AndSize(s, &n);
    if (!utf8) { PyErr_Clear(); return 0; }  /* e.g. lone surrogates */
    int preserve = (n > 0 && (utf8[0] == ' ' || utf8[0] == '\t'));
    if (!_xbuf_cstr(b, "<c t=\"inlineStr\"")) return -1;
    if (style_id)
    {
        if (!_xlsx_style_attr(b, style_id)) return -1;
    }
    else if (is_header && !_xbuf_cstr(b, " s=\"3\"")) return -1;
    if (!_xbuf_cstr(b, "><is><t")) return -1;
    if (preserve && !_xbuf_cstr(b, " xml:space=\"preserve\"")) return -1;
    if (!_xbuf_byte(b, '>')) return -1;
    int escaped = _xbuf_escaped(b, utf8, n);
    if (escaped == 2) return 0;  /* needs _xHHHH_ escapes: Python writes it */
    if (!escaped) return -1;
    if (!_xbuf_cstr(b, "</t></is></c>")) return -1;
    return 1;
}

/* Exact `_is_formatted_cell`: tuple + len 2 + a str or Style second item;
   *out is its `_cell_overlay_kind`. */
static int _is_fmt_tuple(PyObject *cell, int *out)
{
    *out = 0;
    if (!PyTuple_Check(cell))
        return 1;
    Py_ssize_t tlen = PyObject_Length(cell);
    if (tlen < 0) { PyErr_Clear(); return 1; }
    if (tlen != 2)
        return 1;
    PyObject *fmt = PySequence_GetItem(cell, 1);
    if (!fmt) { PyErr_Clear(); return 1; }
    *out = _cell_overlay_kind(fmt);
    Py_DECREF(fmt);
    return 1;
}

/* Resolve a (value, format) or (value, Style) cell through fmt_xf (1 ok,
   0 row fallback, -1 error). The key is the format string, or the Style's
   id(), for an unstyled cell and (style id, that key) for a styled one; the
   writer keeps every keyed Style alive, so an id cannot be reused. On
   success *style becomes the xf and *value a new reference to the cell's
   value. A None value (written as the text "None" by the Python path) and
   nested tuples fall back. */
static int _xlsx_fmt_xf(PyObject *cell, PyObject *fmt_xf, unsigned int *style,
                        PyObject **value)
{
    if (fmt_xf == Py_None) return 0;
    PyObject *fmt = PySequence_GetItem(cell, 1);
    if (!fmt) { PyErr_Clear(); return 0; }
    if (!PyUnicode_Check(fmt))
    {
        PyObject *ident = PyLong_FromVoidPtr(fmt);
        Py_DECREF(fmt);
        if (!ident) return -1;
        fmt = ident;
    }
    PyObject *key = fmt;
    if (*style)
    {
        key = Py_BuildValue("(IO)", *style, fmt);
        Py_DECREF(fmt);
        if (!key) return -1;
    }
    PyObject *xf = PyDict_GetItemWithError(fmt_xf, key);  /* borrowed */
    Py_DECREF(key);
    if (!xf) return PyErr_Occurred() ? -1 : 0;
    unsigned long xf_value = PyLong_AsUnsignedLong(xf);
    if (PyErr_Occurred()) return -1;
    PyObject *inner = PySequence_GetItem(cell, 0);
    if (!inner) { PyErr_Clear(); return 0; }
    if (inner == Py_None || PyTuple_Check(inner)) { Py_DECREF(inner); return 0; }
    *style = (unsigned int)xf_value;
    *value = inner;
    return 1;
}

/* Encode one cell (1 ok, 0 row fallback, -1 OOM). */
static int _xlsx_cell(_XBuf *b, PyObject *cell, int is_header,
                      unsigned int style_id, int use_sst,
                      PyObject *ss_dict, PyObject *ss_list,
                      unsigned int *uniq, unsigned int *all)
{
    if (!cell || cell == Py_None)
    {
        if (!style_id) return _xbuf_cstr(b, "<c/>") ? 1 : -1;
        if (!_xbuf_cstr(b, "<c")) return -1;
        if (!_xlsx_style_attr(b, style_id)) return -1;
        return _xbuf_cstr(b, "/>") ? 1 : -1;
    }
    if (PyUnicode_Check(cell))
        return _xlsx_string_cell(b, cell, is_header, style_id, use_sst,
                                 ss_dict, ss_list, uniq, all);
    if (cell == Py_True || cell == Py_False)
    {
        if (!style_id)
            return _xbuf_cstr(b, cell == Py_True ? "<c t=\"b\"><v>1</v></c>"
                                                 : "<c t=\"b\"><v>0</v></c>") ? 1 : -1;
        if (!_xbuf_cstr(b, "<c t=\"b\"") ||
            !_xlsx_style_attr(b, style_id) ||
            !_xbuf_cstr(b, cell == Py_True ? "><v>1</v></c>" : "><v>0</v></c>"))
            return -1;
        return 1;
    }
    if (PyLong_Check(cell))
    {
        double d;
        if (PyLong_CheckExact(cell))
        {
            long long v = PyLong_AsLongLong(cell);
            if (v == -1 && PyErr_Occurred())
            {
                /* Beyond int64: float(cell), which yields inf and never
                   raises for ints (PyLong_AsDouble alone would leave
                   OverflowError set). */
                PyErr_Clear();
                PyObject *f = PyNumber_Float(cell);
                if (!f) { PyErr_Clear(); return 0; }
                d = PyFloat_AsDouble(f);
                Py_DECREF(f);
            }
            else
                d = (double)v;
        }
        else
        {
            PyObject *f = PyNumber_Float(cell);
            if (!f) { PyErr_Clear(); return 0; }
            d = PyFloat_AsDouble(f);
            Py_DECREF(f);
        }
        if (!_xlsx_open_value(b, style_id)) return -1;
        if (!_xbuf_double(b, d)) return -1;
        if (!_xbuf_cstr(b, "</v></c>")) return -1;
        return 1;
    }
    if (PyFloat_Check(cell))
    {
        double d;
        int nan_like = 0, pos_inf = 0, neg_inf = 0;
        if (PyFloat_CheckExact(cell))
        {
            d = PyFloat_AS_DOUBLE(cell);
            nan_like = (d != d);
            pos_inf = (d == Py_HUGE_VAL);
            neg_inf = (d == -Py_HUGE_VAL);
        }
        else
        {
            PyObject *f = PyNumber_Float(cell);
            if (!f) { PyErr_Clear(); return 0; }
            d = PyFloat_AsDouble(f);
            Py_DECREF(f);
            /* Exact routing for subclasses (mirrors ==/!= dispatch). */
            int rb = PyObject_RichCompareBool(cell, cell, Py_NE);
            if (rb < 0) { PyErr_Clear(); return 0; }
            nan_like = rb;
            if (!nan_like)
            {
                PyObject *pinf = PyFloat_FromDouble(Py_HUGE_VAL);
                PyObject *ninf = pinf ? PyFloat_FromDouble(-Py_HUGE_VAL) : NULL;
                if (!pinf || !ninf)
                {
                    Py_XDECREF(pinf);
                    Py_XDECREF(ninf);
                    return -1;
                }
                int e1 = PyObject_RichCompareBool(cell, pinf, Py_EQ);
                int e2 = (e1 == 0) ? PyObject_RichCompareBool(cell, ninf, Py_EQ) : 0;
                Py_DECREF(pinf);
                Py_DECREF(ninf);
                if (e1 < 0 || e2 < 0) { PyErr_Clear(); return 0; }
                pos_inf = (e1 == 1);
                neg_inf = (e2 == 1);
            }
        }
        if (nan_like || pos_inf || neg_inf)
        {
            PyObject *s = PyUnicode_FromString(
                nan_like ? "NaN" : (pos_inf ? "\xe2\x88\x9e" : "-\xe2\x88\x9e"));
            if (!s) return -1;
            int rc = _xlsx_string_cell(b, s, is_header, style_id, use_sst,
                                       ss_dict, ss_list, uniq, all);
            Py_DECREF(s);
            return rc;
        }
        if (!_xlsx_open_value(b, style_id)) return -1;
        if (!_xbuf_double(b, d)) return -1;
        if (!_xbuf_cstr(b, "</v></c>")) return -1;
        return 1;
    }
    if (decimal_type && PyObject_IsInstance(cell, decimal_type) == 1)
    {
        PyObject *f = PyNumber_Float(cell);
        if (!f) { PyErr_Clear(); return 0; }
        double d = PyFloat_AsDouble(f);
        Py_DECREF(f);
        if (!_xlsx_open_value(b, style_id)) return -1;
        if (!_xbuf_double(b, d)) return -1;
        if (!_xbuf_cstr(b, "</v></c>")) return -1;
        return 1;
    }
    int is_dt = (datetime_type && PyObject_IsInstance(cell, datetime_type) == 1);
    if (is_dt || (date_type && PyObject_IsInstance(cell, date_type) == 1))
    {
        double serial = 0.0;
        int sr = _xlsx_serial(cell, is_dt, &serial);
        if (sr == 0)
            return 0;  /* row fallback (unreadable/defensive ranges) */
        if (sr == 2)
        {
            /* Year out of range: str(cell), like the Python path. */
            PyObject *s = PyObject_Str(cell);
            if (!s) { PyErr_Clear(); return 0; }
            int rc = _xlsx_string_cell(b, s, is_header, style_id, use_sst,
                                       ss_dict, ss_list, uniq, all);
            Py_DECREF(s);
            return rc;
        }
        if (!(style_id ? _xlsx_open_value(b, style_id)
                       : _xbuf_cstr(b, is_dt ? "<c s=\"2\"><v>" : "<c s=\"1\"><v>")))
            return -1;
        if (!_xbuf_double(b, serial)) return -1;
        if (!_xbuf_cstr(b, "</v></c>")) return -1;
        return 1;
    }
    /* Any other object goes to the Python encoder, which writes Formula and
       Hyperlink cells and str() of everything else.  Such cells are rare, so
       the caller need not scan every row for them. */
    return 0;
}

/* Remove SST entries appended during an aborted row encode, so a row
   fallback leaves caller-visible structures untouched (the caller keeps
   its old counters and reprocesses the row in Python).  Only entries
   added after len_before are removed; pre-existing ones are never
   touched. */
static int _xlsx_rollback(PyObject *ss_dict, PyObject *ss_list,
                          Py_ssize_t len_before)
{
    /* The deletes can run Python code, so keep the error the caller is about
       to return (if any) and put it back afterwards. */
    PyObject *pending = PyErr_GetRaisedException();
    Py_ssize_t n = PyList_GET_SIZE(ss_list);
    for (Py_ssize_t i = len_before; i < n; i++)
    {
        PyObject *key = PyList_GET_ITEM(ss_list, i);  /* borrowed */
        if (PyDict_DelItem(ss_dict, key) < 0)
            PyErr_Clear();  /* the key was never inserted: nothing to undo */
    }
    if (n > len_before &&
        PySequence_DelSlice(ss_list, len_before, PY_SSIZE_T_MAX) < 0)
        PyErr_Clear();
    PyErr_SetRaisedException(pending);
    return 0;
}

static PyObject *C_encode_xlsx_row(PyObject *self, PyObject *args)
{
    PyObject *row, *ss_dict, *ss_list;
    PyObject *style_ids = Py_None;
    PyObject *fmt_xf = Py_None;
    PyObject *row_attrs = Py_None;
    unsigned int sst_uniq, sst_all, row_idx;
    int use_sst;

    if (!PyArg_ParseTuple(args, "OOOIIIp|OOO", &row, &ss_dict, &ss_list,
                          &sst_uniq, &sst_all, &row_idx, &use_sst, &style_ids,
                          &fmt_xf, &row_attrs))
        return NULL;
    const char *attrs = NULL;
    if (row_attrs != Py_None)
    {
        if (!PyUnicode_Check(row_attrs))
        {
            PyErr_SetString(PyExc_TypeError, "row_attrs must be a str or None");
            return NULL;
        }
        /* ASCII attributes built by the writer (ht, outlineLevel, hidden...). */
        attrs = PyUnicode_AsUTF8(row_attrs);
        if (!attrs) return NULL;
    }
    if (fmt_xf != Py_None && !PyDict_Check(fmt_xf))
    {
        PyErr_SetString(PyExc_TypeError, "fmt_xf must be a dict or None");
        return NULL;
    }
    if (!PyDict_Check(ss_dict) || !PyList_Check(ss_list))
    {
        PyErr_SetString(PyExc_TypeError, "ss_dict must be a dict, ss_list a list");
        return NULL;
    }
    if (style_ids != Py_None && !PyList_Check(style_ids) &&
        !PyTuple_Check(style_ids) && !PyLong_Check(style_ids))
    {
        PyErr_SetString(PyExc_TypeError,
                        "style_ids must be an integer, list, tuple, or None");
        return NULL;
    }
    unsigned int row_style = 0;
    if (PyLong_Check(style_ids))
    {
        unsigned long style_value = PyLong_AsUnsignedLong(style_ids);
        if (PyErr_Occurred()) return NULL;
        row_style = (unsigned int)style_value;
    }

    _ensure_types();

    PyObject *seq = PySequence_Fast(row, "row must be a sequence");
    if (!seq) { PyErr_Clear(); Py_RETURN_NONE; }
    Py_ssize_t n = PySequence_Fast_GET_SIZE(seq);
    PyObject **items = PySequence_Fast_ITEMS(seq);
    Py_ssize_t list_before = PyList_GET_SIZE(ss_list);
    Py_ssize_t style_len = PyList_Check(style_ids) ? PyList_GET_SIZE(style_ids)
                         : PyTuple_Check(style_ids) ? PyTuple_GET_SIZE(style_ids) : 0;
    Py_ssize_t style_pair_index = 0;

    _XBuf b = {NULL, 0, 0};
    /* <row r="{n}"> */
    {
        char head[64];
        int hn = snprintf(head, sizeof(head), attrs ? "<row r=\"%u\"" : "<row r=\"%u\">",
                          row_idx + 1);
        if (hn <= 0 || !_xbuf_cstr(&b, head)) goto oom;
        if (attrs && (!_xbuf_cstr(&b, attrs) || !_xbuf_byte(&b, '>'))) goto oom;
    }
    for (Py_ssize_t i = 0; i < n; i++)
    {
        PyObject *cell = items[i];
        unsigned int cell_style = row_style;
        if (!row_style && PyList_Check(style_ids) && i < style_len)
        {
            PyObject *style_obj = PyList_GET_ITEM(style_ids, i);
            if (style_obj != Py_None)
            {
                unsigned long style_value = PyLong_AsUnsignedLong(style_obj);
                if (PyErr_Occurred()) goto oom;
                cell_style = (unsigned int)style_value;
            }
        }
        else if (!row_style && PyTuple_Check(style_ids))
        {
            while (style_pair_index < style_len)
            {
                PyObject *pair = PyTuple_GET_ITEM(style_ids, style_pair_index);
                if (!PyTuple_Check(pair) || PyTuple_GET_SIZE(pair) != 2)
                {
                    PyErr_SetString(PyExc_TypeError,
                                    "sparse style entries must be (column, style_id) pairs");
                    goto oom;
                }
                long style_col = PyLong_AsLong(PyTuple_GET_ITEM(pair, 0));
                if (PyErr_Occurred()) goto oom;
                if (style_col < i)
                {
                    style_pair_index++;
                    continue;
                }
                if (style_col == i)
                {
                    unsigned long style_value =
                        PyLong_AsUnsignedLong(PyTuple_GET_ITEM(pair, 1));
                    if (PyErr_Occurred()) goto oom;
                    cell_style = (unsigned int)style_value;
                    style_pair_index++;
                }
                break;
            }
        }
        PyObject *value = NULL;  /* owned: the value of a (value, format) cell */
        if (cell != Py_None)
        {
            int fmt = 0;
            if (PyTuple_Check(cell))
            {
                if (!_is_fmt_tuple(cell, &fmt)) goto oom;
                /* (value, format): use the xf the Python writer recorded in
                   fmt_xf for this format (and cell style); otherwise the
                   Python path writes the row and records it. */
                int xf_rc = fmt ? _xlsx_fmt_xf(cell, fmt_xf, &cell_style, &value) : 1;
                if (xf_rc < 0) goto oom;
                if (xf_rc == 0)
                {
                    Py_DECREF(seq);
                    PyMem_Free(b.data);
                    _xlsx_rollback(ss_dict, ss_list, list_before);
                    Py_RETURN_NONE;
                }
            }
        }
        int rc = _xlsx_cell(&b, value ? value : cell, row_idx == 0, cell_style,
                            use_sst, ss_dict, ss_list, &sst_uniq, &sst_all);
        Py_XDECREF(value);
        if (rc == 0)
        {
            Py_DECREF(seq);
            PyMem_Free(b.data);
            _xlsx_rollback(ss_dict, ss_list, list_before);
            Py_RETURN_NONE;
        }
        if (rc < 0) goto oom;
    }
    if (!_xbuf_cstr(&b, "</row>")) goto oom;
    Py_DECREF(seq);

    {
        PyObject *s = PyUnicode_DecodeUTF8((const char *)b.data, (Py_ssize_t)b.size, "strict");
        PyMem_Free(b.data);
        /* "O" creates a new reference: release the local one (no leaks). */
        PyObject *res = s ? Py_BuildValue("OII", s, sst_uniq, sst_all) : NULL;
        Py_XDECREF(s);
        if (!res) _xlsx_rollback(ss_dict, ss_list, list_before);
        return res;
    }

oom:
    Py_DECREF(seq);
    PyMem_Free(b.data);
    /* Undo this row's table entries; keeps any error raised above. */
    _xlsx_rollback(ss_dict, ss_list, list_before);
    /* Keep a Python error (from the shared-string dict) over a generic one. */
    if (!PyErr_Occurred()) PyErr_NoMemory();
    return NULL;
}

/* ===== Streaming XLSB worksheet reader (iterator type) ===== */

/* Forward declaration for iternext */
static PyObject *XlsbReader_iternext(PyObject *self);

typedef struct {
    PyObject_HEAD
    /* Data buffer */
    Py_buffer view;
    const unsigned char *data;
    Py_ssize_t data_len;
    Py_ssize_t pos;
    int done;

    /* Lookup structures */
    PyObject *shared_strings;
    Py_ssize_t ss_len;
    PyObject *styles_list;
    Py_ssize_t styles_len;
    PyObject *date_num_fmts;

    /* Per-row cell buffer */
    PyObject **cur_cells;
    Py_ssize_t cur_cells_len;

    /* State tracking */
    long cur_row_num;          /* row number of currently accumulating row */
    long gap_rows_remaining;   /* number of empty rows to emit before next data row */
    /* first_col/cols_per_row: layout frozen from the first row with cells.
       first_col == -1 means unset (no prescan layout); then the first
       streamed cell freezes it, exactly like the reference, which latches
       the first cell's column. */
    Py_ssize_t first_col;      /* first column of the first row that has cells */
    Py_ssize_t cols_per_row;   /* width of the first row that has cells */
    int is_initialized;        /* 1 after first Row record consumed */
    int seen_row_record;       /* 1 after the first Row record was consumed */
    int fallback_to_python;    /* prescan found a layout too wide for C storage */
    int date_1904;             /* 1 when the workbook uses the 1904 date system */
} XlsbReader;

/* Bound C's dense per-row cell array. The XLSB format maximum is 16384
   columns; wider nonstandard layouts are delegated to the Python reference
   rather than assigning them a C-specific result or risking huge allocation. */
#define _XLSB_MAX_COL_SPAN 1048576L

static void XlsbReader_dealloc(XlsbReader *self)
{
    if (self->cur_cells)
    {
        /* Uniform ownership: every slot holds an owned reference (singletons
           are incref'd on store), so a mid-row abandonment frees cleanly. */
        for (Py_ssize_t c = 0; c < self->cur_cells_len; c++)
            Py_XDECREF(self->cur_cells[c]);
        free(self->cur_cells);
    }
    if (self->view.buf)
        PyBuffer_Release(&self->view);
    Py_XDECREF(self->shared_strings);
    Py_XDECREF(self->styles_list);
    Py_XDECREF(self->date_num_fmts);
    Py_TYPE(self)->tp_free((PyObject *)self);
}

static PyTypeObject XlsbReader_Type = {
    PyVarObject_HEAD_INIT(NULL, 0)
    .tp_name = "_c_core.XlsbReader",
    .tp_basicsize = sizeof(XlsbReader),
    .tp_dealloc = (destructor)XlsbReader_dealloc,
    .tp_flags = Py_TPFLAGS_DEFAULT,
    .tp_doc = "Streaming XLSB worksheet reader (iterator)",
    .tp_iter = PyObject_SelfIter,
    .tp_iternext = XlsbReader_iternext,
};

static PyObject *_build_row_and_free_cells(XlsbReader *self)
{
    /* Unset layout (-1) falls back to the accumulated width; a frozen
       degenerate layout (<= 0, from an unordered first row) yields empty
       rows, exactly like the reference's range(<=0). */
    Py_ssize_t out_cols;
    if (self->cols_per_row == -1)
        out_cols = (self->cur_cells_len > 0 ? self->cur_cells_len : 1);
    else
        out_cols = (self->cols_per_row > 0) ? self->cols_per_row : 0;
    if (out_cols > _XLSB_MAX_COL_SPAN)
    {
        /* Absurd width: the reference dies building range(huge). */
        _free_cur_cells(self->cur_cells, self->cur_cells_len);
        self->cur_cells = NULL; self->cur_cells_len = 0;
        PyErr_NoMemory();
        return NULL;
    }
    PyObject *row_list = PyList_New(out_cols);
    if (!row_list)
    {
        _free_cur_cells(self->cur_cells, self->cur_cells_len);
        self->cur_cells = NULL; self->cur_cells_len = 0;
        return NULL;
    }
    for (Py_ssize_t c = 0; c < out_cols; c++)
    {
        if (c < self->cur_cells_len && self->cur_cells[c])
            PyList_SET_ITEM(row_list, c, self->cur_cells[c]);
        else
        {
            Py_INCREF(Py_None);
            PyList_SET_ITEM(row_list, c, Py_None);
        }
    }
    /* Cells past the output width are dropped like the reference's
       range(width); they are owned, so release them. */
    for (Py_ssize_t c = out_cols < 0 ? 0 : out_cols; c < self->cur_cells_len; c++)
        Py_XDECREF(self->cur_cells[c]);
    /* Cells stolen by PyList_SET_ITEM - just free the array */
    free(self->cur_cells);
    self->cur_cells = NULL;
    self->cur_cells_len = 0;
    return row_list;
}

static PyObject *_make_empty_row(XlsbReader *self)
{
    Py_ssize_t cols;
    if (self->cols_per_row == -1)
        cols = 1;
    else
        cols = (self->cols_per_row > 0) ? self->cols_per_row : 0;
    if (cols > _XLSB_MAX_COL_SPAN)
    {
        /* Absurd width: the reference dies building [None] * huge. */
        PyErr_NoMemory();
        return NULL;
    }
    PyObject *row = PyList_New(cols);
    if (!row) return NULL;
    for (Py_ssize_t c = 0; c < cols; c++)
    {
        Py_INCREF(Py_None);
        PyList_SET_ITEM(row, c, Py_None);
    }
    return row;
}

/* Whether a cell record counts as a cell for layout purposes: mirrors the
   reference, which ignores undersized bodies and overrunning string lengths
   (the streaming pass skips them too, so the prescan must not freeze the
   layout on records the stream pass will drop). */
static int _counts_for_layout(long rec_id, long rec_len,
                              const unsigned char *data, Py_ssize_t data_len,
                              Py_ssize_t body_start)
{
    if (rec_len < 8) return 0;
    if (body_start + (Py_ssize_t)rec_len > data_len) return 0;
    if (rec_id == 0x02 && rec_len < 12) return 0;
    if ((rec_id == 0x04 || rec_id == 0x0A) && rec_len < 9) return 0;
    if ((rec_id == 0x05 || rec_id == 0x09) && rec_len < 16) return 0;
    if (rec_id == 0x07 && rec_len < 12) return 0;
    if (rec_id == 0x06 || rec_id == 0x08)
    {
        if (rec_len < 12) return 0;
        long len_raw;
        if (!_read_uint32_safe(data, data_len, body_start + 8, &len_raw))
            return 0;
        Py_ssize_t need = (Py_ssize_t)(unsigned long)len_raw * 2 + 12;
        if (need > (Py_ssize_t)rec_len) return 0;
    }
    return 1;
}

/* Pre-scan worksheet records to determine the column layout (first_col,
   cols_per_row) of the first row that contains any cell.  Doing this upfront
   keeps empty leading rows aligned with the data width; tracking the layout
   only while streaming would leave it unset whenever the first Row record has
   no cells and would mis-index later cells. */
static void _prescan_first_data_row(XlsbReader *self)
{
    const unsigned char *data = self->data;
    Py_ssize_t data_len = self->data_len;
    Py_ssize_t pos = 0;
    long cur_row = -1;
    long span_row = -2;        /* row of the first row with cells (-2 = none) */
    /* First/last cell column in record order (unsigned 64-bit: dword columns
       >= 2^31 are huge positive indices to the reference). */
    Py_ssize_t first_c = -1, last_c = -1;

    while (pos < data_len)
    {
        long rec_id = _read_varint(data, data_len, &pos);
        if (rec_id < 0) break;
        long rec_len = _read_varint(data, data_len, &pos);
        if (rec_len < 0) break;

        Py_ssize_t body_start = pos;
        pos += rec_len;

        if (rec_id == 0x00)
        {
            /* The next row header marks the end of the first data row.
               Short bodies read as row -1, like the reference. */
            if (span_row != -2 && cur_row != span_row)
                break;
            long new_row;
            if (!_read_row_index(data, data_len, body_start, &new_row))
            {
                /* Iteration reaches this record too and raises there. */
                PyErr_Clear();
                break;
            }
            cur_row = new_row;
        }
        else if (rec_id >= 0x01 && rec_id <= 0x0B)
        {
            /* Skipped records (undersized/overrunning bodies) are not cells
               to the reference: they must neither open the span nor feed it,
               or leading empty rows would freeze a wrong width. */
            if (!_counts_for_layout(rec_id, rec_len, data, data_len, body_start))
                continue;
            if (span_row == -2) span_row = cur_row;
            if (cur_row == span_row)
            {
                /* First/last cell column in record order, exactly like the
                   reference (which remembers the first cell's column and
                   overwrites the last column per cell).  Full unsigned range:
                   dword columns >= 2^31 are huge positive indices there. */
                long col_raw;
                if (_read_uint32_safe(data, data_len, body_start, &col_raw))
                {
                    Py_ssize_t col = (Py_ssize_t)(unsigned long)col_raw;
                    if (first_c < 0) first_c = col;
                    last_c = col;
                }
            }
        }
        else if (rec_id == 0x92)
        {
            break;
        }
    }

    if (span_row != -2 && first_c >= 0)
    {
        self->first_col = first_c;
        /* May be <= 0 for unordered first rows; the reference then emits
           empty rows, which the builders below reproduce. */
        self->cols_per_row = last_c - first_c + 1;
        if (self->cols_per_row > _XLSB_MAX_COL_SPAN)
            self->fallback_to_python = 1;
    }
    else
    {
        self->first_col = -1;
        self->cols_per_row = -1;
    }
}

/* Helper: read a cell record and store in cur_cells.
   Returns 1 (stored or reference-identical skip), -1 (fatal error pending),
   0 (out of memory, PyErr set). */
static int _process_cell_record(XlsbReader *self, const unsigned char *data,
                                 Py_ssize_t data_len, Py_ssize_t body_start,
                                 long rec_id, long rec_len)
{
    /* BiffReader first rejects every cell record shorter than the 8-byte
       column/style header, regardless of record type. */
    if (rec_len < 8) return 1;
    long col_raw, xf_idx;
    if (!_read_uint32_safe(data, data_len, body_start, &col_raw)) return 1;
    if (!_read_uint32_safe(data, data_len, body_start + 4, &xf_idx)) return 1;
    xf_idx &= 0xFFFFFF;
    /* Truncated framing (the record claims more bytes than the stream
       holds) kills the record, exactly like the reference, whose exact-size
       body read fails and stops.  Length checks below mirror the reference's
       per-type minimums against the real bytes, not the claim. */
    if (body_start + (Py_ssize_t)rec_len > data_len) return 1;
    /* Bit-preserving conversion: dword columns >= 2^31 read negative into
       a 32-bit long but are huge positive indices to the reference. */
    Py_ssize_t col = (Py_ssize_t)(unsigned long)col_raw;

    int has_cell_alloc;
    int fatal = 0, skip = 0;
    PyObject *cell_val = _parse_cell_value(data, data_len, body_start,
                                             rec_id, rec_len,
                                             self->shared_strings, self->ss_len,
                                             self->styles_list, self->styles_len,
                                             self->date_num_fmts, xf_idx, &has_cell_alloc,
                                             self->date_1904, &fatal, &skip);
    if (fatal) return -1;  /* exception pending: abort like the reference */
    if (skip || !cell_val) return 1;

    /* Frozen layout bounds the visible window, exactly like the reference's
       range(width): degenerate layouts (width <= 0) show nothing, absurd
       ones (> cap) die at build time like the reference's range(huge). */
    if (self->cols_per_row != -1)
    {
        if (self->cols_per_row <= 0 ||
            self->cols_per_row > _XLSB_MAX_COL_SPAN ||
            col < self->first_col ||
            col - self->first_col >= self->cols_per_row)
        {
            /* Only the rare out-of-window branch does extra work. Ordinary
               rectangular exports execute the same checks as before. Valid
               later rows may be wider than their header: keep their cells. */
            if (self->cols_per_row > 0 &&
                self->cols_per_row <= _XLSB_MAX_COL_SPAN &&
                col >= self->first_col && col < 16384 &&
                col - self->first_col >= self->cols_per_row)
            {
                self->cols_per_row = col - self->first_col + 1;
            }
            else
            {
                if (has_cell_alloc) Py_DECREF(cell_val);
                return 1;
            }
        }
    }
    else
    {
        /* Unset layout: freeze on the first stored cell; absurd spans are
           corrupt-input noise the reference never shows. */
        if (self->first_col < 0)
            self->first_col = col;
        if (col < self->first_col ||
            col - self->first_col > _XLSB_MAX_COL_SPAN)
        {
            if (has_cell_alloc) Py_DECREF(cell_val);
            return 1;
        }
    }
    Py_ssize_t idx = col - self->first_col;

    if (idx >= self->cur_cells_len)
    {
        PyObject **new_cells = realloc(self->cur_cells,
                                       sizeof(PyObject*) * (size_t)(idx + 1));
        if (!new_cells)
        {
            if (has_cell_alloc) Py_DECREF(cell_val);
            PyErr_NoMemory();
            return 0;
        }
        for (Py_ssize_t c = self->cur_cells_len; c <= idx; c++)
            new_cells[c] = NULL;
        self->cur_cells = new_cells;
        self->cur_cells_len = idx + 1;
    }
    /* Uniform array ownership: own a reference on borrowed singletons. */
    if (!has_cell_alloc) Py_INCREF(cell_val);
    Py_XDECREF(self->cur_cells[idx]);
    self->cur_cells[idx] = cell_val;
    return 1;
}

static PyObject *XlsbReader_iternext(PyObject *self_ptr)
{
    XlsbReader *self = (XlsbReader *)self_ptr;
    if (self->done) return NULL;

    const unsigned char *data = self->data;
    Py_ssize_t data_len = self->data_len;

    /* Phase 0: Emit pending gap rows (one per __next__ call) */
    if (self->gap_rows_remaining > 0)
    {
        self->gap_rows_remaining--;
        PyObject *row = _make_empty_row(self);
        if (!row) { self->done = 1; return NULL; }
        return row;
    }

    /* Phase 1: On first call, skip past metadata to the first Row record
       or the first cell record.  Cells before any Row record form a phantom
       row, exactly like the reference, which latches its first cell. */
    if (!self->is_initialized)
    {
        while (self->pos < data_len)
        {
            Py_ssize_t rec_start = self->pos;
            long rec_id = _read_varint(data, data_len, &self->pos);
            if (rec_id < 0) { self->done = 1; return NULL; }
            long rec_len = _read_varint(data, data_len, &self->pos);
            if (rec_len < 0) { self->done = 1; return NULL; }
            Py_ssize_t body_start = self->pos;
            self->pos += rec_len;
            if (rec_id == 0x00)
            {
                long new_row;
                if (!_read_row_index(data, data_len, body_start, &new_row))
                { self->done = 1; return NULL; }
                self->cur_row_num = new_row;
                self->seen_row_record = 1;
                self->is_initialized = 1;
                break;
            }
            else if (rec_id >= 0x01 && rec_id <= 0x0B)
            {
                /* Phantom row: rewind so Phase 2 consumes this cell. */
                self->pos = rec_start;
                self->is_initialized = 1;
                break;
            }
        }
        if (!self->is_initialized) { self->done = 1; return NULL; }
    }

    /* Phase 2: Read cell records until next Row record or EOF */
    while (self->pos < data_len)
    {
        long rec_id = _read_varint(data, data_len, &self->pos);
        if (rec_id < 0) break;
        long rec_len = _read_varint(data, data_len, &self->pos);
        if (rec_len < 0) break;

        Py_ssize_t body_start = self->pos;
        self->pos += rec_len;

        if (rec_id == 0x00)
        {
            /* Next Row record: build and return the accumulated row */
            PyObject *row = _build_row_and_free_cells(self);
            if (!row) { self->done = 1; return NULL; }

            long new_row;
            if (!_read_row_index(data, data_len, body_start, &new_row))
            {
                Py_DECREF(row);
                self->done = 1;
                return NULL;
            }

            /* Gap rows only once a Row record was seen: the reference
               suppresses gaps while its previous-row tracker is still -1
               (cells before the first Row record). */
            if (self->seen_row_record)
                self->gap_rows_remaining = new_row - self->cur_row_num - 1;
            else
                self->gap_rows_remaining = 0;

            self->cur_row_num = new_row;
            self->seen_row_record = 1;
            return row;
        }
        else if (rec_id >= 0x01 && rec_id <= 0x0B)
        {
            int pr = _process_cell_record(self, data, data_len, body_start, rec_id, rec_len);
            if (pr < 0) { self->done = 1; return NULL; }  /* ref error pending */
            if (!pr) { self->done = 1; return NULL; }
        }
        else if (rec_id == 0x92)
        {
            break;
        }
    }

    /* Phase 3: EOF - emit last row */
    self->done = 1;
    if (self->cur_cells && self->cur_cells_len > 0)
        return _build_row_and_free_cells(self);
    if (self->cur_row_num >= 0)
        return _make_empty_row(self);
    return NULL;
}

/* ===== Factory function ===== */
static PyObject *C_read_xlsb_worksheet(PyObject *self_mod, PyObject *args)
{
    Py_buffer view;
    PyObject *shared_strings;
    PyObject *styles_list;
    PyObject *date_num_fmts;
    int date_1904 = 0;

    if (!PyArg_ParseTuple(args, "y*OOO|p", &view, &shared_strings, &styles_list, &date_num_fmts, &date_1904))
        return NULL;

    if (!PyList_Check(shared_strings)) { PyErr_SetString(PyExc_TypeError, "shared_strings must be a list"); PyBuffer_Release(&view); return NULL; }
    if (!PyList_Check(styles_list)) { PyErr_SetString(PyExc_TypeError, "styles_list must be a list"); PyBuffer_Release(&view); return NULL; }
    if (!PySet_Check(date_num_fmts)) { PyErr_SetString(PyExc_TypeError, "date_num_fmts must be a set"); PyBuffer_Release(&view); return NULL; }

    _ensure_types();

    XlsbReader *reader = PyObject_New(XlsbReader, &XlsbReader_Type);
    if (!reader) { PyBuffer_Release(&view); return NULL; }

    memcpy(&reader->view, &view, sizeof(Py_buffer));
    reader->data = (const unsigned char *)view.buf;
    reader->data_len = (Py_ssize_t)view.len;
    reader->pos = 0;
    reader->done = 0;

    reader->shared_strings = shared_strings;
    reader->ss_len = PyList_Size(shared_strings);
    reader->styles_list = styles_list;
    reader->styles_len = PyList_Size(styles_list);
    reader->date_num_fmts = date_num_fmts;

    reader->cur_cells = NULL;
    reader->cur_cells_len = 0;
    reader->cur_row_num = -1;
    reader->gap_rows_remaining = 0;
    reader->first_col = -1;
    reader->cols_per_row = -1;
    reader->is_initialized = 0;
    reader->seen_row_record = 0;
    reader->fallback_to_python = 0;
    reader->date_1904 = date_1904;

    _prescan_first_data_row(reader);

    Py_INCREF(shared_strings);
    Py_INCREF(styles_list);
    Py_INCREF(date_num_fmts);

    if (reader->fallback_to_python)
    {
        /* Keep behavior exact for unusually wide layouts; the high-level
           reader already treats None as a signal to use BiffReader. */
        Py_DECREF(reader);
        Py_RETURN_NONE;
    }

    return (PyObject *)reader;
}

/* ===== Module definition ===== */
static PyMethodDef CMethods[] = {
    {"read_xlsb_worksheet", C_read_xlsb_worksheet, METH_VARARGS,
     "Read an XLSB worksheet from raw bytes. Returns an iterator of rows.\n\n"
     "Args: (data, shared_strings, styles_list, date_num_fmts)\n"
     "Yields: rows one at a time"},
    {"encode_xlsb_row", C_encode_xlsb_row, METH_VARARGS,
     "Encode a single XLSB row. Returns (bytes, sst_uniq, sst_all)."},
    {"encode_xlsb_row_inline", C_encode_xlsb_row_inline, METH_VARARGS,
     "Encode a single XLSB row with inline text cells. Returns bytes, or None to use the Python encoder."},
    {"calc_column_widths", C_calc_column_widths, METH_VARARGS,
     "Calculate column widths from row data."},
    {"xlsb_row_has_formatted", C_xlsb_row_has_formatted, METH_VARARGS,
     "Return True if a row contains any (value, format) or (value, Style) cell."},
    {"set_style_type", C_set_style_type, METH_O,
     "Register xlspy.Style, the class of the second item of (value, Style) cells."},
    {"encode_xlsb_sst", C_encode_xlsb_sst, METH_VARARGS,
     "Encode the sharedStrings.bin part. Returns bytes."},
    {"read_xlsb_shared_strings", C_read_xlsb_shared_strings, METH_VARARGS,
     "Parse a sharedStrings.bin part. Returns a list of strings."},
    {"encode_xlsx_row", C_encode_xlsx_row, METH_VARARGS,
     "Encode one XLSX sheet row. Returns (row_xml, sst_uniq, sst_all)."},
    {NULL, NULL, 0, NULL}
};

static struct PyModuleDef c_core_module = {
    PyModuleDef_HEAD_INIT,
    "_c_core",
    "C-optimized functions for xlspy",
    -1,
    CMethods
};

PyMODINIT_FUNC PyInit__c_core(void)
{
    PyDateTime_IMPORT;
    if (PyType_Ready(&XlsbReader_Type) < 0)
        return NULL;
    PyObject *m = PyModule_Create(&c_core_module);
    return m;
}
