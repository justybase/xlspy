//! Workbook-model cells (`xlspy.workbook._cell.Cell`) created natively.
//!
//! The loaders fill the `__slots__` of each Cell directly, through the member
//! descriptors the caller passes (`CELL_SLOTS`), exactly as the C extension
//! did; the descriptor layout comes from pyo3-ffi's CPython definitions.

use pyo3::exceptions::PyTypeError;
use pyo3::ffi;
use pyo3::prelude::*;
use pyo3::types::{
    PyBool, PyDate, PyDateTime, PyDict, PyFloat, PyInt, PyList, PySet, PyString, PyTuple, PyType,
};

/// Slot indices used after creation; `CELL_SLOTS` (workbook/_model.py) is
/// _sheet, _row, _column, _value, _formula, _data_type, _style_id, _comment,
/// _hyperlink, _named_style.
pub const ROW: usize = 1;
pub const COLUMN: usize = 2;
const VALUE: usize = 3;
const FORMULA: usize = 4;
const STYLE_ID: usize = 6;
const COMMENT: usize = 7;
const HYPERLINK: usize = 8;
const COUNT: usize = 10;

/// The Cell class and the byte offsets of its slots.
pub struct CellClass {
    type_: Py<PyType>,
    offsets: [isize; COUNT],
}

impl CellClass {
    /// Read the slot offsets from a tuple of `__slots__` member descriptors.
    /// Each must be an object member (`T_OBJECT_EX`) of `cell_type` or one
    /// of its bases, so that writing through the offset is sound.
    pub fn new(cell_type: &Bound<'_, PyAny>, descriptors: &Bound<'_, PyAny>) -> PyResult<Self> {
        let cell_type = cell_type
            .cast::<PyType>()
            .map_err(|_| PyTypeError::new_err("model needs the Cell type and a cells dict"))?;
        let descriptors = match descriptors.cast::<PyTuple>() {
            Ok(t) if t.len() == COUNT => t,
            _ => {
                return Err(PyTypeError::new_err(
                    "expected a tuple of the Cell slot descriptors",
                ));
            }
        };
        let mut offsets = [0isize; COUNT];
        for (i, d) in descriptors.iter().enumerate() {
            // SAFETY: the type check guarantees the member-descriptor layout.
            let member = unsafe {
                if ffi::PyObject_TypeCheck(d.as_ptr(), &raw mut ffi::PyMemberDescr_Type) == 0 {
                    None
                } else {
                    let descr = d.as_ptr() as *mut ffi::PyMemberDescrObject;
                    let owner = (*descr).d_common.d_type;
                    let def = (*descr).d_member;
                    let owns = !owner.is_null()
                        && ffi::PyType_IsSubtype(cell_type.as_type_ptr(), owner) != 0;
                    (owns && !def.is_null() && (*def).type_code == ffi::Py_T_OBJECT_EX)
                        .then(|| (*def).offset)
                }
            };
            offsets[i] = member.ok_or_else(|| {
                PyTypeError::new_err("Cell slot descriptors must be __slots__ members")
            })?;
        }
        Ok(CellClass {
            type_: cell_type.clone().unbind(),
            offsets,
        })
    }

    /// Visit the class, for the owner's `__traverse__`.
    pub fn traverse(&self, visit: &pyo3::PyVisit<'_>) -> Result<(), pyo3::PyTraverseError> {
        visit.call(&self.type_)
    }

    /// A new, uninitialised Cell (every slot unset), as `tp_alloc` makes it.
    pub fn alloc<'py>(&self, py: Python<'py>) -> PyResult<Bound<'py, PyAny>> {
        let tp = self.type_.bind(py).as_type_ptr();
        // SAFETY: tp is a live type object; tp_alloc returns a new reference.
        unsafe {
            let alloc = (*tp).tp_alloc.unwrap_or(ffi::PyType_GenericAlloc);
            Bound::from_owned_ptr_or_err(py, alloc(tp, 0))
        }
    }

    /// A new Cell with every slot set, in `CELL_SLOTS` order (the ROW slot
    /// may be None to leave it unset until the row is complete).
    pub fn new_cell<'py>(
        &self,
        py: Python<'py>,
        values: [Option<&Bound<'py, PyAny>>; COUNT],
    ) -> PyResult<Bound<'py, PyAny>> {
        let cell = self.alloc(py)?;
        let base = cell.as_ptr() as *mut u8;
        for (offset, value) in self.offsets.iter().zip(values) {
            if let Some(value) = value {
                // SAFETY: a fresh instance of the class the offsets belong to;
                // its object slots start NULL, so nothing is released.
                unsafe {
                    *(base.offset(*offset) as *mut *mut ffi::PyObject) = value.clone().into_ptr()
                };
            }
        }
        Ok(cell)
    }

    /// Store `value` in slot `slot` of `cell` (an instance of the Cell class).
    #[inline]
    pub fn set(&self, cell: &Bound<'_, PyAny>, slot: usize, value: &Bound<'_, PyAny>) {
        // SAFETY: `cell` is an instance of the class whose descriptors gave the
        // offsets (checked in `new`), and object slots hold owned references.
        unsafe {
            let p =
                (cell.as_ptr() as *mut u8).offset(self.offsets[slot]) as *mut *mut ffi::PyObject;
            let old = *p;
            *p = value.clone().into_ptr();
            ffi::Py_XDECREF(old);
        }
    }

    /// The object in slot `slot` of `cell`, if set.
    #[inline]
    pub fn get<'py>(&self, cell: &Bound<'py, PyAny>, slot: usize) -> Option<Bound<'py, PyAny>> {
        // SAFETY: as in `set`; a borrowed slot value is turned into a new reference.
        unsafe {
            let p =
                (cell.as_ptr() as *mut u8).offset(self.offsets[slot]) as *mut *mut ffi::PyObject;
            Bound::from_borrowed_ptr_or_opt(cell.py(), *p)
        }
    }
}

/// The row dict of a cells dict: C used PyDict_SetItem on it, which raises
/// SystemError for anything but a dict.
pub fn as_row_dict<'py>(row_dict: Bound<'py, PyAny>) -> PyResult<Bound<'py, PyDict>> {
    row_dict
        .cast_into::<PyDict>()
        .map_err(|_| pyo3::exceptions::PySystemError::new_err("bad argument to internal function"))
}

/// A row of plain values as `Worksheet.append` stores them, the rows of a
/// workbook as the writers take them, and the cells that need more than that.
/// Each leaves what it does not handle to the Python code it stands in for
/// (`workbook/_worksheet.py::_append_new_row`, `workbook/_writer.py`).
#[pyclass(module = "xlspy._rs_core", frozen)]
pub struct ModelCells {
    class: CellClass,
    number: Py<PyString>,
    text: Py<PyString>,
    boolean: Py<PyString>,
}

#[pymethods]
impl ModelCells {
    #[new]
    fn new(
        py: Python<'_>,
        cell_type: &Bound<'_, PyAny>,
        descriptors: &Bound<'_, PyAny>,
    ) -> PyResult<Self> {
        Ok(ModelCells {
            class: CellClass::new(cell_type, descriptors)?,
            number: PyString::intern(py, "n").unbind(),
            text: PyString::intern(py, "s").unbind(),
            boolean: PyString::intern(py, "b").unbind(),
        })
    }

    fn __traverse__(&self, visit: pyo3::PyVisit<'_>) -> Result<(), pyo3::PyTraverseError> {
        self.class.traverse(&visit)
    }

    /// Store `values[start..]` in `row_cells`, up to the first value that is not
    /// None or an exact int, float, bool or str without a leading `=`. Returns
    /// its index (`len(values)` when there is none).
    fn append_plain<'py>(
        &self,
        sheet: &Bound<'py, PyAny>,
        row: &Bound<'py, PyAny>,
        values: &Bound<'py, PyList>,
        start: usize,
        row_cells: &Bound<'py, PyDict>,
        dirty: Option<&Bound<'py, PySet>>,
    ) -> PyResult<usize> {
        let py = sheet.py();
        let none = py.None().into_bound(py);
        let zero = 0i64.into_pyobject(py)?.into_any();
        let (number, text, boolean) = (
            self.number.bind(py).as_any(),
            self.text.bind(py).as_any(),
            self.boolean.bind(py).as_any(),
        );
        let mut position = start;
        while position < values.len() {
            let value = values.get_item(position)?;
            if value.is_none() {
                position += 1;
                continue;
            }
            let data_type = if value.is_exact_instance_of::<PyInt>()
                || value.is_exact_instance_of::<PyFloat>()
            {
                number
            } else if value.is_exact_instance_of::<PyBool>() {
                boolean
            } else if let Ok(string) = value.cast_exact::<PyString>() {
                match string.to_str() {
                    Ok(content) if !content.starts_with('=') => text,
                    _ => return Ok(position),
                }
            } else {
                return Ok(position);
            };
            let column = (position + 1).into_pyobject(py)?.into_any();
            let cell = self.class.new_cell(
                py,
                [
                    Some(sheet),
                    Some(row),
                    Some(&column),
                    Some(&value),
                    Some(&none),
                    Some(data_type),
                    Some(&zero),
                    Some(&none),
                    Some(&none),
                    Some(&none),
                ],
            )?;
            row_cells.set_item(&column, cell)?;
            if let Some(dirty) = dirty {
                dirty.add((row, &column))?;
            }
            position += 1;
        }
        Ok(position)
    }

    /// One row for the writers: each value, or `(value, Style)` for a styled one,
    /// up to the row's last cell. None for a row with a formula, a hyperlink or
    /// any other value than None, int, float, bool, str, date and datetime.
    fn row_values<'py>(
        &self,
        py: Python<'py>,
        row_cells: &Bound<'py, PyDict>,
        styles: &Bound<'py, PyDict>,
        style_at: &Bound<'py, PyAny>,
    ) -> PyResult<Option<Bound<'py, PyList>>> {
        let none = py.None().into_bound(py);
        let cell_type = self.class.type_.bind(py);
        let mut row: Vec<Bound<'py, PyAny>> = Vec::new();
        for (key, cell) in row_cells.iter() {
            if !cell.is_instance(cell_type)? {
                return Ok(None);
            }
            let (Some(value), Some(formula), Some(link), Some(style_id)) = (
                self.class.get(&cell, VALUE),
                self.class.get(&cell, FORMULA),
                self.class.get(&cell, HYPERLINK),
                self.class.get(&cell, STYLE_ID),
            ) else {
                return Ok(None);
            };
            if !formula.is_none() || !link.is_none() {
                return Ok(None);
            }
            let plain = value.is_none()
                || value.is_exact_instance_of::<PyString>()
                || value.is_exact_instance_of::<PyFloat>()
                || value.is_exact_instance_of::<PyInt>()
                || value.is_exact_instance_of::<PyBool>()
                || value.is_exact_instance_of::<PyDateTime>()
                || value.is_exact_instance_of::<PyDate>();
            let Ok(column) = key.extract::<usize>() else {
                return Ok(None);
            };
            if !plain || !(1..=16384).contains(&column) {
                return Ok(None);
            }
            let item = if !value.is_none() && style_id.is_truthy()? {
                let style = match styles.get_item(&style_id)? {
                    Some(style) => style,
                    None => {
                        let style = style_at.call1((&style_id,))?;
                        styles.set_item(&style_id, &style)?;
                        style
                    }
                };
                PyTuple::new(py, [value, style])?.into_any()
            } else {
                value
            };
            if column > row.len() {
                row.resize(column, none.clone());
            }
            row[column - 1] = item;
        }
        PyList::new(py, row).map(Some)
    }

    /// The cells of a sheet with a style and nothing else, and those with a note.
    /// None for anything but rows of Cells.
    fn styled_and_commented<'py>(
        &self,
        py: Python<'py>,
        cells: &Bound<'py, PyDict>,
    ) -> PyResult<Option<(Bound<'py, PyList>, Bound<'py, PyList>)>> {
        let cell_type = self.class.type_.bind(py);
        let (styled, commented) = (PyList::empty(py), PyList::empty(py));
        for row_cells in cells.values() {
            let Ok(row_cells) = row_cells.cast::<PyDict>() else {
                return Ok(None);
            };
            for cell in row_cells.values() {
                if !cell.is_instance(cell_type)? {
                    return Ok(None);
                }
                let (Some(value), Some(formula), Some(link), Some(style_id), Some(comment)) = (
                    self.class.get(&cell, VALUE),
                    self.class.get(&cell, FORMULA),
                    self.class.get(&cell, HYPERLINK),
                    self.class.get(&cell, STYLE_ID),
                    self.class.get(&cell, COMMENT),
                ) else {
                    return Ok(None);
                };
                if value.is_none() && formula.is_none() && link.is_none() && style_id.is_truthy()? {
                    styled.append(&cell)?;
                }
                if !comment.is_none() {
                    commented.append(&cell)?;
                }
            }
        }
        Ok(Some((styled, commented)))
    }
}

#[cfg(test)]
mod tests {
    use super::{COLUMN, COUNT, CellClass, ROW};
    use pyo3::exceptions::PyTypeError;
    use pyo3::prelude::*;
    use pyo3::types::{PyDict, PyTuple};

    const SLOTS: [&str; COUNT] = [
        "_sheet",
        "_row",
        "_column",
        "_value",
        "_formula",
        "_data_type",
        "_style_id",
        "_comment",
        "_hyperlink",
        "_named_style",
    ];

    /// A class with the Cell slot layout, a subclass of it, an unrelated
    /// class with the same slot names, and a class whose slot is a property.
    fn classes(py: Python<'_>) -> Bound<'_, PyDict> {
        let globals = PyDict::new(py);
        let code = format!(
            "class Cell:\n    __slots__ = {SLOTS:?}\n\
             class SubCell(Cell):\n    __slots__ = ()\n\
             class Other:\n    __slots__ = {SLOTS:?}\n\
             class Prop:\n    _row = property(lambda self: 1)\n"
        );
        let code = std::ffi::CString::new(code).unwrap();
        py.run(&code, Some(&globals), None).unwrap();
        globals
    }

    fn descriptors<'py>(owner: &Bound<'py, PyAny>) -> Bound<'py, PyTuple> {
        let items: Vec<_> = SLOTS
            .iter()
            .map(|name| owner.getattr(*name).unwrap())
            .collect();
        PyTuple::new(owner.py(), items).unwrap()
    }

    #[test]
    fn slots_are_written_and_read_through_the_member_offsets() {
        Python::initialize();
        Python::attach(|py| {
            let g = classes(py);
            let cell_type = g.get_item("Cell").unwrap().unwrap();
            let class = CellClass::new(&cell_type, &descriptors(&cell_type)).unwrap();

            let value = "hello".into_pyobject(py).unwrap().into_any();
            let mut values: [Option<&Bound<'_, PyAny>>; COUNT] = [None; COUNT];
            values[3] = Some(&value);
            let cell = class.new_cell(py, values).unwrap();
            assert!(cell.is_exact_instance(cell_type.cast::<pyo3::types::PyType>().unwrap()));
            assert_eq!(
                cell.getattr("_value").unwrap().extract::<String>().unwrap(),
                "hello"
            );
            assert!(
                class.get(&cell, ROW).is_none(),
                "a slot left out stays unset"
            );
            assert!(cell.getattr("_row").is_err());

            let seven = 7i64.into_pyobject(py).unwrap().into_any();
            class.set(&cell, COLUMN, &seven);
            class.set(&cell, COLUMN, &seven); // replacing releases the old reference
            assert_eq!(
                cell.getattr("_column").unwrap().extract::<i64>().unwrap(),
                7
            );
            assert!(class.get(&cell, COLUMN).unwrap().is(&seven));
        });
    }

    #[test]
    fn a_subclass_may_use_the_descriptors_of_its_base() {
        Python::initialize();
        Python::attach(|py| {
            let g = classes(py);
            let base = g.get_item("Cell").unwrap().unwrap();
            let sub = g.get_item("SubCell").unwrap().unwrap();
            assert!(CellClass::new(&sub, &descriptors(&base)).is_ok());
        });
    }

    #[test]
    fn descriptors_that_would_make_the_offsets_unsound_are_refused() {
        Python::initialize();
        Python::attach(|py| {
            let g = classes(py);
            let cell_type = g.get_item("Cell").unwrap().unwrap();
            let other = g.get_item("Other").unwrap().unwrap();
            let prop = g.get_item("Prop").unwrap().unwrap();

            // Members of an unrelated class: their offsets mean nothing for Cell.
            let err = CellClass::new(&cell_type, &descriptors(&other))
                .err()
                .unwrap();
            assert!(err.is_instance_of::<PyTypeError>(py));

            // A property in place of a member descriptor.
            let mut items: Vec<_> = descriptors(&cell_type).iter().collect();
            items[ROW] = prop.getattr("_row").unwrap();
            let mixed = PyTuple::new(py, items).unwrap();
            assert!(CellClass::new(&cell_type, &mixed).is_err());

            // The wrong number of descriptors, and a non-class.
            let short = PyTuple::new(py, descriptors(&cell_type).iter().take(3)).unwrap();
            assert!(CellClass::new(&cell_type, &short).is_err());
            let not_a_class = 1i64.into_pyobject(py).unwrap().into_any();
            assert!(CellClass::new(&not_a_class, &descriptors(&cell_type)).is_err());
        });
    }
}
