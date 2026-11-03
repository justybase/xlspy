//! The column-width meter's hot path: the Rust twin of
//! `ColumnWidthMeter.add` in xlspy/_widths.py.
//!
//! Python owns the rules and hands over what it decided once: each cell
//! *kind* (font, number format compiled by `_format_plan`, wrap) with its
//! character widths, and a callback that measures any cell this module does
//! not handle itself. Rust measures text and numbers (General or a number
//! format), `(value, format)` / `(value, Style)` overlays, formulas and
//! links, and caches the width of booleans and dates, whose text depends
//! only on the kind. A cell whose widest rendering fits the column's width
//! so far is skipped: the bounds are exact, so the widths do not change.

use std::collections::HashMap;

use pyo3::exceptions::PyValueError;
use pyo3::ffi;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyFloat, PyInt, PyString, PyTuple, PyType};
use pyo3::{PyTraverseError, PyVisit};

use crate::cells::fmt_tuple_kind;
use crate::pyseq::FastSeq;
use crate::text::CodePoints;
use crate::types::{
    decimal_type, formula_type, hyperlink_type, int_as_i64, style_type, type_check,
};

/// One pixel more than Excel's AutoFit (see `_SPARE_PX` in _widths.py).
const SPARE_PX: i64 = 1;

/// A font, a number format and a wrap flag: what a cell's width depends on.
struct Kind {
    /// Pixel advance by code point; -1 where Calibri has no glyph.
    table: Vec<i32>,
    /// Characters outside the table, as the Python callback measured them.
    extra: HashMap<u32, i64>,
    /// `char_px(char) -> int`: `_char_pixels` for a character the table lacks.
    char_px: Py<PyAny>,
    margin: i64,
    /// Pixels an alignment indent adds to the text (`_indent_pixels`).
    indent: i64,
    /// Numbers are shown in General (no number format, or "General").
    general: bool,
    wrap: bool,
    /// The widest column a number in General can need: positive with an
    /// exponent of at most two digits (below 1e100), and any number (i64::MAX:
    /// unknown).
    number_max: i64,
    negative_max: i64,
    /// Same with the minus sign, below 1e100.
    negative_small_max: i64,
    /// The widest digit, `-`, and the widest character of Latin-1 text (or i64::MAX).
    digit_max: i64,
    minus: i64,
    latin1_max: i64,
    /// The number format, compiled by Python (`_format_plan`); None when the
    /// numbers are in General, or when Python measures them itself.
    format: Option<FormatPlan>,
}

/// `(advances, margin, general, wrap, char_px, format plan, indent)` from `_kind_spec`.
type KindSpec<'py> = (
    Bound<'py, PyDict>,
    i64,
    bool,
    bool,
    Py<PyAny>,
    Option<Bound<'py, PyAny>>,
    i64,
);

/// One section of a number format, as `_format_plan` compiles it.
#[derive(Clone, Copy)]
enum Section {
    /// General: `_general_number`, measured as shown.
    General,
    /// Text whatever the number (a date or a literal), `px` wide.
    Constant(i64),
    /// `@` sections: the literal text `px` wide and `count` General renderings.
    Text { px: i64, count: i64 },
    /// Digit placeholders between a head and a tail of literal text.
    Number {
        head: i64,
        tail: i64,
        decimals: usize,
        thousands: bool,
        min_digits: usize,
        percent: u32,
        /// Digits of the exponent (`0.00E+00`: 2), None for a fixed form.
        exponent: Option<usize>,
    },
}

/// A compiled number format: the first section, the negative one when the
/// format has a non-empty second section, and whether it has two or more.
#[derive(Clone, Copy)]
struct FormatPlan {
    positive: Section,
    negative: Option<Section>,
    sections: bool,
}

/// A number as `_shown` formats it: an int keeps its exact value.
#[derive(Clone, Copy)]
enum Num {
    Int(i64),
    Float(f64),
}

impl Num {
    fn is_negative(self) -> bool {
        match self {
            Num::Int(v) => v < 0,
            Num::Float(v) => v < 0.0,
        }
    }

    fn negate(self) -> Option<Num> {
        match self {
            Num::Int(v) => v.checked_neg().map(Num::Int),
            Num::Float(v) => Some(Num::Float(-v)),
        }
    }

    /// `value * 100 ** percent` as the float `format()` sees (None: out of range).
    fn scaled(self, percent: u32) -> Option<f64> {
        let value = match self {
            Num::Int(v) => {
                let factor = 100i128.checked_pow(percent)?;
                (v as i128).checked_mul(factor)? as f64
            }
            // 100 ** percent is exact as a float up to 1e22.
            Num::Float(v) if percent <= 11 => v * 100f64.powi(percent as i32),
            Num::Float(_) => return None,
        };
        value.is_finite().then_some(value)
    }

    /// `_general_number` of the value.
    fn general(self) -> Buf {
        match self {
            Num::Int(v) => general(Dec::from_i64(v)),
            Num::Float(v) => general(Dec::from_f64(v)),
        }
    }
}

/// `format(value, ".{decimals}f")`, with `,` thousands and zero-padded to
/// `min_digits` as `_formatted_number_text` does; `value` is not negative.
fn fixed_text(value: f64, decimals: usize, thousands: bool, min_digits: usize) -> String {
    let plain = format!("{value:.decimals$}");
    let (whole, rest) = match plain.find('.') {
        Some(dot) => plain.split_at(dot),
        None => (plain.as_str(), ""),
    };
    let mut out = String::with_capacity(plain.len() + whole.len() / 3 + min_digits);
    if thousands {
        for (i, ch) in whole.chars().enumerate() {
            if i > 0 && (whole.len() - i) % 3 == 0 {
                out.push(',');
            }
            out.push(ch);
        }
    } else {
        out.push_str(whole);
    }
    if whole.len() < min_digits {
        // str.zfill: zeros in front up to min_digits characters, commas included.
        let pad = min_digits.saturating_sub(out.len());
        out.insert_str(0, &"0".repeat(pad));
    }
    out.push_str(rest);
    out
}

/// `f"{value:.{decimals}E}"` with the exponent zero-padded to `digits`.
fn scientific_text(value: f64, decimals: usize, digits: usize) -> Option<String> {
    let plain = format!("{value:.decimals$e}");
    let (mantissa, power) = plain.split_once('e')?;
    let power: i32 = power.parse().ok()?;
    let sign = if power < 0 { '-' } else { '+' };
    Some(format!(
        "{mantissa}E{sign}{:0digits$}",
        power.unsigned_abs()
    ))
}

impl Kind {
    fn char_width(&mut self, py: Python<'_>, cp: u32) -> PyResult<i64> {
        if let Some(&px) = self.table.get(cp as usize)
            && px >= 0
        {
            return Ok(px as i64);
        }
        if let Some(&px) = self.extra.get(&cp) {
            return Ok(px);
        }
        // SAFETY: PyUnicode_FromOrdinal takes any code point up to 0x10FFFF
        // (lone surrogates included) and returns a new reference or NULL with
        // an exception set.
        let ch = unsafe { Bound::from_owned_ptr_or_err(py, ffi::PyUnicode_FromOrdinal(cp as _))? };
        let px: i64 = self.char_px.bind(py).call1((ch,))?.extract()?;
        self.extra.insert(cp, px);
        Ok(px)
    }

    /// Pixels of a str: the widest line when wrapped, all lines joined otherwise.
    fn text_pixels(&mut self, py: Python<'_>, text: &Bound<'_, PyString>) -> PyResult<i64> {
        let points = CodePoints::of(text)?;
        let (mut line, mut widest) = (0i64, 0i64);
        let wrap = self.wrap;
        let mut step = |kind: &mut Kind, cp: u32| -> PyResult<()> {
            if cp == 0x0A {
                if wrap {
                    widest = widest.max(line);
                    line = 0;
                }
            } else {
                line += kind.char_width(py, cp)?;
            }
            Ok(())
        };
        match points {
            CodePoints::Ucs1(d) => {
                for &c in d {
                    step(self, c as u32)?;
                }
            }
            CodePoints::Ucs2(d) => {
                for &c in d {
                    step(self, c as u32)?;
                }
            }
            CodePoints::Ucs4(d) => {
                for &c in d {
                    step(self, c)?;
                }
            }
        }
        Ok(widest.max(line))
    }

    /// Pixels of a short literal text (NaN, ∞).
    fn literal_pixels(&mut self, py: Python<'_>, text: &str) -> PyResult<i64> {
        let mut px = 0;
        for ch in text.chars() {
            px += self.char_width(py, ch as u32)?;
        }
        Ok(px)
    }

    /// Pixels of ASCII text as shown (no room rules).
    fn ascii_pixels(&mut self, py: Python<'_>, text: &[u8]) -> PyResult<i64> {
        let mut px = 0;
        for &b in text {
            px += self.char_width(py, b as u32)?;
        }
        Ok(px)
    }

    /// Pixels of a number in this kind's number format
    /// (`_formatted_number_text`); None when Python must measure it.
    fn formatted_pixels(&mut self, py: Python<'_>, number: Num) -> PyResult<Option<i64>> {
        let Some(plan) = self.format else {
            return Ok(None);
        };
        let (section, number) = match plan.negative {
            Some(negative) if number.is_negative() => match number.negate() {
                Some(positive) => (negative, positive),
                None => return Ok(None),
            },
            _ => (plan.positive, number),
        };
        Ok(Some(match section {
            Section::General => self.ascii_pixels(py, number.general().as_bytes())?,
            Section::Constant(px) => px,
            Section::Text { px, count } => {
                px + count * self.ascii_pixels(py, number.general().as_bytes())?
            }
            Section::Number {
                head,
                tail,
                decimals,
                thousands,
                min_digits,
                percent,
                exponent,
            } => {
                let Some(value) = number.scaled(percent) else {
                    return Ok(None);
                };
                let text = match exponent {
                    Some(digits) => match scientific_text(value.abs(), decimals, digits) {
                        Some(text) => text,
                        None => return Ok(None),
                    },
                    None => fixed_text(value.abs(), decimals, thousands, min_digits),
                };
                // A negative number that rounds to zero shows no sign.
                let sign = value < 0.0
                    && !plan.sections
                    && text.bytes().any(|b| !matches!(b, b'0' | b'.' | b','));
                let sign_px = if sign {
                    self.char_width(py, b'-' as u32)?
                } else {
                    0
                };
                head + sign_px + self.ascii_pixels(py, text.as_bytes())? + tail
            }
        }))
    }

    /// Pixels of the room AutoFit keeps for a General number: its text with
    /// the decimal point as wide as a digit and a negative exponent as E+.
    fn room_pixels(&mut self, py: Python<'_>, text: &[u8]) -> PyResult<i64> {
        let mut px = 0;
        let mut previous = 0u8;
        for &b in text {
            let shown = match b {
                b'.' => b'0',
                b'-' if previous == b'E' => b'+',
                _ => b,
            };
            px += self.char_width(py, shown as u32)?;
            previous = b;
        }
        Ok(px)
    }

    /// True when a finite float in General cannot need more than `best`.
    fn float_fits(&self, value: f64, best: i64) -> bool {
        // An exponent of two digits at most: 1e-99 <= |value| < 1e100, with
        // room for rounding up to the next power.
        let magnitude = value.abs();
        (magnitude == 0.0 || (1.00001e-99..9.99999e99).contains(&magnitude))
            && best
                >= if value < 0.0 {
                    self.negative_small_max
                } else {
                    self.number_max
                }
    }

    /// Pixels of the room for a finite float shown in General.
    fn float_room_pixels(&mut self, py: Python<'_>, value: f64) -> PyResult<i64> {
        let mut buffer = ryu::Buffer::new();
        let shortest = buffer.format_finite(value).as_bytes();
        let short = shortest.strip_suffix(b".0").unwrap_or(shortest);
        if value != 0.0 && short.len() - (value < 0.0) as usize <= 11 && !short.contains(&b'e') {
            // A short shortest form is what General shows (1.5, -42, 0.25).
            self.room_pixels(py, short)
        } else {
            self.room_pixels(py, general(Dec::from_shortest(shortest)).as_bytes())
        }
    }

    /// The bounds that let a cell be skipped without measuring it: a cell
    /// whose widest possible rendering fits the column's width cannot widen it.
    fn set_bounds(&mut self) {
        let px = |cp: u8| -> Option<i64> {
            self.table
                .get(cp as usize)
                .copied()
                .filter(|&p| p >= 0)
                .map(i64::from)
        };
        let digits: Option<Vec<i64>> = (b'0'..=b'9').map(px).collect();
        let (Some(digits), Some(minus), Some(e), Some(plus)) =
            (digits, px(b'-'), px(b'E'), px(b'+'))
        else {
            self.number_max = i64::MAX;
            self.negative_max = i64::MAX;
            self.negative_small_max = i64::MAX;
            self.digit_max = i64::MAX;
            self.minus = i64::MAX;
            self.latin1_max = i64::MAX;
            return;
        };
        let digit = *digits.iter().max().unwrap_or(&0);
        // General: a sign and 11 characters (the point as wide as a digit),
        // or a sign, 7 mantissa characters, E+ and up to 3 exponent digits.
        let widest = (11 * digit).max(9 * digit + e + plus);
        self.number_max = self.column(widest);
        self.negative_small_max = self.column(minus + widest);
        self.negative_max = self.column(minus + (11 * digit).max(10 * digit + e + plus));
        self.digit_max = digit;
        self.minus = minus;
        // A Latin-1 character without a glyph is drawn as wide as a digit, or
        // not at all (`_char_pixels`).
        self.latin1_max = self
            .table
            .iter()
            .take(256)
            .copied()
            .map(i64::from)
            .fold(digits[0], i64::max);
    }

    /// The column AutoFit makes for text `px` wide (`_autofit_pixels`),
    /// the indent included.
    fn column(&self, px: i64) -> i64 {
        let px = px + self.indent;
        (px + self.margin).max(px * 29 / 28 + self.margin - 2) + SPARE_PX
    }
}

// --- Excel's General format ------------------------------------------------
//
// Numbers are formatted on the stack: a heap allocation costs more than the
// formatting itself.

/// ASCII text of at most 48 bytes, on the stack.
struct Buf {
    bytes: [u8; 48],
    len: usize,
}

impl Buf {
    fn new() -> Self {
        Buf {
            bytes: [0; 48],
            len: 0,
        }
    }

    fn push(&mut self, b: u8) {
        self.bytes[self.len] = b;
        self.len += 1;
    }

    fn extend(&mut self, s: &[u8]) {
        self.bytes[self.len..self.len + s.len()].copy_from_slice(s);
        self.len += s.len();
    }

    fn as_bytes(&self) -> &[u8] {
        &self.bytes[..self.len]
    }
}

impl std::fmt::Write for Buf {
    fn write_str(&mut self, s: &str) -> std::fmt::Result {
        if self.len + s.len() > self.bytes.len() {
            return Err(std::fmt::Error);
        }
        self.extend(s.as_bytes());
        Ok(())
    }
}

/// A nonzero decimal `digits` x 10^`exp`: digits without leading or trailing
/// zeros, as Python's `Decimal(repr(value))` holds it.
struct Dec {
    negative: bool,
    digits: [u8; 24],
    n: usize,
    exp: i32,
}

impl Dec {
    fn new(negative: bool, raw: &[u8], mut exp: i32) -> Option<Dec> {
        let start = raw.iter().position(|&d| d != b'0')?;
        let mut end = raw.len();
        while raw[end - 1] == b'0' {
            end -= 1;
            exp += 1;
        }
        let mut digits = [0u8; 24];
        digits[..end - start].copy_from_slice(&raw[start..end]);
        Some(Dec {
            negative,
            digits,
            n: end - start,
            exp,
        })
    }

    /// The shortest decimal that round-trips, as `repr(float)` gives it.
    /// ryu writes it as `1234.5`, `1e-7` or `1.2345e20`.
    fn from_f64(value: f64) -> Option<Dec> {
        Dec::from_shortest(ryu::Buffer::new().format_finite(value).as_bytes())
    }

    /// The decimal ryu wrote for a finite float (sign included).
    fn from_shortest(text: &[u8]) -> Option<Dec> {
        let (negative, text) = match text {
            [b'-', rest @ ..] => (true, rest),
            rest => (false, rest),
        };
        let (mantissa, power) = match text.iter().position(|&b| b == b'e') {
            Some(e) => {
                let (negative, digits) = match &text[e + 1..] {
                    [b'-', rest @ ..] => (true, rest),
                    rest => (false, rest),
                };
                let power = digits.iter().fold(0i32, |p, &d| p * 10 + (d - b'0') as i32);
                (&text[..e], if negative { -power } else { power })
            }
            None => (text, 0),
        };
        // ryu writes at most 24 bytes.
        let mut raw = [0u8; 24];
        let (mut n, mut fraction, mut after_point) = (0usize, 0i32, false);
        for &b in mantissa {
            if b == b'.' {
                after_point = true;
            } else {
                raw[n] = b;
                n += 1;
                fraction += after_point as i32;
            }
        }
        Dec::new(negative, &raw[..n], power - fraction)
    }

    fn from_i64(value: i64) -> Option<Dec> {
        let mut raw = [0u8; 24];
        let mut v = value.unsigned_abs();
        let mut i = raw.len();
        loop {
            i -= 1;
            raw[i] = b'0' + (v % 10) as u8;
            v /= 10;
            if v == 0 {
                break;
            }
        }
        Dec::new(value < 0, &raw[i..], 0)
    }

    fn digits(&self) -> &[u8] {
        &self.digits[..self.n]
    }

    /// Exponent of the leading digit (`Decimal.adjusted()`).
    fn adjusted(&self) -> i32 {
        self.n as i32 - 1 + self.exp
    }
}

/// Digits on the stack that rounding can lengthen by one.
struct Digits {
    d: [u8; 40],
    len: usize,
}

impl Digits {
    fn of(s: &[u8]) -> Self {
        let mut d = [0u8; 40];
        d[..s.len()].copy_from_slice(s);
        Digits { d, len: s.len() }
    }

    fn as_bytes(&self) -> &[u8] {
        &self.d[..self.len]
    }

    /// Add one; a carry out of the first digit makes it one digit longer.
    fn increment(&mut self) {
        for i in (0..self.len).rev() {
            if self.d[i] == b'9' {
                self.d[i] = b'0';
            } else {
                self.d[i] += 1;
                return;
            }
        }
        self.d.copy_within(0..self.len, 1);
        self.d[0] = b'1';
        self.len += 1;
    }

    fn zeros(&mut self, count: usize) {
        self.d[self.len..self.len + count].fill(b'0');
        self.len += count;
    }

    /// Zeros in front until there are `len` digits.
    fn pad_front(&mut self, len: usize) {
        if self.len < len {
            let shift = len - self.len;
            self.d.copy_within(0..self.len, shift);
            self.d[..shift].fill(b'0');
            self.len = len;
        }
    }
}

/// `_scientific`: 6 significant digits rounded half even (Decimal's `.5E`),
/// zeros trimmed, a two-digit exponent at least.
fn scientific(dec: &Dec, out: &mut Buf) {
    use std::fmt::Write;
    let digits = dec.digits();
    let mut power = dec.adjusted();
    let mut kept = Digits::of(&digits[..digits.len().min(6)]);
    if digits.len() > 6 {
        let first = digits[6];
        let rest_nonzero = digits[7..].iter().any(|&d| d != b'0');
        let odd = (kept.d[5] - b'0') % 2 == 1;
        if first > b'5' || (first == b'5' && (rest_nonzero || odd)) {
            kept.increment();
            if kept.len > 6 {
                kept.len = 6;
                power += 1;
            }
        }
    }
    while kept.len > 1 && kept.d[kept.len - 1] == b'0' {
        kept.len -= 1;
    }
    if dec.negative {
        out.push(b'-');
    }
    out.push(kept.d[0]);
    if kept.len > 1 {
        out.push(b'.');
        out.extend(&kept.d[1..kept.len]);
    }
    out.push(b'E');
    out.push(if power < 0 { b'-' } else { b'+' });
    let _ = write!(out, "{:02}", power.unsigned_abs());
}

/// `_general_number`: the text Excel's General format shows for a number.
fn general_text(dec: &Dec, out: &mut Buf) {
    let digits = dec.digits();
    let n = digits.len() as i32;
    if dec.exp >= 0 {
        if n + dec.exp <= 11 {
            if dec.negative {
                out.push(b'-');
            }
            out.extend(digits);
            for _ in 0..dec.exp {
                out.push(b'0');
            }
            return;
        }
        return scientific(dec, out);
    }
    let adjusted = dec.adjusted();
    if (-9..=10).contains(&adjusted) {
        let whole_len = if adjusted >= 0 { adjusted + 1 } else { 1 };
        let decimals = (10 - whole_len).max(0);
        // Round to `decimals` places, half up: `rounded` x 10^-decimals.
        let cut = -decimals - dec.exp; // digits dropped
        let (mut rounded, exact) = if cut <= 0 {
            let mut d = Digits::of(digits);
            d.zeros((-cut) as usize);
            (d, true)
        } else {
            let keep = n - cut;
            let mut d = if keep > 0 {
                Digits::of(&digits[..keep as usize])
            } else {
                Digits::of(b"0")
            };
            if keep >= 0 && digits[keep as usize] >= b'5' {
                d.increment();
            }
            (d, false)
        };
        let decimals = decimals as usize;
        rounded.pad_front(decimals + 1);
        let all = rounded.as_bytes();
        let split = all.len() - decimals;
        let whole = &all[..split];
        let whole = match whole.iter().position(|&d| d != b'0') {
            Some(i) => &whole[i..],
            None => b"0".as_slice(),
        };
        let mut fraction = &all[split..];
        while let [rest @ .., b'0'] = fraction {
            fraction = rest;
        }
        let room = if adjusted >= 0 {
            decimals as i32 + whole_len
        } else {
            decimals as i32 - (-adjusted - 1)
        };
        let zero = whole == b"0" && fraction.is_empty();
        if !zero && (exact || room >= 6) {
            if dec.negative {
                out.push(b'-');
            }
            out.extend(whole);
            if !fraction.is_empty() {
                out.push(b'.');
                out.extend(fraction);
            }
            return;
        }
    }
    scientific(dec, out)
}

/// General's text of a number (zero for None).
fn general(dec: Option<Dec>) -> Buf {
    let mut out = Buf::new();
    match dec {
        None => out.push(b'0'),
        Some(dec) => general_text(&dec, &mut out),
    }
    out
}

/// `_shown(value, None)[0]` of an int or a finite float, for the tests; None
/// for anything else.
#[pyfunction]
pub fn general_number_text(value: &Bound<'_, PyAny>) -> Option<String> {
    let dec = if value.is_exact_instance_of::<PyFloat>() {
        let v: f64 = value.extract().ok()?;
        if !v.is_finite() {
            return None;
        }
        Dec::from_f64(v)
    } else if value.is_exact_instance_of::<PyInt>() {
        Dec::from_i64(int_as_i64(value)?)
    } else {
        return None;
    };
    Some(String::from_utf8_lossy(general(dec).as_bytes()).into_owned())
}

// --- the meter ---------------------------------------------------------------

/// A column the row's styles leave out (kind 0, or 1 in the header row).
const NO_STYLE: usize = usize::MAX;
/// The kind of an unstyled header cell: Calibri bold (`_cell_pixels`).
const HEADER_KIND: usize = 1;

/// What the cells of one row share.
#[derive(Clone, Copy)]
struct RowContext<'a, 'py> {
    styles: Option<&'a Bound<'py, PyDict>>,
    /// The first row of the sheet: unstyled cells are bold.
    header: bool,
    /// Measure every cell, without skipping those that cannot widen the column.
    exact: bool,
    /// The cell a measured value came from (a formula, link or overlay
    /// tuple): what the Python fallback is given.
    outer: Option<&'a Bound<'py, PyAny>>,
}

/// `(sections, positive, negative or None)` from `_format_plan`; each section
/// `(tag, a, b, c, d, e, f, g)`: 0 General, 1 constant (a: px), 2 text (a:
/// px, b: count), 3 number (head, tail, decimals, thousands, min digits,
/// percent, exponent digits or -1).
fn format_plan(plan: &Bound<'_, PyAny>) -> PyResult<FormatPlan> {
    type Raw = (i64, i64, i64, i64, i64, i64, i64, i64);
    let (sections, positive, negative): (bool, Raw, Option<Raw>) = plan.extract()?;
    let section = |raw: Raw| -> PyResult<Section> {
        let (tag, a, b, c, d, e, f, g) = raw;
        let count =
            |v: i64| usize::try_from(v).map_err(|_| PyValueError::new_err("bad format plan"));
        Ok(match tag {
            0 => Section::General,
            1 => Section::Constant(a),
            2 => Section::Text { px: a, count: b },
            3 => Section::Number {
                head: a,
                tail: b,
                decimals: count(c)?,
                thousands: d != 0,
                min_digits: count(e)?,
                percent: u32::try_from(f).map_err(|_| PyValueError::new_err("bad format plan"))?,
                exponent: if g < 0 { None } else { Some(count(g)?) },
            },
            _ => return Err(PyValueError::new_err("bad format plan")),
        })
    };
    Ok(FormatPlan {
        positive: section(positive)?,
        negative: negative.map(section).transpose()?,
        sections,
    })
}

/// What `_shown` writes for a float that is not finite.
fn non_finite_text(value: f64) -> &'static str {
    if value.is_nan() {
        "NaN"
    } else if value < 0.0 {
        "-∞"
    } else {
        "∞"
    }
}

/// A formula's cached value or a link's text (`_shown`), or None for any other cell.
fn shown_value<'py>(
    py: Python<'py>,
    cell: &Bound<'py, PyAny>,
) -> PyResult<Option<Bound<'py, PyAny>>> {
    if formula_type(py).is_some_and(|t| type_check(cell, &t)) {
        return cell.getattr("value").map(Some);
    }
    if hyperlink_type(py).is_some_and(|t| type_check(cell, &t)) {
        return cell.getattr("display_text").map(Some);
    }
    Ok(None)
}

/// Column widths in pixels of the rows Rust measures.
///
/// `WidthMeter(fallback, kind_of, constant_types)`: `fallback(cell, style, header) -> int`
/// is the column width the Python meter gives a cell this module does not
/// measure (0: none); cells whose exact type is in `constant_types` (bool,
/// date, datetime) are measured once per kind through it.
#[pyclass(module = "xlspy._rs_core", unsendable)]
pub struct WidthMeter {
    kinds: Vec<Option<Kind>>,
    best: Vec<i64>,
    fallback: Py<PyAny>,
    constant_types: Vec<Py<PyType>>,
    constants: HashMap<(usize, usize), i64>,
    /// `kind_of(style) -> (id, spec or None)` for a Style not seen yet.
    kind_of: Py<PyAny>,
    style_kinds: HashMap<usize, (Py<PyAny>, usize)>,
    /// The last row-styles dict and the kind of each column it gave.
    plan_dict: Option<Py<PyAny>>,
    plan: Vec<usize>,
    /// `overlay_kind(style, overlay, header) -> (id, spec or None)`: the kind
    /// of a `(value, format)` or `(value, Style)` cell under a column style.
    overlay_kind: Py<PyAny>,
    /// (column style, overlay, header) -> kind; the objects are kept alive.
    overlay_kinds: HashMap<(usize, usize, bool), OverlayKind>,
}

/// A remembered overlay kind and the objects its key points to.
struct OverlayKind {
    _style: Option<Py<PyAny>>,
    _overlay: Py<PyAny>,
    id: usize,
}

/// Overlay pairs remembered at most: a sheet that builds a new format string
/// or Style for every cell gets the callback each time instead.
const OVERLAY_CACHE: usize = 4096;

impl WidthMeter {
    fn register(&mut self, py: Python<'_>, id: usize, spec: &Bound<'_, PyAny>) -> PyResult<()> {
        if self.kinds.len() <= id {
            self.kinds.resize_with(id + 1, || None);
        }
        if self.kinds[id].is_some() || spec.is_none() {
            return Ok(());
        }
        let (advances, margin, general, wrap, char_px, format, indent): KindSpec<'_> =
            spec.extract()?;
        let format = match format {
            Some(plan) if !general => Some(format_plan(&plan)?),
            _ => None,
        };
        let mut table = Vec::with_capacity(0x3000);
        for (key, value) in advances.iter() {
            let Ok(ch) = key.extract::<char>() else {
                return Err(PyValueError::new_err(
                    "advances keys must be single characters",
                ));
            };
            let cp = ch as usize;
            if table.len() <= cp {
                table.resize(cp + 1, -1);
            }
            table[cp] = value.extract()?;
        }
        let mut kind = Kind {
            table,
            extra: HashMap::new(),
            char_px,
            margin,
            indent,
            general,
            wrap,
            number_max: i64::MAX,
            negative_max: i64::MAX,
            negative_small_max: i64::MAX,
            digit_max: i64::MAX,
            minus: i64::MAX,
            latin1_max: i64::MAX,
            format,
        };
        kind.set_bounds();
        self.kinds[id] = Some(kind);
        let _ = py;
        Ok(())
    }

    fn style_kind(&mut self, py: Python<'_>, style: &Bound<'_, PyAny>) -> PyResult<usize> {
        let key = style.as_ptr() as usize;
        if let Some((_, id)) = self.style_kinds.get(&key) {
            return Ok(*id);
        }
        let answer = self.kind_of.bind(py).call1((style,))?;
        let (id, spec): (usize, Bound<'_, PyAny>) = answer.extract()?;
        self.register(py, id, &spec)?;
        // The Style is kept alive, so its address cannot be reused.
        self.style_kinds.insert(key, (style.clone().unbind(), id));
        Ok(id)
    }

    /// The kind of each column of a row styled by `styles` (a {column: Style} dict).
    fn plan_for(
        &mut self,
        py: Python<'_>,
        styles: &Bound<'_, PyDict>,
        length: usize,
    ) -> PyResult<()> {
        let same = self
            .plan_dict
            .as_ref()
            .is_some_and(|d| d.as_ptr() == styles.as_ptr());
        if same && self.plan.len() >= length {
            return Ok(());
        }
        let mut plan = vec![NO_STYLE; length];
        for (column, style) in styles.iter() {
            let Ok(column) = column.extract::<usize>() else {
                continue;
            };
            if column < length && !style.is_none() {
                plan[column] = self.style_kind(py, &style)?;
            }
        }
        self.plan = plan;
        self.plan_dict = Some(styles.clone().into_any().unbind());
        Ok(())
    }
}

#[pymethods]
impl WidthMeter {
    #[new]
    #[pyo3(signature = (fallback, kind_of, constant_types, overlay_kind=None))]
    fn new(
        py: Python<'_>,
        fallback: Py<PyAny>,
        kind_of: Py<PyAny>,
        constant_types: Vec<Py<PyType>>,
        overlay_kind: Option<Py<PyAny>>,
    ) -> Self {
        WidthMeter {
            kinds: Vec::new(),
            best: Vec::new(),
            fallback,
            constant_types,
            constants: HashMap::new(),
            kind_of,
            style_kinds: HashMap::new(),
            plan_dict: None,
            plan: Vec::new(),
            overlay_kind: overlay_kind.unwrap_or_else(|| py.None()),
            overlay_kinds: HashMap::new(),
        }
    }

    /// Give kind `id` its spec: `(advances, margin, general, wrap, char_px, plan, indent)`.
    fn add_kind(&mut self, py: Python<'_>, id: usize, spec: &Bound<'_, PyAny>) -> PyResult<()> {
        self.register(py, id, spec)
    }

    /// Measure a row; `styles` is its {column: Style} dict (None or empty:
    /// kind 0). `header`: the sheet's first row, where unstyled cells take
    /// kind 1 (bold). `keep`: return the width of each cell (0: none) as well.
    #[pyo3(signature = (row, styles=None, header=false, keep=false))]
    fn add(
        &mut self,
        py: Python<'_>,
        row: &Bound<'_, PyAny>,
        styles: Option<&Bound<'_, PyDict>>,
        header: bool,
        keep: bool,
    ) -> PyResult<Option<Vec<i64>>> {
        let row = FastSeq::new(row, c"row must be a sequence")?;
        let length = row.len();
        let styles = styles.filter(|s| !s.is_empty());
        if let Some(s) = styles {
            self.plan_for(py, s, length)?;
        }
        if self.best.len() < length {
            self.best.resize(length, 0);
        }
        let registered = |id: usize| self.kinds.get(id).is_some_and(Option::is_some);
        if !registered(0) || (header && !registered(HEADER_KIND)) {
            return Err(PyValueError::new_err("kinds 0 and 1 must be registered"));
        }
        let unstyled = if header { HEADER_KIND } else { 0 };
        let context = RowContext {
            styles,
            header,
            exact: keep,
            outer: None,
        };
        let mut kept = if keep { vec![0; length] } else { Vec::new() };
        for column in 0..length {
            let Some(cell) = row.get(column) else { break };
            if cell.is_none() {
                continue;
            }
            let id = match styles {
                Some(_) if self.plan[column] != NO_STYLE => self.plan[column],
                _ => unstyled,
            };
            let width = self.measure(py, id, column, &cell, context)?;
            if width > self.best[column] {
                self.best[column] = width;
            }
            if let Some(slot) = kept.get_mut(column) {
                *slot = width;
            }
        }
        Ok(keep.then_some(kept))
    }

    /// Column widths in pixels (0: nothing measured).
    fn best(&self) -> Vec<i64> {
        self.best.clone()
    }

    // The callbacks are bound methods of the Python meter that owns this one:
    // the garbage collector must see them to free the pair.
    fn __traverse__(&self, visit: PyVisit<'_>) -> Result<(), PyTraverseError> {
        visit.call(&self.fallback)?;
        visit.call(&self.kind_of)?;
        visit.call(&self.overlay_kind)?;
        for t in &self.constant_types {
            visit.call(t)?;
        }
        for kind in self.kinds.iter().flatten() {
            visit.call(&kind.char_px)?;
        }
        for (style, _) in self.style_kinds.values() {
            visit.call(style)?;
        }
        if let Some(plan) = &self.plan_dict {
            visit.call(plan)?;
        }
        for known in self.overlay_kinds.values() {
            if let Some(style) = &known._style {
                visit.call(style)?;
            }
            visit.call(&known._overlay)?;
        }
        Ok(())
    }

    fn __clear__(&mut self) {
        Python::attach(|py| {
            self.fallback = py.None();
            self.kind_of = py.None();
            self.overlay_kind = py.None();
        });
        self.constant_types.clear();
        self.kinds.clear();
        self.style_kinds.clear();
        self.plan_dict = None;
        self.plan.clear();
        self.overlay_kinds.clear();
    }
}

impl WidthMeter {
    fn measure(
        &mut self,
        py: Python<'_>,
        id: usize,
        column: usize,
        cell: &Bound<'_, PyAny>,
        context: RowContext<'_, '_>,
    ) -> PyResult<i64> {
        let best = if context.exact {
            0
        } else {
            self.best.get(column).copied().unwrap_or(0)
        };
        let Some(kind) = self.kinds.get_mut(id).and_then(Option::as_mut) else {
            return self.call_fallback(py, column, cell, context);
        };
        // Exact types only, as the Python meter checks them (a str subclass
        // such as CellError, or a bool, goes the other way).
        // SAFETY: `cell` is a live object, so its type pointer is valid; the
        // type objects are statics of the interpreter, only their addresses taken.
        let (type_ptr, str_type, float_type, int_type) = unsafe {
            (
                ffi::Py_TYPE(cell.as_ptr()),
                &raw mut ffi::PyUnicode_Type,
                &raw mut ffi::PyFloat_Type,
                &raw mut ffi::PyLong_Type,
            )
        };
        // Numbers in General cannot need more than `negative_max`, positive
        // ones more than `number_max`.
        let number_fits = kind.general && best >= kind.negative_max;
        let text_px = if type_ptr == str_type {
            // SAFETY: an exact str, checked just above.
            let text = unsafe { cell.cast_unchecked::<PyString>() };
            // SAFETY: an exact str; PyUnicode_KIND and _GET_LENGTH only read its header.
            let (one_byte, length) = unsafe {
                (
                    ffi::PyUnicode_KIND(cell.as_ptr()) == ffi::PyUnicode_1BYTE_KIND,
                    ffi::PyUnicode_GET_LENGTH(cell.as_ptr()) as i64,
                )
            };
            if one_byte
                && kind.latin1_max != i64::MAX
                && kind.column(length * kind.latin1_max) <= best
            {
                return Ok(0);
            }
            kind.text_pixels(py, text)?
        } else if number_fits && type_ptr == int_type {
            return Ok(0);
        } else if type_ptr == float_type {
            // SAFETY: an exact float, checked just above.
            let value = unsafe { ffi::PyFloat_AS_DOUBLE(cell.as_ptr()) };
            if !value.is_finite() {
                // Any number format shows them as text (`_shown`).
                kind.literal_pixels(py, non_finite_text(value))?
            } else if !kind.general {
                match kind.formatted_pixels(py, Num::Float(value))? {
                    Some(px) => px,
                    None => return self.call_fallback(py, column, cell, context),
                }
            } else if kind.float_fits(value, best) {
                return Ok(0);
            } else {
                kind.float_room_pixels(py, value)?
            }
        } else if (kind.general || kind.format.is_some())
            && decimal_type(py).is_some_and(|d| type_ptr == d.as_type_ptr())
        {
            // A Decimal is shown as the float it converts to (`_shown`), which
            // is float(str(decimal)), correctly rounded like Rust's parse.
            let value = cell
                .str()
                .ok()
                .and_then(|text| text.to_str().ok().and_then(|t| t.parse::<f64>().ok()));
            let Some(value) = value else {
                return self.call_fallback(py, column, cell, context);
            };
            if !value.is_finite() {
                kind.literal_pixels(py, non_finite_text(value))?
            } else if !kind.general {
                match kind.formatted_pixels(py, Num::Float(value))? {
                    Some(px) => px,
                    None => return self.call_fallback(py, column, cell, context),
                }
            } else if kind.float_fits(value, best) {
                return Ok(0);
            } else {
                kind.float_room_pixels(py, value)?
            }
        } else if kind.format.is_some() && type_ptr == int_type {
            let Some(value) = int_as_i64(cell) else {
                return self.call_fallback(py, column, cell, context);
            };
            match kind.formatted_pixels(py, Num::Int(value))? {
                Some(px) => px,
                None => return self.call_fallback(py, column, cell, context),
            }
        } else if kind.general && type_ptr == int_type {
            let Some(value) = int_as_i64(cell) else {
                return self.call_fallback(py, column, cell, context);
            };
            let digits = value.unsigned_abs().checked_ilog10().unwrap_or(0) as i64 + 1;
            let bound = if digits > 11 {
                // An i64 has at most 19 digits: a two-digit exponent.
                if value < 0 {
                    kind.negative_small_max
                } else {
                    kind.number_max
                }
            } else if kind.digit_max == i64::MAX {
                i64::MAX
            } else {
                kind.column((value < 0) as i64 * kind.minus + digits * kind.digit_max)
            };
            if bound <= best {
                return Ok(0);
            }
            kind.room_pixels(py, general(Dec::from_i64(value)).as_bytes())?
        } else {
            if type_ptr == (&raw mut ffi::PyTuple_Type) {
                return self.measure_overlay(py, column, cell, context);
            }
            if let Some(value) = shown_value(py, cell)? {
                // A formula's cached value or a link's text, measured as such;
                // a tuple there is shown as str(tuple), which Python measures.
                if value.is_none() {
                    return Ok(0);
                }
                if value.is_instance_of::<PyTuple>() {
                    return self.call_fallback(py, column, cell, context);
                }
                let inner = RowContext {
                    outer: context.outer.or(Some(cell)),
                    ..context
                };
                return self.measure(py, id, column, &value, inner);
            }
            let type_ptr = type_ptr.cast::<ffi::PyObject>();
            let Some(index) = self
                .constant_types
                .iter()
                .position(|t| t.as_ptr() == type_ptr)
            else {
                return self.call_fallback(py, column, cell, context);
            };
            if let Some(&width) = self.constants.get(&(id, index)) {
                return Ok(width);
            }
            let width = self.call_fallback(py, column, cell, context)?;
            self.constants.insert((id, index), width);
            return Ok(width);
        };
        Ok(if text_px > 0 { kind.column(text_px) } else { 0 })
    }

    /// A `(value, format)` or `(value, Style)` cell: its value in the kind
    /// the overlay makes of the column's style.
    fn measure_overlay(
        &mut self,
        py: Python<'_>,
        column: usize,
        cell: &Bound<'_, PyAny>,
        context: RowContext<'_, '_>,
    ) -> PyResult<i64> {
        let style_cls = style_type(py);
        if context.outer.is_some()
            || self.overlay_kind.is_none(py)
            || fmt_tuple_kind(cell, style_cls.as_ref()) == 0
        {
            return self.call_fallback(py, column, cell, context);
        }
        // SAFETY: an exact tuple of two items, checked by the caller and fmt_tuple_kind.
        let tuple = unsafe { cell.cast_unchecked::<PyTuple>() };
        let (value, overlay) = (tuple.get_item(0)?, tuple.get_item(1)?);
        if value.is_none() {
            return Ok(0);
        }
        let style = match context.styles {
            Some(styles) => styles.get_item(column)?.filter(|s| !s.is_none()),
            None => None,
        };
        let key = (
            style.as_ref().map_or(0, |s| s.as_ptr() as usize),
            overlay.as_ptr() as usize,
            context.header,
        );
        let id = match self.overlay_kinds.get(&key) {
            Some(known) => known.id,
            None => {
                let answer =
                    self.overlay_kind
                        .bind(py)
                        .call1((style.as_ref(), &overlay, context.header))?;
                let (id, spec): (usize, Bound<'_, PyAny>) = answer.extract()?;
                self.register(py, id, &spec)?;
                if self.overlay_kinds.len() < OVERLAY_CACHE {
                    let known = OverlayKind {
                        _style: style.map(Bound::unbind),
                        _overlay: overlay.unbind(),
                        id,
                    };
                    self.overlay_kinds.insert(key, known);
                }
                id
            }
        };
        // Whatever the value needs from Python is measured on the whole cell,
        // so that Python applies the overlay as well.
        let inner = RowContext {
            outer: Some(cell),
            ..context
        };
        self.measure(py, id, column, &value, inner)
    }

    /// `fallback(cell, style, header)`: the Python meter's width of a cell
    /// (the outer cell when a value inside a formula, link or overlay needs it).
    fn call_fallback(
        &self,
        py: Python<'_>,
        column: usize,
        cell: &Bound<'_, PyAny>,
        context: RowContext<'_, '_>,
    ) -> PyResult<i64> {
        let cell = context.outer.unwrap_or(cell);
        let style = match context.styles {
            Some(styles) => styles.get_item(column)?,
            None => None,
        };
        self.fallback
            .bind(py)
            .call1((cell, style, context.header))?
            .extract()
    }
}
