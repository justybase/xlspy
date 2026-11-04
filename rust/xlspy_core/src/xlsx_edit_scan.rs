//! The scan of a worksheet part for `XlsxUpdater.update_cells` (_xlsx_edit.py).
//!
//! One pass over `<sheetData>` finds what an edit has to know: the bounds of all
//! cells, the cells being edited, and the formulas that forbid an edit. It reads only
//! the plain markup that Excel and the writers produce (unprefixed names,
//! double-quoted attributes, references in their normal form). Anything else returns
//! None and the caller reads the element tree, which gives the same answers
//! (tests/test_xlsx_edit_scan.py compares the two).

use pyo3::prelude::*;
use std::collections::{HashMap, HashSet};

const MAX_ROWS: i64 = 1_048_576;
const MAX_COLUMNS: i64 = 16_384;
/// The key of a shared formula without an `si` attribute (a byte no XML text holds).
const NO_SI: &[u8] = b"\xff";

fn find(hay: &[u8], needle: &[u8], from: usize) -> Option<usize> {
    memchr::memmem::find(hay.get(from..)?, needle).map(|at| at + from)
}

fn name_byte(c: u8) -> bool {
    c.is_ascii_alphanumeric() || matches!(c, b'_' | b':' | b'-' | b'.')
}

fn space(c: u8) -> bool {
    matches!(c, b' ' | b'\t' | b'\n' | b'\r')
}

/// Reads the attributes at `*i`, up to the end of the tag, and calls `each(name,
/// value)` for each. Returns whether the tag closes itself, with `*i` after the `>`;
/// None for anything unusual (single quotes, no space between attributes) or when
/// `each` says so.
fn attributes<'a>(
    b: &'a [u8],
    i: &mut usize,
    mut each: impl FnMut(&'a [u8], &'a [u8]) -> bool,
) -> Option<bool> {
    loop {
        let start = *i;
        while *i < b.len() && space(b[*i]) {
            *i += 1;
        }
        match *b.get(*i)? {
            b'>' => {
                *i += 1;
                return Some(false);
            }
            b'/' => {
                if b.get(*i + 1) != Some(&b'>') {
                    return None;
                }
                *i += 2;
                return Some(true);
            }
            _ => {}
        }
        if *i == start {
            return None;
        }
        let name_start = *i;
        while *i < b.len() && name_byte(b[*i]) {
            *i += 1;
        }
        if *i == name_start || b.get(*i) != Some(&b'=') || b.get(*i + 1) != Some(&b'"') {
            return None;
        }
        let name = &b[name_start..*i];
        let value_start = *i + 2;
        let close = memchr::memchr(b'"', &b[value_start..])? + value_start;
        *i = close + 1;
        if !each(name, &b[value_start..close]) {
            return None;
        }
    }
}

/// The zero-based column and row of a reference in its normal form (`[A-Z]{1,3}[1-9][0-9]{0,6}`).
fn plain_ref(r: &[u8]) -> Option<(i64, i64)> {
    let letters = r.iter().take_while(|c| c.is_ascii_uppercase()).count();
    if !(1..=3).contains(&letters) {
        return None;
    }
    let digits = &r[letters..];
    if digits.is_empty()
        || digits.len() > 7
        || digits[0] == b'0'
        || !digits.iter().all(u8::is_ascii_digit)
    {
        return None;
    }
    let row = digits
        .iter()
        .fold(0i64, |v, &d| v * 10 + i64::from(d - b'0'));
    let column = r[..letters]
        .iter()
        .fold(0i64, |k, &c| k * 26 + i64::from(c - b'A' + 1));
    if row > MAX_ROWS || column > MAX_COLUMNS {
        return None;
    }
    Some((column - 1, row - 1))
}

/// The row number of a `<row r="...">`: digits, in the sheet.
fn row_number(value: &[u8]) -> Option<i64> {
    if value.is_empty() || value.len() > 7 || !value.iter().all(u8::is_ascii_digit) {
        return None;
    }
    let n = value
        .iter()
        .fold(0i64, |v, &d| v * 10 + i64::from(d - b'0'));
    (1..=MAX_ROWS).contains(&n).then_some(n)
}

fn reference(column: i64, row: i64) -> Vec<u8> {
    let mut letters = Vec::with_capacity(3);
    let mut n = column + 1;
    while n > 0 {
        n -= 1;
        letters.push(b'A' + (n % 26) as u8);
        n /= 26;
    }
    letters.reverse();
    letters.extend_from_slice(row.to_string().as_bytes());
    letters
}

/// A cell being edited: its reference, `t` attribute and whether it has a formula.
pub type Edited = (String, Option<String>, bool);

/// What the scan found, see `scan`.
#[derive(Debug, Default, PartialEq)]
pub struct Found {
    /// `(first column, first row, last column, last row)`, zero-based, of all cells.
    pub bounds: Option<(i64, i64, i64, i64)>,
    pub edited: Vec<Edited>,
    /// The index (in document order, among all `<f>`) of the first shared formula
    /// that shares its group with a formula of an edited cell.
    pub shared_hit: Option<usize>,
    /// `(ref, index)` of each array or data-table formula.
    pub arrays: Vec<(String, usize)>,
}

fn text(b: &[u8]) -> String {
    String::from_utf8_lossy(b).into_owned()
}

/// Scans `<sheetData>...</sheetData>`; `targets` are the references being edited.
pub fn scan(body: &[u8], targets: &HashSet<&[u8]>) -> Option<Found> {
    let wanted_rows: HashSet<i64> = targets
        .iter()
        .filter_map(|t| {
            let letters = t.iter().take_while(|c| c.is_ascii_uppercase()).count();
            std::str::from_utf8(&t[letters..]).ok()?.parse().ok()
        })
        .collect();
    let mut found = Found::default();
    let mut bounds: Option<(i64, i64, i64, i64)> = None;
    let mut f_index = 0usize;
    let mut first_by_group: HashMap<&[u8], usize> = HashMap::new();
    let mut edited_groups: Vec<&[u8]> = Vec::new();

    if !body.starts_with(b"<sheetData") {
        return None;
    }
    let mut pos = b"<sheetData".len();
    if !body
        .get(pos)
        .is_some_and(|&c| c == b'>' || c == b'/' || space(c))
    {
        return None;
    }
    if attributes(body, &mut pos, |_, _| true)? {
        return (pos == body.len()).then_some(found);
    }
    let mut next_row = 1i64;
    loop {
        while pos < body.len() && space(body[pos]) {
            pos += 1;
        }
        if body[pos..].starts_with(b"</sheetData>") {
            if pos + b"</sheetData>".len() != body.len() {
                return None;
            }
            break;
        }
        if !body[pos..].starts_with(b"<row")
            || !body
                .get(pos + 4)
                .is_some_and(|&c| c == b'>' || c == b'/' || space(c))
        {
            return None;
        }
        pos += 4;
        let mut row_attr: Option<&[u8]> = None;
        let closed = attributes(body, &mut pos, |name, value| {
            if name == b"r" {
                row_attr = Some(value);
            }
            true
        })?;
        let row = match row_attr {
            Some(value) => row_number(value)?,
            None => next_row,
        };
        if !(1..=MAX_ROWS).contains(&row) {
            return None;
        }
        next_row = row + 1;
        if closed {
            continue;
        }
        let mut next_column = 0i64;
        loop {
            while pos < body.len() && space(body[pos]) {
                pos += 1;
            }
            if body[pos..].starts_with(b"</row>") {
                pos += b"</row>".len();
                break;
            }
            if !body[pos..].starts_with(b"<c")
                || !body
                    .get(pos + 2)
                    .is_some_and(|&c| c == b'>' || c == b'/' || space(c))
            {
                return None;
            }
            pos += 2;
            let (mut r, mut t): (Option<&[u8]>, Option<&[u8]>) = (None, None);
            let closed = attributes(body, &mut pos, |name, value| {
                if matches!(name, b"r" | b"t") && value.contains(&b'&') {
                    return false;
                }
                match name {
                    b"r" => r = Some(value),
                    b"t" => t = Some(value),
                    _ => {}
                }
                true
            })?;
            let (column, cell_row) = match r {
                Some(r) => plain_ref(r)?,
                None => {
                    if next_column >= MAX_COLUMNS {
                        return None;
                    }
                    (next_column, row - 1)
                }
            };
            next_column = column + 1;
            let mut has_f = false;
            let mut groups_of_this: Vec<&[u8]> = Vec::new();
            if !closed {
                let close = find(body, b"</c>", pos)?;
                let content = &body[pos..close];
                let mut p = 0;
                while let Some(offset) = memchr::memchr(b'<', &content[p..]) {
                    let at = p + offset;
                    let after = content.get(at + 2).copied();
                    let plain = !after.is_some_and(name_byte);
                    match content.get(at + 1) {
                        Some(b'!' | b'?') => return None,
                        Some(b'c') if plain => return None,
                        Some(b'f') if plain => {
                            let mut i = at + 2;
                            let mut ft: Option<&[u8]> = None;
                            let mut si: Option<&[u8]> = None;
                            let mut fref: Option<&[u8]> = None;
                            attributes(content, &mut i, |name, value| {
                                if matches!(name, b"t" | b"si" | b"ref") && value.contains(&b'&') {
                                    return false;
                                }
                                match name {
                                    b"t" => ft = Some(value),
                                    b"si" => si = Some(value),
                                    b"ref" => fref = Some(value),
                                    _ => {}
                                }
                                true
                            })?;
                            has_f = true;
                            if ft == Some(b"shared") {
                                let group = si.unwrap_or(NO_SI);
                                first_by_group.entry(group).or_insert(f_index);
                                groups_of_this.push(group);
                            }
                            if let (Some(b"array" | b"dataTable"), Some(fref)) = (ft, fref)
                                && !fref.is_empty()
                            {
                                found.arrays.push((text(fref), f_index));
                            }
                            f_index += 1;
                            p = i;
                            continue;
                        }
                        _ => {}
                    }
                    p = at + 1;
                }
                pos = close + b"</c>".len();
            }
            bounds = Some(match bounds {
                None => (column, cell_row, column, cell_row),
                Some((c1, r1, c2, r2)) => (
                    c1.min(column),
                    r1.min(cell_row),
                    c2.max(column),
                    r2.max(cell_row),
                ),
            });
            let reference_text: Option<Vec<u8>> = match r {
                Some(r) => targets.contains(r).then(|| r.to_vec()),
                None if wanted_rows.contains(&row) => {
                    let made = reference(column, row);
                    targets.contains(made.as_slice()).then_some(made)
                }
                None => None,
            };
            if let Some(reference_text) = reference_text {
                found
                    .edited
                    .push((text(&reference_text), t.map(text), has_f));
                edited_groups.extend(groups_of_this);
            }
        }
    }
    found.bounds = bounds;
    found.shared_hit = edited_groups
        .iter()
        .filter_map(|g| first_by_group.get(g))
        .min()
        .copied();
    Some(found)
}

/// `(bounds, edited cells, first shared-formula hit, arrays)`, or None for markup
/// this scan does not read.
type Scanned = (
    Option<(i64, i64, i64, i64)>,
    Vec<Edited>,
    Option<usize>,
    Vec<(String, usize)>,
);

/// The scan of one `<sheetData>` element, see the module comment.
#[pyfunction]
pub fn xlsx_edit_scan(body: &str, targets: Vec<String>) -> Option<Scanned> {
    let wanted: HashSet<&[u8]> = targets.iter().map(|t| t.as_bytes()).collect();
    scan(body.as_bytes(), &wanted).map(|f| (f.bounds, f.edited, f.shared_hit, f.arrays))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn run(body: &str, targets: &[&str]) -> Option<Found> {
        let set: HashSet<&[u8]> = targets.iter().map(|t| t.as_bytes()).collect();
        scan(body.as_bytes(), &set)
    }

    #[test]
    fn bounds_and_edited_cells_come_from_references_or_positions() {
        let body = r#"<sheetData><row r="2"><c r="B2" t="s"><v>0</v></c><c><v>1</v></c></row><row><c/><c t="b"><f>1</f><v>1</v></c></row></sheetData>"#;
        // Row 2 holds B2 and C2; the next row has no number and is row 3: A3 and B3.
        let found = run(body, &["B2", "B3", "D3", "A9"]).unwrap();
        assert_eq!(found.bounds, Some((0, 1, 2, 2)));
        assert_eq!(
            found.edited,
            vec![
                ("B2".to_string(), Some("s".to_string()), false),
                ("B3".to_string(), Some("b".to_string()), true),
            ]
        );
    }

    #[test]
    fn a_cell_without_a_reference_follows_the_one_before_it() {
        let found = run(
            r#"<sheetData><row><c/><c/><c/></row></sheetData>"#,
            &["C1", "B1"],
        )
        .unwrap();
        assert_eq!(found.bounds, Some((0, 0, 2, 0)));
        let refs: Vec<&str> = found.edited.iter().map(|e| e.0.as_str()).collect();
        assert_eq!(refs, vec!["B1", "C1"]);
    }

    #[test]
    fn an_empty_sheet_has_no_bounds() {
        assert_eq!(run("<sheetData/>", &["A1"]), Some(Found::default()));
        assert_eq!(
            run("<sheetData></sheetData>", &["A1"]),
            Some(Found::default())
        );
    }

    #[test]
    fn shared_and_array_formulas_are_numbered_in_document_order() {
        let body = concat!(
            r#"<sheetData><row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0">B1</f><v>1</v></c>"#,
            r#"<c r="B1"><f t="array" ref="B1:B2">1</f></c></row>"#,
            r#"<row r="2"><c r="A2"><f t="shared" si="0"/><v>2</v></c></row></sheetData>"#
        );
        let found = run(body, &["A2"]).unwrap();
        assert_eq!(found.shared_hit, Some(0));
        assert_eq!(found.arrays, vec![("B1:B2".to_string(), 1)]);
        assert_eq!(run(body, &["C9"]).unwrap().shared_hit, None);
    }

    #[test]
    fn markup_that_is_not_plain_is_left_to_the_tree() {
        for body in [
            "<sheetData><row r='1'><c r=\"A1\"/></row></sheetData>",
            r#"<sheetData><row r="1"><c r="a1"/></row></sheetData>"#,
            r#"<sheetData><row r="1"><c r="A01"/></row></sheetData>"#,
            r#"<sheetData><row r="1"><c r="XFE1"/></row></sheetData>"#,
            r#"<sheetData><row r="0"><c r="A1"/></row></sheetData>"#,
            r#"<sheetData><row r="1"><c r="A1"><!-- x --></c></row></sheetData>"#,
            r#"<sheetData><row r="1"><c r="A1"><![CDATA[<f>]]></c></row></sheetData>"#,
            r#"<sheetData><row r="1"><x:c r="A1"/></row></sheetData>"#,
            r#"<sheetData><row r="1"><c r="A1"/></row></sheetData>junk"#,
            r#"<sheetData><row r="1"><c r="A1" t="&#115;"/></row></sheetData>"#,
        ] {
            assert_eq!(run(body, &["A1"]), None, "{body}");
        }
    }
}
