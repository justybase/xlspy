//! Shared strings of an XLSX workbook (`xl/sharedStrings.xml`), parsed natively.
//!
//! The native parser takes the markup the regex parser in reader.py takes and
//! gives the same strings: the text of every `<t>` in an `<si>` is joined (rich
//! and phonetic runs included), references and line breaks are decoded as the
//! XML parser decodes them, and `_xHHHH_` escapes are applied. It refuses
//! anything else by returning None (CDATA, comments, processing instructions,
//! prefixed names, a declared encoding other than UTF-8, self-closing items or
//! text elements, `<t>` with attributes other than xml:space="preserve", and
//! references XML does not allow). reader.py then runs the regex and expat
//! parsers as before, so their results and their errors do not change.

use memchr::memmem::{self, Finder};
use pyo3::prelude::*;
use pyo3::types::PyList;

use crate::bytes::ByteView;
use crate::xlsx_read::{has_prefix, unescape};
use crate::xml_text;

/// The finders the scan reuses for every item.
struct Finders {
    item_open: Finder<'static>,
    item_close: Finder<'static>,
    text_close: Finder<'static>,
}

/// The shared strings of a sharedStrings part, in order, or None when the
/// part needs the regex or expat parser.
#[pyfunction]
pub fn read_xlsx_shared_strings<'py>(
    py: Python<'py>,
    data: &Bound<'py, PyAny>,
) -> PyResult<Option<Bound<'py, PyList>>> {
    let view = ByteView::new(data)?;
    parse(py, view.as_slice())
}

fn parse<'py>(py: Python<'py>, data: &[u8]) -> PyResult<Option<Bound<'py, PyList>>> {
    let Ok(text) = std::str::from_utf8(data) else {
        return Ok(None);
    };
    let mut body = text;
    if text.starts_with("<?xml") {
        let Some(end) = text.find("?>") else {
            return Ok(None);
        };
        if !declares_utf8(&text[..end]) {
            return Ok(None);
        }
        body = &text[end + 2..];
    }
    let bytes = body.as_bytes();
    if memmem::find(bytes, b"<!").is_some()
        || memmem::find(bytes, b"<?").is_some()
        || has_prefix(bytes)
    {
        return Ok(None);
    }

    let finders = Finders {
        item_open: Finder::new(b"<si"),
        item_close: Finder::new(b"</si>"),
        text_close: Finder::new(b"</t>"),
    };
    let strings = PyList::empty(py);
    let mut pos = 0usize;
    let mut items = 0usize;
    while let Some(off) = finders.item_open.find(&bytes[pos..]) {
        let at = pos + off;
        // "<si>" exactly: attributes, "<si/>" and "<sip..." go to the other parsers.
        if bytes.get(at + 3) != Some(&b'>') {
            return Ok(None);
        }
        let content_start = at + 4;
        let Some(close) = finders.item_close.find(&bytes[content_start..]) else {
            return Ok(None);
        };
        let content = &bytes[content_start..content_start + close];
        let Some(raw) = item_text(&finders, content) else {
            return Ok(None);
        };
        strings.append(unescape(py, &raw)?)?;
        items += 1;
        pos = content_start + close + "</si>".len();
    }
    // A closing </si> without its opening tag is malformed: leave it to the parsers.
    if memmem::find_iter(bytes, b"</si>").count() != items {
        return Ok(None);
    }
    Ok(Some(strings))
}

/// The text of one `<si>` content: the decoded text of its `<t>` elements,
/// joined. None for markup this parser does not take.
fn item_text(finders: &Finders, content: &[u8]) -> Option<Vec<u8>> {
    let mut out = Vec::new();
    let mut pos = 0usize;
    while let Some(lt) = memchr::memchr(b'<', &content[pos..]) {
        let at = pos + lt;
        let rest = &content[at..];
        if rest.starts_with(b"</t>") {
            return None; // A closing tag with no opening tag.
        }
        if rest.starts_with(b"<t") && rest.get(2).is_some_and(|&c| !name_byte(c)) {
            let open = if rest.starts_with(b"<t>") {
                3
            } else if rest.starts_with(b"<t xml:space=\"preserve\">") {
                24
            } else {
                return None; // "<t/>", "<t\n>" and attributes other than xml:space.
            };
            let text_start = at + open;
            let close = finders.text_close.find(&content[text_start..])?;
            let raw = &content[text_start..text_start + close];
            if memchr::memchr(b'<', raw).is_some() {
                return None;
            }
            out.extend_from_slice(&xml_text::decode(raw)?);
            pos = text_start + close + "</t>".len();
        } else {
            // Another element (a run, its properties, a phonetic run): its text
            // is not part of the string, but its tag must be well formed.
            let end = tag_end(rest)?;
            let name = tag_name(&rest[1..]);
            if name == b"si" {
                return None; // A nested item.
            }
            pos = at + end;
        }
    }
    Some(out)
}

/// The name of an element at the start of ``p`` (after "<" or "</").
fn tag_name(p: &[u8]) -> &[u8] {
    let p = p.strip_prefix(b"/").unwrap_or(p);
    let len = p.iter().take_while(|&&c| name_byte(c)).count();
    &p[..len]
}

fn name_byte(c: u8) -> bool {
    c.is_ascii_alphanumeric() || matches!(c, b'_' | b'-' | b'.' | b':')
}

/// The length of the tag starting at ``p[0] == '<'``, through its ">".
/// A ">" inside a quoted attribute value does not end the tag.
fn tag_end(p: &[u8]) -> Option<usize> {
    let mut quote: Option<u8> = None;
    for (i, &c) in p.iter().enumerate().skip(1) {
        match quote {
            Some(q) => {
                if c == q {
                    quote = None;
                } else if c == b'<' {
                    return None;
                }
            }
            None => match c {
                b'"' | b'\'' => quote = Some(c),
                b'>' => return Some(i + 1),
                b'<' => return None,
                _ => {}
            },
        }
    }
    None
}

/// True when an XML declaration names UTF-8 (in any case) or names no encoding.
/// A spelling with spaces around "=" is refused rather than guessed at.
fn declares_utf8(declaration: &str) -> bool {
    let Some(at) = declaration.find("encoding") else {
        return true;
    };
    let Some(rest) = declaration[at + "encoding".len()..].strip_prefix('=') else {
        return false;
    };
    let mut chars = rest.chars();
    let Some(quote @ ('"' | '\'')) = chars.next() else {
        return false;
    };
    chars
        .as_str()
        .split(quote)
        .next()
        .is_some_and(|v| v.eq_ignore_ascii_case("utf-8"))
}

#[cfg(test)]
mod tests {
    use super::{Finder, Finders, declares_utf8, item_text, tag_end};

    #[test]
    fn declarations() {
        assert!(declares_utf8(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes""#
        ));
        assert!(declares_utf8(r#"<?xml version="1.0" encoding='utf-8'"#));
        assert!(declares_utf8(r#"<?xml version="1.0""#));
        assert!(!declares_utf8(r#"<?xml version="1.0" encoding="UTF-16""#));
        assert!(!declares_utf8(r#"<?xml version="1.0" encoding = "utf-8""#));
    }

    #[test]
    fn quoted_angle_brackets_stay_in_the_tag() {
        assert_eq!(tag_end(br#"<rPh sb="0" x="a>b">"#), Some(20));
        assert_eq!(tag_end(br#"<r>"#), Some(3));
        assert_eq!(tag_end(br#"<b x="<"/>"#), None);
    }

    fn text_of(content: &[u8]) -> Option<Vec<u8>> {
        let finders = Finders {
            item_open: Finder::new(b"<si"),
            item_close: Finder::new(b"</si>"),
            text_close: Finder::new(b"</t>"),
        };
        item_text(&finders, content)
    }

    #[test]
    fn item_text_joins_the_runs_of_an_item() {
        assert_eq!(text_of(b"<t>abc</t>").as_deref(), Some(&b"abc"[..]));
        assert_eq!(
            text_of(br#"<r><rPr><b/></rPr><t>bo</t></r><r><t xml:space="preserve"> ld</t></r>"#)
                .as_deref(),
            Some(&b"bo ld"[..])
        );
        assert_eq!(text_of(b"").as_deref(), Some(&b""[..]));
    }

    #[test]
    fn item_text_leaves_what_it_cannot_read_to_the_xml_parsers() {
        for content in [
            &b"</t>"[..],             // a closing tag with no opening tag
            b"<t/>",                  // an empty element
            b"<t\n>x</t>",            // white space in the tag
            br#"<t other="1">x</t>"#, // another attribute
            b"<t>x",                  // never closed
            b"<t>a<b/>c</t>",         // markup in the text
            b"<si><t>x</t></si>",     // a nested item
            b"<r x=\"never closed",   // a broken tag
        ] {
            assert_eq!(
                text_of(content),
                None,
                "{:?}",
                String::from_utf8_lossy(content)
            );
        }
    }
}
