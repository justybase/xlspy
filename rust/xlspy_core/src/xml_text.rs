//! Decoding of XML character data for the native readers, as the XML parser
//! (ElementTree, which the Python paths use) applies it: the predefined entity
//! references, numeric character references, and end-of-line handling.
//!
//! `decode` returns None for anything it does not reproduce exactly: unknown
//! or malformed references, references to characters XML forbids (NUL, other
//! control characters, surrogates, U+FFFE and U+FFFF). The caller then hands
//! the row to ElementTree, which decides and may raise. The input is element
//! content without markup, so `<` never appears in it.

use std::borrow::Cow;

/// Element text as the XML parser reports it: `&amp;`, `&lt;`, `&gt;`,
/// `&quot;`, `&apos;`, `&#N;` and `&#xH;` are replaced. A line break written as
/// CR LF or a lone CR becomes LF, while a carriage return written as a
/// character reference (`&#13;`) stays CR.
pub fn decode(p: &[u8]) -> Option<Cow<'_, [u8]>> {
    if memchr::memchr2(b'&', b'\r', p).is_none() {
        return Some(Cow::Borrowed(p));
    }
    let mut out = Vec::with_capacity(p.len());
    let mut i = 0;
    while i < p.len() {
        match p[i] {
            b'&' => {
                // The longest reference that can be valid is "#x10FFFF;" (8 characters).
                let window = &p[i + 1..];
                let end = window.iter().take(12).position(|&c| c == b';')?;
                let ch = match &window[..end] {
                    b"amp" => '&',
                    b"lt" => '<',
                    b"gt" => '>',
                    b"quot" => '"',
                    b"apos" => '\'',
                    [b'#', b'x', hex @ ..] => char_reference(hex, 16)?,
                    [b'#', dec @ ..] => char_reference(dec, 10)?,
                    _ => return None,
                };
                let mut buf = [0u8; 4];
                out.extend_from_slice(ch.encode_utf8(&mut buf).as_bytes());
                i += end + 2;
            }
            b'\r' => {
                out.push(b'\n');
                i += if p.get(i + 1) == Some(&b'\n') { 2 } else { 1 };
            }
            c => {
                out.push(c);
                i += 1;
            }
        }
    }
    Some(Cow::Owned(out))
}

/// The character of a numeric reference's digits, when XML allows it.
fn char_reference(digits: &[u8], radix: u32) -> Option<char> {
    if digits.is_empty() || digits.len() > 8 {
        return None;
    }
    let mut value: u32 = 0;
    for &d in digits {
        value = value * radix + (d as char).to_digit(radix)?;
    }
    // Char ::= #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]
    match value {
        0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF => {
            char::from_u32(value)
        }
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::decode;

    fn text(p: &[u8]) -> Option<Vec<u8>> {
        decode(p).map(|c| c.into_owned())
    }

    #[test]
    fn plain_text_is_borrowed_unchanged() {
        assert!(matches!(
            decode(b"plain text"),
            Some(std::borrow::Cow::Borrowed(_))
        ));
    }

    #[test]
    fn predefined_entities() {
        assert_eq!(
            text(b"a &amp; b &lt; c &gt; d &quot;e&quot; &apos;f&apos;").unwrap(),
            "a & b < c > d \"e\" 'f'".as_bytes()
        );
    }

    #[test]
    fn numeric_references_in_both_radixes() {
        assert!(text(b"&#X44;").is_none()); // an uppercase X is not XML
        assert_eq!(text(b"&#65;&#x42;&#x0063;").unwrap(), b"ABc");
        assert_eq!(text("&#xe9;&#233;".as_bytes()).unwrap(), "éé".as_bytes());
        assert_eq!(
            text("&#x1F600;".as_bytes()).unwrap(),
            "\u{1F600}".as_bytes()
        );
    }

    #[test]
    fn carriage_returns_are_normalized_unless_referenced() {
        assert_eq!(text(b"a\r\nb\rc").unwrap(), b"a\nb\nc");
        assert_eq!(text(b"x&#13;y").unwrap(), b"x\ry");
        assert_eq!(text(b"x&#xD;\r\n").unwrap(), b"x\r\n");
    }

    #[test]
    fn references_are_not_rescanned() {
        assert_eq!(text(b"&amp;lt;").unwrap(), b"&lt;");
        assert_eq!(text(b"&#38;amp;").unwrap(), b"&amp;");
    }

    #[test]
    fn unusual_input_is_left_to_the_xml_parser() {
        // Not predefined (the parser would need a DTD), no semicolon, no digits,
        // not an XML character, a surrogate, beyond Unicode, too many digits.
        let bad: [&[u8]; 12] = [
            b"&nbsp;",
            b"&amp",
            b"&#;",
            b"&#x;",
            b"&#0;",
            b"&#1;",
            b"&#xD800;",
            b"&#xFFFE;",
            b"&#x110000;",
            b"&#99999999999;",
            b"&#x1g;",
            b"x &",
        ];
        for input in bad {
            assert_eq!(text(input), None, "{:?}", std::str::from_utf8(input));
        }
    }
}
