//! Little-endian readers, BIFF12 varints and RK numbers, plus a borrowed
//! view of a Python bytes-like object (the C extension's `y*` arguments).
//!
//! The C code stored these values in `long`, which is 32 bits on Windows;
//! here widths are explicit and match the 64-bit-`long` platforms.

use pyo3::exceptions::PyBufferError;
use pyo3::ffi;
use pyo3::prelude::*;

/// A C-contiguous byte buffer (`PyBUF_SIMPLE`), released on drop.
pub struct ByteView {
    view: ffi::Py_buffer,
}

impl ByteView {
    pub fn new(obj: &Bound<'_, PyAny>) -> PyResult<Self> {
        // SAFETY: Py_buffer is a plain C struct of pointers and integers, for
        // which all-zero bytes are a valid (empty) value.
        let mut view: ffi::Py_buffer = unsafe { std::mem::zeroed() };
        // SAFETY: `obj` is live with the GIL held and `view` is writable; on
        // success the exporter fills `view` and must later get PyBuffer_Release.
        if unsafe { ffi::PyObject_GetBuffer(obj.as_ptr(), &mut view, ffi::PyBUF_SIMPLE) } < 0 {
            return Err(PyErr::take(obj.py())
                .unwrap_or_else(|| PyBufferError::new_err("cannot get a buffer")));
        }
        Ok(ByteView { view })
    }

    pub fn as_slice(&self) -> &[u8] {
        if self.view.buf.is_null() || self.view.len <= 0 {
            return &[];
        }
        // SAFETY: a PyBUF_SIMPLE view is `len` contiguous bytes at `buf`, valid
        // until PyBuffer_Release in drop; the slice borrows `self`, so it cannot
        // outlive the view. Null and empty buffers returned early above.
        unsafe { std::slice::from_raw_parts(self.view.buf as *const u8, self.view.len as usize) }
    }
}

// SAFETY: the exporter keeps the buffer valid until release, whatever thread
// reads it; it is only read, and released in drop with the GIL held
// (pyclass deallocation).
unsafe impl Send for ByteView {}
// SAFETY: as for Send: the view is never written through, so shared reads
// from several threads are sound.
unsafe impl Sync for ByteView {}

impl Drop for ByteView {
    fn drop(&mut self) {
        // SAFETY: `view` was filled by a successful PyObject_GetBuffer (new()
        // returns no ByteView otherwise) and is released exactly once, here.
        // Every ByteView lives inside a call or a pyclass that holds the GIL.
        unsafe { ffi::PyBuffer_Release(&mut self.view) };
    }
}

/// BIFF12 variable-length integer (record id or length): at most 4 bytes of
/// 7 bits each, plus a fifth byte as the C reader allowed (so up to 35
/// bits).  `None` on truncation or a fifth continuation byte.
#[inline]
pub fn read_varint(data: &[u8], pos: &mut usize) -> Option<u64> {
    if *pos >= data.len() {
        return None;
    }
    let mut value: u64 = 0;
    let mut shift = 0;
    loop {
        let b = data[*pos];
        *pos += 1;
        value |= u64::from(b & 0x7F) << shift;
        if b & 0x80 == 0 {
            break;
        }
        shift += 7;
        if shift > 28 || *pos >= data.len() {
            return None;
        }
    }
    Some(value)
}

#[inline]
pub fn le_u32(data: &[u8], pos: usize) -> Option<u32> {
    let b = data.get(pos..pos.checked_add(4)?)?;
    Some(u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
}

/// IEEE double at `pos`, or 0.0 when the record is too short.
#[inline]
pub fn le_f64(data: &[u8], pos: usize) -> f64 {
    match pos.checked_add(8).and_then(|end| data.get(pos..end)) {
        Some(b) => f64::from_le_bytes(b.try_into().unwrap()),
        None => 0.0,
    }
}

/// Decoded RK number (MS-XLSB 2.5.122), or 0.0 when the record is too short.
#[inline]
pub fn read_rk(data: &[u8], pos: usize) -> f64 {
    let Some(raw) = le_u32(data, pos) else {
        return 0.0;
    };
    let mut result = if raw & 2 != 0 {
        // 30-bit signed integer in the high bits (arithmetic shift).
        f64::from((raw as i32) >> 2)
    } else {
        f64::from_bits(u64::from(raw & 0xFFFF_FFFC) << 32)
    };
    if raw & 1 != 0 {
        result /= 100.0;
    }
    result
}

#[inline]
pub fn write_varint(out: &mut Vec<u8>, mut value: u32) {
    loop {
        let b = (value & 0x7F) as u8;
        value >>= 7;
        if value != 0 {
            out.push(b | 0x80);
        } else {
            out.push(b);
            break;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn varint_roundtrip() {
        for v in [
            0u32,
            1,
            0x7F,
            0x80,
            0x3FFF,
            0x4000,
            0x1F_FFFF,
            0x20_0000,
            0x0FFF_FFFF,
        ] {
            let mut buf = Vec::new();
            write_varint(&mut buf, v);
            let mut pos = 0;
            assert_eq!(read_varint(&buf, &mut pos), Some(u64::from(v)));
            assert_eq!(pos, buf.len());
        }
    }

    #[test]
    fn varint_truncated_and_overlong() {
        let mut pos = 0;
        assert_eq!(read_varint(&[0x80], &mut pos), None);
        let mut pos = 0;
        assert_eq!(
            read_varint(&[0x80, 0x80, 0x80, 0x80, 0x01], &mut pos),
            Some(1 << 28)
        );
        let mut pos = 0;
        assert_eq!(
            read_varint(&[0x80, 0x80, 0x80, 0x80, 0x80, 0x01], &mut pos),
            None
        );
        let mut pos = 1;
        assert_eq!(read_varint(&[0x01], &mut pos), None);
    }

    #[test]
    fn rk_numbers() {
        // Integer 1, integer -1 (the UBSan case of the C code), integer 1 / 100.
        assert_eq!(read_rk(&(1u32 << 2 | 2).to_le_bytes(), 0), 1.0);
        assert_eq!(read_rk(&((-1i32 << 2) as u32 | 2).to_le_bytes(), 0), -1.0);
        assert_eq!(read_rk(&(1u32 << 2 | 3).to_le_bytes(), 0), 0.01);
        assert_eq!(
            read_rk(&((-536_870_912i32 << 2) as u32 | 2).to_le_bytes(), 0),
            -536_870_912.0
        );
        // Double 1.0: high 30 bits of its IEEE representation.
        let raw = (1.0f64.to_bits() >> 32) as u32;
        assert_eq!(read_rk(&raw.to_le_bytes(), 0), 1.0);
        assert_eq!(read_rk(&[1, 2, 3], 0), 0.0);
    }

    #[test]
    fn le_readers_bounds() {
        assert_eq!(le_u32(&[1, 0, 0, 0], 0), Some(1));
        assert_eq!(le_u32(&[1, 0, 0], 0), None);
        assert_eq!(le_u32(&[0xFF; 4], 0), Some(u32::MAX));
        assert_eq!(le_f64(&2.5f64.to_le_bytes(), 0), 2.5);
        assert_eq!(le_f64(&[0; 7], 0), 0.0);
    }
}
