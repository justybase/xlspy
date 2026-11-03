//! Workbook-model cells (`xlspy.workbook._cell.Cell`) created natively.
//!
//! The loaders fill the `__slots__` of each Cell directly, through the member
//! descriptors the caller passes (`CELL_SLOTS`), exactly as the C extension
//! did; the descriptor layout comes from pyo3-ffi's CPython definitions.

use pyo3::exceptions::PyTypeError;
use pyo3::ffi;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyTuple, PyType};

/// Slot indices used after creation; `CELL_SLOTS` (workbook/_model.py) is
/// _sheet, _row, _column, _value, _formula, _data_type, _style_id, _comment,
/// _hyperlink, _named_style, _dirty.
pub const ROW: usize = 1;
pub const COLUMN: usize = 2;
const COUNT: usize = 11;

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
        "_dirty",
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
