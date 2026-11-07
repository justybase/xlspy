---
description: "Complete xlspy API reference generated from the public docstrings."
---

# API reference

Generated from the docstrings of every public name (`xlspy.__all__` and
`xlspy.workbook.__all__`). Build it with `zensical serve` or `zensical build`
(`pip install -e .[docs]`); the [tutorial](../learn/index.md) shows how the pieces
fit together.

| Page | Contents |
| :--- | :--- |
| [Writing](writing.md) | New workbooks: the writers, sheet options and DataFrame export. |
| [Reading](reading.md) | Fast value extraction with `ExcelReader`. |
| [Updating](updating.md) | Edits to existing files that keep everything else in the package. |
| [Workbook model](model.md) | `load_workbook` and the openpyxl-like objects it returns. |
| [Styles and formatting](styles.md) | Cell styles, conditional formats, tables and rich text. |
| [Charts and other objects](objects.md) | Charts, formulas, notes, links, images, pivots and printing. |
| [Errors](errors.md) | The exceptions xlspy raises. |

## Conventions

- Signatures show the exact parameter names and types. Keyword-only parameters
  are marked with `*`.
- Every public class and function has a docstring; the examples use the same
  names as the [tutorial](../learn/index.md).
- See [Performance](../performance.md) for the cost of the hot paths, and
  [openpyxl parity](../openpyxl-parity.md) for behavioural differences.
