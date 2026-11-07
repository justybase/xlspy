---
description: "What to read next after the xlspy tutorial: how-to guides, API reference and examples."
---

# Next steps

You have written, styled, charted, read, updated and streamed workbooks. Pick the
next path based on what you need.

<div class="xlspy-cards" markdown>

<div class="xlspy-card" markdown>
### [How-to guides](../guide/index.md)
Choose a format, add images and printing, build pivot tables, protect a workbook.
</div>

<div class="xlspy-card" markdown>
### [API reference](../api/index.md)
Every public class and function, with signatures and examples.
</div>

<div class="xlspy-card" markdown>
### [Examples](https://github.com/justybase/xlspy/blob/main/examples/README.md)
Runnable scripts for reports, DataFrames, dashboards and large exports.
</div>

<div class="xlspy-card" markdown>
### [Performance](../performance.md)
Why XLSB is fast, what the native module accelerates, and how to measure it.
</div>

</div>

## A checklist for production

- [ ] Use `with XlsxWriter(...) as workbook:` so a failure publishes nothing.
- [ ] Pass `use_shared_strings=False` for high-cardinality exports.
- [ ] Stream from a generator or cursor instead of building a list.
- [ ] Keep XLSX for interchange; use XLSB for large internal exports.
- [ ] Test the generated file in Excel or LibreOffice at least once.
- [ ] Read values with `ExcelReader`, not `load_workbook()`, when you do not need
      styles or edits.

## Common questions

**Does xlspy need Excel installed?**
No. The core read and write APIs are pure Python plus an optional Rust module.

**Can I write `.xlsm`?**
You can update an existing `.xlsm` with `XlsmUpdater` and preserve its VBA
project. Creating a new macro-enabled workbook from scratch is not supported.

**Can I read `.xlsb` with other libraries?**
Some tools read XLSB; many do not. If the file is for interchange, write XLSX.

**Why does a formula show `#NAME?` in Excel?**
A function added after Excel 2007 may need the `_xlfn.` prefix. xlspy adds it
for known future functions; an unknown function raises or is written as given.
See [Formulas](../guide/objects.md#formulas).

**Where do I report a bug?**
[GitHub issues](https://github.com/justybase/xlspy/issues). Include the smallest
script and the error.

[Back to the tutorial](index.md){ .md-button } [Read the how-to guides](../guide/index.md){ .md-button .md-button--primary }
