# Changelog

🇬🇧 English · [🇷🇺 Русский](CHANGELOG.ru.md)

<!-- NOTE for maintainers: this file exists in two languages.
     When editing CHANGELOG.md, please update CHANGELOG.ru.md accordingly. -->

All notable changes to this project are documented in this file.

This file starts at version 3.2.0; for earlier history see the
[releases page](https://github.com/aVadim483/fast-excel-reader/releases).

## 4.0.0

### Added

* **Reading of legacy XLS workbooks (Excel 97-2003, BIFF8).** `Excel::open()` now chooses the reader
  from the file signature, so XLSX and XLS are opened the same way; the file extension is never
  consulted. `Excel::openXls()` opens a workbook explicitly and `Excel::isXls()` exposes the test.
  See `docs/21-xls.md`.
  * Values and types, with dates detected through number formats
  * Cell styles: fonts, fills, borders, alignment, number formats and palette colours
  * Formula text, including shared formulas
  * Embedded images
  * Multiple worksheets, hidden and very hidden sheets, merged cells, sheet dimensions
  * Encrypted workbooks and BIFF5/BIFF7 files are rejected with a message saying which
* `AbstractBook` and `AbstractSheet`, holding the format-independent half of the readers: read areas,
  key modes, result-mode flags, the row generator and every `read*` helper. XLSX and XLS share one
  implementation of the whole public reading API.
* `withHeader()` now accepts an optional list of column names: `withHeader(['name', 'birthday'])`. The
  header row is still skipped, but the names come from the list instead of from its values. Names are
  positional - the first name goes to the first column of the read area - so no column letters are
  involved and the same call works on a sheet whose data does not start at `A1`. A shorter list renames
  only the columns it covers. Supported for XLSX, XLS and CSV. Calling it with no argument is
  unchanged. This mirrors the naming of `writeHeader()` in the sibling fast-excel-writer.

### Fixed

* `Sheet::readCellsWithStylesFrom()` returned bare cell values instead of values with styles: it called
  `readCells()` rather than `readCellsWithStyles()`, and passed the style key into a bool parameter.
* `Sheet::readCellsWithStyles($styleKey)` never narrowed the result to the requested property. The key
  was looked up on the nested style, where properties sit inside their group, so `'fill-color'` - the
  example in the method's own docblock - always returned the complete style instead.
* Complete cell styles - `getCompleteStyleByIdx()`, `readCellsWithStyles()` and everything built on
  them - died with `Call to undefined method DOMText::getAttribute()` on workbooks whose `styles.xml`
  is written with indentation, which several writers do.

### Changed

* **Reading XLSX is about 1.5 times faster.** Values, types and peak memory are unchanged.
* The accessors that returned a concrete `Sheet` now return `AbstractSheet`, and the fluent setters on
  the workbook return `AbstractBook`. The objects handed back are unchanged, and a subclass may still
  narrow the return type back, so this only affects explicit type declarations in calling code.

## 3.2.0

### Fixed

Four defects, all in methods that had no test coverage.

* `Sheet::readFirstRowCellsFrom()` always threw a `TypeError`: it forwarded `$columnKeys` into
  `readFirstRowCells(?bool $styleIdxInclude)`, so even a default call failed. Because the result is
  keyed by cell address, column keys cannot apply, and the parameter was removed to match
  `readCellsFrom()`. No working call can break, since every call to the old signature threw.
* Restricting columns while requesting numeric column keys returned only `null`s. With a read area in
  place the row template was keyed by column letter while values were stored under numeric keys, and
  the values were then filtered out. Affected `setReadArea()` and `setReadAreaColumns()` combined with
  `KEYS_COL_ZERO_BASED` or `KEYS_COL_ONE_BASED`, and therefore also `KEYS_ZERO_BASED` and
  `KEYS_ONE_BASED`.
* `Sheet::rewind()` discarded its `$columnKeys` argument although it is documented as an alias of
  `reset()`: the body assigned to the parameter instead of forwarding it.
* `Sheet::firstCol()` ignored the column bounds of the read area, reporting the first cell of the row
  as stored in the file. `firstRow()` was unaffected.

### Added

* A regression suite covering the reading path: characterization snapshots over every `read*` method,
  the full `KEYS_*` matrix, result-mode flags, read areas, styles, dates, metadata and degenerate
  inputs, plus targeted tests for generator semantics, read areas, merged cells and streaming memory
  behaviour. `RESULT_MODE_ROW`, `TRIM_STRINGS`, `TREAT_EMPTY_STRING_AS_EMPTY_CELL` and `KEYS_RELATIVE`
  had no coverage at all before.

### Known limitations

* XLSX files that use namespace-prefixed tags (`<x:row>`, `<x:c>`) read back as empty.
