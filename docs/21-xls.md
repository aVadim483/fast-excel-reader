# Reading XLS (Excel 97-2003) in FastExcelReader

[← Back to README](../README.md) | [Documentation index](../README.md#documentation) | [🇷🇺 Русский](ru/21-xls.md)

FastExcelReader reads legacy `.xls` workbooks (BIFF8, written by Excel 97-2003) with the same API used
for XLSX. Values, dates, cell styles, formula text and images all come back in exactly the same shape,
so code written against XLSX works unchanged.

## Table of Contents

1. [Opening a file](#opening-a-file)
2. [The API is the same](#the-api-is-the-same)
3. [Memory and speed](#memory-and-speed)
4. [What is supported](#what-is-supported)
5. [Limitations](#limitations)
6. [Differences from XLSX](#differences-from-xlsx)

## Opening a file

`Excel::open()` picks the reader from the file signature, so nothing special is needed:

```php
use avadim\FastExcelReader\Excel;

$excel = Excel::open(__DIR__ . '/files/report.xls');
$rows = $excel->readRows();
```

The file **extension is never consulted** — only the first bytes of the file. That matters in practice,
because `.xls` is a very common extension for files that are actually XLSX, HTML or CSV.

If you want to be explicit, or to reject anything that is not a legacy workbook:

```php
$excel = Excel::openXls($file);      // throws unless the file really is XLS

if (Excel::isXls($file)) {
    // ...
}
```

## The API is the same

Everything documented in the other guides applies to XLS as well:

```php
$excel = Excel::open('report.xls');

$excel->readRows(true, Excel::KEYS_ONE_BASED);
$excel->sheet()->setReadArea('B2:D100')->readRows();
$excel->selectSheet('Summary')->readCells();

foreach ($excel->sheet()->nextRow() as $rowNum => $row) {
    // ...
}
```

Read areas, key modes, `withHeader()`, the row generator, date formatting, style reading and image
extraction are implemented once and shared by both formats.

## Memory and speed

XLS suits streaming even better than XLSX. A BIFF workbook is a flat sequence of records, so a sheet is
walked in a single forward pass with one record in memory at a time. Two consequences worth knowing:

* Row iteration with `nextRow()` uses constant memory, whatever the size of the sheet.
* Sheets are reached by seeking directly to their offset, so opening a workbook and reading one sheet
  does not touch the others.

The container's allocation table is held as a compact byte string rather than a PHP array, which keeps
the fixed overhead small even for very large files.

## What is supported

* Multiple worksheets, including hidden and very hidden ones
* All cell types — text, numbers, booleans, error values, empty cells
* Automatic date detection through number formats, and the same date formatter API as XLSX
* Cell styles: fonts, fills, borders, alignment, number formats, palette colours
* Merged cells, sheet dimensions, column widths, row heights
* Formula text, including shared formulas
* Embedded images

## Limitations

* **BIFF8 only.** Workbooks written by Excel 5.0/95 (BIFF5/BIFF7) are rejected with a clear message.
  Re-save them as Excel 97-2003 or as XLSX.
* **Encrypted workbooks are not supported** and are rejected rather than returning garbage.
* **Charts, macro sheets and dialogue sheets are skipped** — they carry no cell data.
* **Formula text is best-effort.** Formulas using tokens the decompiler does not render — 3D references
  across sheets, named ranges, array formulas, add-in calls — report `null` as their text. The *cached
  result* of such a formula is always available as the cell value, so reading data is unaffected.
* The format itself is limited to 65 536 rows and 256 columns per sheet.

## Differences from XLSX

These are properties of the formats, not of the reader:

* **Number formats are renumbered by the writing application.** A cell showing `General` may carry the
  built-in format id `0` in XLSX and a custom id such as `164` in XLS. The *pattern* and the resulting
  value type are the same; only `format-num-id` differs.
* **Unformatted date serials.** With `dateFormatter(null)` the XLSX reader returns the literal text as
  written in the XML, for example `'30681.222951388889'` (a string). XLS only ever stored a binary
  double, so it returns a float. The numeric values agree; the exact digits of the original text do not
  exist in the file.
* **Image names.** XLSX preserves the shape name given by the application (`Picture 4`); XLS reports a
  generated name. The image bytes, the anchor cell and the file name are identical.

## See also

* [Getting Started](10-getting-started.md)
* [Reading Data](11-reading-data.md)
* [Cell Styles](14-cell-styles.md)
* [Images](15-images.md)
