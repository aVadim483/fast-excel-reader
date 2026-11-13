# Upgrading to 5.0

[Русский](UPGRADE.ru.md)

5.0 is under development. PHP >=7.4 remains supported.

## Workbook classes and extension points

`Excel` is now an opening facade with a private constructor. XLSX, XLSM, XLTX and
XLTM return `avadim\FastExcelReader\Xlsx\XlsxBook`; XLS, CSV and ODS return their
respective `XlsBook`, `CsvBook` and `OdsBook`. Generic factories return `AbstractBook`.

Replace `Excel` type hints and `instanceof Excel` with `AbstractBook` for any format,
or `XlsxBook` for the XML workbook family. Replace `new Excel($file)` with
`Excel::openXlsx($file)` or `new XlsxBook($file)`.

Subclass `XlsxBook` for XLSX customization. Public `createSheet()` and `createReader()`
extension points remain on that class, and sheet readers use the owning book's factory.
Calling the generic `Excel::open()` does not instantiate a custom subclass.
`Excel` retains the reading constants, `INDEXED_COLORS`, `colNum()`, `colLetter()`
and `setTempDir()` as convenience APIs.

## CSV

`Excel::openCsv()` now returns `CsvBook`, with Excel column letters by default.

```php
// 4.x low-level code:
$reader = Excel::openCsv($file, $options);
// 5.0 replacement retaining CsvReader methods and numeric column keys:
$reader = Excel::openCsvReader($file, $options);

// Common book/sheet API:
$book = Excel::openCsv($file, $options); // openCsvBook() remains available
$rows = $book->sheet()->readRows();
```

Use `$book->getReader()` when low-level access is needed on an existing CSV book.
Explicit CSV factories still accept empty files. Generic `open()` rejects empty input.
CSV options and `open($file, ['format' => 'csv'])` retain their meaning.

## Detection and validation

`isXlsx()` now checks workbook package parts and a supported workbook content type,
including the XML default-type form used by some producers. A ZIP signature alone
no longer returns true. XLSM/XLTX/XLTM use the same reader; VBA is never executed.
Templates are read as saved content, not instantiated as new documents.

`isOds()` identifies the ODS MIME entry. Generic factories dispatch by content,
not extension. XLSB is rejected explicitly. `validate()` checks required XML-family
or ODS package parts and XML well-formedness, not schemas, formula calculations,
encryption or complete semantic validity. Unreadable input throws an exception.

## Explicit resource ownership

Call `close()` in long-running processes. It closes owned readers and deletes
temporary files produced by `openString()`/`openStream()`. Repeated calls are harmless.
Source files and caller-owned input streams are not deleted or closed.

```php
$book = Excel::openStream($input);
try {
    foreach ($book->sheet()->nextRow() as $row) {
        // Import the row.
    }
} finally {
    $book->close();
}
```

Closing is irreversible. New reads and resuming a suspended row iterator throw
`The workbook is closed`. Reading an entire result with `readRows()` still uses
memory proportional to that result. XLSX/XLS shared strings and XLS images also
occupy workbook memory. Shutdown cleanup remains a fallback for temporary files;
do not rely on garbage collection timing to release files promptly.

## ODS differences

See [ODS reading](docs/22-ods.md). Dates without a timezone use UTC, explicit
offsets are honored, and `dateFormatter(null/false)` returns original ISO date text.
Time/duration values remain ISO duration strings, including values beyond 24 hours.
Full styles, images and translated formula text are not part of the initial ODS reader.
