# CSV Parsing in FastExcelReader

[← Back to README](../README.md) | [Documentation index](../README.md#documentation) | [🇷🇺 Русский](ru/20-csv.md)

A fast and efficient CSV reader for PHP, integrated into the `fast-excel-reader` package. It supports various encodings, automatic delimiter detection, and large file processing.

## Table of Contents

1. [Features](#features)
2. [Basic Usage](#basic-usage)
3. [Advanced Reading](#advanced-reading)
   - [Reading with Headers](#reading-with-headers)
   - [Reading via Generator (Memory Efficient)](#reading-via-generator-memory-efficient)
   - [Reading into Array](#reading-into-array)
4. [Configuration Options](#configuration-options)
   - [Automatic Delimiter Detection](#automatic-delimiter-detection)
   - [Encoding Support](#encoding-support)
   - [Parsing Modes (Strict vs Tolerant)](#parsing-modes-strict-vs-tolerant)
5. [Error Handling](#error-handling)
6. [Examples](#examples)
7. [API Reference](#api-reference)

## Features

* **Delimiter Detection**: Automatic detection or manual specification of column delimiters.
* **Wide Encoding Support**:
    - UTF-8, UTF-16LE, UTF-16BE, UTF-32LE, UTF-32BE
    - Windows-1251, KOI8-R, ISO-8859-5
    - CP932, Shift_JIS, EUC-JP
    - Any encoding supported by your PHP environment.
* **RFC 4180 Compliant**: Supports fields with or without quotes, and escaped quotes (doubled quotes).
* **Multi-line Fields**: Handles line breaks within quoted fields.
* **Flexible Parsing**:
    - **Strict Mode**: Strictly follows RFC 4180.
    - **Tolerant Mode**: More forgiving with non-standard CSV files (e.g., quotes inside unquoted fields).
* **Memory Efficient**: Capable of reading very large files line by line using Generators.
* **Additional Utilities**: Skip empty lines, skip comment lines, trim fields, custom enclosures, and escape characters.
* **BOM Handling**: Automatic processing of files with Byte Order Mark.

## Basic Usage

The easiest way to open a CSV file is through the `Excel::openCsv()` method.

```php
use avadim\FastExcelReader\Excel;

$file = 'data.csv';
$csv = Excel::openCsv($file);

foreach ($csv->nextRow() as $rowNum => $row) {
    // $row is a simple numerical array
    print_r($row);
}
```

### Opening a CSV as a workbook with `Excel::open()`

Since v4.1.0 the generic `Excel::open()` entry point also accepts CSV files and returns a normal
workbook, so a CSV behaves like any other spreadsheet: it exposes a single sheet through the same
`sheet()`, `nextRow()`, `readRows()`, `readColumns()`, `setReadArea()`, `withHeader()` and key-mode
API as XLSX and XLS. This is what lets format-agnostic code accept CSV without special cases.

```php
use avadim\FastExcelReader\Excel;

$book  = Excel::open('data.csv');          // returns a Csv\CsvBook
$sheet = $book->sheet();                    // the single CSV sheet, named "CSV"

$rows = $book->readRows();
// keys are Excel column letters, exactly as for XLSX:
// [1 => ['A' => 'ID', 'B' => 'Name'], 2 => ['A' => '1', 'B' => 'John'], ...]
```

The format is chosen by the file signature, not the extension: an OLE2 file is XLS, a ZIP file is
XLSX, and anything else is read as delimited text. Pass options as the second argument — including any
[configuration option](#configuration-options), and `['format' => 'csv']` to force the CSV reader:

```php
$book = Excel::open('data.txt', ['delimiter' => "\t", 'encoding' => 'Windows-1251']);
$book = Excel::open($file, ['format' => 'csv']); // force CSV regardless of signature
```

**`open()` vs `openCsv()`.** Both read the same file; they differ in what you get back and in the
default column keys:

| | `Excel::open($csv)` | `Excel::openCsv($csv)` |
|---|---|---|
| Returns | `Csv\CsvBook` (a workbook) | `Csv\CsvReader` (the engine) |
| Default column keys | Excel letters `A`, `B`, ... | zero-based integers `0`, `1`, ... |
| API | full shared `AbstractSheet` API | the `CsvReader` low-level API |

`openCsv()` is unchanged and remains the way to reach the low-level engine (`getCsvField()`,
`getCsvLine()`, `onError()`, `setBufferSize()`, ...). From a `CsvBook` the same engine is available via
`$book->getReader()`.

**What a CSV does not have.** CSV stores no cell formats, so there is no automatic number/date typing
(every value is a string) and no styles: `readRowsWithStyles()`, `readCellStyles()` and `readStyles()`
return empty styles rather than throwing, and there are no merged cells or images. CSV also stores no
dimension, so `dimension()` is empty until you ask for the actual extent with `actualDimension()`,
`countRows()` or `countColumns()`, which scan the file once (the streaming `nextRow()` path never does).

## Advanced Reading

### Reading with Headers

If your CSV has a header row, you can use the `withHeader()` method to use the first row values as keys for subsequent rows.

```php
$csv = Excel::openCsv($file);
$rows = $csv->withHeader()->nextRow();

foreach ($rows as $rowNum => $row) {
    // $row = ['ID' => '1', 'Name' => 'John', 'City' => 'New York']
    echo $row['Name'];
}
```

You can also name the columns yourself. The header row is still skipped, but the names come from your
list instead of from its values. Names are applied in column order, so no column letters are involved:

```php
$rows = $csv->withHeader(['id', 'name', 'city'])->readRows();
// $row = ['id' => '1', 'name' => 'John', 'city' => 'New York']
```

A shorter list renames only the columns it covers; the rest keep the name from the header row.

### Reading via Generator (Memory Efficient)

The `nextRow()` method returns a `\Generator`, which is ideal for processing large files without loading them entirely into memory.

```php
$csv = Excel::openCsv($file);
$generator = $csv->nextRow();

foreach ($generator as $row) {
    // Process each row
}
```

### Reading into Array

If the file is small and you need all data at once:

```php
$csv = Excel::openCsv($file);
$allRows = $csv->readRows();
```

## Configuration Options

You can pass an array of options or a `CsvOptions` object to `openCsv()`.

```php
use avadim\FastExcelReader\Csv\CsvOptions;

$options = [
    'delimiter' => ';',
    'enclosure' => '"',
    'encoding' => 'UTF-8',
    'trim_fields' => true,
    'skip_empty_lines' => true,
];

$csv = Excel::openCsv($file, $options);

// Other ways to set options:
$options = new CsvOptions($options)
    ->setDelimiter(';')
    ->setEnclosure('"')
    ->setEncoding('UTF-8')
    ->setTrimFields(true)
    ->setSkipEmptyLines(true)
;

$csv = Excel::openCsv($file, $options);
```

Available options:

| Option             | Type     | Default    | Description                                                                  |
|--------------------|----------|------------|------------------------------------------------------------------------------|
| `mode`             | `string` | `'strict'` | Parsing mode: `'strict'` or `'tolerant'`                                     |
| `delimiter`        | `string` | `null`     | Column delimiter (e.g., `,`, `;`, `\t`). `null` or `'auto'` for auto-detect. |
| `enclosure`        | `string` | `"`        | Field enclosure character.                                                   |
| `encoding`         | `string` | `null`     | Input file encoding. `null` for auto-detect.                                 |
| `double_quotes`    | `bool`   | `true`     | Whether to handle doubled quotes as escaped quotes.                          |
| `escape`           | `string` | `''`       | Escape character (e.g., `\`).                                                |
| `trim_fields`      | `bool`   | `true`     | Whether to trim leading/trailing whitespace from unquoted fields.            |
| `skip_empty_lines` | `bool`   | `true`     | Whether to skip lines that are empty.                                        |
| `comment_prefix`   | `string` | `null`     | Character(s) that indicate a comment line (e.g., `#`).                       |

### Automatic Delimiter Detection

If `delimiter` is set to `null` or `'auto'`, the reader will attempt to detect the delimiter by analyzing the first few lines of the file.

### Encoding Support

The reader automatically detects most common encodings including UTF-8, UTF-16, and various regional encodings like Windows-1251 or Shift_JIS. You can also manually specify any encoding supported by PHP's `mb_convert_encoding`.

### Parsing Modes (Strict vs Tolerant)

* **Strict Mode (`'strict'`)**: Throws errors if the CSV does not strictly follow RFC 4180 (e.g., if there is text after a closing quote or unescaped quotes inside a field).
* **Tolerant Mode (`'tolerant'`)**: Attempts to recover and read as much data as possible when encountering malformed CSV structures.

## Error Handling

You can define a custom error handler to manage parsing issues, especially useful in `tolerant` mode.

```php
$csv = Excel::openCsv($file);
$csv->onError(function($code, $error, $line, $lineNo, $colNo) {
    echo "Error on line $lineNo, col $colNo: $error\n";
    echo "Line content: $line\n";
});
```

## Examples

### Reading a Tab-Separated File (TSV)

```php
$csv = Excel::openCsv('data.tsv', ['delimiter' => "\t"]);
foreach ($csv->nextRow() as $row) {
    // ...
}
```

### Handling Windows-1251 Encoded Files

```php
$csv = Excel::openCsv('russian_data.csv', ['encoding' => 'Windows-1251']);
foreach ($csv->nextRow() as $row) {
    // ...
}
```

### Skipping Comments

```php
$csv = Excel::openCsv('config.csv', ['comment_prefix' => '#']);
foreach ($csv->nextRow() as $row) {
    // Rows starting with # will be ignored
}
```

## API Reference

* [Class Csv\CsvReader](94-api-class-csv-reader.md)
* [Class Csv\CsvOptions](95-api-class-csv-options.md)
