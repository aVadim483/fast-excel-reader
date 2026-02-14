# CSV Parsing in FastExcelReader

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

## Features

* **BOM Handling**: Automatic processing of files with Byte Order Mark.
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
