# Class \avadim\FastExcelReader\Excel

---

* [colLetter()](#colletter)
* [colNum()](#colnum)
* [createReader()](#createreader) – Create XML reader object
* [createSheet()](#createsheet) – Create sheet object
* [isOds()](#isods)
* [isXls()](#isxls) – TRUE if the file starts with the OLE2 compound file signature
* [isXlsx()](#isxlsx) – TRUE for the supported XML workbook family, including macro books and templates.
* [open()](#open) – Open a spreadsheet, choosing the reader by the file signature
* [openCsv()](#opencsv) – Open CSV through the common workbook API
* [openCsvBook()](#opencsvbook) – Open CSV as a single-sheet workbook without format detection
* [openCsvReader()](#opencsvreader) – Open the low-level CSV reader without a workbook wrapper
* [openOds()](#openods)
* [openStream()](#openstream) – Open a spreadsheet from an open stream resource
* [openString()](#openstring) – Open a spreadsheet held in a string, choosing the reader by its signature
* [openXls()](#openxls) – Open an XLS (Excel 97-2003, BIFF8) file
* [openXlsx()](#openxlsx) – Open an XLSX workbook without format detection
* [setTempDir()](#settempdir)
* [validate()](#validate) – Check required XML workbook/ODS package parts and XML well-formedness

---

## colLetter()

---

```php
public static function colLetter(int $column): string
```


### Parameters

* `$column`

---

## colNum()

---

```php
public static function colNum(string $column): int
```


### Parameters

* `$column`

---

## createReader()

---

```php
public static function createReader(string $file,
                                    ?array $parserProperties = []): Interfaces\InterfaceXmlReader
```
_Create XML reader object_

### Parameters

* `string $file`
* `array|null $parserProperties`

---

## createSheet()

---

```php
public static function createSheet(string $sheetName, $sheetId, $file, $path,
                                   $excel): Interfaces\InterfaceSheetReader
```
_Create sheet object_

### Parameters

* `string $sheetName`
* `int|string $sheetId`
* `string $file`
* `string $path`
* `XlsxBook $excel`

---

## isOds()

---

```php
public static function isOds(string $file): bool
```


### Parameters

* `$file`

---

## isXls()

---

```php
public static function isXls(string $file): bool
```
_TRUE if the file starts with the OLE2 compound file signature_

### Parameters

* `string $file`

---

## isXlsx()

---

```php
public static function isXlsx(string $file): bool
```
_TRUE for the supported XML workbook family, including macro books and templates._

### Parameters

* `$file`

---

## open()

---

```php
public static function open(string $file, $options): AbstractBook
```
_Open a spreadsheet, choosing the reader by the file signatureThe OLE2 magic number is a legacy XLS workbook, a supported ZIP container is an XML workbook or ODS, and anything else is treated as delimited text (CSV). The file extension is not consulted, because it is often wrong on files arriving from other systems. Pass $options\['format'] = 'csv' to force the CSV reader, and any CsvOptions keys (delimiter, enclosure, encoding, ...) to configure it._

### Parameters

* `string $file`
* `CsvOptions|array|null $options`

---

## openCsv()

---

```php
public static function openCsv(string $file, $options): Csv\CsvBook
```
_Open CSV through the common workbook APIReturns CsvBook in 5.0. Use openCsvReader() for getCsvLine(), fromRow() and numeric CSV column keys._

### Parameters

* `string $file`
* `CsvOptions|array|null $options`

---

## openCsvBook()

---

```php
public static function openCsvBook(string $file, $options): Csv\CsvBook
```
_Open CSV as a single-sheet workbook without format detectionUnlike open(), this factory also accepts an empty CSV file._

### Parameters

* `string $file`
* `CsvOptions|array|null $options`

---

## openCsvReader()

---

```php
public static function openCsvReader(string $file, $options): Csv\CsvReader
```
_Open the low-level CSV reader without a workbook wrapperUse this factory for low-level CSV access. openCsv() and openCsvBook() return a workbook._

### Parameters

* `string $file`
* `CsvOptions|array|null $options`

---

## openOds()

---

```php
public static function openOds(string $file): Ods\OdsBook
```


### Parameters

* `$file`

---

## openStream()

---

```php
public static function openStream($stream, $options): AbstractBook
```
_Open a spreadsheet from an open stream resourceThe stream is copied (from its current position, without seeking, so non-rewindable streams such as HTTP wrappers work) into a temporary file and then opened like open(). This is the entry point for URLs (fopen('https://...')), php://memory and Flysystem/S3 read streams. The caller keeps ownership of the stream; it is not closed here. The temporary file is owned by the returned workbook and removed on close(), with shutdown cleanup as a fallback._

### Parameters

* `resource $stream` – An open, readable stream resource
* `CsvOptions|array|null $options` – Same options as open()

---

## openString()

---

```php
public static function openString(string $content, $options): AbstractBook
```
_Open a spreadsheet held in a string, choosing the reader by its signatureThe content is written to a temporary file and then opened exactly like open() - format is detected from the bytes, not from any file name, so an XLSX/XLS/CSV payload each read back the same as its on-disk counterpart. The temporary file is owned by the returned workbook and removed on close(), with shutdown cleanup as a fallback. Handy for content coming from a database blob, an HTTP response body or an S3/Flysystem read._

### Parameters

* `string $content` – Raw bytes of the workbook
* `CsvOptions|array|null $options` – Same options as open()

---

## openXls()

---

```php
public static function openXls(string $file): Xls\XlsBook
```
_Open an XLS (Excel 97-2003, BIFF8) file_

### Parameters

* `string $file`

---

## openXlsx()

---

```php
public static function openXlsx(string $file): Xlsx\XlsxBook
```
_Open an XLSX workbook without format detection_

### Parameters

* `string $file`

---

## setTempDir()

---

```php
public static function setTempDir($directory)
```


### Parameters

* `$directory`

---

## validate()

---

```php
public static function validate(string $file, ?array &$errors = []): bool
```
_Check required XML workbook/ODS package parts and XML well-formednessThis is not OOXML schema validation. The libxml error mode is restored; its diagnostic buffer is cleared before and after validation. XML errors are returned in $errors; missing package parts return false without XML errors._

### Parameters

* `string $file`
* `array|null $errors`

---
