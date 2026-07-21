# Class \avadim\FastExcelReader\Excel

---

* [__construct()](#__construct) – Excel constructor
* [colLetter()](#colletter) – Convert column number to letter
* [colNum()](#colnum) – Converts an alphabetic column index to a numeric
* [createReader()](#createreader) – Create XML reader object
* [createSheet()](#createsheet) – Create sheet object
* [isXls()](#isxls) – TRUE if the file starts with the OLE2 compound file signature
* [open()](#open) – Open a spreadsheet, choosing the reader by the file signature
* [openCsv()](#opencsv) – Open CSV file
* [openXls()](#openxls) – Open an XLS (Excel 97-2003, BIFF8) file
* [setTempDir()](#settempdir) – Set directory for temporary files
* [validate()](#validate) – Validate XLSX file
* [countExtraImages()](#countextraimages) – Count "extra" images (images that are in the media folder but not in the drawings)
* [countImages()](#countimages) – Returns the total count of images in the workbook
* [countSheets()](#countsheets) – Returns the number of sheets in the workbook
* [dateFormatter()](#dateformatter) – Set custom date formatter
* [formatDate()](#formatdate) – Format date value
* [from()](#from) – Set top left of read area
* [getCompleteStyleByIdx()](#getcompletestylebyidx) – Get complete style by style index
* [getDateFormat()](#getdateformat) – Get current date format
* [getDateFormatPattern()](#getdateformatpattern) – Get PHP date format pattern by style index
* [getDateFormatter()](#getdateformatter) – Get date formatter
* [getDefinedNames()](#getdefinednames) – Get defined names of workbook
* [getFirstSheet()](#getfirstsheet) – Returns the first sheet as default
* [getFormatPattern()](#getformatpattern) – Get format pattern by style index
* [getImageList()](#getimagelist) – Get the list of images from the workbook
* [getSheet()](#getsheet) – Get sheet object by name and optionally set read area and options
* [getSheetById()](#getsheetbyid) – Returns a sheet by ID
* [getSheetNames()](#getsheetnames) – Get names array of all sheets
* [hasDrawings()](#hasdrawings) – Returns TRUE if the workbook contains an any draw objects (not images only)
* [hasExtraImages()](#hasextraimages) – Returns TRUE if there are any "extra" images
* [hasImages()](#hasimages) – Returns TRUE if any sheet contains an image object
* [hiddenSheets()](#hiddensheets) – Array of hidden sheets only
* [innerFileList()](#innerfilelist) – Get list of inner files in XLSX
* [mediaImageFiles()](#mediaimagefiles) – Get list of media image files in the workbook
* [metadataImage()](#metadataimage) – Get image file name from metadata by index
* [readCallback()](#readcallback) – Reads cell values and passes them to a callback function
* [readCells()](#readcells) – Returns the values of all cells as array
* [readCellStyles()](#readcellstyles) – Returns the styles of all cells as array
* [readCellsWithStyles()](#readcellswithstyles) – Returns the values and styles of all cells as array
* [readColumns()](#readcolumns) – Returns cell values as a two-dimensional array from default sheet [col][row]
* [readColumnsWithStyles()](#readcolumnswithstyles) – Returns cell values and styles as a two-dimensional array from default sheet [col][row]
* [readRows()](#readrows) – Returns cell values as a two-dimensional array from default sheet [row][col]
* [readRowsWithStyles()](#readrowswithstyles) – Returns cell values and styles as a two-dimensional array from default sheet [row][col]
* [readStyles()](#readstyles) – Read all workbook styles
* [selectFirstSheet()](#selectfirstsheet) – Selects the first sheet as default
* [selectSheet()](#selectsheet) – Selects default sheet by name
* [selectSheetById()](#selectsheetbyid) – Selects default sheet by ID
* [setDateFormat()](#setdateformat) – Set date format for reading
* [setReadArea()](#setreadarea) – Set top left and right bottom of read area
* [sharedString()](#sharedstring) – Get string by index
* [sheet()](#sheet) – Get current or specified sheet
* [sheetExists()](#sheetexists) – Returns TRUE if a sheet with the given name exists
* [sheets()](#sheets) – Array of all sheets
* [stat()](#stat) – Returns statistics of the workbook: per-sheet breakdown and totals
* [styleByIdx()](#stylebyidx) – Get style array by style index
* [timestamp()](#timestamp) – Convert date to timestamp
* [visibleSheets()](#visiblesheets) – Array of visible sheets only

---

## __construct()

---

```php
public function __construct(?string $file = null, ?string $tempDir = '')
```
_Excel constructor_

### Parameters

* `string|null $file`
* `string|null $tempDir`

---

## colLetter()

---

```php
public static function colLetter(int $colNumber): string
```
_Convert column number to letter_

### Parameters

* `int $colNumber` – ONE based

---

## colNum()

---

```php
public static function colNum(string $colLetter): int
```
_Converts an alphabetic column index to a numeric_

### Parameters

* `string $colLetter`

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
* `Excel $excel`

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

## open()

---

```php
public static function open(string $file): AbstractBook
```
_Open a spreadsheet, choosing the reader by the file signature_

_A ZIP container is XLSX, the OLE2 magic number is a legacy XLS workbook. The file extension is not consulted, because it is often wrong on files arriving from other systems._

### Parameters

* `string $file`

---

## openCsv()

---

```php
public static function openCsv(string $file, $options): Csv\CsvReader
```
_Open CSV file_

### Parameters

* `string $file`
* `CsvOptions|array|null $options`

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

## setTempDir()

---

```php
public static function setTempDir($tempDir)
```
_Set directory for temporary files_

### Parameters

* `string $tempDir`

---

## validate()

---

```php
public static function validate(string $file, ?array &$errors = []): bool
```
_Validate XLSX file_

### Parameters

* `string $file`
* `array|null $errors`

---

## countExtraImages()

---

```php
public function countExtraImages(): int
```
_Count "extra" images (images that are in the media folder but not in the drawings)_

### Parameters

_None_

---

## countImages()

---

```php
public function countImages(): int
```
_Returns the total count of images in the workbook_

### Parameters

_None_

---

## countSheets()

---

```php
public function countSheets(): int
```
_Returns the number of sheets in the workbook_

### Parameters

_None_

---

## dateFormatter()

---

```php
public function dateFormatter($formatter): AbstractBook
```
_Set custom date formatter_

### Parameters

* `\Closure|callable|string|bool|null $formatter`

---

## formatDate()

---

```php
public function formatDate($value, $format, $styleIdx): false|mixed|string
```
_Format date value_

### Parameters

* `mixed $value`
* `string|null $format`
* `int|null $styleIdx`

---

## from()

---

```php
public function from(string $topLeftCell, 
                     ?bool $firstRowKeys = false): AbstractSheet
```
_Set top left of read area_

### Parameters

* `string $topLeftCell`
* `bool|null $firstRowKeys`

---

## getCompleteStyleByIdx()

---

```php
public function getCompleteStyleByIdx(int $styleIdx, 
                                      ?bool $flat = false): array
```
_Get complete style by style index_

### Parameters

* `int $styleIdx`
* `bool|null $flat`

---

## getDateFormat()

---

```php
public function getDateFormat(): ?string
```
_Get current date format_

### Parameters

_None_

---

## getDateFormatPattern()

---

```php
public function getDateFormatPattern(int $styleIdx): ?string
```
_Get PHP date format pattern by style index_

### Parameters

* `int $styleIdx`

---

## getDateFormatter()

---

```php
public function getDateFormatter(): callable|\Closure|bool|null
```
_Get date formatter_

### Parameters

_None_

---

## getDefinedNames()

---

```php
public function getDefinedNames(): array
```
_Get defined names of workbook_

### Parameters

_None_

---

## getFirstSheet()

---

```php
public function getFirstSheet(?string $areaRange = null, 
                              ?bool $firstRowKeys = false): AbstractSheet
```
_Returns the first sheet as default_

### Parameters

* `string|null $areaRange`
* `bool|null $firstRowKeys`

---

## getFormatPattern()

---

```php
public function getFormatPattern(int $styleIdx): mixed|string
```
_Get format pattern by style index_

### Parameters

* `int $styleIdx`

---

## getImageList()

---

```php
public function getImageList(): array
```
_Get the list of images from the workbook_

### Parameters

_None_

---

## getSheet()

---

```php
public function getSheet(?string $name = null, ?string $areaRange = null, 
                         ?bool $firstRowKeys = false): AbstractSheet
```
_Get sheet object by name and optionally set read area and options_

### Parameters

* `string|null $name`
* `string|null $areaRange`
* `bool|null $firstRowKeys`

---

## getSheetById()

---

```php
public function getSheetById(int $sheetId, ?string $areaRange = null, 
                             ?bool $firstRowKeys = false): AbstractSheet
```
_Returns a sheet by ID_

### Parameters

* `int $sheetId`
* `string|null $areaRange`
* `bool|null $firstRowKeys`

---

## getSheetNames()

---

```php
public function getSheetNames(): array
```
_Get names array of all sheets_

### Parameters

_None_

---

## hasDrawings()

---

```php
public function hasDrawings(): bool
```
_Returns TRUE if the workbook contains an any draw objects (not images only)_

### Parameters

_None_

---

## hasExtraImages()

---

```php
public function hasExtraImages(): bool
```
_Returns TRUE if there are any "extra" images_

### Parameters

_None_

---

## hasImages()

---

```php
public function hasImages(): bool
```
_Returns TRUE if any sheet contains an image object_

### Parameters

_None_

---

## hiddenSheets()

---

```php
public function hiddenSheets(): array
```
_Array of hidden sheets only_

### Parameters

_None_

---

## innerFileList()

---

```php
public function innerFileList(): array
```
_Get list of inner files in XLSX_

### Parameters

_None_

---

## mediaImageFiles()

---

```php
public function mediaImageFiles(): array
```
_Get list of media image files in the workbook_

### Parameters

_None_

---

## metadataImage()

---

```php
public function metadataImage(int $vmIndex): ?string
```
_Get image file name from metadata by index_

### Parameters

* `int $vmIndex`

---

## readCallback()

---

```php
public function readCallback(callable $callback, ?int $resultMode = null, 
                             ?bool $styleIdxInclude = null)
```
_Reads cell values and passes them to a callback function_

### Parameters

* `callback $callback`
* `int|null $resultMode`
* `bool|null $styleIdxInclude`

---

## readCells()

---

```php
public function readCells(): array
```
_Returns the values of all cells as array_

### Parameters

_None_

---

## readCellStyles()

---

```php
public function readCellStyles(?bool $flat = false): array
```
_Returns the styles of all cells as array_

### Parameters

* `bool|null $flat`

---

## readCellsWithStyles()

---

```php
public function readCellsWithStyles(): array
```
_Returns the values and styles of all cells as array_

### Parameters

_None_

---

## readColumns()

---

```php
public function readColumns($columnKeys, ?int $resultMode = null): array
```
_Returns cell values as a two-dimensional array from default sheet \[col]\[row]_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`

---

## readColumnsWithStyles()

---

```php
public function readColumnsWithStyles($columnKeys, 
                                      ?int $resultMode = null): array
```
_Returns cell values and styles as a two-dimensional array from default sheet \[col]\[row]_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`

---

## readRows()

---

```php
public function readRows($columnKeys, ?int $resultMode = null, 
                         ?bool $styleIdxInclude = null): array
```
_Returns cell values as a two-dimensional array from default sheet \[row]\[col]_

_readRows()readRows(true)readRows(false, Excel::KEYS_ZERO_BASED)readRows(Excel::KEYS_ZERO_BASED | Excel::KEYS_RELATIVE)_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`
* `bool|null $styleIdxInclude`

---

## readRowsWithStyles()

---

```php
public function readRowsWithStyles($columnKeys, 
                                   ?int $resultMode = null): array
```
_Returns cell values and styles as a two-dimensional array from default sheet \[row]\[col]_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`

---

## readStyles()

---

```php
public function readStyles(): array
```
_Read all workbook styles_

### Parameters

_None_

---

## selectFirstSheet()

---

```php
public function selectFirstSheet(?string $areaRange = null, 
                                 ?bool $firstRowKeys = false): AbstractSheet
```
_Selects the first sheet as default_

### Parameters

* `string|null $areaRange`
* `bool|null $firstRowKeys`

---

## selectSheet()

---

```php
public function selectSheet(string $name, ?string $areaRange = null, 
                            ?bool $firstRowKeys = false): AbstractSheet
```
_Selects default sheet by name_

### Parameters

* `string $name`
* `string|null $areaRange`
* `bool|null $firstRowKeys`

---

## selectSheetById()

---

```php
public function selectSheetById(int $sheetId, ?string $areaRange = null, 
                                ?bool $firstRowKeys = false): AbstractSheet
```
_Selects default sheet by ID_

### Parameters

* `int $sheetId`
* `string|null $areaRange`
* `bool|null $firstRowKeys`

---

## setDateFormat()

---

```php
public function setDateFormat(string $dateFormat): AbstractBook
```
_Set date format for reading_

### Parameters

* `string $dateFormat`

---

## setReadArea()

---

```php
public function setReadArea(string $areaRange, 
                            ?bool $firstRowKeys = false): AbstractSheet
```
_Set top left and right bottom of read area_

### Parameters

* `string $areaRange`
* `bool|null $firstRowKeys`

---

## sharedString()

---

```php
public function sharedString($stringId): ?string
```
_Get string by index_

### Parameters

* `int $stringId`

---

## sheet()

---

```php
public function sheet(?string $name = null): ?AbstractSheet
```
_Get current or specified sheet_

### Parameters

* `string|null $name`

---

## sheetExists()

---

```php
public function sheetExists(string $name): bool
```
_Returns TRUE if a sheet with the given name exists_

### Parameters

* `string $name`

---

## sheets()

---

```php
public function sheets(): array
```
_Array of all sheets_

### Parameters

_None_

---

## stat()

---

```php
public function stat(): array
```
_Returns statistics of the workbook: per-sheet breakdown and totals_

_\['sheets' => \['<sheetName>' => \['rows' => \[...], 'cols' => \[...], 'cells' => \['total' => int, 'filled' => int]],...],'total' => \['sheets'  => int,   // number of sheets'visible' => int,   // number of visible sheets'hidden'  => int,   // number of hidden sheets'rows'    => int,   // sum of actual rows over all sheets'cells'   => \['total' => int, 'filled' => int],],]_

_Note: scans every sheet fully (see Sheet::stat()); expensive on large workbooks._

### Parameters

_None_

---

## styleByIdx()

---

```php
public function styleByIdx($styleIdx): array
```
_Get style array by style index_

### Parameters

* `int $styleIdx`

---

## timestamp()

---

```php
public function timestamp($excelDateTime): int
```
_Convert date to timestamp_

### Parameters

* `$excelDateTime`

---

## visibleSheets()

---

```php
public function visibleSheets(): array
```
_Array of visible sheets only_

### Parameters

_None_

---

