# Class \avadim\FastExcelReader\Ods\OdsSheet

---

* [__construct()](#__construct)
* [actualDimension()](#actualdimension)
* [countCols()](#countcols) – Count columns by dimension value, alias of countColumns()
* [countColumns()](#countcolumns)
* [countRows()](#countrows)
* [dimension()](#dimension) – Get sheet dimension range (e.g. "A1:C10")
* [dimensionArray()](#dimensionarray)
* [firstCol()](#firstcol) – Get letter of the first column in the read area
* [firstRow()](#firstrow) – Get number of the first row in the read area
* [from()](#from) – Set top left of read area. Alias of setReadArea()
* [getMergedCells()](#getmergedcells)
* [getReadRowNum()](#getreadrownum) – Get the number of the last row read
* [id()](#id) – Get sheet ID
* [isActive()](#isactive) – Returns true if the sheet is active
* [isHidden()](#ishidden) – Returns true if the sheet is hidden
* [isMerged()](#ismerged) – Returns true if the cell is merged
* [isName()](#isname) – Case-insensitive name checking
* [isVisible()](#isvisible) – Returns true if the sheet is visible
* [maxColumn()](#maxcolumn)
* [maxRow()](#maxrow)
* [mergedRange()](#mergedrange) – Get merged range for the cell
* [minColumn()](#mincolumn)
* [minRow()](#minrow)
* [name()](#name) – Get sheet name
* [nextRow()](#nextrow) – Read cell values row by row, returns either an array of values or an array of arrays
* [path()](#path)
* [readCallback()](#readcallback) – Reads cell values and passes them to a callback function
* [readCells()](#readcells) – Returns values and styles of cells as array
* [readCellsFrom()](#readcellsfrom) – Set read area and returns cell values as a one-dimensional array [address => value]
* [readCellStyles()](#readcellstyles) – Returns styles of cells as array
* [readCellsWithStyles()](#readcellswithstyles) – Returns cell values and styles as a one-dimensional array [address => value]:
* [readCellsWithStylesFrom()](#readcellswithstylesfrom) – Set read area and returns cell values and styles as a one-dimensional array [address => value]
* [readColumns()](#readcolumns) – Returns cell values as a two-dimensional array from default sheet [col][row]
* [readColumnsFrom()](#readcolumnsfrom) – Set read area and returns cell values as a two-dimensional array from default sheet [col][row]
* [readColumnsWithStyles()](#readcolumnswithstyles) – Returns cell values and styles as a two-dimensional array [column][row]
* [readColumnsWithStylesFrom()](#readcolumnswithstylesfrom) – Set read area and returns cell values and styles as a two-dimensional array [column][row]
* [readFirstRow()](#readfirstrow) – Returns values of cells of 1st row as array
* [readFirstRowCells()](#readfirstrowcells) – Returns values and styles of cells of 1st row as array
* [readFirstRowCellsFrom()](#readfirstrowcellsfrom) – Set read area and returns cell values of 1st row as array [address => value]
* [readFirstRowFrom()](#readfirstrowfrom) – Set read area and returns values of cells of 1st row as array
* [readFirstRowWithStyles()](#readfirstrowwithstyles) – Returns values and styles of cells of 1st row as array
* [readFirstRowWithStylesFrom()](#readfirstrowwithstylesfrom) – Set read area and returns values and styles of cells of 1st row as array
* [readNextRow()](#readnextrow) – Read the next row from the generator
* [readRows()](#readrows) – Returns cell values as a two-dimensional array
* [readRowsFrom()](#readrowsfrom) – Read rows from a given area $areaRange
* [readRowsWithStyles()](#readrowswithstyles) – Returns values, styles, and other info of cells as array
* [readRowsWithStylesFrom()](#readrowswithstylesfrom) – Set read area and returns values, styles, and other info of cells as array
* [reset()](#reset) – Reset read generator
* [rewind()](#rewind) – Rewind read generator, alias of reset()
* [setDateFormat()](#setdateformat) – Set date format
* [setReadArea()](#setreadarea) – Set top left and right bottom of read area
* [setReadAreaColumns()](#setreadareacolumns) – setReadArea('C:AZ') - set left and right columns of read area
* [setState()](#setstate) – Set sheet state (visible, hidden, veryHidden)
* [state()](#state) – Get sheet state
* [withHeader()](#withheader) – Enables header mode

---

## __construct()

---

```php
public function __construct(string $name, string $sheetId,
                            avadim\FastExcelReader\Ods\OdsBook $book)
```


### Parameters

* `$name`
* `$sheetId`
* `$book`

---

## actualDimension()

---

```php
public function actualDimension(): string
```


### Parameters

_None_

---

## countCols()

---

```php
public function countCols(?string $range = null): int
```
_Count columns by dimension value, alias of countColumns()_

### Parameters

* `string|null $range`

---

## countColumns()

---

```php
public function countColumns(?string $range = null): int
```


### Parameters

* `$range`

---

## countRows()

---

```php
public function countRows(?string $range = null): int
```


### Parameters

* `$range`

---

## dimension()

---

```php
public function dimension(): ?string
```
_Get sheet dimension range (e.g. "A1:C10")_

### Parameters

_None_

---

## dimensionArray()

---

```php
public function dimensionArray(): array
```


### Parameters

_None_

---

## firstCol()

---

```php
public function firstCol(): string
```
_Get letter of the first column in the read area_

### Parameters

_None_

---

## firstRow()

---

```php
public function firstRow(): int
```
_Get number of the first row in the read area_

### Parameters

_None_

---

## from()

---

```php
public function from(string $topLeftCell,
                     ?bool $firstRowKeys = false): avadim\FastExcelReader\AbstractSheet
```
_Set top left of read area. Alias of setReadArea()_

### Parameters

* `string $topLeftCell`
* `bool|null $firstRowKeys`

---

## getMergedCells()

---

```php
public function getMergedCells(): ?array
```


### Parameters

_None_

---

## getReadRowNum()

---

```php
public function getReadRowNum(): int
```
_Get the number of the last row read_

### Parameters

_None_

---

## id()

---

```php
public function id(): string
```
_Get sheet ID_

### Parameters

_None_

---

## isActive()

---

```php
public function isActive(): bool
```
_Returns true if the sheet is active_

### Parameters

_None_

---

## isHidden()

---

```php
public function isHidden(): bool
```
_Returns true if the sheet is hidden_

### Parameters

_None_

---

## isMerged()

---

```php
public function isMerged(string $cellAddress): bool
```
_Returns true if the cell is merged_

### Parameters

* `string $cellAddress`

---

## isName()

---

```php
public function isName(string $name): bool
```
_Case-insensitive name checking_

### Parameters

* `string $name`

---

## isVisible()

---

```php
public function isVisible(): bool
```
_Returns true if the sheet is visible_

### Parameters

_None_

---

## maxColumn()

---

```php
public function maxColumn(?string $range = null): string
```


### Parameters

* `$range`

---

## maxRow()

---

```php
public function maxRow(?string $range = null): int
```


### Parameters

* `$range`

---

## mergedRange()

---

```php
public function mergedRange(string $cellAddress): ?string
```
_Get merged range for the cell_

### Parameters

* `string $cellAddress`

---

## minColumn()

---

```php
public function minColumn(?string $range = null): string
```


### Parameters

* `$range`

---

## minRow()

---

```php
public function minRow(?string $range = null): int
```


### Parameters

* `$range`

---

## name()

---

```php
public function name(): string
```
_Get sheet name_

### Parameters

_None_

---

## nextRow()

---

```php
public function nextRow($columnKeys, ?int $resultMode = null,
                        ?bool $styleIdxInclude = null,
                        ?int $rowLimit = 0): ?Generator
```
_Read cell values row by row, returns either an array of values or an array of arrays_

_nextRow(..., ...) : <rowNum> => \[<colNum1> => <value1>, <colNum2> => <value2>, ...]nextRow(..., ..., true) : <rowNum> => \[<colNum1> => \['v' => <value1>, 's' => <style1>], <colNum2> => \['v' => <value2>, 's' => <style2>], ...]_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`
* `bool|null $styleIdxInclude`
* `int|null $rowLimit`

---

## path()

---

```php
public function path(): string
```


### Parameters

_None_

---

## readCallback()

---

```php
public function readCallback(callable $callback, $columnKeys,
                             ?int $resultMode = null,
                             ?bool $styleIdxInclude = null)
```
_Reads cell values and passes them to a callback function_

### Parameters

* `callback $callback` – Callback function($row, $col, $value)
* `array|bool|int|null $columnKeys`
* `int|null $resultMode`
* `bool|null $styleIdxInclude`

---

## readCells()

---

```php
public function readCells(?bool $styleIdxInclude = null): array
```
_Returns values and styles of cells as array_

### Parameters

* `bool|null $styleIdxInclude`

---

## readCellsFrom()

---

```php
public function readCellsFrom(string $areaRange,
                              ?bool $styleIdxInclude = null): array
```
_Set read area and returns cell values as a one-dimensional array \[address => value]_

### Parameters

* `string $areaRange`
* `bool|null $styleIdxInclude`

---

## readCellStyles()

---

```php
public function readCellStyles(?bool $flat = false,
                               ?string $part = null): array
```
_Returns styles of cells as array_

### Parameters

* `bool|null $flat`
* `string|null $part`

---

## readCellsWithStyles()

---

```php
public function readCellsWithStyles(?string $styleKey = null): array
```
_Returns cell values and styles as a one-dimensional array \[address => value]:'v' => _value_'s' => _styles_'f' => _formula_'t' => _type_'o' => _original_value__

### Parameters

* `string|null $styleKey` – If specified, only this style property will be returned (e.g. 'fill-color')

---

## readCellsWithStylesFrom()

---

```php
public function readCellsWithStylesFrom(string $areaRange,
                                        ?string $styleKey = null): array
```
_Set read area and returns cell values and styles as a one-dimensional array \[address => value]_

### Parameters

* `string $areaRange`
* `string|null $styleKey` – If specified, only this style property will be returned (e.g. 'fill-color')

---

## readColumns()

---

```php
public function readColumns($columnKeys, ?int $resultMode = null,
                            ?bool $styleIdxInclude = null): array
```
_Returns cell values as a two-dimensional array from default sheet \[col]\[row]\['A' => \[1 => _value_A1_], \[2 => _value_A2_]],\['B' => \[1 => _value_B1_], \[2 => _value_B2_]]_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`
* `bool|null $styleIdxInclude`

---

## readColumnsFrom()

---

```php
public function readColumnsFrom(string $areaRange, $columnKeys,
                                ?int $resultMode = null,
                                ?bool $styleIdxInclude = null): array
```
_Set read area and returns cell values as a two-dimensional array from default sheet \[col]\[row]_

### Parameters

* `string $areaRange`
* `array|bool|int|null $columnKeys`
* `int|null $resultMode`
* `bool|null $styleIdxInclude`

---

## readColumnsWithStyles()

---

```php
public function readColumnsWithStyles($columnKeys,
                                      ?int $resultMode = null): array
```
_Returns cell values and styles as a two-dimensional array \[column]\[row]_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`

---

## readColumnsWithStylesFrom()

---

```php
public function readColumnsWithStylesFrom(string $areaRange, $columnKeys,
                                          ?int $resultMode = null): array
```
_Set read area and returns cell values and styles as a two-dimensional array \[column]\[row]_

### Parameters

* `string $areaRange`
* `$columnKeys`
* `int|null $resultMode`

---

## readFirstRow()

---

```php
public function readFirstRow($columnKeys,
                             ?bool $styleIdxInclude = null): array
```
_Returns values of cells of 1st row as array_

### Parameters

* `array|bool|int|null $columnKeys`
* `bool|null $styleIdxInclude`

---

## readFirstRowCells()

---

```php
public function readFirstRowCells(?bool $styleIdxInclude = null): array
```
_Returns values and styles of cells of 1st row as array_

### Parameters

* `bool|null $styleIdxInclude`

---

## readFirstRowCellsFrom()

---

```php
public function readFirstRowCellsFrom(string $areaRange,
                                      ?bool $styleIdxInclude = null): array
```
_Set read area and returns cell values of 1st row as array \[address => value]_

_Like readCellsFrom(), this method takes no column keys, because the result is keyed by cell address and renaming a column would corrupt it._

### Parameters

* `string $areaRange`
* `bool|null $styleIdxInclude`

---

## readFirstRowFrom()

---

```php
public function readFirstRowFrom(string $areaRange, $columnKeys,
                                 ?bool $styleIdxInclude = null): array
```
_Set read area and returns values of cells of 1st row as array_

### Parameters

* `string $areaRange`
* `array|bool|int|null $columnKeys`
* `bool|null $styleIdxInclude`

---

## readFirstRowWithStyles()

---

```php
public function readFirstRowWithStyles($columnKeys): array
```
_Returns values and styles of cells of 1st row as array_

### Parameters

* `array|bool|int|null $columnKeys`

---

## readFirstRowWithStylesFrom()

---

```php
public function readFirstRowWithStylesFrom(string $areaRange,
                                           $columnKeys): array
```
_Set read area and returns values and styles of cells of 1st row as array_

### Parameters

* `string $areaRange`
* `array|bool|int|null $columnKeys`

---

## readNextRow()

---

```php
public function readNextRow(): mixed
```
_Read the next row from the generator_

### Parameters

_None_

---

## readRows()

---

```php
public function readRows($columnKeys, ?int $resultMode = null,
                         ?bool $styleIdxInclude = null): array
```
_Returns cell values as a two-dimensional array\[1 => \['A' => _value_A1_], \['B' => _value_B1_]],\[2 => \['A' => _value_A2_], \['B' => _value_B2_]]_

_readRows()readRows(true)readRows(false, Excel::KEYS_ZERO_BASED)readRows(Excel::KEYS_ZERO_BASED | Excel::KEYS_RELATIVE)_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`
* `bool|null $styleIdxInclude`

---

## readRowsFrom()

---

```php
public function readRowsFrom(string $areaRange, $columnKeys,
                             ?int $resultMode = null,
                             ?bool $styleIdxInclude = null): array
```
_Read rows from a given area $areaRange_

### Parameters

* `string $areaRange`
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
_Returns values, styles, and other info of cells as array_

_\['v' => _value_,'s' => _styles_,'f' => _formula_,'t' => _type_,'o' => '_original_value_]_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`

---

## readRowsWithStylesFrom()

---

```php
public function readRowsWithStylesFrom(string $areaRange, $columnKeys,
                                       ?int $resultMode = null): array
```
_Set read area and returns values, styles, and other info of cells as array_

### Parameters

* `string $areaRange`
* `array|bool|int|null $columnKeys`
* `int|null $resultMode`

---

## reset()

---

```php
public function reset($columnKeys, ?int $resultMode = null,
                      ?bool $styleIdxInclude = null,
                      ?int $rowLimit = 0): ?Generator
```
_Reset read generator_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`
* `bool|null $styleIdxInclude`
* `int|null $rowLimit`

---

## rewind()

---

```php
public function rewind($columnKeys, ?int $resultMode = null,
                       ?bool $styleIdxInclude = null,
                       ?int $rowLimit = 0): ?Generator
```
_Rewind read generator, alias of reset()_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`
* `bool|null $styleIdxInclude`
* `int|null $rowLimit`

---

## setDateFormat()

---

```php
public function setDateFormat($dateFormat): avadim\FastExcelReader\AbstractSheet
```
_Set date format_

### Parameters

* `$dateFormat`

---

## setReadArea()

---

```php
public function setReadArea(string $areaRange,
                            ?bool $firstRowKeys = false): avadim\FastExcelReader\AbstractSheet
```
_Set top left and right bottom of read area_

### Parameters

* `string $areaRange`
* `bool|null $firstRowKeys`

---

### Examples

```php
setReadArea('C3:AZ28'); // set top left and right bottom of read area
setReadArea('C3'); // set top left only
```


---

## setReadAreaColumns()

---

```php
public function setReadAreaColumns(string $columnsRange,
                                   ?bool $firstRowKeys = false): avadim\FastExcelReader\AbstractSheet
```
_setReadArea('C:AZ') - set left and right columns of read areasetReadArea('C') - set left column only_

### Parameters

* `string $columnsRange`
* `bool|null $firstRowKeys`

---

## setState()

---

```php
public function setState(string $state): avadim\FastExcelReader\AbstractSheet
```
_Set sheet state (visible, hidden, veryHidden)_

### Parameters

* `string $state`

---

## state()

---

```php
public function state(): string
```
_Get sheet state_

### Parameters

_None_

---

## withHeader()

---

```php
public function withHeader(?array $columnNames = null): avadim\FastExcelReader\AbstractSheet
```
_Enables header mode_

_Treats the first row of the read area as a header row and returns subsequent rowsas associative arrays keyed by column names._

_Pass $columnNames to name the columns yourself: the first row is still skipped,but the names are taken from the list instead of from its values. Names areapplied in column order, starting at the first column of the read area, so theyneed no knowledge of column letters. Columns past the end of the list keep thename from the header row._

### Parameters

* `array|null $columnNames` – Column names in order, or NULL to use the header row values

---
