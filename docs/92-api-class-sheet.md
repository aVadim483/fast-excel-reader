# Class \avadim\FastExcelReader\Sheet

---

* [__construct()](#__construct)
* [actualDimension()](#actualdimension) -- Get the actual dimension range (e.g. "A1:C10")
* [countActualColumns()](#countactualcolumns) -- Returns the actual number of columns from the sheet data area
* [countActualDimension()](#countactualdimension) -- Scan sheet data and returns actual number of rows and columns
* [countActualRows()](#countactualrows) -- Returns the actual number of rows from the sheet data area
* [countCols()](#countcols) -- Count columns by dimension value, alias of countColumns()
* [countColumns()](#countcolumns) -- Count columns by dimension value
* [countImages()](#countimages) -- Count images of the sheet
* [countRows()](#countrows) -- Count rows by dimension value
* [dimension()](#dimension) -- Get sheet dimension range (e.g. "A1:C10")
* [dimensionArray()](#dimensionarray) -- Get sheet dimension as an array
* [extractConditionalFormatting()](#extractconditionalformatting) -- Extracts conditional formatting rules from the sheet
* [extractDataValidations()](#extractdatavalidations) -- Extracts data validation rules from the sheet
* [firstCol()](#firstcol) -- Get letter of the first column in the read area
* [firstRow()](#firstrow) -- Get number of the first row in the read area
* [from()](#from) -- Set top left of read area. Alias of setReadArea()
* [getColAttributes()](#getcolattributes) -- Get column attributes
* [getColumnAttributes()](#getcolumnattributes) -- Get column attributes (width, style, etc)
* [getColumnStyle()](#getcolumnstyle) -- Get column style
* [getColumnWidth()](#getcolumnwidth) -- Get width of the column
* [getConditionalFormatting()](#getconditionalformatting) -- Get conditional formatting rules
* [getDataValidations()](#getdatavalidations) -- Get data validation rules
* [getFreezePaneInfo()](#getfreezepaneinfo) -- Get freeze pane info
* [getImageBlob()](#getimageblob) -- Get image content as binary string
* [getImageList()](#getimagelist) -- Get image list
* [getImageListByRow()](#getimagelistbyrow) -- Get image list by row number
* [getImageMimeType()](#getimagemimetype) -- Get image MIME type
* [getImageName()](#getimagename) -- Get image name
* [getMergedCells()](#getmergedcells) -- Get merged cells. Returns an array
* [getReadRowNum()](#getreadrownum) -- Get the number of the last row read
* [getRowHeight()](#getrowheight) -- Get height of the row
* [getTabColorConfiguration()](#gettabcolorconfiguration) -- Get tab color configuration. Alias of getTabColorConfig()
* [getTabColorInfo()](#gettabcolorinfo) -- Get the tab color info of the sheet
* [hasDrawings()](#hasdrawings) -- Returns true if the sheet has drawings
* [hasImage()](#hasimage) -- Returns TRUE if the cell contains an image
* [id()](#id) -- Get sheet ID
* [imageEntryFullPath()](#imageentryfullpath) -- Get full path to the image in the ZIP archive
* [isActive()](#isactive) -- Returns true if the sheet is active
* [isHidden()](#ishidden) -- Returns true if the sheet is hidden
* [isMerged()](#ismerged) -- Returns true if the cell is merged
* [isName()](#isname) -- Case-insensitive name checking
* [isVisible()](#isvisible) -- Returns true if the sheet is visible
* [maxActualColumn()](#maxactualcolumn) -- Get the last actual column letter
* [maxActualRow()](#maxactualrow) -- Get the last actual row number
* [maxColumn()](#maxcolumn) -- Max column from dimension value
* [maxRow()](#maxrow) -- Max row number from dimension value
* [mergedRange()](#mergedrange) -- Get merged range for the cell
* [minActualColumn()](#minactualcolumn) -- Get the first actual column letter
* [minActualRow()](#minactualrow) -- Get the first actual row number
* [minColumn()](#mincolumn) -- Min column from dimension value
* [minRow()](#minrow) -- Min row number from dimension value
* [name()](#name) -- Get sheet name
* [nextRow()](#nextrow) -- Read cell values row by row, returns either an array of values or an array of arrays
* [path()](#path) -- Get path to the sheet XML file in ZIP archive
* [readCallback()](#readcallback) -- Reads cell values and passes them to a callback function
* [readCells()](#readcells) -- Returns values and styles of cells as array
* [readCellsFrom()](#readcellsfrom) -- Set read area and returns cell values as a one-dimensional array
* [readCellStyles()](#readcellstyles) -- Returns styles of cells as array
* [readCellsWithStyles()](#readcellswithstyles) -- Returns cell values and styles as a one-dimensional array
* [readCellsWithStylesFrom()](#readcellswithstylesfrom) -- Set read area and returns cell values and styles as a one-dimensional array
* [readColumns()](#readcolumns) -- Returns cell values as a two-dimensional array from default sheet
* [readColumnsFrom()](#readcolumnsfrom) -- Set read area and returns cell values as a two-dimensional array from default sheet
* [readColumnsWithStyles()](#readcolumnswithstyles) -- Returns cell values and styles as a two-dimensional array
* [readColumnsWithStylesFrom()](#readcolumnswithstylesfrom) -- Set read area and returns cell values and styles as a two-dimensional array
* [readFirstRow()](#readfirstrow) -- Returns values of cells of 1st row as array
* [readFirstRowCells()](#readfirstrowcells) -- Returns values and styles of cells of 1st row as array
* [readFirstRowCellsFrom()](#readfirstrowcellsfrom) -- Set read area and returns cell values of 1st row as array
* [readFirstRowFrom()](#readfirstrowfrom) -- Set read area and returns values of cells of 1st row as array
* [readFirstRowWithStyles()](#readfirstrowwithstyles) -- Returns values and styles of cells of 1st row as array
* [readFirstRowWithStylesFrom()](#readfirstrowwithstylesfrom) -- Set read area and returns values and styles of cells of 1st row as array
* [readNextRow()](#readnextrow) -- Read the next row from the generator
* [readRows()](#readrows) -- Returns cell values as a two-dimensional array
* [readRowsFrom()](#readrowsfrom) -- Read rows from a given area $areaRange
* [readRowsWithStyles()](#readrowswithstyles) -- Returns values, styles, and other info of cells as array
* [readRowsWithStylesFrom()](#readrowswithstylesfrom) -- Set read area and returns values, styles, and other info of cells as array
* [reset()](#reset) -- Reset read generator
* [rewind()](#rewind) -- Rewind read generator, alias of reset()
* [saveImage()](#saveimage) -- Save image to a file
* [saveImageTo()](#saveimageto) -- Save image to a directory
* [setDateFormat()](#setdateformat) -- Set date format
* [setDefaultRowHeight()](#setdefaultrowheight) -- Set default row height
* [setReadArea()](#setreadarea) -- Set top left and right bottom of read area
* [setReadAreaColumns()](#setreadareacolumns) -- setReadArea('C:AZ') - set left and right columns of read area
* [setState()](#setstate) -- Set sheet state (visible, hidden, veryHidden)
* [state()](#state) -- Get sheet state

---

## __construct()

---

```php
public function __construct(string $sheetName, string $sheetId, string $file, 
                            string $path, $excel)
```


### Parameters

* `string $sheetName`
* `string $sheetId`
* `string $file`
* `string $path`
* `$excel`

---

## actualDimension()

---

```php
public function actualDimension(): string
```
_Get the actual dimension range (e.g. "A1:C10")_

### Parameters

_None_

---

## countActualColumns()

---

```php
public function countActualColumns(): int
```
_Returns the actual number of columns from the sheet data area_

### Parameters

_None_

---

## countActualDimension()

---

```php
public function countActualDimension(bool $countColumns = true, 
                                     bool $countRows = true, 
                                     int $blockSize = 4096): array
```
_Scan sheet data and returns actual number of rows and columns_

### Parameters

* `bool $countColumns`
* `bool $countRows`
* `int $blockSize`

---

## countActualRows()

---

```php
public function countActualRows(): int
```
_Returns the actual number of rows from the sheet data area_

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
_Count columns by dimension value_

### Parameters

* `string|null $range`

---

## countImages()

---

```php
public function countImages(): int
```
_Count images of the sheet_

### Parameters

_None_

---

## countRows()

---

```php
public function countRows(?string $range = null): int
```
_Count rows by dimension value_

### Parameters

* `string|null $range`

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
_Get sheet dimension as an array_

### Parameters

_None_

---

## extractConditionalFormatting()

---

```php
public function extractConditionalFormatting(): void
```
_Extracts conditional formatting rules from the sheet_

### Parameters

_None_

---

## extractDataValidations()

---

```php
public function extractDataValidations(): void
```
_Extracts data validation rules from the sheet_

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
public function from(string $topLeftCell, ?bool $firstRowKeys = false): Sheet
```
_Set top left of read area. Alias of setReadArea()_

### Parameters

* `string $topLeftCell`
* `bool|null $firstRowKeys`

---

## getColAttributes()

---

```php
public function getColAttributes(): array
```
_Get column attributes_

### Parameters

_None_

---

## getColumnAttributes()

---

```php
public function getColumnAttributes($col): array|mixed
```
_Get column attributes (width, style, etc)_

### Parameters

* `int|string $col`

---

## getColumnStyle()

---

```php
public function getColumnStyle($col, ?bool $flat = false): array
```
_Get column style_

### Parameters

* `int|string $col`
* `bool|null $flat`

---

## getColumnWidth()

---

```php
public function getColumnWidth(int $colNumber): ?float
```
_Get width of the column_

### Parameters

* `int|string $colNumber`

---

## getConditionalFormatting()

---

```php
public function getConditionalFormatting(): array
```
_Get conditional formatting rules_

### Parameters

_None_

---

## getDataValidations()

---

```php
public function getDataValidations(): array
```
_Get data validation rules_

### Parameters

_None_

---

## getFreezePaneInfo()

---

```php
public function getFreezePaneInfo(): ?array
```
_Get freeze pane info_

### Parameters

_None_

---

## getImageBlob()

---

```php
public function getImageBlob(string $cell): ?string
```
_Get image content as binary string_

### Parameters

* `string $cell`

---

## getImageList()

---

```php
public function getImageList(): array
```
_Get image list_

### Parameters

_None_

---

## getImageListByRow()

---

```php
public function getImageListByRow($row): array
```
_Get image list by row number_

### Parameters

* `$row`

---

## getImageMimeType()

---

```php
public function getImageMimeType(string $cell): ?string
```
_Get image MIME typeRequires fileinfo extension_

### Parameters

* `string $cell`

---

## getImageName()

---

```php
public function getImageName(string $cell): ?string
```
_Get image name_

### Parameters

* `string $cell`

---

## getMergedCells()

---

```php
public function getMergedCells(): ?array
```
_Get merged cells. Returns an array \[min_cell => range]_

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

## getRowHeight()

---

```php
public function getRowHeight(int $rowNumber): ?float
```
_Get height of the row_

### Parameters

* `int $rowNumber`

---

## getTabColorConfiguration()

---

```php
public function getTabColorConfiguration(): ?array
```
_Get tab color configuration. Alias of getTabColorConfig()_

### Parameters

_None_

---

## getTabColorInfo()

---

```php
public function getTabColorInfo(): ?array
```
_Get the tab color info of the sheetContains any of: rgb, theme, tint, indexed_

### Parameters

_None_

---

## hasDrawings()

---

```php
public function hasDrawings(): bool
```
_Returns true if the sheet has drawings_

### Parameters

_None_

---

## hasImage()

---

```php
public function hasImage(string $cell): bool
```
_Returns TRUE if the cell contains an image_

### Parameters

* `string $cell`

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

## imageEntryFullPath()

---

```php
public function imageEntryFullPath(string $cell): ?string
```
_Get full path to the image in the ZIP archive_

### Parameters

* `string $cell`

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

## maxActualColumn()

---

```php
public function maxActualColumn(): string
```
_Get the last actual column letter_

### Parameters

_None_

---

## maxActualRow()

---

```php
public function maxActualRow(): int
```
_Get the last actual row number_

### Parameters

_None_

---

## maxColumn()

---

```php
public function maxColumn(?string $range = null): string
```
_Max column from dimension value_

### Parameters

* `string|null $range`

---

## maxRow()

---

```php
public function maxRow(?string $range = null): int
```
_Max row number from dimension value_

### Parameters

* `string|null $range`

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

## minActualColumn()

---

```php
public function minActualColumn(): string
```
_Get the first actual column letter_

### Parameters

_None_

---

## minActualRow()

---

```php
public function minActualRow(): int
```
_Get the first actual row number_

### Parameters

_None_

---

## minColumn()

---

```php
public function minColumn(?string $range = null): string
```
_Min column from dimension value_

### Parameters

* `string|null $range`

---

## minRow()

---

```php
public function minRow(?string $range = null): int
```
_Min row number from dimension value_

### Parameters

* `string|null $range`

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
_Get path to the sheet XML file in ZIP archive_

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

* `callback $callback` -- Callback function($row, $col, $value)
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

* `string|null $styleKey` -- If specified, only this style property will be returned (e.g. 'fill-color')

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
* `string|null $styleKey` -- If specified, only this style property will be returned (e.g. 'fill-color')

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
public function readFirstRowCellsFrom(string $areaRange, $columnKeys, 
                                      ?bool $styleIdxInclude = null): array
```
_Set read area and returns cell values of 1st row as array \[address => value]_

### Parameters

* `string $areaRange`
* `array|bool|int|null $columnKeys`
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

## saveImage()

---

```php
public function saveImage(string $cell, ?string $filename = null): ?string
```
_Save image to a file_

### Parameters

* `string $cell`
* `string|null $filename`

---

## saveImageTo()

---

```php
public function saveImageTo(string $cell, string $dirname): ?string
```
_Save image to a directory_

### Parameters

* `string $cell`
* `string $dirname`

---

## setDateFormat()

---

```php
public function setDateFormat($dateFormat): Sheet
```
_Set date format_

### Parameters

* `$dateFormat`

---

## setDefaultRowHeight()

---

```php
public function setDefaultRowHeight(float $rowHeight): void
```
_Set default row height_

### Parameters

* `float $rowHeight`

---

## setReadArea()

---

```php
public function setReadArea(string $areaRange, 
                            ?bool $firstRowKeys = false): Sheet
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
                                   ?bool $firstRowKeys = false): Sheet
```
_setReadArea('C:AZ') - set left and right columns of read areasetReadArea('C') - set left column only_

### Parameters

* `string $columnsRange`
* `bool|null $firstRowKeys`

---

## setState()

---

```php
public function setState(string $state): Sheet
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

