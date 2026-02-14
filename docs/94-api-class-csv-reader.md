# Class \avadim\FastExcelReader\Csv\CsvReader

---

* [__construct()](#__construct) – CsvReader constructor
* [close()](#close)
* [defaultErrorHandler()](#defaulterrorhandler)
* [fromCol()](#fromcol)
* [fromRow()](#fromrow)
* [getCsvField()](#getcsvfield) – Get next field from CSV file
* [getCsvLine()](#getcsvline) – Get line from CSV file as array of fields (null - empty field, false - EOF)
* [getOptions()](#getoptions)
* [nextRow()](#nextrow)
* [onError()](#onerror)
* [readRows()](#readrows) – Read rows and return as 2D array
* [rewind()](#rewind)
* [setBufferSize()](#setbuffersize)
* [withHeader()](#withheader) – Enables header mode

---

## __construct()

---

```php
public function __construct(string $file, $options)
```
_CsvReader constructor_

### Parameters

* `string $file`
* `CsvOptions|array|null $options`

---

## close()

---

```php
public function close()
```


### Parameters

_None_

---

## defaultErrorHandler()

---

```php
public function defaultErrorHandler(int $code, string $error, string $line, 
                                    int $lineNo, int $colNo)
```


### Parameters

* `int $code`
* `string $error`
* `string $line`
* `int $lineNo`
* `int $colNo`

---

## fromCol()

---

```php
public function fromCol(int $colNum): CsvReader
```


### Parameters

* `int $colNum`

---

## fromRow()

---

```php
public function fromRow(int $rowNum): CsvReader
```


### Parameters

* `int $rowNum`

---

## getCsvField()

---

```php
public function getCsvField(): ?string
```
_Get next field from CSV file_

### Parameters

_None_

---

## getCsvLine()

---

```php
public function getCsvLine(): array|null|false
```
_Get line from CSV file as array of fields (null - empty field, false - EOF)_

### Parameters

_None_

---

## getOptions()

---

```php
public function getOptions(): CsvOptions
```


### Parameters

_None_

---

## nextRow()

---

```php
public function nextRow($columnKeys, ?int $resultMode = null, 
                        ?int $rowLimit = 0): ?Generator
```


### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`
* `int|null $rowLimit`

---

## onError()

---

```php
public function onError(?callable $handler): CsvReader
```


### Parameters

* `$handler`

---

## readRows()

---

```php
public function readRows($columnKeys, ?int $resultMode = null): array
```
_Read rows and return as 2D array_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`

---

## rewind()

---

```php
public function rewind()
```


### Parameters

_None_

---

## setBufferSize()

---

```php
public function setBufferSize(int $size): CsvReader
```


### Parameters

* `int $size`

---

## withHeader()

---

```php
public function withHeader(): CsvReader
```
_Enables header mode_

_Treats the first row of the CSV file as a header row and returns subsequentrows as associative arrays keyed by column names_

### Parameters

_None_

---

