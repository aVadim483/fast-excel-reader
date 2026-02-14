# Class \avadim\FastExcelReader\Csv\CsvOptions

---

* [__construct()](#__construct) – CsvOptions constructor
* [__get()](#__get) – Magic getter for options
* [__isset()](#__isset) – Magic isset for options
* [__set()](#__set) – Magic setter for options
* [create()](#create) – Create CsvOptions instance
* [setCommentPrefix()](#setcommentprefix) – Set comment prefix
* [setDelimiter()](#setdelimiter) – Set column delimiter character (null for auto-detect)
* [setDoubleQuotes()](#setdoublequotes) – Set whether to handle double quotes
* [setEnclosure()](#setenclosure) – Set enclosure character of fields
* [setEncoding()](#setencoding) – Set input file encoding (null = auto)
* [setEscape()](#setescape) – Set escape character, usually '\' or '' ('' or null for no escape)
* [setMode()](#setmode) – Set parsing mode (strict or tolerant)
* [setSkipEmptyLines()](#setskipemptylines) – Set whether to skip empty lines
* [setStreamFilter()](#setstreamfilter) – Set stream filter
* [setTrimFields()](#settrimfields) – Set whether to trim fields (does not affect spaces inside quotes)
* [toArray()](#toarray) – Return all options as array

---

## __construct()

---

```php
public function __construct(array $options = [])
```
_CsvOptions constructor_

### Parameters

* `array $options`

---

## __get()

---

```php
public function __get($name): mixed|null
```
_Magic getter for options_

### Parameters

* `string $name`

---

## __isset()

---

```php
public function __isset($name): bool
```
_Magic isset for options_

### Parameters

* `string $name`

---

## __set()

---

```php
public function __set($name, $value): void
```
_Magic setter for options_

### Parameters

* `string $name`
* `mixed $value`

---

## create()

---

```php
public static function create(array $options = []): CsvOptions
```
_Create CsvOptions instance_

### Parameters

* `array $options`

---

## setCommentPrefix()

---

```php
public function setCommentPrefix(?string $value): CsvOptions
```
_Set comment prefix_

### Parameters

* `string|null $value`

---

## setDelimiter()

---

```php
public function setDelimiter(?string $delimiter): CsvOptions
```
_Set column delimiter character (null for auto-detect)_

### Parameters

* `string|null $delimiter`

---

## setDoubleQuotes()

---

```php
public function setDoubleQuotes(bool $enable): CsvOptions
```
_Set whether to handle double quotes_

### Parameters

* `bool $enable`

---

## setEnclosure()

---

```php
public function setEnclosure(string $enclosure): CsvOptions
```
_Set enclosure character of fields_

### Parameters

* `string $enclosure`

---

## setEncoding()

---

```php
public function setEncoding(string $encoding): CsvOptions
```
_Set input file encoding (null = auto)_

### Parameters

* `string $encoding`

---

## setEscape()

---

```php
public function setEscape(string $escape): CsvOptions
```
_Set escape character, usually '\' or '' ('' or null for no escape)_

### Parameters

* `string $escape`

---

## setMode()

---

```php
public function setMode(string $mode): CsvOptions
```
_Set parsing mode (strict or tolerant)_

### Parameters

* `string $mode`

---

## setSkipEmptyLines()

---

```php
public function setSkipEmptyLines(bool $enable): CsvOptions
```
_Set whether to skip empty lines_

### Parameters

* `bool $enable`

---

## setStreamFilter()

---

```php
public function setStreamFilter(?string $filter): CsvOptions
```
_Set stream filter_

### Parameters

* `string|null $filter`

---

## setTrimFields()

---

```php
public function setTrimFields(bool $enable): CsvOptions
```
_Set whether to trim fields (does not affect spaces inside quotes)_

### Parameters

* `bool $enable`

---

## toArray()

---

```php
public function toArray(): array
```
_Return all options as array_

### Parameters

_None_

---

