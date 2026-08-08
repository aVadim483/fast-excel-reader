# Продвинутое чтение

[🇬🇧 English](../12-advanced-reading.md) · [← К README](../../README.ru.md) · [Оглавление](../../README.ru.md#документация)

## Продвинутый пример
```php
use \avadim\FastExcelReader\Excel;

$file = __DIR__ . '/files/demo-02-advanced.xlsx';

$excel = Excel::open($file);

$result = [
    'sheets' => $excel->getSheetNames() // получить имена всех листов
];

$result['#1'] = $excel
    // выбрать лист по имени
    ->selectSheet('Demo1') 
    // выбрать область с данными, где первая строка содержит ключи колонок
    ->setReadArea('B4:D11', true)  
    // задать формат даты
    ->setDateFormat('Y-m-d') 
    // задать для колонки 'C' ключ 'Birthday'
    ->readRows(['C' => 'Birthday']); 

// читаем другие массивы с пользовательскими ключами колонок,
// и в этом случае мы задаём диапазон только по колонкам
$columnKeys = ['B' => 'year', 'C' => 'value1', 'D' => 'value2'];
$result['#2'] = $excel
    ->selectSheet('Demo2', 'B:D')
    ->readRows($columnKeys);

$result['#3'] = $excel
    ->setReadArea('F5:H13')
    ->readRows($columnKeys);
```
Область чтения можно задать по именованным диапазонам в книге. Например, если в книге есть
именованный диапазон **Headers** со ссылкой **Demo1!$B$4:$D$4**, то можно читать ячейки по этому имени

```php
$excel->setReadArea('Values');
$cells = $excel->readCells();
```
Обратите внимание: так как значение содержит имя листа, этот лист становится листом по умолчанию.

Область чтения можно задать и на самом листе
```php
$sheet = $excel->getSheet('Demo1')->setReadArea('Headers');
$cells = $sheet->readCells();
```
Но если попытаться использовать это имя на другом листе, вы получите ошибку
```php
$sheet = $excel->getSheet('Demo2')->setReadArea('Headers');
// Exception: Wrong address or range "Values"

```

При необходимости можно полностью контролировать процесс чтения с помощью метода ```readSheetCallback()``` с колбэк-функцией
```php
use \avadim\FastExcelReader\Excel;

$excel = Excel::open($file);

$result = [];
$excel->readCallback(function ($row, $col, $val) use(&$result) {
    // Любые манипуляции здесь
    $result[$row][$col] = (string)$val;

    // если функция вернёт true, чтение данных прерывается
    return false;
});
var_dump($result);
```

## Свойства документа

Вызовите `Excel::getProperties()`, чтобы прочитать свойства документа книги — базовые
(`docProps/core.xml`) и расширенные, прикладные (`docProps/app.xml`) — как единый ассоциативный
массив с нормализованными ключами:

```php
$props = $excel->getProperties();
```
Вы получите массив вида:
```text
Array
(
    [lastModifiedBy] => 'Chris Holland'
    [modified]       => '2025-02-21T00:23:41Z'
    [application]    => 'Microsoft Macintosh Excel'
    [appVersion]     => '16.0300'
)
```

Распознаются ключи `creator`, `lastModifiedBy`, `created`, `modified`, `title`, `subject`,
`description`, `keywords`, `category`, `revision`, `language` (базовые) и `application`, `appVersion`,
`company`, `manager` (прикладные). Возвращаются только реально присутствующие в файле свойства; книга
без части `docProps` возвращает пустой массив.

## Смотрите также

* [Чтение данных](11-reading-data.md) — построчно, ключи массивов, пустые ячейки
* [Типы значений и форматирование дат](13-dates-and-types.md)
* [Справочник API](../90-api-reference.md)
