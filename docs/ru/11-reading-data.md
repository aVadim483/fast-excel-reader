# Чтение данных

[🇬🇧 English](../11-reading-data.md) · [← К README](../../README.ru.md) · [Оглавление](../../README.ru.md#документация)

* [Чтение значений построчно в цикле](#чтение-значений-построчно-в-цикле)
* [Ключи в результирующих массивах](#ключи-в-результирующих-массивах)
* [Пустые ячейки и строки](#пустые-ячейки-и-строки)

## Чтение значений построчно в цикле
```php
$sheet = $excel->sheet();
foreach ($sheet->nextRow() as $rowNum => $rowData) {
    // $rowData — это массив ['A' => ..., 'B' => ...]
    $addr = 'C' . $rowNum;
    if ($sheet->hasImage($addr)) {
        $sheet->saveImageTo($addr, $fullDirectoryPath);
    }
    // обработка $rowData здесь
    // ...
}

// ИЛИ
foreach ($sheet->nextRow() as $rowNum => $rowData) {
    // обработка $rowData здесь
    // ...
    // получить список изображений из текущей строки
    $imageList = $sheet->getImageListByRow();
    foreach ($imageList as $imageInfo) {
        $imageBlob = $sheet->getImageBlob($imageInfo['address']);
    }
}

// ИЛИ
foreach ($sheet->nextRow(['A' => 'One', 'B' => 'Two'], Excel::KEYS_FIRST_ROW) as $rowNum => $rowData) {
    // $rowData — это массив ['One' => ..., 'Two' => ...]
    // ...
}
```
ПРИМЕЧАНИЕ: каждый раз, когда вы запускаете цикл ```foreach ($sheet->nextRow() as $rowIndex => $row)```,
чтение данных начинается с первой строки.

Но есть и альтернативный способ читать построчно — через метод readNextRow().
В этом случае сначала нужно вызвать метод ```$sheet->reset(...)``` с нужными параметрами чтения,
а затем можно вызывать `````$sheet-readNextRow()`````. Если в какой-то момент нужно начать читать данные
сначала, снова вызовите ```$sheet->reset(...)```.

```php
// Инициализировать внутренний генератор чтения
$sheet->reset(['A' => 'One', 'B' => 'Two'], Excel::KEYS_FIRST_ROW);

// прочитать первую строку
$rowData = $sheet->readNextRow();
var_dump($rowData);

// Прочитать следующие 3 строки
for ($i = 0; $i < 3; $i++) {
    $rowData = $sheet->readNextRow();
    var_dump($rowData);
}

// Сбросить внутренний генератор и прочитать все строки, начиная с первой
$sheet->reset(['A' => 'One', 'B' => 'Two'], Excel::KEYS_FIRST_ROW);
$result = [];
while ($rowData = $sheet->readNextRow()) {
    $result[] = $rowData;
}
var_dump($result);
```

## Ключи в результирующих массивах
```php
// Прочитать строки и использовать первую строку как ключи колонок
$result = $excel->readRows(true);

// То же самое в декларативной форме
$result = $excel->sheet()->withHeader()->readRows();

// Пропустить строку заголовков, но задать имена столбцов самостоятельно, по порядку
$result = $excel->sheet()->withHeader(['col1', 'col2'])->readRows();
```
Имена, переданные в `withHeader()`, позиционные: первое имя достаётся первому столбцу области чтения,
поэтому буквы столбцов не нужны, и один и тот же вызов работает на листе, где данные начинаются не с
`A1`. Более короткий список переименует только покрытые им столбцы; остальные сохранят имя из строки
заголовков.

Вы получите такой результат:
```text
Array
(
    [2] => Array
        (
            ['col1'] => 111
            ['col2'] => 'aaa'
        )
    [3] => Array
        (
            ['col1'] => 222
            ['col2'] => 'bbb'
        )
)
```
Необязательный второй аргумент задаёт ключи результирующего массива
```php

// Строки и колонки начинаются с нуля
$result = $excel->readRows(false, Excel::KEYS_ZERO_BASED);
```
Вы получите такой результат:
```text
Array
(
    [0] => Array
        (
            [0] => 'col1'
            [1] => 'col2'
        )
    [1] => Array
        (
            [0] => 111
            [1] => 'aaa'
        )
    [2] => Array
        (
            [0] => 222
            [1] => 'bbb'
        )
)
```
Допустимые значения режима результата

| режим               | описание                                                                        |
|---------------------|---------------------------------------------------------------------------------|
| KEYS_ORIGINAL       | строки с '1', колонки с 'A' (по умолчанию)                                       |
| KEYS_ROW_ZERO_BASED | строки с 0                                                                       |
| KEYS_COL_ZERO_BASED | колонки с 0                                                                      |
| KEYS_ZERO_BASED     | строки с 0, колонки с 0 (то же, что KEYS_ROW_ZERO_BASED + KEYS_COL_ZERO_BASED)   |
| KEYS_ROW_ONE_BASED  | строки с 1                                                                       |
| KEYS_COL_ONE_BASED  | колонки с 1                                                                      |
| KEYS_ONE_BASED      | строки с 1, колонки с 1 (то же, что KEYS_ROW_ONE_BASED + KEYS_COL_ONE_BASED)     |

Дополнительные опции, которые можно комбинировать с режимами результата

| опция           | описание                                       |
|-----------------|------------------------------------------------|
| KEYS_FIRST_ROW  | то же, что _true_ в первом аргументе           |
| KEYS_RELATIVE   | индекс от верхней левой ячейки области (не листа) |
| KEYS_SWAP       | поменять местами строки и колонки              |

Например
```php

$result = $excel->readRows(['A' => 'bee', 'B' => 'honey'], Excel::KEYS_FIRST_ROW | Excel::KEYS_ROW_ZERO_BASED);
```
Вы получите такой результат:
```text
Array
(
    [0] => Array
        (
            [bee] => 111
            [honey] => 'aaa'
        )

    [1] => Array
        (
            [bee] => 222
            [honey] => 'bbb'
        )

)
```

## Пустые ячейки и строки

Библиотека по умолчанию уже пропускает пустые ячейки и пустые строки. Пустые ячейки — это ячейки,
в которые ничего не записано, а пустые строки — строки, в которых все ячейки пусты. Если ячейка
содержит пустую строку, она не считается пустой. Но вы можете изменить это поведение и пропускать
ячейки с пустыми строками.

```php
$sheet = $excel->sheet();

// Пропускать пустые ячейки и пустые строки
foreach ($sheet->nextRow() as $rowNum => $rowData) {
    // обработать $rowData
}

// Пропускать пустые ячейки и ячейки с пустыми строками
foreach ($sheet->nextRow([], Excel::TREAT_EMPTY_STRING_AS_EMPTY_CELL) as $rowNum => $rowData) {
    // обработать $rowData
}

// Пропускать пустые ячейки и пустые строки (строки только из пробельных символов тоже считаются пустыми)
foreach ($sheet->nextRow([], Excel::TRIM_STRINGS | Excel::TREAT_EMPTY_STRING_AS_EMPTY_CELL) as $rowNum => $rowData) {
    // обработать $rowData
}
```
Другой способ
```php
$sheet->reset([], Excel::TRIM_STRINGS | Excel::TREAT_EMPTY_STRING_AS_EMPTY_CELL);
$rowData = $sheet->readNextRow();
// сделать что-нибудь

$rowData = $sheet->readNextRow();
// обработать следующую строку

// ...
```

## Смотрите также

* [Начало работы](10-getting-started.md)
* [Продвинутое чтение](12-advanced-reading.md) — области чтения, именованные диапазоны, колбэки
* [Справочник API](../90-api-reference.md)
