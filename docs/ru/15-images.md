# Изображения

[🇬🇧 English](../15-images.md) · [← К README](../../README.ru.md) · [Оглавление](../../README.ru.md#документация)

Библиотека умеет находить и извлекать изображения из XLSX-файлов.

## Функции для работы с изображениями
```php
// Возвращает количество изображений на всех листах
$excel->countImages();

// Возвращает количество изображений на листе
$sheet->countImages();

// Возвращает список изображений листа
$sheet->getImageList();

// Возвращает список изображений указанной строки
$sheet->getImageListByRow($rowNumber);

// Возвращает TRUE, если в указанной ячейке есть изображение
$sheet->hasImage($cellAddress);

// Возвращает mime-тип изображения в указанной ячейке (или NULL)
$sheet->getImageMimeType($cellAddress);

// Возвращает внутреннее имя изображения в указанной ячейке (или NULL)
$sheet->getImageName($cellAddress);

// Возвращает изображение из ячейки как blob (если есть) или NULL
$sheet->getImageBlob($cellAddress);

// Записывает изображение из ячейки в указанный файл
$sheet->saveImage($cellAddress, $fullFilenamePath);

// Записывает изображение из ячейки в указанную директорию
$sheet->saveImageTo($cellAddress, $fullDirectoryPath);
```

## Смотрите также

* [Чтение данных](11-reading-data.md) — пример извлечения изображений при обходе строк
* [Справочник API](../90-api-reference.md)
