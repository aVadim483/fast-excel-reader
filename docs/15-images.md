# Images

[← Back to README](../README.md) | [Documentation index](../README.md#documentation) | [🇷🇺 Русский](ru/15-images.md)

The library can define and extract images from XLSX files.

## Images functions
```php
// Returns count images on all sheets
$excel->countImages();

// Returns count images on sheet
$sheet->countImages();

// Returns image list of sheet
$sheet->getImageList();

// Returns image list of specified row
$sheet->getImageListByRow($rowNumber);

// Returns TRUE if the specified cell has an image
$sheet->hasImage($cellAddress);

// Returns mime type of image in the specified cell (or NULL)
$sheet->getImageMimeType($cellAddress);

// Returns inner name of image in the specified cell (or NULL)
$sheet->getImageName($cellAddress);

// Returns an image from the cell as a blob (if exists) or NULL
$sheet->getImageBlob($cellAddress);

// Writes an image from the cell to the specified filename
$sheet->saveImage($cellAddress, $fullFilenamePath);

// Writes an image from the cell to the specified directory
$sheet->saveImageTo($cellAddress, $fullDirectoryPath);
```

## See also

* [Reading Data](11-reading-data.md) — example of extracting images while looping rows
* [API Reference](90-api-reference.md)
