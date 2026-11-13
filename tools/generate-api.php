<?php

/** Regenerate API docs with the sibling avadim/autodoc checkout. */
$root = dirname(__DIR__);
$autoload = dirname($root) . '/autodoc/src/autoload.php';
if (!is_file($autoload)) {
    throw new RuntimeException('The sibling ../autodoc checkout is required');
}
require $autoload;
$classes = [];
foreach ([
    'Excel' => ['Excel', '91-api-class-excel.md'],
    'Sheet' => ['Sheet', '92-api-class-sheet.md'],
    'CsvReader' => ['Csv\\CsvReader', '94-api-class-csv-reader.md'],
    'CsvOptions' => ['Csv\\CsvOptions', '95-api-class-csv-options.md'],
    'XlsxBook' => ['Xlsx\\XlsxBook', '96-api-class-xlsx-book.md'],
    'OdsBook' => ['Ods\\OdsBook', '97-api-class-ods-book.md'],
    'OdsSheet' => ['Ods\\OdsSheet', '98-api-class-ods-sheet.md'],
] as $name => [$class, $file]) {
    $classes[$name] = ['name' => $class, 'file' => $file,
        'methods' => ['visibility' => ['public'], 'exclude' => ['__destruct', 'getColAttributes',
            'ownReader', 'releaseReader', 'ownTemporaryFile', 'releaseGenerator', 'assertOpen',
            'checkNode', 'openContent', 'dateValue']]];
}
$config = ['autoload' => '../vendor/autoload.php', 'dir' => '../docs',
    'main' => ['title' => 'FastExcelReader - API Reference', 'file' => '90-api-reference.md'],
    'namespace' => 'avadim\\FastExcelReader', 'classes' => $classes];
$file = __DIR__ . '/.api-reference.generated.json';
try {
    file_put_contents($file, json_encode($config, JSON_PRETTY_PRINT));
    $documentor = new avadim\SimplePhpDocumentor\Documentor($file);
    $documentor->generate();
    // Normalize generator output so newly added references pass git diff --check.
    foreach (array_merge(['90-api-reference.md'], array_column($classes, 'file')) as $reference) {
        $path = $root . '/docs/' . $reference;
        $text = preg_replace('/[\t ]+$/m', '', file_get_contents($path));
        file_put_contents($path, rtrim($text) . "\n");
    }
}
finally {
    if (is_file($file)) {
        unlink($file);
    }
}
