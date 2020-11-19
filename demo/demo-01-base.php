<?php

include_once __DIR__ . '/../src/autoload.php';

$file = __DIR__ . '/files/demo-01-base.xlsx';

$timer = microtime(true);
$excel = \avadim\FastExcelReader\Excel::open($file);

$result = $excel->readRows(true);

echo '<pre>', print_r($result);

// EOF