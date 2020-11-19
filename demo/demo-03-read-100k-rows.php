<?php

include_once __DIR__ . '/../src/autoload.php';

$file = __DIR__ . '/files/demo-100k-rows.xlsx';

$timer = microtime(true);
$excel = \avadim\FastExcelReader\Excel::open($file);

$cnt = 0;
$excel->readSheetCallback(static function ($row, $col, $val) use(&$cnt) {
    $cnt = $row;
});

echo 'Read: ', $cnt, ' rows<br>';
echo 'elapsed time: ', round(microtime(true) - $timer, 3), ' sec';

// EOF