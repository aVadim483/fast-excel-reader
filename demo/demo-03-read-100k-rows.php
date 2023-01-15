<?php

include_once __DIR__ . '/../src/autoload.php';

$file = __DIR__ . '/files/demo-100k-rows.xlsx';

$timer = microtime(true);
$excel = \avadim\FastExcelReader\Excel::open($file);

$cnt = 0;
foreach ($excel->sheet()->nextRow() as $rowNum => $rowData) {
    $cnt++;
}

echo 'Read: ', $cnt, ' rows<br>';
echo 'Elapsed time: ', round(microtime(true) - $timer, 3), ' sec<br>';

// EOF