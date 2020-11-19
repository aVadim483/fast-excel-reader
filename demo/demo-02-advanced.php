<?php

include_once __DIR__ . '/../src/autoload.php';

$file = __DIR__ . '/files/demo-02-advanced.xlsx';

$timer = microtime(true);
$excel = \avadim\FastExcelReader\Excel::open($file);

$result = [];

$result['sheets'] = $excel->getSheetNames();

$result['#1'] = $excel
    ->selectSheet('Demo1')
    ->setReadArea('B4:D11', true)
    ->setDateFormat('Y-m-d')
    ->readRows(['C' => 'Birthday']);

$columnKeys = ['B' => 'year', 'C' => 'value1', 'D' => 'value2'];
$data2 = $excel
    ->selectSheet('Demo2', 'B5:D13')
    ->readRows($columnKeys);

$data3 = $excel
    ->setReadArea('F5:H13')
    ->readRows($columnKeys);

$result['#2'] = array_merge($data2, $data3);

echo '<pre>', print_r($result);

// EOF