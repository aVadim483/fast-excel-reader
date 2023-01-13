<?php

declare(strict_types=1);

namespace avadim\FastExcelReader;

use PHPUnit\Framework\TestCase;

final class FastExcelReaderTest extends TestCase
{
    public function testExcelReader()
    {
        $file = __DIR__ . '/../demo/files/demo-00-test.xlsx';

        $excel = Excel::open($file);

        $result = $excel->readCells();
        $this->assertTrue(isset($result['A1']) && $result['A1'] === 'number');
        $this->assertTrue(isset($result['B3']) && $result['B3'] === 'Word');

        $result = $excel->readRows();
        $this->assertTrue(isset($result['1']['A']) && $result['1']['A'] === 'number');
        $this->assertTrue(isset($result['3']['B']) && $result['3']['B'] === 'Word');

        $result = $excel->readColumns();
        $this->assertTrue(isset($result['A']['1']) && $result['A']['1'] === 'number');
        $this->assertTrue(isset($result['B']['3']) && $result['B']['3'] === 'Word');

        // Read rows and use the first row as column keys
        $result = $excel->readRows(true);
        $this->assertTrue(isset($result['2']['number']) && $result['2']['number'] === 111);
        $this->assertTrue(isset($result['3']['name']) && $result['3']['name'] === 'Word');

        $result = $excel->readRows(false, Excel::KEYS_ZERO_BASED);
        $this->assertTrue(isset($result[0][0]) && $result[0][0] === 'number');
        $this->assertTrue(isset($result[2][1]) && $result[2][1] === 'Word');

        $result = $excel->readRows(false, Excel::KEYS_ONE_BASED);
        $this->assertTrue(isset($result[1][1]) && $result[1][1] === 'number');
        $this->assertTrue(isset($result[3][2]) && $result[3][2] === 'Word');

        $result = $excel->readRows(['A' => 'bee', 'B' => 'honey'], Excel::KEYS_FIRST_ROW | Excel::KEYS_ROW_ZERO_BASED);
        $this->assertTrue(isset($result[0]['bee']) && $result[0]['bee'] === 111);
        $this->assertTrue(isset($result[1]['honey']) && $result[1]['honey'] === 'Word');
    }
}

