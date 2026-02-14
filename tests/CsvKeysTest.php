<?php

namespace avadim\FastExcelReader;

use avadim\FastExcelReader\Csv\CsvReader;
use avadim\FastExcelReader\Csv\CsvOptions;
use PHPUnit\Framework\TestCase;

class CsvKeysTest extends TestCase
{
    protected string $csvFile;

    protected function setUp(): void
    {
        $this->csvFile = __DIR__ . '/test_keys.csv';
        $content = "Name,Age,City\nJohn,30,New York\nJane,25,London";
        file_put_contents($this->csvFile, $content);
    }

    protected function tearDown(): void
    {
        if (file_exists($this->csvFile)) {
            unlink($this->csvFile);
        }
    }

    public function testKeysOriginal()
    {
        $reader = new CsvReader($this->csvFile, []);
        $rows = iterator_to_array($reader->nextRow([], CsvOptions::KEYS_ORIGINAL));
        
        $expected = [
            1 => [0 => 'Name', 1 => 'Age', 2 => 'City'],
            2 => [0 => 'John', 1 => '30', 2 => 'New York'],
            3 => [0 => 'Jane', 1 => '25', 2 => 'London'],
        ];
        $this->assertEquals($expected, $rows);
    }

    public function testKeysFirstRow()
    {
        $reader = new CsvReader($this->csvFile, []);
        $rows = iterator_to_array($reader->nextRow([], CsvOptions::KEYS_FIRST_ROW));
        
        $expected = [
            2 => ['Name' => 'John', 'Age' => '30', 'City' => 'New York'],
            3 => ['Name' => 'Jane', 'Age' => '25', 'City' => 'London'],
        ];
        $this->assertEquals($expected, $rows);
    }

    public function testKeysRowZeroBased()
    {
        $reader = new CsvReader($this->csvFile, []);
        $rows = iterator_to_array($reader->nextRow([], CsvOptions::KEYS_ROW_ZERO_BASED));
        
        $expected = [
            0 => [0 => 'Name', 1 => 'Age', 2 => 'City'],
            1 => [0 => 'John', 1 => '30', 2 => 'New York'],
            2 => [0 => 'Jane', 1 => '25', 2 => 'London'],
        ];
        
        $this->assertEquals($expected, $rows);
    }

    public function testKeysColZeroBased()
    {
        $reader = new CsvReader($this->csvFile, []);
        $rows = iterator_to_array($reader->nextRow([], CsvOptions::KEYS_COL_ZERO_BASED));
        
        $expected = [
            1 => [0 => 'Name', 1 => 'Age', 2 => 'City'],
            2 => [0 => 'John', 1 => '30', 2 => 'New York'],
            3 => [0 => 'Jane', 1 => '25', 2 => 'London'],
        ];
        $this->assertEquals($expected, $rows);
    }

    public function testKeysRowOneBased()
    {
        $reader = new CsvReader($this->csvFile, []);
        $rows = iterator_to_array($reader->nextRow([], CsvOptions::KEYS_ROW_ONE_BASED));
        
        $expected = [
            2 => [0 => 'Name', 1 => 'Age', 2 => 'City'],
            3 => [0 => 'John', 1 => '30', 2 => 'New York'],
            4 => [0 => 'Jane', 1 => '25', 2 => 'London'],
        ];
        $this->assertEquals($expected, $rows);
    }

    public function testKeysColOneBased()
    {
        $reader = new CsvReader($this->csvFile, []);
        $rows = iterator_to_array($reader->nextRow([], CsvOptions::KEYS_COL_ONE_BASED));
        
        $expected = [
            1 => [1 => 'Name', 2 => 'Age', 3 => 'City'],
            2 => [1 => 'John', 2 => '30', 3 => 'New York'],
            3 => [1 => 'Jane', 2 => '25', 3 => 'London'],
        ];
        $this->assertEquals($expected, $rows);
    }

    public function testKeysColExcel()
    {
        $reader = new CsvReader($this->csvFile, []);
        $rows = iterator_to_array($reader->nextRow([], CsvOptions::KEYS_COL_EXCEL));
        
        $expected = [
            1 => ['A' => 'Name', 'B' => 'Age', 'C' => 'City'],
            2 => ['A' => 'John', 'B' => '30', 'C' => 'New York'],
            3 => ['A' => 'Jane', 'B' => '25', 'C' => 'London'],
        ];
        $this->assertEquals($expected, $rows);
    }

    public function testKeysZeroBased()
    {
        $reader = new CsvReader($this->csvFile, []);
        $rows = iterator_to_array($reader->nextRow([], CsvOptions::KEYS_ZERO_BASED));
        
        // KEYS_ZERO_BASED = KEYS_ROW_ZERO_BASED | KEYS_COL_ZERO_BASED = 2 | 4 = 6
        $expected = [
            0 => [0 => 'Name', 1 => 'Age', 2 => 'City'],
            1 => [0 => 'John', 1 => '30', 2 => 'New York'],
            2 => [0 => 'Jane', 1 => '25', 2 => 'London'],
        ];
        $this->assertEquals($expected, $rows);
    }

    public function testKeysOneBased()
    {
        $reader = new CsvReader($this->csvFile, []);
        $rows = iterator_to_array($reader->nextRow([], CsvOptions::KEYS_ONE_BASED));
        
        // KEYS_ONE_BASED = KEYS_ROW_ONE_BASED | KEYS_COL_ONE_BASED = 8 | 16 = 24
        $expected = [
            2 => [1 => 'Name', 2 => 'Age', 3 => 'City'],
            3 => [1 => 'John', 2 => '30', 3 => 'New York'],
            4 => [1 => 'Jane', 2 => '25', 3 => 'London'],
        ];
        $this->assertEquals($expected, $rows);
    }
}
