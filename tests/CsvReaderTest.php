<?php

namespace avadim\FastExcelReader;

use PHPUnit\Framework\TestCase;
use Throwable;

class CsvReaderTest extends TestCase
{
    protected string $csvFile;

    protected function setUp(): void
    {
        $this->csvFile = __DIR__ . '/test_files/test.csv';
    }

    protected function tearDown(): void
    {
    }

    public function testCsvOptions()
    {
        $options = new CsvOptions();
        $options->setDelimiter(';')
            ->setEnclosure('"');

        $reader = new CsvReader($this->csvFile, $options);
        $rows = $reader->readRows();

        $this->assertCount(3, $rows);
        $this->assertEquals('John; Doe', $rows[2]['B']);
        $this->assertEquals('New York', $rows[2]['C']);
    }

    public function testCsvOptionsArray()
    {
        $options = [
            'delimiter' => ';',
            'enclosure' => '"',
        ];

        $reader = new CsvReader($this->csvFile, $options);
        $rows = $reader->readRows();

        $this->assertCount(3, $rows);
        $this->assertEquals('John; Doe', $rows[2]['B']);
    }

    public function testCsvOptionsStaticCreate()
    {
        $options = CsvOptions::create([
            'delimiter' => ';',
            'enclosure' => '"',
        ]);

        $reader = new CsvReader($this->csvFile, $options);
        $rows = $reader->readRows();

        $this->assertCount(3, $rows);
        $this->assertEquals('John; Doe', $rows[2]['B']);
    }

    public function testCsvWithBom()
    {
        $file = __DIR__ . '/test_files/test-bom.csv';
        // Use Excel::KEYS_FIRST_ROW to use the first row as keys
        $reader = new CsvReader($file, ['delimiter' => ';']);
        $rows = [];
        // nextRow is a generator, we need to collect results
        foreach ($reader->nextRow([], Excel::KEYS_FIRST_ROW) as $rowNum => $row) {
            $rows[$rowNum] = $row;
        }

        $this->assertCount(2, $rows); // 2 data rows, header excluded from result
        // If BOM is handled correctly, the first key will be "ID"
        $firstRow = reset($rows);
        $keys = array_keys($firstRow);
        $this->assertEquals('ID', $keys[0]);
        $this->assertEquals('1', $firstRow['ID']);
    }

    public function testCsvAutoDelimiter()
    {
        $reader = new CsvReader($this->csvFile, ['delimiter' => 'auto']);
        $rows = $reader->readRows();

        $this->assertCount(3, $rows);
        $this->assertEquals('John; Doe', $rows[2]['B']);
        $this->assertEquals('New York', $rows[2]['C']);
    }

    protected function makeReader($input, $options = [])
    {
        $file = __DIR__ . '/test_files/test_strict.csv';
        file_put_contents($file, $input);

        return new CsvReader($file, $options);
    }

    /**
     * @dataProvider \CsvDataProvider::provideCsvRecords
     */
    public function testParseCsvRecord(string $input, ?array $expectedStrict, ?array $expectedLenient, string $note)
    {
        // strict
        if ($expectedStrict === null) {
            $this->expectException(\RuntimeException::class);
        }
        $strictReader = $this->makeReader($input, ['mode' => 'strict', 'trim_fields' => false]);

        echo '# ' . $note;
        if ($expectedStrict === null) {
            $this->expectException(Exception::class);
        }
        $strictRow = $strictReader->getCsvLine();
        if ($expectedStrict !== null) {
            $this->assertSame($expectedStrict, $strictRow, $note . ' (strict)');
        }

        // lenient
        $lenientReader = $this->makeReader($input, ['mode' => 'tolerant', "escape" => '\\']);
        $lenientRow = $lenientReader->getCsvLine();
        $this->assertSame($expectedLenient, $lenientRow, $note . ' (lenient)');
    }

}
