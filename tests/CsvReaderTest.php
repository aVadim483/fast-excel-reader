<?php

namespace avadim\FastExcelReader;

use avadim\FastExcelReader\Csv\CsvOptions;
use avadim\FastExcelReader\Csv\CsvReader;
use PHPUnit\Framework\TestCase;

class CsvReaderTest extends TestCase
{
    protected string $csvFile;

    protected function setUp(): void
    {
        $this->csvFile = __DIR__ . '/test.csv';
        $content = "ID;Name;City\r\n1;\"John; Doe\";New York\n2;Jane;London\r3;\"Josse, Dominique\";Brussels";
        file_put_contents($this->csvFile, $content);
    }

    protected function tearDown(): void
    {
        if (isset($this->csvFile) && is_file($this->csvFile)) {
            @unlink($this->csvFile);
        }
    }

    public function testCsvOptions()
    {
        $options = new CsvOptions();
        $options->setDelimiter(';')
            ->setEnclosure('"');

        $reader = new CsvReader($this->csvFile, $options);
        $rows = $reader->readRows();

        $this->assertCount(4, $rows);
        $this->assertEquals('John; Doe', $rows[2][1]);
        $this->assertEquals('New York', $rows[2][2]);
    }

    public function testCsvOptionsArray()
    {
        $options = [
            'delimiter' => ';',
            'enclosure' => '"',
        ];

        $reader = new CsvReader($this->csvFile, $options);
        $rows = $reader->readRows();

        $this->assertCount(4, $rows);
        $this->assertEquals('John; Doe', $rows[2][1]);
    }

    public function testCsvOptionsStaticCreate()
    {
        $options = CsvOptions::create([
            'delimiter' => ';',
            'enclosure' => '"',
        ]);

        $reader = new CsvReader($this->csvFile, $options);
        $rows = $reader->readRows();

        $this->assertCount(4, $rows);
        $this->assertEquals('John; Doe', $rows[2][1]);
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

        $this->assertCount(4, $rows);
        $this->assertEquals('John; Doe', $rows[2][1]);
        $this->assertEquals('New York', $rows[2][2]);
    }

    public function testCsvOptionsToArray()
    {
        $options = [
            'delimiter' => ';',
            'enclosure' => '"',
            'escape' => '\\',
            'encoding' => 'UTF-8',
            'mode' => CsvOptions::STRICT_MODE,
            'double_quotes' => true,
            'trim_fields' => true,
            'skip_empty_lines' => true,
            'comment_prefix' => null,
        ];

        $csvOptions = new CsvOptions($options);
        $expected = $options;
        $expected['stream_filter'] = null;
        $this->assertEquals($expected, $csvOptions->toArray());
    }

    public function testCsvStreamFilterOption()
    {
        $options = new CsvOptions();
        $this->assertNull($options->stream_filter);
        $this->assertNull($options->streamFilter);

        $options->setStreamFilter('string.rot13');
        $this->assertEquals('string.rot13', $options->stream_filter);
        $this->assertEquals('string.rot13', $options->streamFilter);

        $reader = new CsvReader($this->csvFile, $options);
        $this->assertEquals('string.rot13', $reader->getOptions()->stream_filter);
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
            $this->assertSame($expectedStrict, $strictRow, $note . ' (STRICT)');
        }


        // lenient
        $lenientReader = $this->makeReader($input, ['mode' => 'tolerant', "escape" => '\\']);
        $lenientRow = $lenientReader->getCsvLine();
        $this->assertSame($expectedLenient, $lenientRow, $note . ' (TOLERANT)');
    }

}
