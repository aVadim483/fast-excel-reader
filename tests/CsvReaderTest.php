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

    public function testCsvHeader()
    {
        $reader = new CsvReader($this->csvFile);
        $rows = $reader->readRows([], CsvOptions::KEYS_FIRST_ROW);

        $expected = ['ID' => '1', 'Name' => 'John; Doe', 'City' => 'New York'];
        $this->assertCount(3, $rows);
        $this->assertSame($expected, $rows[2]);

        $reader = new CsvReader($this->csvFile);
        $rows = $reader->readRows(true);
        $this->assertSame($expected, $rows[2]);

        $reader = new CsvReader($this->csvFile);
        $rows = $reader->withHeader()->readRows();
        $this->assertSame($expected, $rows[2]);
    }


}
