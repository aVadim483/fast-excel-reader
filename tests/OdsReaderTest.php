<?php

declare(strict_types=1);

namespace avadim\FastExcelReader\Tests;

use avadim\FastExcelReader\Excel;
use avadim\FastExcelReader\Exception;
use avadim\FastExcelReader\Ods\OdsBook;
use avadim\FastExcelReader\Tests\Support\TemporaryFilesTestCase;

final class OdsReaderTest extends TemporaryFilesTestCase
{
    private function ods(string $tables, string $doctype = ''): string
    {
        $file = $this->directory . '/input.ods';
        $zip = new \ZipArchive();
        $zip->open($file, \ZipArchive::CREATE | \ZipArchive::OVERWRITE);
        $zip->addFromString('mimetype', 'application/vnd.oasis.opendocument.spreadsheet');
        $zip->addFromString('META-INF/manifest.xml', '<?xml version="1.0"?><manifest/>');
        $zip->addFromString('content.xml', '<?xml version="1.0"?>' . $doctype
            . '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"'
            . ' xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">'
            . '<office:body><office:spreadsheet>' . $tables . '</office:spreadsheet></office:body></office:document-content>');
        $zip->close();
        return $file;
    }

    private function table(string $rows, string $name = 'Data'): string
    {
        return '<table:table table:name="' . $name . '">' . $rows . '</table:table>';
    }

    private function stringCell(string $value, string $attributes = ''): string
    {
        return '<table:table-cell office:value-type="string" ' . $attributes . '><text:p>' . $value . '</text:p></table:table-cell>';
    }

    public function testRealExcelFileAndIndependentSheetCursors(): void
    {
        $file = __DIR__ . '/test_files/ods/family.ods';
        self::assertTrue(Excel::isOds($file));
        self::assertFalse(Excel::isXlsx($file));
        self::assertTrue(Excel::validate($file));
        $book = Excel::openOds($file);
        self::assertSame(['Second', 'ReaderData'], array_values($book->getSheetNames()));
        $sheet = $book->sheet('ReaderData');
        $rows = $sheet->nextRow();
        $rows->rewind();
        self::assertSame(['A' => 'ID', 'B' => 'Name'], $rows->current());
        self::assertSame('Second sheet', $book->sheet('Second')->readCells()['A1']);
        $rows->next();
        self::assertSame('Привет', $rows->current()['B']);
        self::assertSame(3, $rows->current()['D']);
        self::assertTrue($rows->current()['E']);
        self::assertTrue($sheet->isMerged('B4'));
        self::assertSame('A1:E4', $sheet->actualDimension());
        $book->close();
    }

    public function testRowsAreasHeadersAndKeyModes(): void
    {
        $file = $this->ods($this->table('<table:table-row>' . $this->stringCell('ID') . $this->stringCell('Name') . '</table:table-row>'
            . '<table:table-row table:number-rows-repeated="3"><table:table-cell office:value-type="float" office:value="7"/>' . $this->stringCell(' Alice ') . '</table:table-row>'));
        $book = Excel::open($file);
        self::assertInstanceOf(OdsBook::class, $book);
        $sheet = $book->sheet();
        self::assertCount(4, $sheet->readRows());
        self::assertSame([2 => ['ID' => 7, 'Name' => ' Alice '], 3 => ['ID' => 7, 'Name' => ' Alice '], 4 => ['ID' => 7, 'Name' => ' Alice ']], $sheet->withHeader()->readRows());
        self::assertSame([2 => [0 => 'Alice'], 3 => [0 => 'Alice']], $sheet->setReadArea('B2:B3')->readRows([], Excel::KEYS_COL_ZERO_BASED | Excel::TRIM_STRINGS));
        self::assertSame(['B2' => ' Alice ', 'B3' => ' Alice '], $sheet->readCells());
        $book->close();
    }

    public function testEmptyRepetitionsStayCompactAndKeepPhysicalCoordinates(): void
    {
        $rows = '<table:table-row table:number-rows-repeated="1000000"><table:table-cell table:number-columns-repeated="16384"/></table:table-row>'
            . '<table:table-row><table:table-cell table:number-columns-repeated="16383"/>' . $this->stringCell('last') . '</table:table-row>';
        $file = $this->ods($this->table($rows));
        $before = memory_get_usage(true);
        $started = microtime(true);
        $book = Excel::open($file);
        self::assertSame([1000001 => ['XFD' => 'last']], $book->readRows());
        self::assertSame('XFD1000001:XFD1000001', $book->sheet()->actualDimension());
        self::assertLessThan(8 * 1024 * 1024, memory_get_usage(true) - $before);
        self::assertLessThan(5.0, microtime(true) - $started);
        $book->close();
    }

    public function testNonemptyRepetitionIsClippedBeforeExpansion(): void
    {
        $file = $this->ods($this->table('<table:table-row table:number-rows-repeated="1000000">'
            . $this->stringCell('x', 'table:number-columns-repeated="16384"') . '</table:table-row>'));
        $book = Excel::open($file);
        self::assertSame([999999 => ['XFC' => 'x', 'XFD' => 'x']], $book->sheet()->readRowsFrom('XFC999999:XFD999999'));
        $book->close();
    }

    public function testRowLimitStopsBeforeConvertingNextRun(): void
    {
        foreach ([1, 3] as $repeat) {
            $file = $this->ods($this->table('<table:table-row table:number-rows-repeated="' . $repeat . '">'
                . '<table:table-cell office:value-type="float" office:value="1"/></table:table-row>'
                . '<table:table-row><table:table-cell office:value-type="float" office:value="invalid"/></table:table-row>'));
            $book = Excel::open($file);
            try {
                foreach (array_unique([1, 2, $repeat]) as $limit) {
                    if ($limit > $repeat) {
                        continue;
                    }
                    self::assertSame(
                        array_fill(1, $limit, ['A' => 1]),
                        iterator_to_array($book->sheet()->nextRow([], 0, null, $limit))
                    );
                }
            }
            finally {
                $book->close();
            }
        }
    }

    public function testTypedValuesTextAndDateContract(): void
    {
        $file = $this->ods($this->table('<table:table-row>'
            . '<table:table-cell office:value-type="date" office:date-value="2024-01-02T03:04:05Z"/>'
            . '<table:table-cell office:value-type="time" office:time-value="PT27H15M"/>'
            . '<table:table-cell office:value-type="boolean" office:boolean-value="false"/>'
            . '<table:table-cell office:value-type="float" office:value="1.25" table:formula="of:=1.25"/>'
            . $this->stringCell('A<text:s text:c="3"/>B<text:tab/>C<text:line-break/>D')
            . '</table:table-row>'));
        $book = Excel::open($file);
        self::assertSame([1 => ['A' => 1704164645, 'B' => 'PT27H15M', 'C' => false, 'D' => 1.25, 'E' => "A   B\tC\nD"]], $book->readRows());
        $book->dateFormatter(null);
        self::assertSame('2024-01-02T03:04:05Z', $book->readCells()['A1']);
        $book->dateFormatter('Y-m-d');
        self::assertSame('2024-01-02', $book->readCells()['A1']);
        $book->close();
    }

    public function testMergedRangeAndDimensionsDoNotExpandCoveredCells(): void
    {
        $file = $this->ods($this->table('<table:table-row>' . $this->stringCell('M', 'table:number-columns-spanned="3" table:number-rows-spanned="2"')
            . '<table:covered-table-cell table:number-columns-repeated="2"/></table:table-row><table:table-row/>'));
        $book = Excel::open($file);
        self::assertSame('A1:C2', $book->sheet()->mergedRange('C2'));
        self::assertSame('A1:C2', $book->sheet()->actualDimension());
        self::assertSame(['A1' => 'M'], $book->readCells());
        $book->close();
    }

    public function testCellsOutsideAreaAreNotConverted(): void
    {
        $file = $this->ods($this->table('<table:table-row><table:table-cell office:value-type="float" office:value="invalid"/></table:table-row>'
            . '<table:table-row><table:table-cell office:value-type="float" office:value="invalid"/>' . $this->stringCell('ok') . '</table:table-row>'));
        $book = Excel::open($file);
        self::assertSame(['B2' => 'ok'], $book->sheet()->readCellsFrom('B2:B2'));
        $book->close();
    }

    public function testDoctypeIsRejectedBeforeEntityExpansion(): void
    {
        $file = $this->ods($this->table('<table:table-row>' . $this->stringCell('&secret;') . '</table:table-row>'), '<!DOCTYPE x [<!ENTITY secret SYSTEM "file:///no-such-secret">]>');
        $this->expectException(Exception::class);
        $this->expectExceptionMessage('Document type');
        Excel::open($file);
    }

    public function testInvalidRepetitionIsRejected(): void
    {
        $file = $this->ods($this->table('<table:table-row table:number-rows-repeated="9999999999999999999999"/>'));
        $book = Excel::open($file);
        $this->expectException(Exception::class);
        $book->readRows();
    }

    public function testSpecialFilenameFallbackAndFormulaWithoutCachedResult(): void
    {
        $file = $this->ods($this->table('<table:table-row><table:table-cell table:formula="of:=1+1"/>' . $this->stringCell('ok') . '</table:table-row>'));
        $special = $this->directory . '/book#with%chars.ods';
        rename($file, $special);
        $book = Excel::open($special);
        self::assertSame(['A1' => null, 'B1' => 'ok'], $book->readCells());
        $book->close();
        self::assertSame([$special], glob($this->directory . '/*'));
    }

    public function testLibreOfficeDatesMatchXlsxRegardlessOfServerTimezone(): void
    {
        $timezone = date_default_timezone_get();
        date_default_timezone_set('Europe/Moscow');
        try {
            $xlsx = Excel::open(__DIR__ . '/../demo/files/demo-05-datetime.xlsx');
            $ods = Excel::open(__DIR__ . '/test_files/ods/demo-05-datetime.ods');
            $expected = $xlsx->readCells();
            $actual = $ods->readCells();
            // The first four date cells survive conversion unchanged; clock/duration cells differ by format.
            foreach (['A1', 'B1', 'A2', 'B2', 'A3', 'B3', 'A4', 'B4'] as $address) {
                self::assertSame($expected[$address], $actual[$address]);
            }
            $ods->close();
            $xlsx->close();
        }
        finally {
            date_default_timezone_set($timezone);
        }
    }
}
