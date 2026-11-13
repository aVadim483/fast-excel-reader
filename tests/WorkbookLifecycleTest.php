<?php

declare(strict_types=1);

namespace avadim\FastExcelReader\Tests;

use avadim\FastExcelReader\Excel;
use avadim\FastExcelReader\Exception;
use avadim\FastExcelReader\Csv\CsvBook;
use avadim\FastExcelReader\Tests\Support\TemporaryFilesTestCase;

final class WorkbookLifecycleTest extends TemporaryFilesTestCase
{
    public static function formats(): array
    {
        return [['ooxml/family.xlsx'], ['ooxml/family.xlsm'], ['ooxml/family.xltx'], ['ooxml/family.xltm'], ['ods/family.ods'], ['xls/demo-00-test.xls'], ['test.csv']];
    }

    /** @dataProvider formats */
    public function testCloseReleasesTempAndInvalidatesSuspendedReader(string $fixture): void
    {
        $book = Excel::openString(file_get_contents(__DIR__ . '/test_files/' . $fixture));
        $sheet = $book->sheet();
        $iterator = $sheet->nextRow();
        $iterator->rewind();
        self::assertTrue($iterator->valid());
        self::assertNotEmpty(glob($this->directory . '/*'));
        $book->close();
        $book->close();
        self::assertTrue($book->isClosed());
        self::assertSame([], glob($this->directory . '/*'));
        try {
            $iterator->next();
            self::fail('A suspended reader must reject a closed workbook');
        }
        catch (Exception $e) {
            self::assertSame('The workbook is closed', $e->getMessage());
        }
        $this->expectException(Exception::class);
        $sheet->readRows();
    }

    public function testCallerStreamIsNotClosedOrRewound(): void
    {
        $input = fopen('php://memory', 'w+b');
        fwrite($input, "skip\na,b\n1,2\n");
        fseek($input, 5);
        $book = Excel::openStream($input);
        self::assertInstanceOf(CsvBook::class, $book);
        self::assertSame('a', $book->readRows()[1]['A']);
        $book->close();
        self::assertIsResource($input);
        self::assertSame(13, ftell($input));
        fclose($input);
    }

    public function testOpenCsvReturnsBookAndReaderRemainsExplicit(): void
    {
        $file = $this->temporaryFile("a;b\n1;2\n");
        $book = Excel::openCsv($file, ['delimiter' => ';']);
        self::assertInstanceOf(CsvBook::class, $book);
        self::assertSame(['A' => '1', 'B' => '2'], $book->readRows()[2]);
        $reader = Excel::openCsvReader($file, ['delimiter' => ';']);
        self::assertSame(['a', 'b'], $reader->getCsvLine());
        $reader->close();
        $book->close();
    }
}
