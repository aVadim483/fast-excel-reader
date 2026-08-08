<?php

use avadim\FastExcelReader\Excel;
use avadim\FastExcelReader\Exception;
use PHPUnit\Framework\TestCase;

/**
 * Excel::openString() and Excel::openStream(): a workbook read from memory or a
 * stream must be byte-for-byte identical to the same file read from disk, for
 * every format the signature dispatcher supports (XLSX, XLS, CSV).
 */
class OpenStreamStringTest extends TestCase
{
    private function assertSameAsFile(string $file, callable $openInMemory): void
    {
        self::assertFileExists($file);

        $expected = Excel::open($file)->readRows();
        $actual = $openInMemory(file_get_contents($file))->readRows();

        self::assertSame($expected, $actual);
    }

    public function testOpenStringXlsx()
    {
        $this->assertSameAsFile(
            __DIR__ . '/test_files/standard-file.xlsx',
            static fn(string $content) => Excel::openString($content)
        );
    }

    public function testOpenStreamXlsx()
    {
        $this->assertSameAsFile(
            __DIR__ . '/test_files/standard-file.xlsx',
            static function (string $content) {
                $stream = fopen('php://memory', 'r+b');
                fwrite($stream, $content);
                rewind($stream);
                $excel = Excel::openStream($stream);
                fclose($stream);

                return $excel;
            }
        );
    }

    public function testOpenStringDispatchesXls()
    {
        $this->assertSameAsFile(
            __DIR__ . '/test_files/xls/demo-00-test.xls',
            static fn(string $content) => Excel::openString($content)
        );
    }

    public function testOpenStringDispatchesCsv()
    {
        $this->assertSameAsFile(
            __DIR__ . '/test_files/test.csv',
            static fn(string $content) => Excel::openString($content)
        );
    }

    public function testOpenStreamDoesNotCloseCallerStream()
    {
        $file = __DIR__ . '/test_files/standard-file.xlsx';
        $stream = fopen($file, 'rb');
        Excel::openStream($stream);

        self::assertIsResource($stream, 'openStream() must not close the caller-owned stream');
        fclose($stream);
    }

    public function testEmptyStringRejected()
    {
        $this->expectException(Exception::class);
        Excel::openString('');
    }

    public function testNonResourceStreamRejected()
    {
        $this->expectException(Exception::class);
        /** @noinspection PhpParamsInspection */
        Excel::openStream('not a resource');
    }
}
