<?php

declare(strict_types=1);

namespace avadim\FastExcelReader\Tests;

use avadim\FastExcelReader\Excel;
use avadim\FastExcelReader\Exception;
use avadim\FastExcelReader\Reader;
use avadim\FastExcelReader\Internal\StreamIO;
use avadim\FastExcelReader\Tests\Support\FailureStream;
use avadim\FastExcelReader\Tests\Support\TemporaryFilesTestCase;

final class StreamFailureTest extends TemporaryFilesTestCase
{
    protected function setUp(): void
    {
        parent::setUp();
        stream_wrapper_register('fxrfailure', FailureStream::class);
        FailureStream::$closed = [];
        FailureStream::$written = [];
    }

    protected function tearDown(): void
    {
        stream_wrapper_unregister('fxrfailure');
        parent::tearDown();
    }

    /** @dataProvider failedInputs */
    public function testFailedInputDoesNotProducePartialCsvOrLeaveTemp(string $mode): void
    {
        $input = fopen('fxrfailure://' . $mode, 'rb');
        try {
            Excel::openStream($input);
            self::fail('An incomplete CSV must not be returned');
        }
        catch (Exception $e) {
            self::assertNotNull($e->getPrevious());
            self::assertSame([], glob($this->directory . '/*'));
            self::assertIsResource($input);
            self::assertNotContains($mode, FailureStream::$closed);
        }
        finally {
            fclose($input);
        }
    }

    public static function failedInputs(): array
    {
        return [['fail-read'], ['throw-read']];
    }

    public function testNonSeekableInputIsAccepted(): void
    {
        $input = fopen('fxrfailure://valid', 'rb');
        try {
            self::assertSame([1 => ['A' => 'a', 'B' => 'b'], 2 => ['A' => '1', 'B' => '2']], Excel::openStream($input)->readRows());
            self::assertIsResource($input);
        }
        finally {
            fclose($input);
        }
    }

    public function testCurrentPositionIsRespected(): void
    {
        $input = fopen('php://memory', 'w+b');
        fwrite($input, "skip\na,b\n1,2\n");
        fseek($input, 5);
        try {
            self::assertSame(['A' => 'a', 'B' => 'b'], Excel::openStream($input)->readRows()[1]);
            self::assertIsResource($input);
        }
        finally {
            fclose($input);
        }
    }

    public function testNonStreamResourceIsRejectedBeforeCreatingTemp(): void
    {
        try {
            Excel::openStream(stream_context_create());
            self::fail('Stream contexts are not readable streams');
        }
        catch (Exception $e) {
            self::assertSame([], glob($this->directory . '/*'));
        }
    }

    public function testEmptyInputIsCleanedUp(): void
    {
        $input = fopen('php://memory', 'rb');
        try {
            Excel::openStream($input);
            self::fail('Empty stream must be rejected');
        }
        catch (Exception $e) {
            self::assertSame([], glob($this->directory . '/*'));
            self::assertIsResource($input);
        }
        finally {
            fclose($input);
        }
    }

    /** @dataProvider writeFailures */
    public function testIncompleteStringWriteIsRejected(string $mode): void
    {
        try {
            StreamIO::writeString('fxrfailure://' . $mode, 'abcdefgh');
            self::fail('Write failure must be reported');
        }
        catch (Exception $e) {
            self::assertNotNull($e->getPrevious());
            if ($mode !== 'deny-open') {
                self::assertContains($mode, FailureStream::$closed);
            }
        }
    }

    public static function writeFailures(): array
    {
        return [['short-write'], ['throw-write'], ['deny-open'], ['throw-flush'], ['throw-close']];
    }

    /** @dataProvider copyFailures */
    public function testCopyClosesOwnedOutputOnWriteFailure(string $mode): void
    {
        $input = fopen('php://memory', 'w+b');
        fwrite($input, 'abcdefgh');
        rewind($input);
        try {
            StreamIO::copyToFile($input, 'fxrfailure://' . $mode);
            self::fail('Copy failure must be reported');
        }
        catch (Exception $e) {
            self::assertIsResource($input);
            self::assertNotNull($e->getPrevious());
            if ($mode !== 'deny-open') {
                self::assertContains($mode, FailureStream::$closed);
            }
        }
        finally {
            fclose($input);
        }
    }

    public static function copyFailures(): array
    {
        return [['short-write'], ['throw-write'], ['deny-open'], ['fail-flush'], ['throw-flush'], ['throw-close']];
    }

    public function testOpenFailureDeletesWholeWorkbookTemp(): void
    {
        try {
            Excel::openString("PK\x03\x04not-a-zip");
            self::fail('Invalid ZIP must be rejected');
        }
        catch (Exception $e) {
            self::assertSame([], glob($this->directory . '/*'));
        }
    }

    public function testFailedWorkbookPreparationRemovesTempBeforeGc(): void
    {
        $file = $this->workbook(['xl/_rels/workbook.xml.rels' => '<Relationships><Relationship Id="bad" Type="worksheet" Target="../../outside.xml"/></Relationships>']);
        try {
            Excel::openString(file_get_contents($file));
            self::fail('Invalid internal target must be rejected');
        }
        catch (Exception $e) {
            self::assertSame([], glob($this->directory . '/xlsx_reader_*'));
        }
    }

    public function testAlterModePreservesParserPropertiesAndDeletesTemp(): void
    {
        $file = $this->workbook(['custom.xml' => '<!DOCTYPE root [<!ENTITY value "expanded">]><root>&value;</root>']);
        $reader = new class($file, [\XMLReader::SUBST_ENTITIES => true]) extends Reader {
            protected bool $alterMode = true;
        };
        try {
            self::assertTrue($reader->openZip('custom.xml'));
            self::assertTrue($reader->getParserProperty(\XMLReader::SUBST_ENTITIES));
            $reader->seekOpenTag('root');
            self::assertSame('expanded', $reader->readString());
        }
        finally {
            $reader->close();
        }
        self::assertSame([], glob($this->directory . '/xlsx_reader_*'));
    }

    public function testAlterModeFailureReleasesArchiveAndOutput(): void
    {
        $file = $this->workbook();
        $reader = new class($file) extends Reader {
            protected bool $alterMode = true;
            protected function makeTempFile()
            {
                throw new Exception('Injected temporary file failure');
            }
        };
        try {
            $reader->openZip('xl/workbook.xml');
            self::fail('Temporary file failure must be reported');
        }
        catch (Exception $e) {
            self::assertSame([], glob($this->directory . '/xlsx_reader_*'));
            self::assertTrue(unlink($file), 'The archive must not remain locked on Windows');
        }
        finally {
            $reader->close();
        }
    }
}
