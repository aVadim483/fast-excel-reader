<?php

declare(strict_types=1);

namespace avadim\FastExcelReader\Tests;

use avadim\FastExcelReader\Exception;
use avadim\FastExcelReader\Tests\Support\GuardTestCase;
use avadim\FastExcelReader\Xls\BiffReader;
use avadim\FastExcelReader\Xls\BiffRecord;
use avadim\FastExcelReader\Xls\BiffString;
use avadim\FastExcelReader\Xls\OleReader;

/**
 * The two lowest layers of the XLS reader: the OLE2 container and the BIFF
 * record stream.
 *
 * Fixtures under tests/test_files/xls were produced by LibreOffice from the
 * matching .xlsx demo files, so they are real BIFF8 written by a third party
 * rather than by this project - a reader validated against its own writer
 * proves very little.
 */
final class XlsContainerTest extends GuardTestCase
{
    private const XLS_DIR = __DIR__ . '/test_files/xls/';

    /**
     * @return void
     */
    public function testOpensCompoundFileAndListsStreams(): void
    {
        $ole = new OleReader(self::XLS_DIR . 'demo-00-test.xls');

        $streams = $ole->streamList();

        $this->assertContains('Workbook', $streams);
        $this->assertTrue($ole->streamExists('Workbook'));
        $this->assertFalse($ole->streamExists('NoSuchStream'));
    }

    /**
     * A BIFF5/BIFF7 workbook stores its data in a stream named "Book"; only
     * "Workbook" is BIFF8
     *
     * @return void
     */
    public function testFixturesAreBiff8(): void
    {
        foreach (glob(self::XLS_DIR . '*.xls') as $file) {
            $ole = new OleReader($file);

            $this->assertTrue($ole->streamExists('Workbook'), basename($file) . ' must hold a Workbook stream');
            $this->assertFalse($ole->streamExists('Book'), basename($file) . ' must not be BIFF5');
        }
    }

    /**
     * @return void
     */
    public function testRejectsFilesThatAreNotCompoundFiles(): void
    {
        $this->expectException(Exception::class);
        $this->expectExceptionMessageMatches('/Not an OLE2 compound file/');

        new OleReader(self::fixture('demo-00-test.xlsx'));
    }

    /**
     * @return void
     */
    public function testRejectsMissingFile(): void
    {
        $this->expectException(Exception::class);

        new OleReader(self::XLS_DIR . 'no-such-file.xls');
    }

    /**
     * @return void
     */
    public function testUnknownStreamThrows(): void
    {
        $ole = new OleReader(self::XLS_DIR . 'demo-00-test.xls');

        $this->expectException(Exception::class);

        $ole->openStream('NoSuchStream');
    }

    /**
     * Reading in small chunks must produce exactly the same bytes as one big
     * read - this is what proves the sector chain is walked correctly
     *
     * @dataProvider fixtureProvider
     *
     * @param string $file
     *
     * @return void
     */
    public function testChunkedReadsMatchWholeStream(string $file): void
    {
        $ole = new OleReader($file);

        $whole = $ole->openStream('Workbook');
        $size = $whole->size();
        $expected = $whole->read($size);

        $this->assertSame($size, strlen($expected), 'the stream must yield exactly its declared size');

        $chunked = '';
        $stream = $ole->openStream('Workbook');
        while (!$stream->eof()) {
            $chunked .= $stream->read(97); // deliberately not a divisor of the sector size
        }

        $this->assertSame($expected, $chunked);
    }

    /**
     * @return array<string, array{0: string}>
     */
    public function fixtureProvider(): array
    {
        $data = [];
        foreach (glob(self::XLS_DIR . '*.xls') as $file) {
            $data[basename($file)] = [$file];
        }

        return $data;
    }

    /**
     * demo-05-datetime.xls has a Workbook stream below the 4 KB cutoff, so it
     * lives in the mini stream and takes the other code path entirely
     *
     * @return void
     */
    public function testMiniStreamPath(): void
    {
        $ole = new OleReader(self::XLS_DIR . 'demo-05-datetime.xls');
        $stream = $ole->openStream('Workbook');

        $this->assertLessThan(4096, $stream->size(), 'this fixture is expected to exercise the mini stream');
        $this->assertSame("\x09\x08", substr($stream->read(2), 0, 2), 'a BIFF stream starts with a BOF record');
    }

    /**
     * @return void
     */
    public function testSeekAndTell(): void
    {
        $ole = new OleReader(self::XLS_DIR . 'demo-00-test.xls');
        $stream = $ole->openStream('Workbook');

        $stream->seek(1000);
        $this->assertSame(1000, $stream->tell());
        $atThousand = $stream->read(16);
        $this->assertSame(1016, $stream->tell());

        $stream->seek(1000);
        $this->assertSame($atThousand, $stream->read(16), 'seeking back must return the same bytes');

        $stream->seek($stream->size());
        $this->assertTrue($stream->eof());
        $this->assertSame('', $stream->read(10), 'reading past the end yields nothing');
    }

    /**
     * The record stream must be consumed exactly, with no bytes left over and
     * no overrun - the strongest single check that record framing is right
     *
     * @dataProvider fixtureProvider
     *
     * @param string $file
     *
     * @return void
     */
    public function testRecordStreamIsConsumedExactly(string $file): void
    {
        $ole = new OleReader($file);
        $stream = $ole->openStream('Workbook');
        $size = $stream->size();
        $biff = new BiffReader($stream);

        $count = 0;
        foreach ($biff->records() as $record) {
            $this->assertIsInt($record['type']);
            $count++;
        }

        $this->assertGreaterThan(10, $count);
        $this->assertSame($size, $biff->tell(), 'the whole stream must be accounted for by records');
    }

    /**
     * @return void
     */
    public function testFirstRecordIsBiff8Bof(): void
    {
        $ole = new OleReader(self::XLS_DIR . 'demo-00-test.xls');
        $biff = new BiffReader($ole->openStream('Workbook'));

        $record = $biff->nextRecord();

        $this->assertSame(BiffRecord::BOF, $record['type']);
        $this->assertSame(0, $record['offset']);

        $version = unpack('v', substr($record['data'], 0, 2))[1];
        $substream = unpack('v', substr($record['data'], 2, 2))[1];

        $this->assertSame(BiffRecord::VERSION_BIFF8, $version);
        $this->assertSame(BiffRecord::SUBSTREAM_GLOBALS, $substream);
    }

    /**
     * BOUNDSHEET carries the absolute offset of each sheet's BOF, which is what
     * lets a single sheet be reached without reading the ones before it
     *
     * @return void
     */
    public function testSeekToSheetSubstream(): void
    {
        $ole = new OleReader(self::XLS_DIR . 'demo-02-advanced.xls');
        $biff = new BiffReader($ole->openStream('Workbook'));

        $sheets = [];
        foreach ($biff->records() as $record) {
            if ($record['type'] === BiffRecord::BOUNDSHEET) {
                [$name] = BiffString::readShort($record['data'], 6);
                $sheets[$name] = unpack('V', substr($record['data'], 0, 4))[1];
            }
            elseif ($record['type'] === BiffRecord::EOF) {
                break;
            }
        }

        $this->assertSame(['Demo1', 'Demo2', 'Demo3'], array_keys($sheets));

        foreach ($sheets as $name => $offset) {
            $biff->seek($offset);
            $bof = $biff->nextRecord();

            $this->assertSame(BiffRecord::BOF, $bof['type'], $name . ' must start with a BOF');
            $this->assertSame(
                BiffRecord::SUBSTREAM_WORKSHEET,
                unpack('v', substr($bof['data'], 2, 2))[1],
                $name . ' must be a worksheet substream'
            );
        }
    }

    /**
     * DIMENSIONS is zero based, so the XLSX range B2:D11 appears as rows 1..11
     * and columns 1..4
     *
     * @return void
     */
    public function testSheetDimensionsMatchTheXlsxCounterpart(): void
    {
        $ole = new OleReader(self::XLS_DIR . 'demo-02-advanced.xls');
        $biff = new BiffReader($ole->openStream('Workbook'));

        $offset = null;
        foreach ($biff->records() as $record) {
            if ($record['type'] === BiffRecord::BOUNDSHEET) {
                $offset = unpack('V', substr($record['data'], 0, 4))[1];
                break;
            }
        }
        $this->assertNotNull($offset);

        $biff->seek($offset);
        $dimensions = null;
        foreach ($biff->records() as $record) {
            if ($record['type'] === BiffRecord::DIMENSIONS) {
                $dimensions = unpack('VrowFirst/VrowLast/vcolFirst/vcolLast', substr($record['data'], 0, 12));
                break;
            }
        }

        $this->assertNotNull($dimensions);
        $this->assertSame(1, $dimensions['rowFirst']);
        $this->assertSame(11, $dimensions['rowLast']);
        $this->assertSame(1, $dimensions['colFirst']);
        $this->assertSame(4, $dimensions['colLast']);
    }

    /**
     * @return void
     */
    public function testSheetNamesAndSharedStrings(): void
    {
        $ole = new OleReader(self::XLS_DIR . 'demo-00-test.xls');
        $biff = new BiffReader($ole->openStream('Workbook'));

        $names = [];
        $strings = [];
        $date1904 = null;
        foreach ($biff->records() as $record) {
            if ($record['type'] === BiffRecord::BOUNDSHEET) {
                [$name] = BiffString::readShort($record['data'], 6);
                $names[] = $name;
            }
            elseif ($record['type'] === BiffRecord::SST) {
                $strings = BiffString::readSharedStrings($record['parts']);
            }
            elseif ($record['type'] === BiffRecord::DATEMODE) {
                $date1904 = unpack('v', $record['data'])[1] === 1;
            }
            elseif ($record['type'] === BiffRecord::EOF) {
                break;
            }
        }

        $this->assertSame(['Sheet1', 'Sheet2', 'Sheet3'], $names);
        $this->assertFalse($date1904);

        // the same values the XLSX fixture carries in its first row
        $this->assertContains('#', $strings);
        $this->assertContains('name', $strings);
        $this->assertContains('birthday', $strings);
        $this->assertContains('random_int', $strings);
        $this->assertContains('James Bond', $strings);
    }

    /**
     * The nastiest corner of BIFF8: a shared string table too large for one
     * record continues in CONTINUE records, and a string may be cut in half by
     * the boundary. The continuation then restates the encoding flag, and the
     * encoding may flip mid-string.
     *
     * The fixture mixes single-byte and UTF-16 strings of a fixed shape, so any
     * mis-handled boundary shows up as a malformed value rather than as a
     * plausible one.
     *
     * @return void
     */
    public function testSharedStringsSplitAcrossContinueRecords(): void
    {
        $ole = new OleReader(self::XLS_DIR . 'continue-sst.xls');
        $biff = new BiffReader($ole->openStream('Workbook'));

        $record = null;
        foreach ($biff->records() as $candidate) {
            if ($candidate['type'] === BiffRecord::SST) {
                $record = $candidate;
                break;
            }
        }

        $this->assertNotNull($record, 'the fixture must contain an SST');
        $this->assertGreaterThan(1, count($record['parts']), 'the fixture must span CONTINUE records');
        $this->assertGreaterThan(8224, strlen($record['data']));

        $strings = BiffString::readSharedStrings($record['parts']);

        $this->assertCount(440, $strings);

        $malformed = [];
        foreach ($strings as $string) {
            if (!preg_match('/^(ascii-\d{4}-x{34}|кир-\d{4}-я{28})$/u', $string)) {
                $malformed[] = $string;
            }
        }

        $this->assertSame([], $malformed, 'every string must survive the segment boundaries intact');
    }

    /**
     * @return void
     */
    public function testShortAndLongStringDecoding(): void
    {
        // ShortXLUnicodeString: 1 byte length, 1 byte flags, compressed payload
        [$value, $consumed] = BiffString::readShort("\x03\x00abc", 0);
        $this->assertSame('abc', $value);
        $this->assertSame(5, $consumed);

        // the same, wide
        [$value, $consumed] = BiffString::readShort("\x02\x01a\x00b\x00", 0);
        $this->assertSame('ab', $value);
        $this->assertSame(6, $consumed);

        // XLUnicodeString: 2 byte length
        [$value, $consumed] = BiffString::readLong("\x03\x00\x00xyz", 0);
        $this->assertSame('xyz', $value);
        $this->assertSame(6, $consumed);

        // U+042F, little endian
        [$value] = BiffString::readLong("\x01\x00\x01\x2f\x04", 0);
        $this->assertSame('Я', $value, 'UTF-16LE payload must decode');
    }

    /**
     * @return void
     */
    public function testRecordNames(): void
    {
        $this->assertSame('BOF', BiffRecord::name(BiffRecord::BOF));
        $this->assertSame('LABELSST', BiffRecord::name(BiffRecord::LABELSST));
        $this->assertSame('0x1234', BiffRecord::name(0x1234), 'unknown records fall back to their number');
    }

    /**
     * Reading a workbook must not scale with the file size: the container keeps
     * the allocation tables, not the data
     *
     * @return void
     */
    public function testMemoryDoesNotScaleWithStreamSize(): void
    {
        $before = memory_get_usage();

        $ole = new OleReader(self::XLS_DIR . 'continue-sst.xls');
        $stream = $ole->openStream('Workbook');
        $biff = new BiffReader($stream);

        $count = 0;
        foreach ($biff->records() as $record) {
            $count++;
        }
        $growth = memory_get_usage() - $before;

        $this->assertGreaterThan(100, $count);
        $this->assertLessThan(
            2 * 1024 * 1024,
            $growth,
            sprintf('walking %d records over a %d byte stream grew the heap by %d bytes', $count, $stream->size(), $growth)
        );
    }
}
