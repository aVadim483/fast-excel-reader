<?php

declare(strict_types=1);

namespace avadim\FastExcelReader\Tests;

use avadim\FastExcelReader\Excel;
use avadim\FastExcelReader\Tests\Support\GuardTestCase;

/**
 * Workbook-level accessors that forward to the current sheet, plus the image
 * accessors on the sheet itself.
 *
 * All of these are format-agnostic and therefore move into the shared base
 * classes during the refactoring, yet most had no coverage. Delegation is
 * exactly the kind of code that survives a refactoring "by accident" - it keeps
 * compiling while silently pointing at the wrong sheet.
 */
final class WorkbookDelegationTest extends GuardTestCase
{
    /**
     * Sheets can be addressed by id as well as by name
     *
     * @return void
     */
    public function testGetSheetByIdMatchesGetSheetByName(): void
    {
        $excel = Excel::open(self::fixture('demo-00-test.xlsx'));
        $names = array_values($excel->getSheetNames());

        $byName = $excel->getSheet($names[1]);
        $byId = $excel->getSheetById((int)$byName->id());

        $this->assertSame($byName->name(), $byId->name());
        $this->assertSame($byName->path(), $byId->path());
    }

    /**
     * selectSheetById() switches the current sheet, so the delegating readers
     * follow it
     *
     * @return void
     */
    public function testSelectSheetByIdSwitchesTheCurrentSheet(): void
    {
        $excel = Excel::open(self::fixture('demo-00-test.xlsx'));

        $excel->selectSheetById(2);
        $this->assertSame('Sheet2', $excel->sheet()->name());
        $this->assertSame($excel->sheet()->readRows(), $excel->readRows());

        $excel->selectSheetById(3);
        $this->assertSame('Sheet3', $excel->sheet()->name());
        $this->assertSame($excel->sheet()->readRows(), $excel->readRows());
    }

    /**
     * An unknown sheet id must not silently fall back to the first sheet
     *
     * @return void
     */
    public function testSelectSheetByUnknownIdThrows(): void
    {
        $excel = Excel::open(self::fixture('demo-00-test.xlsx'));

        $this->expectException(\Throwable::class);

        $excel->selectSheetById(9999);
    }

    /**
     * readCallback() at workbook level walks the current sheet cell by cell and
     * honours a truthy return as "stop"
     *
     * @return void
     */
    public function testWorkbookReadCallbackWalksCurrentSheet(): void
    {
        $excel = Excel::open(self::fixture('demo-00-test.xlsx'));

        $seen = [];
        $excel->readCallback(static function ($row, $col, $value) use (&$seen) {
            $seen[$col . $row] = $value;

            return null;
        });

        $this->assertSame($excel->readCells(), $seen);
    }

    /**
     * A truthy return from the callback stops the walk immediately
     *
     * @return void
     */
    public function testReadCallbackStopsOnTruthyReturn(): void
    {
        $sheet = Excel::open(self::fixture('demo-00-test.xlsx'))->sheet();

        $seen = 0;
        $sheet->readCallback(static function () use (&$seen) {
            $seen++;

            return $seen >= 3;
        });

        $this->assertSame(3, $seen);
    }

    /**
     * getDateFormat() reports what setDateFormat() stored
     *
     * @return void
     */
    public function testDateFormatRoundTrip(): void
    {
        $excel = Excel::open(self::fixture('demo-05-datetime.xlsx'));

        $this->assertNull($excel->getDateFormat(), 'no format is configured by default');

        $excel->setDateFormat('Y-m-d');

        $this->assertSame('Y-m-d', $excel->getDateFormat());
    }

    /**
     * Image presence flags at workbook level
     *
     * @return void
     */
    public function testImagePresenceFlags(): void
    {
        $withImages = Excel::open(self::fixture('demo-03-images.xlsx'));
        $without = Excel::open(self::fixture('demo-00-test.xlsx'));

        $this->assertTrue($withImages->hasImages());
        $this->assertSame(2, $withImages->countImages());

        $this->assertFalse($without->hasImages());
        $this->assertFalse($without->hasExtraImages());
        $this->assertSame(0, $without->countImages());
    }

    /**
     * The per-cell image accessors agree with the image list
     *
     * @return void
     */
    public function testSheetImageAccessors(): void
    {
        $sheet = Excel::open(self::fixture('demo-03-images.xlsx'))->sheet();

        $this->assertSame(['C2', 'C3'], array_keys($sheet->getImageList()));

        $this->assertTrue($sheet->hasImage('C2'));
        $this->assertFalse($sheet->hasImage('A1'));

        $this->assertSame('image1.jpeg', $sheet->getImageList()['C2']['file_name']);
        $this->assertNotNull($sheet->getImageName('C2'));
        $this->assertNull($sheet->getImageName('A1'), 'a cell without an image has no name');
    }

    /**
     * getImageListByRow() filters the same data by row number
     *
     * @return void
     */
    public function testGetImageListByRow(): void
    {
        $sheet = Excel::open(self::fixture('demo-03-images.xlsx'))->sheet();

        $this->assertSame(['C2'], array_keys($sheet->getImageListByRow(2)));
        $this->assertSame(['C3'], array_keys($sheet->getImageListByRow(3)));
        $this->assertSame([], $sheet->getImageListByRow(99));
    }

    /**
     * The MIME type is resolved from the stored entry, and is null for a cell
     * that holds no image
     *
     * @return void
     */
    public function testGetImageMimeType(): void
    {
        $sheet = Excel::open(self::fixture('demo-03-images.xlsx'))->sheet();

        $this->assertNull($sheet->getImageMimeType('A1'));

        if (function_exists('mime_content_type')) {
            $this->assertSame('image/jpeg', $sheet->getImageMimeType('C2'));
        }
    }

    /**
     * Workbook readers delegate to whichever sheet is current
     *
     * @return void
     */
    public function testWorkbookReadersFollowTheSelectedSheet(): void
    {
        $excel = Excel::open(self::fixture('demo-00-test.xlsx'));
        $names = array_values($excel->getSheetNames());

        $excel->selectSheet($names[0]);
        $first = $excel->readRows();

        $excel->selectSheet($names[1]);
        $second = $excel->readRows();

        $this->assertNotSame($first, $second);
        $this->assertSame($excel->sheet()->readRows(), $second);
    }
}
