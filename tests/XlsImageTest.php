<?php

declare(strict_types=1);

namespace avadim\FastExcelReader\Tests;

use avadim\FastExcelReader\Excel;
use avadim\FastExcelReader\Tests\Support\GuardTestCase;
use avadim\FastExcelReader\Xls\XlsEscher;

/**
 * Pictures read from XLS.
 *
 * XLSX keeps pictures as separate files inside the package; XLS embeds them in
 * the Office Drawing ("Escher") record tree. The bytes live once in the
 * workbook-level blip store, and each sheet only names a blip index and the
 * cell the shape is anchored to.
 *
 * The decisive check is that the extracted bytes are identical to the ones the
 * XLSX reader returns for the same picture: a wrong offset in the blip header
 * would still produce plausible-looking data of nearly the right size.
 */
final class XlsImageTest extends GuardTestCase
{
    private const XLS_DIR = __DIR__ . '/test_files/xls/';

    /**
     * @return void
     */
    public function testImageBytesAreIdenticalToXlsx(): void
    {
        $xls = Excel::open(self::XLS_DIR . 'demo-03-images.xls')->sheet();
        $xlsx = Excel::open(self::fixture('demo-03-images.xlsx'))->sheet();

        foreach (['C2', 'C3'] as $cell) {
            $fromXlsx = $xlsx->getImageBlob($cell);
            $fromXls = $xls->getImageBlob($cell);

            $this->assertNotEmpty($fromXlsx, 'guard: the XLSX fixture must have a picture at ' . $cell);
            $this->assertSame($fromXlsx, $fromXls, $cell);
        }
    }

    /**
     * The pictures are anchored to the same cells as in XLSX
     *
     * @return void
     */
    public function testImageCellsMatchXlsx(): void
    {
        $xls = Excel::open(self::XLS_DIR . 'demo-03-images.xls')->sheet();
        $xlsx = Excel::open(self::fixture('demo-03-images.xlsx'))->sheet();

        $this->assertSame(array_keys($xlsx->getImageList()), array_keys($xls->getImageList()));
        $this->assertSame(['C2', 'C3'], array_keys($xls->getImageList()));
    }

    /**
     * File names follow the picture type recorded in the blip store
     *
     * @return void
     */
    public function testImageList(): void
    {
        $sheet = Excel::open(self::XLS_DIR . 'demo-03-images.xls')->sheet();
        $list = $sheet->getImageList();

        $this->assertSame(['image_name', 'file_name'], array_keys($list['C2']));
        $this->assertSame('image1.jpeg', $list['C2']['file_name']);
        $this->assertSame('image2.jpeg', $list['C3']['file_name']);
    }

    /**
     * @return void
     */
    public function testCountAndPresence(): void
    {
        $book = Excel::open(self::XLS_DIR . 'demo-03-images.xls');
        $sheet = $book->sheet();

        $this->assertSame(2, $sheet->countImages());
        $this->assertSame(2, $book->countImages());
        $this->assertTrue($book->hasImages());
        $this->assertTrue($sheet->hasDrawings());

        $this->assertTrue($sheet->hasImage('C2'));
        $this->assertTrue($sheet->hasImage('c2'), 'addresses are case-insensitive');
        $this->assertFalse($sheet->hasImage('A1'));
    }

    /**
     * A workbook without pictures reports none rather than failing
     *
     * @return void
     */
    public function testWorkbookWithoutImages(): void
    {
        $book = Excel::open(self::XLS_DIR . 'demo-00-test.xls');

        $this->assertSame(0, $book->countImages());
        $this->assertFalse($book->hasImages());
        $this->assertSame([], $book->sheet()->getImageList());
        $this->assertNull($book->sheet()->getImageBlob('A1'));
    }

    /**
     * The picture type comes from the blip store, so no filesystem access and
     * no fileinfo extension are involved
     *
     * @return void
     */
    public function testMimeType(): void
    {
        $sheet = Excel::open(self::XLS_DIR . 'demo-03-images.xls')->sheet();

        $this->assertSame('image/jpeg', $sheet->getImageMimeType('C2'));
        $this->assertNull($sheet->getImageMimeType('A1'));
    }

    /**
     * @return void
     */
    public function testImagesByRow(): void
    {
        $sheet = Excel::open(self::XLS_DIR . 'demo-03-images.xls')->sheet();

        $this->assertSame(['C2'], array_keys($sheet->getImageListByRow(2)));
        $this->assertSame(['C3'], array_keys($sheet->getImageListByRow(3)));
        $this->assertSame([], $sheet->getImageListByRow(99));
    }

    /**
     * @return void
     */
    public function testSaveImage(): void
    {
        $sheet = Excel::open(self::XLS_DIR . 'demo-03-images.xls')->sheet();
        $target = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'fxr-image-' . getmypid() . '.jpeg';

        $saved = $sheet->saveImage('C2', $target);

        $this->assertNotNull($saved);
        $this->assertFileExists($saved);
        $this->assertSame($sheet->getImageBlob('C2'), file_get_contents($saved));

        unlink($saved);
    }

    /**
     * @return void
     */
    public function testSaveImageTo(): void
    {
        $sheet = Excel::open(self::XLS_DIR . 'demo-03-images.xls')->sheet();
        $dir = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'fxr-images-' . getmypid();
        if (!is_dir($dir)) {
            mkdir($dir, 0777, true);
        }

        $saved = $sheet->saveImageTo('C3', $dir);

        $this->assertNotNull($saved);
        $this->assertSame('image2.jpeg', basename($saved));
        $this->assertFileExists($saved);

        unlink($saved);
        rmdir($dir);
    }

    /**
     * The picture bytes must not be pulled in while merely reading values
     *
     * @return void
     */
    public function testReadingValuesIsUnaffected(): void
    {
        $xls = Excel::open(self::XLS_DIR . 'demo-03-images.xls')->readRows();
        $xlsx = Excel::open(self::fixture('demo-03-images.xlsx'))->readRows();

        $this->assertSame($xlsx, $xls);
    }

    /**
     * Malformed drawing data must not read past the end of the record
     *
     * @return void
     */
    public function testTruncatedDrawingDataIsSafe(): void
    {
        $this->assertSame([], XlsEscher::blipStore(''));
        $this->assertSame([], XlsEscher::shapes(''));

        // a container claiming far more bytes than it has
        $truncated = pack('vvV', 0x000F, 0xF000, 0xFFFF) . 'short';

        $this->assertSame([], XlsEscher::blipStore($truncated));
        $this->assertSame([], XlsEscher::shapes($truncated));
    }
}
