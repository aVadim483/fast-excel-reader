<?php

use avadim\FastExcelReader\Excel;
use PHPUnit\Framework\TestCase;

/**
 * Sheet::getHyperlinks(): read the <hyperlinks> section of a worksheet and
 * resolve external targets through the sheet relationships file.
 */
class XlsxHyperlinksTest extends TestCase
{
    public function testExternalHyperlinkResolvedFromRels()
    {
        $excel = Excel::open(__DIR__ . '/test_files/standard-file.xlsx');
        $sheet = $excel->sheet('Filters');

        $links = $sheet->getHyperlinks();

        self::assertArrayHasKey('A7', $links);
        self::assertSame('https://go.servicetitan.com/', $links['A7']['link']);
        self::assertStringStartsWith('/new/reports/218?', $links['A7']['location']);
        self::assertSame('', $links['A7']['display']);
        self::assertSame('', $links['A7']['tooltip']);
        // the internal r:id must not leak into the public result
        self::assertSame(['link', 'location', 'display', 'tooltip'], array_keys($links['A7']));
    }

    public function testSheetWithoutHyperlinksReturnsEmptyArray()
    {
        $excel = Excel::open(__DIR__ . '/test_files/standard-file.xlsx');

        self::assertSame([], $excel->sheet('Sheet1')->getHyperlinks());
    }

    public function testResultIsCached()
    {
        $sheet = Excel::open(__DIR__ . '/test_files/standard-file.xlsx')->sheet('Filters');

        self::assertSame($sheet->getHyperlinks(), $sheet->getHyperlinks());
    }
}
