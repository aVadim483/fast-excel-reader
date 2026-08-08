<?php

use avadim\FastExcelReader\Excel;
use PHPUnit\Framework\TestCase;

/**
 * Excel::getProperties(): read the workbook document properties from
 * docProps/core.xml and docProps/app.xml.
 */
class XlsxPropertiesTest extends TestCase
{
    public function testCoreAndAppProperties()
    {
        $props = Excel::open(__DIR__ . '/test_files/standard-file.xlsx')->getProperties();

        self::assertSame('Chris Holland', $props['lastModifiedBy']);
        self::assertSame('2025-02-21T00:23:41Z', $props['modified']);
        self::assertSame('Microsoft Macintosh Excel', $props['application']);
    }

    public function testCreatedAndCreatorKeys()
    {
        $props = Excel::open(__DIR__ . '/test_files/styles.xlsx')->getProperties();

        self::assertSame('Vadim Shemarov', $props['lastModifiedBy']);
        self::assertSame('2015-06-05T18:19:34Z', $props['created']);
        // an element present but empty is reported as an empty string
        self::assertArrayHasKey('creator', $props);
        self::assertSame('', $props['creator']);
    }

    public function testWorkbookWithoutDocPropsReturnsEmptyArray()
    {
        $props = Excel::open(__DIR__ . '/test_files/nonstandard-file.xlsx')->getProperties();

        self::assertSame([], $props);
    }

    public function testResultIsCached()
    {
        $excel = Excel::open(__DIR__ . '/test_files/standard-file.xlsx');

        self::assertSame($excel->getProperties(), $excel->getProperties());
    }
}
