<?php

declare(strict_types=1);

namespace avadim\FastExcelReader\Tests;

use avadim\FastExcelReader\Excel;
use avadim\FastExcelReader\Tests\Support\GuardTestCase;
use avadim\FastExcelReader\Tests\Support\XlsxBuilder;

/**
 * Regression tests for issue #53.
 *
 * A cell using a built-in date format (numFmtId 14 and siblings 15-22) carries no
 * explicit pattern in the file, so the reader has to pick one. It used to overwrite
 * the built-in patterns from the ambient ICU locale whenever ext-intl was loaded,
 * which made the formatted output of a single file depend on the server locale and
 * on whether the extension happened to be present.
 *
 * The output is now deterministic by default; locale rendering is opt-in.
 */
final class BuiltinDateFormatTest extends GuardTestCase
{
    /** Serial 31070 is 1985-01-23. */
    private const SERIAL = 31070;

    /**
     * Build a one-cell workbook: B2 holds the date serial formatted with a built-in code.
     *
     * @param int $numFmtId
     *
     * @return string cell B2 value after formatting via dateFormatter(true)
     */
    private function readB2(int $numFmtId, ?string $locale = null): string
    {
        $file = XlsxBuilder::withRows([2 => ['B' => self::SERIAL]])
            ->withCellFormats(['B2' => $numFmtId])
            ->build();

        $excel = Excel::open($file);
        if ($locale !== null) {
            $excel->useLocaleFormats($locale);
        }
        $excel->dateFormatter(true);

        return (string)$excel->readCells()['B2'];
    }

    /**
     * The default output for the short-date code 14 does not depend on the process locale.
     *
     * @return void
     */
    public function testBuiltinShortDateIsDeterministicAcrossLocales(): void
    {
        \Locale::setDefault('ru_RU');
        $ru = $this->readB2(14);

        \Locale::setDefault('en_US');
        $en = $this->readB2(14);

        \Locale::setDefault('ja_JP');
        $jp = $this->readB2(14);

        $this->assertSame('01-23-85', $ru);
        $this->assertSame($ru, $en);
        $this->assertSame($ru, $jp);
    }

    /**
     * useLocaleFormats() opts in to locale-dependent patterns with an explicit locale,
     * regardless of the process default locale.
     *
     * @return void
     */
    public function testUseLocaleFormatsRendersWithTheGivenLocale(): void
    {
        // Pin an unrelated process locale to prove the argument, not the ambient default, wins.
        \Locale::setDefault('ja_JP');

        $this->assertSame('23.01.1985', $this->readB2(14, 'ru_RU'));
        $this->assertSame('1/23/85', $this->readB2(14, 'en_US'));
    }

    /**
     * The sibling built-in codes 15-22 are deterministic by default too.
     *
     * @return void
     */
    public function testOtherBuiltinDateCodesAreDeterministic(): void
    {
        \Locale::setDefault('ru_RU');

        $this->assertSame('23-Jan-85', $this->readB2(15)); // d-mmm-yy
        $this->assertSame('23-Jan', $this->readB2(16));    // d-mmm
        $this->assertSame('Jan-85', $this->readB2(17));    // mmm-yy
    }

    /**
     * useLocaleFormats() requires ext-intl; without it the call fails loudly rather than
     * silently doing nothing.
     *
     * @return void
     */
    public function testUseLocaleFormatsRequiresIntl(): void
    {
        if (class_exists('IntlDateFormatter', false)) {
            $this->markTestSkipped('ext-intl is loaded, cannot test the missing-extension path');
        }

        $file = XlsxBuilder::withRows([2 => ['B' => self::SERIAL]])
            ->withCellFormats(['B2' => 14])
            ->build();

        $this->expectException(\avadim\FastExcelReader\Exception::class);
        Excel::open($file)->useLocaleFormats('ru_RU');
    }

    /**
     * A numFmt the file declares wins over the builtin meaning of its id.
     *
     * Ids 14-22, 27-36, 45-47, 50-58 and 71-81 stand for builtin date formats,
     * several of them localized, and a cell using one carries no format code.
     * A file that declares a code for such an id has redefined it, though -
     * converters coming from XLS do exactly that - and then the declared code
     * is what the cell is formatted with. Reading it by id turned an article
     * number like 043, stored as 43 under the pattern "000", into a date.
     *
     * @return void
     */
    public function testDeclaredFormatCodeWinsOverABuiltinId(): void
    {
        $file = XlsxBuilder::withRows([2 => ['A' => 43, 'B' => self::SERIAL, 'C' => self::SERIAL]])
            ->withNumberFormats([51 => '000', 52 => 'dd\.mm\.yyyy'])
            ->withCellFormats(['A2' => 51, 'B2' => 52, 'C2' => 14])
            ->build();

        $cells = Excel::open($file)->sheet()->readCells(true);

        $this->assertSame(43, $cells['A2']['v']);
        $this->assertSame('number', $cells['A2']['t']);

        // a genuine date declared under the same range of ids stays a date
        $this->assertSame('date', $cells['B2']['t']);

        // and an id used without a declaration still means the builtin format
        $this->assertSame('date', $cells['C2']['t']);
    }
}
