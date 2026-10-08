<?php

declare(strict_types=1);

namespace avadim\FastExcelReader\Tests;

use avadim\FastExcelReader\Excel;
use avadim\FastExcelReader\Reader;
use avadim\FastExcelReader\Exception;
use avadim\FastExcelReader\Internal\PackagePath;
use avadim\FastExcelReader\Tests\Support\TemporaryFilesTestCase;

final class XlsxPackageValidationTest extends TemporaryFilesTestCase
{
    /** @dataProvider relationshipPaths */
    public function testRelationshipTarget(string $target, string $entry): void
    {
        $file = $this->workbook();
        $expected = array_map(static function ($sheet) { return $sheet->readRows(); }, Excel::open($file)->sheets());
        gc_collect_cycles();
        $zip = new \ZipArchive();
        $zip->open($file);
        $rels = $zip->getFromName('xl/_rels/workbook.xml.rels');
        $dom = new \DOMDocument();
        $dom->loadXML($rels);
        foreach ($dom->getElementsByTagName('Relationship') as $rel) {
            if (basename($rel->getAttribute('Type')) === 'worksheet') {
                $original = PackagePath::resolve('xl/workbook.xml', $rel->getAttribute('Target'));
                $zip->addFromString($entry, $zip->getFromName($original));
                $rel->setAttribute('Target', $target);
                break;
            }
        }
        $zip->addFromString('xl/_rels/workbook.xml.rels', $dom->saveXML());
        $zip->close();
        self::assertSame($expected, array_map(static function ($sheet) { return $sheet->readRows(); }, Excel::open($file)->sheets()));
        $reader = new Reader($file);
        self::assertTrue($reader->openSheetByIndex(0));
        self::assertTrue($reader->seekOpenTag('sheetData'));
        $reader->close();
    }

    public static function relationshipPaths(): array
    {
        return [
            ['local/sheet1.xml', 'xl/local/sheet1.xml'],
            ['xl-data/sheet1.xml', 'xl/xl-data/sheet1.xml'],
            ['/xl/local/sheet1.xml', 'xl/local/sheet1.xml'],
            ['../sheets/sheet1.xml', 'sheets/sheet1.xml'],
            ['./worksheets/../local/sheet1.xml', 'xl/local/sheet1.xml'],
        ];
    }

    public function testExternalRelationshipIsNotReadAsPackagePart(): void
    {
        $file = $this->workbook();
        $expected = array_map(static function ($sheet) { return $sheet->readRows(); }, Excel::open($file)->sheets());
        gc_collect_cycles();
        $zip = new \ZipArchive();
        $zip->open($file);
        $rels = $zip->getFromName('xl/_rels/workbook.xml.rels');
        $rels = str_replace('</Relationships>', '<Relationship Id="external" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" TargetMode="External" Target="https://invalid.example/strings.xml"/></Relationships>', $rels);
        $zip->addFromString('xl/_rels/workbook.xml.rels', $rels);
        $zip->close();
        self::assertSame($expected, array_map(static function ($sheet) { return $sheet->readRows(); }, Excel::open($file)->sheets()));
    }

    /** @dataProvider invalidTargets */
    public function testInvalidInternalTargetIsRejected(string $target): void
    {
        $this->expectException(Exception::class);
        PackagePath::resolve('xl/workbook.xml', $target);
    }

    public static function invalidTargets(): array
    {
        return [['../../outside.xml'], ['https://example.org/a.xml'], ['//example.org/a.xml'], [''], ['a\\b.xml']];
    }

    /** @dataProvider wrapTextValues */
    public function testXmlBooleanWrapText(string $value, bool $expected): void
    {
        $styles = '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="1"><xf numFmtId="0"><alignment wrapText="' . $value . '"/></xf></cellXfs></styleSheet>';
        $file = $this->workbook(['xl/styles.xml' => $styles]);
        $style = Excel::open($file)->getCompleteStyleByIdx(0);
        self::assertSame($expected, isset($style['format']['format-wrap-text']));
    }

    public static function wrapTextValues(): array
    {
        return [['1', true], ['true', true], ['0', false], ['false', false]];
    }

    /** @dataProvider libxmlModes */
    public function testValidationRestoresModeAndClearsDiagnostics(bool $mode): void
    {
        $previous = libxml_use_internal_errors(true);
        try {
            (new \DOMDocument())->loadXML('<broken>');
            libxml_use_internal_errors($mode);
            $errors = ['stale'];
            self::assertTrue(Excel::validate($this->workbook(), $errors));
            self::assertSame([], $errors);
            self::assertSame($mode, libxml_use_internal_errors());
            self::assertSame([], libxml_get_errors());
        }
        finally {
            libxml_clear_errors();
            libxml_use_internal_errors($previous);
        }
    }

    public static function libxmlModes(): array
    {
        return [[true], [false]];
    }

    public function testMalformedPartsAreAllReported(): void
    {
        $file = $this->workbook(['custom1.xml' => '<one>', 'custom2.xml' => '<two>']);
        self::assertFalse(Excel::validate($file, $errors));
        self::assertGreaterThanOrEqual(2, count($errors));
        self::assertContainsOnlyInstancesOf(\LibXMLError::class, $errors);
        self::assertSame([], libxml_get_errors());
    }

    public function testMissingWorkbookIsNotValidXlsx(): void
    {
        self::assertFalse(Excel::validate($this->workbook(['xl/workbook.xml' => null]), $errors));
        self::assertSame([], $errors);
    }

    public function testValidationSupportsSpecialFilename(): void
    {
        self::assertTrue(Excel::validate(__DIR__ . '/test_files/spec#name%sym _.xlsx', $errors));
        self::assertSame([], $errors);
    }

    public function testValidationRestoresModeAfterException(): void
    {
        $previous = libxml_use_internal_errors(false);
        try {
            Excel::validate($this->directory . '/missing.xlsx');
            self::fail('Missing input must be rejected');
        }
        catch (Exception $e) {
            self::assertFalse(libxml_use_internal_errors());
        }
        finally {
            libxml_use_internal_errors($previous);
        }
    }
}
