<?php

namespace avadim\FastExcelReader\Ods;

use avadim\FastExcelReader\AbstractBook;
use avadim\FastExcelReader\Exception;
use avadim\FastExcelReader\Internal\SpreadsheetPackage;
use avadim\FastExcelReader\Reader;

/** Streaming OpenDocument spreadsheet workbook. */
class OdsBook extends AbstractBook
{
    public const TABLE_NS = 'urn:oasis:names:tc:opendocument:xmlns:table:1.0';
    public const OFFICE_NS = 'urn:oasis:names:tc:opendocument:xmlns:office:1.0';
    public const TEXT_NS = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0';

    public static function setTempDir($tempDir)
    {
        Reader::setTempDir($tempDir);
    }

    protected function _prepare(string $file): void
    {
        if (!SpreadsheetPackage::isOds($file)) {
            throw new Exception('Not an ODS workbook');
        }
        $zip = new \ZipArchive();
        $zip->open($file);
        try {
            if ($zip->locateName('content.xml') === false || $zip->locateName('META-INF/manifest.xml') === false) {
                throw new Exception('ODS package is missing content.xml or manifest.xml');
            }
        }
        finally {
            $zip->close();
        }
        $reader = $this->openContent();
        $index = 0;
        $spreadsheet = false;
        $spreadsheetDepth = -1;
        $complete = false;
        try {
            while ($reader->read()) {
                self::checkNode($reader);
                if ($reader->nodeType === \XMLReader::END_ELEMENT && $reader->depth === 0) {
                    $complete = true;
                }
                if ($reader->nodeType === \XMLReader::ELEMENT && $reader->namespaceURI === self::OFFICE_NS && $reader->localName === 'spreadsheet') {
                    $spreadsheet = true;
                    $spreadsheetDepth = $reader->depth;
                }
                if ($reader->nodeType === \XMLReader::ELEMENT && $reader->namespaceURI === self::TABLE_NS && $reader->localName === 'table' && $reader->depth === $spreadsheetDepth + 1) {
                    $index++;
                    $name = (string)$reader->getAttributeNs('name', self::TABLE_NS);
                    if ($name === '') {
                        throw new Exception('ODS worksheet has no name');
                    }
                    $this->sheets[$index] = new OdsSheet($name, (string)$index, $this);
                }
            }
        }
        finally {
            $reader->close();
            $this->releaseReader($reader);
        }
        if (!$complete || !$spreadsheet || !$this->sheets) {
            throw new Exception('ODS content has no spreadsheet worksheets');
        }
        $this->defaultSheetId = 1;
    }

    /** @internal Each iterator owns its cursor, so reading another sheet cannot move it. */
    public function openContent(): Reader
    {
        $this->assertOpen();
        $reader = new Reader($this->file, [\XMLReader::LOADDTD => false, \XMLReader::SUBST_ENTITIES => false]);
        $this->ownReader($reader);
        $reader->openZip('content.xml');
        return $reader;
    }

    /** @internal Reject document type declarations before any cell is expanded. */
    public static function checkNode(Reader $reader): void
    {
        if ($reader->nodeType === \XMLReader::DOC_TYPE) {
            throw new Exception('Document type declarations are not supported in ODS');
        }
    }

    /** ODS has ISO date values, not Excel serials. Null/false formatters return ISO text. */
    public function dateValue(string $value)
    {
        if (!$this->getDateFormatter()) {
            return $value;
        }
        try {
            // Like XLSX serials, timezone-free spreadsheet dates use a UTC baseline.
            $date = new \DateTimeImmutable($value, new \DateTimeZone('UTC'));
            $errors = \DateTimeImmutable::getLastErrors();
            if ($errors && ($errors['warning_count'] || $errors['error_count'])) {
                throw new \UnexpectedValueException('Invalid calendar date');
            }
        }
        catch (\Throwable $e) {
            throw new Exception('Invalid ODS date: ' . $value, 0, $e);
        }
        return $this->formatDate($date->getTimestamp());
    }

    protected function _loadCompleteStyles()
    {
        $this->styles['_'] = ['numFmts' => [], 'fonts' => [], 'fills' => [], 'borders' => [], 'cellStyleXfs' => [], 'cellXfs' => []];
    }
}
