<?php

namespace avadim\FastExcelReader\Internal;

use avadim\FastExcelReader\Exception;

/** Content-based dispatch for ZIP spreadsheet containers. */
final class SpreadsheetPackage
{
    public const ODS_MIME = 'application/vnd.oasis.opendocument.spreadsheet';

    public static function isOds(string $file): bool
    {
        $zip = new \ZipArchive();
        if (@$zip->open($file) !== true) {
            return false;
        }
        try {
            return $zip->getFromName('mimetype', 256) === self::ODS_MIME;
        }
        finally {
            $zip->close();
        }
    }

    public static function isXlsx(string $file): bool
    {
        $zip = new \ZipArchive();
        if (@$zip->open($file) !== true) {
            return false;
        }
        try {
            if ($zip->locateName('xl/workbook.xml') === false || $zip->locateName('xl/_rels/workbook.xml.rels') === false) {
                return false;
            }
            $types = $zip->getFromName('[Content_Types].xml');
            if ($types === false) {
                return false;
            }
            $xml = new \DOMDocument();
            if (!@$xml->loadXML($types, LIBXML_NONET) || $xml->doctype) {
                return false;
            }
            $allowed = [
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
                'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
                'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml',
                'application/vnd.ms-excel.template.macroEnabled.main+xml',
            ];
            foreach ($xml->getElementsByTagNameNS('http://schemas.openxmlformats.org/package/2006/content-types', 'Override') as $part) {
                if ($part->getAttribute('PartName') === '/xl/workbook.xml') {
                    return in_array($part->getAttribute('ContentType'), $allowed, true);
                }
            }
            // Some producers declare the workbook MIME as the default for .xml parts.
            foreach ($xml->getElementsByTagNameNS('http://schemas.openxmlformats.org/package/2006/content-types', 'Default') as $part) {
                if ($part->getAttribute('Extension') === 'xml' && in_array($part->getAttribute('ContentType'), $allowed, true)) {
                    return true;
                }
            }
            return false;
        }
        finally {
            $zip->close();
        }
    }

    public static function requireXlsx(string $file): void
    {
        if (self::isXlsx($file)) {
            return;
        }
        $zip = new \ZipArchive();
        $binary = false;
        if (@$zip->open($file) === true) {
            $binary = $zip->locateName('xl/workbook.bin') !== false;
            $zip->close();
        }
        throw new Exception($binary ? 'XLSB workbooks are not supported' : 'Not an XLSX workbook: missing or unsupported workbook package parts (DOCX/PPTX and other ZIP files are not supported)');
    }
}
