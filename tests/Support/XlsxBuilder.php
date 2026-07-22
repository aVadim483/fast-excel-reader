<?php

declare(strict_types=1);

namespace avadim\FastExcelReader\Tests\Support;

/**
 * Builds a minimal valid XLSX file on the fly.
 *
 * The library only reads spreadsheets, and no shipped fixture contains strings
 * with surrounding whitespace, so TRIM_STRINGS cannot be exercised against the
 * committed files. This builder covers that gap without adding binary blobs to
 * the repository.
 *
 * Values are written as inline strings or raw numbers - no shared string table
 * and no styles beyond the default - which keeps the output small and keeps the
 * assertions focused on the reader logic rather than on the fixture.
 */
final class XlsxBuilder
{
    /** @var array<int, array<string, string>> */
    private $rows = [];

    /** @var string|null */
    private $stylesXml = null;

    /** @var string[] */
    private static $tempFiles = [];

    /**
     * @param array<int, array<string, string|int|float|null>> $rows [rowNum => [colLetter => value]]
     *
     * @return self
     */
    public static function withRows(array $rows): self
    {
        $builder = new self();
        $builder->rows = $rows;

        return $builder;
    }

    /**
     * Replace the default styles.xml, e.g. to write it pretty-printed
     *
     * @param string $stylesXml
     *
     * @return self
     */
    public function withStyles(string $stylesXml): self
    {
        $this->stylesXml = $stylesXml;

        return $this;
    }

    /**
     * Write the workbook to a temporary file and return its path.
     * The file is removed when the process ends.
     *
     * @return string
     */
    public function build(): string
    {
        $file = tempnam(sys_get_temp_dir(), 'fxr') . '.xlsx';

        $zip = new \ZipArchive();
        if ($zip->open($file, \ZipArchive::CREATE | \ZipArchive::OVERWRITE) !== true) {
            throw new \RuntimeException('Cannot create ' . $file);
        }

        $zip->addFromString('[Content_Types].xml', self::contentTypes());
        $zip->addFromString('_rels/.rels', self::rootRels());
        $zip->addFromString('xl/workbook.xml', self::workbook());
        $zip->addFromString('xl/_rels/workbook.xml.rels', self::workbookRels());
        $zip->addFromString('xl/styles.xml', $this->stylesXml ?? self::styles());
        $zip->addFromString('xl/worksheets/sheet1.xml', $this->sheet());
        $zip->close();

        self::$tempFiles[] = $file;
        register_shutdown_function(static function () use ($file) {
            if (is_file($file)) {
                @unlink($file);
            }
        });

        return $file;
    }

    /**
     * @return string
     */
    private function sheet(): string
    {
        $minRow = $this->rows ? min(array_keys($this->rows)) : 1;
        $maxRow = $this->rows ? max(array_keys($this->rows)) : 1;
        $cols = [];
        foreach ($this->rows as $cells) {
            $cols = array_merge($cols, array_keys($cells));
        }
        $cols = $cols ?: ['A'];
        sort($cols);
        $dimension = reset($cols) . $minRow . ':' . end($cols) . $maxRow;

        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            . '<dimension ref="' . $dimension . '"/><sheetData>';

        foreach ($this->rows as $rowNum => $cells) {
            $xml .= '<row r="' . (int)$rowNum . '">';
            foreach ($cells as $col => $value) {
                $addr = $col . (int)$rowNum;
                if ($value === null) {
                    $xml .= '<c r="' . $addr . '"/>';
                }
                elseif (is_int($value) || is_float($value)) {
                    $xml .= '<c r="' . $addr . '"><v>' . $value . '</v></c>';
                }
                else {
                    $xml .= '<c r="' . $addr . '" t="inlineStr"><is><t xml:space="preserve">'
                        . htmlspecialchars((string)$value, ENT_QUOTES | ENT_XML1) . '</t></is></c>';
                }
            }
            $xml .= '</row>';
        }

        return $xml . '</sheetData></worksheet>';
    }

    /**
     * @return string
     */
    private static function contentTypes(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            . '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            . '<Default Extension="xml" ContentType="application/xml"/>'
            . '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            . '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            . '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            . '</Types>';
    }

    /**
     * @return string
     */
    private static function rootRels(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            . '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            . '</Relationships>';
    }

    /**
     * @return string
     */
    private static function workbook(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
            . ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            . '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>';
    }

    /**
     * @return string
     */
    private static function workbookRels(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            . '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            . '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            . '</Relationships>';
    }

    /**
     * @return string
     */
    private static function styles(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            . '<numFmts count="0"/>'
            . '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            . '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
            . '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
            . '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            . '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
            . '</styleSheet>';
    }
}
