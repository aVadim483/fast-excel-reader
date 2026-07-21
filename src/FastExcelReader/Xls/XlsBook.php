<?php

namespace avadim\FastExcelReader\Xls;

use avadim\FastExcelReader\AbstractBook;
use avadim\FastExcelReader\Exception;

/**
 * XLS (BIFF8) workbook reader
 *
 * Implements the format-specific half of AbstractBook by reading the workbook
 * globals substream: the sheet list, the shared string table, number formats
 * and cell formats.
 *
 * The globals substream is read once, up front, in a single forward pass. Sheet
 * data is not touched here at all: BOUNDSHEET stores the absolute offset of
 * each sheet's BOF, so a sheet is reached by seeking straight to it.
 */
class XlsBook extends AbstractBook
{
    protected OleReader $ole;

    /** Number format strings by format index, from FORMAT records */
    protected array $numberFormats = [];

    /** Format index of every XF record, in file order */
    protected array $xfFormatIndex = [];

    protected int $codepage = 1252;

    /**
     * Open an XLS file
     *
     * @param string $file
     *
     * @return XlsBook
     */
    public static function open(string $file): XlsBook
    {
        return new static($file);
    }

    /**
     * XLS is read straight from the file, nothing is ever extracted
     *
     * @param string $tempDir
     *
     * @return void
     */
    public static function setTempDir($tempDir)
    {
    }

    /**
     * A fresh record reader over the workbook stream
     *
     * Every sheet gets its own reader: a suspended row generator must not have
     * its position moved by a read of another sheet.
     *
     * @return BiffReader
     */
    public function newBiffReader(): BiffReader
    {
        return new BiffReader($this->ole->openStream('Workbook'));
    }

    /**
     * @param string $file
     *
     * @return void
     *
     * @throws Exception
     */
    protected function _prepare(string $file): void
    {
        $this->ole = new OleReader($file);

        if (!$this->ole->streamExists('Workbook')) {
            if ($this->ole->streamExists('Book')) {
                throw new Exception('Excel 5.0/95 workbooks (BIFF5/BIFF7) are not supported, only Excel 97-2003 (BIFF8)');
            }

            throw new Exception('Not an XLS workbook: the compound file has no Workbook stream');
        }

        $this->readGlobals();

        if ($this->sheets) {
            $this->selectFirstSheet();
        }
    }

    /**
     * Read the workbook globals substream
     *
     * @return void
     *
     * @throws Exception
     */
    protected function readGlobals(): void
    {
        $biff = $this->newBiffReader();
        $sheetIndex = 0;

        foreach ($biff->records() as $record) {
            switch ($record['type']) {
                case BiffRecord::BOF:
                    $version = unpack('v', substr($record['data'], 0, 2))[1];
                    if ($version !== BiffRecord::VERSION_BIFF8) {
                        throw new Exception(sprintf('Unsupported BIFF version 0x%04X, only BIFF8 (Excel 97-2003) is supported', $version));
                    }
                    break;

                case BiffRecord::FILEPASS:
                    throw new Exception('The workbook is encrypted, reading encrypted files is not supported');

                case BiffRecord::DATEMODE:
                    $this->date1904 = unpack('v', substr($record['data'], 0, 2))[1] === 1;
                    break;

                case BiffRecord::CODEPAGE:
                    $this->codepage = unpack('v', substr($record['data'], 0, 2))[1];
                    break;

                case BiffRecord::FORMAT:
                    $formatIndex = unpack('v', substr($record['data'], 0, 2))[1];
                    [$pattern] = BiffString::readLong($record['data'], 2);
                    $this->numberFormats[$formatIndex] = $pattern;
                    break;

                case BiffRecord::XF:
                    // cells reference XF records by their position in the file,
                    // style and cell formats sharing one sequence
                    $this->xfFormatIndex[] = unpack('v', substr($record['data'], 2, 2))[1];
                    break;

                case BiffRecord::SST:
                    $this->sharedStrings = BiffString::readSharedStrings($record['parts']);
                    break;

                case BiffRecord::BOUNDSHEET:
                    $this->addSheet($record['data'], ++$sheetIndex);
                    break;

                case BiffRecord::EOF:
                    break 2;
            }
        }

        $this->buildStyles();
    }

    /**
     * @param string $data
     * @param int $sheetId
     *
     * @return void
     */
    protected function addSheet(string $data, int $sheetId): void
    {
        $offset = unpack('V', substr($data, 0, 4))[1];
        $visibility = ord($data[4]) & 0x03;
        $sheetType = ord($data[5]);
        [$name] = BiffString::readShort($data, 6);

        // 0x00 is a worksheet; charts, macro sheets and dialogue sheets carry no cells
        if ($sheetType !== 0x00) {
            return;
        }

        $sheet = new XlsSheet($name, (string)$sheetId, $offset, $this);
        if ($visibility === 1) {
            $sheet->setState('hidden');
        }
        elseif ($visibility === 2) {
            $sheet->setState('veryHidden');
        }

        $this->sheets[$sheetId] = $sheet;
        if (!isset($this->defaultSheetId)) {
            $this->defaultSheetId = $sheetId;
        }
    }

    /**
     * Turn the collected XF records into the style table the readers expect
     *
     * Only the number format is derived for now, which is what date and number
     * detection needs. Fonts, fills and borders are also carried by XF records
     * but are not decoded yet.
     *
     * @return void
     */
    protected function buildStyles(): void
    {
        $this->styles['cellXfs'] = [];
        foreach ($this->xfFormatIndex as $formatIndex) {
            $this->styles['cellXfs'][] = $this->styleFromFormatIndex($formatIndex);
        }
    }

    /**
     * Same classification the XLSX reader applies to numFmtId plus format code
     *
     * @param int $formatIndex
     *
     * @return array|null
     */
    protected function styleFromFormatIndex(int $formatIndex): ?array
    {
        $pattern = $this->numberFormats[$formatIndex] ?? '';

        if ($this->_isDatePattern($formatIndex, $pattern)) {
            return ['format' => $pattern, 'formatType' => 'd'];
        }
        if ($pattern !== '') {
            if ($this->_isNumberPattern($formatIndex, $pattern)) {
                return ['format' => $pattern, 'formatType' => 'n'];
            }
            // XLS writers often register a custom format whose pattern is one of
            // the builtin ones - "@" in particular - instead of referencing the
            // builtin index. Classify by the pattern so that such a cell is
            // typed the same as it would be in XLSX.
            $category = $this->categoryByPattern($pattern);
            if ($category !== null && $category !== '') {
                return ['format' => $pattern, 'formatType' => $category];
            }

            return ['format' => $pattern];
        }
        if ($formatIndex > 0 && isset($this->builtinFormats[$formatIndex]['category'])) {
            return ['formatType' => $this->builtinFormats[$formatIndex]['category']];
        }

        return null;
    }

    /**
     * Category of a builtin format with the same pattern, if there is one
     *
     * @param string $pattern
     *
     * @return string|null
     */
    protected function categoryByPattern(string $pattern): ?string
    {
        foreach ($this->builtinFormats as $format) {
            if ($format['pattern'] === $pattern) {
                return $format['category'];
            }
        }

        return null;
    }

    /**
     * @return void
     */
    protected function _loadCompleteStyles()
    {
        $numFmts = [];
        foreach ($this->builtinFormats as $index => $format) {
            $numFmts[$index] = [
                'format-num-id' => $index,
                'format-pattern' => $format['pattern'],
                'format-category' => $format['category'],
            ];
        }
        foreach ($this->numberFormats as $index => $pattern) {
            $numFmts[$index] = [
                'format-num-id' => $index,
                'format-pattern' => $pattern,
                'format-category' => $this->_isDatePattern($index, $pattern) ? 'date' : '',
            ];
        }

        $cellXfs = [];
        foreach ($this->xfFormatIndex as $formatIndex) {
            $cellXfs[] = ['numFmtId' => $formatIndex];
        }

        $this->styles['_'] = [
            'numFmts' => $numFmts,
            'fonts' => [],
            'fills' => [],
            'borders' => [],
            'cellStyleXfs' => [],
            'cellXfs' => $cellXfs,
        ];
    }

    /**
     * Code page declared by the workbook
     *
     * @return int
     */
    public function codepage(): int
    {
        return $this->codepage;
    }
}
