<?php

namespace avadim\FastExcelReader;

/**
 * Class Excel
 *
 * @package avadim\FastExcelReader
 */
class Excel
{
    public const EXCEL_2007_MAX_ROW = 1048576;
    public const EXCEL_2007_MAX_COL = 16384;

    public const KEYS_ORIGINAL = 0;
    public const KEYS_FIRST_ROW = 1;
    public const KEYS_ROW_ZERO_BASED = 2;
    public const KEYS_COL_ZERO_BASED = 4;
    public const KEYS_ZERO_BASED = 6;
    public const KEYS_ROW_ONE_BASED = 8;
    public const KEYS_COL_ONE_BASED = 16;
    public const KEYS_ONE_BASED = 24;
    public const KEYS_RELATIVE = 32;
    public const KEYS_SWAP = 64;

    protected string $file;

    /** @var Reader */
    protected Reader $xmlReader;

    protected array $fileList = [];

    protected array $relations = [];

    protected array $sharedStrings = [];

    protected array $styles = [];

    /** @var Sheet[] */
    protected array $sheets = [];

    protected int $defaultSheetId;

    protected ?string $dateFormat = null;

    /**
     * Excel constructor
     *
     * @param string|null $file
     */
    public function __construct(string $file = null)
    {
        if ($file) {
            $this->file = $file;
            $this->_prepare($file);
        }
    }

    /**
     * @param string $file
     */
    protected function _prepare(string $file)
    {
        $this->xmlReader = new Reader($file);
        $this->fileList = $this->xmlReader->fileList();
        foreach ($this->fileList as $fileName) {
            if (strpos($fileName, 'xl/drawings/drawing') === 0) {
                $this->relations['drawings'][] = $fileName;
            }
            elseif (strpos($fileName, 'xl/media/') === 0) {
                $this->relations['media'][] = $fileName;
            }
            elseif (strpos($fileName, 'xl/theme/') === 0) {
                $this->relations['theme'][] = $fileName;
            }
        }

        $innerFile = 'xl/_rels/workbook.xml.rels';
        $this->xmlReader->openZip($innerFile);
        while ($this->xmlReader->read()) {
            if ($this->xmlReader->nodeType === \XMLReader::ELEMENT && $this->xmlReader->name === 'Relationship') {
                $type = basename($this->xmlReader->getAttribute('Type'));
                if ($type) {
                    $this->relations[$type][$this->xmlReader->getAttribute('Id')] = 'xl/' . $this->xmlReader->getAttribute('Target');
                }
            }
        }
        $this->xmlReader->close();

        if (isset($this->relations['worksheet'])) {
            $this->_loadSheets();
        }
        if (isset($this->relations['sharedStrings'])) {
            $this->_loadSharedStrings(reset($this->relations['sharedStrings']));
        }
        if (isset($this->relations['styles'])) {
            $this->_loadStyles(reset($this->relations['styles']));
        }

        if ($this->sheets) {
            // set current sheet
            $this->selectFirstSheet();
        }
    }

    /**
     * @param string|null $innerFile
     */
    protected function _loadSheets(string $innerFile = null)
    {
        if (!$innerFile) {
            $innerFile = 'xl/workbook.xml';
        }
        $this->xmlReader->openZip($innerFile);
        $sheetCnt = count($this->relations['worksheet']);
        while ($this->xmlReader->read()) {
            if ($this->xmlReader->nodeType === \XMLReader::ELEMENT && $this->xmlReader->name === 'sheet') {
                $rId = $this->xmlReader->getAttribute('r:id');
                $sheetId = $this->xmlReader->getAttribute('sheetId');
                $path = $this->relations['worksheet'][$rId];
                if ($path) {
                    $sheetName = $this->xmlReader->getAttribute('name');
                    $this->sheets[$sheetId] = new Sheet($this->file, $sheetId, $sheetName, $this->relations['worksheet'][$rId]);
                    $this->sheets[$sheetId]->excel = $this;
                }
                if (--$sheetCnt < 1) {
                    break;
                }
            }
        }
        $this->xmlReader->close();
    }

    /**
     * @param string|null $innerFile
     */
    protected function _loadSharedStrings(string $innerFile = null)
    {
        if (!$innerFile) {
            $innerFile = 'xl/sharedStrings.xml';
        }
        $this->xmlReader->openZip($innerFile);
        while ($this->xmlReader->read()) {
            if ($this->xmlReader->nodeType === \XMLReader::ELEMENT && $this->xmlReader->name === 'si' && $node = $this->xmlReader->expand()) {
                $this->sharedStrings[] = $node->textContent;
            }
        }
        $this->xmlReader->close();
    }

    /**
     * @param string|null $innerFile
     */
    protected function _loadStyles(string $innerFile = null)
    {
        if (!$innerFile) {
            $innerFile = 'xl/styles.xml';
        }
        $this->xmlReader->openZip($innerFile);
        $styleType = '';
        while ($this->xmlReader->read()) {
            if ($this->xmlReader->nodeType === \XMLReader::ELEMENT) {
                if ($this->xmlReader->name === 'cellStyleXfs' || $this->xmlReader->name === 'cellXfs') {
                    $styleType = $this->xmlReader->name;
                    continue;
                }
                if ($this->xmlReader->name === 'numFmt') {
                    $numFmtId = (int)$this->xmlReader->getAttribute('numFmtId');
                    $formatCode = $this->xmlReader->getAttribute('formatCode');
                    $numFmts[$numFmtId] = $formatCode;
                }
                elseif ($this->xmlReader->name === 'xf') {
                    $numFmtId = (int)$this->xmlReader->getAttribute('numFmtId');
                    if (isset($numFmts[$numFmtId])) {
                        $format = $numFmts[$numFmtId];
                        if (strpos($format, 'M') !== false || strpos($format, 'm') !== false) {
                            $this->styles[$styleType][] = ['format' => $numFmts[$numFmtId], 'formatType' => 'd'];
                        }
                        else {
                            $this->styles[$styleType][] = ['format' => $numFmts[$numFmtId]];
                        }
                    }
                    elseif (($numFmtId >= 14 && $numFmtId <= 22) || ($numFmtId >= 45 && $numFmtId <= 47)) {
                            $this->styles[$styleType][] = ['formatType' => 'd'];
                    }
                    else {
                        $this->styles[$styleType][] = null;
                    }
                }
            }
        }
        $this->xmlReader->close();
    }

    /**
     * Open XLSX file
     *
     * @param string $file
     *
     * @return Excel
     */
    public static function open(string $file): Excel
    {
        return new self($file);
    }

    /**
     * Converts an alphabetic column index to a numeric
     *
     * @param string $colLetter
     *
     * @return int
     */
    public static function colNum(string $colLetter): int
    {
        static $colNumbers = [];

        if (isset($colNumbers[$colLetter])) {
            return $colNumbers[$colLetter];
        }
        // Strip cell reference down to just letters
        $letters = preg_replace('/[^A-Z]/', '', strtoupper($colLetter));

        if (strlen($letters) >= 3 && $letters > 'XFD') {
            return self::EXCEL_2007_MAX_COL;
        }
        // Iterate through each letter, starting at the back to increment the value
        for ($index = 0, $i = 0; $letters !== ''; $letters = substr($letters, 0, -1), $i++) {
            $index += (ord(substr($letters, -1)) - 64) * (26 ** $i);
        }

        $colNumbers[$colLetter] = ($index <= self::EXCEL_2007_MAX_COL) ? (int)$index: self::EXCEL_2007_MAX_COL;

        return $colNumbers[$colLetter];
    }

    /**
     * Convert column number to letter
     *
     * @param int $colNumber ONE based
     *
     * @return string
     */
    public static function colLetter(int $colNumber): string
    {
        static $colLetters = ['',
            'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
            'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ',
        ];

        if (isset($colLetters[$colNumber])) {
            return $colLetters[$colNumber];
        }

        if ($colNumber > 0 && $colNumber <= self::EXCEL_2007_MAX_COL) {
            $num = $colNumber - 1;
            for ($letter = ''; $num >= 0; $num = (int)($num / 26) - 1) {
                $letter = chr($num % 26 + 0x41) . $letter;
            }
            $colLetters[$colNumber] = $letter;

            return $letter;
        }

        return '';
    }

    /**
     * @param $excelDateTime
     *
     * @return int
     */
    public static function timestamp($excelDateTime): int
    {
        $d = floor($excelDateTime);
        $t = $excelDateTime - $d;
        // $d += 1462; // days since 1904

        $t = (abs($d) > 0) ? ($d - 25569) * 86400 + round($t * 86400) : round($t * 86400);

        return (int)$t;
    }

    /**
     * @param string $dateFormat
     *
     * @return $this
     */
    public function setDateFormat(string $dateFormat): Excel
    {
        $this->dateFormat = $dateFormat;

        return $this;
    }

    /**
     * @return string|null
     */
    public function getDateFormat(): ?string
    {
        return $this->dateFormat;
    }

    public function formatDate($value, $format = null)
    {
        if ($format || $this->dateFormat) {
            return gmdate($format ?: $this->dateFormat, $value);
        }

        return $value;
    }

    /**
     * Returns style array by style Idx
     *
     * @param $styleIdx
     *
     * @return array
     */
    public function styleByIdx($styleIdx): array
    {
        return $this->styles['cellXfs'][$styleIdx] ?? [];
    }

    /**
     * Returns string array by index
     *
     * @param $stringId
     *
     * @return string|null
     */
    public function sharedString($stringId): ?string
    {
        return $this->sharedStrings[$stringId] ?? null;
    }

    /**
     * Returns names array of all sheets
     *
     * @return array
     */
    public function getSheetNames(): array
    {
        $result = [];
        foreach ($this->sheets as $sheetId => $sheet) {
            $result[$sheetId] = $sheet->name();
        }
        return $result;
    }

    /**
     * Returns current or specified sheet
     *
     * @param string|null $name
     *
     * @return Sheet|null
     */
    public function sheet(?string $name = null): ?Sheet
    {
        $sheetId = null;
        if (!$name) {
            $sheetId = $this->defaultSheetId;
        }
        else {
            foreach ($this->sheets as $sheetId => $sheet) {
                if ($sheet->isName($name)) {
                    break;
                }
            }
        }
        if ($sheetId && isset($this->sheets[$sheetId])) {
            return $this->sheets[$sheetId];
        }

        return null;
    }

    /**
     * Select default sheet by name
     *
     * @param string $name
     * @param string|null $areaRange
     * @param bool|null $firstRowKeys
     *
     * @return Sheet
     */
    public function selectSheet(string $name, string $areaRange = null, ?bool $firstRowKeys = false): Sheet
    {
        foreach ($this->sheets as $sheetId => $sheet) {
            if ($sheet->isName($name)) {
                $this->defaultSheetId = $sheetId;
                if ($areaRange) {
                    $sheet->setReadArea($areaRange, $firstRowKeys);
                }

                return $sheet;
            }
        }
        throw new Exception('Sheet name "' . $name . '" not found');
    }

    /**
     * Select default sheet by ID
     *
     * @param int $sheetId
     * @param string|null $areaRange
     * @param bool|null $firstRowKeys
     *
     * @return Sheet
     */
    public function selectSheetById(int $sheetId, string $areaRange = null, ?bool $firstRowKeys = false): Sheet
    {
        if (!isset($this->sheets[$sheetId])) {
            throw new Exception('Sheet ID "' . $sheetId . '" not found');
        }
        $this->defaultSheetId = $sheetId;
        if ($areaRange) {
            $this->sheets[$sheetId]->setReadArea($areaRange, $firstRowKeys);
        }

        return $this->sheets[$sheetId];
    }

    /**
     * Select the first sheet as default
     *
     * @param string|null $areaRange
     * @param bool|null $firstRowKeys
     *
     * @return Sheet
     */
    public function selectFirstSheet(string $areaRange = null, ?bool $firstRowKeys = false): Sheet
    {
        $this->defaultSheetId = array_key_first($this->sheets);
        if ($areaRange) {
            $this->sheets[$this->defaultSheetId]->setReadArea($areaRange, $firstRowKeys);
        }

        return $this->sheets[$this->defaultSheetId];
    }

    public function setReadArea(string $areaRange, ?bool $firstRowKeys = false): Sheet
    {
        return $this->sheets[$this->defaultSheetId]->setReadArea($areaRange, $firstRowKeys);
    }

    /**
     * Reads cell values and passes them to a callback function
     *
     * @param callback $callback
     * @param int|null $indexStyle
     */
    public function readCallback(callable $callback, int $indexStyle = null)
    {
        $this->sheets[$this->defaultSheetId]->readCallback($callback, $indexStyle);
    }

    /**
     * Returns cell values as a two-dimensional array from default sheet [row][col]
     *
     *  readRows()
     *  readRows(true)
     *  readRows(false, Excel::INDEX_ZERO_BASED)
     *  readRows(Excel::INDEX_ZERO_BASED | Excel::INDEX_RELATIVE)
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $indexStyle
     *
     * @return array
     */
    public function readRows($columnKeys = [], int $indexStyle = null): array
    {
        return $this->sheets[$this->defaultSheetId]->readRows($columnKeys, $indexStyle);
    }

    /**
     * Returns the values of all cells as array
     *
     * @return array
     */
    public function readCells(): array
    {
        return $this->sheets[$this->defaultSheetId]->readCells();
    }

    /**
     * Returns cell values as a two-dimensional array from default sheet [col][row]
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $indexStyle
     *
     * @return array
     */
    public function readColumns($columnKeys = null, int $indexStyle = null): array
    {
        return $this->sheets[$this->defaultSheetId]->readColumns($columnKeys, $indexStyle);
    }

    public function innerFileList(): array
    {
        return $this->fileList;
    }

    public function hasDrawings(): bool
    {
        return !empty($this->relations['drawings']);
    }

    public function countImages(): int
    {
        $result = 0;
        if ($this->hasDrawings()) {
            foreach ($this->sheets as $sheet) {
                $result += $sheet->countImages();
            }
        }

        return $result;
    }

    public function getImageList(): array
    {
        $result = [];
        if ($this->hasDrawings()) {
            foreach ($this->sheets as $sheet) {
                $result[$sheet->name()] = $sheet->getImageList();
            }
        }

        return $result;
    }
}

// EOF