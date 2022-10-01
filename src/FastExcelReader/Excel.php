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

    /** @var Reader */
    protected $xmlReader;

    protected $relations = [];

    protected $sharedStrings = [];

    protected $styles = [];

    protected $sheets = [];

    protected $defaultSheet;

    protected $dateFormat;

    /**
     * Excel constructor
     *
     * @param string|null $file
     */
    public function __construct(string $file = null)
    {
        if ($file) {
            $this->_prepare($file);
        }
    }

    /**
     * @param string $file
     */
    protected function _prepare(string $file)
    {
        $this->xmlReader = new Reader($file);

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
            $this->defaultSheet = key($this->sheets);
            foreach ($this->sheets as $sheetId => $sheet) {
                $this->sheets[$sheetId]['area'] = [
                    'row_min' => 1,
                    'col_min' => 1,
                    'row_max' => self::EXCEL_2007_MAX_ROW,
                    'col_max' => self::EXCEL_2007_MAX_COL,
                    'first_row' => false,
                ];
            }
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
                    $this->sheets[$sheetId] = [
                        'name' => $this->xmlReader->getAttribute('name'),
                        'path' => $this->relations['worksheet'][$rId],
                        'sheet_id' => $sheetId,
                    ];
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
                } elseif ($this->xmlReader->name === 'xf') {
                    $numFmtId = (int)$this->xmlReader->getAttribute('numFmtId');
                    if (isset($numFmts[$numFmtId])) {
                        $format = $numFmts[$numFmtId];
                        if (strpos($format, 'M') !== false || strpos($format, 'm') !== false) {
                            $this->styles[$styleType][] = ['format' => $numFmts[$numFmtId], 'formatType' => 'd'];
                        } else {
                            $this->styles[$styleType][] = ['format' => $numFmts[$numFmtId]];
                        }
                    } elseif (($numFmtId >= 14 && $numFmtId <= 22) || ($numFmtId >= 45 && $numFmtId <= 47)) {
                            $this->styles[$styleType][] = ['formatType' => 'd'];
                    } else {
                        $this->styles[$styleType][] = null;
                    }
                }
            }
        }
        $this->xmlReader->close();
    }

    /**
     * @param $excelDateTime
     *
     * @return int
     */
    protected function _timestamp($excelDateTime): int
    {
        $d = floor($excelDateTime);
        $t = $excelDateTime - $d;
        // $d += 1462; // days since 1904

        $t = (abs($d) > 0) ? ($d - 25569) * 86400 + round($t * 86400) : round($t * 86400);

        return (int)$t;
    }

    /**
     * @param $cell
     *
     * @return mixed
     */
    protected function _cellValue($cell)
    {
        // Determine data type
        $dataType = (string)$cell->getAttribute('t');
        $cellValue = null;
        foreach($cell->childNodes as $node) {
            if ($node->nodeName === 'v') {
                $cellValue = $node->nodeValue;
                break;
            }
        }

        $format = null;
        if ( $dataType === '' || $dataType === 'n' ) { // number
            $styleIdx = (int)$cell->getAttribute('s');
            if ($styleIdx > 0) {
                $format = $this->styles['cellXfs'][$styleIdx]['format'] ?? null;
                if (isset($this->styles['cellXfs'][$styleIdx]['formatType'])) {
                    $dataType = $this->styles['cellXfs'][$styleIdx]['formatType'];
                }
            }
        }

        $value = '';

        switch ( $dataType ) {
            case 's':
                // Value is a shared string
                if (is_numeric($cellValue) && isset($this->sharedStrings[(int)$cellValue])) {
                    $value = $this->sharedStrings[(int)$cellValue];
                }
                break;

            case 'b':
                // Value is boolean
                $value = (bool)$cellValue;
                break;

            case 'inlineStr':
                // Value is rich text inline
                $value = $cell->textContent;
                break;

            case 'e':
                // Value is an error message
                $value = (string)$cellValue;
                break;

            case 'd':
                // Value is a date and non-empty
                if (!empty($cellValue)) {
                    $value = $this->_timestamp($cellValue);
                    if ($this->dateFormat) {
                        $value = gmdate($this->dateFormat, $value);
                    }
                }
                break;

            default:
                // Value is a string
                $value = (string) $cellValue;

                // Check for numeric values
                if (is_numeric($value) && $dataType !== 's') {
                    /** @noinspection TypeUnsafeComparisonInspection */
                    if ($value == (int)$value) {
                        $value = (int)$value;
                    }
                    /** @noinspection TypeUnsafeComparisonInspection */
                    elseif ($value == (float)$value) {
                        $value = (float)$value;
                    }
                }
        }

        return $value;
    }

    /**
     * @param string $colLetter
     *
     * @return int
     */
    public static function colNum(string $colLetter): int
    {
        static $colIndex = [];

        if (isset($colIndex[$colLetter])) {
            return $colIndex[$colLetter];
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

        $colIndex[$colLetter] = ($index <= self::EXCEL_2007_MAX_COL) ? (int)$index: self::EXCEL_2007_MAX_COL;

        return $colIndex[$colLetter];
    }

    /**
     * Returns names of all sheets
     *
     * @return array
     */
    public function getSheetNames(): array
    {
        return array_column($this->sheets, 'name', 'sheet_id');
    }

    /**
     * @param $dateFormat
     *
     * @return $this
     */
    public function setDateFormat($dateFormat): Excel
    {
        $this->dateFormat = $dateFormat;

        return $this;
    }

    /**
     * Select default sheet by name
     *
     * @param string $name
     * @param string|null $areaRange
     * @param bool|null $firstRowKeys
     *
     * @return $this
     */
    public function selectSheet(string $name, string $areaRange = null, ?bool $firstRowKeys = false)
    {
        foreach ($this->sheets as $sheetId => $sheet) {
            if (strcasecmp($sheet['name'], $name) === 0) {
                $this->defaultSheet = $sheetId;

                if ($areaRange) {
                    $this->setReadArea($areaRange, $firstRowKeys);
                }

                return $this;
            }
        }
        throw new Exception('Sheet name "' . $name . '" not found');
    }

    /**
     * Select default sheet by ID
     *
     * @param int $sheetId
     * @param string|null $areaRange
     *
     * @return $this
     */
    public function selectSheetById(int $sheetId, string $areaRange = null)
    {
        if (!isset($this->sheets[$sheetId])) {
            throw new Exception('Sheet ID "' . $sheetId . '" not found');
        }
        $this->defaultSheet = $sheetId;

        if ($areaRange) {
            $this->setReadArea($areaRange);
        }

        return $this;
    }

    /**
     * Select the first sheet as default
     *
     * @return $this
     */
    public function selectFirstSheet()
    {
        reset($this->sheets);
        $this->defaultSheet = key($this->sheets);

        return $this;
    }

    /**
     * setReadArea('C3:AZ28') - set top left and right bottom of read area
     * setReadArea('C3') - set top left only
     *
     * @param string $areaRange
     * @param bool|null $firstRowKeys
     *
     * @return $this
     */
    public function setReadArea(string $areaRange, ?bool $firstRowKeys = false)
    {
        if (preg_match('/^([A-Z]+)(\d+)(:([A-Z]+)(\d+))?$/', $areaRange, $matches)) {
            $this->sheets[$this->defaultSheet]['area']['col_min'] = self::colNum($matches[1]);
            $this->sheets[$this->defaultSheet]['area']['row_min'] = (int)$matches[2];
            if (empty($matches[3])) {
                $this->sheets[$this->defaultSheet]['area']['col_max'] = self::EXCEL_2007_MAX_COL;
                $this->sheets[$this->defaultSheet]['area']['row_max'] = self::EXCEL_2007_MAX_ROW;
            }
            else {
                $this->sheets[$this->defaultSheet]['area']['col_max'] = self::colNum($matches[4]);
                $this->sheets[$this->defaultSheet]['area']['row_max'] = (int)$matches[5];
            }
            $this->sheets[$this->defaultSheet]['area']['first_row'] = $firstRowKeys;

            return $this;
        }
        throw new Exception('Wrong address or range "' . $areaRange . '"');
    }

    /**
     * setReadArea('C:AZ') - set left and right columns of read area
     * setReadArea('C') - set left column only
     *
     * @param string $columnsRange
     * @param bool|null $firstRowKeys
     *
     * @return $this
     */
    public function setReadAreaColumns(string $columnsRange, ?bool $firstRowKeys = false)
    {
        if (preg_match('/^([A-Z]+)(:([A-Z]+))?$/', $columnsRange, $matches)) {
            $this->sheets[$this->defaultSheet]['area']['col_min'] = self::colNum($matches[1]);
            if (empty($matches[2])) {
                $this->sheets[$this->defaultSheet]['area']['col_max'] = self::EXCEL_2007_MAX_COL;
            }
            else {
                $this->sheets[$this->defaultSheet]['area']['col_max'] = self::colNum($matches[3]);
            }
            $this->sheets[$this->defaultSheet]['area']['first_row'] = $firstRowKeys;

            return $this;
        }
        throw new Exception('Wrong address or range "' . $columnsRange . '"');
    }

    /**
     * Reads cell values and passes them to a callback function
     *
     * @param callback $callback
     * @param string|int $sheetId
     * @param int|null $indexStyle
     *
     * @return array
     */
    public function readSheetCallback(callable $callback, $sheetId = null, int $indexStyle = null): array
    {
        if (null === $sheetId) {
            $sheetId = $this->defaultSheet;
        }
        elseif (!isset($this->sheets[$sheetId])) {
            throw new Exception('Sheet ID "' . $sheetId . '" not found');
        }

        $this->xmlReader->openZip($this->sheets[$sheetId]['path']);
        $readArea = $this->sheets[$sheetId]['area'];

        $data = [];
        $rowNum = 0;
        $rowOffset = $colOffset = -1;
        if ($this->xmlReader->seekOpenTag('sheetData')) {
            while ($this->xmlReader->read()) {
                if ($this->xmlReader->nodeType === \XMLReader::END_ELEMENT && $this->xmlReader->name === 'sheetData') {
                    break;
                }
                if ($this->xmlReader->nodeType === \XMLReader::ELEMENT) {
                    if ($this->xmlReader->name === 'row') {
                        $rowNum = (int)$this->xmlReader->getAttribute('r');
                        if ($rowOffset === -1) {
                            $rowOffset = $rowNum - 1;
                        }
                    }
                    elseif ($this->xmlReader->name === 'c') {
                        $addr = $this->xmlReader->getAttribute('r');
                        if ($addr && preg_match('/^([A-Z]+)(\d+)$/', $addr, $m)) {
                            $col = $m[1];
                            $colNum = self::colNum($col);
                            if ($colNum >= $readArea['col_min'] && $colNum <= $readArea['col_max']
                                && $rowNum >= $readArea['row_min'] && $rowNum <= $readArea['row_max']) {
                                if ($colOffset === -1) {
                                    $colOffset = $colNum - 1;
                                }
                                $cell = $this->xmlReader->expand();
                                if ($indexStyle & self::KEYS_ROW_ZERO_BASED) {
                                    $row = $rowNum - (($indexStyle & self::KEYS_FIRST_ROW) ? 2 : 1);
                                }
                                elseif ($indexStyle & self::KEYS_ROW_ONE_BASED) {
                                    $row = $rowNum - (($indexStyle & self::KEYS_FIRST_ROW) ? 0 : 1);
                                }
                                else {
                                    $row = (string)$rowNum;
                                }
                                if ($indexStyle & self::KEYS_COL_ZERO_BASED) {
                                    $col = $colNum - 1;
                                }
                                elseif ($indexStyle & self::KEYS_COL_ONE_BASED) {
                                    $col = $colNum;
                                }
                                if (($indexStyle & self::KEYS_RELATIVE)
                                    && (($indexStyle & self::KEYS_ROW_ZERO_BASED) || ($indexStyle & self::KEYS_ROW_ONE_BASED))
                                    && (($indexStyle & self::KEYS_COL_ZERO_BASED) || ($indexStyle & self::KEYS_COL_ONE_BASED))
                                ) {
                                    $row -= $rowOffset;
                                    $col -= $colOffset;
                                }
                                $needBreak = $callback($row, $col, $this->_cellValue($cell));
                                if ($needBreak) {
                                    break;
                                }
                            }
                        }
                    }
                }
            }
        }
        $this->xmlReader->close();

        return $data;
    }

    /**
     * Returns cell values as a two-dimensional array
     *
     * @param int|null $sheetId
     * @param array $columnKeys
     * @param bool|null $firstRowKeys
     * @param int|null $indexStyle
     *
     * @return array
     */
    public function readSheetRows(int $sheetId = null, array $columnKeys = [], bool $firstRowKeys = null, int $indexStyle = null): array
    {
        $data = [];
        if ($firstRowKeys === null) {
            $firstRowKeys = !empty($this->sheets[$sheetId]['area']['first_row']);
        }
        if ($columnKeys) {
            $columnKeys = array_combine(array_map('strtoupper', array_keys($columnKeys)), array_values($columnKeys));
        }
        if ($firstRowKeys) {
            $indexStyle = (int)$indexStyle | self::KEYS_FIRST_ROW;
        }
        $this->readSheetCallback(static function($row, $col, $val) use (&$firstRowKeys, &$columnKeys, &$data) {
            static $firstRowNum = null;

            if ($firstRowKeys) {
                if ($firstRowNum === null) {
                    // the first call
                    $firstRowNum = $row;
                }
                elseif ($firstRowNum < $row) {
                    if ($columnKeys) {
                        $columnKeys = array_merge($data[$firstRowNum], $columnKeys);
                    }
                    else {
                        $columnKeys = $data[$firstRowNum];
                    }
                    unset($data[$firstRowNum]);
                    $firstRowKeys = false;
                }

            }
            if (isset($columnKeys[$col])) {
                $data[$row][$columnKeys[$col]] = $val;
            }
            else {
                $data[$row][$col] = $val;
            }
        }, $sheetId, $indexStyle);

        if ($data && ($indexStyle & self::KEYS_SWAP)) {
            $newData = [];
            $rowKeys = array_keys($data);
            foreach (array_keys(reset($data)) as $colKey) {
                $newData[$colKey] = array_combine($rowKeys, array_column($data, $colKey));
            }
            return $newData;
        }

        return $data;
    }

    /**
     * Returns the values of all cells as array
     *
     * @param $sheetId
     *
     * @return array
     */
    public function readSheetCells($sheetId = null)
    {
        $data = [];
        $this->readSheetCallback(static function($row, $col, $val) use (&$data) {
            $data[$col . $row] = $val;
        }, $sheetId);

        return $data;
    }

    /**
     * Returns cell values as a two-dimensional array from default sheet [row][col]
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
    public function readRows($columnKeys = null, int $indexStyle = null): array
    {
        if (!is_array($columnKeys)) {
            if (is_int($columnKeys) && $columnKeys > 1 && $indexStyle === null) {
                $firstRowKeys = $columnKeys & self::KEYS_FIRST_ROW;
                $indexStyle = $columnKeys;
            }
            else {
                $firstRowKeys = (bool)$columnKeys;
            }
            $columnKeys = [];
        }
        elseif (is_int($indexStyle) && $indexStyle & self::KEYS_FIRST_ROW) {
            $firstRowKeys = true;
        }
        else {
            $firstRowKeys = null;
        }
        return $this->readSheetRows($this->defaultSheet, $columnKeys, $firstRowKeys, $indexStyle);
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
        if (is_int($columnKeys) && $columnKeys > 1 && $indexStyle === null) {
            $indexStyle = $columnKeys | Excel::KEYS_RELATIVE;
            $columnKeys = $columnKeys & self::KEYS_FIRST_ROW;
        }
        else {
            $indexStyle = $indexStyle | Excel::KEYS_RELATIVE;
        }

        return $this->readRows($columnKeys, $indexStyle | Excel::KEYS_SWAP);
    }

    /**
     * Returns the values of all cells as array from default sheet
     *
     * @return array
     */
    public function readCells(): array
    {
        return $this->readSheetCells($this->defaultSheet);
    }

    /**
     * Open XLSX file
     *
     * @param string $file
     *
     * @return Excel
     */
    public static function open(string $file)
    {
        return new self($file);
    }

}

// EOF