<?php

namespace avadim\FastExcelReader;

class Sheet
{
    public Excel $excel;

    protected string $file;

    protected string $sheetId;

    protected string $name;

    protected string $path;

    protected array $area = [];

    /** @var Reader */
    protected $xmlReader;

    public function __construct($file, $sheetId, $name, $path)
    {
        $this->file = $file;
        $this->sheetId = $sheetId;
        $this->name = $name;
        $this->path = $path;

        $this->area = [
            'row_min' => 1,
            'col_min' => 1,
            'row_max' => Excel::EXCEL_2007_MAX_ROW,
            'col_max' => Excel::EXCEL_2007_MAX_COL,
            'first_row' => false,
        ];
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
            if ($styleIdx > 0 && ($style = $this->excel->styleByIdx($styleIdx))) {
                $format = $style['format'] ?? null;
                if (isset($style['formatType'])) {
                    $dataType = $style['formatType'];
                }
            }
        }

        $value = '';

        switch ( $dataType ) {
            case 's':
                // Value is a shared string
                if (is_numeric($cellValue) && ($str = $this->excel->sharedString((int)$cellValue)) !== null) {
                    $value = $str;
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
                    $value = $this->excel->dateFormat(Excel::timestamp($cellValue));
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


    public function name(): string
    {
        return $this->name;
    }
    
    
    public function isName($name): bool
    {
        return strcasecmp($this->name, $name) === 0;
    }
    
    protected function getReader($file = null)
    {
        if (empty($this->xmlReader)) {
            if (!$file) {
                $file = $this->file;
            }
            $this->xmlReader = new Reader($file);
        }

        return $this->xmlReader;
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
    public function setReadArea(string $areaRange, ?bool $firstRowKeys = false): Sheet
    {
        if (preg_match('/^([A-Z]+)(\d+)(:([A-Z]+)(\d+))?$/', $areaRange, $matches)) {
            $this->area['col_min'] = Excel::colNum($matches[1]);
            $this->area['row_min'] = (int)$matches[2];
            if (empty($matches[3])) {
                $this->area['col_max'] = Excel::EXCEL_2007_MAX_COL;
                $this->area['row_max'] = Excel::EXCEL_2007_MAX_ROW;
            }
            else {
                $this->area['col_max'] = Excel::colNum($matches[4]);
                $this->area['row_max'] = (int)$matches[5];
            }
            $this->area['first_row'] = $firstRowKeys;

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
    public function setReadAreaColumns(string $columnsRange, ?bool $firstRowKeys = false): Sheet
    {
        if (preg_match('/^([A-Z]+)(:([A-Z]+))?$/', $columnsRange, $matches)) {
            $this->area['col_min'] = Excel::colNum($matches[1]);
            if (empty($matches[2])) {
                $this->area['col_max'] = Excel::EXCEL_2007_MAX_COL;
            }
            else {
                $this->area['col_max'] = Excel::colNum($matches[3]);
            }
            $this->area['first_row'] = $firstRowKeys;

            return $this;
        }
        throw new Exception('Wrong address or range "' . $columnsRange . '"');
    }

    /**
     * Returns cell values as a two-dimensional array
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $indexStyle
     *
     * @return array
     */
    public function readRows($columnKeys = [], int $indexStyle = null): array
    {
        if (!is_array($columnKeys)) {
            if (is_int($columnKeys) && $columnKeys > 1 && $indexStyle === null) {
                $firstRowKeys = $columnKeys & Excel::KEYS_FIRST_ROW;
                $indexStyle = $columnKeys;
            }
            else {
                $firstRowKeys = (bool)$columnKeys;
            }
            $columnKeys = [];
        }
        elseif (is_int($indexStyle) && $indexStyle & Excel::KEYS_FIRST_ROW) {
            $firstRowKeys = true;
        }
        else {
            $firstRowKeys = null;
        }

        $data = [];
        if ($firstRowKeys === null) {
            $firstRowKeys = !empty($this->area['first_row']);
        }
        if ($columnKeys) {
            $columnKeys = array_combine(array_map('strtoupper', array_keys($columnKeys)), array_values($columnKeys));
        }
        if ($firstRowKeys) {
            $indexStyle = (int)$indexStyle | Excel::KEYS_FIRST_ROW;
        }
        $this->readCallback(static function($row, $col, $val) use (&$firstRowKeys, &$columnKeys, &$data) {
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
        }, $indexStyle);

        if ($data && ($indexStyle & Excel::KEYS_SWAP)) {
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
            $columnKeys = $columnKeys & Excel::KEYS_FIRST_ROW;
        }
        else {
            $indexStyle = $indexStyle | Excel::KEYS_RELATIVE;
        }

        return $this->readRows($columnKeys, $indexStyle | Excel::KEYS_SWAP);
    }

    /**
     * Returns the values of all cells as array
     *
     * @return array
     */
    public function readCells(): array
    {
        $data = [];
        $this->readCallback(static function($row, $col, $val) use (&$data) {
            $data[$col . $row] = $val;
        });

        return $data;
    }

    /**
     * Reads cell values and passes them to a callback function
     *
     * @param callback $callback
     * @param int|null $indexStyle
     *
     * @return array
     */
    public function readCallback(callable $callback, int $indexStyle = null): array
    {
        $xmlReader = $this->getReader();
        $xmlReader->openZip($this->path);
        $readArea = $this->area;

        $data = [];
        $rowNum = 0;
        $rowOffset = $colOffset = -1;
        if ($xmlReader->seekOpenTag('sheetData')) {
            while ($xmlReader->read()) {
                if ($xmlReader->nodeType === \XMLReader::END_ELEMENT && $xmlReader->name === 'sheetData') {
                    break;
                }
                if ($xmlReader->nodeType === \XMLReader::ELEMENT) {
                    if ($xmlReader->name === 'row') {
                        $rowNum = (int)$xmlReader->getAttribute('r');
                        if ($rowOffset === -1) {
                            $rowOffset = $rowNum - 1;
                        }
                    }
                    elseif ($xmlReader->name === 'c') {
                        $addr = $xmlReader->getAttribute('r');
                        if ($addr && preg_match('/^([A-Z]+)(\d+)$/', $addr, $m)) {
                            $col = $m[1];
                            $colNum = Excel::colNum($col);
                            if ($colNum >= $readArea['col_min'] && $colNum <= $readArea['col_max']
                                && $rowNum >= $readArea['row_min'] && $rowNum <= $readArea['row_max']) {
                                if ($colOffset === -1) {
                                    $colOffset = $colNum - 1;
                                }
                                $cell = $xmlReader->expand();
                                if ($indexStyle & Excel::KEYS_ROW_ZERO_BASED) {
                                    $row = $rowNum - (($indexStyle & Excel::KEYS_FIRST_ROW) ? 2 : 1);
                                }
                                elseif ($indexStyle & Excel::KEYS_ROW_ONE_BASED) {
                                    $row = $rowNum - (($indexStyle & Excel::KEYS_FIRST_ROW) ? 0 : 1);
                                }
                                else {
                                    $row = (string)$rowNum;
                                }
                                if ($indexStyle & Excel::KEYS_COL_ZERO_BASED) {
                                    $col = $colNum - 1;
                                }
                                elseif ($indexStyle & Excel::KEYS_COL_ONE_BASED) {
                                    $col = $colNum;
                                }
                                if (($indexStyle & Excel::KEYS_RELATIVE)
                                    && (($indexStyle & Excel::KEYS_ROW_ZERO_BASED) || ($indexStyle & Excel::KEYS_ROW_ONE_BASED))
                                    && (($indexStyle & Excel::KEYS_COL_ZERO_BASED) || ($indexStyle & Excel::KEYS_COL_ONE_BASED))
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
        $xmlReader->close();

        return $data;
    }


}