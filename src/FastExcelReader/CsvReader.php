<?php

namespace avadim\FastExcelReader;

use avadim\FastExcelHelper\Helper;

class CsvReader
{
    protected string $file;
    protected string $delimiter = ',';
    protected string $enclosure = '"';
    protected string $escape = '\\';
    protected ?string $encoding = null;

    /**
     * CsvReader constructor
     *
     * @param string $file
     * @param CsvOptions|array|null $options
     */
    public function __construct(string $file, $options = [])
    {
        if (!file_exists($file)) {
            throw new Exception("File $file not found");
        }
        $this->file = $file;
        if (!empty($options)) {
            if ($options instanceof CsvOptions) {
                $this->delimiter = $options->delimiter;
                $this->enclosure = $options->enclosure;
                $this->escape = $options->escape;
                $this->encoding = $options->encoding;
            } else {
                if (isset($options['delimiter'])) {
                    $this->delimiter = $options['delimiter'];
                }
                if (isset($options['enclosure'])) {
                    $this->enclosure = $options['enclosure'];
                }
                if (isset($options['escape'])) {
                    $this->escape = $options['escape'];
                }
                if (isset($options['encoding'])) {
                    $this->encoding = $options['encoding'];
                }
            }
        }
    }

    /**
     * @param string $delimiter
     *
     * @return $this
     */
    public function setDelimiter(string $delimiter): CsvReader
    {
        $this->delimiter = $delimiter;

        return $this;
    }

    /**
     * @param string $enclosure
     *
     * @return $this
     */
    public function setEnclosure(string $enclosure): CsvReader
    {
        $this->enclosure = $enclosure;

        return $this;
    }

    /**
     * @param string $encoding
     *
     * @return $this
     */
    public function setEncoding(string $encoding): CsvReader
    {
        $this->encoding = $encoding;

        return $this;
    }

    /**
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     * @param int|null $rowLimit
     *
     * @return \Generator|null
     */
    public function nextRow($columnKeys = [], ?int $resultMode = null, ?int $rowLimit = 0): ?\Generator
    {
        $handle = fopen($this->file, 'rb');
        if (!$handle) {
            return null;
        }

        $rowNum = 0;
        $rowCnt = 0;
        $firstRowKeys = false;
        if (is_array($columnKeys)) {
            if (is_int($resultMode) && ($resultMode & Excel::KEYS_FIRST_ROW)) {
                $firstRowKeys = true;
            }
        } elseif ($columnKeys === true) {
            $firstRowKeys = true;
            $columnKeys = [];
        }

        while (($row = fgetcsv($handle, 0, $this->delimiter, $this->enclosure, $this->escape)) !== false) {
            $rowNum++;

            if ($rowNum === 1 && isset($row[0])) {
                if (strpos($row[0], "\xEF\xBB\xBF") === 0) {
                    $row[0] = substr($row[0], 3);
                }
            }

            if ($this->encoding && $this->encoding !== 'UTF-8') {
                foreach ($row as &$value) {
                    $value = mb_convert_encoding($value, 'UTF-8', $this->encoding);
                }
            }

            if ($rowNum === 1 && $firstRowKeys) {
                if (empty($columnKeys)) {
                    $columnKeys = $row;
                } else {
                    $columnKeys = array_merge($row, $columnKeys);
                }
                continue;
            }

            $rowCnt++;
            if ($rowLimit > 0 && $rowCnt > $rowLimit) {
                break;
            }

            $rowData = [];
            foreach ($row as $colIdx => $value) {
                if (isset($columnKeys[$colIdx])) {
                    $key = $columnKeys[$colIdx];
                } else {
                    $key = Helper::colLetter($colIdx + 1);
                }
                $rowData[$key] = $value;
            }

            yield $rowNum => $rowData;
        }
        fclose($handle);
    }

    /**
     * Reads cell values and passes them to a callback function
     *
     * @param callable $callback Callback function($row, $col, $value)
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     */
    public function readCallback(callable $callback, $columnKeys = [], ?int $resultMode = null)
    {
        foreach ($this->nextRow($columnKeys, $resultMode) as $row => $rowData) {
            foreach ($rowData as $col => $val) {
                $needBreak = $callback($row, $col, $val);
                if ($needBreak) {
                    return;
                }
            }
        }
    }

    /**
     * Read rows and return as 2D array
     *
     * @return array
     */
    public function readRows(): array
    {
        $data = [];
        foreach ($this->nextRow() as $rowNum => $row) {
            $data[$rowNum] = $row;
        }

        return $data;
    }

    /**
     * Read cells and return as 1D array [address => value]
     *
     * @return array
     */
    public function readCells(): array
    {
        $data = [];
        foreach ($this->nextRow() as $rowNum => $row) {
            foreach ($row as $colLetter => $value) {
                $data[$colLetter . $rowNum] = $value;
            }
        }

        return $data;
    }
}
