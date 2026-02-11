<?php

namespace avadim\FastExcelReader\Csv;

use avadim\FastExcelHelper\Helper;
use avadim\FastExcelReader\Excel;
use avadim\FastExcelReader\Exception;

class CsvReader
{
    const ERR_UNCLOSED_FIELD = 1001;
    const ERR_UNEXPECTED_QUOTES = 1002;
    const ERR_UNEXPECTED_CHAR = 1003;
    const ERR_UNEXPECTED_EOF = 1004;

    protected string $file;
    protected $fp = null;
    protected ?string $delimiter = null;
    protected string $enclosure = '"';
    protected string $escape = '\\';
    protected ?string $encoding = null;
    protected bool $doubleQuotes = true;
    protected bool $trimFields = true;
    protected bool $skipEmptyLines = false;
    protected bool $strictMode = true;
    protected $errorHandler = null;

    protected string $buffer = '';
    protected int $bufferPos = 0;
    protected int $bufferLen = 0;
    protected int $bufferSize = 4096;
    protected ?string $pushback = null;
    protected int $lineNo = 0;
    protected int $colNo = 0;
    protected string $currentLine = '';
    protected ?string $bom;
    protected ?string $streamFilter = null;

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
        $this->errorHandler = [$this, 'errorHandler'];

        if (empty($options)) {
            $options = new CsvOptions();
        }
        if (!empty($options)) {
            if ($options instanceof CsvOptions) {
                $options = $options->toArray();
            }
            if ($options) {
                foreach ($options as $key => $value) {
                    switch ($key) {
                        case 'delimiter':
                            $this->delimiter = ($value === 'auto' ? null : $value);
                            break;
                        case 'enclosure':
                            $this->enclosure = $value;
                            break;
                        case 'escape':
                            $this->escape = $value;
                            break;
                        case 'encoding':
                            $this->encoding = ($value ? strtoupper($value) : null);
                            break;
                        case 'double_quotes':
                        case 'doubleQuotes':
                            $this->doubleQuotes = $value;
                            break;
                        case 'trim_fields':
                        case 'trimFields':
                            $this->trimFields = $value;
                            break;
                        case 'skip_empty_lines':
                        case 'skipEmptyLines':
                            $this->skipEmptyLines = (bool)$value;
                            break;
                        case 'mode':
                            $this->strictMode = ($value === CsvOptions::STRICT_MODE);
                            break;
                        case 'stream_filter':
                        case 'streamFilter':
                            $this->streamFilter = $value;
                            break;
                    }
                }
            }
        }

        if ($this->delimiter === null || $this->encoding === null) {
            $sample = CsvHelper::readSample($this->file);
        } else {
            $sample = CsvHelper::readSample($this->file, 4);
        }
        if ($sample === null) {
            throw new Exception("Cannot read file $file");
        }
        $this->bom = CsvHelper::detectEncodingByBomBytes($sample);
        if ($this->bom && !$this->encoding) {
            $this->encoding = $this->bom;
        }

        if ($this->encoding === null) {
            $this->encoding = CsvHelper::detectEncoding($sample);
        }
        if ($this->delimiter === null) {
            if ($this->encoding !== 'UTF-8' && $this->encoding !== null) {
                $sample = mb_convert_encoding($sample, 'UTF-8', $this->encoding);
            }
            $candidates = [",", ";", "\t", "|", ":"];
            $res = CsvHelper::detectCsvDelimiter($sample, $candidates);
            if ($res['delimiter'] !== null && $res['confidence'] >= 0.6) {
                $this->delimiter = $res['delimiter'];
            } else {
                $this->delimiter = $res['guess'];
            }
        }
        if ($this->streamFilter === null && (stripos($this->encoding, 'UTF-16') === 0 || stripos($this->encoding, 'UTF-32') === 0)) {
            $this->streamFilter = "convert.iconv.$this->encoding/UTF-8";
        }
    }

    public function __destruct()
    {
        $this->close();
    }

    /**
     * @return bool
     */
    protected function open(): bool
    {
        $this->fp = @fopen($this->file, 'rb');
        if (!$this->fp) {
            throw new Exception("Cannot open file: {$this->file}");
        }

        if ($this->bom && isset(CsvHelper::$bomMap[$this->bom])) {
            fread($this->fp, strlen(CsvHelper::$bomMap[$this->bom]));
        }
        // Attach iconv conversion filter if needed
        if ($this->streamFilter) {
            $ok = stream_filter_append($this->fp, $this->streamFilter, STREAM_FILTER_READ);
            if ($ok === false) {
                fclose($this->fp);
                throw new Exception("Cannot attach stream filter: {$this->streamFilter}");
            }
        }

        $this->lineNo = 0;
        $this->colNo = 0;

        return true;
    }

    protected function close()
    {
        if ($this->fp) {
            fclose($this->fp);
            $this->fp = null;
        }
    }

    /**
     * @param int $size
     *
     * @return $this
     */
    public function setBufferSize(int $size): CsvReader
    {
        $this->bufferSize = $size;

        return $this;
    }

    public function getOptions(): CsvOptions
    {
        return new CsvOptions([
            'delimiter' => $this->delimiter,
            'quote' => $this->enclosure,
            'escape' => $this->escape,
            'encoding' => $this->encoding,
            'mode' => $this->strictMode ? CsvOptions::STRICT_MODE : CsvOptions::TOLERANT_MODE,
            'stream_filter' => $this->streamFilter,
        ]);
    }


    public function setErrorHandler(callable $handler): CsvReader
    {
        $this->errorHandler = $handler;
        
        return $this;
    }

    /**
     * @param int $code
     * @param string $error
     * @param string $line
     * @param int $lineNo
     * @param int $colNo
     */
    public function errorHandler(int $code, string $error, string $line, int $lineNo, int $colNo)
    {
        throw new Exception($error, $code);
    }

    /**
     * @param int $errCode
     * @param string $errText
     */
    protected function error(int $errCode, string $errText)
    {
        if ($this->errorHandler) {
            call_user_func($this->errorHandler, $errCode, $errText, $this->currentLine, $this->lineNo, $this->colNo);
        }
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
        if (!$this->fp && !$this->open()) {
            return null;
        }

        $rowNum = 0;
        $rowCnt = 0;
        $firstRowKeys = false;
        if (is_array($columnKeys)) {
            if (is_int($resultMode) && ($resultMode & Excel::KEYS_FIRST_ROW)) {
                $firstRowKeys = true;
            }
        }
        elseif ($columnKeys === true) {
            $firstRowKeys = true;
            $columnKeys = [];
        }

        while (($row = $this->getCsvLine()) !== false) {
            $rowNum++;

            if ($row === null) {
                if ($this->skipEmptyLines) {
                    continue;
                }
                $row = [];
            }

            if ($rowNum === 1 && isset($row[0])) {
                if (strpos($row[0], "\xEF\xBB\xBF") === 0) {
                    $row[0] = substr($row[0], 3);
                }
            }

            if ($this->encoding && $this->encoding !== 'UTF-8' && $this->encoding !== $this->bom) {
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
        $this->close();
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

    private function newLine()
    {
        $this->lineNo++;
        $this->colNo = 0;
        $this->currentLine = '';
    }

    /**
     * @return string|false False -> EOF
     */
    private function getChar()
    {
        $this->colNo++;

        if ($this->pushback !== null) {
            $ch = $this->pushback;
            $this->pushback = null;

            return $ch;
        }

        if (($this->bufferPos >= ($this->bufferLen - 4)) && !feof($this->fp)) {
            $next = fread($this->fp, $this->bufferSize);
            if ($next !== false) {
                $this->buffer .= $next;
                if ($this->bufferPos > 0) {
                    $this->buffer = substr($this->buffer, $this->bufferPos);
                }
                $this->bufferLen = strlen($this->buffer);
                $this->bufferPos = 0;
            }
        }
        if ($this->bufferPos >= $this->bufferLen) {
            return false;
        }

        $char = $this->buffer[$this->bufferPos++];
        $this->currentLine .= $char;

        return $char;
    }

    /**
     * @param string $ch
     *
     * @return void
     */
    private function ungetChar(string $ch): void
    {
        $this->pushback = $ch;
        $this->colNo--;
    }

    /**
     * Get next field from CSV file
     *
     * @return string|null|false
     */
    public function getCsvField(): ?string
    {
        $field = null;
        $ignore = false;
        $inQuotes = false;
        $quotedField = false;
        $endOfField = false;

        $ch = $this->getChar();
        if ($ch === false) {

            return false; // EOF
        }

        // Empty field: if delimiter or EOL is encountered immediately — return '' and terminator is returned to the stream
        if ($ch === $this->delimiter || $ch === "\n" || $ch === "\r") {
            $this->ungetChar($ch);

            return null;
        }

        if ($ch === $this->enclosure) {
            // Quoted field
            $inQuotes = true;
            $quotedField = true;
        }
        else {
            $field = $ch;
        }

        while (($char = $this->getChar()) !== false) {
            if ($char === $this->enclosure) {
                if (!$quotedField && $this->strictMode) {
                    $this->error(self::ERR_UNEXPECTED_QUOTES, 'Unexpected quotes in ' . $this->lineNo . ':' . $this->colNo);
                }
                if (!$quotedField && !$this->strictMode) {
                    $field .= $char;
                    continue;
                }
                $next = $this->getChar();
                if ($next === $this->enclosure && ($inQuotes || !$this->strictMode)) {
                    // double quotes
                    $field .= $this->enclosure;
                    continue;
                }
                elseif ($next !== $this->enclosure && $inQuotes) {
                    $inQuotes = false;
                    $endOfField = true;
                    $this->ungetChar($next);
                    continue;
                }
            }
            elseif ($char === $this->escape) {
                $next = $this->getChar();
                if ($next === false) { // EOF
                    return $field;
                }
                elseif ($next === "\n" || $next === "\r") { // EOL
                    $this->ungetChar($char);
                    return $field;
                }
                else {
                    $field .= $next;
                }
            }
            elseif ($char === $this->delimiter || $char === "\n" || $char === "\r") {
                if ($inQuotes) {
                    // quoted string
                    $field .= $char;
                }
                else {
                    // end of field
                    if (!$quotedField && $this->trimFields) {
                        $field = trim($field);
                    }
                    $this->ungetChar($char);
                    return $field;
                }
            }
            else {
                if ($endOfField) {
                    if ($this->trimFields && ($char === ' ' || $char === "\t")) {
                        continue;
                    }
                    else {
                        $qch = ($char === '`') ? '`' : '`' . $char . '`';
                        $this->error(self::ERR_UNEXPECTED_CHAR, "Unexpected character {$qch} after field in {$this->lineNo}:{$this->colNo}");
                    }
                }
                $field .= $char;
            }
        }

        if ($inQuotes) {
            // EOF inside quotes
            $this->error(self::ERR_UNEXPECTED_EOF, 'Unexpected EOF inside quoted CSV field');
        }

        return $field;
    }

    /**
     * Get line from CSV file as array of fields
     *
     * @return array|null|false
     */
    public function getCsvLine()
    {
        if (!$this->fp) {
            $this->open();
        }

        $row = [];
        $this->lineNo++;
        $this->colNo = 0;

        // Check for EOF
        if (($first = $this->getChar()) === false) {
            return false;
        }
        $this->ungetChar($first);

        while (($csvField = $this->getCsvField()) !== false) {
            if ($csvField && $this->encoding && substr($this->encoding, 0, 3) !== 'UTF') {
                $field = mb_convert_encoding($csvField, 'UTF-8', $this->encoding);
            }
            else {
                $field = (string)$csvField;
            }
            $row[] = $field;

            // read terminator of field (delimiter / EOL / EOF)
            $sep = $this->getChar();
            if ($sep === $this->delimiter) {
                // the next field (may be empty)
                continue;
            }

            $eol = false;
            if ($sep === false) {
                $eol = true;
            }
            elseif ($sep === "\n") {
                $eol = true;
            }
            elseif ($sep === "\r") {
                // CRLF or just CR
                $n = $this->getChar();
                if ($n && $n !== "\n") {
                    $this->ungetChar($n);
                }
                $eol = true;
            }

            if ($eol) {
                return (count($row) === 1 && $csvField === null) ? null : $row;
            }

            // unexpected character (dirty CSV)
            if ($this->strictMode) {
                $qch = ($sep === '`') ? '`' : '`' . $sep . '`';
                $this->error(self::ERR_UNEXPECTED_CHAR, "Unexpected character {$qch} in {$this->lineNo}:{$this->colNo}");
            }

            // tolerant mode: the field value continues
            $row[count($row) - 1] .= $sep;
        } // while

        return $row ?: false;
    }

}
