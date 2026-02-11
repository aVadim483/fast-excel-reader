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

        if (!empty($options)) {
            if ($options instanceof CsvOptions) {
                $this->delimiter = $options->delimiter;
                $this->enclosure = $options->enclosure;
                $this->escape = $options->escape;
                $this->encoding = strtoupper($options->encoding);
                $this->doubleQuotes = $options->doubleQuotes;
                $this->trimFields = $options->trimFields;
            }
            else {
                if (isset($options['delimiter'])) {
                    $this->delimiter = $options['delimiter'];
                    if ($this->delimiter === 'auto') {
                        $this->delimiter = null;
                    }
                }
                if (isset($options['quote'])) {
                    $this->enclosure = $options['quote'];
                }
                if (isset($options['escape'])) {
                    $this->escape = $options['escape'];
                }
                if (isset($options['encoding'])) {
                    $this->encoding = $options['encoding'];
                    if ($this->encoding === 'auto') {
                        $this->encoding = null;
                    }
                }
                if (isset($options['double_quotes'])) {
                    $this->doubleQuotes = $options['double_quotes'];
                }
                if (isset($options['trim_fields'])) {
                    $this->trimFields = $options['trim_fields'];
                }
                if (isset($options['mode'])) {
                    $this->strictMode = ($options['mode'] === CsvOptions::STRICT_MODE);
                }
            }
        }
        if ($options) {
            foreach ($options as $key => $value) {
                switch ($key) {
                    case 'delimiter':
                        $this->delimiter = $value;
                        break;
                    case 'enclosure':
                        $this->enclosure = $value;
                        break;
                        case 'escape':
                            $this->escape = $value;
                            break;
                    case 'encoding':
                        $this->encoding = strtoupper($value);
                        break;
                    case 'double_quotes':
                        $this->doubleQuotes = $value;
                        break;
                    case 'trim_fields':
                        $this->trimFields = $value;
                        break;
                    case 'mode':
                        $this->strictMode = ($value === CsvOptions::STRICT_MODE);
                        break;
                }
            }
        }

        if ($this->delimiter === null || $this->encoding === null) {
            $sample = CsvHelper::readSample($this->file);
        }
        else {
            $sample = CsvHelper::readSample($this->file, 4);
        }
        if ($sample === null) {
            throw new Exception("Cannot read file $file");
        }
        $this->bom = CsvHelper::detectEncodingByBomBytes($sample);

        if ($this->encoding === null) {
            $this->encoding = CsvHelper::detectEncoding($sample);
        }
        if ($this->delimiter === null) {
            if ($this->encoding !== 'UTF-8') {
                $sample = mb_convert_encoding($sample, 'UTF-8', $this->bom);
            }
            $candidates = [",", ";", "\t", "|", ":"];
            $res = CsvHelper::detectCsvDelimiter($sample, $candidates);
            if ($res['delimiter'] !== null && $res['confidence'] >= 0.6) {
                $this->delimiter = $res['delimiter'];
            }
            else {
                $this->delimiter = $res['guess'];
            }
        }
        if (stripos($this->encoding, 'UTF-16') === 0 || stripos($this->encoding, 'UTF-32') === 0) {
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
        } elseif ($columnKeys === true) {
            $firstRowKeys = true;
            $columnKeys = [];
        }

        while (($row = $this->getCsvLine()) !== null) {
            if (!$row) {
                break;
            }
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
                }
                else {
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
                $this->buffer = substr($this->buffer, $this->bufferPos);
                $this->bufferLen = strlen($this->buffer);
                $this->bufferPos = 0;
            }
        }
        if ($this->bufferPos >= $this->bufferLen) {
            return false;
        }

        $char = $this->buffer[$this->bufferPos++];
        $res = $char;
/*
        $byte = ord($char);

        // Определяем длину UTF-8 по первому байту
        if (($byte & 0xE0) === 0xC0) {
            $length = 2;
        } elseif (($byte & 0xF0) === 0xE0) {
            $length = 3;
        } elseif (($byte & 0xF8) === 0xF0) {
            $length = 4;
        } else {
            // Некорректный стартовый байт
            //return $b1; // или бросить исключение
        }

        $res = $char;
        /* ///

        if (($byte & 0b10000000) === 0) {
            $res = $char;
        }
        elseif (($byte & 0b11100000) === 0b11000000) {
            // utf 2 bytes
            $res = $char . $this->buffer[$this->bufferPos++];
        }
        elseif (($byte & 0b11110000) === 0b11100000) {
            // utf 3 bytes
            $res = $char . $this->buffer[$this->bufferPos++] . $this->buffer[$this->bufferPos++];
        }
        elseif (($byte & 0b11111000) === 0b11110000) {
            // utf 4 bytes
            $res = $char . $this->buffer[$this->bufferPos++] . $this->buffer[$this->bufferPos++] . $this->buffer[$this->bufferPos++];
        }
        else {
            // ($byte & 0b10000000) === 0
            $res = $char;
        }
        */
        $this->currentLine .= $res;

        return $res;
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
                    $this->error(self::ERR_UNEXPECTED_QUOTES, $char, 'Unexpected quotes in ' . $this->lineNo . ':' . $this->colNo);
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
                        $this->error(self::ERR_UNEXPECTED_CHAR, $char, "Unexpected character {$qch} after field in {$this->lineNo}:{$this->colNo}");
                    }
                }
                $field .= $char;
            }
        }

        if ($inQuotes) {
            // EOF inside quotes
            $this->error(self::ERR_UNEXPECTED_EOF, '', 'Unexpected EOF inside quoted CSV field');
        }

        return $field;
    }

    /**
     * Get line from CSV file as array of fields
     *
     * @return array|null
     */
    public function getCsvLine(): ?array
    {
        if (!$this->fp) {
            $this->open();
        }

        $row = [];
        $this->lineNo++;
        $this->colNo = 0;

        // Check for EOF
        if (($first = $this->getChar()) === false) {
            return null;
        }
        $this->ungetChar($first);

        while (($field = $this->getCsvField()) !== false) {
            //$field = mb_convert_encoding($field, 'UTF-8', $this->encoding);
            $row[] = (string)$field;

            // теперь читаем терминатор (delimiter / EOL / EOF)
            $sep = $this->getChar();
            if ($sep === false) {
                return $row;
            }

            if ($sep === $this->delimiter) {
                // следующее поле (в т.ч. может быть пустым)
                continue;
            }

            if ($sep === "\n") {
                return $row;
            }

            if ($sep === "\r") {
                // CRLF or just CR
                $n = $this->getChar();
                if ($n && $n !== "\n") {
                    $this->ungetChar($n);
                }
                return $row;
            }

            // неожиданный символ (грязный CSV)
            if ($this->strictMode) {
                $qch = ($sep === '`') ? '`' : '`' . $sep . '`';
                $this->error(self::ERR_UNEXPECTED_CHAR, $sep, "Unexpected character {$qch} in {$this->lineNo}:{$this->colNo}");
            }

            // lenient: считаем, что это продолжение значения (очень редкий случай), “приклеим” к последнему полю
            $row[count($row) - 1] .= $sep;
        } // while

        return $row ?: null;
    }

}
