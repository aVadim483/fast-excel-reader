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
    protected ?string $commentPrefix = null;
    protected int $startRow = 1;
    protected int $startCol = 1;
    protected bool $withHeader = false;
    protected array $lineErrors = [];

    /** @var callable|null  */
    protected $onError = null;

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
        $this->onError = [$this, 'defaultErrorHandler'];

        $this->setOptions(new CsvOptions());
        if (!empty($options)) {
            $this->setOptions($options);
        }

        if ($this->encoding) {
            $enc = CsvHelper::availableEncoding($this->encoding);
            if (!$enc) {
                throw new Exception("Encoding {$this->encoding} is not supported");
            }
            if ($enc !== $this->encoding) {
                $this->encoding = $enc;
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

    protected function setOptions($options)
    {
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
                    case 'comment_prefix':
                    case 'commentPrefix':
                        $this->commentPrefix = $value;
                        break;
                }
            }
        }
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

    public function close()
    {
        if ($this->fp) {
            fclose($this->fp);
            $this->fp = null;
        }
    }

    public function rewind()
    {
        $this->close();
        $this->open();
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

    /**
     * @param int $rowNum
     *
     * @return $this
     */
    public function fromRow(int $rowNum): CsvReader
    {
        $this->startRow = $rowNum;

        return $this;
    }

    /**
     * @param int $colNum
     *
     * @return $this
     */
    public function fromCol(int $colNum): CsvReader
    {
        $this->startCol = $colNum;

        return $this;
    }

    /**
     * Enables header mode
     *
     * Treats the first row of the CSV file as a header row and returns subsequent
     * rows as associative arrays keyed by column names
     *
     * @return $this
     */
    public function withHeader(): CsvReader
    {
        $this->withHeader = true;

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


    public function onError(?callable $handler): CsvReader
    {
        $this->onError = $handler;
        
        return $this;
    }

    /**
     * @param int $code
     * @param string $error
     * @param string $line
     * @param int $lineNo
     * @param int $colNo
     */
    public function defaultErrorHandler(int $code, string $error, string $line, int $lineNo, int $colNo)
    {
        throw new Exception($error, $code);
    }

    protected function callErrorHandler()
    {
        if ($this->lineErrors && $this->onError) {
            foreach ($this->lineErrors as $lineError) {
                call_user_func($this->onError, $lineError['code'], $lineError['error'], $this->currentLine, $lineError['row_no'], $lineError['col_no']);
            }
        }
    }

    /**
     * @param int $errCode
     * @param string $errText
     */
    protected function setError(int $errCode, string $errText)
    {
        $this->lineErrors[] = ['row_no' => $this->lineNo, 'col_no' => $this->colNo, 'code' => $errCode, 'error' => $errText];
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
        $firstRowKeys = $this->withHeader;
        if (is_array($columnKeys)) {
            if (is_int($resultMode) && ($resultMode & Excel::KEYS_FIRST_ROW)) {
                $firstRowKeys = true;
            }
        }
        elseif ($columnKeys === true) {
            $firstRowKeys = true;
            $columnKeys = [];
        }
        $resColumnKeys = [];

        while (($row = $this->getCsvLine()) !== false) {
            $rowNum++;

            if (count($row) === 0) {
                if ($this->skipEmptyLines) {
                    continue;
                }
                $row = [];
            }

            if ($this->commentPrefix && strpos($this->currentLine, $this->commentPrefix) === 0) {
                continue;
            }

            if ($rowNum < $this->startRow) {
                continue;
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

            if ($this->startCol) {
                $row = array_slice($row, $this->startCol - 1);
            }

            // define column keys
            if (!$resColumnKeys) {
                // original indexes of fields
                $colKeys = array_keys($row);
                if (is_int($resultMode) && ($resultMode & CsvOptions::KEYS_COL_EXCEL)) {
                    foreach ($colKeys as $colIdx) {
                        $resColumnKeys[] = Helper::colLetter($colIdx + 1);
                    }
                }
                elseif (is_int($resultMode) && ($resultMode & CsvOptions::KEYS_COL_ONE_BASED)) {
                    foreach ($colKeys as $colIdx) {
                        $resColumnKeys[] = $colIdx + 1;
                    }
                }
                else {
                    $resColumnKeys = $colKeys;
                }
                if ($columnKeys) {
                    foreach ($columnKeys as $colIdx => $colName) {
                        if ($colName) {
                            $resColumnKeys[$colIdx] = $colName;
                        }
                    }
                }
            }

            if (count($resColumnKeys) < count($row)) {
                $min = count($resColumnKeys);
                $max = count($row);
                for ($idx = $min; $idx < $max; $idx++) {
                    if (is_int($resultMode) && ($resultMode & CsvOptions::KEYS_COL_EXCEL)) {
                        $resColumnKeys[] = Helper::colLetter($idx + 1);
                    }
                    elseif (is_int($resultMode) && ($resultMode & CsvOptions::KEYS_COL_ONE_BASED)) {
                        $resColumnKeys[] = $idx + 1;
                    }
                    else {
                        $resColumnKeys[] = $idx;
                    }
                }
            }
            $rowData = array_combine($resColumnKeys, array_values($row));
            if (is_int($resultMode) && ($resultMode & CsvOptions::KEYS_ROW_ONE_BASED)) {
                $rowKey = $rowNum + 1;
            }
            elseif (is_int($resultMode) && ($resultMode & CsvOptions::KEYS_ROW_ZERO_BASED)) {
                $rowKey = $rowNum - 1;
            }
            else {
                $rowKey = $rowNum;
            }

            yield $rowKey => $rowData;
        }
        $this->close();
    }

    /**
     * Read rows and return as 2D array
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     *
     * @return array
     */
    public function readRows($columnKeys = [], ?int $resultMode = null): array
    {
        $data = [];
        foreach ($this->nextRow($columnKeys, $resultMode) as $rowNum => $row) {
            $data[$rowNum] = $row;
        }

        return $data;
    }

    private function newLine()
    {
        $this->lineNo++;
        $this->colNo = 0;
        $this->currentLine = '';
        $this->lineErrors = [];
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

        if (($this->bufferPos >= $this->bufferLen)) {
            if (!feof($this->fp)) {
                $this->buffer = fread($this->fp, $this->bufferSize);
                $this->bufferLen = strlen($this->buffer);
                $this->bufferPos = 0;
            }
            else {
                return false;
            }
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
        elseif ($ch === $this->escape) {
            $next = $this->getChar();
            if ($next === false) { // EOF
                return $ch;
            }
            elseif ($next === "\n" || $next === "\r") { // EOL
                $this->ungetChar($next);
                return $ch;
            }
            else {
                $field = $next;
            }
        }
        else {
            $field = $ch;
        }

        while (($char = $this->getChar()) !== false) {
            if ($ignore && $char !== $this->delimiter && $char !== "\n" && $char !== "\r") {
                continue;
            }

            if ($char === $this->enclosure) {
                if (!$quotedField && $this->strictMode) {
                    $this->setError(self::ERR_UNEXPECTED_QUOTES, 'Unexpected quotes in ' . $this->lineNo . ':' . $this->colNo);
                    // If there is no break, then ignore characters until the end of the field
                    $ignore = true;
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
                    elseif ($this->strictMode) {
                        $qch = ($char === '`') ? '`' : '`' . $char . '`';
                        $this->setError(self::ERR_UNEXPECTED_CHAR, "Unexpected character {$qch} after field in {$this->lineNo}:{$this->colNo}");
                        // If there is no break, then ignore characters until the end of the field
                        $ignore = true;
                    }
                }
                if (!$ignore) {
                    $field .= $char;
                }
            }
        }

        if ($inQuotes) {
            // EOF inside quotes
            $this->setError(self::ERR_UNEXPECTED_EOF, 'Unexpected EOF inside quoted CSV field');
        }

        return $field;
    }

    /**
     * Get line from CSV file as array of fields (null - empty field, false - EOF)
     *
     * @return array|null|false
     */
    public function getCsvLine()
    {
        if (!$this->fp) {
            $this->open();
        }

        $row = [];
        $this->newLine();

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
                // the next field
                continue;
            }

            if ($sep === false || $sep === "\n") {
                break; // EOL
            }
            elseif ($sep === "\r") {
                // CRLF or just CR
                $n = $this->getChar();
                if ($n && $n !== "\n") {
                    $this->ungetChar($n);
                }
                break; // EOL
            }

            // unexpected character (dirty CSV)
            if ($this->strictMode) {
                $qch = ($sep === '`') ? '`' : '`' . $sep . '`';
                $this->setError(self::ERR_UNEXPECTED_CHAR, "Unexpected character {$qch} in {$this->lineNo}:{$this->colNo}");
            }

            // tolerant mode: the field value continues
            $row[count($row) - 1] .= $sep;
        } // while

        $this->callErrorHandler();

        if (count($row) === 1 && $csvField === null) {
            return [];
        }

        return $row ?: false;
    }

}
