<?php

namespace avadim\FastExcelReader;

use avadim\FastExcelHelper\Helper;

class CsvReader
{
    protected string $file;
    protected $fp = null;
    protected ?string $delimiter = null;
    protected string $quote = '"';
    protected string $escape = '\\';
    protected ?string $encoding = null;
    protected ?string $pushback = null;
    protected bool $strictMode = true;
    protected int $line = 0;
    protected int $col = 0;

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
                $this->quote = $options->quote;
                $this->escape = $options->escape;
                $this->encoding = $options->encoding;
            }
            else {
                if (isset($options['delimiter'])) {
                    $this->delimiter = $options['delimiter'];
                }
                if (isset($options['quote'])) {
                    $this->quote = $options['quote'];
                }
                if (isset($options['escape'])) {
                    $this->escape = $options['escape'];
                }
                if (isset($options['encoding'])) {
                    $this->encoding = $options['encoding'];
                }
                if (isset($options['mode'])) {
                    $this->strictMode = ($options['mode'] === CsvOptions::STRICT_MODE);
                }
            }
        }
        if ($this->delimiter === null || $this->delimiter === 'auto') {
            $candidates = [",", ";", "\t", "|", ":"];
            $res = $this->detectCsvDelimiter($candidates);
            if ($res['delimiter'] !== null && $res['confidence'] >= 0.6) {
                $this->delimiter = $res['delimiter'];
            }
            else {
                $this->delimiter = $res['guess'];
            }
        }
    }

    public function __destruct()
    {
        $this->close();
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
     * @param string $quote
     *
     * @return $this
     */
    public function setQuote(string $quote): CsvReader
    {
        $this->quote = $quote;

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

    public function getOptions(): CsvOptions
    {
        return new CsvOptions([
            'delimiter' => $this->delimiter,
            'quote' => $this->quote,
            'escape' => $this->escape,
            'encoding' => $this->encoding,
            'mode' => $this->strictMode ? CsvOptions::STRICT_MODE : CsvOptions::TOLERANT_MODE,
        ]);
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
        $this->fp = fopen($this->file, 'rb');
        if (!$this->fp) {
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

    /**
     * @return string|null
     */
    private function getChar(): ?string
    {
        if ($this->pushback !== null) {
            $ch = $this->pushback;
            $this->pushback = null;

            return $ch;
        }

        $ch = fgetc($this->fp);
        $this->col++;

        return $ch === false ? null : $ch;
    }

    /**
     * @param string $ch
     *
     * @return void
     */
    private function ungetChar(string $ch): void
    {
        $this->pushback = $ch;
        $this->col--;
    }

    /**
     * Get next field from CSV file
     *
     * @return array|null
     */
    public function getCsvField(): ?string
    {
        $field = null;
        $ignore = false;
        $inQuotes = false;
        $quotedField = false;
        $endOfField = false;

        while (($ch = $this->getChar()) !== null) {
            if ($inQuotes && $ch !== $this->quote && $ch !== $this->escape) {
                $field .= $ch;
                continue;
            }
            if ($endOfField || $ch === $this->delimiter || $ch === "\n" || $ch === "\r") {
                // ignore white spaces after end of field
                if ($ch === ' ' || $ch === "\t") {
                    // ignore white spaces after end of field
                    continue;
                }
                else {
                    if ($quotedField && $inQuotes) {
                        throw new Exception("Unclosed quoted field in line {$this->line} col {$this->col}");
                    }
                    if ($ch === "\n" || $ch === "\r") {
                        $this->ungetChar($ch);
                    }

                    return $field;
                }
            }
            if ($this->escape && $ch === $this->escape) {
                $next = $this->getChar();
                if ($next === $this->escape || $next === $this->quote || $next === $this->delimiter) {
                    $field .= $next;
                    continue;
                }
                elseif ($next !== "\n" && $next !== "\r") {
                    $field .= $next;
                    $ch = $next;
                }
            }
            if ($ch === $this->quote) {
                if (!$field) {
                    $quotedField = true;
                    $inQuotes = true;
                }
                elseif (!$quotedField) {
                    // non-RFC 4180
                    if ($this->strictMode) {
                        $qch = ($this->quote === '`') ? '`' : '`' . $ch . '`';
                        throw new Exception("Invalid CSV quote: $qch in line {$this->line} col {$this->col}");
                    }
                }
                elseif (($next = $this->getChar()) !== null) {
                    if ($next === $this->quote) {
                        // double quote ""
                        $field .= $this->quote;
                    }
                    elseif ($inQuotes) {
                        $inQuotes = false;
                        $endOfField = true;
                        $this->ungetChar($next);
                    }
                }
            }
            else {
                $field .= $ch;
            }
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
            $this->fp = fopen($this->file, 'rb');
        }

        $row = [];
        $this->line++;
        $this->col = 0;

        while (($field = $this->getCsvField()) !== null) {
            $row[] = $field;
            $ch = $this->getChar();
            if ($ch === "\n" || $ch === "\r") {
                if ($ch === "\r") {
                    $next = $this->getChar();
                    if ($next !== "\n" && $next !== null) {
                        $this->ungetChar($next);
                    }
                }

                return $row;
            }
            elseif ($ch !== null) {
                $this->ungetChar($ch);
            }
        }

        return $row ?: null;
    }

    /**
     * Detect delimiter in CSV/TSV by sampling the beginning of a stream or file.
     *
     * @param array|null $candidates Delimiters to try.
     * @param int $sampleBytes How many bytes to read from the beginning.
     * @param int $maxLines How many logical lines to analyze.
     *
     * @return array{
     *   delimiter: ?string,
     *   confidence: float,
     *   details: array<string, array{score: float, mode: int, mode_ratio: float, stdev: float, lines: int}>
     * }
     */
    public function detectCsvDelimiter(?array $candidates = null, int $sampleBytes = 131072, int $maxLines = 80): array
    {
        $candidates ??= [",", ";", "\t", "|", ":"];

        [$text, $readOk] = $this->readSample($sampleBytes);
        if (!$readOk || $text === '') {
            return [
                'delimiter' => null,
                'confidence' => 0.0,
                'details' => [],
            ];
        }

        // Strip UTF-8 BOM if present
        if (strncmp($text, "\xEF\xBB\xBF", 3) === 0) {
            $text = substr($text, 3);
        }

        // Get first logical lines (handles newlines inside quotes)
        $lines = $this->extractLogicalLines($text, $maxLines);
        if (count($lines) < 3) {
            // Not enough data; best effort: pick delimiter with max occurrences in first line
            $first = $lines[0] ?? $text;
            $best = null;
            $bestCnt = 0;
            foreach ($candidates as $d) {
                $cnt = $this->countDelimiterOutsideQuotes($first, $d);
                if ($cnt > $bestCnt) {
                    $bestCnt = $cnt;
                    $best = $d;
                }
            }
            return [
                'delimiter' => $bestCnt > 0 ? $best : null,
                'confidence' => $bestCnt > 0 ? 0.4 : 0.0,
                'details' => [],
                'guess' => $this->guessByLocale($lines),
            ];
        }

        // Heuristic: detect many European decimals "12,34" (comma decimal separator)
        $commaDecimalHits = $this->countCommaDecimals($lines);

        $details = [];
        $bestDelim = null;
        $bestScore = -INF;

        foreach ($candidates as $d) {
            $fields = [];
            $usedLines = 0;

            foreach ($lines as $line) {
                $trim = trim($line);
                if ($trim === '') {
                    continue;
                }
                $cnt = $this->countDelimiterOutsideQuotes($line, $d);
                $fields[] = $cnt + 1;
                $usedLines++;
                if ($usedLines >= $maxLines) {
                    break;
                }
            }

            if ($usedLines < 3) {
                $details[$d] = ['score' => -INF, 'mode' => 0, 'mode_ratio' => 0.0, 'stdev' => 0.0, 'lines' => $usedLines];
                continue;
            }

            $mode = $this->modeInt($fields);
            $modeRatio = ($mode > 0) ? (count(array_filter($fields, fn($v) => $v === $mode)) / count($fields)) : 0.0;
            $stdev = $this->stddev($fields);

            // Base score: stability first
            if ($mode < 2) {
                $score = -INF; // delimiter must yield at least 2 columns consistently
            } else {
                $score = ($modeRatio * 100.0) + min($mode, 20) - ($stdev * 5.0);

                // Penalize too-rare delimiter usage
                $totalDelims = array_sum(array_map(fn($v) => $v - 1, $fields));
                if ($totalDelims < 5) {
                    $score -= 20.0;
                }

                // Tie-break bias: if comma-decimals are common, prefer ';' over ','
                if ($commaDecimalHits > 0) {
                    if ($d === ';') $score += min(10, $commaDecimalHits);
                    if ($d === ',') $score -= min(10, $commaDecimalHits);
                }

                // TSV bias: if there are many tabs across lines, give \t a bit of boost
                if ($d === "\t") {
                    $tabHits = $this->countTabs($lines);
                    if ($tabHits > 0) $score += min(10, $tabHits);
                }
            }

            $details[$d] = [
                'score' => $score,
                'mode' => (int)$mode,
                'mode_ratio' => (float)$modeRatio,
                'stdev' => (float)$stdev,
                'lines' => (int)$usedLines,
            ];

            if ($score > $bestScore) {
                $bestScore = $score;
                $bestDelim = $d;
            }
        }

        // Compute a simple confidence: based on score gap vs runner-up
        $sorted = $details;
        uasort($sorted, fn($a, $b) => $b['score'] <=> $a['score']);
        $top = array_values($sorted)[0] ?? null;
        $second = array_values($sorted)[1] ?? null;

        $confidence = 0.0;
        if ($top && is_finite($top['score']) && $top['score'] > 0) {
            $gap = ($second && is_finite($second['score'])) ? ($top['score'] - $second['score']) : 999;
            // Map gap roughly to 0..1
            $confidence = max(0.0, min(1.0, 0.5 + ($gap / 50.0)));
            // Also require decent mode_ratio
            $confidence *= max(0.0, min(1.0, ($top['mode_ratio'] - 0.5) / 0.5));
        }

        return [
            'delimiter' => (is_finite($bestScore) ? $bestDelim : null),
            'confidence' => (float)$confidence,
            'details' => $details,
            'guess' => $this->guessByLocale($lines),
        ];
    }

    /**
     * Read a sample from resource or file.
     * Returns [$text, $ok].
     *
     * @param int $sampleBytes
     *
     * @return array
     */
    protected function readSample(int $sampleBytes): array
    {
        if ($this->file && is_file($this->file)) {
            $fp = @fopen($this->file, 'rb');
            if (!$fp) return ['', false];
            $data = fread($fp, $sampleBytes);
            fclose($fp);

            return [$data !== false ? $data : '', true];
        }

        return ['', false];
    }

    /**
     * Extract up to $max logical lines from text, handling newlines inside quotes.
     *
     * @param string $text
     * @param int $max
     *
     * @return array
     */
    protected function extractLogicalLines(string $text, int $max): array
    {
        $lines = [];
        $buf = '';
        $inQuotes = false;
        $len = strlen($text);

        for ($i = 0; $i < $len; $i++) {
            $ch = $text[$i];

            if ($ch === '"') {
                // If in quotes and next is also a quote -> escaped quote
                if ($inQuotes && $i + 1 < $len && $text[$i + 1] === '"') {
                    $buf .= '""';
                    $i++;
                    continue;
                }
                $inQuotes = !$inQuotes;
                $buf .= $ch;
                continue;
            }

            // Newline handling (CRLF, LF, CR) only if not in quotes
            if (!$inQuotes && ($ch === "\n" || $ch === "\r")) {
                if ($ch === "\r" && $i + 1 < $len && $text[$i + 1] === "\n") {
                    $i++; // consume \n of CRLF
                }
                $lines[] = $buf;
                $buf = '';
                if (count($lines) >= $max) break;
                continue;
            }

            $buf .= $ch;
        }

        if ($buf !== '' && count($lines) < $max) {
            $lines[] = $buf;
        }

        return $lines;
    }

    /**
     * Count delimiter occurrences outside quotes in a single logical line.
     *
     * @param string $line
     * @param string $delim
     *
     * @return int
     */
    protected function countDelimiterOutsideQuotes(string $line, string $delim): int
    {
        $inQuotes = false;
        $cnt = 0;
        $len = strlen($line);

        for ($i = 0; $i < $len; $i++) {
            $ch = $line[$i];

            if ($ch === '"') {
                if ($inQuotes && $i + 1 < $len && $line[$i + 1] === '"') {
                    $i++; // escaped quote
                    continue;
                }
                $inQuotes = !$inQuotes;
                continue;
            }

            if (!$inQuotes) {
                // delim is single-byte in our candidates; if you add multi-byte, adjust
                if ($ch === $delim) {
                    $cnt++;
                }
            }
        }

        return $cnt;
    }

    /**
     * Find mode (most frequent) of an int array.
     *
     * @param array $values
     *
     * @return int
     */
    protected function modeInt(array $values): int
    {
        $freq = [];
        foreach ($values as $v) {
            $v = (int)$v;
            $freq[$v] = ($freq[$v] ?? 0) + 1;
        }
        arsort($freq);

        return (int)array_key_first($freq);
    }

    /**
     * Standard deviation of numeric array.
     *
     * @param array $values
     *
     * @return float
     */
    protected function stddev(array $values): float
    {
        $n = count($values);
        if ($n <= 1) return 0.0;

        $mean = array_sum($values) / $n;
        $var = 0.0;
        foreach ($values as $v) {
            $d = $v - $mean;
            $var += $d * $d;
        }
        $var /= ($n - 1); // sample stdev

        return sqrt($var);
    }

    /**
     * Count patterns like 12,34 (comma decimals) across sampled lines.
     *
     * @param array $lines
     *
     * @return int
     */
    protected function countCommaDecimals(array $lines): int
    {
        $hits = 0;
        // Keep it lightweight: simple scan, not heavy regex loops over huge strings.
        foreach ($lines as $line) {
            $len = strlen($line);
            for ($i = 1; $i < $len - 1; $i++) {
                if ($line[$i] === ',' && ctype_digit($line[$i - 1]) && ctype_digit($line[$i + 1])) {
                    $hits++;
                    if ($hits >= 20) return $hits; // cap
                }
            }
        }
        return $hits;
    }

    /**
     * @param array $lines
     *
     * @return int
     */
    protected function countTabs(array $lines): int
    {
        $hits = 0;
        foreach ($lines as $line) {
            $hits += substr_count($line, "\t");
            if ($hits >= 20) return $hits; // cap
        }
        return $hits;
    }


    protected function guessByLocale(array $lines): ?string
    {
        $decimalPoint = setlocale(LC_NUMERIC, 0);
        if (!$decimalPoint || $decimalPoint === 'C') {
            $decimalPoint = '';
            $lc = @localeconv();
            if (is_array($lc) && !empty($lc['decimal_point'])) {
                $decimalPoint = $lc['decimal_point'];
            }
        }
        if ($decimalPoint === ',') {
            return ';';
        }

        return ',';
    }

}
