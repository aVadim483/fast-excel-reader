<?php

namespace avadim\FastExcelReader\Csv;

use RuntimeException;

class CsvHelper
{
    public static array $bomMap = [
        'UTF-8' => "\xEF\xBB\xBF",
        'UTF-32LE' => "\xFF\xFE\x00\x00",
        'UTF-32BE' => "\x00\x00\xFE\xFF",
        'UTF-16LE' => "\xFF\xFE",
        'UTF-16BE' => "\xFE\xFF",
    ];

    /**
     * Detect encoding by BOM from a binary string.
     *
     * @return string|null One of: 'UTF-8', 'UTF-16LE', 'UTF-16BE', 'UTF-32LE', 'UTF-32BE' or null
     */
    public static function detectEncodingByBomBytes(string $bytes): ?string
    {
        foreach (self::$bomMap as $enc => $bom) {
            if (strncmp($bytes, $bom, strlen($bom)) === 0) {
                return $enc;
            }
        }

        return null;
    }

    /**
     * Return BOM length for known Unicode BOMs.
     */
    public static function bomLength(string $encoding): int
    {
        switch (strtoupper($encoding)) {
            case 'UTF-8':
                return 3;
            case 'UTF-16LE':
            case 'UTF-16BE':
                return 2;
            case 'UTF-32LE':
            case 'UTF-32BE':
                return 4;
            default:
                return 0;
        }
    }

    /**
     * Best-effort: check if bytes look like valid UTF-8 (strict).
     */
    public static function looksLikeUtf8(string $bytes): bool
    {
        // If ext/mbstring exists, it's a very good validator.
        if (function_exists('mb_check_encoding')) {
            return mb_check_encoding($bytes, 'UTF-8');
        }

        $len = strlen($bytes);
        for ($i = 0; $i < $len; $i++) {
            $c = ord($bytes[$i]);
            if ($c <= 0x7F) {
                continue;
            }

            // Determine sequence length
            if (($c & 0xE0) === 0xC0) {
                $n = 1;
                $min = 0x80;
                $code = $c & 0x1F;
            } elseif (($c & 0xF0) === 0xE0) {
                $n = 2;
                $min = 0x800;
                $code = $c & 0x0F;
            } elseif (($c & 0xF8) === 0xF0) {
                $n = 3;
                $min = 0x10000;
                $code = $c & 0x07;
            } else {
                return false;
            }

            if ($i + $n >= $len) {
                return false;
            }

            for ($j = 0; $j < $n; $j++) {
                $i++;
                $cc = ord($bytes[$i]);
                if (($cc & 0xC0) !== 0x80) {
                    return false;
                }
                $code = ($code << 6) | ($cc & 0x3F);
            }

            // Reject overlongs and invalid ranges
            if ($code < $min) {
                return false;
            }
            if ($code >= 0xD800 && $code <= 0xDFFF) {
                return false; // surrogates
            }
            if ($code > 0x10FFFF) {
                return false;
            }
        }

        return true;
    }

    /**
     * Very simple heuristic for Cyrillic-heavy data (helps pick cp1251 vs "unknown").
     */
    public static function hasCyrillicLikely(string $bytes): bool
    {
        // If it's valid UTF-8 and contains Cyrillic in Unicode, iconv might help,
        // but we want a cheap byte-based heuristic:
        // cp1251 Cyrillic uppercase/lowercase: 0xC0-0xFF (plus 0xA8,0xB8)
        $len = strlen($bytes);
        $hits = 0;
        $limit = min($len, 65536);

        for ($i = 0; $i < $limit; $i++) {
            $b = ord($bytes[$i]);
            if (($b >= 0xC0 && $b <= 0xFF) || $b === 0xA8 || $b === 0xB8) {
                $hits++;
                if ($hits >= 32) {
                    return true;
                }
            }
        }

        return false;
    }

    /**
     * Pick best Cyrillic single-byte encoding by decoding sample and scoring output.
     * Works best for natural-language text (headers, comments, descriptions).
     */
    public static function detectCyrillicEncoding(string $sample, array $candidates = null): string
    {
        if ($candidates === null) {
            $candidates = ['Windows-1251', 'CP866', 'KOI8-R', 'ISO-8859-5', 'MacCyrillic'];
        }

        $bestEnc = $candidates[0];
        $bestScore = -INF;

        foreach ($candidates as $enc) {
            $utf8 = @iconv($enc, 'UTF-8//IGNORE', $sample);
            if ($utf8 === false || $utf8 === '') {
                continue;
            }

            $score = self::scoreDecodedUtf8Text($utf8);

            // Small tie-break: if CP866 and CP1251 close, prefer CP866 when many box-drawing chars present in raw bytes
            // (CP866 часто содержит псевдографику в 0xB0-0xDF)
            if (($enc === 'CP866' || $enc === 'IBM866') && self::looksLikeCp866BoxDrawing($sample)) {
                $score += 2.0;
            }

            if ($score > $bestScore) {
                $bestScore = $score;
                $bestEnc = $enc;
            }
        }

        return $bestEnc;
    }

    /**
     * Score UTF-8 text: prefer many Cyrillic letters and generally printable text.
     */
    public static function scoreDecodedUtf8Text(string $utf8): float
    {
        // Count total letters, Cyrillic letters, printable chars, control chars
        $len = mb_strlen($utf8, 'UTF-8');
        if ($len === 0) {
            return -INF;
        }

        // Cyrillic: \p{Cyrillic} includes letters + some marks; we'll count letters mainly
        preg_match_all('/\p{Cyrillic}/u', $utf8, $m1);
        $cyr = count($m1[0]);

        // Printable-ish: letters/digits/punct/space
        preg_match_all('/[\p{L}\p{N}\p{P}\p{Zs}]/u', $utf8, $m2);
        $printable = count($m2[0]);

        // Controls (except \r\n\t)
        preg_match_all('/[\p{Cc}]/u', $utf8, $m3);
        $ctrl = count($m3[0]);

        // A few replacement chars can appear if some conversion produced them (rare with //IGNORE, but keep)
        preg_match_all("/\xEF\xBF\xBD/u", $utf8, $m4);
        $repl = count($m4[0]);

        // Score: printable density + Cyrillic density, minus bad stuff
        $printableRatio = $printable / $len;
        $cyrRatio = $cyr / $len;

        $score = 0.0;
        $score += $printableRatio * 50.0;
        $score += $cyrRatio * 80.0;
        $score -= $ctrl * 1.5;
        $score -= $repl * 5.0;

        return $score;
    }

    /**
     * Heuristic: CP866 pseudo-graphics bytes are common in console dumps.
     */
    public static function looksLikeCp866BoxDrawing(string $bytes): bool
    {
        $hits = 0;
        $limit = min(strlen($bytes), 65536);
        for ($i = 0; $i < $limit; $i++) {
            $b = ord($bytes[$i]);
            if ($b >= 0xB0 && $b <= 0xDF) {
                $hits++;
                if ($hits >= 64) {
                    return true;
                }
            }
        }

        return false;
    }

    /**
     * Best-effort encoding guess:
     *  - BOM wins
     *  - else valid UTF-8 -> UTF-8
     *  - else if Cyrillic bytes -> Windows-1251 (or other cyrillic encoding)
     *  - else ISO-8859-1 (or keep as binary; choose what you prefer)
     */
    public static function guessInputEncodingFromSample(string $sample): string
    {
        $bomEnc = self::detectEncodingByBomBytes(substr($sample, 0, 4));
        if ($bomEnc) {
            return $bomEnc;
        }

        if (self::looksLikeUtf8($sample)) {
            return 'UTF-8';
        }

        if (self::hasCyrillicLikely($sample)) {
            return self::detectCyrillicEncoding($sample, ['Windows-1251', 'CP866', 'KOI8-R', 'ISO-8859-5']);
        }

        return 'ISO-8859-1';
    }

    /**
     * @param string $sample
     *
     * @return string
     */
    public static function detectEncoding(string $sample): string
    {
        // If BOM exists: advance pointer past BOM.
        $bom = self::detectEncodingByBomBytes($sample);
        if ($bom) {
            // Trust BOM more than heuristic
            $encoding = $bom;
        }
        else {
            $encoding = self::guessInputEncodingFromSample($sample);
        }

        return $encoding;
    }

    /**
     * Detect delimiter in CSV/TSV by sampling the beginning of a stream or file
     *
     * @param string $sample
     * @param array|null $candidates Delimiters to try
     * @param int $sampleBytes How many bytes to read from the beginning
     * @param int $maxLines How many logical lines to analyze
     *
     * @return array{
     *   delimiter: ?string,
     *   confidence: float,
     *   details: array<string, array{score: float, mode: int, mode_ratio: float, stdev: float, lines: int}>
     * }
     */
    public static function detectCsvDelimiter(string $sample, ?array $candidates = null, int $sampleBytes = 131072, int $maxLines = 80): array
    {
        $candidates ??= [",", ";", "\t", "|", ":"];

        // Get first logical lines (handles newlines inside quotes)
        $lines = self::extractLogicalLines($sample, $maxLines);
        if (count($lines) < 3) {
            // Not enough data; best effort: pick delimiter with max occurrences in first line
            $first = $lines[0] ?? $sample;
            $best = null;
            $bestCnt = 0;
            foreach ($candidates as $d) {
                $cnt = self::countDelimiterOutsideQuotes($first, $d);
                if ($cnt > $bestCnt) {
                    $bestCnt = $cnt;
                    $best = $d;
                }
            }
            return [
                'delimiter' => $bestCnt > 0 ? $best : null,
                'confidence' => $bestCnt > 0 ? 0.4 : 0.0,
                'details' => [],
                'guess' => self::guessByLocale($lines),
            ];
        }

        // Heuristic: detect many European decimals "12,34" (comma decimal separator)
        $commaDecimalHits = self::countCommaDecimals($lines);

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
                $cnt = self::countDelimiterOutsideQuotes($line, $d);
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

            $mode = self::modeInt($fields);
            $modeRatio = ($mode > 0) ? (count(array_filter($fields, fn($v) => $v === $mode)) / count($fields)) : 0.0;
            $stdev = self::stddev($fields);

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
                    $tabHits = self::countTabs($lines);
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
            'guess' => self::guessByLocale(),
        ];
    }

    /**
     * Extract up to $max logical lines from text, handling newlines inside quotes.
     *
     * @param string $text
     * @param int $max
     *
     * @return array
     */
    protected static function extractLogicalLines(string $text, int $max): array
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
    protected static function countDelimiterOutsideQuotes(string $line, string $delim): int
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
    protected static function modeInt(array $values): int
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
    protected static function stddev(array $values): float
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
    protected static function countCommaDecimals(array $lines): int
    {
        $hits = 0;
        // Keep it lightweight: simple scan, not heavy regex loops over huge strings.
        foreach ($lines as $line) {
            $len = strlen($line);
            for ($i = 1; $i < $len - 1; $i++) {
                if ($line[$i] === ',' && ctype_digit($line[$i - 1]) && ctype_digit($line[$i + 1])) {
                    $hits++;
                    if ($hits >= 20) {
                        return $hits;
                    } // cap
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
    protected static function countTabs(array $lines): int
    {
        $hits = 0;
        foreach ($lines as $line) {
            $hits += substr_count($line, "\t");
            if ($hits >= 20) return $hits; // cap
        }
        return $hits;
    }


    protected static function guessByLocale(): ?string
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

    /**
     * Read a sample from resource or file.
     *
     * @param string $path
     * @param int $sampleBytes
     *
     * @return string|null
     */
    public static function readSample(string $path, int $sampleBytes = 65536): ?string
    {
        if ($path && is_file($path)) {
            $fp = fopen($path, 'rb');
            if ($fp) {
                $sample = fread($fp, $sampleBytes);
                fclose($fp);

                return $sample === false ? null : $sample;
            }
        }

        return null;
    }
}