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
     * Check if encoding is available in mbstring (return null if not available)
     *
     * @param string $encoding
     *
     * @return string|null
     */
    public static function availableEncoding(string $encoding): ?string
    {
        try {
            $aliases = mb_encoding_aliases($encoding);
        } catch (\Throwable $e) {
            $aliases = [];
        }
        if (strcasecmp($encoding, 'SHIFT_JIS') === 0) {
            $aliases[] = 'SJIS-WIN';
        }
        $list = mb_list_encodings();
        foreach ($list as $enc) {
            if (strcasecmp($enc, $encoding) === 0) {
                return $encoding;
            }
            foreach ($aliases as $alias) {
                if (strcasecmp($enc, $alias) === 0) {
                    return $alias;
                }
            }
        }

        return null;
    }

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
     * @param string $bytes
     * 
     * @return bool
     */
    public static function looksLikeJapaneseBytes(string $bytes): bool
    {
        $len = strlen($bytes);
        if ($len < 64) {
            return false;
        }

        $limit = min($len, 65536);

        $hi = 0;
        $sjisLead = 0;
        $eucLead = 0;
        $eucPrefix = 0;

        for ($i = 0; $i < $limit; $i++) {
            $b = ord($bytes[$i]);

            if ($b >= 0x80) {
                $hi++;
            }

            // Shift_JIS lead bytes (most common ranges)
            if (($b >= 0x81 && $b <= 0x9F) || ($b >= 0xE0 && $b <= 0xFC)) {
                $sjisLead++;
            }

            // EUC-JP lead range and prefixes
            if ($b === 0x8E || $b === 0x8F) {
                $eucPrefix++;
            } elseif ($b >= 0xA1 && $b <= 0xFE) {
                $eucLead++;
            }
        }

        // If most bytes are ASCII, it is definitely not a Japanese legacy encoding
        if ($hi < 16) {
            return false;
        }

        // Threshold heuristics: there must be noticeable signs of SJIS/EUC
        // (tuned for text CSVs with Japanese headers/values)
        if ($sjisLead >= 8) {
            return true;
        }
        if ($eucLead >= 12 || $eucPrefix >= 2) {
            return true;
        }

        return false;
    }

    /**
     * Choose the best Japanese encoding by decoding a sample and evaluating the result.
     *
     * Note: In Windows practice, CP932 (Windows-932 / SJIS-win) is more common.
     */
    public static function detectJapaneseEncoding(string $sample, array $candidates): array
    {
        $bestEnc = $candidates[0];
        $bestScore = -INF;

        foreach ($candidates as $enc) {
            // Try conversion
            $utf8 = @mb_convert_encoding($sample, 'UTF-8', $enc);
            if ($utf8 === false || $utf8 === '') {
                continue;
            }

            $score = self::scoreJapaneseUtf8Text($utf8);
            // CP932 vs EUC-JP
            $rt = self::jpRoundTripSimilarity($sample, $enc); // 0..1
            if ($rt < 0.90) {
                $score -= (0.90 - $rt) * 200.0; // сильный штраф
            }
            else {
                $score += ($rt - 0.90) * 200.0; // небольшой бонус за идеальность
            }

            // Small bonus for CP932 (Windows practice)
            if (strcasecmp($enc, 'CP932') === 0) {
                $score += 0.05;
            }

            if ($score > $bestScore) {
                $bestScore = $score;
                $bestEnc = $enc;
            }
        }

        return ['enc' => $bestEnc, 'score' => $bestScore];
    }

    /**
     * Scoring for decoding result: the more Japanese scripts, the better.
     * We encourage:
     * - Hiragana (\p{Hiragana})
     * - Katakana (\p{Katakana})
     * - Han (Kanji, \p{Han})
     * + overall "printability" of the text
     * and penalize:
     * - control characters
     */
    public static function scoreJapaneseUtf8Text(string $utf8): float
    {
        if (!function_exists('mb_strlen')) {
            // One can live without mbstring, but accuracy is slightly worse.
            $len = strlen($utf8);
            if ($len === 0) {
                return -INF;
            }

            // Count at least the presence of Japanese Unicode ranges via preg
            $jp = 0;
            if (preg_match_all('/[\x{3040}-\x{30FF}\x{4E00}-\x{9FFF}]/u', $utf8, $m)) {
                $jp = count($m[0]);
            }

            $printable = 0;
            if (preg_match_all('/[\p{L}\p{N}\p{P}\p{Zs}]/u', $utf8, $m2)) {
                $printable = count($m2[0]);
            }

            $ctrl = 0;
            if (preg_match_all('/[\p{Cc}]/u', $utf8, $m3)) {
                $ctrl = count($m3[0]);
            }

            $score = 0.0;
            $score += ($printable / max(1, $len)) * 50.0;
            $score += ($jp / max(1, $len)) * 120.0;
            $score -= $ctrl * 1.5;

            return $score;
        }

        $len = mb_strlen($utf8, 'UTF-8');
        if ($len === 0) {
            return -INF;
        }

        preg_match_all('/\p{Hiragana}/u', $utf8, $mH);
        preg_match_all('/\p{Katakana}/u', $utf8, $mK);
        preg_match_all('/\p{Han}/u', $utf8, $mHan);
        $hir = count($mH[0]);
        $kat = count($mK[0]);
        $han = count($mHan[0]);

        preg_match_all('/[\p{L}\p{N}\p{P}\p{Zs}]/u', $utf8, $mP);
        $printable = count($mP[0]);

        preg_match_all('/[\p{Cc}]/u', $utf8, $mC);
        $ctrl = count($mC[0]);

        // Weights can be adjusted, but these usually work well:
        // - kanji/kana give a strong signal
        // - printable density protects against "garbage"
        $printableRatio = $printable / $len;
        $jpRatio = ($hir + $kat + $han) / $len;

        $score = $printableRatio * 50.0;
        $score += $jpRatio * 120.0;
        $score += ($hir / $len) * 10.0;  // small bonus for hiragana
        $score += ($kat / $len) * 6.0;   // and katakana
        $score -= $ctrl * 1.5;

        return $score;
    }

    /**
     * Round-trip similarity (for Japanese)
     *
     * @param string $bytes
     * @param string $enc
     *
     * @return float
     */
    protected static function jpRoundTripSimilarity(string $bytes, string $enc): float
    {
        if (!function_exists('mb_convert_encoding')) {
            return 0.0;
        }

        $utf8 = @mb_convert_encoding($bytes, 'UTF-8', $enc);
        if ($utf8 === false || $utf8 === '') {
            return 0.0;
        }

        $back = @mb_convert_encoding($utf8, $enc, 'UTF-8');
        if ($back === false || $back === '') {
            return 0.0;
        }

        $n = min(strlen($bytes), strlen($back));
        if ($n === 0) {
            return 0.0;
        }

        $same = 0;
        for ($i = 0; $i < $n; $i++) {
            if ($bytes[$i] === $back[$i]) {
                $same++;
            }
        }

        // 0..1
        return $same / $n;
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
        $nonAsciiCount = 0;

        for ($i = 0; $i < $limit; $i++) {
            $b = ord($bytes[$i]);
            // non-ASCII
            if ($b > 127) {
                $nonAsciiCount++;
            }
            if (($b >= 0xC0 && $b <= 0xFF) || $b === 0xA8 || $b === 0xB8) {
                $hits++;
                if ($hits >= 32) {
                    return true;
                }
            }
        }

        return $hits === $nonAsciiCount;
    }

    /**
     * Pick the best Cyrillic single-byte encoding by decoding sample and scoring output.
     * Works best for natural-language text (headers, comments, descriptions).
     */
    public static function detectCyrillicEncoding(string $sample, array $candidates): array
    {
        $bestEnc = $candidates[0];
        $bestScore = -INF;

        foreach ($candidates as $enc) {
            $utf8 = @mb_convert_encoding($sample, 'UTF-8', $enc);
            if ($utf8 === false || $utf8 === '') {
                continue;
            }

            $score = self::scoreDecodedUtf8Text($utf8);
            $score += self::russianLanguageScore($utf8) * 200;

            // Small tie-break: if CP866 and CP1251 close, prefer CP866 when many box-drawing chars present in raw bytes
            // (CP866 often contains pseudographics in 0xB0-0xDF)
            if (($enc === 'CP866' || $enc === 'IBM866') && self::looksLikeCp866BoxDrawing($sample)) {
                $score += 2.0;
            }

            if ($score > $bestScore) {
                $bestScore = $score;
                $bestEnc = $enc;
            }
        }

        return ['enc' => $bestEnc, 'score' => $bestScore];
    }

    /**
     * @param string $utf8
     *
     * @return float
     */
    protected static function russianLanguageScore(string $utf8): float
    {
        $s = mb_strtolower($utf8, 'UTF-8');

        // оставим только кириллицу+пробелы, чтобы не шумели цифры/пунктуация
        $s = preg_replace('/[^\p{Cyrillic}\s]+/u', ' ', $s);
        $s = preg_replace('/\s+/u', ' ', trim($s));
        if ($s === '') return 0.0;

        $bigrams = [
            'ст'=>3.0,'но'=>2.8,'то'=>2.6,'на'=>2.6,'ен'=>2.4,'ов'=>2.2,'ни'=>2.0,'ра'=>1.9,'ко'=>1.8,'ро'=>1.7,
            'го'=>1.7,'по'=>1.6,'пр'=>1.6,'ос'=>1.5,'ло'=>1.5,'ли'=>1.5,'ер'=>1.4,'ал'=>1.4,'ет'=>1.4,'ан'=>1.3,
        ];

        $score = 0.0;
        foreach ($bigrams as $bg => $w) {
            $score += substr_count($s, $bg) * $w;
        }

        // бонус за частые слова (пробелы важны)
        $words = [' и '=>1.5,' в '=>1.2,' на '=>1.2,' не '=>1.1,' что '=>1.0,' это '=>0.9,' как '=>0.8,' для '=>0.8];
        $padded = ' ' . $s . ' ';
        foreach ($words as $w => $k) {
            $score += substr_count($padded, $w) * $k;
        }

        // нормализуем на длину
        $len = max(1, mb_strlen($s, 'UTF-8'));
        return $score / $len;
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

        // Japanese branch (before Cyrillic): CP932/Shift_JIS/EUC-JP
        if (self::looksLikeJapaneseBytes($sample)) {
            $jpEnc = self::detectJapaneseEncoding($sample, ['CP932', 'Shift_JIS', 'EUC-JP']);
        }
        else {
            $jpEnc = ['enc' => null, 'score' => -INF];
        }

        if (self::hasCyrillicLikely($sample)) {
            $cyEnc = self::detectCyrillicEncoding($sample, ['Windows-1251', 'CP866', 'KOI8-R', 'ISO-8859-5'],);
        }
        else {
            $cyEnc = ['enc' => null, 'score' => -INF];
        }

        // The "margin" is needed to prevent random noise from interrupting the normal text
        if ($jpEnc['enc'] !== null && $cyEnc['enc'] !== null) {
            $margin = 6.0;

            // Minimum thresholds to prevent "garbage" selection
            $minJp = 10.0;
            $minCy = 10.0;

            if ($jpEnc['score'] >= $minJp && $jpEnc['score'] > $cyEnc['score'] + $margin) {
                return $jpEnc['enc'];
            }
            if ($cyEnc['score'] >= $minCy && $cyEnc['score'] >= $jpEnc['score'] - $margin) {
                return $cyEnc['enc'];
            }

            // If both are low, you're not sure, but it's better to return the one that's "less trashy."
            return ($jpEnc['score'] > $cyEnc['score']) ? $jpEnc['enc'] : $cyEnc['enc'];
        }

        // If only one branch was triggered, we return it, but with a minimum threshold.
        if ($cyEnc['enc'] !== null && $cyEnc['score'] >= 10.0) {
            return $cyEnc['enc'];
        }
        if ($jpEnc['enc'] !== null && $jpEnc['score'] >= 10.0) {
            return $jpEnc['enc'];
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

    /**
     * @return string|null
     */
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