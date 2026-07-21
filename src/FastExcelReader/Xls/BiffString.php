<?php

namespace avadim\FastExcelReader\Xls;

/**
 * Decoding of BIFF8 string structures
 *
 * A BIFF8 string is either "compressed" - one byte per character, the low half
 * of UCS-2, i.e. Latin-1 - or plain UTF-16LE, selected by a flag byte. Rich
 * text runs and Far East phonetic data may follow the characters.
 *
 * The awkward part is the shared string table. A record longer than 8224 bytes
 * continues in CONTINUE records, and a string may be cut in the middle: the
 * continuation then starts with a *fresh* flag byte, and the encoding may
 * differ from the one the string started with. Decoding therefore has to know
 * where the segment boundaries are, which is why readSharedStrings() takes the
 * record parts rather than one concatenated buffer.
 */
class BiffString
{
    /**
     * Read a string whose length is stored in one byte (ShortXLUnicodeString)
     *
     * Used by BOUNDSHEET for sheet names.
     *
     * @param string $data
     * @param int $offset
     *
     * @return array{0: string, 1: int} The string and the number of bytes consumed
     */
    public static function readShort(string $data, int $offset): array
    {
        if ($offset >= strlen($data)) {
            return ['', 0];
        }
        $charCount = ord($data[$offset]);
        $flags = ord($data[$offset + 1] ?? "\0");
        $wide = ($flags & 0x01) !== 0;
        $bytes = $charCount * ($wide ? 2 : 1);
        $raw = substr($data, $offset + 2, $bytes);

        return [self::decode($raw, $wide), 2 + $bytes];
    }

    /**
     * Read a string whose length is stored in two bytes (XLUnicodeString)
     *
     * @param string $data
     * @param int $offset
     *
     * @return array{0: string, 1: int} The string and the number of bytes consumed
     */
    public static function readLong(string $data, int $offset): array
    {
        if ($offset + 2 > strlen($data)) {
            return ['', 0];
        }
        $charCount = unpack('v', substr($data, $offset, 2))[1];
        $flags = ord($data[$offset + 2] ?? "\0");
        $wide = ($flags & 0x01) !== 0;
        $bytes = $charCount * ($wide ? 2 : 1);
        $raw = substr($data, $offset + 3, $bytes);

        return [self::decode($raw, $wide), 3 + $bytes];
    }

    /**
     * Decode the shared string table
     *
     * @param string[] $parts Payload of the SST record followed by its CONTINUE records
     *
     * @return string[] Strings indexed as referenced by LABELSST
     */
    public static function readSharedStrings(array $parts): array
    {
        $data = implode('', $parts);
        if (strlen($data) < 8) {
            return [];
        }

        // offsets at which a CONTINUE segment starts, used to spot a split string;
        // kept sorted so that locating the next one is a walk, not a search
        $boundaries = [];
        $at = 0;
        foreach ($parts as $index => $part) {
            if ($index > 0) {
                $boundaries[] = $at;
            }
            $at += strlen($part);
        }

        $uniqueCount = unpack('V', substr($data, 4, 4))[1];
        $position = 8;
        $length = strlen($data);
        $strings = [];
        $cursor = 0;

        for ($i = 0; $i < $uniqueCount && $position < $length; $i++) {
            $strings[] = self::readRichExtended($data, $position, $boundaries, $cursor);
        }

        return $strings;
    }

    /**
     * Read one XLUnicodeRichExtendedString, following CONTINUE boundaries
     *
     * @param string $data
     * @param int $position Advanced past the string
     * @param int[] $boundaries
     * @param int $cursor Index of the first boundary not yet passed
     *
     * @return string
     */
    private static function readRichExtended(string $data, int &$position, array $boundaries, int &$cursor): string
    {
        $length = strlen($data);
        if ($position + 3 > $length) {
            $position = $length;

            return '';
        }

        $charCount = unpack('v', substr($data, $position, 2))[1];
        $position += 2;
        $flags = ord($data[$position]);
        $position++;

        $wide = ($flags & 0x01) !== 0;
        $hasPhonetic = ($flags & 0x04) !== 0;
        $hasRichRuns = ($flags & 0x08) !== 0;

        $runCount = 0;
        if ($hasRichRuns) {
            $runCount = unpack('v', substr($data, $position, 2))[1];
            $position += 2;
        }
        $phoneticSize = 0;
        if ($hasPhonetic) {
            $phoneticSize = unpack('V', substr($data, $position, 4))[1];
            $position += 4;
        }

        $result = '';
        $remaining = $charCount;
        while ($remaining > 0 && $position < $length) {
            $charSize = $wide ? 2 : 1;
            $nextBoundary = self::nextBoundary($boundaries, $cursor, $position, $length);
            $take = min($remaining, intdiv($nextBoundary - $position, $charSize));

            if ($take > 0) {
                $result .= self::decode(substr($data, $position, $take * $charSize), $wide);
                $position += $take * $charSize;
                $remaining -= $take;
            }

            if ($remaining > 0) {
                // the string continues in the next segment, which restates the flag byte
                if ($position >= $length) {
                    break;
                }
                $wide = (ord($data[$position]) & 0x01) !== 0;
                $position++;
            }
        }

        $position += $runCount * 4 + $phoneticSize;

        return $result;
    }

    /**
     * Offset of the first segment boundary strictly after $position
     *
     * $boundaries is ascending, and $position only ever moves forward, so the
     * cursor into it moves forward too - the whole table costs one pass.
     *
     * @param int[] $boundaries
     * @param int $cursor Index of the first boundary not yet passed
     * @param int $position
     * @param int $length
     *
     * @return int
     */
    private static function nextBoundary(array $boundaries, int &$cursor, int $position, int $length): int
    {
        $count = count($boundaries);
        while ($cursor < $count && $boundaries[$cursor] <= $position) {
            $cursor++;
        }

        return $cursor < $count ? $boundaries[$cursor] : $length;
    }

    /**
     * @param string $raw
     * @param bool $wide
     *
     * @return string
     */
    private static function decode(string $raw, bool $wide): string
    {
        if ($raw === '') {
            return '';
        }

        return (string)mb_convert_encoding($raw, 'UTF-8', $wide ? 'UTF-16LE' : 'ISO-8859-1');
    }
}
