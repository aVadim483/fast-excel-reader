<?php

namespace avadim\FastExcelReader\Xls;

/**
 * BIFF8 record type numbers
 *
 * Only the records this reader acts on are listed. Everything else is skipped
 * by length, which is what makes the format tolerant of unknown records.
 *
 * @see [MS-XLS] section 2.3 for the full catalogue
 */
final class BiffRecord
{
    // stream and substream structure
    public const BOF = 0x0809;
    public const EOF = 0x000A;
    public const CONTINUE = 0x003C;

    // BOF substream types
    public const SUBSTREAM_GLOBALS = 0x0005;
    public const SUBSTREAM_WORKSHEET = 0x0010;
    public const SUBSTREAM_CHART = 0x0020;
    public const SUBSTREAM_MACRO = 0x0040;

    // workbook globals
    public const FILEPASS = 0x002F;
    public const CODEPAGE = 0x0042;
    public const DATEMODE = 0x0022;
    public const BOUNDSHEET = 0x0085;
    public const SST = 0x00FC;
    public const FORMAT = 0x041E;
    public const FONT = 0x0031;
    public const XF = 0x00E0;
    public const PALETTE = 0x0092;
    public const NAME = 0x0018;
    public const MSODRAWINGGROUP = 0x00EB;

    // worksheet
    public const DIMENSIONS = 0x0200;
    public const ROW = 0x0208;
    public const BLANK = 0x0201;
    public const MULBLANK = 0x00BE;
    public const NUMBER = 0x0203;
    public const RK = 0x027E;
    public const MULRK = 0x00BD;
    public const LABEL = 0x0204;
    public const LABELSST = 0x00FD;
    public const RSTRING = 0x00D6;
    public const BOOLERR = 0x0205;
    public const FORMULA = 0x0006;
    public const STRING = 0x0207;
    public const SHRFMLA = 0x04BC;
    public const ARRAY_RECORD = 0x0221;
    public const MERGEDCELLS = 0x00E5;
    public const COLINFO = 0x007D;
    public const WINDOW2 = 0x023E;
    public const MSODRAWING = 0x00EC;
    public const OBJ = 0x005D;
    public const DEFAULTROWHEIGHT = 0x0225;

    /** BIFF version stored in the BOF record */
    public const VERSION_BIFF8 = 0x0600;
    public const VERSION_BIFF5 = 0x0500;

    /**
     * Error codes carried by BOOLERR and by formula results
     *
     * @var array<int, string>
     */
    public const ERROR_CODES = [
        0x00 => '#NULL!',
        0x07 => '#DIV/0!',
        0x0F => '#VALUE!',
        0x17 => '#REF!',
        0x1D => '#NAME?',
        0x24 => '#NUM!',
        0x2A => '#N/A',
    ];

    /**
     * Human readable name of a record, for diagnostics
     *
     * @param int $type
     *
     * @return string
     */
    public static function name(int $type): string
    {
        static $names = null;
        if ($names === null) {
            $names = [];
            foreach ((new \ReflectionClass(self::class))->getConstants() as $constant => $value) {
                if (is_int($value) && !isset($names[$value]) && strpos($constant, 'SUBSTREAM_') !== 0 && strpos($constant, 'VERSION_') !== 0) {
                    $names[$value] = $constant;
                }
            }
        }

        return $names[$type] ?? sprintf('0x%04X', $type);
    }
}
