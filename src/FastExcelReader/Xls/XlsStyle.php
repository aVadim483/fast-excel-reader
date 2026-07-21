<?php

namespace avadim\FastExcelReader\Xls;

/**
 * Decoding of the BIFF8 formatting records
 *
 * Where OOXML keeps fonts, fills and borders in separate tables and has the
 * cell format reference them by index, BIFF packs the fill and the borders
 * directly into the XF record as bit fields. Everything here is therefore
 * mostly bit extraction; turning the result back into the indexed tables that
 * the rest of the library expects is XlsBook's job.
 *
 * Key names deliberately match the ones the XLSX reader produces, so that a
 * style read from either format is the same array.
 */
class XlsStyle
{
    /**
     * Border line styles by their BIFF index, named as in OOXML
     *
     * @var array<int, string|null>
     */
    public const BORDER_STYLES = [
        0 => null,
        1 => 'thin',
        2 => 'medium',
        3 => 'dashed',
        4 => 'dotted',
        5 => 'thick',
        6 => 'double',
        7 => 'hair',
        8 => 'mediumDashed',
        9 => 'dashDot',
        10 => 'mediumDashDot',
        11 => 'dashDotDot',
        12 => 'mediumDashDotDot',
        13 => 'slantDashDot',
    ];

    /**
     * Fill patterns by their BIFF index, named as in OOXML
     *
     * @var array<int, string>
     */
    public const FILL_PATTERNS = [
        0 => 'none',
        1 => 'solid',
        2 => 'mediumGray',
        3 => 'darkGray',
        4 => 'lightGray',
        5 => 'darkHorizontal',
        6 => 'darkVertical',
        7 => 'darkDown',
        8 => 'darkUp',
        9 => 'darkGrid',
        10 => 'darkTrellis',
        11 => 'lightHorizontal',
        12 => 'lightVertical',
        13 => 'lightDown',
        14 => 'lightUp',
        15 => 'lightGrid',
        16 => 'lightTrellis',
        17 => 'gray125',
        18 => 'gray0625',
    ];

    /**
     * Horizontal alignments by their BIFF index
     *
     * @var array<int, string|null>
     */
    public const ALIGN_HORIZONTAL = [
        0 => null, // general, the default
        1 => 'left',
        2 => 'center',
        3 => 'right',
        4 => 'fill',
        5 => 'justify',
        6 => 'centerContinuous',
        7 => 'distributed',
    ];

    /**
     * Vertical alignments by their BIFF index
     *
     * @var array<int, string|null>
     */
    public const ALIGN_VERTICAL = [
        0 => 'top',
        1 => 'center',
        2 => null, // bottom, the default
        3 => 'justify',
        4 => 'distributed',
    ];

    /**
     * Decode a FONT record
     *
     * @param string $data
     *
     * @return array
     */
    public static function font(string $data): array
    {
        $height = unpack('v', substr($data, 0, 2))[1];
        $flags = unpack('v', substr($data, 2, 2))[1];
        $colorIndex = unpack('v', substr($data, 4, 2))[1];
        $weight = unpack('v', substr($data, 6, 2))[1];
        $underline = ord($data[10]);
        $family = ord($data[11]);
        $charset = ord($data[12]);
        [$name] = BiffString::readShort($data, 14);

        // values are strings because that is what the XLSX reader produces,
        // where they come straight from XML attributes
        $font = [
            'font-size' => (string)($height / 20),
            'font-name' => $name,
        ];
        if ($family) {
            $font['font-family'] = (string)$family;
        }
        if ($charset) {
            $font['font-charset'] = (string)$charset;
        }
        if ($weight >= 700) {
            $font['font-style-bold'] = 1;
        }
        if ($flags & 0x0002) {
            $font['font-style-italic'] = 1;
        }
        if ($flags & 0x0008) {
            $font['font-style-strike'] = 1;
        }
        if ($underline) {
            // 0x02 and 0x22 are the double variants, accounting or not
            $font['font-style-underline'] = ($underline === 0x02 || $underline === 0x22) ? 2 : 1;
        }

        return ['font' => $font, 'colorIndex' => $colorIndex];
    }

    /**
     * Decode an XF record into its independent parts
     *
     * @param string $data
     *
     * @return array
     */
    public static function xf(string $data): array
    {
        $fontIndex = unpack('v', substr($data, 0, 2))[1];
        $formatIndex = unpack('v', substr($data, 2, 2))[1];
        $parentFlags = unpack('v', substr($data, 4, 2))[1];
        $alignByte = ord($data[6]);

        $borderStyles = unpack('v', substr($data, 10, 2))[1];
        $borderColors = unpack('v', substr($data, 12, 2))[1];
        $more = unpack('V', substr($data, 14, 4))[1];
        $fillColors = unpack('v', substr($data, 18, 2))[1];

        return [
            'fontIndex' => $fontIndex,
            'formatIndex' => $formatIndex,
            // bit 2 marks a style XF; cells only ever reference cell XFs
            'isStyleXf' => (bool)($parentFlags & 0x0004),
            'parentXf' => ($parentFlags >> 4) & 0x0FFF,
            'align' => [
                'horizontal' => self::ALIGN_HORIZONTAL[$alignByte & 0x07] ?? null,
                'vertical' => self::ALIGN_VERTICAL[($alignByte >> 4) & 0x07] ?? null,
                'wrap' => (bool)($alignByte & 0x08),
            ],
            'border' => [
                'left' => ['style' => $borderStyles & 0x0F, 'color' => $borderColors & 0x7F],
                'right' => ['style' => ($borderStyles >> 4) & 0x0F, 'color' => ($borderColors >> 7) & 0x7F],
                'top' => ['style' => ($borderStyles >> 8) & 0x0F, 'color' => $more & 0x7F],
                'bottom' => ['style' => ($borderStyles >> 12) & 0x0F, 'color' => ($more >> 7) & 0x7F],
                'diagonal' => ['style' => ($more >> 21) & 0x0F, 'color' => ($more >> 14) & 0x7F],
            ],
            'fill' => [
                'pattern' => ($more >> 26) & 0x3F,
                'foreground' => $fillColors & 0x7F,
                'background' => ($fillColors >> 7) & 0x7F,
            ],
        ];
    }

    /**
     * Decode a PALETTE record
     *
     * The palette replaces colour indices 8 and up; 0 to 7 are fixed.
     *
     * @param string $data
     *
     * @return array<int, string> Colour index => #RRGGBB
     */
    public static function palette(string $data): array
    {
        $count = unpack('v', substr($data, 0, 2))[1];
        $palette = [];
        for ($i = 0; $i < $count; $i++) {
            $rgb = substr($data, 2 + $i * 4, 3);
            if (strlen($rgb) < 3) {
                break;
            }
            $palette[8 + $i] = '#' . strtoupper(bin2hex($rgb));
        }

        return $palette;
    }
}
