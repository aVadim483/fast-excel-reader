<?php

namespace avadim\FastExcelReader\Xls;

/**
 * Parser for the Office Drawing ("Escher") structures that carry pictures
 *
 * Escher is a tree of records nested inside the MSODRAWINGGROUP and MSODRAWING
 * BIFF records. Each record has an 8 byte header - version+instance, type,
 * length - and a record whose version nibble is 0xF is a container holding more
 * records rather than data of its own.
 *
 * Two parts matter for reading pictures:
 *
 * - the workbook-level blip store, which holds the image bytes themselves, and
 * - the per-sheet drawing, where each shape names the blip it displays and the
 *   cell it is anchored to.
 */
class XlsEscher
{
    private const CONTAINER_VERSION = 0x0F;

    // record types
    private const BSE = 0xF007;          // blip store entry
    private const SP_CONTAINER = 0xF004; // shape
    private const OPT = 0xF00B;          // shape property table
    private const CLIENT_ANCHOR = 0xF010;

    /** Shape property holding the 1-based index of the blip to display */
    private const PROPERTY_BLIP_INDEX = 0x0104;

    /**
     * File extensions by MSOBLIPTYPE
     *
     * @var array<int, string>
     */
    private const BLIP_TYPES = [
        2 => 'emf',
        3 => 'wmf',
        4 => 'pict',
        5 => 'jpeg',
        6 => 'png',
        7 => 'dib',
        17 => 'tiff',
    ];

    /**
     * MIME types by MSOBLIPTYPE
     *
     * @var array<int, string>
     */
    private const BLIP_MIME = [
        2 => 'image/x-emf',
        3 => 'image/x-wmf',
        4 => 'image/x-pict',
        5 => 'image/jpeg',
        6 => 'image/png',
        7 => 'image/bmp',
        17 => 'image/tiff',
    ];

    /**
     * Extract the picture data from the workbook drawing group
     *
     * @param string $data Payload of MSODRAWINGGROUP
     *
     * @return array<int, array{ext: string, mime: string, data: string}> Keyed by 1-based blip index
     */
    public static function blipStore(string $data): array
    {
        $blips = [];
        $index = 0;

        foreach (self::records($data) as $record) {
            if ($record['type'] !== self::BSE) {
                continue;
            }
            $index++;
            $blip = self::readBlip($record['data']);
            if ($blip !== null) {
                $blips[$index] = $blip;
            }
        }

        return $blips;
    }

    /**
     * Extract the shapes of one sheet: which picture sits in which cell
     *
     * @param string $data Payload of MSODRAWING
     *
     * @return array<int, array{blip: int, row: int, col: int}> Row and column are zero-based
     */
    public static function shapes(string $data): array
    {
        $shapes = [];

        foreach (self::records($data) as $record) {
            if ($record['type'] !== self::SP_CONTAINER) {
                continue;
            }
            $blip = null;
            $anchor = null;

            foreach (self::records($record['data'], false) as $child) {
                if ($child['type'] === self::OPT) {
                    $blip = self::blipIndex($child['data'], $child['instance']);
                }
                elseif ($child['type'] === self::CLIENT_ANCHOR && strlen($child['data']) >= 18) {
                    $fields = unpack('vflag/vcol1/vdx1/vrow1', substr($child['data'], 0, 8));
                    $anchor = ['row' => $fields['row1'], 'col' => $fields['col1']];
                }
            }

            if ($blip !== null && $anchor !== null) {
                $shapes[] = ['blip' => $blip, 'row' => $anchor['row'], 'col' => $anchor['col']];
            }
        }

        return $shapes;
    }

    /**
     * Walk the record tree
     *
     * Containers are descended into, so a caller sees every record at any depth
     * in document order. Pass FALSE to stay at one level, which is what reading
     * the children of a single shape needs.
     *
     * @param string $data
     * @param bool $recursive
     *
     * @return \Generator
     */
    private static function records(string $data, bool $recursive = true): \Generator
    {
        $pos = 0;
        $length = strlen($data);

        while ($pos + 8 <= $length) {
            $verInst = unpack('v', substr($data, $pos, 2))[1];
            $type = unpack('v', substr($data, $pos + 2, 2))[1];
            $size = unpack('V', substr($data, $pos + 4, 4))[1];

            if ($size < 0 || $pos + 8 + $size > $length) {
                // truncated or malformed: stop rather than read past the end
                $size = $length - $pos - 8;
            }
            $body = substr($data, $pos + 8, $size);

            $record = [
                'type' => $type,
                'instance' => $verInst >> 4,
                'data' => $body,
            ];

            if (($verInst & 0x0F) === self::CONTAINER_VERSION) {
                // a shape container is reported itself, its children are read by
                // the caller; anything else is descended into
                if ($type === self::SP_CONTAINER) {
                    yield $record;
                }
                elseif ($recursive) {
                    yield from self::records($body);
                }
            }
            else {
                yield $record;
            }

            $pos += 8 + $size;
        }
    }

    /**
     * Pull the image bytes out of a blip store entry
     *
     * The entry begins with a 36 byte header naming the picture type, followed
     * by the blip record itself. The blip carries one or two 16 byte checksums
     * and a tag byte before the actual file content.
     *
     * @param string $data
     *
     * @return array{ext: string, mime: string, data: string}|null
     */
    private static function readBlip(string $data): ?array
    {
        if (strlen($data) < 44) {
            return null;
        }
        $blipType = ord($data[0]);
        if (!isset(self::BLIP_TYPES[$blipType])) {
            return null;
        }

        $pos = 36;
        $verInst = unpack('v', substr($data, $pos, 2))[1];
        $size = unpack('V', substr($data, $pos + 4, 4))[1];
        $instance = $verInst >> 4;
        $body = substr($data, $pos + 8, $size);

        // an odd instance means the blip repeats its checksum
        $skip = (($instance & 1) ? 32 : 16) + 1;
        if (strlen($body) <= $skip) {
            return null;
        }

        return [
            'ext' => self::BLIP_TYPES[$blipType],
            'mime' => self::BLIP_MIME[$blipType],
            'data' => substr($body, $skip),
        ];
    }

    /**
     * Find the blip index in a shape property table
     *
     * @param string $data
     * @param int $count Number of properties, carried by the record instance
     *
     * @return int|null
     */
    private static function blipIndex(string $data, int $count): ?int
    {
        for ($i = 0; $i < $count; $i++) {
            $offset = $i * 6;
            if ($offset + 6 > strlen($data)) {
                break;
            }
            $id = unpack('v', substr($data, $offset, 2))[1] & 0x3FFF;
            if ($id === self::PROPERTY_BLIP_INDEX) {
                return unpack('V', substr($data, $offset + 2, 4))[1];
            }
        }

        return null;
    }
}
