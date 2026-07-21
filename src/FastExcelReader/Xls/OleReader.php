<?php

namespace avadim\FastExcelReader\Xls;

use avadim\FastExcelReader\Exception;

/**
 * Reader for the OLE2 Compound File Binary container used by .xls files
 *
 * Only what a spreadsheet reader needs: the sector allocation table, the
 * directory, and named streams.
 *
 * Memory matters here. The FAT is kept as a raw binary string and entries are
 * unpacked one at a time, rather than being expanded into a PHP array of
 * integers - for a 100 MB workbook that is roughly 800 KB instead of tens of
 * megabytes. Stream contents are never buffered whole; see OleStream.
 */
class OleReader
{
    private const SIGNATURE = "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1";

    public const ENDOFCHAIN = 0xFFFFFFFE;
    public const FREESECT = 0xFFFFFFFF;

    private const TYPE_STREAM = 2;
    private const TYPE_ROOT = 5;

    /** @var resource */
    private $handle;

    private int $sectorSize;

    private int $miniSectorSize;

    private int $miniCutoff;

    /** Sector allocation table as a raw byte string, 4 bytes per sector */
    private string $fat = '';

    /** Mini sector allocation table, same encoding */
    private string $miniFat = '';

    /** @var array<string, array{start: int, size: int, type: int}> */
    private array $entries = [];

    private ?string $miniStream = null;

    private int $miniStreamStart = 0;

    private int $miniStreamSize = 0;

    /**
     * @param string $file
     *
     * @throws Exception
     */
    public function __construct(string $file)
    {
        if (!is_readable($file)) {
            throw new Exception('File not found or not readable: "' . $file . '"');
        }
        $handle = fopen($file, 'rb');
        if (!$handle) {
            throw new Exception('Cannot open file "' . $file . '"');
        }
        $this->handle = $handle;

        $header = (string)fread($this->handle, 512);
        if (strlen($header) < 512 || strncmp($header, self::SIGNATURE, 8) !== 0) {
            fclose($this->handle);

            throw new Exception('Not an OLE2 compound file: "' . $file . '"');
        }

        $this->sectorSize = 1 << self::uint16($header, 30);
        $this->miniSectorSize = 1 << self::uint16($header, 32);
        $this->miniCutoff = self::uint32($header, 56);

        $this->readFat($header);
        $this->readMiniFat(self::uint32($header, 60), self::uint32($header, 64));
        $this->readDirectory(self::uint32($header, 48));
    }

    public function __destruct()
    {
        if (is_resource($this->handle)) {
            fclose($this->handle);
        }
    }

    /**
     * Byte offset of a sector in the file
     *
     * @param int $sector
     *
     * @return int
     */
    public function sectorOffset(int $sector): int
    {
        return 512 + $sector * $this->sectorSize;
    }

    /**
     * Follow the sector chain by one step
     *
     * @param int $sector
     *
     * @return int|null NULL at the end of the chain
     */
    public function nextSector(int $sector): ?int
    {
        $offset = $sector * 4;
        if ($offset < 0 || $offset + 4 > strlen($this->fat)) {
            return null;
        }
        $next = unpack('V', substr($this->fat, $offset, 4))[1];

        return ($next >= self::ENDOFCHAIN) ? null : $next;
    }

    /**
     * Names of all streams in the container
     *
     * @return string[]
     */
    public function streamList(): array
    {
        return array_keys($this->entries);
    }

    /**
     * @param string $name
     *
     * @return bool
     */
    public function streamExists(string $name): bool
    {
        return isset($this->entries[$name]);
    }

    /**
     * Open a named stream for reading
     *
     * @param string $name
     *
     * @return OleStream
     *
     * @throws Exception
     */
    public function openStream(string $name): OleStream
    {
        if (!isset($this->entries[$name])) {
            throw new Exception('Stream "' . $name . '" not found in the compound file');
        }
        $entry = $this->entries[$name];

        if ($entry['size'] < $this->miniCutoff) {
            return new OleStream($this, null, null, $entry['size'], $this->miniSectorSize, $this->readMiniStream($entry['start'], $entry['size']));
        }

        return new OleStream($this, $this->handle, $entry['start'], $entry['size'], $this->sectorSize);
    }

    /**
     * Assemble the FAT from its own sectors, following the DIFAT
     *
     * @param string $header
     *
     * @return void
     */
    private function readFat(string $header): void
    {
        $fatSectors = [];
        $numFat = self::uint32($header, 44);

        // the first 109 FAT sector numbers live in the header itself
        for ($i = 0; $i < 109 && count($fatSectors) < $numFat; $i++) {
            $sector = self::uint32($header, 76 + $i * 4);
            if ($sector >= self::ENDOFCHAIN) {
                break;
            }
            $fatSectors[] = $sector;
        }

        // the rest is chained through DIFAT sectors
        $difatSector = self::uint32($header, 68);
        $numDifat = self::uint32($header, 72);
        $entriesPerDifat = intdiv($this->sectorSize, 4) - 1;
        for ($i = 0; $i < $numDifat && $difatSector < self::ENDOFCHAIN; $i++) {
            $data = $this->readSector($difatSector);
            for ($j = 0; $j < $entriesPerDifat; $j++) {
                $sector = self::uint32($data, $j * 4);
                if ($sector >= self::ENDOFCHAIN) {
                    break;
                }
                $fatSectors[] = $sector;
            }
            $difatSector = self::uint32($data, $entriesPerDifat * 4);
        }

        foreach ($fatSectors as $sector) {
            $this->fat .= $this->readSector($sector);
        }
    }

    /**
     * @param int $start
     * @param int $count
     *
     * @return void
     */
    private function readMiniFat(int $start, int $count): void
    {
        $sector = $start;
        for ($i = 0; $i < $count && $sector !== null && $sector < self::ENDOFCHAIN; $i++) {
            $this->miniFat .= $this->readSector($sector);
            $sector = $this->nextSector($sector);
        }
    }

    /**
     * Walk the directory chain and index every stream by name
     *
     * @param int $start
     *
     * @return void
     */
    private function readDirectory(int $start): void
    {
        $sector = $start;
        $perSector = intdiv($this->sectorSize, 128);
        $guard = 0;

        while ($sector !== null && $sector < self::ENDOFCHAIN && $guard++ < 65536) {
            $data = $this->readSector($sector);
            for ($i = 0; $i < $perSector; $i++) {
                $entry = substr($data, $i * 128, 128);
                if (strlen($entry) < 128) {
                    break;
                }
                $nameLength = self::uint16($entry, 64);
                $type = ord($entry[66]);
                if ($nameLength < 2 || ($type !== self::TYPE_STREAM && $type !== self::TYPE_ROOT)) {
                    continue;
                }
                // the length includes the terminating null
                $name = (string)mb_convert_encoding(substr($entry, 0, $nameLength - 2), 'UTF-8', 'UTF-16LE');
                $entryStart = self::uint32($entry, 116);
                $entrySize = self::uint32($entry, 120);

                if ($type === self::TYPE_ROOT) {
                    // the root entry describes the mini stream container
                    $this->miniStreamStart = $entryStart;
                    $this->miniStreamSize = $entrySize;

                    continue;
                }
                $this->entries[$name] = ['start' => $entryStart, 'size' => $entrySize, 'type' => $type];
            }
            $sector = $this->nextSector($sector);
        }
    }

    /**
     * Read a stream that lives in the mini stream
     *
     * Such streams are below the cutoff - 4 KB by default - so the simplest
     * correct thing is to gather them at once.
     *
     * @param int $start
     * @param int $size
     *
     * @return string
     */
    private function readMiniStream(int $start, int $size): string
    {
        if ($this->miniStream === null) {
            $this->miniStream = '';
            $sector = $this->miniStreamStart;
            $guard = 0;
            while ($sector !== null && $sector < self::ENDOFCHAIN && $guard++ < 1048576) {
                $this->miniStream .= $this->readSector($sector);
                if (strlen($this->miniStream) >= $this->miniStreamSize) {
                    break;
                }
                $sector = $this->nextSector($sector);
            }
        }

        $result = '';
        $miniSector = $start;
        $guard = 0;
        while ($miniSector !== null && $miniSector < self::ENDOFCHAIN && strlen($result) < $size && $guard++ < 1048576) {
            $result .= substr($this->miniStream, $miniSector * $this->miniSectorSize, $this->miniSectorSize);
            $offset = $miniSector * 4;
            if ($offset + 4 > strlen($this->miniFat)) {
                break;
            }
            $next = unpack('V', substr($this->miniFat, $offset, 4))[1];
            $miniSector = ($next >= self::ENDOFCHAIN) ? null : $next;
        }

        return substr($result, 0, $size);
    }

    /**
     * @param int $sector
     *
     * @return string
     */
    private function readSector(int $sector): string
    {
        fseek($this->handle, $this->sectorOffset($sector));

        return (string)fread($this->handle, $this->sectorSize);
    }

    /**
     * @param string $data
     * @param int $offset
     *
     * @return int
     */
    private static function uint16(string $data, int $offset): int
    {
        return unpack('v', substr($data, $offset, 2))[1];
    }

    /**
     * @param string $data
     * @param int $offset
     *
     * @return int
     */
    private static function uint32(string $data, int $offset): int
    {
        return unpack('V', substr($data, $offset, 4))[1];
    }
}
