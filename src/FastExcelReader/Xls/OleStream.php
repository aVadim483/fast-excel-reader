<?php

namespace avadim\FastExcelReader\Xls;

/**
 * A seekable byte stream stored inside an OLE2 compound file
 *
 * Large streams are read sector by sector straight from the file handle: only
 * the sector currently being read is held in memory, whatever the stream size.
 * Small streams live in the mini stream and are at most a few kilobytes, so
 * they are simply loaded once.
 */
class OleStream
{
    /** @var resource|null */
    private $handle;

    private OleReader $ole;

    private int $size;

    private int $position = 0;

    /** First sector of the chain, or null for a stream held in memory */
    private ?int $startSector;

    private int $sectorSize;

    /** Whole content, for mini streams only */
    private ?string $content;

    /** Number of the sector currently held in $buffer */
    private int $bufferSector = -1;

    private string $buffer = '';

    /** Sector chain resolved so far: index in chain => sector number */
    private array $chain = [];

    /**
     * @param OleReader $ole
     * @param resource|null $handle
     * @param int|null $startSector
     * @param int $size
     * @param int $sectorSize
     * @param string|null $content
     */
    public function __construct(OleReader $ole, $handle, ?int $startSector, int $size, int $sectorSize, ?string $content = null)
    {
        $this->ole = $ole;
        $this->handle = $handle;
        $this->startSector = $startSector;
        $this->size = $size;
        $this->sectorSize = $sectorSize;
        $this->content = $content;

        if ($startSector !== null) {
            $this->chain[0] = $startSector;
        }
    }

    /**
     * @return int
     */
    public function size(): int
    {
        return $this->size;
    }

    /**
     * @return int
     */
    public function tell(): int
    {
        return $this->position;
    }

    /**
     * @return bool
     */
    public function eof(): bool
    {
        return $this->position >= $this->size;
    }

    /**
     * @param int $position
     *
     * @return void
     */
    public function seek(int $position): void
    {
        $this->position = max(0, $position);
    }

    /**
     * Read $length bytes from the current position
     *
     * @param int $length
     *
     * @return string Fewer bytes than requested at the end of the stream
     */
    public function read(int $length): string
    {
        if ($length <= 0 || $this->position >= $this->size) {
            return '';
        }
        if ($this->position + $length > $this->size) {
            $length = $this->size - $this->position;
        }

        if ($this->content !== null) {
            $result = substr($this->content, $this->position, $length);
            $this->position += $length;

            return $result;
        }

        $result = '';
        while ($length > 0) {
            $chainIndex = intdiv($this->position, $this->sectorSize);
            $offset = $this->position % $this->sectorSize;
            $take = min($length, $this->sectorSize - $offset);

            $sector = $this->sectorAt($chainIndex);
            if ($sector === null) {
                break;
            }
            if ($sector !== $this->bufferSector) {
                fseek($this->handle, $this->ole->sectorOffset($sector));
                $this->buffer = (string)fread($this->handle, $this->sectorSize);
                $this->bufferSector = $sector;
            }

            $result .= substr($this->buffer, $offset, $take);
            $this->position += $take;
            $length -= $take;
        }

        return $result;
    }

    /**
     * Resolve the n-th sector of the chain, extending the resolved part as needed
     *
     * The chain is walked lazily and only forwards, so seeking near the end of a
     * large stream costs one FAT lookup per sector - a few array writes, not a
     * copy of the data.
     *
     * @param int $chainIndex
     *
     * @return int|null
     */
    private function sectorAt(int $chainIndex): ?int
    {
        if (isset($this->chain[$chainIndex])) {
            return $this->chain[$chainIndex];
        }
        $last = count($this->chain) - 1;
        if ($last < 0) {
            return null;
        }
        $sector = $this->chain[$last];
        for ($i = $last; $i < $chainIndex; $i++) {
            $sector = $this->ole->nextSector($sector);
            if ($sector === null) {
                return null;
            }
            $this->chain[$i + 1] = $sector;
        }

        return $sector;
    }
}
