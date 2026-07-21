<?php

namespace avadim\FastExcelReader\Xls;

/**
 * Sequential reader over a BIFF record stream
 *
 * BIFF is a flat sequence of [type:u16][length:u16][data:length] records, which
 * is what makes .xls readable in a single forward pass with constant memory:
 * only one record is held at a time.
 *
 * Records longer than 8224 bytes are split, with the tail carried by CONTINUE
 * records. Those are merged transparently, but the segment boundaries are kept
 * as well - the SST needs them, because a string split across a boundary
 * restarts with a fresh encoding flag.
 */
class BiffReader
{
    private OleStream $stream;

    /** Absolute offset of the record returned last */
    private int $recordOffset = 0;

    public function __construct(OleStream $stream)
    {
        $this->stream = $stream;
    }

    /**
     * Position the reader at an absolute offset within the stream
     *
     * BOUNDSHEET records store exactly such an offset for the BOF of each
     * sheet, so a single sheet can be reached without reading the ones before.
     *
     * @param int $offset
     *
     * @return void
     */
    public function seek(int $offset): void
    {
        $this->stream->seek($offset);
    }

    /**
     * @return int
     */
    public function tell(): int
    {
        return $this->stream->tell();
    }

    /**
     * Offset of the record returned by the last nextRecord() call
     *
     * @return int
     */
    public function recordOffset(): int
    {
        return $this->recordOffset;
    }

    /**
     * Read the next record, merging any CONTINUE records that follow it
     *
     * @return array|null ['type' => int, 'data' => string, 'parts' => string[], 'offset' => int], NULL at the end
     */
    public function nextRecord(): ?array
    {
        $this->recordOffset = $this->stream->tell();
        $header = $this->stream->read(4);
        if (strlen($header) < 4) {
            return null;
        }
        $head = unpack('vtype/vlength', $header);
        $type = $head['type'];
        $length = $head['length'];
        $data = $length > 0 ? $this->stream->read($length) : '';

        $parts = [$data];
        // CONTINUE records are never meaningful on their own
        while (($continued = $this->readContinue()) !== null) {
            $parts[] = $continued;
            $data .= $continued;
        }

        return [
            'type' => $type,
            'data' => $data,
            'parts' => $parts,
            'offset' => $this->recordOffset,
        ];
    }

    /**
     * Iterate records from the current position
     *
     * @param int|null $stopAtType Stop after yielding a record of this type
     *
     * @return \Generator
     */
    public function records(?int $stopAtType = null): \Generator
    {
        while (($record = $this->nextRecord()) !== null) {
            yield $record;
            if ($stopAtType !== null && $record['type'] === $stopAtType) {
                break;
            }
        }
    }

    /**
     * Consume the next record if it is a CONTINUE, otherwise rewind
     *
     * @return string|null
     */
    private function readContinue(): ?string
    {
        $position = $this->stream->tell();
        $header = $this->stream->read(4);
        if (strlen($header) < 4) {
            $this->stream->seek($position);

            return null;
        }
        $head = unpack('vtype/vlength', $header);
        $type = $head['type'];
        $length = $head['length'];
        if ($type !== BiffRecord::CONTINUE) {
            $this->stream->seek($position);

            return null;
        }

        return $length > 0 ? $this->stream->read($length) : '';
    }
}
