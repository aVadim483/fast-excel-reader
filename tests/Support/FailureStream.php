<?php

namespace avadim\FastExcelReader\Tests\Support;

/** Controlled, non-seekable input and output failures without disk quotas or network. */
final class FailureStream
{
    public $context;
    public static array $closed = [];
    public static array $written = [];
    private string $mode;
    private int $position = 0;
    private string $data = "a,b\n1,2\n";

    public function stream_open($path, $mode, $options, &$openedPath): bool
    {
        $this->mode = substr($path, strlen('fxrfailure://'));
        return $this->mode !== 'deny-open';
    }

    public function stream_read($count)
    {
        if ($this->position >= strlen($this->data)) {
            if ($this->mode === 'throw-read') {
                throw new \RuntimeException('Injected read failure');
            }
            return $this->mode === 'fail-read' ? false : '';
        }
        $chunk = substr($this->data, $this->position, $count);
        $this->position += strlen($chunk);
        return $chunk;
    }

    public function stream_eof(): bool
    {
        return $this->mode === 'valid' && $this->position >= strlen($this->data);
    }

    public function stream_write($data): int
    {
        if ($this->mode === 'throw-write') {
            throw new \RuntimeException('Injected write failure');
        }
        $bytes = $this->mode === 'short-write' ? min(strlen($data), max(0, 3 - $this->position)) : strlen($data);
        $this->position += $bytes;
        self::$written[$this->mode] = (self::$written[$this->mode] ?? '') . substr($data, 0, $bytes);
        return $bytes;
    }

    public function stream_flush(): bool
    {
        if ($this->mode === 'throw-flush') {
            throw new \RuntimeException('Injected flush failure');
        }
        return $this->mode !== 'fail-flush';
    }

    public function stream_stat(): array
    {
        return [];
    }

    public function stream_close(): void
    {
        self::$closed[] = $this->mode;
        if ($this->mode === 'throw-close') {
            throw new \RuntimeException('Injected close failure');
        }
    }
}
