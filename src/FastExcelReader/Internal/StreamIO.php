<?php

namespace avadim\FastExcelReader\Internal;

use avadim\FastExcelReader\Exception;

/** @internal Checked writes; the caller owns the destination file and its removal. */
final class StreamIO
{
    public static function writeString(string $file, string $content): void
    {
        self::withOutput($file, static function ($out) use ($content): int {
            $bytes = @fwrite($out, $content);
            if ($bytes !== strlen($content)) {
                throw new Exception('Cannot write the complete spreadsheet content to a temporary file');
            }
            return $bytes;
        });
    }

    /** @param resource $stream Caller-owned input, read from its current position. */
    public static function copyToFile($stream, string $file): int
    {
        return self::withOutput($file, static function ($out) use ($stream): int {
            $bytes = @stream_copy_to_stream($stream, $out);
            if ($bytes === false) {
                throw new Exception('Cannot copy the complete stream to a temporary file');
            }
            return $bytes;
        });
    }

    private static function withOutput(string $file, callable $write): int
    {
        $out = null;
        try {
            $out = @fopen($file, 'wb');
            if (!$out) {
                throw new Exception('Cannot open a temporary output file');
            }
            $bytes = $write($out);
            if (!@fflush($out)) {
                throw new Exception('Cannot flush the temporary output stream');
            }
            if (!@fclose($out)) {
                throw new Exception('Cannot close the temporary output stream');
            }
            $out = null;

            return $bytes;
        }
        catch (\Throwable $e) {
            throw new Exception('Cannot write the spreadsheet to a temporary file', 0, $e);
        }
        finally {
            if (is_resource($out)) {
                try {
                    @fclose($out);
                }
                catch (\Throwable $ignored) {
                    // Keep the original write/flush failure as the exception cause.
                }
            }
        }
    }
}
