<?php

namespace avadim\FastExcelReader\Internal;

use avadim\FastExcelReader\Exception;

/** @internal Resolves internal relationship targets within a ZIP package. */
final class PackagePath
{
    public static function resolve(string $source, string $target): string
    {
        if ($target === '' || preg_match('~^[a-z][a-z0-9+.-]*:|^//|[\\\\?#\x00]~i', $target)) {
            throw new Exception('Invalid internal relationship target: ' . $target);
        }
        $path = $target[0] === '/' ? substr($target, 1) : dirname($source) . '/' . $target;
        $parts = [];
        foreach (explode('/', $path) as $part) {
            if ($part === '' || $part === '.') {
                continue;
            }
            if ($part === '..') {
                if (!$parts) {
                    throw new Exception('Relationship target escapes the package: ' . $target);
                }
                array_pop($parts);
            }
            else {
                $parts[] = $part;
            }
        }
        if (!$parts) {
            throw new Exception('Empty internal relationship target: ' . $target);
        }

        return implode('/', $parts);
    }
}
