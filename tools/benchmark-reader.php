<?php

/** Run each input in a fresh PHP process; reports PHP heap, not total process RSS. */
require __DIR__ . '/../vendor/autoload.php';

use avadim\FastExcelReader\Excel;

if (empty($argv[1]) || !is_file($argv[1])) {
    fwrite(STDERR, "Usage: php tools/benchmark-reader.php file [--string]\n");
    exit(1);
}

$directory = sys_get_temp_dir() . '/fxr_benchmark_' . bin2hex(random_bytes(8));
mkdir($directory);
Excel::setTempDir($directory);
$diskBytes = static function () use ($directory): int {
    $bytes = 0;
    foreach (glob($directory . '/*') as $file) {
        $bytes += filesize($file);
    }
    return $bytes;
};
$started = microtime(true);
$book = null;
try {
    $book = ($argv[2] ?? '') === '--string' ? Excel::openString(file_get_contents($argv[1])) : Excel::open($argv[1]);
    $result = ['php' => PHP_VERSION, 'file' => basename($argv[1]), 'reader' => get_class($book),
        'open_seconds' => microtime(true) - $started, 'open_heap_peak_bytes' => memory_get_peak_usage(true)];
    $temporaryPeak = $diskBytes();
    $started = microtime(true);
    $rows = $cells = 0;
    foreach ($book->sheets() as $sheet) {
        foreach ($sheet->nextRow() as $row) {
            $rows++;
            $cells += count($row);
            $temporaryPeak = max($temporaryPeak, $diskBytes());
        }
    }
    unset($row);
    $result += ['read_seconds' => microtime(true) - $started, 'rows' => $rows, 'cells' => $cells,
        'read_heap_peak_bytes' => memory_get_peak_usage(true), 'temporary_bytes_sampled_peak' => $temporaryPeak];
    $book->close();
    gc_collect_cycles();
    $result += ['closed_heap_bytes' => memory_get_usage(true), 'temporary_bytes_after_close' => $diskBytes()];
    echo json_encode($result, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES), "\n";
}
finally {
    if ($book) {
        $book->close();
    }
    Excel::setTempDir('');
    rmdir($directory);
}
