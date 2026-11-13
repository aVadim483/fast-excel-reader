# Reader resource measurements

[Русский](ru/23-benchmarks.md)

Run each input in a fresh PHP process:

```sh
php -d xdebug.mode=off tools/benchmark-reader.php tests/test_files/ods/family.ods --string
php -d xdebug.mode=off tools/benchmark-reader.php demo/files/demo-100k-rows.xlsx --string
```

`--string` includes the whole-file copy used by `openString()`. Omit it for direct
file opening. Results distinguish opening and reading, PHP heap peaks, sampled
temporary disk usage and disk usage after `close()`. PHP allocator memory is not
process RSS. Reading time includes disk sampling overhead; this is a diagnostic,
not a comparison with competing readers. The iterator does not retain returned rows.

Local measurements on 2026-11-13, Windows, PHP 8.4.19, Xdebug disabled, one fresh
process per input, `--string`:

| Input | Rows / cells | PHP heap peak | Temporary disk after close |
|---|---|---|---|
| Excel family.xlsx | 4 / 10 | 2 MiB | 0 bytes |
| Excel family.ods | 4 / 9 | 2 MiB | 0 bytes |
| BIFF8 demo-00-test.xls | 22 / 45 | 2 MiB | 0 bytes |
| test.csv | 3 / 9 | 2 MiB | 0 bytes |
| ODS: 1,000,000 empty repeated rows, then one value | 1 / 1 emitted | 2 MiB | 0 bytes |
| demo-100k-rows.xlsx | 100,000 / 500,000 | 6 MiB | 0 bytes |

The repeated ODS case uses 16,384 empty cells per repeated row; neither rows nor
cells are materialized. It is reproduced by OdsReaderTest's repetition case.
The two family fixtures differ by one empty covered cell omitted by Excel in ODS.
Other inputs are different datasets: these numbers do not establish relative speed
or memory efficiency between formats. Shared strings, images and selected row width
can dominate other workbooks; a disk-backed shared-string cache is not implemented.
