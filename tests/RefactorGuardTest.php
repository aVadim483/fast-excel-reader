<?php

declare(strict_types=1);

namespace avadim\FastExcelReader\Tests;

use avadim\FastExcelReader\Excel;
use avadim\FastExcelReader\Tests\Support\GuardTestCase;

/**
 * Characterization ("golden master") tests.
 *
 * Purpose: prove that a refactoring did not change observable behaviour. Every
 * scenario runs a public read API against a fixture and compares the result -
 * whole array, exact types - against a stored snapshot.
 *
 * These assertions describe what the library DOES today, not what it SHOULD do.
 * Known quirks are recorded deliberately and marked with @todo.
 */
final class RefactorGuardTest extends GuardTestCase
{
    /**
     * @dataProvider scenarioProvider
     *
     * @param string $slug
     * @param callable $scenario
     *
     * @return void
     */
    public function testScenarioMatchesSnapshot(string $slug, callable $scenario): void
    {
        $this->assertMatchesSnapshot($slug, $scenario());
    }

    /**
     * @return array<string, array{0: string, 1: callable}>
     */
    public function scenarioProvider(): array
    {
        $scenarios = array_merge(
            self::keyModeScenarios('demo-00-test.xlsx', 'keys-origin'),
            // Same matrix on a sheet whose data starts at B2, not A1: this is
            // where the lazily computed $rowOffset/$colOffset actually differ
            self::keyModeScenarios('demo-02-advanced.xlsx', 'keys-offset'),
            self::columnKeyScenarios(),
            self::resultModeScenarios(),
            self::readAreaScenarios(),
            self::cellAndColumnScenarios(),
            self::styleScenarios(),
            self::dateScenarios(),
            self::metadataScenarios(),
            self::degenerateScenarios()
        );

        $data = [];
        foreach ($scenarios as $slug => $scenario) {
            $data[$slug] = [$slug, $scenario];
        }

        return $data;
    }

    /**
     * Every KEYS_* mode against one fixture
     *
     * @param string $file
     * @param string $prefix
     *
     * @return array<string, callable>
     */
    private static function keyModeScenarios(string $file, string $prefix): array
    {
        $modes = [
            'default' => null,
            'original' => Excel::KEYS_ORIGINAL,
            'first-row' => Excel::KEYS_FIRST_ROW,
            'row-zero-based' => Excel::KEYS_ROW_ZERO_BASED,
            'row-one-based' => Excel::KEYS_ROW_ONE_BASED,
            'col-zero-based' => Excel::KEYS_COL_ZERO_BASED,
            'col-one-based' => Excel::KEYS_COL_ONE_BASED,
            'zero-based' => Excel::KEYS_ZERO_BASED,
            'one-based' => Excel::KEYS_ONE_BASED,
            'relative' => Excel::KEYS_RELATIVE,
            'swap' => Excel::KEYS_SWAP,
            'first-row-swap' => Excel::KEYS_FIRST_ROW | Excel::KEYS_SWAP,
            'first-row-row-zero-based' => Excel::KEYS_FIRST_ROW | Excel::KEYS_ROW_ZERO_BASED,
            'relative-first-row' => Excel::KEYS_RELATIVE | Excel::KEYS_FIRST_ROW,
        ];

        $scenarios = [];
        foreach ($modes as $name => $mode) {
            $scenarios[$prefix . '--rows-' . $name] = static function () use ($file, $mode) {
                return Excel::open(self::fixture($file))->readRows(false, $mode);
            };
        }

        // $columnKeys given as a bool interacts with the mode flags separately
        foreach (['true' => true, 'null' => null] as $name => $columnKeys) {
            $scenarios[$prefix . '--rows-colkeys-' . $name] = static function () use ($file, $columnKeys) {
                return Excel::open(self::fixture($file))->readRows($columnKeys);
            };
        }

        return $scenarios;
    }

    /**
     * Explicit column key maps, alone and combined with KEYS_FIRST_ROW
     *
     * @return array<string, callable>
     */
    private static function columnKeyScenarios(): array
    {
        $keys = ['A' => 'Number', 'B' => 'Hero', 'D' => 'Secret'];

        return [
            'colkeys--plain' => static function () use ($keys) {
                return Excel::open(self::fixture('demo-00-test.xlsx'))->readRows($keys);
            },
            'colkeys--first-row' => static function () use ($keys) {
                return Excel::open(self::fixture('demo-00-test.xlsx'))->readRows($keys, Excel::KEYS_FIRST_ROW);
            },
            'colkeys--first-row-row-zero-based' => static function () use ($keys) {
                return Excel::open(self::fixture('demo-00-test.xlsx'))
                    ->readRows($keys, Excel::KEYS_FIRST_ROW | Excel::KEYS_ROW_ZERO_BASED);
            },
            'colkeys--row-zero-based' => static function () use ($keys) {
                return Excel::open(self::fixture('demo-00-test.xlsx'))->readRows($keys, Excel::KEYS_ROW_ZERO_BASED);
            },
            'colkeys--partial' => static function () {
                return Excel::open(self::fixture('demo-02-advanced.xlsx'))->readRows(['B' => 'first', 'D' => 'third']);
            },
            'colkeys--list-not-map' => static function () {
                return Excel::open(self::fixture('demo-00-test.xlsx'))->readRows([], Excel::KEYS_FIRST_ROW);
            },
        ];
    }

    /**
     * RESULT_MODE_ROW, TRIM_STRINGS, TREAT_EMPTY_STRING_AS_EMPTY_CELL -
     * none of these had any test coverage before
     *
     * @return array<string, callable>
     */
    private static function resultModeScenarios(): array
    {
        return [
            'result-mode--row' => static function () {
                $sheet = Excel::open(self::fixture('demo-00-test.xlsx'))->sheet();

                return iterator_to_array($sheet->nextRow([], Excel::RESULT_MODE_ROW));
            },
            'result-mode--row-with-keys' => static function () {
                $sheet = Excel::open(self::fixture('demo-02-advanced.xlsx'))->sheet();

                return iterator_to_array($sheet->nextRow([], Excel::RESULT_MODE_ROW | Excel::KEYS_ONE_BASED));
            },
            'result-mode--row-read-rows' => static function () {
                // readCallback() unwraps __cells, so readRows() should be unaffected
                return Excel::open(self::fixture('demo-00-test.xlsx'))->readRows([], Excel::RESULT_MODE_ROW);
            },
            'result-mode--trim-strings' => static function () {
                return Excel::open(self::fixture('demo-00-test.xlsx'))->readRows([], Excel::TRIM_STRINGS);
            },
            'result-mode--empty-string-as-empty-cell' => static function () {
                return Excel::open(self::fixture('empty.xlsx'))->readRows([], Excel::TREAT_EMPTY_STRING_AS_EMPTY_CELL);
            },
            'result-mode--empty-string-kept' => static function () {
                return Excel::open(self::fixture('empty.xlsx'))->readRows();
            },
            'result-mode--empty-string-cells' => static function () {
                return Excel::open(self::fixture('empty.xlsx'))->readCells();
            },
        ];
    }

    /**
     * Read areas, column ranges, header handling and every *From() variant
     *
     * @return array<string, callable>
     */
    private static function readAreaScenarios(): array
    {
        return [
            'area--set-read-area' => static function () {
                return Excel::open(self::fixture('demo-01-base.xlsx'))->setReadArea('B2:D6')->readRows();
            },
            'area--set-read-area-first-row-keys' => static function () {
                return Excel::open(self::fixture('demo-01-base.xlsx'))->setReadArea('A1:D6', true)->readRows();
            },
            'area--set-read-area-columns' => static function () {
                $sheet = Excel::open(self::fixture('demo-01-base.xlsx'))->sheet();

                return $sheet->setReadAreaColumns('B:D')->readRows(false, Excel::KEYS_ROW_ONE_BASED);
            },
            'area--set-read-area-columns-first-row-keys' => static function () {
                $sheet = Excel::open(self::fixture('demo-01-base.xlsx'))->sheet();

                return $sheet->setReadAreaColumns('B:D', true)->readRows();
            },
            'area--from' => static function () {
                return Excel::open(self::fixture('demo-02-advanced.xlsx'))->from('C3')->readRows();
            },
            'area--from-first-row-keys' => static function () {
                return Excel::open(self::fixture('demo-02-advanced.xlsx'))->sheet()->from('B2', true)->readRows();
            },
            'area--with-header' => static function () {
                return Excel::open(self::fixture('demo-00-test.xlsx'))->sheet()->withHeader()->readRows();
            },
            'area--read-rows-from' => static function () {
                return Excel::open(self::fixture('demo-01-base.xlsx'))->sheet()->readRowsFrom('B2:D6');
            },
            'area--read-cells-from' => static function () {
                return Excel::open(self::fixture('demo-01-base.xlsx'))->sheet()->readCellsFrom('B2:D6');
            },
            'area--read-columns-from' => static function () {
                return Excel::open(self::fixture('demo-01-base.xlsx'))->sheet()->readColumnsFrom('B2:D6');
            },
            'area--read-first-row-from' => static function () {
                return Excel::open(self::fixture('demo-01-base.xlsx'))->sheet()->readFirstRowFrom('B2:D6');
            },
            // readFirstRowCellsFrom() cannot be snapshotted: it always throws.
            // See KnownBugsTest::testReadFirstRowCellsFromIsBroken()
            'area--read-first-row-cells-after-area' => static function () {
                return Excel::open(self::fixture('demo-01-base.xlsx'))->sheet()->setReadArea('B2:D6')->readFirstRowCells();
            },
            'area--read-rows-from-then-plain' => static function () {
                // setReadArea() mutates the sheet, so the second call must stay narrowed
                $sheet = Excel::open(self::fixture('demo-01-base.xlsx'))->sheet();
                $first = $sheet->readRowsFrom('B2:C4');

                return ['from' => $first, 'after' => $sheet->readRows()];
            },
            'area--area-then-key-mode' => static function () {
                return Excel::open(self::fixture('demo-02-advanced.xlsx'))
                    ->setReadArea('C3:D8')
                    ->readRows(false, Excel::KEYS_ZERO_BASED);
            },
            'area--area-relative' => static function () {
                return Excel::open(self::fixture('demo-02-advanced.xlsx'))
                    ->setReadArea('C3:D8')
                    ->readRows(false, Excel::KEYS_RELATIVE);
            },
        ];
    }

    /**
     * readCells / readColumns families, including style-index payloads
     *
     * @return array<string, callable>
     */
    private static function cellAndColumnScenarios(): array
    {
        return [
            'cells--plain' => static function () {
                return Excel::open(self::fixture('demo-00-test.xlsx'))->readCells();
            },
            'cells--style-idx' => static function () {
                return Excel::open(self::fixture('demo-00-test.xlsx'))->sheet()->readCells(true);
            },
            'cells--offset-sheet' => static function () {
                return Excel::open(self::fixture('demo-02-advanced.xlsx'))->readCells();
            },
            'columns--plain' => static function () {
                return Excel::open(self::fixture('demo-00-test.xlsx'))->readColumns();
            },
            'columns--one-based' => static function () {
                return Excel::open(self::fixture('demo-00-test.xlsx'))->readColumns(null, Excel::KEYS_ONE_BASED);
            },
            'columns--offset-sheet' => static function () {
                return Excel::open(self::fixture('demo-02-advanced.xlsx'))->readColumns();
            },
            'rows--style-idx' => static function () {
                return Excel::open(self::fixture('demo-00-test.xlsx'))->sheet()->readRows([], null, true);
            },
            'rows--with-styles' => static function () {
                return Excel::open(self::fixture('demo-00-test.xlsx'))->readRowsWithStyles();
            },
            'rows--with-styles-from' => static function () {
                return Excel::open(self::fixture('demo-02-advanced.xlsx'))->sheet()->readRowsWithStylesFrom('B2:D5');
            },
            'columns--with-styles' => static function () {
                return Excel::open(self::fixture('demo-00-test.xlsx'))->readColumnsWithStyles();
            },
            'columns--with-styles-from' => static function () {
                return Excel::open(self::fixture('demo-02-advanced.xlsx'))->sheet()->readColumnsWithStylesFrom('B2:D5');
            },
            'first-row--plain' => static function () {
                return Excel::open(self::fixture('demo-00-test.xlsx'))->sheet()->readFirstRow();
            },
            'first-row--cells' => static function () {
                return Excel::open(self::fixture('demo-00-test.xlsx'))->sheet()->readFirstRowCells();
            },
            'first-row--with-styles' => static function () {
                return Excel::open(self::fixture('demo-00-test.xlsx'))->sheet()->readFirstRowWithStyles();
            },
            'first-row--with-styles-from' => static function () {
                return Excel::open(self::fixture('demo-02-advanced.xlsx'))->sheet()->readFirstRowWithStylesFrom('B2:D5');
            },
            'first-row--offset-sheet' => static function () {
                return Excel::open(self::fixture('demo-02-advanced.xlsx'))->sheet()->readFirstRow();
            },
        ];
    }

    /**
     * Style payloads. Areas are kept small on purpose - these snapshots are the
     * bulkiest ones and a whole styled sheet adds noise without adding signal.
     *
     * @return array<string, callable>
     */
    private static function styleScenarios(): array
    {
        return [
            'styles--cells-with-styles' => static function () {
                return Excel::open(self::fixture('demo-04-styles.xlsx'))->sheet()->readCellsWithStylesFrom('A1:E8');
            },
            'styles--cells-with-styles-key' => static function () {
                return Excel::open(self::fixture('demo-04-styles.xlsx'))
                    ->sheet()->setReadArea('A1:E8')->readCellsWithStyles('fill-color');
            },
            'styles--cell-styles-nested' => static function () {
                return Excel::open(self::fixture('demo-04-styles.xlsx'))->sheet()->setReadArea('A1:E8')->readCellStyles();
            },
            'styles--cell-styles-flat' => static function () {
                return Excel::open(self::fixture('demo-04-styles.xlsx'))->sheet()->setReadArea('A1:E8')->readCellStyles(true);
            },
            'styles--read-styles' => static function () {
                return Excel::open(self::fixture('demo-00-test.xlsx'))->readStyles();
            },
            'styles--complete-by-idx' => static function () {
                $excel = Excel::open(self::fixture('demo-04-styles.xlsx'));
                $result = [];
                foreach ([0, 1, 2, 3, 5, 8] as $idx) {
                    $result[$idx] = [
                        'nested' => $excel->getCompleteStyleByIdx($idx),
                        'flat' => $excel->getCompleteStyleByIdx($idx, true),
                        'pattern' => $excel->getFormatPattern($idx),
                    ];
                }

                return $result;
            },
        ];
    }

    /**
     * Date and number typing - the part most likely to break silently
     *
     * @return array<string, callable>
     */
    private static function dateScenarios(): array
    {
        return [
            'dates--timestamps' => static function () {
                return Excel::open(self::fixture('demo-05-datetime.xlsx'))->readRows();
            },
            'dates--formatted' => static function () {
                $excel = Excel::open(self::fixture('demo-05-datetime.xlsx'));
                $excel->setDateFormat('Y-m-d H:i:s');

                return $excel->readRows();
            },
            'dates--formatter-false' => static function () {
                $excel = Excel::open(self::fixture('demo-05-datetime.xlsx'));
                $excel->dateFormatter(false);

                return $excel->readRows();
            },
            'dates--formatter-null' => static function () {
                $excel = Excel::open(self::fixture('demo-05-datetime.xlsx'));
                $excel->dateFormatter(null);

                return $excel->readRows();
            },
            'dates--cells-typed' => static function () {
                return Excel::open(self::fixture('demo-05-datetime.xlsx'))->sheet()->readCells(true);
            },
            'dates--sheet-level-format' => static function () {
                $sheet = Excel::open(self::fixture('demo-05-datetime.xlsx'))->sheet();

                return $sheet->setDateFormat('d/m/Y')->readRows();
            },
        ];
    }

    /**
     * Sheet and workbook metadata accessors
     *
     * @return array<string, callable>
     */
    private static function metadataScenarios(): array
    {
        $collect = static function (string $file): array {
            $excel = Excel::open(self::fixture($file));
            $result = [
                'sheetNames' => $excel->getSheetNames(),
                'countSheets' => $excel->countSheets(),
                'visibleSheets' => array_keys($excel->visibleSheets()),
                'hiddenSheets' => array_keys($excel->hiddenSheets()),
                'definedNames' => $excel->getDefinedNames(),
                'sheetExistsHit' => $excel->sheetExists($excel->getSheetNames()[0] ?? ''),
                'sheetExistsMiss' => $excel->sheetExists('no-such-sheet-xyz'),
            ];

            foreach ($excel->sheets() as $name => $sheet) {
                $result['sheets'][$name] = [
                    'id' => $sheet->id(),
                    'name' => $sheet->name(),
                    'path' => $sheet->path(),
                    'state' => $sheet->state(),
                    'isVisible' => $sheet->isVisible(),
                    'isHidden' => $sheet->isHidden(),
                    'dimension' => $sheet->dimension(),
                    'dimensionArray' => $sheet->dimensionArray(),
                    'countRows' => $sheet->countRows(),
                    'countCols' => $sheet->countCols(),
                    'countColumns' => $sheet->countColumns(),
                    'minRow' => $sheet->minRow(),
                    'maxRow' => $sheet->maxRow(),
                    'minColumn' => $sheet->minColumn(),
                    'maxColumn' => $sheet->maxColumn(),
                    'firstRow' => $sheet->firstRow(),
                    'firstCol' => $sheet->firstCol(),
                    'mergedCells' => $sheet->getMergedCells(),
                ];
            }

            return $result;
        };

        $scenarios = [];
        foreach (['demo-00-test.xlsx', 'demo-02-advanced.xlsx', 'demo-04-styles.xlsx', 'demo-07-size-freeze-tabs.xlsx'] as $file) {
            $slug = 'meta--' . pathinfo($file, PATHINFO_FILENAME);
            $scenarios[$slug] = static function () use ($collect, $file) {
                return $collect($file);
            };
        }

        $scenarios['meta--merged-lookup'] = static function () {
            $sheet = Excel::open(self::fixture('demo-02-advanced.xlsx'))->sheet();
            $result = [];
            foreach (['B2', 'C3', 'B4', 'D11', 'A1'] as $cell) {
                $result[$cell] = [
                    'isMerged' => $sheet->isMerged($cell),
                    'mergedRange' => $sheet->mergedRange($cell),
                ];
            }

            return $result;
        };

        $scenarios['meta--actual-dimension'] = static function () {
            $sheet = Excel::open(self::fixture('wrong-dimension.xlsx'))->sheet();

            return [
                'dimension' => $sheet->dimension(),
                'actualDimension' => $sheet->actualDimension(),
                'countActualRows' => $sheet->countActualRows(),
                'countActualColumns' => $sheet->countActualColumns(),
                'minActualRow' => $sheet->minActualRow(),
                'maxActualRow' => $sheet->maxActualRow(),
                'minActualColumn' => $sheet->minActualColumn(),
                'maxActualColumn' => $sheet->maxActualColumn(),
                'stat' => $sheet->stat(),
            ];
        };

        return $scenarios;
    }

    /**
     * Empty, malformed and otherwise degenerate inputs
     *
     * @return array<string, callable>
     */
    private static function degenerateScenarios(): array
    {
        return [
            'degenerate--empty-rows' => static function () {
                return Excel::open(self::fixture('empty.xlsx'))->readRows();
            },
            'degenerate--empty-first-row' => static function () {
                return Excel::open(self::fixture('empty.xlsx'))->sheet()->readFirstRow();
            },
            'degenerate--wrong-dimension-rows' => static function () {
                return Excel::open(self::fixture('wrong-dimension.xlsx'))->readRows();
            },
            'degenerate--wrong-dimension-cells' => static function () {
                return Excel::open(self::fixture('wrong-dimension.xlsx'))->readCells();
            },
            'degenerate--no-dimension-tag' => static function () {
                // demo-06 has no <dimension> element at all
                $excel = Excel::open(self::fixture('demo-06-data-validation.xlsx'));

                return [
                    'dimension' => $excel->sheet()->dimension(),
                    'rows' => $excel->readRows(),
                ];
            },
            'degenerate--area-outside-data' => static function () {
                return Excel::open(self::fixture('demo-00-test.xlsx'))->setReadArea('Z100:AA200')->readRows();
            },
            'degenerate--single-cell-area' => static function () {
                return Excel::open(self::fixture('demo-00-test.xlsx'))->setReadArea('B2')->readRows();
            },
            'degenerate--nonstandard-file' => static function () {
                // This fixture writes namespace-prefixed tags (<x:row>, <x:c>),
                // which nextRow() does not match, so reading yields nothing at
                // all - silently. Pinned so a refactoring cannot change it
                // unnoticed in either direction.
                $excel = Excel::open(self::fixture('nonstandard-file.xlsx'));

                return [
                    'sheetNames' => $excel->getSheetNames(),
                    'dimension' => $excel->sheet()->dimension(),
                    'rows' => $excel->readRows(),
                    'cells' => $excel->readCells(),
                ];
            },
            'degenerate--standard-file-counterpart' => static function () {
                // Same content as nonstandard-file.xlsx, written without prefixes
                $excel = Excel::open(self::fixture('standard-file.xlsx'));

                return [
                    'sheetNames' => $excel->getSheetNames(),
                    'dimension' => $excel->sheet()->dimension(),
                    'rows' => $excel->readRows(),
                ];
            },
            'degenerate--absolute-path-worksheet' => static function () {
                return Excel::open(self::fixture('worksheet-referenced-with-absolute-path.xlsx'))->readRows();
            },
        ];
    }
}
