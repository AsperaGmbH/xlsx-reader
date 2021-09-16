<?php

namespace Aspera\Spreadsheet\XLSX\Tests;

require_once __DIR__ . '/../vendor/autoload.php';

use Aspera\Spreadsheet\XLSX\ReaderConfiguration;
use Aspera\Spreadsheet\XLSX\ReaderSkipConfiguration;
use Exception;
use PHPUnit\Framework\TestCase;
use Aspera\Spreadsheet\XLSX\Reader;

class EmptyRowsTest extends TestCase
{
    const TEST_FILE = __DIR__ . '/input_files/empty_rows_test.xlsx';

    /**
     * Check if empty rows are detected and reported correctly.
     * Includes test of self-closing row elements caused by e.g. the usage of thick borders in adjacent cells.
     *
     * @dataProvider dataProviderCellContent
     *
     * @param int   $skip_config
     * @param array $expected_values
     *
     * @throws Exception
     */
    public function testCellContent($skip_config, $expected_values)
    {
        $reader = new Reader(
            (new ReaderConfiguration())
                ->setSkipEmptyRows($skip_config)
        );
        $reader->open(self::TEST_FILE);
        $output_cells = array();
        foreach ($reader as $row) {
            $output_cells[] = $row[1]; // All values to check for are in the 2nd column. Ignore all other columns.
        }
        $reader->close();

        self::assertSame(
            $expected_values,
            $output_cells,
            'The retrieved sheet content was not as expected.'
        );
    }

    /**
     * @return array
     */
    public function dataProviderCellContent()
    {
        return array(
            'SKIP_NONE'           => array(
                ReaderSkipConfiguration::SKIP_NONE,
                array('', 'row 2', '', '', 'row 5', '', '', '', '')
            ),
            'SKIP_EMPTY'      => array(
                ReaderSkipConfiguration::SKIP_EMPTY,
                array('row 2', 'row 5')
            ),
            'SKIP_TRAILING_EMPTY' => array(
                ReaderSkipConfiguration::SKIP_TRAILING_EMPTY,
                array('', 'row 2', '', '', 'row 5')
            )
        );
    }
}
