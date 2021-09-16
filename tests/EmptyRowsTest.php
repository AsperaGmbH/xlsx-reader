<?php

namespace Aspera\Spreadsheet\XLSX\Tests;

require_once __DIR__ . '/../vendor/autoload.php';

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
     * @throws Exception
     */
    public function testCellContent()
    {
        $reader = new Reader();
        $reader->open(self::TEST_FILE);
        $output_cells = array();
        foreach ($reader as $row) {
            $output_cells[] = $row[1]; // All values to check for are in the 2nd column. Ignore all other columns.
        }
        $reader->close();

        self::assertSame(
            array('', 'row 2', '', '', 'row 5'),
            $output_cells,
            'The retrieved sheet content was not as expected.'
        );
    }
}
