<?php

namespace Aspera\Spreadsheet\XLSX\Tests;

require_once __DIR__ . '/../vendor/autoload.php';

use Exception;
use PHPUnit\Framework\TestCase;
use Aspera\Spreadsheet\XLSX\Reader;
use Aspera\Spreadsheet\XLSX\ReaderConfiguration;

class CellContentTest extends TestCase
{
    const TEST_FILE = __DIR__ . '/input_files/cell_content_test.xlsx';

    /**
     * Check if potentially problematic values are read correctly.
     * Also check if addressing columns by column name works properly.
     *
     * @throws Exception
     */
    public function testCellContent()
    {
        $reader = new Reader((new ReaderConfiguration())
            ->setOutputColumnNames(true)
        );
        $reader->open(self::TEST_FILE);
        $output_cells = array();
        foreach ($reader as $row) {
            $output_cells[] = $row['B']; // Only the second column ("Value") is of interest.
        }
        $reader->close();

        self::assertSame(
            array('', 'Value', '0123', '0123', '   ', '268.02'),
            $output_cells,
            'The retrieved sheet content was not as expected.'
        );
    }
}
