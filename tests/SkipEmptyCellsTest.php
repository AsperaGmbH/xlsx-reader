<?php

namespace Aspera\Spreadsheet\XLSX\Tests;

require_once __DIR__ . '/../vendor/autoload.php';

use Exception;
use PHPUnit\Framework\TestCase as PHPUnitTestCase;
use Aspera\Spreadsheet\XLSX\Worksheet;
use Aspera\Spreadsheet\XLSX\Reader as XLSXReader;

/**
 * Make sure the SkipEmptyCells option works how it should.
 *
 * @author Aspera GmbH
 */
class SkipEmptyCellsTest extends PHPUnitTestCase
{
    const FILE_PATH = 'input_files/iterator_test.xlsx';

    /**
     * Make sure that the SkipEmptyCells option is properly considered by the reader.
     *
     * @param bool $skip_empty_cells
     * @param int  $exp_num_cols
     *
     * @dataProvider dataProviderEmptyCells
     *
     * @throws Exception
     */
    public function testSkipEmptyCellsOption($skip_empty_cells, $exp_num_cols)
    {
        $options = array(
            'SkipEmptyCells' => $skip_empty_cells
        );

        $reader = new XLSXReader($options);
        $reader->open(self::FILE_PATH);

        $sheet_index = null;
        /** @var Worksheet $worksheet */
        foreach ($reader->getSheets() as $index => $worksheet) {
            if ($worksheet->getName() == 'EmptyCellsSheet') {
                $sheet_index = $index;
                break;
            }
        }
        self::assertNotNull($sheet_index, 'Could not locate worksheet with name "EmptyCellsSheet".');
        $reader->changeSheet($sheet_index);

        $current = $reader->current();
        $num_cols = count($current);

        self::assertEquals($exp_num_cols, $num_cols, 'Number of cells differ');

        $reader->close();
    }

    /**
     * @return array
     */
    public function dataProviderEmptyCells()
    {
        return array(
            array(
                'skipEmptyCells' => true,
                'numTotalCols'   => 5
            ),
            array(
                'skipEmptyCells' => false,
                'numTotalCols'   => 8
            )
        );
    }
}
