<?php

namespace Aspera\Spreadsheet\XLSX\Tests;

require_once __DIR__ . '/../vendor/autoload.php';

use Exception;
use PHPUnit\Framework\TestCase as PHPUnitTestCase;
use Aspera\Spreadsheet\XLSX\Worksheet;
use Aspera\Spreadsheet\XLSX\Reader;
use Aspera\Spreadsheet\XLSX\ReaderConfiguration;
use Aspera\Spreadsheet\XLSX\ReaderSkipConfiguration;

/** Make sure the SkipEmptyCells option works how it should. */
class SkipEmptyCellsTest extends PHPUnitTestCase
{
    const FILE_PATH = __DIR__ . '/input_files/iterator_test.xlsx';

    /**
     * Make sure that the SkipEmptyCells option is properly considered by the reader.
     *
     * @param int $skip_empty_cells
     * @param int $exp_num_cols
     *
     * @dataProvider dataProviderEmptyCells
     *
     * @throws Exception
     */
    public function testSkipEmptyCellsOption($skip_empty_cells, $exp_num_cols)
    {
        $reader = new Reader((new ReaderConfiguration())
            ->setSkipEmptyCells($skip_empty_cells)
        );
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

        $num_cols = array();
        foreach ($reader as $row) {
            $num_cols[] = count($row);
        }

        self::assertEquals($exp_num_cols, $num_cols, 'Number of cells differ');

        $reader->close();
    }

    /**
     * @return array
     */
    public function dataProviderEmptyCells()
    {
        return array(
            'SKIP_NONE' => array(
                'skipEmptyCells' => ReaderSkipConfiguration::SKIP_NONE,
                'numTotalCols'   => [8, 0, 4]
            ),
            'SKIP_EMPTY' => array(
                'skipEmptyCells' => ReaderSkipConfiguration::SKIP_EMPTY,
                'numTotalCols'   => [5, 0, 1]
            ),
            'SKIP_TRAILING_EMPTY' => array(
                'skipEmptyCells' => ReaderSkipConfiguration::SKIP_TRAILING_EMPTY,
                'numTotalCols'   => [8, 0, 2]
            )
        );
    }
}
