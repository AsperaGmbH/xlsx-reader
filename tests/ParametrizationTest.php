<?php

namespace Aspera\Spreadsheet\XLSX\Tests;

require_once __DIR__.'/../vendor/autoload.php';

use Exception;
use PHPUnit\Framework\TestCase as PHPUnitTestCase;
use Aspera\Spreadsheet\XLSX\Reader as XLSXReader;

class ParametrizationTest extends PHPUnitTestCase
{
    const FILE_PATH = 'input_files/iterator_test.xlsx';
    const TEMP_DIR_PATH = __DIR__.'/temp_new_folder';

    /** @var XLSXReader */
    private $reader;

    public function testOptionTempDir()
    {
        // create a directory 'read only'
        @mkdir(self::TEMP_DIR_PATH, 0444);

        $options = array(
            'TempDir' => self::TEMP_DIR_PATH
        );

        try {
            $this->reader = new XLSXReader(self::FILE_PATH, $options);
            self::fail('[TempDir] parameter check should have failed.');
        } catch (Exception $e) {
            // nothing unexpected happened
            self::assertSame(
                'XLSXReader: Provided temporary directory ('.self::TEMP_DIR_PATH.') is not writable',
                $e->getMessage()
            );
        }
    }

    public function testDeletionTemporaryFiles()
    {
        // create a directory 'full access'
        @mkdir(self::TEMP_DIR_PATH);

        $options = array(
            'TempDir' => self::TEMP_DIR_PATH
        );
        $this->reader = new XLSXReader(self::FILE_PATH, $options);
        // invoke destructor
        unset($this->reader);

        // after destruction, temporary directory must be totally emptied
        $folder_scan = @scandir(self::TEMP_DIR_PATH, SCANDIR_SORT_NONE);
        self::assertCount(2, $folder_scan, 'Folder ['.self::TEMP_DIR_PATH.'] is not empty');
    }

    public function tearDown()
    {
        // cleanup temporary directory
        if (file_exists(self::TEMP_DIR_PATH)) {
            @rmdir(self::TEMP_DIR_PATH);
        }
    }

    /**
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

        $this->reader = new XLSXReader(self::FILE_PATH, $options);

        $sheet_index = array_keys($this->reader->getSheets(), 'EmptyCellsSheet');
        $this->reader->changeSheet($sheet_index[0]);

        $current = $this->reader->current();
        $num_cols = count($current);

        self::assertEquals($exp_num_cols, $num_cols, 'Number of cells differ');
    }

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