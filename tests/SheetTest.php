<?php

namespace Aspera\Spreadsheet\XLSX\Tests;

require_once __DIR__ . '/../vendor/autoload.php';

use Exception;
use Aspera\Spreadsheet\XLSX\Worksheet;
use PHPUnit\Framework\TestCase as PHPUnitTestCase;
use Aspera\Spreadsheet\XLSX\Reader as XLSXReader;

/**
 * Tests regarding basic worksheet handling functionality.
 *
 * @author Aspera GmbH
 */
class SheetTest extends PHPUnitTestCase
{
    const FILE_PATH = 'input_files\multiple_sheets_test.xlsx';

    /** @var XLSXReader */
    private $reader;

    /**
     * @throws Exception
     */
    public function setUp()
    {
        $this->reader = new XLSXReader();
        $this->reader->open(self::FILE_PATH);
    }

    public function tearDown()
    {
        $this->reader->close();
    }

    /**
     * Checks if the reader is capable of reading the names of worksheets correctly.
     */
    public function testGetSheetsFunction()
    {
        $exp_sheets = array(
            'First Sheet',
            'Second Sheet',
            'Third Sheet'
        );
        $sheet_name_list = array();
        /** @var Worksheet $worksheet */
        foreach ($this->reader->getSheets() as $worksheet) {
            $sheet_name_list[] = $worksheet->getName();
        }
        self::assertSame($exp_sheets, $sheet_name_list, 'Sheet list differs');
    }

    /**
     * Checks if changing the sheet works as expected and also if it handles faulty inputs correctly.
     *
     * @depends testGetSheetsFunction
     *
     * @throws  Exception
     */
    public function testChangeSheetFunction()
    {
        /** @var Worksheet $worksheet */
        foreach ($this->reader->getSheets() as $index => $worksheet) {
            $sheet_name_in_sheet_data = $worksheet->getName();
            self::assertTrue(
                $this->reader->changeSheet($index),
                'Unable to switch to sheet [' . $index . '] => [' . $sheet_name_in_sheet_data . ']'
            );

            // For testing, the sheet name is written in each sheet's first cell of the first line
            $content = $this->reader->current();
            self::assertTrue(
                is_array($content) && !empty($content),
                'No content found in sheet [' . $index . '] => [' . $sheet_name_in_sheet_data . ']'
            );
            $sheet_name_in_cell = $content[0];
            self::assertSame(
                $sheet_name_in_cell,
                $sheet_name_in_sheet_data,
                'Sheet has been changed to a wrong one'
            );
        }

        // test index out of bounds
        self::assertFalse(
            $this->reader->changeSheet(-1),
            'Error expected when stepping out of bounds'
        );
    }
}
