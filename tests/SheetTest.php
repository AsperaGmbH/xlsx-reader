<?php

namespace Aspera\Spreadsheet\XLSX\Tests;

require_once __DIR__ . '/../vendor/autoload.php';

use Exception;
use PHPUnit\Framework\TestCase as PHPUnitTestCase;
use Aspera\Spreadsheet\XLSX\Reader as XLSXReader;

class SheetTest extends PHPUnitTestCase
{
    const FILE_PATH = 'input_files\multiple_sheets_test.xlsx';

    /** @var XLSXReader */
    private $reader;

    public function setUp()
    {
        $this->reader = new XLSXReader(self::FILE_PATH);
    }

    public function testGetSheetsFunction()
    {
        $exp_sheets = array(
            'First Sheet',
            'Second Sheet',
            'Third Sheet'
        );
        self::assertSame($this->reader->getSheets(), $exp_sheets, 'Sheet list differs');
    }

    /**
     * @depends testGetSheetsFunction
     *
     * @throws  Exception
     */
    public function testChangeSheetFunction()
    {
        foreach ($this->reader->getSheets() as $index => $value) {
            self::assertTrue($this->reader->changeSheet($index),
                'Unable to switch to sheet ['.$index.'] => ['.$value.']');

            // For testing, the sheet name is written in each sheet's first cell of the first line
            $content = $this->reader->current();
            self::assertTrue(is_array($content) && !empty($content),
                'No content found in sheet ['.$index.'] => ['.$value.']');
            $sheet_name = $content[0];
            self::assertSame($sheet_name, $value, 'Sheet has been changed to a wrong one');
        }

        // test index out of bounds
        self::assertFalse($this->reader->changeSheet(-1), 'Error expected when stepping out of bounds');
    }
}
