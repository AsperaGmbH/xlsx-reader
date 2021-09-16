<?php

namespace Aspera\Spreadsheet\XLSX\Tests;

require_once __DIR__ . '/../vendor/autoload.php';

use Exception;
use PHPUnit\Framework\TestCase;
use Aspera\Spreadsheet\XLSX\Reader;

/** Ensure that the reader can work with files using a different internal file part structure than the commonly used one. */
class FilepartLocationTest extends TestCase
{
    const TEST_FILE = __DIR__ . '/input_files/filepart_location_test.xlsx';

    /**
     * Attempt reading a file that has none of its files in the usual folders, except for the ones that absolutely require it.
     * Ensure that the contents read from this file (including shared strings and formatted values) are as expected.
     *
     * @throws Exception
     */
    public function testReadDocumentWithUncommonFilepartPaths()
    {
        $reader = new Reader();
        $reader->open(self::TEST_FILE);
        $actual_row = $reader->current();
        $expected_row = array('1.230000 €', 'test string');
        self::assertSame(
            $expected_row,
            $actual_row,
            'Could not read data from test file; Filepart/Relationship handling might be broken.'
            . ' Retrieved row contents: [' . implode('|', $actual_row) . ']'
        );
        $reader->close();
    }
}
