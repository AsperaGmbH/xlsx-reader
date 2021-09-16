<?php

namespace Aspera\Spreadsheet\XLSX\Tests;

require_once __DIR__ . '/../vendor/autoload.php';

use Exception;
use PHPUnit\Framework\TestCase;
use Aspera\Spreadsheet\XLSX\Reader;

/** Ensure that the reader can work with files making use of XML namespaces. */
class XmlNamespaceTest extends TestCase
{
    /**
     * Things of note about the test file:
     * - All elements in all xml files have a valid root namespace prefix.
     * - workbook.xml does not declare the relationship namespace in the root element. This is valid, but not commonly seen.
     * - workbook.xml uses edition 3 namespaces, while the rest of the document uses edition 1 namespaces. This should not cause issues.
     */
    const TEST_FILE = __DIR__ . '/input_files/xml_namespace_test.xlsx';

    /**
     * Attempt reading a file that uses namespaces everywhere.
     * Ensure that the contents read from this file (including shared strings and formatted values) are as expected.
     *
     * @throws Exception
     */
    public function testReadXmlWithNamespaces()
    {
        $reader = new Reader();
        $reader->open(self::TEST_FILE);

        $actual_row = $reader->current();
        $expected_row = array('1.230000 €', 'test string');
        self::assertSame(
            $expected_row,
            $actual_row,
            'Could not read data from test file; XML namespace handling might be broken.'
            . ' Retrieved row contents: [' . implode('|', $actual_row) . ']'
        );

        $reader->close();
    }
}
