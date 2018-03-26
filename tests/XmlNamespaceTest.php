<?php

namespace Aspera\Spreadsheet\XLSX\Tests;

use Exception;
use PHPUnit\Framework\TestCase;
use Aspera\Spreadsheet\XLSX\Reader;

/**
 * Ensure that the reader can work with files making use of XML namespaces.
 *
 * @author Aspera GmbH
 */
class XmlNamespaceTest extends TestCase
{
    const TEST_FILE = 'input_files/xml_namespace_test.xlsx';

    /**
     * Attempt reading a file that uses namespaces everywhere.
     * Ensure that the contents read from this file (including shared strings and formatted values) are as expected.
     *
     * @throws Exception
     */
    public function testReadXmlWithNamespaces()
    {
        $reader = new Reader(self::TEST_FILE);
        $actual_row = $reader->next();
        $expected_row = array('1.230000 €', 'test string');
        self::assertSame(
            $expected_row,
            $actual_row,
            'Could not read data from test file; XML namespace handling might be broken.'
            . ' Retrieved row contents: [' . implode('|', $actual_row) . ']'
        );
    }
}
