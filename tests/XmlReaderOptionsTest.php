<?php

namespace Aspera\Spreadsheet\XLSX\Tests;

require_once __DIR__ . '/../vendor/autoload.php';

use Exception;
use PHPUnit\Framework\TestCase;
use Aspera\Spreadsheet\XLSX\Reader;
use Aspera\Spreadsheet\XLSX\ReaderConfiguration;

class XmlReaderOptionsTest extends TestCase
{
    /**
     * The file deep_xml_tag_nesting.xlsx contains a stack of dead XML tags near the beginning of each XML file.
     * This results in XMLReader errors when using default configuration values, as the nesting depth limit is exceeded.
     * The only way to successfully read this file with XMLReader is by providing the LIBXML_PARSEHUGE option.
     *
     * @var string
     */
    private const TEST_FILE = __DIR__ . '/input_files/deep_xml_tag_nesting.xlsx';

    /**
     * Check whether the XMLReader options are supplied correctly to each XMLReader instance.
     *
     * @dataProvider dataForTestCellContent
     * @throws Exception
     */
    public function testCellContent(int $xml_reader_flags, bool $expect_exception): void
    {
        $reader = new Reader(
            (new ReaderConfiguration())
                ->setXmlReaderFlags($xml_reader_flags)
        );

        if ($expect_exception) {
            $this->expectException('PHPUnit\Framework\Error\Warning'); // required for expectExceptionMessage to work
            $this->expectExceptionMessage('parser error : Excessive depth in document: 256 use XML_PARSE_HUGE option');
        }

        // Reading just one value from the file requires all XML files of the XLSX to be parsed.
        $output_cells = array();
        $reader->open(self::TEST_FILE);
        foreach ($reader as $row) {
            $output_cells[] = $row[0];
        }
        $reader->close();

        // Assertion is more of a sanity check at this point.
        self::assertSame(
            array('test content'),
            $output_cells,
            'The retrieved sheet content was not as expected.'
        );
    }

    public function dataForTestCellContent(): array
    {
        return array(
            array(0, true),
            array(LIBXML_PARSEHUGE, false)
        );
    }
}
