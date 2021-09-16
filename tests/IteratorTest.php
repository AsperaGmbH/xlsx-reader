<?php

namespace Aspera\Spreadsheet\XLSX\Tests;

require_once __DIR__ . '/../vendor/autoload.php';

use Exception;
use PHPUnit\Framework\TestCase as PHPUnitTestCase;
use Aspera\Spreadsheet\XLSX\Reader;

/** Tests ensuring that the Reader properly implements the Iterator interface */
class IteratorTest extends PHPUnitTestCase
{
    const FILE_PATH = __DIR__ . '/input_files/iterator_test.xlsx';

    /** @var Reader */
    private $reader;

    /**
     * @throws Exception
     */
    public function setUp()
    {
        $this->reader = new Reader();
        $this->reader->open(self::FILE_PATH);
    }

    public function tearDown()
    {
        $this->reader->close();
    }

    /**
     * Tests pointer movements of the iterator.
     *
     * @throws Exception
     */
    public function testIterationFunctions()
    {
        $first_row_content = $this->reader->current();
        $this->reader->next();
        $second_row_content = $this->reader->current();
        self::assertNotEquals($first_row_content, $second_row_content, 'First and second row are identical');

        $this->reader->rewind();
        self::assertEquals($first_row_content, $this->reader->current(),
            'rewind() function did not rewind/reset the pointer. Target should be the first row');
        $this->reader->next();
        self::assertNotEquals($first_row_content, $this->reader->current(),
            'next() function did not move the pointer. Target should be the second row');
        self::assertEquals($second_row_content, $this->reader->current(),
            'current() function did not work. Target should be the second row');
    }

    /**
     * Tests that the return value of the method key() is actually increasing/decreasing.
     * Notice that key() and count() do the same functionality based on the current implementation.
     *
     * @depends testIterationFunctions
     * @throws  Exception
     */
    public function testPositioningFunctions()
    {
        $row_number = $this->reader->key();
        self::assertEquals(0, $row_number, 'Row number should be zero');

        $this->reader->next();
        $current_row_number = $this->reader->key();
        self::assertEquals($row_number + 1, $current_row_number, 'Row number should be one');

        $this->reader->rewind();
        $current_row_number = $this->reader->key();
        self::assertEquals(0, $current_row_number, 'Row number should be zero due to rewind()');

        // are count() and key() doing the same?
        self::assertEquals(
            $this->reader->count(),
            $this->reader->key(),
            'Functions count() and key() should return the same'
        );
    }

    /**
     * Tests if we've iterated to the end of the collection
     *
     * @depends testIterationFunctions
     * @throws  Exception
     */
    public function testFunctionValid()
    {
        $read_file = array();
        while (is_array($this->reader->current()) && $this->reader->valid()) {
            $read_file[] = $this->reader->current();
            $this->reader->next();
        }
        self::assertFalse($this->reader->valid(), 'File reading has finished and it is still valid');
        self::assertEquals($this->getExpectedArray(), $read_file, 'File has not been read correctly');
    }

    private function getExpectedArray()
    {
        return array(
            array(
                'text1',
                'text2',
                'text3',
                '',
                'shared string',
                'inline string'
            ),
            array(),
            array(
                '',
                '',
                '',
                'text1',
                'text1',
                'text1'
            ),
            array(
                '',
                '',
                '',
                'text2',
                'text2',
                'text2'
            ),
            array(
                '',
                '',
                '',
                'text3',
                'text3',
                'text3'
            ),
            array(
                'all borders',
                '',
                '',
                '',
                '',
                '',
                '',
                'border-a1',
                'border-b1',
                'border-c1'
            ),
            array(
                '',
                '',
                '',
                '',
                '',
                '',
                '',
                'border-a2',
                'border-b2',
                'border-c2'
            ),
            array(
                '',
                '',
                '',
                '',
                '',
                '',
                '',
                'border-a3',
                'border-b3',
                'border-c3'
            ),
            array(
                '12345',
                '123.45',
                '123.45',
                '',
                '12468.45',
                '12468'
            ),
            array(),
            array(
                '10000.4',
                '10,000.40 €',
                '18/05/1927',
                '05-19-27',
                '9:36:00 AM',
                '- Wednesday 09:36 -'
            ),
            array(),
            array(
                'a cell',
                'a long cell',
                '',
                '',
                ''
            )
        );
    }
}