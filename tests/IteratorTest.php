<?php

namespace Aspera\Spreadsheet\XLSX\Tests;

require_once __DIR__ . '/../vendor/autoload.php';

use Exception;
use PHPUnit\Framework\TestCase as PHPUnitTestCase;
use Aspera\Spreadsheet\XLSX\Reader as XLSXReader;

/**
 * Tests ensuring that the Reader properly implements the Iterator interface
 *
 * @author Aspera GmbH
 */
class IteratorTest extends PHPUnitTestCase
{
    const FILE_PATH = 'input_files\iterator_test.xlsx';
    const EXPECTED_ARRAY = array(
        0  => array(
            0 => 'text1',
            1 => 'text2',
            2 => 'text3',
            3 => '',
            4 => 'shared string',
            5 => 'inline string'
        ),
        1  => array(),
        2  => array(
            0 => '',
            1 => '',
            2 => '',
            3 => 'text1',
            4 => 'text1',
            5 => 'text1'
        ),
        3  => array(
            0 => '',
            1 => '',
            2 => '',
            3 => 'text2',
            4 => 'text2',
            5 => 'text2'
        ),
        4  => array(
            0 => '',
            1 => '',
            2 => '',
            3 => 'text3',
            4 => 'text3',
            5 => 'text3'
        ),
        5  => array(
            0 => 'all borders',
            1 => '',
            2 => '',
            3 => '',
            4 => '',
            5 => '',
            6 => '',
            7 => 'border-a1',
            8 => 'border-b1',
            9 => 'border-c1'
        ),
        6  => array(
            0 => '',
            1 => '',
            2 => '',
            3 => '',
            4 => '',
            5 => '',
            6 => '',
            7 => 'border-a2',
            8 => 'border-b2',
            9 => 'border-c2'
        ),
        7  => array(
            0 => '',
            1 => '',
            2 => '',
            3 => '',
            4 => '',
            5 => '',
            6 => '',
            7 => 'border-a3',
            8 => 'border-b3',
            9 => 'border-c3'
        ),
        8  => array(
            0 => '12345',
            1 => '123.45',
            2 => '123.45',
            3 => '',
            4 => '12468.45',
            5 => '12468'
        ),
        9  => array(),
        10 => array(
            0 => '10000.4',
            1 => '10000.4',
            2 => '10000.4',
            3 => '10001',
            4 => '10000.4',
            5 => '10000.4'
        ),
        11 => array(),
        12 => array(
            0 => 'a cell',
            1 => 'a long cell',
            2 => '',
            3 => '',
            4 => ''
        )
    );

    /** @var XLSXReader */
    private $reader;

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
     * Tests pointer movements of the iterator.
     *
     * @throws Exception
     */
    public function testIterationFunctions()
    {
        $first_row_content = $this->reader->current();
        $second_row_content = $this->reader->next();
        self::assertNotEquals($first_row_content, $second_row_content, 'First and second row are identical');

        $this->reader->rewind();
        self::assertEquals($first_row_content, $this->reader->current(),
            'rewind() function did not rewind/reset the pointer. Target should be the first row');
        self::assertNotEquals($first_row_content, $this->reader->next(),
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
        self::assertEquals($this->reader->count(), $this->reader->key(),
            'Functions count() and key() should return the same');
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
        self::assertEquals(self::EXPECTED_ARRAY, $read_file, 'File has not be read correctly');
    }
}