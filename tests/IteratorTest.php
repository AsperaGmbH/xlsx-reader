<?php

use Aspera\Spreadsheet\XLSX\Reader as XLSXReader;

class IteratorTest extends PHPUnit\Framework\TestCase
{
    const FILE_PATH = 'input_files\iterator_test.xlsx';

    /** @var XLSXReader */
    private $reader;

    public function setUp()
    {
        $this->reader = new XLSXReader(self::FILE_PATH);
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
        while (is_array($this->reader->current()) && !empty($this->reader->current())) {
            self::assertTrue($this->reader->valid(), 'Reading of the current record has failed');
            $this->reader->next();
        }
        self::assertFalse($this->reader->valid(), 'File reading has finished and it is still valid');
    }
}