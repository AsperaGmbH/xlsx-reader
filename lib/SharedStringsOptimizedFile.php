<?php

namespace Aspera\Spreadsheet\XLSX;

use RuntimeException;

/**
 * Data object to hold all data corresponding to a single optimized shared string file (not the original XML file).
 *
 * @author Aspera GmbH
 */
class SharedStringsOptimizedFile
{
    /** @var string Complete path to the file. */
    private $file = '';

    /** @var resource File handle to access the file contents with. */
    private $handle;

    /** @var int Index of the line the handle currently points at. (Only used during reading from the file) */
    private $handle_current_index = -1;

    /** @var string The shared string value corresponding to the current index. (Only used during reading from the file) */
    private $value_at_current_index = '';

    /** @var int Total number of shared strings contained within the file. */
    private $count = 0;

    /**
     * @return string
     */
    public function getFile()
    {
        return $this->file;
    }

    /**
     * @param string $file
     */
    public function setFile($file)
    {
        $this->file = $file;
    }

    /**
     * @return resource
     */
    public function getHandle()
    {
        return $this->handle;
    }

    /**
     * @param resource $handle
     */
    public function setHandle($handle)
    {
        $this->handle = $handle;
    }

    /**
     * @return int
     */
    public function getHandleCurrentIndex()
    {
        return $this->handle_current_index;
    }

    /**
     * @param int $handle_current_index
     */
    public function setHandleCurrentIndex($handle_current_index)
    {
        $this->handle_current_index = $handle_current_index;
    }

    /**
     * Increase current index of handle by 1.
     */
    public function increaseHandleCurrentIndex()
    {
        $this->handle_current_index++;
    }

    /**
     * @return string
     */
    public function getValueAtCurrentIndex()
    {
        return $this->value_at_current_index;
    }

    /**
     * @param string $value_at_current_index
     */
    public function setValueAtCurrentIndex($value_at_current_index)
    {
        $this->value_at_current_index = $value_at_current_index;
    }

    /**
     * @return int
     */
    public function getCount()
    {
        return $this->count;
    }

    /**
     * @param int $count
     */
    public function setCount($count)
    {
        $this->count = $count;
    }

    /**
     * Increase count of elements contained within the file by 1.
     */
    public function increaseCount()
    {
        $this->count++;
    }

    /**
     * Opens a file handle to the file with the given file access mode.
     * If a file handle is currently still open, closes it first.
     *
     * @param   string      $mode
     * @return  resource    The newly opened file handle
     *
     * @throws  RuntimeException
     */
    public function openHandle($mode)
    {
        $this->closeHandle();
        $new_handle = @fopen($this->getFile(), $mode);
        if (!$new_handle) {
            throw new RuntimeException(
                'Could not open file handle for optimized shared string file with mode ' . $mode . '.'
            );
        }
        $this->setHandle($new_handle);
        return $this->getHandle();
    }

    /**
     * Properly closes the current file handle, if it is currently opened.
     */
    public function closeHandle()
    {
        if (!$this->handle) {
            return; // Nothing to close
        }
        fclose($this->handle);
        $this->handle = null;
        $this->handle_current_index = -1;
        $this->value_at_current_index = '';
    }

    /**
     * Properly rewinds the current file handle and all associated internal data.
     *
     * @return  resource    The rewound file handle
     *
     * @throws  RuntimeException
     */
    public function rewindHandle()
    {
        if (!$this->handle) {
            throw new RuntimeException('Could not rewind file handle; There is no file handle currently open.');
        }
        rewind($this->handle);
        $this->handle_current_index = -1;
        $this->value_at_current_index = null;
        return $this->handle;
    }
}
