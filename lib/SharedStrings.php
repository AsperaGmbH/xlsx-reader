<?php

namespace Aspera\Spreadsheet\XLSX;

use XMLReader;
use Exception;
use SplFixedArray;

require_once('SharedStringsConfiguration.php');
require_once('SharedStringsOptimizedFile.php');

class SharedStrings
{
    /**
     *
     * Amount of array positions to add with each extension of the shared string cache array.
     * Larger values run into memory limits faster, lower values are a tiny bit worse on performance
     *
     * @var int SHARED_STRING_COUNT_PER_FILE
     */
    const SHARED_STRING_CACHE_ARRAY_SIZE_STEP = 100;

    /**
     * Filename (without path) of the shared strings XML file.
     *
     * @var string $shared_strings_filename
     */
    private $shared_strings_filename;

    /**
     * Path to the directory containing all shared strings files used by this instance.
     * Includes trailing slash.
     *
     * @var string $shared_strings_directory
     */
    private $shared_strings_directory;

    /**
     * XML reader object for the shared strings XML file
     *
     * @var XMLReader $shared_strings_reader
     */
    private $shared_strings_reader;

    /**
     * Configuration of shared string reading and caching behaviour.
     *
     * @var SharedStringsConfiguration
     */
    private $shared_strings_configuration;

    /**
     * Shared strings cache, if the number of shared strings is low enough
     *
     * @var SplFixedArray $shared_string_cache
     */
    private $shared_string_cache;

    /**
     * Array of SharedStringsOptimizedFile instances containing filenames and associated data for shared strings that
     * were not saved to $shared_string_cache. Files contain values in seek-optimized format. (one entry per line, JSON encoded)
     * Key per element: the index of the first string contained within the file
     *
     * @var array $prepared_shared_string_files
     */
    private $prepared_shared_string_files = array();

    /**
     * The total number of shared strings available in the file.
     *
     * @var int $shared_string_count
     */
    private $shared_string_count = 0;

    /**
     * The shared string index the shared string reader is currently pointing at.
     *
     * @var int $shared_string_index
     */
    private $shared_string_index = 0;

    /**
     * Temporary cache for the last value that was read from the shared strings xml file.
     *
     * @var null $last_shared_string_value
     */
    private $last_shared_string_value;

    /**
     * SharedStrings constructor. Prepares the data stored within the given shared string file for reading.
     *
     * @param string $shared_strings_directory Directory of the shared strings file
     * @param string $shared_strings_filename Filename of the shared strings file
     * @param SharedStringsConfiguration $shared_strings_configuration Configuration for shared string reading and caching behaviour
     *
     * @throws Exception
     */
    public function __construct(
        $shared_strings_directory,
        $shared_strings_filename,
        SharedStringsConfiguration $shared_strings_configuration = null
    ) {
        $this->shared_strings_configuration = $shared_strings_configuration ?: new SharedStringsConfiguration();
        $this->shared_strings_directory = $shared_strings_directory;
        $this->shared_strings_filename = $shared_strings_filename;
        if (is_readable($this->shared_strings_directory . $this->shared_strings_filename)) {
            $this->prepareSharedStrings();
        }
    }

    /**
     * Closes all file pointers managed by this SharedStrings instance.
     * Note: Does not unlink temporary files. Use getTempFiles() to retrieve the list of created temp files.
     */
    public function close()
    {
        if ($this->shared_strings_reader && $this->shared_strings_reader instanceof XMLReader) {
            $this->shared_strings_reader->close();
            unset($this->shared_strings_reader);
        }
        /** @var SharedStringsOptimizedFile $file_data */
        foreach ($this->prepared_shared_string_files as $file_data) {
            $file_data->closeHandle();
        }
        unset($this->shared_strings_directory);
        unset($this->shared_strings_filename);
    }

    /**
     * @param SharedStringsConfiguration $configuration
     */
    public function setSharedStringsConfiguration(SharedStringsConfiguration $configuration)
    {
        $this->shared_strings_configuration = $configuration;
    }

    /**
     * Returns a list of all temporary work files created in this SharedStrings instance.
     *
     * @return array List of temporary files; With absolute paths.
     */
    public function getTempFiles()
    {
        $ret = array();
        /** @var SharedStringsOptimizedFile $file_details */
        foreach ($this->prepared_shared_string_files as $file_details) {
            $ret[] = $file_details->getFile();
        }
        return $ret;
    }

    /**
     * Retrieves a shared string value by its index
     *
     * @param int $target_index Shared string index
     * @return string Shared string of the given index
     *
     * @throws Exception
     */
    public function getSharedString($target_index)
    {
        // If index of the desired string is larger than possible, don't even bother.
        if ($this->shared_string_count && ($target_index >= $this->shared_string_count)) {
            return '';
        }

        // Read from RAM cache?
        if ($this->shared_strings_configuration->getUseCache() && isset($this->shared_string_cache[$target_index])) {
            return $this->shared_string_cache[$target_index];
        }

        // Read from optimized files?
        if ($this->shared_strings_configuration->getUseOptimizedFiles()) {
            $result = $this->getStringFromOptimizedFile($target_index);
            if ($result) {
                return $result;
            }
        }

        // No cache and no optimized files; Read directly from original XML
        return $this->getStringFromOriginalSharedStringFile($target_index);
    }

    /**
     * Attempts to retrieve a string from the optimized shared string files.
     * May return null if unsuccessful.
     *
     * @param int $target_index
     * @return null|string
     *
     * @throws Exception
     */
    private function getStringFromOptimizedFile($target_index)
    {
        // Determine the target file to read from, given the smallest index obtainable from it.
        $index_of_target_file = null;
        foreach (array_keys($this->prepared_shared_string_files) as $lowest_index) {
            if ($lowest_index > $target_index) {
                break; // Because the array is ksorted, we can assume that we've found our value at this point.
            }
            $index_of_target_file = $lowest_index;
        }
        if ($index_of_target_file === null) {
            return null;
        }

        /** @var SharedStringsOptimizedFile $file_data */
        $file_data = $this->prepared_shared_string_files[$index_of_target_file];

        // Determine our target line in the target file
        $target_index_in_file = $target_index - $index_of_target_file; // note: $index_of_target_file is also the index of the first string within the file
        if ($file_data->getHandleCurrentIndex() == $target_index_in_file) {
            // tiny optimization; If a previous seek already evaluated the target value, return it immediately
            return $file_data->getValueAtCurrentIndex();
        }

        // We found our target file to read from. Open a file handle or retrieve an already opened one.
        $target_handle = $file_data->getHandle();
        if (!$target_handle) {
            $target_handle = $file_data->openHandle('rb');
        }

        // Potentially rewind the file handle.
        if ($file_data->getHandleCurrentIndex() > $target_index_in_file) {
            // Our file pointer points at an index after the one we're looking for; Rewind the file pointer
            $target_handle = $file_data->rewindHandle();
        }

        // Walk through the file up to the index we're looking for and return its value
        $file_line = null;
        while ($file_data->getHandleCurrentIndex() < $target_index_in_file) {
            $file_data->increaseHandleCurrentIndex();
            $file_line = fgets($target_handle);
            if ($file_line === false) {
                break; // unexpected EOF; Silent fallback to original shared string file.
            }
        }
        if (is_string($file_line) && $file_line != '') {
            $file_line = json_decode($file_line);
            if ($this->shared_strings_configuration->getKeepFileHandles()) {
                $file_data->setValueAtCurrentIndex($file_line);
            } else {
                $file_data->closeHandle();
            }
            return $file_line;
        }
        return null;
    }

    /**
     * Retrieves a shared string from the original shared strings XML file.
     *
     * @param int $target_index
     * @return null|string
     */
    private function getStringFromOriginalSharedStringFile($target_index)
    {
        // If the desired index is before the current, rewind the XML
        if ($this->shared_strings_reader && $this->shared_string_index > $target_index) {
            $this->shared_strings_reader->close();
            $this->shared_strings_reader = null;
        }

        // Intialize reader, if not already initialized
        if (!$this->shared_strings_reader) {
            $this->shared_strings_reader = new XMLReader();
            $this->shared_strings_reader->open($this->shared_strings_directory . $this->shared_strings_filename);
            $this->shared_string_index = -1;
            $this->last_shared_string_value = null;
        }

        // If an index with the same value as the last already fetched is requested
        // (any further traversing the tree would get us further away from the node)
        if (($index == $this->shared_string_index) && ($this->last_shared_string_value !== null)) {
            return $this->last_shared_string_value;
        }

        // Move reader to the next <si> node, if it isn't already pointing at one.
        $last_result = true;
        $moved_forward = false;
        while ($last_result && ($this->shared_strings_reader->name != 'si' || $this->shared_strings_reader->nodeType == XMLReader::END_ELEMENT)) {
            $moved_forward = true;
            $last_result = $this->shared_strings_reader->read();
        }
        if (!$last_result) {
            // Unexpected EOF
            $this->shared_strings_reader->close();
            $this->shared_strings_reader = null;
            return '';
        }
        if ($moved_forward) {
            $this->shared_string_index++;
        }

        // Find the <si> node with the desired index
        $last_result = true;
        while ($last_result && $this->shared_string_index < $target_index) {
            $last_result = $this->shared_strings_reader->next('si');
            $this->shared_string_index++;
        }
        if (!$last_result) {
            // Unexpected EOF
            $this->shared_strings_reader->close();
            $this->shared_strings_reader = null;
            return '';
        }

        // Extract the value from the shared string
        $value = '';
        while ($this->shared_strings_reader->read()) {
            switch ($this->shared_strings_reader->name) {
                case 't':
                    if ($this->shared_strings_reader->nodeType == XMLReader::END_ELEMENT) {
                        continue 2;
                    }
                    $value .= $this->shared_strings_reader->readString();
                    break;
                case 'si':
                    if ($this->shared_strings_reader->nodeType == XMLReader::END_ELEMENT) {
                        break 2;
                    }
                    break;
            }
        }
        if ($value) {
            $this->last_shared_string_value = $value;
        }

        if (!$this->shared_strings_configuration->getKeepFileHandles()) {
            $this->shared_strings_reader->close();
            $this->shared_strings_reader = null;
        }

        return $value;
    }

    /**
     * Perform optimizations to increase performance of shared string determination operations.
     * Loads shared string data into RAM up to the configured memory limit. Stores additional shared string data
     * in seek-optimized additional files on the filesystem in order to lower seek times.
     *
     * @return void
     *
     * @throws Exception
     */
    private function prepareSharedStrings()
    {
        $this->shared_strings_reader = new XMLReader;
        $this->shared_strings_reader->open($this->shared_strings_directory . $this->shared_strings_filename);

        // Obtain number of shared strings available
        while ($this->shared_strings_reader->read()) {
            if ($this->shared_strings_reader->name == 'sst') {
                $this->shared_string_count = $this->shared_strings_reader->getAttribute('uniqueCount');
                break;
            }
        }
        if (!$this->shared_string_count) {
            return; // No shared strings available, no preparation necessary
        }
        if ($this->shared_strings_configuration->getUseCache()) {
            // This is why we ask for at least 8 KB of memory; Lower values may already exceed the limit with this assignment:
            $this->shared_string_cache = new SplFixedArray(self::SHARED_STRING_CACHE_ARRAY_SIZE_STEP);
        }

        // Prepare working through the XML file. Declare as many variables as makes sense upfront, for more accurate memory usage retrieval.
        $string_index = 0;
        $string_value = '';
        $write_to_cache = $this->shared_strings_configuration->getUseCache();
        $cache_max_size_byte = $this->shared_strings_configuration->getCacheSizeKilobyte() * 1024;
        $start_memory_byte = memory_get_usage(false); // Note: Get current memory usage as late as possible

        // Work through the XML file and cache/reformat/move string data, according to configuration and situation
        while ($this->shared_strings_reader->read()) {
            switch ($this->shared_strings_reader->name) {
                case 'si':
                    if ($this->shared_strings_reader->nodeType == XMLReader::END_ELEMENT) {
                        if ($write_to_cache) {
                            $cache_current_memory_byte = memory_get_usage(false) - $start_memory_byte;
                            if ($cache_current_memory_byte > $cache_max_size_byte) {
                                // transition from "cache everything" to "memory exhausted, stop caching":
                                $this->shared_string_cache->setSize($string_index); // finalize array size
                                $write_to_cache = false;
                            }
                        }
                        $this->prepareSingleSharedString($string_index, $string_value, $write_to_cache);
                        $string_index++;
                        $string_value = '';
                    }
                    break;
                case 't':
                    if ($this->shared_strings_reader->nodeType != XMLReader::END_ELEMENT) {
                        $string_value .= $this->shared_strings_reader->readString();
                    }
                    break;
            }
        }

        // Small optimization: Sort shared string files by lowest included key for slightly faster reading.
        ksort($this->prepared_shared_string_files);

        // Close all no longer necessary file handles
        $this->shared_strings_reader->close();
        $this->shared_strings_reader = null;
        /** @var SharedStringsOptimizedFile $file_data */
        foreach ($this->prepared_shared_string_files as $file_data) {
            $file_data->closeHandle();
        }
    }

    /**
     * Stores the given shared string either in internal cache or in a seek optimized file, depending on the
     * current configuration and status of the internal cache.
     *
     * @param int $index
     * @param string $string
     * @param bool $write_to_cache
     *
     * @throws Exception
     */
    private function prepareSingleSharedString($index, $string, $write_to_cache = false)
    {
        if ($write_to_cache) {
            // Caching enabled and there's still memory available; Write to internal cache.
            if ($index + 1 > $this->shared_string_cache->getSize()) {
                $this->shared_string_cache->setSize($this->shared_string_cache->getSize() + self::SHARED_STRING_CACHE_ARRAY_SIZE_STEP);
            }
            $this->shared_string_cache[$index] = $string;
            return;
        }

        if (!$this->shared_strings_configuration->getUseOptimizedFiles()) {
            // No preparation possible. This value will have to be read from the original shared string XML file.
            return;
        }

        // Caching not possible. Write shared string to seek-optimized file instead.

        // Check if we have an already existing file that still has room for more entries in it.
        /** @var SharedStringsOptimizedFile $newest_file_data */
        $newest_file_data = null;
        $newest_file_is_full = false;
        $shared_string_file_index = null;
        if (count($this->prepared_shared_string_files) > 0) {
            $shared_string_file_index = max(array_keys($this->prepared_shared_string_files));
            $newest_file_data = $this->prepared_shared_string_files[$shared_string_file_index];
            if ($newest_file_data->getCount() > $this->shared_strings_configuration->getOptimizedFileEntryCount()) {
                $newest_file_is_full = true;
            }
        }

        $create_new_file = !$newest_file_data || $newest_file_is_full;
        if ($create_new_file) {
            // Assemble new filename; Add random hash to avoid conflicts for when the target directory is also used by other processes.
            $hash = base_convert(mt_rand(pow(36, 4), pow(36, 5) - 1), 10, 36); // Possible results: "10000" - "zzzzz"
            $newest_file_data = new SharedStringsOptimizedFile();
            $newest_file_data->setFile($this->shared_strings_directory . 'shared_strings_tmp_' . $index . '_' . $hash . '.txt');
            $fhandle = $newest_file_data->openHandle('wb');
            $this->prepared_shared_string_files[$index] = $newest_file_data;
        } else {
            // Append shared string to the newest file.
            $fhandle = $newest_file_data->getHandle();
            if (!$fhandle) {
                $fhandle = $newest_file_data->openHandle('ab');
            }
        }

        // Write shared string to the chosen file
        if (fwrite($fhandle, json_encode($string) . PHP_EOL) === false) {
            throw new Exception('Could not write shared string to temporary file.');
        }
        $newest_file_data->increaseCount();

        if (!$this->shared_strings_configuration->getKeepFileHandles()) {
            $newest_file_data->closeHandle();
        }
    }
}
