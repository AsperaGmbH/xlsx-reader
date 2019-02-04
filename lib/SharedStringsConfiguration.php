<?php

namespace Aspera\Spreadsheet\XLSX;

use InvalidArgumentException;

/**
 * Holds all configuration options related to shared string related behaviour
 *
 * @author Aspera GmbH
 */
class SharedStringsConfiguration
{
    /**
     * If true: Allow caching shared strings to RAM to increase performance.
     *
     * @var bool
     */
    private $use_cache = true;

    /**
     * Maximum allowed RAM consumption for shared string cache, in kilobyte. (Minimum: 8 KB)
     * Once exceeded, additional shared strings will not be written to RAM and instead get read from file as needed.
     * Note that this is a "soft" limit that only applies to the main cache. The application may slightly exceed it.
     *
     * @var int
     */
    private $cache_size_kilobyte = 256;

    /**
     * If true: Allow creation of new files to reduce seek times for non-cached shared strings.
     *
     * @var bool
     */
    private $use_optimized_files = true;

    /**
     * Amount of shared strings to store per seek optimized shared strings file.
     * Lower values result in higher performance at the cost of more temporary files being created.
     * At extremely low values (< 10) you might be better off increasing the cache size.
     *
     * @var int
     */
    private $optimized_file_entry_count = 2500;

    /**
     * If true: file pointers to shared string files are kept open for more efficient reads.
     * Causes higher memory consumption, especially if $optimized_file_entry_count is low.
     *
     * @var bool
     */
    private $keep_file_handles = true;

    /**
     * Enable/disable caching of shared string values in RAM.
     *
     * @param   bool    $new_use_cache_value
     *
     * @throws  InvalidArgumentException
     */
    public function setUseCache($new_use_cache_value)
    {
        if (!is_bool($new_use_cache_value)) {
            throw new InvalidArgumentException('Invalid parameter value; Expected a boolean.');
        }
        $this->use_cache = $new_use_cache_value;
    }


    /**
     * Set the maximum size of the internal shared string cache, in kilobyte. (minimum: 8 KB)
     * Note that this is a soft limit; Depending on circumstances, it might be exceeded by a few byte/kilobyte.
     *
     * @param   int $new_max_size
     *
     * @throws  InvalidArgumentException
     */
    public function setCacheSizeKilobyte($new_max_size)
    {
        if (!is_numeric($new_max_size) || $new_max_size < 8) {
            throw new InvalidArgumentException('Invalid parameter value; Expected a positive number equal to or greater than 8.');
        }
        $this->cache_size_kilobyte = (int)$new_max_size;
    }

    /**
     * Enable/disable the creation of new temporary files for the purpose of optimizing shared string seek performance.
     *
     * @param   bool    $new_use_files_value
     *
     * @throws  InvalidArgumentException
     */
    public function setUseOptimizedFiles($new_use_files_value)
    {
        if (!is_bool($new_use_files_value)) {
            throw new InvalidArgumentException('Invalid parameter value; Expected a boolean.');
        }
        $this->use_optimized_files = $new_use_files_value;
    }

    /**
     * Set the amount of entries to be stored per single optimized shared string file.
     * Adjusting this value has no effect if the creation of optimized shared string files is disabled.
     *
     * @param   int $new_entry_count
     *
     * @throws  InvalidArgumentException
     */
    public function setOptimizedFileEntryCount($new_entry_count)
    {
        if (!is_numeric($new_entry_count) || $new_entry_count <= 0) {
            throw new InvalidArgumentException('Invalid parameter value; Expected a positive number.');
        }
        $this->optimized_file_entry_count = $new_entry_count;
    }

    /**
     * Enable/disable keeping file pointers to shared string files open to achieve more efficient file reads.
     *
     * @param   bool    $new_keep_file_pointers_value
     *
     * @throws  InvalidArgumentException
     */
    public function setKeepFileHandles($new_keep_file_pointers_value)
    {
        if (!is_bool($new_keep_file_pointers_value)) {
            throw new InvalidArgumentException('Invalid parameter value; Expected a boolean.');
        }
        $this->keep_file_handles = $new_keep_file_pointers_value;
    }

    /**
     * @return bool
     */
    public function getUseCache()
    {
        return $this->use_cache;
    }

    /**
     * @return int
     */
    public function getCacheSizeKilobyte()
    {
        return $this->cache_size_kilobyte;
    }

    /**
     * @return bool
     */
    public function getUseOptimizedFiles()
    {
        return $this->use_optimized_files;
    }

    /**
     * @return int
     */
    public function getOptimizedFileEntryCount()
    {
        return $this->optimized_file_entry_count;
    }

    /**
     * @return bool
     */
    public function getKeepFileHandles()
    {
        return $this->keep_file_handles;
    }
}
