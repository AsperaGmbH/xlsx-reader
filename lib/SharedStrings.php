<?php
namespace Aspera\Spreadsheet\XLSX;

use XMLReader;

class SharedStrings
{
    /**
     * Number of shared strings that can be reasonably cached, i.e., that aren't read from file but stored in memory.
     * If the total number of shared strings is higher than this, caching is not used.
     * If this value is null, shared strings are cached regardless of amount.
     * With large shared string caches there are huge performance gains, however a lot of memory could be used which
     * can be a problem, especially on shared hosting.
     */
    const SHARED_STRING_CACHE_LIMIT = 500000;

    /**
     * @var string Path to shared strings XML file
     */
    private $shared_strings_path = false;

    /**
     * @var XMLReader XML reader object for the shared strings XML file
     */
    private $shared_strings = false;

    /**
     * @var array Shared strings cache, if the number of shared strings is low enough
     */
    private $shared_string_cache = array();

    private $shared_string_count = 0;

    private $shared_string_index = 0;

    private $last_shared_string_value = null;

    private $reader_is_within_shared_strings_tag = false;

    private $shared_strings_forwarded = false;

    /**
     * SharedStrings constructor.
     * @param $shared_strings_path
     */
    public function __construct($shared_strings_path)
    {
        $this->shared_strings_path = $shared_strings_path;
        if (is_readable($this->shared_strings_path)) {
            $this->shared_strings = new XMLReader;
            $this->shared_strings->open($this->shared_strings_path);
            $this->prepareSharedStringCache();
        }
    }

    /**
     * Creating shared string cache if the number of shared strings is acceptably low (or there is no limit on the
     * amount
     */
    private function prepareSharedStringCache()
    {
        while ($this->shared_strings->read()) {
            if ($this->shared_strings->name == 'sst') {
                $this->shared_string_count = $this->shared_strings->getAttribute('count');
                break;
            }
        }

        if (!$this->shared_string_count || (self::SHARED_STRING_CACHE_LIMIT < $this->shared_string_count && self::SHARED_STRING_CACHE_LIMIT !== null)) {
            return false;
        }

        $cache_index = 0;
        $cache_value = '';
        while ($this->shared_strings->read()) {
            switch ($this->shared_strings->name) {
                case 'si':
                    if ($this->shared_strings->nodeType == XMLReader::END_ELEMENT) {
                        $this->shared_string_cache[$cache_index] = $cache_value;
                        $cache_index++;
                        $cache_value = '';
                    }
                    break;
                case 't':
                    if ($this->shared_strings->nodeType == XMLReader::END_ELEMENT) {
                        continue;
                    }
                    $cache_value .= $this->shared_strings->readString();
                    break;
            }
        }

        $this->shared_strings->close();

        return true;
    }

    /**
     * Retrieves a shared string value by its index
     *
     * @param int Shared string index
     *
     * @return string Value
     */
    public function getSharedString($index)
    {
        if ((self::SHARED_STRING_CACHE_LIMIT === null || self::SHARED_STRING_CACHE_LIMIT > 0) && !empty($this->shared_string_cache)) {
            if (isset($this->shared_string_cache[$index])) {
                return $this->shared_string_cache[$index];
            } else {
                return '';
            }
        }

        // If the desired index is before the current, rewind the XML
        if ($this->shared_string_index > $index) {
            $this->reader_is_within_shared_strings_tag = false;
            $this->shared_strings->close();
            $this->shared_strings->open($this->shared_strings_path);
            $this->shared_string_index = 0;
            $this->last_shared_string_value = null;
            $this->shared_strings_forwarded = false;
        }

        // Finding the unique string count (if not already read)
        if ($this->shared_string_index == 0 && !$this->shared_string_count) {
            while ($this->shared_strings->read()) {
                if ($this->shared_strings->name == 'sst') {
                    $this->shared_string_count = $this->shared_strings->getAttribute('uniqueCount');
                    break;
                }
            }
        }

        // If index of the desired string is larger than possible, don't even bother.
        if ($this->shared_string_count && ($index >= $this->shared_string_count)) {
            return '';
        }

        // If an index with the same value as the last already fetched is requested
        // (any further traversing the tree would get us further away from the node)
        if (($index == $this->shared_string_index) && ($this->last_shared_string_value !== null)) {
            return $this->last_shared_string_value;
        }

        // Find the correct <si> node with the desired index
        while ($this->shared_string_index <= $index) {
            // SSForwarded is set further to avoid double reading in case nodes are skipped.
            if ($this->shared_strings_forwarded) {
                $this->shared_strings_forwarded = false;
            } else if (!$this->shared_strings->read()) {
                break;
            }

            if ($this->shared_strings->name == 'si') {
                if ($this->shared_strings->nodeType == XMLReader::END_ELEMENT) {
                    $this->reader_is_within_shared_strings_tag = false;
                    $this->shared_string_index++;
                } else {
                    $this->reader_is_within_shared_strings_tag = true;

                    if ($this->shared_string_index < $index) {
                        $this->reader_is_within_shared_strings_tag = false;
                        $this->shared_strings->next('si');
                        $this->shared_strings_forwarded = true;
                        $this->shared_string_index++;
                        continue;
                    } else {
                        break;
                    }
                }
            }
        }

        $value = '';

        // Extract the value from the shared string
        if ($this->reader_is_within_shared_strings_tag && ($this->shared_string_index == $index)) {
            while ($this->shared_strings->read()) {
                switch ($this->shared_strings->name) {
                    case 't':
                        if ($this->shared_strings->nodeType == XMLReader::END_ELEMENT) {
                            continue;
                        }
                        $value .= $this->shared_strings->readString();
                        break;
                    case 'si':
                        if ($this->shared_strings->nodeType == XMLReader::END_ELEMENT) {
                            $this->reader_is_within_shared_strings_tag = false;
                            $this->shared_strings_forwarded = true;
                            break 2;
                        }
                        break;
                }
            }
        }

        if ($value) {
            $this->last_shared_string_value = $value;
        }

        return $value;
    }

    public function close()
    {
        if ($this->shared_strings && $this->shared_strings instanceof XMLReader) {
            $this->shared_strings->close();
            unset($this->shared_strings);
        }
        unset($this->shared_strings_path);
    }
}
