<?php
namespace Aspera\Spreadsheet\XLSX;
use XMLReader;

class SharedStrings
{
    /**
     * Number of shared strings that can be reasonably cached, i.e., that aren't read from file but stored in memory.
     *    If the total number of shared strings is higher than this, caching is not used.
     *    If this value is null, shared strings are cached regardless of amount.
     *    With large shared string caches there are huge performance gains, however a lot of memory could be used which
     *    can be a problem, especially on shared hosting.
     */
    const SHARED_STRING_CACHE_LIMIT = 500000;

    // Shared strings file
    /**
     * @var string Path to shared strings XML file
     */
    private $SharedStringsPath = false;

    /**
     * @var XMLReader XML reader object for the shared strings XML file
     */
    private $SharedStrings = false;

    /**
     * @var array Shared strings cache, if the number of shared strings is low enough
     */
    private $SharedStringCache = array();

    private $SharedStringCount = 0;

    private $SharedStringIndex = 0;

    private $LastSharedStringValue = null;

    private $SSOpen = false;

    private $SSForwarded = false;

    /**
     * SharedStrings constructor.
     * @param $sharedStringsPath
     */
    public function __construct($sharedStringsPath)
    {
        $this->SharedStringsPath = $sharedStringsPath;
        if (is_readable($this->SharedStringsPath)) {
            $this->SharedStrings = new XMLReader;
            $this->SharedStrings->open($this->SharedStringsPath);
            $this->prepareSharedStringCache();
        }
    }

    /**
     * Creating shared string cache if the number of shared strings is acceptably low (or there is no limit on the
     * amount
     */
    private function prepareSharedStringCache()
    {
        while ($this->SharedStrings->read()) {
            if ($this->SharedStrings->name == 'sst') {
                $this->SharedStringCount = $this->SharedStrings->getAttribute('count');
                break;
            }
        }

        if (!$this->SharedStringCount || (self::SHARED_STRING_CACHE_LIMIT < $this->SharedStringCount && self::SHARED_STRING_CACHE_LIMIT !== null)) {
            return false;
        }

        $CacheIndex = 0;
        $CacheValue = '';
        while ($this->SharedStrings->read()) {
            switch ($this->SharedStrings->name) {
                case 'si':
                    if ($this->SharedStrings->nodeType == XMLReader::END_ELEMENT) {
                        $this->SharedStringCache[$CacheIndex] = $CacheValue;
                        $CacheIndex++;
                        $CacheValue = '';
                    }
                    break;
                case 't':
                    if ($this->SharedStrings->nodeType == XMLReader::END_ELEMENT) {
                        continue;
                    }
                    $CacheValue .= $this->SharedStrings->readString();
                    break;
            }
        }

        $this->SharedStrings->close();

        return true;
    }

    /**
     * Retrieves a shared string value by its index
     *
     * @param int Shared string index
     *
     * @return string Value
     */
    public function getSharedString($Index)
    {
        if ((self::SHARED_STRING_CACHE_LIMIT === null || self::SHARED_STRING_CACHE_LIMIT > 0) && !empty($this->SharedStringCache)) {
            if (isset($this->SharedStringCache[$Index])) {
                return $this->SharedStringCache[$Index];
            } else {
                return '';
            }
        }

        // If the desired index is before the current, rewind the XML
        if ($this->SharedStringIndex > $Index) {
            $this->SSOpen = false;
            $this->SharedStrings->close();
            $this->SharedStrings->open($this->SharedStringsPath);
            $this->SharedStringIndex = 0;
            $this->LastSharedStringValue = null;
            $this->SSForwarded = false;
        }

        // Finding the unique string count (if not already read)
        if ($this->SharedStringIndex == 0 && !$this->SharedStringCount) {
            while ($this->SharedStrings->read()) {
                if ($this->SharedStrings->name == 'sst') {
                    $this->SharedStringCount = $this->SharedStrings->getAttribute('uniqueCount');
                    break;
                }
            }
        }

        // If index of the desired string is larger than possible, don't even bother.
        if ($this->SharedStringCount && ($Index >= $this->SharedStringCount)) {
            return '';
        }

        // If an index with the same value as the last already fetched is requested
        // (any further traversing the tree would get us further away from the node)
        if (($Index == $this->SharedStringIndex) && ($this->LastSharedStringValue !== null)) {
            return $this->LastSharedStringValue;
        }

        // Find the correct <si> node with the desired index
        while ($this->SharedStringIndex <= $Index) {
            // SSForwarded is set further to avoid double reading in case nodes are skipped.
            if ($this->SSForwarded) {
                $this->SSForwarded = false;
            } else {
                $ReadStatus = $this->SharedStrings->read();
                if (!$ReadStatus) {
                    break;
                }
            }

            if ($this->SharedStrings->name == 'si') {
                if ($this->SharedStrings->nodeType == XMLReader::END_ELEMENT) {
                    $this->SSOpen = false;
                    $this->SharedStringIndex++;
                } else {
                    $this->SSOpen = true;

                    if ($this->SharedStringIndex < $Index) {
                        $this->SSOpen = false;
                        $this->SharedStrings->next('si');
                        $this->SSForwarded = true;
                        $this->SharedStringIndex++;
                        continue;
                    } else {
                        break;
                    }
                }
            }
        }

        $Value = '';

        // Extract the value from the shared string
        if ($this->SSOpen && ($this->SharedStringIndex == $Index)) {
            while ($this->SharedStrings->read()) {
                switch ($this->SharedStrings->name) {
                    case 't':
                        if ($this->SharedStrings->nodeType == XMLReader::END_ELEMENT) {
                            continue;
                        }
                        $Value .= $this->SharedStrings->readString();
                        break;
                    case 'si':
                        if ($this->SharedStrings->nodeType == XMLReader::END_ELEMENT) {
                            $this->SSOpen = false;
                            $this->SSForwarded = true;
                            break 2;
                        }
                        break;
                }
            }
        }

        if ($Value) {
            $this->LastSharedStringValue = $Value;
        }

        return $Value;
    }

    public function close()
    {
        if ($this->SharedStrings && $this->SharedStrings instanceof XMLReader) {
            $this->SharedStrings->close();
            unset($this->SharedStrings);
        }
        unset($this->SharedStringsPath);
    }
}