<?php

namespace Aspera\Spreadsheet\XLSX\Tests;

require_once __DIR__ . '/../vendor/autoload.php';

use Exception;
use ReflectionClass;
use ReflectionException;
use Aspera\Spreadsheet\XLSX\Reader;
use Aspera\Spreadsheet\XLSX\ReaderConfiguration;
use Aspera\Spreadsheet\XLSX\SharedStrings;
use Aspera\Spreadsheet\XLSX\SharedStringsConfiguration;
use PHPUnit\Framework\TestCase;

/** Test shared/inline string behaviour and configuration. */
class SharedStringsTest extends TestCase
{
    /** @var string FILE_PATH Path to the test file. */
    const FILE_PATH = __DIR__ . '/input_files/shared_strings_test.xlsx';

    /** @var int SHARED_STRING_ENTRY_COUNT Total number of shared strings in the test file's shared string list. */
    const SHARED_STRING_ENTRY_COUNT = 25005;

    /** @var int CACHE_MAX_SIZE_KB Required cache size in KB to store the entire shared strings table from the test file in memory. */
    const CACHE_MAX_SIZE_KB = 2048; // 2 MB should be enough for the entire shared string file in all supported PHP versions

    /**
     * @return array
     */
    public function dataProviderForTestValues()
    {
        return array(
            'via cache'          => array(true, true),
            'via optimized file' => array(false, true),
            'via original file'  => array(false, false)
        );
    }

    /**
     * Check basic shared/inline string extraction using the configured execution path
     *
     * @dataProvider dataProviderForTestValues
     *
     * @param   bool $use_cache
     * @param   bool $use_optimized_files
     *
     * @throws  Exception
     */
    public function testValues($use_cache, $use_optimized_files)
    {
        $xlsx_reader = new Reader((new ReaderConfiguration())
            ->setSharedStringsConfiguration((new SharedStringsConfiguration())
                ->setUseCache($use_cache)
                ->setUseOptimizedFiles($use_optimized_files)
                ->setCacheSizeKilobyte(self::CACHE_MAX_SIZE_KB))
        );

        $xlsx_reader->open(self::FILE_PATH);

        // Check values; A1 contains a shared string, B1 contains an inline string.
        $test_row = $xlsx_reader->current();
        self::assertSame('shared string', $test_row[0],
            'Could not read shared string. Found value: [' . $test_row[0] . ']');
        self::assertSame('inline string', $test_row[1],
            'Could not read inline string. Found value: [' . $test_row[1] . ']');
        $xlsx_reader->close();
    }

    /**
     * @return array
     */
    public function dataProviderForTestMemoryConfiguration()
    {
        return array(
            'no cache'    => array(false, true),
            'small cache' => array(true, false),
            'large cache' => array(true, true)
        );
    }

    /**
     * Check if the use_cache/max_cache_size configuration values are properly respected.
     *
     * @dataProvider dataProviderForTestMemoryConfiguration
     *
     * @param  bool $use_cache
     * @param  bool $use_large_cache
     *
     * @throws Exception
     */
    public function testMemoryConfiguration($use_cache, $use_large_cache)
    {
        // Pick configuration values based on test data set
        if ($use_large_cache) {
            $cache_size_kb = self::CACHE_MAX_SIZE_KB;

            // 25005 entries in test file + SplFixedArray increases in increments of 100 => expected value: 25100
            $increment = SharedStrings::SHARED_STRING_CACHE_ARRAY_SIZE_STEP;
            $string_count_calc = self::SHARED_STRING_ENTRY_COUNT - 1;
            $min_entry_count = $string_count_calc + ($increment - $string_count_calc % $increment);
            $max_entry_count = $min_entry_count;
        } else {
            $cache_size_kb = 8;

            /* Exact entry counts can differ a bit, based on used PHP version and internal configuration.
             * Hence, use generous min/max range. */
            $min_entry_count = 50;
            $max_entry_count = 2000;
        }

        // Initialize reader
        $xlsx_reader = new Reader((new ReaderConfiguration())
            ->setSharedStringsConfiguration((new SharedStringsConfiguration())
                ->setUseCache($use_cache)
                ->setCacheSizeKilobyte($cache_size_kb)
            )
        );
        $xlsx_reader->open(self::FILE_PATH);

        // Get shared strings cache from shared strings object
        $shared_strings = self::getAccessibleProperty($xlsx_reader, 'shared_strings');
        $shared_strings_cache = self::getAccessibleProperty($shared_strings, 'shared_string_cache');
        $shared_strings_cache_count = count($shared_strings_cache);

        // Check against configured values
        if ($use_cache) {
            if ($shared_strings_cache_count == 0) {
                self::fail('Cache is enabled but contents are empty.');
            }
            if ($shared_strings_cache_count < $min_entry_count) {
                self::fail(
                    'Cache size is lower than expected.'
                    . ' Expected at least ' . $min_entry_count . ' entries.'
                    . ' Actual entry count is ' . $shared_strings_cache_count . '.'
                );
            }
            if ($shared_strings_cache_count > $max_entry_count) {
                self::fail(
                    'Cache size is higher than expected.'
                    . ' Maximum allowed entry count is ' . $max_entry_count . '.'
                    . ' Actual entry count is ' . $shared_strings_cache_count . '.'
                );
            }
        } else {
            if ($shared_strings_cache_count > 0) {
                self::fail('Cache is disabled but still contains contents.');
            }
        }
    }

    /**
     * @return array
     */
    public function dataProviderForTestOptimizedFileConfiguration()
    {
        return array(
            'do not use optimized files' => array(false, false),
            'use small optimized files'  => array(true, false),
            'use large optimized files'  => array(true, true)
        );
    }

    /**
     * Check if optimized shared string files are created/not created on demand
     *
     * @dataProvider dataProviderForTestOptimizedFileConfiguration
     *
     * @param  bool $use_optimized_files
     * @param  bool $use_many_entries_per_file
     *
     * @throws Exception
     */
    public function testOptimizedFileConfiguration($use_optimized_files, $use_many_entries_per_file)
    {
        // Pick configuration values based on test data set
        $entries_per_file = $use_many_entries_per_file ? 5000 : 500;
        $expected_file_count = ceil(self::SHARED_STRING_ENTRY_COUNT / $entries_per_file);

        // Initialize reader
        $xlsx_reader = new Reader((new ReaderConfiguration())
            ->setSharedStringsConfiguration((new SharedStringsConfiguration())
                ->setUseCache(false)
                ->setUseOptimizedFiles($use_optimized_files)
                ->setOptimizedFileEntryCount($entries_per_file)
            )
        );
        $xlsx_reader->open(self::FILE_PATH);

        // Get optimized shared strings file list from shared strings object
        $shared_strings = self::getAccessibleProperty($xlsx_reader, 'shared_strings');
        $prepared_files = self::getAccessibleProperty($shared_strings, 'prepared_shared_string_files');
        $prepared_files_count = count($prepared_files);

        // Check number of created prepared files
        if ($use_optimized_files) {
            if ($prepared_files_count == 0) {
                self::fail('No optimized shared string files were created, despite the configuration requesting it.');
            }
            if ($prepared_files_count != $expected_file_count) {
                self::fail(
                    'The optimized shared string entry count configuration seems to be disregarded.'
                    . ' Expected ' . $expected_file_count . ' files to be created,'
                    . ' found ' . $prepared_files_count . '.'
                );
            }
        } else {
            if ($prepared_files_count > 0) {
                self::fail('Optimized shared string files were created, despite the configuration denying it.');
            }
        }
        $xlsx_reader->close();
    }

    /**
     * From the given object, return the value of the given property, regardless of its access modifier.
     *
     * @param  object $target_object        Object to retrieve the property value from
     * @param  string $target_property_name Name of the property of which the value should be returned
     * @return mixed
     *
     * @throws ReflectionException
     */
    private static function getAccessibleProperty($target_object, $target_property_name)
    {
        $reflection = new ReflectionClass(get_class($target_object));
        $internal_property = $reflection->getProperty($target_property_name);
        $internal_property->setAccessible(true);
        return $internal_property->getValue($target_object);
    }
}
