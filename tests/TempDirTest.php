<?php

namespace Aspera\Spreadsheet\XLSX\Tests;

require_once __DIR__ . '/../vendor/autoload.php';

use Exception;
use Aspera\Spreadsheet\XLSX\Reader;
use Aspera\Spreadsheet\XLSX\ReaderConfiguration;
use PHPUnit\Framework\TestCase as PHPUnitTestCase;

/** Test temporary work directory functionality */
class TempDirTest extends PHPUnitTestCase
{
    /** @var string FILE_PATH XLSX File to be used for testing. */
    const FILE_PATH = __DIR__ . '/input_files/iterator_test.xlsx';

    /** @var string TEMP_DIR_NAME Directory name of the temporary work directory. */
    const TEMP_DIR_NAME = 'temp_new_folder';

    /** @var Reader $reader Reader used by all tests in this test class. */
    private static $reader;

    /**
     * Create temporary work directory and initialize the reader, creating temporary files in the process.
     *
     * @throws Exception
     */
    public function testPrepare()
    {
        // Create target directory and configure reader to use it.
        $temp_dir_path = self::getTempDirPath();
        @mkdir($temp_dir_path);
        self::$reader = new Reader((new ReaderConfiguration())
            ->setTempDir($temp_dir_path)
        );
        self::$reader->open(self::FILE_PATH);
    }

    /**
     * Make sure that the TempDir option was properly considered by the reader.
     *
     * @depends testPrepare
     *
     * @throws Exception
     */
    public function testOptionTempDir()
    {
        $temp_dir_path = self::getTempDirPath();

        // Check directory contents. The reader should have created files in it.
        clearstatcache();
        $file_list = @scandir($temp_dir_path, SCANDIR_SORT_NONE);

        // 2 entries will always exist: ., ..; We expect there to be MORE entries than that, if the reader used the directory.
        self::assertGreaterThan(
            2,
            count($file_list),
            'The configured TempDir [' . $temp_dir_path . '] was not used by the reader.'
        );
    }

    /**
     * Make sure that all temporary work files are deleted after processing of the target file.
     *
     * @depends testOptionTempDir
     *
     * @throws Exception
     */
    public function testDeletionTemporaryFiles()
    {
        $temp_dir_path = self::getTempDirPath();

        // Close -> Invoke destructor; This should delete all files from TempDir again.
        self::$reader->close();

        // After destruction, temporary directory must be totally emptied.
        clearstatcache();
        $file_list = @scandir($temp_dir_path, SCANDIR_SORT_NONE);

        // Only 2 entries are expected: ., ..; Any further entries indicate a cleanup failure.
        self::assertCount(2, $file_list, 'The configured TempDir [' . $temp_dir_path . '] was not emptied.');
    }

    /**
     * Remove the temporary work directory.
     *
     * @throws Exception
     */
    public function testCleanup()
    {
        $temp_dir_path = self::getTempDirPath();
        if (file_exists($temp_dir_path)) {
            @rmdir($temp_dir_path);
        }
    }

    /**
     * Return temporary work path used for this test. Already includes trailing slash.
     *
     * @return string
     * @throws Exception
     */
    private static function getTempDirPath()
    {
        $temp_dir_path = sys_get_temp_dir();
        if (!is_string($temp_dir_path)) {
            throw new Exception(
                'Path to temporary work directory could not be determined. sys_get_temp_dir() returned invalid value.'
            );
        }
        if ($temp_dir_path !== '') {
            $temp_dir_path = str_replace(chr(92), '/', $temp_dir_path);
            if (mb_substr($temp_dir_path, -1) !== '/') {
                $temp_dir_path .= '/';
            }
        }
        return $temp_dir_path . self::TEMP_DIR_NAME;
    }
}
