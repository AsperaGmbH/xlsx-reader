<?php

namespace Aspera\Spreadsheet\XLSX\Tests;

require_once __DIR__ . '/../vendor/autoload.php';

use Exception;
use PHPUnit\Framework\TestCase;
use Aspera\Spreadsheet\XLSX\Reader;

/**
 * Ensure that the reader can parse date/time data correctly and can output it given the user's preferences.
 *
 * @author Aspera GmbH
 */
class DateFormatTest extends TestCase
{
    const TEST_FILE = 'input_files/date_format_test.xlsx';

    /**
     * Enforce some date/time formats and check whether the reader acts accordingly.
     *
     * @throws Exception
     */
    public function testEnforcedDateFormats()
    {
        $reader = new Reader(array(
            'ForceDateFormat'     => 'Y-m_d',
            'ForceTimeFormat'     => 'H:i_s',
            'ForceDateTimeFormat' => 'Y_m-d H_i:s'
        ));
        $reader->open(self::TEST_FILE);
        $output_rows = array();
        while ($row = $reader->next()) {
            $output_rows[] = $row[2]; // [2] = only the third column is of interest
        }
        $reader->close();

        self::assertSame(
            array('', '2017-08_20', '15:22_00', '2017_08-20 15_22:00'),
            $output_rows,
            'Enforced Date/Time formats could not be confirmed; Date/Time Format enforcement might be broken.'
            . ' Retrieved row contents: [' . implode('|', $output_rows) . ']'
        );
    }
}
