<?php

namespace Aspera\Spreadsheet\XLSX\Tests;

use Aspera\Spreadsheet\XLSX\Reader;
use Aspera\Spreadsheet\XLSX\ReaderConfiguration;
use DateTime;
use DateTimeZone;
use Exception;
use PHPUnit\Framework\TestCase;

require_once __DIR__ . '/../vendor/autoload.php';

class CellFormatConfigTest extends TestCase
{
    const TEST_FILE = __DIR__ . '/input_files/cell_format_config_test.xlsx';

    /**
     * @return array[]
     *
     * @throws Exception
     */
    public function dataProviderForFormatConfiguration()
    {
        $return_base = array('', '08-20-17', '22:30', '8/20/17 22:30', '25.50%', '0.50%');

        return array(
            'enforce date formats'    => array(
                (new ReaderConfiguration())
                    ->setForceDateFormat('Y-m_d')
                    ->setForceTimeFormat('')
                    ->setForceDateTimeFormat('Y_m-d H_i:s'),
                array_replace(
                    $return_base,
                    array(
                        1 => '2017-08_20',
                        2 => '',
                        3 => '2017_08-20 22_30:00'
                    )
                )
            ),
            'return datetime objects' => array(
                (new ReaderConfiguration())
                    ->setReturnDateTimeObjects(true)
                    ->setForceDateTimeFormat('Y-m-d His'), // Should have no effect, overruled by returnDateTimeObjects.
                array_replace(
                    $return_base,
                    array(
                        1 => new DateTime('2017-08-20', new DateTimeZone('UTC')),
                        2 => new DateTime('1899-12-31 22:30:00', new DateTimeZone('UTC')), // From base date: 1900-01-00 <- note the 0 day
                        3 => new DateTime('2017-08-20 22:30:00', new DateTimeZone('UTC'))
                    )
                )
            ),
            'use custom format'       => array(
                (new ReaderConfiguration())->setCustomFormats(array(
                    14 => '"It is the year" yyyy',
                    20 => 'His',
                    22 => 'yyyy mm dd His'
                )),
                array_replace(
                    $return_base,
                    array(
                        1 => 'It is the year 2017',
                        2 => '223000',
                        3 => '2017 08 20 223000'
                    )
                )
            ),
            'return unformatted'      => array(
                (new ReaderConfiguration())->setReturnUnformatted(true),
                array_replace(
                    $return_base,
                    array(
                        1 => '42967',
                        2 => '0.9375',
                        3 => '42967.9375',
                        4 => '25.5',
                        5 => '0.5'
                    )
                )
            ),
            'percentage as decimals'  => array(
                (new ReaderConfiguration())
                    ->setReturnPercentageDecimal(true),
                array_replace(
                    $return_base,
                    array(
                        4 => '0.255',
                        5 => '5.0000000000000001E-3' // Common Excel problem. Percentage values are divided by 100, which can cause floating point issues.
                    )
                )
            ),
            'return unformatted is overruled by datetime force format' => array(
                (new ReaderConfiguration())
                    ->setReturnUnformatted(true)
                    ->setForceDateTimeFormat('Y-m-d His'),
                array_replace(
                    $return_base,
                    array(
                        1 => '42967', // date value, not datetime value. Therefore, returnUnformatted applies here.
                        2 => '0.9375', // Same situation for time values.
                        3 => '2017-08-20 223000', // datetime value, forceDateTimeFormat applies here.
                        4 => '25.5',
                        5 => '0.5'
                    )
                )
            ),
            'return unformatted is overruled by return datetime objects' => array(
                (new ReaderConfiguration())
                    ->setReturnUnformatted(true)
                    ->setReturnDateTimeObjects(true),
                array_replace(
                    $return_base,
                    array(
                        1 => new DateTime('2017-08-20', new DateTimeZone('UTC')),
                        2 => new DateTime('1899-12-31 22:30:00', new DateTimeZone('UTC')), // From base date: 1900-01-00 <- note the 0 day
                        3 => new DateTime('2017-08-20 22:30:00', new DateTimeZone('UTC')),
                        4 => '25.5',
                        5 => '0.5'
                    )
                )
            )
        );
    }

    /**
     * Ensures that ReaderConfiguration values are considered and interpreted as expected.
     *
     * @dataProvider dataProviderForFormatConfiguration
     *
     * @param ReaderConfiguration $config
     * @param array               $expected
     *
     * @throws Exception
     */
    public function testCellFormatConfiguration($config, $expected)
    {
        $reader = new Reader($config);
        $reader->open(self::TEST_FILE);
        $output_rows = array();
        foreach ($reader as $row) {
            $output_rows[] = $row[2]; // [2] = only the third column is of interest
        }
        $reader->close();

        // Special handling for DateTime values, as they cannot be compared with assertSame directly.
        foreach (array_keys($expected) as $k) {
            if (($expected[$k] instanceof DateTime) !== ($output_rows[$k] instanceof DateTime)) {
                self::fail('Values at position [' . $k . '] do not share the same type (DateTime).');
            }
            if ($expected[$k] instanceof DateTime) {
                $expected[$k] = $expected[$k]->format('Y-m-d H:i:s');
                $output_rows[$k] = $output_rows[$k]->format('Y-m-d H:i:s');
            }
        }

        self::assertSame(
            $expected,
            $output_rows,
            'ReaderConfiguration for cell content formatting was not considered and interpreted as expected.'
        );
    }
}
