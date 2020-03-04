<?php

namespace Aspera\Spreadsheet\XLSX\Tests;

require_once __DIR__ . '/../vendor/autoload.php';

use Exception;
use PHPUnit\Framework\TestCase;
use Aspera\Spreadsheet\XLSX\NumberFormat;

/** Check whether number formatting works properly for all sorts of custom formats. */
class CustomNumberFormatTest extends TestCase
{
    /**
     * Tests if the "general" format is correctly identified and applied.
     *
     * @dataProvider provideValuesForGeneralFormat
     *
     * @param  string $value
     *
     * @throws Exception
     */
    public function testGeneralFormat($value)
    {
        $cell_format = new NumberFormat();
        $cell_format->injectXfNumFmtIds(
            array(456 => 0) // Fake xf
        );
        $actual_output = $cell_format->formatValue($value, 456); // Apply numFmt of fake xf
        self::assertSame(
            $value, // General format should never change anything about the input value.
            $actual_output,
            'Unexpected formatting result for value [' . $value . '] with general format (number format 0).'
        );
    }

    /**
     * @return array
     */
    public function provideValuesForGeneralFormat()
    {
        return array(
            'number'  => array('123'),
            'decimal' => array('123.45'),
            'string'  => array('abc')
        );
    }

    /**
     * @dataProvider provideFormats
     *
     * @param  string $value
     * @param  string $format
     * @param  string $expected_output
     *
     * @throws Exception
     */
    public function testFormat($value, $format, $expected_output)
    {
        $cell_format = new NumberFormat();
        $cell_format->injectXfNumFmtIds(
            array(456 => 789) // Fake xf
        );
        $cell_format->injectNumberFormats(
            array(789 => $format) // Fake numFmt
        );
        $actual_output = $cell_format->formatValue($value, 456); // Apply numFmt of fake xf
        self::assertSame(
            $expected_output,
            $actual_output,
            'Unexpected formatting result for value [' . $value . '] with format [' . $format . '].'
        );
    }

    /**
     * @return array
     */
    public function provideFormats()
    {
        // Note that language info ( [$€-THIS_RANDOM_INTEGER_RIGHT_HERE] ) is always ignored.
        return array(
            'simple format, positive value'                                 => array(
                '123',
                '0.00',
                '123.00'
            ),
            'simple format, negative value'                                 => array(
                '-123',
                '0.00',
                '-123.00'
            ),
            'simple zero format'                                            => array(
                '0',
                '0',
                '0'
            ),
            'negative value format'                                         => array(
                '-123',
                '0;0.0000',
                '123.0000'
            ),
            'negative value format, do not display'                         => array(
                '-50',
                '0.00;;',
                ''
            ),
            'decimal format, without prefix'                                => array(
                '123.45',
                '.?',
                '123.5'
            ),
            'decimal format does not have to make sense'                    => array(
                '1.2',
                '0.0#########0?',
                '1.20 '
            ),
            'fill empty spaces, positive value'                             => array(
                '1.2',
                '???.???',
                '  1.2  '
            ),
            'fill empty spaces, negative value'                             => array(
                '-1.2',
                '???.???',
                '-  1.2  '
            ),
            'optional spaces, positive value'                               => array(
                '1.2',
                '###.###',
                '1.2'
            ),
            'optional spaces, negative value'                               => array(
                '-1.2',
                '###.###',
                '-1.2'
            ),
            'thousands seperator placement is irrelevant'                   => array(
                '1234567.89',
                '#####,,###########0.000,',
                '1,234.568'
            ),
            'thousands separator may force prepended zeroes'                => array(
                '123',
                '0,000',
                '0,123'
            ),
            'scaling, positive value'                                       => array(
                '123456789',
                '0.00,,',
                '123.46'
            ),
            'scaling, negative value'                                       => array(
                '-123456789',
                '0.00,,',
                '-123.46'
            ),
            'scaling, with decimals in value'                               => array(
                '123456.789',
                '0,',
                '123'
            ),
            'scaling, without decimals in format'                           => array(
                '123456789',
                '0,,',
                '123'
            ),
            'currency upfront, positive value'                              => array(
                '123',
                '[$€-805]0.00',
                '€123.00'
            ),
            'currency upfront, negative value'                              => array(
                '-123',
                '[$€-805]0.00',
                '-€123.00'
            ),
            'currency upfront, explicit negative format'                    => array(
                '-123',
                '0;[$€-805]0.00" (minus)"',
                '€123.00 (minus)'
            ),
            'currency, without currency string'                             => array(
                '123',
                '[$-805]0.00',
                '123.00'
            ),
            'currency, without language id'                                 => array(
                '123',
                '[$€]0.00',
                '€123.00'
            ),
            'quoted/escaped semicolons'                                     => array(
                '-123',
                '0";"\;"0;0""0\;0"\";0.0000',
                '123.0000'
            ),
            'quoted sections, preceeding a positive value'                  => array(
                '123',
                '"[POSITIVE] "0.00;"[NEGATIVE] "0.00',
                '[POSITIVE] 123.00'
            ),
            'quoted sections, preceeding a negative value'                  => array(
                '-123',
                '"[POSITIVE] "0.00;"[NEGATIVE] "0.00',
                '[NEGATIVE] 123.00'
            ),
            'escaping does not work within square brackets'                 => array(
                '100',
                '[$CU\\R\\\\RE"N[CY-]0.00',
                'CU\\R\\\\RE"N[CY100.00'
            ),
            'handling of spaces in format string'                           => array(
                'demo',
                '" t" "e ""x t"" "":" @',
                ' t e x t : demo'
            ),
            'date format, with language code'                               => array(
                '123',
                '[$-123]DD-MM-YYYY',
                '02-05-1900'
            ),
            'date format, prefixed with extras'                             => array(
                '123',
                '"before "DD-MM-YYYY',
                'before 02-05-1900'
            ),
            'time in 12h'                                                   => array(
                '0.75',
                'hh:mm AM/PM',
                '06:00 PM'
            ),
            'time in 24h, but with a confusing string'                      => array(
                '0.75',
                'hh:mm" AM/PM"',
                '18:00 AM/PM'
            ),
            '1 conditional, no default, invalid numeric'                    => array(
                '0',
                '[<0]0',
                '########'
            ),
            '1 conditional, no default, invalid text'                       => array(
                'test',
                '[<0]0',
                'test'
            ),
            '1 conditional, with NUMBER-ONLY default, invalid numeric'      => array(
                '0',
                '[<0]0;"["0"]"',
                '[0]'
            ),
            '1 conditional, with NUMBER-ONLY default, invalid text'         => array(
                'test',
                '[<0]0;"["0"]"',
                'test'
            ),
            '1 conditional, with TEXT-ONLY default, invalid numeric'        => array(
                '0',
                '[<0]0;"["@"]"',
                '########'
            ),
            '1 conditional, with TEXT-ONLY default, invalid text'           => array(
                'test',
                '[<0]0;"["@"]"',
                '[test]'
            ),
            '1 conditional, with 2 defaults, with a number'                 => array(
                '0.123',
                '[>50]0;.0;"["@"]"',
                '.1'
            ),
            '1 conditional, with 2 defaults, with text'                     => array(
                'test',
                '[>50]0;.0;"["@"]"',
                '[test]'
            ),
            '2 conditionals, no default, invalid numeric'                   => array(
                '0',
                '[<0]0;[>0]0',
                '########'
            ),
            '2 conditionals, no default, invalid text'                      => array(
                'test',
                '[<0]0;[>0]0',
                'test'
            ),
            '2 conditionals, with 1 default, invalid numeric'               => array(
                '0',
                '[<0]0;[>0]0;0.000',
                '0.000'
            ),
            '2 conditionals, with 1 default, invalid text'                  => array(
                'test',
                '[<0]0;[>0]0;0.000',
                'test'
            ),
            '2 conditionals, with 2 defaults, invalid numeric'              => array(
                '0',
                '[<0]0;[>0]0;0.000;"["@"]"',
                '0.000'
            ),
            '2 conditionals, with 2 defaults, invalid text'                 => array(
                'test',
                '[<0]0;[>0]0;0.000;"["@"]"',
                '[test]'
            ),
            'conditional in 2nd section, matching value'                    => array(
                '50',
                '"1st "0;[=50]"2nd "0;"3rd "0',
                '1st 50'
            ),
            'conditional in 1st section, matching value'                    => array(
                '50',
                '[=50]"1st "0;"2nd "0;"3rd "0',
                '1st 50'
            ),
            'conditional in 1st section, positive value'                    => array(
                '51',
                '[=50]"1st "0;"2nd "0;"3rd "0',
                '3rd 51'
            ),
            'conditional in 1st section, negative value'                    => array(
                '-50',
                '[=50]"1st "0;"2nd "0;"3rd "0',
                '2nd 50'
            ),
            'conditional in 2nd section, negative value'                    => array(
                '-50',
                '"1st "0;[=50]"2nd "0;"3rd "0',
                '-3rd 50'
            ),
            'conditional + color + whitespace'                              => array(
                '-50',
                ' [red]   [=-50]  "1st "0;"2nd "0',
                '      1st 50'
                // Note: Exact behavior differs between applications. Some keep the spaces, some remove them.
            ),
            'percentage signs change output values'                         => array(
                '0.12',
                '0.00%',
                '12.00%'
            ),
            'percentage signs in quotes don\'t change output values'        => array(
                '0.12',
                '0.00"%"',
                '0.12%'
            ),
            'fraction without whole values'                                 => array(
                '1.2',
                '?/#',
                '6/5'
            ),
            'fraction with a negative whole value'                          => array(
                '-5',
                '0/0',
                '-5'
            ),
            'fraction, negative, with a post-decimal value starting with 0' => array(
                '-2.025',
                '0/0',
                '-81/40'
            ),
            'fraction with unneccessary whole values'                       => array(
                '0.2',
                '? #/0',
                '1/5'
            ),
            'fraction with whole values and whitespace'                     => array(
                '1.2',
                ' # 0/? ',
                ' 1 1/5 '
            ),
            'fractions can be percent values'                               => array(
                '0.005',
                '0/0%',
                '1/2%'
            )
        );
    }
}
