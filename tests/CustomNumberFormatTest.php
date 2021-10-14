<?php

namespace Aspera\Spreadsheet\XLSX\Tests;

require_once __DIR__ . '/../vendor/autoload.php';

use Exception;
use PHPUnit\Framework\TestCase;
use Aspera\Spreadsheet\XLSX\NumberFormat;
use Aspera\Spreadsheet\XLSX\ReaderConfiguration;

/** Check whether number formatting works properly for all sorts of custom formats. */
class CustomNumberFormatTest extends TestCase
{
    /**
     * Tests if the "general" format is correctly identified and applied.
     *
     * @dataProvider provideValuesForGeneralFormat
     *
     * @param string $value
     * @param string $expected_output
     *
     * @throws Exception
     */
    public function testGeneralFormat($value, $expected_output)
    {
        $cell_format = new NumberFormat(new ReaderConfiguration());
        $cell_format->injectXfNumFmtIds(
            array(456 => 0) // Fake xf
        );
        $actual_output = $cell_format->formatValue($value, 456); // Apply numFmt of fake xf
        self::assertSame(
            $expected_output,
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
            'number'         => array('123', '123'),
            'decimal'        => array('123.45', '123.45'),
            'string'         => array('abc', 'abc'),
            'scientific'     => array('5.12E-03', '0.00512'),
            'rounding error' => array('5.0000000000000001E-3', '0.005') // Common Excel issue
        );
    }

    /**
     * @dataProvider provideFormats
     *
     * @param string $value
     * @param string $format
     * @param string $expected_output
     *
     * @throws Exception
     */
    public function testFormat($value, $format, $expected_output)
    {
        $cell_format = new NumberFormat(new ReaderConfiguration());
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
            'simple format, positive value'                                      => array(
                '123',
                '0.00',
                '123.00'
            ),
            'simple format, negative value'                                      => array(
                '-123',
                '0.00',
                '-123.00'
            ),
            'simple zero format'                                                 => array(
                '0',
                '0',
                '0'
            ),
            'negative value format'                                              => array(
                '-123',
                '0;0.0000',
                '123.0000'
            ),
            'negative value format, do not display'                              => array(
                '-50',
                '0.00;;',
                ''
            ),
            'decimal format, without prefix'                                     => array(
                '123.45',
                '.?',
                '123.5'
            ),
            'decimal format does not have to make sense'                         => array(
                '1.2',
                '0.0#########0?',
                '1.20 '
            ),
            'fill with zeroes'                                                   => array(
                '5',
                '000.0',
                '005.0'
            ),
            'fill empty spaces, positive value'                                  => array(
                '1.2',
                '???.???',
                '  1.2  '
            ),
            'fill empty spaces, negative value'                                  => array(
                '-1.2',
                '???.???',
                '-  1.2  '
            ),
            'optional spaces, positive value'                                    => array(
                '1.2',
                '###.###',
                '1.2'
            ),
            'optional spaces, negative value'                                    => array(
                '-1.2',
                '###.###',
                '-1.2'
            ),
            'decimal format split up by text'                                    => array(
                '123.456',
                'qq0"QQ".\qq#QQ?qq',
                // Note the expansion of the whole number and the rounding here.
                'qq123QQ.qq4QQ6qq'
            ),
            'multiple thousands seperators'                                      => array(
                '123456789',
                '0,0.0',
                '123,456,789.0'
            ),
            'thousands seperator placement is mostly irrelevant'                 => array(
                '1234567.89',
                '#####,,#####?#####0.000,',
                ' 1,234.568'
            ),
            'additional thousands seperator may actually be scaling'             => array(
                '1234',
                '#,#.##,',
                '1.23'
            ),
            'thousands separator may force prepended zeroes'                     => array(
                '123',
                '0,000',
                '0,123'
            ),
            'thousands separator split by text'                                  => array(
                '123456',
                '0"hi"0,0.0',
                '123,4hi56.0'
            ),
            'scaling, positive value'                                            => array(
                '123456789',
                '0.00,,',
                '123.46'
            ),
            'scaling, negative value'                                            => array(
                '-123456789',
                '0.00,,',
                '-123.46'
            ),
            'scaling, with decimals in value'                                    => array(
                '123456.789',
                '0,',
                '123'
            ),
            'scaling, without decimals in format'                                => array(
                '123456789',
                '0,,',
                '123'
            ),
            'scaling, character placed right before decimal symbol'              => array(
                '1234',
                '0,.00',
                '1.23'
            ),
            'scaling, defined twice'                                             => array(
                '123456',
                '0,.00,',
                '0.12'
            ),
            'neither scaling nor thousands separator if in front of a format'    => array(
                '1234',
                ',0.00',
                ',1234.00'
            ),
            'currency upfront, positive value'                                   => array(
                '123',
                '[$€-805]0.00',
                '€123.00'
            ),
            'currency upfront, negative value'                                   => array(
                '-123',
                '[$€-805]0.00',
                '-€123.00'
            ),
            'currency upfront, explicit negative format'                         => array(
                '-123',
                '0;[$€-805]0.00" (minus)"',
                '€123.00 (minus)'
            ),
            'currency, without currency string'                                  => array(
                '123',
                '[$-805]0.00',
                '123.00'
            ),
            'currency, without language id'                                      => array(
                '123',
                '[$€]0.00',
                '€123.00'
            ),
            'quoted/escaped semicolons'                                          => array(
                '-123',
                '0";"\;"0;0""0\;0"\";0.0000',
                '123.0000'
            ),
            'quoted sections, preceeding a positive value'                       => array(
                '123',
                '"[POSITIVE] "0.00;"[NEGATIVE] "0.00',
                '[POSITIVE] 123.00'
            ),
            'quoted sections, preceeding a negative value'                       => array(
                '-123',
                '"[POSITIVE] "0.00;"[NEGATIVE] "0.00',
                '[NEGATIVE] 123.00'
            ),
            'escaping does not work within square brackets'                      => array(
                '100',
                '[$CU\\R\\\\RE"N[CY-]0.00',
                'CU\\R\\\\RE"N[CY100.00'
            ),
            'handling of spaces in format string'                                => array(
                'demo',
                '" t" "e ""x t"" "":" @',
                ' t e x t : demo'
            ),
            'date format, with language code'                                    => array(
                '123',
                '[$-123]DD-MM-YYYY',
                '02-05-1900'
            ),
            'date format, prefixed with extras'                                  => array(
                '123',
                '"before "DD-MM-YYYY',
                'before 02-05-1900'
            ),
            'time in 12h'                                                        => array(
                '0.75',
                'hh:mm AM/PM',
                '06:00 PM'
            ),
            'time in 24h, but with a confusing string'                           => array(
                '0.75',
                'hh:mm" AM/PM"',
                '18:00 AM/PM'
            ),
            '1 conditional, no default, invalid numeric'                         => array(
                '0',
                '[<0]0',
                '########'
            ),
            '1 conditional, no default, invalid text'                            => array(
                'test',
                '[<0]0',
                'test'
            ),
            '1 conditional, with NUMBER-ONLY default, invalid numeric'           => array(
                '0',
                '[<0]0;"["0"]"',
                '[0]'
            ),
            '1 conditional, with NUMBER-ONLY default, invalid text'              => array(
                'test',
                '[<0]0;"["0"]"',
                'test'
            ),
            '1 conditional, with TEXT-ONLY default, invalid numeric'             => array(
                '0',
                '[<0]0;"["@"]"',
                '########'
            ),
            '1 conditional, with TEXT-ONLY default, invalid text'                => array(
                'test',
                '[<0]0;"["@"]"',
                '[test]'
            ),
            '1 conditional, with 2 defaults, with a number'                      => array(
                '0.123',
                '[>50]0;.0;"["@"]"',
                '.1'
            ),
            '1 conditional, with 2 defaults, with text'                          => array(
                'test',
                '[>50]0;.0;"["@"]"',
                '[test]'
            ),
            '2 conditionals, no default, invalid numeric'                        => array(
                '0',
                '[<0]0;[>0]0',
                '########'
            ),
            '2 conditionals, no default, invalid text'                           => array(
                'test',
                '[<0]0;[>0]0',
                'test'
            ),
            '2 conditionals, with 1 default, invalid numeric'                    => array(
                '0',
                '[<0]0;[>0]0;0.000',
                '0.000'
            ),
            '2 conditionals, with 1 default, invalid text'                       => array(
                'test',
                '[<0]0;[>0]0;0.000',
                'test'
            ),
            '2 conditionals, with 2 defaults, invalid numeric'                   => array(
                '0',
                '[<0]0;[>0]0;0.000;"["@"]"',
                '0.000'
            ),
            '2 conditionals, with 2 defaults, invalid text'                      => array(
                'test',
                '[<0]0;[>0]0;0.000;"["@"]"',
                '[test]'
            ),
            'conditional in 2nd section, matching value'                         => array(
                '50',
                '"1st "0;[=50]"2nd "0;"3rd "0',
                '1st 50'
            ),
            'conditional in 1st section, matching value'                         => array(
                '50',
                '[=50]"1st "0;"2nd "0;"3rd "0',
                '1st 50'
            ),
            'conditional in 1st section, positive value'                         => array(
                '51',
                '[=50]"1st "0;"2nd "0;"3rd "0',
                '3rd 51'
            ),
            'conditional in 1st section, negative value'                         => array(
                '-50',
                '[=50]"1st "0;"2nd "0;"3rd "0',
                '2nd 50'
            ),
            'conditional in 2nd section, negative value'                         => array(
                '-50',
                '"1st "0;[=50]"2nd "0;"3rd "0',
                '-3rd 50'
            ),
            'conditional + color + whitespace'                                   => array(
                '-50',
                ' [red]   [=-50]  "1st "0;"2nd "0',
                // Exact behavior differs between applications. Some keep the spaces, some remove them.
                '      1st 50'
            ),
            'percentage signs change output values'                              => array(
                '0.12',
                '0.00%',
                '12.00%'
            ),
            'percentage signs in quotes don\'t change output values'             => array(
                '0.12',
                '0.00"%"',
                '0.12%'
            ),
            'fraction without whole values'                                      => array(
                '1.2',
                '?/#',
                '6/5'
            ),
            'fraction with a negative whole value'                               => array(
                '-5',
                '0/0',
                '-5/1'
            ),
            'fraction, negative, with a post-decimal value starting with 0'      => array(
                '-2.025',
                '0/0',
                '-81/40'
            ),
            'fraction with unneccessary whole values'                            => array(
                '0.2',
                '? #/0',
                '1/5'
            ),
            'fraction with whole values and whitespace'                          => array(
                '1.2',
                ' # 0/? ',
                ' 1 1/5 '
            ),
            'fraction optional whole-values'                                     => array(
                '0.25',
                '# 0/0',
                '1/4'
            ),
            'fraction non-optional whole-values'                                 => array(
                '0.25',
                '0 0/0',
                '0 1/4'
            ),
            'fraction whole-value extraction with mandatory fraction part'       => array(
                '5',
                '0 0/0',
                '5 0/1'
            ),
            'fractions can be percent values'                                    => array(
                '0.005',
                '0/0%',
                '1/2%'
            ),
            'fraction with additional digits'                                    => array(
                '0.2',
                '#00/#00',
                '01/05'
            ),
            'fraction with whole values and additional digits'                   => array(
                '5.2',
                '#?00 #?00/#?00',
                ' 05  01/ 05'
            ),
            'fraction with whole values and additional digits in odd order'      => array(
                '5.2',
                '?#00 ?#00/?#00',
                ' 05  01/ 05'
            ),
            'fraction broken up by quoted sections'                              => array(
                '2.2',
                // Note: No actual "space" between whole-value and fraction.
                '#00" and "0/0" and "0',
                // Quoted section counts as a "space" for whole-value separation.
                '02 and 1/5 and 0'
            ),
            'fraction broken up by empty section'                                => array(
                '2.2',
                '#00""0/0',
                '021/5'
            ),
            'fraction left part broken up by some unexpected symbol'             => array(
                '2.2',
                '#00+0/00',
                '02+1/05'
            ),
            'fraction right part broken up by some unexpected symbol'            => array(
                '2.2',
                '#00/0+0',
                '11/5+0'
            ),
            'fraction with exclusion of whole-value parts'                       => array(
                '0.2',
                '"A" # "B" # "C" 0/0',
                'A 1/5' // Excel: 'A  B 1/5'; Calc: 'A  C 1/5'
            ),
            'fraction WITHOUT exclusion of whole-value parts (compare to above)' => array(
                '5.2',
                '"A" # "B" # "C" 0/0',
                // Note that any characters separated from the numerator are considered part of the whole-value.
                'A  B 5 C 1/5'
            ),
            'fraction optional nominator'                                        => array(
                '255',
                '0 #/0',
                // In Excel: '255'; In Calc: '255 0/1'; Given other testcases, Calc seems more reasonable here.
                '255 0/1'
            ),
            'fraction optional denominator'                                      => array(
                '255',
                '0 0/#',
                '255 0/1'
            ),
            'fraction optional nominator and denominator'                        => array(
                '123',
                '0 "A" 0 "B" #/#"C"',
                // Note that everything between end of whole-value and end of denominator is dropped, including spaces.
                '12 A 3C'
            ),

            /* Marking least significant digits as "optional" does not make much sense, and it leads to a lot of very
             * confusing behavior. On top of that, it's difficult to find a "sane" solution for such formats.
             * As such, the reader currently does not deliver any results that resemble those of popular editors here.
             * This case is left commented here, in case anyone's ever wondering about this.
             *
             * 'fraction bizarre format rule' => array(
             *     '0.25',
             *     '00?# 00?#/00?#', // Uncommon: Least significant digit marked as "optional".
             *     '00 0 00 1/00 4' // Excel: '00  00  1/004'; Calc: '00  00 1/004 '
             * ), */

            /* The following is technically an ECMA requirement, but popular editors can't actually parse it.
             * As such, this is left commented for documentation's sake. Any output can be deemed "acceptable" here.
             *
             * 'not a fraction' => array(
             *     '255',
             *     '0.00 / "something"',
             *     '255.00 / something' // No output in popular editors. They just error.
             * ), */

            'scientific notation - small number'                                          => array(
                '0.000123',
                '0.00E+00',
                '1.23E-04'
            ),
            'scientific notation - large number'                                          => array(
                '123456789',
                '0.##E+##', // 1 more # than needed for the exponent
                '1.23E+8'
            ),
            'scientific notation - as scientific notation'                                => array(
                '1.23E+03',
                '##.00e+0', // 1 more 0 than non-0 values after the decimal
                '12.30e+2'
            ),
            'scientific notation - excessive # symbols'                                   => array(
                '0.055',
                '####.##e+0',
                // This results in '550.E-4' in Excel and '550E-4' in Calc. None of the two seem logical.
                // (4 pre-decimal digits available + 2 nonzero digits in value => 3 pre-decimal digits used?)
                '5500e-5'
            ),
            'scientific notation - optional-to-spaces in exponent'                        => array(
                '5',
                '0.0E+??',
                '5.0E+ 0'
            ),
            'scientific notation - nothing to the left of the seperator'                  => array(
                '5',
                '.00##e+#00',
                '.50e+01'
            ),
            'scientific notation - nothing'                                               => array(
                '123',
                '.E+0',
                '123' // ".E+1" in Excel, "123" (no formatting) in Calc. The latter seems more sensible.
            ),
            'scientific notation - only required position after seperator'                => array(
                '5',
                '##.0##e+00', // Using at least 1 digit to the left of the separator is preferred over using only '0' digits.
                '50.0e-01' // "5.0E+00" in Excel and Calc, which doesn't seem logical for this format.
            ),
            'scientific notation - shift separator to the left'                           => array(
                '123.45E-03',
                '##.##E+00',
                '12.35E-02' // This also tests for rounding issues.
            ),
            'scientific notation - shift separator to the right, with too many positions' => array(
                '1.2345E-05',
                '000.0##E+00', // 1 more # than non-0 values after the decimal.
                '123.45E-07' // "012.345E-06" in Excel and Calc, which doesn't seem logical for this format.
            ),
            'scientific notation - special case for minus exponential symbol'             => array(
                '12345',
                '00.00e-#',
                '12.35e3' // "01.23E4" in Excel and Calc, which doesn't seem logical for this format.
            ),
            'scientific notation - including text elements'                               => array(
                '1.2345E-05',
                '#.##"hello"E+##',
                '1.23helloE-5'
            ),
            'scientific notation - including invalid decimal elements'                    => array(
                '123',
                '0.00E+0.,0', // Excel autocorrects to "0.00E+0.0", Calc doesn't accept it.
                '1.23E+02' // Excel: "1.23E+2.0", Calc: "1.23E+02"
            ),
            'scientific notation - but not actually'                                      => array(
                '123',
                '0.00E+', // Exponent digits are missing.
                '123.00' // Note: Popular editors refuse further operation here, so technically, anything goes here.
            ),
            'scientific notation - ineffective when escaped'                              => array(
                '123',
                '[$E+]0.00"E+"\\E+E\\+0',
                'E+123.00E+E+E+0'
            ),
            'scientific notation - exponent larger than 1 digit'                          => array(
                '0.00000000005',
                '0.00E+0', // 1 digit exponent in format
                '5.00E-11' // -> 2 digit exponent in output anyway
            ),
            'scientific notation - split by text before decimal'                          => array(
                '123',
                '0" and" 0.0E+00',
                '1 and 2.3E+01'
            ),
            'scientific notation - split by text after decimal'                           => array(
                '123',
                '0.0 "and "0 E+00',
                '1.2 and 3 E+02'
            ),
            'scientific notation - exponent split by text'                                => array(
                '123.456',
                '0.00E+0 "text" 0',
                '1.23E+0 text 2'
            )
        );
    }
}
