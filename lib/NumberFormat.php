<?php

namespace Aspera\Spreadsheet\XLSX;

use Exception;
use InvalidArgumentException;
use DateTime;
use DateTimeZone;
use DateInterval;

/**
 * Functionality to understand and apply number formatting.
 *
 * Format parsing/application works in the following steps (value formats may differ in actual code):
 *  - Parse the styles.xml document, extracting all necessary formatting data from it.
 *      -> Done in Reader.php. The values are injected via injectXfNumFmtIds() and injectNumberFormats().
 *  - Parse number formats on demand, storing parsing results in cache. Parsing steps:
 *      -- Split format into its sections for different value types.
 *          -> '0,00;[red]0,00' => ['0,00', '[red]0,00']
 *      -- Split each section into logical tokens.
 *          -> '[red]0,00' => [['red', square_bracket_index:0], '0,00']
 *      -- Determine a purpose for each section, indicating the range of values to apply it to, while adding default formats.
 *          -> ['0,00', '[red]0,00'] => ['>0':'0,00', '<0':'[red]0,00', 'default_number':'0,00', 'default_text':'@']
 *      -- Run additional detection logic to add additional, semantic information to the stored format data.
 *          -> ['<0':'[red]0,00'] => ['is_percentage':false, 'prepend_minus_sign':false, ...]
 *  - Apply the format by working through each logical token while regarding the aforementioned, additional information.
 */
class NumberFormat
{
    /** @var array Conversion matrix to convert XLSX date formats to PHP date formats. */
    const DATE_REPLACEMENTS = array(
        'All' => array(
            '\\'    => '',
            'am/pm' => 'A',
            'yyyy'  => 'Y',
            'yy'    => 'y',
            'mmmmm' => 'M',
            'mmmm'  => 'F',
            'mmm'   => 'M',
            ':mm'   => ':i',
            'mm'    => 'm',
            'm'     => 'n',
            'dddd'  => 'l',
            'ddd'   => 'D',
            'dd'    => 'd',
            'd'     => 'j',
            'ss'    => 's',
            '.s'    => ''
        ),
        '24H' => array(
            'hh' => 'H',
            'h'  => 'G'
        ),
        '12H' => array(
            'hh' => 'h',
            'h'  => 'G'
        )
    );

    /** @var array Base formats for XLSX documents, to be made available without former declaration. */
    const BUILTIN_FORMATS = array(
        0 => '',
        1 => '0',
        2 => '0.00',
        3 => '#,##0',
        4 => '#,##0.00',

        9  => '0%',
        10 => '0.00%',
        11 => '0.00E+00',
        12 => '# ?/?',
        13 => '# ??/??',
        14 => 'mm-dd-yy',
        15 => 'd-mmm-yy',
        16 => 'd-mmm',
        17 => 'mmm-yy',
        18 => 'h:mm AM/PM',
        19 => 'h:mm:ss AM/PM',
        20 => 'h:mm',
        21 => 'h:mm:ss',
        22 => 'm/d/yy h:mm',

        37 => '#,##0 ;(#,##0)',
        38 => '#,##0 ;[Red](#,##0)',
        39 => '#,##0.00;(#,##0.00)',
        40 => '#,##0.00;[Red](#,##0.00)',

        45 => 'mm:ss',
        46 => '[h]:mm:ss',
        47 => 'mmss.0',
        48 => '##0.0E+0',
        49 => '@',

        // CHT & CHS
        27 => '[$-404]e/m/d',
        30 => 'm/d/yy',
        36 => '[$-404]e/m/d',
        50 => '[$-404]e/m/d',
        57 => '[$-404]e/m/d',

        // THA
        59 => 't0',
        60 => 't0.00',
        61 => 't#,##0',
        62 => 't#,##0.00',
        67 => 't0%',
        68 => 't0.00%',
        69 => 't# ?/?',
        70 => 't# ??/??'
    );

    /** @var bool Is the gmp_gcd method available for usage? Cached value. */
    private static $gmp_gcd_available = false;

    /** @var DateTime Standardized base date for the document's date/time values. */
    private static $base_date;

    /** @var string Decimal separator character for output of locally formatted values. */
    private $decimal_separator;

    /** @var string Thousands separator character for output of locally formatted values. */
    private $thousand_separator;

    /** @var string Currency character for output of locally formatted values. */
    private $currency_code;

    /** @var bool Do not format date/time values and return DateTime objects instead. Default false. */
    private $return_date_time_objects;

    /** @var string Format to use when outputting dates, regardless of originally set formatting.
     *              (Note: Will also be used if the original formatting omits time information, but the data value contains time information.) */
    private $enforced_date_format;

    /** @var string Format to use when outputting time information, regardless of originally set formatting. */
    private $enforced_time_format;

    /** @var string Format to use when outputting datetime values, regardless of originally set formatting. */
    private $enforced_datetime_format;

    /** @var array List of number formats defined by the current XLSX file; array key = format index */
    private $formats = array();

    /** @var array List of custom formats defined by the user; array key = format index */
    private $customized_formats = array();

    /**
     * "numFmtId" attribute values of all cellXfs > xf elements.
     * Key: xf-index, referred to in the "s" attribute values of r > c elements.
     * Value: index of the number format to apply. 0: "general" format. null: No format.
     *
     * @var array
     */
    private $styles = array();

    /** @var array Cache for already processed format strings. Type of each element: NumberFormatSection[] */
    private $parsed_format_cache = array();

    /**
     * @param  array $options
     *
     * @throws Exception
     */
    public function __construct($options = array())
    {
        if (!empty($options['CustomFormats'])) {
            $this->initCustomFormats($options['CustomFormats']);
        }
        if (!empty($options['ForceDateFormat'])) {
            $this->enforced_date_format = $options['ForceDateFormat'];
        }
        if (!empty($options['ForceTimeFormat'])) {
            $this->enforced_time_format = $options['ForceTimeFormat'];
        }
        if (!empty($options['ForceDateTimeFormat'])) {
            $this->enforced_datetime_format = $options['ForceDateTimeFormat'];
        }
        $this->return_date_time_objects = !empty($options['ReturnDateTimeObjects']);

        $this->initLocale();

        self::initBaseDate();
        self::$gmp_gcd_available = function_exists('gmp_gcd');
    }

    /**
     * @param array $xf_num_fmt_ids
     */
    public function injectXfNumFmtIds($xf_num_fmt_ids)
    {
        $this->styles = $xf_num_fmt_ids;
    }

    /**
     * @param array $number_formats
     */
    public function injectNumberFormats($number_formats)
    {
        $this->formats = $number_formats;
    }

    /**
     * @param string $decimal_separator
     */
    public function setDecimalSeparator($decimal_separator)
    {
        if (!is_string($decimal_separator)) {
            throw new InvalidArgumentException('Given argument is not a string.');
        }
        $this->decimal_separator = $decimal_separator;
    }

    /**
     * @param string $thousand_separator
     */
    public function setThousandsSeparator($thousand_separator)
    {
        if (!is_string($thousand_separator)) {
            throw new InvalidArgumentException('Given argument is not a string.');
        }
        $this->thousand_separator = $thousand_separator;
    }

    /**
     * Set the currency character code to use for the output of locale-oriented formatted values
     *
     * @param string $new_character
     */
    public function setCurrencyCode($new_character)
    {
        if (!is_string($new_character)) {
            throw new InvalidArgumentException('Given argument is not a string.');
        }
        $this->currency_code = $new_character;
    }

    /**
     * @param  string $value
     * @param  int    $style_id In worksheet cells, this is also referred to as the "style" of a cell.
     * @return mixed|string
     *
     * @throws Exception
     */
    public function tryFormatValue($value, $style_id)
    {
        if ($value !== '' && $style_id && isset($this->styles[$style_id])) {
            $value = $this->formatValue($value, $style_id);
        } elseif ($value) {
            $value = $this->generalFormat($value);
        }

        return $value;
    }

    /**
     * Formats the value according to the index.
     *
     * @param   string $value
     * @param   int    $format_index
     * @return  string
     *
     * @throws  Exception
     */
    public function formatValue($value, $format_index)
    {
        if (!is_numeric($value)) {
            // Only numeric values are formatted.
            return $value;
        }

        if (isset($this->styles[$format_index]) && ($this->styles[$format_index] !== false)) {
            $format_index = $this->styles[$format_index];
        } else {
            // Invalid format_index or the style was explicitly declared as "don't format anything".
            return $value;
        }

        if ($format_index === 0) {
            // Special case for the "General" format
            return $this->generalFormat($value);
        }

        $format = array();
        if (isset($this->parsed_format_cache[$format_index])) {
            $format = $this->parsed_format_cache[$format_index];
        }

        if (!$format) {
            $format = array(
                'Code'      => false,
                'Type'      => false,
                'Scale'     => 1,
                'Thousands' => false,
                'Currency'  => false
            );

            if (array_key_exists($format_index, $this->customized_formats)) {
                $format['Code'] = $this->customized_formats[$format_index];
            } elseif (array_key_exists($format_index, self::BUILTIN_FORMATS)) {
                $format['Code'] = self::BUILTIN_FORMATS[$format_index];
            } elseif (isset($this->formats[$format_index])) {
                $format['Code'] = $this->formats[$format_index];
            }

            // Format code found, now parsing the format
            if ($format['Code']) {
                $sections = explode(';', $format['Code']);
                $format['Code'] = $sections[0];

                switch (count($sections)) {
                    case 2:
                        if ($value < 0) {
                            $format['Code'] = $sections[1];
                        }
                        break;
                    case 3:
                    case 4:
                        if ($value < 0) {
                            $format['Code'] = $sections[1];
                        } elseif ($value === 0) {
                            $format['Code'] = $sections[2];
                        }
                        break;
                    default:
                        // nop
                        break;
                }
            }

            // Stripping colors
            $format['Code'] = trim(preg_replace('{^\[[[:alpha:]]+\]}i', '', $format['Code']));

            // Percentages
            if (substr($format['Code'], -1) === '%') {
                $format['Type'] = 'Percentage';
            } elseif (preg_match('{^(\[\$[[:alpha:]]*-[0-9A-F]*\])*[hmsdy]}i', $format['Code'])) {
                $format['Type'] = 'DateTime';

                $format['Code'] = trim(preg_replace('{^(\[\$[[:alpha:]]*-[0-9A-F]*\])}i', '', $format['Code']));
                $format['Code'] = strtolower($format['Code']);

                $format['Code'] = strtr($format['Code'], self::DATE_REPLACEMENTS['All']);
                if (strpos($format['Code'], 'A') === false) {
                    $format['Code'] = strtr($format['Code'], self::DATE_REPLACEMENTS['24H']);
                } else {
                    $format['Code'] = strtr($format['Code'], self::DATE_REPLACEMENTS['12H']);
                }
            } elseif ($format['Code'] === '[$eUR ]#,##0.00_-') {
                $format['Type'] = 'Euro';
            } else {
                // Removing skipped characters
                $format['Code'] = preg_replace('{_.}', '', $format['Code']);
                // Removing unnecessary escaping
                $format['Code'] = preg_replace("{\\\\}", '', $format['Code']);
                // Removing string quotes
                $format['Code'] = str_replace(array('"', '*'), '', $format['Code']);
                // Removing thousands separator
                if (strpos($format['Code'], '0,0') !== false || strpos($format['Code'], '#,#') !== false) {
                    $format['Thousands'] = true;
                }
                $format['Code'] = str_replace(array('0,0', '#,#'), array('00', '##'), $format['Code']);

                // Scaling (Commas indicate the power)
                $scale = 1;
                $matches = array();
                if (preg_match('{(0|#)(,+)}', $format['Code'], $matches)) {
                    $scale = 1000 ** strlen($matches[2]);
                    // Removing the commas
                    $format['Code'] = preg_replace(array('{0,+}', '{#,+}'), array('0', '#'), $format['Code']);
                }

                $format['Scale'] = $scale;

                if (preg_match('{#?.*\?\/\?}', $format['Code'])) {
                    $format['Type'] = 'Fraction';
                } else {
                    $format['Code'] = str_replace('#', '', $format['Code']);

                    $matches = array();
                    if (preg_match('{(0+)(\.?)(0*)}', preg_replace('{\[[^\]]+\]}', '', $format['Code']), $matches)) {
                        $integer = $matches[1];
                        $decimal_point = $matches[2];
                        $decimals = $matches[3];

                        $format['MinWidth'] = strlen($integer) + strlen($decimal_point) + strlen($decimals);
                        $format['Decimals'] = $decimals;
                        $format['Precision'] = strlen($format['Decimals']);
                        $format['Pattern'] = '%0' . $format['MinWidth'] . '.' . $format['Precision'] . 'f';
                    }
                }

                $matches = array();
                if (preg_match('{\[\$(.*)\]}u', $format['Code'], $matches)) {
                    // Format contains a currency code (Syntax: [$<Currency String>-<language info>])
                    $curr_code = explode('-', $matches[1]);
                    if (isset($curr_code[0])) {
                        $curr_code = $curr_code[0];
                    } else {
                        $curr_code = $this->currency_code;
                    }
                    $format['Currency'] = $curr_code;
                }
                $format['Code'] = trim($format['Code']);
            }
            $this->parsed_format_cache[$format_index] = $format;
        }

        // Applying format to value
        if ($format) {
            if ($format['Code'] === '@') {
                return (string) $value;
            }

            if ($format['Type'] === 'Percentage') {
                // Percentages
                if ($format['Code'] === '0%') {
                    $value = round(100 * $value, 0) . '%';
                } else {
                    $value = sprintf('%.2f%%', round(100 * $value, 2));
                }
            } elseif ($format['Type'] === 'DateTime') {
                // Dates and times
                $days = (int) $value;
                // Correcting for Feb 29, 1900
                if ($days > 60) {
                    $days--;
                }

                // At this point time is a fraction of a day
                $time = ($value - (int) $value);
                $seconds = 0;
                if ($time) {
                    // Here time is converted to seconds
                    // Workaround against precision loss: set low precision will round up milliseconds
                    $seconds = (int) round($time * 86400, 0);
                }

                $original_value = $value;
                $value = clone self::$base_date;
                if ($original_value < 0) {
                    // Negative value, subtract interval
                    $days = abs($days) + 1;
                    $seconds = abs($seconds);
                    $value->sub(new DateInterval('P' . $days . 'D' . ($seconds ? 'T' . $seconds . 'S' : '')));
                } else {
                    // Positive value, add interval
                    $value->add(new DateInterval('P' . $days . 'D' . ($seconds ? 'T' . $seconds . 'S' : '')));
                }

                if (!$this->return_date_time_objects) {
                    // Determine if the format is a date/time/datetime format and apply enforced formatting accordingly
                    $contains_date = preg_match('#[DdFjlmMnoStwWmYyz]#u', $format['Code']);
                    $contains_time = preg_match('#[aABgGhHisuv]#u', $format['Code']);
                    if ($contains_date) {
                        if ($contains_time) {
                            if ($this->enforced_datetime_format) {
                                $value = $value->format($this->enforced_datetime_format);
                            }
                        } else if ($this->enforced_date_format) {
                            $value = $value->format($this->enforced_date_format);
                        }
                    } else if ($this->enforced_time_format) {
                        $value = $value->format($this->enforced_time_format);
                    }

                    if ($value instanceof DateTime) {
                        // No format enforcement for this value type found. Format as declared.
                        $value = $value->format($format['Code']);
                    }
                } // else: A DateTime object is returned
            } elseif ($format['Type'] === 'Euro') {
                $value = 'EUR ' . sprintf('%1.2f', $value);
            } else {
                // Fractional numbers; We get "0.25" and have to turn that into "1/4".
                if ($format['Type'] === 'Fraction' && ($value != (int) $value)) {
                    // Split fraction from integer value (2.25 => 2 and 0.25)
                    $integer = floor(abs($value));
                    $decimal = fmod(abs($value), 1);

                    // Turn fraction into non-decimal value (0.25 => 25)
                    $decimal *= 10 ** (strlen($decimal) - 2);

                    // Obtain biggest divisor for the fraction part (25 => 100 => 25/100)
                    $decimal_divisor = 10 ** strlen($decimal);

                    // Determine greatest common divisor for fraction optimization (so that 25/100 => 1/4)
                    if (self::$gmp_gcd_available) {
                        $gcd = gmp_strval(gmp_gcd($decimal, $decimal_divisor));
                    } else {
                        $gcd = self::GCD($decimal, $decimal_divisor);
                    }

                    // Determine fraction parts (1 and 4 => 1/4)
                    $adj_decimal = $decimal / $gcd;
                    $adj_decimal_divisor = $decimal_divisor / $gcd;

                    if (   strpos($format['Code'], '0') !== false
                        || strpos($format['Code'], '#') !== false
                        || strpos($format['Code'], '? ?') === 0
                    ) {
                        // Extract whole values from fraction (2.25 => "2 1/4")
                        $value = ($value < 0 ? '-' : '') .
                            ($integer ? $integer . ' ' : '') .
                            $adj_decimal . '/' .
                            $adj_decimal_divisor;
                    } else {
                        // Show entire value as fraction (2.25 => "9/4")
                        $adj_decimal += $integer * $adj_decimal_divisor;
                        $value = ($value < 0 ? '-' : '') .
                            $adj_decimal . '/' .
                            $adj_decimal_divisor;
                    }
                } else {
                    // Scaling
                    $value /= $format['Scale'];
                    if (!empty($format['MinWidth']) && $format['Decimals']) {
                        if ($format['Thousands']) {
                            $value = number_format($value, $format['Precision'],
                                $this->decimal_separator, $this->thousand_separator);
                        } else {
                            $value = sprintf($format['Pattern'], $value);
                        }
                        $format_code = preg_replace('{\[\$.*\]}', '', $format['Code']);
                        $value = preg_replace('{(0+)(\.?)(0*)}', $value, $format_code);
                    }
                }

                // Currency/Accounting
                if ($format['Currency']) {
                    $value = preg_replace('{\[\$.*\]}u', $format['Currency'], $value);
                }
            }
        }

        return $value;
    }

    /**
     * @param   array $custom_formats
     * @return  void
     */
    public function initCustomFormats(array $custom_formats)
    {
        foreach ($custom_formats as $format_index => $format) {
            if (array_key_exists($format_index, self::BUILTIN_FORMATS) !== false) {
                $this->customized_formats[$format_index] = $format;
            }
        }
    }

    /**
     * Attempts to approximate Excel's "general" format.
     *
     * @param   mixed   $value
     * @return  mixed
     */
    private function generalFormat($value)
    {
        if (is_numeric($value)) {
            $value = (float) $value;
        }

        return $value;
    }

    /**
     * Helper function for greatest common divisor calculation in case GMP extension is not enabled.
     *
     * @param   int $int_1
     * @param   int $int_2
     * @return  int Greatest common divisor
     */
    private static function GCD($int_1, $int_2)
    {
        $int_1 = (int) abs($int_1);
        $int_2 = (int) abs($int_2);

        if ($int_1 + $int_2 === 0) {
            return 0;
        }

        $divisor = 1;
        while ($int_1 > 0) {
            $divisor = $int_1;
            $int_1 = $int_2 % $int_1;
            $int_2 = $divisor;
        }

        return $divisor;
    }

    /**
     * Set base date for calculation of retrieved date/time data.
     *
     * @throws Exception
     */
    private static function initBaseDate() {
        self::$base_date = new DateTime();
        self::$base_date->setTimezone(new DateTimeZone('UTC'));
        self::$base_date->setDate(1900, 1, 0);
        self::$base_date->setTime(0, 0, 0);
    }

    /**
     * Pre-fill locale related data using current system locale.
     */
    private function initLocale() {
        $locale = localeconv();
        $this->decimal_separator = $locale['decimal_point'];
        $this->thousand_separator = $locale['thousands_sep'];
        $this->currency_code = $locale['int_curr_symbol'];
    }
}
