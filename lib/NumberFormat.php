<?php

namespace Aspera\Spreadsheet\XLSX;

use Exception;
use RuntimeException;
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

    /** @var bool Do not format anything. Returns numbers as-is. (Note: Does not affect Date/Time values.) */
    private $return_unformatted;

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
    private $number_formats = array();

    /** @var array List of custom formats defined by the user; array key = format index */
    private $customized_formats = array();

    /**
     * "numFmtId" attribute values of all cellXfs > xf elements.
     * Key: xf-index, referred to in the "s" attribute values of r > c elements.
     * Value: index of the number format to apply. 0: "general" format. null: No format.
     *
     * @var array
     */
    private $xf_num_fmt_ids = array();

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
        $this->return_unformatted = !empty($options['ReturnUnformatted']);
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
        $this->xf_num_fmt_ids = $xf_num_fmt_ids;
    }

    /**
     * @param array $number_formats
     */
    public function injectNumberFormats($number_formats)
    {
        $this->number_formats = $number_formats;
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
     * @param  string $value
     * @param  int    $xf_id In worksheet cells, this is also referred to as the "style" of a cell.
     * @return mixed|string
     *
     * @throws Exception
     */
    public function tryFormatValue($value, $xf_id)
    {
        if ($value !== '' && $xf_id && isset($this->xf_num_fmt_ids[$xf_id])) {
            return $this->formatValue($value, $xf_id);
        }

        /* Do not format value, no style set or quotePrefix is set. ($xf_id = null, starts at 1) */
        return $value;
    }

    /**
     * Converts XLSX-style datetime data (a plain decimal number) to a DateTime object.
     *
     * @param  mixed $value
     * @return DateTime
     *
     * @throws Exception
     */
    private function convertNumberToDateTime($value)
    {
        // Determine days. (value = amount of days since base date)
        $days = (int) $value;
        if ($days > 60) {
            $days--; // Correcting for Feb 29, 1900
        }

        // Determine time. (decimal value = fraction of a day)
        $time = $value - (int) $value;
        $seconds = 0;
        if ($time) {
            // Workaround against precision loss: set low precision will round up milliseconds
            $seconds = (int) round($time * 86400, 0);
        }

        $datetime = clone self::$base_date;
        if ($value < 0) {
            // Negative value, subtract interval
            $days = abs($days) + 1;
            $seconds = abs($seconds);
            $datetime->sub(new DateInterval('P' . $days . 'D' . ($seconds ? 'T' . $seconds . 'S' : '')));
        } else {
            // Positive value, add interval
            $datetime->add(new DateInterval('P' . $days . 'D' . ($seconds ? 'T' . $seconds . 'S' : '')));
        }

        return $datetime;
    }

    /**
     * Formats the value according to the index.
     *
     * @param   string $value
     * @param   int    $xf_id
     * @return  string
     *
     * @throws  Exception
     */
    public function formatValue($value, $xf_id)
    {
        $num_fmt_id = 0;
        if (isset($this->xf_num_fmt_ids[$xf_id]) && $this->xf_num_fmt_ids[$xf_id] !== null) {
            $num_fmt_id = $this->xf_num_fmt_ids[$xf_id];
        }
        if ($num_fmt_id === 0) {
            // Invalid format_index or "general" format should be applied ($num_fmt_id = 0), which just outputs the value as-is.
            return $value;
        }

        // Get definition of format for the given format_index.
        $section = $this->getFormatSectionForValue($value, $num_fmt_id);

        // If percentage values are expected, multiply value accordingly before formatting.
        if ($section->isPercentage()) {
            $value *= 100;
        }

        // If this is a datetime format, prepare datetime values. (And return them immediately, if such is requested.)
        $datetime = null;
        if ($section->getDateTimeType()) {
            $datetime = $this->convertNumberToDateTime($value);

            // Return DateTime objects as-is?
            if ($this->return_date_time_objects) {
                return $datetime;
            }

            // Handle enforced date/time/datetime format.
            switch ($section->getDateTimeType()) {
                case 'date':
                    if ($this->enforced_date_format) {
                        return $datetime->format($this->enforced_date_format);
                    }
                    break;
                case 'time':
                    if ($this->enforced_time_format) {
                        return $datetime->format($this->enforced_time_format);
                    }
                    break;
                case 'datetime':
                    if ($this->enforced_datetime_format) {
                        return $datetime->format($this->enforced_datetime_format);
                    }
                    break;
                default:
                    // Note: Should never happen. Exception is just to be safe.
                    throw new RuntimeException('Specific datetime_type for format_index [' . $num_fmt_id . '] is unknown.');
            }
        }

        // If formatting is not desired, return value as-is.
        if ($this->return_unformatted) {
            return $value;
        }

        // Apply format to value, going token by token.
        $output = '';
        foreach ($section->getTokens() as $token) {
            if ($token->isQuoted()) {
                $output .= $token->getCode();
                continue;
            }

            if ($token->isInSquareBrackets()) {
                // Should not happen, since pre-parsing should have handled these already.
                continue;
            }

            if ($datetime) {
                // For DateTime values, DateTime->format() is all that's left to do at this point.
                $output .= $datetime->format($token->getCode());
                continue;
            }

            // Approach of applying format: Use format code as a base, then replace its functional parts with values.
            $output_part = $token->getCode();

            switch ($token->getNumberType()) {
                case 'decimal':
                    $decimal = $this->applyDecimalFormat($value, $token);
                    $output_part = preg_replace('{[0#?,./]+}', $decimal, $output_part);
                    break;
                case 'fraction':
                    $fraction = $this->applyFractionFormat(abs($value), $token); // abs() to ease minus-sign handling.

                    // preg_replace('{[0#?,./ ]+}', ...) would remove leading/trailing spaces. Alternative:
                    if (preg_match('{[0#?,./ ]+}', $output_part, $matches)) {
                        $match_string = trim($matches[0]);
                        $output_part = str_replace($match_string, $fraction, $output_part);
                    }
                    break;
                default:
                    // This token has no detailed number formatting information that would need to be applied.
                    break;
            }

            // Handle @ symbol.
            $output .= str_replace('@', $value, $output_part);
        }

        if ($section->prependMinusSign() && $value < 0) {
            $output = '-' . $output;
        }

        return $output;
    }

    /**
     * Converts the given value to a fraction string.
     * Note: Purposefully omits minus sign for negative values.
     *
     * @param  mixed             $value
     * @param  NumberFormatToken $token
     * @return string
     */
    private static function applyFractionFormat($value, $token)
    {
        if ($value == (int) $value) {
            return $value; // Value is a whole number. Nothing to do here.
        }

        /* --- Conversion from decimal to fraction using floating-point-safe approach:
         * Step 1: Multiply value by a power of 10 that turns the whole decimal into a natural number.
         *  2.025 * 1000 => 2025
         * Step 2 (only informal): Assemble fraction. numerator: The new value. denominator: The chosen power of 10.
         *  2025/1000
         * Step 3: Simplify by dividing by the greatest common divisor between numerator and denominator.
         *  GCD(2025,1000) => 25 | 2025/25 / 1000/25 => 81/40
         * Step 4: If requested, extract whole values.
         *  81/40 => 2 1/40 */

        $str_value = (string) $value;
        $denominator = 10 ** (strlen($str_value) - strpos($str_value, '.') - 1);
        $numerator = $value * $denominator;
        $gcd = self::$gmp_gcd_available
            ? gmp_strval(gmp_gcd($numerator, $denominator))
            : self::GCD($numerator, $denominator);
        $numerator /= $gcd;
        $denominator /= $gcd;
        if ($token->doExtractWhole() && $value > 1) {
            return floor($value) . ' ' . ($numerator % $denominator) . '/' . $denominator;
        }
        return $numerator . '/' . $denominator;
    }

    /**
     * Applies the given decimal format info to the given number.
     *
     * @param  mixed             $number
     * @param  NumberFormatToken $token
     * @return string
     */
    private static function applyDecimalFormat($number, $token)
    {
        $formatted_number = $number;

        // Handle thousands scaling.
        if ($token->getThousandsScale() > 0) {
            $formatted_number /= pow(1000, $token->getThousandsScale());
        }

        // Apply maximum characters behind decimal symbol limit.
        $formatted_number = number_format($formatted_number, strlen($token->getFormatRight()), '.', '');

        // Remove minus sign for now, as it requires explicit handling.
        $formatted_number = str_replace('-', '', $formatted_number);

        // Remove insignificant zeroes for now, we will (re-)add them based on format_info next.
        if (strpos($formatted_number, '.') !== false) {
            $formatted_number = rtrim($formatted_number, '0');
        }

        // Split number into pre-decimal and post-decimal.
        $number_parts = explode('.', $formatted_number);

        // Handle left side of decimal point. (Note array_reverse() usage; We work from the decimal point to the left.)
        $left_side_chars = array_reverse(str_split($number_parts[0]));
        $format_chars = array_reverse(str_split(str_replace('?', ' ', $token->getFormatLeft())));
        $number_left = '';
        if (strlen($token->getFormatLeft()) < strlen($number_parts[0]) && $number_parts[0] !== '0') {
            // Pre-decimal format string is too short. Full pre-decimal number has to always be included in output.
            $number_left = $number_parts[0];
        } else {
            for ($i = 0; $i < strlen($token->getFormatLeft()); $i++) {
                if (isset($left_side_chars[$i])) {
                    $number_left = $left_side_chars[$i] . $number_left; // Add digit here.
                } elseif ($format_chars[$i] !== '#') {
                    $number_left = $format_chars[$i] . $number_left; // Add filler character here.
                }
            }
        }

        // Handle right side of decimal point.
        $number_right = '';
        if (count($number_parts) > 1) {
            $right_side_chars = str_split($number_parts[1]);
            if ($right_side_chars[0] === '') { // Side-effect of str_split('')
                $right_side_chars = array();
            }
            $format_chars = str_split(str_replace('?', ' ', $token->getFormatRight()));
            if ($format_chars[0] === '') { // Side-effect of str_split('')
                $format_chars = array();
            }
            for ($i = 0; $i < strlen($token->getFormatRight()); $i++) {
                if (isset($right_side_chars[$i])) {
                    $number_right .= $right_side_chars[$i]; // Add digit here.
                } elseif ($format_chars[$i] !== '#') {
                    $number_right .= $format_chars[$i]; // Add filler character here.
                }
            }
        }

        $formatted_number = $number_left . ($number_right !== '' ? ('.' . $number_right) : '');

        // Place thousands separators.
        if ($token->useThousandsSeparators()) {
            $number_parts = explode('.', $formatted_number);
            $number_left = $number_parts[0];
            $number_left_with_separators = '';
            while (strlen($number_left) > 3) {
                if (substr($number_left, -4, 1) === ' ') {
                    // Handle edge case: format: [?,000], value: [123], result: [  123], wrong: [ 123], also wrong: [ ,123]
                    $number_left_with_separators = ' ' . substr($number_left, -3) . $number_left_with_separators;
                } else {
                    // No edge case, just add the separator character.
                    $number_left_with_separators = ',' . substr($number_left, -3) . $number_left_with_separators;
                }
                $number_left = substr($number_left, 0, strlen($number_left) - 3);
            }
            $number_left_with_separators = $number_left . $number_left_with_separators;
            if (count($number_parts) > 1) {
                $formatted_number = $number_left_with_separators . '.' . $number_parts[1];
            } else {
                $formatted_number = $number_left_with_separators;
            }
        }

        return $formatted_number;
    }

    /**
     * Gets the format section to be used for the given value with the given format_index.
     *
     * @param  string $value
     * @param  int    $format_index
     * @return NumberFormatSection
     *
     * @throws RuntimeException
     */
    private function getFormatSectionForValue($value, $format_index)
    {
        $sections = $this->getFormatSections($format_index);
        return $this->getSectionForValue($value, $sections);
    }

    /**
     * @param  string                $value
     * @param  NumberFormatSection[] $sections_ordered
     * @return NumberFormatSection
     *
     * @throws RuntimeException
     */
    private function getSectionForValue($value, $sections_ordered)
    {
        foreach ($sections_ordered as $section) {
            switch ($section->getPurpose()) {
                case 'default':
                    // "default" counts for everything.
                    return $section;
                case 'default_number':
                    if (is_numeric($value)) {
                        return $section;
                    }
                    break;
                case 'default_text':
                    if (!is_numeric($value)) {
                        return $section;
                    }
                    break;
                default:
                    // condition format, can only be applied to numbers
                    if (!is_numeric($value)) {
                        break;
                    }
                    if (!preg_match('{([<>=]+)([-+]?)(\d+)}', $section->getPurpose(), $matches)) {
                        throw new RuntimeException('Unexpected section purpose: [' . $section->getPurpose() . ']');
                    }
                    $comparison_operator = $matches[1];
                    $comparison_value_sign = $matches[2];
                    $comparison_value = $matches[3];

                    if ($comparison_value_sign === '-') {
                        $comparison_value *= -1;
                    }

                    if ($comparison_operator === '<>') {
                        if ($comparison_value != $value) {
                            return $section;
                        }
                    } else {
                        if (   (strpos($comparison_operator, '=') !== false && $value == $comparison_value)
                            || (strpos($comparison_operator, '>') !== false && $value > $comparison_value)
                            || (strpos($comparison_operator, '<') !== false && $value < $comparison_value)
                        ) {
                            return $section;
                        }
                    }
                    break;
            }
        }

        // Section purpose assignment should ensure a section for any type of value, so this *should* never happen.
        throw new RuntimeException('No format found for value [' . $value . ']');
    }

    /**
     * Gets the format data for the given format index.
     *
     * @param  int $format_index
     * @return NumberFormatSection[]
     *
     * @throws RuntimeException
     */
    private function getFormatSections($format_index)
    {
        if (isset($this->parsed_format_cache[$format_index])) {
            return $this->parsed_format_cache[$format_index];
        }

        // Look up base format definition via format index.
        $format_code = null;
        if (array_key_exists($format_index, $this->customized_formats)) {
            $format_code = $this->customized_formats[$format_index];
        } elseif (array_key_exists($format_index, self::BUILTIN_FORMATS)) {
            $format_code = self::BUILTIN_FORMATS[$format_index];
        } elseif (isset($this->number_formats[$format_index])) {
            $format_code = $this->number_formats[$format_index];
        }

        if ($format_code === null) {
            // Definition for requested format_index could not be found.
            throw new RuntimeException('format with index [' . $format_index . '] was not defined.');
        }

        $sections = $this->prepareFormatSections($format_code);
        foreach ($sections as $section_index => $section) {
            // Remove color and condition definitions. (They're either useless to us, or pre-parsed already.)
            $this->removeColorsAndConditions($section);

            // Date/Time formats are handled differently than decimal/fraction formats.
            if ($this->isDateTimeFormat($section)) {
                $this->prepareDateTimeFormat($section);
            } else {
                // Any format matching negative AND positive values gets a minus sign prepended to allow value differentiation.
                $prepend_minus_sign = true;
                if (preg_match('{[><=]+[+-]?\d+}', $section->getPurpose())) {
                    if (strpos($section->getPurpose(), '=') !== false) {
                        // Exact equality check implies that no value differentiation is needed.
                        $prepend_minus_sign = false;
                    } elseif (strpos($section->getPurpose(), '>') !== false && strpos($section->getPurpose(), '-') === false) {
                        // For only-positive formats, do not prepend minus sign. (Technically unnecessary, just for clarity.)
                        $prepend_minus_sign = false;
                    } elseif (strpos($section->getPurpose(), '<') !== false && strpos($section->getPurpose(), '-') !== false) {
                        // For only-negative formats, do not prepend minus sign.
                        $prepend_minus_sign = false;
                    } elseif ($section->getPurpose() === '<0') {
                        // Addendum to previous case: 0 is the only "non-negative" value that can still ensure matching of only negative values.
                        $prepend_minus_sign = false;
                    }
                }
                $section->setPrependMinusSign($prepend_minus_sign);

                $this->prepareNumericFormat($section);
            }

            // Values of percentage formats need to be handled differently later.
            $section->setIsPercentage($this->detectIfPercentage($section));
        }

        // Update cached data, for faster retrieval next time round.
        $this->parsed_format_cache[$format_index] = $sections;

        return $sections;
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
     * Split the given $format_string into sections, which are further split into tokens to assist in parsing.
     * Also determines each format section's purpose (read: which value ranges it should be applied to).
     *
     * @param  string $format_string The full format string, including all sections, completely unparsed
     * @return NumberFormatSection[]
     */
    private function prepareFormatSections($format_string)
    {
        $sections_tokenized = array();
        foreach ($this->splitSections($format_string) as $section_index => $section_string) {
            $sections_tokenized[$section_index] = new NumberFormatSection(
                $this->convertFormatSectionToTokens($section_string)
            );
        }
        return $this->assignSectionPurposes($sections_tokenized);
    }

    /**
     * Splits the given number format string into sections. (Format for positive values, for negative values, etc.)
     * Does not identify the actual purpose of each section. (For example in case of conditional sections.)
     *
     * @param  string $format_string
     * @return array  List of found sections, as substrings of $format_string.
     *
     * @throws RuntimeException
     */
    private function splitSections($format_string)
    {
        $offset = 0;
        $start_pos = 0;
        $in_quoted = false;
        $sections = array();
        while (preg_match('{[;"]}', $format_string, $matches, PREG_OFFSET_CAPTURE, $offset)) {
            $match_character = $matches[0][0];
            $match_offset = $matches[0][1];
            $is_escaped = !$in_quoted && $match_offset > 0 && substr($format_string, $match_offset - 1, 1) === '\\';
            switch ($match_character) {
                case '"':
                    // Quote symbols (unless escaped) toggle the "quoted" scope on/off.
                    if (!$is_escaped) {
                        $in_quoted = !$in_quoted;
                    }
                    $offset = $match_offset + 1;
                    break;
                case ';':
                    // Semicolons act as format definition splitters (unless escaped or quoted).
                    if (!$in_quoted && !$is_escaped) {
                        $sections[] = substr($format_string, $start_pos, $match_offset - $start_pos);
                        $start_pos = $match_offset + 1;
                    }
                    $offset = $match_offset + 1;
                    break;
                default:
                    throw new RuntimeException(
                        'Unexpected character [' . $match_character . '] matched at position [' . $match_offset . '].'
                    );
                    break;
            }
        }

        // Add sub-format trailing the last semicolon (or the whole format string, if no semicolon was found).
        if ($start_pos < strlen($format_string)) { // Only if there are leftover characters.
            $sections[] = substr($format_string, $start_pos);
        }

        return $sections;
    }

    /**
     * Splits the given format section into tokens based on logical context, such as quoted/escaped portions.
     *
     * @param  string $section_string The section string to parse.
     * @return NumberFormatToken[]
     *
     * @throws RuntimeException
     */
    private function convertFormatSectionToTokens($section_string)
    {
        /** @var NumberFormatToken[] $tokens */
        $tokens = array();
        $offset = 0;
        $last_tokenized_character = -1;
        $is_quoted = false;
        $is_square_bracketed = false;
        $square_bracket_index = -1;
        while ($offset < strlen($section_string)
            && preg_match('{["\\\\[\\]]}', $section_string, $matches, PREG_OFFSET_CAPTURE, $offset)
        ) {
            $match_character = $matches[0][0];
            $match_offset = $matches[0][1];

            if (in_array($match_character, array('\\', '"')) || !$is_quoted) { // Read: for "[" and "]": Only if !$is_quoted.
                // Add token between last match and this match.
                if ($last_tokenized_character < $match_offset - 1) {
                    $last_token = substr(
                        $section_string,
                        $last_tokenized_character + 1,
                        $match_offset - ($last_tokenized_character + 1)
                    );
                    $tokens[] = (new NumberFormatToken($last_token))
                        ->setIsQuoted($is_quoted && !$is_square_bracketed)
                        ->setSquareBracketIndex($is_square_bracketed ? $square_bracket_index : null);
                }
            }

            switch ($match_character) {
                case '\\':
                    if ($is_quoted || $is_square_bracketed) {
                        // Backslashes cannot escape anything when within quotes/square brackets. Output as-is. (In: \\ - Out: \\)
                        $tokens[] = (new NumberFormatToken('\\'))
                            ->setIsQuoted(!$is_square_bracketed)
                            ->setSquareBracketIndex($is_square_bracketed ? $square_bracket_index : null);
                        $last_tokenized_character = $match_offset;
                    } else {
                        // This backslash will escape whatever follows it. (In: \\ - Out: \)
                        $escaped_character = substr($section_string, $match_offset + 1, 1);
                        $tokens[] = (new NumberFormatToken($escaped_character))
                            ->setIsQuoted(true)
                            ->setSquareBracketIndex($is_square_bracketed ? $square_bracket_index : null);

                        // Move offset to beyond the escaped character, to implicitly avoid it being matched in next loop iteration.
                        $last_tokenized_character = $match_offset + 1;
                    }
                    $offset = $last_tokenized_character + 1;
                    break;
                case '"':
                    if ($is_square_bracketed) {
                        // Quotes are ineffective in square-bracketed areas.
                        $tokens[] = (new NumberFormatToken('"'))
                            ->setSquareBracketIndex($square_bracket_index);
                    } else {
                        // Flip $is_quoted state.
                        $is_quoted = !$is_quoted;
                    }
                    $last_tokenized_character = $match_offset;
                    $offset = $match_offset + 1;
                    break;
                case '[':
                case ']':
                    if (!$is_quoted) {
                        if ($match_character === '[' && $is_square_bracketed) {
                            // Opening square brackets in square bracket areas must be included in output as-is.
                            $tokens[] = (new NumberFormatToken('['))
                                ->setSquareBracketIndex($square_bracket_index);
                            $last_tokenized_character = $match_offset; // Character has been included in output.
                        } else {
                            // Set $is_square_bracketed state according to actually matched character.
                            if ($match_character === '[' && !$is_square_bracketed) {
                                $square_bracket_index++;
                            }
                            $is_square_bracketed = ($match_character === '[');
                            $last_tokenized_character = $match_offset; // Do not include this character in output.
                        }
                    } // else: Will be included in next "add token between last and this match" execution.
                    $offset = $match_offset + 1;
                    break;
                default:
                    throw new RuntimeException(
                        'Unexpected character [' . $match_character . '] matched at position [' . $match_offset . '].'
                    );
                    break;
            }
        }

        // Handle token following the last matched character (or the entire format string in case of 0 matches).
        if ($last_tokenized_character < strlen($section_string) - 1) {
            $last_token = substr(
                $section_string,
                $last_tokenized_character + 1,
                strlen($section_string) - ($last_tokenized_character + 1)
            );

            /* Note: There is no check for unclosed quoted areas. Behavior in case of this type of fault is undefined,
             * and even differs between modern applications. As such, just regard such areas as quoted and continue. */
            $tokens[] = (new NumberFormatToken($last_token))
                ->setIsQuoted($is_quoted)
                ->setSquareBracketIndex($is_square_bracketed ? $square_bracket_index : null);
        }

        // Cleanup tokens, merging successive tokens with identical rule-sets together.
        /** @var NumberFormatToken[] $tokens_merged */
        $tokens_merged = array();
        $current_index = 0;
        foreach ($tokens as $token) {
            // First loop iteration requires explicit handling.
            if (!isset($tokens_merged[$current_index])) {
                $tokens_merged[$current_index] = $token;
                continue;
            }

            // Group successive tokens with the same rule-sets together.
            if ($token->isQuoted() === $tokens_merged[$current_index]->isQuoted()
                && $token->getSquareBracketIndex() === $tokens_merged[$current_index]->getSquareBracketIndex()
            ) {
                $tokens_merged[$current_index]->appendCode($token->getCode());
            } else {
                $current_index++;
                $tokens_merged[$current_index] = $token;
            }
        }

        return $tokens_merged;
    }

    /**
     * Determines the purpose of each given section(, read: to which types/ranges of values it should be applied)
     * and adds it to the CellFormatSection instance data.
     *
     * @param  NumberFormatSection[] $sections
     * @return NumberFormatSection[] Same as input, but extended with $purpose and additional (default/duplicate) formats,
 *                                 ordered by priority of applicability.
     */
    private function assignSectionPurposes($sections)
    {
        /* Some basic rules on format purposes and default formats:
         *  - Format sections are checked for applicability in order of appearance.
         *      -> If 2 sections are applicable, the leftmost (first) section is used.
         *  - Without conditions, the ordering of purposes is: >0, <0, =0, default_text
         *      -- If <0 or =0 are not given, the >0 format does double-duty as the default_number format.
         *  - If only 1 section is given, it is applied to all numbers, even if were a text-only format otherwise.
         *      -- If the section is a text format, it is applied to text as well.
         *  - If 2-3 sections are given, the last section CAN be a text-only format, making it default_text instead of <0 or =0.
         *  - With a condition in the first section, the ordering of purposes is: condition, <0, default_number, default_text
         *      -- If default_number is not given, the <0 format is applied to all numbers not matching the condition.
         *  - With a condition in the first and second section, the ordering of purposes is: condition1, condition2, default_number, default_text
         *      -- If default_number is not given, or the 3rd element is a text-only format, numbers not matching any condition
         *         are output as ########.
         *  - With a condition in the second section (and not in the first), the ordering of purposes is: >0, condition, default_number, default_text
         *      -- Note that sections are still checked in order, not by relevance. So if the condition only accepts positive
         *         values, the format with a condition is never used.
         *      -- If default_number is not given, or the 3rd section is a text-only format, numbers not matching any condition
         *         are output as ########.
         *  - If default_text is not given, the default format for text is @
         */

        // Formats are to be checked for applicability in order of appearance. Replicate this behavior via element ordering.
        /** @var NumberFormatSection[] $sections_in_order */
        $sections_in_order = array();

        $section_purposes_ordered = array('>0', '<0', '=0', 'default_text'); // default sub-format semantic
        $section_purpose_index = 0;
        $contains_condition = false; // If true, the default "positive,negative,zero,text" semantic is overridden.
        $found_purposes = array();
        foreach ($sections as $section) {
            $section_tokens = $section->getTokens();
            $format_type = $this->detectFormatType($section_tokens);
            if (!$format_type) {
                if (count($section_tokens) === 0) {
                    // Empty format. Indicates not showing anything for this value type. Treat as number format.
                    $format_type = 'number';
                } else {
                    // Faulty format, cannot be applied to anything. Skip this purpose index and try the next.
                    $section_purpose_index++;
                    continue;
                }
            }
            $condition = $this->detectCondition($section_tokens);
            if ($condition) {
                $section->setPurpose($condition);
                $sections_in_order[] = $section;
                $contains_condition = true;
            } else {
                if (count($sections) === 1) {
                    // Shortcut: Single-section format string with no condition is default_number.
                    $purpose = 'default_number';
                } elseif ($contains_condition && count($sections) === 2 && $section_purpose_index === 1) {
                    // Shortcut: Second section of condition-containing format string with no further sections is default.
                    $purpose = 'default_' . $format_type;
                } elseif ($contains_condition && $section_purpose_index >= 2) {
                    // In case of condition, the 3rd and 4th sections are default, type-specific formats.
                    if (count($sections) === 3 && $format_type === 'text') {
                        // Special case: [condition];[default_number];[default_text]
                        $purpose = 'default_text';
                        $sections_in_order[1]->setPurpose('default_number');
                    } else {
                        $purpose = 'default_' . $format_type;
                    }
                } elseif ($format_type === 'text') {
                    // 2-4 formats: Any "text-only" format is not applied to numbers.
                    $purpose = 'default_text';
                } else {
                    $purpose = $section_purposes_ordered[$section_purpose_index];
                }

                $section->setPurpose($purpose);
                $sections_in_order[] = $section;
                $found_purposes[$purpose] = $section;

                if (count($sections) === 1 && $format_type === 'text') {
                    // Shortcut: If the only section contains text-format-only elements, it does double-duty for text values.
                    $new_section = clone $section;
                    $new_section->setPurpose('default_text');
                    $sections_in_order[] = $new_section;
                    $found_purposes['default_text'] = $new_section;
                }
            }
            $section_purpose_index++;
        }

        // In case of default ordering, allow usage of "positive" format for any missing numerical formats.
        if (!$contains_condition && isset($found_purposes['>0']) && (!isset($found_purposes['<0']) || !isset($found_purposes['=0']))) {
            $new_section = clone $found_purposes['>0'];
            $new_section->setPurpose('default_number');
            $sections_in_order[] = $new_section;
            $found_purposes['default_number'] = $new_section;
        }

        if (!isset($found_purposes['default_number'])) {
            /* No default format for non-matching numeric values given. If a non-matching number should be output anyway,
             * output # "across the width of the cell" to indicate an error. */
            $new_token_list = array(
                (new NumberFormatToken('########'))
                    ->setIsQuoted(true)
            );
            $sections_in_order[] = new NumberFormatSection($new_token_list, 'default_number');
        }
        if (!isset($found_purposes['default_text'])) {
            // Default format for text values, if no other format for text values is specified.
            $new_token_list = array(
                new NumberFormatToken('@')
            );
            $sections_in_order[] = new NumberFormatSection($new_token_list, 'default_text');
        }

        return $sections_in_order;
    }

    /**
     * Checks the given section for a format condition and returns it, if one is found.
     *
     * @param  NumberFormatToken[] $section_tokens
     * @return string|null
     */
    private function detectCondition($section_tokens)
    {
        $condition = null;
        foreach ($section_tokens as $token) {
            // Condition has to appear within square-bracketed areas.
            if ($token->isInSquareBrackets()) {
                if (preg_match('{^[<>=]+[-+]?\d+$}', $token->getCode(), $matches)) {
                    // Condition found. There may only be one condition per section, so we can break; after this one.
                    $condition = $matches[0];
                    break;
                }
            }
        }
        return $condition;
    }

    /**
     * Checks the given section for indicators of it being used for a particular value type.
     *
     * @param  NumberFormatToken[] $section_tokens
     * @return string|null Either "text", "number" or null.
     */
    private function detectFormatType($section_tokens)
    {
        foreach ($section_tokens as $tokens) {
            if (!$tokens->isQuoted() && !$tokens->isInSquareBrackets()) {
                if (strpos($tokens->getCode(), '@') !== false) {
                    return 'text';
                }
                if (preg_match('{[0#?ymdhsa]}', strtolower($tokens->getCode()))) {
                    return 'number'; // note: This is also used for date formats.
                }
            }
        }
        return null; // Format type uncertain, should probably not be applied.
    }

    /**
     * Removes tokens unnecessary for our particular parsing intentions from the given section.
     *
     * @param NumberFormatSection $section
     */
    private function removeColorsAndConditions($section)
    {
        // Note: The color/conditions definitions are usually at the start of the section. e.g.: "[red][<1000]0,00"
        $tokens = $section->getTokens();
        foreach ($tokens as $token_index => $token) {
            if ($token->isInSquareBrackets()) {
                // The only thing found in square brackets that we need to keep is the currency string.
                if (strpos($token->getCode(), '$') !== 0) { // $ at pos 0 indicates currency string.
                    unset($tokens[$token_index]);
                }
            }
        }

        $section->setTokens(array_values($tokens)); // array_values() removes gaps in the array index list.
    }

    /**
     * Checks if the given section is a date- or time-format.
     *
     * @param  NumberFormatSection $section
     * @return bool
     */
    private function isDateTimeFormat($section)
    {
        // Note: Either of these formats are exclusive. For example: You can't use decimal and fraction in the same format.
        foreach ($section->getTokens() as $token) {
            if (!$token->isQuoted() && !$token->isInSquareBrackets()) {
                if (preg_match('#[ymdhsa]#', strtolower($token->getCode()))) {
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * Checks if the given section requests usage of percentage values.
     *
     * @param  NumberFormatSection $section
     * @return bool
     */
    private function detectIfPercentage($section)
    {
        foreach ($section->getTokens() as $token) {
            if (!$token->isQuoted()
                && !$token->isInSquareBrackets()
                && strpos($token->getCode(), '%') !== false
            ) {
                return true;
            }
        }
        return false;
    }

    /**
     * Prepares the given decimal/fraction section data for easier parsing while reading values.
     * Adds discovered number formatting information to the given section.
     *
     * @param  NumberFormatSection $section
     */
    private function prepareNumericFormat($section)
    {
        $tokens = $section->getTokens();
        foreach ($tokens as $token_index => $token) {
            if ($token->isQuoted()) {
                continue; // Nothing to do for quoted/escaped areas.
            }

            $code = $token->getCode();

            // For currency/language info: Keep currency symbol, remove the rest.
            if ($token->isInSquareBrackets()) {
                if (strpos($code, '$') === 0) {
                    preg_match('{\$([^-]*)}', $code, $matches);
                    $tokens[$token_index]
                        ->setCode($matches[1])
                        ->setIsQuoted(true)
                        ->setSquareBracketIndex(null);
                }
                continue; // No other parsing requirements for square bracketed areas are known.
            }

            // Handle _ character. (=> Skip width of next character; In our case: Replace next character with space.)
            $code = preg_replace('{_.}', ' ', $code);

            // Handle * character. (=> "Repeat next character until column is filled".)
            // Purposefully ignored here, due to there not being a fixed column size to fill against.
            $code = str_replace('*', '', $code);

            // Pre-parse decimal/fraction format string (e.g.: [##,##0.0???,] or [0/0])
            if (preg_match('{[0#?.,/ ]+}', $code, $matches)) {
                $number_format_code = $matches[0];
                if (strpos($matches[0], '/') !== false) {
                    // This is a fraction format definition.
                    $tokens[$token_index]
                        ->setNumberType('fraction')
                        ->setExtractWhole((bool) preg_match('{[0#?] [0#?]+/}', $code));
                } else {
                    // This is a decimal format definition.
                    /* One-regex-catches-all does not work, as *every* part is optional (which would mean that
                     * such a regex would match an empty string). Instead, do multiple regex checks.
                     * Check for dot-required (read: decimal) format first.
                     * If it doesn't match, check non-dot (read: whole-value) format next.
                     * Both preg_match() calls should result in the same semantics for the resulting $matches array,
                     * so the then-following code can interpret the results regardless of which regex matches.
                     * (Hence the odd-looking empty matching group in the second regex.) */
                    if (!preg_match('{((?:[0?#]|,(?=[0?#]))*)\.([0?#]*)(,*)}', $number_format_code, $matches)) {
                        if (!preg_match('{((?:[0?#]|,(?=[0?#]))+)()(,*)}', $number_format_code, $matches)) {
                            continue; // Not a number format. We matched an empty string in the initial preg_match().
                        }
                    }
                    $tokens[$token_index]
                        ->setNumberType('decimal')
                        ->setFormatLeft(str_replace(',', '', $matches[1]))
                        ->setFormatRight($matches[2])
                        ->setThousandsScale(strlen($matches[3]))
                        ->setUseThousandsSeparators(strpos($matches[1], ',') !== false);
                }
            }

            $tokens[$token_index]->setCode($code);
        }

        $section->setTokens($tokens);
    }

    /**
     * Prepares the given date/time section data for easier parsing while reading values.
     * Also determines the more specific date/time/datetime type.
     *
     * @param NumberFormatSection $section
     */
    private function prepareDateTimeFormat($section)
    {
        // Determine if the contained time data should be displayed in 12h format.
        $time_12h = false;
        foreach ($section->getTokens() as $token) {
            if (!$token->isQuoted()
                && !$token->isInSquareBrackets()
                && strpos(strtolower($token->getCode()), 'a') !== false
            ) {
                $time_12h = true;
                break;
            }
        }

        $contains_date = false;
        $contains_time = false;
        $tokens = $section->getTokens();
        foreach ($tokens as $token_index => $token) {
            if ($token->isQuoted()) {
                continue;
            }

            // For currency/language info: Keep currency symbol, remove the rest.
            if ($token->isInSquareBrackets() && strpos($token->getCode(), '$') === 0) {
                preg_match('{\$([^-]*)-\d+}', $token->getCode(), $matches);
                $tokens[$token_index]
                    ->setCode($matches[1])
                    ->setIsQuoted(true)
                    ->setSquareBracketIndex(null);
                continue;
            }

            if (!$token->isInSquareBrackets()) {
                // Translate XLSX date/time code-characters to php date() code-characters.
                $code = strtolower($token->getCode());
                $code = strtr($code, self::DATE_REPLACEMENTS['All']);
                if ($time_12h) {
                    $code = strtr($code, self::DATE_REPLACEMENTS['12H']);
                } else {
                    $code = strtr($code, self::DATE_REPLACEMENTS['24H']);
                }
                $tokens[$token_index]->setCode($code);

                // Determine more specific date/time/datetime specification.
                $contains_date = $contains_date || preg_match('#[DdFjlmMnoStwWYyz]#u', $code);
                $contains_time = $contains_time || preg_match('#[aABgGhHisuv]#u', $code);
            }
        }

        // Determine more specific date/time/datetime specification.
        if ($contains_date && $contains_time) {
            $section->setDateTimeType('datetime');
        } elseif ($contains_date) {
            $section->setDateTimeType('date');
        } elseif ($contains_time) {
            $section->setDateTimeType('time');
        }

        $section->setTokens($tokens);
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
    }
}
