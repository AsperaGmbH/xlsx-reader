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

    /** @var ReaderConfiguration */
    private $configuration;

    /** @var array List of number formats defined by the current XLSX file; array key = format index */
    private $number_formats = array();

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
     * @param ReaderConfiguration $configuration
     */
    public function __construct($configuration)
    {
        $this->configuration = $configuration;
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
            // ID 0 = "general" format
            return $this->applyGeneralFormat($value);
        }

        // Get definition of format for the given format_index.
        $section = $this->getFormatSectionForValue($value, $num_fmt_id);

        // If percentage values are expected, multiply value accordingly before formatting.
        if ($section->isPercentage() && !$this->configuration->getReturnPercentageDecimal()) {
            $value = (string) ($value * 100);
        }

        // If formatting is not desired, return value as-is.
        if ($this->configuration->getReturnUnformatted()
            || ($section->isPercentage() && $this->configuration->getReturnPercentageDecimal())) {
            return $value;
        }

        if ($section->getNumberType() === 'decimal') {
            return $this->applyDecimalFormat($value, $section);
        }

        if ($section->getNumberType() === 'fraction') {
            return $this->applyFractionFormat($value, $section);
        }

        if ($section->getDateTimeType()) {
            return $this->applyDateTimeFormat($value, $section);
        }

        return $this->applyTextFormat($value, $section);
    }

    /**
     * Applies the "general" format (format id = 0) to the value.
     *
     * @param  mixed $value
     * @return mixed
     */
    private function applyGeneralFormat($value)
    {
        if (preg_match('/\d+(?:\.\d+)E[+-]\d+?/u', $value) === 1) {
            // Values stored using e-notation should be converted to decimals for display.
            $value = rtrim(sprintf('%.10F', floatval($value)), '0');
            $value = rtrim($value, '.'); // Note: Do not combine this with the trim above.
        }

        /* Note: At this point, the "general" format may actually use scientific notation for display after all,
         * if some conditions related to the cell's width apply. However, this reader doesn't care about maximum
         * cell width, so this step is intentionally skipped here. All values are output in decimal format. */

        return $value;
    }

    /**
     * Applies requested formatting to values that do not contain any decimal, fraction or date/time format information.
     *
     * @param  mixed               $value
     * @param  NumberFormatSection $section
     * @return string
     */
    private function applyTextFormat($value, $section)
    {
        // Apply format to value, going token by token.
        $output = '';
        foreach ($section->getTokens() as $token) {
            if ($token->isQuoted()) {
                $output .= $token->getCode();
                continue;
            }

            if ($token->isInSquareBrackets()) {
                continue; // Nothing to do here.
            }

            // Handle @ symbol.
            $output .= str_replace('@', $value, $token->getCode());
        }

        return $output;
    }

    /**
     * Included subtypes: Scientific format, currency format
     *
     * @param  mixed               $value
     * @param  NumberFormatSection $section
     * @return string
     */
    private function applyDecimalFormat($value, $section)
    {
        // If scientific formatting is expected, multiply/divide value up/down to the correct exponent.
        $e_exponent = null;
        $exponent_formatted = '';
        if ($section->getExponentFormat() !== '' // else: non-scientific format, or e.g. [0.0E+]
            && $section->getDecimalFormat() !== '.' // else: e.g. [.E+0] - Nonsensical format. Fallback to non-scientific format.
        ) {
            $e_exponent = $this->getExponentForValueAndSection($value, $section);
            $exponent_formatted = $this->mergeValueIntoNumericalFormat(abs($e_exponent), $section->getExponentFormat());
            $value *= 10 ** ($e_exponent * -1);
        }

        // Convert value to decimal, preparing it for inclusion in the actual format string.
        $decimal_formatted = $this->formatDecimalValue($value, $section);

        // Formatted value's whole value may be of different length than format. Calculate difference.
        // (e.g.: "12345" formatted as "0.0" becomes "12345.0" -> 4 more characters in whole value than expected)
        $decimal_without_commas = str_replace(',', '', $decimal_formatted); // Remove thousands separators; Avoids various issues.
        $decimal_without_commas_decimal_position = strpos($decimal_without_commas, '.');
        if ($decimal_without_commas_decimal_position !== false) {
            $decimal_left = substr($decimal_without_commas, 0, $decimal_without_commas_decimal_position);
        } else {
            $decimal_left = substr($decimal_without_commas, 0);
        }
        $format_left_length_diff = strlen($decimal_left) - strlen(str_replace(',', '', $section->getFormatLeft()));

        // Merge decimal-converted value into format, going token by token.
        $output = '';
        $e_exponent_token_expected = false;
        $e_exponent_pos = 0;
        $extra_decimal_digits_included = false;
        $num_skipped_token_characters = 0;
        $offset_in_section_decimal_format = 0; // Offset of processing of characters in $section->getDecimalFormat()
        foreach ($section->getTokens() as $token) {
            if ($token->isQuoted()) {
                $output .= $token->getCode();
                continue;
            }

            if ($token->isInSquareBrackets()) {
                continue; // Nothing to do here.
            }

            if ($token->isScientificNotationE()) {
                // "E+"/"e-" token. What follows is a number format for the exponent.
                $e_exponent_token_expected = true;

                // Check if exponent is actually determined. Otherwise, drop E+00 from output as error fallback.
                if ($e_exponent !== null) {
                    $output .= substr($token->getCode(), 0, 1);
                    if ($e_exponent >= 0 && substr($token->getCode(), 1, 1) === '+') {
                        $output .= '+';
                    } else if ($e_exponent < 0) {
                        $output .= '-';
                    }
                }
                continue;
            }

            // Approach of applying format: Use format code as a base, then replace its functional parts with values.
            $token_code = $token->getCode();

            if ($e_exponent_token_expected && preg_match('{[0#?]}', $token_code)) {
                // This is the exponent portion of the scientific format.
                if ($e_exponent === null) {
                    // ... but the scientific format specification is invalid. Error fallback: Ignore exponent token(s).
                    $token_code = preg_replace('{[0#?]}', '', $token_code);
                } else {
                    // Check for exponent that's longer than format allows. (Short ones are already handled fine.)
                    $e_format_len_diff = max(strlen($exponent_formatted) - strlen($section->getExponentFormat()), 0);

                    // Iterate through all characters of output_part and replace exponent characters.
                    $output_replaced = '';
                    for ($i = 0; $i < strlen($token_code); $i++) {
                        $character = substr($token_code, $i, 1);
                        if (in_array($character, array('0', '#', '?'), true)) {
                            if ($e_exponent_pos === 0) {
                                // In case that exponent is longer than format allows, output additional symbols first.
                                $output_replaced .= substr($exponent_formatted, 0, $e_format_len_diff + 1);
                                $e_exponent_pos += $e_format_len_diff + 1;
                            } else {
                                $output_replaced .= substr($exponent_formatted, $e_exponent_pos, 1);
                                $e_exponent_pos++;
                            }
                        } else if (!in_array($character, array('.', ','), true)) { // Drop ., from output in exponent.
                            $output_replaced .= $character;
                        }
                    }

                    // Exponent formatting completed. Move to output directly, skip decimal formatting steps.
                    $output .= $output_replaced;
                    continue;
                }
            }

            // Remove thousands-separators and scaling symbols from format to ease processing.
            $token_code = str_replace(',', '', $token_code);

            // Replace decimal format characters with decimal value characters, going left-to-right.
            $output_part = '';
            $decimal_part = '';
            $formatter_offset = 0; // Used to move the preg_match() loop forward.
            while (preg_match('{[0#?.]}', $token_code, $matches, PREG_OFFSET_CAPTURE, $formatter_offset)) {
                // Include skipped characters before matched decimal character(s) in output.
                $match_offset = $matches[0][1];
                $output_part .= substr(
                    $token_code,
                    $formatter_offset,
                    $match_offset - $formatter_offset
                );
                $formatter_offset = $match_offset + 1;

                // Replace decimal format character with formatted decimal value part.
                if ($format_left_length_diff > 0 && !$extra_decimal_digits_included) {
                    // Formatted value for left side of decimal is longer than format_left. Include additional digits.
                    // (e.g.: value=12345, format='0.0' -> '12345.0')
                    $decimal_part = substr(
                        $decimal_formatted,
                        0,
                        $format_left_length_diff + 1
                    );
                    $offset_in_section_decimal_format += $format_left_length_diff + 1;

                    // Decimal format may have added thousand separators (commas), causing incongruencies in character amounts.
                    // If we just picked up commas in $decimal_part, add this many additional digits to the output, too.
                    $num_commas_in_decimal_part = strlen(preg_replace('{[^,]}', '', $decimal_part));
                    while ($num_commas_in_decimal_part > 0) {
                        $additional_decimal_part = substr(
                            $decimal_formatted,
                            $offset_in_section_decimal_format,
                            $num_commas_in_decimal_part
                        );
                        $decimal_part .= $additional_decimal_part;
                        $offset_in_section_decimal_format += $num_commas_in_decimal_part;

                        // The additional digits we just picked up may have added even more commas.
                        $num_commas_in_decimal_part = strlen(preg_replace('{[^,]}', '', $additional_decimal_part));
                    }

                    $extra_decimal_digits_included = true;
                } else if (
                    $format_left_length_diff < 0
                    && $num_skipped_token_characters < abs($format_left_length_diff)
                ) {
                    // Format is longer than formatted value. Skip this format character.
                    $num_skipped_token_characters++;
                } else {
                    // Replace this format digit with a formatted digit. (Or multiple, if there's a comma in formatted here.)
                    $decimal_part = '';
                    do {
                        $additional_decimal_part = substr($decimal_formatted, $offset_in_section_decimal_format, 1);
                        $decimal_part .= $additional_decimal_part;
                        $offset_in_section_decimal_format++;
                    } while ($additional_decimal_part === ','); // Commas in value must be treated as 0-length digits.
                }

                $output_part .= $decimal_part;
            }

            // Include skipped characters after last decimal character in output.
            $output_part .= substr($token_code, $formatter_offset);

            // Handle @ symbol.
            $output .= str_replace('@', $value, $output_part);
        }

        if ($section->prependMinusSign() && $value < 0) {
            $output = '-' . $output;
        }

        return $output;
    }

    /**
     * Determines the correct multiplication exponent to fit the given value into the given format optimally.
     *
     * @param  mixed $value
     * @param  NumberFormatSection $section
     * @return int
     */
    private function getExponentForValueAndSection($value, $section)
    {
        $e_exponent = 0;
        if ($value < 1) {
            // Value < 1: Decimal point must be moved to the right first.
            preg_match('{\.(0+)}', number_format($value, '99', '.', ''), $matches);
            $num_digits_until_nonzero = isset($matches[1]) ? strlen($matches[1]) : 0;
            $e_exponent = ($num_digits_until_nonzero + 1) * -1;
        }

        // Value >= 1: Decimal point must be moved to the left (if at all).
        // Any value: Fill all spaces to the left of the decimal point optimally.
        $num_pre_decimal_digits_in_format = strlen($section->getFormatLeft()); // Note: Can be 0. (e.g. [.00E+0])
        $num_pre_decimal_digits_in_value = strlen(abs(floor($value)));
        $e_exponent = $e_exponent + ($num_pre_decimal_digits_in_value - $num_pre_decimal_digits_in_format);

        return $e_exponent;
    }

    /**
     * Applies the given, semantical decimal format info (read: not the included token codes) to the given number.
     * Does not consider parts of the format that aren't directly related to decimal formatting.
     * e.g.: For the format [000"m" 000"k" 000.00], only [000000000.00] is considered.
     *
     * @param  mixed               $number
     * @param  NumberFormatSection $section
     * @return string
     */
    private function formatDecimalValue($number, $section)
    {
        $format_left = str_replace(',', '', $section->getFormatLeft());
        $format_right = str_replace(',', '', $section->getFormatRight());

        $formatted_number = $number;

        // Handle thousands scaling.
        if ($section->getThousandsScale() > 0) {
            $formatted_number /= pow(1000, $section->getThousandsScale());
        }

        // Apply maximum characters behind decimal symbol limit.
        $formatted_number = number_format((float) $formatted_number, strlen($format_right), '.', '');

        // Remove minus sign for now, as it requires explicit handling.
        $formatted_number = str_replace('-', '', $formatted_number);

        // Remove insignificant zeroes for now, we will (re-)add them based on format_info next.
        if (strpos($formatted_number, '.') !== false) {
            $formatted_number = rtrim($formatted_number, '0');
        }

        // Split number into pre-decimal and post-decimal.
        $number_parts = explode('.', $formatted_number);

        // Handle left side of decimal point.
        $number_left = $this->mergeValueIntoNumericalFormat($number_parts[0], $format_left);

        // Handle right side of decimal point. (Can't use mergeValueIntoNumericalFormat() here.)
        $number_right = '';
        if (count($number_parts) > 1) {
            $right_side_chars = str_split($number_parts[1]);
            if ($right_side_chars[0] === '') { // Side-effect of str_split('')
                $right_side_chars = array();
            }
            $format_chars = str_split(str_replace('?', ' ', $format_right));
            if ($format_chars[0] === '') { // Side-effect of str_split('')
                $format_chars = array();
            }
            for ($i = 0; $i < strlen($format_right); $i++) {
                if (isset($right_side_chars[$i])) {
                    $number_right .= $right_side_chars[$i]; // Add digit here.
                } elseif ($format_chars[$i] !== '#') {
                    $number_right .= $format_chars[$i]; // Add filler character here.
                }
            }
        }

        $formatted_number = $number_left . ($number_right !== '' ? ('.' . $number_right) : '');

        // Place thousands separators.
        if ($section->useThousandsSeparators()) {
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

        // Edge-case: Commas at start of decimal format are non-functional and must be output as-is.
        if (preg_match('{^(,+)}', $section->getFormatLeft(), $matches)) {
            $formatted_number = $matches[1] . $formatted_number;
        }

        return $formatted_number;
    }

    /**
     * Apply the given fraction format to the given value.
     *
     * @param  mixed               $value
     * @param  NumberFormatSection $section
     * @return string
     */
    private function applyFractionFormat($value, $section)
    {
        // Note: abs() to ease minus-sign handling. Minus sign will be added manually later.
        $fraction_parts = $this->convertNumberToFraction(abs($value), $section);

        // If value has no fraction && the numerator is optional => Output as whole-value instead of fraction.
        // e.g.: value=255.0, format='0 #/0' -> '255'
        $skip_fraction_part = $fraction_parts['numerator'] === 0
            && strpos($section->getFormatLeft(), '0') === false
            && strpos($section->getFormatRight(), '0') === false; // Note: This check is absent in Excel.
        if (!$skip_fraction_part && $fraction_parts['numerator'] === 0) {
            // e.g.: value=255.0, format='0 0/0' -> '255 0/1'
            $fraction_parts['denominator'] = 1;
        }

        // Format whole-value.
        $whole_value_formatted = '';
        if ($fraction_parts['whole'] !== 0 || strpos($section->getWholeValuesFormat(), '0') !== false) {
            // If 1st condition matches: Format the calculated whole-value for output.
            // If 2nd condition matches: No whole-value, but whole-value output is not optional. Output 0.
            // e.g.: value=0.25, format='0 0/0' -> '0 1/4'
            $whole_value_formatted = $this->mergeValueIntoNumericalFormat(
                $fraction_parts['whole'],
                $section->getWholeValuesFormat()
            );
        }
        $whole_value_len_diff = strlen($section->getWholeValuesFormat()) - strlen((string) $whole_value_formatted);

        // There may be no output on whole-value despite there being a format for it. This may trigger a lot of
        // the original format (including unrelated content) to be dropped from output.
        // e.g. value=0.2, format='"A" # "B" 0/0' -> 'A 1/5'
        $exclude_whole_value_part = ($whole_value_formatted === '' && $section->getWholeValuesFormat() !== '');

        // Format numerator and denominator.
        $numerator_formatted = '';
        $denominator_formatted = '';
        $exclude_fraction = ($fraction_parts['numerator'] === 0 && $fraction_parts['denominator'] === 0);
        if (!$exclude_fraction) {
            $numerator_formatted = $this->mergeValueIntoNumericalFormat(
                $fraction_parts['numerator'],
                $section->getFormatLeft()
            );
            $denominator_formatted = $this->mergeValueIntoNumericalFormat(
                $fraction_parts['denominator'],
                $section->getFormatRight()
            );
        }
        $numerator_len_diff = strlen($section->getFormatLeft()) - strlen($numerator_formatted);
        $denominator_len_diff = strlen($section->getFormatRight()) - strlen($denominator_formatted);

        $current_state = ($whole_value_formatted !== '' || $exclude_whole_value_part) ? 'whole' : 'numerator';
        $potential_whole_value_output = '';
        $potential_fraction_output = '';
        $fraction_part_completed = false;

        $offset_in_whole_value = 0;
        $offset_in_whole_value_format = 0;
        $offset_in_numerator = 0;
        $offset_in_numerator_format = 0;
        $offset_in_denominator = 0;
        $offset_in_denominator_format = 0;

        // Bring all the parts together. Read the whole format left-to-right, using the digit characters [0#?].
        $output = '';
        foreach ($section->getTokens() as $token) {
            $code = $token->getCode();

            if ($token->isQuoted()) {
                if ($exclude_whole_value_part && $current_state === 'whole') {
                    // This MAY be part of the whole-value, which we are supposed to exclude. Hold on to this for now.
                    $potential_whole_value_output .= $code;
                } else if ($exclude_fraction && !$fraction_part_completed) {
                    // This MAY be part of the fraction, which we are supposed to exclude. Hold on to this for now.
                    $potential_fraction_output .= $code;
                } else {
                    $output .= $code;
                }
                continue;
            }

            if ($token->isInSquareBrackets()) {
                continue; // Nothing to do here.
            }

            $offset_in_token = 0;
            $code_formatted = '';
            while (preg_match('{[0#?]}', $code, $matches, PREG_OFFSET_CAPTURE, $offset_in_token)) {
                $matched_format_char = $matches[0][0];
                $offset_in_match = $matches[0][1];

                // Move non-matched contents of $code (which we just skipped past) to where they belong.
                $skipped_format_content = substr(
                    $code,
                    $offset_in_token,
                    $offset_in_match - $offset_in_token
                );
                $offset_in_token = $offset_in_match + 1;
                if ($exclude_whole_value_part && $current_state === 'whole') {
                    if ($offset_in_whole_value_format === 0) {
                        // Move non-matched contents of $code to $code_formatted. (It's the content BEFORE the whole-value part.)
                        // e.g. value=0.2, format='"A" # "B" # "C" 0/0' -> 'A 1/5' <- we are including ' "A" ' right now.
                        $output .= $potential_whole_value_output;
                        $potential_whole_value_output = '';
                        $code_formatted .= $skipped_format_content;
                    } else {
                        // Do NOT move non-matched contents of $code to $code_formatted.
                        // e.g. value=0.2, format='"A" # "B" # "C" 0/0' -> 'A 1/5' <- we are dropping ' "B" ' right now.
                        $potential_whole_value_output = '';
                    }

                    $offset_in_whole_value_format++;
                    if ($offset_in_whole_value_format > strlen($section->getWholeValuesFormat())) {
                        // Do NOT move non-matched contents of $code to $code_formatted.
                        // e.g. value=0.2, format='"A" # "B" # "C" 0/0' -> 'A 1/5' <- we are dropping ' "C" ' right now.
                        $current_state = 'numerator';
                    } else {
                        continue;
                    }
                } else if ($exclude_fraction && !$fraction_part_completed) {
                    $potential_fraction_output .= $skipped_format_content;
                } else {
                    $code_formatted .= $skipped_format_content;
                }

                switch ($current_state) {
                    case 'whole':
                        // Handle overflow format characters. e.g. value=5, format=#?#0 #/# -> ' 5'
                        if (($offset_in_whole_value_format - $offset_in_whole_value) < $whole_value_len_diff) {
                            // ? and 0 lead to characters in output. In case of #, skip this character.
                            if ($matched_format_char !== '#') {
                                $code_formatted .= $potential_fraction_output; // Guaranteed whole-value-related now.
                                $potential_fraction_output = '';

                                $code_formatted .= substr($whole_value_formatted, $offset_in_whole_value, 1);
                                $offset_in_whole_value++;
                            }
                            $offset_in_whole_value_format++;
                            break;
                        }
                        $offset_in_whole_value_format++;

                        // Value longer than format? -> Include extra characters.
                        if ($offset_in_whole_value === 0 && $whole_value_len_diff < 0) {
                            $code_formatted .= substr($whole_value_formatted, 0, abs($whole_value_len_diff));
                            $offset_in_whole_value += abs($whole_value_len_diff);
                        }

                        if ($offset_in_whole_value < strlen($whole_value_formatted)) {
                            // "Normal" case. Move this character to $code_formatted.
                            $code_formatted .= $potential_fraction_output; // Guaranteed whole-value-related now.
                            $potential_fraction_output = '';

                            $code_formatted .= substr($whole_value_formatted, $offset_in_whole_value, 1);
                            $offset_in_whole_value++;
                            break;
                        }
                        $current_state = 'numerator';
                    // No break. Fall-through to numerator.
                    case 'numerator':
                        // Handle overflow format characters. e.g. value=5, format=#?#0 #/# -> ' 5'
                        if (($offset_in_numerator_format - $offset_in_numerator) < $numerator_len_diff) {
                            // ? and 0 lead to characters in output. In case of #, skip this character.
                            if ($matched_format_char !== '#') {
                                $code_formatted .= substr($numerator_formatted, $offset_in_numerator, 1);
                                $offset_in_numerator++;
                            }
                            $offset_in_numerator_format++;
                            break;
                        }
                        $offset_in_numerator_format++;

                        // Value longer than format? -> Include extra characters.
                        if ($offset_in_numerator === 0 && $numerator_len_diff < 0) {
                            $code_formatted .= substr($numerator_formatted, 0, abs($numerator_len_diff));
                            $offset_in_numerator += abs($numerator_len_diff);
                        }

                        if ($offset_in_numerator < strlen($numerator_formatted)) {
                            // "Normal" case. Move this character to $code_formatted.
                            $code_formatted .= substr($numerator_formatted, $offset_in_numerator, 1);
                            $offset_in_numerator++;
                            break;
                        }
                        $current_state = 'denominator';
                    // No break. Fall-through to denominator.
                    case 'denominator':
                        // Handle overflow format characters. e.g. value=5, format=#?#0 #/# -> ' 5'
                        if (($offset_in_denominator_format - $offset_in_denominator) < $denominator_len_diff) {
                            // ? and 0 lead to characters in output. In case of #, skip this character.
                            if ($matched_format_char !== '#') {
                                $code_formatted .= substr($denominator_formatted, $offset_in_denominator, 1);
                                $offset_in_denominator++;
                            }
                            $offset_in_denominator_format++;

                            if ($offset_in_denominator_format === strlen($section->getFormatRight())) {
                                // In case of $exclude_fraction, we can now output content as-is again.
                                $fraction_part_completed = true;
                                $potential_fraction_output = '';
                            }
                            break;
                        }
                        $offset_in_denominator_format++;

                        // Value longer than format? -> Include extra characters.
                        if ($offset_in_denominator === 0 && $denominator_len_diff < 0) {
                            $code_formatted .= substr($denominator_formatted, 0, abs($denominator_len_diff));
                            $offset_in_denominator += abs($denominator_len_diff);
                        }

                        if ($offset_in_denominator >= strlen($denominator_formatted)) {
                            // Extra character matches beyond the end of the fraction format. Those are output as-is.
                            $code_formatted .= $matched_format_char;
                        } else {
                            // "Normal" case. Move this character to $code_formatted.
                            $code_formatted .= substr($denominator_formatted, $offset_in_denominator, 1);
                            $offset_in_denominator++;
                        }

                        if ($offset_in_denominator === strlen($denominator_formatted)) {
                            // In case of $exclude_fraction, we can now output content as-is again.
                            $fraction_part_completed = true;
                            $potential_fraction_output = '';
                        }
                        break;
                    default:
                        // This should never happen.
                        throw new Exception('Invalid value for $current_state: [' . $current_state . ']');
                }
            }

            // Move remaining contents of $code to $code_formatted.
            if (!$exclude_whole_value_part || $current_state !== 'whole') {
                $code_after_last_match = substr($code, $offset_in_token);
                if ($code_after_last_match !== false) {
                    if ($exclude_fraction && !$fraction_part_completed) {
                        // This code portion *MAY* need to be dropped from output. We can't be sure about this yet though.
                        $potential_fraction_output .= $code_after_last_match;
                    } else {
                        $code_formatted .= $code_after_last_match;
                    }
                }
            }

            // Handle @ symbol.
            $output .= str_replace('@', $value, $code_formatted);
        }

        if ($section->prependMinusSign() && $value < 0) {
            $output = '-' . $output;
        }

        return $output;
    }

    /**
     * Converts the given value to a fraction and returns the individual parts of this fraction.
     * Does not apply any formatting, but already performs whole-value extraction if the format requires it.
     *
     * @param  mixed               $value
     * @param  NumberFormatSection $section
     * @return array               Keys: whole, numerator, denominator
     */
    private function convertNumberToFraction($value, $section)
    {
        if ($value == (int) $value) {
            // Value is a whole number. Only check to do here is whether to extract it from the fraction or not.
            if ($section->getWholeValuesFormat() === '') {
                return array(
                    'whole' => 0,
                    'numerator' => (int) $value,
                    'denominator' => 1
                );
            }
            return array(
                'whole' => (int) $value,
                'numerator' => 0,
                'denominator' => 0
            );
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
        if ($section->getWholeValuesFormat() !== '' && $value > 1) {
            return array(
                'whole' => (int) floor($value),
                'numerator' => (int) ($numerator % $denominator),
                'denominator' => (int) $denominator
            );
        }
        return array(
            'whole' => 0,
            'numerator' => (int) $numerator,
            'denominator' => (int) $denominator
        );
    }

    /**
     * Formats the given value as a Date/Time value, as requested by the given $section.
     *
     * @param  mixed               $value
     * @param  NumberFormatSection $section
     * @return DateTime|string
     *
     * @throws Exception
     */
    private function applyDateTimeFormat($value, $section)
    {
        $datetime = $this->convertNumberToDateTime($value);

        // Return DateTime objects as-is?
        if ($this->configuration->getReturnDateTimeObjects()) {
            return $datetime;
        }

        // Handle enforced date/time/datetime format.
        switch ($section->getDateTimeType()) {
            case 'date':
                if ($this->configuration->getForceDateFormat() !== null) {
                    return $datetime->format($this->configuration->getForceDateFormat());
                }
                break;
            case 'time':
                if ($this->configuration->getForceTimeFormat() !== null) {
                    return $datetime->format($this->configuration->getForceTimeFormat());
                }
                break;
            case 'datetime':
                if ($this->configuration->getForceDateTimeFormat() !== null) {
                    return $datetime->format($this->configuration->getForceDateTimeFormat());
                }
                break;
            default:
                // Note: Should never happen. Exception is just to be safe.
                throw new RuntimeException('Specific datetime_type for format_index [' . $num_fmt_id . '] is unknown.');
        }

        $output = '';
        foreach ($section->getTokens() as $token) {
            if ($token->isQuoted()) {
                $output .= $token->getCode();
            } else if ($token->isInSquareBrackets()) {
                continue; // Nothing to do here.
            } else {
                $output .= $datetime->format($token->getCode());
            }
        }

        return $output;
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
     * Takes a numerical value and a combination of 0#? symbols and merges the two.
     * $value=5, $format='000' -> '005'
     * $value=5, $format='##?' -> '5'
     * Note: Not to be used for entire decimal formats, e.g. '0.00'
     *
     * @param  mixed  $value
     * @param  string $format
     * @return string
     */
    private function mergeValueIntoNumericalFormat($value, $format)
    {
        if ($format === '' && $value == 0) { // Note: Non-typesafe for $value, as this may be string or float.
            return '';
        }

        $value_formatted = (string) $value;

        // Handle extra digits of format. e.g.: value=5, format='##?0?000' -> ' 0 005'
        $len_diff = strlen($format) - strlen((string) $value);
        if ($len_diff > 0) {
            $format_overflow = str_replace(
                array('#', '?'),
                array('', ' '),
                substr($format, 0, $len_diff)
            );
            $value_formatted = $format_overflow . $value_formatted;
        }

        return $value_formatted;
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
        foreach ($this->getFormatSections($format_index) as $section) {
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
        if (array_key_exists($format_index, $this->configuration->getCustomFormats())) {
            $format_code = $this->configuration->getCustomFormats()[$format_index];
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
            && preg_match('{["\\\\[\\]]|[Ee][+-]}', $section_string, $matches, PREG_OFFSET_CAPTURE, $offset)
        ) {
            $match_character = $matches[0][0];
            $match_offset = $matches[0][1];

            if (in_array($match_character, array('\\', '"')) || !$is_quoted) { // Read: Quoted "[", "]" and "Ee+-" don't need to be separated from neighboring characters.
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
                    $last_tokenized_character = $match_offset - 1;
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

                        // This may be a 0-length quoted section. This is relevant for some decision-making later.
                        if (!$is_quoted
                            && $last_tokenized_character === ($match_offset - 1)
                            && substr($section_string, $last_tokenized_character, 1) === '"'
                        ) {
                            $tokens[] = (new NumberFormatToken(''))
                                ->setIsQuoted(true);
                        }
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
                case 'E+':
                case 'E-':
                case 'e+':
                case 'e-':
                    if (!$is_quoted && !$is_square_bracketed) {
                        $tokens[] = new NumberFormatToken($match_character);
                        $last_tokenized_character = $match_offset + 1; // Characters have been included in output.
                    } // else: Will be included in next "add token between last and this match" execution.
                    $offset = $match_offset + 2;
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
        $current_index = -1;
        $prevent_merge_of_passed_token = false;
        foreach ($tokens as $token) {
            // First loop iteration requires explicit handling.
            if (!isset($tokens_merged[$current_index]) || $prevent_merge_of_passed_token) {
                $current_index++;
                $tokens_merged[$current_index] = $token;
                $prevent_merge_of_passed_token = false;
                continue;
            }

            if (    $token->isScientificNotationE()
                &&  !$token->isQuoted()
                &&  !$token->isInSquareBrackets()
            ) {
                // This token should be kept seperated from the rest to make formatting easier.
                $prevent_merge_of_passed_token = true;
            }

            // Group successive tokens with the same rule-sets together.
            if (!$prevent_merge_of_passed_token
                && $token->isQuoted() === $tokens_merged[$current_index]->isQuoted()
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
     *                               ordered by priority of applicability.
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
     * e.g. input (as section): [#"some"00,.0\."garbage"0?] -> set values in section:
     *  number_type: 'decimal'
     *  decimal_format: #00,.00?
     *  format_left: '#00,'
     *  format_right: '0,0?'
     *  thousands_scale: 1
     *
     * @param  NumberFormatSection $section
     */
    private function prepareNumericFormat($section)
    {
        $format_left = ''; // For decimals: Characters before decimal. For fractions: Characters before slash.
        $format_right = ''; // For decimals: Characters after decimal. For fractions: Characters after slash.
        $exponent_format = ''; // Only for scientific format.
        $whole_values_format = ''; // Only for fractions.

        // Step 1: Determine correct number format type.
        foreach ($section->getTokens() as $token) {
            if ($token->isQuoted() || $token->isInSquareBrackets()) {
                continue;
            }
            if (strpos($token->getCode(), '/') !== false) { // Replicates Excel behavior. (SHOULD also check for surrounding 0#?)
                $section->setNumberType('fraction');
                break; // "fraction" format declaration is final and can not be overruled by other discoveries anymore.
            }
            if ($section->getNumberType() === null && preg_match('{[0#?.,/]+}', $token->getCode(), $matches)) {
                $section->setNumberType('decimal');
            }
        }

        // Step 2: Walk through all tokens and assign found format characters to their semantical sections.
        $decimal_character_passed = false;
        $e_token_passed = false;
        $fraction_char_detected = false;
        $end_of_fraction_detected = false;
        $whole_value_or_format_left_part = '';

        $tokens = $section->getTokens();
        foreach ($tokens as $token_index => $token) {
            if ($token->isQuoted()) {
                if ($section->getNumberType() === 'fraction') {
                    // Fraction format is complex, and quoted sections need to be detected for correct formatting.
                    if (!$fraction_char_detected) {
                        // e.g.: "0[ ]0/0", "0[ ]/0" (Note: The latter is not a whole-value format.)
                        $whole_values_format .= $whole_value_or_format_left_part;
                        $whole_value_or_format_left_part = '';
                    }
                    if ($format_right !== '') {
                        $end_of_fraction_detected = true;
                    }
                }
                continue;
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

            $tokens[$token_index]->setCode($code); // Note: No further manipulation of code contents from here on out.

            if (preg_match('{[Ee][+-]}', $code, $matches)) {
                // Scientific format detected. Number formats after the E+/e- position must be interpreted as exponent.
                $e_token_passed = true;
                continue; // Ee+- is always a token by itself. (See convertFormatSectionToTokens())
            }

            if ($section->getNumberType() === 'fraction') {
                // Very complex number format. Walk through string character by character, left-to-right.
                for ($i = 0; $i < strlen($code); $i++) {
                    $char = substr($code, $i, 1);
                    switch ($char) {
                        case '/':
                            $format_left = $whole_value_or_format_left_part;
                            $fraction_char_detected = true;
                            break;
                        case '0':
                        case '#':
                        case '?':
                            if ($end_of_fraction_detected) {
                                // e.g.: "0 0/0+[]", "0/0 []"
                                // Not considered part of the format anymore. Ignore here, show in formatted value later.
                            } else if ($fraction_char_detected) {
                                // e.g.: "0 0/[0]", "0/[0]"
                                $format_right .= $char;
                            } else {
                                // e.g.: "[0]/0", "[0] 0/0", "0 [0]/0", "[0] 0 0/0", "0 [0] 0/0", "0 0 [0]/0"
                                $whole_value_or_format_left_part .= $char;
                            }
                            break;
                        case '.':
                        case ',':
                            // Ignored characters. Will not show up in formatted value, does not trigger state changes.
                            break;
                        default:
                            // Non-format character. Indicates a format section change.
                            // Note: Quoted sections are handled further above.
                            if ($fraction_char_detected) {
                                if ($format_right !== '') {
                                    // e.g.: "0/0[ ]", "0 0/0[ ]"
                                    $end_of_fraction_detected = true;
                                } // else e.g.: "0/[ ]0", which does not trigger a section change.
                            } else if ($whole_value_or_format_left_part !== '') {
                                // e.g.: "0[ ]0/0", "0[ ]/0" (Note: The latter is not a whole-value format.)
                                $whole_values_format .= $whole_value_or_format_left_part;
                                $whole_value_or_format_left_part = '';
                            }
                            break;
                    }
                }
                continue; // Do not proceed with other format logic in case of fraction.
            }

            // Extract 0#?., symbols from format string and assign them to purposes in the decimal/scientific format.
            if ($e_token_passed) {
                // Note: "0.0" or "#,?" etc. are not valid. Ignoring [.,] here handles this gracefully.
                $exponent_format .= preg_replace('{[^0#?]}', '', $code);
                continue;
            }

            if ($decimal_character_passed) {
                $decimal_format_characters = preg_replace('{[^0#?,]}', '', $code);
                $format_right .= $decimal_format_characters;
            } else {
                $decimal_pos = strpos($code, '.');
                if ($decimal_pos !== false) {
                    // This token contains the decimal character. Split left/right parts off it.
                    $decimal_character_passed = true;
                    $format_left .= preg_replace('{[^0#?,]}', '', substr($code, 0, $decimal_pos));
                    $format_right .= preg_replace('{[^0#?,]}', '', substr($code, $decimal_pos + 1));
                } else {
                    // This token contains no decimal character. Move characters to left/right part depending on context.
                    $decimal_format_characters = preg_replace('{[^0#?,]}', '', $code);
                    $format_left .= $decimal_format_characters;
                }
            }
        }

        $section->setTokens($tokens); // Establish changes to $token object array copy in $section.
        if ($section->getNumberType() === 'decimal') {
            $section->setDecimalFormat($format_left . ($decimal_character_passed ? '.' : '') . $format_right);
        }
        $section->setFormatLeft($format_left);
        $section->setFormatRight($format_right);
        $section->setExponentFormat($exponent_format);
        $section->setWholeValuesFormat($whole_values_format);

        // Commas at end of either format left or right indicate scaling.
        $scaling = 0;
        if (preg_match('{(,+)$}', $format_left, $matches)) {
            $scaling += strlen($matches[1]);
        }
        if (preg_match('{(,+)$}', $format_right, $matches)) {
            $scaling += strlen($matches[1]);
        }
        $section->setThousandsScale($scaling);

        // Commas *within* format_left (not at its start/end) indicate a thousands separator.
        if (preg_match('{^[^,]+.*,.*[^,]+$}', $format_left)) { // Note: {^[^,]+,[^,]+$} would match 0,0 but not 0,,0
            $section->setUseThousandsSeparators(true);
        }
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
}
