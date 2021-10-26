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
        // Note: returnUnformatted for date/time values must be handled in applyDateTimeFormat(), due to additional constraints.
        $return_unformatted = $this->configuration->getReturnUnformatted() && !$section->getDateTimeType();
        $return_percentage_decimal = $section->isPercentage() && $this->configuration->getReturnPercentageDecimal();
        if ($return_unformatted || $return_percentage_decimal) {
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

        // Check returnUnformatted HERE, so that returnDateTimeObjects and force...Format can take precedence.
        if ($this->configuration->getReturnUnformatted()) {
            return $value;
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

        $sections = (new NumberFormatTokenizer())->prepareFormatSections($format_code);

        // Update cached data, for faster retrieval next time round.
        $this->parsed_format_cache[$format_index] = $sections;

        return $sections;
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
