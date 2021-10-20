<?php

namespace Aspera\Spreadsheet\XLSX;

class NumberFormatTokenizer
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

    /**
     * Split the given $format_string into sections, which are further split into tokens to assist in parsing.
     * Also determines each format section's purpose (read: which value ranges it should be applied to).
     *
     * @param  string $format_string The full format string, including all sections, completely unparsed
     * @return NumberFormatSection[]
     */
    public function prepareFormatSections($format_string)
    {
        $sections_tokenized = array();
        foreach ($this->splitSections($format_string) as $section_index => $section_string) {
            $sections_tokenized[$section_index] = new NumberFormatSection(
                $this->convertFormatSectionToTokens($section_string)
            );
        }
        $sections = $this->assignSectionPurposes($sections_tokenized);

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

        return $sections;
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

            if ($token->isScientificNotationE()
                && !$token->isQuoted()
                && !$token->isInSquareBrackets()
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
}
