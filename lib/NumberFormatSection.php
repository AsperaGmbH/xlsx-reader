<?php

namespace Aspera\Spreadsheet\XLSX;

/**
 * Data of a single section of a number format, to be applied to a particular value type.
 * (e.g: If $purpose is "<0", this format is only applied to negative values.)
 */
class NumberFormatSection {
    /** @var NumberFormatToken[] */
    private $tokens;

    /** @var string Purpose of this section. Can be a condition (e.g: >=-20) or a default_ definition. (e.g.: default_number) */
    private $purpose;

    /** @var string|null Type of this number format. Possible values: null, decimal, fraction */
    private $number_type;

    /** @var string Specific date/time type value for this section. Possible values: date, time, datetime */
    private $dateTimeType;

    /** @var bool If true, this value is intended to convert its value to a percentage value. (multiply by 100) */
    private $is_percentage;

    /** @var bool If true, a minus sign should be automatically prepended to the formatted value. */
    private $prepend_minus_sign;

    /** @var int Amount of thousands to scale the output value down by. */
    private $thousands_scale;

    /** @var bool If true, will include thousands separators in the formatted output. */
    private $use_thousands_separators;

    /** @var string Contains only characters related to the decimal format. */
    private $decimal_format = '';

    /** @var string Part of decimal_format that's to the left of the decimal symbol. */
    private $format_left = '';

    /** @var string Part of decimal_format that's to the right of the decimal symbol. */
    private $format_right = '';

    /** @var string When $is_scientific_format = true, contains the format string for the exponent. 0.00E+## > "##" */
    private $exponent_format = '';

    /** @var string Format of the whole-values portion of the format, if the format is a fraction. */
    private $whole_values_format = '';

    /**
     * @param NumberFormatToken[] $tokens
     * @param string|null         $purpose
     */
    public function __construct($tokens, $purpose = null)
    {
        $this->tokens = $tokens;
        $this->purpose = $purpose;
    }

    /**
     * @param  NumberFormatToken[] $tokens
     * @return $this
     */
    public function setTokens($tokens)
    {
        $this->tokens = $tokens;
        return $this;
    }

    /**
     * @return NumberFormatToken[]
     */
    public function getTokens()
    {
        return $this->tokens;
    }

    /**
     * @param  string $purpose
     * @return $this
     */
    public function setPurpose($purpose)
    {
        $this->purpose = $purpose;
        return $this;
    }

    /**
     * @return string
     */
    public function getPurpose()
    {
        return $this->purpose;
    }

    /**
     * @param  string|null $number_type
     * @return $this
     */
    public function setNumberType($number_type)
    {
        $this->number_type = $number_type;
        return $this;
    }

    /**
     * @return string|null
     */
    public function getNumberType()
    {
        return $this->number_type;
    }

    /**
     * @param  string $dateTimeType
     * @return $this
     */
    public function setDateTimeType($dateTimeType)
    {
        $this->dateTimeType = $dateTimeType;
        return $this;
    }

    /**
     * @return string
     */
    public function getDateTimeType()
    {
        return $this->dateTimeType;
    }

    /**
     * @param  bool $is_percentage
     * @return $this
     */
    public function setIsPercentage($is_percentage)
    {
        $this->is_percentage = $is_percentage;
        return $this;
    }

    /**
     * @return bool
     */
    public function isPercentage()
    {
        return $this->is_percentage;
    }

    /**
     * @param  bool $prepend_minus_sign
     * @return $this
     */
    public function setPrependMinusSign($prepend_minus_sign)
    {
        $this->prepend_minus_sign = $prepend_minus_sign;
        return $this;
    }

    /**
     * @return bool
     */
    public function prependMinusSign()
    {
        return $this->prepend_minus_sign;
    }

    /**
     * @param  int $thousands_scale
     * @return $this
     */
    public function setThousandsScale($thousands_scale)
    {
        $this->thousands_scale = $thousands_scale;
        return $this;
    }

    /**
     * @return int
     */
    public function getThousandsScale()
    {
        return $this->thousands_scale;
    }

    /**
     * @param  bool $use_thousands_separators
     * @return $this
     */
    public function setUseThousandsSeparators($use_thousands_separators)
    {
        $this->use_thousands_separators = $use_thousands_separators;
        return $this;
    }

    /**
     * @return bool
     */
    public function useThousandsSeparators()
    {
        return $this->use_thousands_separators;
    }

    /**
     * @param  string $decimal_format
     * @return $this
     */
    public function setDecimalFormat($decimal_format)
    {
        $this->decimal_format = $decimal_format;
        return $this;
    }

    /**
     * @return string
     */
    public function getDecimalFormat()
    {
        return $this->decimal_format;
    }

    /**
     * @param  string $format_left
     * @return $this
     */
    public function setFormatLeft($format_left)
    {
        $this->format_left = $format_left;
        return $this;
    }

    /**
     * @return string
     */
    public function getFormatLeft()
    {
        return $this->format_left;
    }

    /**
     * @param  string $format_right
     * @return $this
     */
    public function setFormatRight($format_right)
    {
        $this->format_right = $format_right;
        return $this;
    }

    /**
     * @return string
     */
    public function getFormatRight()
    {
        return $this->format_right;
    }

    /**
     * @param  string $exponent_format
     * @return $this
     */
    public function setExponentFormat($exponent_format)
    {
        $this->exponent_format = $exponent_format;
        return $this;
    }

    /**
     * @return string
     */
    public function getExponentFormat()
    {
        return $this->exponent_format;
    }

    /**
     * @param  string $whole_values_format
     * @return $this
     */
    public function setWholeValuesFormat($whole_values_format)
    {
        $this->whole_values_format = $whole_values_format;
        return $this;
    }

    /**
     * @return string
     */
    public function getWholeValuesFormat()
    {
        return $this->whole_values_format;
    }
}
