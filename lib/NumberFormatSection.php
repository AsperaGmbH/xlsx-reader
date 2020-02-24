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

    /** @var string Specific date/time type value for this section. Possible values: date, time, datetime */
    private $dateTimeType;

    /** @var bool If true, this value is intended to convert its value to a percentage value. (multiply by 100) */
    private $is_percentage;

    /** @var bool If true, a minus sign should be automatically prepended to the formatted value. */
    private $prepend_minus_sign;

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
}
