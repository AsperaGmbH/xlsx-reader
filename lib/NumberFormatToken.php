<?php

namespace Aspera\Spreadsheet\XLSX;

/** Data of a single, syntactical token of a cell format. */
class NumberFormatToken
{
    /** @var string Format code of this token. */
    private $code;

    /** @var bool Is this token in quotes or escaped via a backslash, Y/N. If true, $code should be output as-is. */
    private $is_quoted = false;

    /** @var int|null Index of the current square bracket section, starting at 0. null if this token is not in square brackets. */
    private $square_bracket_index;

    /** @var string|null Type of this number format. Possible values: null, decimal, fraction */
    private $number_type;

    /** @var bool If true, whole values will be extracted from the output fraction. */
    private $do_extract_whole;

    /** @var string Format string portion to the left of the decimal point. */
    private $format_left;

    /** @var string Format string portion to the right of the decimal point. */
    private $format_right;

    /** @var int Amount of thousands to scale the output value down by. */
    private $thousands_scale;

    /** @var bool If true, will include thousands separators in the formatted output. */
    private $use_thousands_separators;

    /**
     * @param string $code
     */
    public function __construct($code)
    {
        $this->code = $code;
    }

    /**
     * @param  string $code
     * @return $this
     */
    public function setCode($code)
    {
        $this->code = $code;
        return $this;
    }

    /**
     * @param string $code
     */
    public function appendCode($code)
    {
        $this->code .= $code;
    }

    /**
     * @return string
     */
    public function getCode()
    {
        return $this->code;
    }

    /**
     * @param  bool $is_quoted
     * @return $this
     */
    public function setIsQuoted($is_quoted)
    {
        $this->is_quoted = $is_quoted;
        return $this;
    }

    /**
     * @return bool
     */
    public function isQuoted()
    {
        return $this->is_quoted;
    }

    /**
     * @param  int|null $square_bracket_index
     * @return $this
     */
    public function setSquareBracketIndex($square_bracket_index)
    {
        $this->square_bracket_index = $square_bracket_index;
        return $this;
    }

    /**
     * @return int|null
     */
    public function getSquareBracketIndex()
    {
        return $this->square_bracket_index;
    }

    /**
     * @return bool
     */
    public function isInSquareBrackets()
    {
        return $this->square_bracket_index !== null;
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
     * @param  bool $do_extract_whole
     * @return $this
     */
    public function setExtractWhole($do_extract_whole)
    {
        $this->do_extract_whole = $do_extract_whole;
        return $this;
    }

    /**
     * @return bool
     */
    public function doExtractWhole()
    {
        return $this->do_extract_whole;
    }

    /**
     * @param string $format_left
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
}
