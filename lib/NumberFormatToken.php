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
     * @return bool
     */
    public function isScientificNotationE()
    {
        return preg_match('{^[Ee][+-]$}', $this->code);
    }
}
