<?php

namespace Aspera\Spreadsheet\XLSX;

/** Data of a single, syntactical token of a cell format. */
class NumberFormatToken
{
    /** @var string Format code of this token. */
    private $code;

    /** @var bool Is this token in quotes or escaped via a backslash, Y/N. If true, $code should be output as-is. */
    private $is_quoted = false;

    /** @var ?int Index of the current square bracket section, starting at 0. null if this token is not in square brackets. */
    private $square_bracket_index;

    public function __construct(string $code)
    {
        $this->code = $code;
    }

    public function setCode(string $code): self
    {
        $this->code = $code;
        return $this;
    }

    public function appendCode(string $code): void
    {
        $this->code .= $code;
    }

    public function getCode(): string
    {
        return $this->code;
    }

    public function setIsQuoted(bool $is_quoted): self
    {
        $this->is_quoted = $is_quoted;
        return $this;
    }

    public function isQuoted(): bool
    {
        return $this->is_quoted;
    }

    public function setSquareBracketIndex(?int $square_bracket_index): self
    {
        $this->square_bracket_index = $square_bracket_index;
        return $this;
    }

    public function getSquareBracketIndex(): ?int
    {
        return $this->square_bracket_index;
    }

    public function isInSquareBrackets(): bool
    {
        return $this->square_bracket_index !== null;
    }

    public function isScientificNotationE(): bool
    {
        return preg_match('{^[Ee][+-]$}', $this->code);
    }
}
