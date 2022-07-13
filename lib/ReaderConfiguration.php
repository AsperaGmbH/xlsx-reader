<?php

namespace Aspera\Spreadsheet\XLSX;

use InvalidArgumentException;

/** Configuration options to control reader behavior. */
class ReaderConfiguration
{
    /**
     * Full path to directory to write temporary work files to. Default: sys_get_temp_dir()
     *
     * @var string
     */
    private $temp_dir;

    /**
     * Configuration of empty cell output. Use ReaderSkipConfiguration constants to configure.
     *
     * @var int
     */
    private $skip_empty_cells = ReaderSkipConfiguration::SKIP_NONE;

    /**
     * Configuration of empty row output. Use ReaderSkipConfiguration constants to configure.
     *
     * @var int
     */
    private $skip_empty_rows = ReaderSkipConfiguration::SKIP_NONE;

    /**
     * If true, output will use Excel-style column names (A-ZZ) instead of numbers as column keys.
     *
     * @var bool
     */
    private $output_column_names = false;

    /**
     * Configuration options to control shared string reading and caching behaviour.
     *
     * @var SharedStringsConfiguration
     */
    private $shared_strings_configuration;

    /**
     * A list of user-defined formats, overriding those given in the XLSX file itself.
     * Given as key_value pairs of format: [format_index (int)] => format_code (string)
     *
     * @var array
     */
    private $custom_formats = array();

    /**
     * Format to use when outputting dates, regardless of originally set formatting.
     * See setForceDateFormat() for more information.
     *
     * @var ?string
     */
    private $force_date_format;

    /**
     * Format to use when outputting time values, regardless of originally set formatting.
     * See setForceTimeFormat() for more information.
     *
     * @var ?string
     */
    private $force_time_format;

    /**
     * Format to use when outputting datetime values, regardless of originally set formatting.
     * See setForceDateTimeFormat() for more information.
     *
     * @var ?string
     */
    private $force_date_time_format;

    /**
     * If true, do not format anything. Returns numbers as-is. (e.g. 42967 25% => 25)
     * See setReturnUnformatted() for more information.
     *
     * @var bool
     */
    private $return_unformatted = false;

    /**
     * If true, percentage values will be returned as decimal point values. (e.g. 0-100% => 0-1, 25% => 0.25)
     * Takes precedence over the value of $return_unformatted.
     * See setReturnPercentageDecimal() for more information.
     *
     * @var bool
     */
    private $return_percentage_decimal = false;

    /**
     * If true, return date/time values as PHP DateTime objects, not strings.
     * Takes precedence over the value of $return_unformatted.
     *
     * @var bool
     */
    private $return_date_time_objects = false;

    public function __construct()
    {
        $this->temp_dir = sys_get_temp_dir();
        $this->shared_strings_configuration = new SharedStringsConfiguration();
    }

    /**
     * Full path to directory to write temporary work files to. Default: sys_get_temp_dir()
     *
     * @param  string $temp_dir
     * @return self
     */
    public function setTempDir(string $temp_dir): self
    {
        $this->temp_dir = $temp_dir;

        return $this;
    }

    /**
     * Configuration of empty cell output.
     * Use ReaderSkipConfiguration constants to configure.
     *
     * @param  int $skip_empty_cells A ReaderSkipConfiguration constant.
     * @return self
     */
    public function setSkipEmptyCells(int $skip_empty_cells): self
    {
        $this->skip_empty_cells = $skip_empty_cells;

        return $this;
    }

    /**
     * Configuration of empty row output.
     * Use ReaderSkipConfiguration constants to configure.
     *
     * @param  int $skip_empty_rows A ReaderSkipConfiguration constant.
     * @return self
     */
    public function setSkipEmptyRows(int $skip_empty_rows): self
    {
        $this->skip_empty_rows = $skip_empty_rows;

        return $this;
    }

    /**
     * If true, output will use Excel-style column names (A-ZZ) instead of numbers as column keys.
     *
     * @param  bool $output_column_names
     * @return self
     */
    public function setOutputColumnNames(bool $output_column_names): self
    {
        $this->output_column_names = $output_column_names;

        return $this;
    }

    /**
     * Configuration options to control shared string reading and caching behaviour.
     *
     * @param  SharedStringsConfiguration $shared_strings_configuration
     * @return self
     */
    public function setSharedStringsConfiguration(SharedStringsConfiguration $shared_strings_configuration): self
    {
        $this->shared_strings_configuration = $shared_strings_configuration;

        return $this;
    }

    /**
     * A list of user-defined formats, overriding those given in the XLSX file itself.
     * Given as key_value pairs of format: [format_index (int)] => format_code (string)
     *
     * @param  array $custom_formats
     * @return self
     *
     * @throws InvalidArgumentException
     */
    public function setCustomFormats(array $custom_formats): self
    {
        foreach ($custom_formats as $key => $value) {
            if (!is_numeric($key) || !is_string($value)) {
                throw new InvalidArgumentException(
                    'CustomFormats elements need to be of the structure [format_index] => "format_string".'
                );
            }
        }
        $this->custom_formats = $custom_formats;

        return $this;
    }

    /**
     * Format to use when outputting dates, regardless of originally set formatting.
     *
     * Note that a cell's type is defined by its format, not content.
     * If a cell contains time information, but its format contains no time information, the value is considered a date.
     *
     * @param  ?string $force_date_format
     * @return self
     */
    public function setForceDateFormat(?string $force_date_format): self
    {
        $this->force_date_format = $force_date_format;

        return $this;
    }

    /**
     * Format to use when outputting time values, regardless of originally set formatting.
     *
     * Note that a cell's type is defined by its format, not content.
     * If a cell contains time information, but its format contains no time information, the value is considered a date.
     *
     * @param  ?string $force_time_format
     * @return self
     */
    public function setForceTimeFormat(?string $force_time_format): self
    {
        $this->force_time_format = $force_time_format;

        return $this;
    }

    /**
     * Format to use when outputting datetime values, regardless of originally set formatting.
     *
     * Note that a cell's type is defined by its format, not content.
     * If a cell contains time information, but its format contains no time information, the value is considered a date.
     *
     * @param  ?string $force_date_time_format
     * @return self
     */
    public function setForceDateTimeFormat(?string $force_date_time_format): self
    {
        $this->force_date_time_format = $force_date_time_format;

        return $this;
    }

    /**
     * Do not format anything. Returns numbers as-is. (e.g. 42967 25% => 25)
     *
     * Note 1: Does not affect returned Date/Time instances or percentage value multiplication.
     *
     * Note 2: Be aware that rounding errors introduced by popular spreadsheet editors may cause the
     * internally stored values to differ a lot from what would be shown as a result of formatting.
     * Be further advised that values may sometimes be stored using E-notation.
     *
     * @param  bool $return_unformatted
     * @return self
     */
    public function setReturnUnformatted(bool $return_unformatted): self
    {
        $this->return_unformatted = $return_unformatted;

        return $this;
    }

    /**
     * If true, percentage values will be returned as decimal point values. (e.g. 0-100% => 0-1, 25% => 0.25)
     * Takes precedence over the value of $return_unformatted.
     *
     * Note: Be aware that rounding errors introduced by popular spreadsheet editors may cause the
     * internally stored values to differ a lot from what would be shown as a result of formatting.
     * Be further advised that values may sometimes be stored using E-notation.
     *
     * @param  bool $return_percentage_decimal
     * @return self
     */
    public function setReturnPercentageDecimal(bool $return_percentage_decimal): self
    {
        $this->return_percentage_decimal = $return_percentage_decimal;

        return $this;
    }

    /**
     * If true, return date/time values as PHP DateTime objects, not strings.
     * Takes precedence over the value of $return_unformatted.
     *
     * @param  bool $return_date_time_objects
     * @return self
     */
    public function setReturnDateTimeObjects(bool $return_date_time_objects): self
    {
        $this->return_date_time_objects = $return_date_time_objects;

        return $this;
    }

    /**
     * @return string
     */
    public function getTempDir(): string
    {
        return $this->temp_dir;
    }

    /**
     * @return int
     */
    public function getSkipEmptyCells(): int
    {
        return $this->skip_empty_cells;
    }

    /**
     * @return int
     */
    public function getSkipEmptyRows(): int
    {
        return $this->skip_empty_rows;
    }

    /**
     * @return bool
     */
    public function getOutputColumnNames(): bool
    {
        return $this->output_column_names;
    }

    /**
     * @return SharedStringsConfiguration
     */
    public function getSharedStringsConfiguration(): SharedStringsConfiguration
    {
        return $this->shared_strings_configuration;
    }

    /**
     * @return array
     */
    public function getCustomFormats(): array
    {
        return $this->custom_formats;
    }

    /**
     * @return ?string
     */
    public function getForceDateFormat(): ?string
    {
        return $this->force_date_format;
    }

    /**
     * @return ?string
     */
    public function getForceTimeFormat(): ?string
    {
        return $this->force_time_format;
    }

    /**
     * @return ?string
     */
    public function getForceDateTimeFormat(): ?string
    {
        return $this->force_date_time_format;
    }

    /**
     * @return bool
     */
    public function getReturnUnformatted(): bool
    {
        return $this->return_unformatted;
    }

    /**
     * @return bool
     */
    public function getReturnPercentageDecimal(): bool
    {
        return $this->return_percentage_decimal;
    }

    /**
     * @return bool
     */
    public function getReturnDateTimeObjects(): bool
    {
        return $this->return_date_time_objects;
    }
}
