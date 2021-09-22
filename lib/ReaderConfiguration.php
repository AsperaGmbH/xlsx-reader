<?php

namespace Aspera\Spreadsheet\XLSX;

use InvalidArgumentException;

/** Configuration options to control reader behavior. */
class ReaderConfiguration
{
    /** @var string */
    private $temp_dir = '';

    /** @var int */
    private $skip_empty_cells = ReaderSkipConfiguration::SKIP_NONE;

    /** @var int */
    private $skip_empty_rows = ReaderSkipConfiguration::SKIP_NONE;

    /** @var bool */
    private $output_column_names = false;

    /** @var SharedStringsConfiguration */
    private $shared_strings_configuration;

    /** @var array */
    private $custom_formats = array();

    /** @var string|null */
    private $force_date_format;

    /** @var string|null */
    private $force_time_format;

    /** @var string|null */
    private $force_date_time_format;

    /** @var bool */
    private $return_unformatted = false;

    /** @var bool */
    private $return_percentage_decimal = false;

    /** @var bool */
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
     *
     * @throws InvalidArgumentException
     */
    public function setTempDir($temp_dir)
    {
        if (!is_string($temp_dir)) {
            throw new InvalidArgumentException('TempDir needs to be a string.');
        }
        $this->temp_dir = $temp_dir;

        return $this;
    }

    /**
     * Configuration of empty cell output.
     * Use ReaderSkipConfiguration constants to configure.
     *
     * @param  int $skip_empty_cells A ReaderSkipConfiguration constant.
     * @return self
     *
     * @throws InvalidArgumentException
     */
    public function setSkipEmptyCells($skip_empty_cells)
    {
        if (!is_numeric($skip_empty_cells)) {
            throw new InvalidArgumentException('SkipEmptyCells needs to be a ReaderSkipConfiguration constant.');
        }
        $this->skip_empty_cells = $skip_empty_cells;

        return $this;
    }

    /**
     * Configuration of empty row output.
     * Use ReaderSkipConfiguration constants to configure.
     *
     * @param  int $skip_empty_rows A ReaderSkipConfiguration constant.
     * @return self
     *
     * @throws InvalidArgumentException
     */
    public function setSkipEmptyRows($skip_empty_rows)
    {
        if (!is_numeric($skip_empty_rows)) {
            throw new InvalidArgumentException('SkipEmptyRows needs to be a ReaderSkipConfiguration constant.');
        }
        $this->skip_empty_rows = $skip_empty_rows;

        return $this;
    }

    /**
     * If true, output will use Excel-style column names (A-ZZ) instead of numbers as column keys.
     *
     * @param  bool $output_column_names
     * @return self
     *
     * @throws InvalidArgumentException
     */
    public function setOutputColumnNames($output_column_names)
    {
        if (!is_bool($output_column_names)) {
            throw new InvalidArgumentException('OutputColumnNames needs to be a boolean.');
        }
        $this->output_column_names = $output_column_names;

        return $this;
    }

    /**
     * Configuration options to control shared string reading and caching behaviour.
     *
     * @param  SharedStringsConfiguration $shared_strings_configuration
     * @return self
     *
     * @throws InvalidArgumentException
     */
    public function setSharedStringsConfiguration($shared_strings_configuration)
    {
        if (!($shared_strings_configuration instanceof SharedStringsConfiguration)) {
            throw new InvalidArgumentException(
                'SharedStringsConfiguration needs to be an instance of SharedStringsConfiguration.'
            );
        }
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
    public function setCustomFormats($custom_formats)
    {
        if (!is_array($custom_formats)) {
            throw new InvalidArgumentException('CustomFormats needs to be an array.');
        }
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
     * @param  string|null $force_date_format
     * @return self
     *
     * @throws InvalidArgumentException
     */
    public function setForceDateFormat($force_date_format)
    {
        if (!is_string($force_date_format) && $force_date_format !== null) {
            throw new InvalidArgumentException('ForceDateFormat needs to be a string (or null to unset).');
        }
        $this->force_date_format = $force_date_format;

        return $this;
    }

    /**
     * Format to use when outputting time values, regardless of originally set formatting.
     *
     * Note that a cell's type is defined by its format, not content.
     * If a cell contains time information, but its format contains no time information, the value is considered a date.
     *
     * @param  string|null $force_time_format
     * @return self
     *
     * @throws InvalidArgumentException
     */
    public function setForceTimeFormat($force_time_format)
    {
        if (!is_string($force_time_format) && $force_time_format !== null) {
            throw new InvalidArgumentException('ForceTimeFormat needs to be a string (or null to unset).');
        }
        $this->force_time_format = $force_time_format;

        return $this;
    }

    /**
     * Format to use when outputting datetime values, regardless of originally set formatting.
     *
     * Note that a cell's type is defined by its format, not content.
     * If a cell contains time information, but its format contains no time information, the value is considered a date.
     *
     * @param  string|null $force_date_time_format
     * @return self
     *
     * @throws InvalidArgumentException
     */
    public function setForceDateTimeFormat($force_date_time_format)
    {
        if (!is_string($force_date_time_format) && $force_date_time_format !== null) {
            throw new InvalidArgumentException('ForceDateTimeFormat needs to be a string (or null to unset).');
        }
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
     *
     * @throws InvalidArgumentException
     */
    public function setReturnUnformatted($return_unformatted)
    {
        if (!is_bool($return_unformatted)) {
            throw new InvalidArgumentException('ReturnUnformatted needs to be a bool.');
        }
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
     *
     * @throws InvalidArgumentException
     */
    public function setReturnPercentageDecimal($return_percentage_decimal)
    {
        if (!is_bool($return_percentage_decimal)) {
            throw new InvalidArgumentException('ReturnPercentageDecimal needs to be a bool.');
        }
        $this->return_percentage_decimal = $return_percentage_decimal;

        return $this;
    }

    /**
     * If true, return date/time values as PHP DateTime objects, not strings.
     * Takes precedence over the value of $return_unformatted.
     *
     * @param  bool $return_date_time_objects
     * @return self
     *
     * @throws InvalidArgumentException
     */
    public function setReturnDateTimeObjects($return_date_time_objects)
    {
        if (!is_bool($return_date_time_objects)) {
            throw new InvalidArgumentException('ReturnDateTimeObjects needs to be a bool.');
        }
        $this->return_date_time_objects = $return_date_time_objects;

        return $this;
    }

    /**
     * @return string
     */
    public function getTempDir()
    {
        return $this->temp_dir;
    }

    /**
     * @return int
     */
    public function getSkipEmptyCells()
    {
        return $this->skip_empty_cells;
    }

    /**
     * @return int
     */
    public function getSkipEmptyRows()
    {
        return $this->skip_empty_rows;
    }

    /**
     * @return bool
     */
    public function getOutputColumnNames()
    {
        return $this->output_column_names;
    }

    /**
     * @return SharedStringsConfiguration
     */
    public function getSharedStringsConfiguration()
    {
        return $this->shared_strings_configuration;
    }

    /**
     * @return array
     */
    public function getCustomFormats()
    {
        return $this->custom_formats;
    }

    /**
     * @return string|null
     */
    public function getForceDateFormat()
    {
        return $this->force_date_format;
    }

    /**
     * @return string|null
     */
    public function getForceTimeFormat()
    {
        return $this->force_time_format;
    }

    /**
     * @return string|null
     */
    public function getForceDateTimeFormat()
    {
        return $this->force_date_time_format;
    }

    /**
     * @return bool
     */
    public function getReturnUnformatted()
    {
        return $this->return_unformatted;
    }

    /**
     * @return bool
     */
    public function getReturnPercentageDecimal()
    {
        return $this->return_percentage_decimal;
    }

    /**
     * @return bool
     */
    public function getReturnDateTimeObjects()
    {
        return $this->return_date_time_objects;
    }
}
