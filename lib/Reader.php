<?php

namespace Aspera\Spreadsheet\XLSX;

use Iterator;
use Countable;
use RuntimeException;
use ZipArchive;
use SimpleXMLElement;
use XMLReader;
use DateTime;
use DateTimeZone;
use DateInterval;
use Exception;

require_once('SharedStrings.php');

/**
 * Class for parsing XLSX files specifically
 *
 * @author Martins Pilsetnieks
 */
class Reader implements Iterator, Countable
{
    const CELL_TYPE_BOOL = 'b';
    const CELL_TYPE_NUMBER = 'n';
    const CELL_TYPE_ERROR = 'e';
    const CELL_TYPE_SHARED_STR = 's';
    const CELL_TYPE_STR = 'str';
    const CELL_TYPE_INLINE_STR = 'inlineStr';

    const BUILTIN_FORMATS = array(
        0 => '',
        1 => '0',
        2 => '0.00',
        3 => '#,##0',
        4 => '#,##0.00',

        9 => '0%',
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

    const DATE_REPLACEMENTS = array(
        'All' => array(
            '\\' => '',
            'am/pm' => 'A',
            'yyyy' => 'Y',
            'yy' => 'y',
            'mmmmm' => 'M',
            'mmmm' => 'F',
            'mmm' => 'M',
            ':mm' => ':i',
            'mm' => 'm',
            'm' => 'n',
            'dddd' => 'l',
            'ddd' => 'D',
            'dd' => 'd',
            'd' => 'j',
            'ss' => 's',
            '.s' => ''
        ),
        '24H' => array(
            'hh' => 'H',
            'h' => 'G'
        ),
        '12H' => array(
            'hh' => 'h',
            'h' => 'G'
        )
    );

    private static $runtime_info = array(
        'GMPSupported' => false
    );

    private $valid = false;

    // Worksheet file
    /**
     * @var string Path to the worksheet XML file
     */
    private $worksheet_path = false;

    /**
     * @var XMLReader XML reader object for the worksheet XML file
     */
    private $worksheet = false;

    // Workbook data
    /**
     * @var SimpleXMLElement XML object for the workbook XML file
     */
    private $workbook_xml = false;

    /**
     * @var SharedStrings
     */
    private $shared_strings;

    // Style data
    /**
     * @var SimpleXMLElement XML object for the styles XML file
     */
    private $styles_xml = false;

    /**
     * @var array Container for cell value style data
     */
    private $styles = array();

    // Constructor options
    /**
     * @var array Temporary file names
     */
    private $temp_files = array();

    /**
     * @var string Full path of the temporary directory that is going to be used to store unzipped files
     */
    private $temp_dir;

    /**
     * @var bool By default do not format date/time values
     */
    private $return_date_time_objects;

    /**
     * @var bool By default all empty cells(values) are considered
     */
    private $skip_empty_cells;

    // Runtime parsing data
    /**
     * @var int Current row number in the file
     */
    private $row_number = 0;

    private $current_row = false;

    /**
     * @var array Data about separate sheets in the file
     */
    private $sheets = false;

    private $row_open = false;

    private $formats = array();

    private static $base_date = false;
    private static $decimal_separator = '.';
    private static $thousand_separator = '';
    private static $currency_code = '';

    /**
     * @var array Cache for already processed format strings
     */
    private $parsed_format_cache = array();

    /**
     * @param string Path to file
     * @param array Options:
     *    TempDir (string)
     *      Path to directory to write temporary work files to
     *    ReturnDateTimeObjects (bool)
     *      If true, date/time data will be returned as PHP DateTime objects.
     *      Otherwise, they will be returned as strings.
     *    SkipEmptyCells (bool)
     *      If true, row content will not contain empty cells
     *    SharedStringsConfiguration (SharedStringsConfiguration)
     *      Configuration options to control shared string reading and caching behaviour
     *
     * @throws RuntimeException
     */
    public function __construct($filepath, array $options = null)
    {
        if (!is_readable($filepath)) {
            throw new RuntimeException('XLSXReader: File not readable (' . $filepath . ')');
        }

        if (isset($options['TempDir']) && !is_writable($options['TempDir'])) {
            throw new RuntimeException('XLSXReader: Provided temporary directory (' . $options['TempDir'] . ') is not writable');
        }
        $this->temp_dir = isset($options['TempDir']) ? $options['TempDir'] : sys_get_temp_dir();

        // set options
        $this->temp_dir = rtrim($this->temp_dir, DIRECTORY_SEPARATOR);
        $this->temp_dir = $this->temp_dir . DIRECTORY_SEPARATOR . uniqid() . DIRECTORY_SEPARATOR;
        $this->skip_empty_cells = isset($options['SkipEmptyCells']) && $options['SkipEmptyCells'];
        $this->return_date_time_objects = isset($options['ReturnDateTimeObjects']) && $options['ReturnDateTimeObjects'];

        $zip = new ZipArchive;
        $status = $zip->open($filepath);

        if ($status !== true) {
            throw new RuntimeException(
                'XLSXReader: File not readable (' . $filepath . ') (Error ' . $status . ')'
            );
        }

        // Getting the general workbook information
        if ($zip->locateName('xl/workbook.xml') !== false) {
            $this->workbook_xml = new SimpleXMLElement($zip->getFromName('xl/workbook.xml'));
        }

        // Extracting the XMLs from the XLSX zip file
        if ($zip->locateName('xl/sharedStrings.xml') !== false) {
            $zip->extractTo($this->temp_dir, 'xl/sharedStrings.xml');
            $this->temp_files[] = $this->temp_dir . 'xl' . DIRECTORY_SEPARATOR . 'sharedStrings.xml';
            $shared_strings_configuration = null;
            if (!empty($options['SharedStringsConfiguration'])) {
                $shared_strings_configuration = $options['SharedStringsConfiguration'];
            }
            $this->shared_strings = new SharedStrings(
                $this->temp_dir . 'xl' . DIRECTORY_SEPARATOR,
                'sharedStrings.xml',
                $shared_strings_configuration
            );
            $this->temp_files = array_merge($this->temp_files, $this->shared_strings->getTempFiles());
        }

        $this->initSheets();

        foreach ($this->sheets as $sheet_num => $name) {
            if ($zip->locateName('xl/worksheets/sheet' . $sheet_num . '.xml') !== false) {
                $zip->extractTo($this->temp_dir, 'xl/worksheets/sheet' . $sheet_num . '.xml');
                $this->temp_files[] = $this->temp_dir . 'xl' . DIRECTORY_SEPARATOR . 'worksheets' . DIRECTORY_SEPARATOR . 'sheet' . $sheet_num . '.xml';
            }
        }

        if (!$this->changeSheet(0)) {
            throw new RuntimeException('XLSXReader: Sheet cannot be changed.');
        }

        // If worksheet is present and is OK, parse the styles already
        if ($zip->locateName('xl/styles.xml') !== false) {
            $this->styles_xml = new SimpleXMLElement($zip->getFromName('xl/styles.xml'));
            if ($this->styles_xml && $this->styles_xml->cellXfs && $this->styles_xml->cellXfs->xf) {
                foreach ($this->styles_xml->cellXfs->xf as $xf) {
                    // Format #0 is a special case - it is the "General" format that is applied regardless of applyNumberFormat
                    if ($xf->attributes()->applyNumberFormat || (0 === (int)$xf->attributes()->numFmtId)) {
                        $format_id = (int)$xf->attributes()->numFmtId;
                        // If format ID >= 164, it is a custom format and should be read from styleSheet\numFmts
                        $this->styles[] = $format_id;
                    } else {
                        // 0 for "General" format
                        $this->styles[] = 0;
                    }
                }
            }

            if ($this->styles_xml->numFmts && $this->styles_xml->numFmts->numFmt) {
                foreach ($this->styles_xml->numFmts->numFmt as $num_ft) {
                    $num_ft_attributes = $num_ft->attributes();
                    $this->formats[(int)$num_ft_attributes->numFmtId] = (string)$num_ft_attributes->formatCode;
                }
            }

            unset($this->styles_xml);
        }

        $zip->close();

        // Setting base date
        if (!self::$base_date) {
            self::$base_date = new DateTime;
            self::$base_date->setTimezone(new DateTimeZone('UTC'));
            self::$base_date->setDate(1900, 1, 0);
            self::$base_date->setTime(0, 0, 0);
        }

        // Decimal and thousand separators
        if (!self::$decimal_separator && !self::$thousand_separator && !self::$currency_code) {
            $locale = localeconv();
            self::$decimal_separator = $locale['decimal_point'];
            self::$thousand_separator = $locale['thousands_sep'];
            self::$currency_code = $locale['int_curr_symbol'];
        }

        if (function_exists('gmp_gcd')) {
            self::$runtime_info['GMPSupported'] = true;
        }
    }

    /**
     * Destructor, destroys all that remains (closes and deletes temp files)
     */
    public function __destruct()
    {
        // First close the worksheet, then delete its temporary files.
        if ($this->worksheet && $this->worksheet instanceof XMLReader) {
            $this->worksheet->close();
            unset($this->worksheet);
        }

        // Close shared string handler; Will also close all still opened shared string temporary work files.
        $this->shared_strings->close();

        // Delete all registered temporary work files.
        foreach ($this->temp_files as $temp_file) {
            @unlink($temp_file);
        }

        // Better safe than sorry - shouldn't try deleting '.' or '/', or '..'.
        if (strlen($this->temp_dir) > 2) {
            @rmdir($this->temp_dir . 'xl' . DIRECTORY_SEPARATOR . 'worksheets');
            @rmdir($this->temp_dir . 'xl');
            @rmdir($this->temp_dir);
        }

        unset($this->worksheet_path);

        if (isset($this->styles_xml)) {
            unset($this->styles_xml);
        }

        if ($this->workbook_xml) {
            unset($this->workbook_xml);
        }
    }

    /**
     * Retrieves an array with information about sheets in the current file
     *
     * @return array List of sheets (key is sheet index, value is name). Sheet's index starts with 1.
     */
    public function getSheets()
    {
        return array_values($this->sheets);
    }

    /**
     * Changes the current sheet in the file to another
     *
     * @param int Sheet index
     *
     * @return bool True if sheet was successfully changed, false otherwise.
     */
    public function changeSheet($sheet_index)
    {
        $real_sheet_index = false;
        $sheets = $this->getSheets();

        if (isset($sheets[$sheet_index])) {
            $sheet_indexes = array_keys($this->sheets);
            $real_sheet_index = $sheet_indexes[$sheet_index];
        }

        $temp_worksheet_path = $this->temp_dir . 'xl/worksheets/sheet' . $real_sheet_index . '.xml';

        if ($real_sheet_index !== false && is_readable($temp_worksheet_path)) {
            $this->worksheet_path = $temp_worksheet_path;

            $this->rewind();

            return true;
        }

        return false;
    }

    /**
     * Attempts to approximate Excel's "general" format.
     *
     * @param mixed Value
     *
     * @return mixed Result
     */
    public function generalFormat($value)
    {
        // Numeric format
        if (is_numeric($value)) {
            $value = (float)$value;
        }

        return $value;
    }

    // !Iterator interface methods

    /**
     * Rewind the Iterator to the first element.
     * Similar to the reset() function for arrays in PHP
     */
    public function rewind()
    {
        // If the worksheet was already iterated, XML file is reopened.
        // Otherwise it should be at the beginning anyway
        if ($this->worksheet instanceof XMLReader) {
            $this->worksheet->close();
        } else {
            $this->worksheet = new XMLReader;
        }

        $this->worksheet->open($this->worksheet_path);

        $this->valid = true;
        $this->row_open = false;
        $this->current_row = false;
        $this->row_number = 0;
    }

    /**
     * Return the current element.
     * Similar to the current() function for arrays in PHP
     *
     * @return mixed current element from the collection
     *
     * @throws Exception
     */
    public function current()
    {
        if ($this->row_number === 0 && $this->current_row === false) {
            $this->next();
        }

        return $this->current_row;
    }

    /**
     * Move forward to next element.
     * Similar to the next() function for arrays in PHP
     *
     * @return array|false
     *
     * @throws Exception
     */
    public function next()
    {
        $this->row_number++;

        $this->current_row = array();

        if (!$this->row_open) {
            while ($this->valid = $this->worksheet->read()) {
                if ($this->worksheet->name === 'row') {
                    // Getting the row spanning area (stored as e.g., 1:12)
                    // so that the last cells will be present, even if empty
                    $row_spans = $this->worksheet->getAttribute('spans');

                    if ($row_spans) {
                        $row_spans = explode(':', $row_spans);
                        $current_row_column_count = $row_spans[1];
                    } else {
                        $current_row_column_count = 0;
                    }

                    if ($current_row_column_count > 0 && !$this->skip_empty_cells) {
                        // fill values with empty strings in case of need
                        $this->current_row = array_fill(0, $current_row_column_count, '');
                    }

                    $this->row_open = true;
                    break;
                }
            }
        }

        // Reading the necessary row, if found
        if ($this->row_open) {
            // These two are needed to control for empty cells
            $max_index = 0;
            $cell_count = 0;

            $cell_has_shared_string = false;

            while ($this->valid = $this->worksheet->read()) {
                switch ($this->worksheet->name) {
                    // End of row
                    case 'row':
                        if ($this->worksheet->nodeType === XMLReader::END_ELEMENT) {
                            $this->row_open = false;
                            break 2;
                        }
                        break;
                    // Cell
                    case 'c':
                        // If it is a closing tag, skip it
                        if ($this->worksheet->nodeType === XMLReader::END_ELEMENT) {
                            continue 2;
                        }

                        $style_id = (int)$this->worksheet->getAttribute('s');

                        // Get the index of the cell
                        $cell_index = $this->worksheet->getAttribute('r');
                        $letter = preg_replace('{[^[:alpha:]]}S', '', $cell_index);
                        $cell_index = self::indexFromColumnLetter($letter);

                        // Determine cell type
                        if ($this->worksheet->getAttribute('t') === self::CELL_TYPE_SHARED_STR) {
                            $cell_has_shared_string = true;
                        } else {
                            $cell_has_shared_string = false;
                        }

                        if (!$this->skip_empty_cells) {
                            $this->current_row[$cell_index] = '';
                        }

                        $cell_count++;

                        if ($cell_index > $max_index) {
                            $max_index = $cell_index;
                        }

                        break;
                    // Cell value
                    case 'v':
                    case 'is':
                        if ($this->worksheet->nodeType === XMLReader::END_ELEMENT) {
                            continue 2;
                        }

                        $value = $this->worksheet->readString();

                        if ($cell_has_shared_string) {
                            $value = $this->shared_strings->getSharedString($value);
                        }

                        // Format value if necessary
                        if ($value !== '' && $style_id && isset($this->styles[$style_id])) {
                            $value = $this->formatValue($value, $style_id);
                        } elseif ($value) {
                            $value = $this->generalFormat($value);
                        }

                        $this->current_row[$cell_index] = $value;
                        break;
                }
            }

            // Adding empty cells, if necessary
            // Only empty cells inbetween and on the left side are added
            if (($max_index + 1 > $cell_count) && !$this->skip_empty_cells) {
                $this->current_row += array_fill(0, $max_index + 1, '');
                ksort($this->current_row);
            }

            if (empty($this->current_row) && $this->skip_empty_cells) {
                $this->current_row[] = null;
            }
        }

        return $this->current_row;
    }

    /**
     * Return the identifying key of the current element.
     * Similar to the key() function for arrays in PHP
     *
     * @return mixed either an integer or a string
     */
    public function key()
    {
        return $this->row_number;
    }

    /**
     * Check if there is a current element after calls to rewind() or next().
     * Used to check if we've iterated to the end of the collection
     *
     * @return boolean FALSE if there's nothing more to iterate over
     */
    public function valid()
    {
        return $this->valid;
    }

    // !Countable interface method

    /**
     * Ostensibly should return the count of the contained items but this just returns the number
     * of rows read so far. It's not really correct but at least coherent.
     */
    public function count()
    {
        return $this->row_number;
    }

    /**
     * Takes the column letter and converts it to a numerical index (0-based)
     *
     * @param string Letter(s) to convert
     *
     * @return mixed Numeric index (0-based) or boolean false if it cannot be calculated
     */
    public static function indexFromColumnLetter($letter)
    {
        $letter = strtoupper($letter);

        $result = 0;
        for ($i = strlen($letter) - 1, $j = 0; $i >= 0; $i--, $j++) {
            $ord = ord($letter[$i]) - 64;
            if ($ord > 26) {
                // Something is very, very wrong
                return false;
            }
            $result += $ord * (26 ** $j);
        }

        return $result - 1;
    }

    /**
     * Helper function for greatest common divisor calculation in case GMP extension is not enabled
     *
     * @param int $int_1 Number #1
     * @param int $int_2 Number #2
     *
     * @return int Greatest common divisor
     */
    public static function GCD($int_1, $int_2)
    {
        $int_1 = (int)abs($int_1);
        $int_2 = (int)abs($int_2);

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
     * Formats the value according to the index
     *
     * @param string Cell value
     * @param int Format index
     *
     * @throws Exception
     *
     * @return string Formatted cell value
     */
    private function formatValue($value, $format_index)
    {
        if (!is_numeric($value)) {
            return $value;
        }

        if (isset($this->styles[$format_index]) && ($this->styles[$format_index] !== false)) {
            $format_index = $this->styles[$format_index];
        } else {
            return $value;
        }

        // A special case for the "General" format
        if ($format_index === 0) {
            return $this->generalFormat($value);
        }

        $format = array();

        if (isset($this->parsed_format_cache[$format_index])) {
            $format = $this->parsed_format_cache[$format_index];
        }

        if (!$format) {
            $format = array(
                'Code' => false,
                'Type' => false,
                'Scale' => 1,
                'Thousands' => false,
                'Currency' => false
            );

            if (array_key_exists($format_index, self::BUILTIN_FORMATS)) {
                $format['Code'] = self::BUILTIN_FORMATS[$format_index];
            } elseif (isset($this->formats[$format_index])) {
                $format['Code'] = $this->formats[$format_index];
            }

            // Format code found, now parsing the format
            if ($format['Code']) {
                $sections = explode(';', $format['Code']);
                $format['Code'] = $sections[0];

                switch (count($sections)) {
                    case 2:
                        if ($value < 0) {
                            $format['Code'] = $sections[1];
                        }
                        break;
                    case 3:
                    case 4:
                        if ($value < 0) {
                            $format['Code'] = $sections[1];
                        } elseif ($value === 0) {
                            $format['Code'] = $sections[2];
                        }
                        break;
                }
            }

            // Stripping colors
            $format['Code'] = trim(preg_replace('{^\[[[:alpha:]]+\]}i', '', $format['Code']));

            // Percentages
            if (substr($format['Code'], -1) === '%') {
                $format['Type'] = 'Percentage';
            } elseif (preg_match('{^(\[\$[[:alpha:]]*-[0-9A-F]*\])*[hmsdy]}i', $format['Code'])) {
                $format['Type'] = 'DateTime';

                $format['Code'] = trim(preg_replace('{^(\[\$[[:alpha:]]*-[0-9A-F]*\])}i', '', $format['Code']));
                $format['Code'] = strtolower($format['Code']);

                $format['Code'] = strtr($format['Code'], self::DATE_REPLACEMENTS['All']);
                if (strpos($format['Code'], 'A') === false) {
                    $format['Code'] = strtr($format['Code'], self::DATE_REPLACEMENTS['24H']);
                } else {
                    $format['Code'] = strtr($format['Code'], self::DATE_REPLACEMENTS['12H']);
                }
            } elseif ($format['Code'] === '[$eUR ]#,##0.00_-') {
                $format['Type'] = 'Euro';
            } else {
                // Removing skipped characters
                $format['Code'] = preg_replace('{_.}', '', $format['Code']);
                // Removing unnecessary escaping
                $format['Code'] = preg_replace("{\\\\}", '', $format['Code']);
                // Removing string quotes
                $format['Code'] = str_replace(array('"', '*'), '', $format['Code']);
                // Removing thousands separator
                if (strpos($format['Code'], '0,0') !== false || strpos($format['Code'], '#,#') !== false) {
                    $format['Thousands'] = true;
                }
                $format['Code'] = str_replace(array('0,0', '#,#'), array('00', '##'), $format['Code']);

                // Scaling (Commas indicate the power)
                $scale = 1;
                $matches = array();
                if (preg_match('{(0|#)(,+)}', $format['Code'], $matches)) {
                    $scale = 1000 ** strlen($matches[2]);
                    // Removing the commas
                    $format['Code'] = preg_replace(array('{0,+}', '{#,+}'), array('0', '#'), $format['Code']);
                }

                $format['Scale'] = $scale;

                if (preg_match('{#?.*\?\/\?}', $format['Code'])) {
                    $format['Type'] = 'Fraction';
                } else {
                    $format['Code'] = str_replace('#', '', $format['Code']);

                    $matches = array();
                    if (preg_match('{(0+)(\.?)(0*)}', preg_replace('{\[[^\]]+\]}', '', $format['Code']), $matches)) {
                        $integer = $matches[1];
                        $decimal_point = $matches[2];
                        $decimals = $matches[3];

                        $format['MinWidth'] = strlen($integer) + strlen($decimal_point) + strlen($decimals);
                        $format['Decimals'] = $decimals;
                        $format['Precision'] = strlen($format['Decimals']);
                        $format['Pattern'] = '%0' . $format['MinWidth'] . '.' . $format['Precision'] . 'f';
                    }
                }

                $matches = array();
                if (preg_match('{\[\$(.*)\]}u', $format['Code'], $matches)) {
                    $curr_code = $matches[1];
                    $curr_code = explode('-', $curr_code);
                    if ($curr_code) {
                        $curr_code = $curr_code[0];
                    }

                    if (!$curr_code) {
                        $curr_code = self::$currency_code;
                    }

                    $format['Currency'] = $curr_code;
                }

                $format['Code'] = trim($format['Code']);
            }

            $this->parsed_format_cache[$format_index] = $format;
        }

        // Applying format to value
        if ($format) {
            if ($format['Code'] === '@') {
                return (string)$value;
            }

            if ($format['Type'] === 'Percentage') {
                // Percentages
                if ($format['Code'] === '0%') {
                    $value = round(100 * $value, 0) . '%';
                } else {
                    $value = sprintf('%.2f%%', round(100 * $value, 2));
                }
            } elseif ($format['Type'] === 'DateTime') {
                // Dates and times
                $days = (int)$value;
                // Correcting for Feb 29, 1900
                if ($days > 60) {
                    $days--;
                }

                // At this point time is a fraction of a day
                $time = ($value - (int)$value);
                $seconds = 0;
                if ($time) {
                    // Here time is converted to seconds
                    // Some loss of precision will occur
                    $seconds = (int)($time * 86400);
                }

                $value = clone self::$base_date;
                $value->add(new DateInterval('P' . $days . 'D' . ($seconds ? 'T' . $seconds . 'S' : '')));

                if (!$this->return_date_time_objects) {
                    $value = $value->format($format['Code']);
                } // else: A DateTime object is returned
            } elseif ($format['Type'] === 'Euro') {
                $value = 'EUR ' . sprintf('%1.2f', $value);
            } else {
                // Fractional numbers; We get "0.25" and have to turn that into "1/4".
                if ($format['Type'] === 'Fraction' && ($value != (int)$value)) {
                    // Split fraction from integer value (2.25 => 2 and 0.25)
                    $integer = floor(abs($value));
                    $decimal = fmod(abs($value), 1);

                    // Turn fraction into non-decimal value (0.25 => 25)
                    $decimal *= 10 ** (strlen($decimal) - 2);

                    // Obtain biggest divisor for the fraction part (25 => 100 => 25/100)
                    $decimal_divisor = 10 ** strlen($decimal);

                    // Determine greatest common divisor for fraction optimization (so that 25/100 => 1/4)
                    if (self::$runtime_info['GMPSupported']) {
                        $gcd = gmp_strval(gmp_gcd($decimal, $decimal_divisor));
                    } else {
                        $gcd = self::GCD($decimal, $decimal_divisor);
                    }

                    // Determine fraction parts (1 and 4 => 1/4)
                    $adj_decimal = $decimal / $gcd;
                    $adj_decimal_divisor = $decimal_divisor / $gcd;

                    if (
                        strpos($format['Code'], '0') !== false ||
                        strpos($format['Code'], '#') !== false ||
                        0 === strpos($format['Code'], '? ?')
                    ) {
                        // Extract whole values from fraction (2.25 => "2 1/4")
                        $value = ($value < 0 ? '-' : '') .
                            ($integer ? $integer . ' ' : '') .
                            $adj_decimal . '/' .
                            $adj_decimal_divisor;
                    } else {
                        // Show entire value as fraction (2.25 => "9/4")
                        $adj_decimal += $integer * $adj_decimal_divisor;
                        $value = ($value < 0 ? '-' : '') .
                            $adj_decimal . '/' .
                            $adj_decimal_divisor;
                    }
                } else {
                    // Scaling
                    $value /= $format['Scale'];

                    if (!empty($format['MinWidth']) && $format['Decimals']) {
                        if ($format['Thousands']) {
                            $value = number_format($value, $format['Precision'],
                                self::$decimal_separator, self::$thousand_separator);
                        } else {
                            $value = sprintf($format['Pattern'], $value);
                        }

                        $value = preg_replace('{(0+)(\.?)(0*)}', $value, $format['Code']);
                    }
                }

                // Currency/Accounting
                if ($format['Currency']) {
                    $value = preg_replace('', $format['Currency'], $value);
                }
            }

        }

        return $value;
    }

    /**
     * Sets 'sheets', an array with information about sheets in the current file
     *
     * @return void
     */
    private function initSheets()
    {
        $this->sheets = array();

        foreach ($this->workbook_xml->sheets->sheet as $sheet) {
            $sheet_id = (string)$sheet['sheetId'];
            $this->sheets[$sheet_id] = (string)$sheet['name'];
        }

        ksort($this->sheets);
    }
}
