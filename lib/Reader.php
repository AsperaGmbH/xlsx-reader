<?php

namespace Aspera\Spreadsheet\XLSX;

require_once __DIR__ . '/../vendor/autoload.php';

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
use InvalidArgumentException;

/**
 * Class for parsing XLSX files
 *
 * @author Aspera GmbH
 * @author Martins Pilsetnieks
 */
class Reader implements Iterator, Countable
{
    // XML Namespaces
    const XMLNS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const XMLNS_DOCUMENT_RELATIONSHIPS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const XMLNS_PACKAGE_RELATIONSHIPS = 'http://schemas.openxmlformats.org/package/2006/relationships';

    const CELL_TYPE_SHARED_STR = 's';

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
     * @var DateTime Standardized base date for the document's date/time values
     */
    private static $base_date;

    /**
     * Is the gmp_gcd method available for usage? Cached value.
     *
     * @var bool
     */
    private static $gmp_gcd_available = false;

    /**
     * @var string Decimal separator character for output of locally formatted values
     */
    private $decimal_separator;

    /**
     * @var string Thousands separator character for output of locally formatted values
     */
    private $thousand_separator;

    /**
     * @var string Currency character for output of locally formatted values
     */
    private $currency_code;

    /**
     * @var array Temporary files
     */
    private $temp_files = array();

    /**
     * RelationshipData instance containing all file paths and file identifiers to all parts of the read XLSX file
     * that are relevant to the reader's functionality.
     *
     * @var RelationshipData $relationship_data
     */
    private $relationship_data;

    /**
     * @var SimpleXMLElement XML object for the workbook XML file
     */
    private $workbook_xml = false;

    /**
     * @var array Data about separate sheets in the file
     */
    private $sheets = false;

    /**
     * @var string Path to the current worksheet XML file
     */
    private $worksheet_path = false;

    /**
     * @var XMLReader XML reader object for the current worksheet XML file
     */
    private $worksheet_reader = false;

    /**
     * @var SharedStrings
     */
    private $shared_strings;

    /**
     * @var SimpleXMLElement XML object for the styles XML file
     */
    private $styles_xml = false;

    /**
     * @var array Container for cell value style data
     */
    private $styles = array();

    /**
     * @var array List of custom formats defined by the current XLSX file; array key = format index
     */
    private $formats = array();

    /**
     * @var array Cache for already processed format strings
     */
    private $parsed_format_cache = array();

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

    /**
     * Internal storage for the result of the valid() method related to the Iterator interface.
     *
     * @var bool
     */
    private $valid = false;

    /**
     * @var bool Whether the reader is currently looking at an element within a <row> node
     */
    private $row_open = false;

    /**
     * @var int Current row number in the file
     */
    private $row_number = 0;

    /**
     * @var bool|array Contents of last read row
     */
    private $current_row = false;

    /**
     * @param string Path to file
     * @param array Options:
     *      - TempDir (string)
     *      Path to directory to write temporary work files to
     *      - ReturnDateTimeObjects (bool)
     *      If true, date/time data will be returned as PHP DateTime objects.
     *      Otherwise, they will be returned as strings.
     *      - SkipEmptyCells (bool)
     *      If true, row content will not contain empty cells
     *      - SharedStringsConfiguration (SharedStringsConfiguration)
     *      Configuration options to control shared string reading and caching behaviour
     *
     * @throws RuntimeException
     */
    public function __construct($filepath, array $options = null)
    {
        if (!is_readable($filepath)) {
            throw new RuntimeException('XLSXReader: File not readable (' . $filepath . ')');
        }

        // Set options
        if (isset($options['TempDir']) && !is_writable($options['TempDir'])) {
            throw new RuntimeException('XLSXReader: Provided temporary directory (' . $options['TempDir'] . ') is not writable');
        }
        $this->temp_dir = isset($options['TempDir']) ? $options['TempDir'] : sys_get_temp_dir();
        $this->temp_dir = rtrim($this->temp_dir, DIRECTORY_SEPARATOR);
        /** @noinspection NonSecureUniqidUsageInspection */
        $this->temp_dir = $this->temp_dir . DIRECTORY_SEPARATOR . uniqid() . DIRECTORY_SEPARATOR;
        $this->skip_empty_cells = isset($options['SkipEmptyCells']) && $options['SkipEmptyCells'];
        $this->return_date_time_objects = isset($options['ReturnDateTimeObjects']) && $options['ReturnDateTimeObjects'];

        $shared_strings_configuration = null;
        if (!empty($options['SharedStringsConfiguration'])) {
            $shared_strings_configuration = $options['SharedStringsConfiguration'];
        }

        // Set base date for calculation of date/time data
        self::$base_date = new DateTime;
        self::$base_date->setTimezone(new DateTimeZone('UTC'));
        self::$base_date->setDate(1900, 1, 0);
        self::$base_date->setTime(0, 0, 0);

        // Prefill locale related data using current system locale
        $locale = localeconv();
        $this->decimal_separator = $locale['decimal_point'];
        $this->thousand_separator = $locale['thousands_sep'];
        $this->currency_code = $locale['int_curr_symbol'];

        // Determine availability of standard functions
        self::$gmp_gcd_available = function_exists('gmp_gcd');

        // Open zip file
        $zip = new ZipArchive;
        $status = $zip->open($filepath);
        if ($status !== true) {
            throw new RuntimeException('XLSXReader: File not readable (' . $filepath . ') (Error ' . $status . ')');
        }

        // Gather up information on the document's file structure
        $this->relationship_data = new RelationshipData();
        $this->relationship_data->loadFromZip($zip);

        // Prepare workbook data
        $this->initWorkbook($zip);

        // Prepare shared strings
        $this->initSharedStrings($zip, $shared_strings_configuration);

        // Prepare worksheet data and set the first worksheet as the active worksheet
        $this->initWorksheets($zip);

        // Prepare styles data
        $this->initStyles($zip);

        // Finish initialization phase
        $zip->close();
    }

    /**
     * Destructor, destroys all that remains (closes and deletes temp files)
     */
    public function __destruct()
    {
        // First close the worksheet, then delete its temporary files.
        if ($this->worksheet_reader && $this->worksheet_reader instanceof XMLReader) {
            $this->worksheet_reader->close();
            unset($this->worksheet_reader);
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
     * Set the decimal separator to use for the output of locale-oriented formatted values
     *
     * @param string $new_character
     */
    public function setDecimalSeparator($new_character)
    {
        if (!is_string($new_character)) {
            throw new InvalidArgumentException('Given argument is not a string.');
        }
        $this->decimal_separator = $new_character;
    }

    /**
     * Set the thousands separator to use for the output of locale-oriented formatted values
     *
     * @param string $new_character
     */
    public function setThousandsSeparator($new_character)
    {
        if (!is_string($new_character)) {
            throw new InvalidArgumentException('Given argument is not a string.');
        }
        $this->thousand_separator = $new_character;
    }

    /**
     * Set the currency character code to use for the output of locale-oriented formatted values
     *
     * @param string $new_character
     */
    public function setCurrencyCode($new_character)
    {
        if (!is_string($new_character)) {
            throw new InvalidArgumentException('Given argument is not a string.');
        }
        $this->currency_code = $new_character;
    }

    /**
     * Retrieves an array with information about sheets in the current file
     *
     * @return array List of sheets (key is sheet index, value is of type Worksheet). Sheet's index starts with 0.
     */
    public function getSheets()
    {
        return array_values($this->sheets);
    }

    /**
     * Changes the current sheet in the file to the sheet with the given index.
     *
     * @param int Sheet index
     *
     * @return bool True if sheet was successfully changed, false otherwise.
     */
    public function changeSheet($sheet_index)
    {
        $sheets = $this->getSheets(); // Note: Realigns indexes to an auto increment.
        if (!isset($sheets[$sheet_index])) {
            return false;
        }
        /** @var Worksheet $target_sheet */
        $target_sheet = $sheets[$sheet_index];

        // The path to the target worksheet file can be obtained via the relationship id reference.
        $target_relationship_id = $target_sheet->getRelationshipId();
        /** @var RelationshipElement $relationship_worksheet */
        foreach ($this->relationship_data->getWorksheets() as $relationship_worksheet) {
            if ($relationship_worksheet->getId() === $target_relationship_id) {
                $worksheet_path = $relationship_worksheet->getAccessPath();
                break;
            }
        }
        if (!isset($worksheet_path) || !is_readable($worksheet_path)) {
            return false;
        }

        // Initialize the determined target sheet as the new current sheet
        $this->worksheet_path = $worksheet_path;
        $this->rewind();
        return true;
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
        if ($this->worksheet_reader instanceof XMLReader) {
            $this->worksheet_reader->close();
        } else {
            $this->worksheet_reader = new XMLReader;
        }

        $this->worksheet_reader->open($this->worksheet_path);

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
            while ($this->valid = $this->worksheet_reader->read()) {
                if ($this->worksheet_reader->namespaceURI === self::XMLNS_MAIN && $this->worksheet_reader->localName === 'row') {
                    // Getting the row spanning area (stored as e.g., 1:12)
                    // so that the last cells will be present, even if empty
                    $row_spans = $this->worksheet_reader->getAttribute('spans');

                    if ($row_spans) {
                        $row_spans = explode(':', $row_spans);
                        $current_row_column_count = $row_spans[1];
                    } else {
                        $current_row_column_count = 0;
                    }

                    // If configured: Return empty strings for empty values
                    if ($current_row_column_count > 0 && !$this->skip_empty_cells) {
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
            $last_cell_index = -1;

            $cell_has_shared_string = false;
            while ($this->valid = $this->worksheet_reader->read()) {
                if ($this->worksheet_reader->namespaceURI !== self::XMLNS_MAIN) {
                    continue;
                }
                switch ($this->worksheet_reader->localName) {
                    // End of row
                    case 'row':
                        if ($this->worksheet_reader->nodeType === XMLReader::END_ELEMENT) {
                            $this->row_open = false;
                            break 2;
                        }
                        break;
                    // Cell
                    case 'c':
                        if ($this->worksheet_reader->nodeType === XMLReader::END_ELEMENT) {
                            continue 2;
                        }

                        $style_id = (int)$this->worksheet_reader->getAttribute('s');

                        // Get the index of the cell
                        $cell_index = $this->worksheet_reader->getAttribute('r');
                        if ($cell_index) {
                            $letter = preg_replace('{[^[:alpha:]]}S', '', $cell_index);
                            $cell_index = self::indexFromColumnLetter($letter);
                        } else {
                            // No "r" attribute available; Just position this cell to the right of the last one.
                            $cell_index = $last_cell_index + 1;
                        }
                        $last_cell_index = $cell_index;

                        // Determine cell type
                        $cell_has_shared_string = $this->worksheet_reader->getAttribute('t') === self::CELL_TYPE_SHARED_STR;

                        // If configured: Return empty strings for empty values
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
                        if ($this->worksheet_reader->nodeType === XMLReader::END_ELEMENT) {
                            continue 2;
                        }

                        $value = $this->worksheet_reader->readString();

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

            // If configured: Return empty strings for empty values
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
                // This does not seem to be a letter. Someone must have given us an invalid value.
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
                'Code'      => false,
                'Type'      => false,
                'Scale'     => 1,
                'Thousands' => false,
                'Currency'  => false
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
                    // Format contains a currency code (Syntax: [$<Currency String>-<language info>])
                    $curr_code = explode('-', $matches[1]);
                    if (isset($curr_code[0])) {
                        $curr_code = $curr_code[0];
                    } else {
                        $curr_code = $this->currency_code;
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
                    if (self::$gmp_gcd_available) {
                        $gcd = gmp_strval(gmp_gcd($decimal, $decimal_divisor));
                    } else {
                        $gcd = self::GCD($decimal, $decimal_divisor);
                    }

                    // Determine fraction parts (1 and 4 => 1/4)
                    $adj_decimal = $decimal / $gcd;
                    $adj_decimal_divisor = $decimal_divisor / $gcd;

                    if (   strpos($format['Code'], '0') !== false
                        || strpos($format['Code'], '#') !== false
                        || strpos($format['Code'], '? ?') === 0
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
                                $this->decimal_separator, $this->thousand_separator);
                        } else {
                            $value = sprintf($format['Pattern'], $value);
                        }
                        $value = preg_replace('{(0+)(\.?)(0*)}', $value, $format['Code']);
                    }
                }

                // Currency/Accounting
                if ($format['Currency']) {
                    $value = preg_replace('{\[\$.*\]}u', $format['Currency'], $value);
                }
            }
        }

        return $value;
    }

    /**
     * Read general workbook information from the given zip into memory
     *
     * @param ZipArchive $zip
     */
    private function initWorkbook(ZipArchive $zip)
    {
        $workbook = $this->relationship_data->getWorkbook();
        if ($workbook) {
            $workbook_xml = $zip->getFromName($workbook->getOriginalPath());

            // Workaround for xpath bug (Default namespace cannot be addressed; Fix by "removing" its declaration:)
            $workbook_xml = str_replace('xmlns=', 'ns=', $workbook_xml);

            $this->workbook_xml = new SimpleXMLElement($workbook_xml);
        }
    }

    /**
     * Read shared strings data from the given zip into memory as configured via the given configuration object
     * and potentially create temporary work files for easy retrieval of shared string data.
     *
     * @param ZipArchive                 $zip
     * @param SharedStringsConfiguration $shared_strings_configuration Optional, default null
     */
    private function initSharedStrings(
        ZipArchive $zip,
        SharedStringsConfiguration $shared_strings_configuration = null
    ) {
        $shared_strings = $this->relationship_data->getSharedStrings();
        if (count($shared_strings) > 0) {
            /* Currently, documents with multiple shared strings files are not supported.
            *  Only the first shared string file will be used. */
            /** @var RelationshipElement $first_shared_string_element */
            $first_shared_string_element = $shared_strings[0];

            // Determine target directory and path for the extracted file
            $inzip_path = $first_shared_string_element->getOriginalPath();
            $inzip_path_for_outzip = str_replace(RelationshipData::ZIP_DIR_SEP, DIRECTORY_SEPARATOR, $inzip_path);
            $dir_of_extracted_file = $this->temp_dir . dirname($inzip_path_for_outzip) . DIRECTORY_SEPARATOR;
            $filename_of_extracted_file = basename($inzip_path_for_outzip);
            $path_to_extracted_file = $dir_of_extracted_file . $filename_of_extracted_file;

            // Extract file and note it in relevant variables
            $zip->extractTo($this->temp_dir, $inzip_path);
            $first_shared_string_element->setAccessPath($path_to_extracted_file);
            $this->temp_files[] = $path_to_extracted_file;

            // Initialize SharedStrings
            $this->shared_strings = new SharedStrings(
                $dir_of_extracted_file,
                $filename_of_extracted_file,
                $shared_strings_configuration
            );

            // Extend temp_files with files created by SharedStrings
            $this->temp_files = array_merge($this->temp_files, $this->shared_strings->getTempFiles());
        }
    }

    /**
     * Initializes internal sheet information storage and sets the first sheet as the current sheet.
     *
     * @param  ZipArchive $zip
     * @return void
     */
    private function initWorksheets(ZipArchive $zip)
    {
        $this->sheets = array();

        // Determine namespaces of interest (main namespace, relationships namespace) of the workbook
        $namespaces = $this->workbook_xml->getDocNamespaces();
        $main_ns = '';
        $rel_ns = '';
        foreach ($namespaces as $namespace_prefix => $namespace_uri) {
            switch ($namespace_uri) {
                case self::XMLNS_MAIN:
                    $main_ns = $namespace_prefix;
                    break;
                case self::XMLNS_DOCUMENT_RELATIONSHIPS:
                    $rel_ns = $namespace_prefix;
                    break;
            }
        }
        $main_ns_pre = $main_ns . ($main_ns != '' ? ':' : '');

        // Iterate through all sheet elements in the workbook
        $xpath_query = '/' . $main_ns_pre . 'workbook/' . $main_ns_pre . 'sheets/' . $main_ns_pre . 'sheet';
        foreach ($this->workbook_xml->xpath($xpath_query) as $sheet) {
            $sheet_id = (string)$sheet['sheetId'];
            $new_sheet = new Worksheet();
            $new_sheet->setId($sheet_id);
            $new_sheet->setName((string)$sheet['name']);
            $new_sheet->setRelationshipId((string)$sheet->attributes($rel_ns, true)['id']);
            $this->sheets[$sheet_id] = $new_sheet;
        }

        // Sheet order determining value: Sheet ID attribute
        ksort($this->sheets);

        // Extract worksheets to temporary work directory
        foreach ($this->relationship_data->getWorksheets() as $worksheet) {
            /** @var RelationshipElement $worksheet */
            $worksheet_path_zip = $worksheet->getOriginalPath();
            $worksheet_path_conv = str_replace(RelationshipData::ZIP_DIR_SEP, DIRECTORY_SEPARATOR, $worksheet_path_zip);
            $worksheet_path_unzipped = $this->temp_dir . $worksheet_path_conv;
            $zip->extractTo($this->temp_dir, $worksheet_path_zip);
            $worksheet->setAccessPath($worksheet_path_unzipped);
            $this->temp_files[] = $worksheet_path_unzipped;
        }

        // Set first sheet as current sheet
        if (!$this->changeSheet(0)) {
            throw new RuntimeException('XLSXReader: Sheet cannot be changed.');
        }
    }

    /**
     * Reads and prepares information on styles declared by the document for later usage.
     *
     * @param ZipArchive $zip
     */
    private function initStyles(ZipArchive $zip)
    {
        $styles = $this->relationship_data->getStyles();
        if (count($styles) > 0) {
            /* Currently, documents with multiple styles files are not supported.
            *  Only the first styles file will be used. */
            /** @var RelationshipElement $first_styles_element */
            $first_styles_element = $styles[0];

            $styles_xml_raw = $zip->getFromName($first_styles_element->getOriginalPath());

            // Workaround for xpath bug (Default namespace cannot be addressed; Fix by "removing" its declaration:)
            $styles_xml_raw = str_replace('xmlns=', 'ns=', $styles_xml_raw);

            $this->styles_xml = new SimpleXMLElement($styles_xml_raw);

            // Determine namespaces of interest (main namespace, relationships namespace) of the workbook
            $namespaces = $this->styles_xml->getDocNamespaces();
            $main_ns = '';
            foreach ($namespaces as $namespace_prefix => $namespace_uri) {
                switch ($namespace_uri) {
                    case self::XMLNS_MAIN:
                        $main_ns = $namespace_prefix;
                        break;
                }
            }
            $main_ns_pre = $main_ns . ($main_ns != '' ? ':' : '');

            // Read cell style definitions and store them in internal variables
            foreach ($this->styles_xml->xpath($main_ns_pre . 'cellXfs/' . $main_ns_pre . 'xf') as $xf) {
                // Check if the found number format should actually be applied; If not, use "General" format (ID: 0)
                if ($xf->attributes()->applyNumberFormat) {
                    // If format ID >= 164, it is a custom format and should be read from styleSheet\numFmts
                    $this->styles[] = (int)$xf->attributes()->numFmtId;
                } else {
                    $this->styles[] = 0;
                }
            }

            // Read number format definitions and store them in internal variables
            foreach ($this->styles_xml->xpath($main_ns_pre . 'numFmts/' . $main_ns_pre . 'numFmt') as $num_ft) {
                $num_ft_attributes = $num_ft->attributes();
                $this->formats[(int)$num_ft_attributes->numFmtId] = (string)$num_ft_attributes->formatCode;
            }

            // Clean up
            unset($this->styles_xml);
        }
    }
}
