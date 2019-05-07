<?php

namespace Aspera\Spreadsheet\XLSX;

use Iterator;
use Countable;
use RuntimeException;
use ZipArchive;
use DateTime;
use DateTimeZone;
use DateInterval;
use Exception;
use InvalidArgumentException;

/**
 * Class for parsing XLSX files.
 *
 * @author Aspera GmbH
 * @author Martins Pilsetnieks
 */
class Reader implements Iterator, Countable
{
    /** @var array Base formats for XLSX documents, to be made available without former declaration. */
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

    /** @var DateTime Standardized base date for the document's date/time values. */
    private static $base_date;

    /** @var bool Is the gmp_gcd method available for usage? Cached value. */
    private static $gmp_gcd_available = false;

    /** @var string Decimal separator character for output of locally formatted values. */
    private $decimal_separator;

    /** @var string Thousands separator character for output of locally formatted values. */
    private $thousand_separator;

    /** @var string Currency character for output of locally formatted values. */
    private $currency_code;

    /** @var SharedStringsConfiguration Configuration of shared strings handling. */
    private $shared_strings_configuration = null;

    /** @var bool Do not format date/time values and return DateTime objects instead. Default false. */
    private $return_date_time_objects;

    /** @var bool Output XLSX-style column names instead of numeric column identifiers. Default false. */
    private $output_column_names;

    /** @var bool Do not consider empty cell values in output. Default false. */
    private $skip_empty_cells;

    /** @var string Full path of the temporary directory that is going to be used to store unzipped files. */
    private $temp_dir;

    /** @var array Temporary files created while reading the document. */
    private $temp_files = array();

    /** @var RelationshipData File paths and -identifiers to all relevant parts of the read XLSX file */
    private $relationship_data;

    /** @var array Data about separate sheets in the file. */
    private $sheets = false;

    /** @var SharedStrings Shared strings handler. */
    private $shared_strings;

    /** @var array Container for cell value style data. */
    private $styles = array();

    /** @var array List of custom formats defined by the current XLSX file; array key = format index */
    private $formats = array();

    /** @var array List of custom formats defined by the user; array key = format index */
    private $customized_formats = array();

    /** @var string Format to use when outputting dates, regardless of originally set formatting.
     *              (Note: Will also be used if the original formatting omits time information, but the data value contains time information.) */
    private $enforced_date_format;

    /** @var string Format to use when outputting time information, regardless of originally set formatting. */
    private $enforced_time_format;

    /** @var string Format to use when outputting datetime values, regardless of originally set formatting. */
    private $enforced_datetime_format;

    /** @var array Cache for already processed format strings. */
    private $parsed_format_cache = array();

    /** @var string Path to the current worksheet XML file. */
    private $worksheet_path = false;

    /** @var OoxmlReader XML reader object for the current worksheet XML file. */
    private $worksheet_reader = false;

    /** @var bool Internal storage for the result of the valid() method related to the Iterator interface. */
    private $valid = false;

    /** @var bool Whether the reader is currently looking at an element within a <row> node. */
    private $row_open = false;

    /** @var int Current row number in the file. */
    private $row_number = 0;

    /** @var bool|array Contents of last read row. */
    private $current_row = false;

    /**
     * @param array $options Reader configuration; Permitted values:
     *      - TempDir (string)
     *          Path to directory to write temporary work files to
     *      - ReturnDateTimeObjects (bool)
     *          If true, date/time data will be returned as PHP DateTime objects.
     *          Otherwise, they will be returned as strings.
     *      - SkipEmptyCells (bool)
     *          If true, row content will not contain empty cells
     *      - SharedStringsConfiguration (SharedStringsConfiguration)
     *          Configuration options to control shared string reading and caching behaviour
     *
     * @throws Exception
     * @throws RuntimeException
     */
    public function __construct(array $options = null)
    {
        if (!isset($options['TempDir'])) {
            $options['TempDir'] = null;
        }
        $this->initTempDir($options['TempDir']);

        if (!empty($options['SharedStringsConfiguration'])) {
            $this->shared_strings_configuration = $options['SharedStringsConfiguration'];
        }
        if (!empty($options['CustomFormats'])) {
            $this->initCustomFormats($options['CustomFormats']);
        }
        if (!empty($options['ForceDateFormat'])) {
            $this->enforced_date_format = $options['ForceDateFormat'];
        }
        if (!empty($options['ForceTimeFormat'])) {
            $this->enforced_time_format = $options['ForceTimeFormat'];
        }
        if (!empty($options['ForceDateTimeFormat'])) {
            $this->enforced_datetime_format = $options['ForceDateTimeFormat'];
        }

        $this->skip_empty_cells = !empty($options['SkipEmptyCells']);
        $this->return_date_time_objects = !empty($options['ReturnDateTimeObjects']);
        $this->output_column_names = !empty($options['OutputColumnNames']);

        $this->initBaseDate();
        $this->initLocale();

        self::$gmp_gcd_available = function_exists('gmp_gcd');
    }

    /**
     * Open the given file and prepare everything for the reading of data.
     *
     * @param   string  $file_path
     *
     * @throws  Exception
     */
    public function open($file_path)
    {
        if (!is_readable($file_path)) {
            throw new RuntimeException('XLSXReader: File not readable (' . $file_path . ')');
        }

        if (!mkdir($this->temp_dir, 0777, true) || !file_exists($this->temp_dir)) {
            throw new RuntimeException(
                'XLSXReader: Could neither create nor confirm existance of temporary directory (' . $this->temp_dir . ')'
            );
        }

        $zip = new ZipArchive;
        $status = $zip->open($file_path);
        if ($status !== true) {
            throw new RuntimeException('XLSXReader: File not readable (' . $file_path . ') (Error ' . $status . ')');
        }

        $this->relationship_data = new RelationshipData($zip);
        $this->initWorkbookData($zip);
        $this->initWorksheets($zip);
        $this->initSharedStrings($zip, $this->shared_strings_configuration);
        $this->initStyles($zip);

        $zip->close();
    }

    /**
     * Free all connected resources.
     */
    public function close()
    {
        if ($this->worksheet_reader && $this->worksheet_reader instanceof OoxmlReader) {
            $this->worksheet_reader->close();
            $this->worksheet_reader = null;
        }

        if ($this->shared_strings && $this->shared_strings instanceof SharedStrings) {
            // Closing the shared string handler will also close all still opened shared string temporary work files.
            $this->shared_strings->close();
            $this->shared_strings = null;
        }

        $this->deleteTempfiles();

        $this->worksheet_path = null;
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
     * @param   int     $sheet_index
     *
     * @return  bool    True if sheet was successfully changed, false otherwise.
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

    // !Iterator interface methods

    /**
     * Rewind the Iterator to the first element.
     * Similar to the reset() function for arrays in PHP.
     */
    public function rewind()
    {
        if ($this->worksheet_reader instanceof OoxmlReader) {
            $this->worksheet_reader->close();
        } else {
            $this->worksheet_reader = new OoxmlReader();
            $this->worksheet_reader->setDefaultNamespaceIdentifierElements(OoxmlReader::NS_XLSX_MAIN);
            $this->worksheet_reader->setDefaultNamespaceIdentifierAttributes(OoxmlReader::NS_NONE);
        }

        $this->worksheet_reader->open($this->worksheet_path);

        $this->valid = true;
        $this->row_open = false;
        $this->current_row = false;
        $this->row_number = 0;
    }

    /**
     * Return the current element.
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

        return self::adjustRowOutput($this->current_row);
    }

    /**
     * Move forward to next element.
     *
     * @return array|false
     *
     * @throws Exception
     */
    public function next()
    {
        $this->row_number++;

        $this->current_row = array();

        // Walk through the document until the beginning of the first spreadsheet row.
        if (!$this->row_open) {
            while ($this->valid = $this->worksheet_reader->read()) {
                if (!$this->worksheet_reader->matchesElement('row')) {
                    continue;
                }

                $this->row_open = true;

                /* Getting the row spanning area (stored as e.g., 1:12)
                 * so that the last cells will be present, even if empty. */
                $row_spans = $this->worksheet_reader->getAttributeNsId('spans');

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

                // Do not read further than here if the current 'row' node is not the one to be read
                if ((int) $this->worksheet_reader->getAttributeNsId('r') !== $this->row_number) {
                    return self::adjustRowOutput($this->current_row);
                }
                break;
            }

            // No (further) rows found for reading.
            if (!$this->row_open) {
                return array();
            }
        }

        // Do not read further than here if the current 'row' node is not the one to be read
        if ((int) $this->worksheet_reader->getAttributeNsId('r') !== $this->row_number) {
            $row_spans = $this->worksheet_reader->getAttributeNsId('spans');
            if ($row_spans) {
                $row_spans = explode(':', $row_spans);
                $current_row_column_count = $row_spans[1];
            } else {
                $current_row_column_count = 0;
            }
            if ($current_row_column_count > 0 && !$this->skip_empty_cells) {
                $this->current_row = array_fill(0, $current_row_column_count, '');
            }
            return self::adjustRowOutput($this->current_row);
        }

        // Variables for empty cell handling.
        $max_index = 0;
        $cell_count = 0;
        $last_cell_index = -1;

        // Pre-loop-declarations. Will be filled by one loop iteration, then read in another.
        $style_id = null;
        $cell_index = null;
        $cell_has_shared_string = false;

        while ($this->valid = $this->worksheet_reader->read()) {
            if (!$this->worksheet_reader->matchesNamespace(OoxmlReader::NS_XLSX_MAIN)) {
                continue;
            }
            switch ($this->worksheet_reader->localName) {
                // </row> tag: Finish row reading.
                case 'row':
                    if ($this->worksheet_reader->isClosingTag()) {
                        $this->row_open = false;
                        break 2;
                    }
                    break;

                // <c> tag: Read cell metadata, such as styling of formatting information.
                case 'c':
                    if ($this->worksheet_reader->isClosingTag()) {
                        continue 2;
                    }

                    $cell_count++;

                    // Get the cell index via the "r" attribute and update max_index.
                    $cell_index = $this->worksheet_reader->getAttributeNsId('r');
                    if ($cell_index) {
                        $letter = preg_replace('{[^[:alpha:]]}S', '', $cell_index);
                        $cell_index = self::indexFromColumnLetter($letter);
                    } else {
                        // No "r" attribute available; Just position this cell to the right of the last one.
                        $cell_index = $last_cell_index + 1;
                    }
                    $last_cell_index = $cell_index;
                    if ($cell_index > $max_index) {
                        $max_index = $cell_index;
                    }

                    // Determine cell styling/formatting.
                    $cell_type = $this->worksheet_reader->getAttributeNsId('t');
                    $cell_has_shared_string = $cell_type === 's'; // s = shared string
                    $style_id = (int) $this->worksheet_reader->getAttributeNsId('s');

                    // If configured: Return empty strings for empty values.
                    if (!$this->skip_empty_cells) {
                        $this->current_row[$cell_index] = '';
                    }
                    break;

                // <v> or <is> tag: Read and store cell value according to current styling/formatting.
                case 'v':
                case 'is':
                    if ($this->worksheet_reader->isClosingTag()) {
                        continue 2;
                    }

                    $value = $this->worksheet_reader->readString();

                    if ($cell_has_shared_string) {
                        $value = $this->shared_strings->getSharedString($value);
                    }

                    // Skip empty values when specified as early as possible
                    if ($value === '' && $this->skip_empty_cells) {
                        break;
                    }

                    // Format value if necessary
                    if ($value !== '' && $style_id && isset($this->styles[$style_id])) {
                        $value = $this->formatValue($value, $style_id);
                    } elseif ($value) {
                        $value = $this->generalFormat($value);
                    }

                    $this->current_row[$cell_index] = $value;
                    break;

                default:
                    // nop
                    break;
            }
        }

        /* If configured: Return empty strings for empty values.
         * Only empty cells inbetween and on the left side are added. */
        if (($max_index + 1 > $cell_count) && !$this->skip_empty_cells) {
            $this->current_row += array_fill(0, $max_index + 1, '');
            ksort($this->current_row);
        }

        if (empty($this->current_row) && $this->skip_empty_cells) {
            $this->current_row[] = null;
        }

        return self::adjustRowOutput($this->current_row);
    }

    /**
     * Return the identifying key of the current element.
     *
     * @return mixed either an integer or a string
     */
    public function key()
    {
        return $this->row_number;
    }

    /**
     * Check if there is a current element after calls to rewind() or next().
     * Used to check if we've iterated to the end of the collection.
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
     * @param   string  $letter Letter(s) to convert
     * @return  mixed   Numeric index (0-based) or boolean false if it cannot be calculated
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
     * Converts the given column index to an XLSX-style [A-Z] column identifier string.
     *
     * @param   int $index
     * @return  string
     */
    public static function columnLetterFromIndex($index)
    {
        $dividend = $index + 1; // Internal counting starts at 0; For easy calculation, it needs to start at 1.
        $output_string = '';
        while ($dividend > 0) {
            $modulo = ($dividend - 1) % 26;
            $output_string = chr($modulo + 65) . $output_string;
            $dividend = floor(($dividend - $modulo) / 26);
        }
        return $output_string;
    }

    /**
     * Helper function for greatest common divisor calculation in case GMP extension is not enabled.
     *
     * @param   int $int_1
     * @param   int $int_2
     * @return  int Greatest common divisor
     */
    private static function GCD($int_1, $int_2)
    {
        $int_1 = (int) abs($int_1);
        $int_2 = (int) abs($int_2);

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
     * If configured, replaces numeric column identifiers in output array with XLSX-style [A-Z] column identifiers.
     * If not configured, returns the input array unchanged.
     *
     * @param   array $column_values
     * @return  array
     */
    private function adjustRowOutput($column_values)
    {
        if (!$this->output_column_names) {
            // Column names not desired in output; Nothing to do here.
            return $column_values;
        }

        $column_values_with_keys = array();
        foreach ($column_values as $k => $v) {
            $column_values_with_keys[self::columnLetterFromIndex($k)] = $v;
        }

        return $column_values_with_keys;
    }

    /**
     * Formats the value according to the index.
     *
     * @param   string  $value
     * @param   int     $format_index
     * @return  string
     *
     * @throws  Exception
     */
    private function formatValue($value, $format_index)
    {
        if (!is_numeric($value)) {
            // Only numeric values are formatted.
            return $value;
        }

        if (isset($this->styles[$format_index]) && ($this->styles[$format_index] !== false)) {
            $format_index = $this->styles[$format_index];
        } else {
            // Invalid format_index or the style was explicitly declared as "don't format anything".
            return $value;
        }

        if ($format_index === 0) {
            // Special case for the "General" format
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

            if (array_key_exists($format_index, $this->customized_formats)) {
                $format['Code'] = $this->customized_formats[$format_index];
            } elseif (array_key_exists($format_index, self::BUILTIN_FORMATS)) {
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
                    default:
                        // nop
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
                return (string) $value;
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
                $days = (int) $value;
                // Correcting for Feb 29, 1900
                if ($days > 60) {
                    $days--;
                }

                // At this point time is a fraction of a day
                $time = ($value - (int) $value);
                $seconds = 0;
                if ($time) {
                    // Here time is converted to seconds
                    // Workaround against precision loss: set low precision will round up milliseconds
                    $seconds = (int) round($time * 86400, 0);
                }

                $value = clone self::$base_date;
                $value->add(new DateInterval('P' . $days . 'D' . ($seconds ? 'T' . $seconds . 'S' : '')));

                if (!$this->return_date_time_objects) {
                    // Determine if the format is a date/time/datetime format and apply enforced formatting accordingly
                    $contains_date = preg_match('#[DdFjlmMnoStwWmYyz]#u', $format['Code']);
                    $contains_time = preg_match('#[aABgGhHisuv]#u', $format['Code']);
                    if ($contains_date) {
                        if ($contains_time) {
                            if ($this->enforced_datetime_format) {
                                $value = $value->format($this->enforced_datetime_format);
                            }
                        } else if ($this->enforced_date_format) {
                            $value = $value->format($this->enforced_date_format);
                        }
                    } else if ($this->enforced_time_format) {
                        $value = $value->format($this->enforced_time_format);
                    }

                    if ($value instanceof DateTime) {
                        // No format enforcement for this value type found. Format as declared.
                        $value = $value->format($format['Code']);
                    }
                } // else: A DateTime object is returned
            } elseif ($format['Type'] === 'Euro') {
                $value = 'EUR ' . sprintf('%1.2f', $value);
            } else {
                // Fractional numbers; We get "0.25" and have to turn that into "1/4".
                if ($format['Type'] === 'Fraction' && ($value != (int) $value)) {
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
     * Attempts to approximate Excel's "general" format.
     *
     * @param   mixed   $value
     * @return  mixed
     */
    private function generalFormat($value)
    {
        if (is_numeric($value)) {
            $value = (float) $value;
        }

        return $value;
    }

    /**
     * Check and set TempDir to use for file operations.
     * A new folder will be created within the given directory, which will contain all work files,
     * and which will be cleaned up after the process have finished.
     *
     * @param string|null $base_temp_dir
     */
    private function initTempDir($base_temp_dir) {
        if ($base_temp_dir === null) {
            $base_temp_dir = sys_get_temp_dir();
        }
        if (!is_writable($base_temp_dir)) {
            throw new RuntimeException('XLSXReader: Provided temporary directory (' . $base_temp_dir . ') is not writable');
        }
        $base_temp_dir = rtrim($base_temp_dir, DIRECTORY_SEPARATOR);
        /** @noinspection NonSecureUniqidUsageInspection */
        $this->temp_dir = $base_temp_dir . DIRECTORY_SEPARATOR . uniqid() . DIRECTORY_SEPARATOR;
    }

    /**
     * Set base date for calculation of retrieved date/time data.
     *
     * @throws Exception
     */
    private function initBaseDate() {
        self::$base_date = new DateTime();
        self::$base_date->setTimezone(new DateTimeZone('UTC'));
        self::$base_date->setDate(1900, 1, 0);
        self::$base_date->setTime(0, 0, 0);
    }

    /**
     * Pre-fill locale related data using current system locale.
     */
    private function initLocale() {
        $locale = localeconv();
        $this->decimal_separator = $locale['decimal_point'];
        $this->thousand_separator = $locale['thousands_sep'];
        $this->currency_code = $locale['int_curr_symbol'];
    }

    /**
     * Read general workbook information from the given zip into memory.
     *
     * @param   ZipArchive  $zip
     *
     * @throws  Exception
     */
    private function initWorkbookData(ZipArchive $zip)
    {
        $workbook = $this->relationship_data->getWorkbook();
        if (!$workbook) {
            throw new Exception('workbook data not found in XLSX file');
        }
        $workbook_xml = $zip->getFromName($workbook->getOriginalPath());

        $this->sheets = array();
        $workbook_reader = new OoxmlReader();
        $workbook_reader->setDefaultNamespaceIdentifierElements(OoxmlReader::NS_XLSX_MAIN);
        $workbook_reader->setDefaultNamespaceIdentifierAttributes(OoxmlReader::NS_NONE);
        $workbook_reader->xml($workbook_xml);
        while ($workbook_reader->read()) {
            if ($workbook_reader->matchesElement('sheet')) {
                // <sheet/> - Read in data about this worksheet.
                $sheet_name = (string) $workbook_reader->getAttributeNsId('name');
                $rel_id = (string) $workbook_reader->getAttributeNsId('id', OoxmlReader::NS_RELATIONSHIPS_DOCUMENTLEVEL);

                $new_sheet = new Worksheet();
                $new_sheet->setName($sheet_name);
                $new_sheet->setRelationshipId($rel_id);

                $sheet_index = str_ireplace('rId', '', $rel_id);
                $this->sheets[$sheet_index] = $new_sheet;
            } elseif ($workbook_reader->matchesElement('sheets') && $workbook_reader->isClosingTag()) {
                // </sheets> - Indicates that the rest of the document is of no further importance to us. Abort.
                break;
            }
        }
        $workbook_reader->close();
    }

    /**
     * Extract worksheet files to temp directory and set the first worksheet as active.
     *
     * @param ZipArchive $zip
     */
    private function initWorksheets(ZipArchive $zip)
    {
        // Sheet order determining value: relative sheet positioning within the document (rId)
        ksort($this->sheets);

        // Extract worksheets to temporary work directory
        foreach ($this->relationship_data->getWorksheets() as $worksheet) {
            /** @var RelationshipElement $worksheet */
            $worksheet_path_zip = $worksheet->getOriginalPath();
            $worksheet_path_conv = str_replace(RelationshipData::ZIP_DIR_SEP, DIRECTORY_SEPARATOR, $worksheet_path_zip);
            $worksheet_path_unzipped = $this->temp_dir . $worksheet_path_conv;
            if (!$zip->extractTo($this->temp_dir, $worksheet_path_zip)) {
                $message = 'XLSXReader: Could not extract file [' . $worksheet_path_zip . '] to directory [' . $this->temp_dir . '].';
                $this->reportZipExtractionFailure($zip, $message);
            }
            $worksheet->setAccessPath($worksheet_path_unzipped);
            $this->temp_files[] = $worksheet_path_unzipped;
        }

        // Set first sheet as current sheet
        if (!$this->changeSheet(0)) {
            throw new RuntimeException('XLSXReader: Sheet cannot be changed.');
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
            if (!$zip->extractTo($this->temp_dir, $inzip_path)) {
                $message = 'XLSXReader: Could not extract file [' . $inzip_path . '] to directory [' . $this->temp_dir . '].';
                $this->reportZipExtractionFailure($zip, $message);
            }

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

            $styles_xml = $zip->getFromName($first_styles_element->getOriginalPath());

            // Read cell style definitions and store them in internal variables
            $styles_reader = new OoxmlReader();
            $styles_reader->setDefaultNamespaceIdentifierElements(OoxmlReader::NS_XLSX_MAIN);
            $styles_reader->setDefaultNamespaceIdentifierAttributes(OoxmlReader::NS_NONE);
            $styles_reader->xml($styles_xml);
            $current_scope_is_cell_xfs = false;
            $current_scope_is_num_fmts = false;
            $switchList = array(
                'numFmts' => array('numFmts'),
                'numFmt'  => array('numFmt'),
                'cellXfs' => array('cellXfs'),
                'xf'      => array('xf')
            );
            while ($styles_reader->read()) {
                switch ($styles_reader->matchesOneOfList($switchList)) {
                    // <numFmts><numFmt/></numFmts> - check for number format definitions
                    case 'numFmts':
                        $current_scope_is_num_fmts = !$styles_reader->isClosingTag();
                        break;
                    case 'numFmt':
                        if (!$current_scope_is_num_fmts || $styles_reader->isClosingTag()) {
                            break;
                        }
                        $format_code = (string) $styles_reader->getAttributeNsId('formatCode');
                        $num_fmt_id = (int) $styles_reader->getAttributeNsId('numFmtId');
                        $this->formats[$num_fmt_id] = $format_code;
                        break;

                    // <cellXfs><xf/></cellXfs> - check for format usages
                    case 'cellXfs':
                        $current_scope_is_cell_xfs = !$styles_reader->isClosingTag();
                        break;
                    case 'xf':
                        if (!$current_scope_is_cell_xfs || $styles_reader->isClosingTag()) {
                            break;
                        }
                        if ($styles_reader->getAttributeNsId('applyNumberFormat')) {
                            // Number formatting has been enabled for this format.
                            // If format ID >= 164, it is a custom format and should be read from styleSheet\numFmts
                            $this->styles[] = (int) $styles_reader->getAttributeNsId('numFmtId');
                        } else if ($styles_reader->getAttributeNsId('quotePrefix')) {
                            // "quotePrefix" automatically preceeds the cell content with a ' symbol. This enforces a text format.
                            $this->styles[] = false; // false = "Do not format anything".
                        } else {
                            $this->styles[] = 0; // 0 = "General" format
                        }
                        break;
                }
            }
            $styles_reader->close();
        }
    }

    /**
     * @param   array $custom_formats
     * @return  void
     */
    private function initCustomFormats(array $custom_formats)
    {
        foreach ($custom_formats as $format_index => $format) {
            if (array_key_exists($format_index, self::BUILTIN_FORMATS) !== false) {
                $this->customized_formats[$format_index] = $format;
            }
        }
    }

    /**
     * Delete all registered temporary work files and -directories.
     */
    private function deleteTempfiles() {
        foreach ($this->temp_files as $temp_file) {
            @unlink($temp_file);
        }

        // Better safe than sorry - shouldn't try deleting '.' or '/', or '..'.
        if (strlen($this->temp_dir) > 2) {
            @rmdir($this->temp_dir . 'xl' . DIRECTORY_SEPARATOR . 'worksheets');
            @rmdir($this->temp_dir . 'xl');
            @rmdir($this->temp_dir);
        }
    }

    /**
     * Gather data on zip extractTo() fault and throw an appropriate Exception.
     *
     * @param   ZipArchive  $zip
     * @param   string      $message    Optional error message to prefix the error details with.
     *
     * @throws  RuntimeException
     */
    private function reportZipExtractionFailure($zip, $message = '')
    {
        $status_code = $zip->status;
        $status_message = $zip->getStatusString();
        if ($status_code || $status_message) {
            $message .= ' Status from ZipArchive:';
            if ($status_code) {
                $message .= ' Code [' . $status_code . '];';
            }
            if ($status_message) {
                $message .= ' Message [' . $status_message . '];';
            }
        }
        throw new RuntimeException($message);
    }
}
