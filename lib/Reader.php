<?php

namespace Aspera\Spreadsheet\XLSX;

use Iterator;
use Countable;
use RuntimeException;
use ZipArchive;
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
    /** @var NumberFormat */
    private $number_format;

    /** @var SharedStringsConfiguration Configuration of shared strings handling. */
    private $shared_strings_configuration = null;

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

        $this->number_format = new NumberFormat($options);

        if (!empty($options['SharedStringsConfiguration'])) {
            if (!($options['SharedStringsConfiguration'] instanceof SharedStringsConfiguration)) {
                throw new InvalidArgumentException('SharedStringsConfiguration has an invalid type.');
            }
            $this->shared_strings_configuration = $options['SharedStringsConfiguration'];
        }

        $this->skip_empty_cells = !empty($options['SkipEmptyCells']);
        $this->output_column_names = !empty($options['OutputColumnNames']);
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
        $this->number_format->setDecimalSeparator($new_character);
    }

    /**
     * Set the thousands separator to use for the output of locale-oriented formatted values
     *
     * @param string $new_character
     */
    public function setThousandsSeparator($new_character)
    {
        $this->number_format->setThousandsSeparator($new_character);
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

                /* Getting the row spanning area (stored as e.g. "1:12", or more rarely, "1:3 6:8 11:12")
                 * so that the last cells will be present, even if empty. */
                $row_spans = $this->worksheet_reader->getAttributeNsId('spans');

                if ($row_spans) {
                    $row_spans = explode(':', $row_spans);
                    $current_row_column_count = array_pop($row_spans); // Always get the last segment, regardless of spans structure.
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
                $current_row_column_count = array_pop($row_spans); // Always get the last segment, regardless of spans structure.
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
                    $value = $this->number_format->tryFormatValue($value, $style_id);

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
            $xf_num_fmt_ids = array();
            $number_formats = array();
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
                        $number_formats[$num_fmt_id] = $format_code;
                        break;

                    // <cellXfs><xf/></cellXfs> - check for format usages
                    case 'cellXfs':
                        $current_scope_is_cell_xfs = !$styles_reader->isClosingTag();
                        break;
                    case 'xf':
                        if (!$current_scope_is_cell_xfs || $styles_reader->isClosingTag()) {
                            break;
                        }

                        // Determine if number formatting is set for this cell.
                        $num_fmt_id = null;
                        if ($styles_reader->getAttributeNsId('numFmtId')) {
                            $applyNumberFormat = $styles_reader->getAttributeNsId('applyNumberFormat');
                            if ($applyNumberFormat === null || $applyNumberFormat === '1' || $applyNumberFormat === 'true') {
                                /* Number formatting is enabled either implicitly ('applyNumberFormat' not given)
                                 * or explicitly ('applyNumberFormat' is a true value). */
                                $num_fmt_id = (int) $styles_reader->getAttributeNsId('numFmtId');
                            }
                        }

                        // Determine and store correct formatting style.
                        if ($num_fmt_id !== null) {
                            // Number formatting has been enabled for this format.
                            // If format ID >= 164, it is a custom format and should be read from styleSheet\numFmts
                            $xf_num_fmt_ids[] = $num_fmt_id;
                        } else if ($styles_reader->getAttributeNsId('quotePrefix')) {
                            // "quotePrefix" automatically prepends the cell content with a ' symbol. This enforces a text format.
                            $xf_num_fmt_ids[] = null; // null = "Do not format anything".
                        } else {
                            $xf_num_fmt_ids[] = 0; // 0 = "General" format
                        }
                        break;
                }
            }

            $this->number_format->injectXfNumFmtIds($xf_num_fmt_ids);
            $this->number_format->injectNumberFormats($number_formats);

            $styles_reader->close();
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
