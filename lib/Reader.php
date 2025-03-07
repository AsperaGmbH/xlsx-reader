<?php

namespace Aspera\Spreadsheet\XLSX;

use Iterator;
use LogicException;
use RuntimeException;
use ZipArchive;
use Exception;
use InvalidArgumentException;

/** Class for reading XLSX file contents. */
class Reader implements Iterator
{
    /** @var ReaderConfiguration The provided configuration data. */
    private $configuration;

    /** @var NumberFormat Singleton for formatting methods. Also stores the provided formatting configuration. */
    private $number_format;

    /**
     * Full path to the directory that should be used for all file operations.
     * Not to be confused with ReaderConfiguration::$temp_dir, which is the user-supplied temporary directory,
     * which is a parent directory of $temp_dir_for_reader.
     *
     * @var string
     */
    private $temp_dir_for_reader = '';

    /** @var array Temporary files created while reading the document. */
    private $temp_files = array();

    /** @var ?RelationshipData File paths and -identifiers to all relevant parts of the read XLSX file */
    private $relationship_data;

    /** @var string Path to the current worksheet XML file. */
    private $worksheet_path;

    /** @var array Data about separate sheets in the file. */
    private $sheets;

    /** @var ?OoxmlReader XML reader object for the current worksheet XML file. */
    private $worksheet_reader = null;

    /** @var ?SharedStrings Shared strings handler. */
    private $shared_strings = null;

    /** @var bool Whether the reader has been initialized with open() yet. */
    private $reader_is_open = false;

    /** @var bool true when the current Iterator position is valid. An "invalid" position indicates EOF. */
    private $valid = false;

    /** @var bool Whether the reader is currently pointing at a starting <row> node. */
    private $reader_points_at_new_row = false;

    /** @var bool If true, current row output was already adjusted by a previous call to current(). */
    private $row_output_adjusted = false;

    /** @var int Current row number in the file. Identical to the <row>'s "r" attribute. 1-based. */
    private $row_number = 0; // Initially 0 to indicate that the first row hasn't been read yet.

    /** @var int Amount of rows skipped due to lookahead. Only used when skippedRows = SKIP_TRAILING_EMPTY. */
    private $skipped_empty_rows = 0;

    /** @var ?array Contents of last read row. null when no row has been read yet. */
    private $current_row = null;

    /** @var ?array Contents of next filled row. Only used when skippedRows = SKIP_TRAILING_EMPTY. */
    private $next_filled_row = null;

    /** @var array Structure of an empty row for this document. Only used when skippedRows = SKIP_TRAILING_EMPTY. */
    private $empty_row_structure = array();

    /**
     * @throws Exception
     */
    public function __construct(?ReaderConfiguration $configuration = null)
    {
        if (!isset($configuration)) {
            $configuration = new ReaderConfiguration();
        }
        $this->configuration = $configuration;
        $this->number_format = new NumberFormat($configuration);

        $this->initTempDir($configuration->getTempDir());
    }

    public function __destruct()
    {
        $this->close();
    }

    /**
     * Open the given file and prepare everything for the reading of data.
     *
     * @throws Exception
     */
    public function open(string $file_path): void
    {
        if ($this->reader_is_open) {
            throw new LogicException('Reader was already opened.');
        }

        if (!is_readable($file_path)) {
            throw new RuntimeException('File not readable (' . $file_path . ')');
        }

        if (!mkdir($this->temp_dir_for_reader, 0777, true) || !file_exists($this->temp_dir_for_reader)) {
            throw new RuntimeException(
                'Could neither create nor confirm existence of temporary directory (' . $this->temp_dir_for_reader . ')'
            );
        }

        $zip = new ZipArchive;
        $status = $zip->open($file_path);
        if ($status !== true) {
            throw new RuntimeException('File not readable (' . $file_path . ') (Error ' . $status . ')');
        }

        // Mark as initialized *now*; The following methods expect this.
        $this->reader_is_open = true;

        $this->relationship_data = new RelationshipData($zip, $this->configuration);
        $this->initWorkbookData($zip);
        $this->initSharedStrings($zip, $this->configuration->getSharedStringsConfiguration());
        $this->initStyles($zip);
        $this->initWorksheets($zip);

        $zip->close();
    }

    /**
     * Free all connected resources.
     */
    public function close(): void
    {
        if ($this->worksheet_reader) {
            $this->worksheet_reader->close();
            $this->worksheet_reader = null;
        }

        if ($this->shared_strings) {
            // Closing the shared string handler will also close all still opened shared string temporary work files.
            $this->shared_strings->close();
            $this->shared_strings = null;
        }

        $this->deleteTempfiles();

        $this->worksheet_path = '';
        $this->sheets = array();
        $this->relationship_data = null;
        $this->current_row = null;

        $this->reader_is_open = false;
    }

    /**
     * Retrieves an array with information about sheets in the current file
     *
     * @return Worksheet[] List of sheets. Sheet index (0-based) is used as the key of each element.
     *
     * @throws LogicException
     */
    public function getSheets(): array
    {
        if (!$this->reader_is_open) {
            throw new LogicException('Reader was not intialized via open() yet.');
        }

        return array_values($this->sheets);
    }

    /**
     * Changes the current sheet in the file to the sheet with the given index.
     *
     * @return bool True if sheet was changed successfully, false otherwise.
     *
     * @throws Exception
     */
    public function changeSheet(int $sheet_index): bool
    {
        if (!$this->reader_is_open) {
            throw new LogicException('Reader was not intialized via open() yet.');
        }

        $sheets = $this->getSheets(); // Note: Realigns indexes to an auto increment.
        if (!isset($sheets[$sheet_index])) {
            return false;
        }

        $target_sheet = $sheets[$sheet_index];

        // The path to the target worksheet file can be obtained via the relationship id reference.
        $target_relationship_id = $target_sheet->getRelationshipId();
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
     * Rewind the reader to the first row.
     *
     * @throws Exception
     */
    public function rewind(): void
    {
        if (!$this->reader_is_open) {
            throw new LogicException('Reader was not intialized via open() yet.');
        }

        if ($this->worksheet_reader) {
            $this->worksheet_reader->close();
        } else {
            $this->worksheet_reader = new OoxmlReader();
            $this->worksheet_reader->setDefaultNamespaceIdentifierElements(OoxmlReader::NS_XLSX_MAIN);
            $this->worksheet_reader->setDefaultNamespaceIdentifierAttributes(OoxmlReader::NS_NONE);
        }

        $this->worksheet_reader->open(
            $this->worksheet_path,
            null,
            $this->configuration->getXmlReaderFlags()
        );

        // Reset read-relevant internal variables to their initial values to avoid conflicts between subsequent rewinds.
        $this->reader_points_at_new_row = false;
        $this->row_output_adjusted = false;
        $this->row_number = 0;
        $this->skipped_empty_rows = 0;
        $this->current_row = null;
        $this->next_filled_row = null;
        $this->empty_row_structure = array();

        $this->next();
    }

    /**
     * Move to the next row.
     *
     * @throws Exception
     */
    public function next(): void
    {
        if (!$this->reader_is_open) {
            throw new LogicException('Reader was not intialized via open() yet.');
        }

        $this->current_row = array();
        $this->row_output_adjusted = false;

        // Handle skipped empty rows due to previous lookahead as part of SKIP_TRAILING_EMPTY.
        if ($this->skipped_empty_rows > 0) {
            $this->row_number++;
            $this->current_row = $this->empty_row_structure;
            $this->skipped_empty_rows--;
            return;
        }
        if ($this->next_filled_row !== null) {
            $this->row_number++;
            $this->current_row = $this->next_filled_row;
            $this->next_filled_row = null;
            return;
        }

        $acceptable_current_state = false; // true when current() would return an acceptable value as per configuration.
        $initial_row_number = $this->row_number;
        while (!$acceptable_current_state) {
            $this->read_next_row();

            if (!$this->valid()) {
                return; // EOF reached.
            }

            // Check if contents of current row fulfill skipEmptyRows requirements.
            switch ($this->configuration->getSkipEmptyRows()) {
                case ReaderSkipConfiguration::SKIP_NONE:
                    $acceptable_current_state = true;
                    break;
                case ReaderSkipConfiguration::SKIP_EMPTY:
                    if (implode('', $this->current_row) !== '') {
                        $acceptable_current_state = true;
                    }
                    // This is an empty row that should not be output. Loop to next row.
                    break;
                case ReaderSkipConfiguration::SKIP_TRAILING_EMPTY:
                    if (implode('', $this->current_row) === '') {
                        // This is an empty row that MAY not be wanted in the output. Look ahead for more content.
                        $this->skipped_empty_rows++;
                        $this->empty_row_structure = $this->current_row;
                    } else {
                        $acceptable_current_state = true;
                        if ($this->skipped_empty_rows > 0) {
                            // Intermediate rows were skipped by lookahead. Start returning them.
                            $this->next_filled_row = $this->current_row;
                            $this->current_row = $this->empty_row_structure;
                            $this->row_number = $initial_row_number + 1;
                            $this->skipped_empty_rows--;
                        }
                    }
                    break;
            }
        }
    }

    /**
     * Reads to and through the next <row> and updates this instance's fields accordingly.
     *
     * @throws Exception
     */
    private function read_next_row(): void
    {
        $this->row_number++;

        // Ensure that the read pointer is pointing at an opening <row> element.
        while (!$this->reader_points_at_new_row && $this->valid = $this->worksheet_reader->read()) {
            if ($this->worksheet_reader->matchesElement('row')) {
                $this->reader_points_at_new_row = true;
            }
        }
        if (!$this->reader_points_at_new_row) {
            // No (further) rows found for reading.
            return;
        }

        /* Get the row spanning area (stored as e.g. "1:12", or more rarely, "1:3 6:8 11:12")
         * so that the last cells will be present, even if empty. */
        $row_spans = $this->worksheet_reader->getAttributeNsId('spans');
        if ($row_spans) {
            $row_spans = explode(':', $row_spans);
            $current_row_column_count = array_pop($row_spans); // Always get the last segment, regardless of spans structure.
        } else {
            $current_row_column_count = 0;
        }

        // If configured: Return empty strings for empty values
        if (    $current_row_column_count > 0
            &&  $this->configuration->getSkipEmptyCells() === ReaderSkipConfiguration::SKIP_NONE
        ) {
            $this->current_row = array_fill(0, $current_row_column_count, '');
        }

        if ((int) $this->worksheet_reader->getAttributeNsId('r') !== $this->row_number) {
            // We just skipped over one or multiple empty row(s). Keep current reader state and return empty cells.
            return;
        }

        // From here on out, successive next() calls must start with a read() for the next row.
        $this->reader_points_at_new_row = false;

        // Handle self-closing row tags (e.g. <row r="1"/>) caused by e.g. usage of thick borders in adjacent cells.
        $this->worksheet_reader->moveToElement(); // Necessary for isEmptyElement to work correctly.
        if ($this->worksheet_reader->isEmptyElement) {
            return;
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
                    if ($this->configuration->getSkipEmptyCells() === ReaderSkipConfiguration::SKIP_NONE) {
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
                    if (    $value === ''
                        &&  $this->configuration->getSkipEmptyCells() !== ReaderSkipConfiguration::SKIP_NONE
                    ) {
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
         * Only empty cells in-between and on the left side are added. */
        if (    $max_index + 1 > $cell_count
            &&  $this->configuration->getSkipEmptyCells() !== ReaderSkipConfiguration::SKIP_EMPTY
        ) {
            $this->current_row += array_fill(0, $max_index + 1, '');
            ksort($this->current_row);
        }

        // If requested, remove trailing empty cells.
        if ($this->configuration->getSkipEmptyCells() === ReaderSkipConfiguration::SKIP_TRAILING_EMPTY) {
            foreach (array_reverse(array_keys($this->current_row)) as $k) {
                if ($this->current_row[$k] !== '') {
                    break;
                }
                unset($this->current_row[$k]);
            }
        }
    }

    /**
     * Returns whether the iterator is pointing at a valid position (read: before EOF) right now.
     *
     * @throws LogicException
     */
    public function valid(): bool
    {
        if (!$this->reader_is_open) {
            throw new LogicException('Reader was not intialized via open() yet.');
        }

        return $this->valid;
    }

    /**
     * Returns the current row, previously iterated to via next().
     *
     * @throws LogicException
     */
    public function current(): array
    {
        if (!$this->valid()) {
            throw new LogicException('Invalid reader state. Check valid() before using current().');
        }

        // Lazy output adjustment. Doing this here allows faster iteration via repeated next() calls.
        if (!$this->row_output_adjusted) {
            $this->current_row = self::adjustRowOutput($this->current_row);
            $this->row_output_adjusted = true;
        }

        return $this->current_row;
    }

    /**
     * Return the current row number.
     *
     * @throws LogicException
     */
    public function key(): int
    {
        if (!$this->valid()) {
            throw new LogicException('Invalid reader state. Check valid() before using key().');
        }

        return $this->row_number;
    }

    /**
     * Takes the column letter and converts it to a numerical index (0-based)
     *
     * @param  string  $letter Letter(s) to convert
     * @return int     Numeric index (0-based)
     *
     * @throws InvalidArgumentException Thrown when the given $letter includes non-letter characters.
     */
    public static function indexFromColumnLetter(string $letter): int
    {
        $letter = strtoupper($letter);
        $result = 0;
        for ($i = strlen($letter) - 1, $j = 0; $i >= 0; $i--, $j++) {
            $ord = ord($letter[$i]) - 64;
            if ($ord > 26) {
                // This does not seem to be a letter. Someone must have given us an invalid value.
                throw new InvalidArgumentException(
                    'The given ASCII ' . ($ord + 64) . ' does not represent a letter character.'
                );
            }
            $result += $ord * (26 ** $j);
        }

        return $result - 1;
    }

    /**
     * Converts the given column index to an XLSX-style [A-Z] column identifier string.
     */
    public static function columnLetterFromIndex(int $index): string
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
     */
    private function adjustRowOutput(array $column_values): array
    {
        if (!$this->configuration->getOutputColumnNames()) {
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
     * @throws RuntimeException
     */
    private function initTempDir(string $base_temp_dir): void
    {
        if (!is_writable($base_temp_dir)) {
            throw new RuntimeException('Temporary directory (' . $base_temp_dir . ') is not writable');
        }
        $base_temp_dir = rtrim($base_temp_dir, DIRECTORY_SEPARATOR);

        /** @noinspection NonSecureUniqidUsageInspection */
        $this->temp_dir_for_reader = $base_temp_dir . DIRECTORY_SEPARATOR . uniqid() . DIRECTORY_SEPARATOR;
    }

    /**
     * Read general workbook information from the given zip into memory.
     *
     * @throws Exception
     */
    private function initWorkbookData(ZipArchive $zip): void
    {
        $workbook = $this->relationship_data->getWorkbook();
        if (!$workbook) {
            throw new RuntimeException('workbook data not found in XLSX file');
        }
        $workbook_xml = $zip->getFromName($workbook->getOriginalPath());

        $this->sheets = array();
        $workbook_reader = new OoxmlReader();
        $workbook_reader->setDefaultNamespaceIdentifierElements(OoxmlReader::NS_XLSX_MAIN);
        $workbook_reader->setDefaultNamespaceIdentifierAttributes(OoxmlReader::NS_NONE);
        $workbook_reader->xml($workbook_xml, null, $this->configuration->getXmlReaderFlags());
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
     * @throws Exception
     */
    private function initWorksheets(ZipArchive $zip): void
    {
        // Sheet order determining value: relative sheet positioning within the document (rId)
        ksort($this->sheets);

        // Extract worksheets to temporary work directory
        foreach ($this->relationship_data->getWorksheets() as $worksheet) {
            $worksheet_path_zip = $worksheet->getOriginalPath();
            $worksheet_path_conv = str_replace(RelationshipData::ZIP_DIR_SEP, DIRECTORY_SEPARATOR, $worksheet_path_zip);
            $worksheet_path_unzipped = $this->temp_dir_for_reader . $worksheet_path_conv;
            if (!$zip->extractTo($this->temp_dir_for_reader, $worksheet_path_zip)) {
                $message = 'Could not extract file [' . $worksheet_path_zip . '] to directory [' . $this->temp_dir_for_reader . '].';
                $this->reportZipExtractionFailure($zip, $message);
            }
            $worksheet->setAccessPath($worksheet_path_unzipped);
            $this->temp_files[] = $worksheet_path_unzipped;
        }

        // Set first sheet as current sheet
        if (!$this->changeSheet(0)) {
            throw new RuntimeException('Sheet cannot be changed.');
        }
    }

    /**
     * Read shared strings data from the given zip into memory as configured via the given configuration object
     * and potentially create temporary work files for easy retrieval of shared string data.
     *
     * @throws Exception
     */
    private function initSharedStrings(
        ZipArchive $zip,
        SharedStringsConfiguration $shared_strings_configuration
    ): void {
        $shared_strings = $this->relationship_data->getSharedStrings();
        if (count($shared_strings) > 0) {
            /* Currently, documents with multiple shared strings files are not supported.
            *  Only the first shared string file will be used. */
            $first_shared_string_element = $shared_strings[0];

            // Determine target directory and path for the extracted file
            $inzip_path = $first_shared_string_element->getOriginalPath();
            $inzip_path_for_outzip = str_replace(RelationshipData::ZIP_DIR_SEP, DIRECTORY_SEPARATOR, $inzip_path);
            $dir_of_extracted_file = $this->temp_dir_for_reader . dirname($inzip_path_for_outzip) . DIRECTORY_SEPARATOR;
            $filename_of_extracted_file = basename($inzip_path_for_outzip);
            $path_to_extracted_file = $dir_of_extracted_file . $filename_of_extracted_file;

            // Extract file and note it in relevant variables
            if (!$zip->extractTo($this->temp_dir_for_reader, $inzip_path)) {
                $message = 'Could not extract file [' . $inzip_path . '] to directory [' . $this->temp_dir_for_reader . '].';
                $this->reportZipExtractionFailure($zip, $message);
            }

            $first_shared_string_element->setAccessPath($path_to_extracted_file);
            $this->temp_files[] = $path_to_extracted_file;

            // Initialize SharedStrings
            $this->shared_strings = new SharedStrings(
                $dir_of_extracted_file,
                $filename_of_extracted_file,
                $shared_strings_configuration,
                $this->configuration
            );
        }
    }

    /**
     * Reads and prepares information on styles declared by the document for later usage.
     *
     * @throws Exception
     */
    private function initStyles(ZipArchive $zip): void
    {
        $styles = $this->relationship_data->getStyles();
        if (count($styles) > 0) {
            /* Currently, documents with multiple styles files are not supported.
            *  Only the first styles file will be used. */
            $first_styles_element = $styles[0];

            $styles_xml = $zip->getFromName($first_styles_element->getOriginalPath());

            // Read cell style definitions and store them in internal variables
            $styles_reader = new OoxmlReader();
            $styles_reader->setDefaultNamespaceIdentifierElements(OoxmlReader::NS_XLSX_MAIN);
            $styles_reader->setDefaultNamespaceIdentifierAttributes(OoxmlReader::NS_NONE);
            $styles_reader->xml($styles_xml, null, $this->configuration->getXmlReaderFlags());
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
    private function deleteTempfiles(): void
    {
        foreach ($this->temp_files as $temp_file) {
            @unlink($temp_file);
        }
        $this->temp_files = array();

        // Better safe than sorry - shouldn't try deleting '.' or '/', or '..'.
        if (strlen($this->temp_dir_for_reader) > 2) {
            @rmdir($this->temp_dir_for_reader . 'xl' . DIRECTORY_SEPARATOR . 'worksheets');
            @rmdir($this->temp_dir_for_reader . 'xl');
            @rmdir($this->temp_dir_for_reader);
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
    private function reportZipExtractionFailure(ZipArchive $zip, string $message = ''): void
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
