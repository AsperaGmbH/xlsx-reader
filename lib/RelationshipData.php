<?php

namespace Aspera\Spreadsheet\XLSX;

use RuntimeException;
use ZipArchive;

/**
 * Functionality to work with relationship data (.rels files)
 * Also contains all relationship data it previously evaluated for structured retrieval.
 */
class RelationshipData
{
    /**
     * Directory separator character used in zip file internal paths.
     * Is supposed to always be a forward slash, even on systems with a different directory separator (e.g. Windows).
     *
     * @var string ZIP_DIR_SEP
     */
    const ZIP_DIR_SEP = '/';

    /** @var RelationshipElement Workbook file meta information. Only one element exists per file. */
    private $workbook;

    /** @var array Worksheet files meta information, saved as a list of RelationshipElement instances. */
    private $worksheets = array();

    /** @var array SharedStrings files meta information, saved as a list of RelationshipElement instances. */
    private $shared_strings = array();

    /** @var array Styles files meta information, saved as a list of RelationshipElement instances. */
    private $styles = array();

    /**
     * Returns the workbook relationship element, if a valid one has been obtained previously.
     * Returns null otherwise.
     *
     * @return null|RelationshipElement
     */
    public function getWorkbook()
    {
        if (isset($this->workbook) && $this->workbook->isValid()) {
            return $this->workbook;
        }
        return null;
    }

    /**
     * Returns data of all found valid shared string elements.
     * Returns array of RelationshipElement elements.
     *
     * @return array[RelationshipElement]
     */
    public function getSharedStrings()
    {
        $return_list = array();
        foreach ($this->shared_strings as $shared_string_element) {
            if ($shared_string_element->isValid()) {
                $return_list[] = $shared_string_element;
            }
        }
        return $return_list;
    }

    /**
     * Returns all worksheet data of all found valid worksheet elements.
     * Returns array of RelationshipElement elements.
     *
     * @return array
     */
    public function getWorksheets()
    {
        $return_list = array();
        foreach ($this->worksheets as $worksheet_element) {
            if ($worksheet_element->isValid()) {
                $return_list[] = $worksheet_element;
            }
        }
        return $return_list;
    }

    /**
     * Returns all styles data of all found valid styles elements.
     * Returns array of RelationshipElement elements
     *
     * @return array
     */
    public function getStyles()
    {
        $return_list = array();
        foreach ($this->styles as $styles_element) {
            if ($styles_element->isValid()) {
                $return_list[] = $styles_element;
            }
        }
        return $return_list;
    }

    /**
     * Navigates through the XLSX file using .rels files, gathering up found file parts along the way.
     * Results are saved in internal variables for later retrieval.
     *
     * @param   ZipArchive  $zip    Handle to zip file to read relationship data from
     *
     * @throws  RuntimeException
     */
    public function __construct(ZipArchive $zip)
    {
        // Start with root .rels file. It will point us towards the worksheet file.
        $root_rel_file = self::toRelsFilePath(''); // empty string returns root path
        $this->evaluateRelationshipFromZip($zip, $root_rel_file);

        // Quick check: Workbook should have been retrieved from root relationship file.
        if (!isset($this->workbook) || !$this->workbook->isValid()) {
            throw new RuntimeException('Could not locate workbook data.');
        }

        // The workbook .rels file should point us towards all other required files.
        $workbook_rels_file_path = self::toRelsFilePath($this->workbook->getOriginalPath());
        $this->evaluateRelationshipFromZip($zip, $workbook_rels_file_path);
    }

    /**
     * Read through the .rels data of the given .rels file from the given zip handle
     * and save all included file data to internal variables.
     *
     * @param   ZipArchive  $zip
     * @param   string      $file_zipname
     *
     * @throws  RuntimeException
     */
    private function evaluateRelationshipFromZip(ZipArchive $zip, $file_zipname)
    {
        if ($zip->locateName($file_zipname) === false) {
            throw new RuntimeException('Could not read relationship data. File [' . $file_zipname . '] could not be found.');
        }

        $rels_reader = new OoxmlReader();
        $rels_reader->setDefaultNamespaceIdentifierElements(OoxmlReader::NS_RELATIONSHIPS_PACKAGELEVEL);
        $rels_reader->setDefaultNamespaceIdentifierAttributes(OoxmlReader::NS_NONE);
        $rels_reader->xml($zip->getFromName($file_zipname));
        while ($rels_reader->read() !== false) {
            if (!$rels_reader->matchesElement('Relationship') || $rels_reader->isClosingTag()) {
                // This element is not important to us. Skip.
                continue;
            }

            // Only the last part of the relationship type definition matters to us.
            $rels_type = $rels_reader->getAttributeNsId('Type');
            if (!preg_match('~([^/]+)/?$~', $rels_type, $type_regexp_matches)) {
                throw new RuntimeException(
                    'Invalid type definition found: [' . $rels_type . ']'
                    . ' Relationship could not be evaluated.'
                );
            }

            // Adjust target path (making it absolute without leading slash) so that we can easily use it for zip checks later.
            $target_path = $rels_reader->getAttributeNsId('Target');
            if (strpos($target_path, self::ZIP_DIR_SEP) === 0) {
                // target_path is already absolute, but we need to remove the leading slash.
                $target_path = substr($target_path, 1);
            } elseif (preg_match('~(.*' . self::ZIP_DIR_SEP . ')_rels~', $file_zipname, $path_matches)) {
                // target_path is relative. Add path of this .rels file to target path
                $target_path = $path_matches[1] . $target_path;
            }

            // Assemble and store element data
            $element_data = new RelationshipElement();
            $element_data->setId($rels_reader->getAttributeNsId('Id'));
            $element_data->setOriginalPath($target_path);
            $element_data->setValidityViaZip($zip);
            switch ($type_regexp_matches[1]) {
                case 'officeDocument':
                    $this->workbook = $element_data;
                    break;
                case 'worksheet':
                    $this->worksheets[] = $element_data;
                    break;
                case 'sharedStrings':
                    $this->shared_strings[] = $element_data;
                    break;
                case 'styles':
                    $this->styles[] = $element_data;
                    break;
                default:
                    // nop
                    break;
            }
        }
    }

    /**
     * Returns the path to the .rels file for the given file path.
     * Example: xl/workbook.xml => xl/_rels/workbook.xml.rels
     *
     * @param   string  $file_path
     * @return  string
     */
    private static function toRelsFilePath($file_path)
    {
        // Normalize directory separator character
        $file_path = str_replace('\\', self::ZIP_DIR_SEP, $file_path);

        // Split path in 2 parts around last dir seperator: [path/to/file]/[file.xml]
        $last_slash_pos = strrpos($file_path, '/');
        if ($last_slash_pos === false) {
            // No final slash; file.xml => _rels/file.xml.rels
            // This also implicitly handles the root .rels file, always found under "_rels/.rels"
            $file_path = '_rels/' . $file_path . '.rels';
        } elseif ($last_slash_pos == strlen($file_path) - 1) {
            // Trailing slash; some/folder/ => some/_rels/folder.rels
            $file_path = $file_path . '_rels/.rels';
        } else {
            // File with path; some/folder/file.xml => some/folder/_rels/file.xml.rels
            $file_path = preg_replace('~([^/]+)$~', '_rels/$1.rels', $file_path);
        }
        return $file_path;
    }
}
