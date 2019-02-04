<?php

namespace Aspera\Spreadsheet\XLSX;

use ZipArchive;

/**
 * Data object containing all data related to a single 1:1 relationship declaration
 *
 * @author Aspera GmbH
 */
class RelationshipElement
{
    /** @var string Internal identifier of this file part */
    private $id;

    /** @var bool Element validity flag; If false, this element was not found or might be corrupted. */
    private $is_valid;

    /** @var string Path to this element, as per the context its information was retrieved from. */
    private $original_path;

    /** @var string Absolute path to the file associated with this element for access. */
    private $access_path;

    /**
     * @return string
     */
    public function getId()
    {
        return $this->id;
    }

    /**
     * @param string $id
     */
    public function setId($id)
    {
        $this->id = $id;
    }

    /**
     * @return bool
     */
    public function isValid()
    {
        return $this->is_valid;
    }

    /**
     * @param bool $is_valid
     */
    public function setIsValid($is_valid)
    {
        $this->is_valid = $is_valid;
    }

    /**
     * @return string
     */
    public function getOriginalPath()
    {
        return $this->original_path;
    }

    /**
     * @param string $original_path
     */
    public function setOriginalPath($original_path)
    {
        $this->original_path = $original_path;
    }

    /**
     * @return string
     */
    public function getAccessPath()
    {
        return $this->access_path;
    }

    /**
     * @param string $access_path
     */
    public function setAccessPath($access_path)
    {
        $this->access_path = $access_path;
    }

    /**
     * Checks the given zip file for the element described by this object and sets validity flag accordingly.
     *
     * @param ZipArchive $zip
     */
    public function setValidityViaZip($zip)
    {
        $this->setIsValid($zip->locateName($this->getOriginalPath()) !== false);
    }
}
