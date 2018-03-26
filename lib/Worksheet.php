<?php

namespace Aspera\Spreadsheet\XLSX;

/**
 * Data object for worksheet related data
 *
 * @author Aspera GmbH
 */
class Worksheet
{
    /**
     * Internal ID of this worksheet (Note: Not to be confused with the access index)
     *
     * @var int $id
     */
    private $id;

    /**
     * Name of the worksheet.
     *
     * @var string $name
     */
    private $name;

    /**
     * Relationship ID of this worksheet for matching with workbook data
     *
     * @var string $relationship_id
     */
    private $relationship_id;

    /**
     * @return int
     */
    public function getId()
    {
        return $this->id;
    }

    /**
     * @param int $id
     */
    public function setId($id)
    {
        $this->id = $id;
    }

    /**
     * @return string
     */
    public function getName()
    {
        return $this->name;
    }

    /**
     * @param string $name
     */
    public function setName($name)
    {
        $this->name = $name;
    }

    /**
     * @return string
     */
    public function getRelationshipId()
    {
        return $this->relationship_id;
    }

    /**
     * @param string $relationship_id
     */
    public function setRelationshipId($relationship_id)
    {
        $this->relationship_id = $relationship_id;
    }
}
