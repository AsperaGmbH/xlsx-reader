<?php

namespace Aspera\Spreadsheet\XLSX;

/** Data object for worksheet related data */
class Worksheet
{
    /** @var string */
    private $name;

    /** @var string Relationship ID of this worksheet for matching with workbook data. */
    private $relationship_id;

    /**
     * @return string
     */
    public function getName(): string
    {
        return $this->name;
    }

    /**
     * @param string $name
     */
    public function setName(string $name): void
    {
        $this->name = $name;
    }

    /**
     * @return string
     */
    public function getRelationshipId(): string
    {
        return $this->relationship_id;
    }

    /**
     * @param string $relationship_id
     */
    public function setRelationshipId(string $relationship_id): void
    {
        $this->relationship_id = $relationship_id;
    }
}
