<?php

namespace Aspera\Spreadsheet\XLSX;

/** Data object for worksheet related data */
class Worksheet
{
    /** @var string */
    private $name;

    /** @var string Relationship ID of this worksheet for matching with workbook data. */
    private $relationship_id;

    public function getName(): string
    {
        return $this->name;
    }

    public function setName(string $name): void
    {
        $this->name = $name;
    }

    public function getRelationshipId(): string
    {
        return $this->relationship_id;
    }

    public function setRelationshipId(string $relationship_id): void
    {
        $this->relationship_id = $relationship_id;
    }
}
