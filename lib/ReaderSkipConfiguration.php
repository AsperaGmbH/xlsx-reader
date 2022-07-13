<?php

namespace Aspera\Spreadsheet\XLSX;

/** Constants used to configure exclusion of undesired output. Used by ReaderConfiguration. */
class ReaderSkipConfiguration
{
    public const SKIP_NONE = 0;
    public const SKIP_EMPTY = 1;
    public const SKIP_TRAILING_EMPTY = 2;
}
