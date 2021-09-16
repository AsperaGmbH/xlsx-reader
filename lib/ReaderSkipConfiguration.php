<?php

namespace Aspera\Spreadsheet\XLSX;

/** Constants used to configure exclusion of undesired output. Used by ReaderConfiguration. */
class ReaderSkipConfiguration
{
    const SKIP_NONE = 0;
    const SKIP_EMPTY = 1;
    const SKIP_TRAILING_EMPTY = 2;
}
