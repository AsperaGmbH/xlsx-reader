SpreadsheetReader_XLSX is an extension of the XLSX-targeted spreadsheet reader that is part of https://github.com/nuovo/spreadsheet-reader.
It delivers functionality to efficiently read in data contained within the given XLSX file.

As of now, the library does not support every potential type of content that can be included in XLSX. It can interpret only with a
very restricted subset of XLSX capabilities, such as basic cell value formatting and shared string evaluation.

### Requirements:
*  PHP 5.6.0 or newer
*  PHP must have Zip file support (see http://php.net/manual/en/zip.installation.php)

### Usage:

All data is read from the file sequentially, with each row being returned as a numeric array.
This is about the easiest way to read a file:

    <?php
        require_once('SpreadsheetReader_XLSX.php');
    
        $Reader = new SpreadsheetReader_XLSX('example.xlsx');
        foreach ($Reader as $Row)
        {
            print_r($Row);
        }
    ?>

Multiple sheet reading is also supported.

You can retrieve information about sheets contained in the file by calling the `Sheets()` method which returns an array with
sheet indexes as keys and sheet names as values. Then you can change the sheet that's currently being read by passing that index
to the `ChangeSheet($Index)` method.

Example:

    <?php
        $Reader = new SpreadsheetReader_XLSX('example.xlsx');
        $Sheets = $Reader -> Sheets();
    
        foreach ($Sheets as $Index => $Name)
        {
            echo 'Sheet #'.$Index.': '.$Name;
    
            $Reader -> ChangeSheet($Index);
    
            foreach ($Reader as $Row)
            {
                print_r($Row);
            }
        }
    ?>

If a sheet is changed to the same that is currently open, the position in the file still reverts to the beginning, so as to conform
to the same behavior as when changed to a different sheet.

### Testing

From the command line:

    php test.php path-to-spreadsheet.xls

In the browser:

    http://path-to-library/test.php?File=/path/to/spreadsheet.xls

### Notes about library performance
*  XLSX files use so called "shared strings" internally to optimize for cases where the same string is repeated multiple times.
	Internally XLSX is an XML text that is parsed sequentially to extract data from it, however, in some cases these shared strings are a problem -
	sometimes Excel may put all, or nearly all of the strings from the spreadsheet in the shared string file (which is a separate XML text), and not necessarily in the same
	order. Worst case scenario is when it is in reverse order - for each string we need to parse the shared string XML from the beginning, if we want to avoid keeping the data in memory.
	To that end, the XLSX parser has a cache for shared strings that is used if the total shared string count is not too high. In case you get out of memory errors, you can
	try adjusting the *SHARED_STRING_CACHE_LIMIT* constant in SpreadsheetReader_XLSX to a lower one.

### Licensing
All of the code in this library is licensed under the MIT license as included in the LICENSE file.