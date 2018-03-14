# xlsx-reader

__xlsx-reader__ is an extension of the XLSX-targeted spreadsheet reader that is part of [spreadsheet-reader](https://github.com/nuovo/spreadsheet-reader).
It delivers functionality to efficiently read in data contained within the given XLSX file.

As of now, the library does not support every potential type of content that can be included in XLSX. It can interpret only with a
very restricted subset of XLSX capabilities, such as basic cell value formatting and shared string evaluation.

### Requirements
*  PHP 5.6.0 or newer
*  PHP must have Zip file support (see <http://php.net/manual/en/zip.installation.php>)

### Usage

All data is read from the file sequentially, with each row being returned as a numeric array.
This is about the easiest way to read a file:

```php
<?php
use Aspera\Spreadsheet\XLSX\Reader;

$reader = new Reader('example.xlsx');
foreach ($reader as $row) {
    print_r($row);
}
```

Multiple sheet reading is also supported.

You can retrieve information about sheets contained in the file by calling the `getSheets()` method which returns an array with
sheet indexes as keys and sheet names as values. Then you can change the sheet that's currently being read by passing that index
to the `changeSheet($index)` method.

Example:

```php
<?php
use Aspera\Spreadsheet\XLSX\Reader;

$reader = new Reader('example.xlsx');
$sheets = $reader->getSheets();

foreach ($sheets as $index => $name) {
    echo 'Sheet #' . $index . ': ' . $name;

    $reader->changeSheet($index);

    foreach ($reader as $row) {
        print_r($row);
    }
}
```


If a sheet is changed to the same that is currently open, the position in the file still reverts to the beginning, so as to conform
to the same behavior as when changed to a different sheet.

### Testing

From the command line:

    php test.php path-to-spreadsheet.xls

In the browser:

    http://path-to-library/test.php?file=/path/to/spreadsheet.xls

### Notes about library performance
XLSX files use so called "shared strings" internally to optimize for cases where the same string is repeated multiple times.

Internally XLSX is an XML text that is parsed sequentially to extract data from it, however, in some cases these shared strings are a problem -
sometimes Excel may put all, or nearly all of the strings from the spreadsheet in the shared string file (which is a separate XML text), and not necessarily in the same
order. 

Worst case scenario is when it is in reverse order - for each string we need to parse the shared string XML from the beginning, if we want to avoid keeping the data in memory.
To that end, the XLSX parser has a cache for shared strings that is used if the total shared string count is not too high. In case you get out of memory errors, you can
try adjusting the *SHARED_STRING_CACHE_LIMIT* constant in XLSXReader to a lower one.

### Licensing
All of the code in this library is licensed under the MIT license as included in the [LICENSE.md](LICENSE.md) file.