# xlsx-reader

__xlsx-reader__ is an extension of the XLSX-targeted spreadsheet reader that is part of [spreadsheet-reader](https://github.com/nuovo/spreadsheet-reader).
It delivers functionality to efficiently read in data contained within the given XLSX file.

As of now, the library does not support every potential type of content that can be included in XLSX. It can interpret only with a
very restricted subset of XLSX capabilities, such as basic cell value formatting and shared string evaluation.

### Requirements
*  PHP 5.6.0 or newer, with at least the following optional features enabled:
    *  Zip (enabled by default; see <http://php.net/manual/en/zip.installation.php>)
    *  XMLReader (enabled by default; see <http://php.net/manual/en/xmlreader.installation.php>)

### Installation using Composer
The package is available on [Packagist](https://packagist.org/packages/aspera/xlsx-reader), you can install it using [Composer](https://getcomposer.org/)

```
composer require aspera/xlsx-reader
```

### Usage

All data is read from the file sequentially, with each row being returned as a numeric array.
This is about the easiest way to read a file:

```php
<?php
use Aspera\Spreadsheet\XLSX\Reader;
use Aspera\Spreadsheet\XLSX\SharedStringsConfiguration;

$options = array(
    'TempDir'                    => 'C:/Temp/',
    'SkipEmptyCells'             => false,
    'ReturnDateTimeObjects'      => true,
    'SharedStringsConfiguration' => new SharedStringsConfiguration(),
    'CustomFormats'              => array(20 => 'hh:mm')
);

$reader = new Reader($options);
$reader->open('example.xlsx');

foreach ($reader as $row) {
    print_r($row);
}

$reader->close();
```

Multiple sheet reading is also supported.

You can retrieve information about sheets contained in the file by calling the `getSheets()` method which returns an array with
sheet indexes as keys and Worksheet objects as values. Then you can change the sheet that's currently being read by passing that index
to the `changeSheet($index)` method.

Example:

```php
<?php
use Aspera\Spreadsheet\XLSX\Reader;
use Aspera\Spreadsheet\XLSX\Worksheet;

$reader = new Reader();
$reader->open('example.xlsx');
$sheets = $reader->getSheets();

/** @var Worksheet $sheet_data */
foreach ($sheets as $index => $sheet_data) {
    echo 'Sheet #' . $index . ': ' . $sheet_data->getName();

    $reader->changeSheet($index);

    foreach ($reader as $row) {
        print_r($row);
    }
}

$reader->close();
```

Extra configuration options available when constructing a new Reader() object:
- TempDir: provided temporary directory (used for unzipping all files like Styles.xml, Worksheet.xml...) must be writable and accessible by the XLSX Reader. 
- SkipEmptyCells: will skip empty values within any cell. If an entire row does not contain any value, only one empty (NULL) entry will be returned. 
- ReturnUnformatted: will return numeric values without number formatting. (Exception: Date/Time values. Those are controlled by the ReturnDateTimeObjects parameter.)
- ReturnDateTimeObjects: will return DateTime objects instead of formatted date-time strings.
- SharedStringsConfiguration: explained in "Notes about library performance".
- CustomFormats: matrix that will overwrite any format read by the parser. Array format must match the BUILT-IN formats list documented by Microsoft.
- ForceDateFormat: A date format that will be used for all date values read from the document.
- ForceTimeFormat: A time format that will be used for all time values read from the document.
- ForceDateTimeFormat: A datetime format that will be used for all datetime values read from the document.
- OutputColumnNames: If true, read data will be returned using alphabetical column indexes (A, B, AA, ZX, ...) instead of numeric indexes.

If a sheet is changed to the same that is currently open, the position in the file still reverts to the beginning, so as to conform
to the same behavior as when changed to a different sheet.

### Notes about library performance
XLSX files use so called "shared strings" to optimize file size for cases where the same string is repeated multiple times.
For larger documents, this list of shared strings can become quite large, causing either performance bottlenecks or
insane memory consumption when parsing the document.

To deal with this, several configuration options are supplied that you can use to control shared string handling behavior.
You can introduce them to a Reader instance via a SharedStringsConfiguration object, supplied to the constructor via the 
"SharedStringsConfiguration" option.

For a full list of available configuration values and their effects on performance/memory consumption, check the
code documentation found within the SharedStringsConfiguration class.

### Notes about unsupported features
This reader's purpose is to allow reading of basic data (text, numbers, dates...) from XLSX documents. As such,
there are no plans to extend support to include every single feature available for XLSX files. Only a minimal
subset of XLSX capabilities is supported.

In particular, the following should be noted in regards to unsupported features:
- Files with multiple internal shared strings files are not supported.
- Files with multiple internal styles definition files are not supported.
- Fractions are only partially supported. The results delivered by the reader might be slightly off from the original input.

### Licensing
All of the code in this library is licensed under the MIT license as included in the [LICENSE.md](LICENSE.md) file.