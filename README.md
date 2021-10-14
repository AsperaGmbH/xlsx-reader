# xlsx-reader

__xlsx-reader__ is an extension of the XLSX-targeted spreadsheet reader that is part of [spreadsheet-reader](https://github.com/nuovo/spreadsheet-reader). 

It delivers functionality to efficiently read in data contained within the given XLSX file.

The focus of this library is on delivering data contained in XLSX spreadsheet cells, not the document's styling.
As such, the library offers no support for XLSX capabilities that aren't strictly necessary to achieve this goal.
Only basic cell value formatting and shared string functionalities are supported.

### Requirements
*  PHP 5.6.0 or newer, with at least the following optional features enabled:
    *  Zip (enabled by default; see <http://php.net/manual/en/zip.installation.php>)
    *  XMLReader (enabled by default; see <http://php.net/manual/en/xmlreader.installation.php>)

### Installation using Composer
The package is available on [Packagist](https://packagist.org/packages/aspera/xlsx-reader).
You can install it using [Composer](https://getcomposer.org/).

```
composer require aspera/xlsx-reader
```

### Usage

All data is read from the file sequentially, with each row being returned as an array of columns.

```php
<?php
use Aspera\Spreadsheet\XLSX\Reader;

$reader = new Reader();
$reader->open('example.xlsx');

foreach ($reader as $row) {
    print_r($row);
}

$reader->close();
```

XLSX files with multiple worksheets are also supported.
The method getSheets() returns an array with sheet indexes as keys and Worksheet objects as values.
The method changeSheet($index) is used to switch between sheets to read.

```php
<?php
use Aspera\Spreadsheet\XLSX\Reader;
use Aspera\Spreadsheet\XLSX\Worksheet;

$reader = new Reader();
$reader->open('example.xlsx');
$sheets = $reader->getSheets();

/** @var Worksheet $sheet_data */
foreach ($sheets as $index => $sheet_data) {
    $reader->changeSheet($index);
    echo 'Sheet #' . $index . ': ' . $sheet_data->getName();

    // Note: Any call to changeSheet() resets the current read position to the beginning of the selected sheet.
    foreach ($reader as $row) {
        print_r($row);
    }
}

$reader->close();
```

Options to tune the reader's behavior and output can be specified via a ReaderConfiguration instance.

For a full list of supported options and their effects, consult the in-code documentation of ReaderConfiguration.

```php
<?php
use Aspera\Spreadsheet\XLSX\Reader;
use Aspera\Spreadsheet\XLSX\ReaderConfiguration;
use Aspera\Spreadsheet\XLSX\ReaderSkipConfiguration;

$reader_configuration = (new ReaderConfiguration())
  ->setTempDir('C:/Temp/')
  ->setSkipEmptyCells(ReaderSkipConfiguration::SKIP_EMPTY)
  ->setReturnDateTimeObjects(true)
  ->setCustomFormats(array(20 => 'hh:mm'));
// For a full list of supported options and their effects, consult the in-code documentation of ReaderConfiguration.

$spreadsheet = new Reader($reader_configuration);
```

### Notes about library performance
XLSX files use so-called "shared strings" to optimize file sizes for cases where the same string is repeated multiple times.
For larger documents, this list of shared strings can become quite large, causing either performance bottlenecks or
high memory consumption when parsing the document.

To deal with this, the reader selects sensible defaults for maximum RAM consumption. Once this memory limit has been
exhausted, the file system is used for further optimization strategies.

To configure this behavior in detail, e.g. to increase the amount of memory available to the reader, a SharedStringsConfiguration
instance can be attached to the ReaderConfiguration instance supplied to the reader's constructor.

For a full list of supported options and their effects, consult the in-code documentation of SharedStringsConfiguration.

```php
<?php
use Aspera\Spreadsheet\XLSX\Reader;
use Aspera\Spreadsheet\XLSX\ReaderConfiguration;
use Aspera\Spreadsheet\XLSX\SharedStringsConfiguration;

$shared_strings_configuration = (new SharedStringsConfiguration())
    ->setCacheSizeKilobyte(16 * 1024)
    ->setUseOptimizedFiles(false);
// For a full list of supported options and their effects, consult the in-code documentation of SharedStringsConfiguration.

$reader_configuration = (new ReaderConfiguration())
  ->setSharedStringsConfiguration($shared_strings_configuration);

$spreadsheet = new Reader($reader_configuration);
```

### Notes about unsupported features
This reader's purpose is to allow reading of basic data (text, numbers, dates...) from XLSX documents. As such,
there are no plans to extend support to include all features available for XLSX files. Only a minimal
subset of XLSX capabilities is supported.

In particular, the following should be noted in regard to unsupported features:
- Display cell width is disregarded. As a result, in cases in which popular xlsx editors would shorten values using
  scientific notation or "#####"-placeholders, the reader will return un-shortened values instead.
- Files with multiple internal shared strings files are not supported.
- Files with multiple internal styles definition files are not supported.
- Fractions are only partially supported. The results delivered by the reader might be slightly off from the original input.

### Licensing
All the code in this library is licensed under the MIT license as included in the [LICENSE.md](LICENSE.md) file.