### v.0.10.1  2021-10-26
- Fixed an issue that made returnUnformatted overrule all configuration options for date/time values.

### v.0.10.0  2021-10-22
Breaking changes:
- next() no longer returns the current row. Use current() instead.
- SkipEmptyCells needs to be supplied as a ReaderSkipConfiguration constant now.

Non-breaking changes:
- New configuration option "SkipEmptyRows".
  Use it to exclude either all empty rows or all empty rows at the end of the document from the output. 
  Use ReaderSkipConfiguration values to configure it.
- Configuration option "SkipEmptyCells" can now be configured to only skip trailing empty cells.
- Added support for scientific notation format.
- Fraction formatting support was enhanced.
- Fixed: "General" format does not output values as decimal, if they are stored using scientific notation internally.
- Fixed: Assorted edge cases in number formatting.
- Added notes to documentation of "ReturnUnformatted" and "ReturnPercentageDecimal" about possible gotchas.
- Internal refactorings.

### v.0.9.0  2021-07-20
Breaking changes:
- Reader configuration options must now be supplied to the Reader constructor via a ReaderConfiguration instance.
  Supplying configuration options via an array is no longer supported.
- When the "ReturnUnformatted" option is set, percentage values are now returned as strings instead of numbers.
  This aligns their behavior with that of other values.
- setDecimalSeparator() and setThousandsSeparator() methods have been removed, as they no longer had any function.
- Forced date/time format '' (empty string) gets interpreted correctly now.

Non-breaking changes:
- New configuration option "ReturnPercentageDecimal".
  When set to true, percentage values will be returned using their technical, internal representation ('50%' => '0.5')
  rather than how they are displayed within a document ('50%' => '50').
- Remove unnecessary restriction of custom formats to predetermined formats from the official specification documents.
- SharedStringsConfiguration calls can now be chained.
- Fix potential resource leaks caused by not closing reader instances.
- Update README.md to reflect the current code state.

### v.0.8.2  2021-07-14
- Added support for empty rows with attributes (or: self-closing row tags).
- Minor improvement of test handling.

### v.0.8.1  2020-10-05
- Added support for multi-range row span values, fixing issues caused by sheets that use them.

### v.0.8.0  2020-03-09
Breaking changes:
- Public-facing method "setCurrencyCode" has been removed, as the currency_code value had no effect to begin with.

Non-breaking changes:
- New configuration option "ReturnUnformatted". If set to true, cell values will be returned without number formatting applied. (Note: Date/Time values are still controlled by the "ReturnDateTimeObjects" option.)
- Number format parsing has been improved. The reader is now capable of parsing more complex number formats.
- General format now outputs cell values as-is, instead of attempting to cast them to a float.

### v.0.7.7  2019-10-08
- Fixed issues regarding negative date/time values, causing very early date definitions to lead to unexpected errors.
- Fixed number formatting not being applied in all expected cases.

### v.0.7.6  2019-09-11
- Fixed a bug that caused empty shared strings to be treated incorrectly under certain conditions.

### v.0.7.5  2019-07-23
- Fixed a bug that caused cell formats making use of currency strings and language ids to break.
- Fixed a "continue in switch" warning in PHP 7.3.

### v.0.7.4  2019-05-07
- Added the option to use alphabetical column names (A, B, AA, ZX) instead of numeric indexes in returned row contents, using the parameter "OutputColumnNames".
- Fixed a bug that caused leading zeros in text cell content to get removed if the cell was set to text via an apostrophe prefix.

### v.0.7.3  2019-03-25
- Fixed an issue that prevented empty rows from being properly output in all appropriate cases.

### v.0.7.2  2019-03-14
- Fixed an issue that caused format parsing to cease working for some files.

### v.0.7.1  2019-02-27
- New configuration parameters to control automatic re-formatting of found Date/Time values: forceDateFormat, forceTimeFormat, forceDateTimeFormat
- Improved handling of potential errors when working with subdirectories of the configured temporary directory
- Fixed composer.json lacking ext-xmlreader requirement

### v.0.7.0  2019-02-05
- Improved support for different XLSX file generators:
  - Improved awareness of XML namespaces.
- Improved support for newer OOXML editions:
  - Namespace URIs from newer versions of the OOXML standard are now recognized and handled accordingly.
- Dropped requirement for SimpleXMLElement.
- Minor improvements in handling used document resources.

### v.0.6.3  2018-12-04
- Bugfix: Check if current row, that is to be read, is also the one which the read() function takes, return empty row if not.

### v.0.6.2  2018-11-20
- Bugfix: differentiate between internal sheet ID and positioning ordering of the sheet within the document

### v.0.6.1  2018-05-16
- Removed unneccessary test files.
- Minor code quality improvements.

### v.0.6.0  2018-05-01
- Added option 'SkipEmptyCells' in order to consider or not possible empty values in cells. 
- Added option 'CustomFormats' to define and overwrite format values.
- Ensure deletion of temporary files after run.
- Fix: MAP Toolkit xlsx files can be parsed.
- PHP 7 compliance.
- Allow configuration of locale based values.
- Include PHPUnit and tests for iterator, file location, shared strings, sheet handling, namespaces and temporary directories handling. 
- Major structural refactoring and appliance of PSR1, PSR2 and PSR4 (namespace directory structure)

### v.0.5.11  2015-04-30

- Added a special case for cells formatted as text in XLSX. Previously leading zeros would get truncated if a text cell contained only numbers.

### v.0.5.10  2015-04-18

- Implemented SeekableIterator. Thanks to [paales](https://github.com/paales) for suggestion ([Issue #54](https://github.com/nuovo/spreadsheet-reader/issues/54) and [Pull request #55](https://github.com/nuovo/spreadsheet-reader/pull/55)).
- Fixed a bug in CSV and ODS reading where reading position 0 multiple times in a row would result in internal pointer being advanced and reading the next line. (E.g. reading row #0 three times would result in rows #0, #1, and #2.). This could have happened on multiple calls to `current()` while in #0 position, or calls to `seek(0)` and `current()`.

### v.0.5.9  2015-04-18

- [Pull request #85](https://github.com/nuovo/spreadsheet-reader/pull/85): Fixed an index check. (Thanks to [pa-m](https://github.com/pa-m)).

### v.0.5.8  2015-01-31

- [Issue #50](https://github.com/nuovo/spreadsheet-reader/issues/50): Fixed an XLSX rewind issue. (Thanks to [osuwariboy](https://github.com/osuwariboy))
- [Issue #52](https://github.com/nuovo/spreadsheet-reader/issues/52), [#53](https://github.com/nuovo/spreadsheet-reader/issues/53): Apache POI compatibility for XLSX. (Thanks to [dimapashkov](https://github.com/dimapashkov))
- [Issue #61](https://github.com/nuovo/spreadsheet-reader/issues/61): Autoload fix in the main class. (Thanks to [i-bash](https://github.com/i-bash))
- [Issue #60](https://github.com/nuovo/spreadsheet-reader/issues/60), [#69](https://github.com/nuovo/spreadsheet-reader/issues/69), [#72](https://github.com/nuovo/spreadsheet-reader/issues/72): Fixed an issue where XLSX changeSheet may not work. (Thanks to [jtresponse](https://github.com/jtresponse), [osuwariboy](https://github.com/osuwariboy))
- [Issue #70](https://github.com/nuovo/spreadsheet-reader/issues/70): Added a check for constructor parameter correctness.


### v.0.5.7  2013-10-29

- Attempt to replicate Excel's "General" format in XLSX files that is applied to otherwise unformatted cells.
Currently only decimal number values are converted to PHP's floats.

### v.0.5.6  2013-09-04

- Fix for formulas being returned along with values in XLSX files. (Thanks to [marktag](https://github.com/marktag))

### v.0.5.5  2013-08-23

- Fix for macro sheets appearing when parsing XLS files. (Thanks to [osuwariboy](https://github.com/osuwariboy))

### v.0.5.4  2013-08-22

- Fix for a PHP warning that occurs with completely empty sheets in XLS files.
- XLSM (macro-enabled XLSX) files are recognized and read, too.
- composer.json file is added to the repository (thanks to [matej116](https://github.com/matej116))

### v.0.5.3  2013-08-12

- Fix for repeated columns in ODS files not reading correctly (thanks to [etfb](https://github.com/etfb))
- Fix for filename extension reading (Thanks to [osuwariboy](https://github.com/osuwariboy))

### v.0.5.2  2013-06-28

- A fix for the case when row count wasn't read correctly from the sheet in a XLS file.

### v.0.5.1  2013-06-27

- Fixed file type choice when using mime-types (previously there were problems with  
XLSX and ODS mime-types) (Thanks to [incratec](https://github.com/incratec))

- Fixed an error in XLSX iterator where `current()` would advance the iterator forward  
with each call. (Thanks to [osuwariboy](https://github.com/osuwariboy))

### v.0.5.0  2013-06-17

- Multiple sheet reading is now supported:
	- The `getSheets()` method lets you retrieve a list of all sheets present in the file.
	- `changeSheet($Index)` method changes the sheet in the reader to the one specified.

- Previously temporary files that were extracted, were deleted after the SpreadsheetReader  
was destroyed but the empty directories remained. Now those are cleaned up as well.  

### v.0.4.3  2013-06-14

- Bugfix for shared string caching in XLSX files. When the shared string count was larger  
than the caching limit, instead of them being read from file, empty strings were returned.  

### v.0.4.2  2013-06-02

- XLS file reading relies on the external Spreadsheet_Excel_Reader class which, by default,  
reads additional information about cells like fonts, styles, etc. Now that is disabled  
to save some memory since the style data is unnecessary anyway.  
(Thanks to [ChALkeR](https://github.com/ChALkeR) for the tip.)

Martins Pilsetnieks  <pilsetnieks@gmail.com>