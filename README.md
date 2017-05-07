# Piexl - Php Import EXport Library(for Excel)

Exporting PHP to Excel or Importing Excel to PHP. Excel Library for generate Excel File or for load Excel File.

Installation
------------

The preferred way to install this extension is through [composer](http://getcomposer.org/download/).

Either run

```
php composer.phar require --prefer-dist atishamte/piexl
OR
composer require atishamte/piexl
```

or add

```
"atishamte/piexl": "*"
```

to the require section of your `composer.json` file.


Usage 
-----
### Importing Data

Import file excel and return into an array.


```php
<?php
$config = [
	'fileNames'            => 'SampleData.xlsx', // String or Array of file names
	'setFirstRecordAsKeys' => true,              // Set column keys as a index of each element
	'setIndexSheetByName'  => true,              // Set worksheet name as a index of sheet data array
	'getOnlySheet'         => 'worksheetname',   // Get data of particular worksheet
	'getOnlySheetNames'    => true,              // Get only worksheet names as a array
	'getSheetRangeInfo'    => true,              // Get range of filled cells in worksheets
];
$ExcelData = \atishamte\Piexl\Excel::import($config);
```
