<img src="https://static.germania-kg.com/logos/ga-logo-2016-web.svgz" width="250px">

<h1 align="center">Germania KG · PhpExcelWrapper</h1>

This package is for outsourcing our PhpSpreadsheet stuff.

## Installation

```bash
$ composer require germania-kg/phpexcel-wrapper
```

## Deprecations

- Namespace **Germania\PhpExcelWrapper** will change to *PhpSpreadsheetWriter* with next major release
- Class **Germania\PhpExcelWrapper\PhpExcelWriter:** use *PhpSpreadsheetWriter* instead
- Class **Germania\PhpExcelWrapper\PhpExcelCreator:** use *PhpSpreadsheetCreator* instead

## Interfaces

**PhpOffice\PhpSpreadsheet\SpreadsheetProvider**

```php
// Returns \PhpOffice\PhpSpreadsheet\Spreadsheet
public getSpreadsheet();
```

**PhpOffice\PhpSpreadsheet\SpreadsheetCreator** extends *SpreadsheetProvider* and additionally provides:

```php
public getActiveSheet();
```

**PhpOffice\PhpSpreadsheet\SpreadsheetWriter**

```php
public function write( $spreadsheet_document, $filename ) : ?string;
```



## Usage

Class **Germania\PhpExcelWrapper\PhpSpreadsheetCreator** implements `SpreadsheetCreator`.
Its constrcutor optionally accepts a `PhpOffice\PhpSpreadsheet\Spreadsheet` instance.

```php
$optional = new \PhpOffice\PhpSpreadsheet\Spreadsheet;

$creator = (new Germania\PhpExcelWrapper\PhpSpreadsheetCreator($optional))
->setHeaders(["foo", "bar"])
->setData([
  [ "data_B1", "data_B2" ],
  [ "data_B3", "data_B3" ],  
]);
```

Class **Germania\PhpExcelWrapper\PhpSpreadsheetWriter** implements `SpreadsheetWriter`. 
Its **write** method accepts any `SpreadsheetProvider` or `PhpOffice\PhpSpreadsheet\Spreadsheet` instance and returns the output file path on success, or `null` on error.

```php
$writer = new PhpSpreadsheetWriter( "/path" );
$output = $writer->write($creator, "output.xlsx");
// /path/output.xlsx
```



## Unit tests and development

```bash
$ composer test
```

