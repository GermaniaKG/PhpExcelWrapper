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

## Unit tests and development

1. Copy `phpunit.xml.dist` to `phpunit.xml` 
2. Run [PhpUnit](https://phpunit.de/) like this:

```bash
$ composer test
# or
$ vendor/bin/phpunit
```

And there's more in the `scripts` section of **composer.json**.

