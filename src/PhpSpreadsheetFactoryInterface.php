<?php
namespace Germania\PhpExcelWrapper;

use PhpOffice\PhpSpreadsheet\Spreadsheet;

interface PhpSpreadsheetFactoryInterface
{
    public function create() : Spreadsheet;
}
