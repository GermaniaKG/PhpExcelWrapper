<?php
namespace Germania\PhpExcelWrapper;

use PhpOffice\PhpSpreadsheet\Spreadsheet;

class PhpSpreadsheetFactory implements PhpSpreadsheetFactoryInterface
{
    public function create() : Spreadsheet
    {
        return new Spreadsheet();
    }
}
