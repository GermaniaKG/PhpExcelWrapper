<?php
namespace Germania\PhpExcelWrapper;

use PhpOffice\PhpSpreadsheet\Spreadsheet;

interface SpreadsheetProvider
{

    /**
     * @return \PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    public function getSpreadsheet();

}
