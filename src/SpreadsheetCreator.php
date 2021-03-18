<?php
namespace Germania\PhpExcelWrapper;

interface SpreadsheetCreator extends SpreadsheetProvider
{
    public function getActiveSheet();
}
