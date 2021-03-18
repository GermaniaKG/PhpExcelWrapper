<?php
namespace Germania\PhpExcelWrapper;

use Germania\PhpExcelWrapper\SpreadsheetProvider;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

interface SpreadsheetWriter
{

    /**
     * @param  SpreadsheetProvider|Spreadsheet $document
     * @param  string|string[] $filename
     * @return string|FALSE
     */
    public function __invoke( $document, $filename );


    /**
     * @param  SpreadsheetProvider|Spreadsheet $document
     * @param  string|string[] $filename
     * @return string|null
     */
    public function write( $spreadsheet_document, $filename ) : ?string;
}
