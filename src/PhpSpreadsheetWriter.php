<?php
namespace Germania\PhpExcelWrapper;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class PhpSpreadsheetWriter implements SpreadsheetWriter
{

    /**
     * @var string
     */
    public $base_path;

    /**
     * @var string
     */
    public $filename_join_separator = "-";

    /**
     * @var string
     */
    public $filename_extension = "xlsx";



    public function __construct( string $base_path = null)
    {
        $this->setBasePath($base_path ?: getcwd());
    }





    public function setBasePath( string $base_path)
    {
        $this->base_path = $base_path ?: getcwd();

        if (!is_dir($this->base_path)):
            throw new \RuntimeException("Directory not found: " . $base_path);
        endif;

        if (!is_writable($this->base_path)):
            throw new \RuntimeException("Directory not writeable: " . $base_path);
        endif;

        return $this;
    }

    public function getBasePath() : string
    {
        return $this->base_path;
    }



    /**
     * @inheritDoc
     */
    public function __invoke( $spreadsheet_document, $filename ) {
        return $this->write( $spreadsheet_document, $filename ) ?: false;
    }


    /**
     * @inheritDoc
     */
    public function write( $spreadsheet_document, $filename ) : ?string
    {
        if ($spreadsheet_document instanceOf SpreadsheetProvider):
            $spreadsheet_document = $spreadsheet_document->getSpreadsheet();
        elseif (!$spreadsheet_document instanceOf Spreadsheet):
            throw new \InvalidArgumentException("Spreadsheet or SpreadsheetProvider expected");
        endif;

        $writer = new Xlsx( $spreadsheet_document );

        $filename = $this->generateFilename( $filename );

        $output_path = join(\DIRECTORY_SEPARATOR, [
            $this->getBasePath(),
            $filename
        ]);

        $writer->save( $output_path );
        return is_file( $output_path ) ? $filename : null;
    }



    /**
     * @param  strinv|string[] $filename
     * @return string
     */
    protected function generateFilename( $filename ) : string
    {
        if (is_array($filename)) {
            $filename = join( $this->filename_join_separator, array_filter($filename));
        }

        if (!$ext = pathinfo($filename, \PATHINFO_EXTENSION)
        or $ext != $this->filename_extension) {
            $filename .= ".{$this->filename_extension}";
        }

       return $filename;
    }


}
