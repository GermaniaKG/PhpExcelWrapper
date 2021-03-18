<?php
namespace Germania\PhpExcelWrapper;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment as PHPExcel_Style_Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill as PHPExcel_Style_Fill;


class PhpSpreadsheetCreator implements SpreadsheetCreator {

    /**
     * @var Spreadsheet
     */
    public $document;
    public $sheet;

    public $headers_start_cell = 'A1';
    public $data_start_cell = 'A2';

    public $document_properties = [
        'creator' => null,
        'last_modified_by' => null,
        'title' => null,
        'subject' => null,
        'description' => null,
        'keywords' => null,
        'category' => null
    ];


    public $headers_styles = array(
        'font'      => ['bold' => true],
        'alignment' => [ 'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT ],
        'fill' => [
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => [ 'argb' => 'FFEEEEEE' ]
        ]
    );

    public $align_left = array(
        'alignment' => [ 'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT ]
    );

    public $align_right = array(
        'alignment' => [ 'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT ]
    );




    public function __construct( Spreadsheet $document = null )
    {
        $this->setSpreadsheet($document ?: (new PhpSpreadsheetFactory)->create());
    }


    public function getSpreadsheet() {
        return $this->document;
    }

    public function setSpreadsheet( Spreadsheet $document ) {
        $this->document = $document;
        $this->sheet = $document->getActiveSheet();
        return $this;
    }



    public function getActiveSheet() {
        return $this->getSpreadsheet()->getActiveSheet();
    }



    public function setStyle( $where, $style ) {
        $this->getActiveSheet()->getStyle( $where )->applyFromArray( $style );
        return $this;
    }

    public function setStyles( $styles ) {
        if ($styles instanceOf FormatterInterface):
            $styles = $styles->getStyles();
        elseif (!is_array($styles)):
            throw new \InvalidArgumentException( "Array or FormatterInterface expexted");
        endif;


        foreach($styles as $where => $style_array):
            $this->setStyle($where, $style_array);
        endforeach;


        return $this;

    }



    public function makeHeader( $where )
    {
        $this->setStyle( $where, $this->headers_styles );
        return $this;
    }

    public function makeAlignLeft( $where )
    {
        $this->setStyle( $where, $this->align_left );
        return $this;
    }




    public function setProperties( array $user_props )
    {
        $props = array_merge( $this->document_properties, $user_props);

        $this->document->getProperties()
           ->setCreator( $props['creator'] )
           ->setLastModifiedBy( $props['last_modified_by'] ?: $props['creator'] )
           ->setTitle( $props['title'] )
           ->setSubject( $props['subject'] )
           ->setDescription( $props['description'] )
           ->setKeywords( $props['keywords'] )
           ->setCategory( $props['category'] );

        return $this;
    }





    public function setWorksheetTitle( $title )
    {
        if (mb_strlen($title) > 31) {
            $title = mb_substr($title, 0,31);
        }

        // Replace disallowed characters
        $title = preg_replace('![\/\[\]\*\?\\\]!', "-", $title);

        $this->sheet->setTitle( $title );
        return $this;
    }



    public function setHeaders( $headers, $autosize = true, $autofilter = true )
    {
        if ($headers instanceOf FormatterInterface):
            $headers = $headers->getHeadlines();
        elseif (!is_array($headers)):
            throw new \InvalidArgumentException( "Array or FormatterInterface expexted");
        endif;

        $this->sheet->fromArray( $headers, ' ', $this->headers_start_cell);


        if ($autofilter):
            $this->sheet->setAutoFilter($this->sheet->calculateWorksheetDimension());
        endif;

        if ($autosize):
            for ($i = 0; $i < count($headers); $i++):
                if ($i < 26):
                    $this->sheet->getColumnDimension(chr(65 + $i))->setAutoSize(true);
                endif;
            endfor;
        endif;

        return $this;
    }



    public function setData( $rows )
    {
        if (!is_array($rows)
        and !$rows instanceOf \Traversable) {
            throw new \InvalidArgumentException("Array or Traversable instance expected.");
        }

        if ($rows instanceOf \Traversable) {
            $rows = iterator_to_array( $rows );
        }

        $this->sheet->fromArray( $rows, ' ', $this->data_start_cell);
        return $this;
    }

}
