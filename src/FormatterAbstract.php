<?php
namespace Germania\PhpExcelWrapper;

abstract class FormatterAbstract implements FormatterInterface
{
    public $styles = array();
    public $headlines = array();


    /**
     * @implements FormatterInterface
     */
    public function getStyles()
    {
        return $this->styles;
    }


    /**
     * @implements FormatterInterface
     */
    public function getHeadlines()
    {
        return $this->headlines;
    }
}
