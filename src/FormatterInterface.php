<?php
namespace Germania\PhpExcelWrapper;

interface FormatterInterface
{

    /**
     * @return array
     */
    public function getStyles();

    /**
     * @return array
     */
    public function getHeadlines();
}
