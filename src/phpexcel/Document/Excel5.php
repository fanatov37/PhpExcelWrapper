<?php

namespace PhpOfficeWrapper\PhpExcel\Document;

use PHPExcel;
use PHPExcel_Writer_Excel5;
use PhpOfficeWrapper\PhpExcel\Document;


abstract class Excel5 extends Document
{
    /**
     * (non-PHPDoc)
     *
     * @see Document::_createWriter()
     */
    protected function _createWriter(PHPExcel $phpExcel)
    {
        return new PHPExcel_Writer_Excel5($phpExcel);
    }

    /**
     * (non-PHPDoc)
     *
     * @see Document::_getContentType()
     */
    protected function _getContentType() : string
    {
        return 'application/vnd.ms-excel';
    }

    /**
     * @return string
     */
    protected function _getFileExtension() : string
    {
        return 'xls';
    }
}