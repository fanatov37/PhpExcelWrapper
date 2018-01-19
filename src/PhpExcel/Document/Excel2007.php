<?php

namespace PhpOfficeWrapper\PhpExcel\Document;

use PHPExcel;
use PHPExcel_IOFactory;
use PhpOfficeWrapper\PhpExcel\Document;


abstract class Excel2007 extends Document
{
    /**
     * @param PHPExcel $phpExcel
     *
     * @return \PHPExcel_Writer_IWriter
     * @throws \PHPExcel_Reader_Exception
     */
    final protected function createWriter(PHPExcel $phpExcel)
    {
        return PHPExcel_IOFactory::createWriter($phpExcel, 'Excel2007');
    }
    /**
     * (non-PHPDoc)
     *
     * @see Document::getContentType()
     */
    final protected function getContentType() : string
    {
        return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    }

    /**
     * @return string
     */
    protected function getFileExtension() : string
    {
        return 'xlsx';
    }
}