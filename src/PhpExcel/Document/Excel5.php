<?php
namespace PhpOfficeWrapper\PhpExcel\Document;

use PHPExcel;
use PHPExcel_Writer_Excel5;
use PhpOfficeWrapper\PhpExcel\Document;
/**
 * Class Excel5
 *
 * @link https://github.com/fanatov37/PhpOfficeWrapper.git for the canonical source repository
 * @copyright Copyright (c)
 * @license MIT
 * @author VladFanatov
 * @package PhpOfficeWrapper\PhpExcel\Document
 */
abstract class Excel5 extends Document
{
    /**
     * (non-PHPDoc)
     *
     * @see Document::createWriter()
     */
    final protected function createWriter(PHPExcel $phpExcel)
    {
        return new PHPExcel_Writer_Excel5($phpExcel);
    }
    /**
     * (non-PHPDoc)
     *
     * @see Document::getContentType()
     */
    final protected function getContentType() : string
    {
        return 'application/vnd.ms-excel';
    }
    /**
     * @return string
     */
    protected function getFileExtension() : string
    {
        return 'xls';
    }
}