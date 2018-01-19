<?php

namespace PhpOfficeWrapper\PhpExcel;

use PHPExcel;
use Exception;
use PHPExcel_Writer_IWriter;
use DateTime;


abstract class Document
{
    public const MAX_STRLENGTH_FOR_TITLE = 31;

    /**
     * Invalid characters in sheet title
     *
     * @see PHPExcel_Worksheet::$_invalidCharacters
     *
     * @var array
     */
    public static $invalidCharacters = ['*', ':', '/', '\\', '?', '[', ']'];

    /**
     * @var \PHPExcel
     */
    protected $_phpExcel;

    /**
     * @var PHPExcel_Writer_IWriter
     */
    protected $_writer;

    /**
     * @var boolean
     */
    protected $_isRendered = false;

    /**
     * (non-PHPDoc)
     */
    abstract protected function _renderColumn();

    /**
     * (non-PHPDoc)
     */
    abstract protected function _renderValue();

    /**
     * (non-PHPDoc)
     */
    abstract protected function _renderStyle();

    /**
     * @return string
     */
    abstract protected function _getContentType() : string;

    /**
     * @return PHPExcel_Writer_IWriter
     */
    abstract protected function _createWriter(PHPExcel $phpExcel);

    /**
     * @return string
     */
    abstract protected function _getFileExtension() : string;

    /**
     * @return string
     */
    abstract protected function _getFilename() : string;

    /**
     * (non-PHPDoc)
     */
    protected function _render()
    {
        $this->_renderColumn();
        $this->_renderValue();
        $this->_renderStyle();
    }

    /**
     * @return PHPExcel
     *
     * @throws Exception
     */
    public function getPHPExcel() : PHPExcel
    {
        if ($this->_phpExcel === null) {
            $phpExcel = $this->_createPHPExcel();

            if (!($phpExcel instanceof PHPExcel)) {
                throw new Exception('phpExcel does not instanceof PHPExcel');
            }

            $this->_phpExcel = $phpExcel;
        }
        return $this->_phpExcel;
    }

    /**
     * @return PHPExcel
     */
    protected function _createPHPExcel() : PHPExcel
    {
        return new PHPExcel();
    }

    /**
     * @param $creatorName
     *
     * @return Document
     * @throws Exception
     */
    public function setCreator($creatorName) : self
    {
        $this->getPHPExcel()->getProperties()->setCreator($creatorName);

        return $this;
    }

    /**
     * @param DateTime|NULL $time
     *
     * @return Document
     * @throws Exception
     */
    public function setCreated(DateTime $time = null) : self
    {
        if ($time === null) {
            $time = new DateTime();
        }

        $this->getPHPExcel()->getProperties()->setCreated($time->getTimestamp());

        return $this;
    }

    /**
     * @return self
     */
    public function render() : self
    {
        if (!$this->_isRendered) {
            $this->_render();
            $this->_isRendered = true;
        }

        return $this;
    }

    /**
     * @return PHPExcel_Writer_IWriter
     *
     * @throws Exception
     */
    final public function getWriter() : PHPExcel_Writer_IWriter
    {
        if ($this->_writer === null) {
            $writer = $this->_createWriter($this->getPHPExcel());

            if (!($writer instanceof PHPExcel_Writer_IWriter)) {
                throw new Exception('writer not implements PHPExcel_Writer_IWriter');
            }

            $this->_writer = $writer;
        }
        return $this->_writer;
    }

    /**
     * @param $fileName
     *
     * @return Document
     * @throws Exception
     * @throws \PHPExcel_Writer_Exception
     */
    public function save($fileName) : self
    {
        $this->render()->getWriter()->save($fileName);

        return $this;
    }

    /**
     * @throws Exception
     */
    final public function download()
    {
        /**
         * todo need use ZF Response header for it.
         * todo example https://github.com/Golbut/zf2Traits/blob/e4f1c0d6fdf126230595c6171ab2b463298088df/zf2Traits/ForceDownloadProvider.php
         */

        header("Content-type: {$this->_getContentType()}");
        header('Content-Description: File Transfer');
        header("Content-Disposition: attachment; filename={$this->_getFilename()}.{$this->_getFileExtension()}");
        header('Expires: Mon, 1 Apr 1974 05:00:00 GMT');
        header('Last-Modified: ' . gmdate('D,d M YH:i:s') . ' GMT');
        header('Cache-Control: no-cache, must-revalidate');
        header('Pragma: no-cache');

        $this->save('php://output');
        die();
    }

    /**
     * currently need refactoring
     *
     * @return string
     * @throws Exception
     * @throws \PHPExcel_Writer_Exception
     * @deprecated
     */
    final public function outputToFile() : string
    {
        $tmpFilePath = sprintf('%s/%s.%s',
                sys_get_temp_dir(),
                $this->_getFilename(),
                $this->_getFileExtension()
        );
        $this->save($tmpFilePath);

        $file = [
            'tmp_name' => $tmpFilePath,
            'name' => $this->_getFilename() . '.' . $this->_getFileExtension(),
            'size' => true
        ];

        return TmpFileStorage::addFile($file);
    }

    /**
     * @return string
     * @throws Exception
     * @throws \PHPExcel_Writer_Exception
     */
    final public function outputToBase64() : string
    {
        ob_start();
        $this->save("php://output");
        $xlsData = ob_get_contents();
        ob_end_clean();

        return sprintf('data:%s;base64,%s',
            $this->_getContentType(),
            base64_encode($xlsData)
        );
    }
}