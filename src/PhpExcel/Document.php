<?php
namespace PhpOfficeWrapper\Phpexcel;

use PHPExcel;
use Exception;
use PHPExcel_Writer_IWriter;
use DateTime;
/**
 * Class Document
 *
 * @link https://github.com/fanatov37/PhpOfficeWrapper.git for the canonical source repository
 * @copyright Copyright (c)
 * @license MIT
 * @author VladFanatov
 * @package PhpOfficeWrapper\Phpexcel
 */
abstract class Document
{
    public const MAX_STRLENGTH_FOR_TITLE = 31;
    public const INVALID_TITLE_CHARACTERS = ['*', ':', '/', '\\', '?', '[', ']'];
    /**
     * @var \PHPExcel
     */
    protected $phpExcel;
    /**
     * @var PHPExcel_Writer_IWriter
     */
    protected $writer;
    /**
     * @var boolean
     */
    protected $isRendered = false;
    /**
     * @return string
     */
    abstract protected function getContentType() : string;
    /**
     * @return PHPExcel_Writer_IWriter
     */
    abstract protected function createWriter(PHPExcel $phpExcel);
    /**
     * (non-PHPDoc)
     */
    abstract protected function renderColumn();
    /**
     * (non-PHPDoc)
     */
    abstract protected function renderValue();
    /**
     * (non-PHPDoc)
     */
    abstract protected function renderStyle();
    /**
     * @return string
     */
    abstract protected function getFileExtension() : string;
    /**
     * @return string
     */
    abstract protected function getFilename() : string;
    /**
     * @return PHPExcel
     */
    private function createPHPExcel() : PHPExcel
    {
        return new PHPExcel();
    }
    /**
     * @return PHPExcel
     *
     * @throws Exception
     */
    final public function getPHPExcel() : PHPExcel
    {
        if ($this->phpExcel === null) {
            $phpExcel = $this->createPHPExcel();

            if (!($phpExcel instanceof PHPExcel)) {
                throw new Exception('phpExcel does not instanceof PHPExcel');
            }

            $this->phpExcel = $phpExcel;
        }
        return $this->phpExcel;
    }
    /**
     * (non-PHPDoc)
     */
    protected function render()
    {
        $this->renderValue();
        $this->renderColumn();
        $this->renderStyle();
    }
    /**
     * @return self
     */
    private function _render() : self
    {
        if (!$this->isRendered) {
            $this->render();
            $this->isRendered = true;
        }

        return $this;
    }
    /**
     * @return PHPExcel_Writer_IWriter
     *
     * @throws Exception
     */
    private function getWriter() : PHPExcel_Writer_IWriter
    {
        if ($this->writer === null) {
            $writer = $this->createWriter($this->getPHPExcel());

            if (!($writer instanceof PHPExcel_Writer_IWriter)) {
                throw new Exception('writer not implements PHPExcel_Writer_IWriter');
            }

            $this->writer = $writer;
        }
        return $this->writer;
    }
    /**
     * @param $fileName
     *
     * @return Document
     * @throws Exception
     * @throws \PHPExcel_Writer_Exception
     */
    private function save($fileName) : self
    {
        $this->_render()->getWriter()->save($fileName);

        return $this;
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
     * @throws Exception
     * @throws \PHPExcel_Writer_Exception
     */
    final public function download()
    {
        /**
         * todo need use ZF Response header for it.
         * todo example https://github.com/Golbut/zf2Traits/blob/e4f1c0d6fdf126230595c6171ab2b463298088df/zf2Traits/ForceDownloadProvider.php
         */

        header("Content-type: {$this->getContentType()}");
        header('Content-Description: File Transfer');
        header("Content-Disposition: attachment; filename={$this->getFilename()}.{$this->getFileExtension()}");
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
                $this->getFilename(),
                $this->getFileExtension()
        );
        $this->save($tmpFilePath);

        $file = [
            'tmp_name' => $tmpFilePath,
            'name' => $this->getFilename() . '.' . $this->getFileExtension(),
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
            $this->getContentType(),
            base64_encode($xlsData)
        );
    }
}