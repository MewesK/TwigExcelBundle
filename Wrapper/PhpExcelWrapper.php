<?php

namespace MewesK\TwigExcelBundle\Wrapper;
use Twig_Environment;

/**
 * Class PhpExcelWrapper
 *
 * @package MewesK\TwigExcelBundle\Wrapper
 */
class PhpExcelWrapper
{
    /**
     * @var XlsDocumentWrapper
     */
    private $documentWrapper;
    /**
     * @var XlsSheetWrapper
     */
    private $sheetWrapper;
    /**
     * @var XlsRowWrapper
     */
    private $rowWrapper;
    /**
     * @var XlsCellWrapper
     */
    private $cellWrapper;
    /**
     * @var XlsHeaderFooterWrapper
     */
    private $headerFooterWrapper;
    /**
     * @var XlsDrawingWrapper
     */
    private $drawingWrapper;

    /**
     * @var int
     */
    private $cellIndex;
    /**
     * @var int
     */
    private $rowIndex;

    /**
     * PhpExcelWrapper constructor.
     * @param array $context
     * @param Twig_Environment $environment
     */
    public function __construct(array $context = [], Twig_Environment $environment)
    {
        $this->documentWrapper = new XlsDocumentWrapper($context, $environment);
        $this->sheetWrapper = new XlsSheetWrapper($context, $environment, $this->documentWrapper);
        $this->rowWrapper = new XlsRowWrapper($context, $environment, $this->sheetWrapper);
        $this->cellWrapper = new XlsCellWrapper($context, $environment, $this->sheetWrapper);
        $this->headerFooterWrapper = new XlsHeaderFooterWrapper($context, $environment, $this->sheetWrapper);
        $this->drawingWrapper = new XlsDrawingWrapper($context, $environment, $this->sheetWrapper, $this->headerFooterWrapper);
    }

    //
    // Tags
    //

    /**
     * @param null|array $properties
     *
     * @throws \PHPExcel_Exception
     */
    public function startDocument(array $properties = null)
    {
        $this->documentWrapper->start($properties);
    }

    /**
     * @throws \PHPExcel_Reader_Exception
     */
    public function endDocument()
    {
        $this->documentWrapper->end();
    }

    /**
     * @param string $index
     * @param null|array $properties
     *
     * @throws \PHPExcel_Exception
     */
    public function startSheet($index, array $properties = null)
    {
        $this->sheetWrapper->start($index, $properties);
    }

    public function endSheet()
    {
        $this->sheetWrapper->end();
    }

    public function startRow()
    {
        $this->rowWrapper->start($this->rowIndex);
    }

    public function endRow()
    {
        $this->rowWrapper->end();
    }

    /**
     * @param null|mixed $value
     * @param null|array $properties
     *
     * @throws \PHPExcel_Exception
     */
    public function startCell($value = null, array $properties = null)
    {
        $this->cellWrapper->start($this->cellIndex, $value, $properties);
    }

    public function endCell()
    {
        $this->cellWrapper->end();
    }

    /**
     * @param string $type
     * @param null|array $properties
     */
    public function startHeaderFooter($type, array $properties = null)
    {
        $this->headerFooterWrapper->start($type, $properties);
    }

    public function endHeaderFooter()
    {
        $this->headerFooterWrapper->end();
    }

    /**
     * @param null|string $type
     * @param null|array $properties
     */
    public function startAlignment($type = null, array $properties = null)
    {
        $this->headerFooterWrapper->startAlignment($type, $properties);
    }

    /**
     * @param null|string $value
     */
    public function endAlignment($value = null)
    {
        $this->headerFooterWrapper->endAlignment($value);
    }

    /**
     * @param string $path
     * @param array $properties
     *
     * @throws \PHPExcel_Exception
     */
    public function startDrawing($path, array $properties = null)
    {
        $this->drawingWrapper->start($path, $properties);
    }

    public function endDrawing()
    {
        $this->drawingWrapper->end();
    }

    // Getter / Setter

    /**
     * @return int
     */
    public function getCellIndex()
    {
        return $this->cellIndex;
    }

    /**
     * @param int $cellIndex
     */
    public function setCellIndex($cellIndex)
    {
        $this->cellIndex = $cellIndex;
    }

    /**
     * @return int
     */
    public function getRowIndex()
    {
        return $this->rowIndex;
    }

    /**
     * @param int $rowIndex
     */
    public function setRowIndex($rowIndex)
    {
        $this->rowIndex = $rowIndex;
    }
}
