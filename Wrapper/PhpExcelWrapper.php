<?php

namespace MewesK\TwigExcelBundle\Wrapper;

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
     * @param array $context
     */
    public function __construct(array $context = [])
    {
        $this->documentWrapper = new XlsDocumentWrapper($context);
        $this->sheetWrapper = new XlsSheetWrapper($context, $this->documentWrapper);
        $this->rowWrapper = new XlsRowWrapper($context, $this->sheetWrapper);
        $this->cellWrapper = new XlsCellWrapper($context, $this->sheetWrapper);
        $this->headerFooterWrapper = new XlsHeaderFooterWrapper($context, $this->sheetWrapper);
        $this->drawingWrapper = new XlsDrawingWrapper($context, $this->sheetWrapper, $this->headerFooterWrapper);
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

    /**
     * @param null|int $index
     */
    public function startRow($index = null)
    {
        $this->rowWrapper->start($index);
    }

    public function endRow()
    {
        $this->rowWrapper->end();
    }

    /**
     * @param null|int $index
     * @param null|mixed $value
     * @param null|array $properties
     *
     * @throws \PHPExcel_Exception
     */
    public function startCell($index = null, $value = null, array $properties = null)
    {
        $this->cellWrapper->start($index, $value, $properties);
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
}
