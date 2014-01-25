<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Twig;


class PhpExcelWrapper {

    /**
     * @var \PHPExcel
     */
    protected $document;
    /**
     * @var \PHPExcel_Worksheet
     */
    public $sheet = null;
    /**
     * @var int
     */
    protected $row = 1;
    /**
     * @var \PhpExcel_Cell
     */
    protected $cell = null;
    /**
     * @var \PHPExcel_Worksheet_Drawing
     */
    protected $drawing = null;
    
    /**
     * @var array
     */
    protected $documentMappings = [];
    /**
     * @var array
     */
    protected $sheetMappings = [];
    /**
     * @var array
     */
    protected $rowMappings = [];
    /**
     * @var array
     */
    protected $cellMappings = [];
    /**
     * @var array
     */
    protected $drawingMappings = [];
    

    public function __construct() {
        $this->document = new \PHPExcel();
        $this->document->removeSheetByIndex(0);

        $this->initDocumentPropertyMappings();
        $this->initSheetPropertyMappings();
        $this->initRowPropertyMappings();
        $this->initCellPropertyMappings();
        $this->initDrawingPropertyMappings();
    }

    protected function initDocumentPropertyMappings() {
        $this->documentMappings['category'] = function($value) { $this->document->getProperties()->setCategory($value); };
        $this->documentMappings['company'] = function($value) { $this->document->getProperties()->setCompany($value); };
        $this->documentMappings['created'] = function($value) { $this->document->getProperties()->setCreated($value); };
        $this->documentMappings['creator'] = function($value) { $this->document->getProperties()->setCreator($value); };
        $this->documentMappings['defaultStyle'] = function($value) { $this->document->getDefaultStyle()->applyFromArray($value); };
        $this->documentMappings['description'] = function($value) { $this->document->getProperties()->setDescription($value); };
        $this->documentMappings['keywords'] = function($value) { $this->document->getProperties()->setKeywords($value); };
        $this->documentMappings['lastModifiedBy'] = function($value) { $this->document->getProperties()->setLastModifiedBy($value); };
        $this->documentMappings['manager'] = function($value) { $this->document->getProperties()->setManager($value); };
        $this->documentMappings['modified'] = function($value) { $this->document->getProperties()->setModified($value); };
        $this->documentMappings['security']['lockRevision'] = function($value) { $this->document->getSecurity()->setLockRevision($value); };
        $this->documentMappings['security']['lockStructure'] = function($value) { $this->document->getSecurity()->setLockStructure($value); };
        $this->documentMappings['security']['lockWindows'] = function($value) { $this->document->getSecurity()->setLockWindows($value); };
        $this->documentMappings['security']['revisionsPassword'] = function($value) { $this->document->getSecurity()->setRevisionsPassword($value); };
        $this->documentMappings['security']['workbookPassword'] = function($value) { $this->document->getSecurity()->setWorkbookPassword($value); };
        $this->documentMappings['subject'] = function($value) { $this->document->getProperties()->setSubject($value); };
        $this->documentMappings['title'] = function($value) { $this->document->getProperties()->setTitle($value); };
    }

    protected function initSheetPropertyMappings() {
        $this->sheetMappings['columnDimension']['__multi'] = true;
        $this->sheetMappings['columnDimension']['__object'] = function($key) { return $key == 'default' ? $this->sheet->getDefaultColumnDimension() : $this->sheet->getColumnDimension($key); };
        $this->sheetMappings['columnDimension']['autoSize'] = function($key, $value) { $this->sheetMappings['columnDimension']['__object']($key)->setAutoSize($value); };
        $this->sheetMappings['columnDimension']['collapsed'] = function($key, $value) { $this->sheetMappings['columnDimension']['__object']($key)->setCollapsed($value); };
        $this->sheetMappings['columnDimension']['columnIndex'] = function($key, $value) { $this->sheetMappings['columnDimension']['__object']($key)->setColumnIndex($value); };
        $this->sheetMappings['columnDimension']['outlineLevel'] = function($key, $value) { $this->sheetMappings['columnDimension']['__object']($key)->setOutlineLevel($value); };
        $this->sheetMappings['columnDimension']['visible'] = function($key, $value) { $this->sheetMappings['columnDimension']['__object']($key)->setVisible($value); };
        $this->sheetMappings['columnDimension']['width'] = function($key, $value) { $this->sheetMappings['columnDimension']['__object']($key)->setWidth($value); };
        $this->sheetMappings['columnDimension']['xfIndex'] = function($key, $value) { $this->sheetMappings['columnDimension']['__object']($key)->setXfIndex($value); };
        $this->sheetMappings['footer'] = function($value) { $this->sheet->getHeaderFooter()->setOddFooter($value); };
        $this->sheetMappings['header'] = function($value) { $this->sheet->getHeaderFooter()->setOddHeader($value); };
        $this->sheetMappings['pageMargins']['top'] = function($value) { $this->sheet->getPageMargins()->setTop($value); };
        $this->sheetMappings['pageMargins']['bottom'] = function($value) { $this->sheet->getPageMargins()->setBottom($value); };
        $this->sheetMappings['pageMargins']['left'] = function($value) { $this->sheet->getPageMargins()->setLeft($value); };
        $this->sheetMappings['pageMargins']['right'] = function($value) { $this->sheet->getPageMargins()->setRight($value); };
        $this->sheetMappings['pageMargins']['header'] = function($value) { $this->sheet->getPageMargins()->setHeader($value); };
        $this->sheetMappings['pageMargins']['footer'] = function($value) { $this->sheet->getPageMargins()->setFooter($value); };
        $this->sheetMappings['pageSetup']['fitToHeight'] = function($value) { $this->sheet->getPageSetup()->setFitToHeight($value); };
        $this->sheetMappings['pageSetup']['fitToPage'] = function($value) { $this->sheet->getPageSetup()->setFitToPage($value); };
        $this->sheetMappings['pageSetup']['fitToWidth'] = function($value) { $this->sheet->getPageSetup()->setFitToWidth($value); };
        $this->sheetMappings['pageSetup']['horizontalCentered'] = function($value) { $this->sheet->getPageSetup()->setHorizontalCentered($value); };
        $this->sheetMappings['pageSetup']['orientation'] = function($value) { $this->sheet->getPageSetup()->setOrientation($value); };
        $this->sheetMappings['pageSetup']['paperSize'] = function($value) { $this->sheet->getPageSetup()->setPaperSize($value); };
        $this->sheetMappings['pageSetup']['printArea'] = function($value) { $this->sheet->getPageSetup()->setPrintArea($value); };
        $this->sheetMappings['pageSetup']['scale'] = function($value) { $this->sheet->getPageSetup()->setScale($value); };
        $this->sheetMappings['pageSetup']['verticalCentered'] = function($value) { $this->sheet->getPageSetup()->setVerticalCentered($value); };
        $this->sheetMappings['printGridlines'] = function($value) { $this->sheet->setPrintGridlines($value); };
        $this->sheetMappings['protection']['autoFilter'] = function($value) { $this->sheet->getProtection()->setAutoFilter($value); };
        $this->sheetMappings['protection']['deleteColumns'] = function($value) { $this->sheet->getProtection()->setDeleteColumns($value); };
        $this->sheetMappings['protection']['deleteRows'] = function($value) { $this->sheet->getProtection()->setDeleteRows($value); };
        $this->sheetMappings['protection']['formatCells'] = function($value) { $this->sheet->getProtection()->setFormatCells($value); };
        $this->sheetMappings['protection']['formatColumns'] = function($value) { $this->sheet->getProtection()->setFormatColumns($value); };
        $this->sheetMappings['protection']['formatRows'] = function($value) { $this->sheet->getProtection()->setFormatRows($value); };
        $this->sheetMappings['protection']['insertColumns'] = function($value) { $this->sheet->getProtection()->setInsertColumns($value); };
        $this->sheetMappings['protection']['insertHyperlinks'] = function($value) { $this->sheet->getProtection()->setInsertHyperlinks($value); };
        $this->sheetMappings['protection']['insertRows'] = function($value) { $this->sheet->getProtection()->setInsertRows($value); };
        $this->sheetMappings['protection']['objects'] = function($value) { $this->sheet->getProtection()->setObjects($value); };
        $this->sheetMappings['protection']['pivotTables'] = function($value) { $this->sheet->getProtection()->setPivotTables($value); };
        $this->sheetMappings['protection']['scenarios'] = function($value) { $this->sheet->getProtection()->setSelectLockedCells($value); };
        $this->sheetMappings['protection']['selectLockedCells'] = function($value) { $this->sheet->getProtection()->setVerticalCentered($value); };
        $this->sheetMappings['protection']['selectUnlockedCells'] = function($value) { $this->sheet->getProtection()->setSelectUnlockedCells($value); };
        $this->sheetMappings['protection']['sheet'] = function($value) { $this->sheet->getProtection()->setSheet($value); };
        $this->sheetMappings['protection']['sort'] = function($value) { $this->sheet->getProtection()->setSort($value); };
        $this->sheetMappings['rightToLeft'] = function($value) { $this->sheet->setRightToLeft($value); };
        $this->sheetMappings['rowDimension']['__multi'] = true;
        $this->sheetMappings['rowDimension']['__object'] = function($key) { return $key == 'default' ? $this->sheet->getDefaultRowDimension() : $this->sheet->getRowDimension($key); };
        $this->sheetMappings['rowDimension']['collapsed'] = function($key, $value) { $this->sheetMappings['rowDimension']['__object']($key)->setCollapsed($value); };
        $this->sheetMappings['rowDimension']['outlineLevel'] = function($key, $value) { $this->sheetMappings['rowDimension']['__object']($key)->setOutlineLevel($value); };
        $this->sheetMappings['rowDimension']['rowHeight'] = function($key, $value) { $this->sheetMappings['rowDimension']['__object']($key)->setRowHeight($value); };
        $this->sheetMappings['rowDimension']['rowIndex'] = function($key, $value) { $this->sheetMappings['rowDimension']['__object']($key)->setRowIndex($value); };
        $this->sheetMappings['rowDimension']['visible'] = function($key, $value) { $this->sheetMappings['rowDimension']['__object']($key)->setVisible($value); };
        $this->sheetMappings['rowDimension']['xfIndex'] = function($key, $value) { $this->sheetMappings['rowDimension']['__object']($key)->setXfIndex($value); };
        $this->sheetMappings['rowDimension']['zeroHeight'] = function($key, $value) { $this->sheetMappings['rowDimension']['__object']($key)->setZeroHeight($value); };
        $this->sheetMappings['sheetState'] = function($value) { $this->sheet->setSheetState($value); };
        $this->sheetMappings['showGridlines'] = function($value) { $this->sheet->setShowGridlines($value); };
        $this->sheetMappings['tabColor'] = function($value) { $this->sheet->getTabColor()->setRGB($value); };
        $this->sheetMappings['zoomScale'] = function($value) { $this->sheet->getSheetView()->setZoomScale($value); };
    }

    protected function initRowPropertyMappings() {
        // nothing
    }

    protected function initCellPropertyMappings() {
        $this->cellMappings['break'] = function($value) { $this->sheet->setBreak($this->cell->getCoordinate(), $value); };
        $this->cellMappings['dataValidation']['allowBlank'] = function($value) { $this->cell->getDataValidation()->setAllowBlank($value); };
        $this->cellMappings['dataValidation']['error'] = function($value) { $this->cell->getDataValidation()->setError($value); };
        $this->cellMappings['dataValidation']['errorStyle'] = function($value) { $this->cell->getDataValidation()->setErrorStyle($value); };
        $this->cellMappings['dataValidation']['errorTitle'] = function($value) { $this->cell->getDataValidation()->setErrorTitle($value); };
        $this->cellMappings['dataValidation']['formula1'] = function($value) { $this->cell->getDataValidation()->setFormula1($value); };
        $this->cellMappings['dataValidation']['formula2'] = function($value) { $this->cell->getDataValidation()->setFormula2($value); };
        $this->cellMappings['dataValidation']['operator'] = function($value) { $this->cell->getDataValidation()->setOperator($value); };
        $this->cellMappings['dataValidation']['prompt'] = function($value) { $this->cell->getDataValidation()->setPrompt($value); };
        $this->cellMappings['dataValidation']['promptTitle'] = function($value) { $this->cell->getDataValidation()->setPromptTitle($value); };
        $this->cellMappings['dataValidation']['showDropDown'] = function($value) { $this->cell->getDataValidation()->setShowDropDown($value); };
        $this->cellMappings['dataValidation']['showErrorMessage'] = function($value) { $this->cell->getDataValidation()->setShowErrorMessage($value); };
        $this->cellMappings['dataValidation']['showInputMessage'] = function($value) { $this->cell->getDataValidation()->setShowInputMessage($value); };
        $this->cellMappings['dataValidation']['type'] = function($value) { $this->cell->getDataValidation()->setType($value); };
        $this->cellMappings['style'] = function($value) { $this->sheet->getStyle($this->cell->getCoordinate())->applyFromArray($value); };
        $this->cellMappings['url'] = function($value) { $this->cell->getHyperlink()->setUrl($value); };
    }

    protected function initDrawingPropertyMappings() {
        $this->drawingMappings['coordinates'] = function($value) { $this->drawing->getCoordinates($value); };
        $this->drawingMappings['description'] = function($value) { $this->drawing->getDescription($value); };
        $this->drawingMappings['height'] = function($value) { $this->drawing->getHeight($value); };
        $this->drawingMappings['name'] = function($value) { $this->drawing->getName($value); };
        $this->drawingMappings['offsetX'] = function($value) { $this->drawing->getOffsetX($value); };
        $this->drawingMappings['offsetY'] = function($value) { $this->drawing->getOffsetY($value); };
        $this->drawingMappings['resizeProportional'] = function($value) { $this->drawing->getResizeProportional($value); };
        $this->drawingMappings['rotation'] = function($value) { $this->drawing->getRotation($value); };
        $this->drawingMappings['shadow']['alignment'] = function($value) { $this->drawing->getShadow()->setAlignment($value); };
        $this->drawingMappings['shadow']['alpha'] = function($value) { $this->drawing->getShadow()->setAlpha($value); };
        $this->drawingMappings['shadow']['blurRadius'] = function($value) { $this->drawing->getShadow()->setBlurRadius($value); };
        $this->drawingMappings['shadow']['color'] = function($value) { $this->drawing->getShadow()->setColor()->setRgb($value); };
        $this->drawingMappings['shadow']['direction'] = function($value) { $this->drawing->getShadow()->setDirection($value); };
        $this->drawingMappings['shadow']['distance'] = function($value) { $this->drawing->getShadow()->setDistance($value); };
        $this->drawingMappings['shadow']['visible'] = function($value) { $this->drawing->getShadow()->setVisible($value); };
        $this->drawingMappings['width'] = function($value) { $this->drawing->getWidth($value); };
    }

    public function setProperties(array $properties, array $mappings) {
        foreach ($properties as $key => $value) {
            if (array_key_exists($key, $mappings)) {
                if (is_array($value)) {
                    if ($mappings['__multi']) {
                        foreach ($value as $_key => $_value) {
                            $this->setPropertiesByKey($_key, $_value, $mappings[$key]);
                        }
                    } else {
                        $this->setProperties($value, $mappings[$key]);
                    }
                } else {
                    $mappings[$key](trim($value));
                }
            }
        }
    }

    public function setPropertiesByKey($key, array $properties, array $mappings) {
        foreach ($properties as $_key => $value) {
            if (array_key_exists($_key, $mappings)) {
                if (is_array($value)) {
                    $this->setPropertiesByKey($key, $value, $mappings[$_key]);
                } else {
                    $mappings[$_key]($key, trim($value));
                }
            }
        }
    }

    public function tagDocument(array $properties = null) {
        if ($properties != null) {
            $this->setProperties($properties, $this->documentMappings);
        }
    }

    public function tagSheet($index, array $properties = null) {
        if ($index == null) {
            throw new \LogicException();
        }

        if (!$this->document->sheetNameExists($index)) {
            $this->document->createSheet()->setTitle($index);
        }

        $this->sheet = $this->document->setActiveSheetIndexByName($index);
        $this->row = 1;
        $this->cell = null;
        
        if ($properties != null) {
            $this->setProperties($properties, $this->sheetMappings);
        }
    }

    public function tagRow($index, array $properties = null) {
        if ($this->sheet == null || !is_int($index)) {
            throw new \LogicException();
        }
        
        $this->row = $index;
        $this->cell = null;

        if ($properties != null) {
            $this->setProperties($properties, $this->rowMappings);
        }
    }

    public function tagCell($index, $value = null, array $properties = null) {
        if ($this->sheet == null || !preg_match('/^[A-Z]$/', $index)) {
            throw new \LogicException();
        }

        $this->cell = $this->sheet->getCellByColumnAndRow($index, $this->row);

        if ($value != null) {
            $this->cell->setValue($value);
        }
        
        if ($properties != null) {
            $this->setProperties($properties, $this->cellMappings);
        }
    }
    
    public function tagDrawing($path, array $properties = null) {
        if ($this->sheet == null || !file_exists($path)) {
            throw new \LogicException();
        }

        $this->drawing = new \PHPExcel_Worksheet_Drawing();
        $this->drawing->setWorksheet($this->sheet);
        $this->drawing->setPath($path);

        if ($properties != null) {
            $this->setProperties($properties, $this->cellMappings);
        }

        $this->drawing = null;
    }

    public function save() {
        \PHPExcel_IOFactory::createWriter($this->phpExcel, 'Excel5')->save('php://output');
    }
}