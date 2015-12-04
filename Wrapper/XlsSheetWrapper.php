<?php

namespace MewesK\TwigExcelBundle\Wrapper;

/**
 * Class XlsSheetWrapper
 *
 * @package MewesK\TwigExcelBundle\Wrapper
 */
class XlsSheetWrapper extends AbstractWrapper
{
    /**
     * @var int
     */
    public static $COLUMN_DEFAULT = 0;
    /**
     * @var int
     */
    public static $ROW_DEFAULT = 1;

    /**
     * @var array
     */
    protected $context;
    /**
     * @var XlsDocumentWrapper
     */
    protected $documentWrapper;

    /**
     * @var null|int
     */
    protected $row;
    /**
     * @var null|int
     */
    protected $column;

    /**
     * @var \PHPExcel_Worksheet
     */
    protected $object;
    /**
     * @var array
     */
    protected $attributes;
    /**
     * @var array
     */
    protected $mappings;

    /**
     * @param array $context
     * @param XlsDocumentWrapper $documentWrapper
     */
    public function __construct(array $context, XlsDocumentWrapper $documentWrapper)
    {
        $this->context = $context;
        $this->documentWrapper = $documentWrapper;

        $this->row = null;
        $this->column = null;

        $this->object = null;
        $this->attributes = [];
        $this->mappings = [];

        $this->initializeMappings();
    }

    protected function initializeMappings()
    {
        $wrapper = $this; // PHP 5.3 fix

        $this->mappings['columnDimension']['__multi'] = true;
        $this->mappings['columnDimension']['__object'] = function ($key = 'default') use ($wrapper) {
            return $key === 'default' ? $wrapper->object->getDefaultColumnDimension() : $wrapper->object->getColumnDimension($key);
        };
        $this->mappings['columnDimension']['autoSize'] = function ($key, $value) use ($wrapper) {
            $wrapper->mappings['columnDimension']['__object']($key)->setAutoSize($value);
        };
        $this->mappings['columnDimension']['collapsed'] = function ($key, $value) use ($wrapper) {
            $wrapper->mappings['columnDimension']['__object']($key)->setCollapsed($value);
        };
        $this->mappings['columnDimension']['columnIndex'] = function ($key, $value) use ($wrapper) {
            $wrapper->mappings['columnDimension']['__object']($key)->setColumnIndex($value);
        };
        $this->mappings['columnDimension']['outlineLevel'] = function ($key, $value) use ($wrapper) {
            $wrapper->mappings['columnDimension']['__object']($key)->setOutlineLevel($value);
        };
        $this->mappings['columnDimension']['visible'] = function ($key, $value) use ($wrapper) {
            $wrapper->mappings['columnDimension']['__object']($key)->setVisible($value);
        };
        $this->mappings['columnDimension']['width'] = function ($key, $value) use ($wrapper) {
            $wrapper->mappings['columnDimension']['__object']($key)->setWidth($value);
        };
        $this->mappings['columnDimension']['xfIndex'] = function ($key, $value) use ($wrapper) {
            $wrapper->mappings['columnDimension']['__object']($key)->setXfIndex($value);
        };
        $this->mappings['pageMargins']['top'] = function ($value) use ($wrapper) {
            $wrapper->object->getPageMargins()->setTop($value);
        };
        $this->mappings['pageMargins']['bottom'] = function ($value) use ($wrapper) {
            $wrapper->object->getPageMargins()->setBottom($value);
        };
        $this->mappings['pageMargins']['left'] = function ($value) use ($wrapper) {
            $wrapper->object->getPageMargins()->setLeft($value);
        };
        $this->mappings['pageMargins']['right'] = function ($value) use ($wrapper) {
            $wrapper->object->getPageMargins()->setRight($value);
        };
        $this->mappings['pageMargins']['header'] = function ($value) use ($wrapper) {
            $wrapper->object->getPageMargins()->setHeader($value);
        };
        $this->mappings['pageMargins']['footer'] = function ($value) use ($wrapper) {
            $wrapper->object->getPageMargins()->setFooter($value);
        };
        $this->mappings['pageSetup']['fitToHeight'] = function ($value) use ($wrapper) {
            $wrapper->object->getPageSetup()->setFitToHeight($value);
        };
        $this->mappings['pageSetup']['fitToPage'] = function ($value) use ($wrapper) {
            $wrapper->object->getPageSetup()->setFitToPage($value);
        };
        $this->mappings['pageSetup']['fitToWidth'] = function ($value) use ($wrapper) {
            $wrapper->object->getPageSetup()->setFitToWidth($value);
        };
        $this->mappings['pageSetup']['horizontalCentered'] = function ($value) use ($wrapper) {
            $wrapper->object->getPageSetup()->setHorizontalCentered($value);
        };
        $this->mappings['pageSetup']['orientation'] = function ($value) use ($wrapper) {
            $wrapper->object->getPageSetup()->setOrientation($value);
        };
        $this->mappings['pageSetup']['paperSize'] = function ($value) use ($wrapper) {
            $wrapper->object->getPageSetup()->setPaperSize($value);
        };
        $this->mappings['pageSetup']['printArea'] = function ($value) use ($wrapper) {
            $wrapper->object->getPageSetup()->setPrintArea($value);
        };
        $this->mappings['pageSetup']['scale'] = function ($value) use ($wrapper) {
            $wrapper->object->getPageSetup()->setScale($value);
        };
        $this->mappings['pageSetup']['verticalCentered'] = function ($value) use ($wrapper) {
            $wrapper->object->getPageSetup()->setVerticalCentered($value);
        };
        $this->mappings['printGridlines'] = function ($value) use ($wrapper) {
            $wrapper->object->setPrintGridlines($value);
        };
        $this->mappings['protection']['autoFilter'] = function ($value) use ($wrapper) {
            $wrapper->object->getProtection()->setAutoFilter($value);
        };
        $this->mappings['protection']['deleteColumns'] = function ($value) use ($wrapper) {
            $wrapper->object->getProtection()->setDeleteColumns($value);
        };
        $this->mappings['protection']['deleteRows'] = function ($value) use ($wrapper) {
            $wrapper->object->getProtection()->setDeleteRows($value);
        };
        $this->mappings['protection']['formatCells'] = function ($value) use ($wrapper) {
            $wrapper->object->getProtection()->setFormatCells($value);
        };
        $this->mappings['protection']['formatColumns'] = function ($value) use ($wrapper) {
            $wrapper->object->getProtection()->setFormatColumns($value);
        };
        $this->mappings['protection']['formatRows'] = function ($value) use ($wrapper) {
            $wrapper->object->getProtection()->setFormatRows($value);
        };
        $this->mappings['protection']['insertColumns'] = function ($value) use ($wrapper) {
            $wrapper->object->getProtection()->setInsertColumns($value);
        };
        $this->mappings['protection']['insertHyperlinks'] = function ($value) use ($wrapper) {
            $wrapper->object->getProtection()->setInsertHyperlinks($value);
        };
        $this->mappings['protection']['insertRows'] = function ($value) use ($wrapper) {
            $wrapper->object->getProtection()->setInsertRows($value);
        };
        $this->mappings['protection']['objects'] = function ($value) use ($wrapper) {
            $wrapper->object->getProtection()->setObjects($value);
        };
        $this->mappings['protection']['password'] = function ($value) use ($wrapper) {
            $wrapper->object->getProtection()->setPassword($value);
        };
        $this->mappings['protection']['pivotTables'] = function ($value) use ($wrapper) {
            $wrapper->object->getProtection()->setPivotTables($value);
        };
        $this->mappings['protection']['scenarios'] = function ($value) use ($wrapper) {
            $wrapper->object->getProtection()->setScenarios($value);
        };
        $this->mappings['protection']['selectLockedCells'] = function ($value) use ($wrapper) {
            $wrapper->object->getProtection()->setSelectLockedCells($value);
        };
        $this->mappings['protection']['selectUnlockedCells'] = function ($value) use ($wrapper) {
            $wrapper->object->getProtection()->setSelectUnlockedCells($value);
        };
        $this->mappings['protection']['sheet'] = function ($value) use ($wrapper) {
            $wrapper->object->getProtection()->setSheet($value);
        };
        $this->mappings['protection']['sort'] = function ($value) use ($wrapper) {
            $wrapper->object->getProtection()->setSort($value);
        };
        $this->mappings['rightToLeft'] = function ($value) use ($wrapper) {
            $wrapper->object->setRightToLeft($value);
        };
        $this->mappings['rowDimension']['__multi'] = true;
        $this->mappings['rowDimension']['__object'] = function ($key) use ($wrapper) {
            return $key === 'default' ? $wrapper->object->getDefaultRowDimension() : $wrapper->object->getRowDimension($key);
        };
        $this->mappings['rowDimension']['collapsed'] = function ($key, $value) use ($wrapper) {
            $wrapper->mappings['rowDimension']['__object']($key)->setCollapsed($value);
        };
        $this->mappings['rowDimension']['outlineLevel'] = function ($key, $value) use ($wrapper) {
            $wrapper->mappings['rowDimension']['__object']($key)->setOutlineLevel($value);
        };
        $this->mappings['rowDimension']['rowHeight'] = function ($key, $value) use ($wrapper) {
            $wrapper->mappings['rowDimension']['__object']($key)->setRowHeight($value);
        };
        $this->mappings['rowDimension']['rowIndex'] = function ($key, $value) use ($wrapper) {
            $wrapper->mappings['rowDimension']['__object']($key)->setRowIndex($value);
        };
        $this->mappings['rowDimension']['visible'] = function ($key, $value) use ($wrapper) {
            $wrapper->mappings['rowDimension']['__object']($key)->setVisible($value);
        };
        $this->mappings['rowDimension']['xfIndex'] = function ($key, $value) use ($wrapper) {
            $wrapper->mappings['rowDimension']['__object']($key)->setXfIndex($value);
        };
        $this->mappings['rowDimension']['zeroHeight'] = function ($key, $value) use ($wrapper) {
            $wrapper->mappings['rowDimension']['__object']($key)->setZeroHeight($value);
        };
        $this->mappings['sheetState'] = function ($value) use ($wrapper) {
            $wrapper->object->setSheetState($value);
        };
        $this->mappings['showGridlines'] = function ($value) use ($wrapper) {
            $wrapper->object->setShowGridlines($value);
        };
        $this->mappings['tabColor'] = function ($value) use ($wrapper) {
            $wrapper->object->getTabColor()->setRGB($value);
        };
        $this->mappings['zoomScale'] = function ($value) use ($wrapper) {
            $wrapper->object->getSheetView()->setZoomScale($value);
        };
    }

    /**
     * @param $index
     * @param array $properties
     */
    public function start($index, array $properties = null)
    {
        if ($index === null || !is_string($index)) {
            throw new \InvalidArgumentException('Invalid index');
        }
        if (!$this->documentWrapper->getObject()->sheetNameExists($index)) {
            $this->documentWrapper->getObject()->createSheet()->setTitle($index);
        }

        $this->object = $this->documentWrapper->getObject()->setActiveSheetIndexByName($index);
        $this->attributes['index'] = $index;
        $this->attributes['properties'] = $properties ?: [];

        if ($properties !== null) {
            $this->setProperties($properties, $this->mappings);
        }
    }

    /**
     * @throws \PHPExcel_Reader_Exception
     */
    public function end()
    {
        $this->object = null;
        $this->attributes = [];
        $this->row = null;
    }

    //
    // Helpers
    //

    public function increaseRow()
    {
        $this->row = $this->row === null ? self::$ROW_DEFAULT : $this->row + 1;
    }

    public function increaseColumn()
    {
        $this->column = $this->column === null ? self::$COLUMN_DEFAULT : $this->column + 1;
    }

    //
    // Getters/Setters
    //

    /**
     * @return int|null
     */
    public function getRow()
    {
        return $this->row;
    }

    /**
     * @param int|null $row
     */
    public function setRow($row)
    {
        $this->row = $row;
    }

    /**
     * @return int|null
     */
    public function getColumn()
    {
        return $this->column;
    }

    /**
     * @param int|null $column
     */
    public function setColumn($column)
    {
        $this->column = $column;
    }

    /**
     * @return \PHPExcel_Worksheet
     */
    public function getObject()
    {
        return $this->object;
    }

    /**
     * @param \PHPExcel_Worksheet $object
     */
    public function setObject($object)
    {
        $this->object = $object;
    }

    /**
     * @return array
     */
    public function getAttributes()
    {
        return $this->attributes;
    }

    /**
     * @param array $attributes
     */
    public function setAttributes($attributes)
    {
        $this->attributes = $attributes;
    }

    /**
     * @return array
     */
    public function getMappings()
    {
        return $this->mappings;
    }

    /**
     * @param array $mappings
     */
    public function setMappings($mappings)
    {
        $this->mappings = $mappings;
    }
}
