<?php

namespace MewesK\TwigExcelBundle\Wrapper;

/**
 * Class XlsCellWrapper
 *
 * @package MewesK\TwigExcelBundle\Wrapper
 */
class XlsCellWrapper extends AbstractWrapper
{
    /**
     * @var array
     */
    protected $context;
    /**
     * @var XlsSheetWrapper
     */
    protected $sheetWrapper;

    /**
     * @var \PHPExcel_Cell
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
     * @param XlsSheetWrapper $sheetWrapper
     */
    public function __construct(array $context, XlsSheetWrapper $sheetWrapper)
    {
        $this->context = $context;
        $this->sheetWrapper = $sheetWrapper;

        $this->object = null;
        $this->attributes = [];
        $this->mappings = [];

        $this->initializeMappings();
    }

    protected function initializeMappings()
    {
        $wrapper = $this; // PHP 5.3 fix

        $this->mappings['break'] = function ($value) use ($wrapper) {
            $wrapper->sheetWrapper->getObject()->setBreak($wrapper->object->getCoordinate(), $value);
        };
        $this->mappings['dataType'] = function ($value) use ($wrapper) {
            $wrapper->object->setDataType($value);
        };
        $this->mappings['dataValidation']['allowBlank'] = function ($value) use ($wrapper) {
            $wrapper->object->getDataValidation()->setAllowBlank($value);
        };
        $this->mappings['dataValidation']['error'] = function ($value) use ($wrapper) {
            $wrapper->object->getDataValidation()->setError($value);
        };
        $this->mappings['dataValidation']['errorStyle'] = function ($value) use ($wrapper) {
            $wrapper->object->getDataValidation()->setErrorStyle($value);
        };
        $this->mappings['dataValidation']['errorTitle'] = function ($value) use ($wrapper) {
            $wrapper->object->getDataValidation()->setErrorTitle($value);
        };
        $this->mappings['dataValidation']['formula1'] = function ($value) use ($wrapper) {
            $wrapper->object->getDataValidation()->setFormula1($value);
        };
        $this->mappings['dataValidation']['formula2'] = function ($value) use ($wrapper) {
            $wrapper->object->getDataValidation()->setFormula2($value);
        };
        $this->mappings['dataValidation']['operator'] = function ($value) use ($wrapper) {
            $wrapper->object->getDataValidation()->setOperator($value);
        };
        $this->mappings['dataValidation']['prompt'] = function ($value) use ($wrapper) {
            $wrapper->object->getDataValidation()->setPrompt($value);
        };
        $this->mappings['dataValidation']['promptTitle'] = function ($value) use ($wrapper) {
            $wrapper->object->getDataValidation()->setPromptTitle($value);
        };
        $this->mappings['dataValidation']['showDropDown'] = function ($value) use ($wrapper) {
            $wrapper->object->getDataValidation()->setShowDropDown($value);
        };
        $this->mappings['dataValidation']['showErrorMessage'] = function ($value) use ($wrapper) {
            $wrapper->object->getDataValidation()->setShowErrorMessage($value);
        };
        $this->mappings['dataValidation']['showInputMessage'] = function ($value) use ($wrapper) {
            $wrapper->object->getDataValidation()->setShowInputMessage($value);
        };
        $this->mappings['dataValidation']['type'] = function ($value) use ($wrapper) {
            $wrapper->object->getDataValidation()->setType($value);
        };
        $this->mappings['style'] = function ($value) use ($wrapper) {
            $wrapper->sheetWrapper->getObject()->getStyle($wrapper->object->getCoordinate())->applyFromArray($value);
        };
        $this->mappings['url'] = function ($value) use ($wrapper) {
            $wrapper->object->getHyperlink()->setUrl($value);
        };
    }

    /**
     * @param null|int $index
     * @param null|mixed $value
     * @param null|array $properties
     *
     * @throws \PHPExcel_Exception
     */
    public function start($index = null, $value = null, array $properties = null)
    {
        if ($this->sheetWrapper->getObject() === null) {
            throw new \LogicException();
        }
        if ($index !== null && !is_int($index)) {
            throw new \InvalidArgumentException(sprintf('Invalid index'));
        }

        if ($index === null) {
            $this->sheetWrapper->increaseColumn();
        } else {
            $this->sheetWrapper->setColumn($index);
        }

        $this->object = $this->sheetWrapper->getObject()->getCellByColumnAndRow($this->sheetWrapper->getColumn(),
            $this->sheetWrapper->getRow());

        if ($value !== null) {
            $this->object->setValue($value);
        }

        if ($properties !== null) {
            $this->setProperties($properties, $this->mappings);
        }

        $this->attributes['value'] = $value;
        $this->attributes['properties'] = $properties ?: [];
    }

    public function end()
    {
        $this->object = null;
        $this->attributes = [];
    }

    //
    // Getters/Setters
    //

    /**
     * @return \PHPExcel_Cell
     */
    public function getObject()
    {
        return $this->object;
    }

    /**
     * @param \PHPExcel_Cell $object
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
