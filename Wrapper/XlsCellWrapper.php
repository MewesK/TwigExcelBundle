<?php

namespace MewesK\TwigExcelBundle\Wrapper;
use Twig_Environment;

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
     * @var Twig_Environment
     */
    protected $environment;
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
     * XlsCellWrapper constructor.
     * @param array $context
     * @param Twig_Environment $environment
     * @param XlsSheetWrapper $sheetWrapper
     */
    public function __construct(array $context, Twig_Environment $environment, XlsSheetWrapper $sheetWrapper)
    {
        $this->context = $context;
        $this->environment = $environment;
        $this->sheetWrapper = $sheetWrapper;

        $this->object = null;
        $this->attributes = [];
        $this->mappings = [];

        $this->initializeMappings();
    }

    protected function initializeMappings()
    {
        $this->mappings['break'] = function ($value) {
            $this->sheetWrapper->getObject()->setBreak($this->object->getCoordinate(), $value);
        };
        $this->mappings['dataType'] = function ($value) {
            $this->object->setDataType($value);
        };
        $this->mappings['dataValidation']['allowBlank'] = function ($value) {
            $this->object->getDataValidation()->setAllowBlank($value);
        };
        $this->mappings['dataValidation']['error'] = function ($value) {
            $this->object->getDataValidation()->setError($value);
        };
        $this->mappings['dataValidation']['errorStyle'] = function ($value) {
            $this->object->getDataValidation()->setErrorStyle($value);
        };
        $this->mappings['dataValidation']['errorTitle'] = function ($value) {
            $this->object->getDataValidation()->setErrorTitle($value);
        };
        $this->mappings['dataValidation']['formula1'] = function ($value) {
            $this->object->getDataValidation()->setFormula1($value);
        };
        $this->mappings['dataValidation']['formula2'] = function ($value) {
            $this->object->getDataValidation()->setFormula2($value);
        };
        $this->mappings['dataValidation']['operator'] = function ($value) {
            $this->object->getDataValidation()->setOperator($value);
        };
        $this->mappings['dataValidation']['prompt'] = function ($value) {
            $this->object->getDataValidation()->setPrompt($value);
        };
        $this->mappings['dataValidation']['promptTitle'] = function ($value) {
            $this->object->getDataValidation()->setPromptTitle($value);
        };
        $this->mappings['dataValidation']['showDropDown'] = function ($value) {
            $this->object->getDataValidation()->setShowDropDown($value);
        };
        $this->mappings['dataValidation']['showErrorMessage'] = function ($value) {
            $this->object->getDataValidation()->setShowErrorMessage($value);
        };
        $this->mappings['dataValidation']['showInputMessage'] = function ($value) {
            $this->object->getDataValidation()->setShowInputMessage($value);
        };
        $this->mappings['dataValidation']['type'] = function ($value) {
            $this->object->getDataValidation()->setType($value);
        };
        $this->mappings['style'] = function ($value) {
            $this->sheetWrapper->getObject()->getStyle($this->object->getCoordinate())->applyFromArray($value);
        };
        $this->mappings['url'] = function ($value) {
            $this->object->getHyperlink()->setUrl($value);
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
            throw new \InvalidArgumentException('Invalid index');
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
