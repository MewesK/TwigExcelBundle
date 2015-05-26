<?php

namespace MewesK\TwigExcelBundle\Wrapper;
use PHPExcel_Writer_Abstract;

/**
 * Class XlsDocumentWrapper
 *
 * @package MewesK\TwigExcelBundle\Wrapper
 */
class XlsDocumentWrapper extends AbstractWrapper
{
    /**
     * @var array
     */
    protected $context;

    /**
     * @var \PHPExcel
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
     */
    public function __construct(array $context)
    {
        $this->context = $context;

        $this->object = null;
        $this->attributes = [];
        $this->mappings = [];

        $this->initializeMappings();
    }

    protected function initializeMappings()
    {
        $wrapper = $this; // PHP 5.3 fix

        $this->mappings['category'] = function ($value) use ($wrapper) {
            $wrapper->object->getProperties()->setCategory($value);
        };
        $this->mappings['company'] = function ($value) use ($wrapper) {
            $wrapper->object->getProperties()->setCompany($value);
        };
        $this->mappings['created'] = function ($value) use ($wrapper) {
            $wrapper->object->getProperties()->setCreated($value);
        };
        $this->mappings['creator'] = function ($value) use ($wrapper) {
            $wrapper->object->getProperties()->setCreator($value);
        };
        $this->mappings['defaultStyle'] = function ($value) use ($wrapper) {
            $wrapper->object->getDefaultStyle()->applyFromArray($value);
        };
        $this->mappings['description'] = function ($value) use ($wrapper) {
            $wrapper->object->getProperties()->setDescription($value);
        };
        $this->mappings['format'] = function ($value) use ($wrapper) {
            $wrapper->attributes['format'] = $value;
        };
        $this->mappings['keywords'] = function ($value) use ($wrapper) {
            $wrapper->object->getProperties()->setKeywords($value);
        };
        $this->mappings['lastModifiedBy'] = function ($value) use ($wrapper) {
            $wrapper->object->getProperties()->setLastModifiedBy($value);
        };
        $this->mappings['manager'] = function ($value) use ($wrapper) {
            $wrapper->object->getProperties()->setManager($value);
        };
        $this->mappings['modified'] = function ($value) use ($wrapper) {
            $wrapper->object->getProperties()->setModified($value);
        };
        $this->mappings['security']['lockRevision'] = function ($value) use ($wrapper) {
            $wrapper->object->getSecurity()->setLockRevision($value);
        };
        $this->mappings['security']['lockStructure'] = function ($value) use ($wrapper) {
            $wrapper->object->getSecurity()->setLockStructure($value);
        };
        $this->mappings['security']['lockWindows'] = function ($value) use ($wrapper) {
            $wrapper->object->getSecurity()->setLockWindows($value);
        };
        $this->mappings['security']['revisionsPassword'] = function ($value) use ($wrapper) {
            $wrapper->object->getSecurity()->setRevisionsPassword($value);
        };
        $this->mappings['security']['workbookPassword'] = function ($value) use ($wrapper) {
            $wrapper->object->getSecurity()->setWorkbookPassword($value);
        };
        $this->mappings['subject'] = function ($value) use ($wrapper) {
            $wrapper->object->getProperties()->setSubject($value);
        };
        $this->mappings['title'] = function ($value) use ($wrapper) {
            $wrapper->object->getProperties()->setTitle($value);
        };
    }

    /**
     * @param null|array $properties
     *
     * @throws \PHPExcel_Exception
     */
    public function start(array $properties = null)
    {
        $this->object = new \PHPExcel();
        $this->object->removeSheetByIndex(0);
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
        // try document property
        if (array_key_exists('format', $this->attributes)) {
            $format = $this->attributes['format'];
        } // try symfony request
        else {
            $app = is_array($this->context) && array_key_exists('app', $this->context) ? $this->context['app'] : null;
            $request = $app && is_callable([$app, 'getRequest']) ? $app->getRequest() : null;
            $format = $request && is_callable([$request, 'getRequestFormat']) ? $request->getRequestFormat() : null;
        }
        // set default
        if ($format === null || !is_string($format)) {
            $format = 'xlsx';
        }

        switch (strtolower($format)) {
            case 'csv':
                $writerType = 'CSV';
                break;
            case 'ods':
                $writerType = 'OpenDocument';
                break;
            case 'pdf':
                $writerType = 'PDF';
                break;
            case 'xls':
                $writerType = 'Excel5';
                break;
            case 'xlsx':
                $writerType = 'Excel2007';
                break;
            default:
                throw new \InvalidArgumentException();
        }

        /**
         * @var $writer PHPExcel_Writer_Abstract
         */
        $writer = \PHPExcel_IOFactory::createWriter($this->object, $writerType);
        //TODO: make configurable
        $writer->setPreCalculateFormulas(true);
        //$writer->setUseDiskCaching(true);
        $writer->save('php://output');

        $this->object = null;
        $this->attributes = [];
    }

    //
    // Getters/Setters
    //

    /**
     * @return \PHPExcel
     */
    public function getObject()
    {
        return $this->object;
    }

    /**
     * @param \PHPExcel $object
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
