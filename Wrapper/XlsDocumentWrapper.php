<?php

namespace MewesK\TwigExcelBundle\Wrapper;

use PHPExcel_Settings;
use PHPExcel_Writer_Abstract;
use ReflectionClass;
use Symfony\Bridge\Twig\AppVariable;

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
     * @param bool $preCalculateFormulas
     * @param null|string $diskCachingDirectory
     *
     * @throws \InvalidArgumentException
     * @throws \PHPExcel_Exception
     * @throws \PHPExcel_Reader_Exception
     * @throws \PHPExcel_Writer_Exception
     */
    public function end($preCalculateFormulas = true, $diskCachingDirectory = null)
    {
        $format = null;

        // try document property
        if (array_key_exists('format', $this->attributes)) {
            $format = $this->attributes['format'];
        }

         // try Symfony request
        else if (array_key_exists('app', $this->context)) {
            /**
             * @var $appVariable AppVariable
             */
            $appVariable = $this->context['app'];
            if ($appVariable instanceof AppVariable && $appVariable->getRequest() !== null) {
                $format = $appVariable->getRequest()->getRequestFormat();
            }
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
                try {
                    $reflectionClass = new ReflectionClass('mPDF');
                    $path = dirname($reflectionClass->getFileName());
                    if (!PHPExcel_Settings::setPdfRenderer(PHPExcel_Settings::PDF_RENDERER_MPDF, $path)) {
                        throw new \PHPExcel_Exception();
                    }
                } catch (\Exception $e) {
                    throw new \PHPExcel_Exception('Error loading mPDF. Is mPDF correctly installed?', $e->getCode(), $e);
                }
                break;
            case 'xls':
                $writerType = 'Excel5';
                break;
            case 'xlsx':
                $writerType = 'Excel2007';
                break;
            default:
                throw new \InvalidArgumentException(sprintf('Unknown format "%s"', $format));
        }

        /**
         * @var $writer PHPExcel_Writer_Abstract
         */
        $writer = \PHPExcel_IOFactory::createWriter($this->object, $writerType);
        $writer->setPreCalculateFormulas($preCalculateFormulas);
        $writer->setUseDiskCaching($diskCachingDirectory !== null, $diskCachingDirectory);
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
