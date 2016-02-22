<?php

namespace MewesK\TwigExcelBundle\Wrapper;

use PHPExcel_IOFactory;
use PHPExcel_Settings;
use PHPExcel_Writer_Abstract;
use ReflectionClass;
use Symfony\Bridge\Twig\AppVariable;
use Twig_Environment;
use Twig_Loader_Filesystem;

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
     * @var Twig_Environment
     */
    protected $environment;

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
     * XlsDocumentWrapper constructor.
     * @param array $context
     * @param Twig_Environment $environment
     */
    public function __construct(array $context, Twig_Environment $environment)
    {
        $this->context = $context;
        $this->environment = $environment;

        $this->object = null;
        $this->attributes = [];
        $this->mappings = [];

        $this->initializeMappings();
    }

    protected function initializeMappings()
    {
        $this->mappings['category'] = function ($value) {
            $this->object->getProperties()->setCategory($value);
        };
        $this->mappings['company'] = function ($value) {
            $this->object->getProperties()->setCompany($value);
        };
        $this->mappings['created'] = function ($value) {
            $this->object->getProperties()->setCreated($value);
        };
        $this->mappings['creator'] = function ($value) {
            $this->object->getProperties()->setCreator($value);
        };
        $this->mappings['defaultStyle'] = function ($value) {
            $this->object->getDefaultStyle()->applyFromArray($value);
        };
        $this->mappings['description'] = function ($value) {
            $this->object->getProperties()->setDescription($value);
        };
        $this->mappings['format'] = function ($value) {
            $this->attributes['format'] = $value;
        };
        $this->mappings['keywords'] = function ($value) {
            $this->object->getProperties()->setKeywords($value);
        };
        $this->mappings['lastModifiedBy'] = function ($value) {
            $this->object->getProperties()->setLastModifiedBy($value);
        };
        $this->mappings['manager'] = function ($value) {
            $this->object->getProperties()->setManager($value);
        };
        $this->mappings['modified'] = function ($value) {
            $this->object->getProperties()->setModified($value);
        };
        $this->mappings['security']['lockRevision'] = function ($value) {
            $this->object->getSecurity()->setLockRevision($value);
        };
        $this->mappings['security']['lockStructure'] = function ($value) {
            $this->object->getSecurity()->setLockStructure($value);
        };
        $this->mappings['security']['lockWindows'] = function ($value) {
            $this->object->getSecurity()->setLockWindows($value);
        };
        $this->mappings['security']['revisionsPassword'] = function ($value) {
            $this->object->getSecurity()->setRevisionsPassword($value);
        };
        $this->mappings['security']['workbookPassword'] = function ($value) {
            $this->object->getSecurity()->setWorkbookPassword($value);
        };
        $this->mappings['subject'] = function ($value) {
            $this->object->getProperties()->setSubject($value);
        };
        $this->mappings['template'] = function ($value) {
            $this->attributes['template'] = $value;
        };
        $this->mappings['title'] = function ($value) {
            $this->object->getProperties()->setTitle($value);
        };
    }

    /**
     * @param null|array $properties
     *
     * @throws \PHPExcel_Exception
     */
    public function start(array $properties = null)
    {
        // load template
        if (array_key_exists('template', $properties)) {
            $templatePath = $this->expandPath($properties['template']);
            $reader = PHPExcel_IOFactory::createReaderForFile($templatePath);
            $this->object = $reader->load($templatePath);
        }

        // create new
        else {
            $this->object = new \PHPExcel();
            $this->object->removeSheetByIndex(0);
        }

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
    // Helpers
    //

    /**
     * Resolves properties containing paths using namespaces.
     *
     * @param string $path
     * @return bool
     */
    private function expandPath($path)
    {
        $loader = $this->environment->getLoader();
        if ($loader instanceof Twig_Loader_Filesystem) {
            /**
             * @var Twig_Loader_Filesystem $loader
             */
            foreach ($loader->getNamespaces() as $namespace) {
                if (strpos($path, $namespace) === 1) {
                    foreach ($loader->getPaths($namespace) as $namespacePath) {
                        $expandedPathAttribute = str_replace('@' . $namespace, $namespacePath, $path);
                        if (file_exists($expandedPathAttribute)) {
                            return $expandedPathAttribute;
                        }
                    }
                }
            }
        }
        return $path;
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
