<?php

namespace MewesK\TwigExcelBundle\Wrapper;
use Twig_Environment;

/**
 * Class XlsDrawingWrapper
 *
 * @package MewesK\TwigExcelBundle\Wrapper
 */
class XlsDrawingWrapper extends AbstractWrapper
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
     * @var XlsHeaderFooterWrapper
     */
    protected $headerFooterWrapper;

    /**
     * @var \PHPExcel_Worksheet_Drawing | \PHPExcel_Worksheet_HeaderFooterDrawing
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
     * XlsDrawingWrapper constructor.
     * @param array $context
     * @param Twig_Environment $environment
     * @param XlsSheetWrapper $sheetWrapper
     * @param XlsHeaderFooterWrapper $headerFooterWrapper
     */
    public function __construct(array $context, Twig_Environment $environment, XlsSheetWrapper $sheetWrapper, XlsHeaderFooterWrapper $headerFooterWrapper)
    {
        $this->context = $context;
        $this->environment = $environment;
        $this->sheetWrapper = $sheetWrapper;
        $this->headerFooterWrapper = $headerFooterWrapper;

        $this->object = null;
        $this->attributes = [];
        $this->mappings = [];

        $this->initializeMappings();
    }

    protected function initializeMappings()
    {
        $this->mappings['coordinates'] = function ($value) {
            $this->object->setCoordinates($value);
        };
        $this->mappings['description'] = function ($value) {
            $this->object->setDescription($value);
        };
        $this->mappings['height'] = function ($value) {
            $this->object->setHeight($value);
        };
        $this->mappings['name'] = function ($value) {
            $this->object->setName($value);
        };
        $this->mappings['offsetX'] = function ($value) {
            $this->object->setOffsetX($value);
        };
        $this->mappings['offsetY'] = function ($value) {
            $this->object->setOffsetY($value);
        };
        $this->mappings['resizeProportional'] = function ($value) {
            $this->object->setResizeProportional($value);
        };
        $this->mappings['rotation'] = function ($value) {
            $this->object->setRotation($value);
        };
        $this->mappings['shadow']['alignment'] = function ($value) {
            $this->object->getShadow()->setAlignment($value);
        };
        $this->mappings['shadow']['alpha'] = function ($value) {
            $this->object->getShadow()->setAlpha($value);
        };
        $this->mappings['shadow']['blurRadius'] = function ($value) {
            $this->object->getShadow()->setBlurRadius($value);
        };
        $this->mappings['shadow']['color'] = function ($value) {
            $this->object->getShadow()->getColor()->setRGB($value);
        };
        $this->mappings['shadow']['direction'] = function ($value) {
            $this->object->getShadow()->setDirection($value);
        };
        $this->mappings['shadow']['distance'] = function ($value) {
            $this->object->getShadow()->setDistance($value);
        };
        $this->mappings['shadow']['visible'] = function ($value) {
            $this->object->getShadow()->setVisible($value);
        };
        $this->mappings['width'] = function ($value) {
            $this->object->setWidth($value);
        };
    }

    /**
     * @param $path
     * @param array|null $properties
     * @throws \PHPExcel_Exception
     */
    public function start($path, array $properties = null)
    {
        if ($this->sheetWrapper->getObject() === null) {
            throw new \LogicException();
        }

        // create local copy of the asset
        $tempPath = $this->createTempCopy($path);

        // add to header/footer
        if ($this->headerFooterWrapper->getObject()) {
            $headerFooterAttributes = $this->headerFooterWrapper->getAttributes();
            $location = '';

            switch (strtolower($this->headerFooterWrapper->getAlignmentAttributes()['type'])) {
                case 'left':
                    $location .= 'L';
                    $headerFooterAttributes['value']['left'] .= '&G';
                    break;
                case 'center':
                    $location .= 'C';
                    $headerFooterAttributes['value']['center'] .= '&G';
                    break;
                case 'right':
                    $location .= 'R';
                    $headerFooterAttributes['value']['right'] .= '&G';
                    break;
                default:
                    throw new \InvalidArgumentException(sprintf('Unknown alignment type "%s"', $this->headerFooterWrapper->getAlignmentAttributes()['type']));
            }

            switch (strtolower($headerFooterAttributes['type'])) {
                case 'header':
                case 'oddheader':
                case 'evenheader':
                case 'firstheader':
                    $location .= 'H';
                    break;
                case 'footer':
                case 'oddfooter':
                case 'evenfooter':
                case 'firstfooter':
                    $location .= 'F';
                    break;
                default:
                    throw new \InvalidArgumentException(sprintf('Unknown type "%s"', $headerFooterAttributes['type']));
            }

            $this->object = new \PHPExcel_Worksheet_HeaderFooterDrawing();
            $this->object->setPath($tempPath);
            $this->headerFooterWrapper->getObject()->addImage($this->object, $location);
            $this->headerFooterWrapper->setAttributes($headerFooterAttributes);
        }

        // add to worksheet
        else {
            $this->object = new \PHPExcel_Worksheet_Drawing();
            $this->object->setWorksheet($this->sheetWrapper->getObject());
            $this->object->setPath($tempPath);
        }

        if ($properties !== null) {
            $this->setProperties($properties, $this->mappings);
        }
    }

    public function end()
    {
        $this->object = null;
        $this->attributes = [];
    }

    //
    // Helpers
    //

    /**
     * @param $path
     * @return string
     */
    private function createTempCopy($path)
    {
        // create temp path
        $pathExtension = pathinfo($path, PATHINFO_EXTENSION);
        $tempPath = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'xlsdrawing' . '_' . md5($path) . ($pathExtension ? '.' . $pathExtension : '');

        // create local copy
        if (!file_exists($tempPath)) {
            $data = file_get_contents($path);
            if ($data === false) {
                throw new \InvalidArgumentException($path . ' does not exist.');
            }
            $temp = fopen($tempPath, 'w+');
            if ($temp === false) {
                throw new \RuntimeException('Cannot open ' . $tempPath);
            }
            fwrite($temp, $data);
            if (fclose($temp) === false) {
                throw new \RuntimeException('Cannot close ' . $tempPath);
            }
            unset($data, $temp);
        }

        return $tempPath;
    }

    //
    // Getters/Setters
    //

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

    /**
     * @return \PHPExcel_Worksheet_Drawing|\PHPExcel_Worksheet_HeaderFooterDrawing
     */
    public function getObject()
    {
        return $this->object;
    }

    /**
     * @param \PHPExcel_Worksheet_Drawing|\PHPExcel_Worksheet_HeaderFooterDrawing $object
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
}
