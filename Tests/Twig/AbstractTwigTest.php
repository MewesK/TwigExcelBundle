<?php

namespace MewesK\TwigExcelBundle\Tests\Twig;

use InvalidArgumentException;
use MewesK\TwigExcelBundle\Twig\TwigExcelExtension;
use PHPExcel_Reader_Excel2007;
use PHPExcel_Reader_Excel5;
use PHPExcel_Reader_OOCalc;
use PHPUnit_Framework_TestCase;
use Symfony\Bridge\Twig\AppVariable;
use Symfony\Component\HttpFoundation\Request;
use Symfony\Component\HttpFoundation\RequestStack;
use Twig_Environment;
use Twig_Loader_Filesystem;

/**
 * Class AbstractTwigTest
 * @package MewesK\TwigExcelBundle\Tests\Twig
 */
abstract class AbstractTwigTest extends PHPUnit_Framework_TestCase
{
    protected static $TEMP_PATH = '/../../tmp/';
    protected static $RESOURCE_PATH = '/../Resources/views/';
    protected static $TEMPLATE_PATH = '/../Resources/templates/';

    /**
     * @var Twig_Environment
     */
    protected static $environment;

    //
    // Helper
    //

    /**
     * @param string $templateName
     * @param string $format
     *
     * @return \PHPExcel
     */
    protected function getDocument($templateName, $format)
    {
        // prepare global variables
        $request = new Request();
        $request->setRequestFormat($format);

        $requestStack = new RequestStack();
        $requestStack->push($request);

        $appVariable = new AppVariable();
        $appVariable->setRequestStack($requestStack);

        // generate source from template
        $source = static::$environment->loadTemplate($templateName . '.twig')->render(['app' => $appVariable]);

        // create paths
        $tempDirPath = __DIR__ . static::$TEMP_PATH;
        $tempFilePath = $tempDirPath . $templateName . '.' . $format;

        // create source directory if necessary
        if (!file_exists($tempDirPath) && !@mkdir($tempDirPath) && !@is_dir($tempDirPath)) {
            throw new \RuntimeException('Cannot create temp folder');
        }

        // save source
        file_put_contents($tempFilePath, $source);

        // load source
        switch ($format) {
            case 'ods':
                $reader = new PHPExcel_Reader_OOCalc();
                break;
            case 'xls':
                $reader = new PHPExcel_Reader_Excel5();
                break;
            case 'xlsx':
                $reader = new PHPExcel_Reader_Excel2007();
                break;
            case 'pdf':
            case 'csv':
                return $tempFilePath;
            default:
                throw new InvalidArgumentException();
        }

        return $reader->load($tempFilePath);
    }

    //
    // PhpUnit
    //

    /**
     * @return array
     */
    abstract public function formatProvider();

    /**
     * {@inheritdoc}
     */
    public static function setUpBeforeClass()
    {
        $fileSystem = new Twig_Loader_Filesystem([__DIR__ . static::$RESOURCE_PATH]);
        $fileSystem->addPath( __DIR__ . static::$TEMPLATE_PATH, 'templates');
        static::$environment = new Twig_Environment($fileSystem, ['strict_variables' => true]);
        static::$environment->addExtension(new TwigExcelExtension());
        static::$environment->setCache(__DIR__ . static::$TEMP_PATH);
    }

    /**
     * {@inheritdoc}
     */
    public static function tearDownAfterClass()
    {
        if (in_array(getenv('DELETE_TEMP_FILES'), ['true', '1', 1, true], true)) {
            exec('rm -rf ' . __DIR__ . static::$TEMP_PATH);
        }
    }
}
