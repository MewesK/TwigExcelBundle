<?php

namespace MewesK\TwigExcelBundle\Tests\Functional;

use InvalidArgumentException;
use MewesK\TwigExcelBundle\Tests\Kernel\AppKernel;
use PHPExcel_Reader_Excel2007;
use PHPExcel_Reader_Excel5;
use PHPExcel_Reader_OOCalc;
use Symfony\Bundle\FrameworkBundle\Routing\Router;
use Symfony\Bundle\FrameworkBundle\Test\WebTestCase;
use Symfony\Component\BrowserKit\Client;
use Symfony\Component\Filesystem\Filesystem;

/**
 * Class AbstractControllerTest
 * 
 * @package MewesK\TwigExcelBundle\Tests\Functional
 */
abstract class AbstractControllerTest extends WebTestCase
{
    protected static $CONFIG_FILE;
    protected static $TEMP_PATH = '/../../tmp/functional/';

    /**
     * @var Filesystem
     */
    protected static $fileSystem;
    /**
     * @var Client
     */
    protected static $client;
    /**
     * @var Router
     */
    protected static $router;

    //
    // Helper
    //

    /**
     * @param $uri
     * @param $format
     * @return \PHPExcel
     * @throws \InvalidArgumentException
     * @throws \Symfony\Component\Filesystem\Exception\IOException
     */
    protected function getDocument($uri, $format = 'xlsx')
    {
        // generate source
        static::$client->request('GET', $uri);
        $source = static::$client->getResponse()->getContent();
        
        // create paths
        $tempDirPath = __DIR__ . static::$TEMP_PATH;
        $tempFilePath = $tempDirPath . 'simple' . '.' . $format;

        // save source
        static::$fileSystem->dumpFile($tempFilePath, $source);

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
            default:
                throw new InvalidArgumentException();
        }

        return $reader->load(__DIR__ . static::$TEMP_PATH . 'simple' . '.' . $format);
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
     * @throws \RuntimeException
     */
    protected static function createKernel(array $options = [])
    {
        return static::$kernel = new AppKernel(
            array_key_exists('config', $options) && is_string($options['config']) ? $options['config'] : 'config.yml'
        );
    }

    /**
     * {@inheritdoc}
     * @throws \Exception
     */
    protected function setUp()
    {
        static::$client = static::createClient(['config' => static::$CONFIG_FILE]);
        static::$router = static::$kernel->getContainer()->get('router');
    }

    /**
     * {@inheritdoc}
     * @throws \Exception
     */
    public static function setUpBeforeClass()
    {
        static::$fileSystem = new Filesystem();
        static::$fileSystem->remove(__DIR__ . static::$TEMP_PATH);
    }

    /**
     * {@inheritdoc}
     * @throws \Symfony\Component\Filesystem\Exception\IOException
     */
    public static function tearDownAfterClass()
    {
        if (in_array(getenv('DELETE_TEMP_FILES'), ['true', '1', 1, true], true)) {
            static::$fileSystem->remove(__DIR__ . static::$TEMP_PATH);
        }
    }
}
