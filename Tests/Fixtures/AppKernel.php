<?php

use Symfony\Component\Filesystem\Filesystem;
use Symfony\Component\Config\Loader\LoaderInterface;
use Symfony\Component\HttpKernel\Kernel;

/**
 * Class AppKernel
 *
 * @package MewesK\TwigExcelBundle\Tests\Fixtures
 */
class AppKernel extends Kernel
{
    /**
     * @var string
     */
    private $config;

    /**
     * @param string $config
     * @throws \RuntimeException
     */
    public function __construct($config)
    {
        parent::__construct('test', true);

        $fs = new Filesystem();
        if (!$fs->isAbsolutePath($config)) {
            $config = __DIR__ . '/config/' .$config;
        }

        if (!file_exists($config)) {
            throw new \RuntimeException(sprintf('The config file "%s" does not exist.', $config));
        }

        $this->config = $config;
    }

    /**
     * @return array
     */
    public function registerBundles()
    {
        return [
            new Symfony\Bundle\FrameworkBundle\FrameworkBundle(),
            new Symfony\Bundle\TwigBundle\TwigBundle(),
            new Sensio\Bundle\FrameworkExtraBundle\SensioFrameworkExtraBundle(),
            new MewesK\TwigExcelBundle\MewesKTwigExcelBundle(),
            new MewesK\TwigExcelBundle\Tests\Fixtures\TestBundle\TestBundle()
        ];
    }

    /**
     * @param LoaderInterface $loader
     * @throws \Exception
     */
    public function registerContainerConfiguration(LoaderInterface $loader)
    {
        $loader->load($this->config);
    }

    /**
     * @return string
     */
    public function getCacheDir()
    {
        return __DIR__ . '/../../tmp/cache';
    }

    /**
     * @return string
     */
    public function getLogDir()
    {
        return __DIR__ . '/../../tmp/logs';
    }
}
