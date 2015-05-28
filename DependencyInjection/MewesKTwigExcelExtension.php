<?php

namespace MewesK\TwigExcelBundle\DependencyInjection;

use Symfony\Component\Config\FileLocator;
use Symfony\Component\DependencyInjection\ContainerBuilder;
use Symfony\Component\DependencyInjection\Loader;
use Symfony\Component\HttpKernel\DependencyInjection\ConfigurableExtension;

/**
 * Class MewesKTwigExcelExtension
 *
 * @package MewesK\TwigExcelBundle\DependencyInjection
 */
class MewesKTwigExcelExtension extends ConfigurableExtension
{
    /**
     * {@inheritDoc}
     */
    protected function loadInternal(array $mergedConfig, ContainerBuilder $container)
    {
        $loader = new Loader\XmlFileLoader($container, new FileLocator(__DIR__ . '/../Resources/config'));
        $loader->load('services.xml');

        $container->setParameter('mewes_k_twig_excel.pre_calculate_formulas', $mergedConfig['pre_calculate_formulas']);
        $container->setParameter('mewes_k_twig_excel.disk_caching_directory', $mergedConfig['disk_caching_directory']);
    }
}
