<?php

namespace MewesK\TwigExcelBundle\DependencyInjection;

use Symfony\Component\Config\Definition\Builder\TreeBuilder;
use Symfony\Component\Config\Definition\ConfigurationInterface;

/**
 * Class Configuration
 * @package MewesK\TwigExcelBundle\DependencyInjection
 */
class Configuration implements ConfigurationInterface
{
    /**
     * {@inheritDoc}
     */
    public function getConfigTreeBuilder()
    {
        $treeBuilder = new TreeBuilder();
        $rootNode = $treeBuilder->root('mewes_k_twig_excel');

        $rootNode
            ->children()
                ->booleanNode('pre_calculate_formulas')->defaultTrue()->info('Pre-calculating formulas can be slow in certain cases. Disabling this option can improve the performance but the resulting documents won\'t show the result of any formulas when opened in a external spreadsheet software.')->end()
                ->scalarNode('disk_caching_directory')->defaultNull()->info('Using disk caching can improve memory consumption by writing data to disk temporary. Works only for .XLSX and .ODS documents.')->example('/tmp')->end()
            ->end();

        return $treeBuilder;
    }
}
