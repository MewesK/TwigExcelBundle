<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Twig\Node;

class XlsDrawingNode extends \Twig_Node
{
    public function __construct(\Twig_Node_Expression $path, \Twig_Node_Expression $properties, $line, $tag = 'xlsdrawing')
    {
        parent::__construct(['path' => $path, 'properties' => $properties], [], $line, $tag);
    }

    public function compile(\Twig_Compiler $compiler)
    {
        $compiler
            ->addDebugInfo($this)

            ->write('$drawingPath = ')
            ->subcompile($this->getNode('path'))
            ->raw(';'.PHP_EOL)

            ->write('$drawingProperties = ')
            ->subcompile($this->getNode('properties'))
            ->raw(';'.PHP_EOL)

            ->write('$phpExcel->tagSheet($drawingPath, $drawingProperties);'.PHP_EOL)
            ->write('unset($drawingPath, $drawingProperties);'.PHP_EOL);
    }
}