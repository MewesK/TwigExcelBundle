<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Twig\Node;
use Twig_Compiler;
use Twig_Node;
use Twig_Node_Expression;

class XlsDrawingNode extends Twig_Node
{
    public function __construct(Twig_Node_Expression $path, Twig_Node_Expression $properties, $line, $tag = 'xlsdrawing')
    {
        parent::__construct(array('path' => $path, 'properties' => $properties), array(), $line, $tag);
    }

    public function compile(Twig_Compiler $compiler)
    {
        $compiler
            ->addDebugInfo($this)

            ->write('$drawingPath = ')
            ->subcompile($this->getNode('path'))
            ->raw(';'.PHP_EOL)

            ->write('$drawingProperties = ')
            ->subcompile($this->getNode('properties'))
            ->raw(';'.PHP_EOL)

            ->write('$phpExcel->tagDrawing($drawingPath, $drawingProperties);'.PHP_EOL)
            ->write('unset($drawingPath, $drawingProperties);'.PHP_EOL);
    }
}