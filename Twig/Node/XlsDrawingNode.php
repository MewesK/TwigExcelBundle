<?php

namespace MewesK\TwigExcelBundle\Twig\Node;

use Twig_Compiler;
use Twig_Node;
use Twig_Node_Expression;

/**
 * Class XlsDrawingNode
 *
 * @package MewesK\TwigExcelBundle\Twig\Node
 */
class XlsDrawingNode extends Twig_Node
{
    /**
     * @param Twig_Node_Expression $path
     * @param Twig_Node_Expression $properties
     * @param int $line
     * @param string $tag
     */
    public function __construct(Twig_Node_Expression $path, Twig_Node_Expression $properties, $line = 0, $tag = 'xlsdrawing')
    {
        parent::__construct(['path' => $path, 'properties' => $properties], [], $line, $tag);
    }

    /**
     * @param Twig_Compiler $compiler
     */
    public function compile(Twig_Compiler $compiler)
    {
        $compiler->addDebugInfo($this)
            ->write('$drawingPath = ')
            ->subcompile($this->getNode('path'))
            ->raw(';' . PHP_EOL)
            ->write('$drawingProperties = ')
            ->subcompile($this->getNode('properties'))
            ->raw(';' . PHP_EOL)
            ->write('$phpExcel->startDrawing($drawingPath, $drawingProperties);' . PHP_EOL)
            ->write('unset($drawingPath, $drawingProperties);' . PHP_EOL)
            ->write('$phpExcel->endDrawing();' . PHP_EOL);
    }
}
