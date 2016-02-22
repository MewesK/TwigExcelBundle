<?php

namespace MewesK\TwigExcelBundle\Twig\Node;

use Twig_Compiler;
use Twig_Node_Expression;

/**
 * Class XlsDrawingNode
 *
 * @package MewesK\TwigExcelBundle\Twig\Node
 */
class XlsDrawingNode extends XlsNode
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
            ->write('$context[\'phpExcel\']->startDrawing($drawingPath, $drawingProperties);' . PHP_EOL)
            ->write('unset($drawingPath, $drawingProperties);' . PHP_EOL)
            ->write('$context[\'phpExcel\']->endDrawing();' . PHP_EOL);
    }

    /**
     * @return string[]
     */
    public function getAllowedParents()
    {
        return [
            'MewesK\TwigExcelBundle\Twig\Node\XlsSheetNode',
            'MewesK\TwigExcelBundle\Twig\Node\XlsLeftNode',
            'MewesK\TwigExcelBundle\Twig\Node\XlsCenterNode',
            'MewesK\TwigExcelBundle\Twig\Node\XlsRightNode'
        ];
    }

    /**
     * @return bool
     */
    public function canContainText()
    {
        return false;
    }
}
