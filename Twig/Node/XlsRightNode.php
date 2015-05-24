<?php

namespace MewesK\TwigExcelBundle\Twig\Node;

use Twig_Compiler;
use Twig_Node;

/**
 * Class XlsRightNode
 *
 * @package MewesK\TwigExcelBundle\Twig\Node
 */
class XlsRightNode extends Twig_Node
{
    /**
     * @param Twig_Node $body
     * @param int $line
     * @param string $tag
     */
    public function __construct(Twig_Node $body, $line = 0, $tag = 'xlsright')
    {
        parent::__construct(['body' => $body], [], $line, $tag);
    }

    /**
     * @param Twig_Compiler $compiler
     */
    public function compile(Twig_Compiler $compiler)
    {
        $compiler
            ->addDebugInfo($this)

            ->write('$phpExcel->startAlignment(\'right\');'.PHP_EOL)

            ->write("ob_start();\n")
            ->subcompile($this->getNode('body'))
            ->write('$rightValue = trim(ob_get_clean());'.PHP_EOL)

            ->write('$phpExcel->endAlignment($rightValue);'.PHP_EOL)
            ->write('unset($rightValue);'.PHP_EOL);
    }
}
