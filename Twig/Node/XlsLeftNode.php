<?php

namespace MewesK\TwigExcelBundle\Twig\Node;

use Twig_Compiler;
use Twig_Node;

/**
 * Class XlsLeftNode
 *
 * @package MewesK\TwigExcelBundle\Twig\Node
 */
class XlsLeftNode extends Twig_Node
{
    /**
     * @param Twig_Node $body
     * @param int $line
     * @param string $tag
     */
    public function __construct(Twig_Node $body, $line = 0, $tag = 'xlsleft')
    {
        parent::__construct(['body' => $body], [], $line, $tag);
    }

    /**
     * @param Twig_Compiler $compiler
     */
    public function compile(Twig_Compiler $compiler)
    {
        $compiler->addDebugInfo($this)
            ->write('$phpExcel->startAlignment(\'left\');' . PHP_EOL)
            ->write("ob_start();\n")
            ->subcompile($this->getNode('body'))
            ->write('$leftValue = trim(ob_get_clean());' . PHP_EOL)
            ->write('$phpExcel->endAlignment($leftValue);' . PHP_EOL)
            ->write('unset($leftValue);' . PHP_EOL);
    }
}
