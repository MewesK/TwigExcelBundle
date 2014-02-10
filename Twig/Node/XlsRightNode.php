<?php

namespace MewesK\TwigExcelBundle\Twig\Node;

use Twig_Compiler;
use Twig_Node;
use Twig_NodeInterface;

class XlsRightNode extends Twig_Node
{
    public function __construct(Twig_NodeInterface $body, $line, $tag = 'xlsright')
    {
        parent::__construct(array('body' => $body), array(), $line, $tag);
    }

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