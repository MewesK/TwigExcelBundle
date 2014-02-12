<?php

namespace MewesK\TwigExcelBundle\Twig\Node;

use Twig_Compiler;
use Twig_Node;

class XlsLeftNode extends Twig_Node
{
    public function __construct(Twig_Node $body, $line, $tag = 'xlsleft')
    {
        parent::__construct(array('body' => $body), array(), $line, $tag);
    }

    public function compile(Twig_Compiler $compiler)
    {
        $compiler
            ->addDebugInfo($this)

            ->write('$phpExcel->startAlignment(\'left\');'.PHP_EOL)

            ->write("ob_start();\n")
            ->subcompile($this->getNode('body'))
            ->write('$leftValue = trim(ob_get_clean());'.PHP_EOL)

            ->write('$phpExcel->endAlignment($leftValue);'.PHP_EOL)
            ->write('unset($leftValue);'.PHP_EOL);
    }
}