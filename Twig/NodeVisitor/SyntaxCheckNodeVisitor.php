<?php

namespace MewesK\TwigExcelBundle\Twig\NodeVisitor;

use MewesK\TwigExcelBundle\Twig\Node\XlsNode;
use MewesK\TwigExcelBundle\Twig\NodeHelper;
use Twig_BaseNodeVisitor;
use Twig_Environment;
use Twig_Error_Syntax;
use Twig_Node;
use Twig_Node_Block;
use Twig_Node_Macro;

/**
 * Class SyntaxCheckNodeVisitor
 *
 * @package MewesK\TwigExcelBundle\Twig\NodeVisitor
 */
class SyntaxCheckNodeVisitor extends Twig_BaseNodeVisitor
{
    /**
     * @var array
     */
    protected $path = [];

    /**
     * {@inheritdoc}
     *
     * @throws Twig_Error_Syntax
     */
    protected function doEnterNode(Twig_Node $node, Twig_Environment $env)
    {
        if (($node instanceof Twig_Node_Block || $node instanceof Twig_Node_Macro) && !$node->hasAttribute('twigExcelBundle') && NodeHelper::checkContainsXlsNode($node)) {
            if ($node instanceof Twig_Node_Block) {
                throw new Twig_Error_Syntax('Block tags do not work together with Twig tags provided by TwigExcelBundle. Please use \'xlsblock\' instead.');
            }
            elseif ($node instanceof Twig_Node_Macro) {
                throw new Twig_Error_Syntax('Macro tags do not work together with Twig tags provided by TwigExcelBundle. Please use \'xlsmacro\' instead.');
            }
        }
        elseif ($node instanceof XlsNode) {
            /**
             * @var XlsNode $node
             */
            try {
                NodeHelper::checkAllowedParents($node, $this->path);
            } catch(Twig_Error_Syntax $e) {
                // reset path since throwing an error prevents doLeaveNode to be called
                $this->path = [];
                throw $e;
            }
        }

        $this->path[] = get_class($node);

        return $node;
    }

    /**
     * {@inheritdoc}
     */
    protected function doLeaveNode(Twig_Node $node, Twig_Environment $env)
    {
        array_pop($this->path);

        return $node;
    }

    /**
     * {@inheritdoc}
     */
    public function getPriority()
    {
        return 0;
    }
}
