<?php

namespace MewesK\TwigExcelBundle\Twig\NodeVisitor;

use MewesK\TwigExcelBundle\Twig\Node\XlsNode;
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
        if ($node instanceof Twig_Node_Block) {
            if ($this->checkContainsXlsNode($node) && $node->hasAttribute('twigExcelBundle')) {
                throw new Twig_Error_Syntax('Block tags do not work together with Twig tags provided by TwigExcelBundle. Please use \'xlsblock\' instead.');
            }
        }
        elseif ($node instanceof Twig_Node_Macro && $node->hasAttribute('twigExcelBundle')) {
            if ($this->checkContainsXlsNode($node)) {
                throw new Twig_Error_Syntax('Macro tags do not work together with Twig tags provided by TwigExcelBundle. Please use \'xlsmacro\' instead.');
            }
        }
        elseif ($node instanceof XlsNode) {
            /**
             * @var XlsNode $node
             */
            $this->checkAllowedParents($node, $node->getAllowedParents());
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

    //
    // Helper
    //

    /**
     * @param Twig_Node $node
     * @param array $allowedParents
     *
     * @throws Twig_Error_Syntax
     */
    protected function checkAllowedParents(Twig_Node $node, array $allowedParents)
    {
        $parentName = null;

        foreach (array_reverse($this->path) as $className) {
            if (strpos($className, 'MewesK\TwigExcelBundle\Twig\Node\Xls') === 0) {
                $parentName = $className;
                break;
            }
        }

        if ($parentName === null) {
            return;
        }

        foreach ($allowedParents as $className) {
            if ($className === $parentName) {
                return;
            }
        }

        $this->path = [];
        throw new Twig_Error_Syntax(sprintf('Node "%s" is not allowed inside of Node "%s".', get_class($node), $parentName));
    }

    /**
     * @param Twig_Node $node
     * @return bool
     */
    private function checkContainsXlsNode(Twig_Node $node)
    {
        foreach ($node->getIterator() as $key => $subNode) {
            if ($node instanceof XlsNode) {
                return true;
            } elseif ($subNode instanceof Twig_Node && $subNode->count() > 0) {
                if ($this->checkContainsXlsNode($subNode)) {
                    return true;
                }
            }
        }
        return false;
    }
}
