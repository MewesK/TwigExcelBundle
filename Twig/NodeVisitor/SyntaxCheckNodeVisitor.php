<?php

namespace MewesK\TwigExcelBundle\Twig\NodeVisitor;

use MewesK\TwigExcelBundle\Twig\Node\XlsNode;
use Twig_BaseNodeVisitor;
use Twig_Environment;
use Twig_Error_Syntax;
use Twig_Node;

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
        // TODO warn if using normal blocks

        if ($node instanceof XlsNode) {
            // check allowed parents
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

        throw new Twig_Error_Syntax(sprintf('Node "%s" is not allowed inside of Node "%s".', get_class($node), $parentName));
    }
}
