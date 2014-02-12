<?php

namespace MewesK\TwigExcelBundle\Twig\NodeVisitor;

use MewesK\TwigExcelBundle\Twig\Node\XlsCellNode;
use MewesK\TwigExcelBundle\Twig\Node\XlsCenterNode;
use MewesK\TwigExcelBundle\Twig\Node\XlsDocumentNode;
use MewesK\TwigExcelBundle\Twig\Node\XlsDrawingNode;
use MewesK\TwigExcelBundle\Twig\Node\XlsFooterNode;
use MewesK\TwigExcelBundle\Twig\Node\XlsHeaderNode;
use MewesK\TwigExcelBundle\Twig\Node\XlsLeftNode;
use MewesK\TwigExcelBundle\Twig\Node\XlsRightNode;
use MewesK\TwigExcelBundle\Twig\Node\XlsRowNode;
use MewesK\TwigExcelBundle\Twig\Node\XlsSheetNode;
use Twig_Environment;
use Twig_Error_Syntax;
use Twig_NodeInterface;
use Twig_NodeVisitorInterface;

class SyntaxCheckNodeVisitor implements Twig_NodeVisitorInterface {

    protected $path = array();

    /**
     * {@inheritdoc}
     */
    public function enterNode(Twig_NodeInterface $node, Twig_Environment $env)
    {
        if ($node instanceof XlsDocumentNode) {
            $this->checkAllowedParents($node, array());
        }
        elseif ($node instanceof XlsSheetNode) {
            $this->checkAllowedParents($node, array(
                'MewesK\TwigExcelBundle\Twig\Node\XlsDocumentNode'
            ));
        }
        elseif ($node instanceof XlsRowNode || $node instanceof XlsFooterNode || $node instanceof XlsHeaderNode) {
            $this->checkAllowedParents($node, array(
                'MewesK\TwigExcelBundle\Twig\Node\XlsSheetNode'
            ));
        }
        elseif ($node instanceof XlsLeftNode || $node instanceof XlsCenterNode || $node instanceof XlsRightNode) {
            $this->checkAllowedParents($node, array(
                'MewesK\TwigExcelBundle\Twig\Node\XlsFooterNode',
                'MewesK\TwigExcelBundle\Twig\Node\XlsHeaderNode'
            ));
        }
        elseif ($node instanceof XlsCellNode) {
            $this->checkAllowedParents($node, array(
                'MewesK\TwigExcelBundle\Twig\Node\XlsRowNode'
            ));
        }
        elseif ($node instanceof XlsDrawingNode) {
            $this->checkAllowedParents($node, array(
                'MewesK\TwigExcelBundle\Twig\Node\XlsSheetNode',
                'MewesK\TwigExcelBundle\Twig\Node\XlsLeftNode',
                'MewesK\TwigExcelBundle\Twig\Node\XlsCenterNode',
                'MewesK\TwigExcelBundle\Twig\Node\XlsRightNode'
            ));
        }

        array_push($this->path, get_class($node));

        return $node;
    }

    /**
     * {@inheritdoc}
     */
    public function leaveNode(Twig_NodeInterface $node, Twig_Environment $env)
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

    private function checkAllowedParents($node, array $allowedParents) {
        $parentName = null;
        foreach(array_reverse($this->path) as $className) {
            if (strpos($className, 'MewesK\TwigExcelBundle\Twig\Node\Xls') === 0) {
                $parentName = $className;
                break;
            }
        }
        if ($parentName == null && count($allowedParents) == 0) {
            return;
        }
        foreach($allowedParents as $className) {
            if ($className === $parentName) {
                return;
            }
        }
        throw new Twig_Error_Syntax(
            sprintf('Node "%s" is not allowed inside of Node "%s".', get_class($node), $parentName)
        );
    }
}