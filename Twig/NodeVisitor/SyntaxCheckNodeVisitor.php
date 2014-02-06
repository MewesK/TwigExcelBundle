<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Twig\NodeVisitor;

use MewesK\PhpExcelTwigExtensionBundle\Twig\Node\XlsCellNode;
use MewesK\PhpExcelTwigExtensionBundle\Twig\Node\XlsDocumentNode;
use MewesK\PhpExcelTwigExtensionBundle\Twig\Node\XlsDrawingNode;
use MewesK\PhpExcelTwigExtensionBundle\Twig\Node\XlsFooterNode;
use MewesK\PhpExcelTwigExtensionBundle\Twig\Node\XlsHeaderNode;
use MewesK\PhpExcelTwigExtensionBundle\Twig\Node\XlsRowNode;
use MewesK\PhpExcelTwigExtensionBundle\Twig\Node\XlsSheetNode;
use Twig_Environment;
use Twig_Error_Syntax;
use Twig_NodeInterface;
use Twig_NodeVisitorInterface;

class SyntaxCheckNodeVisitor implements Twig_NodeVisitorInterface {

    protected $lastDocument = null;
    protected $lastSheet = null;
    protected $lastFooter = null;
    protected $lastHeader = null;
    protected $lastRow = null;
    protected $lastCell = null;
    protected $lastDrawing = null;

    /**
     * Called before child nodes are visited.
     *
     * @param Twig_NodeInterface $node The node to visit
     * @param Twig_Environment $env The Twig environment instance
     *
     * @return Twig_NodeInterface The modified node
     * @throws Twig_Error_Syntax
     */
    public function enterNode(Twig_NodeInterface $node, Twig_Environment $env)
    {
        if ($node instanceof XlsDocumentNode) {
            if ($this->lastDocument) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed inside of Node "%s".', get_class($node), get_class($this->lastDocument))
                );
            }
            if ($this->lastSheet) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed inside of Node "%s".', get_class($node), get_class($this->lastSheet))
                );
            }
            if ($this->lastRow) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed inside of Node "%s".', get_class($node), get_class($this->lastRow))
                );
            }
            if ($this->lastCell) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed inside of Node "%s".', get_class($node), get_class($this->lastCell))
                );
            }
            $this->lastDocument = $node;
        }
        elseif ($node instanceof XlsSheetNode) {
            if (!$this->lastDocument) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed outside of Node "%s".', get_class($node), get_class($this->lastDocument))
                );
            }
            if ($this->lastSheet) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed inside of Node "%s".', get_class($node), get_class($this->lastSheet))
                );
            }
            if ($this->lastRow) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed inside of Node "%s".', get_class($node), get_class($this->lastRow))
                );
            }
            if ($this->lastCell) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed inside of Node "%s".', get_class($node), get_class($this->lastCell))
                );
            }
            $this->lastSheet = $node;
        }
        elseif ($node instanceof XlsFooterNode) {
            if (!$this->lastDocument) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed outside of Node "%s".', get_class($node), get_class($this->lastDocument))
                );
            }
            if (!$this->lastSheet) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed outside of Node "%s".', get_class($node), get_class($this->lastSheet))
                );
            }
            if ($this->lastFooter) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed inside of Node "%s".', get_class($node), get_class($this->lastFooter))
                );
            }
            if ($this->lastHeader) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed inside of Node "%s".', get_class($node), get_class($this->lastHeader))
                );
            }
            if ($this->lastRow) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed inside of Node "%s".', get_class($node), get_class($this->lastRow))
                );
            }
            if ($this->lastCell) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed inside of Node "%s".', get_class($node), get_class($this->lastCell))
                );
            }
            $this->lastFooter = $node;
        }
        elseif ($node instanceof XlsHeaderNode) {
            if (!$this->lastDocument) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed outside of Node "%s".', get_class($node), get_class($this->lastDocument))
                );
            }
            if (!$this->lastSheet) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed outside of Node "%s".', get_class($node), get_class($this->lastSheet))
                );
            }
            if ($this->lastFooter) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed inside of Node "%s".', get_class($node), get_class($this->lastFooter))
                );
            }
            if ($this->lastHeader) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed inside of Node "%s".', get_class($node), get_class($this->lastHeader))
                );
            }
            if ($this->lastRow) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed inside of Node "%s".', get_class($node), get_class($this->lastRow))
                );
            }
            if ($this->lastCell) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed inside of Node "%s".', get_class($node), get_class($this->lastCell))
                );
            }
            $this->lastHeader = $node;
        }
        elseif ($node instanceof XlsRowNode) {
            if (!$this->lastDocument) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed outside of Node "%s".', get_class($node), get_class($this->lastDocument))
                );
            }
            if (!$this->lastSheet) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed outside of Node "%s".', get_class($node), get_class($this->lastSheet))
                );
            }
            if ($this->lastRow) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed inside of Node "%s".', get_class($node), get_class($this->lastRow))
                );
            }
            if ($this->lastCell) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed inside of Node "%s".', get_class($node), get_class($this->lastCell))
                );
            }
            $this->lastRow = $node;
        }
        elseif ($node instanceof XlsCellNode) {
            if (!$this->lastDocument) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed outside of Node "%s".', get_class($node), get_class($this->lastDocument))
                );
            }
            if (!$this->lastSheet) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed outside of Node "%s".', get_class($node), get_class($this->lastSheet))
                );
            }
            if (!$this->lastRow) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed outside of Node "%s".', get_class($node), get_class($this->lastRow))
                );
            }
            if ($this->lastCell) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed inside of Node "%s".', get_class($node), get_class($this->lastCell))
                );
            }
            $this->lastCell = $node;
        }
        elseif ($node instanceof XlsDrawingNode) {
            if (!$this->lastDocument) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed outside of Node "%s".', get_class($node), get_class($this->lastDocument))
                );
            }
            if (!$this->lastSheet) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed outside of Node "%s".', get_class($node), get_class($this->lastSheet))
                );
            }
            if ($this->lastRow) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed inside of Node "%s".', get_class($node), get_class($this->lastRow))
                );
            }
            if ($this->lastCell) {
                throw new Twig_Error_Syntax(
                    sprintf('Node "%s" is not allowed inside of Node "%s".', get_class($node), get_class($this->lastCell))
                );
            }
            $this->lastDrawing = $node;
        }
        return $node;
    }

    /**
     * Called after child nodes are visited.
     *
     * @param Twig_NodeInterface $node The node to visit
     * @param Twig_Environment $env The Twig environment instance
     *
     * @return Twig_NodeInterface|false The modified node or false if the node must be removed
     */
    public function leaveNode(Twig_NodeInterface $node, Twig_Environment $env)
    {
        if ($node instanceof XlsDocumentNode) {
            $this->lastDocument = null;
        }
        elseif ($node instanceof XlsSheetNode) {
            $this->lastSheet = null;
        }
        elseif ($node instanceof XlsFooterNode) {
            $this->lastFooter = null;
        }
        elseif ($node instanceof XlsHeaderNode) {
            $this->lastHeader = null;
        }
        elseif ($node instanceof XlsRowNode) {
            $this->lastRow = null;
        }
        elseif ($node instanceof XlsCellNode) {
            $this->lastCell = null;
        }
        elseif ($node instanceof XlsDrawingNode) {
            $this->lastDrawing = null;
        }
        return $node;
    }

    /**
     * Returns the priority for this visitor.
     *
     * Priority should be between -10 and 10 (0 is the default).
     *
     * @return integer The priority level
     */
    public function getPriority()
    {
        return 0;
    }
}