<?php

namespace MewesK\TwigExcelBundle\Twig\Node;

/**
 * Class XlsNode
 *
 * @package MewesK\TwigExcelBundle\Twig\Node
 */
abstract class XlsNode extends \Twig_Node
{
    /**
     * @return string[]
     */
    public abstract function getAllowedParents();

    /**
     * @return bool
     */
    public abstract function canContainText();
}
