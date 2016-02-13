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
    abstract public function getAllowedParents();

    /**
     * @return bool
     */
    abstract public function canContainText();
}
