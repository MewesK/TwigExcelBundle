<?php

namespace MewesK\TwigExcelBundle\Twig\Node;

/**
 * Interface SyntaxAwareNode
 *
 * @package MewesK\TwigExcelBundle\Twig\Node
 */
interface SyntaxAwareNode
{
    /**
     * @return string[]
     */
    public function getAllowedParents();

    /**
     * @return bool
     */
    public function canContainText();
}
