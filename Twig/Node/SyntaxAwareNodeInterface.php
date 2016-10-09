<?php

namespace MewesK\TwigExcelBundle\Twig\Node;

/**
 * Interface SyntaxAwareNodeInterface
 *
 * @package MewesK\TwigExcelBundle\Twig\Node
 */
interface SyntaxAwareNodeInterface
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
