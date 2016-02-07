<?php

namespace MewesK\TwigExcelBundle\Twig\Node;

interface XlsNode
{
    /**
     * @return string[]
     */
    public function getAllowedParents();
}
