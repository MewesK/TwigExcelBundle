<?php

namespace MewesK\TwigExcelBundle\Twig\TokenParser;

use Twig_Error_Syntax;
use Twig_Node_BlockReference;
use Twig_Token;
use Twig_TokenParser_Block;

/**
 * Class XlsBlockTokenParser
 *
 * @package MewesK\TwigExcelBundle\Twig\TokenParser
 */
class XlsBlockTokenParser extends Twig_TokenParser_Block
{
    use RemoveTextNodeTrait;

    /**
     * @param Twig_Token $token
     * @return Twig_Node_BlockReference
     * @throws Twig_Error_Syntax
     */
    public function parse(Twig_Token $token)
    {
        /**
         * @var Twig_Node_BlockReference $blockReference
         */
        $blockReference = parent::parse($token);

        $this->removeTextNodesRecursively($this->parser->getBlock($blockReference->getAttribute('name')));

        return $blockReference;
    }

    /**
     * @param Twig_Token $token
     * @return bool
     */
    public function decideBlockEnd(Twig_Token $token)
    {
        return $token->test('endxlsblock');
    }

    /**
     * @return string
     */
    public function getTag()
    {
        return 'xlsblock';
    }
}
