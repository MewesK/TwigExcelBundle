<?php

namespace MewesK\TwigExcelBundle\Twig\TokenParser;

use MewesK\TwigExcelBundle\Twig\TokenParser\Traits\FixMacroCallsTrait;
use MewesK\TwigExcelBundle\Twig\TokenParser\Traits\RemoveTextNodeTrait;
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
    use FixMacroCallsTrait;
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
        $block = $this->parser->getBlock($blockReference->getAttribute('name'));
        $block->setAttribute('twigExcelBundle', true);

        $this->removeTextNodesRecursively($block);
        $this->fixMacroCallsRecursively($block);

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
