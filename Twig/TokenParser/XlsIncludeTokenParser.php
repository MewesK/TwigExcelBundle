<?php

namespace MewesK\TwigExcelBundle\Twig\TokenParser;

use MewesK\TwigExcelBundle\Twig\Node\XlsIncludeNode;
use MewesK\TwigExcelBundle\Twig\NodeHelper;
use Twig_Node;
use Twig_Token;

/**
 * Class XlsIncludeTokenParser
 *
 * @package MewesK\TwigExcelBundle\Twig\TokenParser
 */
class XlsIncludeTokenParser extends AbstractTokenParser
{
    /**
     * @param Twig_Token $token
     *
     * @return Twig_Node
     * @throws \Twig_Error_Syntax
     */
    public function parse(Twig_Token $token)
    {
        // parse body
        $this->parser->getStream()->expect(Twig_Token::BLOCK_END_TYPE);
        $body = $this->parseBody();
        NodeHelper::removeTextNodesRecursively($body, $this->parser);
        NodeHelper::fixMacroCallsRecursively($body);

        // return node
        return $body;
    }

    /**
     * @return string
     */
    public function getTag()
    {
        return 'xlsinclude';
    }
}
