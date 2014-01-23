<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Twig\TokenParsers;

use MewesK\PhpExcelTwigExtensionBundle\Twig\Nodes\XlsStyleNode;

class XlsStyleTokenParser extends \Twig_TokenParser
{
    public function parse(\Twig_Token $token)
    {
        $coordinates = $this->parser->getExpressionParser()->parseExpression();
        $properties = $this->parser->getExpressionParser()->parseExpression();
        $this->parser->getStream()->expect(\Twig_Token::BLOCK_END_TYPE);

        return new XlsStyleNode($coordinates, $properties, $token->getLine(), $this->getTag());
    }

    public function getTag()
    {
        return 'xlsstyle';
    }
}