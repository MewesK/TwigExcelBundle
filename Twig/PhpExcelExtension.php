<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Twig;

use MewesK\PhpExcelTwigExtensionBundle\Twig\TokenParsers\XlsCellTokenParser;
use MewesK\PhpExcelTwigExtensionBundle\Twig\TokenParsers\XlsDocumentTokenParser;
use MewesK\PhpExcelTwigExtensionBundle\Twig\TokenParsers\XlsSheetTokenParser;
use MewesK\PhpExcelTwigExtensionBundle\Twig\TokenParsers\XlsStyleTokenParser;

class PhpExcelExtension extends \Twig_Extension
{
    /**
     * @inheritdoc
     */
    public function getFunctions()
    {
        return [
            new \Twig_SimpleFunction('xlsmergestyles', [$this, 'mergeStyles'])
        ];
    }

    /**
     * @inheritdoc
     */
    public function getTokenParsers()
    {
        return [
            new XlsDocumentTokenParser(),
            new XlsSheetTokenParser(),
            new XlsCellTokenParser(),
            new XlsStyleTokenParser()
        ];
    }

    /**
     * @inheritdoc
     */
    public function getName()
    {
        return 'phpexcel_extension';
    }

    public function mergeStyles(array $style1, array $style2) {
        if (!is_array($style1) || !is_array($style2)) {
            throw new \Twig_Error_Runtime('The xlsmergestyles function only works with arrays or hashes.');
        }

        return array_merge_recursive($style1, $style2);
    }
} 