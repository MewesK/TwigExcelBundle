<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Twig\Nodes;

class XlsCellNode extends \Twig_Node
{
    public function __construct(\Twig_Node_Expression $coordinates, \Twig_Node_Expression $properties, \Twig_NodeInterface $body, $lineno, $tag = 'xlscell')
    {
        parent::__construct(['coordinates' => $coordinates, 'properties' => $properties, 'body' => $body], [], $lineno, $tag);
    }

    public function compile(\Twig_Compiler $compiler)
    {
        $compiler
            ->addDebugInfo($this)
            ->write('$strCoordinates = ')
            ->subcompile($this->getNode('coordinates'), true)
            ->raw(';'.PHP_EOL)
            ->write("ob_start();\n")
            ->subcompile($this->getNode('body'), true)
            ->write('$objActiveSheet->setCellValue($strCoordinates, trim(ob_get_clean()));'.PHP_EOL)
            ->write('$arrCellProperties = ')
            ->subcompile($this->getNode('properties'), true)
            ->raw(';'.PHP_EOL)

            ->write('
                if ($arrCellProperties) > 0) {
                    if (array_key_exists(\'break\', $arrCellProperties)) { $objActiveSheet->setBreak($strCoordinates, $arrCellProperties[\'break\']); }

                    if (array_key_exists(\'dataValidation\', $arrCellProperties)) {
                        $arrValidation = $arrCellProperties[\'dataValidation\'];
                        $objValidation = $objActiveSheet->getCell($strCoordinates)->getDataValidation();;

                        if (array_key_exists(\'allowBlank\', $arrValidation)) { $objValidation->setAllowBlank($arrValidation[\'allowBlank\']); }
                        if (array_key_exists(\'error\', $arrValidation)) { $objValidation->setError($arrValidation[\'error\']); }
                        if (array_key_exists(\'errorStyle\', $arrValidation)) { $objValidation->setErrorStyle($arrValidation[\'errorStyle\']); }
                        if (array_key_exists(\'errorTitle\', $arrValidation)) { $objValidation->setErrorTitle($arrValidation[\'errorTitle\']); }
                        if (array_key_exists(\'formula1\', $arrValidation)) { $objSecurity->setFormula1($arrValidation[\'formula1\']); }
                        if (array_key_exists(\'formula2\', $arrValidation)) { $objValidation->setFormula2($arrValidation[\'formula2\']); }
                        if (array_key_exists(\'operator\', $arrValidation)) { $objValidation->setOperator($arrValidation[\'operator\']); }
                        if (array_key_exists(\'prompt\', $arrValidation)) { $objValidation->setPrompt($arrValidation[\'prompt\']); }
                        if (array_key_exists(\'promptTitle\', $arrValidation)) { $objValidation->setPromptTitle($arrValidation[\'promptTitle\']); }
                        if (array_key_exists(\'showDropDown\', $arrValidation)) { $objValidation->setShowDropDown($arrValidation[\'showDropDown\']); }
                        if (array_key_exists(\'showErrorMessage\', $arrValidation)) { $objValidation->setShowErrorMessage($arrValidation[\'showErrorMessage\']); }
                        if (array_key_exists(\'showInputMessage\', $arrValidation)) { $objValidation->setShowInputMessage($arrValidation[\'showInputMessage\']); }
                        if (array_key_exists(\'type\', $arrValidation)) { $objValidation->setType($arrValidation[\'type\']); }

                        unset($arrValidation, $objValidation);
                    }

                    if (array_key_exists(\'style\', $arrCellProperties)) { $objActiveSheet->getStyle($strCoordinates)->applyFromArray($arrCellProperties[\'style\']); }
                    if (array_key_exists(\'url\', $arrCellProperties)) { $objActiveSheet->getCell($strCoordinates)->getHyperlink()->setUrl($arrCellProperties[\'url\']); }
                }

                unset($strCoordinates, $arrCellProperties);
            ');
    }
}