<?php

namespace MewesK\PhpExcelTwigExtensionBundle\Twig;

class XlsDocumentNode extends \Twig_Node
{
    public function __construct(\Twig_Node_Expression $properties, \Twig_NodeInterface $body, $lineno, $tag = 'xlsdocument')
    {
        parent::__construct(['properties' => $properties, 'body' => $body], [], $lineno, $tag);
    }

    public function compile(\Twig_Compiler $compiler)
    {
        $compiler
            ->addDebugInfo($this)
            ->write('$arrDocumentProperties = ')
            ->subcompile($this->getNode('properties'), true)
            ->raw(';'.PHP_EOL)

            ->write('
                $objPHPExcel = new PHPExcel();
                $objPHPExcel->removeSheetByIndex(0);

                if (count($arrDocumentProperties) > 0) {
                    $objDocumentProperties = new PHPExcel_DocumentProperties();

                    if (array_key_exists(\'category\', $arrDocumentProperties)) { $objDocumentProperties->setCategory($arrDocumentProperties[\'category\']); }
                    if (array_key_exists(\'company\', $arrDocumentProperties)) { $objDocumentProperties->setCompany($arrDocumentProperties[\'company\']); }
                    if (array_key_exists(\'created\', $arrDocumentProperties)) { $objDocumentProperties->setCreated($arrDocumentProperties[\'created\']); }
                    if (array_key_exists(\'creator\', $arrDocumentProperties)) { $objDocumentProperties->setCreator($arrDocumentProperties[\'creator\']); }
                    if (array_key_exists(\'defaultStyle\', $arrDocumentProperties)) { $objPHPExcel->getDefaultStyle()->applyFromArray($arrDocumentProperties[\'defaultStyle\']); }
                    if (array_key_exists(\'description\', $arrDocumentProperties)) { $objDocumentProperties->setDescription($arrDocumentProperties[\'description\']); }
                    if (array_key_exists(\'keywords\', $arrDocumentProperties)) { $objDocumentProperties->setKeywords($arrDocumentProperties[\'keywords\']); }
                    if (array_key_exists(\'lastModifiedBy\', $arrDocumentProperties)) { $objDocumentProperties->setLastModifiedBy($arrDocumentProperties[\'lastModifiedBy\']); }
                    if (array_key_exists(\'manager\', $arrDocumentProperties)) { $objDocumentProperties->setManager($arrDocumentProperties[\'manager\']); }
                    if (array_key_exists(\'modified\', $arrDocumentProperties)) { $objDocumentProperties->setModified($arrDocumentProperties[\'modified\']); }

                    if (array_key_exists(\'security\', $arrDocumentProperties)) {
                        $arrSecurity = $arrDocumentProperties[\'security\'];
                        $objSecurity = $objPHPExcel->getSecurity();

                        if (array_key_exists(\'lockRevision\', $arrSecurity)) { $objSecurity->setLockRevision($arrSecurity[\'lockRevision\']); }
                        if (array_key_exists(\'lockStructure\', $arrSecurity)) { $objSecurity->setLockStructure($arrSecurity[\'lockStructure\']); }
                        if (array_key_exists(\'lockWindows\', $arrSecurity)) { $objSecurity->setLockWindows($arrSecurity[\'lockWindows\']); }
                        if (array_key_exists(\'revisionsPassword\', $arrSecurity)) { $objSecurity->setRevisionsPassword($arrSecurity[\'revisionsPassword\']); }
                        if (array_key_exists(\'workbookPassword\', $arrSecurity)) { $objSecurity->setWorkbookPassword($arrSecurity[\'workbookPassword\']); }

                        unset($arrSecurity, $objSecurity);
                    }

                    if (array_key_exists(\'subject\', $arrDocumentProperties)) { $objDocumentProperties->setSubject($arrDocumentProperties[\'subject\']); }
                    if (array_key_exists(\'title\', $arrDocumentProperties)) { $objDocumentProperties->setTitle($arrDocumentProperties[\'title\']); }

                    $objPHPExcel->setProperties($objDocumentProperties);

                    unset($objDocumentProperties);
                }

                unset($arrDocumentProperties);
            ')

            ->subcompile($this->getNode('body'), true)

            ->write('
                $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, \'Excel5\');
                $objWriter->save(\'php://output\');

                unset($objWriter, $objPHPExcel);
            ');
    }
}