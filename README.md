# TwigExcelBundle

[![Teavis CI: Build Status](https://travis-ci.org/MewesK/TwigExcelBundle.png?branch=master)](https://travis-ci.org/MewesK/TwigExcelBundle)
[![Scrutinizer: Code Quality](https://scrutinizer-ci.com/g/MewesK/TwigExcelBundle/badges/quality-score.png?b=master)](https://scrutinizer-ci.com/g/MewesK/TwigExcelBundle/?branch=master)
[![SensioLabsInsight: Code Quality](https://insight.sensiolabs.com/projects/283cfe57-6ee4-4102-8fff-da3f6e668e8f/mini.png)](https://insight.sensiolabs.com/projects/283cfe57-6ee4-4102-8fff-da3f6e668e8f)

This Symfony bundle provides a PhpExcel integration for Twig.

## Supported output formats

The supported output formats are directly based on the capabilities of PhpExcel.

 * .CSV (only basic data output)
 * .ODS (only basic data output)
 * .PDF (requires mPDF)
 * .XLS (limited functionality)
 * .XLSX 

## Software requirements

The following software is required to use PHPExcel/TwigExcelBundle.

**Required by this bundle:**

 * PHP 5.5 or newer
 * Symfony 2.7 or newer

**Required by PhpExcel:**

 * PHP extension php_zip enabled
 * PHP extension php_xml enabled
 * PHP extension php_gd2 enabled (if not compiled in)

## Documentation

The source of the documentation is stored in the Resources/doc/ folder in this bundle:
    
[Resources/doc/index.rst](https://github.com/MewesK/TwigExcelBundle/blob/master/Resources/doc/index.rst)

You can find a prettier version on [readthedocs.org](httsp://readthedocs.org):

[https://twigexcelbundle.readthedocs.org](https://twigexcelbundle.readthedocs.org/)

## Installation

All the installation instructions are located in the documentation.

## License

This bundle is under the MIT license. See the complete license in the bundle:

[Resources/meta/LICENSE](https://github.com/MewesK/TwigExcelBundle/blob/master/Resources/meta/LICENSE)
