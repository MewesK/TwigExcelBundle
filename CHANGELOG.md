## 2.0 (2016-02-27)

**Attention**

Some mayor changes required to increase the minimum requirements significantly. Therefore I'm bumping the version to 2.0.

**Improvements**

 * Converted help to RST-format and moved it to Resources/doc
 * Rendered version of the documentation is now hosted on readthedocs.org
 * Removed more deprecated function calls (for compatibility with Twig 2.0)
 * Removed unnecessary mock objects
 * Added messages to InvalidArgumentExceptions (darookee)
  
**Features**

 * Added support for blocks and macros in form of the new xlsblock/xlsmacro tags. The standard tags do sadly not work since they won't allow to alter the internal node structure.
 * Added support for document templates. Use existing spreadsheet files as templates.

## 1.2 (2015-12-02)

**Fixes**

 * Added Symfony 2.7 and 2.8 to the automated tests

**Improvements**

 * Added Symfony 3 support
 * Added PHP 7 automated tests

## 1.1 (2015-12-02)

**Fixes**

 * Fixed minimum PHP version to >=5.4 (short array syntax)
 * Fixed .PDF support
  
**Features**

 * Added .ODS to the supported formats
 * Added 'dataType' property to the 'xlscell' tag. Used to manually override data type of the cell.
 * Added bundle config 'mewesk_twig_excel.pre_calculate_formulas'. Pre-calculating formulas can be slow in certain cases. Disabling this option can improve the performance but the resulting documents won't show the result of any formulas when opened in a external spreadsheet software.
 * Added bundle config 'mewesk_twig_excel.disk_caching_directory'. Using disk caching can improve memory consumption by writing data to disk temporarily. Works only for .XLSX and .ODS documents.

**Improvements**

 * Updated to latest stable PHPExcel
 * Updated to latest stable mPDF
 * Improved inline documentation
 * Improved code quality
 * Added latest PHP/HHVM and Symfony versions to the automated tests
 * Added more unit tests and functional tests

## 1.0 (2014-03-04)

 * Initial release
