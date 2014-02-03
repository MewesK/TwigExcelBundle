PhpExcel Twig Extension Bundle (alpha)
========================

[![Build Status](https://travis-ci.org/MewesK/PhpExcelTwigExtensionBundle.png?branch=dev)](https://travis-ci.org/MewesK/PhpExcelTwigExtensionBundle)

Helper Functions
----------------------------------

  * xlsmergestyles([literal], [literal])

Full Syntax
----------------------------------

```lua
{#
    xlsdocument tag

    {% xlsdocument [properties:literal] %}[Twig_NodeInterface]{% endxlsdocument %}
    {% xlsdocument {
        category: [string],
        company: [string],
        created: [datetime],
        creator: [string],
        defaultStyle: [literal],
        description: [string],
        format: [string],
        keywords: [string],
        lastModifiedBy: [string],
        manager: [string],
        modified: [datetime],
        security: [literal],
        subject: [string],
        title: [string]
    } %}[Twig_NodeInterface]{% endxlsdocument %}

    'properties' are optional
    cannot contain other 'xlsdocument' tags
    may contain multiple 'xlssheet' tags
    possible formats are 'csv', 'html', 'pdf', 'xls, 'xlsx'

    see: PHPExcel_Document
    see: PHPExcel_DocumentSecurity
    see: PHPExcel_Style:applyFromArray
#}
{% xlsdocument {
    category: 'Test category',
    company: 'Test company',
    created: 'now'|date('DATE_ATOM'),
    creator: 'Tester',
    defaultStyle: {
        font: {
            name: 'Verdana',
            size: '18'
        }
    },
    description: 'Test document',
    format: 'xls',
    keywords: 'Test',
    lastModifiedBy: 'Tester',
    manager: 'Tester',
    modified: 'now'|date('DATE_ATOM'),
    security: {
        lockRevision: true,
        lockStructure: true,
        lockWindows: true,
        revisionsPassword: 'test'
        workbookPassword: 'test'
    },
    subject: 'Test',
    title: 'Test'
} %}
    {#
        xlssheet tag

        {% xlssheet [title:string] [properties:literal] %}[Twig_NodeInterface]{% endxlssheet %}
        {% xlssheet [title:string] {
            columnDimension: [literal],
            pageMargins: [literal],
            pageSetup: [literal],
            protection: [literal],
            printGridlines: [boolean],
            rightToLeft: [boolean],
            rowDimension: [literal],
            sheetState: [string],
            showGridlines: [boolean],
            tabColor: [string],
            zoomScale: [integer]
        } %}[Twig_NodeInterface]{% endxlssheet %}

        'title' is required, 'properties' are optional
        cannot contain other 'xlssheet' tags
        may contain multiple 'xlscell' tags

        see: PHPExcel_Worksheet_HeaderFooter
        see: PHPExcel_Worksheet_PageMargins
        see: PHPExcel_Worksheet_PageSetup
        see: PHPExcel_Worksheet_RowDimension
        see: PHPExcel_Worksheet_ColumnDimension
    #}
    {% xlssheet 'Worksheet' {
        columnDimension: {
            'default': {
                autoSize: false,
                collapsed: false,
                columnIndex: 'A',
                outlineLevel: 0,
                visible: true,
                width: -1,
                xfIndex: 0
            },
            'D': {
                visible: false
            }
        },
        footer: 'Header',
        header: 'Footer',
        pageMargins: {
            top: 1,
            bottom: 1,
            left: 0.75,
            right: 0.75,
            header: 0.5,
            footer: 0.5
        },
        pageSetup: {
            fitToHeight: 1,
            fitToPage: false,
            fitToWidth: 1,
            horizontalCentered: false,
            orientation: 'landscape',
            paperSize: 9,
            printArea: 'A1:B1',
            scale: 100,
            verticalCentered: false
        },
        protection: {
            autoFilter: true,
            deleteColumns: true,
            deleteRows: true,
            formatCells: true,
            formatColumns: true,
            formatRows: true,
            insertColumns: true,
            insertHyperlinks: true,
            insertRows: true,
            objects: true,
            pivotTables: true,
            scenarios: true,
            selectLockedCells: true,
            selectUnlockedCells: true,
            sheet: true,
            sort: true
        },
        printGridlines: true,
        rightToLeft: false,
        rowDimension: {
            'default': {
                collapsed: false,
                outlineLevel: 0,
                rowHeight: -1,
                rowIndex: '1',
                visible: true,
                xfIndex: 0
                zeroHeight:false,
            },
            '2': {
                visible: false
            }
        },
        sheetState: 'visible',
        showGridlines: true,
        tabColor: 'c0c0c0',
        zoomScale: 75
    }%}
        {#
            xlsheader tag

            {% xlsheader [type:string] [properties:literal] %}[Twig_NodeInterface]{% endxlsheader %}

            'type' and 'properties' are optional
            'type' can be null, 'odd', 'even' or 'first'
            'type' null makes all headers the same
            cannot contain other 'xlsheader' tags
        #}
        {% xlsheader 'first' %}Test Header{% endxlsrow %}
        {#
            xlsrow tag

            {% xlsrow [index:integer] %}[Twig_NodeInterface]{% endxlsrow %}

            'index' is optional
            if 'index' is not defined it will default to 1 for the first usage per sheet
            for each further usage it will increase the index by 1 automatically (1, 2, 3, ...)
        #}
        {% xlsrow 1 %}
            {#
                xlscell tag

                {% xlscell [index:string] [properties:literal] %}[string]{% endxlscell %}
                {% xlscell [index:string] {
                    break: [integer],
                    dataValidation: [literal],
                    style: [literal],
                    url: [string]
                } %}[string]{% endxlscell %}

                'index' and 'properties' are optional
                if 'index' is not defined it will default to 0 for the first usage per row
                for each further usage it will increase the index by 1 automatically (0, 1, 2, ...)
                cannot contain other 'xlscell' tags

                see: PHPExcel_Cell_DataValidation
                see: PHPExcel_Style:applyFromArray
                see: PHPExcel_Style_Border
            #}
            {% xlscell 0 {
                break: 1,
                dataValidation: {
                    allowBlank: false,
                    error: '',
                    errorStyle: 'stop',
                    errorTitle: '',
                    formula1: '',
                    formula2: '',
                    operator: '',
                    prompt: ''
                    promptTitle: '',
                    showDropDown: false,
                    showErrorMessage: false,
                    showInputMessage: false,
                    type: 'none',
                },
                style: {
                    borders: {
                        bottom: {
                            style: 'thin',
                            color: {
                                rgb: '000000'
                            }
                        }
                    }
                },
                url: 'http://www.example.com'
            } %}
                Test
            {% endxlscell %}
            {% xlscell 2 %}Foo{% endxlscell %}
            {% xlscell %}Bar{% endxlscell %}
        {% endxlsrow %}
        {#
            xlsdrawing tag

            {% xlsdrawing [path:string] [properties:literal] %}
            {% xlsdrawing [path:string] {
                style: [literal]
            } %}

            'path' is required, 'properties' are optional

            see: PHPExcel_Worksheet_Drawing
            see: PHPExcel_Worksheet_Drawing_Shadow
        #}
        {% xlsdrawing '/test.png' {
            coordinates: 'A1',
            description: 'Test',
            height: 0,
            name: '',
            offsetX: 0,
            offsetY: 0,
            resizeProportional: true,
            rotation: 0,
            shadow: {
                alignment: 'br',
                alpha: 50,
                blurRadius: 6,
                color: '000000',
                direction: 0,
                distance: 2,
                visible: false
            },
            width: 0
        } %}
        {#
            xlsfooter tag

            {% xlsfooter [type:string] [properties:literal] %}[Twig_NodeInterface]{% endxlsfooter %}

            'type' and 'properties' are optional
            'type' can be null, 'odd', 'even' or 'first'
            'type' null makes all footers the same
        #}
        {% xlsfooter 'first' %}Test Header{% xlsfooter %}
    {% endxlssheet %}
{% endxlsdocument %}
```