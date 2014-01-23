PhpExcel Twig Extension Bundle
========================

[![Build Status](https://travis-ci.org/MewesK/PhpExcelTwigExtensionBundle.png?branch=master)](https://travis-ci.org/MewesK/PhpExcelTwigExtensionBundle)

Helper Functions
----------------------------------

  * xlsmergestyles([literal], [literal])

Full Syntax
----------------------------------

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
                footer: [string],
                header: [string],
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
                xlscell tag

                {% xlssheet [coordinates:string] [properties:literal] %}[string]{% endxlscell %}
                {% xlssheet [coordinates:string] {
                    break: [integer],
                    dataValidation: [literal],
                    style: [literal],
                    url: [string]
                } %}[string]{% endxlscell %}

                'coordinates' are required, 'properties' are optional
                cannot contain other 'xlscell' tags

                see: PHPExcel_Cell_DataValidation
                see: PHPExcel_Style:applyFromArray
                see: PHPExcel_Style_Border
            #}
            {% xlscell 'A1' {
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
            {% xlscell 'B1' %}Foo{% endxlscell %}
            {% xlscell 'C1' %}Bar{% endxlscell %}
            {#
                xlsstyle tag

                {% xlsstyle [coordinates:string] [properties:literal] %}
                {% xlsstyle [coordinates:string] {
                    style: [literal]
                } %}

                'coordinates' and 'properties' are required

                see: PHPExcel_Style:applyFromArray
                see: PHPExcel_Style_Border
            #}
            {% xlsstyle 'B1:C1' {
                style: {
                    font: {
                        color: 'ff0000'
                    }
                }
            } %}
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
        {% endxlssheet %}
    {% endxlsdocument %}