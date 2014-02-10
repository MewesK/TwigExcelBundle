# TwigExcelBundle (alpha)

[![Build Status](https://travis-ci.org/MewesK/PhpExcelTwigExtensionBundle.png?branch=master)](https://travis-ci.org/MewesK/PhpExcelTwigExtensionBundle)

This Symfony2 bundle provides a PhpExcel integration for Twig.

## Installation

### Step 1: Download using composer

```js
{
    "require": {
        "mewesk/twig-excel-bundle": "1.0.*@dev"
    }
}
```

```bash
$ php composer.phar update mewesk/twig-excel-bundle
```

### Step 2: Enable the bundle

```php
// app/AppKernel.php

public function registerBundles()
{
    $bundles = array(
        // ...
        new MewesK\TwigExcelBundle\TwigExcelBundle(),
    );
}
```

## Getting started

### Step 1: Create your controller

```php
// src/Acme/HelloBundle/Controller/HelloController.php

namespace Acme\HelloBundle\Controller;

use Sensio\Bundle\FrameworkExtraBundle\Configuration\Route;
use Sensio\Bundle\FrameworkExtraBundle\Configuration\Template;
use Symfony\Component\HttpFoundation\Response;

class HelloController
{
    /**
     * @Route("/hello.{_format}", defaults={"_format"="xls"}, requirements={"_format"="csv|xls|xlsx"})
     * @Template("AcmeHelloBundle:Hello:index.xls.twig")
     */
    public function indexAction($name)
    {
        return array('data' => array('La', 'Le', 'Lu'));
    }
}
```

### Step 2: Create your template

```lua
{# src/Acme/HelloBundle/Resources/views/Hello/index.xls.twig #}

{% xlsdocument %}
    {% xlssheet 'Worksheet' %}
        {% xlsrow %}
            {% xlscell { style: { font: { size: '18' } } } %}Values{% endxlscell %}
        {% endxlsrow %}
        {% for value in data %}
            {% xlsrow %}
                {% xlscell %}{{ value }}{% endxlscell %}
            {% endxlsrow %}
        {% endfor %}
    {% endxlssheet %}
{% endxlsdocument %}
```

## Twig Functions

### xlsmergestyles

    xlsmergestyles([style1:array], [style2:array])
    
 * Merges two style arrays recursively
 * Returns a new array

#### Parameters

Name | Type | Optional | Description
---- | ---- | -------- | -----------
style1 | array | | Standard PhpExcel style array
style2 | array | | Standard PhpExcel style array

#### Example

```lua
{% set mergedStyle = xlsmergestyles({ font: { name: 'Verdana' } }, { font: { size: '18' } }) %}
```

## Twig Tags

### xlsdocument

    {% xlsdocument [properties:array] %}
        ...
    {% endxlsdocument %}

 * Cannot contain 'xlsdocument', 'xlsrow', 'xlsheader', 'xlsfooter', 'xlscell', 'xlsleft', 'xlscenter', 'xlsright' or 'xlsdrawing' tags
 * May contain one or more 'xlssheet' tags

#### Attributes

Name | Type | Optional | Description
---- | ---- | -------- | -----------
properties | array | X

#### Properties

Name | Type | Description | XLS
---- | ---- | ----------- | ---
category | string
company | string
created | datetime
creator | string
defaultStyle | array | Standard PhpExcel style array
description | string
format | string | Possible formats are 'csv', 'html', 'pdf', 'xls, 'xlsx'
keywords | string
lastModifiedBy | string
manager | string
modified | datetime
security | array
+ lockRevision | boolean
+ lockStructure | boolean
+ lockWindows | boolean
+ revisionsPassword | string
+ workbookPassword | string
subject | string
title | string

#### Example

```lua
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
    {# ... #}
{% endxlsdocument %}
```

### xlssheet

        {% xlssheet [title:string] [properties:array] %}
        ...
        {% endxlssheet %}

 * Cannot contain 'xlsdocument', 'xlssheet', 'xlscell', 'xlsleft', 'xlscenter' or 'xlsright' tags
 * May contain one or more 'xlsheader', 'xlsfooter', 'xlsrow' and 'xlsdrawing' tags

#### Attributes

Name | Type | Optional | Description
---- | ---- | -------- | -----------
title | string
properties | array | X

#### Properties

Name | Type | Description | XLS
---- | ---- | ----------- | ---
columnDimension | array | Contains one or more arrays. Possible keys are 'default' or a valid column name like 'A'
 + autoSize | boolean
 + collapsed | boolean
 + columnIndex | int | A column index >=0
 + outlineLevel | int
 + visible | boolean
 + width | double
 + xfIndex | int
pageMargins | array
 + top | double
 + bottom | double
 + left | double
 + right | double
 + header | double
 + footer | double
pageSetup | array
 + fitToHeight | int
 + fitToPage | boolean
 + fitToWidth | int
 + horizontalCentered | boolean
 + orientation | string | Possible orientations are 'default', 'landscape', 'portrait'
 + paperSize | int | Possible values are defined in PHPExcel_Worksheet_PageSetup
 + printArea | string | An area like 'A1:B1'
 + scale | int
 + verticalCentered | boolean
protection | array
 + autoFilter | boolean
 + deleteColumns | boolean
 + deleteRows | boolean
 + formatCells | boolean
 + formatColumns | boolean
 + formatRows | boolean
 + insertColumns | boolean
 + insertHyperlinks | boolean
 + insertRows | boolean
 + objects | boolean
 + pivotTables | boolean
 + scenarios | boolean
 + selectLockedCells | boolean
 + selectUnlockedCells | boolean
 + sheet | boolean
 + sort | boolean
printGridlines | boolean
rightToLeft | boolean
rowDimension | array | Contains one or more arrays. Possible keys are 'default' or a row index >=1
 + collapsed | boolean
 + outlineLevel | int
 + rowHeight | double
 + rowIndex | int | A row index >=1
 + visible | boolean
 + xfIndex | int
 + zeroHeight | boolean
sheetState | string
showGridlines | boolean
tabColor | string
zoomScale | int

#### Example

```lua
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
            xfIndex: 0,
            zeroHeight:false
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
    {# ... #}
{% endxlssheet %}
```

### xlsheader

    {% xlsheader [type:string] [properties:array] %}
        ...
    {% endxlsheader %}

 * Cannot contain 'xlsdocument', 'xlssheet', 'xlsrow', 'xlsheader', 'xlsfooter', 'xlscell' or 'xlsdrawing' tags
 * May contain one 'xlsleft', 'xlscenter' and 'xlsright'

#### Attributes

Name | Type | Optional | Description
---- | ---- | -------- | -----------
type | string | X | Possible types are 'header' (default), 'oddHeader', 'evenHeader', 'firstHeader'
properties | array | X

#### Properties

Name | Type | Description | XLS
---- | ---- | ----------- | ---
scaleWithDocument | boolean
alignWithMargins | boolean

#### Example

```lua
{% xlsheader 'first' %}
    {# ... #}
{% endxlsrow %}
```

### xlsfooter

    {% xlsfooter [type:string] [properties:array] %}
        ...
    {% endxlsfooter %}

 * Cannot contain 'xlsdocument', 'xlssheet', 'xlsrow', 'xlsheader', 'xlsfooter', 'xlscell' or 'xlsdrawing' tags
 * May contain one 'xlsleft', 'xlscenter' and 'xlsright'

#### Attributes

Name | Type | Optional | Description
---- | ---- | -------- | -----------
type | string | X | Possible types are 'footer' (default), 'oddFooter', 'evenFooter', 'firstFooter'
properties | array | X

#### Properties

Name | Type | Description | XLS
---- | ---- | ----------- | ---
scaleWithDocument | boolean
alignWithMargins | boolean

#### Example

```lua
{% xlsfooter 'first' %}
    {# ... #}
{% xlsfooter %}
```

### TODO xlsleft, xlscenter, xlsright

### xlsrow

    {% xlsrow [index:int] %}
        ...
    {% endxlsrow %}

 * Cannot contain 'xlsdocument', 'xlssheet', 'xlsrow', 'xlsheader', 'xlsfooter' or 'xlsdrawing' tags
 * May contain one or more 'xlscell' tags
 * If 'index' is not defined it will default to 1 for the first usage per sheet
 * For each further usage it will increase the index by 1 automatically (1, 2, 3, ...)

#### Attributes

Name | Type | Optional | Description
---- | ---- | -------- | -----------
index | int | | A row index >=1

#### Example

```lua
{% xlsrow 1 %}
    {# ... #}
{% endxlsrow %}
```

### xlscell

    {% xlscell [index:string] [properties:array] %}
        ...
    {% endxlscell %}

 * Cannot contain 'xlsdocument', 'xlssheet', 'xlsrow', 'xlsheader', 'xlsfooter', 'xlscell' or 'xlsdrawing' tags
 * If 'index' is not defined it will default to 0 for the first usage per row
 * For each further usage it will increase the index by 1 automatically (0, 1, 2, ...)

#### Attributes

Name | Type | Optional | Description
---- | ---- | -------- | -----------
index | int | | A column index >=0
properties | array | X

#### Properties

Name | Type | Description | XLS
---- | ---- | ----------- | ---
break | int | Possible values are defined in PHPExcel_Worksheet
dataValidation | array
 + allowBlank | boolean
 + error | string
 + errorStyle | string | Possible values are defined in PHPExcel_Cell_DataValidation
 + errorTitle | string
 + formula1 | string
 + formula2 | string
 + operator | string | Possible values are defined in PHPExcel_Cell_DataValidation
 + prompt | string
 + promptTitle | string
 + showDropDown | boolean
 + showErrorMessage | boolean
 + showInputMessage | boolean
 + type | string | Possible values are defined in PHPExcel_Cell_DataValidation
style | array | Standard PhpExcel style array
url | string
    
#### Example

```lua
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
    {# ... #}
{% endxlscell %}
```

### xlsdrawing

    {% xlsdrawing [path:string] [properties:array] %}

 * Cannot contain anything

#### Attributes

Name | Type | Optional | Description
---- | ---- | -------- | -----------
path | string
properties | array | X

#### Properties

Name | Type | Description | XLS
---- | ---- | ----------- | ---
coordinates | string | Cell coordinates like 'A1' | X
description | string
height | int | | X
name | string
offsetX | int
offsetY | int
resizeProportional | boolean | | X
rotation | int
shadow | array
 + alignment: | string | Possible values are defined in PHPExcel_Worksheet_Drawing_Shadow
 + alpha | int
 + blurRadius | int
 + color | string | A hexadecimal color string like '000000' (without #)
 + direction | int
 + distance | int
 + visible | boolean
width | int | | X

#### Example

```lua
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
```
