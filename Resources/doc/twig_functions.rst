Twig Functions
==============

xlsmergestyles
--------------

.. code-block:: twig

    xlsmergestyles([style1:array], [style2:array])

- Merges two style arrays recursively
- Returns a new array

Parameters
``````````

==========  ======  ========  ===========
Name        Type    Optional  Description
==========  ======  ========  ===========
style1      array             Standard PhpExcel style array
style2      array             Standard PhpExcel style array
==========  ======  ========  ===========

Example
```````

.. code-block:: twig

    {% set mergedStyle = xlsmergestyles({ font: { name: 'Verdana' } }, { font: { size: '18' } }) %}

