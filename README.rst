========
rst2pptx
========

Converts reStructuredText to PowerPoint.


Installation
============

From pip::

    $ pip install --upgrade rst2pptx

Example
=======

::

    $ rst2pptx input.rst output.pptx

Input:

.. code-block:: rst

    Slide 1
    =======

    - it has some bullets
    - bullet 2

Output:


+----------------------------------------------------------------------------------------+
| .. image:: https://raw.githubusercontent.com/myint/rst2pptx/master/examples/output.png |
+----------------------------------------------------------------------------------------+

Warning
=======

This tool is in a very early experimental phase. Also, if you don't
specifically need to output PowerPoint, rst2beamer_, which generates PDFs is a
better choice.

.. _rst2beamer: https://github.com/myint/rst2beamer
