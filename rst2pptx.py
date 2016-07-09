#!/usr/bin/env python
# encoding: utf-8

# Copyright (C) 2016 Steven Myint
#
# Permission is hereby granted, free of charge, to any person obtaining
# a copy of this software and associated documentation files (the
# "Software"), to deal in the Software without restriction, including
# without limitation the rights to use, copy, modify, merge, publish,
# distribute, sublicense, and/or sell copies of the Software, and to
# permit persons to whom the Software is furnished to do so, subject to
# the following conditions:
#
# The above copyright notice and this permission notice shall be
# included in all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
# MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
# NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS
# BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
# ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
# CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

"""Converts reStructuredText to PowerPoint."""

from __future__ import absolute_import
from __future__ import division
from __future__ import unicode_literals

import os

import docutils.core
import docutils.nodes
import docutils.utils
import pptx


class PowerPointTranslator(docutils.nodes.NodeVisitor):

    """A translator for converting docutils elements to PowerPoint."""

    def __init__(self, document):
        docutils.nodes.NodeVisitor.__init__(self, document)

        self.bullet_level = 0
        self.presentation = pptx.Presentation()
        self.root_path = None
        self.slides = self.presentation.slides
        self.table_rows = None

    def visit_document(self, node):
        self.root_path = os.path.dirname(node['source'])

    def depart_document(self, node):
        pass

    def visit_docinfo_item(self, node, name):
        pass

    def visit_image(self, node):
        print(type(node))
        self.slides[-1].shapes.add_picture(
            os.path.join(self.root_path, node.attributes['uri']),
            left=pptx.util.Inches(1.),
            top=pptx.util.Inches(2.))

    def visit_Text(self, node):
        pass

    def depart_Text(self, node):
        pass

    def visit_list_item(self, node):
        pass

    def depart_list_item(self, node):
        pass

    def visit_paragraph(self, node):
        if self.table_rows is None:
            text_frame = self.slides[-1].shapes.placeholders[1].text_frame
            paragraph = text_frame.add_paragraph()
            paragraph.text = node.astext()
            paragraph.level = self.bullet_level
        else:
            print('self.table_rows:', self.table_rows)
            self.table_rows[-1].append(node.astext())

    def depart_paragraph(self, node):
        print('depart_paragraph({})'.format(node))

    def visit_section(self, node):
        self.slides.add_slide(self.presentation.slide_layouts[1])

    def visit_title(self, node):
        print('visit_title({})'.format(node))
        if len(self.slides):
            self.slides[-1].shapes.title.text = node.astext()
        else:
            # Title slide.
            slide = self.slides.add_slide(self.presentation.slide_layouts[0])
            slide.shapes.title.text = node.astext()
            # TODO: Author.

    def depart_title(self, node):
        pass

    def visit_literal_block(self, node):
        pass

    def depart_literal_block(self, node):
        pass

    def visit_bullet_list(self, node):
        self.bullet_level += 1

    def depart_bullet_list(self, node):
        self.bullet_level -= 1
        assert self.bullet_level >= 0

    def visit_enumerated_list(self, node):
        pass

    def depart_enumerated_list(self, node):
        pass

    def visit_tgroup(self, node):
        self.table_rows = []

    def depart_tgroup(self, node):
        if self.table_rows and self.table_rows[0]:
            table = self.slides[-1].shapes.add_table(
                rows=len(self.table_rows),
                cols=len(self.table_rows[0]),
                left=pptx.util.Inches(1.),
                top=pptx.util.Inches(2.),
                width=pptx.util.Inches(8.),
                height=pptx.util.Inches(4.)).table

            for (row_index, row) in enumerate(self.table_rows):
                for (col_index, col) in enumerate(row):
                    table.cell(row_idx=row_index, col_idx=col_index).text = col

            self.table_rows = None

    def visit_row(self, node):
        assert self.table_rows is not None
        self.table_rows.append([])

    def unknown_visit(self, node):
        print('unknown_visit({})'.format(node))

    def unknown_departure(self, node):
        print('unknown_visit({})'.format(node))

    def astext(self):
        # TODO
        pass


class PowerPointWriter(docutils.core.writers.Writer):

    """A docutils writer that produces PowerPoint."""

    def __init__(self):
        docutils.core.writers.Writer.__init__(self)
        self.translator_class = PowerPointTranslator

    def translate(self):
        visitor = self.translator_class(self.document)
        self.document.walkabout(visitor)
        self.output = visitor.astext()

        # TODO: Take name from command line.
        visitor.presentation.save('test.pptx')


def main():
    description = (
        'Generates PowerPoint presentations. ' +
        docutils.core.default_description)

    docutils.core.publish_cmdline(
        writer=PowerPointWriter(),
        description=description,
        settings_overrides={'halt_level': docutils.utils.Reporter.ERROR_LEVEL})


if __name__ == '__main__':
    main()
