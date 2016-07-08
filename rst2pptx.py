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
        self.slides = self.presentation.slides

    def visit_document(self, node):
        pass

    def depart_document(self, node):
        pass

    def visit_docinfo_item(self, node, name):
        pass

    def visit_image(self, node):
        pass

    def visit_Text(self, node):
        pass

    def depart_Text(self, node):
        pass

    def visit_list_item(self, node):
        print('visit_list_item({})'.format(node))

    def depart_list_item(self, node):
        print('depart_list_item({})'.format(node))

    def visit_paragraph(self, node):
        text_frame = self.slides[-1].shapes.placeholders[1].text_frame
        paragraph = text_frame.add_paragraph()
        paragraph.text = node.astext()
        paragraph.level = self.bullet_level

    def depart_paragraph(self, node):
        print('depart_paragraph({})'.format(node))

    def visit_section(self, node):
        print('visit_section({})'.format(node))
        # TODO: Handle title page.
        self.slides.add_slide(self.presentation.slide_layouts[1])

    def depart_section(self, node):
        pass

    def visit_title(self, node):
        print('visit_title({})'.format(node))
        self.slides[-1].shapes.title.text = node.astext()
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
