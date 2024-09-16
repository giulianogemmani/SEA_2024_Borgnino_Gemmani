###############################################################################
#
# Comments - A class for writing the Excel XLSX Worksheet file.
#
# Copyright 2013-2018, John McNamara, jmcnamara@cpan.org
#

import re

from . import xmlwriter
from .utility import xl_rowcol_to_cell


class Comments(xmlwriter.XMLwriter):
    """
    A class for writing the Excel XLSX Comments file.


    """

    ###########################################################################
    #
    # Public API.
    #
    ###########################################################################

    def __init__(self):
        """
        Constructor.

        """

        super(Comments, self).__init__()
        self.author_ids = {}

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################

    def _assemble_xml_file(self, comments_data=[]):
        # Assemble and write the XML file.

        # Write the XML declaration.
        self._xml_declaration()

        # Write the comments element.
        self._write_comments()

        # Write the authors element.
        self._write_authors(comments_data)

        # Write the commentList element.
        self._write_comment_list(comments_data)

        self._xml_end_tag('comments')

        # Close the file.
        self._xml_close()

    ###########################################################################
    #
    # XML methods.
    #
    ###########################################################################

    def _write_comments(self):
        # Write the <comments> element.
        xmlns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

        attributes = [('xmlns', xmlns)]

        self._xml_start_tag('comments', attributes)

    def _write_authors(self, comment_data):
        # Write the <authors> element.
        author_count = 0

        self._xml_start_tag('authors')

        for comment in comment_data:
            author = comment[3]

            if author is not None and author not in self.author_ids:
                # Store the author id.
                self.author_ids[author] = author_count
                author_count += 1

                # Write the author element.
                self._write_author(author)

        self._xml_end_tag('authors')

    def _write_author(self, data):
        # Write the <author> element.
        self._xml_data_element('author', data)

    def _write_comment_list(self, comment_data):
        # Write the <commentList> element.
        self._xml_start_tag('commentList')

        for comment in comment_data:
            row = comment[0]
            col = comment[1]
            text = comment[2]
            author = comment[3]

            # Look up the author id.
            author_id = None
            if author is not None:
                author_id = self.author_ids[author]

            # Write the comment element.
            self._write_comment(row, col, text, author_id)

        self._xml_end_tag('commentList')

    def _write_comment(self, row, col, text, author_id):
        # Write the <comment> element.
        ref = xl_rowcol_to_cell(row, col)

        attributes = [('ref', ref)]

        if author_id is not None:
            attributes.append(('authorId', author_id))

        self._xml_start_tag('comment', attributes)

        # Write the text element.
        self._write_text(text)

        self._xml_end_tag('comment')

    def _write_text(self, text):
        # Write the <text> element.
        self._xml_start_tag('text')

        # Write the text r element.
        self._write_text_r(text)

        self._xml_end_tag('text')

    def _write_text_r(self, text):
        # Write the <r> element.
        self._xml_start_tag('r')

        # Write the rPr element.
        self._write_r_pr()

        # Write the text r element.
        self._write_text_t(text)

        self._xml_end_tag('r')

    def _write_text_t(self, text):
        # Write the text <t> element.
        attributes = []

        if re.search('^\s', text) or re.search('\s$', text):
            attributes.append(('xml:space', 'preserve'))

        self._xml_data_element('t', text, attributes)

    def _write_r_pr(self):
        # Write the <rPr> element.
        self._xml_start_tag('rPr')

        # Write the sz element.
        self._write_sz()

        # Write the color element.
        self._write_color()

        # Write the rFont element.
        self._write_r_font()

        # Write the family element.
        self._write_family()

        self._xml_end_tag('rPr')

    def _write_sz(self):
        # Write the <sz> element.
        attributes = [('val', 8)]

        self._xml_empty_tag('sz', attributes)

    def _write_color(self):
        # Write the <color> element.
        attributes = [('indexed', 81)]

        self._xml_empty_tag('color', attributes)

    def _write_r_font(self):
        # Write the <rFont> element.
        attributes = [('val', 'Tahoma')]

        self._xml_empty_tag('rFont', attributes)

    def _write_family(self):
        # Write the <family> element.
        attributes = [('val', 2)]

        self._xml_empty_tag('family', attributes)
