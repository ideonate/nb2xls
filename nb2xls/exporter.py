import copy
from io import BytesIO
import re
import base64
from collections.abc import Iterable
from collections import defaultdict
from math import ceil, isnan

from nbconvert.exporters import Exporter

from traitlets import Bool

from bs4 import BeautifulSoup
from bs4.element import NavigableString, Tag
import xlsxwriter

import mistune
from .mdrenderer import Md2XLSRenderer, \
    MdStyleInstructionCell, MdStyleInstructionText, MdStyleInstructionLink, MdStyleInstructionListItem, \
    MdStyleInstructionLineBreak, MdStyleInstructionListStart, MdStyleInstructionListEnd

from .mdxlsstyles import MdXlsStyleRegistry

try:
    import cv2 # Prefer Open CV2 but don't put in requirements.txt because it can be difficult to install
    import numpy as np
    usecv2 = True
except ImportError:
    import png # PyPNG is a pure-Python PNG manipulator
    usecv2 = False


class XLSExporter(Exporter):
    """
    XLSX custom exporter
    """

    # If this custom exporter should add an entry to the
    # "File -> Download as" menu in the notebook, give it a name here in the
    # `export_from_notebook` class member
    export_from_notebook = "Excel Spreadsheet"

    ignore_markdown_errors = Bool(True, help="""
        Set ignore_markdown_errors to False in order to throw an exception with any md errors. 
        From nbconvert command line for example:
        jupyter nbconvert --to xls Examples/Test.ipynb --XLSExporter.ignore_markdown_errors=False
    """).tag(config=True)

    def __init__(self, config=None, **kw):
        """
        Public constructor

        Parameters
        ----------
        config : :class:`~traitlets.config.Config`
            User configuration instance.
        `**kw`
            Additional keyword arguments passed to parent __init__

        """

        super(XLSExporter, self).__init__(config=config, **kw)

        self.msxlsstylereg = None
        self.workbook = None
        self.row = 0

    def _file_extension_default(self):
        """
        The new file extension is `.xlsx`
        """
        return '.xlsx'

    def from_notebook_node(self, nb, resources=None, **kw):
        """
        Convert a notebook from a notebook node instance.
        Parameters
        ----------
        nb : :class:`~nbformat.NotebookNode`
          Notebook node (dict-like with attr-access)
        resources : dict
          Additional resources that can be accessed read/write by
          preprocessors and filters.
        `**kw`
          Ignored
        """
        nb_copy = copy.deepcopy(nb)
        resources = self._init_resources(resources)

        if 'language' in nb['metadata']:
            resources['language'] = nb['metadata']['language'].lower()

        # Preprocess
        nb_copy, resources = super(XLSExporter, self).from_notebook_node(nb, resources, **kw)

        output = BytesIO()
        self.workbook = xlsxwriter.Workbook(output, {'nan_inf_to_errors': True})

        self.msxlsstylereg = MdXlsStyleRegistry(self.workbook)

        self.worksheet = self.workbook.add_worksheet()

        self.row = 0
        for cellno, cell in enumerate(nb_copy.cells):
            self.worksheet.write(self.row, 0, str(cellno+1))

            # Convert depending on nbformat
            # https://nbformat.readthedocs.io/en/latest/format_description.html#cell-types

            if cell.cell_type == 'markdown':
                self._write_markdown(cell.source)

            elif cell.cell_type == 'code':
                self._write_code(cell)

            else:
                self._write_textplain('No convertible outputs available for cell: {}'.format(cell.source))

            self.row += 1

        self.workbook.close()

        xlsx_data = output.getvalue()

        return xlsx_data, resources

    def _write_code(self, cell):
        """
        Main handler for code cells
        :param celloutputs:
        :return:
        """

        for i,o in enumerate(cell.outputs):

            display_data = None
            if o.output_type in ('execute_result', 'display_data'):
                if 'text/html' in o.data:
                    self._write_texthtml(o.data['text/html'])
                elif 'text/markdown' in o.data:
                    self._write_markdown(o.data['text/markdown'])
                elif 'image/png' in o.data:
                    width, height = 0, 0
                    if 'image/png' in o.metadata and set(o.metadata['image/png'].keys()) == {'width', 'height'} :
                        width, height = o.metadata['image/png']['width'], o.metadata['image/png']['height']
                    self._write_image(o.data['image/png'], width, height)
                elif 'application/json' in o.data:
                    self._write_textplain(repr(o.data['application/json']))
                elif 'text/plain' in o.data:
                    self._write_textplain(o.data['text/plain'])
                else:
                    self._write_textplain('No convertible mimetype available for source (output {}): {}'.format(i, cell.source))

            elif o.output_type == 'stream':
                self._write_textplain(o.text)

            if i < len(cell.outputs)-1:
                # Blank row between outputs, but not at the end
                self.row += 1

    ###
    # Sub-handlers for code cells

    def _write_textplain(self, text):
        lines = text.split("\n")
        for l in lines:
            self.worksheet.write(self.row, 1, l)
            self.row += 1

    # HTML functions start here

    def _write_texthtml(self, html):
        soup = BeautifulSoup(html, 'html.parser')
        self._write_soup(soup)

    def _write_soup(self, soup):
        s = ''
        for child in soup.children:

            if isinstance(child, NavigableString):
                s += child.string

            elif isinstance(child, Tag):

                if len(s) > 0:
                    # Write accumulated string first
                    re.sub(r'\s+', ' ', s)
                    s = s.strip()
                    if len(s) > 0:
                        self.worksheet.write(self.row, 1, s.strip())
                        self.row += 1
                        s = ''

                if child.name in ('div', 'body', 'span', 'p'):
                    self._write_soup(child)

                elif child.name == 'table':
                    self._write_htmltable(soup)

    def _write_htmltable(self, soup):
        double_emphasis_fmt = self.msxlsstylereg.use_style(['double_emphasis'])
        rowspans = defaultdict(int)
        for tablerow in soup('tr'):
            col = 1
            for child in tablerow.children:
                if isinstance(child, Tag):
                    if child.name == 'th' or child.name == 'td':
                        while rowspans[col] > 1:
                            rowspans[col] -= 1
                            col += 1

                        s = child.get_text()

                        fmt = double_emphasis_fmt if child.name == 'th' else None

                        try:
                            f = float(s)
                            if isnan(f):
                                self.worksheet.write_formula(self.row, col, '=NA()', fmt)
                            else:
                                self.worksheet.write_number(self.row, col, f, fmt)
                        except ValueError:
                            self.worksheet.write(self.row, col, s, fmt)

                        if 'rowspan' in child.attrs and child.attrs['rowspan'].isdigit():
                            rowspans[col] = int(child.attrs['rowspan'])

                        if 'colspan' in child.attrs and child.attrs['colspan'].isdigit():
                            col += int(child.attrs['colspan'])-1

                        col += 1
            self.row += 1

    # Image handler

    def _write_image(self, image, want_width, want_height):

        image = base64.b64decode(image)

        image_data = BytesIO(image)

        if usecv2:
            nparr = np.frombuffer(image, np.uint8)
            img = cv2.imdecode(nparr, cv2.IMREAD_ANYCOLOR)
            height, width = img.shape[:2]
        else:
            img = png.Reader(image_data)
            (width, height, _, _) = img.asDirect()

        x_scale, y_scale = 1.0, 1.0

        if want_height > 0 and height > 0:
            y_scale = want_height / height

        if want_width > 0 and width > 0:
            x_scale = want_width / width

        self.row += 1

        self.worksheet.insert_image(self.row, 1, 'image.png',
                                    {'image_data': image_data, 'x_scale': x_scale, 'y_scale': y_scale})

        self.row += ceil(height*y_scale / 15) # 15 is default row height in Excel

    # Markdown handler

    def _write_markdown(self, md):
        if self.ignore_markdown_errors:
            try:
                self._write_markdown_core(md)
            except Exception as e:
                print('Markdown Exception: ', e)
                self._write_textplain(md)
        else:
            self._write_markdown_core(md)

    def _write_markdown_core(self, md):
        markdown = mistune.Markdown(renderer=Md2XLSRenderer())
        lines = markdown(md)

        def flatten(l):
            """
            Nested lists cause problems due to the way the MD parser works.
            :param l: arbitrary-depth nested list of strs and MdStyleInstruction objects
            :return: single-depth flattened version of the array containing only the leaves
            """
            for el in l:
                if isinstance(el, Iterable) and not isinstance(el, str):
                    for sub in flatten(el):
                        yield sub
                else:
                    yield el

        list_counters = []
        list_ordered = []

        all_o = []

        for l in lines:
            in_softnewline = False
            is_indented = 0
            link_url = ''
            already_outputted_text = False
            cell_format_mdname = ''
            o = []
            mdtextstylenames = []
            for i,s in enumerate(flatten(l)):
                if isinstance(s, MdStyleInstructionText):
                    mdtextstylenames += [s.mdname]

                elif isinstance(s, MdStyleInstructionCell):
                    cell_format_mdname = s.mdname

                elif isinstance(s, MdStyleInstructionLink):
                    if already_outputted_text:
                        all_o.append([o, cell_format_mdname, link_url, is_indented])
                        o = []
                        already_outputted_text = False
                    in_softnewline = True
                    link_url = s.link

                elif isinstance(s, MdStyleInstructionListStart):
                    if already_outputted_text:
                        all_o.append([o, cell_format_mdname, link_url, is_indented])
                        o = []
                        already_outputted_text = False

                    is_indented += 1
                    list_counters.append(1)
                    list_ordered.append(s.ordered)

                elif isinstance(s, MdStyleInstructionListEnd):

                    if already_outputted_text:
                        all_o.append([o, cell_format_mdname, link_url, is_indented])
                        o = []
                        already_outputted_text = False

                    is_indented -= 1

                    list_counters.pop()
                    list_ordered.pop()

                    in_softnewline = True
                    link_url = ''

                elif isinstance(s, MdStyleInstructionListItem):
                    if already_outputted_text:
                        all_o.append([o, cell_format_mdname, link_url, is_indented])
                        o = []
                        already_outputted_text = False
                    in_softnewline = True
                    link_url = ''

                    li_count = list_counters[-1]
                    if list_ordered[-1]:
                        o = ['{}. '.format(li_count)]
                    list_counters[-1] += 1

                elif isinstance(s, MdStyleInstructionLineBreak):
                    if already_outputted_text:
                        all_o.append([o, cell_format_mdname, link_url, is_indented])
                        o = []
                        already_outputted_text = False
                    in_softnewline = True
                    link_url = ''

                elif len(s) > 0:
                    if len(mdtextstylenames) > 0 or (cell_format_mdname != '' and i >= 2):
                        fmt = self.msxlsstylereg.use_style([cell_format_mdname] + mdtextstylenames)
                        o.append(fmt)
                        mdtextstylenames = []

                    o.append(s)
                    already_outputted_text = True

                    if in_softnewline and link_url != '':
                        all_o.append([o, cell_format_mdname, link_url, is_indented])
                        o = []
                        already_outputted_text = False
                        in_softnewline = False
                        link_url = ''

            if len(o) > 0:
                all_o.append([o, cell_format_mdname, link_url, is_indented])

        for o, cell_format_mdname, link_url, is_indented in all_o:

            if cell_format_mdname != '':
                o.append(self.msxlsstylereg.use_style(cell_format_mdname))

            if link_url != '':
                if len(o) >= 2:
                    self.worksheet.write_url(self.row, 1+is_indented, link_url, o[1], o[0])
                elif len(o) == 1:
                    self.worksheet.write_url(self.row, 1+is_indented, link_url, None, o[0])

            else:

                if len(o) > 2:
                    self.worksheet.write_rich_string(self.row, 1+is_indented, *o)
                elif len(o) == 2:
                    if isinstance(o[0], xlsxwriter.format.Format) and not isinstance(o[1], xlsxwriter.format.Format):
                        self.worksheet.write(self.row, 1+is_indented, o[1], o[0])
                    elif not isinstance(o[0], xlsxwriter.format.Format) and not isinstance(o[1], xlsxwriter.format.Format):
                        self.worksheet.write(self.row, 1+is_indented, o[0] + ' ' + o[1])
                    else:
                        self.worksheet.write(self.row, 1+is_indented, o[0], o[1])
                elif len(o) == 1 and not isinstance(o[0], xlsxwriter.format.Format):
                    self.worksheet.write(self.row, 1+is_indented, o[0])

            self.row += 1

