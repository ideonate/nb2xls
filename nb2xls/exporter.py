import copy
from io import BytesIO
import re
import base64

from nbconvert.exporters import Exporter

from bs4 import BeautifulSoup
from bs4.element import NavigableString, Tag
import xlsxwriter

import mistune
from .mdrenderer import Md2XLSRenderer, \
    MdStyleInstructionCell, MdStyleInstructionText, MdStyleInstructionLink, MdStyleInstructionList
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
        self.workbook = xlsxwriter.Workbook(output)

        self.msxlsstylereg = MdXlsStyleRegistry(self.workbook)

        self.worksheet = self.workbook.add_worksheet()

        self.row = 0
        for cellno, cell in enumerate(nb_copy.cells):
            self.worksheet.write(self.row, 0, str(cellno+1))

            if cell.cell_type == 'markdown':
                self._write_markdown(cell.source)

            elif cell.cell_type == 'code':
                for o in cell.outputs:
                    self._write_code(o)

            self.row += 1

        self.workbook.close()

        xlsx_data = output.getvalue()

        return xlsx_data, resources

    def _write_code(self, o):
        """
        Main handler for code cells
        :param o:
        :return:
        """
        if o.output_type == 'stream':
            self._write_textplain(o.text)
        if o.output_type in ('display_data', 'execute_result'):
            if 'text/html' in o.data:
                self._write_texthtml(o.data['text/html'])
            elif 'image/png' in o.data:
                width, height = 0, 0
                if 'image/png' in o.metadata and set(o.metadata['image/png'].keys()) == {'width', 'height'} :
                    width, height = o.metadata['image/png']['width'], o.metadata['image/png']['height']
                self._write_image(o.data['image/png'], width, height)
            elif 'application/json' in o.data:
                self._write_textplain(o.data['application/json'])
            elif 'text/plain' in o.data:
                self._write_textplain(o.data['text/plain'])

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
                    re.sub('\s+', ' ', s)
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
        for tablerow in soup('tr'):
            col = 1
            for child in tablerow.children:
                if isinstance(child, Tag):
                    if child.name == 'th' or child.name == 'td':
                        s = child.get_text()
                        try:
                            f = float(s)
                            self.worksheet.write_number(self.row, col, f)
                        except ValueError:
                            self.worksheet.write(self.row, col, s)

                        col += 1
            self.row += 1

    # Image handler

    def _write_image(self, image, want_width, want_height):

        image = base64.b64decode(image)

        image_data = BytesIO(image)

        if usecv2:
            nparr = np.fromstring(image, np.uint8)
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
        self.worksheet.set_row(self.row, height*y_scale)

        self.row += 1

    # Markdown handler

    def _write_markdown(self, md):
        try:
            markdown = mistune.Markdown(renderer=Md2XLSRenderer())
            lines = markdown(md)

            all_o = []

            for l in lines:
                in_softnewline = False
                is_indented = False
                link_url = ''
                already_outputted_text = False
                cell_format_mdname = ''
                o = []
                mdtextstylenames = []
                for i,s in enumerate(l):
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

                    elif isinstance(s, MdStyleInstructionList):
                        if already_outputted_text:
                            all_o.append([o, cell_format_mdname, link_url, is_indented])
                            o = []
                            already_outputted_text = False
                        in_softnewline = True
                        link_url = ''
                        is_indented = True

                    elif len(s) > 0:
                        if len(mdtextstylenames) > 0 or (cell_format_mdname != '' and i >= 2):
                            format = self.msxlsstylereg.use_style([cell_format_mdname] + mdtextstylenames)
                            o.append(format)
                            mdtextstylenames = []

                        o.append(s)
                        already_outputted_text = True

                        if in_softnewline:
                            all_o.append([o, cell_format_mdname, link_url, is_indented])
                            o = []
                            already_outputted_text = False
                            in_softnewline = False
                            is_indented = False
                            link_url = ''

                if len(o) > 0:
                    all_o.append([o, cell_format_mdname, link_url, is_indented])

            for o, cell_format_mdname, link_url, is_indented in all_o:

                is_indented = int(is_indented)

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

        except Exception as e:
            print('Markdown Exception: ', e)
            self._write_textplain(md)
