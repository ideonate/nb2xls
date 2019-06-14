import copy
from io import BytesIO
import re
import base64

from nbconvert.exporters import Exporter

from bs4 import BeautifulSoup
from bs4.element import NavigableString, Tag
import xlsxwriter

try:
    import cv2 # Prefer Open CV2 but don't put in requirements.txt because it can be difficult to install
    import numpy as np
    usecv2 = True
except ImportError:
    import png # PyPNG is a pure-Python PNG manipulator
    usecv2 = False


class XLSExporter(Exporter):
    """
    My custom exporter
    """

    # If this custom exporter should add an entry to the
    # "File -> Download as" menu in the notebook, give it a name here in the
    # `export_from_notebook` class member
    export_from_notebook = "Excel Spreadsheet"

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
        self.worksheet = self.workbook.add_worksheet()

        self.row = 0
        for cellno, cell in enumerate(nb_copy.cells):
            self.worksheet.write(self.row, 0, str(cellno+1))

            if cell.cell_type == 'markdown':
                self._write_textplain(cell.source)

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
