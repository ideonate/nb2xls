import copy
from io import BytesIO
import xlsxwriter

from nbconvert.exporters import Exporter

#-----------------------------------------------------------------------------
# Classes
#-----------------------------------------------------------------------------

class XLSExporter(Exporter):
    """
    My custom exporter
    """

    # If this custom exporter should add an entry to the
    # "File -> Download as" menu in the notebook, give it a name here in the
    # `export_from_notebook` class member
    export_from_notebook = "XLS format"

    def _file_extension_default(self):
        """
        The new file extension is `.test_ext`
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
        if o.output_type == 'stream':
            self._write_textplain('stream : '+o.text)
        if o.output_type in ('display_data', 'execute_result'):
            self._write_textplain('text plain : '+o.data['text/plain'])

    def _write_textplain(self, text):
        lines = text.split("\n")
        print(lines)
        for l in lines:
            self.worksheet.write(self.row, 1, l)
            self.row += 1
