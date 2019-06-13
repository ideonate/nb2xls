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
        nb_copy, resources = self._preprocess(nb_copy, resources)

        output = BytesIO()
        workbook = xlsxwriter.Workbook(output)
        worksheet = workbook.add_worksheet()

        worksheet.write('A1', 'Hello')
        workbook.close()

        xlsx_data = output.getvalue()

        return xlsx_data, resources
