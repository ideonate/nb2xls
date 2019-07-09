import os
from nbconvert.exporters.tests.base import ExportersTestsBase
from nb2xls.exporter import XLSExporter

from localxlsxdiff.compare import diff


class LocalExportersTestsBase(ExportersTestsBase):

    def _get_notebook(self, nb_name='notebook.ipynb'):
        return os.path.join(self._get_files_path(), nb_name)

    def _get_files_path(self):

        #Get the relative path to this module in the IPython directory.
        names = self.__module__.split('.')[1:-1]
        names.append('files')

        #Build a path using this directory and the relative path we just
        #found.
        path = os.path.dirname(__file__)
        return os.path.join(path, *names)


class TestsExcelExporter(LocalExportersTestsBase):

    exporter_class = XLSExporter

    def test_constructor(self):
        """
        Can a XLSExporter be constructed?
        """
        XLSExporter()

    def test_export_basic(self):
        """
        Can a XLSExporter export?
        """
        (output, resources) = XLSExporter().from_filename(self._get_notebook('ExcelTest5.ipynb'))
        assert len(output) > 0

    def test_export_compare(self):
        """
        Does a XLSExporter export the same thing as before?
        """
        (other, resources) = XLSExporter().from_filename(self._get_notebook('ExcelTest4.ipynb'))

        with self.create_temp_cwd() as temp_cwd:

            other_fn = os.path.join(temp_cwd,'ExcelTest4.ipynb.xlsx')
            with open(other_fn, "wb") as f:
                f.write(other)

            ref_fn = self._get_notebook('ExcelTest4.ipynb')+'.xlsx'

            wb_diff, sheet_diffs = diff(ref_fn, other_fn)

            assert len(wb_diff) == 0

            assert len(sheet_diffs) == 0
