import os
import pytest

from testpath.tempdir import TemporaryWorkingDirectory
from nb2xls.exporter import XLSExporter

# This should be discoverable by pytest only
from localxlsxdiff.compare import diff


class LocalExportersTestsBase(object):

    """
    Some functions taken from nbconvert.exporters.tests.base
    """

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

    def create_temp_cwd(self, copy_filenames=None):
        return TemporaryWorkingDirectory()


class TestsExcelExporter(LocalExportersTestsBase):

    exporter_class = XLSExporter

    def test_constructor(self):
        """
        Can a XLSExporter be constructed?
        """
        xls = XLSExporter()
        assert xls.export_from_notebook == "Excel Spreadsheet"
        assert xls._file_extension_default() == '.xlsx'

    def test_export_basic(self):
        """
        Can a XLSExporter export?
        """
        (output, resources) = XLSExporter().from_filename(self._get_notebook('ExcelTest5.ipynb'))
        assert len(output) > 0

    @pytest.mark.parametrize("ipynb_filename,expected_size",
                             [
                                ("ExcelTest4.ipynb", 42343),
                                ("NestedMarkdown1.ipynb", 6227),
                                ("ExcelTest.ipynb", 5455),
                                ("PandasNA.ipynb", 6010),
                                ("MarkdownReprDisplay.ipynb", 5557),
                             ])
    def test_export_compare(self, ipynb_filename, expected_size):
        """
        Does a XLSExporter export the same thing as before?
        """
        (other, resources) = XLSExporter().from_filename(self._get_notebook(ipynb_filename))

        if expected_size != -1:
            # Check file size is about right
            len_other = len(other)
            assert len_other >= expected_size-10 and len_other <= expected_size+10

        with self.create_temp_cwd() as temp_cwd:

            other_fn = os.path.join(temp_cwd, ipynb_filename+'.xlsx')
            with open(other_fn, "wb") as f:
                f.write(other)

            ref_fn = self._get_notebook(ipynb_filename)+'.xlsx'

            wb_diff, sheet_diffs = diff(ref_fn, other_fn)

            assert len(wb_diff) == 0

            assert len(sheet_diffs) == 0
