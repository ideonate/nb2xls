# nb2xls - Jupyter notebooks to Excel Spreadsheets

Convert Jupyter notebooks to Excel Spreadsheets, through a new 'Download As' option or via nbconvert on the command 
line.

Respects tables such as Pandas DataFrames. Also exports image data such as matplotlib output.

Markdown also supported where possible (some elements still need work).

This allows you to share your results with non-programmers such that they can still easily play with the data.

## Installation

### Install via pip (recommended)

```
pip install nb2xls
```

### Development Mode

If you want to contribute or debug:

```
git clone https://github.com/ideonate/nb2xls
cd nb2xls
pip install -e .
```

Restart Jupyter to pick up the new 'Excel Spreadsheet (.xlsx)' option under 'Download As' in the File menu.

## Usage

In Jupyter Notebook, just select the 'Excel Spreadsheet (.xlsx)' option under 'Download As' in the File menu.

To run from the command line try:

```
jupyter nbconvert --to nb2xls.XLSExporter Examples/ExcelTest1.ipynb
```

This should output ExcelTest1.xlsx in the same folder as the ipynb file specified.

## Contact for Feedback

Please get in touch with any feedback or questions: [dan@ideonate.com](dan@ideonate.com). It is very helpful to send 
example notebooks, especially if you have bug reports or feature suggestions.

## License

This code is released under an MIT license.