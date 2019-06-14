class MdXlsStyleRegistry(object):

    default_formats = {
        'double_emphasis': {'bold': True},
        'emphasis': {'italic': True},
        'strikethrough': {'font_strikeout': True},
        'codespan': {'font_name': 'Courier'},
        'h1': {'font_size': 30},
        'h2': {'font_size': 25},
        'h3': {'font_size': 20},
        'h4': {'font_size': 15},
        'h5': {'font_size': 14},
        'h6': {'font_size': 13},
    }

    def __init__(self, workbook):
        self.workbook = workbook
        self.stylereg = {}

    def use_style(self, mdnames):

        if not isinstance(mdnames, list):
            mdnames = [mdnames]

        mdname = '-'.join(mdnames)

        if not mdname in self.stylereg:

            style = self._create_style(mdnames)

            self.stylereg[mdname] = style

        return self.stylereg[mdname]

    def _create_style(self, mdnames):

        d = {}

        for submdname in mdnames:

            if submdname in self.default_formats:
                d = {**d, **self.default_formats[submdname]}

        if len(d) > 0:
            return self.workbook.add_format(d)

        return ''
