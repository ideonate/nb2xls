
def _get_version(version_info):
    dic = {'alpha': 'a',
           'beta': 'b',
           'candidate': 'rc',
           'dev': 'dev',
           'final': ''}
    vi = version_info
    specifier = '' if vi[3] == 'final' else dic[vi[3]] + str(vi[4])
    version = '%s.%s.%s%s' % (vi[0], vi[1], vi[2], specifier)
    return version


# meta data - change alpha/dev to final for release

version_info = (0, 1, 6, 'final', 0)
__version__ = _get_version(version_info)

