import zipfile
import sys

__version__ = '0.1.3'

from functools import partial
if sys.version_info < (2, 7):
    # add context manager to ZipFile
    from contextlib import contextmanager

    @contextmanager
    def ZipFile(file, *args, **kwargs):
        f = zipfile.ZipFile(file, *args, **kwargs)
        yield f
        f.close()
else:
    from zipfile import ZipFile
from pkg_resources import resource_string
from six import string_types
from lxml import etree
from lxml.builder import ElementMaker

nsmap = {
    'mo': 'http://schemas.microsoft.com/office/mac/office/2008/main',
    'o': 'urn:schemas-microsoft-com:office:office',
    've': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    # Text Content
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w10': 'urn:schemas-microsoft-com:office:word',
    'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
    # Drawing
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'mv': 'urn:schemas-microsoft-com:mac:vml',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'v': 'urn:schemas-microsoft-com:vml',
    'wp': ('http://schemas.openxmlformats.org/drawingml/2006/'
           'wordprocessingDrawing'),
    # Properties (core and extended)
    'cp': ('http://schemas.openxmlformats.org/package/2006/metadata/'
           'core-properties'),
    'dc': 'http://purl.org/dc/elements/1.1/',
    'ep': ('http://schemas.openxmlformats.org/officeDocument/2006/'
           'extended-properties'),
    'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
    # Content Types
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
    # Package Relationships
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'pr': 'http://schemas.openxmlformats.org/package/2006/relationships',
    # Dublin Core document properties
    'dcmitype': 'http://purl.org/dc/dcmitype/',
    'dcterms':  'http://purl.org/dc/terms/'
}


def qname(namespace, name):
    '''decorate the name with fully qualified namespace.'''
    namespace = nsmap.get(namespace, namespace)
    assert(namespace in nsmap.values())
    return '{%s}%s' % (namespace, name)

typemap = {}


def add_dict(elem, item):
    attrib = elem.attrib
    for k, v in item.items():
        # add the namespace prefix to attribute key
        if k.startswith('{'):
            pass
        elif elem.prefix:
            k = '{%s}%s' % (elem.nsmap[elem.prefix], k)

        if isinstance(v, string_types):
            attrib[k] = v
        else:
            attrib[k] = typemap[type(v)](None, v)
typemap[dict] = add_dict

E = ElementMaker(namespace=nsmap['w'], nsmap=nsmap, typemap=typemap)


def run(text, style=None):
    """
    Returns a ``r`` (text run) element with the specified style for the text.

    *text* is either a string or ``t`` (text) element or ``br`` (break)
    element.

    *style*, if specified, MUST be a ``rPr`` (run property) element or the
    combination of the following font styles:

    .. _font-style-table:

    +-------+---------------------+
    | style | Font style rendered |
    +=======+=====================+
    | b     | bold                |
    +-------+---------------------+
    | i     | italic              |
    +-------+---------------------+
    | u     | underline           |
    +-------+---------------------+

    For example::

        run('bold and italic', 'bi')

    """
    # TODO: the atomic block should be t, we will revise this if we want
    # support wTabBefore and wTabAfter.
    run = E.r()
    if hasattr(style, 'tag') and style.tag == qname('w', 'rPr'):
        run.append(style)
    elif style is not None:
        run.append(E.rPr(*map(E, style)))

    if hasattr(text, 'tag') and text.tag in (
            qname('w', 't'), qname('w', 'br')):
        run.append(text)
    else:
        run.append(E.t(text))
    return run


def paragraph(runs, style=None):
    """
    Returns a ``p`` (paragraph) element with specified style for text runs.

    *runs* are a list of ``r`` (text run) element, see :func:`run`.

    *style*, if specified, MUST be a ``pPr`` (paragraph property) element with
    nested ``pStyle`` (paragraph style) or a string for ``pStyle``. Commonly
    used styles include: ``ListParagraph``, ``Heading1``, ``Title``,
    ``Subtitle`` and ``Code``.

    """
    para = E.p()
    if hasattr(style, 'tag') and style.tag == qname('w', 'pPr'):
        para.append(style)
    elif style is not None:
        para.append(
            E.pPr(
                E.pStyle(val=style)
            )
        )
    para.extend(runs)
    return para


def li(runs, style):
    """
    Returns a ``p`` (paragraph) with ordered or unordered list style for
    text runs.

    *runs* are a list of ``r`` (run) element, see :func:`run`.

    *style*, an string of supported list style: ``circle``, ``number``,
    ``disc``and ``square``.

    """
    listmap = {
        'circle': '1',
        'number': '2',
        'disc': '3',
        'square': '4',
    }
    assert style in listmap
    # TODO: support nested list
    return paragraph(
        runs,
        E.pPr(
            E.pStyle(val='ListParagraph'),
            E.numPr(
                E.ilvl(val='0'),
                E.numId(val=listmap[style])
            )
        )
    )

title = partial(paragraph, style='Title')
subtitle = partial(paragraph, style='Subtitle')
h1 = partial(paragraph, style='Heading1')
h2 = partial(paragraph, style='Heading2')
h3 = partial(paragraph, style='Heading3')


def spaces(count=1):
    """
    Returns a ``t`` (text) element with the number of *count* of spaces.

    """
    t = E.t(' ' * count)
    t.set("{http://www.w3.org/XML/1998/namespace}space", 'preserve')
    return run(t)


def pagebreak():
    """
    Return a ``br`` (break) element with ``page`` type, aka a pagebreak.

    """
    return paragraph([run(E.br(type='page'))])


def table(cells, style=None):
    """
    Returns a ``tbl`` (table) element with specified style from the cells.

    *style*, if specified, MUST be a ``tblPr`` (table property) element or a
    string of supported table style defined in the theme.

    *cells* contains the table content in the two-dimension array fashion.
    Each cell MUST be a ``tc`` (table cell) element; or a paragraph or a text
    run.

    """
    assert(len(cells) > 0)
    assert(len(cells[0]) > 0)

    tbl = E.tbl()
    if hasattr(style, 'tag') and style.tag == qname('w', 'tblPr'):
        run.append(style)
    else:
        tbl.append(
            E.tblPr(
                E.tblStyle(val=style)
            )
        )
    # TODO: support tblGrid

    # iterate all rows
    for row in cells:
        tr = E.tr(
            E.trPr(
                E.cnfStyle(val='000000100000')
            )
        )
        for cell in row:
            assert hasattr(cell, 'tag')
            if cell.tag == qname(cell.prefix, 'r'):
                cell = paragraph([cell])

            if cell.tag == qname(cell.prefix, 'p'):
                cell = E.tc(
                    E.tcPr(
                        E.tcW(w='0', type='auto')
                    ),
                    cell
                )

            if cell.tag == qname(cell.prefix, 'tc'):
                tr.append(cell)
        tbl.append(tr)
    return tbl


class Document(object):
    """
    A Document instance contains all the parts of Microsoft Word document.
    """
    def __init__(self, doc=None):
        self.doc = doc or E.document(
            E.body()
        )
        self.meta = {}

    @property
    def body(self):
        """
        returns the document.body.
        """
        return self.doc[0]

    def update(self, *args, **kwargs):
        """
        update the core properties of the document.
        """
        self.meta.update(*args, **kwargs)

    @classmethod
    def load(cls, f):
        with ZipFile(f) as zippy:
            root = etree.parse(zippy.open('word/document.xml'))
            return cls(root.getroot())

    def dumps(self, pretty_print=False):
        """
        Serialize the main document stream to a XML formated :class:`str`.

        If *pretty_print* is ``True`` (default: ``False``), then the XML
        elements will be pretty-printed with indention.
        """
        return etree.tostring(self.doc, pretty_print=pretty_print)

    def get_core_props(self):
        _nsmap = dict((k, v) for k, v in nsmap.items() if k in (
            'cp', 'dc', 'dcterms', 'dcmitype', 'xsi'))
        core = etree.Element(qname('cp', 'coreProperties'), nsmap=_nsmap)

        if 'lastModifiedBy' in self.meta:
            key = 'lastModifiedBy'
            el = etree.SubElement(core, qname('cp', key))
            el.text = self.meta[key]

        if 'keywords' in self.meta:
            key = 'keywords'
            el = etree.SubElement(core, qname('cp', key))
            if isinstance(self.meta[key], string_types):
                el.text = self.meta[key]
            else:
                el.text = ','.join(self.meta[key])

        for key in ('title', 'subject', 'creator', 'description'):
            if key in self.meta:
                el = etree.SubElement(core, qname('dc', key))
                el.text = self.meta[key]
        for key in ('created', 'modified'):
            if key in self.meta:
                el = etree.SubElement(core, qname('dcterms', key))
                el.set(qname('xsi', 'type'), 'dcterms:W3CDTF')
                el.text = self.meta[key].strftime('%Y-%m-%dT%H:%M:%SZ')
        return core

    def save(self, fp, pretty_print=False):
        """
        Serialize all document parts to *fp* (a :func:`.write()`-supporting
        file-like object) or a pathname.

        If *pretty_priint* is ``True`` (default: ``False``), then the XML
        elements will be pretty-printed with indention.

        """
        parts = ('[Content_Types].xml',
                 '_rels/.rels',
                 'docProps/app.xml',
                 'word/fontTable.xml',
                 'word/numbering.xml',
                 'word/settings.xml',
                 'word/styles.xml',
                 'word/stylesWithEffects.xml',
                 'word/webSettings.xml',
                 'word/theme/theme1.xml',
                 'word/_rels/document.xml.rels',
                 )
        with ZipFile(fp, mode='w', compression=zipfile.ZIP_DEFLATED) as zippy:
            for part in parts:
                bytes = resource_string(__name__, 'templates/%s' % part)
                # CODE DEBT: use stream?
                zippy.writestr(part, bytes)

            # serialize the document.xml
            zippy.writestr('word/document.xml', etree.tostring(
                self.doc, xml_declaration=True, standalone=True,
                encoding='UTF-8', pretty_print=pretty_print))

            # serialize docProps/core.xml
            zippy.writestr(
                'docProps/core.xml',
                etree.tostring(
                    self.get_core_props(),
                    pretty_print=pretty_print)
            )

            # TODO: save word/_rels/document.xml.rels if blip is supported
