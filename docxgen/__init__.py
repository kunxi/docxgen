import zipfile
import sys

from functools import partial
import sys
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
from lxml import etree
from lxml.builder import ElementMaker

nsmap = {
    'mo': 'http://schemas.microsoft.com/office/mac/office/2008/main',
    'o':  'urn:schemas-microsoft-com:office:office',
    've': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    # Text Content
    'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w10': 'urn:schemas-microsoft-com:office:word',
    'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
    # Drawing
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'm':   'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'mv':  'urn:schemas-microsoft-com:mac:vml',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'v':   'urn:schemas-microsoft-com:vml',
    'wp':  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    # Properties (core and extended)
    'cp':  'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    'dc':  'http://purl.org/dc/elements/1.1/',
    'ep':  'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
    # Content Types
    'ct':  'http://schemas.openxmlformats.org/package/2006/content-types',
    # Package Relationships
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'pr':  'http://schemas.openxmlformats.org/package/2006/relationships',
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

        if isinstance(v, basestring):
            attrib[k] = v
        else:
            attrib[k] = typemap[type(v)](None, v)
typemap[dict] = add_dict

E = ElementMaker(namespace=nsmap['w'], nsmap=nsmap, typemap=typemap)

def run(text='', style=None):
    '''the smallest build block.'''
    # CODE DEBT: the atomic block should be t, we will revise this if we want
    # support wTabBefore and wTabAfter.
    run = E.r()
    if style is not None:
        # bold and/or italic
        if isinstance(style, etree._Element):
            run.append(style)
        else:
            if isinstance(style, basestring):
                style = [style]
            run.append(E.rPr(*[E(x) for x in style]))

    if isinstance(text, basestring):
        text = E.t(text)
    run.append(text)
    return run

def paragraph(style, *runs):
    '''
    Render a paragraph with specified style for text runs.
    @params style: a string or pPr element
    @params runs: a list of run element
    '''
    para = E.p()
    if isinstance(style, basestring):
        para.append(
            E.pPr(
                E.pStyle(val=style)
            )
        )
    elif style is not None:
        para.append(style)

    para.extend(runs)
    return para

def li(style, *runs):
    '''
    Render a list item with text runs.
    @params style: string, enum of circle, square and number.
    @params runs: a list of run element
    '''
    listmap = {
        'circle': '1',
        'number': '2',
        'disc': '3',
        'square': '4',
    }
    assert style in listmap
    return paragraph(
        E.pPr(
            E.pStyle(val='ListParagraph'),
            E.numPr(
                E.ilvl(val='0'),
                E.numId(val=listmap[style])
            )
        ), *runs)

def heading(level, *runs):
    return paragraph('Heading%d' % level, *runs)

h1 = partial(heading, 1)
h2 = partial(heading, 2)
h3 = partial(heading, 3)
title = partial(paragraph, 'Title')
sub_title = partial(paragraph, 'Subtitle')

def space(num=1):
    '''Insert some spaces'''
    t = E.t(' ' * num)
    t.set("{http://www.w3.org/XML/1998/namespace}space", 'preserve')
    return run(t)

def pagebreak():
    '''Insert a hard pagebreak.'''
    return paragraph(None,
        run(
            E.br(type='page')
        )
    )

def table(cells, style=None):
    '''Render a table with cells

    @style: global table style
    @cells: content in 2D array format. If the cell is plain text, using the
            default cell style, else we assume the cell is hand-crafted wordml
            element, we just append it to the row
    '''
    assert(len(cells) > 0)
    assert(len(cells[0]) > 0)

    tbl = E.tbl()
    if isinstance(style, basestring):
        tbl.append(
            E.tblPr(
                E.tblStyle(val=style)
            )
        )
    elif style is not None:
        tbl.append(style)
    # TODO: support tblGrid

    # iterate all rows
    for row in cells:
        tr = E.tr(
            E.trPr(
                E.cnfStyle(val='000000100000')
            )
        )
        for cell in row:
            if isinstance(cell, basestring):
                cell = paragraph(None, run(cell))

            if cell.tag == qname(cell.prefix, 'tc'):
                tr.append(cell)
            else:
                tc = E.tc(
                    E.tcPr(
                        E.tcW(w='0', type='auto')
                    ),
                )
                if cell.tag == qname(cell.prefix, 'p'):
                    tc.append(cell)
                elif cell.tag == qname(cell.prefix, 'r'):
                    tc.append(paragraph(None, cell))
                tr.append(tc)
        tbl.append(tr)
    return tbl


class Document(object):
    '''Encapsulate the docx serialization.'''
    def __init__(self, doc):
        self.doc = doc
        self.meta = {}

    @property
    def body(self):
        return self.doc[0]

    def update(self, *args, **kwargs):
        self.meta.update(*args, **kwargs)

    @classmethod
    def create(cls):
        doc = E.document(
            E.body()
        )
        return cls(doc)

    @classmethod
    def load(cls, file):
        with ZipFile(file) as zippy:
            root = etree.parse(zippy.open('word/document.xml'))
            return cls(root.getroot())

    def dumps(self, pretty_print=False):
        '''Dumps the main document.'''
        return etree.tostring(self.doc, pretty_print=pretty_print)

    def get_core_props(self):
        _nsmap = dict((k, v) for k, v in nsmap.iteritems() if k in ('cp', 'dc', 'dcterms', 'dcmitype', 'xsi'))
        core = etree.Element(qname('cp', 'coreProperties'), nsmap=_nsmap)

        if 'lastModifiedBy' in self.meta:
            key = 'lastModifiedBy'
            el = etree.SubElement(core, qname('cp', key))
            el.text = self.meta[key]

        if 'keywords' in self.meta:
            key = 'keywords'
            el = etree.SubElement(core, qname('cp', key))
            if isinstance(self.meta[key], basestring):
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

    def save(self, file, pretty_print=False):
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
        with ZipFile(file, mode='w', compression=zipfile.ZIP_DEFLATED) as zippy:
            for part in parts:
                bytes = resource_string(__name__, 'templates/%s' % part)
                # CODE DEBT: use stream?
                zippy.writestr(part, bytes)

            # serialize the document.xml
            zippy.writestr('word/document.xml', etree.tostring(
                self.doc, xml_declaration=True, standalone=True,
                encoding='UTF-8', pretty_print=pretty_print))

            # serialize docProps/core.xml
            zippy.writestr('docProps/core.xml', etree.tostring(self.get_core_props(), pretty_print=pretty_print))

            # CODE DEBT: serialize word/_rels/document.xml.rels if we support blip
