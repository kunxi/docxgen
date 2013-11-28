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

E = ElementMaker(namespace=nsmap['w'], nsmap=nsmap)


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
