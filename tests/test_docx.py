from datetime import datetime
from zipfile import ZipFile
from io import BytesIO
from lxml import etree
from docxgen import *
from docxgen import nsmap
from . import check_tag

def test_init():
    doc = Document()
    check_tag(doc.doc, ['document', 'body'])
    check_tag(doc.body, ['body'])

def test_save():
    doc = Document()
    tmp = BytesIO()
    doc.save(tmp)

    with ZipFile(tmp) as zippy:
        assert(zippy.testzip() is None)
        assert(set(zippy.namelist()) == set([
            '[Content_Types].xml', '_rels/.rels', 'docProps/app.xml',
            'word/fontTable.xml', 'word/numbering.xml', 'word/settings.xml',
            'word/styles.xml', 'word/stylesWithEffects.xml',
            'word/webSettings.xml', 'word/theme/theme1.xml',
            'word/_rels/document.xml.rels', 'word/document.xml',
            'docProps/core.xml'
        ]))
        zxf = zippy.open('word/document.xml')
        root = etree.parse(zxf)
        check_tag(root, 'document body'.split())

def test_load():
    # assume save works
    d = Document()
    tmp = BytesIO()
    d.save(tmp)

    doc = Document.load(tmp)
    check_tag(doc.doc, 'document body'.split())
    check_tag(doc.body, 'body'.split())

def test_dumps():
    doc = Document()
    assert doc.dumps()

def test_core_props():
    doc = Document()
    attrs = dict(lastModifiedBy='Joe Smith',
        keywords=['egg', 'spam'],
        title='Testing Doc',
        subject='A boilerplate document',
        creator='Jill Smith',
        description='Core properties test.',
        created=datetime.now()
    )
    doc.update(attrs)
    core = doc.get_core_props()

    for key in ('title', 'subject', 'creator', 'description'):
        attr = core.find('.//dc:%s' % key, namespaces=nsmap)
        assert attr is not None
        assert attr.text == attrs[key]

    attr = core.find('.//cp:keywords', namespaces=nsmap)
    assert attr is not None
    assert attr.text == ','.join(attrs['keywords'])

    attr = core.find('.//dcterms:created', namespaces=nsmap)
    assert attr is not None
    assert datetime.strptime(attr.text, '%Y-%m-%dT%H:%M:%SZ') == datetime(*attrs['created'].timetuple()[:6])
