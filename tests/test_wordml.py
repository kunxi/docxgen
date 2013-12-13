from re import split
from docxgen import *
from docxgen import E
from . import check_tag

def test_run():
    for style, expected in [
            (None, ['r', 't']),
            ('b', ['r', 'rPr', 'b', 't']),
            ('i', ['r', 'rPr', 'i', 't']),
            (['b', 'i'], ['r', 'rPr', 'b', 'i', 't']),
            ('bu', ['r', 'rPr', 'b', 'u', 't']),
            ('bi', ['r', 'rPr', 'b', 'i', 't']),
        ]:
        yield check_tag, run('sample text', style), expected

    colored = E.rPr(
        E.color(val="C0504D", themeColor="accent2")
    )
    yield check_tag, run('red text', colored), ['r', 'rPr', 'color', 't']

def test_paragraph():
    check_tag(paragraph([run('Example')]), ['p', 'r', 't'])

def test_list_item():
    for style in ['circle', 'number', 'square', 'disc']:
        root = li([run('item')], style)
        yield check_tag, root, ['p', 'pPr', 'pStyle', 'numPr', 'ilvl', 'numId', 'r', 't']
        numId = root.find('.//w:numId', namespaces=nsmap)
        assert (numId.get(qname('w', 'val')) in '1234')

def test_heading():
    for _heading, style in zip((h1, h2, h3, title, subtitle),
            ('Heading1', 'Heading2', 'Heading3', 'Title', 'Subtitle')):
        root = _heading([run('Heading')])
        yield check_tag, root, ['p', 'pPr', 'pStyle', 'r', 't']
        pstyle = root.find('.//w:pStyle', namespaces=nsmap)
        assert pstyle.get(qname('w', 'val')) == style

def test_spaces():
    root = spaces(3)
    check_tag(root, ['r', 't'])
    t = root[0]
    assert t.get("{http://www.w3.org/XML/1998/namespace}space") == 'preserve'

def test_pagebreak():
    root = pagebreak()
    check_tag(root, ['p', 'r', 'br'])

def test_table():
    root = table([
        [run('2'), run('7'), run('6')],
        [run('9'), paragraph([run('5')]), run('1')],
        [run('4'), run('3'), run('9')],
    ], 'LightShading-Accent1')
    check_tag(root, split(r'\s+', '''tbl tblPr tblStyle
    tr trPr cnfStyle tc tcPr tcW p r t tc tcPr tcW p r t tc tcPr tcW p r t
    tr trPr cnfStyle tc tcPr tcW p r t tc tcPr tcW p r t tc tcPr tcW p r t
    tr trPr cnfStyle tc tcPr tcW p r t tc tcPr tcW p r t tc tcPr tcW p r t
    '''))
