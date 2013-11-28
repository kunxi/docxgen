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
        ]:
        yield check_tag, run('sample text', style), expected

    colored = E.rPr(
        E.color(val="C0504D", themeColor="accent2")
    )
    yield check_tag, run('red text', colored), ['r', 'rPr', 'color', 't']

def test_paragraph():
    check_tag(paragraph(None, run('Example')), ['p', 'r', 't'])

def test_list_item():
    for style in ['circle', 'number', 'square', 'disc']:
        root = li(style, run('item'))
        yield check_tag, root, ['p', 'pPr', 'pStyle', 'numPr', 'ilvl', 'numId', 'r', 't']
        numId = root.find('.//w:numId', namespaces=nsmap)
        assert (numId.get(qname('w', 'val')) in '1234')