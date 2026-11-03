"""Small OOXML package helpers; no third-party dependencies."""
from pathlib import Path
from xml.dom import minidom
import zipfile

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
STORY_NAMES = {'document.xml', 'footnotes.xml', 'endnotes.xml', 'comments.xml', 'glossary/document.xml'}
ORDERS = {
    'docDefaults': 'rPrDefault pPrDefault',
    'rPr': 'rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath rPrChange',
    'tblPr': 'tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize tblStyleColBandSize tblW jc tblCellSpacing tblInd tblBorders shd tblLayout tblCellMar tblLook tblCaption tblDescription tblPrChange',
    'trPr': 'cnfStyle divId gridBefore gridAfter wBefore wAfter cantSplit trHeight tblHeader tblCellSpacing jc hidden ins del trPrChange',
    'tcPr': 'cnfStyle tcW gridSpan hMerge vMerge tcBorders shd noWrap tcMar textDirection tcFitText vAlign hideMark cellIns cellDel cellMerge tcPrChange',
    'pPr': 'pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange',
}


def elements(node, local=None, namespace=W):
    return [n for n in node.getElementsByTagNameNS(namespace, local or '*')]


def children(node, local=None, namespace=W):
    return [n for n in node.childNodes if n.nodeType == n.ELEMENT_NODE
            and n.namespaceURI == namespace and (local is None or n.localName == local)]


def table_rows(table):
    result = []
    for row in elements(table, 'tr'):
        owner = row.parentNode
        while owner is not None and not (getattr(owner, 'namespaceURI', None) == W and getattr(owner, 'localName', None) == 'tbl'):
            owner = owner.parentNode
        if owner is table:
            result.append(row)
    return result


def attr(node, name, default=''):
    return node.getAttributeNS(W, name) if node.hasAttributeNS(W, name) else default


def put(node, name, value):
    node.setAttributeNS(W, 'w:' + name, str(value))


def drop(node):
    node.parentNode.removeChild(node)


def ensure(node, name, first=False):
    found = children(node, name)
    if found:
        return found[0]
    new = node.ownerDocument.createElementNS(W, 'w:' + name)
    order = ORDERS.get(node.localName, '').split()
    index = order.index(name) if name in order else len(order)
    before = next((n for n in children(node) if n.localName in order
                   and order.index(n.localName) > index), None)
    if before:
        node.insertBefore(new, before)
    elif first and node.firstChild:
        node.insertBefore(new, node.firstChild)
    else:
        node.appendChild(new)
    return new


def is_story(name):
    if not name.startswith('word/') or not name.endswith('.xml'):
        return False
    tail = name[5:]
    return tail in STORY_NAMES or tail.startswith(('header', 'footer', 'dcHeader', 'dcFooter', 'dcFirstFooter'))


def read_package(path):
    with zipfile.ZipFile(path) as z:
        infos = z.infolist()
        names = [i.filename for i in infos]
        if len(names) != len(set(names)):
            raise ValueError('Duplicate ZIP entries are not supported')
        if '[Content_Types].xml' not in names or 'word/document.xml' not in names:
            raise ValueError('Not a Word OOXML package')
        if sum(i.file_size for i in infos) > 200 * 1024 * 1024:
            raise ValueError('Unpacked DOCX exceeds 200 MiB; inspect it with a suitable tool')
        return infos, {i.filename: z.read(i.filename) for i in infos}, z.comment


def parse(data):
    # Reject DTDs; these helpers only accept ordinary OOXML parts.
    if b'<!DOCTYPE' in data.upper() or b'<!ENTITY' in data.upper():
        raise ValueError('DTD/entity declarations are not supported')
    return minidom.parseString(data)


def xml_bytes(doc):
    return doc.toxml(encoding='UTF-8')


def text_signature(data):
    doc = parse(data)
    tags = {'t', 'delText', 'instrText', 'tab', 'br', 'cr', 'fldChar', 'sym'}
    return [(n.localName, n.toxml()) for n in elements(doc) if n.localName in tags]


def format_signature(data):
    doc = parse(data)
    names = {'pPr', 'rPr', 'tblPr', 'trPr', 'tcPr', 'sectPr', 'background'}
    return [n.toxml() for n in elements(doc) if n.localName in names]


def write_package(output, infos, parts, comment):
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    # 'x' prevents accidental overwriting, including a concurrently created file.
    created = False
    try:
        with output.open('xb') as stream:
            created = True
            with zipfile.ZipFile(stream, 'w') as z:
                z.comment = comment
                for info in infos:
                    z.writestr(info, parts[info.filename])
    except Exception:
        # A failed write should not leave a plausible but partial output.
        if created and output.exists():
            output.unlink()
        raise
