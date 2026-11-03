#!/usr/bin/env python3
"""Apply the doc-format page/type profile to a NEW DOCX or explicit restyle.
Not for a targeted edit. Supports one ordinary, final body section only.
"""
import argparse
from pathlib import Path
import re
import sys
import zipfile
from xml.dom import minidom
from docx_xml import (W, elements, children, attr, put, ensure, drop, parse,
                      xml_bytes, read_package, write_package, is_story, ORDERS)

R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
MARKER = 'dcBlankFirstPage'
FONT = 'Times New Roman'
ORDERS['sectPr'] = 'headerReference footerReference footnotePr endnotePr type pgSz pgMar paperSrc pgBorders lnNumType pgNumType cols formProt vAlign noEndnote titlePg textDirection bidi rtlGutter docGrid printerSettings sectPrChange'
ORDERS['footnotePr'] = 'pos numFmt numStart numRestart footnote'


def twips(mm):
    return str(round(float(mm) * 1440 / 25.4))


def flag(pr, name, value):
    put(ensure(pr, name), 'val', '1' if value else '0')


def font_props(pr, size, bold=False):
    fonts = ensure(pr, 'rFonts')
    for key in list(fonts.attributes.keys()):
        if fonts.attributes[key].namespaceURI == W:
            fonts.removeAttribute(key)
    for key in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
        put(fonts, key, FONT)
    put(ensure(pr, 'sz'), 'val', int(size * 2))
    put(ensure(pr, 'szCs'), 'val', int(size * 2))
    flag(pr, 'b', bold)
    flag(pr, 'bCs', bold)


def paragraph_spacing(pr, single=False):
    spacing = ensure(pr, 'spacing')
    put(spacing, 'line', '240' if single else '360')
    put(spacing, 'lineRule', 'auto')


def style_kind(style_id, style_map=None, seen=None):
    normalized = re.sub(r'[\s_-]+', '', style_id).lower()
    if re.fullmatch(r'(heading|заголовок)[1-9]', normalized):
        return 'heading'
    if normalized in ('caption', 'название', 'подпись'):
        return 'caption'
    if normalized in ('footer', 'нижнийколонтитул'):
        return 'footer'
    if normalized in ('footnotetext', 'footnotereference', 'текстсноски', 'знаксноски'):
        return 'footnote'
    if style_map and style_id in style_map:
        seen = set() if seen is None else seen
        if style_id not in seen:
            seen.add(style_id)
            node = style_map[style_id]
            for name in children(node, 'name'):
                direct = style_kind(attr(name, 'val'))
                if direct != 'body':
                    return direct
            based = children(node, 'basedOn')
            if based:
                return style_kind(attr(based[0], 'val'), style_map, seen)
    return 'body'


def paragraph_kind(para, styles):
    props = children(para, 'pPr')
    refs = children(props[0], 'pStyle') if props else []
    return style_kind(attr(refs[0], 'val'), styles) if refs else 'body'


def in_header_row(node):
    parent = node.parentNode
    while parent is not None:
        if getattr(parent, 'namespaceURI', None) == W and getattr(parent, 'localName', None) == 'tr':
            props = children(parent, 'trPr')
            markers = children(props[0], 'tblHeader') if props else []
            return bool(markers and attr(markers[0], 'val', '1') not in ('0', 'false', 'off'))
        parent = parent.parentNode
    return False


def inside_table(node):
    parent = node.parentNode
    while parent is not None:
        if getattr(parent, 'namespaceURI', None) == W and getattr(parent, 'localName', None) == 'tc':
            return True
        parent = parent.parentNode
    return False


def first_line_target(para, styles):
    """Body prose gets 12.5 mm; lists retain their own numbering indentation."""
    pp = children(para, 'pPr')
    if pp and children(pp[0], 'numPr'):
        return None
    refs = children(pp[0], 'pStyle') if pp else []
    ident = attr(refs[0], 'val') if refs else ''
    seen = set()
    special = paragraph_kind(para, styles) != 'body'
    while ident and ident not in seen:
        seen.add(ident)
        style = styles.get(ident)
        names = [ident] + ([attr(n, 'val') for n in children(style, 'name')] if style else [])
        normalized = [re.sub(r'[\s_-]+', '', n).lower() for n in names]
        if any(n.startswith(('list', 'список')) for n in normalized):
            return None
        special = special or any(n in ('title', 'subtitle', 'заглавие', 'подзаголовок')
                                 or n.startswith(('toc', 'оглавление')) for n in normalized)
        sp = children(style, 'pPr') if style else []
        if sp and children(sp[0], 'numPr'):
            return None
        base = children(style, 'basedOn') if style else []
        ident = attr(base[0], 'val') if base else ''
    math = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
    has_math = bool(para.getElementsByTagNameNS(math, 'oMath')
                    or para.getElementsByTagNameNS(math, 'oMathPara'))
    text = any(n.firstChild and n.firstChild.nodeValue.strip() for n in elements(para, 't'))
    return '0' if special or inside_table(para) or has_math or not text else twips(12.5)


def first_line_indent(pr, target):
    if target is None:
        return
    ind = ensure(pr, 'ind')
    for key in ('firstLineChars', 'hanging', 'hangingChars'):
        if ind.hasAttributeNS(W, key):
            ind.removeAttributeNS(W, key)
    put(ind, 'firstLine', target)


def normalize_story(doc, body_size, styles, footnotes=False, footer=False):
    doc.documentElement.setAttribute('xmlns:w', W)
    for para in elements(doc, 'p'):
        kind = 'footnote' if footnotes else paragraph_kind(para, styles)
        size = 10 if kind == 'footnote' else 12 if kind in ('caption', 'footer') or footer or inside_table(para) else body_size
        pr = ensure(para, 'pPr', first=True)
        paragraph_spacing(pr, single=footnotes or footer)
        target = first_line_target(para, styles) if doc.documentElement.localName == 'document' else '0'
        first_line_indent(pr, target)
        font_props(ensure(pr, 'rPr'), size, kind == 'heading')
        if kind == 'caption':
            put(ensure(pr, 'jc'), 'val', 'center')
            flag(ensure(pr, 'rPr'), 'i', False)
            flag(ensure(pr, 'rPr'), 'iCs', False)
        for run in elements(para, 'r'):
            # Nested content (e.g. a text box) keeps its own paragraph context.
            owner = run.parentNode
            while owner is not None and getattr(owner, 'localName', None) != 'p':
                owner = owner.parentNode
            if owner is not para:
                continue
            rp = ensure(run, 'rPr', first=True)
            font_props(rp, size, (kind == 'heading' or in_header_row(run)) and not footer)
            if footer:
                color = ensure(rp, 'color')
                for key in list(color.attributes.keys()):
                    if color.attributes[key].namespaceURI == W:
                        color.removeAttribute(key)
                put(color, 'val', '000000')
            if kind == 'caption':
                flag(rp, 'i', False)
                flag(rp, 'iCs', False)
            if children(run, 'footnoteReference') or children(run, 'footnoteRef'):
                put(ensure(rp, 'rStyle'), 'val', 'FootnoteReference')
                put(ensure(rp, 'vertAlign'), 'val', 'superscript')


def normalize_styles(doc, body_size):
    root = doc.documentElement
    root.setAttribute('xmlns:w', W)
    styles = {attr(s, 'styleId'): s for s in elements(doc, 'style')}
    for ident, typ, name in (('FootnoteReference', 'character', 'footnote reference'), ('FootnoteText', 'paragraph', 'footnote text'), ('Footer', 'paragraph', 'footer')):
        if ident not in styles:
            node = doc.createElementNS(W, 'w:style')
            put(node, 'type', typ); put(node, 'styleId', ident)
            put(ensure(node, 'name'), 'val', name)
            root.appendChild(node)
            styles[ident] = node
    defaults = ensure(root, 'docDefaults', first=True)
    font_props(ensure(ensure(defaults, 'rPrDefault'), 'rPr'), body_size)
    paragraph_spacing(ensure(ensure(defaults, 'pPrDefault'), 'pPr'))
    for style_id, node in styles.items():
        kind = style_kind(style_id, styles)
        size = 10 if kind == 'footnote' else 12 if kind in ('caption', 'footer') or attr(node, 'type') == 'table' else body_size
        font_props(ensure(node, 'rPr'), size, kind == 'heading')
        if kind == 'caption':
            flag(ensure(node, 'rPr'), 'i', False)
            flag(ensure(node, 'rPr'), 'iCs', False)
        if style_id == 'FootnoteReference':
            put(ensure(ensure(node, 'rPr'), 'vertAlign'), 'val', 'superscript')
        if attr(node, 'type') == 'paragraph':
            paragraph_spacing(ensure(node, 'pPr'), kind in ('footnote', 'footer'))
    return styles


def blank_marker(body):
    direct = children(body)
    markers = [n for n in elements(body, 'bookmarkStart') if attr(n, 'name') == MARKER]
    if markers:
        if len(markers) != 1 or not direct or direct[0].localName != 'p' or markers[0].parentNode is not direct[0]:
            raise ValueError('Blank-page marker exists outside the expected first paragraph; inspect the document')
        return
    first = body.ownerDocument.createElementNS(W, 'w:p')
    existing_ids = [int(attr(n, 'id')) for n in elements(body, 'bookmarkStart') if attr(n, 'id').isdigit()]
    ident = str(max(existing_ids, default=-1) + 1)
    start = ensure(first, 'bookmarkStart'); put(start, 'id', ident); put(start, 'name', MARKER)
    put(ensure(ensure(first, 'r'), 'br'), 'type', 'page')
    put(ensure(first, 'bookmarkEnd'), 'id', ident)
    body.insertBefore(first, body.firstChild)


def fresh_part_name(parts, stem):
    index = 1
    while 'word/dc' + stem + str(index) + '.xml' in parts:
        index += 1
    return 'word/dc' + stem + str(index) + '.xml'


def add_part(parts, infos, types, rels, name, data, content_type, relation):
    parts[name] = data
    info = zipfile.ZipInfo(name)
    info.compress_type = zipfile.ZIP_DEFLATED
    infos.append(info)
    over = types.createElementNS(CT, (types.documentElement.prefix + ':' if types.documentElement.prefix else '') + 'Override')
    over.setAttribute('PartName', '/' + name)
    over.setAttribute('ContentType', content_type)
    types.documentElement.appendChild(over)
    used = {n.getAttribute('Id') for n in rels.getElementsByTagNameNS(REL, 'Relationship')}
    index = 1
    while 'rId' + str(index) in used:
        index += 1
    ident = 'rId' + str(index)
    rel = rels.createElementNS(REL, (rels.documentElement.prefix + ':' if rels.documentElement.prefix else '') + 'Relationship')
    for key, value in {'Id': ident, 'Type': R + '/' + relation, 'Target': name[5:]}.items():
        rel.setAttribute(key, value)
    rels.documentElement.appendChild(rel)
    return ident


def footnote_properties(parent):
    pr = ensure(parent, 'footnotePr')
    for name, value in (('numFmt', 'decimal'), ('numStart', '1'), ('numRestart', 'continuous')):
        put(ensure(pr, name), 'val', value)
    return pr


def apply_layout(infos, parts, body_size, right_margin):
    if any(name.startswith('_xmlsignatures/') for name in parts):
        raise ValueError('Signed packages are not modified; use an authorized unsigned copy')
    doc = parse(parts['word/document.xml'])
    doc.documentElement.setAttribute('xmlns:w', W)
    doc.documentElement.setAttribute('xmlns:r', R)
    body = elements(doc, 'body')[0]
    sections = elements(doc, 'sectPr')
    if len(sections) != 1 or sections[0].parentNode is not body or children(body)[-1] is not sections[0]:
        raise ValueError('Only a single final body sectPr is supported; preserve/edit multisection documents manually')
    if len(children(body)) == 1:
        raise ValueError('The document has no body content')
    section = sections[0]
    already_managed = any(attr(n, 'name') == MARKER for n in elements(body, 'bookmarkStart'))
    blank_marker(body)
    for ref_name in ('headerReference', 'footerReference'):
        for old in children(section, ref_name):
            drop(old)
    size = ensure(section, 'pgSz')
    for key, value in {'w': twips(210), 'h': twips(297), 'orient': 'portrait'}.items():
        put(size, key, value)
    margins = ensure(section, 'pgMar')
    for key, value in {'left': twips(30), 'right': twips(right_margin), 'top': twips(20),
                       'bottom': twips(20), 'header': twips(10), 'footer': twips(10), 'gutter': '0'}.items():
        put(margins, key, value)
    for old in children(section, 'type'):
        drop(old)
    put(ensure(section, 'titlePg'), 'val', '1')
    page = ensure(section, 'pgNumType')
    put(page, 'fmt', 'decimal'); put(page, 'start', '1')
    footnote_properties(section)
    # Referenced headers/footers are replaced; orphaned original parts remain intact.
    types = parse(parts['[Content_Types].xml'])
    relname = 'word/_rels/document.xml.rels'
    if relname not in parts:
        raise ValueError('Missing document relationships part')
    rels = parse(parts[relname])
    xmlstart = '<?xml version="1.0" encoding="UTF-8"?>'
    header = (xmlstart + '<w:hdr xmlns:w="' + W + '"><w:p/></w:hdr>').encode()
    empty_footer = (xmlstart + '<w:ftr xmlns:w="' + W + '"><w:p/></w:ftr>').encode()
    page_footer = (xmlstart + '<w:ftr xmlns:w="' + W + '"><w:p><w:pPr><w:pStyle w:val="Footer"/><w:ind w:left="0" w:right="' + twips(30 - right_margin) + '" w:firstLine="0"/><w:jc w:val="center"/></w:pPr><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>').encode()
    # Reuse previously created dc parts on a second application, avoiding package growth.
    def owned_part(stem, data, kind):
        candidates = [n for n in parts if re.fullmatch('word/dc' + stem + r'\d+\.xml', n)]
        if already_managed and len(candidates) == 1:
            name = candidates[0]
            parts[name] = data
            targets = [n for n in rels.getElementsByTagNameNS(REL, 'Relationship')
                       if n.getAttribute('Target') == name[5:] and n.getAttribute('Type') == R + '/' + kind]
            if len(targets) == 1:
                return targets[0].getAttribute('Id')
        name = fresh_part_name(parts, stem)
        return add_part(parts, infos, types, rels, name, data,
                        'application/vnd.openxmlformats-officedocument.wordprocessingml.' + kind + '+xml', kind)
    hid = owned_part('Header', header, 'header')
    fid = owned_part('Footer', page_footer, 'footer')
    firstid = owned_part('FirstFooter', empty_footer, 'footer')
    for tag, typ, ident in reversed((('headerReference', 'default', hid), ('headerReference', 'first', hid),
                            ('footerReference', 'default', fid), ('footerReference', 'first', firstid))):
        node = doc.createElementNS(W, 'w:' + tag)
        put(node, 'type', typ); node.setAttributeNS(R, 'r:id', ident)
        section.insertBefore(node, section.firstChild)
    # Settings-level defaults complement the section-level footnote properties.
    settingsname = 'word/settings.xml'
    if settingsname not in parts:
        data = (xmlstart + '<w:settings xmlns:w="' + W + '"/>').encode()
        add_part(parts, infos, types, rels, settingsname, data,
                 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml', 'settings')
    settings = parse(parts[settingsname]); settings.documentElement.setAttribute('xmlns:w', W)
    for old in elements(settings, 'evenAndOddHeaders'):
        drop(old)
    fnpr = footnote_properties(settings.documentElement)
    successors = {'compat', 'docVars', 'rsids', 'attachedSchema', 'themeFontLang', 'clrSchemeMapping'}
    before = next((n for n in children(settings.documentElement) if n.localName in successors), None)
    if before:
        settings.documentElement.insertBefore(fnpr, before)
    flag(settings.documentElement, 'updateFields', True)
    parts[settingsname] = xml_bytes(settings)
    if 'word/styles.xml' not in parts:
        raise ValueError('Missing styles.xml; author the document with ordinary Word styles first')
    stylesdoc = parse(parts['word/styles.xml'])
    styles = normalize_styles(stylesdoc, body_size)
    parts['word/styles.xml'] = xml_bytes(stylesdoc)
    normalize_story(doc, body_size, styles)
    parts['word/document.xml'] = xml_bytes(doc)
    for name, data in list(parts.items()):
        if name != 'word/document.xml' and (is_story(name) or name == 'word/footnotes.xml'):
            story = parse(data)
            normalize_story(story, body_size, styles, footnotes=name == 'word/footnotes.xml', footer=bool(re.match(r'word/(footer|dcFooter|dcFirstFooter)', name)))
            parts[name] = xml_bytes(story)
    parts['[Content_Types].xml'] = xml_bytes(types)
    parts[relname] = xml_bytes(rels)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input', type=Path)
    parser.add_argument('output', type=Path)
    parser.add_argument('--body-size', type=int, choices=(12, 14), default=14)
    parser.add_argument('--right-margin-mm', type=float, default=10)
    args = parser.parse_args()
    try:
        if args.input.resolve() == args.output.resolve() or args.output.exists():
            raise ValueError('Choose a different output path that does not already exist')
        if not 1 <= args.right_margin_mm <= 50:
            raise ValueError('--right-margin-mm must be between 1 and 50')
        infos, parts, comment = read_package(args.input)
        apply_layout(infos, parts, args.body_size, args.right_margin_mm)
        write_package(args.output, infos, parts, comment)
        print('Written:', args.output)
        print('Run docx_layout_check.py, then render every page. XML cannot prove actual pagination.')
        return 0
    except Exception as error:
        print('ERROR:', error, file=sys.stderr)
        return 2


if __name__ == '__main__':
    sys.exit(main())
