#!/usr/bin/env python3
"""Apply monochrome OOXML formatting to a NEW document or explicit restyle.
Never use globally for a targeted edit. Preserves unmodified package parts.
"""
import argparse
from pathlib import Path
import sys
from docx_xml import (W, elements, children, table_rows, attr, put, drop, ensure, is_story,
                      read_package, parse, xml_bytes, text_signature, write_package)


def black(node):
    for key in list(node.attributes.keys()):
        if node.attributes[key].namespaceURI == W:
            node.removeAttribute(key)
    put(node, 'val', '000000')


def white(node):
    for key in list(node.attributes.keys()):
        if node.attributes[key].namespaceURI == W:
            node.removeAttribute(key)
    put(node, 'val', 'clear')
    put(node, 'color', 'auto')
    put(node, 'fill', 'FFFFFF')


def plain_border(node):
    for key in list(node.attributes.keys()):
        if node.attributes[key].namespaceURI == W:
            node.removeAttribute(key)
    for key, value in {'val': 'single', 'sz': '4', 'space': '0', 'color': '000000'}.items():
        put(node, key, value)


def clean_subtree(root):
    for n in elements(root, 'color'):
        black(n)
    for n in elements(root, 'shd'):
        white(n)
    for n in elements(root):
        if n.localName not in ('shd', 'background') and (n.hasAttributeNS(W, 'color') or n.hasAttributeNS(W, 'themeColor')):
            for key in ('themeColor', 'themeTint', 'themeShade'):
                if n.hasAttributeNS(W, key):
                    n.removeAttributeNS(W, key)
            put(n, 'color', '000000')
    for name in ('highlight', 'textFill', 'textOutline', 'glow', 'shadow', 'reflection'):
        for n in list(root.getElementsByTagName('*')):
            if n.localName == name and n.namespaceURI in (W, 'http://schemas.microsoft.com/office/word/2010/wordml'):
                drop(n)
    # Direct formatting defeats character-style and theme colors on actual text.
    for run in elements(root, 'r'):
        black(ensure(ensure(run, 'rPr', first=True), 'color'))
    # Paragraph marks can carry color/shading even in empty cells.
    for para in elements(root, 'p'):
        ppr = ensure(para, 'pPr', first=True)
        black(ensure(ensure(ppr, 'rPr'), 'color'))
        white(ensure(ppr, 'shd'))
    for border_name in ('pBdr', 'bdr'):
        for n in elements(root, border_name):
            # Paragraph/text borders are decorative; table borders are set below.
            drop(n)


def format_table(table, header_rows):
    pr = ensure(table, 'tblPr', first=True)
    for old in children(pr, 'tblStyle'):
        drop(old)
    white(ensure(pr, 'shd'))
    look = ensure(pr, 'tblLook')
    for key, value in {'val': '0000', 'firstRow': '0', 'lastRow': '0',
                       'firstColumn': '0', 'lastColumn': '0', 'noHBand': '1', 'noVBand': '1'}.items():
        put(look, key, value)
    borders = ensure(pr, 'tblBorders')
    for old in list(children(borders)):
        drop(old)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        plain_border(ensure(borders, edge))
    # Direct cell formatting overrides even a customized default table style.
    for cell in elements(table, 'tc'):
        cp = ensure(cell, 'tcPr', first=True)
        white(ensure(cp, 'shd'))
        cb = ensure(cp, 'tcBorders')
        for old in list(children(cb)):
            drop(old)
        for edge in ('top', 'left', 'bottom', 'right'):
            plain_border(ensure(cb, edge))
    rows = table_rows(table)
    if header_rows is not None and header_rows > len(rows):
        raise ValueError('--header-rows exceeds the row count of a table')
    for index, row in enumerate(rows):
        rp = ensure(row, 'trPr', first=True)
        put(ensure(rp, 'cantSplit'), 'val', '1')
        # Remove fixed row heights that may clip content; preserve atLeast height.
        for h in children(rp, 'trHeight'):
            if attr(h, 'hRule') == 'exact':
                put(h, 'hRule', 'atLeast')
        if header_rows is not None:
            put(ensure(rp, 'tblHeader'), 'val', '1' if index < header_rows else '0')


def transform(parts, scope, header_rows):
    changed = []
    for name, original in list(parts.items()):
        story = is_story(name)
        if not story and not (scope == 'document' and name in ('word/styles.xml', 'word/numbering.xml')):
            continue
        doc = parse(original)
        if scope == 'tables' and not elements(doc, 'tbl'):
            continue
        doc.documentElement.setAttribute('xmlns:w', W)
        if scope == 'document':
            clean_subtree(doc)
            if name == 'word/styles.xml':
                defaults = ensure(doc.documentElement, 'docDefaults', first=True)
                black(ensure(ensure(ensure(defaults, 'rPrDefault'), 'rPr'), 'color'))
            for bg in elements(doc, 'background'):
                for key in list(bg.attributes.keys()):
                    if bg.attributes[key].namespaceURI == W:
                        bg.removeAttribute(key)
                put(bg, 'color', 'FFFFFF')
                for child in list(bg.childNodes):
                    bg.removeChild(child)
        for table in elements(doc, 'tbl'):
            if scope == 'tables':
                clean_subtree(table)
            format_table(table, header_rows)
        # No table bodies live in styles/numbering; normalize their border colors.
        if scope == 'document' and not story:
            for n in elements(doc):
                if n.hasAttributeNS(W, 'color') and n.localName not in ('shd', 'background'):
                    for key in ('themeColor', 'themeTint', 'themeShade'):
                        if n.hasAttributeNS(W, key):
                            n.removeAttributeNS(W, key)
                    put(n, 'color', '000000')
        new = xml_bytes(doc)
        if story and text_signature(original) != text_signature(new):
            raise ValueError('Text/field content changed unexpectedly in ' + name)
        if new != original:
            parts[name] = new
            changed.append(name)
    return changed


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input', type=Path)
    parser.add_argument('output', type=Path)
    parser.add_argument('--scope', choices=('document', 'tables'), default='document')
    parser.add_argument('--header-rows', type=int, help='Explicit first N header rows in EVERY table, including nested tables. Omit to retain existing markers.')
    args = parser.parse_args()
    try:
        if args.input.resolve() == args.output.resolve():
            raise ValueError('Input and output must differ')
        if args.output.exists():
            raise ValueError('Output already exists; choose a new filename')
        if args.header_rows is not None and args.header_rows < 0:
            raise ValueError('--header-rows must be non-negative')
        infos, parts, comment = read_package(args.input)
        if any(name.startswith('_xmlsignatures/') for name in parts):
            raise ValueError('Signed packages are not modified; first create an authorized unsigned copy')
        changed = transform(parts, args.scope, args.header_rows)
        write_package(args.output, infos, parts, comment)
        print('Written:', args.output)
        print('Changed OOXML parts:', ', '.join(changed) or 'none')
        print('Next: run docx_check.py and render/inspect every page. This tool does not verify layout.')
    except Exception as error:
        print('ERROR:', error, file=sys.stderr)
        return 2
    return 0


if __name__ == '__main__':
    sys.exit(main())
