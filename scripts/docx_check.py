#!/usr/bin/env python3
"""Read-only structural DOCX audit. Not a layout, factual, or authorship check."""
import argparse
import json
from pathlib import Path
import sys
from docx_xml import (W, A, elements, children, table_rows, attr, is_story, read_package,
                      parse, format_signature)

OFF = {'0', 'false', 'off'}
BLACK = {'000000', 'auto'}
WHITE = {'FFFFFF', 'auto', ''}


def audit(parts, profile='monochrome', scope='document', baseline=None):
    findings = []
    counts = {'tables': 0, 'rows': 0, 'story_parts': 0}
    def add(level, code, part, message):
        item = {'level': level, 'code': code, 'part': part, 'message': message}
        # One finding per kind/part is enough; counts remain available separately.
        if item not in findings:
            findings.append(item)
    for name, data in parts.items():
        if not name.endswith('.xml'):
            continue
        try:
            doc = parse(data)
        except Exception as error:
            add('error', 'xml_invalid', name, str(error))
            continue
        story = is_story(name)
        if story:
            counts['story_parts'] += 1
            for tag in ('drawing', 'pict', 'object', 'txbxContent', 'altChunk'):
                if elements(doc, tag):
                    add('warning', 'visual_object', name, 'Contains ' + tag + '; inspect rendered output. Embedded graphics/colors are not validated.')
            if elements(doc, 'ins') or elements(doc, 'del'):
                add('warning', 'tracked_changes', name, 'Contains tracked changes; inspect the intended review/final view.')
        if profile != 'monochrome':
            if story:
                tables = elements(doc, 'tbl')
                counts['tables'] += len(tables)
                counts['rows'] += sum(len(table_rows(t)) for t in tables)
            continue
        relevant = story or (scope == 'document' and name in ('word/styles.xml', 'word/numbering.xml'))
        if not relevant:
            continue
        roots = [doc] if scope == 'document' else elements(doc, 'tbl')
        for root in roots:
            for n in elements(root, 'color'):
                if attr(n, 'val').upper() not in BLACK or any(n.hasAttributeNS(W, k) for k in ('themeColor', 'themeTint', 'themeShade')):
                    add('error', 'text_color', name, 'Non-black or theme-dependent text color is present.')
            for n in elements(root, 'shd'):
                if (attr(n, 'fill').upper() not in WHITE or attr(n, 'val') not in ('', 'clear', 'nil')
                        or any(n.hasAttributeNS(W, k) for k in ('themeFill', 'themeFillTint', 'themeFillShade', 'themeColor', 'themeTint', 'themeShade'))):
                    add('error', 'shading', name, 'Colored/patterned or theme-dependent shading is present.')
            for n in elements(root, 'highlight'):
                if attr(n, 'val') != 'none':
                    add('error', 'highlight', name, 'Text highlighting is present.')
            for n in root.getElementsByTagName('*'):
                if n.localName not in ('shd', 'background') and (n.hasAttributeNS(W, 'color') or n.hasAttributeNS(W, 'themeColor')):
                    if attr(n, 'color').upper() not in BLACK or any(n.hasAttributeNS(W, k) for k in ('themeColor', 'themeTint', 'themeShade')):
                        add('error', 'decoration_color', name, 'A border/underline/other property has non-black or theme-dependent color.')
                if n.localName in ('textFill', 'textOutline', 'glow', 'shadow', 'reflection'):
                    add('warning', 'text_effect', name, 'Text effects require visual review.')
            for n in elements(root, 'background'):
                if attr(n, 'color').upper() not in WHITE or any(n.hasAttributeNS(W, k) for k in ('themeColor', 'themeTint', 'themeShade')) or any(c.nodeType == c.ELEMENT_NODE for c in n.childNodes):
                    add('error', 'page_background', name, 'A colored, themed, or graphical page background is present.')
            if elements(root, 'pBdr') or elements(root, 'bdr'):
                add('warning', 'paragraph_border', name, 'Paragraph/text borders may be decorative; inspect their purpose.')
        if story:
            for table in elements(doc, 'tbl'):
                counts['tables'] += 1
                rows = table_rows(table)
                counts['rows'] += len(rows)
                pr = children(table, 'tblPr')
                pr = pr[0] if pr else None
                if pr is not None and children(pr, 'tblStyle'):
                    add('warning', 'table_style', name, 'A table style is referenced. Its conditional/default inheritance needs review; definitions are audited in document scope.')
                if pr is None or not children(pr, 'tblBorders'):
                    add('error', 'table_borders', name, 'A table has no explicit borders; inherited appearance is not established.')
                else:
                    border_root = children(pr, 'tblBorders')[0]
                    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
                        nodes = children(border_root, edge)
                        if not nodes or attr(nodes[0], 'val') != 'single' or attr(nodes[0], 'sz') != '4' or attr(nodes[0], 'color').upper() != '000000':
                            add('error', 'table_borders', name, 'Expected explicit 0.5 pt single black outer/internal table borders.')
                has_header = False
                for row in rows:
                    rpr = children(row, 'trPr')
                    rpr = rpr[0] if rpr else None
                    cant = children(rpr, 'cantSplit') if rpr else []
                    if not cant or attr(cant[0], 'val', '1') in OFF:
                        add('warning', 'row_split', name, 'A row can split across pages; verify long-row behavior visually.')
                    headers = children(rpr, 'tblHeader') if rpr else []
                    has_header = has_header or bool(headers and attr(headers[0], 'val', '1') not in OFF)
                    if rpr:
                        for h in children(rpr, 'trHeight'):
                            if attr(h, 'hRule') == 'exact':
                                add('warning', 'fixed_row_height', name, 'A fixed row height can clip text.')
                if len(rows) > 1 and not has_header:
                    add('warning', 'header_unknown', name, 'No repeating header is marked. Decide whether the first rows are headers; do not guess automatically.')
                # A cell border can override tblBorders and must be checked too.
                for cell in elements(table, 'tc'):
                    for cb in elements(cell, 'tcBorders'):
                        for border in children(cb):
                            if attr(border, 'val') not in ('nil', 'none') and (attr(border, 'color').upper() != '000000' or attr(border, 'sz') != '4' or attr(border, 'val') != 'single'):
                                add('error', 'cell_borders', name, 'Cell border overrides the 0.5 pt single black border rule.')
    if baseline is not None:
        old = baseline
        changed = sorted(name for name in set(old) | set(parts) if old.get(name) != parts.get(name))
        add('info', 'changed_parts', '(package)', ', '.join(changed) if changed else 'No package parts changed.')
        for name in changed:
            if name in ('word/styles.xml', 'word/numbering.xml') or name.startswith('word/theme/'):
                add('warning', 'shared_format_changed', name, 'Shared styles/numbering/theme changed; check whether this was authorized.')
            elif is_story(name) and name in old and name in parts:
                try:
                    if format_signature(old[name]) != format_signature(parts[name]):
                        add('warning', 'format_changed', name, 'Formatting property sequence differs from baseline. Review the intended edit; this is not a visual comparison.')
                except Exception:
                    pass
    elif profile == 'preserve':
        add('warning', 'baseline_missing', '(package)', 'No baseline supplied: preservation of the original formatting cannot be checked.')
    if any(name.startswith('_xmlsignatures/') for name in parts):
        add('warning', 'signed_package', '(package)', 'Digital signature parts exist; changing the document can invalidate signatures.')
    return {'profile': profile, 'scope': scope, 'counts': counts, 'findings': findings,
            'structural_errors': sum(f['level'] == 'error' for f in findings),
            'limitations': 'Structural audit only. Render every page and review content, typography, inherited appearance, images, clipping and pagination separately.'}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input', type=Path)
    parser.add_argument('--profile', choices=('monochrome', 'preserve'), default='monochrome')
    parser.add_argument('--scope', choices=('document', 'tables'), default='document')
    parser.add_argument('--baseline', type=Path)
    parser.add_argument('--json', action='store_true')
    args = parser.parse_args()
    try:
        _, parts, _ = read_package(args.input)
        old = read_package(args.baseline)[1] if args.baseline else None
        result = audit(parts, args.profile, args.scope, old)
        if args.json:
            print(json.dumps(result, ensure_ascii=False, indent=2))
        else:
            print('Structural errors:', result['structural_errors'])
            for item in result['findings']:
                print('{level}: {code}: {part}: {message}'.format(**item))
            print(result['limitations'])
        return 1 if result['structural_errors'] else 0
    except Exception as error:
        if args.json:
            print(json.dumps({'error': str(error)}, ensure_ascii=False))
        else:
            print('ERROR:', error, file=sys.stderr)
        return 2


if __name__ == '__main__':
    sys.exit(main())
