#!/usr/bin/env python3
"""Read-only audit of the doc-format page/type profile. Render is mandatory."""
import argparse
import json
from pathlib import Path, PurePosixPath
import sys
from docx_xml import W, elements, children, attr, parse, read_package
from docx_layout import (R, REL, FONT, MARKER, twips, style_kind, paragraph_kind,
                         in_header_row, inside_table, first_line_target)

OFF = {'0', 'false', 'off'}


def props(node):
    values = {}
    for rp in children(node, 'rPr'):
        for n in children(rp):
            if n.localName == 'rFonts':
                for key in ('ascii', 'hAnsi', 'eastAsia', 'cs', 'asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme'):
                    if n.hasAttributeNS(W, key): values['font_' + key] = attr(n, key)
            elif n.localName in ('sz', 'szCs', 'vertAlign'):
                values[n.localName] = attr(n, 'val')
            elif n.localName in ('b', 'bCs', 'i', 'iCs'):
                values[n.localName] = attr(n, 'val', '1') not in OFF
    return values


def audit(parts, body_size=14, right_margin=10):
    findings = []
    def add(level, code, part, message):
        finding = {'level': level, 'code': code, 'part': part, 'message': message}
        if finding not in findings: findings.append(finding)
    try:
        doc = parse(parts['word/document.xml'])
        stylesdoc = parse(parts['word/styles.xml'])
        rels = parse(parts['word/_rels/document.xml.rels'])
    except Exception as error:
        return {'structural_errors': 1, 'findings': [{'level':'error','code':'package','part':'(package)','message':str(error)}]}
    styles = {attr(n, 'styleId'): n for n in elements(stylesdoc, 'style')}
    defaults = elements(stylesdoc, 'rPrDefault')
    default_props = props(defaults[0]) if defaults else {}
    cache = {}
    def inherited(style_id, seen=None):
        if style_id in cache: return dict(cache[style_id])
        seen = set() if seen is None else set(seen)
        if style_id in seen: return {}
        seen.add(style_id)
        value = dict(default_props)
        node = styles.get(style_id)
        if node is not None:
            base = children(node, 'basedOn')
            if base: value.update(inherited(attr(base[0], 'val'), seen))
            value.update(props(node))
        cache[style_id] = value
        return dict(value)
    def fonts_check(values, size, allow_bold, part, context):
        if any(values.get('font_' + k) != FONT for k in ('ascii', 'hAnsi', 'eastAsia', 'cs')):
            add('error','font',part,context + ': expected Times New Roman in all script font slots.')
        if any(values.get('font_' + k) for k in ('asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme')):
            add('error','theme_font',part,context + ': theme-dependent font can override the requested typeface.')
        if values.get('sz') != str(size * 2) or values.get('szCs') != str(size * 2):
            add('error','font_size',part,context + ': expected ' + str(size) + ' pt, including complex-script size.')
        if not allow_bold and (values.get('b', False) or values.get('bCs', False)):
            add('error','bold',part,context + ': bold is reserved for headings 1–9 and marked table header rows.')
    # Check definitions as well as effective direct formatting; unused colored/bold title styles should not leak into future edits.
    fonts_check(default_props, body_size, False, 'word/styles.xml', 'docDefaults')
    for style_id,node in styles.items():
        kind = style_kind(style_id, styles)
        size = 10 if kind == 'footnote' else 12 if kind in ('caption', 'footer') or attr(node, 'type') == 'table' else body_size
        fonts_check(inherited(style_id), size, kind == 'heading', 'word/styles.xml', 'style ' + style_id)
    bodies = elements(doc, 'body')
    body = bodies[0] if bodies else None
    sections = elements(doc, 'sectPr')
    if body is None or len(sections) != 1 or sections[0].parentNode is not body:
        add('error','sections','word/document.xml','Only one final body section is supported by this profile checker.')
        section = None
    else:
        section = sections[0]
        pg = children(section,'pgSz')
        if not pg or any(attr(pg[0],k) != v for k,v in {'w':twips(210),'h':twips(297)}.items()) or attr(pg[0],'orient','portrait') != 'portrait':
            add('error','page_size','word/document.xml','Expected portrait A4 (210 × 297 mm).')
        margins=children(section,'pgMar')
        expected={'left':twips(30),'right':twips(right_margin),'top':twips(20),'bottom':twips(20),'gutter':'0'}
        if not margins or any(attr(margins[0],k,'0' if k=='gutter' else '')!=v for k,v in expected.items()):
            add('error','margins','word/document.xml','Expected margins 30/' + str(right_margin) + '/20/20 mm and no gutter.')
        title=children(section,'titlePg')
        if not title or attr(title[0],'val','1') in OFF:
            add('error','first_page_footer','word/document.xml','Different first-page header/footer must be enabled.')
        num=children(section,'pgNumType')
        if not num or attr(num[0],'start')!='1' or attr(num[0],'fmt','decimal')!='decimal':
            add('error','page_numbering','word/document.xml','Numbering must start at 1 in Arabic numerals; physical page two then has number 2.')
    if body is not None:
        direct=children(body)
        first=direct[0] if direct else None
        marker=elements(first,'bookmarkStart') if first is not None else []
        breaks=elements(first,'br') if first is not None else []
        if (first is None or first.localName!='p' or not any(attr(n,'name')==MARKER for n in marker)
            or len([n for n in breaks if attr(n,'type')=='page'])!=1
            or any(''.join(n.childNodes[i].data for i in range(len(n.childNodes)) if n.childNodes[i].nodeType==n.TEXT_NODE).strip() for n in elements(first,'t'))
            or any(elements(first,t) for t in ('drawing','pict','object','tbl','footnoteReference'))):
            add('error','blank_first_structure','word/document.xml','Expected an empty first paragraph with the managed bookmark and a page break. Render must confirm a physically blank first page.')
    # Resolve only the active header/footer parts, leaving orphaned originals outside the active document.
    relationships={n.getAttribute('Id'):n for n in rels.getElementsByTagNameNS(REL,'Relationship')}
    active={}
    if section is not None:
        for kind in ('header','footer'):
            refs=children(section,kind+'Reference')
            for ref in refs:
                typ=attr(ref,'type')
                ident=ref.getAttributeNS(R,'id')
                rel=relationships.get(ident)
                if rel is None or rel.getAttribute('Type') != R+'/'+kind or rel.getAttribute('TargetMode') == 'External':
                    add('error','reference','word/document.xml','Broken ' + kind + ' reference.'); continue
                path=PurePosixPath('word')/rel.getAttribute('Target')
                # Word normally uses local non-parent targets for these managed parts.
                name=str(path)
                if '..' in path.parts or name not in parts:
                    add('error','reference','word/document.xml','Missing/unsupported header/footer target: '+name); continue
                try: active[(kind,typ)]=(name,parse(parts[name]))
                except Exception as error: add('error','xml',name,str(error))
        for key in (('header','default'),('header','first'),('footer','default'),('footer','first')):
            if key not in active: add('error','header_footer_missing','word/document.xml','Missing explicit '+key[1]+' '+key[0]+'.')
    for key,(name,partdoc) in active.items():
        texts=[''.join(n.childNodes[i].data for i in range(len(n.childNodes)) if n.childNodes[i].nodeType==n.TEXT_NODE) for n in elements(partdoc,'t')]
        fields=elements(partdoc,'fldSimple')
        instr=elements(partdoc,'instrText')
        graphic=any(elements(partdoc,t) for t in ('drawing','pict','object','tbl'))
        if key != ('footer','default'):
            if any(t.strip() for t in texts) or fields or instr or graphic:
                add('error','header_footer_content',name,'Headers and the first-page footer must be empty.')
        else:
            pagefields=[f for f in fields if f.getAttributeNS(W,'instr').strip().upper()=='PAGE']
            if len(pagefields)!=1 or instr or graphic:
                add('error','footer_page_field',name,'Expected exactly one simple PAGE field and no other content.')
            for textnode in elements(partdoc,'t'):
                parent=textnode.parentNode
                while parent is not None and parent not in fields: parent=parent.parentNode
                if parent not in pagefields:
                    add('error','footer_extra_text',name,'Footer contains text outside the PAGE field.')
            paras=elements(partdoc,'p')
            if len(paras)!=1:
                add('error','footer_paragraph',name,'Expected one footer paragraph.')
            for para in paras:
                pp=children(para,'pPr'); pstyle=children(pp[0],'pStyle') if pp else []
                if not pstyle or attr(pstyle[0],'val')!='Footer':
                    add('error','footer_style',name,'PAGE paragraph must use the explicit Footer style at 12 pt, so field refresh preserves its size.')
                jc=children(pp[0],'jc') if pp else []; ind=children(pp[0],'ind') if pp else []
                if not jc or attr(jc[0],'val')!='center' or not ind or any(attr(ind[0],k)!=v for k,v in {'left':'0','right':twips(30-right_margin),'firstLine':'0'}.items()):
                    add('error','footer_position',name,'PAGE must be centered on the physical sheet, compensating for unequal left/right margins.')
    check_names=['word/document.xml']+[n for n,_ in active.values()]
    if 'word/footnotes.xml' in parts: check_names.append('word/footnotes.xml')
    for name in dict.fromkeys(check_names):
        partdoc=doc if name=='word/document.xml' else parse(parts[name])
        footnotes=name=='word/footnotes.xml'
        footer=any(kind=='footer' and target[0]==name for (kind,typ),target in active.items())
        for para in elements(partdoc,'p'):
            kind='footnote' if footnotes else paragraph_kind(para,styles)
            size=10 if kind=='footnote' else 12 if kind in ('caption','footer') or footer or inside_table(para) else body_size
            pp=children(para,'pPr'); spacing=children(pp[0],'spacing') if pp else []
            if not spacing or attr(spacing[0],'line')!=('240' if footnotes or footer else '360') or attr(spacing[0],'lineRule','auto')!='auto':
                add('error','line_spacing',name,'Expected single spacing for footnotes/footers and 1.5 spacing elsewhere.')
            target = first_line_target(para, styles) if name == 'word/document.xml' else '0'
            ind = children(pp[0], 'ind') if pp else []
            if target is not None and (not ind or attr(ind[0], 'firstLine') != target
                    or any(ind[0].hasAttributeNS(W, k) for k in ('firstLineChars', 'hanging', 'hangingChars'))):
                add('error','first_line_indent',name,'Expected a 1.25 cm first-line indent for body prose and zero for headings, titles, tables, captions, formulas, blank paragraphs and page numbers; list indentation is preserved.')
            style_refs=children(pp[0],'pStyle') if pp else []
            base=inherited(attr(style_refs[0],'val')) if style_refs else dict(default_props)
            for run in elements(para,'r'):
                owner=run.parentNode
                while owner is not None and getattr(owner,'localName',None)!='p': owner=owner.parentNode
                if owner is not para: continue
                value=dict(base)
                rpr=children(run,'rPr'); rs=children(rpr[0],'rStyle') if rpr else []
                if rs: value.update(inherited(attr(rs[0],'val')))
                value.update(props(run))
                fonts_check(value,size,(kind=='heading' or in_header_row(run)) and not footer,name,kind+' run')
                if footer:
                    colors=children(rpr[0],'color') if rpr else []
                    if not colors or attr(colors[0],'val')!='000000' or colors[0].hasAttributeNS(W,'themeColor'):
                        add('error','footer_color',name,'PAGE must have explicit black text.')
                if (children(run,'footnoteReference') or children(run,'footnoteRef')) and value.get('vertAlign')!='superscript':
                    add('error','footnote_superscript',name,'Footnote reference marks must be superscript.')
            if kind=='caption':
                jc=children(pp[0],'jc') if pp else []
                if not jc or attr(jc[0],'val')!='center': add('error','caption_alignment',name,'Caption paragraph must be centered.')
        if footnotes:
            notes=elements(partdoc,'footnote')
            if not any(attr(n,'type')=='separator' and elements(n,'separator') for n in notes):
                add('error','footnote_separator',name,'Automatic Word footnote separator is absent.')
    for name in ('word/settings.xml','word/document.xml'):
        if name not in parts: continue
        pd=parse(parts[name])
        for pr in elements(pd,'footnotePr'):
            for prop,val in (('numFmt','decimal'),('numStart','1'),('numRestart','continuous')):
                nodes=children(pr,prop)
                if not nodes or attr(nodes[0],'val')!=val:
                    add('error','footnote_numbering',name,'Footnotes must use continuous Arabic numbering from 1.')
        if elements(pd,'evenAndOddHeaders'):
            add('error','even_odd_headers',name,'Separate even-page headers/footers are unsupported; disable them.')
    add('warning','visual_review','(render)','XML checks cannot prove physical blankness, actual page numbers, footer coordinates, pagination, font substitution, formula numbering, caption numbering, or punctuation around footnote references. Inspect rendered pages and text.')
    return {'body_size_pt':body_size,'right_margin_mm':right_margin,'structural_errors':sum(f['level']=='error' for f in findings),'findings':findings}


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input',type=Path)
    parser.add_argument('--body-size',type=int,choices=(12,14),default=14)
    parser.add_argument('--right-margin-mm',type=float,default=10)
    parser.add_argument('--json',action='store_true')
    args=parser.parse_args()
    try:
        result=audit(read_package(args.input)[1],args.body_size,args.right_margin_mm)
        if args.json: print(json.dumps(result,ensure_ascii=False,indent=2))
        else:
            print('Structural errors:',result['structural_errors'])
            for f in result['findings']: print('{level}: {code}: {part}: {message}'.format(**f))
        return 1 if result['structural_errors'] else 0
    except Exception as error:
        print(json.dumps({'error':str(error)}) if args.json else 'ERROR: '+str(error))
        return 2


if __name__=='__main__': sys.exit(main())
