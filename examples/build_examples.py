#!/usr/bin/env python3
"""Build realistic before/after DOCX examples with local doc-format tools.

Requires python-docx, matplotlib, pdfplumber and Pillow. No network requests.
The source data are synthetic; bibliography URLs are recorded in source-data.json.
"""
import argparse
from datetime import datetime, timezone
from io import BytesIO
import hashlib
import json
from pathlib import Path
import re
import shutil
import subprocess
import sys
import zipfile

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.patches import FancyArrowPatch, Rectangle
import pdfplumber
from PIL import Image
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from lxml import etree as ET

HERE = Path(__file__).resolve().parent
PACKAGE = HERE.parent
sys.path.insert(0,str(PACKAGE/'scripts'))
from docx_xml import ORDERS
PROPERTY_ORDER = dict(ORDERS)
PROPERTY_ORDER['lvl'] = 'start numFmt lvlRestart pStyle isLgl suff lvlText lvlPicBulletId legacy lvlJc pPr rPr'
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
NS = {'w': W, 'm': M}
STAMP = datetime(2000, 1, 1, tzinfo=timezone.utc)


def call(script, *args):
    result = subprocess.run([sys.executable, str(PACKAGE/'scripts'/script), *map(str,args)],
                            capture_output=True, text=True, timeout=180)
    if result.returncode:
        raise RuntimeError(result.stderr or result.stdout)
    return result.stdout


def load_parts(path):
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def save_parts(path, parts):
    temporary = path.with_suffix('.tmp.docx')
    with zipfile.ZipFile(temporary, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            item = zipfile.ZipInfo(name, (2000, 1, 1, 0, 0, 0))
            item.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(item, data)
    temporary.replace(path)


def xml(data):
    return ET.fromstring(data)


def encoded(root):
    return ET.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)


def prop(parent, name, **attrs):
    node = parent.find('w:'+name, NS)
    if node is None:
        node = ET.Element('{%s}%s'%(W,name))
        order=PROPERTY_ORDER.get(ET.QName(parent).localname,'').split()
        index=order.index(name)if name in order else len(order)
        before=next((child for child in parent if ET.QName(child).localname in order and order.index(ET.QName(child).localname)>index),None)
        if before is not None:parent.insert(parent.index(before),node)
        else:parent.append(node)
    for key,value in attrs.items(): node.set('{%s}%s'%(W,key),str(value))
    return node


def font(run, family='Times New Roman', size=14, bold=False):
    rp=run.find('w:rPr',NS)
    if rp is None: rp=ET.Element('{%s}rPr'%W);run.insert(0,rp)
    rf=prop(rp,'rFonts')
    rf.attrib.clear()
    for key in ('ascii','hAnsi','eastAsia','cs'):rf.set('{%s}%s'%(W,key),family)
    prop(rp,'sz',val=size*2);prop(rp,'szCs',val=size*2)
    prop(rp,'b',val='1' if bold else '0');prop(rp,'bCs',val='1' if bold else '0')
    color=prop(rp,'color');color.attrib.clear();color.set('{%s}val'%W,'000000')


def anonymize(path, title):
    parts=load_parts(path)
    core=xml(parts['docProps/core.xml'])
    for child in list(core):core.remove(child)
    for ns,name,text in [('http://purl.org/dc/elements/1.1/','title',title),
                         ('http://purl.org/dc/elements/1.1/','creator','Document Craft Examples'),
                         ('http://schemas.openxmlformats.org/package/2006/metadata/core-properties','lastModifiedBy','Document Craft Examples')]:
        ET.SubElement(core,'{%s}%s'%(ns,name)).text=text
    parts['docProps/core.xml']=encoded(core)
    if 'docProps/app.xml'in parts:
        app=xml(parts['docProps/app.xml'])
        for child in list(app):
            if ET.QName(child).localname in ('Company','Manager','Template','HyperlinkBase'):child.text=''
        parts['docProps/app.xml']=encoded(app)
    parts.pop('docProps/custom.xml',None)
    save_parts(path,parts)


def figures(work, values):
    plt.rcParams.update({'font.family':'DejaVu Sans','font.size':11,'axes.spines.top':False,'axes.spines.right':False})
    try:
        path=font_manager.findfont(font_manager.FontProperties(family='Times New Roman'),fallback_to_default=False)
        font_manager.fontManager.addfont(path);plt.rcParams['font.family']='Times New Roman'
    except ValueError:pass
    outputs={}
    for key,labels in [('request_flow',['Новая','В работе','Закрыта']),
                       ('equipment_schema',['Карточка\nasset','Перемещение\nmovement'])]:
        fig,ax=plt.subplots(figsize=(7,1.65));ax.set_xlim(0,10);ax.set_ylim(0,2);ax.axis('off')
        xs=[.25,3.7,7.15]if len(labels)==3 else[1.0,6.3]
        for x,label in zip(xs,labels):
            ax.add_patch(Rectangle((x,.55),2.6,.9,facecolor='white',edgecolor='black',linewidth=1))
            ax.text(x+1.3,1,label,ha='center',va='center')
        for left,right in zip(xs,xs[1:]):ax.add_patch(FancyArrowPatch((left+2.6,1),(right,1),arrowstyle='->',mutation_scale=12,color='black',linewidth=1))
        if key=='equipment_schema':ax.text(5,1.28,'1 : N',ha='center')
        if key=='request_flow':
            ax.add_patch(FancyArrowPatch((8.45,.55),(5,.55),connectionstyle='arc3,rad=-.30',arrowstyle='->',mutation_scale=11,color='black',linewidth=1))
            ax.text(6.8,-.12,'повторное открытие',ha='center',fontsize=10)
        path=work/(key+'.png');fig.savefig(path,dpi=180,bbox_inches='tight',facecolor='white',metadata={'Software':'Document Craft Examples'});plt.close(fig);outputs[key]=path
    fig,ax=plt.subplots(figsize=(6.6,2.35));mean=sum(values)/len(values)
    ax.plot(range(1,len(values)+1),values,color='black',marker='o',markersize=4,linewidth=1)
    ax.axhline(mean,color='black',linestyle='--',linewidth=.8,label='Среднее')
    ax.set_xlabel('Номер отсчёта');ax.set_ylabel('Длина, мм');ax.set_xticks(range(1,len(values)+1));ax.grid(axis='y',color='.85',linewidth=.5);ax.legend(frameon=False,loc='upper right')
    path=work/'measurement_plot.png';fig.tight_layout();fig.savefig(path,dpi=180,facecolor='white',metadata={'Software':'Document Craft Examples'});plt.close(fig);outputs['measurement_plot']=path
    return outputs


def mr(parent,text):
    run=ET.SubElement(parent,'{%s}r'%M);ET.SubElement(run,'{%s}t'%M).text=text


def subscript(parent,base,index):
    node=ET.SubElement(parent,'{%s}sSub'%M)
    mr(ET.SubElement(node,'{%s}e'%M),base)
    mr(ET.SubElement(node,'{%s}sub'%M),index)


def mean_symbol(parent,base):
    node=ET.SubElement(parent,'{%s}acc'%M)
    pr=ET.SubElement(node,'{%s}accPr'%M)
    ET.SubElement(pr,'{%s}chr'%M).set('{%s}val'%M,'\u0304')
    mr(ET.SubElement(node,'{%s}e'%M),base)


def fraction(parent):
    node=ET.SubElement(parent,'{%s}f'%M);num=ET.SubElement(node,'{%s}num'%M);den=ET.SubElement(node,'{%s}den'%M)
    return num,den


def equation(p,key):
    math=ET.Element('{%s}oMath'%M,nsmap={'m':M})
    if key in ('mean_time','sample_mean'):
        mean_symbol(math,'T'if key=='mean_time'else'x');mr(math,' = ')
        num,den=fraction(math)
        for i,index in enumerate(('1','2','n')):
            if i:mr(num,' + 'if i==1 else' + ⋯ + ')
            subscript(num,'t'if key=='mean_time'else'x',index)
        mr(den,'n')
    elif key in ('closed_share','verified_share'):
        mr(math,'P = 'if key=='closed_share'else'K = ')
        num,den=fraction(math)
        subscript(num,'N','закр'if key=='closed_share'else'подтв')
        subscript(den,'N','всего'if key=='closed_share'else'учёт');mr(math,' × 100 %')
    else:
        mr(math,'s = ');rad=ET.SubElement(math,'{%s}rad'%M);pr=ET.SubElement(rad,'{%s}radPr'%M);ET.SubElement(pr,'{%s}degHide'%M).set('{%s}val'%M,'1');ET.SubElement(rad,'{%s}deg'%M)
        expr=ET.SubElement(rad,'{%s}e'%M);num,den=fraction(expr)
        for i,index in enumerate(('1','n')):
            if i:mr(num,' + ⋯ + ')
            node=ET.SubElement(num,'{%s}sSup'%M);base=ET.SubElement(node,'{%s}e'%M)
            mr(base,'(');subscript(base,'x',index);mr(base,' − ');mean_symbol(base,'x');mr(base,')')
            mr(ET.SubElement(node,'{%s}sup'%M),'2')
        mr(den,'n − 1')
    p._p.append(math)


def add_bookmark(p,name,number):
    start=OxmlElement('w:bookmarkStart');start.set(qn('w:id'),str(number));start.set(qn('w:name'),name)
    end=OxmlElement('w:bookmarkEnd');end.set(qn('w:id'),str(number))
    p._p.insert(1 if p._p.find(qn('w:pPr'))is not None else 0,start);p._p.append(end)


def toc_container(doc):
    doc.add_paragraph('Содержание',style='TOC Heading')
    node=OxmlElement('w:sdt');pr=OxmlElement('w:sdtPr');tag=OxmlElement('w:tag');tag.set(qn('w:val'),'dc-toc');pr.append(tag);node.append(pr)
    content=OxmlElement('w:sdtContent');node.append(content)
    body=doc._element.body;body.insert(len(body)-1,node)


def add_notes(path,notes):
    parts=load_parts(path);root=ET.Element('{%s}footnotes'%W,nsmap={'w':W})
    for ident,typ in [('-1','separator'),('0','continuationSeparator')]:
        note=ET.SubElement(root,'{%s}footnote'%W);note.set('{%s}id'%W,ident);note.set('{%s}type'%W,typ)
        p=ET.SubElement(note,'{%s}p'%W);r=ET.SubElement(p,'{%s}r'%W);ET.SubElement(r,'{%s}%s'%(W,typ))
    for ident,text in enumerate(notes,1):
        note=ET.SubElement(root,'{%s}footnote'%W);note.set('{%s}id'%W,str(ident));p=ET.SubElement(note,'{%s}p'%W);pp=ET.SubElement(p,'{%s}pPr'%W);prop(pp,'pStyle',val='FootnoteText')
        r=ET.SubElement(p,'{%s}r'%W);font(r,size=10);prop(r.find('w:rPr',NS),'vertAlign',val='superscript');ET.SubElement(r,'{%s}footnoteRef'%W)
        r=ET.SubElement(p,'{%s}r'%W);font(r,size=10);t=ET.SubElement(r,'{%s}t'%W);t.set('{http://www.w3.org/XML/1998/namespace}space','preserve');t.text=' '+text
    parts['word/footnotes.xml']=encoded(root)
    rels=xml(parts['word/_rels/document.xml.rels']);r=ET.SubElement(rels,'{%s}Relationship'%REL);r.set('Id','rIdExampleNotes');r.set('Type','http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes');r.set('Target','footnotes.xml');parts['word/_rels/document.xml.rels']=encoded(rels)
    types=xml(parts['[Content_Types].xml']);over=ET.SubElement(types,'{%s}Override'%CT);over.set('PartName','/word/footnotes.xml');over.set('ContentType','application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml');parts['[Content_Types].xml']=encoded(types)
    save_parts(path,parts)


def draft(item,path,images,access_date):
    doc=Document();normal=doc.styles['Normal'];normal.font.name='Calibri'if item['draft']=='office'else'Times New Roman';normal.font.size=Pt(11 if item['draft']=='office'else 12 if item['draft']=='mixed'else 14)
    normal.paragraph_format.line_spacing=1.15 if item['draft']!='mixed'else 1.5;normal.paragraph_format.space_after=Pt(8);normal.paragraph_format.first_line_indent=Cm(0)
    for name in ('CodeListing','TOCText'):
        if name not in doc.styles:doc.styles.add_style(name,1)
    doc.styles['CodeListing'].font.name='Courier New';doc.styles['CodeListing'].font.size=Pt(11);doc.styles['CodeListing'].paragraph_format.line_spacing=1;doc.styles['CodeListing'].paragraph_format.space_after=Pt(0)
    doc.styles['TOCText'].font.size=Pt(12)
    prop(doc.styles['TOC Heading'].element.get_or_add_pPr(),'outlineLvl',val=9)
    for style in ('Heading 1','Heading 2'):
        doc.styles[style].font.name='Times New Roman';doc.styles[style].font.size=Pt(14)
        if item['draft']!='office':doc.styles[style].font.color.rgb=RGBColor(0,0,0)
    section=doc.sections[0];section.page_width=Cm(21);section.page_height=Cm(29.7);section.left_margin=section.right_margin=Cm(2.5);section.top_margin=section.bottom_margin=Cm(2)
    footer=section.footer.paragraphs[0];footer.alignment=WD_ALIGN_PARAGRAPH.CENTER
    field=OxmlElement('w:fldSimple');field.set(qn('w:instr'),' PAGE ');run=OxmlElement('w:r');t=OxmlElement('w:t');t.text='1';run.append(t);field.append(run);footer._p.append(field)
    doc.core_properties.author='Document Craft Examples';doc.core_properties.last_modified_by='Document Craft Examples';doc.core_properties.title=item['title'];doc.core_properties.created=STAMP;doc.core_properties.modified=STAMP
    doc.add_paragraph(item['title'],style='Title');doc.add_paragraph(item['subtitle'],style='Subtitle');toc_container(doc)
    notes=[]
    for index,section_data in enumerate(item['sections'],1):
        heading=doc.add_paragraph(section_data['title'],style='Heading 1');heading.paragraph_format.page_break_before=True;add_bookmark(heading,'dc_h'+str(index),index)
        for block in section_data['blocks']:
            kind=block['kind']
            if kind=='p':doc.add_paragraph(block['text'])
            elif kind=='footnote':
                p=doc.add_paragraph(block['text']);r=p.add_run()._r;rp=r.get_or_add_rPr();prop(rp,'vertAlign',val='superscript');ref=OxmlElement('w:footnoteReference');ref.set(qn('w:id'),str(len(notes)+1));r.append(ref);p.add_run('.');notes.append(block['note'])
            elif kind=='table':
                doc.add_paragraph(block['caption'],style='Caption');table=doc.add_table(rows=1,cols=len(block['columns']));table.style='Light Shading Accent 1'if item['draft']=='office'else'Table Grid'
                for cell,value in zip(table.rows[0].cells,block['columns']):cell.text=value
                marker=OxmlElement('w:tblHeader');table.rows[0]._tr.get_or_add_trPr().append(marker)
                for values in block['rows']:
                    for cell,value in zip(table.add_row().cells,values):cell.text=value
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            for r in p.runs:r.font.size=Pt(11 if item['draft']=='mixed'else 12)
            elif kind=='formula':
                p=doc.add_paragraph();p.paragraph_format.space_before=p.paragraph_format.space_after=Pt(12);p.paragraph_format.tab_stops.add_tab_stop(Cm(8),WD_TAB_ALIGNMENT.CENTER);p.paragraph_format.tab_stops.add_tab_stop(Cm(16),WD_TAB_ALIGNMENT.RIGHT);p.add_run('\t');equation(p,block['key']);p.add_run('\t('+str(block['number'])+')');doc.add_paragraph(block['explanation'])
            elif kind=='code':
                doc.add_paragraph(block['caption'],style='Caption')
                for line in block['text'].split('\n'):doc.add_paragraph(line,style='CodeListing')
            elif kind=='figure':
                p=doc.add_paragraph();p.add_run().add_picture(str(images[block['key']]),width=Cm(14.5));doc.add_paragraph(block['caption'],style='Caption')
    p=doc.add_paragraph('Список источников',style='Heading 2');add_bookmark(p,'dc_sources',100)
    for n,source in enumerate(item['sources'],1):
        doc.add_paragraph(f"{n}. {source['title']}. URL: {source['url']} (дата обращения: {access_date}).")
    doc.save(path);add_notes(path,notes);anonymize(path,item['title'])


def academic_after(path,item):
    parts=load_parts(path);doc=xml(parts['word/document.xml']);styles=xml(parts['word/styles.xml']);numbering=xml(parts['word/numbering.xml'])
    aid=max([int(n.get('{%s}abstractNumId'%W))for n in numbering.findall('w:abstractNum',NS)]+[-1])+1
    nid=max([int(n.get('{%s}numId'%W))for n in numbering.findall('w:num',NS)]+[0])+1
    abstract=ET.SubElement(numbering,'{%s}abstractNum'%W);abstract.set('{%s}abstractNumId'%W,str(aid));prop(abstract,'multiLevelType',val='singleLevel');lvl=ET.SubElement(abstract,'{%s}lvl'%W);lvl.set('{%s}ilvl'%W,'0');prop(lvl,'start',val=1);prop(lvl,'numFmt',val='decimal');prop(lvl,'pStyle',val='Heading1');prop(lvl,'lvlText',val='%1');prop(lvl,'suff',val='space');prop(lvl,'lvlJc',val='left')
    num=ET.SubElement(numbering,'{%s}num'%W);num.set('{%s}numId'%W,str(nid));prop(num,'abstractNumId',val=aid)
    for para in doc.findall('.//w:p',NS):
        pp=para.find('w:pPr',NS)
        if pp is None:pp=ET.Element('{%s}pPr'%W);para.insert(0,pp)
        ref=pp.find('w:pStyle',NS);style=ref.get('{%s}val'%W)if ref is not None else''
        text=''.join(para.itertext())
        in_table=any(ET.QName(n).localname=='tc'for n in para.iterancestors())
        prop(pp,'widowControl',val=1)
        if style=='CodeListing':
            prop(pp,'ind',left=0,right=0,firstLine=0);prop(pp,'spacing',before=0,after=0,line=240,lineRule='auto');prop(pp,'keepNext',val=1)
            for run in para.findall('.//w:r',NS):font(run,'Courier New',12)
        elif style in('Heading1','Heading2'):
            prop(pp,'keepNext',val=1);prop(pp,'spacing',before=0,after=200,line=360,lineRule='auto')
            if style=='Heading1'and re.match(r'^\d+ ',text):
                first=para.find('w:r/w:t',NS)
                if first is not None:first.text=re.sub(r'^\d+\s+','',first.text)
                np=prop(pp,'numPr');prop(np,'ilvl',val=0);prop(np,'numId',val=nid)
        elif para.find('.//m:oMath',NS)is not None:
            prop(pp,'spacing',before=420,after=420,line=360,lineRule='auto');prop(pp,'ind',firstLine=0)
            tabs=pp.find('w:tabs',NS)
            if tabs is not None:pp.remove(tabs)
            tabs=ET.SubElement(pp,'{%s}tabs'%W)
            for position,kind in [('4819','center'),('9638','right')]:tab=ET.SubElement(tabs,'{%s}tab'%W);tab.set('{%s}pos'%W,position);tab.set('{%s}val'%W,kind)
        elif para.find('.//w:drawing',NS)is not None:
            prop(pp,'jc',val='center');prop(pp,'spacing',before=80,after=80,line=360,lineRule='auto')
        elif style=='Caption':
            prop(pp,'keepNext',val=1 if text.startswith(('Таблица','Листинг'))else 0);prop(pp,'spacing',before=100,after=100,line=360,lineRule='auto')
        elif not in_table and text.strip() and style not in('Title','Subtitle','TOCHeading','TOCText'):
            prop(pp,'jc',val='both');prop(pp,'spacing',before=0,after=0,line=360,lineRule='auto')
        elif in_table:prop(pp,'spacing',before=0,after=0,line=360,lineRule='auto')
    tables=doc.findall('.//w:tbl',NS)
    specifications=[b for s in item['sections']for b in s['blocks']if b['kind']=='table']
    for table,spec in zip(tables,specifications):
        pr=table.find('w:tblPr',NS);prop(pr,'tblW',w=9638,type='dxa');prop(pr,'tblLayout',type='fixed')
        grid=table.find('w:tblGrid',NS)
        for child in list(grid):grid.remove(child)
        widths=[round(v*1440/2.54)for v in spec['widths']]
        for width in widths:ET.SubElement(grid,'{%s}gridCol'%W).set('{%s}w'%W,str(width))
        for row in table.findall('w:tr',NS):
            for cell,width in zip(row.findall('w:tc',NS),widths):prop(cell.find('w:tcPr',NS),'tcW',w=width,type='dxa')
    code_style=styles.find('w:style[@w:styleId="CodeListing"]',NS)
    if code_style is not None:
        rp=code_style.find('w:rPr',NS)
        if rp is not None:code_style.remove(rp)
        rp=ET.SubElement(code_style,'{%s}rPr'%W);rf=prop(rp,'rFonts')
        for key in('ascii','hAnsi','eastAsia','cs'):rf.set('{%s}%s'%(W,key),'Courier New')
        prop(rp,'sz',val=24);prop(rp,'szCs',val=24);prop(rp,'b',val=0);prop(rp,'color',val='000000')
        pp=prop(code_style,'pPr');prop(pp,'spacing',before=0,after=0,line=240,lineRule='auto');prop(pp,'ind',firstLine=0)
    parts['word/document.xml']=encoded(doc);parts['word/styles.xml']=encoded(styles);parts['word/numbering.xml']=encoded(numbering);save_parts(path,parts);anonymize(path,item['title'])


def toc_rows(item):
    return [(s['title'],'dc_h'+str(i),1)for i,s in enumerate(item['sections'],1)]+[('Список источников','dc_sources',2)]


def write_toc(path,item,pages,after):
    parts=load_parts(path);doc=xml(parts['word/document.xml']);targets=doc.xpath('.//w:sdt[w:sdtPr/w:tag[@w:val="dc-toc"]]',namespaces=NS);target=targets[0] if targets else None
    if target is None:raise ValueError('Missing TOC container')
    content=target.find('w:sdtContent',NS)
    for child in list(content):content.remove(child)
    p=ET.SubElement(content,'{%s}p'%W);pp=ET.SubElement(p,'{%s}pPr'%W);prop(pp,'pStyle',val='TOCText');prop(pp,'ind',firstLine=0);prop(pp,'spacing',before=0,after=0,line=360,lineRule='auto')
    tabs=ET.SubElement(pp,'{%s}tabs'%W);tab=ET.SubElement(tabs,'{%s}tab'%W);tab.set('{%s}val'%W,'right');tab.set('{%s}leader'%W,'dot');tab.set('{%s}pos'%W,'9638'if after else'9071')
    run=ET.SubElement(p,'{%s}r'%W);ET.SubElement(run,'{%s}fldChar'%W).set('{%s}fldCharType'%W,'begin')
    run=ET.SubElement(p,'{%s}r'%W);instr=ET.SubElement(run,'{%s}instrText'%W);instr.set('{http://www.w3.org/XML/1998/namespace}space','preserve');instr.text=' TOC \\o "1-2" \\h \\z \\u '
    run=ET.SubElement(p,'{%s}r'%W);ET.SubElement(run,'{%s}fldChar'%W).set('{%s}fldCharType'%W,'separate')
    rows=toc_rows(item)
    for i,(title,anchor,level)in enumerate(rows):
        link=ET.SubElement(p,'{%s}hyperlink'%W);link.set('{%s}anchor'%W,anchor)
        run=ET.SubElement(link,'{%s}r'%W);font(run,size=12);ET.SubElement(run,'{%s}t'%W).text=('    'if level==2 else'')+title
        run=ET.SubElement(link,'{%s}r'%W);font(run,size=12);ET.SubElement(run,'{%s}tab'%W);ET.SubElement(run,'{%s}t'%W).text=str(pages[title])
        if i<len(rows)-1:ET.SubElement(run,'{%s}br'%W)
    run=ET.SubElement(p,'{%s}r'%W);ET.SubElement(run,'{%s}fldChar'%W).set('{%s}fldCharType'%W,'end')
    parts['word/document.xml']=encoded(doc);save_parts(path,parts)


def page_map(pdf,item,after):
    pages={}
    with pdfplumber.open(pdf)as doc:
        text=[' '.join((p.extract_text()or'').split())for p in doc.pages]
        for title,_,_ in toc_rows(item):
            for i,content in enumerate(text[2 if after else 1:],3 if after else 2):
                if title in content:pages[title]=i;break
            if title not in pages:raise ValueError('Heading not found in PDF: '+title)
    return pages,len(text)


def render(path,outdir,font_dirs):
    args=[path,'--out-dir',outdir]
    for directory in font_dirs:args+=['--font-dir',directory]
    return json.loads(call('render_pages.py',*args))


def payload(path):
    parts=load_parts(path);doc=xml(parts['word/document.xml'])
    for node in doc.findall('.//w:sdt',NS):node.getparent().remove(node)
    text=[]
    for para in doc.findall('.//w:p',NS):
        value=''.join(para.findall('.//w:t',NS)[i].text or''for i in range(len(para.findall('.//w:t',NS))))
        pp=para.find('w:pPr',NS);sty=pp.find('w:pStyle',NS)if pp is not None else None
        if sty is not None and sty.get('{%s}val'%W)=='Heading1':value=re.sub(r'^\d+\s+','',value)
        if value:text.append(value)
    return {'text':text,'math':[''.join(n.itertext())for n in doc.findall('.//m:oMath',NS)],
            'notes':[''.join(n.itertext())for n in xml(parts['word/footnotes.xml']).findall('w:footnote',NS)],
            'media':{name:hashlib.sha256(data).hexdigest()for name,data in parts.items()if name.startswith('word/media/')}}


def main():
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('--work-dir',type=Path,required=True);parser.add_argument('--font-dir',type=Path,action='append',default=[]);args=parser.parse_args()
    work=args.work_dir.resolve();work.mkdir(parents=True,exist_ok=True)
    source=json.loads((HERE/'source-data.json').read_text());docs=source['documents'];images=figures(work,next(d['values']for d in docs if'values'in d));summary=[]
    for item in docs:
        folder=HERE/item['slug'];folder.mkdir(exist_ok=True);local=work/item['slug'];local.mkdir(exist_ok=True)
        before=folder/'before.docx';after=folder/'after.docx';draft(item,before,images,source['access_date'])
        initial={title:i+2 for i,(title,_,_)in enumerate(toc_rows(item))};write_toc(before,item,initial,False)
        first=render(before,local/'before-initial',args.font_dir);mapped,_=page_map(first['pdf'],item,False);write_toc(before,item,mapped,False)
        final_before=render(before,local/'before',args.font_dir);check_before,before_count=page_map(final_before['pdf'],item,False);assert check_before==mapped
        mono=local/'monochrome.docx';base=local/'layout.docx'
        call('docx_format.py',before,mono);call('docx_layout.py',mono,base);shutil.copy2(base,after);academic_after(after,item)
        expected={k:v+1 for k,v in mapped.items()};write_toc(after,item,expected,True)
        first=render(after,local/'after-initial',args.font_dir);mapped_after,_=page_map(first['pdf'],item,True);write_toc(after,item,mapped_after,True)
        final_after=render(after,local/'after',args.font_dir);check_after,after_count=page_map(final_after['pdf'],item,True);assert check_after==mapped_after
        anonymize(before,item['title']);anonymize(after,item['title'])
        same=payload(before)==payload(after)
        if not same:
            (local/'payload-before.json').write_text(json.dumps(payload(before),ensure_ascii=False,indent=2));(local/'payload-after.json').write_text(json.dumps(payload(after),ensure_ascii=False,indent=2));raise ValueError('Content changed: '+item['slug'])
        before_page=check_before[item['sections'][0]['title']];after_page=check_after[item['sections'][0]['title']]
        for mode,manifest,number in [('before',final_before,before_page),('after',final_after,after_page)]:
            image=Image.open(Path(manifest['pdf']).parent/f'page-{number:04d}.png').convert('RGB');image.save(folder/(mode+'.png'))
        with pdfplumber.open(final_after['pdf'])as pdf:
            assert not(pdf.pages[0].extract_text()or'').strip()
            for index,page in enumerate(pdf.pages[1:],2):
                number=[c for c in page.chars if c['top']>page.height-45]
                assert ''.join(c['text']for c in number).strip()==str(index),(item['slug'],index)
                center=(min(c['x0']for c in number)+max(c['x1']for c in number))/2
                assert abs(center-page.width/2)<.3
        summary.append({'slug':item['slug'],'title':item['title'],'before_pages':before_count,'after_pages':after_count,'before_preview_page':before_page,'after_preview_page':after_page,'table_count':sum(b['kind']=='table'for s in item['sections']for b in s['blocks']),'formula_count':sum(b['kind']=='formula'for s in item['sections']for b in s['blocks']),'code_listings':sum(b['kind']=='code'for s in item['sections']for b in s['blocks']),'figures':sum(b['kind']=='figure'for s in item['sections']for b in s['blocks']),'content_preserved':same,'toc_pages_verified':True,'page_numbers_verified':True,'docx_metadata':'generic','notes':'General DOCX checker exceptions for Courier New listing and generated TOC are checked separately against the academic profile.'})
        print(item['slug'],before_count,'->',after_count,flush=True)
    (HERE/'validation.json').write_text(json.dumps({'synthetic_examples':True,'format_only':True,'documents':summary},ensure_ascii=False,indent=2)+'\n')


if __name__=='__main__':main()
