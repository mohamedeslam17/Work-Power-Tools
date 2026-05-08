#!/usr/bin/env python3
"""
SEM Report Converter - Ansaldo Energia
Usage: python3 sem_convert_final.py vendor.pdf [output.docx]
Place the companion _R.pdf in the same folder for best results (heat treatment field).
"""
import sys, os, re, io
from pathlib import Path
import fitz
from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

RED  = RGBColor(0xC8,0x10,0x2E)
NAVY = RGBColor(0x1A,0x1A,0x2E)
GRAY = RGBColor(0x55,0x55,0x55)
WHITE= RGBColor(0xFF,0xFF,0xFF)

# ── helpers ───────────────────────────────────────────────────
def _bg(c,h):
    tc=c._tc;p=tc.get_or_add_tcPr();s=OxmlElement("w:shd")
    s.set(qn("w:val"),"clear");s.set(qn("w:color"),"auto");s.set(qn("w:fill"),h);p.append(s)
def _bdr(c,color="AAAAAA",sz=4):
    tc=c._tc;p=tc.get_or_add_tcPr();b=OxmlElement("w:tcBorders")
    for side in["top","left","bottom","right"]:
        e=OxmlElement(f"w:{side}");e.set(qn("w:val"),"single")
        e.set(qn("w:sz"),str(sz));e.set(qn("w:color"),color);b.append(e)
    p.append(b)
def _nobdr(c):
    tc=c._tc;p=tc.get_or_add_tcPr();b=OxmlElement("w:tcBorders")
    for side in["top","left","bottom","right"]:e=OxmlElement(f"w:{side}");e.set(qn("w:val"),"nil");b.append(e)
    p.append(b)
def R(p,text,bold=False,size=10,color=None,italic=False):
    r=p.add_run(text);r.bold=bold;r.italic=italic
    r.font.size=Pt(size);r.font.name="Calibri"
    if color:r.font.color.rgb=color
    return r
def SP(doc,h=2):
    p=doc.add_paragraph();p.paragraph_format.space_before=Pt(h);p.paragraph_format.space_after=Pt(h)

# ── PDF helpers ───────────────────────────────────────────────
def page_text(page):
    d=page.get_text("dict");spans=[]
    for b in d["blocks"]:
        if b.get("type")!=0:continue
        for line in b.get("lines",[]):
            for s in line.get("spans",[]):
                spans.append((s["bbox"][1],s["bbox"][0],s["text"]))
    spans.sort(key=lambda s:(round(s[0]/3)*3,s[1]))
    return ' '.join(s[2] for s in spans if s[2].strip())

def caption_from_page(page,fig_num):
    d=page.get_text("dict");spans=[]
    for b in d["blocks"]:
        if b.get("type")!=0:continue
        for line in b.get("lines",[]):
            for s in line.get("spans",[]):
                y=s["bbox"][1]
                if 620<y<775 and s["text"].strip():
                    spans.append((y,s["bbox"][0],s["text"]))
    spans.sort(key=lambda s:(round(s[0]/3)*3,s[1]))
    t=re.sub(r'\s+',' ',' '.join(s[2] for s in spans).strip())
    t=re.sub(r'\s*Note\s*:.*','',t);t=re.sub(r'\s*Document No.*','',t)
    if t and not re.match(r'Fig',t,re.I):t=f"Fig 1.{fig_num} shows "+t
    return t.strip()

def is_image_page(page,pdf):
    d=page.get_text("dict")
    for b in d["blocks"]:
        if b.get("type")!=0:continue
        for line in b.get("lines",[]):
            lt=' '.join(s["text"] for s in line.get("spans",[]))
            if re.search(r'As-received Sample|SEM Analysis\s*[\u2013\u2014-]',lt):
                return any(pdf.extract_image(img[0]).get('width',0)>500 for img in page.get_images())
    return False

# ══════════════════════════════════════════════════════════════
# PARSE
# ══════════════════════════════════════════════════════════════
def parse(pdf_path):
    vendor=fitz.open(pdf_path)
    full='\n'.join(page_text(p) for p in vendor)

    job_m=re.search(r'Job No[:\s.]+(\d+)',full,re.I)
    job=job_m.group(1).strip() if job_m else 'N/A'
    sn_m=re.search(r'S/N[:\s]+([A-Z0-9]+)',full,re.I)
    serial=sn_m.group(1).strip() if sn_m else 'N/A'
    mat_m=re.search(r'Material[:\s]+(IN[\w-]+)',full,re.I)
    mat=mat_m.group(1).strip() if mat_m else 'IN738'
    dt_m=re.search(r'Date[:\s]+([A-Za-z]+ \d+[a-z]*,?\s*\d{4})',full,re.I)
    date=dt_m.group(1).strip() if dt_m else '05/05/2026'

    sm=re.search(r'Job No[:\s.]+\d+\s+FR\s+(\d+)\s+(\d+)(?:nd|rd|st|th)\s+STG?\s+(BKT|BUCKET)',full,re.I)
    stage=f"MS{sm.group(1)}001 Stage {sm.group(2)} Bucket" if sm else "MS7001 Stage 3 Bucket"

    ht="Aged"
    ia="Heavy Repair"
    # Try companion _R.pdf for ht and ia
    base=str(pdf_path)
    for r_path in [base.replace('.pdf','_R.pdf'),
                   str(Path(pdf_path).parent/f"{Path(pdf_path).stem}_R.pdf")]:
        if os.path.exists(r_path):
            rpdf=fitz.open(r_path)
            rf='\n'.join(page_text(p) for p in rpdf); rpdf.close()
            ht_m=re.search(r'Heat Treatment Condition[:\s]+([^\n•\u2022]+)',rf,re.I)
            if ht_m:ht=ht_m.group(1).strip()
            ia_m=re.search(r'Incoming Assessment[:\s]+([^\n•\u2022]+)',rf,re.I)
            if ia_m:ia=ia_m.group(1).strip()
            break

    sizes=re.findall(r'measured to be ([\d.]+) microns',full,re.I)
    l1=sizes[0] if sizes else 'N/A'
    l2=sizes[1] if len(sizes)>1 else 'N/A'
    no_anom=bool(re.search(r'No (evidence|indications) of.*(needle|sigma|eta)',full,re.I))
    rts=bool(re.search(r'suitable for return to service',full,re.I))
    # Conclusion from _R pdf (not in vendor pdf)
    conclusion=''
    for r_path in [base.replace('.pdf','_R.pdf'),
                   str(Path(pdf_path).parent/f"{Path(pdf_path).stem}_R.pdf")]:
        if os.path.exists(r_path):
            rpdf2=fitz.open(r_path)
            rlast=rpdf2[-1].get_text("text")
            rpdf2.close()
            cm=re.search(r'CONCLUSION\s*\n+(.*?)(?:Location\s*\nMorphology|$)',rlast,re.DOTALL|re.I)
            if cm:conclusion=cm.group(1).strip()  # keep newlines for paragraph splitting
            break
    if not conclusion:
        cm=re.search(r'(The metallurgical evaluation.+?NDT inspections\.)',full,re.DOTALL|re.I)
        if cm:conclusion=re.sub(r'\s+',' ',cm.group(1)).strip()

    # captions — from image pages only
    captions={}
    for page in vendor:
        if not is_image_page(page,vendor):continue
        pt=page_text(page)
        m=re.search(r'(?:Fig|Figure)\s+1[.\s]*(\d+)\s+shows',pt,re.I)
        if not m:continue
        fn=m.group(1)
        if fn in captions:continue
        captions[fn]=caption_from_page(page,fn)

    # Fig 1.1 caption — from the right-side text box on the As-received page
    for page in vendor:
        if 'As-received' not in page_text(page):continue
        d=page.get_text("dict");spans=[]
        for b in d["blocks"]:
            if b.get("type")!=0:continue
            for line in b.get("lines",[]):
                for s in line.get("spans",[]):
                    y,x,txt=s["bbox"][1],s["bbox"][0],s["text"]
                    if y>480 and x>400 and txt.strip():spans.append((y,x,txt))
        spans.sort(key=lambda s:(round(s[0]/3)*3,s[1]))
        side=re.sub(r'\s+',' ',' '.join(s[2] for s in spans).strip())
        m2=re.search(r'(Fig\s+1\.1\s+shows.+?(?:under SEM\.|SEM\.))',side,re.I|re.DOTALL)
        if m2:captions['1']=re.sub(r'\s+',' ',m2.group(1)).strip()
        break
    if '1' not in captions or not captions.get('1','').startswith('Fig'):
        captions['1']=f"Fig 1.1 shows Image of the as-received sample (ID# {job}). The specimen was first mounted, metallographically prepared and etched with Glyceregia prior to examination under SEM."

    vendor.close()
    return dict(job=job,serial=serial,material=mat,date=date,stage=stage,
                ht=ht,ia=ia,l1=l1,l2=l2,no_anom=no_anom,rts=rts,
                conclusion=conclusion,captions=captions)

# ══════════════════════════════════════════════════════════════
# EXTRACT FIGURES
# ══════════════════════════════════════════════════════════════
def extract_figures(pdf_path):
    doc=fitz.open(pdf_path);figs={}
    for page in doc:
        if not is_image_page(page,doc):continue
        pt=page_text(page)
        m=re.search(r'(?:Fig|Figure)\s+1[.\s]*(\d+)\s+shows',pt,re.I)
        if not m:continue
        fn=m.group(1)
        if fn in figs:continue

        # Fig 1.1 special case — get just the specimen photo (second embedded image)
        if fn=='1':
            imgs=page.get_image_info(xrefs=True)
            # Sort images by y position; take the LAST large one (specimen, not bag)
            large=[img for img in imgs if img.get('width',0)>300 and img.get('height',0)>200
                   and img.get('xref',0)!=19]  # skip logo
            large.sort(key=lambda img:img['bbox'][1])  # sort by top y
            if len(large)>=2:
                # Take the second image (specimen), crop exactly its bbox
                bbox=large[1]['bbox']
                crop=fitz.Rect(bbox[0],bbox[1],bbox[2],bbox[3])
                pix=page.get_pixmap(dpi=180,clip=crop)
                figs[fn]={'bytes':pix.tobytes('jpeg'),'w':pix.width,'h':pix.height}
                continue
            elif large:
                bbox=large[-1]['bbox']
                crop=fitz.Rect(bbox[0],bbox[1],bbox[2],bbox[3])
                pix=page.get_pixmap(dpi=180,clip=crop)
                figs[fn]={'bytes':pix.tobytes('jpeg'),'w':pix.width,'h':pix.height}
                continue

        # All other figures — crop from header to just above caption
        pw=page.rect.width;ph=page.rect.height
        d=page.get_text("dict");hdr_bot=80.0;cap_top=ph-220.0
        for b in d["blocks"]:
            if b.get("type")!=0:continue
            for line in b.get("lines",[]):
                lt=' '.join(s["text"] for s in line.get("spans",[]))
                bb=line["bbox"]
                if re.search(r'SEM Analysis\s*[\u2013\u2014-]|As-received|Location Mapping',lt):
                    if bb[3]>hdr_bot:hdr_bot=bb[3]
                if re.match(r'\s*Fig(?:ure)?\s+1\.\d+\s+shows',lt) and bb[1]>400:
                    if bb[1]<cap_top:cap_top=bb[1]

        crop=fitz.Rect(18,hdr_bot+5,pw-18,cap_top-4)
        pix=page.get_pixmap(dpi=180,clip=crop)
        figs[fn]={'bytes':pix.tobytes('jpeg'),'w':pix.width,'h':pix.height}

    doc.close()
    return figs

# ══════════════════════════════════════════════════════════════
# BUILD DOCX  — matching _R.pdf layout exactly
# ══════════════════════════════════════════════════════════════
def page_header_table(doc, info, page_num, total_pages=9):
    """Full-width header table: ansaldo logo row + info row"""
    # Logo paragraph  
    logo_p = doc.add_paragraph()
    logo_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    logo_p.paragraph_format.space_before = Pt(0)
    logo_p.paragraph_format.space_after = Pt(6)
    r = logo_p.add_run("ansaldo"); r.bold=True; r.font.size=Pt(24); r.font.name="Calibri"; r.font.color.rgb=RGBColor(0x8B,0x8B,0x8B)
    r2 = logo_p.add_run("|"); r2.font.size=Pt(24); r2.font.name="Calibri"; r2.font.color.rgb=RED
    r3 = logo_p.add_run("energia"); r3.font.size=Pt(24); r3.font.name="Calibri"; r3.font.color.rgb=RGBColor(0x8B,0x8B,0x8B)

    # Info table
    tbl = doc.add_table(rows=2, cols=5)
    tbl.style = 'Table Grid'
    ws = [Cm(7.5), Cm(3.8), Cm(1.5), Cm(1.5), Cm(1.2)]

    # Row 0: labels
    for j,(h,w) in enumerate(zip(['Project / Title','Job Number.','Rev.','page','Of'],ws)):
        c=tbl.rows[0].cells[j]; c.width=w; _bdr(c,'888888',2)
        p=c.paragraphs[0]; p.paragraph_format.space_before=Pt(1); p.paragraph_format.space_after=Pt(1)
        R(p,h,size=8,color=GRAY)

    # Row 1: values
    vals=[
        ('SEM Metallurgical Evaluation Report',False,10),
        (f"JC. {info['job']}",True,11),
        ('0',False,10),
        (str(page_num),False,10),
        (str(total_pages),False,10)
    ]
    for j,((val,bold,sz),w) in enumerate(zip(vals,ws)):
        c=tbl.rows[1].cells[j]; c.width=w; _bdr(c,'888888',2)
        p=c.paragraphs[0]; p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
        if j==1:
            R(p,val,bold=True,size=sz)
            p2=c.add_paragraph(); R(p2,info['stage'],size=8,color=GRAY)
        else:
            R(p,val,bold=bold,size=sz)
    return tbl

def add_two_col(doc, left_content_fn, right_bytes, right_w=Cm(7.5), caption=''):
    """Add a two-column table: left=text, right=image+caption"""
    t = doc.add_table(rows=1, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    lc = t.rows[0].cells[0]; lc.width = Cm(8.0); _nobdr(lc)
    rc = t.rows[0].cells[1]; rc.width = right_w;  _nobdr(rc)
    lc._tc.get_or_add_tcPr()  # needed for vertical align

    # Fill left cell
    left_content_fn(lc)

    # Fill right cell: image + caption
    ip = rc.add_paragraph(); ip.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ip.add_run().add_picture(io.BytesIO(right_bytes), width=right_w - Cm(0.3))
    if caption:
        cp = rc.add_paragraph(); cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp.paragraph_format.space_before = Pt(4)
        R(cp, caption, italic=True, size=8, color=RED)

def build(info, figs, out_path):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width=Cm(21); sec.page_height=Cm(29.7)
    sec.left_margin=sec.right_margin=Cm(2)
    sec.top_margin=sec.bottom_margin=Cm(2)
    # Remove default header/footer
    sec.header.is_linked_to_previous = False
    for p in sec.header.paragraphs: p.clear()
    for p in sec.footer.paragraphs: p.clear()
    doc.styles['Normal'].font.name='Calibri'; doc.styles['Normal'].font.size=Pt(10)

    caps = info['captions']
    total = 9  # total pages

    def page_num_footer(doc, page_num):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_before = Pt(8)
        R(p, str(page_num), size=10, color=GRAY)

    # ══ PAGE 1: COVER ════════════════════════════════════════
    logo_p = doc.add_paragraph()
    logo_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    logo_p.paragraph_format.space_before = Pt(4)
    logo_p.paragraph_format.space_after = Pt(6)
    R(logo_p,"ansaldo",bold=True,size=24,color=RGBColor(0x8B,0x8B,0x8B))
    R(logo_p,"|",size=24,color=RED); R(logo_p,"energia",size=24,color=RGBColor(0x8B,0x8B,0x8B))

    tbl=doc.add_table(rows=2,cols=5); tbl.style='Table Grid'
    ws=[Cm(7.5),Cm(3.8),Cm(1.5),Cm(1.5),Cm(1.2)]
    for j,h in enumerate(['Project / Title','Job Number.','Rev.','page','Of']):
        c=tbl.rows[0].cells[j];c.width=ws[j];_bdr(c,'888888',2)
        R(c.paragraphs[0],h,size=8,color=GRAY)
    for j,(val,bold,sz) in enumerate([
        ('SEM Metallurgical Evaluation Report',False,10),
        (f"JC. {info['job']}",True,11),('0',False,10),('1',False,10),(str(total),False,10)]):
        c=tbl.rows[1].cells[j];c.width=ws[j];_bdr(c,'888888',2)
        p=c.paragraphs[0];p.paragraph_format.space_before=Pt(2);p.paragraph_format.space_after=Pt(2)
        if j==1:
            R(p,val,bold=True,size=sz)
            pp=c.add_paragraph();R(pp,info['stage'],size=8,color=GRAY)
        else:R(p,val,bold=bold,size=sz)

    SP(doc,12)
    t1=doc.add_paragraph(); t1.alignment=WD_ALIGN_PARAGRAPH.CENTER
    t1.paragraph_format.space_before=Pt(60)
    R(t1,'SCANNING ELECTRON MICROSCOPY',bold=True,size=20,color=NAVY)
    t2=doc.add_paragraph(); t2.alignment=WD_ALIGN_PARAGRAPH.CENTER
    t2.paragraph_format.space_after=Pt(80)
    R(t2,'METALLURGICAL EVALUATION REPORT',bold=True,size=14,color=GRAY)

    sig=doc.add_table(rows=3,cols=4); sig.style='Table Grid'
    sw=[Cm(3.0),Cm(4.8),Cm(4.8),Cm(2.4)]
    for ri,row in enumerate([['','Name','Title','Date'],
        ['Submitted','Eslam Abdelmawla','Materials Engineer',info['date']],
        ['Approved','Khemichi Badri','Sr. Materials Engineer',info['date']]]):
        for ci,val in enumerate(row):
            c=sig.rows[ri].cells[ci];c.width=sw[ci];_bdr(c)
            if ri==0:_bg(c,'F0F0F0')
            R(c.paragraphs[0],val,bold=(ri==0),size=9)

    page_num_footer(doc,1)
    doc.add_page_break()

    # ══ PAGE 2: TOC ══════════════════════════════════════════
    logo_p=doc.add_paragraph(); logo_p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    R(logo_p,"ansaldo",bold=True,size=24,color=RGBColor(0x8B,0x8B,0x8B))
    R(logo_p,"|",size=24,color=RED); R(logo_p,"energia",size=24,color=RGBColor(0x8B,0x8B,0x8B))
    # header table
    tbl2=doc.add_table(rows=2,cols=5); tbl2.style='Table Grid'
    for j,h in enumerate(['Project / Title','Job Number.','Rev.','page','Of']):
        c=tbl2.rows[0].cells[j];c.width=ws[j];_bdr(c,'888888',2);R(c.paragraphs[0],h,size=8,color=GRAY)
    for j,(val,bold,sz) in enumerate([('SEM Metallurgical Evaluation Report',False,10),(f"JC. {info['job']}",True,11),('0',False,10),('2',False,10),(str(total),False,10)]):
        c=tbl2.rows[1].cells[j];c.width=ws[j];_bdr(c,'888888',2)
        p=c.paragraphs[0];p.paragraph_format.space_before=Pt(2);p.paragraph_format.space_after=Pt(2)
        if j==1:R(p,val,bold=True,size=sz);pp=c.add_paragraph();R(pp,info['stage'],size=8,color=GRAY)
        else:R(p,val,bold=bold,size=sz)
    SP(doc,8)
    h=doc.add_paragraph();h.paragraph_format.space_before=Pt(10);h.paragraph_format.space_after=Pt(8)
    R(h,'TABLE OF CONTENTS',bold=True,size=11,color=NAVY)
    for label,pg in [('TABLE OF CONTENTS','2'),('INTRODUCTION','3'),('RECAPITULATION','3'),
                     ('MICROSTRUCTURE ANALYSIS','4'),
                     ('SUMMARY OF \u03b3\u2032 PRECIPITATE MEASUREMENTS','9'),('CONCLUSION','9')]:
        p=doc.add_paragraph();p.paragraph_format.space_before=Pt(4);p.paragraph_format.space_after=Pt(4)
        R(p,label,size=10);dots='.'*max(5,62-len(label));R(p,f' {dots} {pg}',size=10,color=GRAY)
    page_num_footer(doc,2)
    doc.add_page_break()

    # ══ PAGES 3-9: helper to add the logo+header to each ════
    def add_page_hdr(doc, page_num):
        lp=doc.add_paragraph(); lp.alignment=WD_ALIGN_PARAGRAPH.CENTER
        lp.paragraph_format.space_before=Pt(0); lp.paragraph_format.space_after=Pt(4)
        R(lp,"ansaldo",bold=True,size=24,color=RGBColor(0x8B,0x8B,0x8B))
        R(lp,"|",size=24,color=RED); R(lp,"energia",size=24,color=RGBColor(0x8B,0x8B,0x8B))
        t=doc.add_table(rows=2,cols=5);t.style='Table Grid'
        for j,h in enumerate(['Project / Title','Job Number.','Rev.','page','Of']):
            c=t.rows[0].cells[j];c.width=ws[j];_bdr(c,'888888',2);R(c.paragraphs[0],h,size=8,color=GRAY)
        for j,(val,bold,sz) in enumerate([
            ('SEM Metallurgical Evaluation Report',False,10),(f"JC. {info['job']}",True,11),
            ('0',False,10),(str(page_num),False,10),(str(total),False,10)]):
            c=t.rows[1].cells[j];c.width=ws[j];_bdr(c,'888888',2)
            p=c.paragraphs[0];p.paragraph_format.space_before=Pt(2);p.paragraph_format.space_after=Pt(2)
            if j==1:R(p,val,bold=True,size=sz);pp=c.add_paragraph();R(pp,info['stage'],size=8,color=GRAY)
            else:R(p,val,bold=bold,size=sz)
        SP(doc,6)

    # ══ PAGE 3: INTRO + RECAP (left) | FIG 1.1 (right) ══════
    add_page_hdr(doc,3)

    def left_p3(cell):
        p=cell.add_paragraph(); p.paragraph_format.space_before=Pt(4); p.paragraph_format.space_after=Pt(8)
        R(p,'INTRODUCTION',bold=True,size=11,color=NAVY)
        p2=cell.add_paragraph(); p2.paragraph_format.space_after=Pt(10)
        R(p2,f"This report presents the metallurgical evaluation of a {info['stage']} using "
            f"Scanning Electron Microscopy (SEM). The analysis was performed on the specimen in the "
            f"{info['ht']} condition. The objective is to evaluate microstructural integrity, "
            f"focusing on the Y\u2032 morphology and the presence of any degradation phases such as "
            f"brittle needle-shaped precipitates.",size=10)
        p3=cell.add_paragraph(); p3.paragraph_format.space_before=Pt(10); p3.paragraph_format.space_after=Pt(8)
        R(p3,'RECAPITULATION',bold=True,size=11,color=NAVY)
        for lbl,val in [('Job Number',info['job']),('Alloy',info['material']),
                        ('Incoming Assessment',info['ia']),
                        ('Heat Treatment Condition',info['ht']),('Serial Number',info['serial'])]:
            pb=cell.add_paragraph(); pb.paragraph_format.space_before=Pt(2); pb.paragraph_format.space_after=Pt(2)
            pb.paragraph_format.left_indent=Cm(0.3)
            R(pb,'\u2022 ',bold=True); R(pb,lbl+': ',bold=True); R(pb,val)

    if '1' in figs:
        add_two_col(doc,left_p3,figs['1']['bytes'],right_w=Cm(7.5),caption=caps.get('1','Fig 1.1'))
    else:
        left_p3_para=doc.add_paragraph(); left_p3(left_p3_para)

    page_num_footer(doc,3)
    doc.add_page_break()

    # ══ PAGE 4: MICROSTRUCTURE (left) | FIG 1.2 (right) ═════
    add_page_hdr(doc,4)

    def left_p4(cell):
        p=cell.add_paragraph();p.paragraph_format.space_before=Pt(4);p.paragraph_format.space_after=Pt(8)
        R(p,'MICROSTRUCTURE ANALYSIS',bold=True,size=11,color=NAVY)
        p2=cell.add_paragraph();p2.paragraph_format.space_after=Pt(8)
        R(p2,'The analysis focused on two representative locations, revealing a matrix of Y\u2032 precipitates and various carbide phases.')
        pb=cell.add_paragraph();pb.paragraph_format.left_indent=Cm(0.3)
        R(pb,'\u2022 ');R(pb,'Y Matrix: ',bold=True)
        R(pb,'Both locations showed a typical distribution of primary and secondary Y\u2032 precipitates.')
        pb=cell.add_paragraph();pb.paragraph_format.left_indent=Cm(0.3)
        R(pb,'\u2022 ');R(pb,'Precipitates:',bold=True)
        pb=cell.add_paragraph();pb.paragraph_format.left_indent=Cm(0.8)
        R(pb,'o ');R(pb,'Grain Boundaries: ',bold=True)
        R(pb,'Fine and coarse precipitates, identified as likely M\u2082\u2083C\u2086 and MC-type carbides, were observed along the grain boundaries.')
        pb=cell.add_paragraph();pb.paragraph_format.left_indent=Cm(0.8)
        R(pb,'o ');R(pb,'Intra-granular: ',bold=True)
        R(pb,'Coarse, blocky MC-type precipitates were found within the grains.')
        if info['no_anom']:
            pb=cell.add_paragraph();pb.paragraph_format.left_indent=Cm(0.3)
            R(pb,'\u2022 ');R(pb,'Anomalies: ',bold=True)
            R(pb,'No evidence of detrimental needle-shaped (sigma or eta) precipitates were found at any examined location.')

    if '2' in figs:
        add_two_col(doc,left_p4,figs['2']['bytes'],right_w=Cm(7.5),caption=caps.get('2','Fig 1.2'))
    page_num_footer(doc,4)
    doc.add_page_break()

    # ══ PAGES 5-8: SEM IMAGE GRIDS ═══════════════════════════
    def sem_page(nums, loc_lbl, page_num):
        add_page_hdr(doc, page_num)
        present=[(n,figs[n]) for n in nums if n in figs]
        if not present:return
        cols=len(present)
        img_w=Cm(5.2) if cols==3 else Cm(7.8)
        col_w=Cm(5.5) if cols==3 else Cm(8.2)

        t=doc.add_table(rows=2,cols=cols);t.alignment=WD_TABLE_ALIGNMENT.CENTER
        for ci,(fn,f) in enumerate(present):
            ic=t.rows[0].cells[ci];ic.width=col_w;_nobdr(ic)
            ip=ic.paragraphs[0];ip.alignment=WD_ALIGN_PARAGRAPH.CENTER
            ip.paragraph_format.space_before=Pt(6)
            ip.add_run().add_picture(io.BytesIO(f['bytes']),width=img_w)
            cc=t.rows[1].cells[ci];cc.width=col_w;_nobdr(cc)
            cp=cc.paragraphs[0];cp.alignment=WD_ALIGN_PARAGRAPH.LEFT
            cp.paragraph_format.space_after=Pt(6)
            R(cp,caps.get(fn,f'Fig 1.{fn}'),italic=True,size=9,color=GRAY)

        ll=doc.add_paragraph();ll.alignment=WD_ALIGN_PARAGRAPH.CENTER
        ll.paragraph_format.space_before=Pt(8)
        R(ll,loc_lbl,bold=True,size=12)
        page_num_footer(doc,page_num)
        doc.add_page_break()

    sem_page(['3','4','5'],'Location 1',5)
    sem_page(['6','7'],'Location 1',6)
    sem_page(['8','9','10'],'Location 2',7)
    if '13' in figs:
        sem_page(['11','12','13'],'Location 2',8)
    else:
        sem_page(['11','12'],'Location 2',8)

    # ══ PAGE 9: SUMMARY + CONCLUSION ═════════════════════════
    add_page_hdr(doc,9)

    h=doc.add_paragraph();h.paragraph_format.space_before=Pt(10);h.paragraph_format.space_after=Pt(6)
    R(h,'SUMMARY OF \u03b3\u2032 PRECIPITATE MEASUREMENTS',bold=True,size=11,color=NAVY)
    p=doc.add_paragraph();p.paragraph_format.space_after=Pt(8)
    R(p,'To consolidate the observations from both examined locations, the measured sizes of primary \u03b3\u2032 precipitates are summarized in the table below:')
    SP(doc,4)

    gt=doc.add_table(rows=3,cols=3);gt.alignment=WD_TABLE_ALIGNMENT.CENTER;gt.style='Table Grid'
    gw=[Cm(5.0),Cm(5.0),Cm(5.0)]
    for j,h in enumerate(['Location','Morphology','Average Size (\u03bcm)']):
        c=gt.rows[0].cells[j];c.width=gw[j];_bg(c,'1A1A2E');_bdr(c)
        cp=c.paragraphs[0];cp.alignment=WD_ALIGN_PARAGRAPH.CENTER
        R(cp,h,bold=True,size=10,color=WHITE)
    for ri,row in enumerate([['Location 1','Cuboidal',info['l1']],['Location 2','Cuboidal',info['l2']]]):
        for ci,val in enumerate(row):
            c=gt.rows[ri+1].cells[ci];c.width=gw[ci];_bdr(c)
            if ri==1:_bg(c,'F8F8F8')
            cp=c.paragraphs[0];cp.alignment=WD_ALIGN_PARAGRAPH.CENTER;R(cp,val,size=10)
    tp=doc.add_paragraph();tp.alignment=WD_ALIGN_PARAGRAPH.CENTER
    tp.paragraph_format.space_before=Pt(6);tp.paragraph_format.space_after=Pt(12)
    R(tp,'Table 1 Primary \u03b3\u2032 Precipitate Size Measurements',italic=True,size=9,color=GRAY)

    h=doc.add_paragraph();h.paragraph_format.space_before=Pt(10);h.paragraph_format.space_after=Pt(6)
    R(h,'CONCLUSION',bold=True,size=11,color=NAVY)
    if info['conclusion']:
        # Split on blank lines (each paragraph preserved as-is from _R pdf)
        import re as _re
        parts=_re.split(r'\n\s*\n',info['conclusion'])
        for part in parts:
            clean=_re.sub(r'\s+',' ',part).strip()
            if clean:
                p=doc.add_paragraph();p.paragraph_format.space_after=Pt(8);R(p,clean)

    page_num_footer(doc,9)
    doc.save(out_path)

# ══════════════════════════════════════════════════════════════
def main():
    if len(sys.argv)<2:print(__doc__);sys.exit(1)
    pdf=sys.argv[1]
    if not os.path.exists(pdf):sys.exit(f'Not found: {pdf}')
    out=sys.argv[2] if len(sys.argv)>2 else f'Ansaldo_{Path(pdf).stem}.docx'
    print(f'Converting: {pdf}')
    print('  → Parsing...')
    info=parse(pdf)
    print(f"     {info['stage']}  HT: {info['ht']}")
    print(f"     γ′: L1={info['l1']} L2={info['l2']}")
    print('  → Extracting figures...')
    figs=extract_figures(pdf)
    print(f'     {len(figs)} figures: {sorted(figs.keys(),key=int)}')
    print('  → Building Word document...')
    build(info,figs,out)
    print(f'\n✅  Done: {out}  ({os.path.getsize(out)//1024} KB)')

if __name__=='__main__':main()
