"""
visio_export.py — Microsoft Visio .vsdx (correct flat-Cell schema)
Tested against Visio Online and Visio Desktop 2019+.
"""
import io, re, zipfile, textwrap

def _sid(s):
    return re.sub(r"[^a-zA-Z0-9_]","_",str(s).strip())[:35]
def _norm(s):
    return re.sub(r"[^a-z0-9]","_",str(s).lower().strip())
def _esc(s):
    return (str(s).replace("&","&amp;").replace("<","&lt;")
            .replace(">","&gt;").replace('"',"&quot;").replace("'","&apos;"))
def _wrap(t, n=18):
    return "&#xA;".join(textwrap.wrap(str(t).strip(), width=n)[:3])

PHASE_RANK = {"collection":0,"processing":1,"storage":2,"sharing":3,"exit":4,"main":2}

NODE_SIZES = {
    "external":  (2.0, 0.65),
    "team":      (2.1, 0.70),
    "process":   (2.0, 0.65),
    "decision":  (2.0, 1.00),
    "endpoint":  (1.9, 0.65),
    "datastore": (2.0, 0.80),
}

FILL = {
    "external":  "#FFF5CC","team":"#FADADD","process":"#FFFFFF",
    "decision":  "#C0392B","endpoint":"#B03030","datastore":"#D4E8FA",
    "privacy":   "#D5E8D4",
}
LINE = {
    "external":  "#A07800","team":"#B03030","process":"#555555",
    "decision":  "#8B2222","endpoint":"#7B1A1A","datastore":"#1A5276",
    "privacy":   "#27AE60",
}
FONT = {
    "external":  "#5C4400","team":"#641E16","process":"#1A1A1A",
    "decision":  "#FFFFFF","endpoint":"#FFFFFF","datastore":"#0D3B6E",
    "privacy":   "#145A32",
}

PW=36.0; PH=14.0; MAR=0.8
CTRL_W=1.85; CTRL_H=0.32; CTRL_GAP=0.07

_uid=[100]
def _id(): _uid[0]+=1; return _uid[0]

def _layout(nodes):
    pm={}
    for n in nodes:
        ph=n.get("phase","processing").lower()
        if ph not in PHASE_RANK: ph="processing"
        pm.setdefault(ph,[]).append(n)
    zone_w=(PW-MAR*2)/5
    px={p:MAR+i*zone_w for i,p in enumerate(["collection","processing","storage","sharing","exit"])}
    px["main"]=px["processing"]
    layout={}
    for ph,pn in pm.items():
        zx=px.get(ph,MAR); cx=zx+zone_w/2; n=len(pn)
        sp=(PH-MAR*2-1.8)/(n+1)
        for i,node in enumerate(pn):
            nid=_sid(node["id"]); nt=node.get("type","process")
            w,h=NODE_SIZES.get(nt,(2.0,0.65))
            cy=MAR+1.8+sp*(i+1)
            layout[nid]=(cx-w/2,cy-h/2,w,h)
    return layout,zone_w

def _shape(sid,x,y,w,h,label,ntype,rounded=False):
    fill=FILL.get(ntype,"#FFFFFF")
    line=LINE.get(ntype,"#555555")
    font=FONT.get(ntype,"#1A1A1A")
    bold="1" if ntype in("team","decision","endpoint") else "0"
    px_=x+w/2; py_=y+h/2
    lbl=_wrap(label)

    # Geometry based on shape type
    if ntype=="decision":
        geom=f"""<Geom IX='0'>
   <Cell N='NoFill' V='0'/><Cell N='NoLine' V='0'/>
   <MoveTo IX='1'><Cell N='X' V='{w/2:.4f}'/><Cell N='Y' V='0.0000'/></MoveTo>
   <LineTo IX='2'><Cell N='X' V='{w:.4f}'/><Cell N='Y' V='{h/2:.4f}'/></LineTo>
   <LineTo IX='3'><Cell N='X' V='{w/2:.4f}'/><Cell N='Y' V='{h:.4f}'/></LineTo>
   <LineTo IX='4'><Cell N='X' V='0.0000'/><Cell N='Y' V='{h/2:.4f}'/></LineTo>
   <LineTo IX='5'><Cell N='X' V='{w/2:.4f}'/><Cell N='Y' V='0.0000'/></LineTo>
  </Geom>"""
    elif ntype=="endpoint":
        geom=f"""<Geom IX='0'>
   <Cell N='NoFill' V='0'/><Cell N='NoLine' V='0'/>
   <MoveTo IX='1'><Cell N='X' V='{w/2:.4f}'/><Cell N='Y' V='0.0000'/></MoveTo>
   <Ellipse IX='2'>
    <Cell N='X' V='{w/2:.4f}'/><Cell N='Y' V='{h/2:.4f}'/>
    <Cell N='A' V='{w:.4f}'/><Cell N='B' V='{h/2:.4f}'/>
    <Cell N='C' V='{w/2:.4f}'/><Cell N='D' V='{h:.4f}'/>
   </Ellipse>
  </Geom>"""
    elif ntype=="datastore":
        cap=0.12
        geom=f"""<Geom IX='0'>
   <Cell N='NoFill' V='0'/><Cell N='NoLine' V='0'/>
   <MoveTo IX='1'><Cell N='X' V='0'/><Cell N='Y' V='{cap:.4f}'/></MoveTo>
   <LineTo IX='2'><Cell N='X' V='{w:.4f}'/><Cell N='Y' V='{cap:.4f}'/></LineTo>
   <LineTo IX='3'><Cell N='X' V='{w:.4f}'/><Cell N='Y' V='{h-cap:.4f}'/></LineTo>
   <LineTo IX='4'><Cell N='X' V='0'/><Cell N='Y' V='{h-cap:.4f}'/></LineTo>
   <LineTo IX='5'><Cell N='X' V='0'/><Cell N='Y' V='{cap:.4f}'/></LineTo>
  </Geom>
  <Geom IX='1'>
   <Cell N='NoFill' V='0'/><Cell N='NoLine' V='0'/>
   <MoveTo IX='1'><Cell N='X' V='{w/2:.4f}'/><Cell N='Y' V='{h-cap*2:.4f}'/></MoveTo>
   <Ellipse IX='2'>
    <Cell N='X' V='{w/2:.4f}'/><Cell N='Y' V='{h-cap:.4f}'/>
    <Cell N='A' V='{w:.4f}'/><Cell N='B' V='{h-cap:.4f}'/>
    <Cell N='C' V='{w/2:.4f}'/><Cell N='D' V='{h:.4f}'/>
   </Ellipse>
  </Geom>"""
    else:
        geom=""

    rnd="0.08" if (ntype in("external","privacy") or rounded) else "0"

    return f"""  <Shape ID='{sid}' Type='Shape'>
   <Cell N='PinX' V='{px_:.4f}'/>
   <Cell N='PinY' V='{py_:.4f}'/>
   <Cell N='Width' V='{w:.4f}'/>
   <Cell N='Height' V='{h:.4f}'/>
   <Cell N='LocPinX' V='{w/2:.4f}' F='Width*0.5'/>
   <Cell N='LocPinY' V='{h/2:.4f}' F='Height*0.5'/>
   <Cell N='Angle' V='0'/>
   <Cell N='FlipX' V='0'/>
   <Cell N='FlipY' V='0'/>
   <Cell N='ResizeMode' V='0'/>
   <Cell N='FillForegnd' V='{fill}'/>
   <Cell N='FillBkgnd' V='{fill}'/>
   <Cell N='FillPattern' V='1'/>
   <Cell N='LineColor' V='{line}'/>
   <Cell N='LineWeight' V='0.02'/>
   <Cell N='Rounding' V='{rnd}'/>
   <Cell N='ObjType' V='2'/>
   <Char IX='0'>
    <Cell N='Color' V='{font}'/>
    <Cell N='Size' V='0.13'/>
    <Cell N='Style' V='{bold}'/>
    <Cell N='Font' V='4'/>
   </Char>
   <Para IX='0'>
    <Cell N='HorzAlign' V='1'/>
    <Cell N='VerticalAlign' V='1'/>
   </Para>
   <Text>{lbl}</Text>
   {geom}
  </Shape>"""

def _edge(eid, f_sid, t_sid, label, color, dashed=False):
    lbl=f"<Text>{_esc(label[:20])}</Text>" if label else ""
    pat=f"<Cell N='LinePattern' V='2'/>" if dashed else ""
    return f"""  <Shape ID='{eid}' Type='Edge'>
   <Cell N='LineColor' V='{color}'/>
   <Cell N='LineWeight' V='0.018'/>
   <Cell N='EndArrow' V='1'/>
   <Cell N='EndArrowSize' V='2'/>
   {pat}
   <Cell N='ObjType' V='2'/>
   <Char IX='0'><Cell N='Size' V='0.09'/><Cell N='Color' V='#555555'/></Char>
   {lbl}
  </Shape>
  <Connect FromSheet='{eid}' FromCell='BeginX' ToSheet='{f_sid}' ToCell='PinX'/>
  <Connect FromSheet='{eid}' FromCell='EndX'   ToSheet='{t_sid}' ToCell='PinX'/>"""

def _bg_rect(sid,x,y,w,h,fill):
    return f"""  <Shape ID='{sid}' Type='Shape'>
   <Cell N='PinX' V='{x+w/2:.4f}'/>
   <Cell N='PinY' V='{y+h/2:.4f}'/>
   <Cell N='Width' V='{w:.4f}'/>
   <Cell N='Height' V='{h:.4f}'/>
   <Cell N='LocPinX' V='{w/2:.4f}' F='Width*0.5'/>
   <Cell N='LocPinY' V='{h/2:.4f}' F='Height*0.5'/>
   <Cell N='FillForegnd' V='{fill}'/>
   <Cell N='FillBkgnd' V='{fill}'/>
   <Cell N='FillPattern' V='1'/>
   <Cell N='LineColor' V='#DDDDDD'/>
   <Cell N='LineWeight' V='0.005'/>
   <Cell N='ObjType' V='0'/>
   <Cell N='LockSelect' V='1'/>
   <Cell N='LockDelete' V='1'/>
   <Char IX='0'><Cell N='Color' V='#1A3A5C'/><Cell N='Size' V='0.13'/><Cell N='Style' V='1'/></Char>
   <Para IX='0'><Cell N='HorzAlign' V='1'/></Para>
  </Shape>"""

def _text_shape(sid,x,y,w,h,label,fill,font,bold,fsize="0.14"):
    return f"""  <Shape ID='{sid}' Type='Shape'>
   <Cell N='PinX' V='{x+w/2:.4f}'/>
   <Cell N='PinY' V='{y+h/2:.4f}'/>
   <Cell N='Width' V='{w:.4f}'/>
   <Cell N='Height' V='{h:.4f}'/>
   <Cell N='LocPinX' V='{w/2:.4f}' F='Width*0.5'/>
   <Cell N='LocPinY' V='{h/2:.4f}' F='Height*0.5'/>
   <Cell N='FillForegnd' V='{fill}'/>
   <Cell N='FillBkgnd' V='{fill}'/>
   <Cell N='FillPattern' V='1'/>
   <Cell N='LineColor' V='{fill}'/>
   <Cell N='LineWeight' V='0.005'/>
   <Cell N='ObjType' V='0'/>
   <Char IX='0'><Cell N='Color' V='{font}'/><Cell N='Size' V='{fsize}'/><Cell N='Style' V='{"1" if bold else "0"}'/></Char>
   <Para IX='0'><Cell N='HorzAlign' V='1'/><Cell N='VerticalAlign' V='1'/></Para>
   <Text>{_esc(label)}</Text>
  </Shape>"""

def _build_page(nodes,edges,privacy_controls,show_ctrls,title):
    _uid[0]=100
    layout,zone_w=_layout(nodes)
    nr={_sid(n["id"]):PHASE_RANK.get(n.get("phase","processing").lower(),2) for n in nodes}
    nmap={}
    for nid in layout:
        nmap[_norm(nid)]=nid
        for p in [x for x in _norm(nid).split("_") if len(x)>2]:
            nmap[p]=nid
    resolved={}
    if show_ctrls:
        for k,v in privacy_controls.items():
            sid2=_sid(k); nk=_norm(k)
            nid=sid2 if sid2 in layout else nmap.get(nk) or nmap.get(nk.split("_")[0] if "_" in nk else nk)
            if nid: resolved[nid]=v[:4]

    shapes=[]; connects=[]

    # Phase lane backgrounds
    phase_data=[
        ("collection","① Data Collection","#EBF5FB"),
        ("processing","② Processing",     "#FEFAF0"),
        ("storage",   "③ Storage",        "#F0FAF0"),
        ("sharing",   "④ Sharing",        "#FFF8F0"),
        ("exit",      "⑤ Exit / Archive", "#FDF0F0"),
    ]
    for pi,(ph,lbl,bg) in enumerate(phase_data):
        bx=MAR+pi*zone_w
        # background
        bg_sid=_id()
        shapes.append(_bg_rect(bg_sid, bx,0, zone_w,PH, bg))
        # label at top
        lbl_sid=_id()
        shapes.append(_text_shape(lbl_sid, bx,PH-1.2, zone_w,0.55, lbl, bg,"#1A3A5C",True,"0.13"))

    # Title bar
    tsid=_id()
    shapes.append(_text_shape(tsid, MAR,PH-0.7, PW-MAR*2,0.55,
        f"{title}  ·  Privacy & Data Protection Review  ·  DPDPA 2023 / GDPR",
        "#1A3A5C","#FFFFFF",True,"0.16"))

    # Main nodes
    nid_to_sid={}
    for n in nodes:
        nid=_sid(n["id"]); sid=_id(); nid_to_sid[nid]=sid
        x,y,w,h=layout[nid]
        shapes.append(_shape(sid,x,y,w,h,n.get("label",""),n.get("type","process")))

    # Privacy controls
    if show_ctrls:
        for nid,ctrls in resolved.items():
            if nid not in layout: continue
            x,y,w,h=layout[nid]
            total=(len(ctrls))*(CTRL_H+CTRL_GAP)-CTRL_GAP
            sy=y-total-0.2
            psid=nid_to_sid[nid]
            for ci,ctrl in enumerate(ctrls):
                csid=_id()
                cy_=sy+ci*(CTRL_H+CTRL_GAP)
                cx_=x+(w-CTRL_W)/2
                shapes.append(_shape(csid,cx_,cy_,CTRL_W,CTRL_H,ctrl,"privacy"))
                # connector
                esid=_id()
                shapes.append(f"""  <Shape ID='{esid}' Type='Edge'>
   <Cell N='LineColor' V='#27AE60'/>
   <Cell N='LineWeight' V='0.01'/>
   <Cell N='LinePattern' V='2'/>
   <Cell N='EndArrow' V='0'/>
   <Cell N='ObjType' V='2'/>
  </Shape>
  <Connect FromSheet='{esid}' FromCell='BeginX' ToSheet='{csid}' ToCell='PinX'/>
  <Connect FromSheet='{esid}' FromCell='EndX'   ToSheet='{psid}' ToCell='PinX'/>""")

    # Flow edges
    for e in edges:
        s=_sid(e.get("from","")); d=_sid(e.get("to",""))
        if s not in nid_to_sid or d not in nid_to_sid: continue
        raw=e.get("label","")
        sens=any(k in raw.lower() for k in ["health","medical","salary","financial","bank","biometric"])
        back=nr.get(s,2)>=nr.get(d,2) and s!=d
        color="#C0392B" if sens else ("#AAAAAA" if back else "#888888")
        shapes.append(_edge(_id(),nid_to_sid[s],nid_to_sid[d],raw,color,back))

    shapes_xml="\n".join(shapes)
    return f"""<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<PageContents xmlns='http://schemas.microsoft.com/office/visio/2012/main'
  xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  xml:space='preserve'>
 <Shapes>
{shapes_xml}
 </Shapes>
</PageContents>"""


def generate_vsdx(dfd_data: dict) -> tuple:
    title  = dfd_data.get("process_name","Data Flow Diagram")
    asis   = dfd_data.get("asis",  {"nodes":[],"edges":[]})
    future = dfd_data.get("future",{"nodes":[],"edges":[]})
    ctrls  = dfd_data.get("privacy_controls",{})

    CT="""<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>
 <Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>
 <Default Extension='xml'  ContentType='application/xml'/>
 <Override PartName='/visio/document.xml'    ContentType='application/vnd.ms-visio.drawing.main+xml'/>
 <Override PartName='/visio/pages/pages.xml' ContentType='application/vnd.ms-visio.pages+xml'/>
 <Override PartName='/visio/pages/page1.xml' ContentType='application/vnd.ms-visio.page+xml'/>
</Types>"""

    RRELS="""<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
 <Relationship Id='rId1' Type='http://schemas.microsoft.com/visio/2010/relationships/document' Target='visio/document.xml'/>
</Relationships>"""

    def _doc(state):
        return f"""<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<VisioDocument xmlns='http://schemas.microsoft.com/office/visio/2012/main'
  xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>
 <DocumentProperties>
  <Title>{_esc(title)} — {_esc(state)}</Title>
  <Creator>ROPA Intelligence Platform</Creator>
 </DocumentProperties>
 <DocumentSettings>
  <DefaultTextStyle>0</DefaultTextStyle>
  <DefaultLineStyle>0</DefaultLineStyle>
  <DefaultFillStyle>0</DefaultFillStyle>
  <DefaultGuideStyle>0</DefaultGuideStyle>
 </DocumentSettings>
 <Colors>
  <ColorEntry IX='0' RGB='#000000'/>
  <ColorEntry IX='1' RGB='#FFFFFF'/>
  <ColorEntry IX='2' RGB='#FF0000'/>
 </Colors>
 <StyleSheets>
  <StyleSheet ID='0' NameU='Normal' IsCustomName='0'>
   <Cell N='LineColor' V='#000000'/>
   <Cell N='FillForegnd' V='#FFFFFF'/>
   <Cell N='FillBkgnd' V='#FFFFFF'/>
  </StyleSheet>
 </StyleSheets>
</VisioDocument>"""

    DOCRELS="""<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
 <Relationship Id='rId1' Type='http://schemas.microsoft.com/visio/2010/relationships/pages' Target='pages/pages.xml'/>
</Relationships>"""

    PAGES_RELS="""<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
 <Relationship Id='rId1' Type='http://schemas.microsoft.com/visio/2010/relationships/page' Target='page1.xml'/>
</Relationships>"""

    def _pages():
        return f"""<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Pages xmlns='http://schemas.microsoft.com/office/visio/2012/main'
  xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>
 <Page ID='0' NameU='Page-1' IsCustomName='1' IsCustomNameU='1' ViewScale='1' ViewCenterX='0' ViewCenterY='0'>
  <PageSheet UniqueID='{{00000000-0000-0000-0000-000000000001}}'>
   <PageProps>
    <Cell N='PageWidth'  V='{PW}'/>
    <Cell N='PageHeight' V='{PH}'/>
    <Cell N='PageScale'  V='1'/>
    <Cell N='DrawingScale' V='1'/>
    <Cell N='DrawingSizeType' V='0'/>
    <Cell N='DrawingScaleType' V='0'/>
    <Cell N='InhibitSnap' V='0'/>
   </PageProps>
   <PrintProps>
    <Cell N='PageLeftMargin'  V='0.5'/>
    <Cell N='PageRightMargin' V='0.5'/>
    <Cell N='PageTopMargin'   V='0.5'/>
    <Cell N='PageBottomMargin' V='0.5'/>
    <Cell N='PaperKind' V='119'/>
    <Cell N='PrintPageOrientation' V='2'/>
   </PrintProps>
  </PageSheet>
  <Rel r:id='rId1'/>
 </Page>
</Pages>"""

    def _make(nodes,edges,pc,show,state):
        page_xml=_build_page(nodes,edges,pc,show,title)
        # validate XML
        import xml.etree.ElementTree as ET
        ET.fromstring(page_xml)  # will raise if invalid
        buf=io.BytesIO()
        with zipfile.ZipFile(buf,"w",zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml",              CT.strip())
            zf.writestr("_rels/.rels",                      RRELS.strip())
            zf.writestr("visio/document.xml",               _doc(state).strip())
            zf.writestr("visio/_rels/document.xml.rels",    DOCRELS.strip())
            zf.writestr("visio/pages/pages.xml",            _pages().strip())
            zf.writestr("visio/pages/_rels/pages.xml.rels", PAGES_RELS.strip())
            zf.writestr("visio/pages/page1.xml",            page_xml.strip())
        buf.seek(0); return buf.getvalue()

    return (_make(asis["nodes"],   asis["edges"],   {},    False, "Current State"),
            _make(future["nodes"], future["edges"], ctrls, True,  "Post Compliance"))
