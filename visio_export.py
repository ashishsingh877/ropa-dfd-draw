"""
visio_export.py — Microsoft Visio .vsdx generator (fixed schema)
Uses the exact XML structure Visio Online and Desktop expect.
"""
import io, re, zipfile, math, textwrap

def _sid(s):
    return re.sub(r"[^a-zA-Z0-9_]","_",str(s).strip())[:35]

def _norm(s):
    return re.sub(r"[^a-z0-9]","_",str(s).lower().strip())

def _wrap(t, n=18):
    return "\n".join(textwrap.wrap(str(t).strip(), width=n)[:3])

# Colour helpers
def _rgb(h):
    """#RRGGBB → Visio BGR format #BBGGRR"""
    h = h.lstrip("#")
    return f"#{h[4:6]}{h[2:4]}{h[0:2]}"

PHASE_RANK = {"collection":0,"processing":1,"storage":2,"sharing":3,"exit":4,"main":2}

STYLES = {
    "external":  ("#FFF5CC","#A07800","#5C4400",False,"rounded"),
    "team":      ("#FADADD","#B03030","#641E16",True, "rect"),
    "process":   ("#FFFFFF","#555555","#1A1A1A",False,"rect"),
    "decision":  ("#C0392B","#8B2222","#FFFFFF",True, "diamond"),
    "endpoint":  ("#B03030","#7B1A1A","#FFFFFF",True, "ellipse"),
    "datastore": ("#D4E8FA","#1A5276","#0D3B6E",False,"cylinder"),
    "privacy":   ("#D5E8D4","#27AE60","#145A32",False,"rounded"),
}

# Page dimensions (inches)
PW = 36.0; PH = 14.0; MAR = 0.8

NODE_SIZES = {
    "external":  (2.0, 0.65),
    "team":      (2.1, 0.70),
    "process":   (2.0, 0.65),
    "decision":  (2.0, 1.00),
    "endpoint":  (1.9, 0.65),
    "datastore": (2.0, 0.80),
}
CTRL_W = 1.85; CTRL_H = 0.32; CTRL_GAP = 0.07

def _layout(nodes):
    pm = {}
    for n in nodes:
        ph = n.get("phase","processing").lower()
        if ph not in PHASE_RANK: ph = "processing"
        pm.setdefault(ph,[]).append(n)
    zone_w = (PW - MAR*2)/5
    phase_x = {p: MAR + i*zone_w for i,p in enumerate(
        ["collection","processing","storage","sharing","exit"])}
    phase_x["main"] = phase_x["processing"]
    layout = {}
    for ph, pn in pm.items():
        zx = phase_x.get(ph, MAR)
        cx = zx + zone_w/2
        n  = len(pn)
        sp = (PH - MAR*2 - 1.5)/(n+1)
        for i,node in enumerate(pn):
            nid = _sid(node["id"])
            nt  = node.get("type","process")
            w,h = NODE_SIZES.get(nt,(2.0,0.65))
            cy  = MAR + 1.5 + sp*(i+1)
            layout[nid] = (cx-w/2, cy-h/2, w, h)
    return layout, zone_w

_uid = [1]
def _id(): _uid[0]+=1; return _uid[0]

def _esc(s):
    return (str(s).replace("&","&amp;").replace("<","&lt;")
            .replace(">","&gt;").replace('"',"&quot;"))

def _shape_xml(sid, x, y, w, h, label, ntype):
    fill,line,font,bold,shape = STYLES.get(ntype, STYLES["process"])
    ls = "1" if bold else "0"
    lbl = _esc(_wrap(label,16))

    if shape == "diamond":
        geom = f"""
  <Geom IX='0'>
   <Cell N='NoFill' V='0'/>
   <Cell N='NoLine' V='0'/>
   <MoveTo IX='1'><Cell N='X' V='{w/2:.3f}'/><Cell N='Y' V='0'/></MoveTo>
   <LineTo IX='2'><Cell N='X' V='{w:.3f}'/><Cell N='Y' V='{h/2:.3f}'/></LineTo>
   <LineTo IX='3'><Cell N='X' V='{w/2:.3f}'/><Cell N='Y' V='{h:.3f}'/></LineTo>
   <LineTo IX='4'><Cell N='X' V='0'/><Cell N='Y' V='{h/2:.3f}'/></LineTo>
   <LineTo IX='5'><Cell N='X' V='{w/2:.3f}'/><Cell N='Y' V='0'/></LineTo>
  </Geom>"""
    elif shape == "ellipse":
        geom = f"""
  <Geom IX='0'>
   <Cell N='NoFill' V='0'/><Cell N='NoLine' V='0'/>
   <MoveTo IX='1'><Cell N='X' V='{w/2:.3f}'/><Cell N='Y' V='0'/></MoveTo>
   <Ellipse IX='2'>
    <Cell N='X' V='{w/2:.3f}'/><Cell N='Y' V='{h/2:.3f}'/>
    <Cell N='A' V='{w:.3f}'/><Cell N='B' V='{h/2:.3f}'/>
    <Cell N='C' V='{w/2:.3f}'/><Cell N='D' V='{h:.3f}'/>
   </Ellipse>
  </Geom>"""
    elif shape == "cylinder":
        cap = 0.15
        geom = f"""
  <Geom IX='0'>
   <Cell N='NoFill' V='0'/><Cell N='NoLine' V='0'/>
   <MoveTo IX='1'><Cell N='X' V='0'/><Cell N='Y' V='{cap:.3f}'/></MoveTo>
   <LineTo IX='2'><Cell N='X' V='{w:.3f}'/><Cell N='Y' V='{cap:.3f}'/></LineTo>
   <LineTo IX='3'><Cell N='X' V='{w:.3f}'/><Cell N='Y' V='{h-cap:.3f}'/></LineTo>
   <LineTo IX='4'><Cell N='X' V='0'/><Cell N='Y' V='{h-cap:.3f}'/></LineTo>
   <LineTo IX='5'><Cell N='X' V='0'/><Cell N='Y' V='{cap:.3f}'/></LineTo>
  </Geom>
  <Geom IX='1'>
   <Cell N='NoFill' V='0'/><Cell N='NoLine' V='0'/>
   <MoveTo IX='1'><Cell N='X' V='{w/2:.3f}'/><Cell N='Y' V='{h-cap*2:.3f}'/></MoveTo>
   <Ellipse IX='2'>
    <Cell N='X' V='{w/2:.3f}'/><Cell N='Y' V='{h-cap:.3f}'/>
    <Cell N='A' V='{w:.3f}'/><Cell N='B' V='{h-cap:.3f}'/>
    <Cell N='C' V='{w/2:.3f}'/><Cell N='D' V='{h:.3f}'/>
   </Ellipse>
  </Geom>"""
    elif shape == "rounded":
        geom = ""
    else:
        geom = ""

    rounding = "0.08" if shape in ("rounded","cylinder") else "0"

    return f"""<Shape ID='{sid}' Type='Shape' LineStyle='0' FillStyle='0' TextStyle='0'>
 <XForm>
  <PinX>{x+w/2:.3f}</PinX><PinY>{y+h/2:.3f}</PinY>
  <Width>{w:.3f}</Width><Height>{h:.3f}</Height>
  <LocPinX F='Width*0.5'>0</LocPinX><LocPinY F='Height*0.5'>0</LocPinY>
  <Angle>0</Angle><FlipX>0</FlipX><FlipY>0</FlipY>
 </XForm>
 <Fill><FillForegnd>{_rgb(fill)}</FillForegnd><FillBkgnd>{_rgb(fill)}</FillBkgnd><FillPattern>1</FillPattern></Fill>
 <Line><LineColor>{_rgb(line)}</LineColor><LineWeight>0.02</LineWeight><Rounding>{rounding}</Rounding></Line>
 <Char IX='0'><Color>{_rgb(font)}</Color><Size>0.13</Size><Style>{ls}</Style></Char>
 <Text><cp IX='0'/>{lbl}</Text>{geom}
</Shape>"""


def _connector_xml(cid, from_id, to_id, label, color, dashed=False):
    lbl_xml = f"<Text><cp IX='0'/>{_esc(label[:20])}</Text>" if label else ""
    pattern = "<LinePattern>2</LinePattern>" if dashed else ""
    return f"""<Shape ID='{cid}' Type='Edge' LineStyle='0' FillStyle='0' TextStyle='0'>
 <XForm1D>
  <BeginX F='Sheet.{from_id}!PinX'>0</BeginX><BeginY F='Sheet.{from_id}!PinY'>0</BeginY>
  <EndX F='Sheet.{to_id}!PinX'>1</EndX><EndY F='Sheet.{to_id}!PinY'>1</EndY>
 </XForm1D>
 <Line><LineColor>{_rgb(color)}</LineColor><LineWeight>0.02</LineWeight>{pattern}</Line>
 <Char IX='0'><Color>{_rgb('#555555')}</Color><Size>0.09</Size></Char>
 {lbl_xml}
</Shape>
<Connect FromSheet='{cid}' FromCell='BeginX' ToSheet='{from_id}' ToCell='PinX'/>
<Connect FromSheet='{cid}' FromCell='EndX'   ToSheet='{to_id}'   ToCell='PinX'/>"""


def _build_page_xml(nodes, edges, privacy_controls, show_ctrls, title):
    _uid[0] = 100
    layout, zone_w = _layout(nodes)

    nr = {_sid(n["id"]): PHASE_RANK.get(n.get("phase","processing").lower(),2) for n in nodes}

    # Fuzzy ctrl resolution
    nmap = {}
    for nid in layout:
        nmap[_norm(nid)] = nid
        parts=[p for p in _norm(nid).split("_") if len(p)>2]
        if parts: nmap[parts[0]] = nid
    resolved = {}
    if show_ctrls:
        for k,v in privacy_controls.items():
            sid=_sid(k)
            nid=sid if sid in layout else nmap.get(_norm(k)) or nmap.get(_norm(k).split("_")[0] if "_" in _norm(k) else _norm(k))
            if nid: resolved[nid]=v[:4]

    shapes = []
    connects = []

    # Phase lane BGs
    phase_data = [
        ("collection","① Data Collection","#EAF4FB"),
        ("processing","② Processing",     "#FEF9F0"),
        ("storage",   "③ Storage",        "#F0FAF0"),
        ("sharing",   "④ Sharing",        "#FFF8F0"),
        ("exit",      "⑤ Exit / Archive", "#FDF0F0"),
    ]
    for pi,(ph,lbl,bg) in enumerate(phase_data):
        sid = _id()
        bx = MAR + pi*zone_w; by = 0
        shapes.append(f"""<Shape ID='{sid}' Type='Shape'>
 <XForm><PinX>{bx+zone_w/2:.3f}</PinX><PinY>{PH/2:.3f}</PinY>
  <Width>{zone_w:.3f}</Width><Height>{PH:.3f}</Height>
  <LocPinX F='Width*0.5'>0</LocPinX><LocPinY F='Height*0.5'>0</LocPinY></XForm>
 <Fill><FillForegnd>{_rgb(bg)}</FillForegnd><FillBkgnd>{_rgb(bg)}</FillBkgnd><FillPattern>1</FillPattern></Fill>
 <Line><LineColor>{_rgb('#DDDDDD')}</LineColor><LineWeight>0.01</LineWeight></Line>
 <Char IX='0'><Color>{_rgb('#1A3A5C')}</Color><Size>0.14</Size><Style>1</Style></Char>
 <Para IX='0'><HorzAlign>1</HorzAlign></Para>
 <Text><cp IX='0'/>{_esc(lbl)}</Text>
</Shape>""")

    # Title bar
    tsid = _id()
    shapes.append(f"""<Shape ID='{tsid}' Type='Shape'>
 <XForm><PinX>{PW/2:.3f}</PinX><PinY>{PH-0.4:.3f}</PinY>
  <Width>{PW-MAR*2:.3f}</Width><Height>0.6</Height>
  <LocPinX F='Width*0.5'>0</LocPinX><LocPinY F='Height*0.5'>0</LocPinY></XForm>
 <Fill><FillForegnd>{_rgb('#1A3A5C')}</FillForegnd><FillBkgnd>{_rgb('#1A3A5C')}</FillBkgnd><FillPattern>1</FillPattern></Fill>
 <Line><LineColor>{_rgb('#1A3A5C')}</LineColor><LineWeight>0.01</LineWeight></Line>
 <Char IX='0'><Color>{_rgb('#FFFFFF')}</Color><Size>0.18</Size><Style>1</Style></Char>
 <Para IX='0'><HorzAlign>1</HorzAlign></Para>
 <Text><cp IX='0'/>{_esc(title)}  ·  Privacy &amp; Data Protection Review  ·  DPDPA 2023 / GDPR</Text>
</Shape>""")

    # Node shapes
    nid_sid = {}
    for n in nodes:
        nid = _sid(n["id"])
        sid = _id()
        nid_sid[nid] = sid
        x,y,w,h = layout[nid]
        shapes.append(_shape_xml(sid, x, y, w, h, n.get("label",""), n.get("type","process")))

    # Privacy controls
    if show_ctrls:
        for nid, ctrls in resolved.items():
            if nid not in layout: continue
            x,y,w,h = layout[nid]
            total_h = len(ctrls)*(CTRL_H+CTRL_GAP)-CTRL_GAP
            start_y = y - total_h - 0.18
            parent_sid = nid_sid[nid]
            for ci, ctrl in enumerate(ctrls):
                csid = _id()
                cy = start_y + ci*(CTRL_H+CTRL_GAP)
                cx = x + (w-CTRL_W)/2
                shapes.append(_shape_xml(csid, cx, cy, CTRL_W, CTRL_H, ctrl, "privacy"))
                # dashed connector
                esid = _id()
                shapes.append(f"""<Shape ID='{esid}' Type='Edge'>
 <XForm1D>
  <BeginX F='Sheet.{csid}!PinX'>0</BeginX><BeginY F='Sheet.{csid}!PinY'>0</BeginY>
  <EndX F='Sheet.{parent_sid}!PinX'>1</EndX><EndY F='Sheet.{parent_sid}!PinY'>1</EndY>
 </XForm1D>
 <Line><LineColor>{_rgb('#27AE60')}</LineColor><LineWeight>0.01</LineWeight><LinePattern>2</LinePattern><EndArrow>0</EndArrow></Line>
</Shape>""")
                connects.append(f"<Connect FromSheet='{esid}' FromCell='BeginX' ToSheet='{csid}'      ToCell='PinX'/>")
                connects.append(f"<Connect FromSheet='{esid}' FromCell='EndX'   ToSheet='{parent_sid}' ToCell='PinX'/>")

    # Edge connectors
    for e in edges:
        src=_sid(e.get("from","")); dst=_sid(e.get("to",""))
        if src not in nid_sid or dst not in nid_sid: continue
        raw=e.get("label","")
        sens=any(k in raw.lower() for k in ["health","medical","salary","financial","bank","biometric"])
        back=nr.get(src,2)>=nr.get(dst,2) and src!=dst
        color="#C0392B" if sens else ("#AAAAAA" if back else "#888888")
        xml = _connector_xml(_id(), nid_sid[src], nid_sid[dst], raw, color, back)
        # Split into shape and connect parts
        parts = xml.split("\n")
        shape_lines = []
        for line in parts:
            if line.startswith("<Connect"):
                connects.append(line)
            else:
                shape_lines.append(line)
        shapes.append("\n".join(shape_lines))

    shapes_xml   = "\n".join(shapes)
    connects_xml = "\n".join(connects)

    return f"""<?xml version='1.0' encoding='utf-8' ?>
<PageContents xmlns='http://schemas.microsoft.com/office/visio/2012/main'
  xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  xml:space='preserve'>
 <Shapes>
{shapes_xml}
 </Shapes>
 <Connects>
{connects_xml}
 </Connects>
</PageContents>"""


def generate_vsdx(dfd_data: dict) -> tuple:
    title  = dfd_data.get("process_name","Data Flow Diagram")
    asis   = dfd_data.get("asis",  {"nodes":[],"edges":[]})
    future = dfd_data.get("future",{"nodes":[],"edges":[]})
    ctrls  = dfd_data.get("privacy_controls",{})

    content_types = """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>
 <Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>
 <Default Extension='xml'  ContentType='application/xml'/>
 <Override PartName='/visio/document.xml'     ContentType='application/vnd.ms-visio.drawing.main+xml'/>
 <Override PartName='/visio/pages/page1.xml'  ContentType='application/vnd.ms-visio.page+xml'/>
 <Override PartName='/visio/pages/pages.xml'  ContentType='application/vnd.ms-visio.pages+xml'/>
</Types>"""

    root_rels = """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
 <Relationship Id='rId1'
   Type='http://schemas.microsoft.com/visio/2010/relationships/document'
   Target='visio/document.xml'/>
</Relationships>"""

    doc_rels = """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
 <Relationship Id='rId1'
   Type='http://schemas.microsoft.com/visio/2010/relationships/pages'
   Target='pages/pages.xml'/>
</Relationships>"""

    pages_rels = """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
 <Relationship Id='rId1'
   Type='http://schemas.microsoft.com/visio/2010/relationships/page'
   Target='page1.xml'/>
</Relationships>"""

    def _make(nodes, edges, pc, show_ctrls, state):
        doc_xml = f"""<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<VisioDocument xmlns='http://schemas.microsoft.com/office/visio/2012/main'
  xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>
 <DocumentProperties><Title>{_esc(title)} — {_esc(state)}</Title></DocumentProperties>
 <DocumentSettings><DefaultTextStyle>0</DefaultTextStyle></DocumentSettings>
 <StyleSheets>
  <StyleSheet ID='0' NameU='Normal'>
   <Cell N='LineColor' V='#000000'/><Cell N='FillForegnd' V='#FFFFFF'/>
  </StyleSheet>
 </StyleSheets>
</VisioDocument>"""

        pages_xml = f"""<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Pages xmlns='http://schemas.microsoft.com/office/visio/2012/main'
  xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>
 <Page ID='0' NameU='Page-1' IsCustomName='1'>
  <PageSheet>
   <PageProps>
    <PageWidth>{PW}</PageWidth><PageHeight>{PH}</PageHeight>
    <PageScale>1</PageScale><DrawingScale>1</DrawingScale>
    <DrawingSizeType>0</DrawingSizeType><DrawingScaleType>0</DrawingScaleType>
    <InhibitSnap>0</InhibitSnap><PageLockReplace>0</PageLockReplace>
   </PageProps>
  </PageSheet>
  <Rel r:id='rId1'/>
 </Page>
</Pages>"""

        page_xml = _build_page_xml(nodes, edges, pc, show_ctrls, title)

        buf = io.BytesIO()
        with zipfile.ZipFile(buf,"w",zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml",               content_types.strip())
            zf.writestr("_rels/.rels",                       root_rels.strip())
            zf.writestr("visio/document.xml",                doc_xml.strip())
            zf.writestr("visio/_rels/document.xml.rels",     doc_rels.strip())
            zf.writestr("visio/pages/pages.xml",             pages_xml.strip())
            zf.writestr("visio/pages/_rels/pages.xml.rels",  pages_rels.strip())
            zf.writestr("visio/pages/page1.xml",             page_xml.strip())
        buf.seek(0)
        return buf.getvalue()

    asis_vsdx   = _make(asis["nodes"],   asis["edges"],   {},    False, "Current State")
    future_vsdx = _make(future["nodes"], future["edges"], ctrls, True,  "Post Compliance")
    return asis_vsdx, future_vsdx
