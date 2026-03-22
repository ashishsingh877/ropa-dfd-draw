"""
visio_export.py
===============
Generates a proper .vsdx file (Microsoft Visio) from DFD data.
VSDX is a ZIP archive containing XML files following the Open Packaging Convention.
Opens natively in Visio 2013+ and Visio for Microsoft 365.
All shapes are fully editable — resize, restyle, reconnect, add text.
"""
import io, re, zipfile, math, textwrap
from xml.etree import ElementTree as ET

def _sid(s):
    return re.sub(r"[^a-zA-Z0-9_]", "_", str(s).strip())[:35]

def _norm(s):
    return re.sub(r"[^a-z0-9]", "_", str(s).lower().strip())

def _wrap(text, n=18):
    return "\n".join(textwrap.wrap(str(text).strip(), width=n)[:3])

# ── Visio shape colours ────────────────────────────────────────────────────────
SHAPE_STYLES = {
    "external":  dict(fill="#FFF5CC", line="#A07800", font="#5C4400",
                      shape="rounded_rect", bold=False),
    "team":      dict(fill="#FADADD", line="#B03030", font="#641E16",
                      shape="rect",         bold=True),
    "process":   dict(fill="#FFFFFF", line="#555555", font="#1A1A1A",
                      shape="rect",         bold=False),
    "decision":  dict(fill="#C0392B", line="#8B2222", font="#FFFFFF",
                      shape="diamond",      bold=True),
    "endpoint":  dict(fill="#B03030", line="#7B1A1A", font="#FFFFFF",
                      shape="ellipse",      bold=True),
    "datastore": dict(fill="#D4E8FA", line="#1A5276", font="#0D3B6E",
                      shape="cylinder",     bold=False),
    "privacy":   dict(fill="#D5E8D4", line="#27AE60", font="#145A32",
                      shape="rounded_rect", bold=False),
}

# ── Layout constants (in Visio inches, 1 inch = 914400 EMUs) ─────────────────
PAGE_W   = 36.0    # inches — landscape A0-ish
PAGE_H   = 14.0
MARGIN   = 0.8
NODE_W   = 2.0
NODE_H   = 0.75
DEC_W    = 2.0
DEC_H    = 1.0
DS_W     = 2.0
DS_H     = 0.85
EP_W     = 1.9
EP_H     = 0.7
CTRL_W   = 1.9
CTRL_H   = 0.38
CTRL_GAP = 0.08

PHASE_RANK = {"collection":0,"processing":1,"storage":2,"sharing":3,"exit":4,"main":2}

def _hex_to_rgb01(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i+2],16)/255.0 for i in (0,2,4))

def _rgb_to_visio(h):
    """Convert #RRGGBB to Visio color format #BBGGRR."""
    h = h.lstrip("#")
    r,g,b = h[0:2],h[2:4],h[4:6]
    return f"#{b}{g}{r}"


# ── XML namespace helpers ──────────────────────────────────────────────────────
NS = {
    "vsdx": "http://schemas.microsoft.com/office/visio/2012/main",
    "r":    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
VISIO_NS = "http://schemas.microsoft.com/office/visio/2012/main"
REL_NS   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

def _el(tag, attrib=None, text=None, ns=VISIO_NS):
    e = ET.Element(f"{{{ns}}}{tag}", attrib or {})
    if text is not None: e.text = str(text)
    return e

def _sub(parent, tag, attrib=None, text=None, ns=VISIO_NS):
    e = _el(tag, attrib, text, ns)
    parent.append(e)
    return e


# ── Layout calculator ──────────────────────────────────────────────────────────
def _compute_layout(nodes):
    """Return {node_id: (x,y,w,h)} in Visio inches."""
    phase_map = {}
    for n in nodes:
        ph = n.get("phase","processing").lower()
        if ph not in PHASE_RANK: ph = "processing"
        phase_map.setdefault(ph,[]).append(n)

    usable_w = PAGE_W - MARGIN*2
    n_phases  = 5
    zone_w    = usable_w / n_phases
    usable_h  = PAGE_H  - MARGIN*2 - 2.0   # leave room for phase labels

    phase_x_start = {
        "collection": MARGIN,
        "processing": MARGIN + zone_w,
        "storage":    MARGIN + zone_w*2,
        "sharing":    MARGIN + zone_w*3,
        "exit":       MARGIN + zone_w*4,
        "main":       MARGIN + zone_w,
    }

    NODE_SIZES = {
        "external":  (NODE_W, NODE_H),
        "team":      (NODE_W+0.2, NODE_H+0.1),
        "process":   (NODE_W, NODE_H),
        "decision":  (DEC_W, DEC_H),
        "endpoint":  (EP_W, EP_H),
        "datastore": (DS_W, DS_H),
    }

    layout = {}
    for ph, ph_nodes in phase_map.items():
        x_start = phase_x_start.get(ph, MARGIN)
        zone_cx  = x_start + zone_w/2
        n = len(ph_nodes)
        spacing  = usable_h / (n+1)
        for i, node in enumerate(ph_nodes):
            nid   = _sid(node["id"])
            ntype = node.get("type","process")
            w,h   = NODE_SIZES.get(ntype,(NODE_W,NODE_H))
            cy    = MARGIN + 1.5 + spacing*(i+1)   # 1.5 for phase labels
            cx    = zone_cx
            layout[nid] = (cx - w/2, cy - h/2, w, h)
    return layout


# ── Shape ID counter ───────────────────────────────────────────────────────────
_shape_id = [1000]
def _next_id():
    _shape_id[0] += 1
    return _shape_id[0]


# ── Shape XML builders ─────────────────────────────────────────────────────────
def _make_rect_shape(sid, x, y, w, h, label, fill, line, font, bold,
                     rounded=False):
    """Build a Visio Shape XML element for a rectangle/rounded-rect."""
    shape = _el("Shape", {
        "ID":     str(sid),
        "Type":   "Shape",
        "Width":  f"{w:.4f}",
        "Height": f"{h:.4f}",
        "PinX":   f"{x+w/2:.4f}",
        "PinY":   f"{y+h/2:.4f}",
    })
    # Geometry (rectangle or rounded)
    _sub(shape,"Cell",{"N":"Rounding","V":"0.1" if rounded else "0"})
    # Fill
    _sub(shape,"Cell",{"N":"FillForegnd","V":_rgb_to_visio(fill)})
    _sub(shape,"Cell",{"N":"FillBkgnd",  "V":_rgb_to_visio(fill)})
    # Line
    _sub(shape,"Cell",{"N":"LineColor","V":_rgb_to_visio(line)})
    _sub(shape,"Cell",{"N":"LineWeight","V":"0.02"})
    # Font
    _sub(shape,"Cell",{"N":"CharColor","V":_rgb_to_visio(font)})
    _sub(shape,"Cell",{"N":"CharStyle","V":"1" if bold else "0"})
    _sub(shape,"Cell",{"N":"CharSize", "V":"0.14"})
    # Text
    txt = _sub(shape,"Text")
    txt.text = _wrap(label,18)
    return shape


def _make_diamond_shape(sid, x, y, w, h, label, fill, line, font):
    """Build a Visio diamond (decision) shape."""
    shape = _el("Shape",{
        "ID":str(sid),"Type":"Shape",
        "Width":f"{w:.4f}","Height":f"{h:.4f}",
        "PinX":f"{x+w/2:.4f}","PinY":f"{y+h/2:.4f}",
    })
    # Use Visio's built-in diamond geometry (master shape trick via NameU)
    _sub(shape,"Cell",{"N":"Rounding","V":"0"})
    _sub(shape,"Cell",{"N":"FillForegnd","V":_rgb_to_visio(fill)})
    _sub(shape,"Cell",{"N":"FillBkgnd",  "V":_rgb_to_visio(fill)})
    _sub(shape,"Cell",{"N":"LineColor",  "V":_rgb_to_visio(line)})
    _sub(shape,"Cell",{"N":"LineWeight", "V":"0.025"})
    _sub(shape,"Cell",{"N":"CharColor",  "V":_rgb_to_visio(font)})
    _sub(shape,"Cell",{"N":"CharStyle",  "V":"1"})
    _sub(shape,"Cell",{"N":"CharSize",   "V":"0.13"})
    # Diamond geometry
    geom = _sub(shape,"Geom",{"IX":"0"})
    _sub(geom,"Cell",{"N":"NoFill","V":"0"})
    _sub(geom,"Cell",{"N":"NoLine","V":"0"})
    mv = _sub(geom,"MoveTo",{"IX":"1"})
    _sub(mv,"Cell",{"N":"X","V":f"{w/2:.4f}"})
    _sub(mv,"Cell",{"N":"Y","V":"0"})
    ln1 = _sub(geom,"LineTo",{"IX":"2"})
    _sub(ln1,"Cell",{"N":"X","V":f"{w:.4f}"})
    _sub(ln1,"Cell",{"N":"Y","V":f"{h/2:.4f}"})
    ln2 = _sub(geom,"LineTo",{"IX":"3"})
    _sub(ln2,"Cell",{"N":"X","V":f"{w/2:.4f}"})
    _sub(ln2,"Cell",{"N":"Y","V":f"{h:.4f}"})
    ln3 = _sub(geom,"LineTo",{"IX":"4"})
    _sub(ln3,"Cell",{"N":"X","V":"0"})
    _sub(ln3,"Cell",{"N":"Y","V":f"{h/2:.4f}"})
    ln4 = _sub(geom,"LineTo",{"IX":"5"})
    _sub(ln4,"Cell",{"N":"X","V":f"{w/2:.4f}"})
    _sub(ln4,"Cell",{"N":"Y","V":"0"})
    txt = _sub(shape,"Text")
    txt.text = _wrap(label,14)
    return shape


def _make_ellipse_shape(sid, x, y, w, h, label, fill, line, font, bold):
    shape = _el("Shape",{
        "ID":str(sid),"Type":"Shape",
        "Width":f"{w:.4f}","Height":f"{h:.4f}",
        "PinX":f"{x+w/2:.4f}","PinY":f"{y+h/2:.4f}",
    })
    _sub(shape,"Cell",{"N":"FillForegnd","V":_rgb_to_visio(fill)})
    _sub(shape,"Cell",{"N":"FillBkgnd",  "V":_rgb_to_visio(fill)})
    _sub(shape,"Cell",{"N":"LineColor",  "V":_rgb_to_visio(line)})
    _sub(shape,"Cell",{"N":"LineWeight", "V":"0.025"})
    _sub(shape,"Cell",{"N":"CharColor",  "V":_rgb_to_visio(font)})
    _sub(shape,"Cell",{"N":"CharStyle",  "V":"1" if bold else "0"})
    _sub(shape,"Cell",{"N":"CharSize",   "V":"0.13"})
    # Ellipse geometry
    geom = _sub(shape,"Geom",{"IX":"0"})
    _sub(geom,"Cell",{"N":"NoFill","V":"0"})
    _sub(geom,"Cell",{"N":"NoLine","V":"0"})
    mv = _sub(geom,"MoveTo",{"IX":"1"})
    _sub(mv,"Cell",{"N":"X","V":f"{w/2:.4f}"})
    _sub(mv,"Cell",{"N":"Y","V":"0"})
    el = _sub(geom,"Ellipse",{"IX":"2"})
    _sub(el,"Cell",{"N":"X",  "V":f"{w/2:.4f}"})
    _sub(el,"Cell",{"N":"Y",  "V":f"{h/2:.4f}"})
    _sub(el,"Cell",{"N":"A",  "V":f"{w:.4f}"})
    _sub(el,"Cell",{"N":"B",  "V":f"{h/2:.4f}"})
    _sub(el,"Cell",{"N":"C",  "V":f"{w/2:.4f}"})
    _sub(el,"Cell",{"N":"D",  "V":f"{h:.4f}"})
    txt = _sub(shape,"Text")
    txt.text = _wrap(label,16)
    return shape


def _make_cylinder_shape(sid, x, y, w, h, label, fill, line, font):
    """Cylinder using a rect + top ellipse."""
    shape = _el("Shape",{
        "ID":str(sid),"Type":"Shape",
        "Width":f"{w:.4f}","Height":f"{h:.4f}",
        "PinX":f"{x+w/2:.4f}","PinY":f"{y+h/2:.4f}",
    })
    _sub(shape,"Cell",{"N":"FillForegnd","V":_rgb_to_visio(fill)})
    _sub(shape,"Cell",{"N":"FillBkgnd",  "V":_rgb_to_visio(fill)})
    _sub(shape,"Cell",{"N":"LineColor",  "V":_rgb_to_visio(line)})
    _sub(shape,"Cell",{"N":"LineWeight", "V":"0.018"})
    _sub(shape,"Cell",{"N":"CharColor",  "V":_rgb_to_visio(font)})
    _sub(shape,"Cell",{"N":"CharSize",   "V":"0.13"})
    cap = 0.18
    # Body
    geom0 = _sub(shape,"Geom",{"IX":"0"})
    _sub(geom0,"Cell",{"N":"NoFill","V":"0"})
    mv0 = _sub(geom0,"MoveTo",{"IX":"1"})
    _sub(mv0,"Cell",{"N":"X","V":"0"})
    _sub(mv0,"Cell",{"N":"Y","V":f"{cap:.4f}"})
    for nx,ny,ix in [(w,cap,2),(w,h-cap,3),(0,h-cap,4),(0,cap,5)]:
        ln=_sub(geom0,"LineTo",{"IX":str(ix)})
        _sub(ln,"Cell",{"N":"X","V":f"{nx:.4f}"})
        _sub(ln,"Cell",{"N":"Y","V":f"{ny:.4f}"})
    # Bottom ellipse
    geom1 = _sub(shape,"Geom",{"IX":"1"})
    mv1 = _sub(geom1,"MoveTo",{"IX":"1"})
    _sub(mv1,"Cell",{"N":"X","V":f"{w/2:.4f}"})
    _sub(mv1,"Cell",{"N":"Y","V":"0"})
    el1 = _sub(geom1,"Ellipse",{"IX":"2"})
    _sub(el1,"Cell",{"N":"X","V":f"{w/2:.4f}"})
    _sub(el1,"Cell",{"N":"Y","V":f"{cap:.4f}"})
    _sub(el1,"Cell",{"N":"A","V":f"{w:.4f}"})
    _sub(el1,"Cell",{"N":"B","V":f"{cap:.4f}"})
    _sub(el1,"Cell",{"N":"C","V":f"{w/2:.4f}"})
    _sub(el1,"Cell",{"N":"D","V":f"{cap*2:.4f}"})
    # Top ellipse
    geom2 = _sub(shape,"Geom",{"IX":"2"})
    mv2 = _sub(geom2,"MoveTo",{"IX":"1"})
    _sub(mv2,"Cell",{"N":"X","V":f"{w/2:.4f}"})
    _sub(mv2,"Cell",{"N":"Y","V":f"{h-cap*2:.4f}"})
    el2 = _sub(geom2,"Ellipse",{"IX":"2"})
    _sub(el2,"Cell",{"N":"X","V":f"{w/2:.4f}"})
    _sub(el2,"Cell",{"N":"Y","V":f"{h-cap:.4f}"})
    _sub(el2,"Cell",{"N":"A","V":f"{w:.4f}"})
    _sub(el2,"Cell",{"N":"B","V":f"{h-cap:.4f}"})
    _sub(el2,"Cell",{"N":"C","V":f"{w/2:.4f}"})
    _sub(el2,"Cell",{"N":"D","V":f"{h:.4f}"})
    txt = _sub(shape,"Text")
    txt.text = _wrap(label,16)
    return shape


def _build_node_shape(sid, nid, ntype, label, x, y, w, h):
    st = SHAPE_STYLES.get(ntype, SHAPE_STYLES["process"])
    fill,line,font,bold = st["fill"],st["line"],st["font"],st["bold"]
    shape_type = st["shape"]
    if shape_type == "diamond":
        return _make_diamond_shape(sid,x,y,w,h,label,fill,line,font)
    elif shape_type == "ellipse":
        return _make_ellipse_shape(sid,x,y,w,h,label,fill,line,font,bold)
    elif shape_type == "cylinder":
        return _make_cylinder_shape(sid,x,y,w,h,label,fill,line,font)
    elif shape_type == "rounded_rect":
        return _make_rect_shape(sid,x,y,w,h,label,fill,line,font,bold,rounded=True)
    else:
        return _make_rect_shape(sid,x,y,w,h,label,fill,line,font,bold)


def _make_connector(cid, from_sid, to_sid, label, color, dashed=False):
    conn = _el("Connect",{
        "FromSheet":str(cid),
        "ToSheet":  str(to_sid),
        "FromCell": "EndX",
        "ToCell":   "PinX",
    })
    shape = _el("Shape",{
        "ID":str(cid),"Type":"Edge",
    })
    _sub(shape,"Cell",{"N":"LineColor",  "V":_rgb_to_visio(color)})
    _sub(shape,"Cell",{"N":"LineWeight", "V":"0.018"})
    if dashed:
        _sub(shape,"Cell",{"N":"LinePattern","V":"2"})
    if label:
        _sub(shape,"Cell",{"N":"CharSize","V":"0.10"})
        txt = _sub(shape,"Text")
        txt.text = label[:20]
    # BeginX/Y and EndX/Y will be set by Visio via Connect elements
    _sub(shape,"Cell",{"N":"BeginX","V":"0","F":f"Sheet.{from_sid}!PinX"})
    _sub(shape,"Cell",{"N":"BeginY","V":"0","F":f"Sheet.{from_sid}!PinY"})
    _sub(shape,"Cell",{"N":"EndX",  "V":"0","F":f"Sheet.{to_sid}!PinX"})
    _sub(shape,"Cell",{"N":"EndY",  "V":"0","F":f"Sheet.{to_sid}!PinY"})
    return shape, conn


def _make_phase_label(sid, ph_idx, label, color):
    zone_w = (PAGE_W - MARGIN*2) / 5
    x = MARGIN + ph_idx * zone_w
    shape = _el("Shape",{
        "ID":str(sid),"Type":"Shape",
        "Width":f"{zone_w:.4f}","Height":"0.5",
        "PinX":f"{x+zone_w/2:.4f}","PinY":f"{MARGIN+0.25:.4f}",
    })
    _sub(shape,"Cell",{"N":"FillForegnd","V":_rgb_to_visio(color)})
    _sub(shape,"Cell",{"N":"FillBkgnd",  "V":_rgb_to_visio(color)})
    _sub(shape,"Cell",{"N":"LineColor",  "V":_rgb_to_visio("#CCCCCC")})
    _sub(shape,"Cell",{"N":"CharColor",  "V":_rgb_to_visio("#1A3A5C")})
    _sub(shape,"Cell",{"N":"CharStyle",  "V":"1"})
    _sub(shape,"Cell",{"N":"CharSize",   "V":"0.14"})
    txt = _sub(shape,"Text")
    txt.text = label
    return shape


# ── Main VSDX builder ──────────────────────────────────────────────────────────
def _build_page_xml(nodes, edges, privacy_controls, show_controls, title):
    """Build the visio/pages/page1.xml content."""
    _shape_id[0] = 1000   # reset

    layout = _compute_layout(nodes)

    # Build node_id → shape_id mapping
    nid_to_sid = {}
    for nid in layout:
        nid_to_sid[nid] = _next_id()

    # Fuzzy control key resolution
    ctrl_resolved = {}
    if show_controls:
        nmap = {}
        for nid in layout:
            nmap[_norm(nid)] = nid
            parts=[p for p in _norm(nid).split("_") if len(p)>2]
            if parts: nmap[parts[0]] = nid
        for key,clist in privacy_controls.items():
            sid2 = _sid(key)
            if sid2 in layout: nid=sid2
            else:
                nk=_norm(key)
                nid=nmap.get(nk) or nmap.get(nk.split("_")[0] if "_" in nk else nk)
            if nid: ctrl_resolved[nid]=clist[:4]

    # Build node rank lookup for edge coloring
    node_rank={}
    for n in nodes:
        ph=n.get("phase","processing").lower()
        if ph not in PHASE_RANK: ph="processing"
        node_rank[_sid(n["id"])]=PHASE_RANK[ph]
    node_map={n.get("type","process"):n for n in nodes}
    node_type_map={_sid(n["id"]):n.get("type","process") for n in nodes}
    node_label_map={_sid(n["id"]):n.get("label","") for n in nodes}

    # Page XML
    page_root = _el("PageContents",{"xmlns":VISIO_NS})
    shapes_el = _sub(page_root,"Shapes")
    connects_el= _sub(page_root,"Connects")

    # Phase lane backgrounds + labels
    phases = [
        ("collection","① Collection","#EAF4FB"),
        ("processing","② Processing","#FEF9F0"),
        ("storage",   "③ Storage",   "#F0FAF0"),
        ("sharing",   "④ Sharing",   "#FFF8F0"),
        ("exit",      "⑤ Exit/Archive","#FDF0F0"),
    ]
    zone_w = (PAGE_W - MARGIN*2)/5
    for pi,(ph,lbl,bg) in enumerate(phases):
        # Lane background
        bg_sid = _next_id()
        bg_shape = _el("Shape",{
            "ID":str(bg_sid),"Type":"Shape",
            "Width":f"{zone_w:.4f}","Height":f"{PAGE_H:.4f}",
            "PinX":f"{MARGIN+pi*zone_w+zone_w/2:.4f}",
            "PinY":f"{PAGE_H/2:.4f}",
        })
        _sub(bg_shape,"Cell",{"N":"FillForegnd","V":_rgb_to_visio(bg)})
        _sub(bg_shape,"Cell",{"N":"FillBkgnd",  "V":_rgb_to_visio(bg)})
        _sub(bg_shape,"Cell",{"N":"LineColor",  "V":_rgb_to_visio("#DDDDDD")})
        _sub(bg_shape,"Cell",{"N":"LineWeight", "V":"0.01"})
        _sub(bg_shape,"Cell",{"N":"NoObjHandles","V":"1"})
        shapes_el.append(bg_shape)
        # Phase label
        lbl_sid = _next_id()
        shapes_el.append(_make_phase_label(lbl_sid, pi, lbl, bg))

    # Title text box
    t_sid = _next_id()
    t_shape = _el("Shape",{
        "ID":str(t_sid),"Type":"Shape",
        "Width":f"{PAGE_W-MARGIN*2:.4f}","Height":"0.6",
        "PinX":f"{PAGE_W/2:.4f}","PinY":f"{PAGE_H-0.5:.4f}",
    })
    _sub(t_shape,"Cell",{"N":"FillForegnd","V":_rgb_to_visio("#1A3A5C")})
    _sub(t_shape,"Cell",{"N":"FillBkgnd",  "V":_rgb_to_visio("#1A3A5C")})
    _sub(t_shape,"Cell",{"N":"LineColor",  "V":_rgb_to_visio("#1A3A5C")})
    _sub(t_shape,"Cell",{"N":"CharColor",  "V":_rgb_to_visio("#FFFFFF")})
    _sub(t_shape,"Cell",{"N":"CharStyle",  "V":"1"})
    _sub(t_shape,"Cell",{"N":"CharSize",   "V":"0.20"})
    txt_el = _sub(t_shape,"Text")
    txt_el.text = f"{title}  ·  Privacy & Data Protection Review  ·  DPDPA 2023 / GDPR"
    shapes_el.append(t_shape)

    # Node shapes
    for n in nodes:
        nid   = _sid(n["id"])
        ntype = n.get("type","process")
        label = n.get("label","")
        sid   = nid_to_sid[nid]
        x,y,w,h = layout[nid]
        shape = _build_node_shape(sid,nid,ntype,label,x,y,w,h)
        shapes_el.append(shape)

    # Privacy control shapes (below each node)
    ctrl_sid_map = {}
    if show_controls:
        for nid, ctrls in ctrl_resolved.items():
            if nid not in layout: continue
            x,y,w,h = layout[nid]
            start_y = y - (CTRL_H + CTRL_GAP) * len(ctrls) - 0.15
            ctrl_sid_map[nid] = []
            for ci, ctrl in enumerate(ctrls):
                csid = _next_id()
                ctrl_sid_map[nid].append(csid)
                cy_ = start_y + ci*(CTRL_H+CTRL_GAP)
                cx_ = x + (w - CTRL_W)/2
                c_shape = _make_rect_shape(
                    csid, cx_, cy_, CTRL_W, CTRL_H,
                    ctrl, "#D5E8D4","#27AE60","#145A32", False, rounded=True)
                shapes_el.append(c_shape)
                # Dashed connector from control to node
                conn_sid = _next_id()
                conn_shape = _el("Shape",{
                    "ID":str(conn_sid),"Type":"Edge",
                })
                _sub(conn_shape,"Cell",{"N":"LineColor",  "V":_rgb_to_visio("#27AE60")})
                _sub(conn_shape,"Cell",{"N":"LineWeight", "V":"0.012"})
                _sub(conn_shape,"Cell",{"N":"LinePattern","V":"2"})
                _sub(conn_shape,"Cell",{"N":"EndArrow",   "V":"0"})
                _sub(conn_shape,"Cell",{"N":"BeginX","V":"0","F":f"Sheet.{csid}!PinX"})
                _sub(conn_shape,"Cell",{"N":"BeginY","V":"0","F":f"Sheet.{csid}!PinY"})
                _sub(conn_shape,"Cell",{"N":"EndX",  "V":"0","F":f"Sheet.{nid_to_sid[nid]}!PinX"})
                _sub(conn_shape,"Cell",{"N":"EndY",  "V":"0","F":f"Sheet.{nid_to_sid[nid]}!PinY"})
                shapes_el.append(conn_shape)
                # Connect element
                c_conn = _el("Connect",{
                    "FromSheet":str(conn_sid),
                    "ToSheet":  str(nid_to_sid[nid]),
                    "FromCell": "EndX",
                    "ToCell":   "PinX",
                })
                connects_el.append(c_conn)

    # Edge connectors
    for e in edges:
        src = _sid(e.get("from",""))
        dst = _sid(e.get("to",""))
        if src not in nid_to_sid or dst not in nid_to_sid: continue
        raw  = e.get("label","")
        sens = any(k in raw.lower() for k in
                   ["health","medical","biometric","salary","financial","bank"])
        back = node_rank.get(src,2) >= node_rank.get(dst,2) and src!=dst
        color= "#C0392B" if sens else ("#AAAAAA" if back else "#888888")
        e_sid = _next_id()
        conn_shape = _el("Shape",{"ID":str(e_sid),"Type":"Edge"})
        _sub(conn_shape,"Cell",{"N":"LineColor",  "V":_rgb_to_visio(color)})
        _sub(conn_shape,"Cell",{"N":"LineWeight", "V":"0.022" if sens else "0.015"})
        if back: _sub(conn_shape,"Cell",{"N":"LinePattern","V":"2"})
        if raw:
            _sub(conn_shape,"Cell",{"N":"CharSize","V":"0.09"})
            t=_sub(conn_shape,"Text"); t.text=raw[:18]
        _sub(conn_shape,"Cell",{"N":"BeginX","V":"0","F":f"Sheet.{nid_to_sid[src]}!PinX"})
        _sub(conn_shape,"Cell",{"N":"BeginY","V":"0","F":f"Sheet.{nid_to_sid[src]}!PinY"})
        _sub(conn_shape,"Cell",{"N":"EndX",  "V":"0","F":f"Sheet.{nid_to_sid[dst]}!PinX"})
        _sub(conn_shape,"Cell",{"N":"EndY",  "V":"0","F":f"Sheet.{nid_to_sid[dst]}!PinY"})
        shapes_el.append(conn_shape)
        c1=_el("Connect",{"FromSheet":str(e_sid),"ToSheet":str(nid_to_sid[src]),
                           "FromCell":"BeginX","ToCell":"PinX"})
        c2=_el("Connect",{"FromSheet":str(e_sid),"ToSheet":str(nid_to_sid[dst]),
                           "FromCell":"EndX","ToCell":"PinX"})
        connects_el.append(c1); connects_el.append(c2)

    return ET.tostring(page_root, encoding="unicode", xml_declaration=False)


def generate_vsdx(dfd_data: dict) -> tuple:
    """
    Generate two .vsdx files: As-Is and Post Compliance.
    Returns (asis_bytes, future_bytes).
    """
    title   = dfd_data.get("process_name","Data Flow Diagram")
    asis    = dfd_data.get("asis",  {"nodes":[],"edges":[]})
    future  = dfd_data.get("future",{"nodes":[],"edges":[]})
    ctrls   = dfd_data.get("privacy_controls",{})

    def _make_vsdx(nodes, edges, privacy_controls, show_ctrls, state_label):
        page_xml = _build_page_xml(nodes, edges, privacy_controls, show_ctrls, title)

        # [Content_Types].xml
        content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/visio/document.xml"
    ContentType="application/vnd.ms-visio.drawing.main+xml"/>
  <Override PartName="/visio/pages/page1.xml"
    ContentType="application/vnd.ms-visio.page+xml"/>
  <Override PartName="/visio/pages/pages.xml"
    ContentType="application/vnd.ms-visio.pages+xml"/>
</Types>'''

        # _rels/.rels
        rels_root = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/visio/2010/relationships/document"
    Target="visio/document.xml"/>
</Relationships>'''

        # visio/_rels/document.xml.rels
        doc_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/visio/2010/relationships/pages"
    Target="pages/pages.xml"/>
</Relationships>'''

        # visio/document.xml
        doc_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<VisioDocument xmlns="http://schemas.microsoft.com/office/visio/2012/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <DocumentProperties>
    <Title>{title} — {state_label}</Title>
    <Creator>ROPA Intelligence Platform</Creator>
  </DocumentProperties>
</VisioDocument>'''

        # visio/pages/pages.xml
        pages_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Pages xmlns="http://schemas.microsoft.com/office/visio/2012/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <Page ID="1" Name="Data Flow Diagram">
    <PageSheet>
      <PageProps>
        <PageWidth>36</PageWidth>
        <PageHeight>14</PageHeight>
        <PageScale>1</PageScale>
        <DrawingScale>1</DrawingScale>
        <DrawingSizeType>1</DrawingSizeType>
        <DrawingScaleType>0</DrawingScaleType>
      </PageProps>
    </PageSheet>
    <Rel r:id="rId1"/>
  </Page>
</Pages>'''

        # visio/pages/_rels/pages.xml.rels
        pages_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/visio/2010/relationships/page"
    Target="page1.xml"/>
</Relationships>'''

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml",          content_types.strip())
            zf.writestr("_rels/.rels",                  rels_root.strip())
            zf.writestr("visio/document.xml",           doc_xml.strip())
            zf.writestr("visio/_rels/document.xml.rels",doc_rels.strip())
            zf.writestr("visio/pages/pages.xml",        pages_xml.strip())
            zf.writestr("visio/pages/_rels/pages.xml.rels", pages_rels.strip())
            zf.writestr("visio/pages/page1.xml",
                        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'+page_xml)
        buf.seek(0)
        return buf.getvalue()

    asis_vsdx   = _make_vsdx(asis["nodes"],   asis["edges"],   {},     False, "Current State")
    future_vsdx = _make_vsdx(future["nodes"], future["edges"], ctrls,  True,  "Post Compliance")
    return asis_vsdx, future_vsdx
