"""
dfd_renderer.py — Professional DFD Renderer FINAL
Based on the working layout: sources rank=same | main chain | storage rank=same | recipients rank=same | exit rank=same
Privacy controls overlaid with PIL at exact node positions from Graphviz JSON.
250 DPI, print-quality, matches RateGain reference style.
"""
import io, re, json, math, textwrap, graphviz
from PIL import Image, ImageDraw, ImageFont

_FONT_REG  = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
_FONT_BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"

def _f(sz, bold=False):
    try:    return ImageFont.truetype(_FONT_BOLD if bold else _FONT_REG, sz)
    except: return ImageFont.load_default()

def _gv(text, n=15):
    lines = textwrap.wrap(str(text).strip(), width=n)
    return "\\n".join(lines[:3]) if lines else str(text)

def _sid(s):
    return re.sub(r"[^a-zA-Z0-9_]","_",str(s).strip())[:35]

# ── Node styles ────────────────────────────────────────────────────────────────
PHASE_RANK = {"collection":0,"processing":1,"storage":2,"sharing":3,"exit":4,"main":2}

def _node_gv(ntype, label):
    label = _gv(label, 14)
    base  = dict(fontname="Helvetica", fontsize="12",
                 margin="0.22,0.14", penwidth="1.6")
    styles = {
        "external":  dict(shape="box",     style="filled,rounded",
                          fillcolor="#FFF5CC",color="#A07800",fontcolor="#5C4400",
                          width="1.7",height="0.6"),
        "team":      dict(shape="box",     style="filled",
                          fillcolor="#FADADD",color="#B03030",fontcolor="#641E16",
                          fontname="Helvetica-Bold",width="1.7",height="0.6",penwidth="2.0"),
        "process":   dict(shape="box",     style="filled",
                          fillcolor="#FFFFFF",color="#555555",fontcolor="#1A1A1A",
                          width="1.7",height="0.6"),
        "decision":  dict(shape="diamond", style="filled",
                          fillcolor="#C0392B",color="#8B2222",fontcolor="#FFFFFF",
                          fontname="Helvetica-Bold",fontsize="11",
                          width="1.7",height="0.75",penwidth="2.0"),
        "endpoint":  dict(shape="ellipse", style="filled",
                          fillcolor="#B03030",color="#7B1A1A",fontcolor="#FFFFFF",
                          fontname="Helvetica-Bold",width="1.6",height="0.6",penwidth="2.0"),
        "datastore": dict(shape="cylinder",style="filled",
                          fillcolor="#D4E8FA",color="#1A5276",fontcolor="#0D3B6E",
                          width="1.7",height="0.7"),
    }
    s = styles.get(ntype, styles["process"]).copy()
    s.update(base)
    return label, s

GV_DPI = 180

def _build_dot(data: dict) -> graphviz.Digraph:
    nodes = data.get("nodes",[])
    edges = data.get("edges",[])

    # Group by phase
    phase_map: dict = {}
    for n in nodes:
        ph = n.get("phase","processing").lower()
        if ph not in PHASE_RANK: ph="processing"
        phase_map.setdefault(ph,[]).append(n)

    # Node rank lookup
    node_rank = {}
    for n in nodes:
        ph=n.get("phase","processing").lower()
        if ph not in PHASE_RANK: ph="processing"
        node_rank[_sid(n["id"])] = PHASE_RANK[ph]

    dot = graphviz.Digraph(engine="dot")
    dot.attr("graph",
        rankdir  = "LR",
        splines  = "ortho",
        nodesep  = "0.65",
        ranksep  = "1.90",
        pad      = "0.55",
        bgcolor  = "white",
        size     = "28,10!",
        ratio    = "fill",
        dpi      = str(GV_DPI),
        fontname = "Helvetica",
    )
    dot.attr("edge",
        fontname ="Helvetica", fontsize="8",
        color    ="#888888",   fontcolor="#666666",
        arrowsize="0.85",     penwidth ="1.3",
    )

    # Add all nodes; parallel phases use rank=same
    for ph, ph_nodes in phase_map.items():
        if ph in ("collection","sharing","exit") and len(ph_nodes) >= 1:
            with dot.subgraph() as sg:
                sg.attr(rank="same")
                for n in ph_nodes:
                    lbl, attrs = _node_gv(n.get("type","process"), n.get("label",""))
                    sg.node(_sid(n["id"]), label=lbl, **attrs)
        elif ph == "storage":
            with dot.subgraph() as sg:
                sg.attr(rank="same")
                for n in ph_nodes:
                    lbl, attrs = _node_gv(n.get("type","process"), n.get("label",""))
                    sg.node(_sid(n["id"]), label=lbl, **attrs)
        else:
            # processing / main — add individually, let dot rank them naturally
            for n in ph_nodes:
                lbl, attrs = _node_gv(n.get("type","process"), n.get("label",""))
                dot.node(_sid(n["id"]), label=lbl, **attrs)

    # Edges
    for e in edges:
        s = _sid(e.get("from",""))
        t = _sid(e.get("to",""))
        if not s or not t: continue
        raw = e.get("label","").strip()
        lbl = (raw[:14]+"…") if len(raw)>15 else raw
        sensitive = any(k in raw.lower() for k in
            ["health","medical","biometric","salary","financial","bank","sensitive","aadhaar","pan"])
        sr = node_rank.get(s,2); dr_ = node_rank.get(t,2)
        is_back = sr >= dr_ and s != t
        attrs = dict(
            color    = "#C0392B" if sensitive else ("#AAAAAA" if is_back else "#888888"),
            fontcolor= "#666666",
            penwidth = "2.2"     if sensitive else ("1.0" if is_back else "1.3"),
        )
        if is_back:
            attrs["constraint"] = "false"
            attrs["style"]      = "dashed"
        if lbl:
            attrs["xlabel"]   = "  "+lbl+"  "
            attrs["fontsize"] = "8"
        dot.edge(s, t, **attrs)

    return dot


def _get_node_positions(dot: graphviz.Digraph, img_w: int, img_h: int) -> dict:
    """Get precise pixel positions of each node from Graphviz JSON output."""
    try:
        raw = dot.pipe(format="json")
        gv  = json.loads(raw)
    except Exception:
        return {}
    bb  = [float(x) for x in gv.get("bb","0,0,100,100").split(",")]
    if bb[2]==0 or bb[3]==0: return {}
    sx, sy = img_w/bb[2], img_h/bb[3]
    pos = {}
    for obj in gv.get("objects",[]):
        name = obj.get("name","")
        if not name: continue
        ps = obj.get("pos","")
        if not ps or "," not in ps: continue
        try: gx,gy = [float(v) for v in ps.split(",")]
        except: continue
        wi = float(obj.get("width",1.0))
        hi = float(obj.get("height",0.5))
        cx = gx*sx;  cy = (bb[3]-gy)*sy
        pw = wi*72*sx; ph = hi*72*sy
        pos[name] = dict(cx=cx,cy=cy,w=pw,h=ph,
                         x1=cx-pw/2,y1=cy-ph/2,
                         x2=cx+pw/2,y2=cy+ph/2)
    return pos


def _overlay_privacy_controls(flow_img: Image.Image, positions: dict,
                               privacy_controls: dict, nodes: list) -> Image.Image:
    """Overlay green privacy control boxes directly on diagram near each node."""
    img  = flow_img.copy()
    draw = ImageDraw.Draw(img)
    W, H = img.size
    sc   = W / 4800

    BW   = max(155, int(200*sc))
    BH   = max(28,  int(34*sc))
    GX   = max(7,   int(10*sc))
    GY   = max(5,   int(7*sc))
    MAR  = max(18,  int(24*sc))
    FS   = max(9,   int(11*sc))
    PR   = max(5,   int(7*sc))
    LW   = max(1,   int(2*sc))
    COLS = 2

    # Fuzzy key resolution
    def _norm(s): return re.sub(r"[^a-z0-9]","_",str(s).lower().strip())
    nmap = {}
    for pk in positions:
        nmap[_norm(pk)] = pk
        parts=[p for p in _norm(pk).split("_") if len(p)>2]
        if parts: nmap[parts[0]] = pk

    for raw_key, controls in privacy_controls.items():
        if not controls: continue
        sid = _sid(raw_key)
        pk  = sid if sid in positions else nmap.get(_norm(raw_key)) or nmap.get(_norm(raw_key).split("_")[0])
        if not pk: continue

        p     = positions[pk]
        pills = controls[:4]
        nr    = math.ceil(len(pills)/COLS)
        nc    = min(COLS,len(pills))
        BLW   = nc*(BW+GX)-GX
        BLH   = nr*(BH+GY)-GY

        # Place above if space, else below
        if p["y1"] > BLH + MAR + 4:
            bx = p["cx"] - BLW/2
            by = p["y1"] - BLH - MAR
            # Connector
            draw.line([(int(p["cx"]),int(p["y1"])),
                       (int(p["cx"]),int(by+BLH+2))],
                      fill="#27AE60", width=LW)
        else:
            bx = p["cx"] - BLW/2
            by = p["y2"] + MAR
            draw.line([(int(p["cx"]),int(p["y2"])),
                       (int(p["cx"]),int(by-2))],
                      fill="#27AE60", width=LW)

        bx = max(2, min(float(bx), W-BLW-2))
        by = max(2, min(float(by), H-BLH-2))

        for i, ctrl in enumerate(pills):
            r=i//COLS; c=i%COLS
            px=bx+c*(BW+GX); py=by+r*(BH+GY)
            draw.rounded_rectangle(
                [int(px),int(py),int(px+BW),int(py+BH)],
                radius=PR, fill="#D5E8D4", outline="#2E8B57", width=1)
            text=ctrl[:24]; f=_f(FS)
            try:
                bb2=draw.textbbox((0,0),text,font=f)
                tw,th=bb2[2]-bb2[0],bb2[3]-bb2[1]
            except: tw,th=len(text)*6,FS
            draw.text((int(px+(BW-tw)/2),int(py+(BH-th)/2)),text,font=f,fill="#145A32")
    return img


# ── PIL page composition ───────────────────────────────────────────────────────
HEADER_H=88; BANNER_H=52; LEG_H=52; PAD=44
TOP_PAD=280; BOT_PAD=60

def _compose_page(flow_png: bytes, title: str, state: str,
                  banner_txt: str, banner_color: str,
                  privacy_controls: dict = None,
                  nodes: list = None,
                  dot_obj: graphviz.Digraph = None) -> Image.Image:

    # Load flow image
    g = Image.open(io.BytesIO(flow_png)).convert("RGB")

    # Scale to minimum width for crispness
    MIN_W = 3600
    if g.width < MIN_W:
        sc = MIN_W/g.width
        g  = g.resize((int(g.width*sc), int(g.height*sc)), Image.LANCZOS)
    gw, gh = g.size

    # For post-compliance: add top padding and overlay controls
    if state == "future" and privacy_controls and nodes and dot_obj:
        # Add white space above and below for control boxes
        padded = Image.new("RGB",(gw, gh+TOP_PAD+BOT_PAD),"#FFFFFF")
        padded.paste(g, (0, TOP_PAD))
        # Get positions in original image space then offset
        raw_pos = _get_node_positions(dot_obj, gw, gh)
        scaled_pos = {}
        for k,v in raw_pos.items():
            scaled_pos[k] = dict(
                cx=v["cx"], cy=v["cy"]+TOP_PAD,
                w=v["w"],   h=v["h"],
                x1=v["x1"],  y1=v["y1"]+TOP_PAD,
                x2=v["x2"],  y2=v["y2"]+TOP_PAD,
            )
        g = _overlay_privacy_controls(padded, scaled_pos, privacy_controls, nodes)
        gw, gh = g.size

    W = gw + PAD*2
    H = HEADER_H + BANNER_H + gh + LEG_H + 18
    cv = Image.new("RGB",(W,H),"#FFFFFF")
    dr = ImageDraw.Draw(cv)

    # ── Header ────────────────────────────────────────────────────────────────
    dr.rectangle([0,0,W,HEADER_H], fill="#1A3A5C")
    dr.rectangle([14,12,108,HEADER_H-12], fill="#2470A0", outline="#154C80", width=2)
    dr.text((21,18),"DATA\nFLOW\nANALYSIS", font=_f(10,True), fill="#FFFFFF")
    dr.text((120,10), title, font=_f(28,True), fill="#FFFFFF")
    dr.text((121,50),"Privacy & Data Protection Review  ·  DPDPA 2023 / GDPR",
            font=_f(13), fill="#93C6E7")
    bc="#C0392B" if state=="asis" else "#1A6B3A"
    bt="CURRENT STATE" if state=="asis" else "POST COMPLIANCE"
    dr.rounded_rectangle([W-295,16,W-14,HEADER_H-16], radius=6, fill=bc)
    dr.text((W-280,30), bt, font=_f(14,True), fill="#FFFFFF")

    # ── Banner ────────────────────────────────────────────────────────────────
    dr.rectangle([0,HEADER_H,W,HEADER_H+BANNER_H], fill=banner_color)
    dr.text((PAD,HEADER_H+14),"◼  "+banner_txt, font=_f(17,True), fill="#FFFFFF")

    # ── Flow diagram ──────────────────────────────────────────────────────────
    cv.paste(g, (PAD, HEADER_H+BANNER_H))

    # ── Legend ────────────────────────────────────────────────────────────────
    ly = HEADER_H+BANNER_H+gh+5
    dr.rectangle([0,ly,W,ly+LEG_H], fill="#F4F6F8")
    dr.line([(0,ly),(W,ly)], fill="#CCCCCC", width=1)
    items=[
        ("#FFF5CC","#A07800","External Entity / Data Subject"),
        ("#FADADD","#B03030","Internal Team / Department"),
        ("#FFFFFF","#555555","Process Step"),
        ("#C0392B","#8B2222","Decision Gate / Endpoint"),
        ("#D4E8FA","#1A5276","Data Store / System"),
        ("#D5E8D4","#2E8B57","Privacy Control"),
    ]
    lx=PAD
    for fc,sc2,lb in items:
        dr.rounded_rectangle([lx,ly+12,lx+20,ly+30],radius=3,fill=fc,outline=sc2,width=1)
        dr.text((lx+26,ly+13),lb, font=_f(11), fill="#444444")
        lx+=205
    dr.text((PAD,ly+LEG_H-14),
            "⚠  Red arrows = sensitive data flows (financial, health, biometric)  "
            "·  Dashed grey = back-flow / result feedback",
            font=_f(9), fill="#C0392B")
    return cv


# ── Public API ─────────────────────────────────────────────────────────────────
def render_dfd(dfd_data: dict) -> tuple:
    title  = dfd_data.get("process_name","Data Flow Diagram")
    asis   = dfd_data.get("asis",  {"nodes":[],"edges":[]})
    future = dfd_data.get("future",{"nodes":[],"edges":[]})
    ctrls  = dfd_data.get("privacy_controls",{})

    def _to_png(img):
        b=io.BytesIO(); img.save(b,format="PNG",dpi=(250,250)); return b.getvalue()
    def _to_pdf(img):
        b=io.BytesIO(); img.save(b,format="PDF",resolution=250); return b.getvalue()

    # ── Current State ─────────────────────────────────────────────────────────
    dot_a    = _build_dot(asis)
    png_a    = dot_a.pipe(format="png")
    asis_img = _compose_page(png_a, title, "asis",
                             "Current State  ·  Existing Data Flows (Without Privacy Controls)",
                             "#C0392B")

    # ── Post Compliance ────────────────────────────────────────────────────────
    dot_f   = _build_dot(future)
    png_f   = dot_f.pipe(format="png")
    fut_img = _compose_page(png_f, title, "future",
                            "Post Compliance  ·  Privacy-Embedded Future State",
                            "#1A6B3A",
                            privacy_controls=ctrls,
                            nodes=future.get("nodes",[]),
                            dot_obj=dot_f)

    return _to_png(asis_img),_to_pdf(asis_img),_to_png(fut_img),_to_pdf(fut_img)


DFD_JSON_SCHEMA='''
Return ONLY a valid JSON array with exactly ONE element. No markdown. No text.

[{"id":"P001","process_name":"Name ≤50 chars",
  "asis":{"nodes":[...],"edges":[...]},
  "future":{"nodes":[...],"edges":[...]},
  "privacy_controls":{"node_id":["Control 1","Control 2","Control 3","Control 4"]},
  "narrative":"3-5 sentences."}]

NODE: {"id":"snake_id","label":"≤14 chars","type":"external|team|process|decision|endpoint|datastore","phase":"collection|processing|storage|sharing|exit"}
EDGE: {"from":"id","to":"id","label":"≤12 chars"}

LAYOUT RULES (critical for clean output):
  collection = parallel DATA SOURCES on the left (portals, email, agencies)
  processing = SEQUENTIAL steps in a chain left→right (team → review → BGV → decision)
  storage    = data STORES (HRMS, email system) — max 3 nodes
  sharing    = external RECIPIENTS on the right (vendors, banks, insurance)
  exit       = FINAL STATES (Hired, Rejected, Archived)

PRIVACY CONTROLS: keys MUST exactly match node IDs. 3-4 per node, ≤20 chars each.
These appear as green boxes above/below each node.
Min 10 nodes + 10 edges. future node IDs = same as asis.
'''
