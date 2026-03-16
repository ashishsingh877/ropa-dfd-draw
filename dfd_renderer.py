"""
dfd_renderer.py — Professional DFD Renderer v15
Pure PIL drawing — no Graphviz layout engine.
Full pixel control. Clean landscape. Matches RateGain reference.
"""
import io, re, math, textwrap
from PIL import Image, ImageDraw, ImageFont

_FONT_REG  = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
_FONT_BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"

def _f(sz, bold=False):
    try:    return ImageFont.truetype(_FONT_BOLD if bold else _FONT_REG, sz)
    except: return ImageFont.load_default()

def _sid(s):
    return re.sub(r"[^a-zA-Z0-9_]","_",str(s).strip())[:35]

def _wrap(text, max_chars=16):
    return "\n".join(textwrap.wrap(str(text).strip(), width=max_chars)[:3])

# ── Colours ────────────────────────────────────────────────────────────────────
C = {
    "external_fill":  "#FFF5CC", "external_stroke": "#A07800", "external_text": "#5C4400",
    "team_fill":      "#FADADD", "team_stroke":     "#B03030", "team_text":     "#641E16",
    "process_fill":   "#FFFFFF", "process_stroke":  "#555555", "process_text":  "#1A1A1A",
    "decision_fill":  "#C0392B", "decision_stroke": "#8B2222", "decision_text": "#FFFFFF",
    "endpoint_fill":  "#B03030", "endpoint_stroke": "#7B1A1A", "endpoint_text": "#FFFFFF",
    "datastore_fill": "#D4E8FA", "datastore_stroke":"#1A5276", "datastore_text":"#0D3B6E",
    "ctrl_fill":      "#D5E8D4", "ctrl_stroke":     "#27AE60", "ctrl_text":     "#145A32",
    "edge_normal":    "#666666",
    "edge_sensitive": "#C0392B",
    "edge_back":      "#AAAAAA",
    "header_bg":      "#1A3A5C",
    "banner_asis":    "#C0392B",
    "banner_future":  "#1A6B3A",
    "bg":             "#FFFFFF",
    "legend_bg":      "#F4F6F8",
}

# ── Canvas constants ───────────────────────────────────────────────────────────
W_CANVAS  = 5400    # landscape width
HEADER_H  = 90
BANNER_H  = 50
LEG_H     = 50
CHART_PAD = 60      # padding inside chart area

# Node sizes
NODE_W    = 170
NODE_H    = 60
DECISION_W= 170     # diamond width
DECISION_H= 90
ENDPOINT_W= 150
ENDPOINT_H= 60
DS_W      = 160
DS_H      = 70

CTRL_W    = 165
CTRL_H    = 30
CTRL_GAP  = 8

ARROW_COLOR  = "#666666"
ARROW_HEAD   = 12


# ── Low-level drawing primitives ───────────────────────────────────────────────

def _text_in_box(draw, x, y, w, h, text, font, color):
    lines = _wrap(text, int(w/7)).split("\n")
    total_h = len(lines) * (font.size + 2)
    ty = y + (h - total_h) // 2
    for line in lines:
        try:
            bb = draw.textbbox((0,0), line, font=font)
            tw = bb[2]-bb[0]
        except:
            tw = len(line)*7
        draw.text((x + (w-tw)//2, ty), line, font=font, fill=color)
        ty += font.size + 3

def _draw_box(draw, x, y, w, h, fill, stroke, text, bold=False, font_size=13, stroke_w=2):
    draw.rounded_rectangle([x,y,x+w,y+h], radius=6, fill=fill, outline=stroke, width=stroke_w)
    font = _f(font_size, bold)
    _text_in_box(draw, x, y, w, h, text, font, C["process_text"] if fill=="#FFFFFF" else
                 (C["team_text"] if "FAD" in fill else (C["external_text"] if "FFF5" in fill else C["endpoint_text"])))

def _draw_external(draw, x, y, w, h, text):
    draw.rounded_rectangle([x,y,x+w,y+h], radius=10, fill=C["external_fill"], outline=C["external_stroke"], width=2)
    _text_in_box(draw, x, y, w, h, text, _f(12), C["external_text"])

def _draw_team(draw, x, y, w, h, text):
    draw.rectangle([x,y,x+w,y+h], fill=C["team_fill"], outline=C["team_stroke"], width=2)
    _text_in_box(draw, x, y, w, h, text, _f(13, bold=True), C["team_text"])

def _draw_process(draw, x, y, w, h, text):
    draw.rectangle([x,y,x+w,y+h], fill=C["process_fill"], outline=C["process_stroke"], width=1)
    _text_in_box(draw, x, y, w, h, text, _f(12), C["process_text"])

def _draw_decision(draw, x, y, w, h, text):
    # Diamond
    cx, cy = x+w//2, y+h//2
    pts = [(cx,y),(x+w,cy),(cx,y+h),(x,cy)]
    draw.polygon(pts, fill=C["decision_fill"], outline=C["decision_stroke"])
    draw.line([pts[0],pts[1]], fill=C["decision_stroke"], width=2)
    draw.line([pts[1],pts[2]], fill=C["decision_stroke"], width=2)
    draw.line([pts[2],pts[3]], fill=C["decision_stroke"], width=2)
    draw.line([pts[3],pts[0]], fill=C["decision_stroke"], width=2)
    _text_in_box(draw, x+10, y+10, w-20, h-20, text, _f(12, bold=True), C["decision_text"])

def _draw_endpoint(draw, x, y, w, h, text):
    draw.ellipse([x,y,x+w,y+h], fill=C["endpoint_fill"], outline=C["endpoint_stroke"], width=2)
    _text_in_box(draw, x, y, w, h, text, _f(12, bold=True), C["endpoint_text"])

def _draw_datastore(draw, x, y, w, h, text):
    # Cylinder (rectangle with ellipse caps)
    cap = 16
    draw.rectangle([x, y+cap//2, x+w, y+h-cap//2], fill=C["datastore_fill"], outline=C["datastore_fill"])
    draw.ellipse([x, y, x+w, y+cap], fill=C["datastore_fill"], outline=C["datastore_stroke"], width=1)
    draw.ellipse([x, y+h-cap, x+w, y+h], fill=C["datastore_fill"], outline=C["datastore_stroke"], width=1)
    draw.line([x, y+cap//2, x, y+h-cap//2], fill=C["datastore_stroke"], width=1)
    draw.line([x+w, y+cap//2, x+w, y+h-cap//2], fill=C["datastore_stroke"], width=1)
    _text_in_box(draw, x, y+cap//2, w, h-cap, text, _f(12), C["datastore_text"])

def _draw_ctrl_box(draw, x, y, text):
    draw.rounded_rectangle([x,y,x+CTRL_W,y+CTRL_H], radius=7,
                            fill=C["ctrl_fill"], outline=C["ctrl_stroke"], width=1)
    f = _f(9)
    try:
        bb = draw.textbbox((0,0), text[:26], font=f)
        tw,th = bb[2]-bb[0], bb[3]-bb[1]
    except:
        tw,th = len(text)*6, 9
    draw.text((x+(CTRL_W-tw)//2, y+(CTRL_H-th)//2), text[:26], font=f, fill=C["ctrl_text"])

def _arrow(draw, x1,y1, x2,y2, color="#666666", width=2, dashed=False, label=""):
    """Draw an arrow from (x1,y1) to (x2,y2) with optional label."""
    if dashed:
        # Draw dashed line
        dx,dy = x2-x1, y2-y1
        dist = max(1, math.sqrt(dx*dx+dy*dy))
        dash=12; gap=6; step=dash+gap
        n_steps = int(dist/step)
        for i in range(n_steps+1):
            t0 = i*step/dist; t1 = min((i*step+dash)/dist, 1.0)
            draw.line([(int(x1+dx*t0),int(y1+dy*t0)),(int(x1+dx*t1),int(y1+dy*t1))],
                      fill=color, width=width)
    else:
        draw.line([(x1,y1),(x2,y2)], fill=color, width=width)
    # Arrowhead
    dx,dy = x2-x1, y2-y1
    dist  = max(1, math.sqrt(dx*dx+dy*dy))
    dx,dy = dx/dist, dy/dist
    ax1 = int(x2 - ARROW_HEAD*dx - ARROW_HEAD*0.5*(-dy))
    ay1 = int(y2 - ARROW_HEAD*dy - ARROW_HEAD*0.5*dx)
    ax2 = int(x2 - ARROW_HEAD*dx + ARROW_HEAD*0.5*(-dy))
    ay2 = int(y2 - ARROW_HEAD*dy + ARROW_HEAD*0.5*dx)
    draw.polygon([(x2,y2),(ax1,ay1),(ax2,ay2)], fill=color)
    # Label
    if label:
        mx, my = (x1+x2)//2, (y1+y2)//2
        f = _f(9)
        try:
            bb = draw.textbbox((0,0),label,font=f); tw=bb[2]-bb[0]
        except: tw=len(label)*6
        # White bg behind label
        draw.rectangle([mx-tw//2-3, my-8, mx+tw//2+3, my+8], fill="#FFFFFF")
        draw.text((mx-tw//2, my-6), label, font=f, fill="#555555")

def _node_center(pos):
    ntype,x,y,w,h = pos
    return x+w//2, y+h//2

def _node_edge(pos, side):
    """Get connection point on edge: side = 'left','right','top','bottom'"""
    ntype,x,y,w,h = pos[0],pos[1],pos[2],pos[3],pos[4]
    cx,cy = x+w//2, y+h//2
    if side=="right":  return x+w, cy
    if side=="left":   return x, cy
    if side=="top":    return cx, y
    if side=="bottom": return cx, y+h
    return cx,cy


# ── Main render function ───────────────────────────────────────────────────────

def _draw_node(draw, pos):
    ntype,x,y,w,h = pos[:5]
    label = pos[5] if len(pos)>5 else ""
    if ntype=="external":  _draw_external(draw,x,y,w,h,label)
    elif ntype=="team":    _draw_team(draw,x,y,w,h,label)
    elif ntype=="process": _draw_process(draw,x,y,w,h,label)
    elif ntype=="decision":_draw_decision(draw,x,y,w,h,label)
    elif ntype=="endpoint":_draw_endpoint(draw,x,y,w,h,label)
    elif ntype=="datastore":_draw_datastore(draw,x,y,w,h,label)


def _compute_layout(nodes, canvas_w, canvas_h):
    """
    Compute pixel positions for all nodes in a landscape left-to-right layout.
    Phases: collection(left) | processing | storage | sharing | exit(right)
    """
    PHASE_ORDER = ["collection","processing","storage","sharing","exit","main"]
    PHASE_X = {}
    usable_w = canvas_w - CHART_PAD*2

    # Assign X zones
    zone_w = usable_w // 5
    for i,ph in enumerate(["collection","processing","storage","sharing","exit"]):
        PHASE_X[ph] = CHART_PAD + i*zone_w
    PHASE_X["main"] = PHASE_X["processing"]

    # Group by phase
    phase_map = {}
    for n in nodes:
        ph = n.get("phase","processing").lower()
        if ph not in PHASE_X: ph="processing"
        phase_map.setdefault(ph,[]).append(n)

    NODE_SIZES = {
        "external":  (NODE_W,   NODE_H),
        "team":      (NODE_W+10,NODE_H+10),
        "process":   (NODE_W,   NODE_H),
        "decision":  (DECISION_W,DECISION_H),
        "endpoint":  (ENDPOINT_W,ENDPOINT_H),
        "datastore": (DS_W,     DS_H),
    }

    layout = {}   # node_id → (type, x, y, w, h, label)

    for ph in PHASE_ORDER:
        ph_nodes = phase_map.get(ph,[])
        if not ph_nodes: continue
        n_nodes = len(ph_nodes)
        zone_x  = PHASE_X.get(ph, CHART_PAD)
        # Distribute vertically
        spacing = canvas_h // (n_nodes+1)
        for i,n in enumerate(ph_nodes):
            nid  = _sid(n["id"])
            ntype= n.get("type","process")
            lbl  = n.get("label","")
            w,h  = NODE_SIZES.get(ntype,(NODE_W,NODE_H))
            cy   = spacing*(i+1)
            cx   = zone_x + zone_w//2
            layout[nid] = (ntype, cx-w//2, cy-h//2, w, h, lbl)

    return layout, zone_w


def _render_diagram(nodes, edges, privacy_controls,
                    show_controls, canvas_w, canvas_h):
    """Render one diagram on a transparent canvas."""
    img  = Image.new("RGB", (canvas_w, canvas_h), "#FFFFFF")
    draw = ImageDraw.Draw(img)

    layout, zone_w = _compute_layout(nodes, canvas_w, canvas_h)

    # Fuzzy resolve controls
    def _norm(s): return re.sub(r"[^a-z0-9]","_",str(s).lower().strip())
    norm_map = {}
    for nid in layout:
        norm_map[_norm(nid)] = nid
        parts=[p for p in _norm(nid).split("_") if len(p)>2]
        if parts: norm_map[parts[0]] = nid

    resolved_ctrls = {}
    if show_controls:
        for key,clist in privacy_controls.items():
            sid = _sid(key)
            if sid in layout: nid=sid
            else:
                nk=_norm(key)
                nid=norm_map.get(nk) or norm_map.get(nk.split("_")[0] if "_" in nk else nk)
            if nid: resolved_ctrls[nid]=clist[:4]

    # Draw phase lane headers
    PHASE_ORDER = ["collection","processing","storage","sharing","exit"]
    PHASE_LABELS = ["① Collection","② Processing","③ Storage","④ Sharing","⑤ Exit / Archive"]
    PHASE_COLORS = ["#EAF4FB","#FEF9F0","#F0FAF0","#FFF8F0","#FDF0F0"]
    usable_w = canvas_w - CHART_PAD*2
    zone_w2  = usable_w // 5
    for i,(ph,lbl,bg) in enumerate(zip(PHASE_ORDER,PHASE_LABELS,PHASE_COLORS)):
        lx = CHART_PAD + i*zone_w2
        draw.rectangle([lx,0,lx+zone_w2,canvas_h], fill=bg)
        draw.line([(lx,0),(lx,canvas_h)], fill="#DDDDDD", width=1)
        f = _f(11, bold=True)
        try:
            bb=draw.textbbox((0,0),lbl,font=f); tw=bb[2]-bb[0]
        except: tw=len(lbl)*8
        draw.text((lx+(zone_w2-tw)//2, 10), lbl, font=f, fill="#1A3A5C")
    draw.line([(CHART_PAD+5*zone_w2,0),(CHART_PAD+5*zone_w2,canvas_h)], fill="#DDDDDD", width=1)

    # Draw edges
    node_rank = {}
    for n in nodes:
        ph=n.get("phase","processing").lower()
        pr={"collection":0,"processing":1,"storage":2,"sharing":3,"exit":4,"main":2}
        node_rank[_sid(n["id"])] = pr.get(ph,2)

    for e in edges:
        src=_sid(e.get("from","")); dst=_sid(e.get("to",""))
        if src not in layout or dst not in layout: continue
        sp=layout[src]; dp=layout[dst]
        lbl=e.get("label","")
        raw=lbl.lower()
        sensitive=any(k in raw for k in ["health","medical","biometric","salary","financial","bank","sensitive","aadhaar","pan"])
        is_back = node_rank.get(src,2) >= node_rank.get(dst,2) and src!=dst

        color  = C["edge_sensitive"] if sensitive else (C["edge_back"] if is_back else C["edge_normal"])
        width_ = 3 if sensitive else (1 if is_back else 2)

        # Connection points
        sx,sy = _node_edge(sp,"right")
        dx2,dy2 = _node_edge(dp,"left")
        if is_back:
            # Route back edge: go down, then back left
            sx,sy  = _node_edge(sp,"bottom")
            dx2,dy2= _node_edge(dp,"bottom")
            mid_y  = max(sy,dy2) + 40
            pts = [(sx,sy),(sx,mid_y),(dx2,mid_y),(dx2,dy2)]
            for i in range(len(pts)-1):
                _arrow(draw,pts[i][0],pts[i][1],pts[i+1][0],pts[i+1][1],
                       color=color,width=width_,dashed=True)
            # Label in middle
            if lbl:
                mx=(sx+dx2)//2; my=mid_y+6
                f=_f(9)
                try: bb=draw.textbbox((0,0),lbl,font=f); tw=bb[2]-bb[0]
                except: tw=len(lbl)*6
                draw.rectangle([mx-tw//2-3,my-2,mx+tw//2+3,my+12],fill="#FFFFFF")
                draw.text((mx-tw//2,my),lbl,font=f,fill="#888888")
        else:
            _arrow(draw,sx,sy,dx2,dy2,color=color,width=width_,
                   dashed=False,label=lbl)

    # Draw nodes (on top of edges)
    for nid,pos in layout.items():
        _draw_node(draw,pos)

    # Draw privacy controls (below each node)
    if show_controls:
        for nid,ctrls in resolved_ctrls.items():
            if nid not in layout: continue
            ntype,nx,ny,nw,nh,_ = layout[nid]
            # Controls start below node
            start_y = ny + nh + 14
            # Place controls in 2 columns
            cols=2
            for ci,ctrl in enumerate(ctrls[:4]):
                col=ci%cols; row=ci//cols
                cx_ = nx + col*(CTRL_W+6) - (CTRL_W*cols + 6*(cols-1) - nw)//2
                cy_ = start_y + row*(CTRL_H+CTRL_GAP)
                _draw_ctrl_box(draw, cx_, cy_, ctrl)
            # Dashed connector from node bottom to first control
            fx = nx+nw//2; fy = ny+nh
            lx = nx+nw//2; ly = start_y-2
            draw.line([(fx,fy),(lx,ly)], fill="#27AE60", width=1)

    return img


def _make_page(diagram_img: Image.Image, title: str, state: str,
               banner_txt: str, banner_color: str) -> Image.Image:
    """Wrap diagram in professional header/banner/legend."""
    DW, DH = diagram_img.size
    W  = DW + 80
    H  = HEADER_H + BANNER_H + DH + LEG_H + 20

    cv = Image.new("RGB",(W,H),"#FFFFFF")
    dr = ImageDraw.Draw(cv)

    # Header
    dr.rectangle([0,0,W,HEADER_H], fill=C["header_bg"])
    dr.rectangle([14,12,106,HEADER_H-12], fill="#2470A0", outline="#154C80", width=2)
    dr.text((20,18),"DATA\nFLOW\nANALYSIS", font=_f(10,True), fill="#FFFFFF")
    dr.text((118,10), title, font=_f(28,True), fill="#FFFFFF")
    dr.text((119,50),"Privacy & Data Protection Review  ·  DPDPA 2023 / GDPR",
            font=_f(13), fill="#93C6E7")
    bc="#C0392B" if state=="asis" else "#1A6B3A"
    bt="CURRENT STATE" if state=="asis" else "POST COMPLIANCE"
    dr.rounded_rectangle([W-295,16,W-14,HEADER_H-16], radius=6, fill=bc)
    dr.text((W-280,30), bt, font=_f(14,True), fill="#FFFFFF")

    # Banner
    dr.rectangle([0,HEADER_H,W,HEADER_H+BANNER_H], fill=banner_color)
    dr.text((40,HEADER_H+14),"◼  "+banner_txt, font=_f(17,True), fill="#FFFFFF")

    # Diagram
    cv.paste(diagram_img, (40, HEADER_H+BANNER_H))

    # Legend
    ly = HEADER_H+BANNER_H+DH+5
    dr.rectangle([0,ly,W,ly+LEG_H], fill=C["legend_bg"])
    dr.line([(0,ly),(W,ly)], fill="#CCCCCC", width=1)
    items=[
        ("#FFF5CC","#A07800","External Entity / Data Subject"),
        ("#FADADD","#B03030","Internal Team / Department"),
        ("#FFFFFF","#555555","Process Step"),
        ("#C0392B","#8B2222","Decision Gate / Endpoint"),
        ("#D4E8FA","#1A5276","Data Store / System"),
        ("#D5E8D4","#27AE60","Privacy Control"),
    ]
    lx=40
    for fc,sc2,lb in items:
        dr.rounded_rectangle([lx,ly+12,lx+18,ly+30],radius=3,fill=fc,outline=sc2,width=1)
        dr.text((lx+24,ly+13),lb, font=_f(11), fill="#444444")
        lx+=210
    dr.text((40,ly+LEG_H-14),
            "⚠  Red arrows = sensitive data (financial, health, biometric)  ·  Dashed = feedback / back-flow",
            font=_f(9), fill="#C0392B")
    return cv


def render_dfd(dfd_data: dict) -> tuple:
    title  = dfd_data.get("process_name","Data Flow Diagram")
    asis   = dfd_data.get("asis",  {"nodes":[],"edges":[]})
    future = dfd_data.get("future",{"nodes":[],"edges":[]})
    ctrls  = dfd_data.get("privacy_controls",{})

    # Determine canvas height based on max nodes in any phase
    def _max_nodes(data):
        pm={}
        for n in data.get("nodes",[]):
            ph=n.get("phase","processing").lower()
            pm[ph]=pm.get(ph,0)+1
        return max(pm.values()) if pm else 3

    def _chart_h(data, with_ctrls):
        mn = _max_nodes(data)
        base = max(900, mn*180 + 120)
        if with_ctrls: base += 260   # extra space for control boxes
        return base

    asis_h   = _chart_h(asis, False)
    future_h = _chart_h(future, True)

    asis_diag   = _render_diagram(asis["nodes"],   asis["edges"],   {}, False, W_CANVAS-80, asis_h)
    future_diag = _render_diagram(future["nodes"], future["edges"], ctrls, True, W_CANVAS-80, future_h)

    asis_page   = _make_page(asis_diag,   title, "asis",
                             "Current State  ·  Existing Data Flows (Without Privacy Controls)",
                             C["banner_asis"])
    future_page = _make_page(future_diag, title, "future",
                             "Post Compliance  ·  Privacy-Embedded Future State",
                             C["banner_future"])

    def _to_png(img):
        b=io.BytesIO(); img.save(b,format="PNG",dpi=(250,250)); return b.getvalue()
    def _to_pdf(img):
        b=io.BytesIO(); img.save(b,format="PDF",resolution=250); return b.getvalue()

    return _to_png(asis_page),_to_pdf(asis_page),_to_png(future_page),_to_pdf(future_page)


DFD_JSON_SCHEMA='''
Return ONLY a valid JSON array with exactly ONE element. No markdown. No text before or after.

[{"id":"P001","process_name":"Name ≤50 chars",
  "asis":{"nodes":[...],"edges":[...]},
  "future":{"nodes":[...],"edges":[...]},
  "privacy_controls":{"node_id":["Control 1","Control 2","Control 3","Control 4"]},
  "narrative":"3-5 sentences."}]

NODE: {"id":"snake_id","label":"≤14 chars","type":"external|team|process|decision|endpoint|datastore","phase":"collection|processing|storage|sharing|exit"}
EDGE: {"from":"id","to":"id","label":"≤12 chars"}

PHASES (strict left-to-right):
  collection = data SOURCES (portals, email, agencies, external entities)
  processing = sequential STEPS left→right (team review, interview, BGV, decision gate)
  storage    = STORAGE SYSTEMS (HRMS, Email System, database)
  sharing    = EXTERNAL RECIPIENTS (vendors, banks, insurance, regulators)
  exit       = FINAL STATES (Hired, Rejected, Archived, Offboarded)

CRITICAL: privacy_controls keys must EXACTLY match node IDs (snake_case).
Controls: 3-4 per node, ≤20 chars. Min 10 nodes + 10 edges.
future nodes = same IDs as asis. All edge IDs must exist.
'''
