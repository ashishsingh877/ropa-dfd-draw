"""
drawio_export.py
Generate draw.io XML from DFD data.
User can paste this into https://app.diagrams.net to get a professional editable DFD.
"""

import re, html, math

COLORS = {
    "external":  dict(fillColor="#FFF5CC", strokeColor="#A07800", fontColor="#5C4400", rounded="1"),
    "team":      dict(fillColor="#FADADD", strokeColor="#B03030", fontColor="#641E16", rounded="0", style_extra="fontStyle=1"),
    "process":   dict(fillColor="#FFFFFF", strokeColor="#555555", fontColor="#1A1A1A", rounded="0"),
    "decision":  dict(fillColor="#C0392B", strokeColor="#8B2222", fontColor="#FFFFFF", rounded="0", shape="rhombus", style_extra="fontStyle=1"),
    "endpoint":  dict(fillColor="#B03030", strokeColor="#7B1A1A", fontColor="#FFFFFF", rounded="0", shape="ellipse", style_extra="fontStyle=1"),
    "datastore": dict(fillColor="#D4E8FA", strokeColor="#1A5276", fontColor="#0D3B6E", rounded="0", shape="mxgraph.flowchart.stored_data"),
    "privacy":   dict(fillColor="#D5E8D4", strokeColor="#27AE60", fontColor="#145A32", rounded="1"),
}

PHASE_X = {"collection":80, "processing":380, "storage":700, "sharing":1020, "exit":1320, "main":380}
PHASE_SPACING = 280

NODE_W = 140; NODE_H = 50
CTRL_W = 130; CTRL_H = 28

def _sid(s):
    return re.sub(r"[^a-zA-Z0-9_]","_",str(s).strip())[:35]

def _xml_escape(s):
    return html.escape(str(s))

def generate_drawio_xml(dfd_data: dict, state: str = "future") -> str:
    """Generate draw.io XML for a DFD. state='asis' or 'future'."""
    data = dfd_data.get("future" if state=="future" else "asis", {"nodes":[],"edges":[]})
    ctrls = dfd_data.get("privacy_controls", {}) if state=="future" else {}
    title = dfd_data.get("process_name","Data Flow Diagram")
    nodes = data.get("nodes",[])
    edges = data.get("edges",[])

    # Layout: assign positions
    phase_map = {}
    for n in nodes:
        ph = n.get("phase","processing").lower()
        phase_map.setdefault(ph,[]).append(n)

    PHASE_ORDER_X = {"collection":80,"processing":360,"storage":660,"sharing":960,"exit":1260,"main":360}
    PHASE_VERTICAL_GAP = 80

    node_positions = {}
    for ph, ph_nodes in phase_map.items():
        x = PHASE_ORDER_X.get(ph,360)
        total_h = len(ph_nodes)*NODE_H + (len(ph_nodes)-1)*PHASE_VERTICAL_GAP
        start_y = max(80, 400 - total_h//2)
        for i,n in enumerate(ph_nodes):
            nid = _sid(n["id"])
            y   = start_y + i*(NODE_H + PHASE_VERTICAL_GAP)
            node_positions[nid] = (x, y)

    uid = [1]
    def new_id():
        uid[0]+=1; return f"n{uid[0]}"

    cells = []

    # Title label
    cells.append(f'''    <mxCell id="title" value="{_xml_escape(title)}" style="text;html=1;strokeColor=none;fillColor=#1A3A5C;fontColor=#FFFFFF;align=center;verticalAlign=middle;whiteSpace=wrap;fontSize=16;fontStyle=1;rounded=1;" vertex="1" parent="1">
      <mxGeometry x="80" y="10" width="1260" height="48" as="geometry" />
    </mxCell>''')

    # Phase labels
    for ph, x in PHASE_ORDER_X.items():
        if ph not in phase_map: continue
        ph_lbl = {"collection":"① Collection","processing":"② Processing",
                  "storage":"③ Storage","sharing":"④ Sharing",
                  "exit":"⑤ Exit/Archive","main":"Processing"}.get(ph,ph.title())
        cells.append(f'''    <mxCell id="ph_{ph}" value="{ph_lbl}" style="text;html=1;strokeColor=#CCCCCC;fillColor=#F8F9FA;fontColor=#1A3A5C;fontStyle=1;fontSize=10;align=center;" vertex="1" parent="1">
      <mxGeometry x="{x-10}" y="65" width="{NODE_W+20}" height="24" as="geometry" />
    </mxCell>''')

    # Nodes
    node_ids_map = {}   # original id → cell id
    for n in nodes:
        nid   = _sid(n["id"])
        ntype = n.get("type","process")
        label = n.get("label","")
        c     = COLORS.get(ntype, COLORS["process"])
        x,y   = node_positions.get(nid,(400,300))
        cid   = new_id()
        node_ids_map[nid] = cid

        shape_attr = ""
        if c.get("shape") == "rhombus":
            shape_attr = "rhombus;verticalLabelPosition=middle;"
        elif c.get("shape") == "ellipse":
            shape_attr = "ellipse;"
        elif "stored_data" in c.get("shape",""):
            shape_attr = "shape=cylinder3;boundedLbl=1;backgroundOutline=1;"

        extra = c.get("style_extra","")
        style = (f"rounded={c.get('rounded','0')};whiteSpace=wrap;html=1;"
                 f"fillColor={c['fillColor']};strokeColor={c['strokeColor']};"
                 f"fontColor={c['fontColor']};fontSize=11;"
                 f"{shape_attr}{extra};")

        cells.append(f'''    <mxCell id="{cid}" value="{_xml_escape(label)}" style="{style}" vertex="1" parent="1">
      <mxGeometry x="{x}" y="{y}" width="{NODE_W}" height="{NODE_H}" as="geometry" />
    </mxCell>''')

        # Privacy controls (future state)
        if state=="future" and nid in ctrls:
            ctrl_list = ctrls[nid][:6]
            nc  = 2
            nr  = math.ceil(len(ctrl_list)/nc)
            bw  = nc*CTRL_W + (nc-1)*8
            bh  = nr*CTRL_H + (nr-1)*6
            bx  = x + (NODE_W-bw)//2
            by  = y - bh - 18

            for ci, ctrl in enumerate(ctrl_list):
                r=ci//nc; col=ci%nc
                cx2 = bx + col*(CTRL_W+8)
                cy2 = by + r*(CTRL_H+6)
                ccid = new_id()
                ctrl_style = (f"rounded=1;whiteSpace=wrap;html=1;"
                              f"fillColor=#D5E8D4;strokeColor=#27AE60;"
                              f"fontColor=#145A32;fontSize=9;")
                cells.append(f'''    <mxCell id="{ccid}" value="{_xml_escape(ctrl[:28])}" style="{ctrl_style}" vertex="1" parent="1">
      <mxGeometry x="{cx2}" y="{cy2}" width="{CTRL_W}" height="{CTRL_H}" as="geometry" />
    </mxCell>''')
                # Connector
                conn_id = new_id()
                cells.append(f'''    <mxCell id="{conn_id}" style="edgeStyle=none;dashed=1;strokeColor=#27AE60;exitX=0.5;exitY=0;exitDx=0;exitDy=0;" edge="1" source="{cid}" target="{ccid}" parent="1">
      <mxGeometry relative="1" as="geometry" />
    </mxCell>''')

    # Edges
    def _norm(s):
        return re.sub(r"[^a-z0-9]","_",str(s).lower())
    norm_map = {_norm(k):k for k in node_ids_map}

    for e in edges:
        src_sid = _sid(e.get("from",""))
        dst_sid = _sid(e.get("to",""))
        src_cid = node_ids_map.get(src_sid)
        dst_cid = node_ids_map.get(dst_sid)
        if not src_cid or not dst_cid: continue
        lbl = e.get("label","")
        sensitive = any(k in lbl.lower() for k in ["salary","bank","health","medical","biometric"])
        color = "#C0392B" if sensitive else "#666666"
        eid = new_id()
        cells.append(f'''    <mxCell id="{eid}" value="{_xml_escape(lbl)}" style="edgeStyle=orthogonalEdgeStyle;rounded=1;strokeColor={color};fontColor={color};fontSize=9;exitX=1;exitY=0.5;exitDx=0;exitDy=0;entryX=0;entryY=0.5;entryDx=0;entryDy=0;" edge="1" source="{src_cid}" target="{dst_cid}" parent="1">
      <mxGeometry relative="1" as="geometry" />
    </mxCell>''')

    xml = f'''<?xml version="1.0" encoding="UTF-8"?>
<mxfile host="app.diagrams.net">
  <diagram name="{_xml_escape(title)} – {state.upper()}">
    <mxGraphModel dx="1422" dy="762" grid="0" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="0" pageScale="1" pageWidth="1654" pageHeight="1169" math="0" shadow="0">
      <root>
        <mxCell id="0" /><mxCell id="1" parent="0" />
{chr(10).join(cells)}
      </root>
    </mxGraphModel>
  </diagram>
</mxfile>'''
    return xml
