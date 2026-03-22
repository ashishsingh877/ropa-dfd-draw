"""
app.py  -  ROPA Intelligence Platform
Two clean landscape DFDs per process: Current State + Post Compliance.
"""
import io, json, base64, zipfile
from datetime import datetime
import streamlit as st
from ropa_parser  import parse_ropa_excel, processes_to_text
from ai_client    import chat, stream_chat, parse_json_from_response
from prompts      import EXTRACT_SYSTEM, DFD_SYSTEM, RISK_SYSTEM
from dfd_renderer import render_dfd
from drawio_export  import generate_drawio_xml
from visio_export   import generate_vsdx

st.set_page_config(page_title="ROPA — DFD Analyzer", page_icon="🔐",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap');
html,body,[class*="css"]  { font-family:'Inter',sans-serif; }
.stApp { background:#0d1117; color:#c9d1d9; }
section[data-testid="stSidebar"] { background:#161b22; border-right:1px solid #21262d; }
.hero { background:linear-gradient(160deg,#0d1117 0%,#0d2136 50%,#0d1117 100%);
        border:1px solid #1f3a5c; border-radius:14px; padding:2rem; margin-bottom:1.5rem; text-align:center; }
.hero-title { font-family:'JetBrains Mono',monospace; font-size:1.75rem; color:#58a6ff; margin:0; }
.hero-sub   { color:#6e7681; font-size:.9rem; margin-top:.4rem; }
.metric-row { display:flex; gap:.75rem; margin:1rem 0; flex-wrap:wrap; }
.metric-box { flex:1; min-width:100px; text-align:center; background:#161b22;
              border:1px solid #21262d; border-radius:10px; padding:.85rem; }
.metric-num { font-family:'JetBrains Mono',monospace; font-size:1.85rem; color:#58a6ff; font-weight:700; }
.metric-lbl { font-size:.65rem; color:#484f58; text-transform:uppercase; letter-spacing:1.5px; }
.stage-row  { display:flex; gap:.5rem; margin:1rem 0; flex-wrap:wrap; }
.stage      { flex:1; min-width:120px; text-align:center; border:1px solid #21262d; border-radius:7px;
              padding:.45rem; font-family:'JetBrains Mono',monospace; font-size:.72rem; color:#484f58; background:#161b22; }
.stage.active { border-color:#388bfd; color:#388bfd; background:#0d1f35; }
.stage.done   { border-color:#3fb950; color:#3fb950; background:#0d1e14; }
.info-box { background:#0d1f35; border:1px solid #1f3a5c; border-radius:8px; padding:.65rem 1rem; color:#7a8db0; font-size:.82rem; }
.badge { display:inline-block; padding:.15rem .6rem; border-radius:999px; font-size:.7rem; font-weight:600; font-family:'JetBrains Mono',monospace; }
.b-blue  { background:#0d2136; color:#58a6ff; border:1px solid #1f3a5c; }
.b-green { background:#0d1e14; color:#3fb950; border:1px solid #1a4228; }
.stButton>button { background:linear-gradient(135deg,#1f3a5c,#388bfd30); color:#58a6ff; border:1px solid #388bfd;
                   border-radius:8px; font-family:'JetBrains Mono',monospace; font-weight:600; padding:.55rem 1.4rem; transition:all .2s; }
.stButton>button:hover { background:#388bfd; color:#fff; transform:translateY(-1px); }
.stTabs [data-baseweb="tab-list"] { background:#161b22; border-radius:8px; gap:3px; }
.stTabs [data-baseweb="tab"]      { color:#484f58; border-radius:6px; font-size:.85rem; }
.stTabs [aria-selected="true"]    { background:#21262d !important; color:#58a6ff !important; }
div[data-testid="stExpander"]     { background:#161b22; border:1px solid #21262d; border-radius:8px; }
hr { border-color:#21262d; }
</style>
""", unsafe_allow_html=True)

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<p style="font-family:JetBrains Mono,monospace;color:#58a6ff;font-size:1rem;">🔐 ROPA Analyzer</p>', unsafe_allow_html=True)
    st.markdown("---")
    # Auto-load from Streamlit Secrets first
    _secret_key = ""
    try:
        _secret_key = (st.secrets.get("GROQ_API_KEY","") or
                       st.secrets.get("groq_api_key",""))
    except Exception:
        pass
    if _secret_key:
        groq_key = _secret_key
        st.success("🔑 API key loaded from Secrets")
    else:
        groq_key = st.text_input("Groq API Key", type="password", placeholder="gsk_…",
                                  help="Tip: add GROQ_API_KEY to Streamlit Secrets to skip typing this")
    model    = st.selectbox("Model", ["llama-3.3-70b-versatile","llama3-70b-8192","mixtral-8x7b-32768"])
    st.markdown("---")
    st.markdown("""
**Output per process**
- 🔴 Current State DFD (PNG + PDF)
- 🟢 Post Compliance DFD (PNG + PDF)
- ⚠️ Risk & Gap Analysis
""")
    st.markdown("---")
    st.markdown('<span class="badge b-green">Groq Free</span> <span class="badge b-blue">DPDPA 2023</span>', unsafe_allow_html=True)

# ── Hero ───────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
  <div class="hero-title">🔐 ROPA Intelligence Platform</div>
  <div class="hero-sub">Upload ROPA → AI generates professional As-Is &amp; Post-Compliance DFDs → Export PNG / PDF</div>
</div>
""", unsafe_allow_html=True)

# ── Upload ─────────────────────────────────────────────────────────────────────
st.markdown("### 📂 Upload ROPA File")
c1, c2 = st.columns([2, 3])
with c1:
    uploaded = st.file_uploader("ROPA Excel", type=["xlsx","xls"], label_visibility="collapsed")
with c2:
    st.markdown('<div class="info-box">Supports both ROPA layouts:<br>• <b>Vertical</b> — one sheet per process (Data Fiduciary format)<br>• <b>Horizontal</b> — RoPA_Template with 7-section, 53-column layout</div>', unsafe_allow_html=True)

if uploaded:
    st.markdown(f'<span class="badge b-green">✓ {uploaded.name}</span> <span class="badge b-blue">{uploaded.size:,} bytes</span>', unsafe_allow_html=True)

st.markdown("---")

# ── Helpers ────────────────────────────────────────────────────────────────────
def stream_box(api_key, system, user, max_tokens=6000):
    full, ph = "", st.empty()
    try:
        for chunk in stream_chat(api_key, system, user, max_tokens, model):
            full += chunk
            preview = full[-900:].replace("<","&lt;").replace(">","&gt;")
            ph.markdown(f'<div style="background:#010409;border:1px solid #21262d;border-radius:6px;padding:.6rem;font-family:JetBrains Mono,monospace;font-size:.75rem;color:#58a6ff;max-height:160px;overflow:auto;">{preview}</div>', unsafe_allow_html=True)
    except Exception:
        try: full = chat(api_key, system, user, max_tokens, model)
        except Exception as e: st.error(f"API error: {e}"); return ""
    ph.empty()
    return full

def b64(b): return base64.b64encode(b).decode()

def show_dfd_pair(dfd, idx):
    """Show two side-by-side sub-tabs: Current State | Post Compliance."""
    asis_png  = dfd.get("_asis_png")
    asis_pdf  = dfd.get("_asis_pdf")
    future_png= dfd.get("_future_png")
    future_pdf= dfd.get("_future_pdf")
    pid       = dfd.get("id", f"P{idx+1:03d}")
    pname     = dfd.get("process_name", f"Process {idx+1}")
    safe      = pname.replace(" ","_").replace("/","-")[:35]

    t_asis, t_future = st.tabs(["🔴 Current State", "🟢 Post Compliance"])

    with t_asis:
        if asis_png:
            st.markdown(f'<div style="background:#f8f9fa;border:1px solid #dee2e6;border-radius:8px;padding:4px;overflow:auto;"><img src="data:image/png;base64,{b64(asis_png)}" style="width:100%;min-width:900px;" /></div>', unsafe_allow_html=True)
            cc1, cc2 = st.columns(2)
            with cc1: st.download_button("⬇️ PNG — Current State", data=asis_png,  file_name=f"DFD_{pid}_{safe}_CurrentState.png", mime="image/png",        use_container_width=True, key=f"a_png_{idx}")
            with cc2: st.download_button("⬇️ PDF — Current State", data=asis_pdf,  file_name=f"DFD_{pid}_{safe}_CurrentState.pdf", mime="application/pdf",   use_container_width=True, key=f"a_pdf_{idx}")
        else:
            st.error(dfd.get("_render_error","Render failed"))

    with t_future:
        if future_png:
            st.markdown(f'<div style="background:#f8f9fa;border:1px solid #dee2e6;border-radius:8px;padding:4px;overflow:auto;"><img src="data:image/png;base64,{b64(future_png)}" style="width:100%;min-width:900px;" /></div>', unsafe_allow_html=True)
            cc1, cc2 = st.columns(2)
            with cc1: st.download_button("⬇️ PNG — Post Compliance", data=future_png, file_name=f"DFD_{pid}_{safe}_PostCompliance.png", mime="image/png",      use_container_width=True, key=f"f_png_{idx}")
            with cc2: st.download_button("⬇️ PDF — Post Compliance", data=future_pdf, file_name=f"DFD_{pid}_{safe}_PostCompliance.pdf", mime="application/pdf", use_container_width=True, key=f"f_pdf_{idx}")
        else:
            st.error(dfd.get("_render_error","Render failed"))

# ── Analysis trigger ───────────────────────────────────────────────────────────
stage_ph = st.empty()

def set_stages(s1, s2, s3):
    stage_ph.markdown(
        f'<div class="stage-row">'
        f'<div class="stage {s1}">① Extract</div>'
        f'<div class="stage {s2}">② Build DFDs</div>'
        f'<div class="stage {s3}">③ Risk Report</div>'
        f'</div>', unsafe_allow_html=True)

if uploaded and groq_key:
    c_run, c_info = st.columns([1,3])
    with c_run:
        run = st.button("🚀  Analyse ROPA", use_container_width=True)
    with c_info:
        st.markdown('<div class="info-box">3 stages: Extract → DFDs → Risk Report. Each process gets two clean landscape diagrams (Current State + Post Compliance) rendered as real PNG/PDF images.</div>', unsafe_allow_html=True)

    if run:
        raw_bytes = uploaded.read()
        with st.spinner("Parsing ROPA file…"):
            raw_procs = parse_ropa_excel(raw_bytes, uploaded.name)

        if not raw_procs or (len(raw_procs)==1 and raw_procs[0].get("_format")=="B_EMPTY"):
            st.error("No processing activities found. Please upload a filled ROPA file.")
            st.stop()

        set_stages("active","","")
        st.markdown("#### ① Extracting processing activities…")

        raw_ext = stream_box(groq_key, EXTRACT_SYSTEM, f"ROPA DATA:\n\n{processes_to_text(raw_procs)[:18000]}", 4096)
        try:    enriched = parse_json_from_response(raw_ext)
        except: enriched = raw_procs

        st.session_state["enriched"] = enriched
        n         = len(enriched)
        depts     = len({p.get("function_name","") for p in enriched if p.get("function_name")})
        sensitive = sum(1 for p in enriched if str(p.get("sensitive_data","")).strip().lower() not in ("","none","no","n/a","not applicable"))
        transfers = sum(1 for p in enriched if str(p.get("transfer_jurisdictions","")).strip().lower() not in ("","none","no","n/a","not applicable"))

        st.success(f"✓ {n} processing activities extracted")
        st.markdown(
            f'<div class="metric-row">'
            f'<div class="metric-box"><div class="metric-num">{n}</div><div class="metric-lbl">Processes</div></div>'
            f'<div class="metric-box"><div class="metric-num">{depts}</div><div class="metric-lbl">Departments</div></div>'
            f'<div class="metric-box"><div class="metric-num">{sensitive}</div><div class="metric-lbl">Sensitive</div></div>'
            f'<div class="metric-box"><div class="metric-num">{transfers}</div><div class="metric-lbl">Transfers</div></div>'
            f'</div>', unsafe_allow_html=True)

        set_stages("done","active","")
        st.markdown("#### ② Generating Data Flow Diagrams…")

        # Generate ONE DFD per process (avoids token truncation on multi-process responses)
        dfd_list = []
        gen_prog = st.progress(0, text="Generating DFD JSON…")
        for pi, proc in enumerate(enriched):
            pname = proc.get("process_name", f"Process {pi+1}")
            gen_prog.progress(pi / max(len(enriched), 1),
                              text=f"DFD {pi+1}/{len(enriched)}: {pname[:45]}…")
            try:
                raw_dfd = chat(
                    groq_key, DFD_SYSTEM,
                    (f"Generate a DFD for this ONE processing activity only.\n"
                     f"Return a JSON ARRAY with EXACTLY ONE element.\n\n"
                     f"{json.dumps(proc, indent=2)[:8000]}"),
                    16000, model
                )
                result = parse_json_from_response(raw_dfd)
                if isinstance(result, list) and result:
                    item = result[0]
                    if not item.get("id"):
                        item["id"] = f"P{pi+1:03d}"
                    dfd_list.append(item)
                elif isinstance(result, dict):
                    if not result.get("id"):
                        result["id"] = f"P{pi+1:03d}"
                    dfd_list.append(result)
            except Exception as e:
                st.warning(f"⚠️ DFD skipped for **{pname}**: {e}")
        gen_prog.progress(1.0, text=f"✓ {len(dfd_list)} DFD(s) generated")
        gen_prog.empty()

        prog = st.progress(0, text="Rendering diagrams…")
        rendered = []
        for i, dfd in enumerate(dfd_list):
            try:
                a_png, a_pdf, f_png, f_pdf = render_dfd(dfd)
                dfd["_asis_png"]   = a_png
                dfd["_asis_pdf"]   = a_pdf
                dfd["_future_png"] = f_png
                dfd["_future_pdf"] = f_pdf
            except Exception as e:
                dfd["_asis_png"] = dfd["_asis_pdf"] = dfd["_future_png"] = dfd["_future_pdf"] = None
                dfd["_render_error"] = str(e)
            rendered.append(dfd)
            prog.progress((i+1)/max(len(dfd_list),1), text=f"Rendered {i+1}/{len(dfd_list)}")
        prog.empty()
        st.session_state["dfds"] = rendered
        st.success(f"✓ {len(rendered)} DFDs rendered")

        set_stages("done","done","active")
        st.markdown("#### ③ Running Risk & Gap Analysis…")
        risk_md = stream_box(groq_key, RISK_SYSTEM, f"PROCESSING ACTIVITIES:\n\n{json.dumps(enriched,indent=2)[:16000]}", 5000)
        st.session_state["risk_md"] = risk_md
        set_stages("done","done","done")
        st.success("✓ Complete")
        st.balloons()

elif not uploaded:
    st.markdown('<div style="background:#161b22;border:1px solid #21262d;border-radius:10px;text-align:center;color:#21262d;padding:3rem;">⬆️  Upload a ROPA Excel file above to begin</div>', unsafe_allow_html=True)
elif not groq_key:
    st.markdown('<div style="background:#161b22;border:1px solid #21262d;border-radius:10px;text-align:center;color:#21262d;padding:2rem;">🔑  Enter your Groq API key in the sidebar</div>', unsafe_allow_html=True)

# ── Results ────────────────────────────────────────────────────────────────────
if "dfds" in st.session_state or "risk_md" in st.session_state:
    dfds     = st.session_state.get("dfds", [])
    risk_md  = st.session_state.get("risk_md", "")
    enriched = st.session_state.get("enriched", [])

    st.markdown("---")
    tab_dfd, tab_risk, tab_export = st.tabs(["🔷 Data Flow Diagrams", "⚠️ Risk Analysis", "⬇️ Export"])

    with tab_dfd:
        if not dfds:
            if "dfds_raw" in st.session_state:
                st.warning("JSON parse failed:"); st.code(st.session_state["dfds_raw"])
            else:
                st.info("Run the analysis to generate DFDs.")
        else:
            st.markdown(f'<div class="info-box">&#9670; {len(dfds)} process(es) — click a process to view Current State and Post Compliance diagrams. Download PNG or PDF from each tab.</div>', unsafe_allow_html=True)
            st.markdown("")
            for i, dfd in enumerate(dfds):
                pname = dfd.get("process_name", f"Process {i+1}")
                pid   = dfd.get("id", f"P{i+1:03d}")
                with st.expander(f"**[{pid}]** {pname}", expanded=(i==0)):
                    show_dfd_pair(dfd, i)

    with tab_risk:
        if risk_md: st.markdown(risk_md)
        else:       st.info("Run analysis to generate the Risk Report.")

    with tab_export:
        st.markdown("### ⬇️ Export All Outputs")
        ts = datetime.now().strftime("%Y%m%d_%H%M")

        if dfds:
            st.markdown("**All DFDs**")
            for i, dfd in enumerate(dfds):
                pname = dfd.get("process_name", f"P{i+1:03d}")
                safe  = pname.replace(" ","_").replace("/","-")[:30]
                pid   = dfd.get("id", f"P{i+1:03d}")
                cc = st.columns(4)
                with cc[0]: st.markdown(f"**{pid}** {pname}")
                with cc[1]:
                    if dfd.get("_asis_png"):   st.download_button("🔴 PNG As-Is",       data=dfd["_asis_png"],   file_name=f"DFD_{pid}_{safe}_CurrentState.png",   mime="image/png",      use_container_width=True, key=f"ex_ap_{i}")
                with cc[2]:
                    if dfd.get("_asis_pdf"):   st.download_button("🔴 PDF As-Is",       data=dfd["_asis_pdf"],   file_name=f"DFD_{pid}_{safe}_CurrentState.pdf",   mime="application/pdf",use_container_width=True, key=f"ex_apdf_{i}")
                with cc[3]:
                    if dfd.get("_future_png"): st.download_button("🟢 PNG Compliance",  data=dfd["_future_png"], file_name=f"DFD_{pid}_{safe}_PostCompliance.png", mime="image/png",      use_container_width=True, key=f"ex_fp_{i}")

        st.markdown("---")
        col1, col2 = st.columns(2)
        with col1: st.download_button("⚠️ Risk Analysis (.md)", data=risk_md or "Not generated.", file_name=f"risk_analysis_{ts}.md", mime="text/markdown",    use_container_width=True)
        with col2: st.download_button("📊 Processes JSON",      data=json.dumps(enriched,indent=2,ensure_ascii=False), file_name=f"ropa_processes_{ts}.json", mime="application/json", use_container_width=True)

        st.markdown("---")
        zio = io.BytesIO()
        with zipfile.ZipFile(zio,"w",zipfile.ZIP_DEFLATED) as zf:
            for i, dfd in enumerate(dfds):
                pn   = dfd.get("process_name",f"P{i+1:03d}")
                safe = pn.replace(" ","_").replace("/","-")[:30]
                pid  = dfd.get("id",f"P{i+1:03d}")
                try:
                    va,vf = generate_vsdx(dfd)
                    zf.writestr(f"visio/DFD_{pid}_{safe}_CurrentState.vsdx", va)
                    zf.writestr(f"visio/DFD_{pid}_{safe}_PostCompliance.vsdx", vf)
                except Exception: pass
                if dfd.get("_asis_png"):   zf.writestr(f"dfds/DFD_{pid}_{safe}_CurrentState.png",   dfd["_asis_png"])
                if dfd.get("_asis_pdf"):   zf.writestr(f"dfds/DFD_{pid}_{safe}_CurrentState.pdf",   dfd["_asis_pdf"])
                if dfd.get("_future_png"): zf.writestr(f"dfds/DFD_{pid}_{safe}_PostCompliance.png", dfd["_future_png"])
                if dfd.get("_future_pdf"): zf.writestr(f"dfds/DFD_{pid}_{safe}_PostCompliance.pdf", dfd["_future_pdf"])
            if risk_md: zf.writestr(f"risk_analysis_{ts}.md", risk_md)
            zf.writestr(f"ropa_processes_{ts}.json", json.dumps(enriched,indent=2,ensure_ascii=False))
        zio.seek(0)
        # ── Draw.io XML export ─────────────────────────────────────────────────
        st.markdown("---")
        st.markdown("### 🔧 Editable Draw.io Files")
        st.markdown('<div class="info-box">Download as draw.io XML → open at <b>app.diagrams.net</b> → fully editable, resize, restyle, add annotations. Free tool.</div>', unsafe_allow_html=True)
        drawio_cols = st.columns(min(len(dfds)*2, 6))
        dci = 0
        for i, dfd in enumerate(dfds):
            pname = dfd.get("process_name", f"P{i+1:03d}")
            safe  = pname.replace(" ","_").replace("/","-")[:30]
            pid   = dfd.get("id", f"P{i+1:03d}")
            try:
                xml_a = generate_drawio_xml(dfd, "asis")
                xml_f = generate_drawio_xml(dfd, "future")
                with drawio_cols[dci % 6]:
                    st.download_button(f"📐 {pid} As-Is (XML)",
                        data=xml_a, file_name=f"DFD_{pid}_{safe}_AsIs.drawio",
                        mime="application/xml", use_container_width=True,
                        key=f"dio_a_{i}")
                dci+=1
                with drawio_cols[dci % 6]:
                    st.download_button(f"📐 {pid} Post-Compliance (XML)",
                        data=xml_f, file_name=f"DFD_{pid}_{safe}_PostCompliance.drawio",
                        mime="application/xml", use_container_width=True,
                        key=f"dio_f_{i}")
                dci+=1
            except Exception as ex:
                st.warning(f"draw.io export failed for {pid}: {ex}")

        # ── Visio .vsdx export ─────────────────────────────────────────────────
        st.markdown("---")
        st.markdown("### 🔷 Microsoft Visio Files (.vsdx)")
        st.markdown('<div class="info-box">Download as <b>.vsdx</b> → open directly in <b>Microsoft Visio</b> (2013+ or Microsoft 365). All shapes, connectors and privacy controls are fully editable. Resize, restyle, add annotations, change layout.</div>', unsafe_allow_html=True)
        visio_cols = st.columns(min(len(dfds)*2, 6))
        vci = 0
        for i, dfd in enumerate(dfds):
            pname = dfd.get("process_name", f"P{i+1:03d}")
            safe  = pname.replace(" ","_").replace("/","-")[:30]
            pid   = dfd.get("id", f"P{i+1:03d}")
            try:
                v_asis, v_future = generate_vsdx(dfd)
                with visio_cols[vci % 6]:
                    st.download_button(f"🔷 {pid} Visio As-Is",
                        data=v_asis,
                        file_name=f"DFD_{pid}_{safe}_CurrentState.vsdx",
                        mime="application/vnd.ms-visio.drawing",
                        use_container_width=True, key=f"vsdx_a_{i}")
                vci+=1
                with visio_cols[vci % 6]:
                    st.download_button(f"🔷 {pid} Visio Post-Compliance",
                        data=v_future,
                        file_name=f"DFD_{pid}_{safe}_PostCompliance.vsdx",
                        mime="application/vnd.ms-visio.drawing",
                        use_container_width=True, key=f"vsdx_f_{i}")
                vci+=1
            except Exception as ex:
                st.warning(f"Visio export failed for {pid}: {ex}")

        st.markdown("---")
        st.download_button("📦 Full Bundle (ZIP)", data=zio.getvalue(), file_name=f"ropa_analysis_{ts}.zip", mime="application/zip")
