import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import warnings
warnings.filterwarnings('ignore')

# ── Page config ──────────────────────────────────────────────
st.set_page_config(
    page_title="SwingEdge Pro — Bottom Bounce",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ── Custom CSS ────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Mono:wght@400;500&display=swap');

:root {
    --navy: #0a0f1e;
    --navy2: #111827;
    --orange: #f97316;
    --text: #e2e8f0;
    --muted: #94a3b8;
    --border: #1e293b;
    --card: #0f172a;
}

html, body, [class*="css"] {
    font-family: 'Syne', sans-serif;
    background-color: var(--navy) !important;
    color: var(--text) !important;
}
.stApp { background-color: var(--navy) !important; }
#MainMenu, footer, header { visibility: hidden; }

.login-box {
    max-width: 400px;
    margin: 8rem auto;
    background: var(--card);
    border: 1px solid var(--border);
    border-top: 3px solid var(--orange);
    border-radius: 12px;
    padding: 2.5rem;
    text-align: center;
}
.login-logo { font-size: 2rem; font-weight: 800; margin-bottom: 0.25rem; }
.login-logo span { color: var(--orange); }
.login-sub { color: var(--muted); font-size: 0.8rem; margin-bottom: 2rem; font-family: 'DM Mono', monospace; }

.sep-header {
    background: linear-gradient(135deg, #0a0f1e 0%, #111827 50%, #0a0f1e 100%);
    border-bottom: 2px solid var(--orange);
    padding: 1.5rem 2rem;
    margin: -1rem -1rem 2rem -1rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
}
.sep-logo { font-size: 1.8rem; font-weight: 800; letter-spacing: -1px; }
.sep-logo span { color: var(--orange); }
.sep-badge {
    background: var(--orange);
    color: #000;
    font-size: 0.65rem;
    font-weight: 700;
    padding: 2px 8px;
    border-radius: 3px;
    letter-spacing: 1px;
    margin-left: 8px;
    vertical-align: middle;
}
.sep-meta { font-family: 'DM Mono', monospace; font-size: 0.8rem; color: var(--muted); }

.info-bar {
    background: var(--card);
    border: 1px solid var(--border);
    border-left: 3px solid var(--orange);
    border-radius: 6px;
    padding: 0.75rem 1.25rem;
    margin-bottom: 1.5rem;
    display: flex;
    gap: 2rem;
    font-family: 'DM Mono', monospace;
    font-size: 0.8rem;
    color: var(--muted);
}
.info-bar b { color: var(--orange); }

.metric-row { display: flex; gap: 1rem; margin-bottom: 1.5rem; flex-wrap: wrap; }
.metric-card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 1rem 1.5rem;
    flex: 1;
    min-width: 120px;
    text-align: center;
}
.metric-num { font-size: 2rem; font-weight: 800; color: var(--orange); line-height: 1; }
.metric-label { font-size: 0.7rem; color: var(--muted); margin-top: 4px; letter-spacing: 0.5px; }

.upload-zone {
    border: 2px dashed var(--border);
    border-radius: 12px;
    padding: 3rem;
    text-align: center;
    color: var(--muted);
    margin: 2rem 0;
}

.stTabs [data-baseweb="tab-list"] {
    background: var(--card) !important;
    border-radius: 8px 8px 0 0;
    border-bottom: 2px solid var(--orange) !important;
    gap: 0;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    color: var(--muted) !important;
    font-weight: 600 !important;
    font-size: 0.85rem !important;
    padding: 0.6rem 1rem !important;
    border-radius: 0 !important;
}
.stTabs [aria-selected="true"] {
    background: var(--orange) !important;
    color: #000 !important;
    border-radius: 6px 6px 0 0 !important;
}
.stTabs [data-baseweb="tab-panel"] {
    background: var(--card) !important;
    border-radius: 0 0 8px 8px !important;
    padding: 1rem !important;
    border: 1px solid var(--border) !important;
    border-top: none !important;
}

.watermark {
    text-align: center;
    color: var(--border);
    font-size: 0.7rem;
    font-family: 'DM Mono', monospace;
    margin-top: 3rem;
    letter-spacing: 1px;
}

/* Input fields */
.stTextInput input {
    background: #1e293b !important;
    border: 1px solid #334155 !important;
    color: var(--text) !important;
    border-radius: 6px !important;
}
button[kind="primary"] {
    background: var(--orange) !important;
    color: #000 !important;
    font-weight: 700 !important;
    border: none !important;
}
</style>
""", unsafe_allow_html=True)

# ── PASSWORD GATE ─────────────────────────────────────────────
PASSWORD = "SwingEdge@2026"  # ← Change this to your preferred password

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("""
    <div class="login-box">
        <div class="login-logo">Swing<span>Edge</span> Pro</div>
        <div class="login-sub">BOTTOM BOUNCE DASHBOARD · PRIVATE ACCESS</div>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        pwd = st.text_input("Password", type="password", placeholder="Enter password...")
        if st.button("🔓 Login", use_container_width=True):
            if pwd == PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Incorrect password.")
    st.stop()

# ── Header ────────────────────────────────────────────────────
st.markdown("""
<div class="sep-header">
  <div>
    <span class="sep-logo">Swing<span>Edge</span> Pro<span class="sep-badge">PRO</span></span>
    <div style="font-size:0.75rem; color:#94a3b8; margin-top:4px; font-family:'DM Mono',monospace;">
      Bottom Bounce EMA Crossover Scanner
    </div>
  </div>
  <div class="sep-meta">NSE · 1304 Stocks · Daily</div>
</div>
""", unsafe_allow_html=True)

# ── Load Excel ────────────────────────────────────────────────
def load_excel(file):
    wb = openpyxl.load_workbook(file, data_only=True)
    data = {}

    ws_info = wb['📅 Scan Info']
    info = {}
    for row in ws_info.iter_rows(min_row=2, values_only=True):
        if row[0] and row[1]:
            info[str(row[0])] = str(row[1])
    data['info'] = info

    sheet_map = {
        'T0': '👀 T0 — Pre-Recovery',
        'T1': '🌱 T1 — 10W Cross 20W',
        'T2': '🔥 T2 — 10W Cross 40W',
        'T3': '🚀 T3 — 20W Cross 40W',
        'COMBO': '⚡ COMBO T2+T3',
    }

    for key, sheetname in sheet_map.items():
        if sheetname in wb.sheetnames:
            ws = wb[sheetname]
            rows = list(ws.iter_rows(values_only=True))
            if len(rows) < 2:
                data[key] = pd.DataFrame()
                continue
            headers = rows[0]
            df = pd.DataFrame(rows[1:], columns=headers)
            if df.columns[0] is None:
                df = df.iloc[:, 1:]
            df = df.dropna(subset=['Symbol'])
            data[key] = df
        else:
            data[key] = pd.DataFrame()

    return data

# ── File Upload ───────────────────────────────────────────────
uploaded = st.file_uploader(
    "Upload Bottom Bounce Excel",
    type=["xlsx"],
    label_visibility="collapsed"
)

if uploaded is None:
    st.markdown("""
    <div class="upload-zone">
        <div style="font-size:2.5rem; margin-bottom:1rem;">📊</div>
        <div style="font-size:1.1rem; font-weight:700; color:#e2e8f0; margin-bottom:0.5rem;">
            Upload your Bottom Bounce scan result
        </div>
        <div style="font-size:0.85rem;">
            Drop the <b style="color:#f97316">SwingEdge_BottomBounce_YYYY-MM-DD.xlsx</b> file above
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

with st.spinner("Loading scan data..."):
    data = load_excel(uploaded)

info = data.get('info', {})
scan_date = info.get('Scan Date', '—')
run_time = info.get('Run Time', '—')
universe = info.get('Universe', '1304')

counts = {k: len(data.get(k, pd.DataFrame())) for k in ['T0','T1','T2','T3','COMBO']}
total = counts['T0'] + counts['T1'] + counts['T2'] + counts['T3']

st.markdown(f"""
<div class="info-bar">
  <span>📅 <b>Scan Date:</b> {scan_date}</span>
  <span>🕐 <b>Run Time:</b> {run_time} IST</span>
  <span>🌐 <b>Universe:</b> {universe} stocks</span>
  <span>📊 <b>Total Signals:</b> {total}</span>
</div>
""", unsafe_allow_html=True)

st.markdown(f"""
<div class="metric-row">
  <div class="metric-card">
    <div class="metric-num">{counts['T0']}</div>
    <div class="metric-label">👀 T0 PRE-RECOVERY</div>
  </div>
  <div class="metric-card">
    <div class="metric-num">{counts['T1']}</div>
    <div class="metric-label">🌱 T1 EARLY</div>
  </div>
  <div class="metric-card">
    <div class="metric-num">{counts['T2']}</div>
    <div class="metric-label">🔥 T2 MOMENTUM</div>
  </div>
  <div class="metric-card">
    <div class="metric-num">{counts['T3']}</div>
    <div class="metric-label">🚀 T3 TREND RESTORED</div>
  </div>
  <div class="metric-card" style="border-color:#f97316;">
    <div class="metric-num" style="color:#fbbf24;">{counts['COMBO']}</div>
    <div class="metric-label">⚡ COMBO T2+T3</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Tabs ──────────────────────────────────────────────────────
stage_keys = ['T0', 'T1', 'T2', 'T3', 'COMBO']
stage_desc = {
    'T0': 'Pre-Recovery — 10W & 20W both still below 40W but gap narrowing',
    'T1': 'Early Stage — 10W has crossed above 20W',
    'T2': 'Momentum — 10W has crossed above 40W',
    'T3': 'Trend Restored — 20W has crossed above 40W',
    'COMBO': 'High Conviction — T2 + T3 both triggered',
}
tab_labels = [
    f"👀 T0  ({counts['T0']})",
    f"🌱 T1  ({counts['T1']})",
    f"🔥 T2  ({counts['T2']})",
    f"🚀 T3  ({counts['T3']})",
    f"⚡ COMBO  ({counts['COMBO']})",
]

tabs = st.tabs(tab_labels)

for tab, key in zip(tabs, stage_keys):
    with tab:
        df = data.get(key, pd.DataFrame())
        if df.empty:
            st.info("No signals for this stage today.")
        else:
            st.caption(f"_{stage_desc[key]}_")
            search = st.text_input(
                "Filter", key=f"search_{key}",
                placeholder="🔍 Filter by symbol...",
                label_visibility="collapsed"
            )
            if search:
                df = df[df['Symbol'].str.contains(search.upper(), na=False)]

            st.dataframe(df, use_container_width=True,
                         height=min(600, 50 + len(df) * 35),
                         hide_index=True)

            buf = BytesIO()
            df.to_excel(buf, index=False)
            st.download_button(
                f"⬇ Download {key} as Excel",
                data=buf.getvalue(),
                file_name=f"BB_{key}_{scan_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{key}"
            )

# Logout
st.sidebar.markdown("---")
if st.sidebar.button("🔒 Logout"):
    st.session_state.authenticated = False
    st.rerun()

st.markdown('<div class="watermark">SWINGEDGE PRO · CONFIDENTIAL · FOR INTERNAL USE ONLY</div>', unsafe_allow_html=True)
