import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import warnings
warnings.filterwarnings('ignore')

st.set_page_config(
    page_title="SwingEdge Pro — Bottom Bounce",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&family=Space+Mono:wght@400;700&display=swap');

* { margin: 0; padding: 0; box-sizing: border-box; }

:root {
    --black: #0a0a0a;
    --dark: #111111;
    --card: #181818;
    --border: #222222;
    --gold: #f5a623;
    --gold2: #fbbf24;
    --white: #ffffff;
    --muted: #888888;
    --green: #22c55e;
}

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif !important;
    background-color: var(--black) !important;
    color: var(--white) !important;
}
.stApp { background-color: var(--black) !important; }
#MainMenu, footer, header { visibility: hidden; }
section[data-testid="stSidebar"] { display: none; }

/* ── TOP NAV ── */
.top-nav {
    position: fixed;
    top: 0; left: 0; right: 0;
    background: rgba(10,10,10,0.95);
    backdrop-filter: blur(10px);
    border-bottom: 1px solid var(--border);
    padding: 0 2rem;
    height: 56px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    z-index: 999;
}
.nav-logo {
    font-size: 1rem;
    font-weight: 800;
    letter-spacing: -0.5px;
    color: var(--white);
}
.nav-logo .dot { color: var(--gold); }
.nav-badge {
    background: var(--gold);
    color: #000;
    font-size: 0.6rem;
    font-weight: 700;
    padding: 2px 6px;
    border-radius: 3px;
    margin-left: 6px;
    letter-spacing: 0.5px;
}
.nav-right {
    font-family: 'Space Mono', monospace;
    font-size: 0.7rem;
    color: var(--muted);
}

/* ── PAGE BODY ── */
.page-body { padding-top: 72px; padding-left: 1rem; padding-right: 1rem; }

/* ── LOGIN ── */
.login-wrap {
    min-height: 100vh;
    display: flex;
    align-items: center;
    justify-content: center;
    padding-top: 56px;
}
.login-card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 16px;
    padding: 3rem 2.5rem;
    width: 100%;
    max-width: 400px;
    text-align: center;
}
.login-icon {
    width: 56px; height: 56px;
    background: var(--gold);
    border-radius: 14px;
    display: flex; align-items: center; justify-content: center;
    margin: 0 auto 1.5rem;
    font-size: 1.5rem;
}
.login-title { font-size: 1.4rem; font-weight: 800; margin-bottom: 0.25rem; }
.login-sub { color: var(--muted); font-size: 0.8rem; margin-bottom: 2rem; font-family: 'Space Mono', monospace; }
.login-err { color: #ef4444; font-size: 0.8rem; margin-top: 0.5rem; }

/* ── SCAN HEADER ── */
.scan-header {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 1.25rem 1.5rem;
    margin-bottom: 1.25rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
    flex-wrap: wrap;
    gap: 1rem;
}
.scan-title { font-size: 1rem; font-weight: 700; color: var(--gold); }
.scan-meta { font-family: 'Space Mono', monospace; font-size: 0.72rem; color: var(--muted); display:flex; gap:1.5rem; flex-wrap:wrap; }
.scan-meta span b { color: var(--white); }

/* ── METRIC CARDS ── */
.metrics { display: grid; grid-template-columns: repeat(5, 1fr); gap: 0.75rem; margin-bottom: 1.25rem; }
.m-card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 1rem;
    text-align: center;
    transition: border-color 0.2s;
}
.m-card:hover { border-color: var(--gold); }
.m-card.combo { border-color: var(--gold); background: #1a1500; }
.m-num { font-size: 2.2rem; font-weight: 900; color: var(--gold); line-height: 1; }
.m-combo { color: var(--gold2); }
.m-label { font-size: 0.65rem; color: var(--muted); margin-top: 6px; letter-spacing: 0.5px; font-weight: 600; }

/* ── TABS ── */
.stTabs [data-baseweb="tab-list"] {
    background: var(--card) !important;
    border-radius: 10px 10px 0 0 !important;
    border-bottom: 1px solid var(--gold) !important;
    gap: 0 !important;
    padding: 0 !important;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    color: var(--muted) !important;
    font-weight: 600 !important;
    font-size: 0.82rem !important;
    padding: 0.7rem 1.2rem !important;
    border-radius: 0 !important;
    letter-spacing: 0.2px !important;
}
.stTabs [aria-selected="true"] {
    background: var(--gold) !important;
    color: #000 !important;
    font-weight: 800 !important;
    border-radius: 8px 8px 0 0 !important;
}
.stTabs [data-baseweb="tab-panel"] {
    background: var(--card) !important;
    border: 1px solid var(--border) !important;
    border-top: none !important;
    border-radius: 0 0 10px 10px !important;
    padding: 1rem !important;
}

/* ── INPUTS ── */
.stTextInput input {
    background: #1a1a1a !important;
    border: 1px solid var(--border) !important;
    color: var(--white) !important;
    border-radius: 8px !important;
    font-size: 0.85rem !important;
}
.stTextInput input:focus { border-color: var(--gold) !important; box-shadow: 0 0 0 2px rgba(245,166,35,0.15) !important; }

/* ── BUTTONS ── */
.stButton button {
    background: var(--gold) !important;
    color: #000 !important;
    font-weight: 700 !important;
    border: none !important;
    border-radius: 8px !important;
    font-size: 0.85rem !important;
}
.stDownloadButton button {
    background: transparent !important;
    color: var(--gold) !important;
    border: 1px solid var(--gold) !important;
    font-weight: 600 !important;
    border-radius: 8px !important;
    font-size: 0.8rem !important;
}

/* ── DATAFRAME ── */
.stDataFrame { border-radius: 8px !important; overflow: hidden !important; }
iframe { border-radius: 8px !important; }

/* ── FOOTER ── */
.footer {
    text-align: center;
    color: #333;
    font-size: 0.65rem;
    font-family: 'Space Mono', monospace;
    margin-top: 3rem;
    padding-bottom: 2rem;
    letter-spacing: 1px;
}
</style>
""", unsafe_allow_html=True)

# ── PASSWORD ──────────────────────────────────────────────────
PASSWORD = "SwingEdge@2026"

if "auth" not in st.session_state:
    st.session_state.auth = False

# ── NAV BAR ──────────────────────────────────────────────────
st.markdown("""
<div class="top-nav">
  <div class="nav-logo">● SwingEdge<span class="dot"> Pro</span><span class="nav-badge">PRO</span></div>
  <div class="nav-right">NSE · 1304 STOCKS · DAILY SCANS</div>
</div>
""", unsafe_allow_html=True)

# ── LOGIN GATE ────────────────────────────────────────────────
if not st.session_state.auth:
    st.markdown('<div style="height:15vh"></div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("""
        <div class="login-card">
          <div class="login-icon">⚡</div>
          <div class="login-title">SwingEdge Pro</div>
          <div class="login-sub">BOTTOM BOUNCE · MEMBER ACCESS</div>
        </div>
        """, unsafe_allow_html=True)
        pwd = st.text_input("", type="password", placeholder="Enter your password...", label_visibility="collapsed")
        if st.button("Access Dashboard →", use_container_width=True):
            if pwd == PASSWORD:
                st.session_state.auth = True
                st.rerun()
            else:
                st.error("Incorrect password. Please try again.")
    st.stop()

# ── LOAD DATA ─────────────────────────────────────────────────
@st.cache_data
def load_data():
    try:
        wb = openpyxl.load_workbook("bottom_bounce.xlsx", data_only=True)
    except FileNotFoundError:
        return None, None
    
    ws_info = wb['📅 Scan Info']
    info = {}
    for row in ws_info.iter_rows(min_row=2, values_only=True):
        if row[0] and row[1]:
            info[str(row[0])] = str(row[1])

    sheet_map = {
        'T0': '👀 T0 — Pre-Recovery',
        'T1': '🌱 T1 — 10W Cross 20W',
        'T2': '🔥 T2 — 10W Cross 40W',
        'T3': '🚀 T3 — 20W Cross 40W',
        'COMBO': '⚡ COMBO T2+T3',
    }
    data = {}
    for key, sheetname in sheet_map.items():
        if sheetname in wb.sheetnames:
            ws = wb[sheetname]
            rows = list(ws.iter_rows(values_only=True))
            if len(rows) < 2:
                data[key] = pd.DataFrame()
                continue
            headers = [str(h) if h is not None else f'_col{i}' for i, h in enumerate(rows[0])]
            df = pd.DataFrame(rows[1:], columns=headers)
            if headers[0].startswith('_col'):
                df = df.iloc[:, 1:]
            df = df[df['Symbol'].notna()]
            data[key] = df
        else:
            data[key] = pd.DataFrame()
    return data, info

data, info = load_data()

if data is None:
    st.error("⚠️ No scan data found. Please upload bottom_bounce.xlsx to the GitHub repo.")
    st.stop()

scan_date = info.get('Scan Date', '—')
run_time = info.get('Run Time', '—')
counts = {k: len(data.get(k, pd.DataFrame())) for k in ['T0','T1','T2','T3','COMBO']}
total = sum(counts[k] for k in ['T0','T1','T2','T3'])

st.markdown('<div style="height:1rem"></div>', unsafe_allow_html=True)

# ── SCAN HEADER ───────────────────────────────────────────────
st.markdown(f"""
<div class="scan-header">
  <div class="scan-title">📊 Bottom Bounce — EMA Crossover Scanner</div>
  <div class="scan-meta">
    <span>📅 <b>{scan_date}</b></span>
    <span>🕐 <b>{run_time} IST</b></span>
    <span>🌐 <b>1304 stocks</b></span>
    <span>🎯 <b>{total} signals</b></span>
  </div>
</div>
""", unsafe_allow_html=True)

# ── METRICS ───────────────────────────────────────────────────
st.markdown(f"""
<div class="metrics">
  <div class="m-card"><div class="m-num">{counts['T0']}</div><div class="m-label">👀 T0 PRE-RECOVERY</div></div>
  <div class="m-card"><div class="m-num">{counts['T1']}</div><div class="m-label">🌱 T1 EARLY CROSS</div></div>
  <div class="m-card"><div class="m-num">{counts['T2']}</div><div class="m-label">🔥 T2 MOMENTUM</div></div>
  <div class="m-card"><div class="m-num">{counts['T3']}</div><div class="m-label">🚀 T3 TREND RESTORED</div></div>
  <div class="m-card combo"><div class="m-num m-combo">{counts['COMBO']}</div><div class="m-label">⚡ COMBO T2+T3</div></div>
</div>
""", unsafe_allow_html=True)

# ── TABS ──────────────────────────────────────────────────────
stage_keys = ['T0','T1','T2','T3','COMBO']
stage_desc = {
    'T0': 'Pre-Recovery — Gap narrowing, 10W & 20W still below 40W',
    'T1': 'Early Cross — 10W has crossed above 20W this week',
    'T2': 'Momentum — 10W has crossed above 40W',
    'T3': 'Trend Restored — 20W has crossed above 40W',
    'COMBO': 'High Conviction — Both T2 + T3 triggered this week',
}
tabs = st.tabs([
    f"👀 T0  ({counts['T0']})",
    f"🌱 T1  ({counts['T1']})",
    f"🔥 T2  ({counts['T2']})",
    f"🚀 T3  ({counts['T3']})",
    f"⚡ COMBO  ({counts['COMBO']})",
])

for tab, key in zip(tabs, stage_keys):
    with tab:
        df = data.get(key, pd.DataFrame())
        if df.empty:
            st.info("No signals for this stage today.")
        else:
            st.caption(f"_{stage_desc[key]}_")
            col1, col2 = st.columns([3,1])
            with col1:
                search = st.text_input("", key=f"s_{key}", placeholder="🔍 Search symbol...")
            if search:
                df = df[df['Symbol'].astype(str).str.contains(search.upper(), na=False)]
            st.dataframe(df, height=min(550, 55 + len(df)*35), hide_index=True)
            buf = BytesIO()
            df.to_excel(buf, index=False)
            with col2:
                st.download_button(
                    f"⬇ Export {key}",
                    data=buf.getvalue(),
                    file_name=f"BB_{key}_{scan_date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{key}"
                )

st.markdown('<div class="footer">SWINGEDGE PRO · CONFIDENTIAL · FOR MEMBER USE ONLY · © 2026</div>', unsafe_allow_html=True)
