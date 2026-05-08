import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import requests
import warnings
warnings.filterwarnings('ignore')

st.set_page_config(
    page_title="SwingEdge Pro — Member Portal",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow:ital,wght@0,300;0,400;0,500;0,600;0,700;0,800;0,900;1,300;1,400&family=Barlow+Condensed:wght@300;400;500;600;700;800;900&display=swap');

* { margin: 0; padding: 0; box-sizing: border-box; }

:root {
    --black: #080808;
    --dark: #0e0e0e;
    --card: #141414;
    --card2: #1a1a1a;
    --border: #242424;
    --border2: #2e2e2e;
    --gold: #f5a623;
    --gold2: #ffd066;
    --white: #f0f0f0;
    --muted: #666666;
    --muted2: #999999;
    --green: #22c55e;
    --navy: #1a2035;
}

html, body, [class*="css"] {
    font-family: 'Barlow', sans-serif !important;
    background-color: var(--black) !important;
    color: var(--white) !important;
}
.stApp { background-color: var(--black) !important; }
#MainMenu, footer, header { visibility: hidden; }
section[data-testid="stSidebar"] { display: none !important; }

/* ── NAV ── */
.nav {
    position: fixed; top: 0; left: 0; right: 0;
    background: rgba(8,8,8,0.97);
    border-bottom: 1px solid var(--border);
    padding: 0 2.5rem;
    height: 60px;
    display: flex; align-items: center; justify-content: space-between;
    z-index: 9999;
    backdrop-filter: blur(20px);
}
.nav-brand {
    display: flex; align-items: center; gap: 10px;
}
.nav-logo {
    font-family: 'Barlow', sans-serif;
    font-size: 0.95rem;
    font-weight: 700;
    color: var(--white);
}
.nav-logo span { color: var(--muted2); font-weight: 400; }
.nav-link {
    font-size: 0.82rem; color: var(--muted2); cursor: pointer;
}
.nav-cta {
    background: var(--gold); color: #000;
    font-size: 0.78rem; font-weight: 700;
    padding: 6px 14px; border-radius: 4px; cursor: pointer;
}

/* ── LOGIN ── */
.login-page {
    min-height: 100vh;
    display: flex; align-items: center; justify-content: center;
    background: var(--black);
    padding-top: 60px;
}
.login-card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 4px;
    padding: 3rem 2.5rem;
    width: 100%; max-width: 420px;
    text-align: center;
    position: relative;
    overflow: hidden;
}
.login-card::before {
    content: '';
    position: absolute; top: 0; left: 0; right: 0; height: 2px;
    background: linear-gradient(90deg, transparent, var(--gold), transparent);
}
.login-wolf { font-size: 3rem; margin-bottom: 1rem; }
.login-brand {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 1.3rem; font-weight: 800;
    letter-spacing: 4px; text-transform: uppercase;
    margin-bottom: 0.25rem;
}
.login-brand span { color: var(--gold); }
.login-tag {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 0.62rem; letter-spacing: 2px;
    color: var(--muted); text-transform: uppercase;
    margin-bottom: 2.5rem;
}
.login-divider {
    width: 40px; height: 1px;
    background: var(--gold); margin: 0 auto 2rem;
}

/* ── HOME PAGE ── */
.home-hero {
    padding: 6rem 2.5rem 3rem;
    text-align: center;
}
.hero-tag {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 0.7rem; letter-spacing: 3px;
    color: var(--gold); text-transform: uppercase;
    margin-bottom: 1.5rem;
}
.hero-title {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 3.5rem; font-weight: 300;
    letter-spacing: 6px; text-transform: uppercase;
    line-height: 1; color: var(--white);
    margin-bottom: 0.5rem;
}
.hero-title strong { font-weight: 800; color: var(--gold); }
.hero-sub {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 0.75rem; letter-spacing: 4px;
    color: var(--muted); text-transform: uppercase;
    margin-top: 1.5rem;
}

/* ── SCAN GRID ── */
.scans-section {
    padding: 0 2.5rem 4rem;
    max-width: 1400px; margin: 0 auto;
}
.section-label {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 0.65rem; letter-spacing: 3px;
    color: var(--muted); text-transform: uppercase;
    margin-bottom: 1.5rem;
    padding-bottom: 0.75rem;
    border-bottom: 1px solid var(--border);
}
.home-hero .hero-tag {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 0.65rem; letter-spacing: 3px;
    color: var(--muted2); text-transform: uppercase;
    margin-bottom: 1.5rem;
}
.home-hero .hero-title {
    font-size: 2.8rem; font-weight: 800;
    line-height: 1.1; color: var(--white);
    margin-bottom: 0.1rem;
}
.home-hero .hero-title em {
    font-style: italic; color: var(--gold);
    font-weight: 800;
}
.home-hero .hero-sub {
    font-size: 0.9rem; color: var(--muted2);
    max-width: 500px; margin: 1.5rem auto 0;
    line-height: 1.6;
}
.features-row {
    display: flex; gap: 1rem;
    padding: 2rem 2.5rem 0;
    max-width: 1400px; margin: 0 auto;
}
.feat-card {
    flex: 1; background: var(--card);
    border: 1px solid var(--border);
    border-radius: 8px; padding: 1.5rem;
}
.feat-icon { font-size: 1.5rem; margin-bottom: 0.75rem; }
.feat-name { font-size: 0.85rem; font-weight: 700; margin-bottom: 0.5rem; }
.feat-desc { font-size: 0.75rem; color: var(--muted2); line-height: 1.5; }
.scan-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 1px;
    background: var(--border);
    border: 1px solid var(--border);
    border-radius: 4px;
    overflow: hidden;
    margin-bottom: 3rem;
}
.scan-card {
    background: var(--card);
    padding: 1.75rem;
    cursor: pointer;
    transition: background 0.2s;
    position: relative;
}
.scan-card:hover { background: var(--card2); }
.scan-card.active:hover { background: #1a1500; }
.scan-card.active { border-left: 2px solid var(--gold); }
.scan-card.coming { opacity: 0.5; cursor: default; }
.scan-icon { font-size: 1.5rem; margin-bottom: 1rem; }
.scan-name {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 1rem; font-weight: 700;
    letter-spacing: 1px; text-transform: uppercase;
    color: var(--white); margin-bottom: 0.4rem;
}
.scan-desc {
    font-size: 0.78rem; color: var(--muted2);
    line-height: 1.5; margin-bottom: 1rem;
}
.scan-status {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 0.6rem; letter-spacing: 2px;
    text-transform: uppercase; font-weight: 700;
}
.scan-status.live { color: var(--gold); }
.scan-status.soon { color: var(--muted); }
.scan-count {
    position: absolute; top: 1.75rem; right: 1.75rem;
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 1.5rem; font-weight: 800;
    color: var(--gold); opacity: 0.3;
}
.scan-arrow {
    position: absolute; bottom: 1.75rem; right: 1.75rem;
    color: var(--gold); font-size: 1rem; opacity: 0;
    transition: opacity 0.2s;
}
.scan-card.active:hover .scan-arrow { opacity: 1; }

/* ── DASHBOARD ── */
.dash-header {
    padding: 5rem 2.5rem 1.5rem;
}
.dash-back {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 0.7rem; letter-spacing: 2px;
    color: var(--muted); text-transform: uppercase;
    cursor: pointer; margin-bottom: 1.5rem;
    display: inline-block;
}
.dash-title {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 2rem; font-weight: 800;
    letter-spacing: 2px; text-transform: uppercase;
    color: var(--white); margin-bottom: 0.25rem;
}
.dash-title span { color: var(--gold); }
.dash-meta {
    font-size: 0.78rem; color: var(--muted2);
    font-family: 'Barlow Condensed', sans-serif;
    letter-spacing: 1px;
}

.metrics-strip {
    display: grid;
    grid-template-columns: repeat(5, 1fr);
    gap: 1px; background: var(--border);
    margin: 0 2.5rem 1px;
    border: 1px solid var(--border);
    border-radius: 4px 4px 0 0; overflow: hidden;
}
.metric {
    background: var(--card);
    padding: 1.25rem 1rem;
    text-align: center;
}
.metric.highlight { background: #110f00; }
.metric-n {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 2.5rem; font-weight: 900;
    color: var(--gold); line-height: 1;
}
.metric-n.gold2 { color: var(--gold2); }
.metric-l {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 0.6rem; letter-spacing: 2px;
    color: var(--muted); text-transform: uppercase;
    margin-top: 4px; font-weight: 600;
}

/* ── TABS ── */
.stTabs [data-baseweb="tab-list"] {
    background: var(--card) !important;
    border-bottom: 1px solid var(--gold) !important;
    gap: 0 !important; padding: 0 !important;
    border-radius: 0 !important;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    color: var(--muted2) !important;
    font-family: 'Barlow Condensed', sans-serif !important;
    font-weight: 700 !important;
    font-size: 0.78rem !important;
    letter-spacing: 1.5px !important;
    text-transform: uppercase !important;
    padding: 0.75rem 1.5rem !important;
    border-radius: 0 !important;
}
.stTabs [aria-selected="true"] {
    background: var(--gold) !important;
    color: #000 !important;
}
.stTabs [data-baseweb="tab-panel"] {
    background: var(--card) !important;
    border: 1px solid var(--border) !important;
    border-top: none !important;
    padding: 1.25rem !important;
    border-radius: 0 0 4px 4px !important;
}

/* ── INPUTS / BUTTONS ── */
.stTextInput input {
    background: var(--dark) !important;
    border: 1px solid var(--border2) !important;
    color: var(--white) !important;
    border-radius: 3px !important;
    font-family: 'Barlow', sans-serif !important;
    font-size: 0.85rem !important;
}
.stTextInput input:focus { border-color: var(--gold) !important; }
.stButton button {
    background: var(--gold) !important;
    color: #000 !important;
    font-family: 'Barlow Condensed', sans-serif !important;
    font-weight: 800 !important;
    letter-spacing: 2px !important;
    text-transform: uppercase !important;
    border: none !important;
    border-radius: 3px !important;
}
.stDownloadButton button {
    background: transparent !important;
    color: var(--gold) !important;
    border: 1px solid var(--gold) !important;
    font-family: 'Barlow Condensed', sans-serif !important;
    font-weight: 700 !important;
    letter-spacing: 1px !important;
    border-radius: 3px !important;
}

/* Fullscreen fix */
[data-testid="stFullScreenFrame"] {
    background: #080808 !important;
}
.fullscreen-wrapper {
    background: #080808 !important;
}
iframe[title="st.dataframe"] {
    background: #080808 !important;
}
.footer {
    text-align: center;
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 0.6rem; letter-spacing: 2px;
    color: var(--border2); text-transform: uppercase;
    padding: 2rem 0 3rem;
}
</style>
""", unsafe_allow_html=True)

# ── PASSWORD ──────────────────────────────────────────────────
PASSWORD = "SwingEdge@2026"

if "auth" not in st.session_state:
    st.session_state.auth = False
if "page" not in st.session_state:
    st.session_state.page = "home"

# ── NAV ──────────────────────────────────────────────────────
st.markdown("""
<div class="nav">
  <div class="nav-brand">
    <span style="color:#f5a623; font-size:0.6rem;">●</span>
    <span class="nav-logo">SwingEdge<span>Pro</span></span>
  </div>
  <div style="display:flex; gap:2rem; align-items:center;">
    <span class="nav-link">Indicator</span>
    <span class="nav-link">Why It Works</span>
    <span class="nav-link">The Proof</span>
  </div>
  <div style="display:flex; gap:1rem; align-items:center;">
    <span class="nav-link">Log in</span>
    <span class="nav-cta">Member Portal</span>
  </div>
</div>
""", unsafe_allow_html=True)

# ── LOGIN ─────────────────────────────────────────────────────
if not st.session_state.auth:
    st.markdown('<div style="height:12vh"></div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.1, 1])
    with col2:
        st.markdown("""
        <div class="login-card">
          <div style="font-size:0.6rem; color:#f5a623; letter-spacing:2px; margin-bottom:0.5rem;">● SWINGEDGEPRO.IN</div>
          <div class="login-brand">SwingEdge<span>Pro</span></div>
          <div class="login-tag">Member Access · Daily Scans</div>
          <div class="login-divider"></div>
        </div>
        """, unsafe_allow_html=True)
        pwd = st.text_input("", type="password", placeholder="Enter member password...", label_visibility="collapsed")
        if st.button("ACCESS PORTAL →", use_container_width=True):
            if pwd == PASSWORD:
                st.session_state.auth = True
                st.rerun()
            else:
                st.error("Incorrect password.")
        st.markdown("""
        <div style="text-align:center; margin-top:1.5rem;">
          <div style="font-family:'Barlow Condensed',sans-serif; font-size:0.6rem; letter-spacing:2px; color:#444; text-transform:uppercase;">
            ALIGN WITH MOMENTUM · EXECUTE WITH AN EDGE
          </div>
        </div>
        """, unsafe_allow_html=True)
    st.stop()

# ── SCAN CONFIG ───────────────────────────────────────────────
SCANS = [
    {
        "key": "indicator",
        "icon": "⚡",
        "name": "SwingEdge Pro Ultimate",
        "desc": "6-in-1 TradingView indicator. One glance to know if a stock deserves your watchlist.",
        "status": "live",
    },
    {
        "key": "bottom_bounce",
        "icon": "📊",
        "name": "Bottom Bounce",
        "desc": "EMA crossover recovery — stocks emerging from stage 1 base with volume confirmation",
        "status": "live",
        "sheets": {
            'T0': '👀 T0 — Pre-Recovery',
            'T1': '🌱 T1 — 10W Cross 20W',
            'T2': '🔥 T2 — 10W Cross 40W',
            'T3': '🚀 T3 — 20W Cross 40W',
            'COMBO': '⚡ COMBO T2+T3',
        },
        "tab_labels": lambda c: [f"T0 ({c['T0']})", f"T1 ({c['T1']})", f"T2 ({c['T2']})", f"T3 ({c['T3']})", f"COMBO ({c['COMBO']})"],
        "desc_map": {
            'T0': 'Pre-Recovery — Gap narrowing, EMAs still inverted',
            'T1': 'Early Cross — 10W crossed above 20W',
            'T2': 'Momentum — 10W crossed above 40W',
            'T3': 'Trend Restored — 20W crossed above 40W',
            'COMBO': 'High Conviction — T2 + T3 both fired this week',
        }
    },
    {"key": "rs_explosion", "icon": "🚀", "name": "RS Explosion", "desc": "Relative strength breakout — stocks showing institutional accumulation with RS surge above 70", "status": "soon"},
    {"key": "52w_high", "icon": "🏔️", "name": "52W High Breakout", "desc": "Stocks breaking out of multi-week bases near 52-week highs with volume expansion", "status": "soon"},
    {"key": "accumulation", "icon": "🏦", "name": "Accumulation", "desc": "Silent institutional buying — 3+ up days vs down days with 1.5x volume ratio", "status": "soon"},
    {"key": "ath_scanner", "icon": "⚡", "name": "ATH Scanner", "desc": "Stocks within 5% of all-time highs with Stage 2 structure and RS ≥ 70", "status": "soon"},
    {"key": "ema200_cross", "icon": "📈", "name": "200 EMA Cross", "desc": "Fresh crossovers above the 200 EMA with RS confirmation and volume surge", "status": "soon"},
    {"key": "deep_base", "icon": "🔭", "name": "Deep Base Recovery", "desc": "Long consolidation breakout — stocks recovering from extended Stage 1 bases", "status": "soon"},
    {"key": "50ema_shakeout", "icon": "💥", "name": "50 EMA Shakeout", "desc": "Stage 2 leaders retesting 50 EMA — high-conviction pullback entry setups", "status": "soon"},
]

# ── HOME PAGE ─────────────────────────────────────────────────
if st.session_state.page == "home":
    st.markdown("""
    <div class="home-hero">
      <div class="hero-tag">RECLAIM YOUR EDGE</div>
      <div class="hero-title">Designed for the <em>serious</em></div>
      <div class="hero-title">momentum trader.</div>
      <div class="hero-sub">We do the heavy lifting so you can focus on what matters — finding the right stocks at the right time.</div>
    </div>

    <div class="features-row">
      <div class="feat-card">
        <div class="feat-icon">🚫</div>
        <div class="feat-name">Eliminate Dead Stocks</div>
        <div class="feat-desc">Only stocks showing real strength, momentum and volume qualify. No more buying into weakness.</div>
      </div>
      <div class="feat-card">
        <div class="feat-icon">🔗</div>
        <div class="feat-name">Connect the Signals</div>
        <div class="feat-desc">6 powerful metrics combined into one indicator. EMA, RS, Volume, ADR% — all in one glance.</div>
      </div>
      <div class="feat-card">
        <div class="feat-icon">💡</div>
        <div class="feat-name">Instant Clarity</div>
        <div class="feat-desc">Know immediately if a stock deserves your attention. No more second-guessing or manual checks.</div>
      </div>
      <div class="feat-card">
        <div class="feat-icon">🎯</div>
        <div class="feat-name">Impulse + Tightness</div>
        <div class="feat-desc">Built around the exact setup that produces explosive moves — coiling before the breakout.</div>
      </div>
      <div class="feat-card">
        <div class="feat-icon">⏱️</div>
        <div class="feat-name">Save Hours Daily</div>
        <div class="feat-desc">What used to take hours of manual scanning now takes seconds. Your edge, automated.</div>
      </div>
    </div>

    <div style="height:3rem;"></div>
    """, unsafe_allow_html=True)

    # ── MEMBER LOGIN SECTION ──
    st.markdown('<div id="member-portal" style="padding: 3rem 2.5rem; text-align:center; border-top: 1px solid #242424;">', unsafe_allow_html=True)
    st.markdown("""
    <div style="font-family:'Barlow Condensed',sans-serif; font-size:0.65rem; letter-spacing:3px; color:#888; text-transform:uppercase; margin-bottom:1rem;">— Member Portal</div>
    <div style="font-family:'Barlow',sans-serif; font-size:1.5rem; font-weight:800; color:#f0f0f0; margin-bottom:0.5rem;">Access Your Daily Scans</div>
    <div style="font-size:0.82rem; color:#888; margin-bottom:2rem;">Enter your member password to view today's scan results</div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        pwd = st.text_input("", type="password", placeholder="Enter member password...", label_visibility="collapsed", key="home_pwd")
        if st.button("ACCESS PORTAL →", use_container_width=True, key="home_login"):
            if pwd == PASSWORD:
                st.session_state.auth = True
                st.rerun()
            else:
                st.error("Incorrect password.")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div style="padding: 0 1rem;">', unsafe_allow_html=True)
    st.markdown('<div class="section-label">— Scan Intelligence Suite</div>', unsafe_allow_html=True)

    # Render scan cards in 3-column grid using st.columns
    rows = [SCANS[i:i+3] for i in range(0, len(SCANS), 3)]
    for row in rows:
        cols = st.columns(3)
        for col, scan in zip(cols, row):
            with col:
                is_live = scan['status'] == 'live'
                border = "border-left: 3px solid #f5a623;" if is_live else "opacity: 0.5;"
                status = "● LIVE" if is_live else "○ COMING SOON"
                status_color = "#f5a623" if is_live else "#666"
                st.markdown(f"""
                <div style="background:#141414; {border} padding:1.5rem; border-radius:4px; margin-bottom:1px; min-height:180px;">
                  <div style="font-size:1.5rem; margin-bottom:0.75rem;">{scan['icon']}</div>
                  <div style="font-family:'Barlow Condensed',sans-serif; font-size:1rem; font-weight:800; letter-spacing:2px; text-transform:uppercase; color:#f0f0f0; margin-bottom:0.4rem;">{scan['name']}</div>
                  <div style="font-size:0.75rem; color:#888; line-height:1.5; margin-bottom:1rem;">{scan['desc']}</div>
                  <div style="font-family:'Barlow Condensed',sans-serif; font-size:0.6rem; letter-spacing:2px; font-weight:700; color:{status_color};">{status}</div>
                </div>
                """, unsafe_allow_html=True)
                if is_live:
                    if st.button(f"Open {scan['name']} →", key=f"open_{scan['key']}", use_container_width=True):
                        st.session_state.page = scan['key']
                        st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('<div class="footer">SwingEdgePro.in · Confidential · Member Use Only · © 2026</div>', unsafe_allow_html=True)

# ── INDICATOR PAGE ───────────────────────────────────────────
elif st.session_state.page == "indicator":

    if st.button("← Back to Portal", key="back_ind"):
        st.session_state.page = "home"
        st.rerun()

    st.markdown("""
    <div class="dash-header">
      <div style="font-family:'Barlow Condensed',sans-serif; font-size:0.65rem; letter-spacing:3px; color:#f5a623; text-transform:uppercase; margin-bottom:0.75rem;">+ SwingEdge Pro Indicator</div>
      <div style="font-size:2.5rem; font-weight:800; color:#f0f0f0; line-height:1.1; margin-bottom:0.5rem;">One indicator.<br><em style="color:#f5a623; font-style:italic;">Complete stock</em><br>intelligence.</div>
      <div style="font-size:0.85rem; color:#888; max-width:500px; margin-bottom:1.5rem; line-height:1.6;">Built on TradingView. Source protected. Every metric you need to qualify a stock for your watchlist — merged into one powerful, private indicator.</div>
      <div style="background:#1a1a1a; border-left:3px solid #f5a623; padding:1rem 1.25rem; border-radius:4px; max-width:500px; margin-bottom:2rem;">
        <em style="font-size:0.85rem; color:#ccc;">"The only indicator that tells you in one glance whether a stock deserves to be on your watchlist or not."</em>
      </div>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns([1, 1])
    with col2:
        st.markdown("""
        <div style="background:#141414; border:1px solid #2e2e2e; border-radius:12px; padding:2rem; margin-top:2rem;">
          <div style="display:flex; align-items:center; gap:0.75rem; margin-bottom:0.25rem;">
            <div style="width:4px; height:1.2rem; background:#f5a623; border-radius:2px;"></div>
            <div style="font-weight:800; font-size:1rem; color:#f0f0f0;">SwingEdge Pro Ultimate</div>
          </div>
          <div style="font-family:'Barlow Condensed',sans-serif; font-size:0.6rem; letter-spacing:2px; color:#f5a623; text-transform:uppercase; margin-bottom:1.5rem; padding-left:1rem;">+ 6-IN-1 INDICATOR</div>
          <div style="display:flex; flex-direction:column; gap:0;">
            %s
          </div>
          <div style="margin-top:1.5rem; padding-top:1rem; border-top:1px solid #242424; display:flex; justify-content:space-between; align-items:center;">
            <div style="font-size:0.72rem; color:#666;">🔒 Source Protected · Not for resale</div>
            <div style="font-size:0.78rem; font-weight:700; color:#f5a623;">SwingEdge Pro</div>
          </div>
        </div>
        """ % "".join([
            f"""<div style="display:flex; justify-content:space-between; align-items:center; padding:0.85rem 0; border-bottom:1px solid #1e1e1e;">
              <div style="display:flex; gap:0.75rem; align-items:center;">
                <span style="font-family:'Barlow Condensed',sans-serif; font-size:0.7rem; color:#555; font-weight:600;">0{i+1}</span>
                <span style="font-size:0.88rem; font-weight:600; color:#f0f0f0;">{name}</span>
              </div>
              <span style="color:#22c55e; font-size:0.8rem;">✓</span>
            </div>"""
            for i, name in enumerate(["Trend Levels", "Pace Quality", "Recovery %", "Relative Volume", "Money Flow Quality", "Strength Score"])
        ]), unsafe_allow_html=True)

    st.markdown('<div class="footer">SwingEdgePro.in · Confidential · Member Use Only · © 2026</div>', unsafe_allow_html=True)

# ── BOTTOM BOUNCE DASHBOARD ───────────────────────────────────
elif st.session_state.page == "bottom_bounce":

    @st.cache_data(ttl=300)
    def load_bb():
        try:
            api = requests.get("https://api.github.com/repos/SwingEdgeLab/swingEdge-dashboard/contents/").json()
            bb_files = sorted([f['name'] for f in api if isinstance(f, dict) and f['name'].startswith('SwingEdge_BottomBounce') and f['name'].endswith('.xlsx')], reverse=True)
            filename = bb_files[0] if bb_files else 'bottom_bounce.xlsx'
            url = f"https://github.com/SwingEdgeLab/swingEdge-dashboard/raw/main/{filename}"
            r = requests.get(url, headers={"Accept": "application/octet-stream"})
            wb = openpyxl.load_workbook(BytesIO(r.content), data_only=True)

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
                    headers = [str(h) if h is not None else f'_c{i}' for i, h in enumerate(rows[0])]
                    df = pd.DataFrame(rows[1:], columns=headers)
                    if headers[0].startswith('_c'):
                        df = df.iloc[:, 1:]
                    df = df[df['Symbol'].notna()]
                    data[key] = df
                else:
                    data[key] = pd.DataFrame()
            return data, info
        except Exception as e:
            return None, None

    if st.button("← Back to Portal", key="back"):
        st.session_state.page = "home"
        st.rerun()

    with st.spinner("Loading scan data..."):
        data, info = load_bb()

    if data is None:
        st.error("Could not load scan data. Please check GitHub repo.")
        st.stop()

    scan_date = info.get('Scan Date', '—')
    run_time = info.get('Run Time', '—')
    counts = {k: len(data.get(k, pd.DataFrame())) for k in ['T0','T1','T2','T3','COMBO']}
    total = sum(counts[k] for k in ['T0','T1','T2','T3'])

    st.markdown(f"""
    <div style="padding: 4.5rem 1rem 0.5rem; display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:0.5rem;">
      <div>
        <span style="font-family:'Barlow Condensed',sans-serif; font-size:1.1rem; font-weight:800; letter-spacing:2px; text-transform:uppercase; color:#f0f0f0;">📊 Bottom <span style='color:#f5a623;'>Bounce</span></span>
        <span style="font-size:0.72rem; color:#666; margin-left:1rem; font-family:'Barlow Condensed',sans-serif; letter-spacing:1px;">{scan_date} · {run_time} IST · 1304 stocks</span>
      </div>
      <div style="display:flex; gap:1.5rem; font-family:'Barlow Condensed',sans-serif; font-size:0.72rem; letter-spacing:1px;">
        <span style="color:#888;">T0 <b style="color:#f5a623;">{counts['T0']}</b></span>
        <span style="color:#888;">T1 <b style="color:#f5a623;">{counts['T1']}</b></span>
        <span style="color:#888;">T2 <b style="color:#f5a623;">{counts['T2']}</b></span>
        <span style="color:#888;">T3 <b style="color:#f5a623;">{counts['T3']}</b></span>
        <span style="color:#888;">COMBO <b style="color:#fbbf24;">{counts['COMBO']}</b></span>
        <span style="color:#888;">TOTAL <b style="color:#f5a623;">{total}</b></span>
      </div>
    </div>
    """, unsafe_allow_html=True)

    stage_keys = ['T0','T1','T2','T3','COMBO']
    desc_map = {
        'T0': 'Pre-Recovery — Gap narrowing, EMAs still inverted',
        'T1': 'Early Cross — 10W has crossed above 20W',
        'T2': 'Momentum — 10W has crossed above 40W',
        'T3': 'Trend Restored — 20W has crossed above 40W',
        'COMBO': 'High Conviction — T2 + T3 both fired this week',
    }

    with st.container():
        st.markdown('<div style="padding: 0 2.5rem;">', unsafe_allow_html=True)
        tabs = st.tabs([
            f"T0  ({counts['T0']})",
            f"T1  ({counts['T1']})",
            f"T2  ({counts['T2']})",
            f"T3  ({counts['T3']})",
            f"COMBO  ({counts['COMBO']})",
        ])
        for tab, key in zip(tabs, stage_keys):
            with tab:
                df = data.get(key, pd.DataFrame())
                if df.empty:
                    st.info("No signals for this stage today.")
                else:
                    st.caption(f"_{desc_map[key]}_")
                    c1, c2 = st.columns([4, 1])
                    with c1:
                        search = st.text_input("", key=f"s_{key}", placeholder="Search symbol...")
                    if search:
                        df = df[df['Symbol'].astype(str).str.contains(search.upper(), na=False)]
                    st.dataframe(df, height=min(900, max(400, len(df)*38)), hide_index=True, use_container_width=True)
                    buf = BytesIO()
                    df.to_excel(buf, index=False)
                    with c2:
                        st.download_button(f"Export {key}", data=buf.getvalue(),
                            file_name=f"BB_{key}_{scan_date}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_{key}")
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="footer">SwingEdgePro.in · Confidential · Member Use Only · © 2026</div>', unsafe_allow_html=True)
