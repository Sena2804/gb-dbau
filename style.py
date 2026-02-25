import base64
from pathlib import Path

import streamlit as st


# ✅ OPTIMISATION : logo chargé UNE SEULE FOIS et mis en cache par Streamlit
# Sans ce cache, le fichier était relu + encodé en Base64 à chaque rerun.
@st.cache_data
def get_logo_b64() -> str:
    img_path = Path(__file__).resolve().parent / "assets" / "image.png"
    if img_path.exists():
        with open(img_path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return ""


NIVEAU_ORDER = ["Licence", "Master", "Doctorat", "Spécialisation"]


def get_colors(light: bool) -> dict:
    return {
        "bg_page":       "#ffffff"              if light else "#0e1117",
        "bg_card":       "#f6f8fa"              if light else "#161b22",
        "bg_dark":       "#e1e4e8"              if light else "#21262d",
        "text_primary":  "#24292f"              if light else "#e6edf3",
        "text_muted":    "#57606a"              if light else "#8b949e",
        "accent":        "#0969da"              if light else "#6C9FFF",
        "border":        "#d0d7de"              if light else "#30363d",
        "border_subtle": "rgba(0,0,0,0.06)"    if light else "rgba(255,255,255,0.08)",
        "shadow":        "rgba(0,0,0,0.1)"     if light else "rgba(0,0,0,0.3)",
    }


def build_css(colors: dict, light: bool) -> str:
    # On appelle la fonction cachée — pas de re-lecture disque à chaque rerun
    logo_data = get_logo_b64()

    watermark_css = f"""
    [data-testid="stAppViewContainer"]::after {{
        content: "";
        position: fixed;
        top: 0;
        left: 0;
        width: 100vw;
        height: 100vh;
        background-image: url("data:image/png;base64,{logo_data}");
        background-repeat: repeat;
        background-size: 700px;
        opacity: 0.04;
        pointer-events: none;
        z-index: 99999;
    }}
    """ if logo_data else ""

    light_overrides = '''
    /* --- Streamlit native overrides for light mode --- */
    [data-testid="stApp"] {
        background-color: #ffffff !important;
    }
    [data-testid="stAppViewContainer"] {
        background-color: #ffffff !important;
        color: #24292f !important;
    }
    [data-testid="stHeader"] {
        background: #ffffff !important;
    }
    [data-testid="stMainBlockContainer"] {
        background-color: #ffffff !important;
    }
    p, span, label, .stMarkdown, .stCaption, li, td, th, h1, h2, h3, h4 {
        color: #24292f !important;
    }
    [data-testid="stMetricValue"] {
        color: #24292f !important;
    }
    [data-testid="stMetricLabel"] p {
        color: #57606a !important;
    }
    .stSelectbox label, .stMultiSelect label,
    .stTextInput label {
        color: #24292f !important;
    }
    hr {
        border-color: #d0d7de !important;
    }
    button[kind="secondary"], button[data-testid="stBaseButton-secondary"] {
        background-color: #f6f8fa !important;
        color: #24292f !important;
        border-color: #d0d7de !important;
    }
    button[kind="secondary"]:hover, button[data-testid="stBaseButton-secondary"]:hover {
        background-color: #e1e4e8 !important;
        border-color: #d0d7de !important;
    }
    button[kind="primary"], button[data-testid="stBaseButton-primary"] {
        background-color: #0969da !important;
        color: #ffffff !important;
    }
    button[data-testid="stBaseButton-secondary"] {
        background-color: #f6f8fa !important;
        border-color: #d0d7de !important;
    }
    [data-testid="stTextInput"] input,
    [data-testid="stNumberInput"] input {
        background-color: #ffffff !important;
        color: #24292f !important;
        border-color: #d0d7de !important;
    }
    [data-testid="stTextInput"] input::placeholder {
        color: #8c959f !important;
        opacity: 1 !important;
    }
    [data-baseweb="select"] input::placeholder {
        color: #8c959f !important;
        opacity: 1 !important;
    }
    [data-baseweb="select"] {
        background-color: #ffffff !important;
    }
    [data-baseweb="select"] > div {
        background-color: #ffffff !important;
        color: #24292f !important;
        border-color: #d0d7de !important;
    }
    [data-baseweb="select"] span {
        color: #24292f !important;
    }
    [data-baseweb="select"] svg {
        fill: #57606a !important;
    }
    [data-baseweb="select"] div {
        color: #24292f !important;
    }
    [data-baseweb="popover"] {
        background-color: #ffffff !important;
        border-color: #d0d7de !important;
    }
    [data-baseweb="popover"] li {
        background-color: #ffffff !important;
        color: #24292f !important;
    }
    [data-baseweb="popover"] li:hover,
    [data-baseweb="popover"] li[aria-selected="true"] {
        background-color: #f6f8fa !important;
    }
    [data-baseweb="tag"] {
        background-color: #e1e4e8 !important;
        color: #24292f !important;
    }
    [data-testid="stProgress"] [role="progressbar"] > div {
        background-color: #d0d7de !important;
    }
    [data-testid="stProgress"] [role="progressbar"] > div > div {
        background-color: transparent !important;
    }
    [data-testid="stProgress"] [role="progressbar"] > div > div > div {
        background-color: #EAC100 !important;
    }
    [data-testid="stAlert"] {
        background-color: #f6f8fa !important;
        color: #24292f !important;
    }
    .stAlertContainer {
        background-color: rgba(61,157,243,0.1) !important;
        color: #0969da !important;
    }
    .stAlertContainer div { color: inherit !important; }
    .stAlertContainer p { color: #24292f !important; }
    [data-testid="stFileUploader"] {
        background-color: #ffffff !important;
    }
    [data-testid="stFileUploader"] section {
        background-color: #f6f8fa !important;
        border-color: #d0d7de !important;
    }
    [role="tooltip"],
    [role="tooltip"] > div,
    [role="tooltip"] div,
    [role="tooltip"] span,
    [role="tooltip"] p,
    [data-testid="stTooltipContent"],
    .stTooltipContent {
        background-color: #24292f !important;
        color: #ffffff !important;
    }
    .badge-favorable, .badge-favorable .ms { color: #3fb950 !important; }
    .badge-defavorable, .badge-defavorable .ms { color: #f85149 !important; }
    .badge-attente, .badge-attente .ms { color: #8b949e !important; }
    .badge-suppleant, .badge-suppleant .ms { color: #0969da !important; }
    .quota-full .quota-text { color: #3fb950 !important; }
    .kpi-card.kpi-green .kpi-icon, .kpi-card.kpi-green .kpi-icon .ms,
    .kpi-card.kpi-green .kpi-value { color: #3fb950 !important; }
    .kpi-card.kpi-red .kpi-icon, .kpi-card.kpi-red .kpi-icon .ms,
    .kpi-card.kpi-red .kpi-value { color: #f85149 !important; }
    .kpi-card.kpi-muted .kpi-icon, .kpi-card.kpi-muted .kpi-icon .ms,
    .kpi-card.kpi-muted .kpi-value { color: #57606a !important; }
    .kpi-card.kpi-blue .kpi-icon, .kpi-card.kpi-blue .kpi-icon .ms,
    .kpi-card.kpi-blue .kpi-value { color: #0969da !important; }
    .num-badge { color: #ffffff !important; background: #0969da !important; }
    .alert-quota, .alert-quota .ms { color: #f85149 !important; }
    .niveau-label { color: var(--accent) !important; }
    .section-header .ms { color: var(--accent) !important; }
    .app-title .ms { color: var(--accent) !important; }
    #sticky-clone span { color: #24292f !important; }
    #sticky-clone > div:first-child .ms { color: #0969da !important; }
    #sticky-clone .kpi-green .kpi-icon .ms, #sticky-clone .kpi-green .kpi-value { color: #3fb950 !important; }
    #sticky-clone .kpi-red .kpi-icon .ms, #sticky-clone .kpi-red .kpi-value { color: #f85149 !important; }
    #sticky-clone .kpi-muted .kpi-icon .ms, #sticky-clone .kpi-muted .kpi-value { color: #57606a !important; }
    #sticky-clone .kpi-blue .kpi-icon .ms, #sticky-clone .kpi-blue .kpi-value { color: #0969da !important; }
    .export-card-title { color: #24292f !important; }
    .export-card-desc { color: #57606a !important; }
    .transfer-card { background: #f6f8fa !important; border-color: #d0d7de !important; }
    .transfer-summary { background: #e1e4e8 !important; }
    .transfer-summary-title { color: #24292f !important; }
    ''' if light else ''

    return f"""
<link href="https://fonts.googleapis.com/css2?family=Material+Symbols+Rounded:opsz,wght,FILL,GRAD@24,400,1,0" rel="stylesheet">
<style>
    {watermark_css}
    :root {{
        --bg-page: {colors["bg_page"]};
        --bg-card: {colors["bg_card"]};
        --bg-dark: {colors["bg_dark"]};
        --text-primary: {colors["text_primary"]};
        --text-muted: {colors["text_muted"]};
        --accent: {colors["accent"]};
        --border: {colors["border"]};
        --border-subtle: {colors["border_subtle"]};
        --shadow: {colors["shadow"]};
    }}

    .block-container {{
        padding-top: 2.5rem !important;
        max-width: 100% !important;
        padding-left: 3rem !important;
        padding-right: 3rem !important;
        padding-bottom: 3rem !important;
    }}

    .ms {{
        font-family: 'Material Symbols Rounded';
        font-size: 24px;
        vertical-align: middle;
        margin-right: 8px;
        font-weight: normal;
        font-style: normal;
        display: inline-block;
        line-height: 1;
        text-transform: none;
        letter-spacing: normal;
        word-wrap: normal;
        white-space: nowrap;
        direction: ltr;
        -webkit-font-smoothing: antialiased;
    }}

    .app-title {{
        display: flex;
        align-items: center;
        gap: 18px;
        padding-bottom: 1.5rem;
        border-bottom: 2px solid var(--border-subtle);
        margin-bottom: 2rem;
    }}
    .app-title .ms {{ font-size: 48px; color: var(--accent); }}
    .app-title h1 {{
        margin: 0;
        font-size: 2.4rem;
        font-weight: 800;
        color: var(--text-primary);
        letter-spacing: -0.5px;
    }}
    .app-title span.sub {{
        font-size: 1.1rem;
        color: var(--text-muted);
        font-weight: 400;
        display: block;
        margin-top: 3px;
    }}

    .section-header {{
        display: flex;
        align-items: center;
        gap: 12px;
        margin-bottom: 1.5rem;
        margin-top: 0.8rem;
    }}
    .section-header .ms {{ font-size: 32px; color: var(--accent); }}
    .section-header h3 {{
        margin: 0;
        font-size: 1.6rem;
        font-weight: 700;
        color: var(--text-primary);
    }}

    .kpi-row {{
        display: flex;
        gap: 16px;
        margin-bottom: 1rem;
    }}
    .kpi-card {{
        flex: 1;
        background: var(--bg-card);
        border-radius: 16px;
        padding: 1.8rem 1.4rem;
        text-align: center;
        border: 2px solid var(--border);
        transition: border-color 0.2s, transform 0.15s, box-shadow 0.15s;
        box-shadow: 0 2px 8px var(--shadow);
    }}
    .kpi-card:hover {{
        border-color: var(--accent);
        transform: translateY(-2px);
        box-shadow: 0 6px 20px var(--shadow);
    }}
    .kpi-card .kpi-icon {{ font-size: 30px; margin-bottom: 8px; color: var(--text-primary); }}
    .kpi-card .kpi-value {{
        font-size: 3rem;
        font-weight: 800;
        margin: 4px 0;
        color: var(--text-primary);
        line-height: 1;
    }}
    .kpi-card .kpi-label {{
        font-size: 0.95rem;
        color: var(--text-muted);
        text-transform: uppercase;
        letter-spacing: 1px;
        font-weight: 600;
        margin-top: 6px;
    }}
    .kpi-card.kpi-green .kpi-icon, .kpi-card.kpi-green .kpi-value {{ color: #3fb950 !important; }}
    .kpi-card.kpi-red .kpi-icon, .kpi-card.kpi-red .kpi-value {{ color: #f85149 !important; }}
    .kpi-card.kpi-muted .kpi-icon, .kpi-card.kpi-muted .kpi-value {{ color: var(--text-muted) !important; }}
    .kpi-card.kpi-blue .kpi-icon, .kpi-card.kpi-blue .kpi-icon .ms,
    .kpi-card.kpi-blue .kpi-value {{ color: #58a6ff !important; }}

    [data-testid="stTabs"] [role="tablist"] {{
        display: flex !important;
        justify-content: space-between !important;
        width: 100% !important;
        gap: 10px !important;
    }}
    [data-testid="stTabs"] [role="tab"] {{
        flex: 1 !important;
        height: auto !important;
        padding: 1rem 0.5rem !important;
        font-size: 1.1rem !important;
        font-weight: 600 !important;
    }}
    [data-testid="stTabs"] [role="tab"] p {{
        font-size: 1.5rem !important;
        font-weight: 700 !important;
        color: var(--text-primary) !important;
        transition: color 0.3s ease !important;
    }}
    [data-testid="stTabs"] [role="tab"] [data-testid="stIconMaterial"] {{
        font-size: 1.8rem !important;
        margin-right: 10px !important;
        transition: color 0.3s ease !important;
    }}
    [data-testid="stTabs"] [aria-selected="true"] p,
    [data-testid="stTabs"] [aria-selected="true"] [data-testid="stIconMaterial"] {{
        color: var(--accent) !important;
    }}
    [data-testid="stTabs"] [role="tab"]:hover p,
    [data-testid="stTabs"] [role="tab"]:hover [data-testid="stIconMaterial"] {{
        color: var(--accent) !important;
        opacity: 0.8;
    }}

    .num-badge {{
        display: inline-flex;
        align-items: center;
        justify-content: center;
        min-width: 38px;
        height: 38px;
        padding: 2px 10px;
        border-radius: 10px;
        font-size: 1rem;
        font-weight: 800;
        background: var(--accent);
        color: #ffffff;
        box-shadow: 0 2px 6px rgba(9,105,218,0.3);
    }}

    .badge {{
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 6px 16px;
        border-radius: 24px;
        font-size: 1rem;
        font-weight: 700;
    }}
    .badge .ms {{ font-size: 18px; margin-right: 3px; }}
    .badge-favorable  {{ background: rgba(63,185,80,0.15);  color: #3fb950; border: 2px solid rgba(63,185,80,0.4); }}
    .badge-defavorable{{ background: rgba(248,81,73,0.15);  color: #f85149; border: 2px solid rgba(248,81,73,0.4); }}
    .badge-attente    {{ background: rgba(139,148,158,0.15);color: var(--text-muted); border: 2px solid rgba(139,148,158,0.4); }}
    .badge-suppleant  {{ background: rgba(88,166,255,0.15); color: #58a6ff; border: 2px solid rgba(88,166,255,0.4); }}

    .candidate-name {{
        font-weight: 700;
        font-size: 1.1rem;
        color: var(--text-primary);
        line-height: 1.3;
    }}
    .candidate-sub {{
        color: var(--text-muted);
        font-size: 0.92rem;
        margin-top: 3px;
        font-weight: 500;
    }}
    .moyenne-txt {{
        font-family: 'Courier New', monospace;
        font-weight: 800;
        font-size: 1.15rem;
        color: var(--accent);
        background: rgba(9,105,218,0.08);
        padding: 4px 12px;
        border-radius: 6px;
        border: 1px solid rgba(9,105,218,0.15);
    }}

    .quota-grid {{
        display: flex;
        gap: 14px;
        flex-wrap: wrap;
        margin-bottom: 1.2rem;
    }}
    .quota-card {{
        flex: 1;
        min-width: 220px;
        border-radius: 14px;
        padding: 1.5rem !important;
        border: 2px solid var(--border);
        background: var(--bg-card);
        box-shadow: 0 2px 6px var(--shadow);
    }}
    .quota-card .quota-filiere {{
        font-weight: 800;
        font-size: 1.2rem !important;
        color: var(--text-primary);
        margin-bottom: 10px;
        line-height: 1.2;
    }}
    .quota-card .quota-text {{
        font-size: 1.1rem !important;
        color: var(--text-muted);
        font-weight: 700;
        margin-top: 8px;
    }}
    .quota-card .quota-bar {{
        height: 12px !important;
        border-radius: 6px;
        background: var(--bg-dark);
        margin: 12px 0;
        overflow: hidden;
    }}
    .quota-card .quota-bar-fill {{
        height: 100%;
        border-radius: 4px;
        transition: width 0.4s ease;
    }}
    .quota-ok {{ border-color: rgba(9,105,218,0.35); }}
    .quota-ok .quota-bar-fill {{ background: #0969da; }}
    .quota-full {{ border-color: rgba(63,185,80,0.5); }}
    .quota-full .quota-bar-fill {{ background: #3fb950; }}
    .quota-full .quota-text {{ color: #3fb950; }}

    .candidat-card {{
        background: var(--bg-card);
        border: 2px solid var(--border);
        border-radius: 16px;
        padding: 1.8rem;
        margin-bottom: 1.2rem;
        box-shadow: 0 4px 12px var(--shadow);
    }}
    .candidat-card table {{ width: 100%; border-collapse: collapse; }}
    .candidat-card td {{
        padding: 9px 10px;
        font-size: 1.05rem;
        color: var(--text-primary);
        border: none;
    }}
    .candidat-card td:first-child {{
        color: var(--text-muted);
        font-weight: 600;
        width: 140px;
        font-size: 0.95rem;
    }}

    .export-card {{
        background: var(--bg-card);
        border: 2px solid var(--border);
        border-radius: 14px;
        padding: 1.2rem 1.5rem;
        margin-bottom: 1rem;
        box-shadow: 0 2px 6px var(--shadow);
        transition: border-color 0.2s, box-shadow 0.15s;
    }}
    .export-card:hover {{
        border-color: var(--accent);
        box-shadow: 0 4px 14px var(--shadow);
    }}
    .export-card-header {{ display: flex; align-items: flex-start; gap: 14px; }}
    .export-card-title {{
        font-size: 1.2rem;
        font-weight: 700;
        color: var(--text-primary);
        margin-bottom: 4px;
    }}
    .export-card-desc {{
        font-size: 0.95rem;
        color: var(--text-muted);
        line-height: 1.4;
    }}

    .niveau-label {{
        font-size: 1.4rem !important;
        font-weight: 800;
        color: var(--accent);
        text-transform: uppercase;
        letter-spacing: 2px;
        margin-bottom: 15px;
        margin-top: 25px;
        border-left: 5px solid var(--accent);
        padding-left: 15px;
    }}

    .transfer-card {{
        background: var(--bg-card);
        border: 2px solid var(--border);
        border-radius: 14px;
        padding: 1.5rem;
    }}
    .transfer-summary {{
        background: var(--bg-dark);
        border-radius: 10px;
        padding: 1rem;
        margin-top: 1rem;
    }}
    .transfer-summary-title {{
        font-weight: 700;
        font-size: 1rem;
        margin-bottom: 8px;
        color: var(--text-primary);
    }}

    .alert-quota {{
        display: flex;
        align-items: center;
        gap: 10px;
        background: rgba(248,81,73,0.1);
        border: 2px solid rgba(248,81,73,0.35);
        border-radius: 10px;
        padding: 14px 18px;
        color: #f85149;
        font-size: 1rem;
        font-weight: 600;
    }}
    .alert-quota .ms {{ font-size: 22px; }}

    [data-testid="stTextInput"] input,
    [data-testid="stNumberInput"] input {{
        font-size: 1.05rem !important;
        padding: 0.6rem 0.9rem !important;
        min-height: 44px !important;
        border: 1px solid var(--border) !important;
        border-bottom: 1px solid var(--border) !important;
        border-radius: 8px !important;
        background-color: var(--bg-page) !important;
        color: var(--text-primary) !important;
    }}
    [data-testid="stTextInput"] input:hover,
    [data-testid="stNumberInput"] input:hover {{
        border-color: var(--text-primary) !important;
        border-bottom-color: var(--text-primary) !important;
    }}
    [data-testid="stTextInput"] input:focus,
    [data-testid="stNumberInput"] input:focus {{
        border-color: var(--accent) !important;
        border-bottom-color: var(--accent) !important;
        box-shadow: 0 0 0 3px rgba(9,105,218,0.15) !important;
    }}
    [data-baseweb="select"] > div {{ min-height: 44px !important; }}
    label {{ font-size: 1rem !important; font-weight: 600 !important; }}

    button[data-testid^="stBaseButton"] {{
        font-size: 1rem !important;
        font-weight: 600 !important;
        min-height: 42px !important;
        padding: 0.5rem 1rem !important;
        border-radius: 10px !important;
    }}
    button[kind="primary"], button[data-testid="stBaseButton-primary"] {{
        background-color: #0969da !important;
        color: #ffffff !important;
        box-shadow: 0 3px 10px rgba(9,105,218,0.35) !important;
    }}

    div[data-testid="stColumn"]:last-child div[data-testid="stHorizontalBlock"] button {{
        transition: all 0.2s ease-in-out !important;
        border: 1px solid transparent !important;
        background-color: transparent !important;
        border-radius: 8px !important;
    }}
    div[data-testid="stColumn"]:last-child div[data-testid="stHorizontalBlock"] button:hover:not(:disabled) {{
        transform: translateY(-2px);
        box-shadow: 0 4px 8px var(--shadow) !important;
    }}
    div[data-testid="stColumn"]:last-child div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:nth-child(1) button span[data-testid="stIconMaterial"] {{
        color: #3fb950 !important;
    }}
    div[data-testid="stColumn"]:last-child div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:nth-child(1) button:hover:not(:disabled) {{
        background-color: rgba(63, 185, 80, 0.1) !important;
        border-color: rgba(63, 185, 80, 0.4) !important;
    }}
    div[data-testid="stColumn"]:last-child div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:nth-child(2) button span[data-testid="stIconMaterial"] {{
        color: #f85149 !important;
    }}
    div[data-testid="stColumn"]:last-child div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:nth-child(2) button:hover:not(:disabled) {{
        background-color: rgba(248, 81, 73, 0.1) !important;
        border-color: rgba(248, 81, 73, 0.4) !important;
    }}
    div[data-testid="stColumn"]:last-child div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:nth-child(3) button span[data-testid="stIconMaterial"] {{
        color: #58a6ff !important;
    }}
    div[data-testid="stColumn"]:last-child div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:nth-child(3) button:hover:not(:disabled) {{
        background-color: rgba(88, 166, 255, 0.1) !important;
        border-color: rgba(88, 166, 255, 0.4) !important;
    }}
    div[data-testid="stColumn"]:last-child div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:nth-child(4) button span[data-testid="stIconMaterial"] {{
        color: #d29922 !important;
    }}
    div[data-testid="stColumn"]:last-child div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:nth-child(4) button:hover:not(:disabled) {{
        background-color: rgba(210, 153, 34, 0.1) !important;
        border-color: rgba(210, 153, 34, 0.4) !important;
    }}

    button:disabled {{
        opacity: 0.22 !important;
        filter: grayscale(1);
        cursor: not-allowed !important;
    }}
    div[data-testid="stColumn"]:last-child div[data-testid="stHorizontalBlock"] {{
        gap: 0.5rem !important;
        align-items: center;
    }}

    [data-testid="stDivider"] hr {{
        margin: 0.6rem 0 !important;
        border-width: 1px !important;
        opacity: 0.5 !important;
    }}

    [data-testid="stProgress"] [role="progressbar"] > div > div > div {{
        background-color: #0969da !important;
    }}

    header[data-testid="stHeader"] {{ background: transparent; }}
    .stDeployButton {{ display: none; }}

    .sticky-header {{
        position: fixed !important;
        top: 0;
        left: 0;
        right: 0;
        z-index: 999;
        background: var(--bg-page);
        padding: 1rem 3rem 0.6rem 3rem;
        border-bottom: 2px solid var(--border);
        box-shadow: 0 2px 12px var(--shadow);
    }}

    section[data-testid="stSidebar"] .stMarkdown h3 {{ font-size: 1.2rem; }}
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span,
    section[data-testid="stSidebar"] label {{ font-size: 1rem !important; }}

    {light_overrides}
</style>
"""


def get_sidebar_style(progress_color: str) -> str:
    logo_data = get_logo_b64()
    return f"""
    <style>
        [data-testid="stSidebar"] {{
            background-color: rgba(0, 135, 81, 0.1) !important;
            backdrop-filter: blur(12px);
        }}
        [data-testid="stSidebarHeader"] {{
            background-image: url("data:image/png;base64,{logo_data}");
            background-repeat: no-repeat;
            background-size: contain;
            background-position: center;
            height: 120px;
            margin-bottom: 20px;
        }}
        [data-testid="stSidebar"] .stMarkdown p {{
            font-size: 1.2rem !important;
            line-height: 1.4 !important;
            color: #004d2e !important;
            font-weight: 600 !important;
        }}
        [data-testid="stSidebar"] label {{
            font-size: 1.1rem !important;
            color: #004d2e !important;
            font-weight: 700 !important;
        }}
        [data-testid="stSidebarCollapseButton"] {{ display: none; }}
        [data-testid="stSidebarContent"] {{
            padding-top: 0rem !important;
            background-color: transparent !important;
        }}
        div[data-testid="stProgress"] > div > div > div > div {{
            background-color: {progress_color} !important;
            transition: background-color 0.5s ease;
        }}
        [data-testid="stCaptionContainer"] {{
            font-weight: 700;
            letter-spacing: 1px;
            text-transform: uppercase;
            font-size: 1.1rem !important;
            color: rgba(0, 77, 46, 0.7) !important;
            margin-top: 1.2rem;
        }}
        [data-testid="stSidebar"] button[kind="secondary"] {{
            border: 2px solid #ff4b4b !important;
            color: #ff4b4b !important;
            background-color: rgba(255, 255, 255, 0.5) !important;
            transition: all 0.3s ease;
            border-radius: 10px;
            font-size: 1.7rem !important;
            font-weight: 700 !important;
            padding: 0.6rem 1rem !important;
            margin-top: auto;
        }}
        [data-testid="stSidebar"] button[kind="secondary"]:hover {{
            background-color: #ff4b4b !important;
            color: white !important;
            box-shadow: 0 4px 14px rgba(255, 75, 75, 0.35);
        }}
        .sidebar-status-container {{ padding: 1rem 0; }}
        .sidebar-title {{
            color: #004d2e !important;
            font-size: 20px !important;
            font-weight: 800 !important;
            margin-bottom: 1.5rem !important;
        }}
        .progression-row {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
        }}
        .progression-label {{
            font-size: 2.5rem !important;
            font-weight: 700 !important;
            color: #004d2e !important;
        }}
        .progression-pct {{
            font-size: 2.5rem !important;
            font-weight: 800 !important;
            color: #0969da !important;
        }}
        .progression-sub {{
            font-size: 1.2rem !important;
            color: #004d2e !important;
            opacity: 0.8;
            font-weight: 500;
        }}
    </style>
    """


def build_sticky_js(colors: dict) -> str:
    """
    ✅ OPTIMISATION : MutationObserver avec debounce de 150ms.
    Sans ça, chaque micro-modification du DOM (reruns rapides Streamlit)
    déclenchait une mise à jour immédiate → boucle infinie → crash WebSocket.
    """
    return f"""
<script>
(function() {{
    const doc = window.parent.document;
    const titleEl = doc.querySelector('.app-title');
    const kpiEl   = doc.querySelector('#kpi-row');
    if (!titleEl || !kpiEl) return;

    const existing = doc.getElementById('sticky-clone');
    if (existing) existing.remove();

    const header = doc.createElement('div');
    header.id = 'sticky-clone';
    header.style.cssText = `
        position: fixed; top: 0; left: 0; right: 0; z-index: 9999999;
        background: {colors["bg_page"]}; padding: 0.8rem 3rem 0.5rem 3rem;
        border-bottom: 2px solid {colors["border"]};
        box-shadow: 0 2px 10px {colors["shadow"]};
        display: none;
    `;
    header.innerHTML = `
        <div style="display:flex;align-items:center;gap:12px;margin-bottom:0.5rem">
            <span class="ms" style="font-size:28px;color:{colors["accent"]};font-family:'Material Symbols Rounded'">school</span>
            <span style="font-size:1.4rem;font-weight:800;color:{colors["text_primary"]}">CNBAU — Bourse de Russie</span>
        </div>
    `;

    const kpiClone = kpiEl.cloneNode(true);
    kpiClone.removeAttribute('id');
    kpiClone.querySelectorAll('.kpi-card').forEach(c => {{ c.style.padding = '0.5rem 0.7rem'; }});
    kpiClone.querySelectorAll('.kpi-value').forEach(v => {{
        v.style.fontSize  = '2.5rem';
        v.style.fontWeight = '900';
        v.style.color     = '{colors["accent"]}';
    }});
    header.appendChild(kpiClone);

    const stApp = doc.querySelector('[data-testid="stApp"]') || doc.body;
    stApp.appendChild(header);

    // --- Scroll sentinel ---
    const sentinel = titleEl.closest('[data-testid="element-container"]') || titleEl;
    function checkScroll() {{
        const rect = sentinel.getBoundingClientRect();
        header.style.display = rect.bottom < 0 ? 'block' : 'none';
    }}

    const scrollContainer = doc.querySelector('[data-testid="stMainBlockContainer"]')
        || doc.querySelector('section[data-testid="stMain"]')
        || doc.querySelector('.main');

    let el = scrollContainer;
    while (el && el !== doc.documentElement) {{
        if (el.scrollHeight > el.clientHeight) {{
            el.addEventListener('scroll', checkScroll);
            break;
        }}
        el = el.parentElement;
    }}
    window.parent.addEventListener('scroll', checkScroll, true);
    checkScroll();

    // ✅ DEBOUNCE 150ms : évite les mises à jour en rafale lors des reruns rapides
    let _debounceTimer = null;
    const observer = new MutationObserver(() => {{
        clearTimeout(_debounceTimer);
        _debounceTimer = setTimeout(() => {{
            const freshKpi = doc.querySelector('#kpi-row');
            if (!freshKpi) return;
            const oldClone = header.querySelector('.kpi-row');
            if (oldClone) oldClone.remove();
            const newClone = freshKpi.cloneNode(true);
            newClone.removeAttribute('id');
            newClone.querySelectorAll('.kpi-card').forEach(c => {{ c.style.padding = '0.5rem 0.7rem'; }});
            newClone.querySelectorAll('.kpi-value').forEach(v => {{ v.style.fontSize = '1.4rem'; }});
            header.appendChild(newClone);
        }}, 150);  // 150ms de délai — absorbe les reruns rapides
    }});

    observer.observe(kpiEl, {{ childList: true, subtree: true, characterData: true }});
}})();
</script>
"""