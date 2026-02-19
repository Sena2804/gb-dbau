"""Styles, thème et CSS pour l'application CNBAU."""

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
    """Retourne le bloc HTML <link> + <style> complet pour le thème courant."""
    light_overrides = '''
    /* --- Streamlit native overrides for light mode --- */
    [data-testid="stApp"] {
        background-color: #ffffff !important;
    }
    [data-testid="stAppViewContainer"] {
        background-color: #ffffff !important;
        color: #24292f !important;
    }
    [data-testid="stSidebar"] {
        background-color: #f6f8fa !important;
        border-right-color: #d0d7de !important;
    }
    [data-testid="stSidebar"] [data-testid="stSidebarContent"] {
        background-color: #f6f8fa !important;
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
    /* --- Buttons --- */
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
    /* Icon-only action buttons in the table */
    button[data-testid="stBaseButton-secondary"] {
        background-color: #f6f8fa !important;
        border-color: #d0d7de !important;
    }

    /* --- Inputs, selects, dropdowns --- */
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
    /* Selectbox & multiselect */
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
    /* Dropdown menu */
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
    /* Multiselect tags */
    [data-baseweb="tag"] {
        background-color: #e1e4e8 !important;
        color: #24292f !important;
    }

    /* --- Progress bar --- */
    [data-testid="stProgress"] [role="progressbar"] > div {
        background-color: #d0d7de !important;
    }
    [data-testid="stProgress"] [role="progressbar"] > div > div {
        background-color: transparent !important;
    }
    [data-testid="stProgress"] [role="progressbar"] > div > div > div {
        background-color: #0969da !important;
    }

    /* --- Info / success / warning / error boxes --- */
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
    /* Success (Favorable) */
    .stAlertContainer[class*="st-bb"],
    div[data-testid="stAlert"]:has(.stAlertContainer[class*="st-bb"]) .stAlertContainer {
        background-color: rgba(22,128,57,0.1) !important;
        color: #1a7f37 !important;
    }
    /* Error (Défavorable) */
    .stAlertContainer[class*="st-bc"],
    div[data-testid="stAlert"]:has(.stAlertContainer[class*="st-bc"]) .stAlertContainer {
        background-color: rgba(207,34,46,0.1) !important;
        color: #cf222e !important;
    }

    /* --- File uploader --- */
    [data-testid="stFileUploader"] {
        background-color: #ffffff !important;
    }
    [data-testid="stFileUploader"] section {
        background-color: #f6f8fa !important;
        border-color: #d0d7de !important;
    }

    /* Restore semantic colors that the blanket rule above would override */
    .badge-favorable, .badge-favorable .ms { color: #3fb950 !important; }
    .badge-defavorable, .badge-defavorable .ms { color: #f85149 !important; }
    .badge-attente, .badge-attente .ms { color: #8b949e !important; }
    .quota-full .quota-text { color: #f85149 !important; }
    .kpi-card.kpi-green .kpi-icon, .kpi-card.kpi-green .kpi-icon .ms,
    .kpi-card.kpi-green .kpi-value { color: #3fb950 !important; }
    .kpi-card.kpi-red .kpi-icon, .kpi-card.kpi-red .kpi-icon .ms,
    .kpi-card.kpi-red .kpi-value { color: #f85149 !important; }
    .kpi-card.kpi-muted .kpi-icon, .kpi-card.kpi-muted .kpi-icon .ms,
    .kpi-card.kpi-muted .kpi-value { color: #57606a !important; }
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
    ''' if light else ''

    return f"""
<link href="https://fonts.googleapis.com/css2?family=Material+Symbols+Rounded:opsz,wght,FILL,GRAD@24,400,1,0" rel="stylesheet">
<style>
    /* --- Theme variables (driven by Python toggle) --- */
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

    /* --- Reset & base --- */
    .block-container {{ padding-top: 1.5rem; max-width: 1200px; }}

    .ms {{ font-family: 'Material Symbols Rounded'; font-size: 20px;
           vertical-align: middle; margin-right: 6px; font-weight: normal;
           font-style: normal; display: inline-block; line-height: 1;
           text-transform: none; letter-spacing: normal; word-wrap: normal;
           white-space: nowrap; direction: ltr;
           -webkit-font-smoothing: antialiased; }}

    /* --- Section headers --- */
    .section-header {{
        display: flex; align-items: center; gap: 10px;
        margin-bottom: 0.8rem; margin-top: 0.5rem;
    }}
    .section-header .ms {{ font-size: 26px; color: var(--accent); }}
    .section-header h3 {{ margin: 0; font-size: 1.3rem; font-weight: 600; color: var(--text-primary); }}

    /* --- Page title --- */
    .app-title {{
        display: flex; align-items: center; gap: 14px;
        padding-bottom: 1rem; border-bottom: 1px solid var(--border-subtle);
        margin-bottom: 1.2rem;
    }}
    .app-title .ms {{ font-size: 36px; color: var(--accent); }}
    .app-title h1 {{ margin: 0; font-size: 1.8rem; font-weight: 700; color: var(--text-primary); }}
    .app-title span.sub {{ font-size: 0.95rem; color: var(--text-muted); font-weight: 400; display: block; }}

    /* --- KPI cards --- */
    .kpi-row {{ display: flex; gap: 12px; margin-bottom: 0.5rem; }}
    .kpi-card {{
        flex: 1; background: var(--bg-card); border-radius: 12px; padding: 1.1rem 1rem;
        text-align: center; border: 1px solid var(--border);
        transition: border-color 0.2s;
    }}
    .kpi-card:hover {{ border-color: var(--accent); }}
    .kpi-card .kpi-icon {{ font-size: 22px; margin-bottom: 4px; color: var(--text-primary); }}
    .kpi-card .kpi-value {{ font-size: 2rem; font-weight: 700; margin: 2px 0; color: var(--text-primary); }}
    .kpi-card .kpi-label {{ font-size: 0.82rem; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.5px; }}
    .kpi-card.kpi-green .kpi-icon, .kpi-card.kpi-green .kpi-value {{ color: #3fb950 !important; }}
    .kpi-card.kpi-red .kpi-icon, .kpi-card.kpi-red .kpi-value {{ color: #f85149 !important; }}
    .kpi-card.kpi-muted .kpi-icon, .kpi-card.kpi-muted .kpi-value {{ color: var(--text-muted) !important; }}

    /* --- Number badges --- */
    .num-badge {{
        display: inline-flex; align-items: center; justify-content: center;
        min-width: 28px; padding: 2px 8px; border-radius: 8px;
        font-size: 0.85rem; font-weight: 700;
        background: var(--accent); color: #ffffff;
    }}

    /* --- Status badges --- */
    .badge {{
        display: inline-flex; align-items: center; gap: 4px;
        padding: 4px 12px; border-radius: 20px; font-size: 0.85rem; font-weight: 600;
    }}
    .badge .ms {{ font-size: 16px; margin-right: 2px; }}
    .badge-favorable {{ background: rgba(63,185,80,0.15); color: #3fb950; border: 1px solid rgba(63,185,80,0.3); }}
    .badge-defavorable {{ background: rgba(248,81,73,0.15); color: #f85149; border: 1px solid rgba(248,81,73,0.3); }}
    .badge-attente {{ background: rgba(139,148,158,0.15); color: var(--text-muted); border: 1px solid rgba(139,148,158,0.3); }}

    /* --- Quota cards --- */
    .quota-grid {{ display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 0.8rem; }}
    .quota-card {{
        flex: 1; min-width: 160px; border-radius: 10px; padding: 0.75rem 1rem;
        border: 1px solid var(--border); background: var(--bg-card);
    }}
    .quota-card .quota-filiere {{ font-weight: 600; font-size: 0.9rem; color: var(--text-primary); margin-bottom: 4px; }}
    .quota-card .quota-bar {{ height: 6px; border-radius: 3px; background: var(--bg-dark); margin: 6px 0; overflow: hidden; }}
    .quota-card .quota-bar-fill {{ height: 100%; border-radius: 3px; transition: width 0.3s; }}
    .quota-card .quota-text {{ font-size: 0.8rem; color: var(--text-muted); }}
    .quota-ok {{ border-color: rgba(63,185,80,0.3); }}
    .quota-ok .quota-bar-fill {{ background: #3fb950; }}
    .quota-full {{ border-color: rgba(248,81,73,0.4); }}
    .quota-full .quota-bar-fill {{ background: #f85149; }}
    .quota-full .quota-text {{ color: #f85149; }}

    /* --- Candidat info card --- */
    .candidat-card {{
        background: var(--bg-card); border: 1px solid var(--border); border-radius: 12px;
        padding: 1.2rem; margin-bottom: 1rem;
    }}
    .candidat-card table {{ width: 100%; border-collapse: collapse; }}
    .candidat-card td {{ padding: 6px 8px; font-size: 0.95rem; color: var(--text-primary); border: none; }}
    .candidat-card td:first-child {{ color: var(--text-muted); font-weight: 500; width: 120px; }}

    /* --- Niveau label --- */
    .niveau-label {{
        font-size: 0.85rem; font-weight: 600; color: var(--accent);
        text-transform: uppercase; letter-spacing: 0.8px;
        margin-bottom: 6px; margin-top: 8px;
    }}

    /* --- Alert quota --- */
    .alert-quota {{
        display: flex; align-items: center; gap: 8px;
        background: rgba(248,81,73,0.1); border: 1px solid rgba(248,81,73,0.3);
        border-radius: 8px; padding: 10px 14px; color: #f85149; font-size: 0.9rem;
    }}
    .alert-quota .ms {{ font-size: 20px; }}

    /* --- Hide default Streamlit header/footer --- */
    header[data-testid="stHeader"] {{ background: transparent; }}
    .stDeployButton {{ display: none; }}
    div[data-testid="stDataFrame"] {{ font-size: 1rem; }}

    /* --- Sticky header (appliqué via JS) --- */
    .sticky-header {{
        position: fixed !important;
        top: 0;
        left: 0;
        right: 0;
        z-index: 999;
        background: var(--bg-page);
        padding: 1rem 2rem 0.5rem 2rem;
        border-bottom: 1px solid var(--border);
        box-shadow: 0 2px 8px var(--shadow);
    }}
    .sticky-header .app-title h1 {{ font-size: 1.3rem; }}
    .sticky-header .app-title .sub {{ display: none; }}
    .sticky-header .kpi-row {{ margin-bottom: 0; }}
    .sticky-header .kpi-card {{ padding: 0.5rem 0.6rem; }}
    .sticky-header .kpi-value {{ font-size: 1.3rem; }}

    /* --- Action buttons: green / red / yellow by position in 3-col block --- */
    div[data-testid="stColumn"]:last-child div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:nth-child(1) button:not(:disabled) span[data-testid="stIconMaterial"] {{
        color: #3fb950 !important;
    }}
    div[data-testid="stColumn"]:last-child div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:nth-child(2) button:not(:disabled) span[data-testid="stIconMaterial"] {{
        color: #f85149 !important;
    }}
    div[data-testid="stColumn"]:last-child div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:nth-child(3) button:not(:disabled) span[data-testid="stIconMaterial"] {{
        color: #d29922 !important;
    }}
    /* --- Boutons désactivés : grisés --- */
    button:disabled {{
        opacity: 0.25 !important;
    }}
    /* --- Réduire l'espacement entre boutons d'action --- */
    div[data-testid="stColumn"]:last-child div[data-testid="stHorizontalBlock"] {{
        gap: 0.25rem !important;
    }}
    div[data-testid="stColumn"]:last-child div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {{
        padding: 0 !important;
    }}

    /* --- Sidebar --- */
    section[data-testid="stSidebar"] .stMarkdown h3 {{ font-size: 1.1rem; }}

    {light_overrides}
</style>
"""


def build_sticky_js(colors: dict) -> str:
    """Retourne le snippet JavaScript pour le header sticky."""
    return f"""
<script>
(function() {{
    const doc = window.parent.document;
    const titleEl = doc.querySelector('.app-title');
    const kpiEl = doc.querySelector('#kpi-row');
    if (!titleEl || !kpiEl) return;

    const existing = doc.getElementById('sticky-clone');
    if (existing) existing.remove();

    const header = doc.createElement('div');
    header.id = 'sticky-clone';
    header.style.cssText = `
        position: fixed; top: 0; left: 0; right: 0; z-index: 9999999;
        background: {colors["bg_page"]}; padding: 0.8rem 2rem 0.5rem 2rem;
        border-bottom: 1px solid {colors["border"]};
        box-shadow: 0 2px 8px {colors["shadow"]};
        display: none;
    `;
    header.innerHTML = `
        <div style="display:flex;align-items:center;gap:10px;margin-bottom:0.5rem">
            <span class="ms" style="font-size:24px;color:{colors["accent"]};font-family:'Material Symbols Rounded'">school</span>
            <span style="font-size:1.2rem;font-weight:700;color:{colors["text_primary"]}">CNBAU — Bourse de Russie</span>
        </div>
    `;
    const kpiClone = kpiEl.cloneNode(true);
    kpiClone.removeAttribute('id');
    kpiClone.querySelectorAll('.kpi-card').forEach(c => {{
        c.style.padding = '0.4rem 0.5rem';
    }});
    kpiClone.querySelectorAll('.kpi-value').forEach(v => {{
        v.style.fontSize = '1.2rem';
    }});
    header.appendChild(kpiClone);
    const stApp = doc.querySelector('[data-testid="stApp"]') || doc.body;
    stApp.appendChild(header);

    const scrollContainer = doc.querySelector('[data-testid="stMainBlockContainer"]')
        || doc.querySelector('section[data-testid="stMain"]')
        || doc.querySelector('.main');
    const sentinel = titleEl.closest('[data-testid="element-container"]') || titleEl;

    function checkScroll() {{
        const rect = sentinel.getBoundingClientRect();
        header.style.display = rect.bottom < 0 ? 'block' : 'none';
    }}

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

    const observer = new MutationObserver(() => {{
        const freshKpi = doc.querySelector('#kpi-row');
        if (!freshKpi) return;
        const oldClone = header.querySelector('.kpi-row');
        if (oldClone) oldClone.remove();
        const newClone = freshKpi.cloneNode(true);
        newClone.removeAttribute('id');
        newClone.querySelectorAll('.kpi-card').forEach(c => {{
            c.style.padding = '0.4rem 0.5rem';
        }});
        newClone.querySelectorAll('.kpi-value').forEach(v => {{
            v.style.fontSize = '1.2rem';
        }});
        header.appendChild(newClone);
    }});
    observer.observe(kpiEl, {{ childList: true, subtree: true, characterData: true }});
}})();
</script>
"""