"""Application Streamlit - Suivi de session CNBAU Bourse de Russie."""

import io
from html import escape as html_escape
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import database as db

st.set_page_config(
    page_title="CNBAU - Bourse de Russie",
    layout="wide",
)

NIVEAU_ORDER = ["Licence", "Master", "Doctorat", "Spécialisation"]

# Mode clair / sombre
if "_theme" not in st.session_state:
    st.session_state["_theme"] = "light"
light = st.session_state["_theme"] == "light"

COLORS = {
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
    /* Tooltips */
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
    [data-testid="stFileUploader"] small,
    [data-testid="stFileUploader"] span,
    [data-testid="stFileUploader"] div {
        color: #24292f !important;
    }
    [data-testid="stFileUploader"] button svg {
        fill: #57606a !important;
    }

    /* Restore semantic colors that the blanket rule above would override */
    .badge-favorable, .badge-favorable .ms { color: #3fb950 !important; }
    .badge-defavorable, .badge-defavorable .ms { color: #f85149 !important; }
    .badge-attente, .badge-attente .ms { color: #8b949e !important; }
    .quota-full .quota-text { color: #3fb950 !important; }
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
    .quota-ok {{ border-color: rgba(9,105,218,0.3); }}
    .quota-ok .quota-bar-fill {{ background: #0969da; }}
    .quota-full {{ border-color: rgba(63,185,80,0.4); }}
    .quota-full .quota-bar-fill {{ background: #3fb950; }}
    .quota-full .quota-text {{ color: #3fb950; }}

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


st.markdown(build_css(COLORS, light), unsafe_allow_html=True)

# Titre (rendu normalement, le sticky sera appliqué via JS)
st.markdown("""
<div class="app-title">
    <span class="ms" style="font-size:36px">school</span>
    <div>
        <h1>CNBAU — Bourse de Russie</h1>
        <span class="sub">Session de la Commission Nationale des Bourses et Aides Universitaires</span>
    </div>
</div>
""", unsafe_allow_html=True)

# Initialisation
db.init_db()

if not db.is_db_loaded():
    st.info("Chargez le fichier Excel des candidatures pour commencer.")
    uploaded = st.file_uploader("Fichier Excel des candidatures", type=["xlsx"])
    if uploaded:
        with open("_temp_upload.xlsx", "wb") as f:
            f.write(uploaded.getvalue())
        n = db.load_excel_to_db("_temp_upload.xlsx")
        quotas_path = Path("quotas.json")
        if quotas_path.exists():
            db.load_quotas(str(quotas_path))
        Path("_temp_upload.xlsx").unlink(missing_ok=True)
        st.success(f"{n} candidatures chargées.")
        st.rerun()
    st.stop()


# Fonctions utilitaires
def render_status(avis: str) -> str:
    if avis == "Favorable":
        return '<span class="badge badge-favorable"><span class="ms">check_circle</span>Favorable</span>'
    elif avis == "Défavorable":
        return '<span class="badge badge-defavorable"><span class="ms">cancel</span>Défavorable</span>'
    return '<span class="badge badge-attente"><span class="ms">schedule</span>En attente</span>'


def icon(name: str) -> str:
    return f'<span class="ms">{name}</span>'


# Données globales
all_df = db.get_all_candidatures()
quotas = db.get_quotas()
favorables_count = db.get_favorables_count()


# KPIs
stats = db.get_stats()
kpis = [
    ("groups", stats["total"], "Total", ""),
    ("task_alt", stats["traites"], "Traitées", ""),
    ("check_circle", stats["favorables"], "Favorables", "kpi-green"),
    ("cancel", stats["defavorables"], "Défavorables", "kpi-red"),
    ("pending", stats["restants"], "Restantes", "kpi-muted"),
]
kpi_html = '<div class="kpi-row" id="kpi-row">'
for ic, val, label, cls in kpis:
    kpi_html += f"""
    <div class="kpi-card {cls}">
        <div class="kpi-icon">{icon(ic)}</div>
        <div class="kpi-value">{val}</div>
        <div class="kpi-label">{label}</div>
    </div>"""
kpi_html += '</div>'
st.markdown(kpi_html, unsafe_allow_html=True)

# Quotas par filière
@st.fragment
def render_quotas():
    fav = db.get_favorables_count()
    st.markdown(f"""
    <div class="section-header">
        <span class="ms">monitoring</span>
        <h3>Quotas par filière</h3>
    </div>""", unsafe_allow_html=True)

    for niveau in NIVEAU_ORDER:
        niveau_quotas = {k: v for k, v in quotas.items() if k[0] == niveau}
        if not niveau_quotas:
            continue
        niveau_quotas = dict(sorted(niveau_quotas.items(), key=lambda item: item[0][1]))
        st.markdown(f'<div class="niveau-label">{niveau}</div>', unsafe_allow_html=True)
        html = '<div class="quota-grid">'
        for (niv, fil), places in niveau_quotas.items():
            selectionnes = fav.get((niv, fil), 0)
            restantes = places - selectionnes
            pct = min((selectionnes / places) * 100, 100) if places > 0 else 0
            css_class = "quota-full" if restantes <= 0 else "quota-ok"
            status_text = "COMPLET" if restantes <= 0 else f"{restantes} restante(s)"
            html += f"""
            <div class="quota-card {css_class}">
                <div class="quota-filiere">{fil}</div>
                <div class="quota-bar"><div class="quota-bar-fill" style="width:{pct}%"></div></div>
                <div class="quota-text">{selectionnes}/{places} — {status_text}</div>
            </div>"""
        html += '</div>'
        st.markdown(html, unsafe_allow_html=True)


render_quotas()

st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)

# Filtres
st.markdown("""
<div class="section-header">
    <span class="ms">filter_list</span>
    <h3>Filtres</h3>
</div>""", unsafe_allow_html=True)

col_f1, col_f2, col_f3 = st.columns(3)
with col_f1:
    niveaux_dispo = sorted(all_df["niveau_etudes"].unique(), key=lambda x: NIVEAU_ORDER.index(x) if x in NIVEAU_ORDER else 99)
    filtre_niveau = st.multiselect("Niveau d'études", niveaux_dispo, placeholder="Filtrer par niveau…")
with col_f2:
    base_df = all_df[all_df["niveau_etudes"].isin(filtre_niveau)] if filtre_niveau else all_df
    filieres_dispo = sorted(base_df["filiere"].unique())
    filtre_filiere = st.multiselect("Filière", filieres_dispo, placeholder="Filtrer par filière…")
with col_f3:
    avis_options = ["Favorable", "Défavorable", "En attente"]
    filtre_avis = st.multiselect("Avis", avis_options, placeholder="Filtrer par avis…")

# Si aucun filtre sélectionné → tout afficher
if filtre_niveau:
    filtered_df = all_df[all_df["niveau_etudes"].isin(filtre_niveau)]
else:
    filtered_df = all_df.copy()
if filtre_filiere:
    filtered_df = filtered_df[filtered_df["filiere"].isin(filtre_filiere)]
if filtre_avis:
    filtered_df = filtered_df[filtered_df["avis"].isin(filtre_avis)]

filtered_df = filtered_df.copy()

# KPIs filtrés
if len(filtered_df) < len(all_df):
    f_total = len(filtered_df)
    f_fav = len(filtered_df[filtered_df["avis"] == "Favorable"])
    f_def = len(filtered_df[filtered_df["avis"] == "Défavorable"])
    st.caption(f"Filtre actif : {f_total} candidatures | {f_fav} favorables | {f_def} défavorables | {f_total - f_fav - f_def} en attente")

# Tableau des candidatures
st.markdown("""
<div class="section-header">
    <span class="ms">list_alt</span>
    <h3>Liste des candidatures</h3>
</div>""", unsafe_allow_html=True)

filtered_df["niveau_order"] = filtered_df["niveau_etudes"].map(
    {v: i for i, v in enumerate(NIVEAU_ORDER)}
)
sort_cols = ["niveau_order", "filiere"]
if "numero" in filtered_df.columns and filtered_df["numero"].notna().any():
    sort_cols.append("numero")
else:
    sort_cols.append("name")
filtered_df = filtered_df.sort_values(sort_cols).drop(columns=["niveau_order"])

# Tableau paginé avec boutons d'avis
PAGE_SIZE = 15
total_rows = len(filtered_df)
total_pages = max(1, -(-total_rows // PAGE_SIZE))

if "page" not in st.session_state:
    st.session_state["page"] = 1
# Borner la page courante au nombre de pages disponibles
if st.session_state["page"] > total_pages:
    st.session_state["page"] = total_pages

page = st.session_state["page"]
page_df = filtered_df.iloc[(page - 1) * PAGE_SIZE : page * PAGE_SIZE]

# En-tête
header_cols = st.columns([0.5, 2.5, 2, 1, 0.8, 1.5, 0.3, 2])
for col, label in zip(header_cols, ["N°", "Nom", "Filière", "Niveau", "Moy.", "Avis", "", "Actions"]):
    col.markdown(f"**{label}**")

st.divider()

# Lignes
fav_counts_local = db.get_favorables_count()
for _, row in page_df.iterrows():
    cols = st.columns([0.5, 2.5, 2, 1, 0.8, 1.5, 0.3, 2])
    _num = row.get("numero", row.get("id_demande", ""))
    cols[0].markdown(
        f'<span class="num-badge">{_num}</span>',
        unsafe_allow_html=True,
    )
    cols[1].write(row["name"])
    cols[2].write(row["filiere"])
    cols[3].write(row["niveau_etudes"])
    cols[4].write(row.get("moyenne", ""))

    avis = row["avis"]
    cols[5].markdown(render_status(avis), unsafe_allow_html=True)

    id_demande = row["id_demande"]
    key_quota = (row["niveau_etudes"], row["filiere"])
    places = quotas.get(key_quota)
    current_fav = fav_counts_local.get(key_quota, 0)
    quota_full = places is not None and current_fav >= places

    with cols[7]:
        b1, b2, b3 = st.columns(3)
        if b1.button("", key=f"fav_{id_demande}", icon=":material/check_circle:", disabled=(avis == "Favorable") or (quota_full and avis != "Favorable"), help="Favorable"):
            db.update_avis(id_demande, "Favorable")
            st.rerun()
        if b2.button("", key=f"def_{id_demande}", icon=":material/cancel:", disabled=(avis == "Défavorable"), help="Défavorable"):
            db.update_avis(id_demande, "Défavorable")
            st.rerun()
        if b3.button("", key=f"att_{id_demande}", icon=":material/schedule:", disabled=(avis == "En attente"), help="En attente"):
            db.update_avis(id_demande, "En attente")
            st.rerun()

# Pagination bas de tableau
col_prev, col_info, col_next = st.columns([1, 2, 1])
with col_prev:
    if st.button("← Précédent", disabled=(page <= 1), use_container_width=True):
        st.session_state["page"] = page - 1
        st.rerun()
with col_info:
    start_row = (page - 1) * PAGE_SIZE + 1
    end_row = min(page * PAGE_SIZE, total_rows)
    st.markdown(
        f"<div style='text-align:center;color:{COLORS['text_muted']};line-height:2.4rem'>{start_row}–{end_row} sur {total_rows} (page {page}/{total_pages})</div>",
        unsafe_allow_html=True,
    )
with col_next:
    if st.button("Suivant →", disabled=(page >= total_pages), use_container_width=True):
        st.session_state["page"] = page + 1
        st.rerun()

st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)

# Gestion des décisions
@st.fragment
def render_decisions():
    st.markdown("""
    <div class="section-header">
        <span class="ms">gavel</span>
        <h3>Gestion des décisions</h3>
    </div>""", unsafe_allow_html=True)

    ID_RUSSE_PREFIX = "BEN-"
    ID_RUSSE_SUFFIX = "/26"

    CRITERES = {"N°": "numero", "ID Russe": "id_russe", "Nom": "name"}
    col_crit, col_search = st.columns([1, 3])
    with col_crit:
        critere = st.selectbox("Critère", list(CRITERES.keys()), label_visibility="collapsed")
    with col_search:
        placeholders = {"N°": "Ex : 1, 12, 250…", "ID Russe": "Ex : 13256", "Nom": "Ex : DUPONT"}
        if critere == "ID Russe":
            col_prefix, col_input, col_suffix = st.columns([1, 2, 1])
            with col_prefix:
                st.markdown(
                    f'<div style="line-height:2.4rem;font-weight:600;color:{COLORS["accent"]}">{ID_RUSSE_PREFIX}</div>',
                    unsafe_allow_html=True,
                )
            with col_input:
                search_raw = st.text_input(
                    "Numéro ID russe",
                    placeholder=placeholders[critere],
                    label_visibility="collapsed",
                ).strip()
            with col_suffix:
                st.markdown(
                    f'<div style="line-height:2.4rem;font-weight:600;color:{COLORS["accent"]}">{ID_RUSSE_SUFFIX}</div>',
                    unsafe_allow_html=True,
                )
            search_query = f"{ID_RUSSE_PREFIX}{search_raw}{ID_RUSSE_SUFFIX}" if search_raw else ""
        else:
            search_query = st.text_input(
                "Rechercher un candidat",
                placeholder=placeholders[critere],
                label_visibility="collapsed",
            ).strip()

    if not search_query:
        st.info(f"Saisissez un **{critere}** pour rechercher un candidat.")
        return

    field = CRITERES[critere]
    candidat = db.search_by_field(field, search_query)
    results = []
    if not candidat:
        results = db.search_by_field_fuzzy(field, search_query)
        if not results:
            st.warning(f"Aucune candidature trouvée pour **{critere}** = « {search_query} ».")
            return
        if len(results) == 1:
            candidat = results[0]
        else:
            st.info(f"{len(results)} résultats trouvés. Sélectionnez un candidat :")
            options = {f"{r['id_demande']} — {r['name']} ({r['filiere']}, {r['niveau_etudes']})": r["id_demande"] for r in results}
            selected = st.selectbox("Candidat", list(options.keys()), label_visibility="collapsed")
            candidat = db.search_by_field("numero", options[selected])

    fav_counts = db.get_favorables_count()
    niveau = candidat["niveau_etudes"]
    filiere = candidat["filiere"]
    places = quotas.get((niveau, filiere), None)
    selectionnes = fav_counts.get((niveau, filiere), 0)
    quota_atteint = places is not None and selectionnes >= places

    col_info, col_actions = st.columns([2, 1])
    with col_info:
        numero = html_escape(str(candidat.get('numero') or ''))
        id_russe = html_escape(str(candidat.get('id_russe') or candidat.get('id_demande', '')))
        sexe = html_escape(str(candidat.get('sexe') or ''))
        date_lieu = html_escape(str(candidat.get('date_lieu_naissance') or ''))
        diplome = html_escape(str(candidat.get('diplome_filiere_annee') or ''))
        observation = html_escape(str(candidat.get('observation') or ''))
        name = html_escape(str(candidat['name']))
        filiere_val = html_escape(str(candidat['filiere']))
        niveau_val = html_escape(str(candidat['niveau_etudes']))
        obs_row = f'<tr><td>Observation</td><td><em>{observation}</em></td></tr>' if observation else ''
        rows = [
            f'<tr><td>N°</td><td>{numero}</td></tr>',
            f'<tr><td>ID Russe</td><td>{id_russe}</td></tr>',
            f'<tr><td>Nom</td><td><strong>{name}</strong></td></tr>',
            f'<tr><td>Sexe</td><td>{sexe}</td></tr>',
            f'<tr><td>Naissance</td><td>{date_lieu}</td></tr>',
            f'<tr><td>Filière</td><td>{filiere_val}</td></tr>',
            f'<tr><td>Niveau</td><td>{niveau_val}</td></tr>',
            f'<tr><td>Diplôme</td><td>{diplome}</td></tr>',
        ]
        if obs_row:
            rows.append(obs_row)
        rows.append(f'<tr><td>Avis</td><td>{render_status(candidat["avis"])}</td></tr>')
        table_html = ''.join(rows)
        st.markdown(
            f'<div class="candidat-card"><table>{table_html}</table></div>',
            unsafe_allow_html=True,
        )

    with col_actions:
        if places is not None:
            pct = min((selectionnes / places) * 100, 100) if places > 0 else 0
            st.markdown(f"""
            <div style="background:{COLORS['bg_card']}; border:1px solid {COLORS['border']}; border-radius:10px; padding:0.8rem 1rem; margin-bottom:0.8rem;">
                <div style="font-size:0.82rem; color:{COLORS['text_muted']}; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:4px;">
                    Quota {filiere} ({niveau})
                </div>
                <div style="font-size:1.4rem; font-weight:700; color:{COLORS['text_primary']};">{selectionnes}/{places}</div>
                <div class="quota-bar" style="height:6px; border-radius:3px; background:{COLORS['bg_dark']}; margin:6px 0; overflow:hidden;">
                    <div style="height:100%; border-radius:3px; width:{pct}%; background:{'#f85149' if quota_atteint else '#3fb950'};"></div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            if quota_atteint and candidat["avis"] != "Favorable":
                st.markdown(f"""
                <div class="alert-quota">
                    <span class="ms">block</span>
                    Quota atteint — avis favorable impossible
                </div>
                """, unsafe_allow_html=True)
                st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div style="background:{COLORS['bg_card']}; border:1px solid {COLORS['border']}; border-radius:10px; padding:0.8rem 1rem; margin-bottom:0.8rem;">
                <div style="font-size:0.82rem; color:{COLORS['text_muted']}; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:4px;">
                    {filiere} ({niveau})
                </div>
                <div style="font-size:0.9rem; color:{COLORS['text_muted']};">Pas de quota défini — {selectionnes} favorable(s)</div>
            </div>
            """, unsafe_allow_html=True)

        btn_col1, btn_col2 = st.columns(2)
        with btn_col1:
            favorable_disabled = quota_atteint and candidat["avis"] != "Favorable"
            if st.button(
                "Favorable",
                key=f"search_fav_{candidat['id_demande']}",
                disabled=favorable_disabled,
                use_container_width=True,
                type="primary",
            ):
                db.update_avis(candidat["id_demande"], "Favorable")
                st.rerun(scope="app")

        with btn_col2:
            if st.button(
                "Défavorable",
                key=f"search_def_{candidat['id_demande']}",
                use_container_width=True,
            ):
                db.update_avis(candidat["id_demande"], "Défavorable")
                st.rerun(scope="app")

        if candidat["avis"] != "En attente":
            if st.button("Remettre en attente", key=f"search_reset_{candidat['id_demande']}", use_container_width=True):
                db.update_avis(candidat["id_demande"], "En attente")
                st.rerun(scope="app")


render_decisions()

st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)

# Export
st.markdown("""
<div class="section-header">
    <span class="ms">download</span>
    <h3>Export des résultats</h3>
</div>""", unsafe_allow_html=True)

col_exp1, col_exp2 = st.columns([1, 3])
with col_exp1:
    if st.button("Générer l'export Excel", use_container_width=True):
        output = "export_decisions_cnbau.xlsx"
        db.export_to_excel(output)
        with open(output, "rb") as f:
            st.session_state["export_data"] = f.read()
        st.session_state["export_ready"] = True

with col_exp2:
    if st.session_state.get("export_ready"):
        st.download_button(
            label="Télécharger le fichier Excel",
            data=st.session_state["export_data"],
            file_name="export_decisions_cnbau.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

# Sticky header via JS
_sticky_bg = COLORS["bg_page"]
_sticky_border = COLORS["border"]
_sticky_shadow = COLORS["shadow"]
_sticky_accent = COLORS["accent"]
_sticky_text = COLORS["text_primary"]

components.html(f"""
<script>
(function() {{
    const doc = window.parent.document;
    const titleEl = doc.querySelector('.app-title');
    const kpiEl = doc.querySelector('#kpi-row');
    if (!titleEl || !kpiEl) return;

    // Supprimer l'ancien clone pour le recréer avec les bonnes couleurs
    const existing = doc.getElementById('sticky-clone');
    if (existing) existing.remove();

    // Créer le header fixe
    const header = doc.createElement('div');
    header.id = 'sticky-clone';
    header.style.cssText = `
        position: fixed; top: 0; left: 0; right: 0; z-index: 9999999;
        background: {_sticky_bg}; padding: 0.8rem 2rem 0.5rem 2rem;
        border-bottom: 1px solid {_sticky_border};
        box-shadow: 0 2px 8px {_sticky_shadow};
        display: none;
    `;
    header.innerHTML = `
        <div style="display:flex;align-items:center;gap:10px;margin-bottom:0.5rem">
            <span class="ms" style="font-size:24px;color:{_sticky_accent};font-family:'Material Symbols Rounded'">school</span>
            <span style="font-size:1.2rem;font-weight:700;color:{_sticky_text}">CNBAU — Bourse de Russie</span>
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

    // Observer le scroll
    const scrollContainer = doc.querySelector('[data-testid="stMainBlockContainer"]')
        || doc.querySelector('section[data-testid="stMain"]')
        || doc.querySelector('.main');
    const sentinel = titleEl.closest('[data-testid="element-container"]') || titleEl;

    function checkScroll() {{
        const rect = sentinel.getBoundingClientRect();
        header.style.display = rect.bottom < 0 ? 'block' : 'none';
    }}

    // Chercher le bon container scrollable
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

    // Observer les changements du KPI row pour mettre à jour le clone
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
""", height=0)

# Sidebar
with st.sidebar:
    st.markdown(f"""
    <div style="display:flex; align-items:center; gap:8px; margin-bottom:1rem;">
        <span class="ms" style="font-size:24px; color:{COLORS['accent']};">settings</span>
        <h3 style="margin:0; font-size:1.1rem; color:{COLORS['text_primary']};">Administration</h3>
    </div>
    """, unsafe_allow_html=True)

    def _on_theme_toggle():
        st.session_state["_theme"] = "light" if st.session_state["_theme_toggle"] else "dark"

    st.toggle("Mode clair", key="_theme_toggle", value=light, on_change=_on_theme_toggle)

    st.caption("Session en cours")
    stats = db.get_stats()
    st.metric("Progression", f"{stats['traites']}/{stats['total']}")
    st.progress(stats["traites"] / stats["total"] if stats["total"] > 0 else 0)

    st.divider()
    if st.button("Réinitialiser la session", type="secondary", use_container_width=True):
        db.reset_db()
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()
