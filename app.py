from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

import database as db

st.set_page_config(
    page_title="CNBAU - Bourse de Russie",
    layout="wide",
)

NIVEAU_ORDER = ["Licence", "Master", "Doctorat", "Spécialisation"]

# ── Mode clair / sombre ──────────────────────────────────────────────────────
if "_theme" not in st.session_state:
    st.session_state["_theme"] = "dark"
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

# ── Icônes Material Symbols + CSS global ─────────────────────────────────────
st.markdown(f"""
<link href="https://fonts.googleapis.com/css2?family=Material+Symbols+Rounded:opsz,wght,FILL,GRAD@24,400,1,0" rel="stylesheet">
<style>
    /* --- Theme variables (driven by Python toggle) --- */
    :root {{
        --bg-page: {COLORS["bg_page"]};
        --bg-card: {COLORS["bg_card"]};
        --bg-dark: {COLORS["bg_dark"]};
        --text-primary: {COLORS["text_primary"]};
        --text-muted: {COLORS["text_muted"]};
        --accent: {COLORS["accent"]};
        --border: {COLORS["border"]};
        --border-subtle: {COLORS["border_subtle"]};
        --shadow: {COLORS["shadow"]};
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

    {"" if not light else """
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
    """}
</style>
""", unsafe_allow_html=True)

# ── Titre (rendu normalement, le sticky sera appliqué via JS) ────────────────
st.markdown("""
<div class="app-title">
    <span class="ms" style="font-size:48px">school</span>
    <div>
        <h1>CNBAU — Bourse de Russie</h1>
        <span class="sub">Session de la Commission Nationale des Bourses et Aides Universitaires</span>
    </div>
</div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Initialisation BDD — chargement des données
# ---------------------------------------------------------------------------

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

# ---------------------------------------------------------------------------
# Données globales
# ---------------------------------------------------------------------------

all_df = db.get_all_candidatures()
quotas = db.get_quotas()
stats  = db.get_stats()

# ---------------------------------------------------------------------------
# KPIs — toujours visibles en haut
# ---------------------------------------------------------------------------

st.markdown(render_kpi_row(stats), unsafe_allow_html=True)
st.markdown("<div style='height:1.5rem'></div>", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Navigation par onglets
# ---------------------------------------------------------------------------

tab_titles = [
    ":material/description: Liste des candidatures",
    ":material/leaderboard: Suivi des quotas",
    ":material/gavel: Examen individuel",
    ":material/swap_horiz: Réallocation",
    ":material/download: Export",
]
tab_liste, tab_quotas, tab_eval, tab_realloc, tab_export = st.tabs(tab_titles)

# ===========================================================================
# ONGLET 1 — LISTE DES CANDIDATURES
# ===========================================================================

with tab_liste:
    st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)
    st.markdown(section_header("filter_list", "Parcourir les candidatures"), unsafe_allow_html=True)
    st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)

    # Filtres
    col_f1, col_f2, col_f3 = st.columns(3)

    with col_f1:
        niveaux_dispo = sorted(
            all_df["niveau_etudes"].unique(),
            key=lambda x: NIVEAU_ORDER.index(x) if x in NIVEAU_ORDER else 99,
        )
        filtre_niveau = st.multiselect(
            "Niveau d'études",
            niveaux_dispo,
            placeholder="Tous les niveaux…",
        )

    with col_f2:
        base_df = all_df[all_df["niveau_etudes"].isin(filtre_niveau)] if filtre_niveau else all_df
        filieres_dispo = sorted(base_df["filiere"].unique())
        filtre_filiere = st.multiselect(
            "Filière",
            filieres_dispo,
            placeholder="Toutes les filières…",
        )

    with col_f3:
        filtre_avis = st.multiselect(
            "Avis de la commission",
            ["Favorable", "Défavorable", "Suppléant", "En attente"],
            placeholder="Tous les avis…",
        )

    st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)

    # Application des filtres
    filtered_df = all_df.copy()
    if filtre_niveau:
        filtered_df = filtered_df[filtered_df["niveau_etudes"].isin(filtre_niveau)]
    if filtre_filiere:
        filtered_df = filtered_df[filtered_df["filiere"].isin(filtre_filiere)]
    if filtre_avis:
        filtered_df = filtered_df[filtered_df["avis"].isin(filtre_avis)]

    # Tri
    filtered_df["_niv_order"] = filtered_df["niveau_etudes"].map(
        {v: i for i, v in enumerate(NIVEAU_ORDER)}
    )
    sort_cols       = ["_niv_order", "filiere"]
    ascending_flags = [True, True]

    if "moyenne" in filtered_df.columns:
        sort_cols.append("moyenne")
        ascending_flags.append(False)

    filtered_df = (
        filtered_df
        .sort_values(by=sort_cols, ascending=ascending_flags)
        .drop(columns=["_niv_order"])
    )

    # Pagination
    if "page_liste" not in st.session_state:
        st.session_state["page_liste"] = 1

    total_rows  = len(filtered_df)
    total_pages = max(1, -(-total_rows // PAGE_SIZE))
    st.session_state["page_liste"] = min(st.session_state["page_liste"], total_pages)

    page    = st.session_state["page_liste"]
    page_df = filtered_df.iloc[(page - 1) * PAGE_SIZE : page * PAGE_SIZE]

    # En-tête du tableau — proportions ajustées pour plus de lisibilité
    h_cols = st.columns([0.6, 2.5, 1.5, 2.0, 1.0, 1.8, 2.2])
    labels = ["N°", "Candidat", "Niveau", "Filière", "Moy.", "Statut", "Actions"]

    for col, label in zip(h_cols, labels):
        col.markdown(
            f"<span style='font-size:0.9rem;font-weight:700;color:{COLORS['text_muted']};text-transform:uppercase;letter-spacing:0.8px'>{label}</span>",
            unsafe_allow_html=True
        )

    st.markdown("<div style='height:0.3rem'></div>", unsafe_allow_html=True)
    st.divider()

    fav_counts_local = db.get_favorables_count()

    for _, row in page_df.iterrows():
        id_demande = row["id_demande"]
        avis = row["avis"]
        try:
            moyenne = float(row.get("moyenne") or 0)
        except (ValueError, TypeError):
            moyenne = 0.0

        key_quota = (row["niveau_etudes"], row["filiere"])
        quota_full = quotas.get(key_quota) is not None and fav_counts_local.get(key_quota, 0) >= quotas.get(key_quota)
        _num = row.get("numero", id_demande)

        with st.container():
            # On utilise les mêmes ratios que l'en-tête
            cols = st.columns([0.6, 2.5, 1.5, 2.0, 1.0, 1.8, 2.2])

            # 1. Numéro
            cols[0].markdown(
                f'<div style="display:flex;align-items:center;height:60px">'
                f'<span class="num-badge">{_num}</span></div>',
                unsafe_allow_html=True
            )

            # 2. Nom du Candidat
            cols[1].markdown(f"""
                <div style='line-height:1.3;padding:10px 0'>
                    <div class='candidate-name'>{row['name']}</div>
                </div>
            """, unsafe_allow_html=True)

            # 3. Niveau (Nouvelle colonne)
            cols[2].markdown(
                f"<div style='font-size:0.95rem; font-weight:600; padding:15px 0; color:{COLORS['accent']}'>{row['niveau_etudes']}</div>",
                unsafe_allow_html=True
            )

            # 4. Filière
            cols[3].markdown(
                f"<div style='font-size:0.95rem;padding:10px 0;color:{COLORS['text_primary']}'>{row['filiere']}</div>",
                unsafe_allow_html=True
            )

            # 5. Moyenne
            cols[4].markdown(
                f'<div style="padding:12px 0"><span class="moyenne-txt">{moyenne:.2f}</span></div>',
                unsafe_allow_html=True
            )

            # 6. Statut
            cols[5].markdown(
                f'<div style="padding:10px 0">{render_status(avis)}</div>',
                unsafe_allow_html=True
            )

            # 7. Actions
            with cols[6]:
                b_cols = st.columns(4)
                def btn_action(col, icon, key_suffix, target_avis, current_avis, disabled_cond, help_text):
                    if col.button("", key=f"{key_suffix}_{id_demande}", icon=icon,
                                  disabled=(current_avis == target_avis) or disabled_cond, help=help_text):
                        db.update_avis(id_demande, target_avis)
                        st.rerun()
                btn_action(b_cols[0], ":material/check_circle:", "fav",  "Favorable",   avis, (quota_full and avis != "Favorable"), "Favorable")
                btn_action(b_cols[1], ":material/cancel:",       "def",  "Défavorable", avis, False, "Défavorable")
                btn_action(b_cols[2], ":material/group_add:",    "sup",  "Suppléant",   avis, False, "Suppléant")
                btn_action(b_cols[3], ":material/schedule:",     "att",  "En attente",  avis, False, "En attente")

        st.markdown("<div style='margin-bottom: 4px;'></div>", unsafe_allow_html=True)
        st.divider()


    # Pagination
    st.markdown("<div style='height:0.8rem'></div>", unsafe_allow_html=True)
    col_prev, col_info, col_next = st.columns([1, 2, 1])

    with col_prev:
        if st.button("← Précédent", disabled=(page <= 1), use_container_width=True):
            st.session_state["page_liste"] = page - 1
            st.rerun()

    with col_info:
        start_row = (page - 1) * PAGE_SIZE + 1
        end_row   = min(page * PAGE_SIZE, total_rows)
        st.markdown(
            f"<div style='text-align:center;color:{COLORS['text_muted']};line-height:2.6rem;font-size:1.05rem;font-weight:600'>"
            f"{start_row}–{end_row} sur {total_rows} &nbsp;·&nbsp; page {page} / {total_pages}</div>",
            unsafe_allow_html=True,
        )

    with col_next:
        if st.button("Suivant →", disabled=(page >= total_pages), use_container_width=True):
            st.session_state["page_liste"] = page + 1
            st.rerun()

# ===========================================================================
# ONGLET 2 — SUIVI DES QUOTAS
# ===========================================================================

with tab_quotas:
    fav = db.get_favorables_count()
    st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)
    st.markdown(section_header("monitoring", "État d'avancement des quotas"), unsafe_allow_html=True)
    st.caption("Aperçu en temps réel des places disponibles par filière et par niveau.")
    st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)

    for niveau in NIVEAU_ORDER:
        niveau_quotas = {k: v for k, v in quotas.items() if k[0] == niveau}
        if not niveau_quotas:
            continue
        niveau_quotas = dict(sorted(niveau_quotas.items(), key=lambda item: item[0][1]))
        st.markdown(f'<div class="niveau-label">{niveau}</div>', unsafe_allow_html=True)
        st.markdown(render_quota_grid(niveau, niveau_quotas, fav), unsafe_allow_html=True)
        st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)

# ===========================================================================
# ONGLET 3 — EXAMEN INDIVIDUEL
# ===========================================================================

with tab_eval:
    st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)
    st.markdown(section_header("gavel", "Recherche et Décision"), unsafe_allow_html=True)

    CRITERES = {"N°": "numero", "ID Russe": "id_russe", "Nom": "name"}
    col_crit, col_search = st.columns([1, 3])

    with col_crit:
        critere = st.selectbox(
            "Critère de recherche",
            list(CRITERES.keys()),
            label_visibility="collapsed",
        )

    with col_search:
        placeholders = {
            "N°":       "Ex : 1, 12, 250…",
            "ID Russe": "Ex : 13256",
            "Nom":      "Ex : DUPONT",
        }
        if critere == "ID Russe":
            col_prefix, col_input, col_suffix = st.columns([1, 2, 1])
            with col_prefix:
                st.markdown(
                    f'<div style="line-height:2.6rem;font-weight:700;font-size:1.1rem;color:{COLORS["accent"]}">'
                    f"{ID_RUSSE_PREFIX}</div>",
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
                    f'<div style="line-height:2.6rem;font-weight:700;font-size:1.1rem;color:{COLORS["accent"]}">'
                    f"{ID_RUSSE_SUFFIX}</div>",
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
        st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)
        st.info(f"Saisissez un **{critere}** pour afficher et évaluer un candidat.")
    else:
        field    = CRITERES[critere]
        candidat = db.search_by_field(field, search_query)

        if not candidat:
            results = db.search_by_field_fuzzy(field, search_query)
            if not results:
                st.warning(f"Aucune candidature trouvée pour **{critere}** = « {search_query} ».")
            elif len(results) == 1:
                candidat = results[0]
            else:
                st.info(f"{len(results)} résultats trouvés. Sélectionnez un candidat :")
                options = {
                    f"{r['id_demande']} — {r['name']} ({r['filiere']}, {r['niveau_etudes']})": r["id_demande"]
                    for r in results
                }
                selected = st.selectbox("Candidat", list(options.keys()), label_visibility="collapsed")
                candidat = db.search_by_field("numero", options[selected])

        if candidat:
            st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)

            fav_counts    = db.get_favorables_count()
            niveau        = candidat["niveau_etudes"]
            filiere       = candidat["filiere"]
            places        = quotas.get((niveau, filiere))
            selectionnes  = fav_counts.get((niveau, filiere), 0)
            quota_atteint = places is not None and selectionnes >= places

            col_info_panel, col_actions = st.columns([2, 1], gap="large")

            with col_info_panel:
                st.markdown(render_candidat_card(candidat), unsafe_allow_html=True)

            with col_actions:
                st.markdown(
                    render_quota_mini(filiere, niveau, selectionnes, places, COLORS),
                    unsafe_allow_html=True,
                )

                if quota_atteint and candidat["avis"] != "Favorable":
                    st.error("⚠️ Quota atteint — avis favorable impossible")

                st.markdown("<div style='height:0.8rem'></div>", unsafe_allow_html=True)

        btn_col1, btn_col2 = st.columns(2)
        with btn_col1:
            favorable_disabled = quota_atteint and candidat["avis"] != "Favorable"
            if st.button(
                "Favorable",
                key=f"fav_{candidat['id_demande']}",
                disabled=favorable_disabled,
                use_container_width=True,
                type="primary",
            ):
                db.update_avis(candidat["id_demande"], "Favorable")
                st.rerun(scope="app")

        with btn_col2:
            if st.button(
                "Défavorable",
                key=f"def_{candidat['id_demande']}",
                use_container_width=True,
            ):
                db.update_avis(candidat["id_demande"], "Défavorable")
                st.rerun(scope="app")

        if candidat["avis"] != "En attente":
            if st.button("Remettre en attente", key=f"reset_{candidat['id_demande']}", use_container_width=True):
                db.update_avis(candidat["id_demande"], "En attente")
                st.rerun(scope="app")


render_decisions()

st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)

    # Indicateur global
    total_quota = db.get_total_quota()
    st.markdown(
        f'<div class="transfer-summary">'
        f'<div class="transfer-summary-title">Total des bourses : {total_quota} / 150</div>'
        f'</div>',
        unsafe_allow_html=True,
    )
    st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)

    # Données pour les selectbox
    realloc_quotas = db.get_quotas()
    realloc_fav = db.get_favorables_count()

    # Construire les listes par niveau
    niveaux_with_quotas = sorted(
        {k[0] for k in realloc_quotas},
        key=lambda x: NIVEAU_ORDER.index(x) if x in NIVEAU_ORDER else 99,
    )

    st.markdown('<div class="transfer-card">', unsafe_allow_html=True)

    # --- Sélection niveau / filière HORS du form pour réactivité ---
    col_src, col_dest = st.columns(2, gap="large")

    with col_src:
        st.markdown("**Source — retirer des places**")
        src_niveau = st.selectbox(
            "Niveau (source)",
            niveaux_with_quotas,
            key="src_niveau",
        )
        # Filières de la source ayant des places disponibles > 0
        src_filieres = []
        for (niv, fil), places in sorted(realloc_quotas.items(), key=lambda x: x[0][1]):
            if niv == src_niveau:
                fav_count = realloc_fav.get((niv, fil), 0)
                dispo = places - fav_count
                if dispo > 0:
                    src_filieres.append(fil)

        src_filiere = st.selectbox(
            "Filière (source)",
            src_filieres if src_filieres else ["— aucune filière disponible —"],
            key="src_filiere",
        )

        # Info source
        if src_filieres and src_filiere in src_filieres:
            src_key = (src_niveau, src_filiere)
            src_places = realloc_quotas.get(src_key, 0)
            src_fav = realloc_fav.get(src_key, 0)
            src_dispo = src_places - src_fav
            st.info(f"Quota actuel : **{src_places}** | Favorables : **{src_fav}** | Disponibles : **{src_dispo}**")
        else:
            src_dispo = 0
            st.warning("Aucune filière avec des places disponibles pour ce niveau.")

    with col_dest:
        st.markdown("**Destination — ajouter des places**")
        dest_niveau = st.selectbox(
            "Niveau (destination)",
            niveaux_with_quotas,
            key="dest_niveau",
        )
        # Filières de la destination (exclure la source si même niveau)
        dest_filieres = []
        for (niv, fil), places in sorted(realloc_quotas.items(), key=lambda x: x[0][1]):
            if niv == dest_niveau:
                if dest_niveau == src_niveau and fil == src_filiere:
                    continue
                dest_filieres.append(fil)

        dest_filiere = st.selectbox(
            "Filière (destination)",
            dest_filieres if dest_filieres else ["— aucune filière —"],
            key="dest_filiere",
        )

        # Info destination
        if dest_filieres and dest_filiere in dest_filieres:
            dest_key = (dest_niveau, dest_filiere)
            dest_places = realloc_quotas.get(dest_key, 0)
            dest_fav = realloc_fav.get(dest_key, 0)
            st.info(f"Quota actuel : **{dest_places}** | Favorables : **{dest_fav}**")

    # --- Nombre de places + bouton dans un form pour éviter les reruns accidentels ---
    with st.form("transfer_form"):
        nb_transfer = st.number_input(
            "Nombre de places à transférer",
            min_value=1,
            max_value=max(src_dispo, 1),
            value=1,
            key="nb_transfer",
        )
        submitted = st.form_submit_button("Transférer", type="primary", icon=":material/swap_horiz:")

    st.markdown('</div>', unsafe_allow_html=True)

    if submitted:
        if not src_filieres or src_filiere not in src_filieres:
            st.error("La filière source n'a pas de places disponibles.")
        elif not dest_filieres or dest_filiere not in dest_filieres:
            st.error("Veuillez sélectionner une filière de destination valide.")
        else:
            result = db.transfer_quota(
                src_niveau, src_filiere,
                dest_niveau, dest_filiere,
                nb_transfer,
            )
            if result["success"]:
                # Log dans la session
                if "transfer_log" not in st.session_state:
                    st.session_state["transfer_log"] = []
                st.session_state["transfer_log"].append({
                    "source": f"{src_filiere} ({src_niveau})",
                    "destination": f"{dest_filiere} ({dest_niveau})",
                    "places": nb_transfer,
                    "horodatage": datetime.now().strftime("%H:%M:%S"),
                })
                st.success(
                    f"Transfert effectué : **{nb_transfer}** place(s) de "
                    f"*{src_filiere}* → *{dest_filiere}*. "
                    f"Nouveau quota source : {result['source_nouveau']}, "
                    f"destination : {result['dest_nouveau']}."
                )
                st.rerun()
            else:
                st.error(f"Échec du transfert : {result['error']}")

    # Historique des transferts de la session
    if st.session_state.get("transfer_log"):
        st.markdown("<div style='height:1.5rem'></div>", unsafe_allow_html=True)
        st.markdown(section_header("history", "Historique des transferts (session)"), unsafe_allow_html=True)
        import pandas as _pd
        log_df = _pd.DataFrame(st.session_state["transfer_log"])
        log_df.columns = ["Source", "Destination", "Places", "Heure"]
        st.dataframe(log_df, use_container_width=True, hide_index=True)

# ===========================================================================
# ONGLET 5 — EXPORT
# ===========================================================================

with tab_export:
    st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)
    st.markdown(section_header("download", "Génération des documents officiels"), unsafe_allow_html=True)
    st.caption("Générez et téléchargez les documents de décisions finales pour transmission officielle.")
    st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)

    # --- Export 1 : Word — Titulaires & Suppléants ---
    st.markdown(f"""
    <div class="export-card">
        <div class="export-card-header">
            <span class="ms" style="font-size:28px; color:{COLORS['accent']};">description</span>
            <div>
                <div class="export-card-title">Export Word — Titulaires & Suppléants</div>
                <div class="export-card-desc">Document officiel avec la liste des candidats Favorables (Titulaires) et Suppléants, organisée par niveau et filière.</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    col_g1, col_d1, _ = st.columns([1, 1, 2])
    with col_g1:
        if st.button("Générer", key="gen_word", type="primary", use_container_width=True, icon=":material/description:"):
            with st.spinner("Génération en cours…"):
                output = "export_decisions_cnbau.docx"
                db.export_to_docx(output)
                with open(output, "rb") as f:
                    st.session_state["export_word_data"] = f.read()
                st.session_state["export_word_ready"] = True
    with col_d1:
        if st.session_state.get("export_word_ready"):
            st.download_button(
                label="Télécharger (.docx)", key="dl_word",
                data=st.session_state["export_word_data"],
                file_name="export_decisions_cnbau.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True, icon=":material/download:",
            )

    st.divider()

    # --- Export 2 : Word — Toutes les décisions ---
    st.markdown(f"""
    <div class="export-card">
        <div class="export-card-header">
            <span class="ms" style="font-size:28px; color:{COLORS['accent']};">fact_check</span>
            <div>
                <div class="export-card-title">Export Word — Toutes les décisions</div>
                <div class="export-card-desc">Document complet incluant Titulaires, Suppléants et Candidats non retenus (Défavorables), organisé par niveau et filière.</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    col_g2, col_d2, _ = st.columns([1, 1, 2])
    with col_g2:
        if st.button("Générer", key="gen_word_all", type="primary", use_container_width=True, icon=":material/fact_check:"):
            with st.spinner("Génération en cours…"):
                output = "export_toutes_decisions_cnbau.docx"
                db.export_all_avis_to_docx(output)
                with open(output, "rb") as f:
                    st.session_state["export_word_all_data"] = f.read()
                st.session_state["export_word_all_ready"] = True
    with col_d2:
        if st.session_state.get("export_word_all_ready"):
            st.download_button(
                label="Télécharger (.docx)", key="dl_word_all",
                data=st.session_state["export_word_all_data"],
                file_name="export_toutes_decisions_cnbau.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True, icon=":material/download:",
            )

    st.divider()

    # --- Export 3 : Excel — Candidatures par avis ---
    st.markdown(f"""
    <div class="export-card">
        <div class="export-card-header">
            <span class="ms" style="font-size:28px; color:#3fb950;">table_chart</span>
            <div>
                <div class="export-card-title">Export Excel — Candidatures par avis</div>
                <div class="export-card-desc">Fichier Excel avec une feuille par catégorie d'avis (Favorable, Suppléant, Défavorable). Idéal pour l'analyse et le traitement des données.</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    col_g3, col_d3, _ = st.columns([1, 1, 2])
    with col_g3:
        if st.button("Générer", key="gen_excel_avis", type="primary", use_container_width=True, icon=":material/table_chart:"):
            with st.spinner("Génération en cours…"):
                output = "export_decisions_cnbau.xlsx"
                db.export_avis_to_xlsx(output)
                with open(output, "rb") as f:
                    st.session_state["export_excel_avis_data"] = f.read()
                st.session_state["export_excel_avis_ready"] = True
    with col_d3:
        if st.session_state.get("export_excel_avis_ready"):
            st.download_button(
                label="Télécharger (.xlsx)", key="dl_excel_avis",
                data=st.session_state["export_excel_avis_data"],
                file_name="export_decisions_cnbau.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, icon=":material/download:",
            )

    st.divider()

    # --- Export 4 : Excel — Grille des quotas ---
    st.markdown(f"""
    <div class="export-card">
        <div class="export-card-header">
            <span class="ms" style="font-size:28px; color:#d29922;">grid_view</span>
            <div>
                <div class="export-card-title">Export Excel — Grille des quotas</div>
                <div class="export-card-desc">Document de référence listant toutes les filières par niveau avec leurs quotas, le nombre de favorables et les places restantes.</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    col_g4, col_d4, _ = st.columns([1, 1, 2])
    with col_g4:
        if st.button("Générer", key="gen_excel_quotas", type="primary", use_container_width=True, icon=":material/grid_view:"):
            with st.spinner("Génération en cours…"):
                output = "export_quotas_cnbau.xlsx"
                db.export_quotas_to_xlsx(output)
                with open(output, "rb") as f:
                    st.session_state["export_excel_quotas_data"] = f.read()
                st.session_state["export_excel_quotas_ready"] = True
    with col_d4:
        if st.session_state.get("export_excel_quotas_ready"):
            st.download_button(
                label="Télécharger (.xlsx)", key="dl_excel_quotas",
                data=st.session_state["export_excel_quotas_data"],
                file_name="export_quotas_cnbau.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, icon=":material/download:",
            )

# ---------------------------------------------------------------------------
# Header sticky (JS)
# ---------------------------------------------------------------------------

components.html(build_sticky_js(COLORS), height=0)

# ---------------------------------------------------------------------------
# Sidebar — Administration
# ---------------------------------------------------------------------------

with st.sidebar:
    total = sum(quotas.values())

    progression = stats["favorables"] / total if total > 0 else 0
    pct = int(progression * 100)

    color_bar = "#008751" if progression >= 1.0 else "#EAC100"

    st.markdown(get_sidebar_style(color_bar), unsafe_allow_html=True)

    st.markdown(f"""
        <div class="sidebar-status-container">
            <h1 class="sidebar-title">ÉTAT DE LA MISSION</h1>
            <div class="progression-row">
                <span class="progression-label">Progression</span>
                <span class="progression-pct">{pct}%</span>
            </div>
            <div class="progression-sub">
                {stats['favorables']} accordées sur {total} places
            </div>
        </div>
    """, unsafe_allow_html=True)

    st.progress(progression)
    st.markdown('<div style="height:1rem"></div>', unsafe_allow_html=True)
    st.divider()

    if st.button("Réinitialiser la session", type="secondary", use_container_width=True):
        db.reset_db()
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()
