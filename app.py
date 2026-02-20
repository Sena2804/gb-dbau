"""Application Streamlit - Suivi de session CNBAU Bourse de Russie."""

from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

import database as db
from style import NIVEAU_ORDER, build_css, build_sticky_js, get_colors
from ui_helper import (
    render_candidat_card,
    render_kpi_row,
    render_quota_grid,
    render_quota_mini,
    render_status,
    section_header,
)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

st.set_page_config(page_title="CNBAU - Bourse de Russie", layout="wide")

PAGE_SIZE = 15
ID_RUSSE_PREFIX = "BEN-"
ID_RUSSE_SUFFIX = "/26"

# ---------------------------------------------------------------------------
# Thème
# ---------------------------------------------------------------------------

if "_theme" not in st.session_state:
    st.session_state["_theme"] = "light"

light = st.session_state["_theme"] == "light"
COLORS = get_colors(light)

st.markdown(build_css(COLORS, light), unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# En-tête de page
# ---------------------------------------------------------------------------

st.markdown("""
<div class="app-title">
    <span class="ms" style="font-size:36px">school</span>
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

all_df  = db.get_all_candidatures()
quotas  = db.get_quotas()
stats   = db.get_stats()

# ---------------------------------------------------------------------------
# KPIs
# ---------------------------------------------------------------------------

st.markdown(render_kpi_row(stats), unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Quotas par filière
# ---------------------------------------------------------------------------

@st.fragment
def render_quotas_section():
    fav = db.get_favorables_count()
    st.markdown(section_header("monitoring", "Quotas par filière"), unsafe_allow_html=True)

    for niveau in NIVEAU_ORDER:
        niveau_quotas = {k: v for k, v in quotas.items() if k[0] == niveau}
        if not niveau_quotas:
            continue
        niveau_quotas = dict(sorted(niveau_quotas.items(), key=lambda item: item[0][1]))
        st.markdown(f'<div class="niveau-label">{niveau}</div>', unsafe_allow_html=True)
        st.markdown(render_quota_grid(niveau, niveau_quotas, fav), unsafe_allow_html=True)


render_quotas_section()
st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Filtres
# ---------------------------------------------------------------------------

st.markdown(section_header("filter_list", "Filtres"), unsafe_allow_html=True)

col_f1, col_f2, col_f3 = st.columns(3)
with col_f1:
    niveaux_dispo = sorted(
        all_df["niveau_etudes"].unique(),
        key=lambda x: NIVEAU_ORDER.index(x) if x in NIVEAU_ORDER else 99,
    )
    filtre_niveau = st.multiselect("Niveau d'études", niveaux_dispo, placeholder="Filtrer par niveau…")

with col_f2:
    base_df = all_df[all_df["niveau_etudes"].isin(filtre_niveau)] if filtre_niveau else all_df
    filieres_dispo = sorted(base_df["filiere"].unique())
    filtre_filiere = st.multiselect("Filière", filieres_dispo, placeholder="Filtrer par filière…")

with col_f3:
    filtre_avis = st.multiselect(
        "Avis",
        ["Favorable", "Défavorable", "Suppléant", "En attente"],
        placeholder="Filtrer par avis…",
    )

# Application des filtres
filtered_df = all_df.copy()
if filtre_niveau:
    filtered_df = filtered_df[filtered_df["niveau_etudes"].isin(filtre_niveau)]
if filtre_filiere:
    filtered_df = filtered_df[filtered_df["filiere"].isin(filtre_filiere)]
if filtre_avis:
    filtered_df = filtered_df[filtered_df["avis"].isin(filtre_avis)]

if len(filtered_df) < len(all_df):
    f_total = len(filtered_df)
    f_fav   = len(filtered_df[filtered_df["avis"] == "Favorable"])
    f_def   = len(filtered_df[filtered_df["avis"] == "Défavorable"])
    f_sup   = len(filtered_df[filtered_df["avis"] == "Suppléant"])
    st.caption(
        f"Filtre actif : {f_total} candidatures | {f_fav} favorables "
        f"| {f_def} défavorables | {f_sup} suppléants | {f_total - f_fav - f_def - f_sup} en attente"
    )

# ---------------------------------------------------------------------------
# Tableau des candidatures (paginé)
# ---------------------------------------------------------------------------

st.markdown(section_header("list_alt", "Liste des candidatures"), unsafe_allow_html=True)

# Tri
filtered_df["niveau_order"] = filtered_df["niveau_etudes"].map(
    {v: i for i, v in enumerate(NIVEAU_ORDER)}
)
sort_cols = ["niveau_order", "filiere"]
ascending_flags = [True, True]

if "moyenne" in filtered_df.columns:
    sort_cols.append("moyenne")
    ascending_flags.append(False) # Plus haute moyenne en premier

# sort_cols.append("filiere")
# ascending_flags.append(True)
filtered_df = filtered_df.sort_values(by=sort_cols, ascending=ascending_flags).drop(columns=["niveau_order"])

# Pagination
total_rows  = len(filtered_df)
total_pages = max(1, -(-total_rows // PAGE_SIZE))

if "page" not in st.session_state:
    st.session_state["page"] = 1
st.session_state["page"] = min(st.session_state["page"], total_pages)

page    = st.session_state["page"]
page_df = filtered_df.iloc[(page - 1) * PAGE_SIZE : page * PAGE_SIZE]

# En-tête du tableau
for col, label in zip(
    st.columns([0.5, 2.5, 1.5, 1, 1, 1.5, 0.3, 2.5]),
    ["N°", "Nom", "Filière", "Niveau", "Moyenne", "Avis", "", "Actions"],
):
    col.markdown(f"**{label}**")

st.divider()

# Lignes du tableau
fav_counts_local = db.get_favorables_count()

for _, row in page_df.iterrows():
    cols = st.columns([0.5, 2.5, 1.5, 1, 1, 1.5, 0.3, 2.5])
    id_demande = row["id_demande"]
    avis       = row["avis"]
    try:
        moyenne = float(row.get("moyenne") or 0)
    except (ValueError, TypeError):
        moyenne = 0.0
    key_quota  = (row["niveau_etudes"], row["filiere"])
    places     = quotas.get(key_quota)
    quota_full = places is not None and fav_counts_local.get(key_quota, 0) >= places

    _num = row.get("numero", row.get("id_demande", ""))
    cols[0].markdown(f'<span class="num-badge">{_num}</span>', unsafe_allow_html=True)
    cols[1].write(row["name"])
    cols[2].write(row["filiere"])
    cols[3].write(row["niveau_etudes"])
    
    # Affichage de la moyenne (Colonne 4)
    moy_color = COLORS['accent']
    cols[4].markdown(f'<b style="color:{moy_color}">{moyenne:.2f}</b>', unsafe_allow_html=True)
    
    # Statut (Colonne 5)
    cols[5].markdown(render_status(avis), unsafe_allow_html=True)

    # Actions (Colonne 7) - Maintenant l'index 7 existe !
    with cols[7]:
        b1, b2, b3, b4 = st.columns(4)
        if b1.button("", key=f"fav_{id_demande}", icon=":material/check_circle:",
            disabled=(avis == "Favorable") or (quota_full and avis != "Favorable"),
            help="Favorable"):
            db.update_avis(id_demande, "Favorable")
            st.rerun()
        if b2.button("", key=f"def_{id_demande}", icon=":material/cancel:",
            disabled=(avis == "Défavorable"),
            help="Défavorable"):
            db.update_avis(id_demande, "Défavorable")
            st.rerun()
        if b3.button("", key=f"sup_{id_demande}", icon=":material/group_add:",
            disabled=(avis == "Suppléant"),
            help="Suppléant"):
            db.update_avis(id_demande, "Suppléant")
            st.rerun()
        if b4.button("", key=f"att_{id_demande}", icon=":material/schedule:",
            disabled=(avis == "En attente"),
            help="En attente"):
            db.update_avis(id_demande, "En attente")
            st.rerun()
            
            
# Pagination
col_prev, col_info, col_next = st.columns([1, 2, 1])
with col_prev:
    if st.button("← Précédent", disabled=(page <= 1), use_container_width=True):
        st.session_state["page"] = page - 1
        st.rerun()
with col_info:
    start_row = (page - 1) * PAGE_SIZE + 1
    end_row   = min(page * PAGE_SIZE, total_rows)
    st.markdown(
        f"<div style='text-align:center;color:{COLORS['text_muted']};line-height:2.4rem'>"
        f"{start_row}–{end_row} sur {total_rows} (page {page}/{total_pages})</div>",
        unsafe_allow_html=True,
    )
with col_next:
    if st.button("Suivant →", disabled=(page >= total_pages), use_container_width=True):
        st.session_state["page"] = page + 1
        st.rerun()

st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Gestion des décisions (recherche individuelle)
# ---------------------------------------------------------------------------

@st.fragment
def render_decisions():
    st.markdown(section_header("gavel", "Gestion des décisions"), unsafe_allow_html=True)

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
                    f'<div style="line-height:2.4rem;font-weight:600;color:{COLORS["accent"]}">'
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
                    f'<div style="line-height:2.4rem;font-weight:600;color:{COLORS["accent"]}">'
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
        st.info(f"Saisissez un **{critere}** pour rechercher un candidat.")
        return

    # Recherche
    field     = CRITERES[critere]
    candidat  = db.search_by_field(field, search_query)

    if not candidat:
        results = db.search_by_field_fuzzy(field, search_query)
        if not results:
            st.warning(f"Aucune candidature trouvée pour **{critere}** = « {search_query} ».")
            return
        if len(results) == 1:
            candidat = results[0]
        else:
            st.info(f"{len(results)} résultats trouvés. Sélectionnez un candidat :")
            options  = {
                f"{r['id_demande']} — {r['name']} ({r['filiere']}, {r['niveau_etudes']})": r["id_demande"]
                for r in results
            }
            selected = st.selectbox("Candidat", list(options.keys()), label_visibility="collapsed")
            candidat = db.search_by_field("numero", options[selected])

    # Quota
    fav_counts    = db.get_favorables_count()
    niveau        = candidat["niveau_etudes"]
    filiere       = candidat["filiere"]
    places        = quotas.get((niveau, filiere))
    selectionnes  = fav_counts.get((niveau, filiere), 0)
    quota_atteint = places is not None and selectionnes >= places

    col_info_panel, col_actions = st.columns([2, 1])

    with col_info_panel:
        st.markdown(render_candidat_card(candidat), unsafe_allow_html=True)

    with col_actions:
        st.markdown(
            render_quota_mini(filiere, niveau, selectionnes, places, COLORS),
            unsafe_allow_html=True,
        )

        if quota_atteint and candidat["avis"] != "Favorable":
            st.markdown("""
<div class="alert-quota">
    <span class="ms">block</span>
    Quota atteint — avis favorable impossible
</div>""", unsafe_allow_html=True)
            st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)

        btn_col1, btn_col2, btn_col3 = st.columns(3)
        with btn_col1:
            if st.button(
                "Favorable",
                key=f"search_fav_{candidat['id_demande']}",
                disabled=(quota_atteint and candidat["avis"] != "Favorable"),
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

        with btn_col3:
            if st.button(
                "Suppléant",
                key=f"search_sup_{candidat['id_demande']}",
                use_container_width=True,
            ):
                db.update_avis(candidat["id_demande"], "Suppléant")
                st.rerun(scope="app")

        if candidat["avis"] != "En attente":
            if st.button(
                "Remettre en attente",
                key=f"search_reset_{candidat['id_demande']}",
                use_container_width=True,
            ):
                db.update_avis(candidat["id_demande"], "En attente")
                st.rerun(scope="app")


render_decisions()
st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Export Excel
# ---------------------------------------------------------------------------

st.markdown(section_header("download", "Export des résultats"), unsafe_allow_html=True)

col_exp1, col_exp2 = st.columns([1, 3])
with col_exp1:
    if st.button("Générer l'export Word", use_container_width=True):
        output = "export_decisions_cnbau.docx"
        db.export_to_docx(output)
        with open(output, "rb") as f:
            st.session_state["export_data"] = f.read()
        st.session_state["export_ready"] = True

with col_exp2:
    if st.session_state.get("export_ready"):
        st.download_button(
            label="Télécharger le fichier Word",
            data=st.session_state["export_data"],
            file_name="export_decisions_cnbau.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

# ---------------------------------------------------------------------------
# Header sticky (JS)
# ---------------------------------------------------------------------------

components.html(build_sticky_js(COLORS), height=0)

# ---------------------------------------------------------------------------
# Sidebar — Administration
# ---------------------------------------------------------------------------

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
    st.metric("Progression", f"{stats['traites']}/{stats['total']}")
    st.progress(stats["traites"] / stats["total"] if stats["total"] > 0 else 0)

    st.divider()

    if st.button("Réinitialiser la session", type="secondary", use_container_width=True):
        db.reset_db()
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()