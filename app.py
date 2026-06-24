import re
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

import database as db
from style import NIVEAU_ORDER, build_css, build_sticky_js, get_colors, get_sidebar_style
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

st.set_page_config(page_title="CNaBAU - Bourse du Maroc", layout="wide")

PAGE_SIZE = 15

# ---------------------------------------------------------------------------
# Thème
# ---------------------------------------------------------------------------

if "_theme" not in st.session_state:
    st.session_state["_theme"] = "light"

light  = st.session_state["_theme"] == "light"
COLORS = get_colors(light)

st.markdown(build_css(COLORS, light), unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# En-tête de page
# ---------------------------------------------------------------------------

st.markdown("""
<div class="app-title">
    <span class="ms" style="font-size:48px">school</span>
    <div>
        <h1>CNaBAU — Bourse du Maroc</h1>
        <span class="sub">Session de la Commission Nationale des Bourses et Aides Universitaires</span>
    </div>
</div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Initialisation BDD
# ---------------------------------------------------------------------------

db.init_db()

if not db.is_db_loaded():
    st.info("Chargez le fichier Excel des candidatures pour commencer.")
    uploaded = st.file_uploader("Fichier Excel des candidatures", type=["xlsx"])
    if uploaded:
        temp_path = Path("_temp_upload.xlsx")
        with temp_path.open("wb") as f:
            f.write(uploaded.getvalue())
        try:
            n = db.load_excel_to_db(str(temp_path))
        except (ValueError, KeyError) as exc:
            st.error(f"Le fichier Excel n'est pas compatible avec la session Maroc 2026 : {exc}")
            st.stop()
        finally:
            temp_path.unlink(missing_ok=True)
        st.success(f"{n} candidatures et leurs quotas ont été chargés.")
        st.rerun()
    st.stop()

# ---------------------------------------------------------------------------
# ✅ OPTIMISATION : fonctions de chargement cachées (TTL 2s)
# Les données sont rechargées au maximum toutes les 2 secondes, pas à chaque
# rerun. Après un update_avis(), on vide le cache manuellement.
# ---------------------------------------------------------------------------

@st.cache_data(ttl=2)
def cached_get_all_candidatures():
    return db.get_all_candidatures()

@st.cache_data(ttl=2)
def cached_get_quotas():
    return db.get_quotas()

@st.cache_data(ttl=2)
def cached_get_stats():
    return db.get_stats()

@st.cache_data(ttl=2)
def cached_get_favorables_count():
    return db.get_favorables_count()


def invalidate_cache():
    """Vide le cache après toute modification BDD pour forcer un rechargement."""
    cached_get_all_candidatures.clear()
    cached_get_quotas.clear()
    cached_get_stats.clear()
    cached_get_favorables_count.clear()


# ---------------------------------------------------------------------------
# ✅ OPTIMISATION : verrou anti-double-clic
# Empêche deux actions simultanées quand l'utilisateur clique rapidement.
# ---------------------------------------------------------------------------

if "processing" not in st.session_state:
    st.session_state["processing"] = False


def do_update_avis(id_demande: str, avis: str):
    """Met à jour l'avis avec verrou et invalidation du cache."""
    if st.session_state["processing"]:
        return
    st.session_state["processing"] = True
    try:
        db.update_avis(id_demande, avis)
        invalidate_cache()
    finally:
        st.session_state["processing"] = False
    st.rerun()


# ---------------------------------------------------------------------------
# Données globales (cachées)
# ---------------------------------------------------------------------------

all_df = cached_get_all_candidatures()
quotas = cached_get_quotas()
stats  = cached_get_stats()

# ---------------------------------------------------------------------------
# KPIs
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

    col_f1, col_f2, col_f3 = st.columns(3)

    with col_f1:
        niveaux_dispo = sorted(
            all_df["niveau_etudes"].unique(),
            key=lambda x: NIVEAU_ORDER.index(x) if x in NIVEAU_ORDER else 99,
        )
        filtre_niveau = st.multiselect("Niveau d'études", niveaux_dispo, placeholder="Tous les niveaux…")

    with col_f2:
        base_df = all_df[all_df["niveau_etudes"].isin(filtre_niveau)] if filtre_niveau else all_df
        filieres_dispo = sorted(base_df["filiere"].unique())
        filtre_filiere = st.multiselect("Filière", filieres_dispo, placeholder="Toutes les filières…")

    with col_f3:
        filtre_avis = st.multiselect(
            "Avis de la commission",
            ["Favorable", "Défavorable", "Suppléant", "En attente"],
            placeholder="Tous les avis…",
        )

    st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)

    filtered_df = all_df.copy()
    if filtre_niveau:
        filtered_df = filtered_df[filtered_df["niveau_etudes"].isin(filtre_niveau)]
    if filtre_filiere:
        filtered_df = filtered_df[filtered_df["filiere"].isin(filtre_filiere)]
    if filtre_avis:
        filtered_df = filtered_df[filtered_df["avis"].isin(filtre_avis)]

    filtered_df["_niv_order"] = filtered_df["niveau_etudes"].map(
        {v: i for i, v in enumerate(NIVEAU_ORDER)}
    )
    filtered_df["_moyenne_num"] = (
        filtered_df["moyenne"]
        .astype(str)
        .str.extract(r"(\d+(?:[.,]\d+)?)", expand=False)
        .str.replace(",", ".", regex=False)
        .astype(float)
    )
    sort_cols       = ["_niv_order", "filiere"]
    ascending_flags = [True, True]
    if "moyenne" in filtered_df.columns:
        sort_cols.append("_moyenne_num")
        ascending_flags.append(False)

    filtered_df = (
        filtered_df
        .sort_values(by=sort_cols, ascending=ascending_flags)
        .drop(columns=["_niv_order", "_moyenne_num"])
    )

    if "page_liste" not in st.session_state:
        st.session_state["page_liste"] = 1

    total_rows  = len(filtered_df)
    total_pages = max(1, -(-total_rows // PAGE_SIZE))
    st.session_state["page_liste"] = min(st.session_state["page_liste"], total_pages)

    page    = st.session_state["page_liste"]
    page_df = filtered_df.iloc[(page - 1) * PAGE_SIZE : page * PAGE_SIZE]

    h_cols = st.columns([0.6, 2.5, 1, 2.0, 3.0, 1.8, 2.2])
    for col, label in zip(h_cols, ["N°", "Candidat", "Niveau", "Filière", "Moy.", "Statut", "Actions"]):
        col.markdown(
            f"<span style='font-size:0.9rem;font-weight:700;color:{COLORS['text_muted']};text-transform:uppercase;letter-spacing:0.8px'>{label}</span>",
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height:0.3rem'></div>", unsafe_allow_html=True)
    st.divider()

    # ✅ Chargé une seule fois pour toute la page (pas dans la boucle)
    fav_counts_local = cached_get_favorables_count()

    for _, row in page_df.iterrows():
        id_demande = row["id_demande"]
        avis       = row["avis"]
        try:
            match = re.search(r"\d+(?:[.,]\d+)?", str(row.get("moyenne") or ""))
            moyenne = float(match.group(0).replace(",", ".")) if match else 0.0
        except (ValueError, TypeError):
            moyenne = 0.0

        key_quota  = (row["niveau_etudes"], row["filiere"])
        quota_full = (
            quotas.get(key_quota) is not None
            and fav_counts_local.get(key_quota, 0) >= quotas.get(key_quota)
        )
        _num = row.get("numero", id_demande)

        with st.container():
            cols = st.columns([0.6, 2.5, 1.5, 2.0, 2.0, 1.8, 2.2])

            cols[0].markdown(
                f'<div style="display:flex;align-items:center;height:60px">'
                f'<span class="num-badge">{_num}</span></div>',
                unsafe_allow_html=True,
            )
            cols[1].markdown(
                f"<div style='line-height:1.3;padding:10px 0'>"
                f"<div class='candidate-name'>{row['name']}</div></div>",
                unsafe_allow_html=True,
            )
            cols[2].markdown(
                f"<div style='font-size:0.95rem;font-weight:600;padding:15px 0;color:{COLORS['accent']}'>"
                f"{row['niveau_etudes']}</div>",
                unsafe_allow_html=True,
            )
            cols[3].markdown(
                f"<div style='font-size:0.95rem;padding:10px 0;color:{COLORS['text_primary']}'>"
                f"{row['filiere']}</div>",
                unsafe_allow_html=True,
            )
            cols[4].markdown(
                f'<div style="padding:12px 0"><span class="moyenne-txt">{moyenne:.2f}</span></div>',
                unsafe_allow_html=True,
            )
            cols[5].markdown(
                f'<div style="padding:10px 0">{render_status(avis)}</div>',
                unsafe_allow_html=True,
            )

            with cols[6]:
                b_cols = st.columns(4)

                # ✅ Verrou anti-double-clic : disabled si processing=True
                def btn_action(col, icon, key_suffix, target_avis, current_avis, disabled_cond, help_text):
                    is_disabled = (
                        (current_avis == target_avis)
                        or disabled_cond
                        or st.session_state["processing"]
                    )
                    if col.button(
                        "", key=f"{key_suffix}_{id_demande}", icon=icon,
                        disabled=is_disabled, help=help_text,
                    ):
                        do_update_avis(id_demande, target_avis)

                btn_action(b_cols[0], ":material/check_circle:", "fav",  "Favorable",   avis, (quota_full and avis != "Favorable"), "Favorable")
                btn_action(b_cols[1], ":material/cancel:",       "def",  "Défavorable", avis, False, "Défavorable")
                btn_action(b_cols[2], ":material/group_add:",    "sup",  "Suppléant",   avis, False, "Suppléant")
                btn_action(b_cols[3], ":material/schedule:",     "att",  "En attente",  avis, False, "En attente")

        st.markdown("<div style='margin-bottom:4px;'></div>", unsafe_allow_html=True)
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
            f"<div style='text-align:center;color:{COLORS['text_muted']};line-height:2.6rem;"
            f"font-size:1.05rem;font-weight:600'>"
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
    fav = cached_get_favorables_count()
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

    CRITERES = {"N°": "numero", "Nom": "name"}
    col_crit, col_search = st.columns([1, 3])

    with col_crit:
        critere = st.selectbox("Critère de recherche", list(CRITERES.keys()), label_visibility="collapsed")

    with col_search:
        placeholders = {"N°": "Ex : 1, 12, 250…", "Nom": "Ex : DUPONT"}
        search_query = st.text_input(
            "Rechercher un candidat", placeholder=placeholders[critere], label_visibility="collapsed"
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
                options  = {
                    f"N° {r['numero']} — {r['name']} ({r['filiere']}, {r['niveau_etudes']})": r["numero"]
                    for r in results
                }
                selected = st.selectbox("Candidat", list(options.keys()), label_visibility="collapsed")
                candidat = db.search_by_field("numero", options[selected])

        if candidat:
            st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)

            fav_counts    = cached_get_favorables_count()
            niveau        = candidat["niveau_etudes"]
            filiere       = candidat["filiere"]
            places        = quotas.get((niveau, filiere))
            selectionnes  = fav_counts.get((niveau, filiere), 0)
            quota_atteint = places is not None and selectionnes >= places

            col_info_panel, col_actions = st.columns([2, 1], gap="large")

            with col_info_panel:
                st.markdown(render_candidat_card(candidat), unsafe_allow_html=True)

            with col_actions:
                st.markdown(render_quota_mini(filiere, niveau, selectionnes, places, COLORS), unsafe_allow_html=True)

                if quota_atteint and candidat["avis"] != "Favorable":
                    st.error("⚠️ Quota atteint — avis favorable impossible")

                st.markdown("<div style='height:0.8rem'></div>", unsafe_allow_html=True)

                btn_col1, btn_col2 = st.columns(2)

                with btn_col1:
                    if st.button(
                        "✅ Favorable",
                        key=f"eval_fav_{candidat['id_demande']}",
                        disabled=(quota_atteint and candidat["avis"] != "Favorable") or st.session_state["processing"],
                        use_container_width=True,
                        type="primary",
                    ):
                        do_update_avis(candidat["id_demande"], "Favorable")

                    st.markdown("<div style='height:0.4rem'></div>", unsafe_allow_html=True)

                    if st.button(
                        "👥 Suppléant",
                        key=f"eval_sup_{candidat['id_demande']}",
                        use_container_width=True,
                        disabled=st.session_state["processing"],
                    ):
                        do_update_avis(candidat["id_demande"], "Suppléant")

                with btn_col2:
                    if st.button(
                        "❌ Défavorable",
                        key=f"eval_def_{candidat['id_demande']}",
                        use_container_width=True,
                        disabled=st.session_state["processing"],
                    ):
                        do_update_avis(candidat["id_demande"], "Défavorable")

                    st.markdown("<div style='height:0.4rem'></div>", unsafe_allow_html=True)

                    if candidat["avis"] != "En attente":
                        if st.button(
                            "🔄 En attente",
                            key=f"eval_att_{candidat['id_demande']}",
                            use_container_width=True,
                            disabled=st.session_state["processing"],
                        ):
                            do_update_avis(candidat["id_demande"], "En attente")

# ===========================================================================
# ONGLET 4 — RÉALLOCATION DES QUOTAS
# ===========================================================================

with tab_realloc:
    from datetime import datetime

    st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)
    st.markdown(section_header("swap_horiz", "Réallocation des quotas"), unsafe_allow_html=True)
    st.caption(
        "Transférez des places inutilisées d'une filière vers une autre. "
        "Le total des bourses reste toujours inchangé."
    )
    st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)

    total_quota = sum(cached_get_quotas().values())
    st.markdown(
        f'<div class="transfer-summary">'
        f'<div class="transfer-summary-title">Total des bourses : {total_quota}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )
    st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)

    realloc_quotas = cached_get_quotas()
    realloc_fav    = cached_get_favorables_count()

    niveaux_with_quotas = sorted(
        {k[0] for k in realloc_quotas},
        key=lambda x: NIVEAU_ORDER.index(x) if x in NIVEAU_ORDER else 99,
    )

    st.markdown('<div class="transfer-card">', unsafe_allow_html=True)

    col_src, col_dest = st.columns(2, gap="large")

    with col_src:
        st.markdown("**Source — retirer des places**")
        src_niveau   = st.selectbox("Niveau (source)", niveaux_with_quotas, key="src_niveau")
        src_filieres = []
        for (niv, fil), places in sorted(realloc_quotas.items(), key=lambda x: x[0][1]):
            if niv == src_niveau:
                dispo = places - realloc_fav.get((niv, fil), 0)
                if dispo > 0:
                    src_filieres.append(fil)

        src_filiere = st.selectbox(
            "Filière (source)",
            src_filieres if src_filieres else ["— aucune filière disponible —"],
            key="src_filiere",
        )

        if src_filieres and src_filiere in src_filieres:
            src_key   = (src_niveau, src_filiere)
            src_places = realloc_quotas.get(src_key, 0)
            src_fav    = realloc_fav.get(src_key, 0)
            src_dispo  = src_places - src_fav
            st.info(f"Quota actuel : **{src_places}** | Favorables : **{src_fav}** | Disponibles : **{src_dispo}**")
        else:
            src_dispo = 0
            st.warning("Aucune filière avec des places disponibles pour ce niveau.")

    with col_dest:
        st.markdown("**Destination — ajouter des places**")
        dest_niveau   = st.selectbox("Niveau (destination)", niveaux_with_quotas, key="dest_niveau")
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

        if dest_filieres and dest_filiere in dest_filieres:
            dest_key    = (dest_niveau, dest_filiere)
            dest_places = realloc_quotas.get(dest_key, 0)
            dest_fav    = realloc_fav.get(dest_key, 0)
            st.info(f"Quota actuel : **{dest_places}** | Favorables : **{dest_fav}**")

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
            result = db.transfer_quota(src_niveau, src_filiere, dest_niveau, dest_filiere, nb_transfer)
            if result["success"]:
                invalidate_cache()
                if "transfer_log" not in st.session_state:
                    st.session_state["transfer_log"] = []
                st.session_state["transfer_log"].append({
                    "source":      f"{src_filiere} ({src_niveau})",
                    "destination": f"{dest_filiere} ({dest_niveau})",
                    "places":      nb_transfer,
                    "horodatage":  datetime.now().strftime("%H:%M:%S"),
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

    # Export 1 : Word — Titulaires & Suppléants
    st.markdown(f"""
    <div class="export-card">
        <div class="export-card-header">
            <span class="ms" style="font-size:28px;color:{COLORS['accent']};">description</span>
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

    # Export 2 : Word — Toutes les décisions
    st.markdown(f"""
    <div class="export-card">
        <div class="export-card-header">
            <span class="ms" style="font-size:28px;color:{COLORS['accent']};">fact_check</span>
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

    # Export 3 : Excel — Candidatures par avis
    st.markdown(f"""
    <div class="export-card">
        <div class="export-card-header">
            <span class="ms" style="font-size:28px;color:#3fb950;">table_chart</span>
            <div>
                <div class="export-card-title">Export Excel — Candidatures par avis</div>
                <div class="export-card-desc">Fichier Excel avec une feuille par catégorie d'avis. Idéal pour l'analyse et le traitement des données.</div>
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

    # Export 4 : Excel — Grille des quotas
    st.markdown(f"""
    <div class="export-card">
        <div class="export-card-header">
            <span class="ms" style="font-size:28px;color:#d29922;">grid_view</span>
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
    total      = sum(quotas.values())
    progression = stats["favorables"] / total if total > 0 else 0
    pct         = int(progression * 100)
    color_bar   = "#008751" if progression >= 1.0 else "#EAC100"

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
        invalidate_cache()
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()
