"""Fonctions utilitaires de rendu HTML pour l'interface CNBAU."""

from html import escape as html_escape


# ---------------------------------------------------------------------------
# Primitives
# ---------------------------------------------------------------------------

def icon(name: str) -> str:
    """Retourne un <span> Material Symbol."""
    return f'<span class="ms">{name}</span>'


def render_status(avis: str) -> str:
    """Retourne le badge HTML coloré selon l'avis."""
    if avis == "Favorable":
        return '<span class="badge badge-favorable"><span class="ms">check_circle</span>Favorable</span>'
    if avis == "Défavorable":
        return '<span class="badge badge-defavorable"><span class="ms">cancel</span>Défavorable</span>'
    return '<span class="badge badge-attente"><span class="ms">schedule</span>En attente</span>'


def section_header(icon_name: str, title: str) -> str:
    """Retourne le HTML d'un en-tête de section."""
    return f"""
<div class="section-header">
    <span class="ms">{icon_name}</span>
    <h3>{title}</h3>
</div>"""


# ---------------------------------------------------------------------------
# Blocs composites
# ---------------------------------------------------------------------------

def render_kpi_row(stats: dict) -> str:
    """Retourne le HTML de la rangée de KPI."""
    kpis = [
        ("groups",        stats["total"],       "Total",       ""),
        ("task_alt",      stats["traites"],     "Traitées",    ""),
        ("check_circle",  stats["favorables"],  "Favorables",  "kpi-green"),
        ("cancel",        stats["defavorables"],"Défavorables","kpi-red"),
        ("pending",       stats["restants"],    "Restantes",   "kpi-muted"),
    ]
    cards = ""
    for ic, val, label, cls in kpis:
        cards += f"""
    <div class="kpi-card {cls}">
        <div class="kpi-icon">{icon(ic)}</div>
        <div class="kpi-value">{val}</div>
        <div class="kpi-label">{label}</div>
    </div>"""
    return f'<div class="kpi-row" id="kpi-row">{cards}</div>'


def render_quota_grid(niveau: str, niveau_quotas: dict, fav_counts: dict) -> str:
    """Retourne le HTML de la grille de quotas pour un niveau donné."""
    html = '<div class="quota-grid">'
    for (niv, fil), places in niveau_quotas.items():
        selectionnes = fav_counts.get((niv, fil), 0)
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
    html += "</div>"
    return html


def render_candidat_card(candidat: dict) -> str:
    """Retourne la fiche HTML d'un candidat."""
    numero        = html_escape(str(candidat.get("numero") or ""))
    id_russe      = html_escape(str(candidat.get("id_russe") or candidat.get("id_demande", "")))
    sexe          = html_escape(str(candidat.get("sexe") or ""))
    date_lieu     = html_escape(str(candidat.get("date_lieu_naissance") or ""))
    diplome       = html_escape(str(candidat.get("diplome_filiere_annee") or ""))
    observation   = html_escape(str(candidat.get("observation") or ""))
    name          = html_escape(str(candidat["name"]))
    filiere_val   = html_escape(str(candidat["filiere"]))
    niveau_val    = html_escape(str(candidat["niveau_etudes"]))

    obs_row = f"<tr><td>Observation</td><td><em>{observation}</em></td></tr>" if observation else ""

    rows = "".join([
        f"<tr><td>N°</td><td>{numero}</td></tr>",
        f"<tr><td>ID Russe</td><td>{id_russe}</td></tr>",
        f"<tr><td>Nom</td><td><strong>{name}</strong></td></tr>",
        f"<tr><td>Sexe</td><td>{sexe}</td></tr>",
        f"<tr><td>Naissance</td><td>{date_lieu}</td></tr>",
        f"<tr><td>Filière</td><td>{filiere_val}</td></tr>",
        f"<tr><td>Niveau</td><td>{niveau_val}</td></tr>",
        f"<tr><td>Diplôme</td><td>{diplome}</td></tr>",
        obs_row,
        f"<tr><td>Avis</td><td>{render_status(candidat['avis'])}</td></tr>",
    ])
    return f'<div class="candidat-card"><table>{rows}</table></div>'


def render_quota_mini(
    filiere: str,
    niveau: str,
    selectionnes: int,
    places: int | None,
    colors: dict,
) -> str:
    """Retourne le mini-bloc quota dans le panneau de décision."""
    if places is not None:
        quota_atteint = selectionnes >= places
        pct = min((selectionnes / places) * 100, 100) if places > 0 else 0
        bar_color = "#f85149" if quota_atteint else "#3fb950"
        return f"""
<div style="background:{colors['bg_card']}; border:1px solid {colors['border']}; border-radius:10px; padding:0.8rem 1rem; margin-bottom:0.8rem;">
    <div style="font-size:0.82rem; color:{colors['text_muted']}; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:4px;">
        Quota {filiere} ({niveau})
    </div>
    <div style="font-size:1.4rem; font-weight:700; color:{colors['text_primary']};">{selectionnes}/{places}</div>
    <div class="quota-bar" style="height:6px; border-radius:3px; background:{colors['bg_dark']}; margin:6px 0; overflow:hidden;">
        <div style="height:100%; border-radius:3px; width:{pct}%; background:{bar_color};"></div>
    </div>
</div>"""
    return f"""
<div style="background:{colors['bg_card']}; border:1px solid {colors['border']}; border-radius:10px; padding:0.8rem 1rem; margin-bottom:0.8rem;">
    <div style="font-size:0.82rem; color:{colors['text_muted']}; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:4px;">
        {filiere} ({niveau})
    </div>
    <div style="font-size:0.9rem; color:{colors['text_muted']};">Pas de quota défini — {selectionnes} favorable(s)</div>
</div>"""