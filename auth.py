import hmac
from typing import Any

import streamlit as st


def _get_secret_section(name: str) -> dict[str, Any]:
    try:
        section = st.secrets.get(name, {})
    except (FileNotFoundError, KeyError):
        return {}
    return dict(section) if section else {}


def _normalized_list(values: Any) -> set[str]:
    if not values:
        return set()
    return {str(value).strip().lower() for value in values if str(value).strip()}


def _user_email() -> str:
    try:
        return str(st.user.get("email", "")).strip().lower()
    except Exception:
        return ""


def _password_matches(candidate: str, expected: str) -> bool:
    if not candidate or not expected:
        return False
    return hmac.compare_digest(candidate, expected)


def _clear_auth_state() -> None:
    st.session_state.pop("password_authenticated", None)
    st.session_state.pop("auth_method", None)
    st.session_state.pop("auth_email", None)
    st.session_state.pop("is_admin", None)


def _logout_button(label: str = "Se deconnecter") -> None:
    if st.button(label, use_container_width=True):
        _clear_auth_state()
        if getattr(st.user, "is_logged_in", False):
            st.logout()
        st.rerun()


def require_auth() -> dict[str, Any]:
    access = _get_secret_section("access")
    allowed_emails = _normalized_list(access.get("allowed_emails"))
    admin_emails = _normalized_list(access.get("admin_emails"))
    emergency_password = str(access.get("emergency_password", ""))
    oidc_provider = str(access.get("oidc_provider", "")).strip() or None

    if getattr(st.user, "is_logged_in", False):
        email = _user_email()
        if not allowed_emails:
            st.error("La liste des comptes autorises n'est pas configuree.")
            _logout_button()
            st.stop()
        if allowed_emails and email not in allowed_emails:
            st.error("Ce compte n'est pas autorise pour cette application.")
            _logout_button()
            st.stop()

        is_admin = email in admin_emails
        st.session_state["auth_method"] = "oidc"
        st.session_state["auth_email"] = email
        st.session_state["is_admin"] = is_admin

        with st.sidebar:
            st.caption(f"Connecte : {email or 'compte DBAU'}")
            role = "Administrateur" if is_admin else "Membre"
            st.caption(f"Role : {role}")
            _logout_button()

        return {"method": "oidc", "email": email, "is_admin": is_admin}

    if st.session_state.get("password_authenticated"):
        st.session_state["auth_method"] = "password"
        st.session_state["auth_email"] = "urgence@dbau.local"
        st.session_state["is_admin"] = True

        with st.sidebar:
            st.caption("Mode urgence actif")
            st.caption("Role : Administrateur")
            _logout_button("Quitter le mode urgence")

        return {"method": "password", "email": "urgence@dbau.local", "is_admin": True}

    st.markdown(
        """
        <style>
            [data-testid="stAppViewContainer"] {
                background:
                    radial-gradient(circle at top left, rgba(0, 135, 81, 0.14), transparent 28rem),
                    linear-gradient(135deg, #f7faf7 0%, #eef4f1 48%, #fff8df 100%);
            }

            section[data-testid="stSidebar"] {
                display: none;
            }

            #MainMenu,
            footer,
            [data-testid="stToolbar"],
            [data-testid="stDecoration"],
            [data-testid="viewerBadge"],
            [data-testid="stStatusWidget"],
            [data-testid="manage-app-button"],
            [data-testid="stAppDeployButton"],
            .stDeployButton {
                display: none !important;
                visibility: hidden !important;
                height: 0 !important;
            }

            header[data-testid="stHeader"] {
                background: transparent !important;
            }

            .auth-brand {
                padding-top: 10vh;
            }

            .auth-kicker {
                color: #008751;
                font-size: 0.82rem;
                font-weight: 800;
                letter-spacing: 0.08em;
                text-transform: uppercase;
                margin-bottom: 1rem;
            }

            .auth-brand h1 {
                color: #17211c;
                font-size: clamp(2.15rem, 4vw, 4rem);
                line-height: 1.02;
                margin: 0 0 1rem 0;
                letter-spacing: 0;
            }

            .auth-brand p {
                color: #526057;
                font-size: 1.08rem;
                line-height: 1.6;
                max-width: 34rem;
                margin: 0;
            }

            .auth-panel-title {
                color: #17211c;
                font-size: 1.18rem;
                font-weight: 800;
                margin-bottom: 0.25rem;
            }

            .auth-panel-subtitle {
                color: #66746c;
                font-size: 0.92rem;
                margin-bottom: 1.25rem;
            }

            .auth-divider {
                display: flex;
                align-items: center;
                gap: 0.75rem;
                color: #7a867f;
                font-size: 0.78rem;
                font-weight: 700;
                margin: 1.2rem 0 1rem 0;
                text-transform: uppercase;
            }

            .auth-divider::before,
            .auth-divider::after {
                content: "";
                height: 1px;
                background: rgba(23, 33, 28, 0.12);
                flex: 1;
            }

            .auth-lock {
                display: inline-flex;
                align-items: center;
                justify-content: center;
                width: 2.75rem;
                height: 2.75rem;
                border-radius: 8px;
                background: #008751;
                color: white;
                font-size: 1.45rem;
                margin-bottom: 1rem;
            }

            .auth-footer {
                color: #6d7871;
                font-size: 0.82rem;
                margin-top: 1rem;
                line-height: 1.5;
            }

            .stButton > button {
                border-radius: 8px;
                min-height: 2.8rem;
                font-weight: 750;
            }

            .stTextInput input {
                border-radius: 8px;
            }

            @media (max-width: 820px) {
                .auth-brand {
                    padding-top: 1.5rem;
                }
            }
        </style>
        """,
        unsafe_allow_html=True,
    )

    _, content, _ = st.columns([0.08, 0.84, 0.08])
    with content:
        brand_col, login_col = st.columns([1.1, 0.9], gap="large")

        with brand_col:
            st.markdown(
                """
                <div class="auth-brand">
                    <div class="auth-lock">DB</div>
                    <div class="auth-kicker">DBAU - Acces restreint</div>
                    <h1>CNaBAU<br>Session securisee</h1>
                    <p>Les donnees de bourse restent accessibles uniquement aux comptes autorises par la DBAU.</p>
                </div>
                """,
                unsafe_allow_html=True,
            )

        with login_col:
            st.markdown('<div style="height:10vh"></div>', unsafe_allow_html=True)
            with st.container(border=True):
                st.markdown('<div class="auth-panel-title">Connexion</div>', unsafe_allow_html=True)
                st.markdown(
                    "<div class=\"auth-panel-subtitle\">Compte institutionnel ou acces temporaire d'urgence.</div>",
                    unsafe_allow_html=True,
                )

                if hasattr(st, "login"):
                    if st.button("Connexion Microsoft DBAU", type="primary", use_container_width=True):
                        st.login(oidc_provider)
                else:
                    st.warning("La version de Streamlit installee ne supporte pas st.login().")

                st.markdown('<div class="auth-divider">Secours</div>', unsafe_allow_html=True)
                candidate = st.text_input(
                    "Mot de passe d'urgence",
                    type="password",
                    label_visibility="collapsed",
                )
                if st.button("Entrer avec le mot de passe", use_container_width=True):
                    if _password_matches(candidate, emergency_password):
                        st.session_state["password_authenticated"] = True
                        st.rerun()
                    st.error("Mot de passe incorrect.")

                st.markdown(
                    "<div class=\"auth-footer\">Acces journalier recommande par compte Microsoft. Le mot de passe d'urgence doit rester temporaire.</div>",
                    unsafe_allow_html=True,
                )

    st.stop()


def require_admin(message: str = "Action reservee aux administrateurs.") -> bool:
    if st.session_state.get("is_admin"):
        return True
    st.warning(message)
    return False
