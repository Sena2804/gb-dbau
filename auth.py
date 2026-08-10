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

    st.title("Acces securise DBAU")
    st.info("Authentification obligatoire avant d'acceder aux donnees.")

    col_oidc, col_password = st.columns(2)

    with col_oidc:
        st.subheader("Compte DBAU")
        if hasattr(st, "login"):
            if st.button("Connexion DBAU", type="primary", use_container_width=True):
                st.login(oidc_provider)
        else:
            st.warning("La version de Streamlit installee ne supporte pas st.login().")

    with col_password:
        st.subheader("Acces d'urgence")
        candidate = st.text_input("Mot de passe", type="password")
        if st.button("Entrer", use_container_width=True):
            if _password_matches(candidate, emergency_password):
                st.session_state["password_authenticated"] = True
                st.rerun()
            st.error("Mot de passe incorrect.")

    st.stop()


def require_admin(message: str = "Action reservee aux administrateurs.") -> bool:
    if st.session_state.get("is_admin"):
        return True
    st.warning(message)
    return False
