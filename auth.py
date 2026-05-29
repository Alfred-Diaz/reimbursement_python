import os
import streamlit as st


DEFAULT_ADMIN_USERNAME = "admin"
DEFAULT_ADMIN_PASSWORD = "change-this-password"
DEFAULT_USER_USERNAME = "user"
DEFAULT_USER_PASSWORD = "change-this-password"


def get_secret(name, default=""):
    try:
        return st.secrets.get(name, os.environ.get(name, default))
    except Exception:
        return os.environ.get(name, default)


def get_users_from_list_secrets():
    try:
        users_list = st.secrets.get("users", [])
    except Exception:
        users_list = []

    users = {}
    for item in users_list:
        username = str(item.get("username", "")).strip()
        password = str(item.get("password", ""))
        role = str(item.get("role", "user")).strip().lower()

        if not username or not password:
            continue
        if role not in {"admin", "user"}:
            role = "user"

        users[username] = {
            "password": password,
            "role": role,
        }

    return users


def get_users_from_legacy_secrets():
    admin_username = get_secret("ADMIN_USERNAME", DEFAULT_ADMIN_USERNAME)
    admin_password = get_secret("ADMIN_PASSWORD", DEFAULT_ADMIN_PASSWORD)
    user_username = get_secret("USER_USERNAME", DEFAULT_USER_USERNAME)
    user_password = get_secret("USER_PASSWORD", DEFAULT_USER_PASSWORD)

    return {
        admin_username: {
            "password": admin_password,
            "role": "admin",
        },
        user_username: {
            "password": user_password,
            "role": "user",
        },
    }


def get_users():
    users = get_users_from_list_secrets()
    if users:
        return users
    return get_users_from_legacy_secrets()


def secrets_configured():
    if get_users_from_list_secrets():
        return True

    admin_password = get_secret("ADMIN_PASSWORD", "")
    user_password = get_secret("USER_PASSWORD", "")
    return bool(admin_password and user_password)


def using_default_passwords():
    users = get_users()
    for user in users.values():
        if user.get("password") == DEFAULT_ADMIN_PASSWORD:
            return True
        if user.get("password") == DEFAULT_USER_PASSWORD:
            return True
    return False


def init_auth_state():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if "username" not in st.session_state:
        st.session_state.username = ""
    if "role" not in st.session_state:
        st.session_state.role = "guest"


def login_form():
    init_auth_state()

    if st.session_state.authenticated:
        return True

    st.subheader("Login Required")
    st.caption("Use your assigned username and password to access the cleanup tool.")

    if not secrets_configured():
        st.warning(
            "Login secrets are not fully configured. Add [[users]] entries in Streamlit Secrets."
        )

    if using_default_passwords():
        st.error(
            "Default passwords are still active. Change them in Streamlit Secrets before using this app for real data."
        )

    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Login")

    if submitted:
        users = get_users()
        user = users.get(username)

        if user and password == user["password"]:
            st.session_state.authenticated = True
            st.session_state.username = username
            st.session_state.role = user["role"]
            st.rerun()
        else:
            st.error("Invalid username or password.")

    return False


def logout_button():
    init_auth_state()

    with st.sidebar:
        if st.session_state.authenticated:
            st.write(f"Logged in as: {st.session_state.username}")
            st.write(f"Role: {st.session_state.role}")
            if st.button("Logout"):
                st.session_state.authenticated = False
                st.session_state.username = ""
                st.session_state.role = "guest"
                st.rerun()


def current_role():
    init_auth_state()
    return st.session_state.role


def is_admin():
    return current_role() == "admin"
