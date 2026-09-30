"""
Fuel Levy tab for the Data Processing Toolbox
=============================================
Scrapes the "Metropolitan courier fuel levy" table from the Direct Couriers
portal and shows it as a tab that only appears after an admin login.

Nothing sensitive lives in this file. Everything comes from Streamlit secrets.

Who can see the tab - reuses your existing [users] / [access] tables. A user needs
"fuel" in their access list:

    [access]
    Timm1999 = "billing, container, fuel"

The Direct Couriers portal login:

    [directcouriers]
    username = "..."
    password = "..."
    # optional
    login_url      = "https://appsrv.directcouriers.com.au/online_mel/"
    username_field = ""     # only if auto-detection picks the wrong field
    password_field = ""

Needs in requirements.txt:  requests, beautifulsoup4, lxml
"""

import base64
import hashlib
import hmac
import io
from datetime import datetime

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup

DEFAULT_LOGIN_URL = "https://appsrv.directcouriers.com.au/online_mel/"
METRO = "Metro courier levy %"
ROW_LABELS = {
    "metropolitan": METRO,
    "intrastate": "Intrastate/Interstate levy %",
    "international": "International levy %",
}
HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0 Safari/537.36")
}
BLOCK_MSG = ("Direct Couriers' security service blocked the request (common for servers "
             "in the cloud). Save the Fuel Levy page from your own browser (Cmd/Ctrl+S) "
             "and upload it below instead.")


# ------------------------------------------------------------------ secrets
def _secret(name, default=None):
    try:
        return st.secrets[name]
    except Exception:
        return default


# ------------------------------------------------------------------ login gate
FUEL_PERMISSION = "fuel"     # must appear in the user's [access] entry
MAX_FAILS = 5                # failed attempts per browser session


def _dc() -> dict:
    """The [directcouriers] secrets table."""
    return dict(_secret("directcouriers") or {})


def _verify_password(password: str, stored: str) -> bool:
    """Check a password against 'pbkdf2_sha256$iterations$salt$hash'.
    The salt is tried as plain text and as hex bytes, and the hash as hex and as
    base64, because hand-rolled pbkdf2 storage varies."""
    try:
        algo, iters, salt, want = str(stored).strip().split("$")
        iters = int(iters)
    except ValueError:
        return False
    if algo != "pbkdf2_sha256":
        return False
    salts = [salt.encode()]
    try:
        salts.append(bytes.fromhex(salt))
    except ValueError:
        pass
    for sb in salts:
        dk = hashlib.pbkdf2_hmac("sha256", password.encode(), sb, iters)
        if hmac.compare_digest(dk.hex(), want.lower()) or \
           hmac.compare_digest(base64.b64encode(dk).decode(), want):
            return True
    return False


def _user_can_use_fuel(username: str, password: str) -> bool:
    users = {str(k).lower(): v for k, v in dict(_secret("users") or {}).items()}
    access = {str(k).lower(): v for k, v in dict(_secret("access") or {}).items()}
    name = (username or "").strip().lower()
    stored = users.get(name)
    if not stored or not _verify_password(password or "", stored):
        return False
    perms = {p.strip().lower() for p in str(access.get(name, "")).split(",")}
    return FUEL_PERMISSION in perms


def fuel_unlocked() -> bool:
    return bool(st.session_state.get("fuel_ok"))


def admin_login_box():
    """Call once, above st.tabs(). Shows a small locked expander, or a log-out button."""
    if fuel_unlocked():
        st.caption(f"🔓 Admin tools unlocked ({st.session_state.get('fuel_user', '')})")
        if st.button("Log out of admin tools", key="fuel_logout"):
            for k in ("fuel_ok", "fuel_user", "fuel_df"):
                st.session_state.pop(k, None)
            st.rerun()
        return
    with st.expander("🔒 Admin login"):
        if st.session_state.get("fuel_fails", 0) >= MAX_FAILS:
            st.error("Too many failed attempts. Refresh the page to try again.")
            return
        with st.form("fuel_login", clear_on_submit=True):
            u = st.text_input("Username", key="fuel_u")
            p = st.text_input("Password", type="password", key="fuel_p")
            go = st.form_submit_button("Log in", key="fuel_go")
        if go:
            if _user_can_use_fuel(u, p):
                st.session_state["fuel_ok"] = True
                st.session_state["fuel_user"] = u.strip()
                st.session_state["fuel_fails"] = 0
                st.rerun()
            else:
                st.session_state["fuel_fails"] = st.session_state.get("fuel_fails", 0) + 1
                st.error("Incorrect login, or this login doesn't have access.")


# ------------------------------------------------------------------ scraping
def _login_form(soup):
    for form in soup.find_all("form"):
        if form.find("input", {"type": "password"}):
            return form
    return None


def _guess_user_field(inputs) -> str:
    """Pick the username box: visible text-type inputs only, keyword match first."""
    cands = [i for i in inputs
             if i.get("name") and (i.get("type") or "text").lower() in ("text", "email", "tel", "number")]
    for i in cands:
        attrs = " ".join(str(i.get(a, "")) for a in ("name", "id", "placeholder")).lower()
        if any(k in attrs for k in ("user", "login", "logon", "account", "email", "cust")):
            return i["name"]
    return cands[0]["name"] if cands else ""


def _blocked(resp) -> bool:
    t = resp.text.lower()
    return resp.status_code == 403 or ("access denied" in t and "incident id" in t)


def inspect_login_form(login_url: str):
    """Field names/types on the login form (values are not shown - they can hold tokens)."""
    r = requests.get(login_url, headers=HEADERS, timeout=30)
    if _blocked(r):
        raise RuntimeError(BLOCK_MSG)
    form = _login_form(BeautifulSoup(r.text, "lxml"))
    if form is None:
        raise RuntimeError("No password field found at that address - is it the login page?")
    return pd.DataFrame([{"name": i.get("name"), "type": i.get("type", "text")}
                         for i in form.find_all("input")])


def fetch_levy_html(username, password, login_url, user_field="", pass_field="") -> str:
    """Log in, then follow the session-tokenised link to the Fuel Levy page."""
    s = requests.Session()
    s.headers.update(HEADERS)

    page = s.get(login_url, timeout=30)
    if _blocked(page):
        raise RuntimeError(BLOCK_MSG)
    page.raise_for_status()
    form = _login_form(BeautifulSoup(page.text, "lxml"))
    if form is None:
        raise RuntimeError("Couldn't find a login form at the portal login address.")

    inputs = form.find_all("input")
    user_field = user_field or _guess_user_field(inputs)
    pass_field = pass_field or next(
        (i["name"] for i in inputs
         if (i.get("type") or "").lower() == "password" and i.get("name")), "")
    if not (user_field and pass_field):
        raise RuntimeError("Couldn't work out the login field names. Use 'Inspect login form' "
                           "below, then set username_field / password_field under [directcouriers] in secrets.")

    payload = {i["name"]: i.get("value", "") for i in inputs
               if i.get("name") and (i.get("type") or "text").lower() not in ("submit", "button", "image")}
    payload[user_field] = username
    payload[pass_field] = password
    for i in inputs:
        if (i.get("type") or "").lower() == "submit" and i.get("name"):
            payload[i["name"]] = i.get("value", "")

    action = requests.compat.urljoin(page.url, form.get("action") or page.url)
    method = (form.get("method") or "post").lower()
    resp = (s.post(action, data=payload, timeout=30) if method == "post"
            else s.get(action, params=payload, timeout=30))
    if _blocked(resp):
        raise RuntimeError(BLOCK_MSG)
    resp.raise_for_status()

    body = resp.text
    if _login_form(BeautifulSoup(body, "lxml")) is not None:
        raise RuntimeError("Login failed - check the username / password under [directcouriers] in the secrets.")
    if "for week starting" in body.lower():
        return body

    menu = BeautifulSoup(body, "lxml")
    link = next((a["href"] for a in menu.find_all("a", href=True)
                 if "fuellevy" in a["href"].lower()
                 or "fuel levy" in a.get_text(" ", strip=True).lower()), None)
    if not link:
        raise RuntimeError("Logged in, but couldn't find a Fuel Levy link on the page.")
    levy = s.get(requests.compat.urljoin(resp.url, link), timeout=30)
    if _blocked(levy):
        raise RuntimeError(BLOCK_MSG)
    levy.raise_for_status()
    return levy.text


# ------------------------------------------------------------------ parsing
def _cell_text(c) -> str:
    return " ".join(c.get_text(" ", strip=True).split())


def _parse_week(text):
    for fmt in ("%d %b %y", "%d %b %Y", "%d/%m/%Y", "%d/%m/%y", "%d %B %Y"):
        try:
            return datetime.strptime(text.strip(), fmt).date()
        except ValueError:
            continue
    return None


def _to_float(text):
    try:
        return float(text.replace("%", "").replace(",", "").strip())
    except ValueError:
        return None


def parse_levies(html: str) -> pd.DataFrame:
    """Find the table by its 'For Week Starting' row; match rows by their label words."""
    soup = BeautifulSoup(html, "lxml")
    table = header_row = None
    for t in soup.find_all("table"):
        for row in t.find_all("tr"):
            cells = row.find_all(["td", "th"])
            if cells and "week starting" in _cell_text(cells[0]).lower():
                table, header_row = t, row
                break
        if table is not None:
            break
    if table is None:
        raise ValueError("Couldn't find the fuel levy table (no 'For Week Starting' row). "
                         "Is this the Fuel Levy page?")

    weeks = [_parse_week(_cell_text(c)) for c in header_row.find_all(["td", "th"])[1:]]
    data = {}
    for row in table.find_all("tr"):
        cells = row.find_all(["td", "th"])
        if len(cells) < 2:
            continue
        label = _cell_text(cells[0]).lower()
        for key, name in ROW_LABELS.items():
            if key in label:
                data[name] = [_to_float(_cell_text(c)) for c in cells[1:]]
    if METRO not in data:
        raise ValueError("Found the table but not the 'Metropolitan courier fuel levy' row.")

    rows = []
    for i, week in enumerate(weeks):
        if week is None:
            continue
        rec = {"Week starting": week}
        for name, values in data.items():
            rec[name] = values[i] if i < len(values) else None
        rows.append(rec)
    if not rows:
        raise ValueError("Found the table but couldn't read any week dates from it.")
    return pd.DataFrame(rows).sort_values("Week starting", ascending=False).reset_index(drop=True)


def merge_history(df: pd.DataFrame, prev_bytes) -> pd.DataFrame:
    """Add weeks from a previously downloaded history CSV (newest scrape wins)."""
    if not prev_bytes:
        return df
    old = pd.read_csv(io.BytesIO(prev_bytes))
    old["Week starting"] = pd.to_datetime(old["Week starting"], errors="coerce").dt.date
    old = old.dropna(subset=["Week starting"])
    both = pd.concat([df, old], ignore_index=True).drop_duplicates("Week starting", keep="first")
    return both.sort_values("Week starting", ascending=False).reset_index(drop=True)


# ------------------------------------------------------------------ UI
def _show_df(obj):
    try:
        st.dataframe(obj, hide_index=True, width="stretch")
    except TypeError:                                   # older Streamlit
        st.dataframe(obj, hide_index=True, use_container_width=True)


def render_fuel_levy_tab():
    st.write("Pulls the weekly **Metropolitan courier fuel levy** (plus intrastate/interstate "
             "and international) from the Direct Couriers portal.")
    st.caption("🔒 Nothing is stored. Download the history CSV and re-upload it next time to "
               "keep a running history beyond the weeks shown on the site.")

    fetch = st.button("🔄 Fetch latest from Direct Couriers", type="primary", key="fuel_fetch")

    with st.expander("Or upload a saved copy of the Fuel Levy page instead"):
        saved = st.file_uploader("Fuel Levy page (.html) — save it from your browser with Cmd/Ctrl+S",
                                 type=["html", "htm"], key="fuel_saved")
    prev = st.file_uploader("Previous fuel_levy_history.csv (optional) — merges into the history",
                            type="csv", key="fuel_prev")

    df = None
    html = None
    if fetch:
        cfg = _dc()
        user, pw = cfg.get("username"), cfg.get("password")
        if not (user and pw):
            st.error("The [directcouriers] username / password are not set in the app secrets.")
        else:
            with st.spinner("Logging in to Direct Couriers..."):
                try:
                    html = fetch_levy_html(
                        str(user), str(pw),
                        cfg.get("login_url", DEFAULT_LOGIN_URL),
                        str(cfg.get("username_field", "") or ""),
                        str(cfg.get("password_field", "") or ""))
                except Exception as e:
                    st.error(f"{e}")
    elif saved is not None:
        html = saved.getvalue().decode("utf8", errors="ignore")

    if html:
        try:
            df = parse_levies(html)
            st.session_state["fuel_df"] = df
        except ValueError as e:
            st.error(str(e))
    else:
        df = st.session_state.get("fuel_df")

    with st.expander("Troubleshooting: inspect the login form"):
        st.caption("Lists the field names on the portal's login page (no values shown).")
        if st.button("Inspect login form", key="fuel_inspect"):
            try:
                _show_df(inspect_login_form(_dc().get("login_url", DEFAULT_LOGIN_URL)))
            except Exception as e:
                st.error(f"{e}")

    if df is None or df.empty:
        st.info("Click **Fetch latest** (or upload a saved page) to load the levy table.")
        return

    try:
        hist = merge_history(df, prev.getvalue() if prev is not None else None)
    except Exception as e:
        st.warning(f"Couldn't merge that history file ({e}). Showing this scrape only.")
        hist = df

    cur = df.iloc[0]
    cur_val = cur[METRO]
    delta = None
    if len(df) > 1 and pd.notna(cur_val) and pd.notna(df.iloc[1][METRO]):
        delta = f"{cur_val - df.iloc[1][METRO]:+.2f} pts vs previous week"
    st.divider()
    st.metric(f"Metro courier fuel levy — week starting {cur['Week starting']:%d/%m/%Y}",
              f"{cur_val:.2f}%" if pd.notna(cur_val) else "n/a", delta=delta)

    fmt = {c: "{:.2f}" for c in hist.columns if c != "Week starting"}
    fmt["Week starting"] = lambda d: d.strftime("%d/%m/%Y")
    _show_df(hist.style.format(fmt, na_rep="–"))

    chart = hist.copy()
    chart["Week starting"] = pd.to_datetime(chart["Week starting"])
    st.line_chart(chart.sort_values("Week starting").set_index("Week starting")[[METRO]])

    c1, c2 = st.columns(2)
    with c1:
        st.download_button("⬇️ Latest scrape (CSV)", df.to_csv(index=False).encode(),
                           "fuel_levy_latest.csv", "text/csv", key="fuel_dl_latest")
    with c2:
        st.download_button("⬇️ History (CSV) — re-upload next time", hist.to_csv(index=False).encode(),
                           "fuel_levy_history.csv", "text/csv", key="fuel_dl_hist")
