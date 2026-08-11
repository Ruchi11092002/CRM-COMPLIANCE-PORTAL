import hmac
import os
from functools import wraps

from flask import (
    Flask,
    redirect,
    request,
    session,
    url_for,
)

import Generate_compliance as portal


app = Flask(__name__)

app.secret_key = os.environ.get("SESSION_SECRET", "CHANGE-ME-IN-RENDER")

app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SECURE=True,
    SESSION_COOKIE_SAMESITE="Lax",
)


def login_required(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        if not session.get("authenticated"):
            return redirect(url_for("login"))
        return func(*args, **kwargs)

    return wrapper


@app.route("/login", methods=["GET", "POST"])
def login():
    error = ""

    if request.method == "POST":
        submitted_password = request.form.get("password", "")
        real_password = os.environ.get("PORTAL_PASSWORD", "")

        if real_password and hmac.compare_digest(
            submitted_password,
            real_password
        ):
            session.clear()
            session["authenticated"] = True
            return redirect(url_for("home"))

        error = "Incorrect password. Please try again."

    return f"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">

<title>CRM Compliance Portal</title>

<style>

* {{
    box-sizing: border-box;
}}

body {{
    margin: 0;
    font-family: "Segoe UI", Arial, sans-serif;
    min-height: 100vh;
    display: flex;
    align-items: center;
    justify-content: center;

    background:
        radial-gradient(
            circle at top left,
            rgba(31,78,121,.13),
            transparent 420px
        ),
        linear-gradient(
            135deg,
            #edf7ff,
            #f8fbff 50%,
            #fff8e8
        );
}}

.login-card {{
    width: min(420px, 92vw);
    background: white;
    padding: 30px;
    border-radius: 20px;
    border: 1px solid #dfe4ea;
    box-shadow: 0 24px 70px rgba(15,23,42,.16);
}}

h1 {{
    margin: 0;
    font-size: 25px;
    color: #172033;
}}

.subtitle {{
    color: #64748b;
    font-size: 13px;
    line-height: 1.55;
    margin: 10px 0 22px;
}}

label {{
    display: block;
    color: #475569;
    font-size: 12px;
    font-weight: 700;
    margin-bottom: 7px;
}}

input {{
    width: 100%;
    border: 1px solid #d7dbe2;
    border-radius: 11px;
    padding: 12px 13px;
    font-size: 15px;
    outline: none;
}}

input:focus {{
    border-color: #8bb4dd;
    box-shadow: 0 0 0 3px rgba(31,78,121,.12);
}}

button {{
    width: 100%;
    border: 0;
    border-radius: 11px;
    margin-top: 14px;
    padding: 12px 15px;
    background: #1f4e79;
    color: white;
    font-weight: 700;
    font-size: 14px;
    cursor: pointer;
}}

button:hover {{
    background: #183d61;
}}

.error {{
    color: #b91c1c;
    margin-top: 12px;
    min-height: 18px;
    font-size: 13px;
    font-weight: 600;
}}

.security {{
    color: #94a3b8;
    font-size: 11px;
    margin-top: 19px;
    text-align: center;
}}

</style>
</head>

<body>

<div class="login-card">

    <h1>CRM Compliance Portal</h1>

    <div class="subtitle">
        Authorized access only. Enter the portal password
        to view compliance status and documents.
    </div>

    <form method="POST">

        <label>Portal Password</label>

        <input
            type="password"
            name="password"
            placeholder="Enter password"
            required
            autofocus
        >

        <button type="submit">
            Open Portal
        </button>

    </form>

    <div class="error">
        {error}
    </div>

    <div class="security">
        CRM Compliance Management System
    </div>

</div>

</body>
</html>
"""


@app.route("/")
@login_required
def home():

    # For now this still reads the Excel/data from the repository.
    # We will replace this with Microsoft Graph in Stage 2.
    records = portal.load_records()

    html = portal.generate_html(records)

    # ---------------------------------------------------------
    # Disable the OLD browser-side password screen.
    # Authentication has already happened securely on Flask.
    # ---------------------------------------------------------

    html = html.replace(
        '<body class="locked">',
        '<body>'
    )

    html = html.replace(
        '<div id="portalContent" class="portal-content locked">',
        '<div id="portalContent" class="portal-content">'
    )

    html = html.replace(
        '<div id="loginOverlay" class="login-overlay">',
        '<div id="loginOverlay" class="login-overlay" '
        'style="display:none!important">'
    )

    # Keep the old button present so existing JS does not fail,
    # but hide it.
    old_button = (
        '<button class="btn ghost" id="lockPortalBtn" '
        'type="button">Lock Portal</button>'
    )

    new_button = (
        '<a class="btn ghost" href="/logout">'
        'Logout'
        '</a>'
        '<button class="btn ghost" id="lockPortalBtn" '
        'type="button" style="display:none">'
        'Lock Portal'
        '</button>'
    )

    html = html.replace(old_button, new_button)

    return html


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/health")
def health():
    return {
        "status": "ok",
        "application": "CRM Compliance Portal"
    }


if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 10000))
    )