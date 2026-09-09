"""Serves the built ASME Ops single-page app under ``/app`` (and ``/auth/login``)."""

from __future__ import annotations

from pathlib import Path

from flask import Blueprint, current_app, redirect, send_from_directory

from asme.config import settings

bp = Blueprint("ops_web", __name__)

MISSING_BUILD = """<!doctype html><meta charset="utf-8"><title>ASME Ops</title>
<style>body{font-family:system-ui,sans-serif;max-width:640px;margin:10vh auto;color:#111827;line-height:1.5}code{background:#f3f6f9;padding:2px 6px;border-radius:4px}</style>
<h1>ASME Ops is not built yet</h1>
<p>The operations frontend lives in <code>apps/ops-web</code>. Build it once and this route will serve it:</p>
<pre><code>cd apps/ops-web
npm ci
npm run build</code></pre>
<p>For development run <code>npm run dev</code> there instead and open <code>http://localhost:5173/app</code>.</p>
"""


def dist_dir() -> Path:
    configured = settings().ops_web_dist
    return Path(configured) if configured else Path(current_app.root_path).parent / "apps" / "ops-web" / "dist"


@bp.get("/app")
@bp.get("/app/")
@bp.get("/app/<path:path>")
def spa(path: str = ""):
    dist = dist_dir()
    if path and (dist / path).is_file():
        response = send_from_directory(dist, path)
        if path.startswith("assets/"):
            response.headers["Cache-Control"] = "public, max-age=31536000, immutable"
        return response
    index = dist / "index.html"
    if not index.is_file():
        return MISSING_BUILD, 200, {"Content-Type": "text/html; charset=utf-8"}
    response = send_from_directory(dist, "index.html")
    response.headers["Cache-Control"] = "no-store"
    return response


@bp.get("/auth/login")
def auth_login_alias():
    return redirect("/app/login")
