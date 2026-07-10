"""
FastAPI entry point — Mannings Social Dashboard.
"""

import logging
from pathlib import Path

from fastapi import FastAPI, Request, Query, Form
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.responses import HTMLResponse, RedirectResponse
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.middleware.sessions import SessionMiddleware

from app.config import settings, BASE_DIR
from app.db.reader import get_available_periods, default_period, MONTH_NAMES
from app.api import data as data_api
from app.api import chat as chat_api

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(name)s] %(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

app = FastAPI(title=settings.app_name)

# ── Middleware (order matters: last added = outermost = runs first on request) ──

class NoCacheMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request, call_next):
        response = await call_next(request)
        if "/static/" in request.url.path or "/competitors/" in request.url.path:
            response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
            response.headers["Pragma"] = "no-cache"
            response.headers["Expires"] = "0"
        return response


_PUBLIC_PATHS = {"/login", "/logout"}
_PUBLIC_PREFIXES = ("/static/", "/competitors/")


class AuthMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request, call_next):
        path = request.url.path
        if path in _PUBLIC_PATHS or any(path.startswith(p) for p in _PUBLIC_PREFIXES):
            return await call_next(request)
        if not request.session.get("authed"):
            return RedirectResponse("/login", status_code=303)
        return await call_next(request)


# Add innermost first
app.add_middleware(AuthMiddleware)
app.add_middleware(SessionMiddleware, secret_key=settings.session_secret)
app.add_middleware(NoCacheMiddleware)

app.mount("/static", StaticFiles(directory=str(BASE_DIR / "app" / "static")), name="static")

competitors_dir = Path(settings.competitors_dir)
if competitors_dir.exists():
    app.mount("/competitors", StaticFiles(directory=str(competitors_dir)), name="competitors")

templates = Jinja2Templates(directory=str(BASE_DIR / "app" / "templates"))

app.include_router(data_api.router)
app.include_router(chat_api.router)


# ── Auth routes ──

@app.get("/login", response_class=HTMLResponse)
async def login_page(request: Request, error: str = ""):
    return templates.TemplateResponse(
        request=request,
        name="login.html",
        context={"error": error},
    )


@app.post("/login")
async def login_submit(request: Request, passcode: str = Form(...)):
    if passcode == settings.app_passcode:
        request.session["authed"] = True
        return RedirectResponse("/", status_code=303)
    return RedirectResponse("/login?error=1", status_code=303)


@app.get("/logout")
async def logout(request: Request):
    request.session.clear()
    return RedirectResponse("/login", status_code=303)


# ── Dashboard routes ──

@app.get("/api/health")
async def health_check():
    import os, traceback
    from app.config import settings, BASE_DIR
    from app.db.reader import get_available_periods, get_period_data, default_period
    excel = Path(settings.excel_path)
    comp = Path(settings.competitors_dir)
    static_dir = BASE_DIR / "app" / "static" / "js"

    # Test actual data loading
    periods_ok = False
    periods_count = 0
    api_error = None
    try:
        periods = get_available_periods()
        periods_count = len(periods)
        periods_ok = True
    except Exception as e:
        api_error = str(e)

    # Test period data
    data_ok = False
    sheet_count = 0
    try:
        y, m = default_period()
        pd_obj = get_period_data(y, m)
        sheet_count = len(pd_obj.sheets)
        data_ok = True
    except Exception as e:
        api_error = str(e) + " | " + traceback.format_exc()[-200:]

    return {
        "excel_path": settings.excel_path,
        "excel_exists": excel.exists(),
        "excel_size": excel.stat().st_size if excel.exists() else 0,
        "competitors_exists": comp.exists(),
        "static_js_exists": static_dir.exists(),
        "static_js_files": [f.name for f in static_dir.iterdir()] if static_dir.exists() else [],
        "base_dir": str(BASE_DIR),
        "cwd": os.getcwd(),
        "periods_ok": periods_ok,
        "periods_count": periods_count,
        "data_ok": data_ok,
        "sheet_count": sheet_count,
        "error": api_error,
    }


@app.get("/", response_class=HTMLResponse)
async def page_fb_page(request: Request, year: int = Query(0), month: int = Query(0)):
    if year == 0 or month == 0:
        year, month = default_period()
    periods = get_available_periods()
    return templates.TemplateResponse(
        request=request,
        name="dashboard.html",
        context={
            "page": "fb_page",
            "page_title": "FB Page",
            "year": year,
            "month": month,
            "period_label": f"{MONTH_NAMES.get(month, '')} {year}",
            "periods": periods,
        },
    )


@app.get("/fb-posts", response_class=HTMLResponse)
async def page_fb_posts(request: Request, year: int = Query(0), month: int = Query(0)):
    if year == 0 or month == 0:
        year, month = default_period()
    periods = get_available_periods()
    return templates.TemplateResponse(
        request=request,
        name="dashboard.html",
        context={
            "page": "fb_posts",
            "page_title": "FB Posts",
            "year": year,
            "month": month,
            "period_label": f"{MONTH_NAMES.get(month, '')} {year}",
            "periods": periods,
        },
    )


@app.get("/instagram", response_class=HTMLResponse)
async def page_instagram(request: Request, year: int = Query(0), month: int = Query(0)):
    if year == 0 or month == 0:
        year, month = default_period()
    periods = get_available_periods()
    return templates.TemplateResponse(
        request=request,
        name="dashboard.html",
        context={
            "page": "instagram",
            "page_title": "Instagram",
            "year": year,
            "month": month,
            "period_label": f"{MONTH_NAMES.get(month, '')} {year}",
            "periods": periods,
        },
    )


@app.get("/linkedin", response_class=HTMLResponse)
async def page_linkedin(request: Request, year: int = Query(0), month: int = Query(0)):
    if year == 0 or month == 0:
        year, month = default_period()
    periods = get_available_periods()
    return templates.TemplateResponse(
        request=request,
        name="dashboard.html",
        context={
            "page": "linkedin",
            "page_title": "LinkedIn",
            "year": year,
            "month": month,
            "period_label": f"{MONTH_NAMES.get(month, '')} {year}",
            "periods": periods,
        },
    )


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app.main:app", host="0.0.0.0", port=8000, reload=True)
