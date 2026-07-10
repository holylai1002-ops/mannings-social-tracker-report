"""
FastAPI entry point — Mannings Social Dashboard v2 (Calendar Picker).
"""

import logging
from pathlib import Path

from fastapi import FastAPI, Request, Query, Form
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.responses import HTMLResponse, RedirectResponse
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.middleware.sessions import SessionMiddleware

from app_v2.config import settings, BASE_DIR
from app_v2.db.range_reader import default_range, MONTH_NAMES, get_available_periods
from app_v2.api import data as data_api
from app_v2.api import chat as chat_api

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(name)s] %(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

app = FastAPI(title=settings.app_name)


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


app.add_middleware(AuthMiddleware)
app.add_middleware(SessionMiddleware, secret_key=settings.session_secret)
app.add_middleware(NoCacheMiddleware)

app.mount("/static", StaticFiles(directory=str(BASE_DIR / "app_v2" / "static")), name="static")

competitors_dir = Path(settings.competitors_dir)
if competitors_dir.exists():
    app.mount("/competitors", StaticFiles(directory=str(competitors_dir)), name="competitors")

templates = Jinja2Templates(directory=str(BASE_DIR / "app_v2" / "templates"))

app.include_router(data_api.router)
app.include_router(chat_api.router)


@app.get("/login", response_class=HTMLResponse)
async def login_page(request: Request, error: str = ""):
    return templates.TemplateResponse(request=request, name="login.html", context={"error": error})


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


def _ctx(page: str, page_title: str, start: str, end: str):
    periods = get_available_periods()
    ds, de = default_range()
    if not start:
        start = ds
    if not end:
        end = de
    return {
        "page": page,
        "page_title": page_title,
        "start_date": start,
        "end_date": end,
        "periods": periods,
    }


@app.get("/", response_class=HTMLResponse)
async def page_fb_page(request: Request, start: str = "", end: str = ""):
    return templates.TemplateResponse(request=request, name="dashboard.html",
                                      context=_ctx("fb_page", "FB Page", start, end))


@app.get("/fb-posts", response_class=HTMLResponse)
async def page_fb_posts(request: Request, start: str = "", end: str = ""):
    return templates.TemplateResponse(request=request, name="dashboard.html",
                                      context=_ctx("fb_posts", "FB Posts", start, end))


@app.get("/instagram", response_class=HTMLResponse)
async def page_instagram(request: Request, start: str = "", end: str = ""):
    return templates.TemplateResponse(request=request, name="dashboard.html",
                                      context=_ctx("instagram", "Instagram", start, end))


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app_v2.main:app", host="0.0.0.0", port=8010, reload=True)
