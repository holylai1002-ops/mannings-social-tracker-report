"""
FastAPI entry point — Mannings Social Dashboard.
"""

import logging
from pathlib import Path

from fastapi import FastAPI, Request, Query
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.responses import HTMLResponse
from starlette.middleware.base import BaseHTTPMiddleware

from app.config import settings, BASE_DIR
from app.db.reader import get_available_periods, default_period, MONTH_NAMES
from app.api import data as data_api
from app.api import chat as chat_api

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


app.add_middleware(NoCacheMiddleware)

app.mount("/static", StaticFiles(directory=str(BASE_DIR / "app" / "static")), name="static")

competitors_dir = Path(settings.competitors_dir)
if competitors_dir.exists():
    app.mount("/competitors", StaticFiles(directory=str(competitors_dir)), name="competitors")

templates = Jinja2Templates(directory=str(BASE_DIR / "app" / "templates"))

app.include_router(data_api.router)
app.include_router(chat_api.router)


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


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app.main:app", host="0.0.0.0", port=8000, reload=True)
