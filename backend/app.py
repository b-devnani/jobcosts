"""FastAPI app for the Job Cost Projection tool.

Two surfaces:

* **Generate** — upload a Procore budget-detail CSV, pick a project from the
  dropdown, and download the filled-in Job Cost Projection workbook.
* **Admin** — a small data grid to add / edit / remove projects (name plus the
  four milestone dates), protected by a shared admin password.
"""

from __future__ import annotations

import io
import logging
import os
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import Depends, FastAPI, Form, Header, HTTPException, UploadFile
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

from . import db
from .converter import ConversionError, convert_csv_to_workbook_bytes

STATIC_DIR = Path(__file__).resolve().parent / "static"
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "admin")

# Hard cap on the uploaded CSV so a huge file can't exhaust memory. The template
# only holds 146 data rows, so a real export is tiny; this is defence in depth.
MAX_CSV_BYTES = 10 * 1024 * 1024  # 10 MB

XLSX_MEDIA_TYPE = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

log = logging.getLogger("uvicorn.error")


@asynccontextmanager
async def lifespan(app: FastAPI):
    db.init_db()
    if ADMIN_PASSWORD == "admin":
        log.warning(
            "ADMIN_PASSWORD is the insecure default 'admin'. Set the "
            "ADMIN_PASSWORD environment variable to a strong value before "
            "exposing this tool to anyone else."
        )
    yield


app = FastAPI(title="Job Cost Projection Tool", version="1.0.0", lifespan=lifespan)


# --------------------------------------------------------------------------- #
# Auth (intentionally lightweight: a single shared admin password)
# --------------------------------------------------------------------------- #
def require_admin(x_admin_password: str | None = Header(default=None)) -> None:
    if not x_admin_password or x_admin_password != ADMIN_PASSWORD:
        raise HTTPException(status_code=401, detail="Admin authentication required.")


@app.post("/api/admin/login")
def admin_login(payload: dict) -> dict:
    if payload.get("password") != ADMIN_PASSWORD:
        raise HTTPException(status_code=401, detail="Incorrect admin password.")
    return {"ok": True}


# --------------------------------------------------------------------------- #
# Projects CRUD
# --------------------------------------------------------------------------- #
@app.get("/api/projects")
def get_projects() -> list[dict]:
    return db.list_projects()


@app.post("/api/projects", dependencies=[Depends(require_admin)])
def post_project(payload: dict) -> dict:
    try:
        return db.create_project(payload)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@app.put("/api/projects/{project_id}", dependencies=[Depends(require_admin)])
def put_project(project_id: int, payload: dict) -> dict:
    try:
        updated = db.update_project(project_id, payload)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    if updated is None:
        raise HTTPException(status_code=404, detail="Project not found.")
    return updated


@app.delete("/api/projects/{project_id}", dependencies=[Depends(require_admin)])
def remove_project(project_id: int) -> dict:
    if not db.delete_project(project_id):
        raise HTTPException(status_code=404, detail="Project not found.")
    return {"ok": True}


# --------------------------------------------------------------------------- #
# Generate the workbook
# --------------------------------------------------------------------------- #
@app.post("/api/generate")
async def generate(
    csv_file: UploadFile,
    project_id: int | None = Form(default=None),
    name: str | None = Form(default=None),
    orig_substantial_completion: str | None = Form(default=None),
    orig_final_completion: str | None = Form(default=None),
    current_substantial_completion: str | None = Form(default=None),
    current_final_completion: str | None = Form(default=None),
    contract_amount_last_pay_app: str | None = Form(default=None),
    month_last_pay_app: str | None = Form(default=None),
):
    """Build the workbook from the uploaded CSV plus project milestone data.

    Milestone data comes from the chosen project (``project_id``); any field can
    be overridden by an explicit form value, which also lets the tool run
    without a saved project at all.
    """
    if csv_file is None:
        raise HTTPException(status_code=400, detail="A CSV file is required.")

    raw = await csv_file.read()
    if not raw:
        raise HTTPException(status_code=400, detail="The uploaded CSV is empty.")
    if len(raw) > MAX_CSV_BYTES:
        raise HTTPException(
            status_code=413,
            detail=f"CSV file is too large (max {MAX_CSV_BYTES // (1024 * 1024)} MB).",
        )

    project = db.get_project(project_id) if project_id else None
    if project_id and project is None:
        raise HTTPException(status_code=404, detail="Selected project not found.")

    def pick(field: str, override: str | None):
        if override not in (None, ""):
            return override
        return project.get(field) if project else None

    resolved_name = pick("name", name) or "Job Cost Projection"

    try:
        xlsx_bytes, filename = convert_csv_to_workbook_bytes(
            raw,
            name=resolved_name,
            orig_substantial_completion=pick(
                "orig_substantial_completion", orig_substantial_completion
            ),
            orig_final_completion=pick(
                "orig_final_completion", orig_final_completion
            ),
            current_substantial_completion=pick(
                "current_substantial_completion", current_substantial_completion
            ),
            current_final_completion=pick(
                "current_final_completion", current_final_completion
            ),
            contract_amount_last_pay_app=pick(
                "contract_amount_last_pay_app", contract_amount_last_pay_app
            ),
            month_last_pay_app=pick("month_last_pay_app", month_last_pay_app),
        )
    except ConversionError as exc:
        raise HTTPException(status_code=422, detail=str(exc))

    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
    return StreamingResponse(
        io.BytesIO(xlsx_bytes), media_type=XLSX_MEDIA_TYPE, headers=headers
    )


# --------------------------------------------------------------------------- #
# Health check (used by hosting platforms)
# --------------------------------------------------------------------------- #
@app.get("/healthz")
def healthz() -> dict:
    return {"status": "ok"}


# --------------------------------------------------------------------------- #
# Static frontend
# --------------------------------------------------------------------------- #
@app.get("/")
def index() -> FileResponse:
    return FileResponse(STATIC_DIR / "index.html")


app.mount("/", StaticFiles(directory=STATIC_DIR), name="static")
