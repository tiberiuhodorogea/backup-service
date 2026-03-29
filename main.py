"""
Backup Service — FastAPI application.

Run standalone:  python main.py
Or via uvicorn:  uvicorn main:app --host 0.0.0.0 --port 8550
"""

import asyncio
import logging
import os
from contextlib import asynccontextmanager
from datetime import datetime, timedelta

from fastapi import FastAPI, WebSocket, WebSocketDisconnect
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from apscheduler.schedulers.background import BackgroundScheduler

import database as db
import backup_engine as engine

log = logging.getLogger(__name__)

STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")

# ── Scheduler ───────────────────────────────────────────
scheduler = BackgroundScheduler(daemon=True)


def _next_run_dt(hour: int, minute: int, freq_days: int):
    """Return the next datetime matching hour:minute, at least 1 min in the future."""
    now = datetime.now()
    candidate = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
    if candidate <= now:
        candidate += timedelta(days=freq_days)
    return candidate


def refresh_scheduler():
    """Re-read settings and reconfigure the scheduled job."""
    settings = db.get_settings()
    enabled = settings.get("scheduler_enabled", "true") == "true"

    job = scheduler.get_job("backup_job")
    if job:
        scheduler.remove_job("backup_job")

    if enabled:
        freq = int(settings.get("frequency_days", "1"))
        h, m = (int(x) for x in settings.get("backup_time", "02:00").split(":"))
        next_dt = _next_run_dt(h, m, freq)
        scheduler.add_job(
            engine.start_backup_thread,
            "interval",
            days=freq,
            next_run_time=next_dt,
            id="backup_job",
            replace_existing=True,
        )
        log.info("Scheduler configured: every %d day(s) at %02d:%02d, next run: %s",
                 freq, h, m, next_dt.strftime("%Y-%m-%d %H:%M"))
    else:
        log.info("Scheduler disabled")


def get_next_run_time():
    job = scheduler.get_job("backup_job")
    if job and job.next_run_time:
        return job.next_run_time.strftime("%Y-%m-%d %H:%M:%S")
    return None


# ── Lifespan ────────────────────────────────────────────
@asynccontextmanager
async def lifespan(app: FastAPI):
    db.init_db()
    refresh_scheduler()
    scheduler.start()
    log.info("Application started")
    yield
    log.info("Application shutting down")
    scheduler.shutdown(wait=False)


app = FastAPI(title="Custodia", lifespan=lifespan)

# Serve static files
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


# ── Pydantic models ────────────────────────────────────
class PathBody(BaseModel):
    path: str
    label: str = ""


class ToggleBody(BaseModel):
    enabled: bool


class SettingsBody(BaseModel):
    frequency_days: int | None = None
    backup_time: str | None = None
    retention_count: int | None = None
    scheduler_enabled: bool | None = None


# ── Routes ──────────────────────────────────────────────
@app.get("/")
async def index():
    return FileResponse(os.path.join(STATIC_DIR, "index.html"))


# ── Folder browser ───────────────────────────────────────
@app.get("/api/browse")
async def browse_folder(path: str = ""):
    import string
    from pathlib import Path

    if not path:
        drives = []
        for letter in string.ascii_uppercase:
            drive = f"{letter}:\\"
            if os.path.exists(drive):
                drives.append({"name": drive, "path": drive})
        return {"path": "", "parent": None, "entries": drives}

    path = os.path.normpath(path)
    if not os.path.isdir(path):
        return {"path": path, "parent": None, "entries": []}

    parent = str(Path(path).parent)
    if parent == path:
        parent = ""

    entries = []
    try:
        for entry in sorted(os.scandir(path), key=lambda e: e.name.lower()):
            try:
                if entry.is_dir(follow_symlinks=False):
                    entries.append({"name": entry.name, "path": entry.path})
            except OSError:
                pass
    except PermissionError:
        pass

    return {"path": path, "parent": parent, "entries": entries}


# ── Status ──────────────────────────────────────────────
@app.get("/api/status")
async def api_status():
    data = engine.state.to_dict()
    data["next_scheduled"] = get_next_run_time()
    runs = db.get_runs(limit=1)
    if runs:
        last = runs[0]
        data["last_run_status"] = last["status"]
        data["last_run_time"] = last["completed_at"] or last["started_at"]
    return data


# ── Sources ─────────────────────────────────────────────
@app.get("/api/sources")
async def list_sources():
    return db.get_sources()


@app.post("/api/sources")
async def create_source(body: PathBody):
    label = body.label or os.path.basename(body.path.rstrip("/\\")) or body.path
    new_id = db.add_source(body.path, label)
    return {"id": new_id, "path": body.path, "label": label}


@app.delete("/api/sources/{source_id}")
async def delete_source(source_id: int):
    db.remove_source(source_id)
    return {"ok": True}


@app.patch("/api/sources/{source_id}")
async def patch_source(source_id: int, body: ToggleBody):
    db.toggle_source(source_id, body.enabled)
    return {"ok": True}


# ── Destinations ────────────────────────────────────────
@app.get("/api/destinations")
async def list_destinations():
    return db.get_destinations()


@app.post("/api/destinations")
async def create_destination(body: PathBody):
    label = body.label or os.path.basename(body.path.rstrip("/\\")) or body.path
    new_id = db.add_destination(body.path, label)
    return {"id": new_id, "path": body.path, "label": label}


@app.delete("/api/destinations/{dest_id}")
async def delete_destination(dest_id: int):
    db.remove_destination(dest_id)
    return {"ok": True}


@app.patch("/api/destinations/{dest_id}")
async def patch_destination(dest_id: int, body: ToggleBody):
    db.toggle_destination(dest_id, body.enabled)
    return {"ok": True}


# ── Settings ────────────────────────────────────────────
@app.get("/api/settings")
async def get_settings():
    return db.get_settings()


@app.put("/api/settings")
async def put_settings(body: SettingsBody):
    data = {k: v for k, v in body.model_dump().items() if v is not None}
    if "scheduler_enabled" in data:
        data["scheduler_enabled"] = "true" if data["scheduler_enabled"] else "false"
    if "retention_count" in data:
        data["retention_count"] = str(max(2, min(5, int(data["retention_count"]))))
    if "frequency_days" in data:
        data["frequency_days"] = str(max(1, int(data["frequency_days"])))
    db.update_settings(data)
    refresh_scheduler()
    return db.get_settings()


# ── Backup control ──────────────────────────────────────
@app.post("/api/backup/start")
async def start_backup():
    ok = engine.start_backup_thread()
    if not ok:
        log.warning("Backup start requested but one is already running")
        return {"ok": False, "error": "Backup already running"}
    log.info("Manual backup started via dashboard")
    return {"ok": True}


@app.post("/api/backup/cancel")
async def cancel_backup():
    engine.cancel_backup()
    log.info("Backup cancellation requested via dashboard")
    return {"ok": True}


# ── History & Logs ──────────────────────────────────────
@app.get("/api/history")
async def history(limit: int = 20):
    return db.get_runs(limit=limit)


@app.get("/api/logs")
async def logs(run_id: int = None, limit: int = 200, offset: int = 0):
    return db.get_logs(run_id=run_id, limit=limit, offset=offset)


# ── WebSocket ───────────────────────────────────────────
@app.websocket("/ws")
async def ws_endpoint(websocket: WebSocket):
    await websocket.accept()
    try:
        while True:
            data = engine.state.to_dict()
            data["next_scheduled"] = get_next_run_time()
            log_lines = engine.state.drain_logs()
            await websocket.send_json(
                {"type": "update", "state": data, "logs": log_lines}
            )
            await asyncio.sleep(0.5)
    except (WebSocketDisconnect, Exception):
        pass


# ── Entrypoint ──────────────────────────────────────────
if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8550, log_level="info")
