"""
Greenhouse Controller Backend
-----------------------------
FastAPI + SQLite. Handles:
  - ESP32 device ingest (sensor readings -> control response)
  - Dashboard API (readings, settings, devices)
  - Static dashboard hosting

Run locally:
    pip install -r requirements.txt
    uvicorn main:app --host 0.0.0.0 --port 8000 --reload

Deploy: works as-is on Render / Railway / Fly.io (see README).
"""

import os
import sqlite3
import secrets
import math
from contextlib import contextmanager
from datetime import datetime, timezone, timedelta
from typing import Optional

from fastapi import FastAPI, HTTPException, Header, Depends
from fastapi.staticfiles import StaticFiles
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from pydantic import BaseModel

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
DB_PATH = os.environ.get("GREENHOUSE_DB", os.path.join(os.path.dirname(__file__), "greenhouse.db"))
DASHBOARD_USER = os.environ.get("DASHBOARD_USER", "admin")
DASHBOARD_PASS = os.environ.get("DASHBOARD_PASS", "changeme")

app = FastAPI(title="Greenhouse Controller API")
security = HTTPBasic()


# ---------------------------------------------------------------------------
# DB setup
# ---------------------------------------------------------------------------
@contextmanager
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db():
    with get_db() as db:
        db.executescript(
            """
            CREATE TABLE IF NOT EXISTS devices (
                id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                api_key TEXT UNIQUE NOT NULL,
                created_at TEXT NOT NULL,
                last_seen TEXT
            );

            CREATE TABLE IF NOT EXISTS readings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                device_id TEXT NOT NULL,
                ts TEXT NOT NULL,
                temperature REAL,
                humidity REAL,
                vpd REAL,
                FOREIGN KEY(device_id) REFERENCES devices(id)
            );

            CREATE INDEX IF NOT EXISTS idx_readings_device_ts
                ON readings(device_id, ts DESC);

            CREATE TABLE IF NOT EXISTS settings (
                device_id TEXT PRIMARY KEY,
                temp_min REAL DEFAULT 18,
                temp_max REAL DEFAULT 28,
                humidity_min REAL DEFAULT 50,
                humidity_max REAL DEFAULT 70,
                vpd_target REAL DEFAULT 1.0,
                light_on_hour INTEGER DEFAULT 6,
                light_off_hour INTEGER DEFAULT 22,
                manual_override INTEGER DEFAULT 0,
                override_fan INTEGER DEFAULT 0,
                override_heater INTEGER DEFAULT 0,
                override_humidifier INTEGER DEFAULT 0,
                override_lights INTEGER DEFAULT 0,
                FOREIGN KEY(device_id) REFERENCES devices(id)
            );
            """
        )


init_db()


# ---------------------------------------------------------------------------
# Models
# ---------------------------------------------------------------------------
class IngestPayload(BaseModel):
    device: str
    temperature: float
    humidity: float
    vpd: Optional[float] = None  # ESP32 can send its own, or let server compute


class SettingsUpdate(BaseModel):
    temp_min: Optional[float] = None
    temp_max: Optional[float] = None
    humidity_min: Optional[float] = None
    humidity_max: Optional[float] = None
    vpd_target: Optional[float] = None
    light_on_hour: Optional[int] = None
    light_off_hour: Optional[int] = None
    manual_override: Optional[bool] = None
    override_fan: Optional[bool] = None
    override_heater: Optional[bool] = None
    override_humidifier: Optional[bool] = None
    override_lights: Optional[bool] = None


class NewDevice(BaseModel):
    name: str


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def calc_vpd(temp_c: float, rh_percent: float) -> float:
    """Saturation vapour pressure (Tetens eq.) -> VPD in kPa."""
    svp = 0.61078 * math.exp((17.27 * temp_c) / (temp_c + 237.3))
    return round(svp * (1 - rh_percent / 100), 3)


def require_dashboard_auth(creds: HTTPBasicCredentials = Depends(security)):
    ok_user = secrets.compare_digest(creds.username, DASHBOARD_USER)
    ok_pass = secrets.compare_digest(creds.password, DASHBOARD_PASS)
    if not (ok_user and ok_pass):
        raise HTTPException(status_code=401, detail="Invalid credentials", headers={"WWW-Authenticate": "Basic"})
    return creds.username


def get_device_by_api_key(db, api_key: str):
    row = db.execute("SELECT * FROM devices WHERE api_key = ?", (api_key,)).fetchone()
    return row


def ensure_settings_row(db, device_id: str):
    db.execute("INSERT OR IGNORE INTO settings (device_id) VALUES (?)", (device_id,))


def compute_control(settings_row, temperature: float, humidity: float, vpd: float):
    """Threshold-based automation logic. Manual override short-circuits everything."""
    if settings_row["manual_override"]:
        return {
            "fan": bool(settings_row["override_fan"]),
            "heater": bool(settings_row["override_heater"]),
            "humidifier": bool(settings_row["override_humidifier"]),
            "lights": bool(settings_row["override_lights"]),
        }

    fan = temperature > settings_row["temp_max"] or humidity > settings_row["humidity_max"]
    heater = temperature < settings_row["temp_min"]
    humidifier = humidity < settings_row["humidity_min"]

    hour = datetime.now(timezone.utc).hour
    on_h, off_h = settings_row["light_on_hour"], settings_row["light_off_hour"]
    if on_h <= off_h:
        lights = on_h <= hour < off_h
    else:  # schedule wraps past midnight
        lights = hour >= on_h or hour < off_h

    return {"fan": fan, "heater": heater, "humidifier": humidifier, "lights": lights}


# ---------------------------------------------------------------------------
# Device endpoints (ESP32 <-> server)
# ---------------------------------------------------------------------------
@app.post("/api/ingest")
def ingest(payload: IngestPayload, x_api_key: str = Header(...)):
    with get_db() as db:
        device = get_device_by_api_key(db, x_api_key)
        if not device:
            raise HTTPException(status_code=401, detail="Invalid API key")
        if device["id"] != payload.device:
            raise HTTPException(status_code=403, detail="API key does not match device id")

        vpd = payload.vpd if payload.vpd is not None else calc_vpd(payload.temperature, payload.humidity)
        now = datetime.now(timezone.utc).isoformat()

        db.execute(
            "INSERT INTO readings (device_id, ts, temperature, humidity, vpd) VALUES (?, ?, ?, ?, ?)",
            (device["id"], now, payload.temperature, payload.humidity, vpd),
        )
        db.execute("UPDATE devices SET last_seen = ? WHERE id = ?", (now, device["id"]))

        ensure_settings_row(db, device["id"])
        settings_row = db.execute("SELECT * FROM settings WHERE device_id = ?", (device["id"],)).fetchone()

        control = compute_control(settings_row, payload.temperature, payload.humidity, vpd)
        return control


# ---------------------------------------------------------------------------
# Dashboard-facing API (protected by basic auth)
# ---------------------------------------------------------------------------
@app.post("/api/devices")
def create_device(payload: NewDevice, user: str = Depends(require_dashboard_auth)):
    device_id = payload.name.lower().replace(" ", "-")
    api_key = secrets.token_hex(16)
    with get_db() as db:
        try:
            db.execute(
                "INSERT INTO devices (id, name, api_key, created_at) VALUES (?, ?, ?, ?)",
                (device_id, payload.name, api_key, datetime.now(timezone.utc).isoformat()),
            )
            ensure_settings_row(db, device_id)
        except sqlite3.IntegrityError:
            raise HTTPException(status_code=400, detail="Device already exists")
    return {"device_id": device_id, "api_key": api_key}


@app.get("/api/devices")
def list_devices(user: str = Depends(require_dashboard_auth)):
    with get_db() as db:
        rows = db.execute("SELECT id, name, created_at, last_seen FROM devices").fetchall()
        return [dict(r) for r in rows]


@app.get("/api/readings/{device_id}")
def get_readings(
    device_id: str,
    hours: Optional[float] = None,
    limit: int = 200,
    user: str = Depends(require_dashboard_auth),
):
    """
    If `hours` is given, returns all readings from the last N hours (for
    time-scaled charts with pan/zoom). Otherwise falls back to the last
    `limit` rows.
    """
    with get_db() as db:
        if hours is not None:
            cutoff = (datetime.now(timezone.utc) - timedelta(hours=hours)).isoformat()
            rows = db.execute(
                "SELECT ts, temperature, humidity, vpd FROM readings WHERE device_id = ? AND ts >= ? ORDER BY ts ASC",
                (device_id, cutoff),
            ).fetchall()
            return [dict(r) for r in rows]

        rows = db.execute(
            "SELECT ts, temperature, humidity, vpd FROM readings WHERE device_id = ? ORDER BY ts DESC LIMIT ?",
            (device_id, limit),
        ).fetchall()
        return [dict(r) for r in reversed(rows)]


@app.get("/api/readings/{device_id}/latest")
def get_latest_reading(device_id: str, user: str = Depends(require_dashboard_auth)):
    with get_db() as db:
        row = db.execute(
            "SELECT ts, temperature, humidity, vpd FROM readings WHERE device_id = ? ORDER BY ts DESC LIMIT 1",
            (device_id,),
        ).fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="No readings yet")
        return dict(row)


@app.get("/api/settings/{device_id}")
def get_settings(device_id: str, user: str = Depends(require_dashboard_auth)):
    with get_db() as db:
        ensure_settings_row(db, device_id)
        row = db.execute("SELECT * FROM settings WHERE device_id = ?", (device_id,)).fetchone()
        return dict(row)


@app.post("/api/settings/{device_id}")
def update_settings(device_id: str, payload: SettingsUpdate, user: str = Depends(require_dashboard_auth)):
    with get_db() as db:
        ensure_settings_row(db, device_id)
        updates = {k: v for k, v in payload.dict().items() if v is not None}
        if not updates:
            return {"status": "no changes"}
        # booleans -> int for sqlite
        for k in list(updates.keys()):
            if isinstance(updates[k], bool):
                updates[k] = int(updates[k])
        set_clause = ", ".join(f"{k} = ?" for k in updates)
        db.execute(f"UPDATE settings SET {set_clause} WHERE device_id = ?", (*updates.values(), device_id))
        row = db.execute("SELECT * FROM settings WHERE device_id = ?", (device_id,)).fetchone()
        return dict(row)


@app.get("/api/status/{device_id}")
def current_status(device_id: str, user: str = Depends(require_dashboard_auth)):
    """Latest reading + what the control logic would currently output."""
    with get_db() as db:
        ensure_settings_row(db, device_id)
        settings_row = db.execute("SELECT * FROM settings WHERE device_id = ?", (device_id,)).fetchone()
        latest = db.execute(
            "SELECT ts, temperature, humidity, vpd FROM readings WHERE device_id = ? ORDER BY ts DESC LIMIT 1",
            (device_id,),
        ).fetchone()
        if not latest:
            return {"reading": None, "control": None}
        control = compute_control(settings_row, latest["temperature"], latest["humidity"], latest["vpd"])
        return {"reading": dict(latest), "control": control}


# ---------------------------------------------------------------------------
# Static dashboard
# ---------------------------------------------------------------------------
static_dir = os.path.join(os.path.dirname(__file__), "static")
app.mount("/", StaticFiles(directory=static_dir, html=True), name="static")
