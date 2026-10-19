import os
import json
import time
from datetime import datetime, timedelta, timezone

import gspread
from flask import Flask, Response, jsonify, request
from google.oauth2.service_account import Credentials

BASE = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__)

IST = timezone(timedelta(hours=5, minutes=30))
STAGES = ["IMAP", "ZIP", "CONVERT", "TRANSFORM", "SHEET", "DONE"]
COLS = ("run_id", "ts", "stage", "unit", "level", "message")
STALL_SECONDS = int(os.environ.get("STALL_SECONDS", "300"))   # इतनी देर कोई लॉग/हार्टबीट नहीं = STALLED
TAIL_ROWS = int(os.environ.get("TAIL_ROWS", "4000"))          # शीट की आख़िरी इतनी पंक्तियाँ ही पढ़ना
CACHE_TTL = 4                                                 # सेकंड (Google API कोटा बचाने के लिए)
_cache = {"t": 0.0, "rows": []}


def _worksheet():
    info = json.loads(os.environ["GOOGLE_CREDENTIALS"])
    creds = Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
    sh = gspread.authorize(creds).open_by_key(os.environ["LOG_SHEET_ID"])
    return sh.worksheet(os.environ.get("LOG_TAB", "logs"))


def load_rows():
    if time.time() - _cache["t"] < CACHE_TTL:
        return _cache["rows"]
    ws = _worksheet()
    total = len(ws.col_values(1))
    rows = []
    if total >= 2:
        start = max(2, total - TAIL_ROWS + 1)
        for r in ws.get(f"A{start}:F{total}"):
            if r:
                rows.append(dict(zip(COLS, list(r) + [""] * (6 - len(r)))))
    _cache.update(t=time.time(), rows=rows)
    return rows


def _ts(s):
    try:
        return datetime.fromisoformat(s)
    except Exception:
        return None


def summarize(rid, rs, now):
    times = [t for t in (_ts(r["ts"]) for r in rs) if t]
    start, last = (min(times), max(times)) if times else (None, None)
    errors = sum(r["level"] == "ERROR" for r in rs)
    warns = sum(r["level"] == "WARN" for r in rs)
    done = [r for r in rs if r["stage"] == "DONE" and r["message"].startswith(("COMPLETE", "FAILED"))]
    idle = int((now - last).total_seconds()) if last else 0
    if done and done[-1]["message"].startswith("FAILED"):
        status = "FAILED"
    elif done:
        status = "COMPLETED"
    elif idle > STALL_SECONDS:
        status = "STALLED"
    else:
        status = "RUNNING"
    end = last if status != "RUNNING" else now
    return {
        "id": rid, "status": status, "errors": errors, "warns": warns,
        "count": sum(r["level"] != "BEAT" for r in rs), "idle_s": idle,
        "start": start.isoformat() if start else None,
        "last": last.isoformat() if last else None,
        "duration_s": int((end - start).total_seconds()) if start and end else 0,
    }


def build_state(run_id):
    now = datetime.now(IST)
    rows = load_rows()
    runs = {}
    for r in rows:
        runs.setdefault(r["run_id"], []).append(r)
    summaries = sorted((summarize(k, v, now) for k, v in runs.items()),
                       key=lambda s: s["id"], reverse=True)
    sel = run_id if run_id in runs else (summaries[0]["id"] if summaries else None)
    out = {"now": now.isoformat(), "runs": summaries[:30], "run": None,
           "stages": [], "units": [], "logs": []}
    if not sel:
        return out

    rs = runs[sel]
    run = next(s for s in summaries if s["id"] == sel)
    out["run"] = run
    logs = [r for r in rs if r["level"] != "BEAT"]
    out["logs"] = logs[-1500:]

    seen = [r["stage"] for r in logs if r["stage"] in STAGES]
    current = seen[-1] if seen else None
    for s in STAGES:
        sr = [r for r in logs if r["stage"] == s]
        if not sr:
            st = "pending"
        elif any(r["level"] == "ERROR" for r in sr):
            st = "error"
        elif s == current and run["status"] in ("RUNNING", "STALLED"):
            st = "stalled" if run["status"] == "STALLED" else "active"
        else:
            st = "done"
        out["stages"].append({"name": s, "status": st, "logs": len(sr)})

    units = {}
    for r in logs:
        if not r["unit"]:
            continue
        u = units.setdefault(r["unit"], {"name": r["unit"], "stage": "", "errors": 0,
                                         "warns": 0, "done": False, "last": ""})
        u["stage"], u["last"] = r["stage"], r["message"]
        u["errors"] += r["level"] == "ERROR"
        u["warns"] += r["level"] == "WARN"
        if r["stage"] == "SHEET" and r["level"] == "SUCCESS" and "जोड़ीं" in r["message"]:
            u["done"] = True
    for u in units.values():
        u["status"] = "error" if u["errors"] else "done" if u["done"] else (
            "active" if run["status"] == "RUNNING" else "incomplete")
    out["units"] = sorted(units.values(), key=lambda u: u["name"])
    return out


@app.route("/")
def index():
    with open(os.path.join(BASE, "index.html"), encoding="utf-8") as f:
        return Response(f.read(), mimetype="text/html")


@app.route("/api/state")
def state():
    try:
        return jsonify(build_state(request.args.get("run_id", "")))
    except KeyError as e:
        return jsonify(error=f"Environment variable missing: {e}"), 500
    except Exception as e:
        return jsonify(error=f"{type(e).__name__}: {e}"), 500


if __name__ == "__main__":
    app.run(debug=True, port=5000)
