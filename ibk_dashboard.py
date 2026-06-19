from __future__ import annotations

import hashlib
import importlib.util
import json
import os
import re
import secrets
import shutil
import sys
import threading
import time
import traceback
from datetime import datetime, timedelta
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse
from urllib.request import Request, urlopen

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Side, Border


APP_DIR = Path(__file__).resolve().parent
DASHBOARD_DIR = Path(os.environ.get("IBK_GENERATOR_DIR", r"C:\Users\User1\Desktop\IBK_Dashboard_SQLite"))
sys.path.insert(0, str(DASHBOARD_DIR))

import im70_74_excel_com as core  # noqa: E402
from im70_74_analysis_report import build as build_png_report  # noqa: E402
from ibk_store import IBKStore  # noqa: E402


HOST = os.environ.get("IBK_HOST", "0.0.0.0")
PORT = int(os.environ.get("IBK_PORT", "8788"))
DATA_DIR = APP_DIR / "data"
UPLOAD_DIR = DATA_DIR / "uploads"
REPORT_DIR = DATA_DIR / "reports"
ASSET_DIR = APP_DIR / "assets"
USER_PATH = DATA_DIR / "users.json"
SESSION_PATH = DATA_DIR / "sessions.json"
INDEX_PATH = DATA_DIR / "archive.json"
DB_PATH = DATA_DIR / "ibk_dashboard.sqlite3"
TEMPLATE_PATH = Path(os.environ.get("IM70_74_TEMPLATE", DASHBOARD_DIR / "template_im70_74.xlsx"))
TOLOV_OUTPUT_DIR = Path(os.environ.get("IBK_TOLOV_OUTPUT_DIR", r"C:\Users\User1\Documents\Codex\tolov_generated_0406_0706"))
TOLOV_GENERATOR_PATH = Path(os.environ.get("IBK_TOLOV_GENERATOR", r"C:\Users\User1\Documents\Codex\2026-06-04\powershell-https-aka-ms-pswindows-ps\outputs\tolov_generator.py"))
TOLOV_FILES = {
    "1. Sbor11-12.xlsx": "Sbor 11-12: yig'imlar",
    "2. Sbor 44.xlsx": "Sbor 44: TR80 yig'imi",
    "3. Sbor im40.xlsx": "Sbor 10: IM40/ND40",
    "4. Sbor ek10.xlsx": "Sbor 10: EK10",
    "5. Sbor dr.xlsx": "Sbor 10: boshqa rejimlar",
    "6. 29.xlsx": "29-kod to'lovi",
    "7. 20.xlsx": "20-kod to'lovi",
    "8. 27.xlsx": "27-kod to'lovi",
    "9. 25.xlsx": "25-kod to'lovi",
    "10. 30.xlsx": "30-kod to'lovi",
    "11. 74.xlsx": "74-kod to'lovi",
    "11. 79.xlsx": "79-kod to'lovi",
    "12. im42.xlsx": "IM42 bo'yicha ro'yxat",
}

JOBS: dict[str, dict] = {}
SESSIONS: dict[str, dict] = {}
STORE: IBKStore | None = None

TAS_ICAO = "UTTT"  # Toshkent (Islom Karimov) xalqaro aeroporti
FLIGHTS_CACHE: dict = {"ts": 0.0, "data": None}
FLIGHTS_CACHE_TTL = 300  # 5 daqiqa


def _opensky_fetch(kind: str, begin: int, end: int) -> list[dict]:
    url = f"https://opensky-network.org/api/flights/{kind}?airport={TAS_ICAO}&begin={begin}&end={end}"
    req = Request(url, headers={"User-Agent": "Mozilla/5.0 (IBK-Dashboard)"})
    with urlopen(req, timeout=10) as resp:
        raw = json.loads(resp.read().decode("utf-8"))
    rows = []
    for f in raw[-15:]:
        rows.append({
            "callsign": (f.get("callsign") or "").strip(),
            "from": f.get("estDepartureAirport") or "",
            "to": f.get("estArrivalAirport") or "",
            "first_seen": f.get("firstSeen"),
            "last_seen": f.get("lastSeen"),
        })
    rows.reverse()
    return rows


def fetch_tas_flights() -> dict:
    now = int(time.time())
    begin = now - 12 * 3600
    result = {"updated": now, "arrivals": [], "departures": [], "error": ""}
    try:
        result["arrivals"] = _opensky_fetch("arrival", begin, now)
    except Exception as exc:
        result["error"] = f"{type(exc).__name__}: {exc}"
    try:
        result["departures"] = _opensky_fetch("departure", begin, now)
    except Exception as exc:
        if not result["error"]:
            result["error"] = f"{type(exc).__name__}: {exc}"
    return result


def get_tas_flights() -> dict:
    now = time.time()
    if FLIGHTS_CACHE["data"] is None or now - FLIGHTS_CACHE["ts"] > FLIGHTS_CACHE_TTL:
        FLIGHTS_CACHE["data"] = fetch_tas_flights()
        FLIGHTS_CACHE["ts"] = now
    return FLIGHTS_CACHE["data"]


# Toshkent atrofidagi hudud (Markaziy Osiyo) - jonli parvozlar xaritasi uchun
TAS_BBOX = {"lamin": 34.0, "lomin": 56.0, "lamax": 46.0, "lomax": 78.0}
LIVE_FLIGHTS_CACHE: dict = {"ts": 0.0, "data": None}
LIVE_FLIGHTS_CACHE_TTL = 25  # OpenSky anonim limitlariga ehtiyot


def fetch_live_states() -> dict:
    b = TAS_BBOX
    url = (
        "https://opensky-network.org/api/states/all"
        f"?lamin={b['lamin']}&lomin={b['lomin']}&lamax={b['lamax']}&lomax={b['lomax']}"
    )
    result = {"updated": int(time.time()), "planes": [], "error": ""}
    try:
        req = Request(url, headers={"User-Agent": "Mozilla/5.0 (IBK-Dashboard)"})
        with urlopen(req, timeout=10) as resp:
            raw = json.loads(resp.read().decode("utf-8"))
        for s in raw.get("states") or []:
            lon, lat = s[5], s[6]
            if lon is None or lat is None:
                continue
            result["planes"].append({
                "icao24": s[0],
                "callsign": (s[1] or "").strip(),
                "country": s[2],
                "lon": lon,
                "lat": lat,
                "alt": s[7] or s[13] or 0,
                "on_ground": bool(s[8]),
                "speed": s[9] or 0,
                "heading": s[10] or 0,
            })
    except Exception as exc:
        result["error"] = f"{type(exc).__name__}: {exc}"
    return result


def get_live_states() -> dict:
    now = time.time()
    if LIVE_FLIGHTS_CACHE["data"] is None or now - LIVE_FLIGHTS_CACHE["ts"] > LIVE_FLIGHTS_CACHE_TTL:
        LIVE_FLIGHTS_CACHE["data"] = fetch_live_states()
        LIVE_FLIGHTS_CACHE["ts"] = now
    return LIVE_FLIGHTS_CACHE["data"]


def ensure_job(job_id: str) -> dict:
    """JOBS va JOBS[job_id] hech qachon None bo'lib qolmasligi uchun himoya."""
    global JOBS
    if not isinstance(JOBS, dict):
        JOBS = {}
    if not job_id:
        job_id = "unknown_" + str(int(time.time() * 1000))
    job = JOBS.get(job_id)
    if not isinstance(job, dict):
        job = {"status": "navbatda", "data": {}}
    if not isinstance(job.get("data"), dict):
        job["data"] = {}
    JOBS[job_id] = job
    return job


def safe_files_dict(files) -> dict:
    """files har doim dict bo'lishini kafolatlaydi."""
    if isinstance(files, dict):
        return files
    return {"status": "tayyorlanmoqda", "excel": "", "pdf": "", "pngs": []}

DEFAULT_PERMS = ["view", "upload", "export", "release"]
ADMIN_PERMS = ["view", "upload", "export", "release", "admin", "settings"]
ROLES = {"admin", "rahbar", "inspektor", "foydalanuvchi", "user"}
ROLE_LABELS = {
    "admin": "Admin",
    "rahbar": "Rahbar",
    "inspektor": "Inspektor",
    "foydalanuvchi": "Foydalanuvchi",
    "user": "Foydalanuvchi",
}


def ensure_dirs():
    for path in [DATA_DIR, UPLOAD_DIR, REPORT_DIR, ASSET_DIR, TOLOV_OUTPUT_DIR]:
        path.mkdir(parents=True, exist_ok=True)
    if not USER_PATH.exists():
        save_json(USER_PATH, {
            "admin": {
                "salt": "ibk",
                "password": hash_password("admin123", "ibk"),
                "role": "admin",
                "enabled": True,
                "full_name": "Administrator",
                "position": "Admin",
                "phone": "",
                "lang": "uz",
                "post_code": "",
                "role_label": "Admin",
                "perms": ADMIN_PERMS,
            }
        })
    else:
        users = load_json(USER_PATH, {})
        changed = False
        for rec in users.values():
            if "enabled" not in rec:
                rec["enabled"] = True
                changed = True
            if "perms" not in rec:
                rec["perms"] = ADMIN_PERMS if rec.get("role") == "admin" else DEFAULT_PERMS
                changed = True
            for field, default in [("full_name", ""), ("position", ""), ("phone", ""), ("lang", "uz"), ("post_code", "")]:
                if field not in rec:
                    rec[field] = default
                    changed = True
            label = ROLE_LABELS.get(rec.get("role", "user"), "Foydalanuvchi")
            if rec.get("role_label") != label:
                rec["role_label"] = label
                changed = True
        if changed:
            save_json(USER_PATH, users)
    if not SESSION_PATH.exists():
        save_json(SESSION_PATH, {})
    if not INDEX_PATH.exists():
        save_json(INDEX_PATH, {"reports": []})


def load_json(path: Path, default):
    if not path.exists():
        return default
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return default
    return default if data is None else data


def save_json(path: Path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def hash_password(password: str, salt: str) -> str:
    return hashlib.sha256((salt + password).encode("utf-8")).hexdigest()


def safe_name(name: str) -> str:
    return re.sub(r"[^\w .,'`()+-]+", "_", name, flags=re.UNICODE).strip() or "file"


def parse_multipart(content_type: str, body: bytes) -> dict[str, dict]:
    match = re.search(r"boundary=(?P<b>[^;]+)", content_type or "")
    if not match:
        return {}
    boundary = match.group("b").strip().strip('"').encode("utf-8")
    result: dict[str, dict] = {}
    marker = b"--" + boundary
    for part in body.split(marker):
        part = part.strip(b"\r\n")
        if not part or part == b"--":
            continue
        if part.endswith(b"--"):
            part = part[:-2].rstrip(b"\r\n")
        head, sep, content = part.partition(b"\r\n\r\n")
        if not sep:
            continue
        disposition = ""
        for line in head.decode("utf-8", errors="replace").split("\r\n"):
            if line.lower().startswith("content-disposition:"):
                disposition = line
                break
        name_match = re.search(r'name="([^"]+)"', disposition)
        if not name_match:
            continue
        filename_match = re.search(r'filename="([^"]*)"', disposition)
        field_name = name_match.group(1)
        item = {
            "filename": filename_match.group(1) if filename_match else "",
            "content": content.rstrip(b"\r\n"),
        }
        if field_name in result:
            if isinstance(result[field_name], list):
                result[field_name].append(item)
            else:
                result[field_name] = [result[field_name], item]
        else:
            result[field_name] = item
    return result


def report_date_from_name(name: str) -> datetime:
    return core.report_date_from_name(name) or datetime.now()


def fmt_date(dt: datetime) -> str:
    return dt.strftime("%d.%m.%Y")


def duplicate_report(filename: str, report_date: datetime) -> dict | None:
    needle = safe_name(filename)
    for item in load_json(INDEX_PATH, {"reports": []}).get("reports", []):
        if item.get("date") != fmt_date(report_date):
            continue
        source_name = Path(item.get("source", "")).name
        if source_name == needle or source_name == filename:
            return item
    return None


def clean_archive_records(reports: list[dict]) -> list[dict]:
    """Arxivda bir sana bitta qator bo'lib qolsin; eng foydali yozuv saqlanadi."""
    def score(rec: dict):
        source = str(rec.get("source", ""))
        deposit = str(rec.get("deposit", ""))
        return (
            1 if "IBK_Dashboard_SQLite" in source else 0,
            1 if deposit else 0,
            str(rec.get("id", "")),
        )
    best: dict[str, dict] = {}
    for rec in reports or []:
        if not isinstance(rec, dict):
            continue
        date = rec.get("date", "")
        if not date:
            continue
        if date not in best or score(rec) > score(best[date]):
            best[date] = rec
    def date_key(rec: dict):
        try:
            return datetime.strptime(rec.get("date", ""), "%d.%m.%Y")
        except Exception:
            return datetime.min
    return sorted(best.values(), key=date_key, reverse=True)


def digits(value) -> str:
    if isinstance(value, float):
        if value != value:  # NaN
            return ""
        return str(int(value))
    if isinstance(value, int):
        return str(value)
    return re.sub(r"\D", "", str(value or ""))


def source_post_code(value) -> str:
    match = re.search(r"\d{5}", str(value or ""))
    return match.group(0) if match else ""


POST_CLASSIFIER = {
    # Qoraqalpog'iston Respublikasi
    "35001": "Nukus aeroporti chegara bojxona posti",
    "35002": "Nukus TIF bojxona posti",
    "35003": "Xo'jayli chegara bojxona posti",
    "35004": "Dovut-ota chegara bojxona posti",
    "35010": "Qoraqalpog'iston temir yo'l chegara bojxona posti",
    # Andijon viloyati
    "03002": "Do'stlik chegara bojxona posti",
    "03003": "Andijon aeroporti chegara bojxona posti",
    "03005": "Mingtepa chegara bojxona posti",
    "03006": "Qorasuv chegara bojxona posti",
    "03007": "Xonobod chegara bojxona posti",
    "03008": "Pushmon chegara bojxona posti",
    "03009": "Madaniyat chegara bojxona posti",
    "03011": "Andijon TIF bojxona posti",
    "03013": "Keskanyoʻr chegara bojxona posti",
    "03014": "Savay temir yo'l chegara bojxona posti",
    "03015": "Asaka TIF bojxona posti",
    # Buxoro viloyati
    "06001": "Buxoro aeroporti chegara bojxona posti",
    "06006": "Buxoro TIF bojxona posti",
    "06009": "Qorakoʻl TIF bojxona posti",
    "06010": "Olot chegara bojxona posti",
    "06011": "Xo'jadavlat temir yo'l chegara bojxona posti",
    # Jizzax viloyati
    "08003": "Uchturgon chegara bojxona posti",
    "08004": "Jizzax TIF bojxona posti",
    "08007": "Qo'shkent chegara bojxona posti",
    # Qashqadaryo viloyati
    "10002": "Nasaf TIF bojxona posti",
    "10007": "Qamashi-G'uzor TIF bojxona posti",
    "10008": "Qarshi-Kerki chegara bojxona posti",
    "10012": "Qarshi aeroporti chegara bojxona posti",
    # Navoiy viloyati
    "12002": "Navoiy aeroporti chegara bojxona posti",
    "12003": "Navoiy TIF bojxona posti",
    "12008": "Zarafshon TIF bojxona posti",
    # Namangan viloyati
    "14002": "Namangan aeroporti chegara bojxona posti",
    "14003": "Uchqo'rg'on chegara bojxona posti",
    "14004": "Kosonsoy chegara bojxona posti",
    "14005": "Pop chegara bojxona posti",
    "14010": "Namangan TIF bojxona posti",
    # Samarqand viloyati
    "18001": "Samarqand aeroporti chegara bojxona posti",
    "18002": "Jartepa chegara bojxona posti",
    "18005": "Samarqand TIF bojxona posti",
    "18007": "Ulug'bek TIF bojxona posti",
    # Surxondaryo viloyati
    "22002": "Termiz aeroporti chegara bojxona posti",
    "22003": "Sariosiyo chegara bojxona posti",
    "22004": "Sariosiyo temir yo'l chegara bojxona posti",
    "22005": "Termiz TIF bojxona posti",
    "22006": "Denov TIF bojxona posti",
    "22007": "Gulbahor chegara bojxona posti",
    "22011": "Daryo porti chegara bojxona posti",
    "22015": "Boldir temir yo'l chegara bojxona posti",
    "22017": "Ayritom chegara bojxona posti",
    "22022": "Termiz xalqaro savdo markazi TIF bojxona posti",
    # Sirdaryo viloyati
    "24002": "Xovosоbod chegara bojxona posti",
    "24004": "Sirdaryo chegara bojxona posti",
    "24006": "Oq oltin chegara bojxona posti",
    "24009": "Guliston TIF bojxona posti",
    "24014": "Malik chegara bojxona posti",
    # Toshkent viloyati
    "27001": "Yallama chegara bojxona posti",
    "27008": "Navoiy chegara bojxona posti",
    "27009": "S. Najimov chegara bojxona posti",
    "27011": "Oybek chegara bojxona posti",
    "27013": "Bekobod avto chegara bojxona posti",
    "27014": "Chirchiq TIF bojxona posti",
    "27015": "Olmaliq TIF bojxona posti",
    "27016": "Yangiyoʻl TIF bojxona posti",
    "27019": "Nazarbek TIF bojxona posti",
    "27020": "Keles TIF bojxona posti",
    "27021": "G'ishtkoprik chegara bojxona posti",
    "27023": "Farhod chegara bojxona posti",
    "27024": "Bekobod temir yo'l chegara bojxona posti",
    "27028": "Angren TIF bojxona posti",
    # Farg'ona viloyati
    "30001": "Farg'ona aeroporti chegara bojxona posti",
    "30002": "Qo'qon TIF bojxona posti",
    "30004": "Farg'ona chegara bojxona posti",
    "30005": "Andarxon chegara bojxona posti",
    "30006": "Rishton chegara bojxona posti",
    "30008": "Rovot chegara bojxona posti",
    "30009": "Vodiy TIF bojxona posti",
    "30010": "O'zbekiston chegara bojxona posti",
    "30012": "So'x chegara bojxona posti",
    # Xorazm viloyati
    "33001": "Shovot chegara bojxona posti",
    "33004": "Do'stlik chegara bojxona posti",
    "33007": "Urganch TIF bojxona posti",
    "33011": "Urganch aeroporti chegara bojxona posti",
    "33033": "Shovot chegaraoldi savdo zonasi chegara bojxona posti",
    # Toshkent shahri
    "26002": "Toshkent-tovar TIF bojxona posti",
    "26003": "Ark buloq TIF bojxona posti",
    "26004": "Chuqursoy TIF bojxona posti",
    "26009": "Keles temir yo'l chegara bojxona posti",
    "26010": "Sirg'ali TIF bojxona posti",
    "26013": "Chuqursoy texnik idora temir yo'l chegara bojxona posti",
    # Toshkent-AERO
    "00101": "Toshkent xalqaro aeroporti CHBP",
    "00102": "Avia yuklar TIF bojxona posti",
    "00107": "Elektron tijorat TIF bojxona posti",
    "00110": "Toshkent-Humo aeroporti CHBP",
}


def source_post_name(code: str) -> str:
    code = core.clean(code)
    return POST_CLASSIFIER.get(code, f"Post №{code}" if code else "-")


def source_transport(code: str, post_name: str = "") -> str:
    text = f"{code} {post_name}".lower()
    if code.startswith("001") or any(w in text for w in ["avia", "aero", "aeroport", "airport"]):
        return "Avia"
    if any(w in text for w in ["temir", "rail", "poezd", "vokzal", "yol", "yo'l"]):
        return "Temir yo'l"
    return "Avto"


def declaration_key(row) -> str:
    return core.clean(row[core.SRC["decl_no"]])


def read_report_source(source: Path) -> pd.DataFrame:
    df = core.read_source(source)
    df["_decl_key"] = df.apply(declaration_key, axis=1)
    df["_source_post_code"] = df[core.SRC["decl_no"]].map(source_post_code)
    df["_source_post_name"] = df["_source_post_code"].map(source_post_name)
    df["_source_transport"] = df.apply(lambda r: source_transport(r.get("_source_post_code", ""), r.get("_source_post_name", "")), axis=1)
    return df


def item_key(row) -> str:
    goods_no = core.clean(row[core.SRC.get("goods_no", core.SRC["decl_no"])])
    return f"{declaration_key(row)}::{goods_no}"


def snapshot_rows(df: pd.DataFrame) -> list[dict]:
    rows = []
    for _, r in df.iterrows():
        dt = r.get("_date", None)
        if pd.notna(dt):
            gtd_date = pd.Timestamp(dt).strftime("%Y-%m-%d")
        else:
            gtd_date = core.clean(r[core.SRC["decl_date"]])
        rows.append({
            "item_key": item_key(r),
            "decl": core.clean(r[core.SRC["decl_no"]]),
            "item_no": core.clean(r[core.SRC["goods_no"]]),
            "regime": core.clean(r.get("_regime_code", r[core.SRC["regime"]])),
            "post": core.to_latin(core.clean(r.get("_post_group", r[core.SRC["post"]]))),
            "source_post": core.clean(r.get("_source_post_code", "")),
            "source_post_name": core.clean(r.get("_source_post_name", "")),
            "transport": core.clean(r.get("_source_transport", "")),
            "stir": digits(r[core.SRC["stir"]]),
            "company": core.clean(r[core.SRC["company"]]),
            "gtd_date": gtd_date,
            "hs_code": core.clean(r[core.SRC["hs"]]),
            "goods": core.clean(r[core.SRC["goods"]]),
            "weight": float(r.get("_weight_tn", 0) or 0),
            "value": float(r.get("_value_usd_k", 0) or 0),
            "payment": float(r.get("_pay_total_mln", 0) or 0),
            "partiya": int(r.get("_partiya", 1) or 0),
            "country": core.clean(r[core.SRC["country"]]),
            "warehouse": core.clean(r[core.SRC["warehouse"]]),
            "reason": core.clean_reason_display(r[core.SRC["reason"]]),
        })
    return rows


def save_snapshot_to_store(report_id: str, source: Path, deposit: Path | None, report_date: datetime, df: pd.DataFrame | None = None):
    global STORE
    if STORE is None:
        STORE = IBKStore(DB_PATH)
    df = df if df is not None else read_report_source(source)
    snapshot_id = STORE.register_snapshot(report_id, source.name, str(source), str(deposit or ""), fmt_date(report_date))
    STORE.insert_active_items(snapshot_id, snapshot_rows(df))
    return snapshot_id


STORE_BACKFILLED = False


def ensure_store_backfilled():
    """Arxivdagi (2024 yildan boshlab) avval yuklangan, lekin SQLite snapshot
    bazasiga hali tushmagan hisobotlarni bir martalik tarzda bazaga qo'shadi.
    Shu orqali "Aylanma tendensiyasi" kabi tahlillar 2024 yildan boshlab
    mavjud bo'lgan barcha davrlarni qamrab oladi."""
    global STORE, STORE_BACKFILLED
    if STORE_BACKFILLED:
        return
    if STORE is None:
        STORE = IBKStore(DB_PATH)
    try:
        existing = {s.get("report_id") for s in STORE.list_snapshots()}
    except Exception:
        existing = set()
    archive = load_json(INDEX_PATH, {"reports": []})
    reports = archive.get("reports", []) if isinstance(archive, dict) else []
    for rec in reports:
        if not isinstance(rec, dict):
            continue
        rid = rec.get("id")
        if not rid or rid in existing:
            continue
        source = rec.get("source")
        if not source:
            continue
        src_path = Path(source)
        if not src_path.exists():
            continue
        try:
            report_date = report_date_from_name(src_path.name)
            df = read_report_source(src_path)
            deposit_raw = rec.get("deposit") or ""
            deposit = Path(deposit_raw) if deposit_raw else None
            save_snapshot_to_store(rid, src_path, deposit, report_date, df)
            existing.add(rid)
        except Exception:
            continue
    STORE_BACKFILLED = True


def is_own_warehouse_name(value: str) -> bool:
    text = core.clean(value)
    low = text.lower().replace("?", "'").replace("`", "'")
    if not text:
        return True
    address_words = [
        "toshkent shahar", "toshkent shahri", "tashkent", "toshkent viloyati",
        " r-n", "r-n ", "tumani", "tuman", "ko'chasi", "kochasi", "ko`chasi",
        "mfy", "mahalla", "-uy", " uy", "dom", "??????", "??????", "??????",
    ]
    return any(w in low for w in address_words)


def own_warehouse_label(value: str) -> str:
    return "O'z ombor" if is_own_warehouse_name(value) else core.clean(value)


def own_company_rows_web(df: pd.DataFrame, report_date: datetime, over_3_months: bool = False) -> list[dict]:
    own = df[df[core.SRC["warehouse"]].map(is_own_warehouse_name)].copy()
    if over_3_months:
        cutoff = pd.Timestamp(report_date) - pd.DateOffset(months=3)
        own = own[own["_date"].le(cutoff)].copy()
    if own.empty:
        return []
    g = own.groupby(core.SRC["stir"], dropna=False).agg(
        company=(core.SRC["company"], "first"),
        partiya=("_partiya", "sum"), vazn=("_weight_tn", "sum"), qiymat=("_value_usd_k", "sum"), tolov=("_pay_total_mln", "sum"),
        first_date=("_date", "min"), goods=(core.SRC["goods"], core.goods_category),
    ).reset_index().sort_values(["qiymat", "tolov"], ascending=False)
    rows = []
    for r in g.itertuples(index=False):
        kun_hisobi = int((pd.Timestamp(report_date) - pd.Timestamp(r.first_date)).days) if pd.notna(r.first_date) else 0
        rows.append({"key": {"stir": digits(r[0])}, "korxona": core.clean(r.company), "stir": digits(r[0]), "partiya": int(r.partiya), "vazn": float(r.vazn), "qiymat": float(r.qiymat), "tolov": float(r.tolov), "muddat": r.first_date.strftime("%d.%m.%Y") if pd.notna(r.first_date) else "", "kun_hisobi": kun_hisobi, "tovar": core.to_latin(core.clean(r.goods))})
    return rows

def detail_rows(df: pd.DataFrame, filters: dict) -> list[dict]:
    d = df.copy()
    if filters.get("decl"):
        d = d[d[core.SRC["decl_no"]].map(core.clean).eq(filters["decl"])]
    if filters.get("stir"):
        d = d[d[core.SRC["stir"]].map(digits).eq(filters["stir"])]
    if filters.get("tnved"):
        d = d[d["_tnved_name"].map(core.to_latin).eq(filters["tnved"])]
    if filters.get("regime"):
        d = d[d["_regime_code"].eq(filters["regime"])]
    if filters.get("post"):
        d = d[d["_post_group"].map(core.to_latin).eq(filters["post"])]
    if filters.get("source_post"):
        d = d[d["_source_post_code"].map(core.clean).eq(filters["source_post"])]
    if filters.get("transport"):
        d = d[d["_source_transport"].map(core.clean).eq(filters["transport"])]
    if filters.get("food_name"):
        _food_cats = [
            ("Ichimliklar (alkogolsiz)", r"ичимлик|напит|drink|сок|water|вода"),
            ("Shakar va qandolat mahsulotlari", r"шакар|сахар|sugar|qand|конфет|шоколад"),
            ("Sut mahsulotlari, tuxum va asal", r"сут|молок|milk|tuxum|яйц|egg|asal|мед"),
            ("Go'sht va go'sht mahsulotlari", r"go'sht|мяс|meat|колбас|tovuq|куриц"),
            ("Yog' va moy mahsulotlari", r"yog'|yog\b|мой|масло|\boil\b"),
            ("Don, un va yorma mahsulotlari", r"\bdon\b|\bun\b|мука|круп|guruch|рис|bug'doy|пшениц|jo'xori|arpa"),
            ("Meva-sabzavot mahsulotlari", r"мева|сабзавот|овощ|фрукт|картоф|томат|пиёз|лук|pomidor|bodring"),
            ("Boshqa oziq-ovqatlar", r"oziq|овқат|озиқ|confection|кондитер|озуқ|пищ"),
        ]
        fname = filters["food_name"]
        hay = (d["_tnved_name"].map(core.clean) + " " + d[core.SRC["goods"]].map(core.clean))
        used = pd.Series(False, index=d.index)
        matched = False
        for label, pat in _food_cats:
            mask = hay.str.contains(pat, case=False, na=False, regex=True)
            if core.to_latin(core.clean(label)) == fname:
                d = d[mask & ~used]
                matched = True
                break
            used |= mask
        if not matched:
            d = d[~used & d["_tnved_name"].map(core.clean).str.contains(
                "озиқ|овқат|пищ|food|озуқ", case=False, na=False, regex=True)]
    out = []
    for _, r in d.head(2000).iterrows():
        out.append({
            "decl": core.clean(r[core.SRC["decl_no"]]),
            "date": r["_date_text"],
            "regime": r["_regime_code"],
            "post": core.to_latin(core.clean(r["_post_group"])),
            "source_post": core.clean(r.get("_source_post_code", "")),
            "source_post_name": core.clean(r.get("_source_post_name", "")),
            "transport": core.clean(r.get("_source_transport", "")),
            "stir": digits(r[core.SRC["stir"]]),
            "company": core.clean(r[core.SRC["company"]]),
            "hs": core.clean(r[core.SRC["hs"]]),
            "goods": core.clean(r[core.SRC["goods"]]),
            "partiya": int(r["_partiya"]),
            "vazn": float(r["_weight_tn"]),
            "qiymat": float(r["_value_usd_k"]),
            "tolov": float(r["_pay_total_mln"]),
            "reason": core.clean_reason_display(r[core.SRC["reason"]]),
        })
    return out


def released_between(old_source: Path, new_source: Path) -> dict:
    old = read_report_source(old_source)
    new = read_report_source(new_source)
    old_g = old.groupby("_decl_key", dropna=False).agg(
        company=(core.SRC["company"], "first"),
        stir=(core.SRC["stir"], "first"),
        regime=("_regime_code", "first"),
        post=("_post_group", "first"),
        partiya=("_partiya", "sum"),
        vazn=("_weight_tn", "sum"),
        qiymat=("_value_usd_k", "sum"),
        tolov=("_pay_total_mln", "sum"),
    )
    new_g = new.groupby("_decl_key", dropna=False).agg(
        partiya=("_partiya", "sum"),
        vazn=("_weight_tn", "sum"),
        qiymat=("_value_usd_k", "sum"),
        tolov=("_pay_total_mln", "sum"),
    )
    rows = []
    for key, r in old_g.iterrows():
        n = new_g.loc[key] if key in new_g.index else None
        if n is None:
            dp, dv, dq, dt = int(r.partiya), float(r.vazn), float(r.qiymat), float(r.tolov)
        else:
            dp = max(0, int(r.partiya) - int(n.partiya))
            dv = max(0.0, float(r.vazn) - float(n.vazn))
            dq = max(0.0, float(r.qiymat) - float(n.qiymat))
            dt = max(0.0, float(r.tolov) - float(n.tolov))
        if dp > 0 or dv > 0.005 or dq > 0.005 or dt > 0.005:
            rows.append({
                "key": {"decl": key},
                "decl": key,
                "company": core.clean(r.company),
                "stir": digits(r.stir),
                "regime": r.regime,
                "post": core.to_latin(core.clean(r.post)),
                "partiya": dp,
                "vazn": dv,
                "qiymat": dq,
                "tolov": dt,
            })
    rows.sort(key=lambda x: (x["qiymat"], x["tolov"], x["partiya"]), reverse=True)
    return {
        "partiya": sum(x["partiya"] for x in rows),
        "vazn": sum(x["vazn"] for x in rows),
        "qiymat": sum(x["qiymat"] for x in rows),
        "tolov": sum(x["tolov"] for x in rows),
        "rows": rows[:300],
    }


def release_company_table(base_source: Path, final_source: Path) -> dict:
    base = read_report_source(base_source)
    final = read_report_source(final_source)

    def by_company(df: pd.DataFrame):
        g = df.groupby(core.SRC["stir"], dropna=False).agg(
            company=(core.SRC["company"], "first"),
            vazn=("_weight_tn", "sum"),
            qiymat=("_value_usd_k", "sum"),
            partiya=("_partiya", "sum"),
            tolov=("_pay_total_mln", "sum"),
        ).reset_index()
        g["stir"] = g[core.SRC["stir"]].map(digits)
        return g.set_index("stir")

    base_g = by_company(base)
    final_g = by_company(final)
    rows = []
    unreleased = []
    for stir, b in base_g.iterrows():
        f = final_g.loc[stir] if stir in final_g.index else None
        remain_vazn = float(f.vazn) if f is not None else 0.0
        remain_qiymat = float(f.qiymat) if f is not None else 0.0
        remain_partiya = int(f.partiya) if f is not None else 0
        remain_tolov = float(f.tolov) if f is not None else 0.0
        released_vazn = max(0.0, float(b.vazn) - remain_vazn)
        released_qiymat = max(0.0, float(b.qiymat) - remain_qiymat)
        released_partiya = max(0, int(b.partiya) - remain_partiya)
        released_tolov = max(0.0, float(b.tolov) - remain_tolov)
        row = {
            "key": {"stir": stir},
            "korxona": core.clean(b.company),
            "stir": stir,
            "base_vazn": float(b.vazn),
            "base_qiymat": float(b.qiymat),
            "base_partiya": int(b.partiya),
            "remain_vazn": remain_vazn,
            "remain_qiymat": remain_qiymat,
            "remain_partiya": remain_partiya,
            "released_vazn": released_vazn,
            "released_qiymat": released_qiymat,
            "released_pct": (released_qiymat / float(b.qiymat) * 100) if float(b.qiymat) else 0.0,
            "released_partiya": released_partiya,
            "released_tolov": released_tolov,
            "current_vazn": remain_vazn,
            "current_qiymat": remain_qiymat,
            "current_partiya": remain_partiya,
        }
        if released_vazn <= 0.005 and released_qiymat <= 0.005 and released_partiya <= 0 and released_tolov <= 0.005:
            unreleased.append(row)
            continue
        rows.append(row)
    rows.sort(key=lambda x: (x["released_qiymat"], x["released_vazn"], x["released_partiya"]), reverse=True)
    total = {
        "korxona": "Jami",
        "stir": "",
        "base_vazn": sum(r["base_vazn"] for r in rows),
        "base_qiymat": sum(r["base_qiymat"] for r in rows),
        "base_partiya": sum(r["base_partiya"] for r in rows),
        "remain_vazn": sum(r["remain_vazn"] for r in rows),
        "remain_qiymat": sum(r["remain_qiymat"] for r in rows),
        "remain_partiya": sum(r["remain_partiya"] for r in rows),
        "released_vazn": sum(r["released_vazn"] for r in rows),
        "released_qiymat": sum(r["released_qiymat"] for r in rows),
        "released_pct": 0.0,
        "released_partiya": sum(r["released_partiya"] for r in rows),
        "released_tolov": sum(r["released_tolov"] for r in rows),
        "current_vazn": sum(r["current_vazn"] for r in rows),
        "current_qiymat": sum(r["current_qiymat"] for r in rows),
        "current_partiya": sum(r["current_partiya"] for r in rows),
    }
    total["released_pct"] = (total["released_qiymat"] / total["base_qiymat"] * 100) if total["base_qiymat"] else 0.0
    partial = [r for r in rows if r["released_partiya"] == 0 and (r["released_qiymat"] > 0.005 or r["released_vazn"] > 0.005)]
    unreleased.sort(key=lambda x: (x["base_qiymat"], x["base_vazn"], x["base_partiya"]), reverse=True)
    return {"total": total, "rows": rows, "partial": partial[:100], "unreleased": unreleased[:100], "top_released": rows[:20]}


def find_report_by_date(date_text: str):
    for r in load_json(INDEX_PATH, {"reports": []})["reports"]:
        if r.get("date") == date_text:
            return r
    return None


def company_headers():
    return [
        {"k": "korxona", "t": "Korxona", "width": 42},
        {"k": "stir", "t": "STIR", "width": 14},
        {"k": "partiya", "t": "Partiya", "i": True, "width": 10},
        {"k": "vazn", "t": "Vazn (tn)", "n": True, "width": 14},
        {"k": "qiymat", "t": "Qiymat (ming $)", "n": True, "width": 16},
        {"k": "tolov", "t": "To'lov (mln so'm)", "n": True, "width": 18},
        {"k": "depozit", "t": "Depozit (mln so'm)", "n": True, "width": 18},
    ]


def expired_headers():
    return [
        {"k": "korxona", "t": "Korxona", "width": 42},
        {"k": "stir", "t": "STIR", "width": 14},
        {"k": "rejim", "t": "Rejim", "width": 10},
        {"k": "post", "t": "Post", "width": 22},
        {"k": "partiya", "t": "Partiya", "i": True, "width": 10},
        {"k": "kun", "t": "Kun hisobi", "i": True, "width": 12},
        {"k": "qiymat", "t": "Qiymat (ming $)", "n": True, "width": 16},
        {"k": "tolov", "t": "To'lov (mln so'm)", "n": True, "width": 18},
    ]


def goods_headers():
    return [
        {"k": "name", "t": "Tovar", "width": 46},
        {"k": "partiya", "t": "Partiya", "i": True, "width": 10},
        {"k": "qiymat", "t": "Qiymat (ming $)", "n": True, "width": 16},
        {"k": "tolov", "t": "To'lov (mln so'm)", "n": True, "width": 18},
    ]


def food_headers():
    return [
        {"k": "name", "t": "Oziq-ovqat turi", "width": 42},
        {"k": "vazn", "t": "Vazn (tn)", "n": True, "width": 14},
        {"k": "qiymat", "t": "Qiymat (ming $)", "n": True, "width": 16},
        {"k": "over_vazn", "t": "3 oy+ vazn (tn)", "n": True, "width": 16},
        {"k": "over_qiymat", "t": "3 oy+ qiymat (ming $)", "n": True, "width": 18},
        {"k": "ulush", "t": "Qiymatdagi ulushi (%)", "n": True, "width": 18},
    ]


def basic_headers(name_title: str = "Ko'rsatkich"):
    return [
        {"k": "name", "t": name_title, "width": 38},
        {"k": "partiya", "t": "Partiya", "i": True, "width": 10},
        {"k": "vazn", "t": "Vazn (tn)", "n": True, "width": 14},
        {"k": "qiymat", "t": "Qiymat (ming $)", "n": True, "width": 16},
        {"k": "tolov", "t": "To'lov (mln so'm)", "n": True, "width": 18},
    ]


def release_headers():
    return [
        {"k": "korxona", "t": "Korxona", "width": 42},
        {"k": "stir", "t": "STIR", "width": 14},
        {"k": "base_vazn", "t": "Boshlang'ich vazn (tn)", "n": True, "width": 18},
        {"k": "base_qiymat", "t": "Boshlang'ich qiymat (ming $)", "n": True, "width": 22},
        {"k": "base_partiya", "t": "Boshlang'ich partiya", "i": True, "width": 18},
        {"k": "remain_vazn", "t": "Yakuniy sanada boshlang'ichdan qoldiq vazn (tn)", "n": True, "width": 24},
        {"k": "remain_qiymat", "t": "Yakuniy sanada boshlang'ichdan qoldiq qiymat (ming $)", "n": True, "width": 28},
        {"k": "remain_partiya", "t": "Yakuniy sanada boshlang'ichdan qoldiq partiya", "i": True, "width": 24},
        {"k": "released_vazn", "t": "Yechilishi vazn (tn)", "n": True, "width": 18},
        {"k": "released_qiymat", "t": "Yechilishi qiymat (ming $)", "n": True, "width": 20},
        {"k": "released_pct", "t": "Yechilishi (%)", "n": True, "width": 14},
        {"k": "released_partiya", "t": "Yechilishi partiya", "i": True, "width": 16},
        {"k": "released_tolov", "t": "Yechilgan to'lov (mln so'm)", "n": True, "width": 20},
        {"k": "current_vazn", "t": "Yakuniy qoldiq vazn (tn)", "n": True, "width": 18},
        {"k": "current_qiymat", "t": "Yakuniy qoldiq qiymat (ming $)", "n": True, "width": 22},
        {"k": "current_partiya", "t": "Yakuniy qoldiq partiya", "i": True, "width": 18},
    ]


def make_xlsx(headers: list[dict], rows: list[dict], title: str) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Jadval"
    ws.cell(1, 1, title)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(1, len(headers)))
    ws.cell(1, 1).font = Font(bold=True, size=14)
    ws.cell(1, 1).alignment = Alignment(horizontal="center")
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    for c, h in enumerate(headers, 1):
        cell = ws.cell(2, c, h["t"])
        cell.fill = fill
        cell.font = font
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r_idx, row in enumerate(rows, 3):
        for c_idx, h in enumerate(headers, 1):
            value = row.get(h["k"], "")
            cell = ws.cell(r_idx, c_idx, value)
            cell.border = border
            cell.alignment = Alignment(horizontal="right" if h.get("n") else "left", vertical="center")
            if h.get("i"):
                cell.number_format = "# ##0"
            elif h.get("n"):
                cell.number_format = "# ##0.00"
    for c, h in enumerate(headers, 1):
        ws.column_dimensions[ws.cell(2, c).column_letter].width = h.get("width", 14)
    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()


def build_dashboard(report_id: str, source: Path, deposit: Path | None, report_date: datetime) -> dict:
    df = read_report_source(source)
    deposits, deposit_date, deposit_total = core.read_deposit(deposit)
    expired = core.with_expiry(df, report_date)
    expired = expired[expired["_expired"]].copy()

    companies = df.groupby(core.SRC["stir"], dropna=False).agg(
        company=(core.SRC["company"], "first"),
        partiya=("_partiya", "sum"),
        vazn=("_weight_tn", "sum"),
        qiymat=("_value_usd_k", "sum"),
        tolov=("_pay_total_mln", "sum"),
    ).reset_index()
    companies["stir"] = companies[core.SRC["stir"]].map(digits)
    companies["depozit"] = companies["stir"].map(lambda s: float(deposits.get(s, 0.0)))

    def company_records(frame):
        rows = []
        for _, r in frame.iterrows():
            rows.append({
                "key": {"stir": r["stir"]},
                "korxona": core.clean(r.company),
                "stir": r["stir"],
                "partiya": int(r.partiya),
                "vazn": float(r.vazn),
                "qiymat": float(r.qiymat),
                "tolov": float(r.tolov),
                "depozit": float(r.get("depozit", 0.0)),
            })
        return rows

    goods_rows_web = []
    for r in core.goods_rows(df):
        goods_rows_web.append({"key": {"tnved": core.to_latin(core.clean(r[1]))}, "name": core.to_latin(core.clean(r[1])), "partiya": int(r[2]), "korxona": int(r[3]), "vazn": float(r[4]), "qiymat": float(r[5]), "tolov": float(r[6])})

    regimes = df.groupby("_regime_code", dropna=False).agg(
        partiya=("_partiya", "sum"),
        vazn=("_weight_tn", "sum"),
        qiymat=("_value_usd_k", "sum"),
        tolov=("_pay_total_mln", "sum"),
    ).sort_values("qiymat", ascending=False).reset_index()

    exp = expired.groupby([core.SRC["stir"], "_regime_code", "_post_group"], dropna=False).agg(
        company=(core.SRC["company"], "first"),
        partiya=("_partiya", "sum"),
        qiymat=("_value_usd_k", "sum"),
        tolov=("_pay_total_mln", "sum"),
        expiry=("_expiry", "min"),
    ).sort_values(["partiya", "qiymat"], ascending=False).reset_index()
    exp_rows = []
    for _, r in exp.iterrows():
        exp_rows.append({
            "key": {"stir": digits(r[core.SRC["stir"]]), "regime": r["_regime_code"], "post": core.to_latin(core.clean(r["_post_group"]))},
            "korxona": core.clean(r.company),
            "stir": digits(r[core.SRC["stir"]]),
            "rejim": r["_regime_code"],
            "post": core.to_latin(core.clean(r["_post_group"])),
            "partiya": int(r.partiya),
            "kun": max(0, int((report_date - r.expiry).days)) if pd.notna(r.expiry) else 0,
            "qiymat": float(r.qiymat),
            "tolov": float(r.tolov),
        })

    age_rows = []
    for label, cutoff in [("3 oy+", report_date - pd.DateOffset(months=3)), ("6 oy+", report_date - pd.DateOffset(months=6)), ("1 yil+", report_date - pd.DateOffset(years=1))]:
        d = df[df["_date"].le(cutoff)]
        age_rows.append({"muddat": label, "partiya": int(d["_partiya"].sum()), "vazn": float(d["_weight_tn"].sum()), "qiymat": float(d["_value_usd_k"].sum()), "tolov": float(d["_pay_total_mln"].sum())})

    summary_rows = [{
        "name": "Jami",
        "partiya": int(df["_partiya"].sum()),
        "vazn": float(df["_weight_tn"].sum()),
        "qiymat": float(df["_value_usd_k"].sum()),
        "tolov": float(df["_pay_total_mln"].sum()),
    }]
    for _, r in regimes.iterrows():
        summary_rows.append({"key": {"view": "regime_posts", "regime": r["_regime_code"]}, "name": r["_regime_code"], "partiya": int(r.partiya), "vazn": float(r.vazn), "qiymat": float(r.qiymat), "tolov": float(r.tolov)})
    expired_summary = [{
        "name": "Jami muddati o'tgan",
        "partiya": int(expired["_partiya"].sum()),
        "vazn": float(expired["_weight_tn"].sum()) if len(expired) else 0.0,
        "qiymat": float(expired["_value_usd_k"].sum()) if len(expired) else 0.0,
        "tolov": float(expired["_pay_total_mln"].sum()) if len(expired) else 0.0,
    }]

    archive = load_json(INDEX_PATH, {"reports": []})["reports"]
    by_date = {}
    for r in archive:
        by_date.setdefault(r.get("date", ""), r)
    release = {}
    for days in [1, 3, 5, 10, 30]:
        base_dt = report_date - timedelta(days=days)
        base_key = fmt_date(base_dt)
        item = by_date.get(base_key)
        old_path = Path(item["source"]) if item and item.get("source") else None
        if old_path and old_path.exists():
            try:
                release[str(days)] = released_between(old_path, source)
                release[str(days)]["base_date"] = base_key
            except Exception:
                release[str(days)] = {"partiya": 0, "vazn": 0, "qiymat": 0, "tolov": 0, "rows": [], "base_date": base_key, "missing_date": base_key}
        else:
            release[str(days)] = {"partiya": 0, "vazn": 0, "qiymat": 0, "tolov": 0, "rows": [], "base_date": "", "missing_date": base_key}

    food_total_value = 0.0
    food_data = []
    for row in core.food_rows(df):
        item = {"key": {"food_name": core.to_latin(core.clean(row[1]))}, "name": core.to_latin(core.clean(row[1])), "vazn": float(row[2]), "qiymat": float(row[3]), "over_vazn": float(row[4]), "over_qiymat": float(row[5]), "ulush": 0.0}
        food_data.append(item)
        food_total_value += item["qiymat"]
    for item in food_data:
        item["ulush"] = (item["qiymat"] / food_total_value * 100) if food_total_value else 0.0
    food_total = {"key": {}, "name": "IBK bo'yicha Jami", "vazn": sum(x["vazn"] for x in food_data), "qiymat": food_total_value, "over_vazn": sum(x["over_vazn"] for x in food_data), "over_qiymat": sum(x["over_qiymat"] for x in food_data), "ulush": 100.0 if food_total_value else 0.0}

    def list_to_basic(rows, name_idx=1, partiya_idx=2, vazn_idx=3, qiymat_idx=4, tolov_idx=5):
        out = []
        for row in rows:
            out.append({"key": {}, "name": core.to_latin(core.clean(row[name_idx])), "partiya": int(row[partiya_idx] or 0), "vazn": float(row[vazn_idx] or 0), "qiymat": float(row[qiymat_idx] or 0), "tolov": float(row[tolov_idx] or 0)})
        return out

    wh = df.copy()
    wh["_warehouse_web"] = wh[core.SRC["warehouse"]].map(own_warehouse_label)
    warehouse_data = []
    wh_g = wh.groupby("_warehouse_web", dropna=False).agg(partiya=("_partiya", "sum"), vazn=("_weight_tn", "sum"), qiymat=("_value_usd_k", "sum"), tolov=("_pay_total_mln", "sum")).reset_index().sort_values(["qiymat", "tolov"], ascending=False)
    for _, r in wh_g.iterrows():
        name = core.to_latin(core.clean(r["_warehouse_web"]))
        warehouse_data.append({"key": {"warehouse": name}, "name": name, "partiya": int(r.partiya), "vazn": float(r.vazn), "qiymat": float(r.qiymat), "tolov": float(r.tolov), "_class": "own-row" if name.lower().startswith("o'z ombor") else ""})
    reason_data = list_to_basic(core.reason_rows(df))
    own_all_data = own_company_rows_web(df, report_date, False)
    own_3m_data = own_company_rows_web(df, report_date, True)
    country_data = []
    for _, r in df.groupby(core.SRC["country"], dropna=False).agg(partiya=("_partiya", "sum"), vazn=("_weight_tn", "sum"), qiymat=("_value_usd_k", "sum"), tolov=("_pay_total_mln", "sum"), korxona=(core.SRC["stir"], "nunique")).reset_index().sort_values("qiymat", ascending=False).iterrows():
        country_data.append({"key": {"country": core.to_latin(core.clean(r[core.SRC["country"]]))}, "name": core.to_latin(core.clean(r[core.SRC["country"]])), "partiya": int(r.partiya), "vazn": float(r.vazn), "qiymat": float(r.qiymat), "tolov": float(r.tolov), "korxona": int(r.korxona)})
    all_post_summary = [{"key": {"post": r["_post_group"]}, "post": core.to_latin(core.clean(r["_post_group"])), "partiya": int(r.partiya), "vazn": float(r.vazn), "qiymat": float(r.qiymat), "tolov": float(r.tolov)} for _, r in df.groupby("_post_group", dropna=False).agg(partiya=("_partiya", "sum"), vazn=("_weight_tn", "sum"), qiymat=("_value_usd_k", "sum"), tolov=("_pay_total_mln", "sum")).reset_index().sort_values("qiymat", ascending=False).iterrows()]
    source_post_data = [{"key": {"source_post": r["_source_post_code"]}, "post_kodi": core.clean(r["_source_post_code"]) or "-", "post_nomi": core.clean(r["_source_post_name"]) or "-", "transport": core.clean(r["_source_transport"]) or "-", "partiya": int(r.partiya), "vazn": float(r.vazn), "qiymat": float(r.qiymat), "tolov": float(r.tolov), "korxona": int(r.korxona)} for _, r in df.groupby(["_source_post_code", "_source_post_name", "_source_transport"], dropna=False).agg(partiya=("_partiya", "sum"), vazn=("_weight_tn", "sum"), qiymat=("_value_usd_k", "sum"), tolov=("_pay_total_mln", "sum"), korxona=(core.SRC["stir"], "nunique")).reset_index().sort_values("qiymat", ascending=False).iterrows()]
    transport_data = [{"key": {"transport": r["_source_transport"]}, "name": core.clean(r["_source_transport"]) or "-", "partiya": int(r.partiya), "vazn": float(r.vazn), "qiymat": float(r.qiymat), "tolov": float(r.tolov)} for _, r in df.groupby("_source_transport", dropna=False).agg(partiya=("_partiya", "sum"), vazn=("_weight_tn", "sum"), qiymat=("_value_usd_k", "sum"), tolov=("_pay_total_mln", "sum")).reset_index().sort_values("qiymat", ascending=False).iterrows()]
    post_summary = [{"key": {"post": r["_post_group"]}, "post": core.to_latin(core.clean(r["_post_group"])), "partiya": int(r.partiya), "vazn": float(r.vazn), "qiymat": float(r.qiymat), "tolov": float(r.tolov)} for _, r in expired.groupby("_post_group", dropna=False).agg(partiya=("_partiya", "sum"), vazn=("_weight_tn", "sum"), qiymat=("_value_usd_k", "sum"), tolov=("_pay_total_mln", "sum")).reset_index().sort_values("qiymat", ascending=False).iterrows()]
    regime_posts = {}
    for regime, part in df.groupby("_regime_code", dropna=False):
        regime_posts[regime] = [{"key": {"regime": regime, "post": r["_post_group"]}, "post": core.to_latin(core.clean(r["_post_group"])), "partiya": int(r.partiya), "vazn": float(r.vazn), "qiymat": float(r.qiymat), "tolov": float(r.tolov)} for _, r in part.groupby("_post_group", dropna=False).agg(partiya=("_partiya", "sum"), vazn=("_weight_tn", "sum"), qiymat=("_value_usd_k", "sum"), tolov=("_pay_total_mln", "sum")).reset_index().sort_values("qiymat", ascending=False).iterrows()]
    year_series = df["_date_text"].astype(str).str.extract(r"(\d{4})", expand=False).fillna("")
    df["_gtd_year"] = year_series
    regime_year_post = []
    for _, r in df.groupby(["_post_group", "_regime_code", "_gtd_year"], dropna=False).agg(partiya=("_partiya", "sum"), vazn=("_weight_tn", "sum"), qiymat=("_value_usd_k", "sum"), tolov=("_pay_total_mln", "sum")).reset_index().sort_values(["_post_group", "_regime_code", "_gtd_year"]).iterrows():
        regime_year_post.append({"key": {"post": core.to_latin(core.clean(r["_post_group"])), "regime": r["_regime_code"]}, "post": core.to_latin(core.clean(r["_post_group"])), "rejim": core.clean(r["_regime_code"]), "yil": core.clean(r["_gtd_year"]) or "-", "partiya": int(r.partiya), "vazn": float(r.vazn), "qiymat": float(r.qiymat), "tolov": float(r.tolov)})
    expired_post_regime = []
    for row in core.expired_summary_rows(df, report_date):
        expired_post_regime.append({"key": {"post": core.to_latin(core.clean(row[0]))}, "post": core.to_latin(core.clean(row[0])), "jami_partiya": int(row[1]), "jami_qiymat": float(row[2]), "expired_partiya": int(row[3]), "expired_qiymat": float(row[4]), "ulush": float(row[5]) * 100, "im70_partiya": int(row[6]), "im70_qiymat": float(row[7]), "im74_partiya": int(row[8]), "im74_qiymat": float(row[9]), "tr80_partiya": int(row[10]), "tr80_qiymat": float(row[11])})
    expired_block = []
    for row in core.expired_block_rows(df, report_date):
        expired_block.append({"key": {}, "name": core.to_latin(core.clean(row[0])), "korxona": core.clean(row[1]), "stir": core.clean(row[2]), "partiya": int(row[3] or 0), "qiymat": float(row[4] or 0), "vazn": float(row[5] or 0), "tolov": float(row[6] or 0), "reason": core.to_latin(core.clean(row[7]))})

    return {
        "schema": 5,
        "id": report_id,
        "meta": {"date": fmt_date(report_date), "source": source.name, "deposit": deposit.name if deposit else "", "deposit_date": fmt_date(deposit_date) if deposit_date else ""},
        "kpis": {"partiya": int(df["_partiya"].sum()), "vazn": float(df["_weight_tn"].sum()), "qiymat": float(df["_value_usd_k"].sum()), "tolov": float(df["_pay_total_mln"].sum()), "depozit": float(deposit_total), "depozit_matched": float(companies[companies["depozit"].gt(0)]["depozit"].sum()), "expired": int(expired["_partiya"].sum()), "expired_value": float(expired["_value_usd_k"].sum()) if len(expired) else 0.0},
        "top_value": company_records(companies.sort_values("qiymat", ascending=False).head(30)),
        "top_deposit": company_records(companies[companies["depozit"].gt(0)].sort_values("depozit", ascending=False).head(20)),
        "regimes": [{"key": {"regime": r["_regime_code"]}, "rejim": r["_regime_code"], "partiya": int(r.partiya), "vazn": float(r.vazn), "qiymat": float(r.qiymat), "tolov": float(r.tolov)} for _, r in regimes.iterrows()],
        "summary": summary_rows,
        "expired_summary": expired_summary,
        "goods": goods_rows_web,
        "expired": exp_rows,
        "ages": age_rows,
        "all_post_summary": all_post_summary,
        "source_posts": source_post_data,
        "transport": transport_data,
        "post_summary": post_summary,
        "regime_posts": regime_posts,
        "regime_year_post": regime_year_post,
        "expired_post_regime": expired_post_regime,
        "expired_block": expired_block,
        "food": food_data,
        "food_total": food_total,
        "warehouse": warehouse_data,
        "reason": reason_data,
        "countries": country_data,
        "own_all": own_all_data,
        "own_3m": own_3m_data,
        "released": release,
    }


def update_report_files(job_id: str, report_dir: Path, report_date: datetime, error: str | None = None) -> dict:
    """Har doim dict qaytaradi. report_dir yoki report_date None bo'lsa ham xavfsiz ishlaydi."""
    # None himoyasi
    if report_dir is None or report_date is None:
        files: dict = {
            "status": "xatolik",
            "excel": "",
            "pdf": "",
            "pngs": [],
            "error": error or ("report_dir=None" if report_dir is None else "report_date=None"),
        }
        job = ensure_job(job_id)
        job["data"]["files"] = files
        job["files"] = files
        return files

    try:
        excel_path = report_dir / f"Ombor_malumot_{report_date:%Y_%m_%d}.xlsx"
        if not excel_path.exists():
            old_excel = report_dir / f"IBK_jamlanma_{report_date:%Y_%m_%d}.xlsx"
            if old_excel.exists():
                excel_path = old_excel
        png_path = report_dir / f"IBK_tahlil_{report_date:%Y_%m_%d}.png"
        pdf_path = report_dir / f"IBK_tahlil_{report_date:%Y_%m_%d}.pdf"
        pngs = sorted(report_dir.glob(f"{png_path.stem}*.png"))
        files = {
            "status": "tayyor" if excel_path.exists() and (pdf_path.exists() or pngs) else ("excel_tayyor" if excel_path.exists() else "tayyorlanmoqda"),
            "excel": excel_path.name if excel_path.exists() else "",
            "pdf": pdf_path.name if pdf_path.exists() else "",
            "pngs": [p.name for p in pngs if p.exists()],
        }
    except Exception as exc:
        files = {"status": "xatolik", "excel": "", "pdf": "", "pngs": [], "error": f"files_build: {exc}"}

    if error:
        files["error"] = error

    try:
        data_path = report_dir / "dashboard.json"
        data = load_json(data_path, {})
        if not isinstance(data, dict):
            data = {}
        data["files"] = files
        save_json(data_path, data)
    except Exception:
        pass

    job = ensure_job(job_id)
    if not isinstance(job.get("data"), dict):
        job["data"] = {}
    job["data"]["files"] = files
    job["files"] = files
    return files

def generate_artifacts(job_id: str, stored_source: Path, stored_deposit: Path | None, report_date: datetime, report_dir: Path):
    excel_path = report_dir / f"Ombor_malumot_{report_date:%Y_%m_%d}.xlsx"
    png_path = report_dir / f"IBK_tahlil_{report_date:%Y_%m_%d}.png"
    pdf_path = report_dir / f"IBK_tahlil_{report_date:%Y_%m_%d}.pdf"
    try:
        core.build(stored_source, TEMPLATE_PATH, excel_path, report_date, stored_deposit)
        update_report_files(job_id, report_dir, report_date)
    except Exception as exc:
        update_report_files(job_id, report_dir, report_date, f"Excel: {type(exc).__name__}: {exc}")
        return
    try:
        build_png_report(stored_source, png_path, pdf_path, report_date, stored_deposit)
        update_report_files(job_id, report_dir, report_date)
    except Exception as exc:
        update_report_files(job_id, report_dir, report_date, f"PNG/PDF: {type(exc).__name__}: {exc}")


def run_artifact_job(job_id: str, report_id: str):
    job = ensure_job(job_id)
    try:
        job["status"] = "Excel/PNG/PDF tayyorlanmoqda"
        archive = load_json(INDEX_PATH, {"reports": []})
        if not isinstance(archive, dict):
            archive = {"reports": []}
        reports = archive.get("reports") or []
        item = next((r for r in reports if isinstance(r, dict) and r.get("id") == report_id), None)
        if not item:
            raise FileNotFoundError("Hisobot topilmadi")
        report_dir = Path(item["dir"]) if item.get("dir") else None
        if report_dir is None:
            raise ValueError("run_artifact_job: report_dir=None")
        source = Path(item["source"]) if item.get("source") else None
        if source is None:
            raise ValueError("run_artifact_job: source=None")
        deposit = Path(item["deposit"]) if item.get("deposit") else None
        date_str = item.get("date", "")
        if not date_str:
            raise ValueError("run_artifact_job: date bo'sh")
        report_date = datetime.strptime(date_str, "%d.%m.%Y")
        generate_artifacts(report_id, source, deposit, report_date, report_dir)
        data = load_json(report_dir / "dashboard.json", {})
        if not isinstance(data, dict):
            data = {}
        job = ensure_job(job_id)
        job.update({"status": "tayyor", "data": data})
    except Exception as exc:
        job = ensure_job(job_id)
        job["status"] = "xatolik"
        job["error"] = f"{type(exc).__name__}: {exc}\n{traceback.format_exc()}"



def tolov_summary_rows() -> list[dict]:
    summary = load_json(TOLOV_OUTPUT_DIR / "summary.json", {})
    rows = []
    for filename, name in TOLOV_FILES.items():
        rec = summary.get(filename, {})
        rows.append({
            "name": name,
            "rows": int(float(rec.get("rows", 0) or 0)),
            "sum": float(rec.get("sum", 0) or 0),
            "file": filename,
        })
    return rows


def build_tolov_from_source(source: Path) -> list[dict]:
    if not TOLOV_GENERATOR_PATH.exists():
        raise FileNotFoundError(f"To'lov generatori topilmadi: {TOLOV_GENERATOR_PATH}")
    spec = importlib.util.spec_from_file_location("ibk_tolov_generator", TOLOV_GENERATOR_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError("To'lov generatorini yuklab bo'lmadi")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    summary = module.build(source, TOLOV_OUTPUT_DIR)
    save_json(TOLOV_OUTPUT_DIR / "summary.json", summary)
    return tolov_summary_rows()


def run_job(job_id: str, source: Path, deposit: Path | None, report_date: datetime):
    job = ensure_job(job_id)
    try:
        # Kiruvchi parametrlarni tekshirish
        if report_date is None:
            raise ValueError("run_job: report_date=None, fayl nomidan sana aniqlanmadi")
        if source is None:
            raise ValueError("run_job: source=None")

        job["status"] = "Hisob-kitob qilinyapti"
        report_dir = REPORT_DIR / job_id
        report_dir.mkdir(parents=True, exist_ok=True)
        stored_source = report_dir / source.name
        shutil.copy2(source, stored_source)
        stored_deposit = None
        if deposit:
            stored_deposit = report_dir / deposit.name
            shutil.copy2(deposit, stored_deposit)

        df_for_store = read_report_source(stored_source)
        save_snapshot_to_store(job_id, stored_source, stored_deposit, report_date, df_for_store)
        data = build_dashboard(job_id, stored_source, stored_deposit, report_date)
        if not isinstance(data, dict):
            data = {}
        files = update_report_files(job_id, report_dir, report_date)
        if not isinstance(files, dict):
            files = {"status": "kutilmoqda", "excel": "", "pdf": "", "pngs": []}
        data["files"] = files
        if not data["files"].get("excel"):
            data["files"]["status"] = "kutilmoqda"
        save_json(report_dir / "dashboard.json", data)

        archive = load_json(INDEX_PATH, {"reports": []})
        if not isinstance(archive, dict):
            archive = {"reports": []}
        reports = archive.get("reports")
        if not isinstance(reports, list):
            reports = []
        reports = [r for r in reports if isinstance(r, dict) and r.get("id") != job_id]
        reports.append({"id": job_id, "date": fmt_date(report_date), "source": str(stored_source), "deposit": str(stored_deposit or ""), "dir": str(report_dir)})
        archive["reports"] = clean_archive_records(reports)
        save_json(INDEX_PATH, archive)
        job.update({"status": "tayyor", "data": data})
    except Exception as exc:
        job = ensure_job(job_id)
        job["status"] = "xatolik"
        job["error"] = f"{type(exc).__name__}: {exc}\n{traceback.format_exc()}"


HTML = r"""<!doctype html><html lang="uz"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>IBK Dashboard</title>
<style>
:root{--ink:#172033;--muted:#64748b;--line:#dae2ec;--blue:#174a7c;--green:#1b9e77;--orange:#d95f02;--bg:#f6f8fb;--panel:#fff}*{box-sizing:border-box}body{margin:0;color:var(--ink);font:14px/1.45 "Segoe UI",Arial,sans-serif;background:linear-gradient(120deg,#f6f8fb,#eef6fb,#f7fbf5);background-size:280% 280%;animation:bgshift 18s ease-in-out infinite}header{position:sticky;top:0;background:#fff;border-bottom:1px solid var(--line);z-index:3;padding:12px 18px;display:flex;align-items:center;justify-content:space-between;gap:12px}h1{font-size:22px;margin:0}.muted{color:var(--muted)}main{max-width:1500px;margin:auto;padding:14px}.login,.upload,.panel{background:#fff;border:1px solid var(--line);border-radius:8px;padding:14px}.panel{overflow-x:auto}.upload{display:grid;grid-template-columns:1fr 1fr 170px auto;gap:10px;align-items:end}input,select{width:100%;padding:9px;border:1px solid var(--line);border-radius:6px;background:#fff}label{display:block;margin-bottom:5px;font-weight:700;color:var(--muted)}button,.btn{background:var(--blue);color:#fff;border:0;border-radius:6px;padding:10px 13px;font-weight:700;text-decoration:none;cursor:pointer}.btn.light,button.light{background:#e8eef6;color:var(--ink)}.kpis{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin:12px 0}.kpi{background:#fff;border:1px solid var(--line);border-radius:8px;padding:12px}.kpi b{display:block;font-size:24px;margin-top:5px}.tabs{display:flex;gap:8px;flex-wrap:wrap;margin:12px 0}.tab.active{background:var(--blue);color:#fff}.tab{background:#e8eef6;color:#172033}.grid{display:grid;grid-template-columns:1fr 1fr;gap:12px;align-items:start}h2{font-size:17px;margin:0 0 10px}table{width:100%;min-width:760px;border-collapse:collapse;table-layout:auto}.release-table{min-width:1120px;font-size:12px;table-layout:fixed}.release-table th,.release-table td{padding:4px 5px}.release-table th{white-space:normal!important;word-break:normal;line-height:1.12;vertical-align:middle;text-align:center}.release-table th:nth-child(1),.release-table td:nth-child(1){width:34px;text-align:center}.release-table th:nth-child(2),.release-table td:nth-child(2){width:245px}.release-table th:nth-child(3),.release-table td:nth-child(3){width:82px}.release-table th:nth-child(n+4),.release-table td:nth-child(n+4){width:82px}.release-table th:nth-child(6),.release-table td:nth-child(6),.release-table th:nth-child(9),.release-table td:nth-child(9),.release-table th:nth-child(12),.release-table td:nth-child(12){width:96px}th{background:var(--blue);color:#fff;text-align:left;padding:7px;font-size:13px}td{border:1px solid #d7e0ea;padding:5px 6px;vertical-align:middle}td.text,th.text{white-space:normal;overflow:visible;text-overflow:clip;word-break:normal}td.num,th.num{text-align:right;white-space:nowrap}tbody tr:first-child td{font-weight:800;background:#d7e8d2}tbody tr.own-row td{background:#eaf3ec!important;font-weight:800}th{white-space:normal;line-height:1.2}tr:nth-child(even) td{background:#f8fafc}tbody tr{cursor:pointer}.merged-row td{background:#dfeaf6!important;font-weight:800}.merged-row td:first-child{text-align:left}.bars{display:grid;gap:8px}.barrow{display:grid;grid-template-columns:220px 1fr 120px;gap:8px;align-items:center}.bar{height:24px;background:#e8eef6;position:relative}.bar span{display:block;height:100%;min-width:34px;background:var(--green);color:#fff;font-size:12px;font-weight:700;text-align:right;padding-right:5px;line-height:24px}dialog{width:min(1100px,96vw);border:0;border-radius:8px;padding:0}dialog .head{padding:12px 14px;border-bottom:1px solid var(--line);display:flex;justify-content:space-between}dialog .body{padding:14px;max-height:72vh;overflow:auto}.hidden{display:none}@media(max-width:1000px){.upload,.grid{grid-template-columns:1fr}.kpis{grid-template-columns:repeat(2,1fr)}header{align-items:flex-start;flex-direction:column}}
 .workspace{display:grid;grid-template-columns:210px 1fr;gap:14px;align-items:start}.tabs{position:sticky;top:78px;display:flex;flex-direction:column;gap:8px;margin:12px 0}.tab{width:100%;text-align:left;background:#e8eef6;color:#172033;border-left:5px solid transparent;transition:.18s transform,.18s background,.18s box-shadow}.tab:active,button:active,.btn:active{transform:translateY(1px) scale(.99)}.tab.active{background:#174a7c;color:#fff;border-left-color:#1b9e77;box-shadow:0 2px 8px rgba(23,74,124,.18)}.chart{margin-top:12px;padding-top:10px;border-top:1px solid var(--line)}.barrow{grid-template-columns:minmax(145px,220px) 1fr 110px}.barrow div:first-child{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.bar span{animation:growbar .55s ease-out}.wide{grid-column:1/-1}.partial{background:#fff4cc!important;color:#8a5a00;font-weight:800;border:1px solid #e6b84d}.toolbar{display:flex;gap:8px;flex-wrap:wrap;align-items:center}.danger{background:#b42318!important}.dark{--ink:#e5eefb;--muted:#9fb0c6;--line:#2f3f55;--blue:#2c7fb8;--bg:#101826;--panel:#172033}.dark header,.dark .login,.dark .upload,.dark .panel,.dark .kpi{background:#172033}.dark input,.dark select{background:#101826;color:var(--ink);border-color:var(--line)}.dark tr:nth-child(even) td{background:#1d2a3d}@keyframes bgshift{0%,100%{background-position:0% 50%}50%{background-position:100% 50%}}@keyframes growbar{from{width:0}to{}}@media(max-width:1000px){.workspace{grid-template-columns:1fr}.tabs{position:static;flex-direction:row;overflow:auto}.tab{min-width:max-content}.barrow{grid-template-columns:120px 1fr 82px}}

/* Customs emblem + airport movement background */
body::before{content:"";position:fixed;inset:0;z-index:-3;pointer-events:none;background:linear-gradient(120deg,rgba(246,248,251,.96),rgba(238,246,251,.94),rgba(247,251,245,.96)),url('/assets/gerb-bojxona.jpg') center 46% / min(420px,42vw) auto no-repeat;opacity:1}
body::after{content:"";position:fixed;inset:0;z-index:-2;pointer-events:none;background:linear-gradient(115deg,transparent 0 43%,rgba(23,74,124,.08) 44%,transparent 45% 100%),repeating-linear-gradient(115deg,transparent 0 72px,rgba(27,158,119,.08) 74px,transparent 78px);animation:runwayMove 22s linear infinite}
header::after{content:"";position:fixed;left:-120px;top:92px;width:110px;height:2px;background:linear-gradient(90deg,transparent,rgba(23,74,124,.45),transparent);box-shadow:24px -5px 0 -1px rgba(23,74,124,.22);transform:rotate(-12deg);animation:flightPath 18s linear infinite;pointer-events:none}
@keyframes runwayMove{from{background-position:0 0,0 0}to{background-position:360px 0,420px 0}}
@keyframes flightPath{0%{transform:translateX(-8vw) translateY(0) rotate(-12deg);opacity:0}12%{opacity:.5}70%{opacity:.35}100%{transform:translateX(115vw) translateY(34vh) rotate(-12deg);opacity:0}}
.dark body::before{opacity:.38;filter:brightness(.7)}



/* Premium background system */
body{background:#eef4f7;color:var(--ink)}
body::before{content:"";position:fixed;inset:0;z-index:-4;pointer-events:none;background:
  linear-gradient(120deg,rgba(245,248,250,.92),rgba(229,239,239,.88) 48%,rgba(247,250,247,.93)),
  url('/assets/gerb-bojxona.jpg') right 6vw top 96px / min(330px,28vw) auto no-repeat;
  filter:saturate(.92);opacity:1}
body::after{content:"";position:fixed;inset:0;z-index:-3;pointer-events:none;background:
  linear-gradient(118deg,transparent 0 36%,rgba(17,70,84,.12) 36.4%,transparent 37.1% 100%),
  linear-gradient(118deg,transparent 0 49%,rgba(28,113,89,.10) 49.2%,transparent 49.8% 100%),
  repeating-linear-gradient(118deg,transparent 0 118px,rgba(17,70,84,.055) 120px,transparent 124px),
  linear-gradient(180deg,rgba(255,255,255,.36),rgba(255,255,255,.72));
  animation:premiumDrift 38s linear infinite}
main::before{content:"";position:fixed;left:calc(210px + 4vw);right:3vw;top:92px;height:130px;z-index:-2;pointer-events:none;background:
  linear-gradient(90deg,rgba(23,74,124,.16),rgba(27,158,119,.09),transparent 70%),
  repeating-linear-gradient(90deg,rgba(23,74,124,.10) 0 1px,transparent 1px 52px);
  border-top:1px solid rgba(23,74,124,.13);border-bottom:1px solid rgba(23,74,124,.08);transform:skewY(-4deg);opacity:.68}
header::before{content:"";position:fixed;left:-160px;top:118px;width:155px;height:3px;z-index:-1;pointer-events:none;background:linear-gradient(90deg,transparent,rgba(18,74,91,.58),rgba(27,158,119,.32),transparent);box-shadow:38px -8px 0 -1px rgba(18,74,91,.20),76px 8px 0 -2px rgba(27,158,119,.18);transform:rotate(-10deg);animation:premiumFlight 24s ease-in-out infinite}
.panel,.kpi,.login,.upload{box-shadow:0 14px 36px rgba(23,43,77,.07);background:rgba(255,255,255,.92);backdrop-filter:blur(8px)}
header{background:rgba(255,255,255,.88);backdrop-filter:blur(12px)}
body.bg-aero::before{background:
  linear-gradient(120deg,rgba(244,248,250,.95),rgba(235,244,243,.88)),
  url('/assets/gerb-bojxona.jpg') center 58% / min(360px,34vw) auto no-repeat;opacity:1}
body.bg-aero::after{background:
  linear-gradient(62deg,transparent 0 44%,rgba(23,74,124,.15) 44.25%,transparent 44.9%),
  linear-gradient(152deg,transparent 0 53%,rgba(27,158,119,.12) 53.25%,transparent 53.8%),
  repeating-linear-gradient(62deg,transparent 0 86px,rgba(23,74,124,.075) 88px,transparent 93px),
  repeating-linear-gradient(152deg,transparent 0 136px,rgba(217,95,2,.065) 138px,transparent 143px),
  linear-gradient(180deg,rgba(255,255,255,.34),rgba(255,255,255,.76));animation:aeroSweep 34s linear infinite}
body.bg-classic::before{background:linear-gradient(120deg,rgba(246,248,251,.97),rgba(238,246,251,.94),rgba(247,251,245,.97)),url('/assets/gerb-bojxona.jpg') center 46% / min(380px,36vw) auto no-repeat}
body.bg-classic::after{background:repeating-linear-gradient(115deg,transparent 0 92px,rgba(23,74,124,.055) 94px,transparent 99px);animation:premiumDrift 42s linear infinite}
@keyframes premiumDrift{from{background-position:0 0,0 0,0 0,0 0}to{background-position:520px 0,360px 0,480px 0,0 0}}
@keyframes aeroSweep{from{background-position:0 0,0 0,0 0,0 0,0 0}to{background-position:420px 0,-380px 0,520px 0,-460px 0,0 0}}
@keyframes premiumFlight{0%{transform:translateX(-8vw) translateY(0) rotate(-10deg);opacity:0}14%{opacity:.7}62%{opacity:.38}100%{transform:translateX(118vw) translateY(24vh) rotate(-10deg);opacity:0}}
.dark body::before,.dark.bg-aero::before,.dark.bg-classic::before{filter:brightness(.55) saturate(.85);opacity:.52}
.dark .panel,.dark .kpi,.dark .login,.dark .upload{background:rgba(23,32,51,.92)}

/* Animated Toshkent-AERO premium scene */
body::before,body::after,header::before{content:none!important;display:none!important}
.sky-scene{position:fixed;inset:0;z-index:0;overflow:hidden;pointer-events:none;background:linear-gradient(180deg,#dff4ff 0%,#f4fbff 42%,#dbeff3 68%,#c6e6ef 100%)}
.sky-scene:before{content:"";position:absolute;inset:0;background:radial-gradient(circle at 18% 18%,rgba(255,255,255,.88),rgba(255,255,255,0) 22%),linear-gradient(110deg,rgba(255,255,255,.38),rgba(255,255,255,0) 38%);opacity:.9}
.sky-scene:after{content:"";position:absolute;inset:0;background-image:linear-gradient(rgba(18,78,116,.045) 1px,transparent 1px),linear-gradient(90deg,rgba(18,78,116,.04) 1px,transparent 1px);background-size:52px 52px;mask-image:linear-gradient(180deg,transparent 0%,#000 36%,transparent 100%)}
.sky-layer{position:absolute;left:0;right:0}
.mountains{bottom:30vh;height:26vh;background:linear-gradient(135deg,transparent 0 22%,rgba(74,112,126,.28) 22% 34%,transparent 34% 42%,rgba(57,103,118,.36) 42% 58%,transparent 58% 66%,rgba(85,128,139,.24) 66% 78%,transparent 78%),linear-gradient(180deg,rgba(255,255,255,.45),rgba(255,255,255,0));clip-path:polygon(0 100%,0 52%,10% 66%,18% 32%,28% 72%,41% 25%,55% 70%,68% 22%,80% 62%,91% 36%,100% 58%,100% 100%);filter:blur(.2px)}
.water{bottom:0;height:24vh;background:linear-gradient(180deg,rgba(113,180,200,.34),rgba(77,145,172,.55)),repeating-linear-gradient(165deg,rgba(255,255,255,.48) 0 2px,transparent 2px 22px);animation:waterMove 16s linear infinite}
.city{bottom:18vh;height:24vh;width:220%;background:linear-gradient(180deg,transparent 0 20%,rgba(18,45,64,.08) 20% 100%),repeating-linear-gradient(90deg,rgba(24,57,78,.45) 0 30px,transparent 30px 44px,rgba(36,84,105,.35) 44px 76px,transparent 76px 88px,rgba(21,70,96,.42) 88px 120px,transparent 120px 140px);clip-path:polygon(0 100%,0 62%,2% 62%,2% 44%,5% 44%,5% 66%,7% 66%,7% 36%,10% 36%,10% 58%,13% 58%,13% 28%,16% 28%,16% 70%,19% 70%,19% 42%,23% 42%,23% 62%,27% 62%,27% 35%,31% 35%,31% 68%,35% 68%,35% 24%,39% 24%,39% 61%,43% 61%,43% 40%,47% 40%,47% 66%,51% 66%,51% 31%,55% 31%,55% 58%,60% 58%,60% 37%,64% 37%,64% 70%,70% 70%,70% 30%,75% 30%,75% 62%,82% 62%,82% 44%,88% 44%,88% 66%,94% 66%,94% 38%,100% 38%,100% 100%);animation:cityDrift 90s linear infinite;opacity:.78}
.city.front{bottom:15vh;height:18vh;opacity:.62;filter:blur(.3px);animation-duration:62s;background:repeating-linear-gradient(90deg,rgba(15,65,87,.58) 0 44px,transparent 44px 58px,rgba(33,91,112,.5) 58px 92px,transparent 92px 116px)}
.tower{position:absolute;right:18vw;bottom:23vh;width:44px;height:150px;background:linear-gradient(180deg,rgba(32,81,101,.52),rgba(22,56,74,.3));clip-path:polygon(22% 100%,38% 28%,16% 28%,16% 14%,84% 14%,84% 28%,62% 28%,78% 100%);box-shadow:0 0 34px rgba(54,118,145,.16)}.tower:before{content:"";position:absolute;left:-22px;top:0;width:88px;height:28px;border-radius:8px;background:linear-gradient(90deg,rgba(255,255,255,.62),rgba(77,132,151,.44));border:1px solid rgba(255,255,255,.5)}.tower:after{content:"";position:absolute;left:19px;top:-42px;width:6px;height:46px;background:rgba(37,93,115,.42);box-shadow:0 -12px 22px rgba(235,247,255,.9)}.runway{position:absolute;right:5vw;bottom:12vh;width:50vw;height:19vh;transform:perspective(620px) rotateX(62deg) skewX(-18deg);transform-origin:bottom right;background:linear-gradient(90deg,rgba(35,49,57,.3),rgba(29,39,47,.72)),repeating-linear-gradient(90deg,transparent 0 52px,rgba(255,255,255,.82) 52px 66px,transparent 66px 118px);border-left:2px solid rgba(255,255,255,.55);border-right:2px solid rgba(255,255,255,.38);box-shadow:0 -18px 80px rgba(20,65,90,.18)}
.plane-wrap{position:absolute;top:11vh;left:-48vw;width:min(760px,58vw);animation:planeCruise 30s linear infinite;animation-delay:-10s;filter:drop-shadow(0 20px 22px rgba(24,55,72,.22));will-change:transform}
.plane-svg{width:100%;height:auto;display:block}.plane-body{fill:url(#planeSkin);stroke:rgba(23,54,72,.38);stroke-width:2}.plane-wing{fill:#d8e9ef;stroke:rgba(23,54,72,.34);stroke-width:2}.plane-tail{fill:#c8dde6}.plane-window{fill:#183f55}.plane-text{font:700 29px Georgia,serif;fill:#155b49;letter-spacing:.4px}.seal-ring{fill:rgba(255,255,255,.78);stroke:#2b8a57;stroke-width:5}.engine{fill:#eef7fa;stroke:#35596a;stroke-width:2}.engine-dark{fill:#264b5c}.bird{position:absolute;width:38px;height:18px;opacity:.58;animation:birdFly 26s linear infinite}.bird:before,.bird:after{content:"";position:absolute;top:8px;width:18px;height:8px;border-top:2px solid rgba(38,78,94,.72);border-radius:50%}.bird:before{left:2px;transform:rotate(18deg)}.bird:after{right:2px;transform:rotate(-18deg)}.b1{top:14vh;left:-5vw;animation-delay:-4s}.b2{top:25vh;left:-12vw;animation-delay:-14s;transform:scale(.72)}.b3{top:18vh;left:-20vw;animation-delay:-22s;transform:scale(.55)}
header,main,dialog{position:relative;z-index:2}.panel,.kpi,.login,.upload{background:rgba(255,255,255,.9);backdrop-filter:blur(18px) saturate(1.15)}.dark .sky-scene{filter:brightness(.72) saturate(.82)}.dark .panel,.dark .kpi,.dark .login,.dark .upload{background:rgba(23,32,51,.9)}
.exec-summary{grid-column:1/-1}.summary-grid{display:grid;grid-template-columns:repeat(5,minmax(145px,1fr));gap:10px}.summary-item{border:1px solid var(--border);border-radius:16px;padding:12px;background:linear-gradient(180deg,rgba(255,255,255,.72),rgba(255,255,255,.38))}.summary-item b{display:block;color:var(--green);margin-bottom:5px}.kpi{cursor:pointer;transition:transform .18s ease,box-shadow .18s ease}.kpi:hover{transform:translateY(-2px);box-shadow:0 14px 30px rgba(19,67,91,.14)}
@keyframes planeCruise{0%{transform:translate3d(-18vw,8vh,0) scale(.76) rotate(-4deg)}45%{transform:translate3d(58vw,-1vh,0) scale(.95) rotate(-1deg)}100%{transform:translate3d(155vw,-8vh,0) scale(1.08) rotate(-3deg)}}
@keyframes cityDrift{from{transform:translateX(0)}to{transform:translateX(-50%)}}@keyframes waterMove{from{background-position:0 0,0 0}to{background-position:0 0,260px 80px}}@keyframes birdFly{from{transform:translateX(0) translateY(0) scale(.8)}50%{transform:translateX(58vw) translateY(-4vh) scale(1)}to{transform:translateX(112vw) translateY(2vh) scale(.7)}}
@media(max-width:900px){.plane-wrap{width:86vw;animation-duration:24s}.summary-grid{grid-template-columns:1fr 1fr}.runway{width:76vw}.city{height:18vh}}
.tab-group-title{margin:8px 8px 6px;padding:10px 12px;border-radius:14px;background:linear-gradient(135deg,#173d67,#246f78);color:#fff;font-weight:800;letter-spacing:.08em;text-align:center;box-shadow:0 10px 24px rgba(25,70,100,.18)}
.tab.payments{background:linear-gradient(135deg,#e9fff6,#fff7df);border-color:#d5eadf;font-weight:800}.dark .tab.payments{background:linear-gradient(135deg,#163c34,#3a321b)}
.overview-note{font-size:12px;color:var(--muted);margin-top:8px}.mini-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px}.module-intro{display:flex;justify-content:space-between;gap:12px;align-items:center;margin-bottom:10px}.module-intro .muted{max-width:760px}
@media(max-width:900px){.mini-grid{grid-template-columns:1fr}.tab-group-title{text-align:left}}


/* Cinematic aviation background inspired by provided reference videos */
.sky-scene{background:radial-gradient(circle at 78% 18%,rgba(255,177,79,.78),rgba(255,177,79,.12) 18%,transparent 37%),radial-gradient(circle at 18% 20%,rgba(80,137,190,.42),transparent 30%),linear-gradient(180deg,#07111f 0%,#12233b 38%,#2d3141 58%,#070b14 100%)!important}
.sky-scene:before{content:""!important;display:block!important;position:absolute;inset:-8%;background:radial-gradient(ellipse at 76% 29%,rgba(255,213,139,.48),transparent 25%),radial-gradient(ellipse at 46% 25%,rgba(255,255,255,.22),transparent 28%),radial-gradient(ellipse at 20% 45%,rgba(168,194,217,.18),transparent 31%),radial-gradient(ellipse at 70% 70%,rgba(248,164,79,.2),transparent 34%);filter:blur(18px);opacity:.95;animation:cloudCinema 34s ease-in-out infinite alternate}
.sky-scene:after{content:""!important;display:block!important;position:absolute;inset:0;background-image:radial-gradient(circle,rgba(255,218,149,.42) 0 1px,transparent 1.6px),radial-gradient(circle,rgba(134,194,232,.18) 0 1px,transparent 1.8px),linear-gradient(180deg,transparent 0%,rgba(0,0,0,.34) 100%);background-size:86px 86px,132px 132px,100% 100%;background-position:0 0,28px 48px,0 0;opacity:.72;mask-image:linear-gradient(180deg,transparent 0%,#000 46%,#000 100%);animation:citySpark 18s linear infinite}
.mountains{bottom:29vh;height:30vh;opacity:.26;background:linear-gradient(135deg,transparent 0 19%,rgba(101,137,151,.42) 19% 34%,transparent 34% 43%,rgba(72,105,126,.52) 43% 59%,transparent 59% 66%,rgba(107,139,145,.32) 66% 80%,transparent 80%)!important;filter:blur(1.5px)}
.water{bottom:0;height:23vh;background:linear-gradient(180deg,rgba(52,93,111,.2),rgba(8,19,32,.72)),repeating-linear-gradient(170deg,rgba(255,190,112,.12) 0 2px,transparent 2px 34px)!important;animation:waterMove 24s linear infinite!important;opacity:.7}
.city{bottom:8vh;height:44vh;width:260%;opacity:.9;transform-origin:bottom center;transform:perspective(820px) rotateX(58deg);clip-path:none!important;background:radial-gradient(circle,rgba(255,204,111,.92) 0 1.6px,transparent 2.8px),radial-gradient(circle,rgba(122,191,231,.62) 0 1px,transparent 2.4px),linear-gradient(90deg,transparent 0 48%,rgba(255,205,119,.38) 48% 50%,transparent 50% 100%),linear-gradient(0deg,transparent 0 47%,rgba(255,255,255,.08) 47% 50%,transparent 50% 100%)!important;background-size:42px 42px,76px 76px,180px 180px,150px 150px!important;background-position:0 0,22px 18px,0 0,0 0!important;animation:cityDrift 72s linear infinite!important;filter:drop-shadow(0 0 16px rgba(255,178,94,.22))}
.city.front{bottom:2vh;height:34vh;opacity:.82;animation-duration:48s!important;background:radial-gradient(circle,rgba(255,223,146,.95) 0 1.8px,transparent 3px),radial-gradient(circle,rgba(100,176,224,.42) 0 1px,transparent 2.4px),linear-gradient(90deg,transparent 0 46%,rgba(255,255,255,.13) 46% 50%,transparent 50% 100%)!important;background-size:34px 34px,66px 66px,128px 128px!important;filter:blur(.15px) drop-shadow(0 0 18px rgba(255,188,94,.18))}
.runway{left:31vw;right:auto;bottom:2vh;width:38vw;height:46vh;transform:perspective(780px) rotateX(68deg);transform-origin:bottom center;background:linear-gradient(90deg,transparent 0 43%,rgba(19,27,35,.86) 43% 57%,transparent 57% 100%),repeating-linear-gradient(0deg,transparent 0 34px,rgba(255,255,255,.7) 34px 48px,transparent 48px 90px),repeating-linear-gradient(90deg,transparent 0 32px,rgba(94,185,255,.5) 32px 38px,transparent 38px 72px)!important;border:0;box-shadow:0 0 90px rgba(62,159,223,.22),inset 0 0 60px rgba(255,207,128,.12);opacity:.74}
.tower{right:14vw;bottom:25vh;opacity:.62;filter:drop-shadow(0 0 18px rgba(255,199,110,.26))}.tower:after{box-shadow:0 -12px 30px rgba(255,214,139,.9)}
.bird{display:none}.plane-wrap{top:8vh;left:50%;width:min(900px,58vw);animation:planeCinematic 28s cubic-bezier(.45,0,.25,1) infinite!important;animation-delay:-8s!important;filter:drop-shadow(0 30px 30px rgba(0,0,0,.42)) drop-shadow(0 0 20px rgba(255,210,132,.18));transform-origin:center center}.plane-text{font:800 24px Georgia,serif!important;fill:#f8f3df!important;letter-spacing:.6px}.plane-body{fill:url(#planeSkin)!important;stroke:rgba(255,255,255,.38)!important}.plane-wing{fill:#cddde8!important;stroke:rgba(255,255,255,.26)!important}.plane-tail{fill:#b8cedb!important}.engine{fill:#dceaf1!important;stroke:#516675!important}.engine-dark{fill:#07131d!important}.seal-ring{fill:rgba(255,255,255,.82)!important;stroke:#2e9b62!important}.plane-window{fill:#0b1d2d!important}
.panel,.kpi,.login,.upload{background:rgba(255,255,255,.88)!important;box-shadow:0 18px 44px rgba(2,12,24,.12)}.dark .panel,.dark .kpi,.dark .login,.dark .upload{background:rgba(18,27,43,.9)!important}.tab-group-title{box-shadow:0 12px 30px rgba(0,0,0,.22)}
@keyframes planeCinematic{0%{transform:translate3d(-78vw,-6vh,0) scale(.34) rotate(-9deg);opacity:0}8%{opacity:.96}42%{transform:translate3d(-12vw,3vh,0) scale(.78) rotate(-2deg);opacity:1}68%{transform:translate3d(10vw,7vh,0) scale(1.08) rotate(1deg);opacity:1}100%{transform:translate3d(92vw,-1vh,0) scale(.48) rotate(7deg);opacity:0}}
@keyframes cloudCinema{from{transform:translate3d(-2%,1%,0) scale(1)}to{transform:translate3d(3%,-2%,0) scale(1.08)}}@keyframes citySpark{from{background-position:0 0,28px 48px,0 0}to{background-position:86px 86px,160px 180px,0 0}}


/* Cinematic polish: glass header and better plane path */
header{background:linear-gradient(180deg,rgba(255,255,255,.82),rgba(255,255,255,.54))!important;backdrop-filter:blur(18px) saturate(1.1);border-bottom:1px solid rgba(255,255,255,.42);box-shadow:0 12px 35px rgba(4,15,31,.08)}
.dark header{background:linear-gradient(180deg,rgba(15,24,40,.84),rgba(15,24,40,.56))!important}
@keyframes planeCinematic{0%{transform:translate3d(-84vw,-7vh,0) scale(.36) rotate(-8deg);opacity:0}10%{opacity:.95}36%{transform:translate3d(-22vw,0vh,0) scale(.72) rotate(-3deg);opacity:1}58%{transform:translate3d(-4vw,6vh,0) scale(1.08) rotate(0deg);opacity:1}78%{transform:translate3d(18vw,3vh,0) scale(.9) rotate(3deg);opacity:.95}100%{transform:translate3d(82vw,-6vh,0) scale(.42) rotate(8deg);opacity:0}}
@media(min-width:1000px){.plane-wrap{width:min(960px,62vw)}}


/* Final professional video-like background asset */
.sky-scene{background-image:linear-gradient(90deg,rgba(4,10,20,.58),rgba(4,10,20,.22) 38%,rgba(4,10,20,.48)),linear-gradient(180deg,rgba(5,11,22,.18),rgba(5,11,22,.34)),url('/assets/ibk-cinematic-airport.png')!important;background-size:cover,cover,cover!important;background-position:center center!important;animation:cinemaBgDrift 38s ease-in-out infinite alternate!important;transform-origin:center center}
.sky-scene:before{content:""!important;display:block!important;position:absolute;inset:0;background:radial-gradient(circle at 70% 34%,rgba(255,190,95,.16),transparent 23%),linear-gradient(90deg,rgba(0,0,0,.25),transparent 44%,rgba(0,0,0,.2));filter:none!important;opacity:1!important;animation:cinemaLightBreath 9s ease-in-out infinite alternate!important}
.sky-scene:after{content:""!important;display:block!important;position:absolute;inset:0;background:linear-gradient(180deg,rgba(255,255,255,.04),rgba(0,0,0,.1)),radial-gradient(circle at 50% 72%,rgba(255,211,131,.18),transparent 18%);background-size:100% 100%!important;opacity:1!important;mask-image:none!important;animation:none!important}
.sky-scene .mountains,.sky-scene .water,.sky-scene .city,.sky-scene .runway,.sky-scene .tower,.sky-scene .bird,.sky-scene .plane-wrap{display:none!important}
header{background:linear-gradient(180deg,rgba(255,255,255,.78),rgba(255,255,255,.48))!important}.panel,.kpi,.login,.upload{background:rgba(255,255,255,.86)!important;backdrop-filter:blur(20px) saturate(1.1)!important}.tab{background:rgba(245,248,253,.88)!important}.tab.active{background:#254f86!important}.tab-group-title{background:linear-gradient(135deg,rgba(18,54,88,.9),rgba(28,105,119,.86))!important}
@keyframes cinemaBgDrift{0%{transform:scale(1.02) translate3d(-.8%,.4%,0)}100%{transform:scale(1.075) translate3d(.9%,-.6%,0)}}@keyframes cinemaLightBreath{0%{opacity:.72}100%{opacity:1}}
@media(max-width:900px){.sky-scene{background-position:54% center!important}.panel,.kpi,.login,.upload{background:rgba(255,255,255,.9)!important}}


/* Visible premium background motion layers */
.sky-scene{animation:cinemaBgDrift 22s ease-in-out infinite alternate!important;will-change:transform,background-position}
.cinema-clouds,.cinema-runway,.cinema-glow,.cinema-vignette{position:absolute;inset:0;pointer-events:none}
.cinema-clouds{inset:-12%;background:radial-gradient(ellipse at 12% 28%,rgba(203,219,234,.18),transparent 24%),radial-gradient(ellipse at 78% 22%,rgba(255,191,105,.14),transparent 28%),radial-gradient(ellipse at 48% 48%,rgba(255,255,255,.08),transparent 32%);filter:blur(20px);mix-blend-mode:screen;animation:cloudLayerMove 18s ease-in-out infinite alternate;opacity:.95}
.cinema-runway{background:repeating-linear-gradient(180deg,transparent 0 34px,rgba(255,224,151,.72) 35px 37px,transparent 38px 78px),repeating-linear-gradient(90deg,transparent 0 42px,rgba(90,190,255,.42) 43px 45px,transparent 46px 92px);clip-path:polygon(42% 58%,58% 58%,82% 100%,18% 100%);filter:blur(.35px) drop-shadow(0 0 12px rgba(255,201,111,.42));opacity:.46;animation:runwayLightFlow 2.8s linear infinite;mix-blend-mode:screen}
.cinema-glow{background:radial-gradient(circle at 50% 38%,rgba(255,205,129,.22),transparent 18%),radial-gradient(circle at 52% 62%,rgba(255,238,174,.18),transparent 16%);mix-blend-mode:screen;animation:planeAuraPulse 4.8s ease-in-out infinite alternate;opacity:.72}
.cinema-vignette{background:linear-gradient(90deg,rgba(2,7,16,.44),transparent 35%,rgba(2,7,16,.38)),linear-gradient(180deg,rgba(0,0,0,.08),rgba(0,0,0,.46));opacity:.82}
@keyframes cloudLayerMove{0%{transform:translate3d(-2.5%,1.2%,0) scale(1.02)}100%{transform:translate3d(2.2%,-1.6%,0) scale(1.08)}}
@keyframes runwayLightFlow{0%{background-position:0 0,0 0;opacity:.34}50%{opacity:.58}100%{background-position:0 78px,92px 0;opacity:.34}}
@keyframes planeAuraPulse{0%{transform:scale(.96);opacity:.46}100%{transform:scale(1.08);opacity:.9}}
@keyframes cinemaBgDrift{0%{transform:scale(1.03) translate3d(-1.4%,.6%,0)}50%{transform:scale(1.065) translate3d(.2%,-.3%,0)}100%{transform:scale(1.095) translate3d(1.1%,-.9%,0)}}


/* Real video background layer */
.bg-video{position:absolute;inset:0;width:100%;height:100%;object-fit:cover;z-index:0;filter:saturate(1.08) contrast(1.04) brightness(.9)}
#bgCanvas{display:none}.cinema-clouds,.cinema-runway,.cinema-glow,.cinema-vignette{z-index:1}.sky-scene:before,.sky-scene:after{z-index:2}.sky-scene{background-image:linear-gradient(90deg,rgba(4,10,20,.58),rgba(4,10,20,.22) 38%,rgba(4,10,20,.48)),url('/assets/ibk-cinematic-airport.png')!important}


/* Lighter video background and expandable module navigation */
.bg-video{filter:saturate(.96) contrast(.98) brightness(1.18)!important;opacity:.92}.cinema-vignette{opacity:.28!important;background:linear-gradient(90deg,rgba(255,255,255,.10),transparent 45%,rgba(255,255,255,.04)),linear-gradient(180deg,rgba(255,255,255,.12),rgba(255,255,255,.02))!important}.cinema-clouds{opacity:.62!important}.cinema-runway{opacity:.28!important}.cinema-glow{opacity:.5!important}.sky-scene:before{background:radial-gradient(circle at 70% 34%,rgba(255,230,180,.24),transparent 24%),linear-gradient(90deg,rgba(255,255,255,.10),transparent 44%,rgba(255,255,255,.08))!important}.panel,.kpi,.login,.upload{background:rgba(255,255,255,.91)!important}.summary-item{background:linear-gradient(180deg,rgba(255,255,255,.86),rgba(255,255,255,.62))!important}.module-parent{margin:6px 8px;padding:12px 14px;border:0;border-radius:16px;background:linear-gradient(135deg,rgba(26,73,121,.94),rgba(40,129,139,.9));color:white;font-weight:850;text-align:left;box-shadow:0 12px 28px rgba(27,78,118,.2);cursor:pointer}.module-parent.pay{background:linear-gradient(135deg,rgba(36,124,96,.94),rgba(210,151,46,.86))}.module-parent.active{outline:2px solid rgba(255,255,255,.85)}.subtabs{padding:4px 0 10px}.tab.sub{margin-left:18px;width:calc(100% - 28px);font-size:13px}.pay-kpis{display:grid;grid-template-columns:repeat(4,minmax(160px,1fr));gap:12px;margin-bottom:14px}.pay-card{padding:14px;border:1px solid var(--border);border-radius:16px;background:rgba(255,255,255,.78)}.pay-card b{display:block;font-size:20px;margin-top:6px}.module-grid{display:grid;grid-template-columns:1.2fr .8fr;gap:14px}@media(max-width:900px){.pay-kpis,.module-grid{grid-template-columns:1fr}.tab.sub{margin-left:10px;width:calc(100% - 14px)}}
.excel-wrap{overflow:auto;border:1px solid #9fb9ce;border-radius:10px;background:white}.excel-table{width:100%;border-collapse:collapse;font-size:13px;background:white}.excel-table th{background:#1f4e78!important;color:white!important;border:1px solid #6f8da8;padding:7px 8px;text-align:center;vertical-align:middle;white-space:normal}.excel-table td{border:1px solid #b7c8d8;padding:6px 8px;vertical-align:middle;background:#fff}.excel-table td.num{text-align:center;font-variant-numeric:tabular-nums}.excel-table td.text{text-align:left}.excel-table tr.total td{background:#e2f0d9!important;font-weight:800}.excel-table tr:nth-child(even):not(.total) td{background:#f7fbff}.excel-table .download-cell{text-align:center}.excel-download{display:inline-flex;align-items:center;justify-content:center;min-width:74px;padding:6px 10px;border-radius:8px;background:#eef6ff;border:1px solid #9fc2e0;color:#164a73;font-weight:800;text-decoration:none}.excel-download:hover{background:#dff0ff}.excel-actions{display:flex;gap:8px;flex-wrap:wrap;margin:8px 0 10px}.excel-title{display:flex;justify-content:space-between;align-items:center;gap:10px}.dark .excel-table td{background:#182538;color:#eaf1f7;border-color:#36546d}.dark .excel-table tr:nth-child(even):not(.total) td{background:#1f3046}.dark .excel-table tr.total td{background:#253f32!important}


/* AeroInfo-style light aviation portal theme */
.sky-scene{background:linear-gradient(180deg,#eef8ff 0%,#f8fcff 44%,#eaf5fb 100%)!important;animation:none!important;transform:none!important}.bg-video{filter:none!important;opacity:1!important}.cinema-clouds,.cinema-runway,.cinema-glow,.cinema-vignette,.sky-scene:before,.sky-scene:after{display:none!important}header{background:rgba(255,255,255,.86)!important;border-bottom:1px solid rgba(17,90,142,.12)!important;box-shadow:0 8px 28px rgba(20,82,128,.08)!important}.panel,.kpi,.login,.upload{background:rgba(255,255,255,.88)!important;border:1px solid rgba(27,98,157,.13)!important;box-shadow:0 14px 34px rgba(38,98,148,.10)!important;backdrop-filter:blur(14px) saturate(1.04)!important}.summary-item{background:linear-gradient(180deg,rgba(255,255,255,.92),rgba(244,249,253,.78))!important;border-color:rgba(29,112,170,.13)!important}.tab{background:rgba(255,255,255,.84)!important;border:1px solid rgba(30,102,165,.13)!important;color:#14304c}.tab.active{background:linear-gradient(135deg,#2275b8,#2aa6b4)!important;color:white!important}.module-parent{background:linear-gradient(135deg,#1e76b7,#35b7c4)!important;box-shadow:0 12px 26px rgba(37,119,177,.18)!important}.module-parent.pay{background:linear-gradient(135deg,#2a9d8f,#7ccfbd)!important}.bar span{background:linear-gradient(90deg,#2f86c8,#4ecdc4)!important}.kpi b,h1,h2{color:#102a43}.muted{color:#5d7082!important}body{background:#eef8ff!important}
header{display:flex!important;align-items:center!important;justify-content:space-between!important;gap:12px!important}header>div:first-child{flex:0 0 auto!important;min-width:165px!important}h1{white-space:nowrap!important}#meta{max-width:280px!important;white-space:nowrap!important;overflow:hidden!important;text-overflow:ellipsis!important}#actions{margin-left:auto!important;display:flex!important;align-items:center!important;justify-content:flex-end!important;gap:8px!important;flex-wrap:wrap!important}.fixed-table{table-layout:fixed!important;min-width:1080px}.fixed-table th,.fixed-table td{overflow:hidden;text-overflow:ellipsis}.expired-total-table{table-layout:fixed!important;min-width:1180px;font-size:12px}.expired-total-table th,.expired-total-table td{text-align:center;padding:5px 6px}.expired-total-table th{white-space:normal!important;word-break:break-word!important;overflow-wrap:anywhere!important;line-height:1.12!important;height:48px!important;vertical-align:middle!important}.expired-total-table td.text{text-align:left}.admin-layout{display:grid;grid-template-columns:400px 1fr;gap:14px;align-items:start}.admin-card{border:1px solid var(--line);border-radius:12px;padding:16px;background:rgba(255,255,255,.72);box-shadow:0 10px 24px rgba(23,43,77,.06)}.admin-form{display:grid;grid-template-columns:110px 1fr;gap:7px 12px;align-items:center}.admin-form input[type=hidden]{display:none}.admin-form .perm-grid,.admin-form .excel-actions{grid-column:1/-1}.admin-form>label{font-weight:600;color:#3a5a7a;font-size:13px;text-align:right;padding-right:4px}.admin-form input,.admin-form select{padding:7px 10px;border:1px solid var(--line);border-radius:8px;font-size:13px;background:var(--panel);color:var(--ink)}.perm-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:8px;margin-top:2px}.perm-grid label{font-weight:600;color:var(--ink);background:#eef6ff;border:1px solid #c8dded;border-radius:8px;padding:8px}.role-pill{display:inline-block;border-radius:999px;padding:3px 9px;background:#e8f3ff;color:#174a7c;font-weight:800}.live-chart{position:relative;min-height:210px;padding:10px}.sparkline{height:150px;border-left:1px solid var(--line);border-bottom:1px solid var(--line);background:linear-gradient(180deg,rgba(34,117,184,.08),rgba(42,166,180,.02));display:flex;align-items:flex-end;gap:8px;padding:12px}.sparkline i{display:block;flex:1;border-radius:6px 6px 0 0;background:linear-gradient(180deg,#2275b8,#4ecdc4);animation:barRise .8s ease both}.sparkline i:nth-child(2n){background:linear-gradient(180deg,#1b9e77,#8bd3c7)}@keyframes barRise{from{height:0;opacity:.3}to{opacity:1}}@media(max-width:1000px){.admin-layout{grid-template-columns:1fr}}
.dark{--ink:#eaf2ff!important;--muted:#b7c6d8!important;--line:#3b5168!important;--blue:#3f8fd5!important;--panel:#17283c!important}.dark body,.dark{background:#0d1b2a!important;color:#eaf2ff!important}.dark header,.dark .panel,.dark .kpi,.dark .login,.dark .upload,.dark .admin-card,.dark .summary-item,.dark .pay-card{background:rgba(20,35,54,.94)!important;color:#eaf2ff!important;border-color:#39536d!important}.dark td{background:#142337!important;color:#edf6ff!important;border-color:#34506b!important}.dark tr:nth-child(even) td{background:#192b42!important}.dark th{background:#255d92!important;color:white!important}.dark input,.dark select{background:#0f1d2d!important;color:#fff!important;border-color:#4d6b86!important}.dark .muted{color:#bed1e5!important}.dark .tab{background:#16283d!important;color:#eaf2ff!important}.dark .tab.active,.dark .module-parent.active{box-shadow:0 0 0 2px #84d4ff inset!important}.dark .bar{background:#203650!important}.dark .excel-table td{background:#142337!important;color:#edf6ff!important}.dark .excel-download{background:#1d3a58;color:#fff;border-color:#4d80aa}.login{max-width:430px;margin:9vh auto!important;text-align:center;border-radius:24px!important;padding:24px!important;box-shadow:0 22px 60px rgba(22,83,132,.18)!important}.login:before{display:none!important;content:none!important}@keyframes sealSpin{from{rotate:0deg}to{rotate:0deg}}@keyframes sealFloat{0%,100%{translate:0 0}50%{translate:0 -7px}}.sky-scene{background:linear-gradient(180deg,rgba(229,246,255,.92),rgba(247,252,255,.86)),radial-gradient(circle at 72% 22%,rgba(77,184,225,.24),transparent 26%),linear-gradient(135deg,#eaf8ff,#f8fdff 50%,#dff2fb)!important}.sky-scene:after{display:block!important;content:"";position:absolute;inset:0;background:linear-gradient(10deg,transparent 60%,rgba(40,144,198,.08) 61%,transparent 62%),repeating-linear-gradient(115deg,transparent 0 90px,rgba(40,144,198,.055) 92px,transparent 96px);animation:bgPlaneDrift 14s linear infinite}.bg-video{opacity:.28!important;filter:saturate(.75) brightness(1.25)!important}.plane-wrap{opacity:.12!important}.module-parent{font-size:15px!important;letter-spacing:.2px}.tab.sub{border-radius:10px!important;background:rgba(255,255,255,.72)!important;color:#12304e!important}.tab.sub.active{background:linear-gradient(135deg,#1e76b7,#35b7c4)!important;color:white!important;font-weight:900!important}.dark .tab.sub{background:#1a314a!important;color:#edf7ff!important;border-color:#4c6c88!important}.dark .tab.sub.active{background:linear-gradient(135deg,#2d84d0,#38bac8)!important;color:#fff!important}.dark .module-parent,.dark .module-parent.pay{color:#fff!important}.dark h1,.dark h2,.dark h3,.dark b,.dark label,.dark .summary-item span{color:#f4fbff!important}.module-parent:hover,.tab:hover,button:hover,.btn:hover,tr:hover td,.archive-card:hover,.summary-item:hover{transform:translateY(-2px)!important;transition:transform .18s ease,box-shadow .18s ease,background .18s ease;box-shadow:0 12px 28px rgba(35,114,178,.16)!important}tr:hover td{background:#eaf6ff!important}.dark tr:hover td{background:#24435f!important}.viz-card{display:grid;grid-template-columns:170px 1fr;gap:16px;align-items:center}.donut{width:150px;height:150px;border-radius:50%;background:conic-gradient(#2477bd var(--p),#d8edf8 0);display:grid;place-items:center;animation:donutPulse 1.5s ease-in-out infinite}.donut:after{content:attr(data-label);display:grid;place-items:center;width:92px;height:92px;border-radius:50%;background:rgba(255,255,255,.94);font-weight:900;color:#17507b;text-align:center}.trend-svg{width:100%;height:160px}.trend-line{fill:none;stroke:#1f8ec1;stroke-width:4;stroke-linecap:round;stroke-dasharray:900;animation:drawLine 1.8s ease-in-out infinite alternate}.trend-area{fill:url(#trendGrad);opacity:.36}.trend-dot{fill:#1b9e77;animation:dotPulse 1.1s ease-in-out infinite}.dark .donut:after{background:#142337;color:#dff6ff}@keyframes donutPulse{0%,100%{filter:saturate(1)}50%{filter:saturate(1.5) brightness(1.05)}}@keyframes drawLine{from{stroke-dashoffset:900}to{stroke-dashoffset:0}}@keyframes dotPulse{0%,100%{r:4}50%{r:7}}@keyframes bgPlaneDrift{from{background-position:0 0,0 0}to{background-position:240px 0,320px 0}}

.flow-map{min-height:310px;border:1px solid rgba(42,116,176,.16);border-radius:18px;background:linear-gradient(180deg,rgba(239,249,255,.92),rgba(255,255,255,.78));padding:10px;overflow:hidden}.flow-map svg{width:100%;height:310px}.flow-map rect{fill:#f4fbff}.flow-map .land{fill:#dff1f8;stroke:#97c6dd;stroke-width:1}.flow-map .land.small{fill:#edf8fb}.flow-map .grid-map path{stroke:rgba(56,131,184,.13);stroke-width:1}.flow-map .route path{fill:none;stroke:url(#routeG);stroke-linecap:round;opacity:.58;stroke-dasharray:520;animation:routeDraw 2.2s ease-in-out infinite alternate}.flow-map .route circle{fill:#2bb7c5;stroke:white;stroke-width:2;filter:drop-shadow(0 3px 7px rgba(20,101,151,.25))}.flow-map text{font-size:9px;fill:#12304e;font-weight:800}.flow-map .uz circle{fill:#1f72b8;stroke:white;stroke-width:3;animation:dotPulse 1.2s ease-in-out infinite}.flow-map .uz text{font-size:12px;fill:#0d335b}.dark .flow-map{background:#142337;border-color:#3b5975}.dark .flow-map rect{fill:#142337}.dark .flow-map .land{fill:#1e3853;stroke:#517da2}.dark .flow-map text{fill:#eaf6ff}.dark .flow-map .grid-map path{stroke:rgba(172,214,245,.16)}@keyframes routeDraw{from{stroke-dashoffset:520;opacity:.34}to{stroke-dashoffset:0;opacity:.82}}
.globe-map{min-height:390px!important;padding:16px!important;background:radial-gradient(circle at 50% 28%,rgba(255,255,255,.98),rgba(226,247,255,.86) 46%,rgba(214,237,249,.72))!important}.globe-map svg{height:380px!important;display:block;margin:auto;max-width:960px}.globe-map rect{fill:rgba(247,253,255,.72)!important}.globe-shadow{fill:rgba(55,133,184,.18);filter:blur(1px)}.globe-sea{fill:url(#globeSea);stroke:rgba(255,255,255,.88);stroke-width:2.2;filter:drop-shadow(0 24px 34px rgba(43,116,170,.22))}.globe-lines ellipse{fill:none;stroke:rgba(255,255,255,.46);stroke-width:1.15}.continent{fill:rgba(255,255,255,.54);stroke:rgba(79,143,184,.42);stroke-width:1}.globe-map .route path{stroke:url(#routeG);stroke-linecap:round;opacity:.68;stroke-dasharray:520;animation:routeDraw 1.7s ease-in-out infinite alternate}.globe-map .route circle{fill:#37c9d3;stroke:#fff;stroke-width:2.2}.globe-map .route text{font-size:8.3px;fill:#153b5f;font-weight:850;text-shadow:0 1px 0 rgba(255,255,255,.75)}.globe-map .uz circle{fill:#1769aa;stroke:#fff;stroke-width:3.2;animation:dotPulse 1.05s ease-in-out infinite}.globe-map .uz text{font-size:12px;fill:#0c3154;font-weight:900;text-shadow:0 1px 0 rgba(255,255,255,.8)}.chart-under-globe{margin-top:14px;padding-top:14px;border-top:1px solid rgba(42,116,176,.16)}.dark .globe-map{background:radial-gradient(circle at 50% 26%,rgba(42,77,105,.92),rgba(18,39,62,.88) 50%,rgba(12,28,45,.86))!important}.dark .globe-map rect{fill:rgba(15,32,51,.72)!important}.dark .globe-shadow{fill:rgba(54,176,223,.14)}.dark .globe-sea{stroke:rgba(194,237,255,.45);filter:drop-shadow(0 24px 36px rgba(0,0,0,.32))}.dark .globe-lines ellipse{stroke:rgba(230,249,255,.28)}.dark .continent{fill:rgba(224,247,255,.26);stroke:rgba(180,229,252,.28)}.dark .globe-map .route text,.dark .globe-map .uz text{fill:#e9f8ff;text-shadow:0 1px 4px rgba(0,0,0,.55)}.dark .chart-under-globe{border-top-color:#375977}
.sample-release-table{min-width:1620px!important}.sample-release-table th{white-space:normal!important;word-break:break-word!important;overflow-wrap:anywhere!important;line-height:1.08!important;height:58px!important;padding:5px 4px!important;vertical-align:middle!important}.sample-release-table td{vertical-align:middle!important}.sample-release-table td:nth-child(2){text-align:left!important}.sample-release-table tr:first-child td{background:#e2f0d9!important;font-weight:850!important}.dark .sample-release-table tr:first-child td{background:#253f32!important}
.compact-archive{min-width:720px!important;table-layout:fixed!important}.compact-archive th,.compact-archive td{padding:6px 8px!important}.compact-archive td:nth-child(2),.compact-archive td:nth-child(3){white-space:nowrap!important;overflow:hidden!important;text-overflow:ellipsis!important}.compact-archive button{padding:6px 10px!important}
/* Final UX polish */
body{background:#eef8ff!important}.sky-scene{background:linear-gradient(180deg,rgba(236,248,255,.95),rgba(248,253,255,.9)),radial-gradient(circle at 78% 18%,rgba(65,170,220,.18),transparent 28%),linear-gradient(120deg,#e7f6ff,#ffffff 48%,#dff1fb)!important}.sky-scene:before{display:block!important;content:""!important;position:absolute;inset:0;background:repeating-linear-gradient(110deg,transparent 0 78px,rgba(34,126,190,.055) 80px,transparent 84px),linear-gradient(15deg,transparent 58%,rgba(33,128,194,.08) 59%,transparent 61%)!important;animation:bgPlaneDrift 18s linear infinite!important}.sky-scene:after{display:block!important;content:""!important;position:absolute;inset:0;background:radial-gradient(ellipse at 50% 115%,rgba(31,91,145,.16),transparent 42%),linear-gradient(180deg,transparent 0 62%,rgba(65,150,201,.08) 63%,transparent 64%)!important;animation:none!important}
.login{position:relative!important;max-width:min(760px,94vw)!important;min-height:72vh!important;margin:5vh auto!important;display:grid!important;place-items:center!important;background:transparent!important;border:0!important;box-shadow:none!important;overflow:visible!important}.login:before{display:none!important}.login-seal-wrap{position:absolute;inset:0;display:grid;place-items:center;cursor:pointer;transition:opacity .9s ease,transform .9s ease}.login-seal{width:min(560px,82vw);height:auto;animation:sealFloatCalm 4.8s ease-in-out infinite;transform-origin:center center;mix-blend-mode:multiply}.login.active .login-seal-wrap{opacity:0;transform:scale(1.1);pointer-events:none}.login-box{position:relative;z-index:2;width:min(420px,92vw);padding:26px;border-radius:22px;background:rgba(255,255,255,.88);border:1px solid rgba(76,145,196,.18);box-shadow:0 24px 70px rgba(28,87,141,.18);backdrop-filter:blur(18px) saturate(1.08);opacity:0;transform:translateY(18px) scale(.96);pointer-events:none;transition:opacity .55s ease,transform .55s ease}.login.active .login-box{opacity:1;transform:translateY(0) scale(1);pointer-events:auto}.login-form-stack{display:grid!important;grid-template-columns:1fr!important;gap:12px!important}.pass-wrap{display:flex;gap:6px;align-items:center}.pass-wrap input{flex:1}.eye-btn{width:46px;padding:9px!important;border-radius:9px!important;background:#eaf3fb!important;color:#143b61!important}.designer-line{font-weight:800;color:#1f5f93}.header-clock{font-weight:800;color:#1d6a9f}.busy-spinner{display:inline-block;width:14px;height:14px;margin-left:8px;border:2px solid rgba(255,255,255,.45);border-top-color:#fff;border-radius:50%;vertical-align:-2px;animation:spinBusy .75s linear infinite}.btn.light .busy-spinner,button.light .busy-spinner{border-color:rgba(31,80,120,.25);border-top-color:#1f5f93}.is-busy{opacity:.82;pointer-events:none}#actions .btn,#actions button{min-height:36px}.dark #actions button.light,.dark #actions .btn.light{background:#284968!important;color:#f6fbff!important;border:1px solid #5e86a8!important}.dark #actions .btn,.dark #actions button{color:#fff!important}.dark .login-box{background:rgba(16,31,49,.92)!important;color:#eef8ff!important;border-color:#416684!important}.dark .eye-btn{background:#203b57!important;color:#fff!important}
table{table-layout:fixed!important}th{white-space:normal!important;word-break:break-word!important;overflow-wrap:anywhere!important;text-align:center!important;vertical-align:middle!important;line-height:1.12!important;min-height:42px!important}td{height:34px!important;max-height:34px!important;vertical-align:middle!important}td.text{white-space:nowrap!important;overflow:hidden!important;text-overflow:ellipsis!important;word-break:normal!important}td.num{text-align:center!important}.fixed-table th,.fixed-table td{overflow:hidden!important}.excel-table th{white-space:normal!important;word-break:break-word!important;overflow-wrap:anywhere!important}.excel-table td.text{white-space:nowrap!important;overflow:hidden!important;text-overflow:ellipsis!important}.merged-row td,.expired-total-table td:first-child{white-space:nowrap!important;overflow:hidden!important;text-overflow:ellipsis!important}
.transport-table{min-width:1010px!important}.transport-table th{height:52px!important;padding:5px 4px!important}.transport-table td{height:34px!important;padding:4px 6px!important}.transport-table td:nth-child(2){text-align:left!important;white-space:nowrap!important;overflow:hidden!important;text-overflow:ellipsis!important}.transport-viz{display:grid;grid-template-columns:360px 1fr;gap:18px}.ring-grid{display:grid;grid-template-columns:repeat(2,minmax(140px,1fr));gap:12px}.transport-ring{position:relative;min-height:150px;border-radius:18px;background:linear-gradient(180deg,rgba(255,255,255,.92),rgba(236,248,255,.86));border:1px solid rgba(50,123,180,.16);display:grid;place-items:center;text-align:center;cursor:pointer;overflow:hidden;box-shadow:0 10px 24px rgba(30,96,150,.09);transition:.2s}.transport-ring:before{content:"";position:absolute;width:104px;height:104px;border-radius:50%;background:conic-gradient(#1e79bd calc(var(--p)*1%),#d9eef8 0);animation:ringLoad .9s ease both;animation-delay:var(--delay)}.transport-ring:after{content:"";position:absolute;width:68px;height:68px;border-radius:50%;background:rgba(255,255,255,.95)}.transport-ring b,.transport-ring span,.transport-ring small{position:relative;z-index:1}.transport-ring b{align-self:end;font-size:13px;color:#143b61}.transport-ring span{font-weight:900;font-size:22px;color:#0d6da6}.transport-ring small{align-self:start;color:#64748b;font-size:11px;max-width:130px}.transport-ring:hover{transform:translateY(-4px) scale(1.015);box-shadow:0 16px 36px rgba(31,111,178,.18)}.flow-list{display:grid;gap:8px}.flow-row{display:grid;grid-template-columns:255px 1fr 92px;gap:10px;align-items:center;padding:8px 10px;border:1px solid rgba(37,108,166,.13);border-radius:12px;background:rgba(255,255,255,.72);cursor:pointer;transition:.18s}.flow-row:hover{transform:translateX(4px);box-shadow:0 10px 24px rgba(32,103,163,.13)}.flow-name{min-width:0;display:grid;grid-template-columns:58px 1fr;gap:8px;align-items:center}.flow-name span{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.flow-track{height:24px;border-radius:999px;background:#e7f1f8;position:relative;overflow:hidden}.flow-track i{display:block;height:100%;border-radius:999px;background:linear-gradient(90deg,#32c6cf,#1d72b8);animation:flowGrow .75s ease both}.flow-track em{position:absolute;inset:0;display:grid;place-items:center;color:#fff;font-style:normal;font-weight:900;font-size:12px;text-shadow:0 1px 2px rgba(0,0,0,.32)}.flow-num{text-align:right;font-weight:800;color:#174a7c}.dark .transport-ring,.dark .flow-row{background:rgba(20,35,54,.88);border-color:#39536d}.dark .transport-ring:after{background:#142337}.dark .transport-ring b,.dark .flow-name b,.dark .flow-num{color:#eaf6ff}.dark .transport-ring small{color:#b7c6d8}.dark .flow-track{background:#263f59}.dark .flow-name span{color:#eaf6ff}@keyframes ringLoad{from{transform:scale(.84);opacity:.2}to{transform:scale(1);opacity:1}}@keyframes flowGrow{from{width:0}to{}}@media(max-width:1000px){.transport-viz{grid-template-columns:1fr}.flow-row{grid-template-columns:210px 1fr 82px}.ring-grid{grid-template-columns:repeat(2,1fr)}}
.sample-like{border-collapse:collapse!important;table-layout:fixed!important;font-family:"Segoe UI",Arial,sans-serif}.sample-like th{background:#b6dde8!important;color:#111!important;border:1px solid #333!important;font-size:12px!important;font-weight:800!important;white-space:normal!important;line-height:1.08!important;text-align:center!important;vertical-align:middle!important;padding:4px!important}.sample-like td{border:1px solid #333!important;height:24px!important;max-height:24px!important;font-size:12px!important;padding:3px 5px!important;vertical-align:middle!important}.sample-like tbody tr:first-child td,.sample-like .grand-total td{background:#eaf1dd!important;font-weight:800!important}.sample-like .merged-row td{background:#d9eaf7!important;font-weight:800!important;text-align:center!important}.sample-like .sub-total td{background:#f2f7fb!important;font-weight:800!important}.sample-like td.text,.sample-like .col-korxona,.sample-like .col-company,.sample-like .col-goods,.sample-like .col-tovar,.sample-like .col-reason{text-align:left!important;white-space:nowrap!important;overflow:hidden!important;text-overflow:ellipsis!important}.sample-like .col-reason,td.col-reason{white-space:nowrap!important;overflow:hidden!important;text-overflow:ellipsis!important}.expired-sample{min-width:1120px!important}.expired-sample col:nth-child(1),.expired-sample td:nth-child(1),.expired-sample th:nth-child(1){width:330px!important}.food-sample{min-width:820px!important}.food-sample col:nth-child(1){width:56px!important}.food-sample col:nth-child(2){width:340px!important}.regime-year-table{max-width:920px!important;min-width:820px!important}.col-korxona,.col-company{min-width:260px;text-align:left!important;white-space:nowrap!important;overflow:hidden!important;text-overflow:ellipsis!important}.col-goods,.col-tovar{white-space:nowrap!important;overflow:hidden!important;text-overflow:ellipsis!important}.col-reason{white-space:nowrap!important;overflow:hidden!important;text-overflow:ellipsis!important}body.login-screen header{display:flex!important}body.login-screen #actions{display:none!important}body.login-screen #app,body.login-screen #dash,body.login-screen #tabs,body.login-screen #view,body.login-screen #kpis,.login-screen .workspace,.login-screen .tabs{display:none!important}body.login-screen main{max-width:100%!important}.login-screen .plane-wrap,.login-screen .bird,.login-screen .city,.login-screen .mountains,.login-screen .water,.login-screen .runway,.login-screen .tower,.login-screen .cinema-clouds,.login-screen .cinema-runway,.login-screen .cinema-glow,.login-screen .cinema-vignette{display:none!important}.login-screen .bg-video{display:block!important;opacity:.56!important;filter:brightness(1.18) saturate(.82) contrast(.96)!important}.login-screen #bgCanvas{display:block!important}.login-screen .sky-scene{background:radial-gradient(circle at 50% 32%,rgba(255,255,255,.96),rgba(228,245,255,.84) 42%,rgba(207,232,248,.72))!important}.login-screen .sky-scene:before,.login-screen .sky-scene:after{display:none!important}.login-seal{mix-blend-mode:multiply!important;filter:none!important}.eye-btn{font-size:18px!important;line-height:1!important}
dialog{width:min(1760px,98.5vw)!important;max-width:98.5vw!important}dialog .body{max-height:82vh!important}.details-wide{min-width:1580px!important}.real-globe{min-height:430px!important;padding:12px!important}
.flow-legend{display:flex;flex-wrap:wrap;gap:14px;justify-content:center;align-items:center;margin-top:10px;font-size:12px;color:#557086;font-weight:700}
.flow-legend .lg{font-size:16px;line-height:1}
.flow-legend .lg.hi{color:#16a34a}.flow-legend .lg.mid{color:#f59e0b}.flow-legend .lg.lo{color:#5b9bd1}
.flow-legend .lg-size{font-size:11px;color:#7aa9c9}
.dark .flow-legend{color:#b9d4ea}
.globe-caption{text-align:center;margin-top:8px;font-size:12px;color:#557086;font-weight:700}.designer-line{display:block!important;font-size:12px!important;font-weight:700!important;color:#557086!important;margin-top:2px}.dark .designer-line{color:#b9d4ea!important}
.flights-map{height:460px;border-radius:14px;overflow:hidden;border:1px solid var(--line);margin-bottom:8px}
.plane-marker div{font-size:20px;line-height:1;color:#174a7c;text-shadow:0 0 3px #fff,0 0 1px #fff;transform-origin:center center}
.dark .plane-marker div{color:#9fd4ff;text-shadow:0 0 3px #0a1726,0 0 1px #0a1726}
.trend-chart-wrap{width:100%}
.trend-chart{width:100%;height:230px;display:block}
.trend-grid{stroke:rgba(120,150,180,.18);stroke-width:1}
.trend-axis{font-size:10px;fill:#7a93ab;font-weight:700}
.trend-line{fill:none;stroke-width:2.4;stroke-linejoin:round}
.trend-line.c0{stroke:#1f72b8}.trend-line.c1{stroke:#1b9e77}
.trend-dot.c0{fill:#1f72b8}.trend-dot.c1{fill:#1b9e77}
.trend-dot{stroke:#fff;stroke-width:1.4}
.trend-legend-row{display:flex;gap:16px;justify-content:center;margin-top:6px;font-size:12px;font-weight:700;color:#557086}
.trend-legend.c0{color:#1f72b8}.trend-legend.c1{color:#1b9e77}
.dark .trend-grid{stroke:rgba(180,229,252,.12)}
.dark .trend-axis{fill:#9fc3df}
.dark .trend-legend-row{color:#b9d4ea}
@keyframes sealFloatCalm{0%,100%{transform:translateY(0) scale(1)}50%{transform:translateY(-14px) scale(1.015)}}@keyframes spinBusy{to{transform:rotate(360deg)}}
.login-error{min-height:20px;margin-top:8px;color:#b42318;font-weight:800}.dark .login-error{color:#ffb4a8}.danger,.logout-btn{background:#c0392b!important;color:#fff!important}.logout-btn:hover{background:#9f2f25!important}.lang-btn{min-width:44px;padding:8px 10px!important}.regime-year-table{font-size:15px!important}.regime-year-table th,.regime-year-table td{text-align:center!important;font-size:15px!important;padding:5px 6px!important}.regime-year-table .col-name,.regime-year-table td.col-name{text-align:center!important;vertical-align:middle!important}.expired-detail-sample{min-width:1190px!important}.expired-detail-sample th{height:42px!important}.expired-detail-sample td{height:28px!important}.expired-detail-sample .col-reason{white-space:nowrap!important;overflow:hidden!important;text-overflow:ellipsis!important}.sample-like .merged-row td{text-align:center!important}.sample-like .col-korxona{text-align:left!important}.forgot-link{margin-top:6px;background:transparent!important;color:#174a7c!important;padding:4px 0!important;text-align:left}.dark .forgot-link{color:#9bd3ff!important}body.login-screen header{display:flex!important}body.login-screen #actions{display:none!important}body.login-screen #app,body.login-screen #dash,body.login-screen #tabs,body.login-screen #view,body.login-screen #kpis,.login-screen .workspace,.login-screen .tabs{display:none!important}.ibk-map-panel{background:linear-gradient(180deg,rgba(8,27,48,.94),rgba(13,43,70,.90))!important;color:#eaf7ff!important;border-color:rgba(119,201,255,.24)!important}.ibk-map-panel .muted,.ibk-map-panel span{color:#b9dff4!important}body.logged-in main{max-width:1500px!important;margin:auto!important;padding:14px!important;display:block!important}
body.logged-in #app{display:block!important}
body.logged-in #dash{display:block!important}
body.logged-in .workspace{display:grid!important;grid-template-columns:230px minmax(0,1fr)!important;gap:16px!important;align-items:start!important}
body.logged-in #tabs{display:flex!important;position:sticky!important;top:86px!important;flex-direction:column!important;gap:8px!important;margin:12px 0!important}
body.logged-in #view{min-width:0!important;width:100%!important}
body.logged-in .module-parent{display:flex!important;align-items:center!important;gap:9px!important;border-radius:10px!important;padding:12px 13px!important;background:rgba(232,238,246,.92)!important;color:#172033!important;border-left:5px solid transparent!important;text-align:left!important;box-shadow:0 8px 18px rgba(15,50,78,.06)!important}
body.logged-in .module-parent.active{background:#174a7c!important;color:#fff!important;border-left-color:#1b9e77!important}




body.logged-in .subtabs{display:flex!important;flex-direction:column!important;gap:6px!important;margin:2px 0 10px 12px!important}
body.logged-in .tab.sub{font-size:13px!important;padding:9px 10px!important;border-radius:8px!important}
body.login-screen #app,body.login-screen #dash,body.login-screen #tabs,body.login-screen #view,body.login-screen #kpis,.login-screen .workspace,.login-screen .tabs{display:none!important}
/* === HARD LOGIN/APP LAYER FIX === */
body.logged-in #login,
body:not(.login-screen) #login{
  display:none!important;
  visibility:hidden!important;
  opacity:0!important;
  pointer-events:none!important;
  position:fixed!important;
  left:-99999px!important;
  top:-99999px!important;
  width:0!important;
  height:0!important;
  overflow:hidden!important;
}

body.login-screen #login{
  display:grid!important;
  visibility:visible!important;
  opacity:1!important;
  pointer-events:auto!important;
  position:relative!important;
  left:auto!important;
  top:auto!important;
  width:auto!important;
  height:auto!important;
  overflow:visible!important;
}

body.login-screen #app,
body.login-screen #dash{
  display:none!important;
  visibility:hidden!important;
  opacity:0!important;
  pointer-events:none!important;
}

body.logged-in #app,
body.logged-in #dash{
  display:block!important;
  visibility:visible!important;
  opacity:1!important;
  pointer-events:auto!important;
}

body.logged-in main{
  display:block!important;
}

/* hex-bg */body{background:#f8fbff!important}body::before,body::after,header::after,header::before{content:none!important;display:none!important;animation:none!important}.sky-scene{position:fixed!important;top:0!important;left:0!important;right:0!important;bottom:0!important;width:100%!important;height:100%!important;z-index:0!important;overflow:hidden!important;pointer-events:none!important;animation:none!important;transform:none!important;background:#f8fbff!important}.sky-scene *{display:none!important}#hexBg{display:block!important;position:absolute!important;top:0!important;left:0!important;width:100%!important;height:100%!important;z-index:1!important}.sky-scene::before,.sky-scene::after{display:none!important;content:none!important;animation:none!important}.dark .sky-scene{background:#080f1e!important}.panel,.kpi,.upload{background:rgba(255,255,255,.97)!important;backdrop-filter:blur(12px)}.dark .panel,.dark .kpi,.dark .upload{background:rgba(15,28,50,.95)!important}#login,main.login-screen-main{background:transparent!important;border:none!important;box-shadow:none!important;backdrop-filter:none!important}header{background:rgba(255,255,255,.08)!important;backdrop-filter:none!important;border-bottom:1px solid rgba(100,160,220,.15)!important;box-shadow:none!important}.dark header{background:rgba(8,15,30,.12)!important}body.login-screen header{background:transparent!important;border-bottom:none!important;box-shadow:none!important}
@keyframes cfmDash{to{stroke-dashoffset:-28}}.cfm-flow-line{animation:cfmDash 2.2s linear infinite}.cfm-legend{background:rgba(255,255,255,.93);border:1px solid #c8d8ea;border-radius:8px;padding:8px 12px;font-size:11px;line-height:1.7;color:#1a3a5c;box-shadow:0 2px 8px rgba(0,0,0,.1)}.dark .cfm-legend{background:rgba(18,32,50,.93);border-color:#3b5168;color:#eaf2ff}
</style></head><body class="login-screen"><div class="sky-scene" aria-hidden="true"><video id="bgVideo" class="bg-video" autoplay muted playsinline></video><canvas id="bgCanvas" width="1280" height="720"></canvas><div class="cinema-clouds"></div><div class="cinema-runway"></div><div class="cinema-glow"></div><div class="cinema-vignette"></div><div class="sky-layer mountains"></div><div class="sky-layer city"></div><div class="sky-layer city front"></div><div class="runway"></div><div class="tower"></div><div class="sky-layer water"></div><div class="bird b1"></div><div class="bird b2"></div><div class="bird b3"></div><div class="plane-wrap"><svg class="plane-svg" viewBox="0 0 900 360"><defs><linearGradient id="planeSkin" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#ffffff"/><stop offset="0.48" stop-color="#d8e5ec"/><stop offset="1" stop-color="#8ea5b4"/></linearGradient><clipPath id="sealClip"><circle cx="450" cy="132" r="24"/></clipPath></defs><path class="plane-tail" d="M420 93 450 18 480 93 464 112 436 112Z"/><path class="plane-wing" d="M103 168 450 116 797 168 770 204 494 176 465 268 435 268 406 176 130 204Z"/><path class="plane-body" d="M382 105c16-54 120-54 136 0 18 61 9 164-24 211-20 29-68 29-88 0-33-47-42-150-24-211Z"/><path class="plane-body" d="M338 148c47-30 177-30 224 0l-28 35c-38-21-130-21-168 0Z" opacity=".85"/><ellipse class="engine" cx="294" cy="203" rx="48" ry="35"/><ellipse class="engine-dark" cx="294" cy="205" rx="25" ry="20"/><ellipse class="engine" cx="606" cy="203" rx="48" ry="35"/><ellipse class="engine-dark" cx="606" cy="205" rx="25" ry="20"/><circle class="seal-ring" cx="450" cy="132" r="31"/><image href="/assets/gerb-bojxona.jpg" x="426" y="108" width="48" height="48" clip-path="url(#sealClip)"/><text class="plane-text" x="450" y="198" text-anchor="middle">Toshkent-AERO IBK</text><g><rect class="plane-window" x="409" y="91" width="14" height="8" rx="4"/><rect class="plane-window" x="431" y="86" width="14" height="8" rx="4"/><rect class="plane-window" x="453" y="86" width="14" height="8" rx="4"/><rect class="plane-window" x="475" y="91" width="14" height="8" rx="4"/></g><path d="M154 175c210-61 382-61 592 0" fill="none" stroke="rgba(255,255,255,.45)" stroke-width="4"/><path d="M410 302c25 19 55 19 80 0" fill="none" stroke="rgba(255,255,255,.38)" stroke-width="5" stroke-linecap="round"/></svg></div></div><header><div><h1>IBK Dashboard</h1><div class="muted" id="meta">Kirish kerak</div><div class="muted"><span id="clock" class="header-clock"></span><span class="designer-line">by @aero004</span></div></div><div id="actions"></div></header><main>
<section id="login" class="login login-closed"><div class="login-seal-wrap" onclick="activateLogin()"><img class="login-seal" src="/assets/sticker.webp" alt="Bojxona gerbi"></div><div class="login-box"><h2>Kirish</h2><div class="login-form-stack"><div><label>Login</label><input id="user" autocomplete="username" placeholder="Login"></div><div><label>Parol</label><div class="pass-wrap"><input id="pass" type="password" autocomplete="current-password" placeholder="Parol"><button class="eye-btn" type="button" onclick="togglePassword()" title="Ko'rsatish/yashirish">&#128065;</button></div></div><button id="loginBtn" onclick="doLogin()">Kirish</button><div id="loginError" class="login-error"></div><button type="button" class="forgot-link" onclick="forgotPassword()">Parolni unutdingizmi?</button></div><div class="muted">Gerb ustiga bosilganda kirish oynasi ochiladi.</div></div></section>
<section id="app" class="hidden"><div id="status" class="muted"></div><div id="dash" class="hidden"><div class="kpis" id="kpis"></div><div class="workspace"><aside class="tabs" id="tabs"></aside><section id="view"></section></div></div></section>
<dialog id="dlg"><div class="head"><b id="dlgTitle">Asos</b><button class="light" onclick="dlg.close()">Yopish</button></div><div class="body" id="dlgBody"></div></dialog>
<script>
let TOKEN=localStorage.ibk_token||"", DATA=null, TAB="home", GROUP="home", ARCHIVE=[], PAYMENTS=[], ME=null, LANG=localStorage.ibk_lang||"uz", COMPANY_TRENDS={periods:[],companies:[]}, GOODS_TRENDS={periods:[],goods:[]};
const I18N={uz:{archive:"Arxiv",upload:"Fayl yuklash",general:"Umumiy",companies:"Korxonalar",expired:"Muddati o'tgan",released:"Nazoratdan yechish",goods:"Tovarlar",food:"Oziq-ovqat",profile:"Profil",settings:"Sozlamalar",admin:"Admin",dark:"Tungi rejim",logout:"Chiqish"},uzc:{archive:"Arxiv",upload:"Fayl yuklash",general:"Umumiy",companies:"Korxonalar",expired:"Muddati o'tgan",released:"Nazoratdan yechish",goods:"Tovarlar",food:"Oziq-ovqat",profile:"Profil",settings:"Sozlamalar",admin:"Admin",dark:"Tungi rejim",logout:"Chiqish"},ru:{archive:"Arxiv",upload:"Zagruzka",general:"Obshiy",companies:"Kompanii",expired:"Prosrochennie",released:"Snyatie s kontrolya",goods:"Tovari",food:"Produkti",profile:"Profil",settings:"Nastroyki",admin:"Admin",dark:"Temniy rejim",logout:"Vixod"}};
function tr(k){return (I18N[LANG]||I18N.uz)[k]||k} function setBg(v){document.body.classList.toggle("bg-aero",v==="aero");document.body.classList.toggle("bg-classic",v==="classic");localStorage.ibk_bg=v} function setLang(v){LANG=v;localStorage.ibk_lang=v;render()} const $=id=>document.getElementById(id);
const fmtN=v=>Math.abs(+v||0)<.005?"0":(+v).toLocaleString("ru-RU",{minimumFractionDigits:2,maximumFractionDigits:2}).replace(/\u00a0/g," "), fmtI=v=>Math.round(+v||0).toLocaleString("ru-RU").replace(/\u00a0/g," "), esc=s=>String(s??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[m]));
function activateLogin(){let el=$("login");if(el)el.classList.add("active")}
function togglePassword(){let p=$("pass");if(p)p.type=p.type==="password"?"text":"password"}
["user","pass"].forEach(id=>{let el=document.getElementById(id);if(el)el.addEventListener("keydown",e=>{if(e.key==="Enter"){e.preventDefault();doLogin()}})});
function updateClock(){let c=$("clock");if(c)c.textContent=new Date().toLocaleString("uz-UZ",{hour:"2-digit",minute:"2-digit",second:"2-digit",day:"2-digit",month:"2-digit",year:"numeric"}).replace(",", "")}
setInterval(updateClock,1000);setTimeout(updateClock,50);
function setBusy(btn,on,text){if(!btn)return;if(on){btn.dataset.old=btn.innerHTML;btn.classList.add("is-busy");btn.disabled=true;btn.innerHTML=(text||btn.textContent)+` <span class="busy-spinner"></span>`}else{btn.classList.remove("is-busy");btn.disabled=false;if(btn.dataset.old)btn.innerHTML=btn.dataset.old}}
function clearBusy(){document.querySelectorAll(".is-busy").forEach(b=>setBusy(b,false))}
async function api(url,opt={}){opt.headers=Object.assign({"X-Token":TOKEN},opt.headers||{});let r=await fetch(url,opt);if(r.status===401){showLogin();throw Error("login")};return await r.json()} function showLogin(){$("login").classList.remove("hidden");$("app").classList.add("hidden");$("meta").textContent="Kirish kerak"}
async function doLogin(){let btn=$("loginBtn"),err=$("loginError");try{if(err)err.textContent="";setBusy(btn,true,"Kirish");let user=($("user")?.value||"").trim(),pass=$("pass")?.value||"";if(!user||!pass){if(err)err.textContent="Login va parolni kiriting";return;}let j=await api("/api/login",{method:"POST",body:JSON.stringify({user,pass})});TOKEN=j.token;localStorage.ibk_token=TOKEN;ME=j.user;await showApp()}catch(e){if(err)err.textContent=(e&&e.message&&e.message!=="login")?e.message:"Login yoki parol xato";}finally{setBusy(btn,false)}} function logout(){localStorage.removeItem("ibk_token");TOKEN="";DATA=null;showLogin()}
async function showApp(){$("login").classList.add("hidden");$("app").classList.remove("hidden");ME=await api("/api/me");LANG=ME.lang||localStorage.ibk_lang||"uz";await loadArchive();await loadPayments();DATA=null;TAB="home";GROUP="home";render()} async function loadArchive(){let j=await api("/api/archive");ARCHIVE=j.reports||[]} async function loadPayments(){try{let j=await api("/api/tolov");PAYMENTS=j.payments||[]}catch(e){PAYMENTS=[]}} async function loadReport(id){DATA=await api("/api/reports/"+id);if(TAB==="upload")TAB="umumiy";render()}
async function poll(id){let j=await api("/api/jobs/"+id);$("status").textContent=j.status;if(j.status==="xatolik"){$("status").textContent=j.error;return}if(j.status!=="tayyor"){setTimeout(()=>poll(id),1800);return}DATA=j.data;TAB="umumiy";await loadArchive();render()}
async function prepareArtifacts(){if(!DATA)return;$("status").textContent="Excel/PNG/PDF tayyorlash boshlandi...";let j=await api("/api/artifacts",{method:"POST",body:JSON.stringify({report:DATA.id})});pollArtifacts(j.job_id)}
async function pollArtifacts(id){let j=await api("/api/jobs/"+id);$("status").textContent=j.status;if(j.status==="xatolik"){$("status").textContent="Excel/PNG/PDF tayyorlashda xatolik. Qayta tayyorlashni bosing yoki logni tekshiramiz.";return}if(j.status!=="tayyor"){setTimeout(()=>pollArtifacts(id),2500);return}DATA=j.data;$("status").textContent="Excel/PNG/PDF tayyor";render()}
async function openGroup(g,tab){if(GROUP===g){GROUP="home";TAB="home";render();return}GROUP=g;TAB=tab;if(g==="bnrte"&&!DATA&&ARCHIVE.length){await loadReport(ARCHIVE[0].id);return}render()}
function landingPanel(){return `<div class="panel wide"><h2>IBK Dashboard</h2><p class="muted">Toshkent-AERO IBK bo'yicha BNRTE nazoratdagi tovarlar, to'lovlar, arxiv, nazoratdan yechilish va tahliliy ko'rsatkichlar yagona dashboardda jamlanadi.</p><div class="summary-grid"><button class="summary-item" onclick="openGroup('bnrte','umumiy')"><b>BNRTE</b><span>Nazoratdagi tovarlar jamlanmasi, muddatlar, omborlar va muddati o'tgan tahlillar.</span></button><button class="summary-item" onclick="openGroup('payments','payments')"><b>To'lovlar</b><span>Baza fayl asosida to'lov turlari bo'yicha Excel jadvallar va tahlil.</span></button><button class="summary-item" onclick="openGroup('common','upload')"><b>Fayl yuklash</b><span>BNRTE yoki To'lovlar uchun yangi asos fayllarni yuklash.</span></button></div></div>${flightsPanelShell()}`}
function renderKpis(){let k=DATA.kpis||{};
  let depHtml=`<div class=kpi onclick="showKpi('depozit')" style="grid-row:span 1">
    <span>Depozit fayl jami (mln so'm)</span><b>${fmtN(k.depozit)}</b>
    <div style="margin-top:6px;padding-top:6px;border-top:1px solid rgba(0,0,0,.08);font-size:12px;color:var(--muted)">
      Mos korxonalar: <b style="color:var(--blue)">${fmtN(k.depozit_matched||0)}</b>
    </div>
  </div>`;
  let rows=[["Partiya",fmtI(k.partiya),"partiya"],["Vazn (tn)",fmtN(k.vazn),"vazn"],["Qiymat (ming $)",fmtN(k.qiymat),"qiymat"],["Kutilayotgan to'lov (mln so'm)",fmtN(k.tolov),"tolov"],["Muddati o'tgan partiya",fmtI(k.expired),"expired"]];
  return rows.map(x=>`<div class=kpi onclick="showKpi('${x[2]}')"><span>${x[0]}</span><b>${x[1]}</b></div>`).join("")+depHtml;}
function table(h,rows,cls=""){rows=rows||[];return `<table class="${cls}"><colgroup>${h.map(x=>`<col style="${x.w?'width:'+x.w:''}">`).join("")}</colgroup><thead><tr>${h.map(x=>`<th class="${x.n?'num':'text'}">${x.t}</th>`).join("")}</tr></thead><tbody>${rows.map((r,ri)=>`<tr class="${r._class||''}" onclick='detail(${JSON.stringify(r.key||{}).replaceAll("'","&#39;")})'>${h.map(x=>{let cls=x.n?'num':'text', raw=r[x.k], val=x.n&&(raw===""||raw===null||raw===undefined)?"":(x.f?x.f(raw):esc(raw)), tip=esc(raw);if(ri===0&&!x.n&&(val==="Jami"||String(val).startsWith("Jami ")))val="IBK bo'yicha Jami";if(x.k==="released_partiya"&&(+raw||0)===0&&((+r.released_qiymat||0)>0.005||(+r.released_vazn||0)>0.005)){cls+=" partial";tip="Qisman yechilgan";val="0"}return `<td class="${cls}" title="${tip}">${val}</td>`}).join("")}</tr>`).join("")}</tbody></table>`}
function total(rows,map){rows=rows||[];let out={};for(let k in map)out[k]=rows.reduce((a,r)=>a+(+r[map[k]]||0),0);return out} function basicTotal(rows,label="IBK bo'yicha Jami",nameKey="name"){let t=total(rows,{partiya:"partiya",vazn:"vazn",qiymat:"qiymat",tolov:"tolov"});let r={key:{},partiya:t.partiya,vazn:t.vazn,qiymat:t.qiymat,tolov:t.tolov};r[nameKey]=label;return [r].concat(rows||[])} function companyTotal(rows){let k=DATA.kpis||{};return [{key:{},korxona:"IBK bo'yicha Jami",stir:"",partiya:k.partiya||0,vazn:k.vazn||0,qiymat:k.qiymat||0,tolov:k.tolov||0,depozit:k.depozit_matched||0}].concat(rows||[])}function expiredTotal(rows){let t=total(rows,{partiya:"partiya",vazn:"vazn",qiymat:"qiymat",tolov:"tolov"});return [{key:{},korxona:"IBK bo'yicha Jami",stir:"",rejim:"",post:"",kun:"",partiya:t.partiya,vazn:t.vazn,qiymat:t.qiymat,tolov:t.tolov}].concat(rows||[])} function foodRows(){let t=DATA.food_total||{name:"IBK bo'yicha Jami",vazn:0,qiymat:0,over_vazn:0,over_qiymat:0,ulush:100};return [t].concat(DATA.food||[])} function regimeSummaryRows(){let rows=(DATA.summary||basicTotal(DATA.regimes||[])).map(r=>{if(["IM70","IM74","TR80"].includes(r.name||r.rejim))return Object.assign({key:{view:"regime_posts",regime:r.name||r.rejim}},r);return r});return rows}
function expiredSummaryRows(){return (DATA.expired_summary||basicTotal(DATA.expired||[])).map((r,i)=>i===0?Object.assign({key:{view:"expired_inline"}},r):r)} function periodRows(r){let rows=(r&&r.rows)||[], t=total(rows,{partiya:"partiya",qiymat:"qiymat",tolov:"tolov"});return [{key:{},company:"IBK bo'yicha Jami",stir:"",decl:"",partiya:t.partiya,qiymat:t.qiymat,tolov:t.tolov}].concat(rows)} function expiredPostRegimeRows(){let rows=DATA.expired_post_regime||[], t=total(rows,{jami_partiya:"jami_partiya",jami_qiymat:"jami_qiymat",expired_partiya:"expired_partiya",expired_qiymat:"expired_qiymat",im70_partiya:"im70_partiya",im74_partiya:"im74_partiya",tr80_partiya:"tr80_partiya"});t.ulush=t.jami_qiymat? t.expired_qiymat/t.jami_qiymat*100:0;t.post="IBK bo'yicha Jami";t.key={};return [t].concat(rows)} function executiveSummary(){if(!DATA)return "";let k=DATA.kpis||{}, top=(DATA.top_value||[])[0]||{}, dep=(DATA.top_deposit||[])[0]||{}, own=(DATA.warehouse||[]).find(r=>(r.name||"")==="O'z ombor")||{}, exp=(DATA.expired_post_regime||[])[0]||{};let items=[['Umumiy nazorat',`${fmtI(k.partiya)} partiya, ${fmtN(k.qiymat)} ming $ qiymat.`],['Muddati o\'tgan',`${fmtI(k.expired)} partiya. Asosiy kesim: ${esc(exp.post||'postlar')}.`],['Eng yirik korxona',`${esc(top.korxona||'-')} - ${fmtN(top.qiymat||0)} ming $.`],['O\'z ombor',`${fmtI(own.partiya||0)} partiya, ${fmtN(own.qiymat||0)} ming $.`],['Depozit yetakchisi',`${esc(dep.korxona||'-')} - ${fmtN(dep.depozit||0)} mln so\'m.`]];return `<div class="panel exec-summary"><h2>Rahbar uchun qisqa xulosa</h2><div class="summary-grid">${items.map(x=>`<div class="summary-item"><b>${x[0]}</b><span>${x[1]}</span></div>`).join("")}</div></div>`}
function showKpi(kind){if(!DATA)return;let titles={partiya:"Partiya kelib chiqishi",vazn:"Vazn kelib chiqishi",qiymat:"Qiymat kelib chiqishi",tolov:"Kutilayotgan to'lov kelib chiqishi",depozit:"Depozit kelib chiqishi",expired:"Muddati o'tgan partiya kelib chiqishi"};dlgTitle.textContent=titles[kind]||"KPI asosi";if(kind==="depozit"){dlgBody.innerHTML=table(companyCols(),companyTotal(DATA.top_deposit||[]));dlg.showModal();return}if(kind==="expired"){dlgBody.innerHTML=table([{k:"post",t:"Post",w:"22%"},{k:"jami_partiya",t:"Jami partiya",n:1,f:fmtI},{k:"jami_qiymat",t:"Jami qiymat (ming $)",n:1,f:fmtN},{k:"expired_partiya",t:"Muddati o'tgan partiya",n:1,f:fmtI},{k:"expired_qiymat",t:"Muddati o'tgan qiymat (ming $)",n:1,f:fmtN},{k:"ulush",t:"Partiyadagi ulushi (%)",n:1,f:fmtN},{k:"im70_partiya",t:"IM70 partiya",n:1,f:fmtI},{k:"im74_partiya",t:"IM74 partiya",n:1,f:fmtI},{k:"tr80_partiya",t:"TR80 partiya",n:1,f:fmtI}],expiredPostRegimeRows());dlg.showModal();return}dlgBody.innerHTML=table([{k:"rejim",t:"Rejim",w:"24%"},{k:"partiya",t:"Partiya",n:1,f:fmtI},{k:"vazn",t:"Vazn (tn)",n:1,f:fmtN},{k:"qiymat",t:"Qiymat (ming $)",n:1,f:fmtN},{k:"tolov",t:"Kutilayotgan to'lov (mln so'm)",n:1,f:fmtN}],basicTotal(DATA.regimes||[],"IBK bo'yicha Jami","rejim"));dlg.showModal()}
 function numbered(rows){return (rows||[]).map((r,i)=>Object.assign({rn:i?i:""},r))} function dateOptions(){let seen=new Set();return ARCHIVE.filter(r=>!seen.has(r.date)&&seen.add(r.date)).map(r=>`<option value="${r.date}">${r.date}</option>`).join("")} function xls(kind){return DATA?`<a class="btn light" href="/api/export?kind=${kind}&report=${DATA.id}&token=${TOKEN}">Excel</a>`:""}
function bars(rows,label,value,fmt){rows=(rows||[]).filter(r=>(+r[value]||0)>0);let max=rows.reduce((m,r)=>Math.max(m,+r[value]||0),0)||1,sum=rows.reduce((a,r)=>a+(+r[value]||0),0)||1;return `<div class=bars>${rows.map(r=>{let pct=(+r[value]||0)/sum*100,w=(+r[value]||0)/max*100;return `<div class=barrow><div title="${esc(r[label])}">${esc(r[label])}</div><div class=bar><span style="width:${Math.max(5,w)}%">${pct.toFixed(1)}%</span></div><b>${fmt(r[value])}</b></div>`}).join("")}</div>`} function by(rows,k){return (rows||[]).slice().sort((a,b)=>(+b[k]||0)-(+a[k]||0))} function expiredBlockTable(rows){rows=rows||[];let body=rows.map(r=>{let group=!r.korxona&&!r.stir&&r.name;if(group)return `<tr class="merged-row"><td colspan="3" title="${esc(r.name)}">${esc(r.name)}</td><td class="num">${fmtI(r.partiya)}</td><td class="num">${fmtN(r.qiymat)}</td><td class="num">${fmtN(r.vazn)}</td><td class="num">${fmtN(r.tolov)}</td><td></td></tr>`;return `<tr title="${esc((r.korxona||'')+' '+(r.reason||''))}"><td class="num">${esc(r.name)}</td><td class="text col-korxona">${esc(r.korxona)}</td><td class="num">${esc(r.stir)}</td><td class="num">${fmtI(r.partiya)}</td><td class="num">${fmtN(r.qiymat)}</td><td class="num">${fmtN(r.vazn)}</td><td class="num">${fmtN(r.tolov)}</td><td class="text col-reason">${esc(r.reason)}</td></tr>`}).join("");return `<table class="sample-like expired-detail-sample"><colgroup><col style="width:44px"><col style="width:330px"><col style="width:116px"><col style="width:74px"><col style="width:96px"><col style="width:96px"><col style="width:116px"><col style="width:320px"></colgroup><thead><tr><th rowspan=2>T/r</th><th colspan=2>Korxona ma'lumotlari</th><th rowspan=2>Partiya</th><th rowspan=2>Qiymati<br>(ming $)</th><th rowspan=2>Vazni<br>(tn)</th><th rowspan=2>Kutilayotgan<br>(mln so'm)</th><th rowspan=2>Saqlanish sababi</th></tr><tr><th>Korxona nomi</th><th>STIR</th></tr></thead><tbody>${body}</tbody></table>`}
function uploadPanel(){return `<div class=stack><div class=panel><h2>Fayl yuklash</h2><div class=muted>Bu modul umumiy: BNRTE va To'lovlar uchun fayllar alohida yuklanadi.</div></div><div class=panel><h2>BNRTE jamlanma</h2><form class="upload" id="upload"><div><label>Asos fayl</label><input name="source" type="file" accept=".xls,.xlsx,.html,.htm" required></div><div><label>Depozit fayl</label><input name="deposit" type="file" accept=".xlsx"></div><div></div><button>BNRTE yuklash</button></form><div class=muted>Hisobot sanasi asos fayl nomidan avtomatik aniqlanadi.</div></div><div class=panel><h2>To'lovlar jadvallari</h2><form class="upload" id="tolovUpload"><div><label>To'lov baza fayli</label><input name="source" type="file" accept=".xlsx,.xls" required></div><div class=muted>04.06+07.06.2026 kabi asos fayl yuklanadi.</div><div></div><button>To'lov jadvallarini shakllantirish</button></form><div id="tolovUploadResult" class=muted>Natijada 13 ta Excel fayl shakllanadi va To'lovlar modulidagi yuklash tugmalariga ulanadi.</div></div></div>`} function bindUpload(){let f=$("upload");if(f)f.onsubmit=async e=>{e.preventDefault();$("status").textContent="BNRTE fayllari yuklanyapti...";let j=await api("/api/reports",{method:"POST",body:new FormData(f)});poll(j.job_id)};let tf=$("tolovUpload");if(tf)tf.onsubmit=async e=>{e.preventDefault();$("status").textContent="To'lovlar shakllantirilyapti...";let j=await api("/api/tolov",{method:"POST",body:new FormData(tf)});PAYMENTS=j.payments||[];$("tolovUploadResult").innerHTML=`Tayyor: ${fmtI(PAYMENTS.reduce((a,r)=>a+(+r.rows||0),0))} qator, ${fmtN(PAYMENTS.reduce((a,r)=>a+(+r.sum||0),0))} so'm. <button class="light" onclick="GROUP='payments';TAB='pay_lists';render()">To'lovlar jadvaliga o'tish</button>`;$("status").textContent="To'lovlar tayyor"}}
const ownCols=[{k:"korxona",t:"Korxona nomi",w:"20%"},{k:"stir",t:"STIR",w:"9%"},{k:"muddat",t:"Muddat",w:"9%"},{k:"kun_hisobi",t:"Kun hisobi",w:"7%",n:1,f:fmtI},{k:"partiya",t:"Partiya",w:"7%",n:1,f:fmtI},{k:"vazn",t:"Vazn (tn)",w:"9%",n:1,f:fmtN},{k:"qiymat",t:"Qiymat (ming $)",w:"10%",n:1,f:fmtN},{k:"tolov",t:"To'lov (mln so'm)",w:"10%",n:1,f:fmtN},{k:"tovar",t:"Tovar",w:"19%"}];
function companyCols(){let k=DATA&&DATA.kpis||{};let dh=k.depozit?`Depozit ${fmtN(k.depozit)} (mln so'm)`:"Depozit (mln so'm)";return [{k:"korxona",t:"Korxona nomi",w:"48%"},{k:"stir",t:"STIR",w:"10%"},{k:"partiya",t:"Partiya",n:1,f:fmtI},{k:"vazn",t:"Vazn (tn)",n:1,f:fmtN},{k:"qiymat",t:"Qiymat (ming $)",n:1,f:fmtN},{k:"tolov",t:"To'lov (mln so'm)",n:1,f:fmtN},{k:"depozit",t:dh,n:1,f:fmtN}]}
const sumCols=[{k:"name",t:"Ko'rsatkich",w:"38%"},{k:"partiya",t:"Partiya",n:1,f:fmtI},{k:"vazn",t:"Vazn (tn)",n:1,f:fmtN},{k:"qiymat",t:"Qiymat (ming $)",n:1,f:fmtN},{k:"tolov",t:"Kutilayotgan to'lov (mln so'm)",n:1,f:fmtN}];
const regimeCols=[{k:"rejim",t:"Rejim",w:"24%"},{k:"partiya",t:"Partiya",n:1,f:fmtI},{k:"vazn",t:"Vazn (tn)",n:1,f:fmtN},{k:"qiymat",t:"Qiymat (ming $)",n:1,f:fmtN},{k:"tolov",t:"Kutilayotgan to'lov (mln so'm)",n:1,f:fmtN}];

function paymentRows(){let rows=PAYMENTS&&PAYMENTS.length?PAYMENTS:(DATA&&DATA.payments?DATA.payments:null);if(!rows)rows=[{name:"Sbor11-12",rows:402,sum:481739660,file:"1. Sbor11-12.xlsx"},{name:"Sbor 44",rows:2,sum:2472000,file:"2. Sbor 44.xlsx"},{name:"Sbor IM40",rows:470,sum:142964000,file:"3. Sbor im40.xlsx"},{name:"Sbor EK10",rows:89,sum:83018000,file:"4. Sbor ek10.xlsx"},{name:"Sbor DR",rows:2053,sum:292969000,file:"5. Sbor dr.xlsx"},{name:"29",rows:418,sum:15559074509.5,file:"6. 29.xlsx"},{name:"20",rows:296,sum:3935764049.81,file:"7. 20.xlsx"},{name:"27",rows:5,sum:37269405.9,file:"8. 27.xlsx"},{name:"25",rows:2,sum:20941130.06,file:"9. 25.xlsx"},{name:"30",rows:1964,sum:2250053286.17,file:"10. 30.xlsx"},{name:"74",rows:60,sum:2485871.08,file:"11. 74.xlsx"},{name:"79",rows:8,sum:494400000,file:"11. 79.xlsx"},{name:"IM42",rows:0,sum:0,file:"12. im42.xlsx"}];return rows.map(r=>Object.assign({},r,{file:r.file||paymentFileByName(r.name)}))}
function paymentFileByName(name){let m={"Sbor 11-12: yig'imlar":"1. Sbor11-12.xlsx","Sbor 44: TR80 yig'imi":"2. Sbor 44.xlsx","Sbor 10: IM40/ND40":"3. Sbor im40.xlsx","Sbor 10: EK10":"4. Sbor ek10.xlsx","Sbor 10: boshqa rejimlar":"5. Sbor dr.xlsx","29-kod to'lovi":"6. 29.xlsx","20-kod to'lovi":"7. 20.xlsx","27-kod to'lovi":"8. 27.xlsx","25-kod to'lovi":"9. 25.xlsx","30-kod to'lovi":"10. 30.xlsx","74-kod to'lovi":"11. 74.xlsx","79-kod to'lovi":"11. 79.xlsx","IM42 bo'yicha ro'yxat":"12. im42.xlsx","Sbor11-12":"1. Sbor11-12.xlsx","Sbor 44":"2. Sbor 44.xlsx","Sbor IM40":"3. Sbor im40.xlsx","Sbor EK10":"4. Sbor ek10.xlsx","Sbor DR":"5. Sbor dr.xlsx","29":"6. 29.xlsx","20":"7. 20.xlsx","27":"8. 27.xlsx","25":"9. 25.xlsx","30":"10. 30.xlsx","74":"11. 74.xlsx","79":"11. 79.xlsx","IM42":"12. im42.xlsx"};return m[name]||""}
function paymentTotal(){let rows=paymentRows();return {rows:rows.reduce((a,r)=>a+(+r.rows||0),0),sum:rows.reduce((a,r)=>a+(+r.sum||0),0)}}
function paymentTable(rows){let body=[{name:"IBK bo'yicha Jami",rows:paymentTotal().rows,sum:paymentTotal().sum,file:""}].concat(rows||[]);return `<div class="excel-wrap"><table class="excel-table"><thead><tr><th style="width:46px">T/r</th><th>To'lov turi</th><th style="width:110px">Qator</th><th style="width:190px">Summa (so'm)</th><th style="width:120px">Excel</th></tr></thead><tbody>${body.map((r,i)=>`<tr class="${i===0?'total':''}"><td class="num">${i===0?'':i}</td><td class="text">${esc(r.name)}</td><td class="num">${fmtI(r.rows)}</td><td class="num">${fmtN(r.sum)}</td><td class="download-cell">${r.file?`<a class="excel-download" href="/download/tolov/${encodeURIComponent(r.file)}?token=${TOKEN}">Yuklash</a>`:''}</td></tr>`).join("")}</tbody></table></div>`}
function paymentModule(mode="overview"){let rows=paymentRows(), t=paymentTotal(), sorted=by(rows,"sum");if(mode==="lists")return `<div class=stack><div class=panel><div class="excel-title"><h2>Hosil bo'lgan to'lov jadvallari</h2><a class="btn light" href="/download/tolov/_all?token=${TOKEN}">Hammasini ZIP</a></div>${paymentTable(rows)}</div><div class=panel><h2>Summa ulushi</h2>${bars(rows,"name","sum",fmtN)}</div></div>`;if(mode==="analysis")return `<div class=grid2><div class=panel><h2>To'lov turlari bo'yicha tahlil</h2>${paymentTable(sorted)}</div><div class=panel><h2>Eng katta summalar</h2>${bars(sorted.slice(0,8),"name","sum",fmtN)}</div><div class=panel><h2>Qatorlar soni</h2>${bars(by(rows,"rows"),"name","rows",fmtI)}</div></div>`;return `<div class=stack><div class=panel><h2>To'lovlar</h2><div class=pay-kpis><div class=pay-card>Jami qator<b>${fmtI(t.rows)}</b></div><div class=pay-card>Jami summa<b>${fmtN(t.sum)} so'm</b></div><div class=pay-card>To'lov turlari<b>${fmtI(rows.length)}</b></div><div class=pay-card>Eng katta tur<b>${esc(sorted[0]?.name||"-")}</b></div></div><div class=module-grid><div>${paymentTable(rows)}</div><div>${bars(sorted.slice(0,8),"name","sum",fmtN)}</div></div></div></div>`}

function overviewPanels(){return `<div class="grid2">${executiveSummary()}<div class=panel><h2>Jami va 70-74-80</h2>${table(sumCols,regimeSummaryRows())}<div class=overview-note>Rejim ustiga bosilganda postlar kesimida asos ochiladi.</div></div><div class=panel><h2 onclick="detail({view:'expired_inline'})">Jami muddati o'tgan</h2>${table(sumCols,expiredSummaryRows())}<div id="expiredInline" class="chart"></div></div><div class="panel wide"><h2>TOP 20 korxona: qiymat ulushi</h2>${bars(DATA.top_value||[],"korxona","qiymat",fmtN)}</div><div class=panel><h2>Rejimlar qiymat ulushi</h2>${bars(DATA.regimes||[],"rejim","qiymat",fmtN)}</div><div class=panel><h2>Muddati o'tgan postlar</h2>${bars(DATA.post_summary||[],"post","partiya",fmtI)}</div><div class=panel><h2>Tovar guruhlari partiya ulushi</h2>${bars(by(DATA.goods||[],"partiya"),"name","partiya",fmtI)}</div></div>`}

function render(){if(TAB.startsWith("pay"))GROUP="payments";if(["umumiy","rejim","korxona","ombor","expired","released","goods","food","muddat","archive"].includes(TAB))GROUP="bnrte";if(["upload","profile","settings","admin"].includes(TAB))GROUP="common";if(TAB==="home")GROUP="home";$("dash").classList.remove("hidden");$("meta").textContent=DATA?DATA.meta.date+" holatiga":"Tizimga xush kelibsiz";$("kpis").innerHTML=(DATA&&GROUP!=="home")?renderKpis():"";let bnrte=[["umumiy",tr("general")],["rejim","Rejimlar"],["korxona",tr("companies")],["ombor","Omborlar"],["expired",tr("expired")],["released",tr("released")],["goods",tr("goods")],["food",tr("food")],["muddat","Muddatlar"],["archive",tr("archive")]], payTabs=[["payments","Umumiy"],["pay_lists","Hosil bo'lgan jadvallar"],["pay_analysis","Tahlil"]], commonTabs=[["upload",tr("upload")],["profile",tr("profile")],["settings",tr("settings")],["admin",tr("admin")]];let bnrteHtml=GROUP==="bnrte"?`<div class="subtabs">${bnrte.map(t=>`<button class="tab sub ${TAB===t[0]?'active':''}" onclick="TAB='${t[0]}';GROUP='bnrte';render()">${t[1]}</button>`).join("")}</div>`:"";let payHtml=GROUP==="payments"?`<div class="subtabs">${payTabs.map(t=>`<button class="tab sub ${TAB===t[0]?'active':''}" onclick="TAB='${t[0]}';GROUP='payments';render()">${t[1]}</button>`).join("")}</div>`:"";let commonHtml=GROUP==="common"?`<div class="subtabs">${commonTabs.map(t=>`<button class="tab sub ${TAB===t[0]?'active':''}" onclick="TAB='${t[0]}';GROUP='common';render()">${t[1]}</button>`).join("")}</div>`:"";$("tabs").innerHTML=`<button class="module-parent ${GROUP==='bnrte'?'active':''}" onclick="openGroup('bnrte','umumiy')">BNRTE</button>${bnrteHtml}<button class="module-parent pay ${GROUP==='payments'?'active':''}" onclick="openGroup('payments','payments')">To'lovlar</button>${payHtml}<button class="module-parent ${GROUP==='common'?'active':''}" onclick="openGroup('common','upload')">Boshqaruv</button>${commonHtml}`;let f=DATA&&DATA.files||{};let fileParts=[];if(f.excel)fileParts.push(`<a class=btn href="/download/${DATA.id}/${f.excel}?token=${TOKEN}">Jamlanma Excel</a>`);if(f.pdf)fileParts.push(`<a class=btn href="/download/${DATA.id}/${f.pdf}?token=${TOKEN}">PDF</a>`);if((f.pngs||[]).length)fileParts.push(`<a class=btn href="/download/${DATA.id}/${f.pngs[0]}?token=${TOKEN}">PNG</a>`);let prepBtn=`<button class="light" onclick="prepareArtifacts()">Excel/PNG/PDF tayyorlash</button>`;let fileBtns=fileParts.length?fileParts.join(" ")+" "+prepBtn:prepBtn;if(f.status&&f.status!=="tayyor")fileBtns+=` <span class="muted">${esc(f.status)}</span>`;$("status").textContent=f.error?"Excel/PNG/PDF tayyorlashda xatolik bor. Qayta tayyorlash tugmasini bosing yoki logni tekshiramiz.":"";$("actions").innerHTML=DATA?`${fileBtns} <button class="light" onclick="TAB='settings';render()">${tr("settings")}</button> <button class="logout-btn" onclick="logout()">${tr("logout")}</button>`:`<button class="light" onclick="TAB='settings';render()">${tr("settings")}</button> <button class="logout-btn" onclick="logout()">${tr("logout")}</button>`;view()}
function view(){let v=$("view");if(TAB==="home"){v.innerHTML=landingPanel();return}if(TAB==="archive"){let seen=new Set(), rows=ARCHIVE.filter(r=>{let k=r.date+"|"+(r.source||"").split(/[\\/]/).pop();if(seen.has(k))return false;seen.add(k);return true});v.innerHTML=`<div class=panel><h2>Arxiv</h2><div class=cards>${rows.map(r=>`<button class="archive-card" onclick="loadReport('${r.id}')"><b>${r.date}</b><span>Asos: ${esc((r.source||'').split(/[\\/]/).pop())}</span><span>Depozit: ${esc((r.deposit||'Depozitsiz').split(/[\\/]/).pop()||'Depozitsiz')}</span></button>`).join("")}</div></div>`;return}if(TAB==="upload"){v.innerHTML=uploadPanel();bindUpload();return}if(TAB==="profile"){v.innerHTML=`<div class=panel><h2>Profil</h2><b>${esc(ME.full_name||ME.user)}</b><p>${esc(ME.position||ME.role)}</p><p>Vakolatlar: ${(ME.perms||[]).join(", ")}</p></div>`;return}if(TAB==="settings"){v.innerHTML=`<div class=panel><h2>Sozlamalar</h2><div class=settings><label>Til</label><select onchange="setLang(this.value)"><option value=uz ${LANG==='uz'?'selected':''}>O'zbek lotin</option><option value=uzc ${LANG==='uzc'?'selected':''}>O'zbek kirill</option><option value=ru ${LANG==='ru'?'selected':''}>Rus tili</option></select><button onclick="document.body.classList.toggle('dark')">${tr("dark")}</button><label>Fon</label><select onchange="setBg(this.value)"><option value="premium">Premium: rasmiy aeroport</option><option value="aero">Premium: runway xaritasi</option><option value="classic">Klassik: yengil gerb</option></select></div></div>`;return}if(TAB==="admin"){v.innerHTML=adminPanel();bindUserForm();loadUsers();return}if(!DATA){v.innerHTML=landingPanel();return}
if(TAB==="umumiy"){v.innerHTML=overviewPanels();return}
if(TAB==="payments"){v.innerHTML=paymentModule("overview");return}if(TAB==="pay_lists"){v.innerHTML=paymentModule("lists");return}if(TAB==="pay_analysis"){v.innerHTML=paymentModule("analysis");return}
if(TAB==="rejim"){v.innerHTML=`<div class=stack><div class=panel><h2>Jami va 70-74-80</h2>${table(sumCols,regimeSummaryRows())}</div><div class=panel><h2>70-74-80 rejimlar kesimida</h2>${table(regimeCols,basicTotal(DATA.regimes||[],"IBK bo'yicha Jami","rejim"))}</div><div class=panel><h2>Rejimlar qiymat ulushi</h2>${bars(DATA.regimes||[],"rejim","qiymat",fmtN)}</div><div class=panel><h2>Rejimlar to'lov ulushi</h2>${bars(DATA.regimes||[],"rejim","tolov",fmtN)}</div></div>`;return}
if(TAB==="ombor"){v.innerHTML=`<div class=stack><div class=panel><h2>Omborlar kesimida</h2>${table(sumCols,basicTotal(DATA.warehouse||[]))}</div><div class=panel><h2>Omborlar qiymat ulushi</h2>${bars(DATA.warehouse||[],"name","qiymat",fmtN)}</div><div class=panel><h2>O'z ombor jami</h2>${table(ownCols,expiredTotal(DATA.own_all||[]).map(r=>Object.assign({korxona:r.korxona||"IBK bo'yicha Jami"},r)))}</div><div class="panel wide"><h2>O'z ombor 3 oy+</h2>${table(ownCols,expiredTotal(DATA.own_3m||[]).map(r=>Object.assign({korxona:r.korxona||"IBK bo'yicha Jami"},r)))}</div><div class=panel><h2>O'z ombor partiya ulushi</h2>${bars(DATA.own_all||[],"korxona","partiya",fmtI)}</div></div><div class="panel wide"><h2>Omborlar oboroti (nazoratdan yechilgan)</h2><div class="filters compact-filters" style="margin-bottom:8px"><label style="font-size:12px;color:#557086">Boshlang'ich sana:</label><select id="omborOborotBase">${dateOptions()}</select><label style="font-size:12px;color:#557086">Yakuniy sana:</label><select id="omborOborotFinal">${dateOptions()}</select><button onclick="buildOmborOborot()">Hisoblash</button></div><div id="omborOborotResult" class="muted">Boshlang'ich sana (eskiroq) va yakuniy sana (yangiroq) tanlang, so'ng Hisoblash tugmasini bosing.</div></div><div class="panel wide" id="warehouseTrendPanel"><h3>Omborlar bo'yicha yuk oqimi tendensiyasi</h3><div class="muted">Yuklanmoqda...</div></div>`;let bSel=$('omborOborotBase'),fSel=$('omborOborotFinal');if(bSel&&fSel&&bSel.options.length>1){bSel.selectedIndex=bSel.options.length-1;fSel.selectedIndex=0;buildOmborOborot();}loadWarehouseTrends();return}
if(TAB==="muddat"){v.innerHTML=`<div class=stack><div class=panel><h2>Muddatlar kesimida</h2>${table([{k:"muddat",t:"Muddat",w:"30%"},{k:"partiya",t:"Partiya",n:1,f:fmtI},{k:"vazn",t:"Vazn (tn)",n:1,f:fmtN},{k:"qiymat",t:"Qiymat (ming $)",n:1,f:fmtN},{k:"tolov",t:"Kutilayotgan to'lov (mln so'm)",n:1,f:fmtN}],basicTotal(DATA.ages||[],"IBK bo'yicha Jami","muddat"))}</div><div class=panel><h2>Saqlanish sabablari</h2>${table(sumCols,basicTotal(DATA.reason||[]))}</div><div class=panel><h2>Muddatlar partiya ulushi</h2>${bars(DATA.ages||[],"muddat","partiya",fmtI)}</div><div class=panel><h2>Saqlanish sabablari qiymat ulushi</h2>${bars(DATA.reason||[],"name","qiymat",fmtN)}</div></div>`;return}
if(TAB==="korxona"){v.innerHTML=`<div class=stack><div class=panel><h2>TOP 20 korxona (qiymat) ${xls("top_value")}</h2>${table(companyCols(),companyTotal(DATA.top_value||[]))}</div><div class=panel><h2>TOP 20 korxona (depozit mablag'lari) ${xls("top_deposit")}</h2>${table(companyCols(),companyTotal(DATA.top_deposit||[]))}</div><div class=panel><h2>Qiymat ulushi</h2>${bars(DATA.top_value||[],"korxona","qiymat",fmtN)}</div></div><div class="panel" id="trendPanel"><h3>Korxonalar bo'yicha davriy tendensiya</h3><div class="muted">Yuklanmoqda...</div></div><div class="panel wide" id="transportCompanyPanel"><h3>Transport kesimida korxonalar — yuk oqimi tendensiyasi</h3><div class="muted">Yuklanmoqda...</div></div><div class="panel wide" id="transportTrendPanel"><h3>Transport turi bo'yicha davriy tendensiya</h3><div class="muted">Yuklanmoqda...</div></div>`;loadCompanyTrends();loadTransportCompanyTrends();loadTransportTrends();return}
if(TAB==="expired"){v.innerHTML=`<div class=stack><div class=panel><h2>Jami muddati o'tgan: postlar va rejimlar kesimida</h2>${expiredTotalExcelTable()}<div class=overview-note>Jadval ustunlari jamlanma Excel vkladkasidagi ko'rinishga yaqin qat'iy kenglikda berildi.</div></div><div class=panel><h2>Muddati o'tgan postlar kesimida</h2>${table([{k:"post",t:"Post",w:"42%"},{k:"partiya",t:"Partiya",n:1,f:fmtI,w:"90px"},{k:"vazn",t:"Vazn (tn)",n:1,f:fmtN,w:"120px"},{k:"qiymat",t:"Qiymat (ming $)",n:1,f:fmtN,w:"130px"},{k:"tolov",t:"To'lov (mln so'm)",n:1,f:fmtN,w:"135px"}],basicTotal(DATA.post_summary||[],"IBK bo'yicha Jami","post"),"fixed-table")}${bars(DATA.post_summary||[],"post","qiymat",fmtN)}${miniChart(DATA.post_summary||[],"qiymat","post")}</div><div class=panel><h2>Muddati o'tgan jamlanma ${xls("expired")}</h2>${expiredBlockTable(DATA.expired_block||[])}</div><div class=panel><h2>Muddati o'tgan korxonalar</h2>${table([{k:"korxona",t:"Korxona nomi",w:"300px"},{k:"stir",t:"STIR",w:"92px"},{k:"rejim",t:"Rejim",w:"70px"},{k:"post",t:"Nazorat posti",w:"170px"},{k:"kun",t:"Kun hisobi",n:1,f:fmtI,w:"80px"},{k:"partiya",t:"Partiya",n:1,f:fmtI,w:"78px"},{k:"qiymat",t:"Qiymat (ming $)",n:1,f:fmtN,w:"120px"},{k:"tolov",t:"To'lov (mln so'm)",n:1,f:fmtN,w:"125px"}],expiredTotal(DATA.expired||[]),"fixed-table")}</div></div>`;return}
if(TAB==="released"){let rel=DATA.released||{};v.innerHTML=`<div class=panel><h2>Nazoratdan yechilishi</h2><div class=filters><select id=relBase>${dateOptions()}</select><select id=relFinal>${dateOptions()}</select><button onclick="buildRelease()">Shakllantirish</button></div><div id=releaseResult></div></div><div class=stack>${["1","3","5","10","30"].map(d=>{let r=rel[d]||{rows:[]};let title=d==="30"?"1 oy ichida yechilgan":`${d} kun ichida yechilgan`;let note=r.missing_date?`<div class=muted>${r.missing_date} sanasidagi asos fayl arxivda yo'q.</div>`:`<div class=muted>Asos sana: ${r.base_date||"-"}</div>`;return `<div class=panel><h2>${title}</h2>${note}${table([{k:"company",t:"Korxona nomi",w:"40%"},{k:"stir",t:"STIR",w:"10%"},{k:"decl",t:"Deklaratsiya",w:"16%"},{k:"partiya",t:"Partiya",n:1,f:fmtI},{k:"qiymat",t:"Qiymat (ming $)",n:1,f:fmtN},{k:"tolov",t:"To'lov (mln so'm)",n:1,f:fmtN}],periodRows(r))}</div>`}).join("")}</div>`;return}
if(TAB==="goods"){let goodsCols=[{k:"name",t:"Tovar guruhi",w:"36%"},{k:"partiya",t:"Partiya",n:1,f:fmtI},{k:"korxona",t:"Korxona",n:1,f:fmtI},{k:"vazn",t:"Vazn (tn)",n:1,f:fmtN},{k:"qiymat",t:"Qiymat (ming $)",n:1,f:fmtN},{k:"tolov",t:"To'lov (mln so'm)",n:1,f:fmtN}], gp=topNWithOther(DATA.goods||[],30,"qiymat","name"), gpart=topNWithOther(DATA.goods||[],30,"partiya","name"), gvazn=topNWithOther(DATA.goods||[],30,"vazn","name");v.innerHTML=`<div class="panel" id="goodsTrendPanel"><h3>Tovarlar bo'yicha davriy tendensiya</h3><div class="muted">Yuklanmoqda...</div></div><div class=grid2><div class="panel wide"><h2>Tovarlar guruhlari ${xls("goods")}</h2>${table(goodsCols,basicTotal(gp))}</div><div class=panel><h2>Partiya bo'yicha ulushi</h2>${bars(gpart,"name","partiya",fmtI)}</div><div class=panel><h2>Qiymat bo'yicha ulushi</h2>${bars(gp,"name","qiymat",fmtN)}</div><div class=panel><h2>Vazn bo'yicha ulushi</h2>${bars(gvazn,"name","vazn",fmtN)}</div></div>`;loadGoodsTrends();return}
if(TAB==="food"){v.innerHTML=`<div class=panel><h2>Oziq-ovqatlar kesimida ${xls("food")}</h2>${table([{k:"name",t:"Oziq-ovqat turi",w:"42%"},{k:"vazn",t:"Vazn (tn)",n:1,f:fmtN},{k:"qiymat",t:"Qiymat (ming $)",n:1,f:fmtN},{k:"over_vazn",t:"3 oy+ vazn (tn)",n:1,f:fmtN},{k:"over_qiymat",t:"3 oy+ qiymat (ming $)",n:1,f:fmtN},{k:"ulush",t:"Qiymatdagi ulushi (%)",n:1,f:fmtN}],foodRows())}${bars(DATA.food||[],"name","qiymat",fmtN)}</div>`;return}}
async function buildRelease(){let base=$("relBase").value, final=$("relFinal").value, j=await api(`/api/release?base=${encodeURIComponent(base)}&final=${encodeURIComponent(final)}`);if(j.missing){$("releaseResult").innerHTML=`<div class=muted>Arxivda yo'q sana: ${j.missing.join(", ")}. Shu davr uchun asos fayl yuklash kerak.</div>`;return}let raw=(j.rows||[]).map(r=>Object.assign({type_text:r.release_type==="qisman"?"Qisman":"To'liq",korxona:r.korxona||r.company||""},r));let rows=numbered([Object.assign({korxona:"IBK bo'yicha Jami",stir:"",decl:"",item_no:"",regime:"",post:"",type_text:""},j.total||{})].concat(raw));let cols=[{k:"rn",t:"T/r",n:1,f:v=>v,w:"34px"},{k:"korxona",t:"Korxona nomi",w:"240px"},{k:"stir",t:"STIR",w:"82px"},{k:"decl",t:"Deklaratsiya",w:"135px"},{k:"item_no",t:"Tovar tartib raqami",w:"74px"},{k:"regime",t:"Rejim",w:"62px"},{k:"post",t:"Post",w:"145px"},{k:"base_partiya",t:`${base} holatiga partiya`,n:1,f:fmtI,w:"76px"},{k:"base_qiymat",t:`${base} holatiga qiymat (ming $)`,n:1,f:fmtN,w:"104px"},{k:"remain_partiya",t:`${final} holatiga qoldiq partiya`,n:1,f:fmtI,w:"80px"},{k:"remain_qiymat",t:`${final} holatiga qoldiq qiymat (ming $)`,n:1,f:fmtN,w:"104px"},{k:"released_partiya",t:"Yechilgan partiya",n:1,f:fmtI,w:"78px"},{k:"released_vazn",t:"Yechilgan vazn (tn)",n:1,f:fmtN,w:"96px"},{k:"released_qiymat",t:"Yechilgan qiymat (ming $)",n:1,f:fmtN,w:"105px"},{k:"released_pct",t:"Qiymatdagi ulushi (%)",n:1,f:fmtN,w:"88px"},{k:"type_text",t:"Holati",w:"82px"}];$("releaseResult").innerHTML=`<h2>${base} - ${final} <a class="btn light" href="/api/export?kind=release&base=${base}&final=${final}&token=${TOKEN}">Excel</a></h2><div class=overview-note>Boshlang'ich davr: ${base}. Yakuniy davr: ${final}. Qisman yechilgan qatorlar qiymat/vazn kamaygan, lekin partiya soni o'zgarmagan holatlarni bildiradi.</div>${table(cols,rows,"release-table fixed-table")}${miniChart(raw,"released_qiymat","korxona")}`}
async function detail(key){if(!DATA||!key||Object.keys(key).length===0)return;if(key.view==="expired_inline"){let el=$("expiredInline");if(el)el.innerHTML=`<h2>Postlar va rejimlar kesimida</h2>${table([{k:"post",t:"Post",w:"22%"},{k:"jami_partiya",t:"Jami partiya",n:1,f:fmtI},{k:"jami_qiymat",t:"Jami qiymat (ming $)",n:1,f:fmtN},{k:"expired_partiya",t:"Muddati o'tgan partiya",n:1,f:fmtI},{k:"expired_qiymat",t:"Muddati o'tgan qiymat (ming $)",n:1,f:fmtN},{k:"ulush",t:"Partiyadagi ulushi (%)",n:1,f:fmtN},{k:"im70_partiya",t:"IM70 partiya",n:1,f:fmtI},{k:"im74_partiya",t:"IM74 partiya",n:1,f:fmtI},{k:"tr80_partiya",t:"TR80 partiya",n:1,f:fmtI}],expiredPostRegimeRows())}`;return}if(key.view==="regime_posts"){let rows=(DATA.regime_posts||{})[key.regime]||[];dlgTitle.textContent=`${key.regime} - postlar kesimida`;dlgBody.innerHTML=table([{k:"post",t:"Post",w:"42%"},{k:"partiya",t:"Partiya",n:1,f:fmtI},{k:"vazn",t:"Vazn (tn)",n:1,f:fmtN},{k:"qiymat",t:"Qiymat (ming $)",n:1,f:fmtN},{k:"tolov",t:"To'lov (mln so'm)",n:1,f:fmtN}],basicTotal(rows,"IBK bo'yicha Jami","post"));dlg.showModal();return}let q=new URLSearchParams({report:DATA.id,filters:JSON.stringify(key)}), j=await api("/api/details?"+q);dlgTitle.textContent="Asos deklaratsiyalar";dlgBody.innerHTML=table([{k:"decl",t:"Deklaratsiya"},{k:"date",t:"Sana"},{k:"regime",t:"Rejim"},{k:"stir",t:"STIR"},{k:"company",t:"Korxona"},{k:"hs",t:"TIF TN"},{k:"goods",t:"Tovar"},{k:"vazn",t:"Vazn (tn)",n:1,f:fmtN},{k:"qiymat",t:"Qiymat (ming $)",n:1,f:fmtN}],j.rows);dlg.showModal()}
function roleTitle(r){return ({admin:"Admin",rahbar:"Rahbar",inspektor:"Inspektor",foydalanuvchi:"Foydalanuvchi",user:"Foydalanuvchi"}[r]||r||"Foydalanuvchi")}function permTitle(p){return ({view:"Ko'rish",upload:"Fayl yuklash",export:"Excel/PDF/PNG",release:"Nazoratdan yechish",admin:"Admin",settings:"Sozlamalar"}[p]||p)}function miniChart(rows,k,label){rows=by(rows||[],k).slice(0,10);let max=rows.reduce((m,r)=>Math.max(m,+r[k]||0),0)||1;return `<div class="live-chart"><div class="sparkline">${rows.map((r,i)=>`<i title="${esc(r[label]||r.name||r.korxona||'')}: ${fmtN(r[k])}" style="height:${Math.max(8,(+r[k]||0)/max*100)}%;animation-delay:${i*.04}s"></i>`).join("")}</div></div>`}function expiredTotalExcelTable(){let cols=[{k:"post",t:"Post",w:"210px"},{k:"jami_partiya",t:"Jami partiya",n:1,f:fmtI,w:"82px"},{k:"jami_vazn",t:"Jami vazn (tn)",n:1,f:fmtN,w:"96px"},{k:"jami_qiymat",t:"Jami qiymat (ming $)",n:1,f:fmtN,w:"110px"},{k:"expired_partiya",t:"Muddati o'tgan partiya",n:1,f:fmtI,w:"92px"},{k:"expired_vazn",t:"Muddati o'tgan vazn (tn)",n:1,f:fmtN,w:"106px"},{k:"expired_qiymat",t:"Muddati o'tgan qiymat (ming $)",n:1,f:fmtN,w:"116px"},{k:"ulush",t:"Partiyadagi ulushi (%)",n:1,f:fmtN,w:"95px"},{k:"im70_partiya",t:"IM70 partiya",n:1,f:fmtI,w:"78px"},{k:"im74_partiya",t:"IM74 partiya",n:1,f:fmtI,w:"78px"},{k:"tr80_partiya",t:"TR80 partiya",n:1,f:fmtI,w:"78px"},{k:"note",t:"Izoh",w:"120px"}];let rows=expiredPostRegimeRows().map(r=>Object.assign({jami_vazn:r.jami_vazn||0,expired_vazn:r.expired_vazn||0,note:""},r));return table(cols,rows,"expired-total-table")}function adminPanel(){return `<div class="admin-layout"><div class="admin-card"><h2>Yangi xodim</h2><form id=userForm class=admin-form><input name=user placeholder="Login" required><input name=password placeholder="Dastlabki parol" required><input name=full_name placeholder="F.I.Sh."><input name=position placeholder="Lavozim"><input name=phone placeholder="Telefon"><input name=post_code placeholder="Post kodi, masalan 00102"><select name=role><option value=foydalanuvchi>Foydalanuvchi</option><option value=inspektor>Inspektor</option><option value=rahbar>Rahbar</option><option value=admin>Admin</option></select><select name=lang><option value=uz>O'zbek lotin</option><option value=uzc>O'zbek kirill</option><option value=ru>Rus tili</option></select><div class=perm-grid><label><input type=checkbox name=perm_view checked> Ko'rish</label><label><input type=checkbox name=perm_upload> Yuklash</label><label><input type=checkbox name=perm_export> Eksport</label><label><input type=checkbox name=perm_release> Yechish</label></div><button>Saqlash</button></form></div><div class="admin-card"><h2>Hodimlar ro'yxati</h2><div id=users></div></div></div>`}function bindUserForm(){let f=$("userForm");if(!f)return;f.onsubmit=async e=>{e.preventDefault();let perms=[];["view","upload","export","release"].forEach(p=>{if(f[`perm_${p}`]?.checked)perms.push(p)});await api("/api/users",{method:"POST",body:JSON.stringify({user:f.user.value,pass:f.password.value,full_name:f.full_name.value,position:f.position.value,phone:f.phone.value,post_code:f.post_code.value,role:f.role.value,lang:f.lang.value,perms})});loadUsers();f.reset();f.perm_view.checked=true}} async function loadUsers(){let box=$("users");if(!box)return;try{let j=await api("/api/users");let rows=(j.users||[]).map(u=>Object.assign({},u,{role_label:`<span class="role-pill">${roleTitle(u.role)}</span>`,perms_text:(u.perms||[]).map(permTitle).join(", ")}));box.innerHTML=table([{k:"user",t:"Login",w:"90px"},{k:"full_name",t:"F.I.Sh.",w:"210px"},{k:"role_label",t:"Rol",w:"120px"},{k:"post_code",t:"Post",w:"75px"},{k:"perms_text",t:"Vakolatlar",w:"260px"}],rows,"fixed-table").replaceAll("&lt;span class=&quot;role-pill&quot;&gt;","<span class=\"role-pill\">").replaceAll("&lt;/span&gt;","</span>")}catch(e){box.innerHTML="Admin vakolati kerak"}}
function svgLineChart(periods,series){let w=900,h=230,padL=56,padR=16,padT=18,padB=34;let allVals=series.flatMap(s=>s.values);let maxV=Math.max(1,...allVals);let n=Math.max(periods.length,1);let x=i=>padL+(n>1?i/(n-1)*(w-padL-padR):(w-padL-padR)/2);let y=val=>h-padB-(val/maxV)*(h-padT-padB);let grid=[0,.25,.5,.75,1].map(f=>{let val=maxV*f,yy=y(val);return `<line x1="${padL}" y1="${yy.toFixed(1)}" x2="${w-padR}" y2="${yy.toFixed(1)}" class="trend-grid"/><text x="${padL-8}" y="${(yy+4).toFixed(1)}" text-anchor="end" class="trend-axis">${fmtN(val)}</text>`}).join("");let lines=series.map((s,si)=>{let pts=s.values.map((v,i)=>`${x(i).toFixed(1)},${y(v).toFixed(1)}`).join(" ");let dots=s.values.map((v,i)=>`<circle cx="${x(i).toFixed(1)}" cy="${y(v).toFixed(1)}" r="3.6" class="trend-dot c${si}"><title>${esc(periods[i]||"")}: ${fmtN(v)}</title></circle>`).join("");return `<polyline points="${pts}" class="trend-line c${si}"/>${dots}`}).join("");let xlabels=periods.map((p,i)=>{if(periods.length>10&&i%Math.ceil(periods.length/10)!==0&&i!==periods.length-1)return "";return `<text x="${x(i).toFixed(1)}" y="${h-8}" text-anchor="middle" class="trend-axis">${esc(p)}</text>`}).join("");let legend=series.map((s,si)=>`<span class="trend-legend c${si}">\u25CF ${esc(s.label)}</span>`).join(" ");return `<div class="trend-chart-wrap"><svg viewBox="0 0 ${w} ${h}" class="trend-chart" preserveAspectRatio="xMidYMid meet">${grid}${lines}${xlabels}</svg><div class="trend-legend-row">${legend}</div></div>`}
function periodRangeCaption(periods){if(!periods.length)return "";if(periods.length===1)return `Davr: ${periods[0]} (hozircha yagona hisobot davri)`;return `Tahlil qilingan davrlar: ${periods[0]} dan ${periods[periods.length-1]} gacha, jami ${periods.length} ta hisobot davri (har bir davr - shu sanada yuklangan hisobot bo'yicha faol nazoratdagi qoldiq)`}
function seriesNZ(p){return (p.value>0.005)||(p.weight>0.005)}
function firstActiveIdx(series){for(let i=0;i<series.length;i++){if(seriesNZ(series[i]))return i}return -1}
function lastActiveIdx(series){for(let i=series.length-1;i>=0;i--){if(seriesNZ(series[i]))return i}return -1}
function trendBuckets(periods,items){let lastIdx=periods.length-1;function hadAny(series,upTo){return series.slice(0,upTo).some(seriesNZ)}let last2idx=periods.length>=2?periods.length-2:0;let active=items.filter(c=>seriesNZ(c.series[lastIdx]));let stoppedAll=items.filter(c=>hadAny(c.series,lastIdx)&&!seriesNZ(c.series[lastIdx]));let newAll=items.filter(c=>!hadAny(c.series,lastIdx)&&seriesNZ(c.series[lastIdx]));let stopped2=periods.length>=2?items.filter(c=>hadAny(c.series,last2idx)&&c.series.slice(last2idx).every(p=>!seriesNZ(p))):[];let new2=periods.length>=2?items.filter(c=>!hadAny(c.series,last2idx)&&c.series.slice(last2idx).some(seriesNZ)):[];return {active,stoppedAll,newAll,stopped2,new2}}
let COMP_PERIOD_M=3,GOODS_PERIOD_M=3;
function filterPeriodItems(periods,items,months){if(!periods.length)return{newItems:[],stoppedItems:[]};let lastDate=new Date(periods[periods.length-1]);let cutoff=new Date(lastDate);cutoff.setMonth(cutoff.getMonth()-months);let cutoffStr=cutoff.toISOString().slice(0,10);let bi=[],wi=[];periods.forEach((p,i)=>{if(p<cutoffStr)bi.push(i);else wi.push(i)});if(!wi.length)return{newItems:[],stoppedItems:[]};let newItems=items.filter(item=>{let hb=bi.some(i=>seriesNZ(item.series[i]));let hw=wi.some(i=>seriesNZ(item.series[i]));return!hb&&hw});let stoppedItems=items.filter(item=>{let hb=bi.some(i=>seriesNZ(item.series[i]));let hw=wi.some(i=>seriesNZ(item.series[i]));return hb&&!hw});return{newItems,stoppedItems}}
function periodBtns(fn,active){return[[1,'🗓 1 oy'],[3,'📅 3 oy'],[6,'📆 6 oy'],[12,'📊 1 yil']].map(([n,lbl])=>`<button class="${n===active?'btn':'btn light'}" style="padding:4px 12px;font-size:12px" onclick="${fn}(${n})">${lbl}</button>`).join('')}
function filterPeriodArr(periods,months){if(!periods||!periods.length)return[];let lastDate=new Date(periods[periods.length-1]);let cutoff=new Date(lastDate);cutoff.setMonth(cutoff.getMonth()-months);let cutoffStr=cutoff.toISOString().slice(0,10);return periods.filter(p=>p>=cutoffStr)}
function renderCompChart(m){let el=$('compChartSect');if(!el||!COMPANY_TRENDS)return;let periods=COMPANY_TRENDS.periods||[],companies=COMPANY_TRENDS.companies||[];if(!periods.length)return;let fp=filterPeriodArr(periods,m);let wi=fp.map(p=>periods.indexOf(p));let totalValues=fp.map((p,j)=>companies.reduce((a,c)=>a+(c.series[wi[j]]?.value||0),0));let totalWeights=fp.map((p,j)=>companies.reduce((a,c)=>a+(c.series[wi[j]]?.weight||0),0));let range=fp.length>=2?`${fp[0]} — ${fp[fp.length-1]}`:fp[0]||"";el.innerHTML=`<div class="panel"><h3>Umumiy oborot tendensiyasi (${range})</h3>${svgLineChart(fp,[{label:"Qiymat (ming $)",values:totalValues},{label:"Vazn (tn)",values:totalWeights}])}<div class="overview-note">Har qaysi korxona qatoriga bosilganda shu korxonaning alohida qiymat/vazn tendensiyasi ochiladi.</div></div>`}
function renderGoodsChart(m){let el=$('goodsChartSect');if(!el||!GOODS_TRENDS)return;let periods=GOODS_TRENDS.periods||[],items=GOODS_TRENDS.goods||[];if(!periods.length)return;let fp=filterPeriodArr(periods,m);let wi=fp.map(p=>periods.indexOf(p));let totalValues=fp.map((p,j)=>items.reduce((a,g)=>a+(g.series[wi[j]]?.value||0),0));let totalWeights=fp.map((p,j)=>items.reduce((a,g)=>a+(g.series[wi[j]]?.weight||0),0));let range=fp.length>=2?`${fp[0]} — ${fp[fp.length-1]}`:fp[0]||"";el.innerHTML=`<div class="panel"><h3>Umumiy tovar oqimi tendensiyasi (${range})</h3>${svgLineChart(fp,[{label:"Qiymat (ming $)",values:totalValues},{label:"Vazn (tn)",values:totalWeights}])}<div class="overview-note">Har qaysi tovar qatoriga bosilganda shu tovarning alohida qiymat/vazn tendensiyasi ochiladi.</div></div>`}
function renderCompNewStopped(m){COMP_PERIOD_M=m;renderCompChart(m);let el=$('compNewStoppedSect');if(!el||!COMPANY_TRENDS)return;let periods=COMPANY_TRENDS.periods||[],companies=COMPANY_TRENDS.companies||[];let{newItems,stoppedItems}=filterPeriodItems(periods,companies,m);let mLabel=m===1?'1 oy':m+' oy';el.innerHTML=`<div style="display:flex;gap:8px;align-items:center;margin:4px 0 12px;flex-wrap:wrap"><span style="font-size:12px;color:#557086;font-weight:700">Hisobot davri:</span>${periodBtns('renderCompNewStopped',m)}</div><div class="grid2">${companyTrendTable(periods,stoppedItems,'🚫 So\'nggi '+mLabel+' ichida faoliyatini to\'xtatgan korxonalar','Barcha korxonalar faol','stopped')}${companyTrendTable(periods,newItems,'🆕 So\'nggi '+mLabel+' ichida birinchi marta yuk olib kelgan korxonalar','Yangi korxona aniqlanmadi','new')}</div>`}
function renderGoodsNewStopped(m){GOODS_PERIOD_M=m;renderGoodsChart(m);let el=$('goodsNewStoppedSect');if(!el||!GOODS_TRENDS)return;let periods=GOODS_TRENDS.periods||[],items=GOODS_TRENDS.goods||[];let{newItems,stoppedItems}=filterPeriodItems(periods,items,m);let mLabel=m===1?'1 oy':m+' oy';el.innerHTML=`<div style="display:flex;gap:8px;align-items:center;margin:4px 0 12px;flex-wrap:wrap"><span style="font-size:12px;color:#557086;font-weight:700">Hisobot davri:</span>${periodBtns('renderGoodsNewStopped',m)}</div><div class="grid2">${goodsTrendTable(periods,stoppedItems,'🚫 So\'nggi '+mLabel+' ichida IBKdan chiqib ketgan tovarlar','Barcha tovarlar faol','stopped')}${goodsTrendTable(periods,newItems,'🆕 So\'nggi '+mLabel+' ichida yangi qo\'shilgan tovarlar','Yangi tovar aniqlanmadi','new')}</div>`}
async function buildOmborOborot(){let base=$('omborOborotBase').value,final=$('omborOborotFinal').value,box=$('omborOborotResult');if(!box)return;box.innerHTML='<div class=muted>Hisoblanmoqda...</div>';let j=await loadReleaseData(base,final,box);if(!j)return;box.innerHTML=warehouseTurnoverPanel(j)}
function companyTrendTable(periods,rows,title,emptyMsg,mode){if(!rows.length)return `<div class="panel"><h3>${esc(title)} (0)</h3><div class="muted">${esc(emptyMsg)}</div></div>`;let lastIdx=periods.length-1;let dateHead=mode==="new"?"Birinchi faol davr":"So'nggi faol davr";let valHead=mode==="new"?"Joriy qiymat (ming $)":"So'nggi qiymat (ming $)";let wHead=mode==="new"?"Joriy vazn (tn)":"So'nggi vazn (tn)";let body=rows.slice(0,150).map(c=>{let refIdx,dateLabel;if(mode==="new"){refIdx=lastIdx;let fi=firstActiveIdx(c.series);dateLabel=fi>=0?periods[fi]:"-"}else{refIdx=lastActiveIdx(c.series);dateLabel=refIdx>=0?periods[refIdx]:"-"}let s=c.series[refIdx>=0?refIdx:lastIdx]||{value:0,weight:0};return `<tr style="cursor:pointer" onclick="showCompanyTrend('${c.stir}')"><td class=text>${esc(c.company||"-")}</td><td>${esc(c.stir)}</td><td>${esc(dateLabel)}</td><td class=num>${fmtN(s.value)}</td><td class=num>${fmtN(s.weight)}</td></tr>`}).join("");return `<div class="panel"><h3>${esc(title)} (${rows.length})</h3><table class="fixed-table"><thead><tr><th>Korxona</th><th>STIR</th><th>${dateHead}</th><th>${valHead}</th><th>${wHead}</th></tr></thead><tbody>${body}</tbody></table></div>`}
function showCompKpiList(type){let b=window.COMP_BUCKETS||{},companies=window.COMP_ALL_COMPANIES||[],periods=(COMPANY_TRENDS&&COMPANY_TRENDS.periods)||[];let lastIdx=periods.length-1;let rows,title;if(type==='all'){rows=companies;title='Jami korxonalar ('+companies.length+')';}else if(type==='active'){rows=b.active||[];title='So\'nggi davrda faol ('+rows.length+')';}else if(type==='stopped'){rows=b.stoppedAll||[];title='Butunlay to\'xtatgan korxonalar ('+rows.length+')';}else if(type==='new'){rows=b.newAll||[];title='Yangi qo\'shilgan korxonalar ('+rows.length+')';}else return;dlgTitle.textContent=title;dlgBody.innerHTML=table([{k:"company",t:"Korxona nomi",w:"50%"},{k:"stir",t:"STIR",w:"110px"},{k:"last_value",t:"So\'nggi qiymat (ming $)",n:1,f:fmtN},{k:"last_weight",t:"Vazn (tn)",n:1,f:fmtN}],rows.slice(0,300).map(c=>{let s=c.series[lastIdx]||{value:0,weight:0};return {key:{stir:c.stir},company:c.company||c.stir,stir:c.stir,last_value:s.value,last_weight:s.weight}}),"fixed-table");dlg.showModal()}
async function loadCompanyTrends(){let el=$("trendPanel");if(!el)return;try{let j=await api("/api/company_trends");COMPANY_TRENDS=j;if(j.error&&!j.periods?.length){el.innerHTML=`<h3>Korxonalar bo'yicha davriy tendensiya</h3><div class="muted">Server xatoligi: ${esc(j.error)}</div>`;return}let periods=j.periods||[],companies=j.companies||[];if(!periods.length){el.innerHTML=`<h3>Korxonalar bo'yicha davriy tendensiya</h3><div class="muted">Hali tarixiy ma'lumot yetarli emas. Har bir yangi yuklangan fayl bilan tendensiya shakllanadi.</div>`;return}let b=trendBuckets(periods,companies);window.COMP_BUCKETS=b;window.COMP_ALL_COMPANIES=companies;el.innerHTML=`<h3>Korxonalar bo'yicha davriy tendensiya</h3><div class="overview-note">${periodRangeCaption(periods)}.</div><div class="kpis"><div class="kpi" onclick="showCompKpiList('all')" style="cursor:pointer" title="Barcha korxonalar ro\'yxatini ko\'rish"><span>Jami korxonalar</span><b>${fmtI(companies.length)}</b></div><div class="kpi" onclick="showCompKpiList('active')" style="cursor:pointer" title="Faol korxonalar ro\'yxati"><span>So'nggi davrda faol</span><b>${fmtI(b.active.length)}</b></div><div class="kpi" onclick="showCompKpiList('stopped')" style="cursor:pointer" title="To\'xtatgan korxonalar ro\'yxati"><span>Butunlay to'xtagan</span><b>${fmtI(b.stoppedAll.length)}</b></div><div class="kpi" onclick="showCompKpiList('new')" style="cursor:pointer" title="Yangi korxonalar ro\'yxati"><span>Yangi qo'shilgan</span><b>${fmtI(b.newAll.length)}</b></div></div><div id="compChartSect"></div><div id="compNewStoppedSect"></div>`;renderCompChart(COMP_PERIOD_M);renderCompNewStopped(COMP_PERIOD_M)}catch(e){el.innerHTML=`<h3>Korxonalar bo'yicha davriy tendensiya</h3><div class="muted">Ma'lumotni yuklab bo'lmadi: ${esc(e.message||e)}</div>`}}
function showCompanyTrend(stir){let c=(COMPANY_TRENDS.companies||[]).find(x=>x.stir===stir);renderCompNewStopped(COMP_PERIOD_M);if(!c)return;let periods=COMPANY_TRENDS.periods||[];let values=c.series.map(s=>s.value),weights=c.series.map(s=>s.weight);dlgTitle.textContent=`${c.company||c.stir} - davriy tendensiya (STIR: ${c.stir})`;let rows=c.series.map((s,i)=>({date:periods[i],value:s.value,weight:s.weight,partiya:s.partiya,key:{}}));dlgBody.innerHTML=`<div class="overview-note">${periodRangeCaption(periods)}.</div>`+svgLineChart(periods,[{label:"Qiymat (ming $)",values},{label:"Vazn (tn)",values:weights}])+table([{k:"date",t:"Davr",w:"110px"},{k:"value",t:"Qiymat (ming $)",n:1,f:fmtN},{k:"weight",t:"Vazn (tn)",n:1,f:fmtN},{k:"partiya",t:"Partiya",n:1,f:fmtI}],rows,"fixed-table");dlg.showModal()}
function goodsTrendTable(periods,rows,title,emptyMsg,mode){if(!rows.length)return `<div class="panel"><h3>${esc(title)} (0)</h3><div class="muted">${esc(emptyMsg)}</div></div>`;let lastIdx=periods.length-1;let dateHead=mode==="new"?"Birinchi faol davr":"So'nggi faol davr";let valHead=mode==="new"?"Joriy qiymat (ming $)":"So'nggi qiymat (ming $)";let wHead=mode==="new"?"Joriy vazn (tn)":"So'nggi vazn (tn)";let body=rows.slice(0,150).map(g=>{let refIdx,dateLabel;if(mode==="new"){refIdx=lastIdx;let fi=firstActiveIdx(g.series);dateLabel=fi>=0?periods[fi]:"-"}else{refIdx=lastActiveIdx(g.series);dateLabel=refIdx>=0?periods[refIdx]:"-"}let s=g.series[refIdx>=0?refIdx:lastIdx]||{value:0,weight:0};return `<tr style="cursor:pointer" title="${esc(g.goods||"-")}" onclick="showGoodsTrend('${g.hs_code}')"><td class=text>${esc(g.goods_short||g.hs_code||"-")}</td><td>${esc(g.hs_code)}</td><td>${esc(dateLabel)}</td><td class=num>${fmtN(s.value)}</td><td class=num>${fmtN(s.weight)}</td></tr>`}).join("");return `<div class="panel"><h3>${esc(title)} (${rows.length})</h3><table class="fixed-table"><thead><tr><th>Tovar</th><th>TIF TN kodi</th><th>${dateHead}</th><th>${valHead}</th><th>${wHead}</th></tr></thead><tbody>${body}</tbody></table></div>`}
async function loadGoodsTrends(){let el=$("goodsTrendPanel");if(!el)return;try{let j=await api("/api/goods_trends");GOODS_TRENDS=j;if(j.error&&!j.periods?.length){el.innerHTML=`<h3>Tovarlar bo'yicha davriy tendensiya</h3><div class="muted">Server xatoligi: ${esc(j.error)}</div>`;return}let periods=j.periods||[],items=j.goods||[];if(!periods.length){el.innerHTML=`<h3>Tovarlar bo'yicha davriy tendensiya</h3><div class="muted">Hali tarixiy ma'lumot yetarli emas. Har bir yangi yuklangan fayl bilan tendensiya shakllanadi.</div>`;return}let b=trendBuckets(periods,items);el.innerHTML=`<h3>Tovarlar bo'yicha davriy tendensiya</h3><div class="overview-note">${periodRangeCaption(periods)}. Har bir tovar TIF TN (HS) kodi bo'yicha guruhlangan.</div><div class="kpis"><div class="kpi"><span>Jami tovar turlari</span><b>${fmtI(items.length)}</b></div><div class="kpi"><span>So'nggi davrda faol</span><b>${fmtI(b.active.length)}</b></div><div class="kpi"><span>Chiqib ketgan</span><b>${fmtI(b.stoppedAll.length)}</b></div><div class="kpi"><span>Yangi qo'shilgan</span><b>${fmtI(b.newAll.length)}</b></div></div><div id="goodsChartSect"></div><div id="goodsNewStoppedSect"></div>`;renderGoodsChart(GOODS_PERIOD_M);renderGoodsNewStopped(GOODS_PERIOD_M)}catch(e){el.innerHTML=`<h3>Tovarlar bo'yicha davriy tendensiya</h3><div class="muted">Ma'lumotni yuklab bo'lmadi: ${esc(e.message||e)}</div>`}}
function showGoodsTrend(hs){let g=(GOODS_TRENDS.goods||[]).find(x=>x.hs_code===hs);renderGoodsNewStopped(GOODS_PERIOD_M);if(!g)return;let periods=GOODS_TRENDS.periods||[];let values=g.series.map(s=>s.value),weights=g.series.map(s=>s.weight);dlgTitle.textContent=`${g.goods||g.hs_code} - davriy tendensiya (TIF TN: ${g.hs_code})`;let rows=g.series.map((s,i)=>({date:periods[i],value:s.value,weight:s.weight,partiya:s.partiya,key:{}}));dlgBody.innerHTML=`<div class="overview-note">${periodRangeCaption(periods)}.</div>`+svgLineChart(periods,[{label:"Qiymat (ming $)",values},{label:"Vazn (tn)",values:weights}])+table([{k:"date",t:"Davr",w:"110px"},{k:"value",t:"Qiymat (ming $)",n:1,f:fmtN},{k:"weight",t:"Vazn (tn)",n:1,f:fmtN},{k:"partiya",t:"Partiya",n:1,f:fmtI}],rows,"fixed-table");dlg.showModal()}

let WAREHOUSE_TRENDS={},TRANSPORT_TRENDS={},TRANSPORT_CO_TRENDS={};
let WAREHOUSE_PERIOD_M=6,TRANSPORT_PERIOD_M=3;
function renderWarehouseChart(m){WAREHOUSE_PERIOD_M=m;let el=$('whChartSect');if(!el||!WAREHOUSE_TRENDS?.periods?.length)return;let periods=WAREHOUSE_TRENDS.periods,warehouses=WAREHOUSE_TRENDS.warehouses;let fp=filterPeriodArr(periods,m);let wi=fp.map(p=>periods.indexOf(p));let totalValues=wi.map(i=>warehouses.reduce((a,w)=>a+(w.series[i]?.value||0),0));let totalWeights=wi.map(i=>warehouses.reduce((a,w)=>a+(w.series[i]?.weight||0),0));el.innerHTML=`<div style="display:flex;gap:8px;margin:4px 0 8px;flex-wrap:wrap"><span style="font-size:12px;color:#557086;font-weight:700">Hisobot davri:</span>${periodBtns('renderWarehouseChart',m)}</div><div class="panel"><h3>Barcha omborlar bo'yicha yuk oqimi (${fp[0]||''} — ${fp[fp.length-1]||''})</h3>${svgLineChart(fp,[{label:"Qiymat (ming $)",values:totalValues},{label:"Vazn (tn)",values:totalWeights}])}</div>`}
function renderTransportChart(m){TRANSPORT_PERIOD_M=m;let el=$('trChartSect');if(!el||!TRANSPORT_TRENDS?.periods?.length)return;let periods=TRANSPORT_TRENDS.periods,transports=TRANSPORT_TRENDS.transports;let fp=filterPeriodArr(periods,m);let wi=fp.map(p=>periods.indexOf(p));let colors=["#1d72b8","#16a34a","#f59e0b","#dc2626","#7c3aed","#0891b2","#ea580c"];let allSeries=transports.map((t,i)=>({label:t.transport,values:wi.map(idx=>t.series[idx]?.value||0),color:colors[i%colors.length]}));el.innerHTML=`<div style="display:flex;gap:8px;margin:4px 0 8px;flex-wrap:wrap"><span style="font-size:12px;color:#557086;font-weight:700">Hisobot davri:</span>${periodBtns('renderTransportChart',m)}</div><div class="panel"><h3>Transport turlari bo'yicha qiymat tendensiyasi (${fp[0]||''} — ${fp[fp.length-1]||''})</h3>${svgLineChart(fp,allSeries)}</div>`}
async function loadWarehouseTrends(){let el=$("warehouseTrendPanel");if(!el)return;try{let j=await api("/api/warehouse_trends");WAREHOUSE_TRENDS=j;if(j.error&&!j.periods?.length){el.innerHTML=`<h3>Omborlar tendensiyasi</h3><div class="muted">Server xatoligi: ${esc(j.error)}</div>`;return}let periods=j.periods||[],warehouses=j.warehouses||[];if(!periods.length){el.innerHTML=`<h3>Omborlar bo'yicha yuk oqimi tendensiyasi</h3><div class="muted">Hali tarixiy ma'lumot yetarli emas. Har bir yangi yuklangan fayl bilan tendensiya shakllanadi.</div>`;return}let growing=warehouses.filter(w=>{let s=w.series;if(s.length<2)return false;return s[s.length-1].value>s[s.length-2].value&&s[s.length-2].value>0});let shrinking=warehouses.filter(w=>{let s=w.series;if(s.length<2)return false;return s[s.length-1].value<s[s.length-2].value&&s[s.length-1].value>0});let whRows=warehouses.map(w=>{let s=w.series;let last=s[s.length-1]||{},prev=s.length>1?s[s.length-2]:{};let delta=(last.value||0)-(prev.value||0);let trend=delta>0?"📈 o'sish":delta<0?"📉 pasayish":"➡ barqaror";return {warehouse:w.warehouse,last_value:last.value||0,last_weight:last.weight||0,last_partiya:last.partiya||0,trend}});el.innerHTML=`<h3>Omborlar bo'yicha yuk oqimi tendensiyasi</h3><div class="overview-note">${periodRangeCaption(periods)}. Har qaysi ombor qatoriga bosish orqali alohida tendensiya grafiği ochiladi.</div><div class="kpis"><div class="kpi"><span>Jami omborlar</span><b>${fmtI(warehouses.length)}</b></div><div class="kpi"><span>So'nggi davrda faol</span><b>${fmtI(warehouses.filter(w=>w.series[w.series.length-1]?.value>0).length)}</b></div><div class="kpi"><span>Yuk oqimi o'sgan</span><b style="color:#16a34a">${fmtI(growing.length)}</b></div><div class="kpi"><span>Yuk oqimi kamaygan</span><b style="color:#dc2626">${fmtI(shrinking.length)}</b></div></div><div id="whChartSect"></div><div class="panel wide"><h3>Omborlar kesimida so'nggi holat</h3><table class="fixed-table"><colgroup><col style="width:42%"><col style="width:17%"><col style="width:15%"><col style="width:12%"><col style="width:14%"></colgroup><thead><tr><th class="text">Ombor</th><th class="num">Qiymat (ming $)</th><th class="num">Vazn (tn)</th><th class="num">Partiya</th><th class="text">Tendensiya</th></tr></thead><tbody>${whRows.map(r=>`<tr onclick="showWarehouseTrend(${JSON.stringify(r.warehouse)})" style="cursor:pointer"><td class="text">${esc(r.warehouse)}</td><td class="num">${fmtN(r.last_value)}</td><td class="num">${fmtN(r.last_weight)}</td><td class="num">${fmtI(r.last_partiya)}</td><td class="text">${r.trend}</td></tr>`).join("")}</tbody></table></div>`;renderWarehouseChart(WAREHOUSE_PERIOD_M)}catch(e){let el2=$("warehouseTrendPanel");if(el2)el2.innerHTML=`<h3>Omborlar tendensiyasi</h3><div class="muted">Ma'lumotni yuklab bo'lmadi: ${esc(e.message||e)}</div>`}}
function showWarehouseTrend(whName){let w=(WAREHOUSE_TRENDS.warehouses||[]).find(x=>x.warehouse===whName);if(!w)return;let periods=WAREHOUSE_TRENDS.periods||[];let values=w.series.map(s=>s.value),weights=w.series.map(s=>s.weight);dlgTitle.textContent=`${whName} — yuk oqimi tendensiyasi`;let rows=w.series.map((s,i)=>({date:periods[i],value:s.value,weight:s.weight,partiya:s.partiya,key:{}}));dlgBody.innerHTML=`<div class="overview-note">${periodRangeCaption(periods)}.</div>`+svgLineChart(periods,[{label:"Qiymat (ming $)",values},{label:"Vazn (tn)",values:weights}])+table([{k:"date",t:"Davr",w:"110px"},{k:"value",t:"Qiymat (ming $)",n:1,f:fmtN},{k:"weight",t:"Vazn (tn)",n:1,f:fmtN},{k:"partiya",t:"Partiya",n:1,f:fmtI}],rows,"fixed-table");dlg.showModal()}
async function loadTransportTrends(){let el=$("transportTrendPanel");if(!el)return;try{let j=await api("/api/transport_trends");TRANSPORT_TRENDS=j;if(j.error&&!j.periods?.length){el.innerHTML=`<h3>Transport turi tendensiyasi</h3><div class="muted">Server xatoligi: ${esc(j.error)}</div>`;return}let periods=j.periods||[],transports=j.transports||[];if(!periods.length){el.innerHTML=`<h3>Transport turi bo'yicha davriy tendensiya</h3><div class="muted">Hali tarixiy ma'lumot yetarli emas. Har bir yangi yuklangan fayl bilan tendensiya shakllanadi.</div>`;return}let colors=["#1d72b8","#16a34a","#f59e0b","#dc2626","#7c3aed","#0891b2","#ea580c"];el.innerHTML=`<h3>Transport turi bo'yicha davriy tendensiya</h3><div class="overview-note">${periodRangeCaption(periods)}. Har qaysi transport turi bo'yicha qiymat (ming $) tendensiyasi.</div><div class="kpis">${transports.slice(0,4).map(t=>{let last=t.series[t.series.length-1]||{};return `<div class="kpi"><span>${esc(t.transport)}</span><b>${fmtN(last.value||0)}</b></div>`}).join("")}</div><div id="trChartSect"></div><div class="grid2">${transports.map((t,i)=>{let values=t.series.map(s=>s.value),weights=t.series.map(s=>s.weight);return `<div class="panel"><h3 onclick="showTransportTrend(${JSON.stringify(t.transport)})" style="cursor:pointer;color:${colors[i%colors.length]}">${esc(t.transport)}</h3>${svgLineChart(periods,[{label:"Qiymat",values,color:colors[i%colors.length]},{label:"Vazn (tn)",values:weights}])}</div>`}).join("")}</div>`;renderTransportChart(TRANSPORT_PERIOD_M)}catch(e){let el2=$("transportTrendPanel");if(el2)el2.innerHTML=`<h3>Transport turi tendensiyasi</h3><div class="muted">Ma'lumotni yuklab bo'lmadi: ${esc(e.message||e)}</div>`}}
function showTransportTrend(trName){let t=(TRANSPORT_TRENDS.transports||[]).find(x=>x.transport===trName);if(!t)return;let periods=TRANSPORT_TRENDS.periods||[];let values=t.series.map(s=>s.value),weights=t.series.map(s=>s.weight);dlgTitle.textContent=`${trName} — davriy tendensiya`;let rows=t.series.map((s,i)=>({date:periods[i],value:s.value,weight:s.weight,partiya:s.partiya,key:{}}));dlgBody.innerHTML=`<div class="overview-note">${periodRangeCaption(periods)}.</div>`+svgLineChart(periods,[{label:"Qiymat (ming $)",values},{label:"Vazn (tn)",values:weights}])+table([{k:"date",t:"Davr",w:"110px"},{k:"value",t:"Qiymat (ming $)",n:1,f:fmtN},{k:"weight",t:"Vazn (tn)",n:1,f:fmtN},{k:"partiya",t:"Partiya",n:1,f:fmtI}],rows,"fixed-table");dlg.showModal()}
let TRANSPORT_CO_PERIOD_M=3;
function renderTransportCoTable(m){TRANSPORT_CO_PERIOD_M=m;let el=$('trCoSect');if(!el||!TRANSPORT_CO_TRENDS?.periods?.length)return;let periods=TRANSPORT_CO_TRENDS.periods,transports=TRANSPORT_CO_TRENDS.transports;let fp=filterPeriodArr(periods,m);let wi=fp.map(p=>periods.indexOf(p));let colors=["#1d72b8","#16a34a","#f59e0b","#dc2626","#7c3aed","#0891b2","#ea580c"];el.innerHTML=`<div style="display:flex;gap:8px;margin:4px 0 8px;flex-wrap:wrap"><span style="font-size:12px;color:#557086;font-weight:700">Hisobot davri:</span>${periodBtns('renderTransportCoTable',m)}</div>${transports.map((t,ti)=>{let cols=[{k:"company",t:"Korxona",w:"45%"},{k:"stir",t:"STIR",w:"90px"},{k:"period_value",t:"Davr qiymat (ming $)",n:1,f:fmtN},{k:"last_value",t:"So'nggi davr (ming $)",n:1,f:fmtN},{k:"trend",t:"Tendensiya",w:"80px"}];let rows=t.companies.map(c=>{let pVal=wi.reduce((a,i)=>a+(c.series[i]?.value||0),0);let last=c.series[c.series.length-1]||{},prev=c.series.length>1?c.series[c.series.length-2]:{};let delta=(last.value||0)-(prev.value||0);return {key:{stir:c.stir},company:c.company||c.stir,stir:c.stir,period_value:pVal,last_value:last.value||0,trend:delta>0?"📈":delta<0?"📉":"➡"}}).filter(r=>r.period_value>0).sort((a,b)=>b.period_value-a.period_value).slice(0,15);let totalV=rows.reduce((a,r)=>a+r.period_value,0);return rows.length?`<div class="panel wide"><h3 style="color:${colors[ti%colors.length]}">${esc(t.transport)} — TOP korxonalar (tanlangan davr: ${fmtN(totalV)} ming $)</h3>${table(cols,rows,"fixed-table")}</div>`:''}).join('')}`}
async function loadTransportCompanyTrends(){let el=$("transportCompanyPanel");if(!el)return;try{let j=await api("/api/transport_company_trends");TRANSPORT_CO_TRENDS=j;if(j.error&&!j.periods?.length){el.innerHTML=`<h3>Transport kesimida korxonalar</h3><div class="muted">Server xatoligi: ${esc(j.error)}</div>`;return}let periods=j.periods||[],transports=j.transports||[];if(!periods.length){el.innerHTML=`<h3>Transport kesimida korxonalar</h3><div class="muted">Hali tarixiy ma'lumot yetarli emas.</div>`;return}el.innerHTML=`<h3>Transport turlari kesimida korxonalar</h3><div class="overview-note">${periodRangeCaption(periods)}. Tanlangan davrda har bir transport turi bo'yicha eng faol korxonalar.</div><div id="trCoSect"></div>`;renderTransportCoTable(TRANSPORT_CO_PERIOD_M)}catch(e){let el2=$("transportCompanyPanel");if(el2)el2.innerHTML=`<h3>Transport kesimida korxonalar</h3><div class="muted">Ma'lumotni yuklab bo'lmadi: ${esc(e.message||e)}</div>`}}
function topNWithOther(rows,n,sortKey,labelKey){
  rows=by(rows||[],sortKey);let head=rows.slice(0,n), tail=rows.slice(n);
  if(tail.length){let other={key:{},name:"Boshqa tovarlar",korxona:"Boshqa tovarlar"};other[labelKey||"name"]="Boshqa tovarlar";["partiya","korxona","vazn","qiymat","tolov"].forEach(k=>other[k]=tail.reduce((a,r)=>a+(+r[k]||0),0));head.push(other)}
  return head;
}
function countryRows(){let rows=DATA.countries||DATA.country||[];return rows}
function chartBlock(title,rows,label,value,fmt){return `<div class=panel><h2>${title}</h2>${bars(rows,label,value,fmt)}${miniChart(rows,value,label)}</div>`}
function overviewPanels(){let topCompanies=by(DATA.top_value||[],"qiymat").slice(0,30), topGoods=topNWithOther(DATA.goods||[],30,"partiya","name");return `<div class="grid2">${executiveSummary()}<div class=panel><h2>70-74-80: postlar kesimida</h2>${table([{k:"post",t:"Post",w:"42%"},{k:"partiya",t:"Partiya",n:1,f:fmtI},{k:"vazn",t:"Vazn (tn)",n:1,f:fmtN},{k:"qiymat",t:"Qiymat (ming $)",n:1,f:fmtN},{k:"tolov",t:"Kutilayotgan to'lov (mln so'm)",n:1,f:fmtN}],basicTotal(DATA.all_post_summary||[],"IBK bo'yicha Jami","post"),"fixed-table")}<div class=overview-note>Post qatorlari ustiga bosilganda asos deklaratsiyalar ochiladi.</div></div><div class="panel wide"><h2>Jami muddati o'tgan: postlar va rejimlar kesimida</h2>${expiredTotalExcelTable()}</div><div class="panel wide"><h2>Qiymat bo'yicha TOP 30 korxona</h2>${bars(topCompanies,"korxona","qiymat",fmtN)}</div>${chartBlock("Muddati o'tgan partiyalar (postlar kesimida)",DATA.post_summary||[],"post","partiya",fmtI)}${chartBlock("Tovar guruhlari partiya bo'yicha ulushi",topGoods,"name","partiya",fmtI)}<div class=panel><h2>Davlatlar bo'yicha tahlil</h2>${bars(countryRows(),"name","qiymat",fmtN)}</div></div>`}
function releaseAnalytics(rows){rows=rows||[];let byCompany={};rows.forEach(r=>{let k=r.stir||r.korxona||r.company||"";if(!byCompany[k])byCompany[k]={korxona:r.korxona||r.company||"",stir:r.stir||"",partiya:0,qiymat:0,vazn:0,count:0};byCompany[k].partiya+=+r.released_partiya||0;byCompany[k].qiymat+=+r.released_qiymat||0;byCompany[k].vazn+=+r.released_vazn||0;byCompany[k].count++});let list=Object.values(byCompany);let cols=[{k:"korxona",t:"Korxona nomi",w:"42%"},{k:"stir",t:"STIR",w:"12%"},{k:"partiya",t:"Partiya",n:1,f:fmtI},{k:"qiymat",t:"Qiymat (ming $)",n:1,f:fmtN},{k:"vazn",t:"Vazn (tn)",n:1,f:fmtN}];return `<div class=grid2><div class=panel><h2>Eng ko'p nazoratdan yechgan korxonalar</h2>${table(cols,basicTotal(by(list,"qiymat").slice(0,20),"IBK bo'yicha Jami","korxona"),"fixed-table")}</div><div class=panel><h2>Eng tez nazoratdan yechadigan korxonalar</h2><div class=muted>To'liq tezlik tahlili bir nechta sanalar arxivga yuklangandan keyin avtomatik shakllanadi.</div>${table(cols,by(list,"partiya").slice(0,10),"fixed-table")}</div><div class=panel><h2>Eng sekin nazoratdan yechadigan korxonalar</h2><div class=muted>Yakuniy davrda qoldiq ko'p qolgan korxonalar alohida hisoblanadi.</div>${table(cols,by(list,"vazn").slice(-10).reverse(),"fixed-table")}</div></div>`}
async function buildRelease(){let base=$("relBase").value, final=$("relFinal").value, j=await api(`/api/release?base=${encodeURIComponent(base)}&final=${encodeURIComponent(final)}`);if(j.missing){$("releaseResult").innerHTML=`<div class=muted>Arxivda yo'q sana: ${j.missing.join(", ")}. Shu davr uchun asos fayl yuklash kerak.</div>`;return}let raw=(j.rows||[]).map(r=>Object.assign({korxona:r.korxona||r.company||"",partiya:r.released_partiya,qiymat:r.released_qiymat,vazn:r.released_vazn},r));let rows=numbered([Object.assign({korxona:"IBK bo'yicha Jami",stir:"",decl:"",regime:"",partiya:(j.total||{}).released_partiya,qiymat:(j.total||{}).released_qiymat,vazn:(j.total||{}).released_vazn},j.total||{})].concat(raw));let cols=[{k:"rn",t:"T/r",n:1,f:v=>v,w:"38px"},{k:"korxona",t:"Korxona nomi",w:"280px"},{k:"stir",t:"STIR",w:"90px"},{k:"decl",t:"Deklaratsiya",w:"145px"},{k:"regime",t:"Rejim",w:"70px"},{k:"partiya",t:"Partiya",n:1,f:fmtI,w:"80px"},{k:"qiymat",t:"Qiymat (ming $)",n:1,f:fmtN,w:"120px"},{k:"vazn",t:"Vazn (tn)",n:1,f:fmtN,w:"110px"}];$("releaseResult").innerHTML=`<h2>${base} - ${final} <a class="btn light" href="/api/export?kind=release&base=${base}&final=${final}&token=${TOKEN}">Excel</a></h2><div class=overview-note>Boshlang'ich davr: ${base}. Yakuniy davr: ${final}.</div>${table(cols,rows,"release-table fixed-table")}${releaseAnalytics(raw)}${miniChart(raw,"qiymat","korxona")}`}
function startIdleTimer(){let t;function reset(){clearTimeout(t);t=setTimeout(()=>{logout();alert("20 daqiqa foydalanilmagani uchun profil avtomatik yopildi.")},1200000)}["click","keydown","mousemove","scroll","touchstart"].forEach(e=>document.addEventListener(e,reset,{passive:true}));reset()}startIdleTimer();
const oldRender=render;render=function(){oldRender();let a=$("actions");if(a){a.innerHTML=a.innerHTML.replaceAll("Jamlanma Excel","Ombor ma'lumot").replaceAll("Nazoratdan yechilgan","Nazoratdan yechish");a.insertAdjacentHTML("afterbegin",`<button class="light lang-btn" onclick="setLang('uz')">O'zb</button><button class="light lang-btn" onclick="setLang('uzc')">&#1038;&#1079;&#1073;</button><button class="light lang-btn" onclick="setLang('ru')">&#1056;&#1091;&#1089;</button><button class="light lang-btn" onclick="document.body.classList.toggle('dark')">&#9680;</button> `)}}
function parseUzDate(s){let m=String(s||"").match(/(\d{2})\.(\d{2})\.(\d{4})/);return m?new Date(+m[3],+m[2]-1,+m[1]).getTime():0}
topNWithOther=function(rows,n,sortKey,labelKey){rows=by(rows||[],sortKey);let head=rows.slice(0,n),tail=rows.slice(n);if(tail.length){let other={key:{},name:"Boshqalar",korxona:"Boshqalar"};other[labelKey||"name"]="Boshqalar";["partiya","vazn","qiymat","tolov"].forEach(k=>other[k]=tail.reduce((a,r)=>a+(+r[k]||0),0));head.push(other)}return head}
countryRows=function(){return topNWithOther(DATA.countries||DATA.country||[],30,"qiymat","name")}
bars=function(rows,label,value,fmt){rows=(rows||[]).filter(r=>(+r[value]||0)>0);let sum=rows.reduce((a,r)=>a+(+r[value]||0),0)||1;return `<div class=bars>${rows.map((r,i)=>{let pct=(+r[value]||0)/sum*100;return `<div class=barrow style="animation-delay:${i*.03}s"><div title="${esc(r[label])}">${esc(r[label])}</div><div class=bar><span style="width:${Math.max(6,pct)}%">${pct.toFixed(1)}%</span></div><b>${fmt(r[value])}</b></div>`}).join("")}</div>`}
miniChart=function(rows,k,label){rows=by(rows||[],k).slice(0,8);let sum=rows.reduce((a,r)=>a+(+r[k]||0),0)||1, pts=rows.map((r,i)=>[18+i*(260/Math.max(1,rows.length-1)),140-(+r[k]||0)/Math.max(...rows.map(x=>+x[k]||0),1)*104]);let path=pts.map((p,i)=>(i?'L':'M')+p[0].toFixed(1)+' '+p[1].toFixed(1)).join(' '), area=pts.length?`M18 150 ${path.slice(1)} L278 150 Z`:"";let p1=((+rows[0]?.[k]||0)/sum*100).toFixed(0)+"%";return `<div class="viz-card"><div class="donut" style="--p:${p1}" data-label="${p1}"></div><svg class="trend-svg" viewBox="0 0 300 170"><defs><linearGradient id="trendGrad" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#4ecdc4"/><stop offset="1" stop-color="#ffffff" stop-opacity="0"/></linearGradient></defs><path class="trend-area" d="${area}"></path><path class="trend-line" d="${path}"></path>${pts.map(p=>`<circle class="trend-dot" cx="${p[0]}" cy="${p[1]}" r="4"></circle>`).join("")}</svg></div>`}
chartBlock=function(title,rows,label,value,fmt){return `<div class=panel><h2>${title}</h2>${miniChart(rows,value,label)}${bars(rows,label,value,fmt)}</div>`}
executiveSummary=function(){if(!DATA)return "";let k=DATA.kpis||{},top=(DATA.top_value||[])[0]||{},dep=(DATA.top_deposit||[])[0]||{},own=(DATA.warehouse||[]).find(r=>(r.name||"")==="O'z ombor")||{},exp=by(DATA.post_summary||[],"partiya")[0]||{};let items=[["Umumiy nazorat",`${fmtI(k.partiya)} partiya, ${fmtN(k.qiymat)} ming $ qiymat.`],["Muddati o'tgan",`${fmtI(k.expired)} partiya. Eng ko'p partiya qayd etilgan post: ${esc(exp.post||'-')}.`],["Eng yirik korxona",`${esc(top.korxona||'-')} - ${fmtN(top.qiymat||0)} ming $.`],["O'z ombor",`${fmtI(own.partiya||0)} partiya, ${fmtN(own.qiymat||0)} ming $.`],["Depozit yetakchisi",`${esc(dep.korxona||'-')} - ${fmtN(dep.depozit||0)} mln so'm.`]];return `<div class="panel exec-summary"><h2>Rahbar uchun qisqa xulosa</h2><div class="summary-grid">${items.map(x=>`<div class="summary-item"><b>${x[0]}</b><span>${x[1]}</span></div>`).join("")}</div></div>`}
uploadPanel=function(){return `<div class=stack><div class=panel><h2>Fayl yuklash</h2><div class=muted>BNRTE, To'lovlar va yil davomida yig'ilgan asos fayllarni shu yerdan yuklaysiz.</div></div><div class=panel><h2>BNRTE jamlanma</h2><form class="upload" id="upload"><div><label>Asos fayl</label><input name="source" type="file" accept=".xls,.xlsx,.html,.htm" required></div><div><label>Depozit fayl</label><input name="deposit" type="file" accept=".xlsx"></div><div></div><button>BNRTE yuklash</button></form></div><div class=panel><h2>Yillik arxivni birdan yuklash</h2><form class="upload" id="bulkUpload"><div><label>Asos fayllar</label><input name="sources" type="file" accept=".xls,.xlsx,.html,.htm" multiple required></div><div><label>Depozit fayl ixtiyoriy</label><input name="deposit" type="file" accept=".xlsx"></div><div></div><button>Hammasini yuklash</button></form><div id=bulkResult class=muted>Fayllar sanasi nomidan olinadi va arxivga qo'shiladi.</div></div><div class=panel><h2>To'lovlar jadvallari</h2><form class="upload" id="tolovUpload"><div><label>To'lov baza fayli</label><input name="source" type="file" accept=".xlsx,.xls" required></div><div class=muted>04.06+07.06.2026 kabi asos fayl yuklanadi.</div><div></div><button>To'lov jadvallarini shakllantirish</button></form><div id="tolovUploadResult" class=muted>Natijada Excel fayllar shakllanadi.</div></div></div>`}
bindUpload=function(){let f=$("upload");if(f)f.onsubmit=async e=>{e.preventDefault();$("status").textContent="BNRTE fayllari yuklanyapti...";let j=await api("/api/reports",{method:"POST",body:new FormData(f)});poll(j.job_id)};let bf=$("bulkUpload");if(bf)bf.onsubmit=async e=>{e.preventDefault();$("status").textContent="Yillik asos fayllar yuklanyapti...";let j=await api("/api/reports_bulk",{method:"POST",body:new FormData(bf)});$("bulkResult").textContent=`${j.count||0} ta fayl navbatga qo'shildi. Arxiv sanalari yangilanadi.`;$("status").textContent="Bulk yuklash boshlandi";await loadArchive();};let tf=$("tolovUpload");if(tf)tf.onsubmit=async e=>{e.preventDefault();$("status").textContent="To'lovlar shakllantirilyapti...";let j=await api("/api/tolov",{method:"POST",body:new FormData(tf)});PAYMENTS=j.payments||[];$("tolovUploadResult").innerHTML=`Tayyor: ${fmtI(PAYMENTS.reduce((a,r)=>a+(+r.rows||0),0))} qator, ${fmtN(PAYMENTS.reduce((a,r)=>a+(+r.sum||0),0))} so'm.`;$("status").textContent="To'lovlar tayyor"}}
buildRelease=async function(){let base=$("relBase").value,final=$("relFinal").value;if(parseUzDate(base)>=parseUzDate(final)){$("releaseResult").innerHTML=`<div class=muted>Boshlang'ich sana yakuniy sanadan oldin bo'lishi kerak.</div>`;return}let j=await api(`/api/release?base=${encodeURIComponent(base)}&final=${encodeURIComponent(final)}`);if(j.missing){$("releaseResult").innerHTML=`<div class=muted>Arxivda yo'q sana: ${j.missing.join(", ")}. Shu davr uchun asos fayl yuklash kerak.</div>`;return}let raw=(j.rows||[]).map(r=>Object.assign({korxona:r.korxona||r.company||"",partiya:r.released_partiya,qiymat:r.released_qiymat,vazn:r.released_vazn},r));let rows=numbered([Object.assign({korxona:"IBK bo'yicha Jami",stir:"",decl:"",regime:"",partiya:(j.total||{}).released_partiya,qiymat:(j.total||{}).released_qiymat,vazn:(j.total||{}).released_vazn},j.total||{})].concat(raw));let cols=[{k:"rn",t:"T/r",n:1,f:v=>v,w:"38px"},{k:"korxona",t:"Korxona nomi",w:"280px"},{k:"stir",t:"STIR",w:"90px"},{k:"decl",t:"Deklaratsiya",w:"145px"},{k:"regime",t:"Rejim",w:"70px"},{k:"partiya",t:"Partiya",n:1,f:fmtI,w:"80px"},{k:"qiymat",t:"Qiymat (ming $)",n:1,f:fmtN,w:"120px"},{k:"vazn",t:"Vazn (tn)",n:1,f:fmtN,w:"110px"}];$("releaseResult").innerHTML=`<h2>${base} - ${final} <a class="btn light" href="/api/export?kind=release&base=${base}&final=${final}&token=${TOKEN}">Excel</a></h2>${table(cols,rows,"release-table fixed-table")}${releaseAnalytics(raw)}${miniChart(raw,"qiymat","korxona")}`}
detail=async function(key){if(!DATA||!key||Object.keys(key).length===0)return;if(key.view==="expired_inline"){let el=$("expiredInline");if(el)el.innerHTML=expiredTotalExcelTable();return}if(key.view==="regime_posts"){let rows=(DATA.regime_posts||{})[key.regime]||[];dlgTitle.textContent=`${key.regime} - postlar kesimida`;dlgBody.innerHTML=table([{k:"post",t:"Post",w:"42%"},{k:"partiya",t:"Partiya",n:1,f:fmtI},{k:"vazn",t:"Vazn (tn)",n:1,f:fmtN},{k:"qiymat",t:"Qiymat (ming $)",n:1,f:fmtN},{k:"tolov",t:"To'lov (mln so'm)",n:1,f:fmtN}],basicTotal(rows,"IBK bo'yicha Jami","post"));dlg.showModal();return}let filterText=JSON.stringify(key),q=new URLSearchParams({report:DATA.id,filters:filterText}),j=await api("/api/details?"+q);dlgTitle.textContent="Asos deklaratsiyalar";dlgBody.innerHTML=`<p><a class="btn light" href="/api/export_details?report=${DATA.id}&filters=${encodeURIComponent(filterText)}&token=${TOKEN}">Excelga yuklash</a></p>`+table([{k:"decl",t:"Deklaratsiya"},{k:"date",t:"Sana"},{k:"regime",t:"Rejim"},{k:"stir",t:"STIR"},{k:"company",t:"Korxona"},{k:"hs",t:"TIF TN"},{k:"goods",t:"Tovar"},{k:"partiya",t:"Partiya",n:1,f:fmtI},{k:"vazn",t:"Vazn (tn)",n:1,f:fmtN},{k:"qiymat",t:"Qiymat (ming $)",n:1,f:fmtN}],j.rows,"fixed-table details-wide");dlg.showModal()}
const renderFinal=render;render=function(){renderFinal();let tabs=$("tabs");if(tabs){let buttons=[...tabs.querySelectorAll("button")], rej=buttons.find(b=>b.textContent.trim()==="Rejimlar"), exp=buttons.find(b=>b.textContent.includes("Muddati"));if(rej&&exp&&exp.nextSibling!==rej)exp.after(rej)}let a=$("actions");if(a&&DATA){let f=DATA.files||{};if((f.pngs||[]).length>1&&!a.innerHTML.includes("_pngs"))a.insertAdjacentHTML("afterbegin",`<a class=btn href="/download/${DATA.id}/_pngs?token=${TOKEN}">Barcha PNG</a> `)}}
const apiRaw=api;api=async function(url,opt={}){opt.headers=Object.assign({"X-Token":TOKEN},opt.headers||{});let r=await fetch(url,opt);let j=await r.json().catch(()=>({error:`HTTP ${r.status}`}));if(r.status===401){showLogin();throw Error("login")}if(!r.ok||j.error)throw Error(j.error||`HTTP ${r.status}`);return j}
function countryPoint(name,i,total){let s=String(name||"").toLowerCase();let map=[["xitoy",70,126],["china",70,126],["rossiya",210,70],["russia",210,70],["turkiya",165,155],["turkey",165,155],["koreya",105,95],["germaniya",135,115],["germany",135,115],["italiya",125,148],["fransiya",118,130],["aqsh",48,88],["usa",48,88],["hindiston",98,182],["india",98,182],["qozog",238,116],["kazakh",238,116],["baa",146,205],["amirlik",146,205],["eron",170,185],["iran",170,185]];for(let m of map){if(s.includes(m[0]))return {x:m[1],y:m[2]}}let angle=(-120+(i/(Math.max(total-1,1))*240))*Math.PI/180;return {x:155+112*Math.cos(angle),y:150+82*Math.sin(angle)}}
const COUNTRY_COORDS=[["xitoy",35.86,104.19],["china",35.86,104.19],["rossiya",61.52,105.32],["russia",61.52,105.32],["turkiya",38.96,35.24],["turkey",38.96,35.24],["koreya",36.5,127.8],["germaniya",51.17,10.45],["germany",51.17,10.45],["italiya",41.87,12.57],["italy",41.87,12.57],["fransiya",46.23,2.21],["france",46.23,2.21],["aqsh",37.09,-95.71],["usa",37.09,-95.71],["amerika",37.09,-95.71],["hindiston",20.59,78.96],["india",20.59,78.96],["qozog",48.02,66.92],["kazakh",48.02,66.92],["baa",23.42,53.85],["amirlik",23.42,53.85],["eron",32.43,53.69],["iran",32.43,53.69],["polsha",51.92,19.15],["ispaniya",40.46,-3.75],["yaponiya",36.2,138.25],["japan",36.2,138.25],["belarus",53.7,27.95],["latviya",56.88,24.6],["litva",55.17,23.88],["ukraina",48.38,31.17],["misr",26.82,30.8],["egypt",26.82,30.8],["pokiston",30.38,69.35],["pakistan",30.38,69.35],["vietnam",14.06,108.28],["malayziya",4.21,101.98],["tailand",15.87,100.99],["indoneziya",-0.79,113.92],["braziliya",-14.24,-51.93],["meksika",23.63,-102.55],["saudiya",23.89,45.08],["gruziya",41.69,44.03],["georgia",41.69,44.03],["ozarbayjon",40.14,47.58],["azerbaijan",40.14,47.58],["armaniston",40.07,45.04],["armenia",40.07,45.04],["qirg'iziston",41.2,74.77],["kyrgyz",41.2,74.77],["tojikiston",38.86,71.28],["tajik",38.86,71.28],["turkmaniston",38.97,59.56],["turkmen",38.97,59.56],["afg'oniston",33.93,67.71],["afghan",33.93,67.71],["niderlandiya",52.13,5.29],["netherlands",52.13,5.29],["belgiya",50.50,4.47],["avstriya",47.52,14.55],["shveytsariya",46.82,8.23],["chexiya",49.82,15.47],["vengriya",47.16,19.50],["ruminiya",45.94,24.97],["bolgariya",42.73,25.49],["serbiya",44.02,21.01],["xorvatiya",45.10,15.20],["slovakiya",48.67,19.70],["sloveniya",46.15,14.99],["estoniya",58.60,25.01],["finlyandiya",61.92,25.75],["shvetsiya",60.13,18.64],["norvegiya",60.47,8.47],["daniya",56.26,9.50],["gollandiya",52.13,5.29],["avstraly",-25.27,133.78],["australia",-25.27,133.78],["kanada",56.13,-106.35],["canada",56.13,-106.35],["isroil",31.05,34.85],["israel",31.05,34.85],["iordaniya",30.59,36.24],["livan",33.85,35.86],["falastin",31.95,35.23],["suriya",34.80,38.99],["iroq",33.22,43.68],["iraq",33.22,43.68],["quvayt",29.31,47.48],["bahrayn",26.07,50.55],["qatar",25.35,51.18],["ummon",21.51,55.92],["yaman",15.55,48.52],["efiopiya",9.14,40.49],["marokash",31.79,-7.09],["tunis",33.89,9.54],["jazoir",28.03,1.66],["nigeria",9.08,8.68],["gana",7.95,-1.02],["keniya",-0.02,37.91],["tanzaniya",-6.37,34.89],["britaniya",55.38,-3.44],["buyuk brit",55.38,-3.44],["united king",55.38,-3.44],["england",51.50,-0.12],["scotland",56.49,-4.20],["wales",52.13,-3.78],["ireland",53.41,-8.24],["irlandiya",53.41,-8.24],["irlandia",53.41,-8.24],["portugal",39.40,-8.22],["portugaliya",39.40,-8.22],["gretsiya",37.98,23.73],["greece",37.98,23.73],["portugal",39.40,-8.22],["singapur",1.35,103.82],["singapore",1.35,103.82],["gonkong",22.32,114.17],["hong kong",22.32,114.17],["tayvan",23.70,121.00],["taiwan",23.70,121.00],["nz",40.90,174.89],["yangi zelend",40.90,174.89],["peru",-9.19,-75.02],["chili",-35.68,-71.54],["argentina",-38.42,-63.62],["kolumbiya",4.57,-74.30],["venetsuela",6.42,-66.59]];
function countryLatLon(name){let s=String(name||"").toLowerCase();for(let c of COUNTRY_COORDS){if(s.includes(c[0]))return {lat:c[1],lon:c[2]}}return null}
const COUNTRY_FLAGS={"xitoy":"🇨🇳","china":"🇨🇳","rossiya":"🇷🇺","russia":"🇷🇺","turkiya":"🇹🇷","turkey":"🇹🇷","koreya":"🇰🇷","germaniya":"🇩🇪","germany":"🇩🇪","italiya":"🇮🇹","fransiya":"🇫🇷","aqsh":"🇺🇸","usa":"🇺🇸","amerika":"🇺🇸","hindiston":"🇮🇳","india":"🇮🇳","qozog":"🇰🇿","kazakh":"🇰🇿","baa":"🇦🇪","amirlik":"🇦🇪","eron":"🇮🇷","iran":"🇮🇷","polsha":"🇵🇱","ispaniya":"🇪🇸","yaponiya":"🇯🇵","japan":"🇯🇵","belarus":"🇧🇾","latviya":"🇱🇻","litva":"🇱🇹","ukraina":"🇺🇦","misr":"🇪🇬","egypt":"🇪🇬","pokiston":"🇵🇰","pakistan":"🇵🇰","vietnam":"🇻🇳","malayziya":"🇲🇾","tailand":"🇹🇭","indoneziya":"🇮🇩","braziliya":"🇧🇷","meksika":"🇲🇽","saudiya":"🇸🇦","gretsiya":"🇬🇷","gruziya":"🇬🇪","britaniya":"🇬🇧"};
function countryFlag(name){let s=String(name||"").toLowerCase();for(let k in COUNTRY_FLAGS){if(s.includes(k))return COUNTRY_FLAGS[k]}return "🌐"}
let _CFM_ROWS=[];let COUNTRY_FLOW_MAP=null;let _YMAP_LOADED=false;let _YMAP_CBS=[];
function _loadYMaps(cb){
  if(typeof ymaps3!=='undefined'){ymaps3.ready.then(cb);return}
  _YMAP_CBS.push(cb);
  if(_YMAP_LOADED)return;
  _YMAP_LOADED=true;
  let s=document.createElement('script');
  s.src='https://api-maps.yandex.ru/v3/?apikey=0b5695bd-89cd-4d0e-8945-11b2472402ba&lang=ru_RU';
  s.onload=function(){ymaps3.ready.then(function(){let cbs=_YMAP_CBS;_YMAP_CBS=[];cbs.forEach(fn=>fn())})};
  s.onerror=function(){let el=document.getElementById('cfmMap');if(el)el.innerHTML='<div style="height:100%;display:flex;align-items:center;justify-content:center;color:#888;font-size:13px">Yandex Maps yuklanmadi</div>'};
  document.head.appendChild(s);
}
function countryFlowMap(rows){
  _CFM_ROWS=by(rows||[],"qiymat").slice(0,15);
  return `<div style="border-radius:12px;overflow:hidden;border:1px solid var(--line)"><div id="cfmMap" style="height:480px;width:100%"><div style="height:100%;display:flex;align-items:center;justify-content:center;color:#888;font-size:13px;background:#ddeeff">Xarita yuklanmoqda...</div></div></div><div class="globe-caption">Davlatlardan O\'zbekistonga yo\'nalgan yuk oqimi</div>`
}
function initCountryFlowMap(){
  let el=document.getElementById("cfmMap");if(!el)return;
  let rows=_CFM_ROWS;if(!rows.length)return;
  _loadYMaps(function(){_buildCFM(el,rows)});
}
function _buildCFM(el,rows){
  try{
    if(COUNTRY_FLOW_MAP){try{COUNTRY_FLOW_MAP.destroy()}catch(e){}COUNTRY_FLOW_MAP=null}
    el.innerHTML='';
    const {YMap,YMapDefaultSchemeLayer,YMapDefaultFeaturesLayer,YMapMarker,YMapFeature}=ymaps3;
    let map=new YMap(el,{location:{center:[58,41],zoom:2},behaviors:['drag','scrollZoom','pinchZoom']});
    map.addChild(new YMapDefaultSchemeLayer());
    map.addChild(new YMapDefaultFeaturesLayer());
    COUNTRY_FLOW_MAP=map;
    let UZ=[69.3,41.3];
    let uzEl=document.createElement('div');
    uzEl.innerHTML='<div style="background:#1d72b8;color:#fff;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:900;white-space:nowrap;box-shadow:0 2px 6px rgba(0,0,0,.4);transform:translate(-50%,-50%)">🇺🇿 O\'zbekiston</div>';
    map.addChild(new YMapMarker({coordinates:UZ},uzEl));
    let maxQ=Math.max(1,...rows.map(r=>+r.qiymat||0));
    let maxV=Math.max(1,...rows.map(r=>+r.vazn||0));
    let allCoords=[];
    rows.forEach(r=>{
      let ll=countryLatLon(r.name);if(!ll)return;
      let q=+r.qiymat||0,v=+r.vazn||0,p=+r.partiya||0;
      let qS=q/maxQ;
      let clr=qS>=.6?'#16a34a':qS>=.3?'#d97706':'#1d72b8';
      let coord=[ll.lon,ll.lat];
      allCoords.push(coord);
      map.addChild(new YMapFeature({id:'cfmline_'+r.name,geometry:{type:'LineString',coordinates:[coord,UZ]},style:{stroke:[{color:clr,width:1.5+qS*2.5,opacity:.72}]}}));
      let sz=Math.round(Math.max(30,Math.min(54,30+Math.sqrt(v/maxV)*24)));
      let fs=sz>44?12:sz>36?11:10;
      let flag=countryFlag(r.name);let nm=esc(r.name).slice(0,16);
      let mEl=document.createElement('div');
      mEl.style.cssText='transform:translate(-50%,-50%);display:flex;flex-direction:column;align-items:center;gap:2px;cursor:pointer';
      mEl.title=`${flag} ${esc(r.name)}\nPartiya: ${fmtI(p)}\nQiymat: ${fmtN(q)} ming $\nVazn: ${fmtN(v)} tn`;
      mEl.innerHTML=`<div style="width:${sz}px;height:${sz}px;border-radius:50%;background:${clr};border:2.5px solid #fff;display:flex;align-items:center;justify-content:center;font-size:${fs}px;font-weight:800;color:#fff;box-shadow:0 2px 8px rgba(0,0,0,.3);box-sizing:border-box">${fmtI(p)}</div><div style="white-space:nowrap;background:rgba(255,255,255,.93);padding:1px 5px;border-radius:3px;font-size:10px;font-weight:700;color:#1a3a5c;box-shadow:0 1px 3px rgba(0,0,0,.2)">${flag} ${nm}</div>`;
      map.addChild(new YMapMarker({coordinates:coord},mEl));
    });
    if(allCoords.length){
      let lons=allCoords.map(c=>c[0]).concat([UZ[0]]),lats=allCoords.map(c=>c[1]).concat([UZ[1]]),pad=4;
      map.setLocation({bounds:[[Math.min(...lons)-pad,Math.min(...lats)-pad],[Math.max(...lons)+pad,Math.max(...lats)+pad]],duration:400});
    }
    let lgEl=document.createElement('div');
    lgEl.className='cfm-legend';
    lgEl.style.cssText='position:absolute;bottom:12px;right:12px;z-index:10';
    lgEl.innerHTML='<b>Rang — qiymat ulushi:</b><br><span style="color:#16a34a">● Yuqori (60%+)</span><br><span style="color:#d97706">● O\'rta (30-60%)</span><br><span style="color:#1d72b8">● Quyi (&lt;30%)</span><hr style="margin:4px 0"><b>Doira:</b> partiya / vazn';
    el.style.position='relative';
    el.appendChild(lgEl);
  }catch(err){console.error('CFM map error:',err)}
}
const POST_NAMES={"35001":"Nukus aeroporti chegara bojxona posti","35002":"Nukus TIF bojxona posti","35003":"Xo'jayli chegara bojxona posti","35004":"Dovut-ota chegara bojxona posti","35010":"Qoraqalpog'iston temir yo'l chegara bojxona posti","03002":"Do'stlik chegara bojxona posti","03003":"Andijon aeroporti chegara bojxona posti","03005":"Mingtepa chegara bojxona posti","03006":"Qorasuv chegara bojxona posti","03007":"Xonobod chegara bojxona posti","03008":"Pushmon chegara bojxona posti","03009":"Madaniyat chegara bojxona posti","03011":"Andijon TIF bojxona posti","03013":"Keskanyol chegara bojxona posti","03014":"Savay temir yo'l chegara bojxona posti","03015":"Asaka TIF bojxona posti","06001":"Buxoro aeroporti chegara bojxona posti","06006":"Buxoro TIF bojxona posti","06009":"Qorakol TIF bojxona posti","06010":"Olot chegara bojxona posti","06011":"Xo'jadavlat temir yo'l chegara bojxona posti","08003":"Uchturgon chegara bojxona posti","08004":"Jizzax TIF bojxona posti","08007":"Qo'shkent chegara bojxona posti","10002":"Nasaf TIF bojxona posti","10007":"Qamashi-G'uzor TIF bojxona posti","10008":"Qarshi-Kerki chegara bojxona posti","10012":"Qarshi aeroporti chegara bojxona posti","12002":"Navoiy aeroporti chegara bojxona posti","12003":"Navoiy TIF bojxona posti","12008":"Zarafshon TIF bojxona posti","14002":"Namangan aeroporti chegara bojxona posti","14003":"Uchqo'rg'on chegara bojxona posti","14004":"Kosonsoy chegara bojxona posti","14005":"Pop chegara bojxona posti","14010":"Namangan TIF bojxona posti","18001":"Samarqand aeroporti chegara bojxona posti","18002":"Jartepa chegara bojxona posti","18005":"Samarqand TIF bojxona posti","18007":"Ulug'bek TIF bojxona posti","22002":"Termiz aeroporti chegara bojxona posti","22003":"Sariosiyo chegara bojxona posti","22004":"Sariosiyo temir yo'l chegara bojxona posti","22005":"Termiz TIF bojxona posti","22006":"Denov TIF bojxona posti","22007":"Gulbahor chegara bojxona posti","22011":"Daryo porti chegara bojxona posti","22015":"Boldir temir yo'l chegara bojxona posti","22017":"Ayritom chegara bojxona posti","22022":"Termiz xalqaro savdo markazi TIF bojxona posti","24002":"Xovosobod chegara bojxona posti","24004":"Sirdaryo chegara bojxona posti","24006":"Oq oltin chegara bojxona posti","24009":"Guliston TIF bojxona posti","24014":"Malik chegara bojxona posti","27001":"Yallama chegara bojxona posti","27008":"Navoiy chegara bojxona posti","27009":"S. Najimov chegara bojxona posti","27011":"Oybek chegara bojxona posti","27013":"Bekobod avto chegara bojxona posti","27014":"Chirchiq TIF bojxona posti","27015":"Olmaliq TIF bojxona posti","27016":"Yangiyol TIF bojxona posti","27019":"Nazarbek TIF bojxona posti","27020":"Keles TIF bojxona posti","27021":"G'ishtkoprik chegara bojxona posti","27023":"Farhod chegara bojxona posti","27024":"Bekobod temir yo'l chegara bojxona posti","27028":"Angren TIF bojxona posti","30001":"Farg'ona aeroporti chegara bojxona posti","30002":"Qo'qon TIF bojxona posti","30004":"Farg'ona chegara bojxona posti","30005":"Andarxon chegara bojxona posti","30006":"Rishton chegara bojxona posti","30008":"Rovot chegara bojxona posti","30009":"Vodiy TIF bojxona posti","30010":"O'zbekiston chegara bojxona posti","30012":"So'x chegara bojxona posti","33001":"Shovot chegara bojxona posti","33004":"Do'stlik chegara bojxona posti","33007":"Urganch TIF bojxona posti","33011":"Urganch aeroporti chegara bojxona posti","33033":"Shovot chegaraoldi savdo zonasi","26002":"Toshkent-tovar TIF bojxona posti","26003":"Ark buloq TIF bojxona posti","26004":"Chuqursoy TIF bojxona posti","26009":"Keles temir yo'l chegara bojxona posti","26010":"Sirg'ali TIF bojxona posti","26013":"Chuqursoy texnik idora temir yo'l chegara bojxona posti","00101":"Toshkent xalqaro aeroporti CHBP","00102":"Avia yuklar TIF bojxona posti","00107":"Elektron tijorat TIF bojxona posti","00110":"Toshkent-Humo aeroporti CHBP"};
function namedSourcePosts(){return (DATA.source_posts||[]).map(r=>{let nm=r.post_nomi;if(!nm||nm===r.post_kodi||/^\d{5}$/.test(String(nm||""))){nm=POST_NAMES[r.post_kodi]||(r.post_kodi?"Post №"+r.post_kodi:"-")}return Object.assign({},r,{post_nomi:nm})})}
function sourcePostInfographics(){let posts=by(namedSourcePosts(),"qiymat").slice(0,12),trans=DATA.transport||[],tp=trans.reduce((a,r)=>a+(+r.qiymat||0),0)||1,pp=posts.reduce((a,r)=>a+(+r.qiymat||0),0)||1;let rings=trans.map((r,i)=>{let pct=Math.round((+r.qiymat||0)/tp*100),dash=Math.max(0,Math.min(100,pct));return `<div class=transport-ring style="--p:${dash};--delay:${i*.12}s" onclick='detail(${JSON.stringify(r.key||{transport:r.name})})'><b>${esc(r.name||"-")}</b><span>${pct}%</span><small>${fmtI(r.partiya||0)} partiya В· ${fmtN(r.qiymat||0)} ming $</small></div>`}).join("");let barsHtml=posts.map((r,i)=>{let pct=Math.max(2,(+r.qiymat||0)/pp*100);return `<div class=flow-row onclick='detail(${JSON.stringify(r.key||{})})' title="${esc(r.post_nomi||r.post_kodi||"-")}"><div class=flow-name><b>${esc(r.post_kodi||"-")}</b><span>${esc(r.post_nomi||"-")}</span></div><div class=flow-track><i style="width:${pct.toFixed(1)}%;animation-delay:${i*.05}s"></i><em>${pct.toFixed(1)}%</em></div><div class=flow-num>${fmtN(r.qiymat||0)}</div></div>`}).join("");return `<div class="panel wide"><h2>Nazoratga qo'yilgan postlar va transport turlari</h2><div class=transport-viz><div class=ring-grid>${rings}</div><div class=flow-list>${barsHtml}</div></div></div>`}
function transportPanel(){let cols=[{k:"post_kodi",t:"Post kodi",w:"78px"},{k:"post_nomi",t:"Post nomi",w:"260px"},{k:"transport",t:"Transport turi",w:"92px"},{k:"partiya",t:"Partiya",w:"78px",n:1,f:fmtI},{k:"vazn",t:"Vazn (tn)",w:"92px",n:1,f:fmtN},{k:"qiymat",t:"Qiymat (ming $)",w:"112px",n:1,f:fmtN},{k:"tolov",t:"To'lov (mln so'm)",w:"112px",n:1,f:fmtN},{k:"korxona",t:"Korxona",w:"82px",n:1,f:fmtI}];let rows=namedSourcePosts();return `<div class=panel><h2>Deklaratsiya post kodi bo'yicha tahlil</h2>${table(cols,basicTotal(rows,"IBK bo'yicha Jami","post_nomi"),"fixed-table transport-table")}</div>${sourcePostInfographics()}${chartBlock("Transport turi bo'yicha ulushi",DATA.transport||[],"name","qiymat",fmtN)}`}
const overviewPanelsWithMap=overviewPanels;overviewPanels=function(){let html=overviewPanelsWithMap();let countries=countryRows();let countryBlock=`<div class="panel wide"><h2>Davlatlar bo'yicha yo'nalishlar</h2>${countryFlowMap(countries)}<div class="chart-under-globe">${bars(countries,"name","qiymat",fmtN)}</div></div>${transportPanel()}`;return html.replace(`<div class=panel><h2>Davlatlar bo'yicha tahlil</h2>${bars(countryRows(),"name","qiymat",fmtN)}</div>`,countryBlock)}
function flightsPanelShell(){
  return `<div class="panel wide" id="flightsPanelWrap">
<h2>Toshkent xalqaro aeroporti — jonli parvozlar</h2>
<div style="border-radius:14px;overflow:hidden;border:1px solid var(--line);margin-bottom:8px">
  <div id="flightsMap" style="height:520px;width:100%">
    <div style="height:100%;display:flex;align-items:center;justify-content:center;color:#888;font-size:13px;background:#ddeeff">Yandex xarita yuklanmoqda...</div>
  </div>
</div>
<div id="flightsMeta" class="muted" style="font-size:12px;margin-top:4px">Jonli parvozlar, real vaqtda yangilanadi · Toshkent aeroporti (TAS) · manba: OpenSky Network</div>
</div>`}
let FLIGHTS_MAP=null,FLIGHTS_MARKERS=[],FLIGHTS_TIMER=null;
function initFlightsMap(){
  let el=document.getElementById("flightsMap");if(!el)return;
  _loadYMaps(function(){
    if(FLIGHTS_MAP){try{FLIGHTS_MAP.destroy()}catch(e){}FLIGHTS_MAP=null;FLIGHTS_MARKERS=[]}
    el.innerHTML='';
    const {YMap,YMapDefaultSchemeLayer,YMapDefaultFeaturesLayer,YMapMarker}=ymaps3;
    let map=new YMap(el,{location:{center:[69.27,41.3],zoom:6},behaviors:['drag','scrollZoom','pinchZoom']});
    map.addChild(new YMapDefaultSchemeLayer());
    map.addChild(new YMapDefaultFeaturesLayer());
    FLIGHTS_MAP=map;
    let tasEl=document.createElement('div');
    tasEl.innerHTML='<div style="background:#1d72b8;color:#fff;padding:2px 6px;border-radius:4px;font-size:10px;font-weight:700;white-space:nowrap;transform:translate(-50%,-50%);cursor:default" title="Toshkent xalqaro aeroporti (TAS)">TAS</div>';
    map.addChild(new YMapMarker({coordinates:[69.2401,41.2995]},tasEl));
    refreshFlightsMap();
    if(FLIGHTS_TIMER)clearInterval(FLIGHTS_TIMER);
    FLIGHTS_TIMER=setInterval(refreshFlightsMap,30000);
  });
}
async function refreshFlightsMap(){
  if(!FLIGHTS_MAP||typeof ymaps3==='undefined')return;
  let meta=document.getElementById("flightsMeta");
  try{
    let j=await api("/api/flights_live");
    FLIGHTS_MARKERS.forEach(m=>{try{FLIGHTS_MAP.removeChild(m)}catch(e){}});
    FLIGHTS_MARKERS=[];
    const {YMapMarker}=ymaps3;
    (j.planes||[]).forEach(p=>{
      let h=+(p.heading||0);
      let mEl=document.createElement('div');
      mEl.style.cssText='transform:translate(-50%,-50%)';
      mEl.title=esc(p.callsign||p.icao24||'-')+'\nBalandlik: '+Math.round(p.alt||0)+' m\nTezlik: '+Math.round((p.speed||0)*3.6)+' km/soat\nDavlat: '+esc(p.country||'-');
      mEl.innerHTML='<div style="transform:rotate('+h+'deg);font-size:18px;line-height:1;cursor:pointer;text-shadow:0 1px 3px rgba(0,0,0,.5)">&#9992;</div>';
      let marker=new YMapMarker({coordinates:[p.lon,p.lat]},mEl);
      FLIGHTS_MAP.addChild(marker);
      FLIGHTS_MARKERS.push(marker);
    });
    if(meta)meta.textContent=j.error?'Xatolik: '+j.error:'Jonli parvozlar: '+(j.planes||[]).length+' ta \xb7 yangilangan: '+new Date((j.updated||Date.now()/1000)*1000).toLocaleTimeString("uz-UZ",{timeZone:"Asia/Tashkent"});
  }catch(e){if(meta)meta.textContent="Ma'lumot yuklanmadi"}
}
detail=async function(key){if(!DATA||!key||Object.keys(key).length===0)return;if(key.view==="expired_inline"){let el=$("expiredInline");if(el)el.innerHTML=expiredTotalExcelTable();return}if(key.view==="regime_posts"){let rows=(DATA.regime_posts||{})[key.regime]||[];dlgTitle.textContent=`${key.regime} - postlar kesimida`;dlgBody.innerHTML=table([{k:"post",t:"Post",w:"42%"},{k:"partiya",t:"Partiya",n:1,f:fmtI},{k:"vazn",t:"Vazn (tn)",n:1,f:fmtN},{k:"qiymat",t:"Qiymat (ming $)",n:1,f:fmtN},{k:"tolov",t:"To'lov (mln so'm)",n:1,f:fmtN}],basicTotal(rows,"IBK bo'yicha Jami","post"));dlg.showModal();return}let filterText=JSON.stringify(key),q=new URLSearchParams({report:DATA.id,filters:filterText}),j=await api("/api/details?"+q);dlgTitle.textContent="Asos deklaratsiyalar";dlgBody.innerHTML=`<p><a class="btn light" href="/api/export_details?report=${DATA.id}&filters=${encodeURIComponent(filterText)}&token=${TOKEN}">Excelga yuklash</a></p>`+table([{k:"decl",t:"Deklaratsiya",w:"160px"},{k:"source_post",t:"Boshlang'ich post kodi",w:"90px"},{k:"source_post_name",t:"Boshlang'ich post nomi",w:"220px"},{k:"transport",t:"Transport",w:"92px"},{k:"date",t:"Sana",w:"86px"},{k:"regime",t:"Rejim",w:"64px"},{k:"post",t:"Nazorat posti",w:"170px"},{k:"stir",t:"STIR",w:"105px"},{k:"company",t:"Korxona",w:"220px"},{k:"hs",t:"TIF TN",w:"105px"},{k:"goods",t:"Tovar",w:"240px"},{k:"partiya",t:"Partiya",w:"70px",n:1,f:fmtI},{k:"vazn",t:"Vazn (tn)",w:"90px",n:1,f:fmtN},{k:"qiymat",t:"Qiymat (ming $)",w:"105px",n:1,f:fmtN}],j.rows,"fixed-table details-wide");dlg.showModal()}
function cleanTopActions(){let a=$("actions");if(!a)return;[...a.querySelectorAll("button")].forEach(b=>{if(["Sozlamalar","Nastroyki","Settings"].includes(b.textContent.trim()))b.remove()});[...a.querySelectorAll("a")].forEach(x=>{let t=x.textContent.trim();if(/^PNG\s+\d+$/i.test(t)||t==="Barcha PNG")x.remove()});if(DATA){let f=DATA.files||{};let has=a.querySelector("[data-pngzip]");if(!has&&(f.pngs||[]).length){let href=(f.pngs||[]).length>1?`/download/${DATA.id}/_pngs?token=${TOKEN}`:`/download/${DATA.id}/${f.pngs[0]}?token=${TOKEN}`;a.insertAdjacentHTML("afterbegin",`<a data-pngzip class=btn href="${href}">PNG</a> `)}}}
function translateRuPage(){if(LANG!=="ru")return;let pairs=[["Kirish","Р’С…РѕРґ"],["Chiqish","Р’С‹С…РѕРґ"],["Ombor ma'lumot","РЎРєР»Р°РґСЃРєР°СЏ СЃРІРѕРґРєР°"],["Fayl yuklash","Р—Р°РіСЂСѓР·РєР° С„Р°Р№Р»РѕРІ"],["Umumiy","РћР±С‰РµРµ"],["Korxonalar","РџСЂРµРґРїСЂРёСЏС‚РёСЏ"],["Muddati o'tgan","РџСЂРѕСЃСЂРѕС‡РµРЅРЅС‹Рµ"],["Nazoratdan yechish","РЎРЅСЏС‚РёРµ СЃ РєРѕРЅС‚СЂРѕР»СЏ"],["Tovarlar","РўРѕРІР°СЂС‹"],["Omborlar","РЎРєР»Р°РґС‹"],["Rejimlar","Р РµР¶РёРјС‹"],["Oziq-ovqat","РџСЂРѕРґРѕРІРѕР»СЊСЃС‚РІРёРµ"],["Muddatlar","РЎСЂРѕРєРё"],["Boshqaruv","РЈРїСЂР°РІР»РµРЅРёРµ"],["Arxiv","РђСЂС…РёРІ"],["To'lovlar","РџР»Р°С‚РµР¶Рё"],["Davlatlar bo'yicha yo'nalishlar","РњР°СЂС€СЂСѓС‚С‹ РїРѕ СЃС‚СЂР°РЅР°Рј"],["Deklaratsiya post kodi bo'yicha tahlil","РђРЅР°Р»РёР· РїРѕ РєРѕРґСѓ РїРѕСЃС‚Р° РґРµРєР»Р°СЂР°С†РёРё"],["Transport turi bo'yicha ulushi","Р”РѕР»СЏ РїРѕ РІРёРґСѓ С‚СЂР°РЅСЃРїРѕСЂС‚Р°"],["Qiymat bo'yicha TOP 30 korxona","РўРћРџ-30 РїСЂРµРґРїСЂРёСЏС‚РёР№ РїРѕ СЃС‚РѕРёРјРѕСЃС‚Рё"],["Rahbar uchun qisqa xulosa","РљСЂР°С‚РєР°СЏ СЃРІРѕРґРєР° РґР»СЏ СЂСѓРєРѕРІРѕРґРёС‚РµР»СЏ"]];document.querySelectorAll("button,h2,h1,b,label,.muted,.btn").forEach(el=>{let s=el.childNodes.length===1?el.textContent.trim():"";let p=pairs.find(x=>x[0]===s);if(p)el.textContent=p[1]})}
adminPanel=function(){return `<div class=admin-layout><div class="admin-card"><h2>Hodim qo'shish yoki tahrirlash</h2><form id=userForm class=admin-form><input type=hidden name=edit_mode value=""><label>Login</label><input name=user required><label>Yangi parol</label><input name=password type=password placeholder="Tahrirda bo'sh qoldirish mumkin"><label>F.I.Sh.</label><input name=full_name><label>Lavozim</label><input name=position><label>Telefon</label><input name=phone><label>Post kodi</label><input name=post_code placeholder="00 = IBK, aks holda post kodi"><label>Rol</label><select name=role><option value=foydalanuvchi>Foydalanuvchi</option><option value=inspektor>Inspektor</option><option value=rahbar>Rahbar</option><option value=admin>Admin</option></select><label>Til</label><select name=lang><option value=uz>O'zbek lotin</option><option value=uzc>O'zbek kirill</option><option value=ru>Rus tili</option></select><div class=perm-grid><label><input type=checkbox name=perm_view checked> Ko'rish</label><label><input type=checkbox name=perm_upload> Yuklash</label><label><input type=checkbox name=perm_export> Eksport</label><label><input type=checkbox name=perm_release> Yechish</label></div><div class=excel-actions><button>Saqlash</button><button type=button class=light onclick="resetUserForm()">Tozalash</button></div></form></div><div class="admin-card"><h2>Hodimlar ro'yxati</h2><div id=users></div></div></div>`}
function resetUserForm(){let f=$("userForm");if(!f)return;f.reset();f.edit_mode.value="";f.user.disabled=false;f.perm_view.checked=true}
bindUserForm=function(){let f=$("userForm");if(!f)return;f.onsubmit=async e=>{e.preventDefault();let perms=[];["view","upload","export","release"].forEach(p=>{if(f[`perm_${p}`]?.checked)perms.push(p)});let body={user:f.user.value,pass:f.password.value,full_name:f.full_name.value,position:f.position.value,phone:f.phone.value,post_code:f.post_code.value,role:f.role.value,lang:f.lang.value,perms,enabled:true};await api(f.edit_mode.value?"/api/users/update":"/api/users",{method:"POST",body:JSON.stringify(body)});resetUserForm();loadUsers()}}
window.editUser=function(login){api("/api/users").then(j=>{let u=(j.users||[]).find(x=>x.user===login),f=$("userForm");if(!u||!f)return;f.edit_mode.value="1";f.user.value=u.user;f.user.disabled=true;f.password.value="";f.full_name.value=u.full_name||"";f.position.value=u.position||"";f.phone.value=u.phone||"";f.post_code.value=u.post_code||"";f.role.value=u.role||"foydalanuvchi";f.lang.value=u.lang||"uz";["view","upload","export","release"].forEach(p=>f[`perm_${p}`].checked=(u.perms||[]).includes(p));window.scrollTo({top:0,behavior:"smooth"})})}
window.deleteUser=async function(login){if(!confirm(`${login} hodimini o'chirasizmi?`))return;await api("/api/users/delete",{method:"POST",body:JSON.stringify({user:login})});loadUsers()}
loadUsers=async function(){let box=$("users");if(!box)return;try{let j=await api("/api/users");let rows=(j.users||[]).filter(u=>u.enabled!==false).map(u=>Object.assign({},u,{role_text:roleTitle(u.role),perms_text:(u.perms||[]).map(permTitle).join(", "),actions:`<button class="light" onclick="editUser('${u.user.replaceAll("'","")}')">Tahrirlash</button> <button class="danger" onclick="deleteUser('${u.user.replaceAll("'","")}')">O'chirish</button>`}));box.innerHTML=table([{k:"user",t:"Login",w:"90px"},{k:"full_name",t:"F.I.Sh.",w:"210px"},{k:"role_text",t:"Rol",w:"110px"},{k:"post_code",t:"Post",w:"70px"},{k:"perms_text",t:"Vakolatlar",w:"230px"},{k:"actions",t:"Amal",w:"160px"}],rows,"fixed-table").replaceAll("&lt;button class=&quot;light&quot;","<button class=\"light\"").replaceAll("&lt;button class=&quot;danger&quot;","<button class=\"danger\"").replaceAll("&lt;/button&gt;","</button>").replaceAll("&quot;&gt;","\">")}catch(e){box.innerHTML="Admin vakolati kerak"}}
const bindUploadFinal=bindUpload;bindUpload=function(){bindUploadFinal();let f=$("upload");if(f)f.onsubmit=async e=>{e.preventDefault();try{$("status").textContent="BNRTE fayllari yuklanyapti...";let j=await api("/api/reports",{method:"POST",body:new FormData(f)});poll(j.job_id)}catch(err){$("status").textContent=err.message}};let bf=$("bulkUpload");if(bf)bf.onsubmit=async e=>{e.preventDefault();try{$("status").textContent="Yillik asos fayllar yuklanyapti...";let j=await api("/api/reports_bulk",{method:"POST",body:new FormData(bf)});let skip=(j.skipped||[]).length?` O'tkazib yuborildi: ${j.skipped.map(x=>x.filename).join(", ")}`:"";$("bulkResult").textContent=`${j.count||0} ta fayl navbatga qo'shildi.${skip}`;$("status").textContent="Bulk yuklash boshlandi";await loadArchive()}catch(err){$("status").textContent=err.message}}}
const renderClean=render;render=function(){renderClean();cleanTopActions();if(LANG==="ru")translateRuPage()}
function releaseCompanyRows(data){
  let all=[...(data.rows||[]),...(data.unreleased||[])], byc={};
  all.forEach(r=>{
    let k=r.stir||r.korxona||r.company||"-";
    if(!byc[k])byc[k]={key:{stir:r.stir||""},korxona:r.korxona||r.company||"",stir:r.stir||"",base_vazn:0,base_qiymat:0,base_partiya:0,remain_vazn:0,remain_qiymat:0,remain_partiya:0,released_vazn:0,released_qiymat:0,released_partiya:0};
    ["base_vazn","base_qiymat","base_partiya","remain_vazn","remain_qiymat","remain_partiya","released_vazn","released_qiymat","released_partiya"].forEach(x=>byc[k][x]+=+r[x]||0);
  });
  return by(Object.values(byc).map(r=>Object.assign(r,{released_pct:r.base_qiymat?r.released_qiymat/r.base_qiymat*100:0,remain_pct:r.base_qiymat?r.remain_qiymat/r.base_qiymat*100:0,current_vazn:r.remain_vazn,current_qiymat:r.remain_qiymat,current_partiya:r.remain_partiya})),"base_qiymat");
}
function releaseTotalRow(rows){
  let t={korxona:"IBK bo'yicha Jami",stir:"",base_vazn:0,base_qiymat:0,base_partiya:0,remain_vazn:0,remain_qiymat:0,remain_partiya:0,released_vazn:0,released_qiymat:0,released_partiya:0,current_vazn:0,current_qiymat:0,current_partiya:0};
  rows.forEach(r=>Object.keys(t).forEach(k=>{if(typeof t[k]==="number")t[k]+=+r[k]||0}));
  t.released_pct=t.base_qiymat?t.released_qiymat/t.base_qiymat*100:0;
  t.remain_pct=t.base_qiymat?t.remain_qiymat/t.base_qiymat*100:0;
  return t;
}
function releaseCols(base,final){
  return [
    {k:"rn",t:"№",n:1,f:v=>v,w:"34px"},{k:"korxona",t:"Korxona",w:"305px"},{k:"stir",t:"STIR",w:"88px"},
    {k:"base_vazn",t:`${base} holatiga qoldiq vazn (tn)`,n:1,f:fmtN,w:"96px"},{k:"base_qiymat",t:`${base} holatiga qoldiq qiymat (ming $)`,n:1,f:fmtN,w:"104px"},{k:"base_partiya",t:`${base} holatiga qoldiq partiya`,n:1,f:fmtI,w:"72px"},
    {k:"remain_vazn",t:`${final} holatiga ${base}dan qoldiq vazn (tn)`,n:1,f:fmtN,w:"96px"},{k:"remain_qiymat",t:`${final} holatiga ${base}dan qoldiq qiymat (ming $)`,n:1,f:fmtN,w:"104px"},{k:"remain_partiya",t:`${final} holatiga ${base}dan qoldiq partiya`,n:1,f:fmtI,w:"72px"},
    {k:"released_vazn",t:"Yechilishi vazn (tn)",n:1,f:fmtN,w:"88px"},{k:"released_qiymat",t:"Yechilishi qiymat (ming $)",n:1,f:fmtN,w:"96px"},{k:"released_pct",t:"Yechilishi %da",n:1,f:fmtN,w:"68px"},{k:"released_partiya",t:"Yechilishi partiya",n:1,f:fmtI,w:"72px"},
    {k:"current_vazn",t:`${final} holatiga qoldiq vazn (tn)`,n:1,f:fmtN,w:"96px"},{k:"current_qiymat",t:`${final} holatiga qoldiq qiymat (ming $)`,n:1,f:fmtN,w:"104px"},{k:"current_partiya",t:`${final} holatiga qoldiq partiya`,n:1,f:fmtI,w:"72px"}
  ];
}
function releaseSpeedPanels(rows,base,final){
  rows=rows||[];
  let days=Math.max(1,Math.round((parseUzDate(final)-parseUzDate(base))/86400000));
  let fast=rows.filter(r=>r.released_qiymat>0).map(r=>Object.assign({mezoni:`${fmtN(r.released_pct)}% / ${days} kun`},r)).sort((a,b)=>(b.released_pct-a.released_pct)||(b.released_qiymat-a.released_qiymat)).slice(0,15);
  let slow=rows.filter(r=>r.base_qiymat>0).map(r=>Object.assign({mezoni:`qoldiq ${fmtN(r.remain_pct)}%`},r)).sort((a,b)=>(b.remain_pct-a.remain_pct)||(b.remain_qiymat-a.remain_qiymat)).slice(0,15);
  let cols=[{k:"korxona",t:"Korxona nomi",w:"290px"},{k:"stir",t:"STIR",w:"86px"},{k:"released_qiymat",t:"Yechilgan qiymat (ming $)",n:1,f:fmtN,w:"105px"},{k:"released_vazn",t:"Yechilgan vazn (tn)",n:1,f:fmtN,w:"96px"},{k:"released_pct",t:"Yechilish ulushi (%)",n:1,f:fmtN,w:"86px"},{k:"remain_qiymat",t:"Qoldiq qiymat (ming $)",n:1,f:fmtN,w:"105px"},{k:"mezoni",t:"Mezon",w:"110px"}];
  return `<div class=grid2><div class=panel><h2>Eng tez nazoratdan yechadigan korxonalar</h2><div class=muted>Mezon: yechilgan qiymatning boshlang'ich qiymatga nisbati, teng holatda yechilgan qiymat. TOP 15.</div>${table(cols,basicTotal(fast,"IBK bo'yicha Jami","korxona"),"fixed-table")}</div><div class=panel><h2>Eng sekin nazoratdan yechadigan korxonalar</h2><div class=muted>Mezon: yakuniy sanada qolgan qiymat ulushi, teng holatda qoldiq qiymat. TOP 15.</div>${table(cols,basicTotal(slow,"IBK bo'yicha Jami","korxona"),"fixed-table")}</div>${chartBlock("Korxonalar yechilishi qiymat bo'yicha ulushi",fast,"korxona","released_qiymat",fmtN)}</div>`;
}
function warehouseTurnoverRows(data){
  let all=[...(data.rows||[]),...(data.unreleased||[])], map={};
  all.forEach(r=>{
    let name=r.warehouse||"O'z ombor";
    if(!map[name])map[name]={key:{warehouse:name},name,base_vazn:0,remain_vazn:0,released_vazn:0,base_qiymat:0,released_qiymat:0,remain_qiymat:0,partiya:0};
    ["base_vazn","remain_vazn","released_vazn","base_qiymat","released_qiymat","remain_qiymat"].forEach(k=>map[name][k]+=+r[k]||0);
    map[name].partiya+=+r.released_partiya||0;
  });
  return Object.values(map).map(r=>Object.assign(r,{released_pct:r.base_vazn?r.released_vazn/r.base_vazn*100:0,remain_pct:r.base_vazn?r.remain_vazn/r.base_vazn*100:0})).sort((a,b)=>b.released_vazn-a.released_vazn);
}
function warehouseTurnoverPanel(data){
  let rows=warehouseTurnoverRows(data);
  let fast=rows.filter(r=>r.released_vazn>0).sort((a,b)=>(b.released_pct-a.released_pct)||(b.released_vazn-a.released_vazn)).slice(0,15);
  let slow=rows.filter(r=>r.base_vazn>0).sort((a,b)=>(b.remain_pct-a.remain_pct)||(b.remain_vazn-a.remain_vazn)).slice(0,15);
  let cols=[{k:"name",t:"Ombor",w:"320px"},{k:"base_vazn",t:"Boshlang'ich vazn (tn)",n:1,f:fmtN,w:"105px"},{k:"released_vazn",t:"Yechilgan vazn (tn)",n:1,f:fmtN,w:"105px"},{k:"released_pct",t:"Yechilish ulushi (%)",n:1,f:fmtN,w:"90px"},{k:"remain_vazn",t:"Qoldiq vazn (tn)",n:1,f:fmtN,w:"105px"},{k:"partiya",t:"Partiya",n:1,f:fmtI,w:"75px"}];
  return `<div class="panel wide"><h2>Omborlar oboroti: nazoratdan yechilgan vazn bo'yicha</h2>${table(cols,basicTotal(rows,"IBK bo'yicha Jami","name"),"fixed-table")}</div><div class=grid2><div class=panel><h2>Eng tez oborot qiladigan omborlar</h2>${table(cols,basicTotal(fast,"IBK bo'yicha Jami","name"),"fixed-table")}</div><div class=panel><h2>Eng sekin oborot qiladigan omborlar</h2>${table(cols,basicTotal(slow,"IBK bo'yicha Jami","name"),"fixed-table")}</div>${chartBlock("Omborlar yechilgan vazn bo'yicha ulushi",fast,"name","released_vazn",fmtN)}</div>`;
}
buildRelease=async function(){
  let base=$("relBase").value, final=$("relFinal").value;
  if(parseUzDate(base)>=parseUzDate(final)){
    $("releaseResult").innerHTML=`<div class=muted>Boshlang'ich sana yakuniy sanadan oldin bo'lishi kerak.</div>`;
    return;
  }
  let j=await api(`/api/release?base=${encodeURIComponent(base)}&final=${encodeURIComponent(final)}`);
  if(j.missing){
    $("releaseResult").innerHTML=`<div class=muted>Arxivda yo'q sana: ${j.missing.join(", ")}. Shu davr uchun asos fayl yuklash kerak.</div>`;
    return;
  }
  window.LAST_RELEASE=j;
  let companies=releaseCompanyRows(j);
  let rows=numbered([releaseTotalRow(companies)].concat(companies));
  $("releaseResult").innerHTML=`<h2>${base} - ${final} <a class="btn light" href="/api/export?kind=release&base=${base}&final=${final}&token=${TOKEN}">Excel</a></h2><div class=overview-note>Boshlang'ich davr: ${base}. Yakuniy davr: ${final}. Jadval namuna fayldagi kabi korxonalar kesimida yig'iladi; ustun nomlari ixcham bo'lishi uchun keyingi qatorga ko'shiriladi.</div>${table(releaseCols(base,final),rows,"release-table sample-release-table fixed-table")}${releaseSpeedPanels(companies,base,final)}`;
  window.LAST_RELEASE=j;
}
function uniqueArchiveRows(){
  let best={};
  (ARCHIVE||[]).forEach(r=>{
    if(!r||!r.date)return;
    let score=(String(r.source||"").includes("IBK_Dashboard_SQLite")?2:0)+(r.deposit?1:0)+String(r.id||"").length/1000;
    if(!best[r.date]||score>best[r.date]._score)best[r.date]=Object.assign({_score:score},r);
  });
  return Object.values(best).sort((a,b)=>parseUzDate(b.date)-parseUzDate(a.date));
}
function compactArchivePanel(){
  let rows=uniqueArchiveRows();
  let body=rows.map(r=>{
    let source=(r.source||"").split(/[\\/]/).pop(), deposit=r.deposit?((r.deposit||"").split(/[\\/]/).pop()||"Bor"):"-";
    return `<tr><td class=num>${esc(r.date)}</td><td class=text title="${esc(source)}">${esc(source)}</td><td class=text title="${esc(deposit)}">${esc(deposit)}</td><td class=num><button class="light" onclick="loadReport('${esc(r.id)}')">Ochish</button></td></tr>`;
  }).join("");
  let html=`<table class="fixed-table compact-archive"><colgroup><col style="width:95px"><col style="width:360px"><col style="width:160px"><col style="width:90px"></colgroup><thead><tr><th>Sana</th><th>Asos fayl</th><th>Depozit</th><th></th></tr></thead><tbody>${body}</tbody></table>`;
  return `<div class=panel><h2>Arxiv</h2><div class=muted>${rows.length} ta sana bo'yicha yagona arxiv yozuvi. Dublikat sanalar yashirildi va ro'yxatdan tozalandi.</div>${html}</div>`;
}
const renderArchiveCompact=render;render=function(){
  renderArchiveCompact();
  if(TAB==="archive"){
    let v=$("view");
    if(v)v.innerHTML=compactArchivePanel();
  }
}
const showLoginBase=showLogin;showLogin=function(){showLoginBase();let el=$("login");if(el)el.classList.remove("active");clearBusy()}
const pollBase=poll;poll=async function(id){try{await pollBase(id)}finally{if(DATA)clearBusy()}}
const prepareArtifactsBase=prepareArtifacts;prepareArtifacts=async function(){let btn=(typeof event!=="undefined"&&event)?event.currentTarget:null;try{setBusy(btn,true,"Tayyorlash");await prepareArtifactsBase()}catch(e){clearBusy();throw e}}
const bindUploadBusy=bindUpload;bindUpload=function(){
  bindUploadBusy();
  let f=$("upload");if(f)f.onsubmit=async e=>{e.preventDefault();let btn=e.submitter;try{setBusy(btn,true,"Yuklanmoqda");$("status").textContent="BNRTE fayllari yuklanyapti...";let j=await api("/api/reports",{method:"POST",body:new FormData(f)});poll(j.job_id)}catch(err){$("status").textContent=err.message;setBusy(btn,false)}};
  let bf=$("bulkUpload");if(bf)bf.onsubmit=async e=>{e.preventDefault();let btn=e.submitter;try{setBusy(btn,true,"Yuklanmoqda");$("status").textContent="Yillik asos fayllar yuklanyapti...";let j=await api("/api/reports_bulk",{method:"POST",body:new FormData(bf)});let skip=(j.skipped||[]).length?` O'tkazib yuborildi: ${j.skipped.map(x=>x.filename).join(", ")}`:"";$("bulkResult").textContent=`${j.count||0} ta fayl navbatga qo'shildi.${skip}`;$("status").textContent="Bulk yuklash boshlandi";await loadArchive()}catch(err){$("status").textContent=err.message}finally{setBusy(btn,false)}};
  let tf=$("tolovUpload");if(tf)tf.onsubmit=async e=>{e.preventDefault();let btn=e.submitter;try{setBusy(btn,true,"Shakllanmoqda");$("status").textContent="To'lovlar shakllantirilyapti...";let j=await api("/api/tolov",{method:"POST",body:new FormData(tf)});PAYMENTS=j.payments||[];$("tolovUploadResult").innerHTML=`Tayyor: ${fmtI(PAYMENTS.reduce((a,r)=>a+(+r.rows||0),0))} qator, ${fmtN(PAYMENTS.reduce((a,r)=>a+(+r.sum||0),0))} so'm.`;$("status").textContent="To'lovlar tayyor"}catch(err){$("status").textContent=err.message}finally{setBusy(btn,false)}};
}
const cleanTopActionsBase=cleanTopActions;cleanTopActions=function(){
  cleanTopActionsBase();
  let a=$("actions");if(!a)return;
  let pngLinks=[...a.querySelectorAll("a")].filter(x=>x.textContent.trim()==="PNG"||x.hasAttribute("data-pngzip"));
  pngLinks.slice(1).forEach(x=>x.remove());
  ["O'zb","\u040e\u0437\u0431","\u0420\u0443\u0441","\u25d0"].forEach(label=>{
    let buttons=[...a.querySelectorAll("button")].filter(b=>b.textContent.trim()===label);
    buttons.slice(1).forEach(b=>b.remove());
  });
}
const renderFinalPolish=render;render=function(){renderFinalPolish();cleanTopActions();updateClock();if(TOKEN){let l=document.getElementById("login");if(l){l.style.cssText="display:none!important";l.classList.add("hidden");}}}

// ===== UPLOAD YANGILANMASI =====
function showUploadProgress(label,pct){
  let el=$("uploadStatus");if(!el)return;
  if(!label){el.innerHTML='';return}
  let c=pct>=100?'#16a34a':pct===0?'#dc2626':'#1d72b8';
  el.innerHTML=`<div style="background:#f0f6ff;border:1px solid #c8dff5;border-radius:8px;padding:10px 14px"><div style="font-size:13px;font-weight:700;color:${c};margin-bottom:6px">${esc(label)}</div><div style="background:#dde8f5;border-radius:6px;height:8px;overflow:hidden"><div style="width:${Math.max(3,pct)}%;background:${c};height:100%;border-radius:6px;transition:width .4s ease"></div></div></div>`;
}
poll=async function(id){
  try{
    let j=await api("/api/jobs/"+id);
    let pMap={"navbatda":8,"Hisob-kitob qilinyapti":30,"SQLite bazaga saqlanmoqda":55,"Dashboard JSON tayyorlanmoqda":70,"Excel/PNG/PDF tayyorlanmoqda":82,"Depozit qayta ishlanmoqda":45};
    let pct=j.status==="tayyor"?100:j.status==="xatolik"?0:(pMap[j.status]||35);
    showUploadProgress(j.status==="tayyor"?"✓ Tayyor!":j.status==="xatolik"?"✗ Xatolik":j.status+" ...",pct);
    $("status").textContent=j.status;
    if(j.status==="xatolik"){showUploadProgress("✗ "+(j.error||"Noma'lum xato"),0);clearBusy();return}
    if(j.status!=="tayyor"){setTimeout(()=>poll(id),1800);return}
    DATA=j.data;TAB="umumiy";await loadArchive();clearBusy();setTimeout(()=>showUploadProgress("",0),4000);render();
  }catch(err){$("status").textContent=err.message;clearBusy()}
};
async function uploadDepositOnly(btn){
  let f=$("depositOnlyForm");if(!f)return;
  let sel=f.querySelector("[name=report_id]"),dep=f.querySelector("[name=deposit]");
  if(!sel||!dep||!dep.files.length){$("depositResult").textContent="Hisobot va depozit fayl tanlang";return}
  setBusy(btn,true,"Yuklanmoqda");$("depositResult").textContent="Depozit yuklanyapti...";
  showUploadProgress("Depozit fayl yuklanyapti...",10);
  try{
    let fd=new FormData();fd.append("report_id",sel.value);fd.append("deposit",dep.files[0]);
    let j=await api("/api/deposit",{method:"POST",body:fd});
    $("depositResult").textContent="Qayta hisoblanmoqda...";
    poll(j.job_id);
  }catch(err){$("depositResult").textContent="Xato: "+err.message;showUploadProgress("✗ "+err.message,0)}
  finally{setBusy(btn,false)}
}
async function refreshCurrentReport(btn){
  if(!DATA)return;
  setBusy(btn,true,"Yangilanmoqda");$("status").textContent="Ma'lumotlar yangilanmoqda...";
  try{DATA=await api("/api/reports/"+DATA.id);render();$("status").textContent="Yangilandi"}catch(err){$("status").textContent=err.message}finally{setBusy(btn,false)}
}
uploadPanel=function(){
  let arcOpts=(ARCHIVE||[]).slice().reverse().slice(0,40).map(r=>`<option value="${r.id}">${r.date} — ${esc((r.source||"").split(/[\\/]/).pop())}</option>`).join("");
  let refreshHtml=DATA?`<div class=panel style="border-left:4px solid #1d72b8"><h2>&#8635; Joriy ma\'lumotlarni yangilash</h2><p class=muted>Ko\'rinishda: <b>${DATA.meta&&DATA.meta.date||"?"}</b> holatiga hisobot yukli.</p><button onclick="refreshCurrentReport(this)">Ma\'lumotlarni qayta yuklash</button></div>`:"";
  return `<div class=stack>
<div class=panel><h2>Fayl yuklash markazi</h2><div class=muted>BNRTE, Depozit va To\'lovlar uchun fayllar alohida yuklanadi. Jarayon holati real vaqtda ko\'rinadi.</div></div>
<div id=uploadStatus></div>
${refreshHtml}
<div class=panel><h2>BNRTE jamlanma — yangi hisobot yuklash</h2>
<form class="upload" id="upload">
<div><label>Asos fayl (xls/xlsx)</label><input name="source" type="file" accept=".xls,.xlsx,.html,.htm" required></div>
<div><label>Depozit fayl (ixtiyoriy)</label><input name="deposit" type="file" accept=".xlsx"></div>
<div></div><button>Yuklash va hisoblash</button>
</form>
<div class=muted style="margin-top:6px">Hisobot sanasi asos fayl nomidan avtomatik aniqlanadi.</div>
</div>
<div class=panel><h2>Depozit faylni alohida yangilash</h2>
<form class="upload" id="depositOnlyForm">
<div><label>Hisobot tanlang</label><select name="report_id" style="width:100%;padding:8px;border:1px solid var(--line);border-radius:6px">${arcOpts||"<option>Arxiv bo\'sh</option>"}</select></div>
<div><label>Depozit fayl (xlsx)</label><input name="deposit" type="file" accept=".xlsx" required></div>
<div></div><button type="button" onclick="uploadDepositOnly(this)">Depozit yuklash</button>
</form>
<div id="depositResult" class=muted></div>
</div>
<div class=panel><h2>Yillik arxivni birdan yuklash</h2>
<form class="upload" id="bulkUpload">
<div><label>Asos fayllar (bir nechta)</label><input name="sources" type="file" accept=".xls,.xlsx,.html,.htm" multiple required></div>
<div><label>Depozit fayl (ixtiyoriy)</label><input name="deposit" type="file" accept=".xlsx"></div>
<div></div><button>Hammasini yuklash</button>
</form>
<div id=bulkResult class=muted>Fayllar sanasi nomidan olinadi va arxivga qo\'shiladi.</div>
</div>
<div class=panel><h2>To\'lovlar jadvallari</h2>
<form class="upload" id="tolovUpload">
<div><label>To\'lov baza fayli</label><input name="source" type="file" accept=".xlsx,.xls" required></div>
<div class=muted>04.06+07.06.2026 kabi asos fayl yuklanadi.</div>
<div></div><button>To\'lov jadvallarini shakllantirish</button>
</form>
<div id="tolovUploadResult" class=muted></div>
</div></div>`;
};
const bindUploadFull=bindUpload;bindUpload=function(){
  let f=$("upload");
  if(f)f.onsubmit=async e=>{
    e.preventDefault();let btn=e.submitter;
    try{setBusy(btn,true,"Yuklanmoqda");showUploadProgress("Fayl yuklanyapti...",5);
      let j=await api("/api/reports",{method:"POST",body:new FormData(f)});poll(j.job_id);
    }catch(err){$("status").textContent=err.message;showUploadProgress("✗ "+err.message,0)}
    finally{setBusy(btn,false)};
  };
  let bf=$("bulkUpload");
  if(bf)bf.onsubmit=async e=>{
    e.preventDefault();let btn=e.submitter;
    try{setBusy(btn,true,"Yuklanmoqda");showUploadProgress("Yillik fayllar yuklanyapti...",5);
      let j=await api("/api/reports_bulk",{method:"POST",body:new FormData(bf)});
      let skip=(j.skipped||[]).length?` O\'tkazib: ${j.skipped.map(x=>x.filename).join(", ")}`:"";
      $("bulkResult").textContent=`${j.count||0} ta fayl navbatga qo\'shildi.${skip}`;
      showUploadProgress(`${j.count||0} ta fayl navbatga qo\'shildi`,85);
      await loadArchive();setTimeout(()=>showUploadProgress("",0),3000);
    }catch(err){$("status").textContent=err.message;showUploadProgress("✗ "+err.message,0)}
    finally{setBusy(btn,false)};
  };
  let tf=$("tolovUpload");
  if(tf)tf.onsubmit=async e=>{
    e.preventDefault();let btn=e.submitter;
    try{setBusy(btn,true,"Shakllanmoqda");showUploadProgress("To\'lovlar shakllantirilyapti...",10);
      let j=await api("/api/tolov",{method:"POST",body:new FormData(tf)});
      PAYMENTS=j.payments||[];
      $("tolovUploadResult").innerHTML=`Tayyor: ${fmtI(PAYMENTS.reduce((a,r)=>a+(+r.rows||0),0))} qator, ${fmtN(PAYMENTS.reduce((a,r)=>a+(+r.sum||0),0))} so\'m.`;
      showUploadProgress("To\'lovlar tayyor!",100);setTimeout(()=>showUploadProgress("",0),3000);
    }catch(err){$("status").textContent=err.message;showUploadProgress("✗ "+err.message,0)}
    finally{setBusy(btn,false)};
  };
};
// ===== /UPLOAD YANGILANMASI =====

function startBackgroundVideo(){
  const video=document.getElementById('bgVideo'), canvas=document.getElementById('bgCanvas');
  if(!video||!canvas||!canvas.captureStream)return;
  const ctx=canvas.getContext('2d');
  function resize(){const d=Math.min(1.25,Math.max(1,window.devicePixelRatio||1));canvas.width=Math.round(1280*d);canvas.height=Math.round(720*d)}
  resize(); window.addEventListener('resize',resize,{passive:true});
  const routes=[{x1:.08,y1:.68,x2:.44,y2:.38,x3:.82,y3:.58,delay:0},{x1:.18,y1:.28,x2:.54,y2:.18,x3:.92,y3:.34,delay:.22},{x1:.05,y1:.48,x2:.38,y2:.60,x3:.76,y3:.26,delay:.48},{x1:.30,y1:.82,x2:.62,y2:.50,x3:.98,y3:.72,delay:.68}];
  function plane(ctx,x,y,a,scale,alpha){ctx.save();ctx.translate(x,y);ctx.rotate(a);ctx.scale(scale,scale);ctx.globalAlpha=alpha;ctx.fillStyle='#1d77b8';ctx.strokeStyle='rgba(255,255,255,.75)';ctx.lineWidth=1.2;ctx.beginPath();ctx.moveTo(18,0);ctx.lineTo(-14,-6);ctx.lineTo(-8,0);ctx.lineTo(-14,6);ctx.closePath();ctx.fill();ctx.stroke();ctx.fillStyle='#42b9c7';ctx.fillRect(-20,-2,8,4);ctx.restore()}
  function bez(p0,p1,p2,t){let u=1-t;return {x:u*u*p0.x+2*u*t*p1.x+t*t*p2.x,y:u*u*p0.y+2*u*t*p1.y+t*t*p2.y}}
  function draw(t){
    const w=canvas.width,h=canvas.height, time=t/1000;
    let g=ctx.createLinearGradient(0,0,0,h);g.addColorStop(0,'#eaf8ff');g.addColorStop(.46,'#ffffff');g.addColorStop(1,'#e9f5fb');ctx.fillStyle=g;ctx.fillRect(0,0,w,h);
    ctx.save();ctx.globalAlpha=.34;ctx.strokeStyle='rgba(48,128,188,.16)';ctx.lineWidth=1;for(let x=((time*10)%64)-64;x<w;x+=64){ctx.beginPath();ctx.moveTo(x,0);ctx.lineTo(x+h*.15,h);ctx.stroke()}for(let y=((time*7)%56)-56;y<h;y+=56){ctx.beginPath();ctx.moveTo(0,y);ctx.lineTo(w,y+h*.04);ctx.stroke()}ctx.restore();
    ctx.save();ctx.globalCompositeOperation='multiply';for(let i=0;i<7;i++){let x=((time*18+i*260)%(w+360))-180,y=h*(.12+.13*Math.sin(time*.21+i));let rg=ctx.createRadialGradient(x,y,0,x,y,w*.17);rg.addColorStop(0,'rgba(198,228,244,.42)');rg.addColorStop(1,'rgba(198,228,244,0)');ctx.fillStyle=rg;ctx.fillRect(0,0,w,h)}ctx.restore();
    ctx.save();ctx.strokeStyle='rgba(31,127,190,.24)';ctx.lineWidth=2;ctx.setLineDash([10,12]);for(const r of routes){ctx.beginPath();ctx.moveTo(r.x1*w,r.y1*h);ctx.quadraticCurveTo(r.x2*w,r.y2*h,r.x3*w,r.y3*h);ctx.stroke()}ctx.restore();
    ctx.save();ctx.globalCompositeOperation='screen';for(let i=0;i<3;i++){let cx=w*(.16+i*.31),cy=h*(.34+.12*Math.sin(time*.18+i));let rr=((time*38+i*70)%210)+30;ctx.strokeStyle='rgba(67,178,198,'+(0.24-rr/1100)+')';ctx.lineWidth=2;ctx.beginPath();ctx.arc(cx,cy,rr,0,Math.PI*2);ctx.stroke()}ctx.restore();
    for(const r of routes){let tt=(time*.065+r.delay)%1,p0={x:r.x1*w,y:r.y1*h},p1={x:r.x2*w,y:r.y2*h},p2={x:r.x3*w,y:r.y3*h},p=bez(p0,p1,p2,tt),q=bez(p0,p1,p2,Math.min(.995,tt+.01));plane(ctx,p.x,p.y,Math.atan2(q.y-p.y,q.x-p.x),1.0+tt*.25,.75)}
    ctx.save();ctx.globalAlpha=.7;ctx.fillStyle='rgba(255,255,255,.46)';ctx.fillRect(0,0,w,h);ctx.restore();
    requestAnimationFrame(draw);
  }
  function attach(){const stream=canvas.captureStream(30);window.BG_STREAM=stream;video.muted=true;video.playsInline=true;video.autoplay=true;video.srcObject=stream;video.play().catch(()=>{})}
  requestAnimationFrame(draw);setTimeout(attach,120);
}

/* Excel-like table polish and sample-table renderers */
const tableCore=table;
table=function(h,rows,cls=""){rows=rows||[];return `<table class="${cls}"><colgroup>${h.map(x=>`<col style="${x.w?'width:'+x.w:''}">`).join("")}</colgroup><thead><tr>${h.map(x=>`<th class="${x.n?'num':'text'} col-${x.k}">${x.t}</th>`).join("")}</tr></thead><tbody>${rows.map((r,ri)=>`<tr class="${r._class||''}" onclick='detail(${JSON.stringify(r.key||{}).replaceAll("'","&#39;")})'>${h.map(x=>{let cls=(x.n?'num':'text')+` col-${x.k}`,raw=r[x.k],val=x.n&&(raw===""||raw===null||raw===undefined)?"":(x.f?x.f(raw):esc(raw)),tip=esc(raw);if(ri===0&&!x.n&&(val==="Jami"||String(val).startsWith("Jami ")))val="IBK bo'yicha Jami";if(x.k==="released_partiya"&&(+raw||0)===0&&((+r.released_qiymat||0)>0.005||(+r.released_vazn||0)>0.005)){cls+=" partial";tip="Qisman yechilgan";val="0"}return `<td class="${cls}" title="${tip}">${val}</td>`}).join("")}</tr>`).join("")}</tbody></table>`}
function excelExpiredPostRegime(){let rows=expiredPostRegimeRows();let body=rows.map(r=>`<tr onclick='detail(${JSON.stringify(r.key||{}).replaceAll("'","&#39;")})'><td class=text title="${esc(r.post||'')}">${esc(r.post||'')}</td><td class=num>${fmtI(r.jami_partiya)}</td><td class=num>${fmtN(r.jami_qiymat)}</td><td class=num>${fmtI(r.expired_partiya)}</td><td class=num>${fmtN(r.expired_qiymat)}</td><td class=num>${fmtN(r.ulush)}</td><td class=num>${fmtI(r.im70_partiya)}</td><td class=num>${fmtN(r.im70_qiymat||0)}</td><td class=num>${fmtI(r.im74_partiya)}</td><td class=num>${fmtN(r.im74_qiymat||0)}</td><td class=num>${fmtI(r.tr80_partiya)}</td><td class=num>${fmtN(r.tr80_qiymat||0)}</td></tr>`).join("");return `<table class="sample-like expired-sample"><thead><tr><th rowspan=4>Bojxona postlari</th><th colspan=2 rowspan=3>Jami nazoratdagi</th><th colspan=9>Rejim muddati o'tgan tovarlar</th></tr><tr><th colspan=3 rowspan=2>Jami</th><th colspan=6>Shundan</th></tr><tr><th colspan=2>Vaqtincha saqlash IM70</th><th colspan=2>Bojxona ombori IM74</th><th colspan=2>Tranzit TR80</th></tr><tr><th>Partiya</th><th>Qiymati,<br>ming doll.</th><th>Partiya</th><th>Qiymati,<br>ming doll.</th><th>Partiyadagi<br>ulushi (%)</th><th>Partiya</th><th>Qiymati,<br>ming doll.</th><th>Partiya</th><th>Qiymati,<br>ming doll.</th><th>Partiya</th><th>Qiymati,<br>ming doll.</th></tr></thead><tbody>${body}</tbody></table>`}
expiredTotalExcelTable=function(){return excelExpiredPostRegime()}
function excelFoodTable(){let rows=foodRows().map((r,i)=>Object.assign({rn:i?i:"Jami:"},r));let body=rows.map(r=>`<tr onclick='detail(${JSON.stringify(r.key||{}).replaceAll("'","&#39;")})'><td class=num>${esc(r.rn)}</td><td class=text title="${esc(r.name)}">${esc(r.name)}</td><td class=num>${fmtN(r.vazn)}</td><td class=num>${fmtN(r.qiymat)}</td><td class=num>${fmtN(r.over_vazn)}</td><td class=num>${fmtN(r.over_qiymat)}</td><td class=num>${fmtN(r.ulush)}</td></tr>`).join("");return `<table class="sample-like food-sample"><colgroup><col style="width:36px"><col style="width:auto"><col style="width:86px"><col style="width:110px"><col style="width:86px"><col style="width:110px"><col style="width:90px"></colgroup><thead><tr><th rowspan=2>№</th><th rowspan=2>Tovarlar turi</th><th rowspan=2>Vazni,<br>tn</th><th rowspan=2>Qiymati,<br>ming AQSH doll.</th><th colspan=2>Shundan, 3 oydan oshganlari</th><th rowspan=2>Qiymatdagi<br>ulushi (%)</th></tr><tr><th>Vazni,<br>tn</th><th>Qiymati,<br>ming AQSH doll.</th></tr></thead><tbody>${body}</tbody></table>`}
function excelRegimeYearTable(){let rows=DATA.regime_year_post||[],posts=[...new Set(rows.map(r=>r.post))],out=[];posts.forEach(p=>{let pr=rows.filter(r=>r.post===p);out.push({name:p,_class:"merged-row"});["IM70","IM74","TR80"].forEach(reg=>{let rr=pr.filter(x=>x.rejim===reg);if(!rr.length)return;out.push({name:reg,partiya:rr.reduce((a,x)=>a+(+x.partiya||0),0),vazn:rr.reduce((a,x)=>a+(+x.vazn||0),0),qiymat:rr.reduce((a,x)=>a+(+x.qiymat||0),0),tolov:rr.reduce((a,x)=>a+(+x.tolov||0),0),_class:"sub-total"});rr.forEach(x=>out.push({name:(x.yil||'-')+' yil',partiya:x.partiya,vazn:x.vazn,qiymat:x.qiymat,tolov:x.tolov,key:x.key}))})});let total={name:"IBK bo'yicha Jami",partiya:rows.reduce((a,x)=>a+(+x.partiya||0),0),vazn:rows.reduce((a,x)=>a+(+x.vazn||0),0),qiymat:rows.reduce((a,x)=>a+(+x.qiymat||0),0),tolov:rows.reduce((a,x)=>a+(+x.tolov||0),0),_class:"grand-total"};return table([{k:"name",t:"Ko'rsatkich nomi",w:"230px"},{k:"partiya",t:"Partiya soni",w:"90px",n:1,f:fmtI},{k:"vazn",t:"Vazni<br>(tn.)",w:"115px",n:1,f:fmtN},{k:"qiymat",t:"Qiymati<br>(ming $)",w:"126px",n:1,f:fmtN},{k:"tolov",t:"To'lov<br>(mln.so'm)",w:"132px",n:1,f:fmtN}],[total].concat(out),"sample-like regime-year-table")}
const oldTransportPanel=transportPanel;transportPanel=function(){return oldTransportPanel().replace("Deklaratsiya post kodi bo'yicha tahlil","Transport turi bo'yicha tahlil").replace("Post kodi","Jo'natuvchi post")}
const renderSampleTables=render;render=function(){renderSampleTables();let v=$("view");if(!v||!DATA)return;if(TAB==="rejim")v.innerHTML=`<div class=stack><div class=panel><h2>70-74-80: post, rejim va yillar kesimida</h2>${excelRegimeYearTable()}</div><div class=panel><h2>Rejimlar qiymat bo'yicha ulushi</h2>${bars(DATA.regimes||[],"rejim","qiymat",fmtN)}</div></div>`;if(TAB==="food")v.innerHTML=`<div class=panel><h2>Oziq-ovqatlar kesimida ${xls("food")}</h2>${excelFoodTable()}${bars(DATA.food||[],"name","qiymat",fmtN)}</div>`}
const oldOverviewPanels3=overviewPanels;overviewPanels=function(){return oldOverviewPanels3()}
document.body.classList.add("login-screen");setTimeout(()=>{let eye=document.querySelector(".eye-btn");if(eye)eye.innerHTML="&#128065;";},0);
const showLoginScreenBase=showLogin;showLogin=function(){
  document.body.classList.add("login-screen");document.body.classList.remove("logged-in");
  let login=$("login"),app=$("app"),dash=$("dash"),tabs=$("tabs"),view=$("view"),kpis=$("kpis"),actions=$("actions"),err=$("loginError"),meta=$("meta");
  if(login){login.classList.remove("hidden");login.style.display="grid";}
  if(app){app.classList.add("hidden");app.style.display="none";}
  if(dash){dash.classList.add("hidden");dash.style.display="none";}
  if(tabs)tabs.innerHTML="";if(view)view.innerHTML="";if(kpis)kpis.innerHTML="";if(actions)actions.innerHTML="";
  if(err)err.textContent="";if(meta)meta.textContent="Kirish kerak";
}
doLogin=async function(){
  let btn=$("loginBtn"),err=$("loginError");
  if(err)err.textContent="";
  try{
    setBusy(btn,true,"Kirish");
    let user=($("user")?.value||"").trim(),pass=$("pass")?.value||"";
    if(!user||!pass)throw Error("Login va parolni kiriting");
    let j=await api("/api/login",{method:"POST",body:JSON.stringify({user,pass})});
    TOKEN=j.token;localStorage.ibk_token=TOKEN;ME=j.user;
    await showApp();
  }catch(e){
    forceLoginView();
    activateLogin();
    if(err)err.textContent=(e&&e.message&&e.message!=="login")?e.message:"Login yoki parol xato";
  }finally{setBusy(btn,false)}
}
function forgotPassword(){
  let err=$("loginError"),u=($("user")?.value||"").trim();
  if(err)err.textContent=u?`${u} uchun vaqtinchalik parol SMS orqali yuborish moduli tayyorlanmoqda. SMS provayder ulangandan keyin ishlaydi.`:"Avval loginni kiriting, keyin SMS orqali tiklash so'rovi yuboriladi.";
}
const showAppStableBase=showApp;
showApp=async function(){
  document.body.classList.remove("login-screen");
  document.body.classList.add("logged-in");
  let loginEl=$("login"),appEl=$("app"),dashEl=$("dash"),tabsEl=$("tabs"),viewEl=$("view"),kpisEl=$("kpis"),actionsEl=$("actions"),metaEl=$("meta");
  if(loginEl){loginEl.classList.add("hidden");loginEl.style.display="none";}
  if(appEl){appEl.classList.remove("hidden");appEl.style.display="block";}
  if(dashEl){dashEl.classList.remove("hidden");dashEl.style.display="block";}
  ME=await api("/api/me");
  LANG=ME.lang||localStorage.ibk_lang||"uz";
  await loadArchive();
  await loadPayments();
  DATA=null;TAB="home";GROUP="home";
  if(metaEl)metaEl.textContent="Tizimga xush kelibsiz";
  if(kpisEl)kpisEl.innerHTML="";
  if(tabsEl)tabsEl.innerHTML=`<button class="module-parent" onclick="openGroup('bnrte','umumiy')"><span>▦</span>BNRTE</button><button class="module-parent pay" onclick="openGroup('payments','payments')"><span>$</span>To'lovlar</button><button class="module-parent" onclick="openGroup('common','upload')"><span>⚙</span>Boshqaruv</button>`;
  if(actionsEl)actionsEl.innerHTML=`<button class="light lang-btn" onclick="setLang('uz')">O'zb</button><button class="light lang-btn" onclick="setLang('uzc')">Ўзб</button><button class="light lang-btn" onclick="setLang('ru')">Рус</button><button class="light lang-btn" onclick="document.body.classList.toggle('dark')">◐</button> <button class="logout-btn" onclick="logout()">Chiqish</button>`;
  if(viewEl) viewEl.innerHTML=landingPanel();
  setTimeout(()=>{initFlightsMap();initCountryFlowMap();},200);
}
let AUTO_LOGOUT_MS=20*60*1000,autoLogoutTimer=null;
function resetAutoLogout(){clearTimeout(autoLogoutTimer);if(TOKEN)autoLogoutTimer=setTimeout(()=>{logout();let e=$("loginError");if(e)e.textContent="20 daqiqa faollik bo'lmagani uchun qayta kirish kerak."},AUTO_LOGOUT_MS)}
["click","keydown","mousemove","touchstart","scroll"].forEach(ev=>document.addEventListener(ev,resetAutoLogout,{passive:true}));
const showAppAutoLogout=showApp;showApp=async function(){await showAppAutoLogout();resetAutoLogout()}
const logoutAutoBase=logout;logout=function(){clearTimeout(autoLogoutTimer);logoutAutoBase();}
const renderRealGlobe=render;render=function(){renderRealGlobe();setTimeout(()=>{initFlightsMap();initCountryFlowMap();},80)}
setBg(localStorage.ibk_bg||"premium");
startBackgroundVideo();
function releaseDatePair(id,title){return `<div class="release-date-card"><h3>${title}</h3><div class="filters compact-filters"><select id="${id}Base">${dateOptions()}</select><select id="${id}Final">${dateOptions()}</select><button onclick="buildReleaseSection('${id}')">Shakllantirish</button></div><div id="${id}Result" class="release-section-result"></div></div>`}
function releaseDashboardPanel(){return `<div class="stack"><div class="panel"><h2>Nazoratdan yechish</h2><div class="overview-note">Tepadagi sanalar barcha jadvallar uchun umumiy ishlashi yoki har jadval alohida muddat bilan shakllanishi mumkin.</div><div class="filters compact-filters"><label class="inline-check"><input type="checkbox" id="relGlobalUse" checked onchange="syncReleaseDates()"> Barcha jadvallar uchun</label><select id="relGlobalBase" onchange="syncReleaseDates()">${dateOptions()}</select><select id="relGlobalFinal" onchange="syncReleaseDates()">${dateOptions()}</select><button onclick="buildAllReleaseSections()">Barchasini shakllantirish</button></div></div>${releaseDatePair('relMain','Nazoratdan yechilishi jadvali')}${releaseDatePair('relSpeed','Korxonalar: eng ko\'p, eng tez va eng sekin')}${releaseDatePair('relWh','Omborlar oboroti')}</div>`}
function syncReleaseDates(){let use=$('relGlobalUse')?.checked,base=$('relGlobalBase')?.value,final=$('relGlobalFinal')?.value;['relMain','relSpeed','relWh'].forEach(id=>{let b=$(`${id}Base`),f=$(`${id}Final`);if(!b||!f)return;if(use){b.value=base;f.value=final;b.disabled=true;f.disabled=true}else{b.disabled=false;f.disabled=false}})}
function releaseDatesFor(id){let use=$('relGlobalUse')?.checked;if(use)return {base:$('relGlobalBase').value,final:$('relGlobalFinal').value};return {base:$(`${id}Base`).value,final:$(`${id}Final`).value}}
async function loadReleaseData(base,final,target){if(parseUzDate(base)>=parseUzDate(final)){target.innerHTML=`<div class=muted>Boshlang'ich sana yakuniy sanadan oldin bo'lishi kerak.</div>`;return null}let j=await api(`/api/release?base=${encodeURIComponent(base)}&final=${encodeURIComponent(final)}`);if(j.missing){target.innerHTML=`<div class=muted>Arxivda yo'q sana: ${j.missing.join(', ')}. Shu davr uchun asos fayl yuklash kerak.</div>`;return null}window.LAST_RELEASE=j;return j}
async function buildReleaseSection(id){let d=releaseDatesFor(id),box=$(`${id}Result`);if(!box)return;box.innerHTML='<div class=muted>Hisoblanmoqda...</div>';let j=await loadReleaseData(d.base,d.final,box);if(!j)return;let companies=releaseCompanyRows(j);if(id==='relMain'){let rows=numbered([releaseTotalRow(companies)].concat(companies));box.innerHTML=`<h2>${d.base} - ${d.final} <a class="btn light" href="/api/export?kind=release&base=${d.base}&final=${d.final}&token=${TOKEN}">Excel</a></h2><div class=overview-note>Boshlang'ich davr: ${d.base}. Yakuniy davr: ${d.final}.</div>${table(releaseCols(d.base,d.final),rows,'release-table sample-release-table fixed-table')}`;return}if(id==='relSpeed'){box.innerHTML=releaseSpeedPanels(companies,d.base,d.final);return}if(id==='relWh'){box.innerHTML=warehouseTurnoverPanel(j);return}}
async function buildAllReleaseSections(){syncReleaseDates();for(const id of ['relMain','relSpeed','relWh']) await buildReleaseSection(id)}
const renderReleaseDateControls=render;render=function(){renderReleaseDateControls();if(TAB==='released'){let v=$('view');if(v){v.innerHTML=releaseDashboardPanel();setTimeout(syncReleaseDates,0)}}}
/* === BOOT FIX: tokenni o'chirib tashlamaymiz; holatni to'g'ri ochamiz === */
function forceLoginView(){
  document.body.classList.add("login-screen");
  document.body.classList.remove("logged-in");
  let login=$("login"), app=$("app"), dash=$("dash"), tabs=$("tabs"), view=$("view"), kpis=$("kpis"), actions=$("actions"), meta=$("meta");
  if(login){login.classList.remove("hidden");login.style.display="grid";}
  if(app){app.classList.add("hidden");app.style.display="none";}
  if(dash){dash.classList.add("hidden");dash.style.display="none";}
  if(tabs)tabs.innerHTML="";
  if(view)view.innerHTML="";
  if(kpis)kpis.innerHTML="";
  if(actions)actions.innerHTML="";
  if(meta)meta.textContent="Kirish kerak";
}

function forceAppView(){
  document.body.classList.remove("login-screen");
  document.body.classList.add("logged-in");
  let login=$("login"), app=$("app"), dash=$("dash");
  if(login){login.classList.add("hidden");login.style.display="none";}
  if(app){app.classList.remove("hidden");app.style.display="block";}
  if(dash){dash.classList.remove("hidden");dash.style.display="block";}
}

const showLoginFinalBase = showLogin;
showLogin = function(){
  forceLoginView();
}

const showAppFinalBase = showApp;
showApp = async function(){
  forceAppView();
  await showAppFinalBase();
  forceAppView();
  window.scrollTo({top:0, behavior:"auto"});
}

/* doLogin final version is set above */

forceLoginView();
/* Hex monitoring background */
(function(){
  var sc=document.querySelector('.sky-scene');
  if(!sc)return;
  /* Kill all pseudo-element patterns that bleed through the canvas */
  var ks=document.createElement('style');
  ks.textContent='.sky-scene::before,.sky-scene::after{display:none!important;content:none!important;animation:none!important;background:none!important}';
  document.head.appendChild(ks);
  /* Remove every old sky-scene child — eliminates all CSS specificity fights */
  while(sc.firstChild)sc.removeChild(sc.firstChild);
  /* Force full-screen via setProperty so !important beats any stylesheet rule */
  var sp=sc.style;
  ['position:fixed','top:0','left:0','right:0','bottom:0','width:100%','height:100%',
   'z-index:0','overflow:hidden','pointer-events:none','background:#f8fbff',
   'animation:none','transform:none'].forEach(function(r){
    var i=r.indexOf(':');sp.setProperty(r.slice(0,i),r.slice(i+1),'important');
  });
  /* Build hex canvas */
  var cv=document.createElement('canvas');
  cv.id='hexBg';
  sp=cv.style;
  ['position:absolute','top:0','left:0','width:100%','height:100%',
   'display:block','z-index:1','pointer-events:none'].forEach(function(r){
    var i=r.indexOf(':');sp.setProperty(r.slice(0,i),r.slice(i+1),'important');
  });
  sc.appendChild(cv);
  var ctx=cv.getContext('2d');
  var R=28,W=0,H=0,cells=[];
  function dk(){return document.body.classList.contains('dark');}
  function resize(){
    W=window.innerWidth;H=window.innerHeight;
    cv.width=W;cv.height=H;
    var cols=Math.ceil(W/(R*1.73))+2,rows=Math.ceil(H/(R*1.5))+2;
    cells=[];
    for(var row=0;row<rows;row++)
      for(var col=0;col<cols;col++)
        cells.push({x:col*R*1.73+(row%2)*R*.87,y:row*R*1.5,
                    p:Math.random()*Math.PI*2,s:.012+Math.random()*.018});
  }
  function hx(x,y){
    ctx.beginPath();
    for(var i=0;i<6;i++){var a=Math.PI/3*i-Math.PI/6;
      ctx.lineTo(x+R*Math.cos(a),y+R*Math.sin(a));}
    ctx.closePath();
  }
  function draw(){
    var d=dk();
    ctx.clearRect(0,0,W,H);
    ctx.fillStyle=d?'#080f1e':'#f8fbff';ctx.fillRect(0,0,W,H);
    for(var i=0;i<cells.length;i++){
      var c=cells[i];c.p+=c.s;
      var v=(Math.sin(c.p)+1)/2;
      hx(c.x,c.y);
      ctx.fillStyle=d?('rgba(59,130,246,'+(v*.38).toFixed(3)+')')
                     :('rgba(100,160,220,'+(v*.22).toFixed(3)+')');
      ctx.fill();
      ctx.strokeStyle=d?('rgba(96,165,250,'+(.28+v*.55).toFixed(3)+')')
                       :('rgba(100,160,220,'+(.18+v*.35).toFixed(3)+')');
      ctx.lineWidth=1.5;ctx.stroke();
    }
    requestAnimationFrame(draw);
  }
  resize();window.addEventListener('resize',resize,{passive:true});draw();
})();
</script></main></body></html>"""


class Handler(BaseHTTPRequestHandler):
    def json(self, data, status=HTTPStatus.OK):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def body_json(self):
        return json.loads(self.rfile.read(int(self.headers.get("Content-Length", "0") or "0")).decode("utf-8") or "{}")

    def token(self):
        parsed = urlparse(self.path)
        q = parse_qs(parsed.query)
        return self.headers.get("X-Token") or q.get("token", [""])[0]

    def user(self):
        token = self.token()
        sess = SESSIONS.get(token)
        return sess["user"] if sess else None

    def require_user(self):
        user = self.user()
        if not user:
            self.json({"error": "login kerak"}, HTTPStatus.UNAUTHORIZED)
            return None
        return user

    def require_admin(self):
        user = self.require_user()
        if not user:
            return None
        users = load_json(USER_PATH, {})
        rec = users.get(user, {})
        if rec.get("role") != "admin" and "admin" not in rec.get("perms", []):
            self.json({"error": "admin huquqi kerak"}, HTTPStatus.FORBIDDEN)
            return None
        return user

    def require_perm(self, perm: str):
        user = self.require_user()
        if not user:
            return None
        rec = load_json(USER_PATH, {}).get(user, {})
        if rec.get("role") == "admin" or perm in rec.get("perms", []):
            return user
        self.json({"error": f"{perm} vakolati kerak"}, HTTPStatus.FORBIDDEN)
        return None

    def do_GET(self):
        global STORE
        parsed = urlparse(self.path)
        if parsed.path.startswith("/assets/"):
            name = Path(unquote(parsed.path.rsplit("/", 1)[-1])).name
            path = ASSET_DIR / name
            if not path.exists() or not path.is_file():
                self.send_error(404)
                return
            data = path.read_bytes()
            ctype = "image/jpeg" if path.suffix.lower() in {".jpg", ".jpeg"} else "application/octet-stream"
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if parsed.path == "/":
            body = HTML.encode("utf-8")
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        if parsed.path == "/api/flights":
            if not self.require_user():
                return
            self.json(get_tas_flights())
            return
        if parsed.path == "/api/flights_live":
            if not self.require_user():
                return
            self.json(get_live_states())
            return
        if parsed.path == "/api/company_trends":
            if not self.require_user():
                return
            try:
                if STORE is None:
                    STORE = IBKStore(DB_PATH)
                ensure_store_backfilled()
                self.json(STORE.company_series_all())
            except Exception as exc:
                self.json({"periods": [], "companies": [], "error": str(exc)})
            return
        if parsed.path == "/api/goods_trends":
            if not self.require_user():
                return
            try:
                if STORE is None:
                    STORE = IBKStore(DB_PATH)
                ensure_store_backfilled()
                self.json(STORE.goods_series_all())
            except Exception as exc:
                self.json({"periods": [], "goods": [], "error": str(exc)})
            return
        if parsed.path == "/api/warehouse_trends":
            if not self.require_user():
                return
            try:
                if STORE is None:
                    STORE = IBKStore(DB_PATH)
                ensure_store_backfilled()
                self.json(STORE.warehouse_series_all())
            except Exception as exc:
                self.json({"periods": [], "warehouses": [], "error": str(exc)})
            return
        if parsed.path == "/api/transport_trends":
            if not self.require_user():
                return
            try:
                if STORE is None:
                    STORE = IBKStore(DB_PATH)
                ensure_store_backfilled()
                self.json(STORE.transport_series_all())
            except Exception as exc:
                self.json({"periods": [], "transports": [], "error": str(exc)})
            return
        if parsed.path == "/api/transport_company_trends":
            if not self.require_user():
                return
            try:
                if STORE is None:
                    STORE = IBKStore(DB_PATH)
                ensure_store_backfilled()
                self.json(STORE.transport_company_series_all())
            except Exception as exc:
                self.json({"periods": [], "transports": [], "error": str(exc)})
            return
        if parsed.path == "/api/tolov":
            if not self.require_user():
                return
            self.json({"payments": tolov_summary_rows()})
            return
        if parsed.path == "/api/archive":
            if not self.require_user():
                return
            archive = load_json(INDEX_PATH, {"reports": []})
            archive["reports"] = clean_archive_records(archive.get("reports", []))
            save_json(INDEX_PATH, archive)
            self.json(archive)
            return
        if parsed.path == "/api/me":
            user = self.require_user()
            if not user:
                return
            rec = load_json(USER_PATH, {}).get(user, {})
            self.json({"user": user, "role": rec.get("role", "user"), "full_name": rec.get("full_name", ""), "position": rec.get("position", ""), "phone": rec.get("phone", ""), "lang": rec.get("lang", "uz"), "perms": rec.get("perms", [])})
            return
        if parsed.path == "/api/users":
            if not self.require_admin():
                return
            users = load_json(USER_PATH, {})
            self.json({"users": [{"user": u, "role": r.get("role", "user"), "role_label": r.get("role_label", ROLE_LABELS.get(r.get("role", "user"), "Foydalanuvchi")), "post_code": r.get("post_code", ""), "enabled": r.get("enabled", True), "full_name": r.get("full_name", ""), "position": r.get("position", ""), "phone": r.get("phone", ""), "lang": r.get("lang", "uz"), "perms": r.get("perms", [])} for u, r in users.items()]})
            return
        if parsed.path.startswith("/api/jobs/"):
            if not self.require_user():
                return
            self.json(JOBS.get(parsed.path.rsplit("/", 1)[-1], {"status": "xatolik", "error": "Job topilmadi"}))
            return
        if parsed.path.startswith("/api/reports/"):
            if not self.require_user():
                return
            report_id = parsed.path.rsplit("/", 1)[-1]
            item = next((r for r in load_json(INDEX_PATH, {"reports": []})["reports"] if r["id"] == report_id), None)
            if not item:
                self.json({"error": "Topilmadi"}, HTTPStatus.NOT_FOUND)
                return
            data_path = Path(item["dir"]) / "dashboard.json"
            data = load_json(data_path, {})
            needs_rebuild = data and (
                data.get("schema") != 5 or
                "food" not in data or "post_summary" not in data or "regime_posts" not in data or "regime_year_post" not in data or "expired_post_regime" not in data or "expired_block" not in data or "warehouse" not in data or
                "vazn" not in data.get("kpis", {}) or "expired_value" not in data.get("kpis", {}) or
                "depozit_matched" not in data.get("kpis", {})
            )
            if needs_rebuild:
                try:
                    deposit_path = Path(item["deposit"]) if item.get("deposit") else None
                    rebuilt = build_dashboard(item["id"], Path(item["source"]), deposit_path, datetime.strptime(item["date"], "%d.%m.%Y"))
                    rebuilt["files"] = data.get("files") or {"status": "tayyorlanmoqda", "excel": "", "pdf": "", "pngs": []}
                    data = rebuilt
                    save_json(data_path, data)
                except Exception:
                    pass
            if data:
                files = data.setdefault("files", {})
                if files.get("status") == "tayyorlanmoqda" or not files.get("excel"):
                    try:
                        files = update_report_files(report_id, Path(item["dir"]), datetime.strptime(item["date"], "%d.%m.%Y"))
                        data["files"] = files
                    except Exception:
                        pass
                if files.get("excel") and (files.get("pdf") or files.get("pngs")) and not files.get("status"):
                    files["status"] = "tayyor"
                k = data.setdefault("kpis", {})
                exp = data.get("expired", [])
                k.setdefault("vazn", sum(x.get("vazn", 0) for x in data.get("regimes", [])))
                k.setdefault("expired_value", sum(x.get("qiymat", 0) for x in exp))
                if "summary" not in data:
                    data["summary"] = [{"name": "IBK bo'yicha Jami", "partiya": k.get("partiya", 0), "vazn": k.get("vazn", 0), "qiymat": k.get("qiymat", 0), "tolov": k.get("tolov", 0)}] + [
                        {"name": r.get("rejim", ""), "partiya": r.get("partiya", 0), "vazn": r.get("vazn", 0), "qiymat": r.get("qiymat", 0), "tolov": r.get("tolov", 0)}
                        for r in data.get("regimes", [])
                    ]
                if "expired_summary" not in data:
                    data["expired_summary"] = [{"name": "IBK bo'yicha Jami", "partiya": sum(x.get("partiya", 0) for x in exp), "vazn": sum(x.get("vazn", 0) for x in exp), "qiymat": sum(x.get("qiymat", 0) for x in exp), "tolov": sum(x.get("tolov", 0) for x in exp)}]
                data.setdefault("post_summary", [])
                data.setdefault("warehouse", [])
                data.setdefault("reason", [])
                data.setdefault("own_all", [])
                data.setdefault("own_3m", [])
                data.setdefault("food", [])
                data.setdefault("food_total", {"name": "IBK bo'yicha Jami", "vazn": 0, "qiymat": 0, "over_vazn": 0, "over_qiymat": 0, "ulush": 0})
                rel = data.setdefault("released", {})
                for d in ["1", "3", "5", "10", "30"]:
                    rel.setdefault(d, {"partiya": 0, "vazn": 0, "qiymat": 0, "tolov": 0, "rows": [], "base_date": "", "missing_date": ""})
            self.json(data)
            return
        if parsed.path == "/api/details":
            if not self.require_user():
                return
            q = parse_qs(parsed.query)
            report_id = q.get("report", [""])[0]
            filters = json.loads(q.get("filters", ["{}"])[0])
            item = next((r for r in load_json(INDEX_PATH, {"reports": []})["reports"] if r["id"] == report_id), None)
            if not item:
                self.json({"rows": []})
                return
            self.json({"rows": detail_rows(read_report_source(Path(item["source"])), filters)})
            return
        if parsed.path == "/api/export_details":
            if not self.require_perm("export"):
                return
            q = parse_qs(parsed.query)
            report_id = q.get("report", [""])[0]
            filters = json.loads(q.get("filters", ["{}"])[0])
            item = next((r for r in load_json(INDEX_PATH, {"reports": []})["reports"] if r["id"] == report_id), None)
            if not item:
                self.send_error(404)
                return
            rows = detail_rows(read_report_source(Path(item["source"])), filters)
            headers = [
                {"k": "decl", "t": "Deklaratsiya", "width": 20},
                {"k": "date", "t": "Sana", "width": 14},
                {"k": "regime", "t": "Rejim", "width": 10},
                {"k": "post", "t": "Post", "width": 24},
                {"k": "stir", "t": "STIR", "width": 14},
                {"k": "company", "t": "Korxona", "width": 42},
                {"k": "hs", "t": "TIF TN", "width": 14},
                {"k": "goods", "t": "Tovar", "width": 46},
                {"k": "partiya", "t": "Partiya", "i": True, "width": 10},
                {"k": "vazn", "t": "Vazn (tn)", "n": True, "width": 14},
                {"k": "qiymat", "t": "Qiymat (ming $)", "n": True, "width": 18},
                {"k": "tolov", "t": "To'lov (mln so'm)", "n": True, "width": 18},
                {"k": "reason", "t": "Saqlanish sababi", "width": 30},
            ]
            xlsx = make_xlsx(headers, rows, "Asos deklaratsiyalar")
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", "attachment; filename*=UTF-8''asos_deklaratsiyalar.xlsx")
            self.send_header("Content-Length", str(len(xlsx)))
            self.end_headers()
            self.wfile.write(xlsx)
            return
        if parsed.path == "/api/release":
            if not self.require_perm("release"):
                return
            q = parse_qs(parsed.query)
            base_date = q.get("base", [""])[0]
            final_date = q.get("final", [""])[0]
            base_item = find_report_by_date(base_date)
            final_item = find_report_by_date(final_date)
            missing = []
            if not base_item:
                missing.append(base_date or "boshlang'ich sana")
            if not final_item:
                missing.append(final_date or "yakuniy sana")
            if missing:
                self.json({"missing": missing, "rows": [], "total": {}})
                return
            base_snapshot = STORE.snapshot_id_by_date(base_date) if STORE else None
            final_snapshot = STORE.snapshot_id_by_date(final_date) if STORE else None
            if base_snapshot and final_snapshot:
                result = STORE.compute_released(base_snapshot, final_snapshot)
            else:
                result = release_company_table(Path(base_item["source"]), Path(final_item["source"]))
            result.update({"base": base_date, "final": final_date})
            self.json(result)
            return
        if parsed.path == "/api/export":
            if not self.require_perm("export"):
                return
            q = parse_qs(parsed.query)
            kind = q.get("kind", [""])[0]
            report_id = q.get("report", [""])[0]
            headers = []
            rows = []
            title = "IBK Dashboard jadval"
            if kind == "release":
                base_date = q.get("base", [""])[0]
                final_date = q.get("final", [""])[0]
                base_item = find_report_by_date(base_date)
                final_item = find_report_by_date(final_date)
                if not base_item or not final_item:
                    self.send_error(404)
                    return
                base_snapshot = STORE.snapshot_id_by_date(base_date) if STORE else None
                final_snapshot = STORE.snapshot_id_by_date(final_date) if STORE else None
                if base_snapshot and final_snapshot:
                    data = STORE.compute_released(base_snapshot, final_snapshot)
                else:
                    data = release_company_table(Path(base_item["source"]), Path(final_item["source"]))
                rows = [data["total"]] + data["rows"]
                title = f"Nazoratdan yechilishi {base_date} - {final_date}"
                headers = release_headers()
            else:
                item = next((r for r in load_json(INDEX_PATH, {"reports": []})["reports"] if r["id"] == report_id), None)
                if not item:
                    self.send_error(404)
                    return
                data = load_json(Path(item["dir"]) / "dashboard.json", {})
                if kind == "top_value":
                    rows, headers, title = data.get("top_value", []), company_headers(), "TOP 20 korxona qiymat"
                elif kind == "top_deposit":
                    rows, headers, title = data.get("top_deposit", []), company_headers(), "TOP 20 korxona depozit"
                elif kind == "expired":
                    rows, headers, title = data.get("expired", []), expired_headers(), "Muddati o'tgan"
                elif kind == "goods":
                    rows, headers, title = data.get("goods", []), goods_headers(), "Tovar guruhlari"
                elif kind == "food":
                    rows, headers, title = [data.get("food_total", {})] + data.get("food", []), food_headers(), "Oziq-ovqat"
                else:
                    rows, headers, title = data.get("top_value", []), company_headers(), "Jadval"
            xlsx = make_xlsx(headers, rows, title)
            filename = re.sub(r"[^\w.-]+", "_", title, flags=re.UNICODE) + ".xlsx"
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{filename}")
            self.send_header("Content-Length", str(len(xlsx)))
            self.end_headers()
            self.wfile.write(xlsx)
            return
        if parsed.path.startswith("/download/tolov/"):
            if not self.require_perm("export"):
                return
            name = Path(unquote(parsed.path.split("/download/tolov/", 1)[1])).name
            if name == "_all":
                import zipfile
                buf = BytesIO()
                with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for filename in TOLOV_FILES:
                        path = TOLOV_OUTPUT_DIR / filename
                        if path.exists() and path.is_file():
                            zf.write(path, arcname=filename)
                data = buf.getvalue()
                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", "application/zip")
                self.send_header("Content-Disposition", "attachment; filename*=UTF-8''tolov_jadvallari.zip")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
                return
            if name not in TOLOV_FILES:
                self.send_error(404)
                return
            path = TOLOV_OUTPUT_DIR / name
            if not path.exists() or not path.is_file():
                self.send_error(404)
                return
            data = path.read_bytes()
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(path.name)}")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        if parsed.path.startswith("/download/"):
            if not self.require_perm("export"):
                return
            _, _, report_id, filename = parsed.path.split("/", 3)
            item = next((r for r in load_json(INDEX_PATH, {"reports": []})["reports"] if r["id"] == report_id), None)
            if not item:
                self.send_error(404)
                return
            if Path(unquote(filename)).name == "_pngs":
                import zipfile
                report_dir = Path(item["dir"])
                data_path = report_dir / "dashboard.json"
                dashboard = load_json(data_path, {})
                png_names = dashboard.get("files", {}).get("pngs", [])
                buf = BytesIO()
                with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for name in png_names:
                        p = report_dir / Path(name).name
                        if p.exists():
                            zf.write(p, arcname=p.name)
                data = buf.getvalue()
                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", "application/zip")
                self.send_header("Content-Disposition", "attachment; filename*=UTF-8''IBK_tahlil_pnglar.zip")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
                return
            path = Path(item["dir"]) / Path(unquote(filename)).name
            if not path.exists():
                self.send_error(404)
                return
            data = path.read_bytes()
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/octet-stream")
            self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{path.name}")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        self.send_error(404)

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/login":
            data = self.body_json()
            users = load_json(USER_PATH, {})
            rec = users.get(data.get("user", ""))
            if not rec or rec["password"] != hash_password(data.get("pass", ""), rec["salt"]):
                self.json({"error": "Login yoki parol xato"}, HTTPStatus.UNAUTHORIZED)
                return
            if not rec.get("enabled", True):
                self.json({"error": "Foydalanuvchiga ruxsat berilmagan"}, HTTPStatus.FORBIDDEN)
                return
            token = secrets.token_urlsafe(32)
            SESSIONS[token] = {"user": data["user"], "created": time.time()}
            save_json(SESSION_PATH, SESSIONS)
            self.json({"token": token, "user": {"user": data["user"], "role": rec.get("role", "user"), "role_label": rec.get("role_label", ROLE_LABELS.get(rec.get("role", "user"), "Foydalanuvchi")), "post_code": rec.get("post_code", ""), "full_name": rec.get("full_name", ""), "position": rec.get("position", ""), "phone": rec.get("phone", ""), "lang": rec.get("lang", "uz"), "perms": rec.get("perms", [])}})
            return
        if parsed.path == "/api/users":
            if not self.require_admin():
                return
            data = self.body_json()
            user = re.sub(r"[^\w.@-]+", "", data.get("user", "")).strip()
            password = data.get("pass", "")
            role = data.get("role", "user") if data.get("role") in ROLES else "user"
            perms = [p for p in data.get("perms", DEFAULT_PERMS) if p in ADMIN_PERMS]
            if not user or not password:
                self.json({"error": "Login va parol kerak"}, HTTPStatus.BAD_REQUEST)
                return
            users = load_json(USER_PATH, {})
            salt = secrets.token_hex(8)
            users[user] = {"salt": salt, "password": hash_password(password, salt), "role": role, "role_label": ROLE_LABELS.get(role, "Foydalanuvchi"), "enabled": True, "full_name": data.get("full_name", ""), "position": data.get("position", ""), "phone": data.get("phone", ""), "lang": data.get("lang", "uz"), "post_code": data.get("post_code", ""), "perms": ADMIN_PERMS if role == "admin" else perms}
            save_json(USER_PATH, users)
            self.json({"ok": True})
            return
        if parsed.path == "/api/users/delete":
            admin = self.require_admin()
            if not admin:
                return
            data = self.body_json()
            user = data.get("user", "")
            users = load_json(USER_PATH, {})
            if user == admin:
                self.json({"error": "Admin o'zini o'chira olmaydi"}, HTTPStatus.BAD_REQUEST)
                return
            if user in users:
                users[user]["enabled"] = False
                save_json(USER_PATH, users)
            self.json({"ok": True})
            return
        if parsed.path == "/api/users/update":
            if not self.require_admin():
                return
            data = self.body_json()
            user = re.sub(r"[^\w.@-]+", "", data.get("user", "")).strip()
            users = load_json(USER_PATH, {})
            if not user or user not in users:
                self.json({"error": "Hodim topilmadi"}, HTTPStatus.NOT_FOUND)
                return
            role = data.get("role", users[user].get("role", "user"))
            role = role if role in ROLES else "user"
            perms = [p for p in data.get("perms", users[user].get("perms", DEFAULT_PERMS)) if p in ADMIN_PERMS]
            rec = users[user]
            rec.update({
                "role": role,
                "role_label": ROLE_LABELS.get(role, "Foydalanuvchi"),
                "enabled": bool(data.get("enabled", rec.get("enabled", True))),
                "full_name": data.get("full_name", rec.get("full_name", "")),
                "position": data.get("position", rec.get("position", "")),
                "phone": data.get("phone", rec.get("phone", "")),
                "lang": data.get("lang", rec.get("lang", "uz")),
                "post_code": data.get("post_code", rec.get("post_code", "")),
                "perms": ADMIN_PERMS if role == "admin" else perms,
            })
            password = data.get("pass", "")
            if password:
                rec["salt"] = secrets.token_hex(8)
                rec["password"] = hash_password(password, rec["salt"])
            users[user] = rec
            save_json(USER_PATH, users)
            self.json({"ok": True})
            return
        if parsed.path == "/api/artifacts":
            if not self.require_perm("export"):
                return
            data = self.body_json()
            report_id = data.get("report", "")
            item = next((r for r in load_json(INDEX_PATH, {"reports": []})["reports"] if r["id"] == report_id), None)
            if not item:
                self.json({"error": "Hisobot topilmadi"}, HTTPStatus.NOT_FOUND)
                return
            job_id = "artifacts_" + report_id + "_" + str(int(time.time() * 1000))
            JOBS[job_id] = {"status": "navbatda", "data": {}}
            threading.Thread(target=run_artifact_job, args=(job_id, report_id), daemon=True).start()
            self.json({"job_id": job_id})
            return
        if parsed.path == "/api/tolov":
            if not self.require_perm("upload"):
                return
            length = int(self.headers.get("Content-Length", "0") or "0")
            form = parse_multipart(self.headers.get("Content-Type", ""), self.rfile.read(length))
            if "source" not in form or not form["source"]["filename"]:
                self.json({"error": "To'lov baza fayli kerak"}, HTTPStatus.BAD_REQUEST)
                return
            tmp = UPLOAD_DIR / ("tolov_" + str(int(time.time() * 1000)))
            tmp.mkdir(parents=True, exist_ok=True)
            source = tmp / safe_name(form["source"]["filename"])
            source.write_bytes(form["source"]["content"])
            try:
                rows = build_tolov_from_source(source)
            except Exception as exc:
                self.json({"error": f"{type(exc).__name__}: {exc}"}, HTTPStatus.INTERNAL_SERVER_ERROR)
                return
            self.json({"ok": True, "payments": rows})
            return
        if parsed.path == "/api/reports_bulk":
            if not self.require_perm("upload"):
                return
            length = int(self.headers.get("Content-Length", "0") or "0")
            form = parse_multipart(self.headers.get("Content-Type", ""), self.rfile.read(length))
            raw_sources = form.get("sources") or form.get("source") or []
            sources = raw_sources if isinstance(raw_sources, list) else [raw_sources]
            sources = [x for x in sources if x and x.get("filename")]
            if not sources:
                self.json({"error": "Asos fayllar kerak"}, HTTPStatus.BAD_REQUEST)
                return
            deposit_item = form.get("deposit")
            job_ids = []
            skipped = []
            for idx, src in enumerate(sources):
                report_date = report_date_from_name(src["filename"])
                dup = duplicate_report(src["filename"], report_date)
                if dup:
                    skipped.append({"filename": src["filename"], "date": fmt_date(report_date), "reason": "Bu fayl avval yuklangan"})
                    continue
                job_id = report_date.strftime("%Y%m%d_") + str(int(time.time() * 1000)) + f"_{idx}"
                tmp = UPLOAD_DIR / job_id
                tmp.mkdir(parents=True, exist_ok=True)
                source = tmp / safe_name(src["filename"])
                source.write_bytes(src["content"])
                deposit = None
                if isinstance(deposit_item, dict) and deposit_item.get("filename"):
                    deposit = tmp / safe_name(deposit_item["filename"])
                    deposit.write_bytes(deposit_item["content"])
                JOBS[job_id] = {"status": "navbatda", "data": None}
                threading.Thread(target=run_job, args=(job_id, source, deposit, report_date), daemon=True).start()
                job_ids.append(job_id)
            self.json({"job_ids": job_ids, "count": len(job_ids), "skipped": skipped})
            return
        if parsed.path == "/api/deposit":
            if not self.require_perm("upload"):
                return
            length = int(self.headers.get("Content-Length", "0") or "0")
            form = parse_multipart(self.headers.get("Content-Type", ""), self.rfile.read(length))
            report_id = (form.get("report_id") or {}).get("content", b"").decode("utf-8", errors="ignore").strip()
            if not report_id:
                self.json({"error": "report_id kerak"}, HTTPStatus.BAD_REQUEST)
                return
            if "deposit" not in form or not form["deposit"]["filename"]:
                self.json({"error": "Depozit fayl kerak"}, HTTPStatus.BAD_REQUEST)
                return
            archive = load_json(INDEX_PATH, {"reports": []})
            if not isinstance(archive, dict):
                archive = {"reports": []}
            reports = archive.get("reports") or []
            item = next((r for r in reports if isinstance(r, dict) and r.get("id") == report_id), None)
            if not item:
                self.json({"error": "Hisobot topilmadi"}, HTTPStatus.NOT_FOUND)
                return
            report_dir = Path(item.get("dir", ""))
            if not report_dir.exists():
                self.json({"error": "Hisobot papkasi topilmadi"}, HTTPStatus.NOT_FOUND)
                return
            deposit_path = report_dir / safe_name(form["deposit"]["filename"])
            deposit_path.write_bytes(form["deposit"]["content"])
            item["deposit"] = str(deposit_path)
            save_json(INDEX_PATH, archive)
            job_id = "deposit_" + str(int(time.time() * 1000))
            JOBS[job_id] = {"status": "navbatda"}
            def _rebuild_with_deposit(jid=job_id, it=dict(item), dp=deposit_path):
                j = ensure_job(jid)
                try:
                    j["status"] = "Depozit qayta ishlanmoqda"
                    src = Path(it["source"])
                    rd = datetime.strptime(it["date"], "%d.%m.%Y")
                    data = build_dashboard(it["id"], src, dp, rd)
                    rdir = Path(it["dir"])
                    data_path = rdir / "dashboard.json"
                    existing = load_json(data_path, {})
                    data["files"] = existing.get("files", {})
                    save_json(data_path, data)
                    j.update({"status": "tayyor", "data": data})
                except Exception as exc:
                    j["status"] = "xatolik"
                    j["error"] = f"{type(exc).__name__}: {exc}\n{traceback.format_exc()}"
            threading.Thread(target=_rebuild_with_deposit, daemon=True).start()
            self.json({"job_id": job_id, "report_id": report_id})
            return
        if parsed.path == "/api/reports":
            if not self.require_perm("upload"):
                return
            length = int(self.headers.get("Content-Length", "0") or "0")
            form = parse_multipart(self.headers.get("Content-Type", ""), self.rfile.read(length))
            if "source" not in form or not form["source"]["filename"]:
                self.json({"error": "Asos fayl kerak"}, HTTPStatus.BAD_REQUEST)
                return
            report_date = datetime.strptime(form.get("date", {}).get("content", b"").decode("utf-8") or "1900-01-01", "%Y-%m-%d") if form.get("date", {}).get("content") else report_date_from_name(form["source"]["filename"])
            dup = duplicate_report(form["source"]["filename"], report_date)
            if dup:
                self.json({"error": f"Bu fayl avval yuklangan: {fmt_date(report_date)} / {form['source']['filename']}"}, HTTPStatus.CONFLICT)
                return
            job_id = report_date.strftime("%Y%m%d_") + str(int(time.time() * 1000))
            tmp = UPLOAD_DIR / job_id
            tmp.mkdir(parents=True, exist_ok=True)
            source = tmp / safe_name(form["source"]["filename"])
            source.write_bytes(form["source"]["content"])
            deposit = None
            if "deposit" in form and form["deposit"]["filename"]:
                deposit = tmp / safe_name(form["deposit"]["filename"])
                deposit.write_bytes(form["deposit"]["content"])
            JOBS[job_id] = {"status": "navbatda", "data": {}}
            threading.Thread(target=run_job, args=(job_id, source, deposit, report_date), daemon=True).start()
            self.json({"job_id": job_id})
            return
        self.send_error(404)

    def log_message(self, fmt, *args):
        print(f"[IBK] {self.address_string()} {fmt % args}")


def main():
    global STORE
    ensure_dirs()
    STORE = IBKStore(DB_PATH)
    SESSIONS.update(load_json(SESSION_PATH, {}))
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    print(f"IBK_Dashboard: http://localhost:{PORT}", flush=True)
    server.serve_forever()


if __name__ == "__main__":
    main()













