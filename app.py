import io
import json
import os
import re
import shutil
import signal
import sys
import threading
import time
import traceback
import urllib.request
import uuid
import webbrowser
from dataclasses import asdict
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

import keepa_screener as ks

BASE = Path(__file__).resolve().parent
DATA = Path(os.environ.get("KEEPA_DESK_DATA") or BASE / "data")
RUNS = DATA / "runs"
UPLOADS = DATA / "uploads"
PAGE = BASE / "ui.html"
HOST = "127.0.0.1"
PORT = int(os.environ.get("KEEPA_DESK_PORT", "8765"))
MAX_UPLOAD = 200 * 1024 * 1024
EXCEL = (".xlsx", ".xlsm", ".xls")
SUFFIXES = (".csv", ".txt") + EXCEL
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
RUN_ID = r"(\d{8}-\d{6}-[0-9a-f]{4})"
UPLOAD_ID = re.compile(r"^[0-9a-f]{32}$")

LIVE = {}
LOCK = threading.Lock()


class ApiError(Exception):
    def __init__(self, message, status=400):
        super().__init__(message)
        self.message, self.status = message, status


def write_json(path, data):
    temp = path.with_name(path.name + ".tmp")
    with open(temp, "w", encoding="utf-8") as handle:
        handle.write(json.dumps(data, default=str))
        handle.flush()
        try:
            os.fsync(handle.fileno())
        except OSError:
            pass
    for _ in range(5):
        try:
            temp.replace(path)
            return
        except PermissionError:
            time.sleep(0.05)


def read_json(path):
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    return data if isinstance(data, dict) else None


def clean(value):
    if isinstance(value, float) and (value != value or value in (float("inf"), float("-inf"))):
        return None
    return value


def read_rows(path, pos, limit=None):
    rows = []
    try:
        with open(path, "rb") as handle:
            handle.seek(pos)
            while limit is None or len(rows) < limit:
                line = handle.readline()
                if not line.endswith(b"\n"):
                    break
                pos += len(line)
                try:
                    rows.append(json.loads(line))
                except ValueError:
                    continue
    except FileNotFoundError:
        pass
    return rows, pos


def tail(path, count=200):
    try:
        with open(path, "rb") as handle:
            handle.seek(0, os.SEEK_END)
            size = handle.tell()
            handle.seek(max(0, size - 64 * 1024))
            data = handle.read()
    except OSError:
        return []
    lines = data.decode("utf-8", errors="replace").splitlines()
    if size > 64 * 1024 and lines:
        lines = lines[1:]  # first line is probably cut in half
    return lines[-count:]


def safe_name(name):
    return re.sub(r"[^A-Za-z0-9._ ()-]+", "_", name).strip(" .")[:120] or "file"


class Run:
    def __init__(self, rid, meta, key=""):
        self.id, self.meta, self.key = rid, meta, key
        self.folder = RUNS / rid
        self.stop = threading.Event()
        self.pause = threading.Event()
        self.lock = threading.Lock()
        self.wait_reason = None
        self.wait_until = None
        self.saved_at = 0.0
        self.rows_file = self.log_file = None

    def _save(self, force=False):
        if force or time.time() - self.saved_at >= 1.0:
            write_json(self.folder / "run.json", self.meta)
            self.saved_at = time.time()

    def snapshot(self):
        with self.lock:
            reason = self.wait_reason
            left = max(0, round(self.wait_until - time.time())) if self.wait_until else None
            return dict(self.meta, paused=self.pause.is_set(), waiting=reason == "tokens",
                        wait_reason=reason, wait_left=left)

    def on_wait(self, reason, until=None):
        with self.lock:
            self.wait_reason, self.wait_until = reason, until

    def update(self, force=False, **fields):
        with self.lock:
            self.meta.update(fields)
            self._save(force)

    def scrub(self, text):
        return text.replace(self.key, "***") if self.key else text

    def log(self, message):
        message = self.scrub(str(message))
        with self.lock:
            if self.log_file and not self.log_file.closed:
                self.log_file.write(f"{time.strftime('%H:%M:%S')}  {message}\n")
                self.log_file.flush()

    def on_row(self, row):
        self.rows_file.write(json.dumps({k: clean(v) for k, v in row.items()}, default=str) + "\n")
        self.rows_file.flush()
        try:
            os.fsync(self.rows_file.fileno())
        except OSError:
            pass
        with self.lock:
            self.meta["checked"] += 1
            if row.get("status") == "PASS":
                self.meta["passed"] += 1
            self._save()

    def progress(self, label, done, total):
        self.update(label=label, done=done, total=total)

    def tokens(self, left):
        self.update(tokens_left=left)


def work(run, cfg, planned, preface):
    state, error, fields = "done", "", {}
    awake = KeepAwake()
    awake.on()
    try:
        run.rows_file = open(run.folder / "rows.jsonl", "a", encoding="utf-8")
        run.log_file = open(run.folder / "log.txt", "a", encoding="utf-8")
        for line in preface:
            run.log(line)
        result = ks.run_screening(
            cfg["path"], cfg["rules"], cfg["api_key"], cfg["domain_name"], cfg["sheet"], cfg["overrides"],
            out=str(run.folder / "report.xlsx"), limit=cfg["limit"], buybox=cfg["buybox"],
            refresh=cfg["refresh"], log=run.log, cost_basis=cfg["cost_basis"], two_pass=cfg["two_pass"],
            stop=run.stop.is_set, progress=run.progress, on_row=run.on_row, prefilter=cfg["prefilter"],
            cache_only=cfg["cache_only"], on_tokens=run.tokens, name=cfg["name"], planned=planned,
            pause=run.pause.is_set, on_wait=run.on_wait)
        fields = {"report": Path(result["out"]).name, "passed": result["passed"], "checked": result["total"]}
        if result["cancelled"]:
            state = "cancelled"
        elif result.get("error"):
            state, error = "error", run.scrub(result["error"])
    except ks.Cancelled:
        state = "cancelled"
    except Exception as exc:
        if not isinstance(exc, (ValueError, RuntimeError)):
            traceback.print_exc()
        error = run.scrub(str(exc)) or exc.__class__.__name__
        state = "error"
        if run.log_file:
            run.log(f"Error: {error}")
    finally:
        awake.off()
        run.pause.clear()
        run.on_wait(None)
        with run.lock:
            for handle in (run.rows_file, run.log_file):
                if handle:
                    handle.close()
        run.update(force=True, state=state, error=error, **fields)
        with LOCK:
            LIVE.pop(run.id, None)


def run_folder(rid):
    folder = RUNS / rid
    if not folder.is_dir():
        raise ApiError("Run not found", 404)
    return folder


def read_meta(rid):
    run = LIVE.get(rid)
    if run:
        return run.snapshot()
    meta = read_json(run_folder(rid) / "run.json")
    if meta is None:
        raise ApiError("Run not found", 404)
    return meta


def mark_interrupted():
    for folder in RUNS.glob("*"):
        meta = read_json(folder / "run.json")
        if meta is None and folder.is_dir():
            meta = read_json(folder / "start.json")
            if meta:
                meta["state"] = "running"
        if meta and meta.get("state") in ("running", "paused"):
            meta["state"] = "interrupted"
            rows, _ = read_rows(folder / "rows.jsonl", 0)
            meta["checked"] = len(rows)
            meta["passed"] = sum(1 for r in rows if r.get("status") == "PASS")
            write_json(folder / "run.json", meta)


def list_runs():
    out = []
    for folder in sorted(RUNS.glob("*"), reverse=True):
        try:
            meta = read_meta(folder.name)
        except ApiError:
            continue
        state = meta.get("state")
        if state == "running" and meta.get("paused"):
            state = "paused"
        out.append({"id": folder.name, "name": meta.get("name"), "domain": meta.get("domain"),
                    "state": state, "passed": meta.get("passed", 0),
                    "checked": meta.get("checked", 0), "created": meta.get("created"),
                    "report": bool(meta.get("report")) or meta.get("checked", 0) > 0})
    return out


def feed(rid, pos):
    meta = read_meta(rid)
    folder = run_folder(rid)
    rows, pos = read_rows(folder / "rows.jsonl", pos, limit=500)
    return {"pos": pos, "rows": [{k: r.get(k) for k in ks.LIVE_KEYS} for r in rows],
            "state": meta.get("state"), "label": meta.get("label", ""), "done": meta.get("done", 0),
            "total": meta.get("total", 0), "checked": meta.get("checked", 0), "passed": meta.get("passed", 0),
            "tokens_left": meta.get("tokens_left"), "waiting": bool(meta.get("waiting")),
            "paused": bool(meta.get("paused")), "wait_reason": meta.get("wait_reason"),
            "wait_left": meta.get("wait_left"),
            "report": bool(meta.get("report")), "partial": meta.get("checked", 0) > 0 or bool(rows) or pos > 0,
            "error": meta.get("error", ""), "log": tail(folder / "log.txt")}


def report_bytes(rid):
    folder = run_folder(rid)
    meta = read_meta(rid)
    final = folder / "report.xlsx"
    if meta.get("report") and final.exists():
        return final.read_bytes()
    rows, _ = read_rows(folder / "rows.jsonl", 0)
    if not rows:
        raise ApiError("Nothing to export yet.", 404)
    rows.sort(key=lambda r: (r.get("source_row") or 0))
    rules = ks.Rules(**meta["rules"])
    pf = meta.get("prefilter")
    prefilter = ks.PreFilter(**{**pf, "exclude_words": tuple(pf.get("exclude_words") or ())}) if pf else None
    buffer = io.BytesIO()
    ks.write_report(rows, rules, meta["domain"], buffer, prefilter)
    return buffer.getvalue()


def upload_path(upload_id):
    if not isinstance(upload_id, str) or not UPLOAD_ID.match(upload_id):
        raise ApiError("Choose a supplier file first.")
    found = sorted(UPLOADS.glob(upload_id + "__*"))
    if not found:
        raise ApiError("That uploaded file is no longer on disk. Choose it again.")
    return found[0]


def read_catalog(path, sheet):
    try:
        return ks.load_catalog(path, sheet or None)
    except ValueError as exc:
        raise ApiError(str(exc))
    except Exception as exc:
        raise ApiError(f"Could not read that file: {exc}")


def inspect(upload_id, sheet=None):
    path = upload_path(upload_id)
    sheets = []
    if path.suffix.lower() in EXCEL:
        try:
            sheets = ks.excel_sheets(path)
            sheet = sheet if sheet in sheets else ks.best_sheet(path)
        except ValueError as exc:
            raise ApiError(str(exc))
        except Exception as exc:
            raise ApiError(f"Could not read that file: {exc}")
    df = read_catalog(path, sheet if sheets else None)
    mapping, remembered = ks.suggest_mapping(df)
    preview = [[ks._s(v) or "" for v in row] for row in df.head(8).itertuples(index=False, name=None)]
    return {"upload": upload_id, "name": path.name.split("__", 1)[1], "sheets": sheets, "sheet": sheet or "",
            "columns": list(df.columns), "mapping": {f: mapping.get(f) or "" for f in ks.FIELDS},
            "cost_basis": mapping.get("cost_basis") or "auto", "remembered": remembered,
            "rows": len(df), "header": df.attrs["header_row"], "preview": preview}


def save_upload(handler):
    name = Path(unquote(handler.headers.get("X-Filename", ""))).name
    if Path(name).suffix.lower() not in SUFFIXES:
        raise ApiError("Use a .csv, .txt, .xlsx, .xlsm or .xls file.")
    try:
        length = int(handler.headers.get("Content-Length") or 0)
    except ValueError:
        length = 0
    if length <= 0:
        raise ApiError("That file is empty.")
    if length > MAX_UPLOAD:
        raise ApiError("That file is larger than 200 MB.", 413)
    uid = uuid.uuid4().hex
    target = UPLOADS / f"{uid}__{safe_name(name)}"
    left = length
    with open(target, "wb") as handle:
        while left:
            chunk = handler.rfile.read(min(65536, left))
            if not chunk:
                break
            handle.write(chunk)
            left -= len(chunk)
    if left:
        target.unlink()
        raise ApiError("The upload was cut off. Try again.")
    try:
        return inspect(uid)
    except Exception:
        target.unlink()
        raise


def number(b, key, label, default, lo=None, hi=None):
    raw = b.get(key)
    if raw is None or (isinstance(raw, str) and not raw.strip()):
        return default
    try:
        value = float(raw)
    except (TypeError, ValueError):
        raise ApiError(f"{label} must be a number.")
    if value != value or value in (float("inf"), float("-inf")):
        raise ApiError(f"{label} must be a number.")
    if lo is not None and value < lo:
        raise ApiError(f"{label} must be at least {lo:g}.")
    if hi is not None and value > hi:
        raise ApiError(f"{label} must be at most {hi:g}.")
    return value


def parse(b):
    path = upload_path(b.get("upload"))
    domain_name = b.get("domain") or "GB"
    if domain_name not in ks.DOMAINS:
        raise ApiError("Unknown marketplace.")
    cache_only = bool(b.get("cache_only"))
    api_key = str(b.get("api_key") or "").strip()
    if not api_key and not cache_only:
        raise ApiError("Enter your Keepa API key, or tick cache-only.")
    d = ks.Rules()
    checks = b.get("checks") or {}
    basis = b.get("basis") or d.price_basis
    if basis not in ("conservative", "current", "avg90"):
        raise ApiError("Unknown sell price option.")
    rules = ks.Rules(
        min_monthly_sales=int(number(b, "min_sales", "Minimum sales", d.min_monthly_sales, 0)),
        bsr_max=int(number(b, "bsr_max", "Sales rank limit", d.bsr_max, 1)),
        bsr_max_cv=number(b, "bsr_cv", "Rank variation", d.bsr_max_cv, 0),
        min_offers=int(number(b, "min_offers", "Minimum offers", d.min_offers, 0)),
        max_offers=int(number(b, "max_offers", "Maximum offers", d.max_offers, 1)),
        amazon_max_instock_pct=number(b, "amazon_pct", "Amazon in-stock %", d.amazon_max_instock_pct, 0, 100),
        min_roi=number(b, "min_roi", "Minimum ROI", d.min_roi),
        max_suppressed_pct=number(b, "max_suppressed", "Buy Box suppression %", d.max_suppressed_pct, 0, 100),
        price_basis=basis,
        referral_fee_pct=number(b, "referral", "Referral fee", d.referral_fee_pct, 0, 100),
        fba_fee_default=number(b, "fba", "FBA fee", d.fba_fee_default, 0),
        prep_cost=number(b, "prep", "Prep cost", d.prep_cost, 0),
        inbound_ship=number(b, "inbound", "Inbound shipping", d.inbound_ship, 0),
        vat_rate=number(b, "vat", "VAT", d.vat_rate * 100, 0, 100) / 100,
        auto_pack=bool(b.get("auto_pack", False)),
        **{f"check_{c}": bool(checks.get(c, True)) for c in ks.CHECKS})
    if rules.check_offers and rules.min_offers > rules.max_offers:
        raise ApiError("Minimum offers is higher than maximum offers.")
    pf = b.get("prefilter") or {}
    words = tuple(w.strip() for w in str(pf.get("exclude") or "").split(",") if w.strip())
    prefilter = ks.PreFilter(min_qty=number(pf, "min_qty", "Minimum QTY", None),
                             min_cost=number(pf, "min_cost", "Minimum cost", None),
                             max_cost=number(pf, "max_cost", "Maximum cost", None),
                             exclude_words=words, qty_column=pf.get("qty_column") or None)
    if (prefilter.min_cost is not None and prefilter.max_cost is not None
            and prefilter.min_cost > prefilter.max_cost):
        raise ApiError("Minimum cost is higher than maximum cost.")
    mapping = b.get("map") or {}
    cost_basis = b.get("cost_basis") or "auto"
    if cost_basis not in ("auto", "unit", "case"):
        raise ApiError("Unknown cost basis.")
    return {"path": path, "name": Path(path.name.split("__", 1)[1]).stem, "rules": rules,
            "api_key": api_key, "domain_name": domain_name, "sheet": b.get("sheet") or None,
            "overrides": {f: (mapping.get(f) or "") for f in ks.FIELDS}, "cost_basis": cost_basis,
            "prefilter": prefilter if prefilter.active() else None, "cache_only": cache_only,
            "limit": int(number(b, "limit", "Row limit", 0, 0)) or None,
            "buybox": bool(b.get("buybox", True)), "two_pass": bool(b.get("two_pass", True)),
            "refresh": bool(b.get("refresh"))}


def make_plan(cfg):
    lines = []
    try:
        planned = ks.plan(cfg["path"], cfg["rules"], cfg["sheet"], cfg["overrides"], cfg["cost_basis"],
                          cfg["prefilter"], cfg["limit"], lines.append)
    except ValueError as exc:
        raise ApiError(str(exc))
    except Exception as exc:
        raise ApiError(f"Could not read that file: {exc}")
    return planned, lines


def preflight(planned, cfg, tokens, out):
    rules = cfg["rules"]
    active = [i for i in planned.items if i["source_row"] not in planned.skipped_by_filter]
    total = len(planned.items)
    bad_code = sum(1 for i in active if i["bad_code"] and not i["asin"])
    no_id = max(0, planned.no_id - bad_code)
    no_cost = sum(1 for i in active if i["cost"] is None or i["cost"] <= 0)
    multipack = 0
    for i in active:
        for pattern in ks.PACK_PATTERNS:
            m = pattern.search(i["title"] or "")
            if m and int(m.group(1)) > 1:
                multipack += 1
                break
    with_id = sum(1 for i in active if i["asin"] or i["code"])
    repeats = max(0, with_id - out["lookups"])
    facts = []
    for count, one, many in ((no_id, "row has no barcode or ASIN and will be skipped", "rows have no barcode or ASIN and will be skipped"),
                             (bad_code, "row has a barcode that fails the check digit and will be skipped", "rows have a barcode that fails the check digit and will be skipped"),
                             (no_cost, "row has no cost, so ROI cannot be checked for it", "rows have no cost, so ROI cannot be checked for them"),
                             (multipack, "name looks like a multipack, check the cost is per unit", "names look like multipacks, check the cost is per unit"),
                             (repeats, "row repeats a product that is already being looked up", "rows repeat a product that is already being looked up")):
        if count:
            facts.append({"n": count, "text": one if count == 1 else many})
    warnings = []
    if rules.min_roi < 5:
        warnings.append(f"Minimum ROI is {rules.min_roi:g}%. Almost everything will pass that rule.")
    if rules.min_roi > 100:
        warnings.append(f"Minimum ROI is {rules.min_roi:g}%. Almost nothing will pass that rule.")
    if rules.vat_rate == 0 and cfg["domain_name"] == "GB":
        warnings.append("VAT is 0% on a UK list. Profit will look higher than it really is.")
    if rules.min_monthly_sales == 0 and rules.check_sales:
        warnings.append("Minimum monthly sales is 0, so the sales rule checks nothing.")
    if not any(getattr(rules, "check_" + c) for c in ks.CHECKS):
        warnings.append("Every rule is switched off. Every product will pass.")
    if total and len(planned.skipped_by_filter) / total > 0.5:
        warnings.append(f"Your input filters skip {len(planned.skipped_by_filter) * 100 // total}% of the rows.")
    if total and planned.no_id / total > 0.3:
        warnings.append("More than 30% of rows have no usable ID. Check the barcode and ASIN column mapping.")
    if cfg["cache_only"] and out["lookups"] and out["cached"] < out["lookups"]:
        missing = out["lookups"] - out["cached"]
        warnings.append(f"Cache only is on, but {missing:,} of {out['lookups']:,} lookups are not cached and will show as not found.")
    out["preflight"] = {"facts": facts, "warnings": warnings}


def estimate(b):
    cfg = parse(b)
    planned, _ = make_plan(cfg)
    cheap = ks.is_cheap(cfg["rules"], cfg["buybox"], cfg["two_pass"])
    hours = 10 ** 6 if cfg["cache_only"] else (0 if cfg["refresh"] else 24)
    first = ks.KeepaClient("", ks.DOMAINS[cfg["domain_name"]], buybox=cfg["buybox"] and not cheap,
                           cache_hours=hours)
    active = [i for i in planned.items if i["source_row"] not in planned.skipped_by_filter]
    asins = {i["asin"] for i in active if i["asin"]}
    codes = {i["code"] for i in active if not i["asin"] and i["code"]}
    cached = (sum(first.is_cached("asin", a) for a in asins)
              + sum(first.is_cached("code", c) for c in codes))
    tokens = 0 if cfg["cache_only"] else (len(asins) + len(codes) - cached) * (3 if first.buybox else 1)
    out = {"rows": len(planned.items), "lookups": len(asins) + len(codes), "cached": cached,
           "no_id": planned.no_id, "filtered": len(planned.skipped_by_filter), "tokens": tokens,
           "cheap": cheap, "cache_only": cfg["cache_only"], "left": None, "rate": None,
           "token_error": None, "wait_min": None}
    if not cfg["cache_only"]:
        try:
            status = ks.token_status(cfg["api_key"])
            out["left"], out["rate"] = status["left"], status["rate"]
            if status["left"] is not None and status["rate"]:
                out["wait_min"] = max(0, tokens - status["left"]) / status["rate"]
        except RuntimeError as exc:
            out["token_error"] = str(exc)
    preflight(planned, cfg, tokens, out)
    return out


def start_run(b):
    cfg = parse(b)
    planned, lines = make_plan(cfg)
    now = time.strftime("%Y-%m-%d %H:%M")
    with LOCK:
        if LIVE:
            raise ApiError("A screening is already running. Stop it or wait for it to finish.", 409)
        rid = time.strftime("%Y%m%d-%H%M%S") + "-" + uuid.uuid4().hex[:4]
        (RUNS / rid).mkdir(parents=True)
        meta = {"id": rid, "name": cfg["name"], "created": now, "state": "running",
                "domain": cfg["domain_name"], "label": "", "done": 0, "total": 0, "checked": 0, "passed": 0,
                "tokens_left": None, "report": "", "error": "",
                "settings": {k: v for k, v in b.items() if k != "api_key"},
                "rules": asdict(cfg["rules"]),
                "prefilter": asdict(cfg["prefilter"]) if cfg["prefilter"] else None}
        run = Run(rid, meta, cfg["api_key"])
        write_json(run.folder / "run.json", meta)
        write_json(run.folder / "start.json", meta)
        LIVE[rid] = run
    threading.Thread(target=work, args=(run, cfg, planned, lines), daemon=True).start()
    return {"id": rid}


def defaults():
    d = ks.Rules()
    return {"domains": list(ks.DOMAINS),
            "rules": {"min_sales": d.min_monthly_sales, "bsr_max": d.bsr_max, "bsr_cv": d.bsr_max_cv,
                      "min_offers": d.min_offers, "max_offers": d.max_offers,
                      "amazon_pct": d.amazon_max_instock_pct, "min_roi": d.min_roi,
                      "max_suppressed": d.max_suppressed_pct, "referral": d.referral_fee_pct,
                      "fba": d.fba_fee_default, "prep": d.prep_cost, "inbound": d.inbound_ship,
                      "vat": round(d.vat_rate * 100, 4), "basis": d.price_basis},
            "checks": ks.CHECKS, "fields": ks.FIELDS}


ROUTES = {"GET": [], "POST": []}


def route(method, pattern):
    def register(func):
        ROUTES[method].append((re.compile(pattern), func))
        return func
    return register


@route("GET", r"/")
def page(h, url):
    try:
        body = PAGE.read_bytes()
    except OSError:
        raise ApiError("ui.html is missing next to app.py", 404)
    h.send(200, body, "text/html; charset=utf-8")


@route("GET", r"/api/ping")
def ping(h, url):
    h.json({"app": "keepa-screener"})


@route("GET", r"/api/defaults")
def get_defaults(h, url):
    h.json(defaults())


@route("GET", r"/api/runs")
def get_runs(h, url):
    h.json({"runs": list_runs()})


@route("GET", rf"/api/run/{RUN_ID}/meta")
def get_meta(h, url, rid):
    run_folder(rid)
    h.json(read_meta(rid))


@route("GET", rf"/api/run/{RUN_ID}/feed")
def get_feed(h, url, rid):
    try:
        pos = max(0, int(parse_qs(url.query).get("pos", ["0"])[0]))
    except ValueError:
        pos = 0
    h.json(feed(rid, pos))


@route("GET", rf"/api/run/{RUN_ID}/report")
def get_report(h, url, rid):
    body = report_bytes(rid)
    name = safe_name(read_meta(rid).get("name") or "report")
    h.send(200, body, XLSX, {"Content-Disposition": f'attachment; filename="screened_{name}.xlsx"'})


@route("POST", r"/api/upload")
def post_upload(h, url):
    h.json(save_upload(h))


@route("POST", r"/api/inspect")
def post_inspect(h, url):
    b = h.body()
    h.json(inspect(b.get("upload"), b.get("sheet") or None))


@route("POST", r"/api/tokens")
def post_tokens(h, url):
    key = str(h.body().get("api_key") or "").strip()
    if not key:
        raise ApiError("Enter your Keepa API key first.")
    try:
        h.json(ks.token_status(key))
    except RuntimeError as exc:
        raise ApiError(str(exc))


@route("POST", r"/api/estimate")
def post_estimate(h, url):
    h.json(estimate(h.body()))


@route("POST", r"/api/start")
def post_start(h, url):
    h.json(start_run(h.body()))


@route("POST", rf"/api/run/{RUN_ID}/cancel")
def post_cancel(h, url, rid):
    run = LIVE.get(rid)
    if run:
        run.stop.set()
        run.pause.clear()
        run.log("Stop requested, saving the report of what is done")
    h.json({"ok": True})


@route("POST", rf"/api/run/{RUN_ID}/pause")
def post_pause(h, url, rid):
    run = LIVE.get(rid)
    if not run:
        raise ApiError("This run is not running.", 409)
    if not run.pause.is_set():
        run.pause.set()
        run.log("Paused. No Keepa requests are sent until you press Resume")
    h.json({"ok": True, "paused": True})


@route("POST", rf"/api/run/{RUN_ID}/resume")
def post_resume(h, url, rid):
    run = LIVE.get(rid)
    if not run:
        raise ApiError("This run is not running.", 409)
    if run.pause.is_set():
        run.pause.clear()
        run.log("Resumed")
    h.json({"ok": True, "paused": False})


@route("POST", rf"/api/run/{RUN_ID}/delete")
def post_delete(h, url, rid):
    folder = run_folder(rid)
    if rid in LIVE:
        raise ApiError("Stop this run before deleting it.", 409)
    shutil.rmtree(folder, ignore_errors=True)
    h.json({"ok": True})


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def send(self, status, body, ctype, headers=None):
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for key, value in (headers or {}).items():
            self.send_header(key, value)
        self.end_headers()
        self.wfile.write(body)

    def json(self, data, status=200):
        self.send(status, json.dumps(data, default=str).encode("utf-8"), "application/json")

    def body(self):
        try:
            length = int(self.headers.get("Content-Length") or 0)
            data = json.loads(self.rfile.read(min(length, 5_000_000)) or b"{}")
        except ValueError:
            raise ApiError("Bad request.")
        if not isinstance(data, dict):
            raise ApiError("Bad request.")
        return data

    def guard(self, method):
        port = self.server.server_address[1]
        if self.headers.get("Host", "") not in (f"127.0.0.1:{port}", f"localhost:{port}"):
            raise ApiError("Forbidden", 403)
        if method == "POST" and self.headers.get("X-Desk") != "1":
            raise ApiError("Forbidden", 403)

    def dispatch(self, method):
        try:
            self.guard(method)
            url = urlparse(self.path)
            for pattern, func in ROUTES[method]:
                found = pattern.fullmatch(url.path)
                if found:
                    return func(self, url, *found.groups())
            raise ApiError("Not found", 404)
        except ApiError as exc:
            self.fail(exc.message, exc.status)
        except (BrokenPipeError, ConnectionResetError):
            pass
        except Exception as exc:
            traceback.print_exc()
            self.fail(f"Unexpected error: {exc}", 500)

    def fail(self, message, status):
        try:
            self.json({"error": message}, status)
        except OSError:
            pass

    def do_GET(self):
        self.dispatch("GET")

    def do_POST(self):
        self.dispatch("POST")


class Server(ThreadingHTTPServer):
    # On Windows SO_REUSEADDR lets a second program bind a port that is already in use,
    # so two copies of the app could silently share 8765. Only reuse addresses elsewhere.
    allow_reuse_address = os.name != "nt"
    daemon_threads = True


def already_running(port):
    try:
        with urllib.request.urlopen(f"http://{HOST}:{port}/api/ping", timeout=2) as reply:
            return json.loads(reply.read()).get("app") == "keepa-screener"
    except Exception:
        return False


def clear_old_uploads(days=30):
    cutoff = time.time() - days * 86400
    for path in UPLOADS.glob("*__*"):
        try:
            if path.stat().st_mtime < cutoff:
                path.unlink()
        except OSError:
            pass


def disable_quick_edit():
    """Windows pauses a console program when you click inside its window (QuickEdit).
    That looks exactly like a frozen screening, so switch it off for this window."""
    if os.name != "nt":
        return
    try:
        import ctypes
        kernel = ctypes.windll.kernel32
        handle = kernel.GetStdHandle(-10)
        mode = ctypes.c_uint32()
        if kernel.GetConsoleMode(handle, ctypes.byref(mode)):
            kernel.SetConsoleMode(handle, (mode.value & ~0x0040) | 0x0080)
    except Exception:
        pass


class KeepAwake:
    """Stops the computer going to sleep while a screening runs."""

    def __init__(self):
        self.proc = None

    def on(self):
        try:
            if os.name == "nt":
                import ctypes
                ctypes.windll.kernel32.SetThreadExecutionState(0x80000000 | 0x00000001)
            elif sys.platform == "darwin" and shutil.which("caffeinate"):
                import subprocess
                self.proc = subprocess.Popen(["caffeinate", "-i", "-w", str(os.getpid())])
        except Exception:
            pass

    def off(self):
        try:
            if os.name == "nt":
                import ctypes
                ctypes.windll.kernel32.SetThreadExecutionState(0x80000000)
            elif self.proc:
                self.proc.terminate()
                self.proc = None
        except Exception:
            pass


def bind():
    """Use PORT, or the next free port if another program has it."""
    for port in range(PORT, PORT + 20):
        if already_running(port):
            return None, port
        try:
            return Server((HOST, port), Handler), port
        except OSError:
            continue
    sys.exit(f"Ports {PORT}-{PORT + 19} are all in use. Close some programs and try again.")


def main():
    RUNS.mkdir(parents=True, exist_ok=True)
    UPLOADS.mkdir(parents=True, exist_ok=True)
    clear_old_uploads()
    disable_quick_edit()
    server, port = bind()
    url = f"http://127.0.0.1:{port}/"
    if server is None:
        print(f"The screener is already running, opening {url}")
        webbrowser.open(url)
        return
    mark_interrupted()
    print("")
    print(f"  Keepa screener is running at {url}")
    print("  Your browser should open by itself. If not, copy that address into it.")
    print("  Keep this window open while you work. Close it (or press Ctrl+C) to quit.")
    print("")
    if not os.environ.get("KEEPA_DESK_NO_BROWSER"):
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    def leave(*args):
        raise KeyboardInterrupt()
    for name in ("SIGTERM", "SIGHUP", "SIGBREAK"):
        sig = getattr(signal, name, None)
        if sig is not None:
            try:
                signal.signal(sig, leave)
            except (ValueError, OSError):
                pass
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        for run in list(LIVE.values()):
            run.stop.set()
            run.pause.clear()
        deadline = time.time() + 8
        while LIVE and time.time() < deadline:
            time.sleep(0.2)
        server.server_close()


if __name__ == "__main__":
    main()
