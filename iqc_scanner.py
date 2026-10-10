import sys
import os
import json
import re
import csv
import threading
import configparser
from datetime import datetime
from pathlib import Path

import pyodbc

# ─────────────────────────────────────────────────────────────────────────────
# Load SQL Server connection string from db_config.ini (same dir as exe/script)
# ─────────────────────────────────────────────────────────────────────────────
def _load_db_connection_string() -> str:
    cfg_path = Path(getattr(sys, '_MEIPASS', Path(__file__).parent)) / "db_config.ini"
    if not cfg_path.exists():
        raise FileNotFoundError(
            f"db_config.ini not found at:\n{cfg_path}\n\n"
            "Create it with [database] section containing:\n"
            "  driver, server, database, trusted (yes/no)\n"
            "  username, password  (only needed if trusted = no)"
        )
    cp = configparser.ConfigParser()
    cp.read(cfg_path, encoding="utf-8")
    sec = cp["database"]
    trusted = sec.get("trusted", "yes").strip().lower() in ("yes", "y", "true", "1")
    if trusted:
        auth = "Trusted_Connection=yes;"
    else:
        auth = f"UID={sec.get('username','')};PWD={sec.get('password','')};"
    return (
        f"DRIVER={{{sec.get('driver', 'ODBC Driver 17 for SQL Server')}}};"
        f"SERVER={sec.get('server', '')};"
        f"DATABASE={sec.get('database', '')};"
        f"{auth}"
    )

DB_CONNECTION_STRING = _load_db_connection_string()
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

try:
    from openpyxl import load_workbook, Workbook
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False

import ctypes
import pathlib, sys

# ─────────────────────────────────────────────────────────────────────────────
# Force Windows to treat this script as a distinct application on the taskbar
# ─────────────────────────────────────────────────────────────────────────────
try:
    myappid = 'iqc.barcode.scanner.v1'  # arbitrary unique string
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except AttributeError:
    pass # Fallback for non-Windows operating systems

# ─────────────────────────────────────────────────────────────────────────────
# Resolve pathname for .ico referencing in same directory
# ─────────────────────────────────────────────────────────────────────────────
def _resource(filename: str) -> str:
    """Resolve a resource path for both dev and PyInstaller frozen builds."""
    base = pathlib.Path(getattr(sys, '_MEIPASS', pathlib.Path(__file__).parent))
    return str(base / filename)


# ──────────────────────────────────────────────
# XLSX CONFIG LOADER
# ──────────────────────────────────────────────

def _cell(ws, row, col):
    """Return stripped string value of a cell, or empty string."""
    v = ws.cell(row=row, column=col).value
    return str(v).strip() if v is not None else ""

def _int(ws, row, col):
    """Return int value of a cell, or None."""
    v = ws.cell(row=row, column=col).value
    try:
        return int(v) if v not in (None, "") else None
    except (ValueError, TypeError):
        return None

def _bool(ws, row, col):
    """Return True if cell contains yes/y/true/1 (case-insensitive)."""
    return _cell(ws, row, col).lower() in ("yes", "y", "true", "1")

def _rules(ws, row,
           col_fmt, col_prefix, col_suffix,
           col_total_len, col_run_len,
           col_yr_start, col_yr_len,
           col_ww_start, col_ww_len,
           col_regex, col_desc):
    """
    Build a barcode-rules dict from a row+column set.
    Identical structure for single-barcodes, panel SNs, and array SNs
    so the validator treats all three exactly the same way.
    """
    fmt       = _cell(ws, row, col_fmt) or "none"
    prefix    = _cell(ws, row, col_prefix)
    suffix    = _cell(ws, row, col_suffix)
    total_len = _int(ws, row, col_total_len)
    run_len   = _int(ws, row, col_run_len)
    yr_start  = _int(ws, row, col_yr_start)
    yr_len    = _int(ws, row, col_yr_len)
    ww_start  = _int(ws, row, col_ww_start)
    ww_len    = _int(ws, row, col_ww_len)
    regex_pat = _cell(ws, row, col_regex)
    desc      = _cell(ws, row, col_desc)
    return {
        "format":                   fmt,
        "prefix":                   prefix,
        "suffix":                   suffix,
        "length":                   total_len,
        "min_length":               total_len,
        "max_length":               total_len,
        "running_number_length":    run_len,
        "year_section_start":       yr_start,
        "year_section_length":      yr_len,
        "workweek_section_start":   ww_start,
        "workweek_section_length":  ww_len,
        "pattern":                  regex_pat,
        "description":              desc or fmt,
    }


def load_config_xlsx(path: Path) -> dict:
    """
    Read config.xlsx and return the dict structure the app expects.

    ── Sheet "App" ──────────────────────────────────────────────────
    Row 1 = header, Row 2 = values:
      A  station
      B  operator_required   (yes / no)
      C  scan_timeout_seconds

    ── Sheet "Projects" ─────────────────────────────────────────────
    Row 1 = header (ignored), data from row 2 onward.
    One row per PCB.  Columns:

    COMMON (all modes)
      A  project_id
      B  project_name
      C  mode                    single | panel
      D  duplicate_check         yes / no
      E  duplicate_scope         project | global

    SINGLE-BARCODE RULES  (also used as panel-SN rules in panel mode)
      F  format       none | prefix_suffix | fixed_length | regex | structured
      G  prefix
      H  suffix
      I  total_length
      J  running_number_length
      K  year_section_start      (1-based char position)
      L  year_section_length
      M  workweek_section_start
      N  workweek_section_length
      O  regex_pattern
      P  description

    PANEL-MODE EXTRAS
      Q  array_count
      R  array_layout            e.g. 2x2, 4x5
      S  array_duplicate_check   yes / no
      T  array_duplicate_scope   project | global

    ARRAY BARCODE RULES  (same columns as single, offset to Q-block)
      U  array_format
      V  array_prefix
      W  array_suffix
      X  array_total_length
      Y  array_running_number_length
      Z  array_year_section_start
      AA array_year_section_length
      AB array_workweek_section_start
      AC array_workweek_section_length
      AD array_regex_pattern
      AE array_description
    """
    if not _OPENPYXL_OK:
        raise RuntimeError(
            "openpyxl is not installed.\n"
            "Run:  pip install openpyxl\n"
            "then restart the application.")

    wb = load_workbook(path, data_only=True)

    # ── App sheet ──────────────────────────────
    app_cfg = {
        "station":              "IQC-01",
        "operator_required":    True,
        "scan_timeout_seconds": 30,
    }
    if "App" in wb.sheetnames:
        ws = wb["App"]
        app_cfg["station"]           = _cell(ws, 2, 1) or "IQC-01"
        app_cfg["operator_required"] = _bool(ws, 2, 2)
        t = _int(ws, 2, 3)
        if t:
            app_cfg["scan_timeout_seconds"] = t

    # ── Projects sheet ─────────────────────────
    projects = []
    if "Projects" in wb.sheetnames:
        ws = wb["Projects"]
        for row in range(2, ws.max_row + 1):
            pid = _cell(ws, row, 1)
            if not pid:
                continue                    # skip blank rows

            mode      = _cell(ws, row, 3).lower()
            dup_check = _bool(ws, row, 4)
            dup_scope = _cell(ws, row, 5) or "project"

            # Single / panel-SN barcode rules (columns F–P = 6–16)
            single_rules = _rules(
                ws, row,
                col_fmt=6,  col_prefix=7,  col_suffix=8,
                col_total_len=9,  col_run_len=10,
                col_yr_start=11,  col_yr_len=12,
                col_ww_start=13,  col_ww_len=14,
                col_regex=15,     col_desc=16,
            )

            proj = {
                "id":              pid,
                "name":            _cell(ws, row, 2) or pid,
                "panel_mode":      mode == "panel",
                "duplicate_check": dup_check,
                "duplicate_scope": dup_scope,
                "barcode_rules":   single_rules,
            }

            if mode == "panel":
                arr_count      = _int(ws, row, 17) or 4
                arr_layout     = _cell(ws, row, 18) or "2x2"
                arr_dup_check  = _bool(ws, row, 19)
                arr_dup_scope  = _cell(ws, row, 20) or "project"

                # Array barcode rules (columns U–AE = 21–31)
                array_rules = _rules(
                    ws, row,
                    col_fmt=21,  col_prefix=22,  col_suffix=23,
                    col_total_len=24, col_run_len=25,
                    col_yr_start=26,  col_yr_len=27,
                    col_ww_start=28,  col_ww_len=29,
                    col_regex=30,     col_desc=31,
                )

                proj["panel_config"] = {
                    "array_count":           arr_count,
                    "array_layout":          arr_layout,
                    "array_duplicate_check": arr_dup_check,
                    "array_duplicate_scope": arr_dup_scope,
                    "panel_barcode_rules":   single_rules,   # reuse col F–P
                    "array_barcode_rules":   array_rules,    # col U–AE
                }

            projects.append(proj)

    return {
        "app":      app_cfg,
        "ui":       {"scan_timeout_seconds": app_cfg["scan_timeout_seconds"]},
        "projects": projects,
    }


def generate_template_xlsx(path: Path):
    """
    Write a ready-to-fill config.xlsx template.
    Column headers exactly match the loader's expected positions.
    Two example rows are included (one single-mode, one panel-mode).
    """
    wb = Workbook()

    # ── App sheet ──
    ws_app = wb.active
    ws_app.title = "App"
    ws_app.append(["station", "operator_required (yes/no)", "scan_timeout_seconds"])
    ws_app.append(["IQC-01", "yes", 30])
    for col in ws_app.columns:
        ws_app.column_dimensions[col[0].column_letter].width = 28

    # ── Projects sheet ──
    ws = wb.create_sheet("Projects")
    headers = [
        # A–E  common
        "project_id", "project_name", "mode (single/panel)",
        "duplicate_check (yes/no)", "duplicate_scope (project/global)",
        # F–P  single / panel-SN rules
        "format (none/prefix_suffix/fixed_length/regex/structured)",
        "prefix", "suffix", "total_length", "running_number_length",
        "year_section_start", "year_section_length",
        "workweek_section_start", "workweek_section_length",
        "regex_pattern", "description",
        # Q–T  panel extras
        "array_count", "array_layout (e.g. 2x2)",
        "array_duplicate_check (yes/no)", "array_duplicate_scope (project/global)",
        # U–AE  array barcode rules (same set as F–P)
        "array_format", "array_prefix", "array_suffix",
        "array_total_length", "array_running_number_length",
        "array_year_section_start", "array_year_section_length",
        "array_workweek_section_start", "array_workweek_section_length",
        "array_regex_pattern", "array_description",
    ]
    ws.append(headers)

    # Example 1 — single-mode structured barcode  (MA + YY + WW + 6-digit seq = 12 chars)
    ws.append([
        "PCB-001", "Main Board", "single", "yes", "project",
        "structured", "MA", "", 12, 6,
        3, 2, 5, 2,
        "", "MA + YY(pos3) + WW(pos5) + 6-digit seq",
        "", "", "", "",
        "", "", "", "", "", "", "", "", "", "", "",
    ])

    # Example 2 — panel mode
    #   Panel SN: prefix_suffix, prefix "PNL", 16 chars
    #   Array SN: structured, prefix "ARR", 14 chars, YY at pos4, WW at pos6, 6-digit seq
    ws.append([
        "PCB-002", "Sub Panel", "panel", "no", "project",
        "prefix_suffix", "PNL", "", 16, "",
        "", "", "", "",
        "", "Panel SN: PNL + 13 chars",
        4, "2x2", "yes", "project",
        "structured", "ARR", "", 13, 6,
        4, 2, 6, 2,
        "", "ARR + YY(pos4) + WW(pos6) + 6-digit seq",
    ])

    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 20

    wb.save(path)


# ──────────────────────────────────────────────
# THEME
# ──────────────────────────────────────────────

DARK = {
    "bg":        "#0F1117",
    "surface":   "#1A1D27",
    "surface2":  "#232736",
    "border":    "#2E3347",
    "text":      "#E8EAFF",
    "muted":     "#6B7280",
    "pass":      "#00E676",
    "pass_dim":  "#003D1A",
    "fail":      "#FF3D71",
    "fail_dim":  "#3D000F",
    "warn":      "#FFD740",
    "accent":    "#5B8DEF",
    "accent2":   "#7C6EF5",
    "scan_bg":   "#0C0E14",
    "font_mono": ("Consolas", 13),
    "font_ui":   ("Segoe UI", 10),
    "font_hd":   ("Segoe UI", 11, "bold"),
    "font_lg":   ("Segoe UI", 20, "bold"),
    "font_xl":   ("Segoe UI", 32, "bold"),
}


# ──────────────────────────────────────────────
# DATA LAYER
# ──────────────────────────────────────────────

class Database:
    """
    SQL Server–backed traceability store.
    Uses pyodbc with connection pooling disabled so each instance
    manages its own connection — safe for multi-user desktop deployments.

    Connection string is read from DB_CONNECTION_STRING at module level.
    """

    COLUMNS = [
        "id", "project_id", "project_name", "gr_id", "barcode",
        "panel_sn", "array_index", "timestamp", "operator",
        "station", "result", "failure_reason",
    ]

    def __init__(self):
        self._lock = threading.Lock()
        try:
            self.conn = pyodbc.connect(DB_CONNECTION_STRING, autocommit=False)
        except pyodbc.Error as e:
            raise RuntimeError(
                f"Cannot connect to SQL Server.\n\n{e}\n\n"
                "Check DB_CONNECTION_STRING at the top of the script."
            ) from e

    def _cursor(self):
        """Return a fresh cursor; reconnect once if connection dropped."""
        try:
            return self.conn.cursor()
        except pyodbc.ProgrammingError:
            self.conn = pyodbc.connect(DB_CONNECTION_STRING, autocommit=False)
            return self.conn.cursor()

    def insert(self, record: dict) -> int:
        with self._lock:
            cur = self._cursor()
            cur.execute("""
                INSERT INTO scan_records
                  (project_id, project_name, gr_id, barcode, panel_sn,
                   array_index, timestamp, operator, station, result, failure_reason)
                OUTPUT INSERTED.id
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                record["project_id"], record["project_name"], record["gr_id"],
                record["barcode"],    record["panel_sn"],     record["array_index"],
                record["timestamp"],  record["operator"],     record["station"],
                record["result"],     record["failure_reason"],
            ))
            new_id = cur.fetchone()[0]
            self.conn.commit()
            return new_id

    def is_duplicate(self, barcode: str, project_id: str, scope: str) -> bool:
        with self._lock:
            cur = self._cursor()
            if scope == "global":
                cur.execute(
                    "SELECT TOP 1 1 FROM scan_records WHERE barcode=? AND result='PASS'",
                    (barcode,))
            else:
                cur.execute(
                    "SELECT TOP 1 1 FROM scan_records WHERE barcode=? AND project_id=? AND result='PASS'",
                    (barcode, project_id))
            return cur.fetchone() is not None

    _FILTER_COLS = {
        "id", "project_id", "project_name", "gr_id", "barcode",
        "panel_sn", "array_index", "timestamp", "operator",
        "station", "result", "failure_reason",
    }

    def _build_where(self, date_from: str = "", date_to: str = "", **filters):
        clauses, params = [], []
        for key, val in filters.items():
            if val and key in self._FILTER_COLS:
                clauses.append(f"{key} LIKE ?")
                params.append(f"%{val}%")
        if date_from:
            clauses.append("timestamp >= ?")
            params.append(date_from)
        if date_to:
            clauses.append("timestamp <= ?")
            params.append(date_to + "T23:59:59")
        where = ("WHERE " + " AND ".join(clauses)) if clauses else ""
        return where, params

    def count(self, date_from: str = "", date_to: str = "", **filters) -> int:
        where, params = self._build_where(date_from=date_from, date_to=date_to, **filters)
        with self._lock:
            cur = self._cursor()
            cur.execute(f"SELECT COUNT(*) FROM scan_records {where}", params)
            return cur.fetchone()[0]

    def search(self, date_from: str = "", date_to: str = "",
               limit: int = None, offset: int = 0, **filters) -> list:
        where, params = self._build_where(date_from=date_from, date_to=date_to, **filters)
        if limit is not None:
            # MSSQL uses OFFSET … FETCH instead of LIMIT … OFFSET
            order_page = f"ORDER BY id DESC OFFSET {offset} ROWS FETCH NEXT {limit} ROWS ONLY"
            sql = f"SELECT * FROM scan_records {where} {order_page}"
        else:
            sql = f"SELECT * FROM scan_records {where} ORDER BY id DESC"
        with self._lock:
            cur = self._cursor()
            cur.execute(sql, params)
            return cur.fetchall()

    def export_csv(self, path: str, date_from: str = "", date_to: str = "", **filters):
        where, params = self._build_where(date_from=date_from, date_to=date_to, **filters)
        with self._lock:
            cur = self._cursor()
            cur.execute(f"SELECT * FROM scan_records {where} ORDER BY id DESC", params)
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(self.COLUMNS)
            while True:
                chunk = cur.fetchmany(500)
                if not chunk:
                    break
                w.writerows(chunk)

    def get_stats(self, project_id: str = None, gr_id: str = None) -> dict:
        conds, params = [], []
        if project_id:
            conds.append("project_id=?"); params.append(project_id)
        if gr_id:
            conds.append("gr_id=?");     params.append(gr_id)
        base = "FROM scan_records" + ((" WHERE " + " AND ".join(conds)) if conds else "")
        and_ = "AND" if conds else "WHERE"
        with self._lock:
            cur = self._cursor()
            cur.execute(f"SELECT COUNT(*) {base}", params)
            total  = cur.fetchone()[0]
            cur.execute(f"SELECT COUNT(*) {base} {and_} result='PASS'", params)
            passed = cur.fetchone()[0]
        return {"total": total, "passed": passed, "failed": total - passed}


# ──────────────────────────────────────────────
# VALIDATION ENGINE
# ──────────────────────────────────────────────

class BarcodeValidator:
    """
    Stateless validator — applies configurable rule dicts.

    All formats work identically whether applied to a single barcode,
    a panel SN, or an individual array SN.  The same rule dict keys
    are used everywhere so the xlsx config is symmetric.

    Supported format values:
      none           – any non-empty barcode passes
      prefix_suffix  – checks prefix and/or suffix strings
      fixed_length   – exact character count
      regex          – Python re.fullmatch against 'pattern'
      structured     – prefix + year(YY) + workweek(WW) +
                       running-number section, all positions configurable
    """

    def validate(self, barcode: str, rules: dict):
        """Returns (ok: bool, reason: str)."""
        b = barcode.strip()
        if not b:
            return False, "Empty barcode"

        fmt = rules.get("format", "none")
        mn  = rules.get("min_length")
        mx  = rules.get("max_length")

        if mn and len(b) < mn:
            return False, f"Too short ({len(b)} chars, minimum {mn})"
        if mx and len(b) > mx:
            return False, f"Too long ({len(b)} chars, maximum {mx})"

        if fmt == "regex":
            pattern = rules.get("pattern", ".*")
            if not re.fullmatch(pattern, b):
                desc = rules.get("description", pattern)
                return False, f"Format mismatch — expected: {desc}"

        elif fmt == "prefix_suffix":
            prefix = rules.get("prefix", "")
            suffix = rules.get("suffix", "")
            if prefix and not b.startswith(prefix):
                return False, f"Must start with '{prefix}'"
            if suffix and not b.endswith(suffix):
                return False, f"Must end with '{suffix}'"

        elif fmt == "fixed_length":
            length = rules.get("length")
            if length and len(b) != length:
                return False, f"Must be exactly {length} characters (got {len(b)})"

        elif fmt == "structured":
            ok, reason = self._validate_structured(b, rules)
            if not ok:
                return False, reason

        # fmt == "none" → always passes format check
        return True, ""

    # ── Structured-format helper ─────────────

    def _validate_structured(self, b: str, rules: dict):
        """
        Validates prefix · year(YY) · workweek(WW) · running-number · suffix.
        All section positions are 1-based character indices from the xlsx.
        Only sections with non-None config values are checked.
        """
        prefix = rules.get("prefix", "")
        suffix = rules.get("suffix", "")

        if prefix and not b.startswith(prefix):
            return False, f"Must start with '{prefix}'"
        if suffix and not b.endswith(suffix):
            return False, f"Must end with '{suffix}'"

        yr_s = rules.get("year_section_start")
        yr_l = rules.get("year_section_length")
        if yr_s and yr_l:
            section  = b[yr_s - 1: yr_s - 1 + yr_l]
            expected = datetime.now().strftime("%y")
            if not section.isdigit():
                return False, f"Year section (pos {yr_s}, len {yr_l}) must be numeric, got '{section}'"
            if section != expected:
                return False, f"Year mismatch — barcode '{section}', expected '{expected}'"

        ww_s = rules.get("workweek_section_start")
        ww_l = rules.get("workweek_section_length")
        if ww_s and ww_l:
            section = b[ww_s - 1: ww_s - 1 + ww_l]
            if not section.isdigit() or not (1 <= int(section) <= 53):
                return False, f"Work-week (pos {ww_s}, len {ww_l}) must be 01–53, got '{section}'"

        rn_l  = rules.get("running_number_length")
        total = rules.get("length") or rules.get("min_length")
        if rn_l and total:
            suf_len  = len(suffix) if suffix else 0
            rn_start = total - suf_len - rn_l   # 0-based
            section  = b[rn_start: rn_start + rn_l]
            if not section.isdigit():
                return False, f"Running number ({rn_l} digits before suffix) must be numeric, got '{section}'"

        return True, ""


# ──────────────────────────────────────────────
# PANEL SCAN DIALOG
# ──────────────────────────────────────────────

class PanelScanDialog(tk.Toplevel):
    """
    Modal dialog for panel + array scanning.

    Layout strategy:
      ≤ 16 arrays  → visual grid (compact slots, no entry box per slot)
      > 16 arrays  → sequential single-entry mode (one scan at a time,
                     large progress counter, scrollable completion list)

    Validation features (symmetric with single-barcode mode):
      • Same rule dict format as single barcodes (format, prefix, suffix,
        length, structured sections, regex, etc.)
      • Intra-panel duplicate detection (two arrays cannot share a barcode)
      • Cross-panel duplicate detection (optional, uses project DB scope)
    """

    GRID_THRESHOLD = 16   # above this, use sequential mode
    LAYOUT_COLS = {
        "1x2":1, "1x4":1, "2x2":2, "2x3":3, "2x4":4, "2x5":5,
        "3x3":3, "3x4":4, "4x4":4, "4x5":5, "1x6":1, "2x6":6,
    }

    def __init__(self, parent, panel_sn: str, panel_cfg: dict,
                 validator: "BarcodeValidator", db: "Database",
                 project_id: str, on_complete, on_fail=None):
        super().__init__(parent)
        T = DARK
        self.title(f"Panel Array Scan  ·  {panel_sn}")
        self.configure(bg=T["bg"])
        self.resizable(True, True)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._abort)

        self.panel_sn    = panel_sn
        self.cfg         = panel_cfg
        self.validator   = validator
        self.db          = db
        self.project_id  = project_id
        self.on_complete = on_complete
        self.on_fail = on_fail

        self.array_count  = panel_cfg.get("array_count", 4)
        self.array_rules  = panel_cfg.get("array_barcode_rules", {})
        self.arr_dup_check = panel_cfg.get("array_duplicate_check", False)
        self.arr_dup_scope = panel_cfg.get("array_duplicate_scope", "project")
        layout            = panel_cfg.get("array_layout", "2x2")
        self.layout_cols  = self.LAYOUT_COLS.get(layout, 2)

        self.scanned       = {}   # slot_index (int) → barcode string
        self.scanned_set   = set()  # fast intra-panel dedup lookup

        # Grid-mode widget dicts (only used when ≤ GRID_THRESHOLD)
        self.slot_frames   = {}
        self.status_labels = {}
        self.entries       = {}

        self._build_ui()
        if self.array_count <= self.GRID_THRESHOLD:
            self._focus_next()
        else:
            self.seq_entry.focus_set()

    # ── Build ────────────────────────────────

    def _build_ui(self):
        T = DARK

        # ── Header ──
        hdr = tk.Frame(self, bg=T["surface"], pady=10)
        hdr.pack(fill="x")
        tk.Label(hdr, text="📋  Panel Array Scan", font=("Segoe UI", 13, "bold"),
                 bg=T["surface"], fg=T["text"]).pack(side="left", padx=14)
        tk.Label(hdr, text=f"SN: {self.panel_sn}", font=("Consolas", 10),
                 bg=T["surface"], fg=T["accent"]).pack(side="left", padx=6)
        self.prog_lbl = tk.Label(hdr, text=f"0 / {self.array_count}",
                                  font=("Segoe UI", 11, "bold"),
                                  bg=T["surface"], fg=T["muted"])
        self.prog_lbl.pack(side="right", padx=14)

        style = ttk.Style()
        style.configure("Panel.Horizontal.TProgressbar",
                        troughcolor=T["surface2"], background=T["accent"],
                        bordercolor=T["border"])
        self.pbar_var = tk.DoubleVar(value=0)
        ttk.Progressbar(self, variable=self.pbar_var, maximum=self.array_count,
                        mode="determinate",
                        style="Panel.Horizontal.TProgressbar").pack(fill="x")

        # ── Body: grid or sequential ──
        if self.array_count <= self.GRID_THRESHOLD:
            self._build_grid_body()
        else:
            self._build_sequential_body()

        # ── Footer ──
        bot = tk.Frame(self, bg=T["surface"], pady=8)
        bot.pack(fill="x")
        tk.Button(bot, text="  ✕  Abort  ", command=self._abort,
                  bg=T["fail_dim"], fg=T["fail"], relief="flat",
                  font=DARK["font_hd"], padx=10, pady=6).pack(side="left", padx=12)
        self.submit_btn = tk.Button(bot, text="  ✓  Submit Panel  ",
                                    command=self._submit,
                                    bg=T["surface2"], fg=T["muted"], relief="flat",
                                    font=DARK["font_hd"], padx=10, pady=6, state="disabled")
        self.submit_btn.pack(side="right", padx=12)
        self.bot_hint = tk.Label(bot, text="Scan all arrays to enable submit",
                                  font=("Segoe UI", 9),
                                  bg=T["surface"], fg=T["muted"])
        self.bot_hint.pack(side="right", padx=8)
        self._update_progress()

    # ── Grid body (≤ GRID_THRESHOLD arrays) ──

    def _build_grid_body(self):
        T = DARK
        grid_outer = tk.Frame(self, bg=T["bg"])
        grid_outer.pack(fill="both", expand=True, padx=12, pady=10)

        for i in range(self.array_count):
            row_i = i // self.layout_cols
            col_i = i %  self.layout_cols
            grid_outer.columnconfigure(col_i, weight=1)
            grid_outer.rowconfigure(row_i, weight=1)

            slot = tk.Frame(grid_outer, bg=T["surface2"],
                            highlightthickness=2,
                            highlightbackground=T["border"])
            slot.grid(row=row_i, column=col_i, padx=5, pady=5, sticky="nsew")

            tk.Label(slot, text=f" Array {i+1:02d} ",
                     font=("Segoe UI", 8, "bold"),
                     bg=T["surface2"], fg=T["muted"]).pack(anchor="w", pady=(8,0), padx=8)

            entry = tk.Entry(slot, font=("Consolas", 11),
                             bg=T["scan_bg"], fg=T["text"],
                             insertbackground=T["accent"], relief="flat",
                             highlightthickness=1, highlightbackground=T["border"])
            entry.pack(fill="x", padx=8, pady=4, ipady=5)
            entry.bind("<Return>",  lambda e, idx=i: self._on_scan_grid(idx))
            entry.bind("<FocusIn>", lambda e, idx=i: self._set_active(idx))

            status = tk.Label(slot, text="  ⬜  Awaiting scan",
                              font=("Segoe UI", 9),
                              bg=T["surface2"], fg=T["muted"], anchor="w")
            status.pack(fill="x", padx=8, pady=(0, 8))

            self.entries[i]       = entry
            self.status_labels[i] = status
            self.slot_frames[i]   = slot

    # ── Sequential body (> GRID_THRESHOLD arrays) ──

    def _build_sequential_body(self):
        T = DARK
        body = tk.Frame(self, bg=T["bg"])
        body.pack(fill="both", expand=True, padx=16, pady=10)

        # Big current-slot indicator
        self.seq_slot_lbl = tk.Label(body,
                                      text=f"Scan  Array 01  of  {self.array_count}",
                                      font=("Segoe UI", 18, "bold"),
                                      bg=T["bg"], fg=T["accent"])
        self.seq_slot_lbl.pack(pady=(10, 4))

        self.seq_status_lbl = tk.Label(body, text="⬜  Awaiting scan",
                                        font=("Segoe UI", 11),
                                        bg=T["bg"], fg=T["muted"])
        self.seq_status_lbl.pack(pady=(0, 10))

        self.seq_entry = tk.Entry(body, font=("Consolas", 16),
                                   bg=T["scan_bg"], fg=T["text"],
                                   insertbackground=T["accent"],
                                   relief="flat", highlightthickness=2,
                                   highlightbackground=T["accent"])
        self.seq_entry.pack(fill="x", ipady=10, pady=(0, 6))
        self.seq_entry.bind("<Return>", self._on_scan_sequential)

        tk.Label(body, text="↵ Enter or scanner trigger",
                 font=("Segoe UI", 9), bg=T["bg"], fg=T["muted"]).pack()

        # Scrollable completion list
        list_hdr = tk.Frame(body, bg=T["surface2"])
        list_hdr.pack(fill="x", pady=(14, 0))
        tk.Label(list_hdr, text="Completed arrays", font=("Segoe UI", 9, "bold"),
                 bg=T["surface2"], fg=T["muted"], pady=4).pack(side="left", padx=8)
        self.seq_count_lbl = tk.Label(list_hdr, text="0 / " + str(self.array_count),
                                       font=("Segoe UI", 9, "bold"),
                                       bg=T["surface2"], fg=T["muted"])
        self.seq_count_lbl.pack(side="right", padx=8)

        list_frame = tk.Frame(body, bg=T["surface"])
        list_frame.pack(fill="both", expand=True)

        self.seq_canvas = tk.Canvas(list_frame, bg=T["surface"], highlightthickness=0)
        seq_vsb = ttk.Scrollbar(list_frame, orient="vertical",
                                 command=self.seq_canvas.yview)
        self.seq_canvas.configure(yscrollcommand=seq_vsb.set)
        seq_vsb.pack(side="right", fill="y")
        self.seq_canvas.pack(side="left", fill="both", expand=True)

        self.seq_list_inner = tk.Frame(self.seq_canvas, bg=T["surface"])
        self.seq_canvas.create_window((0, 0), window=self.seq_list_inner,
                                       anchor="nw", tags="inner")
        self.seq_list_inner.bind("<Configure>",
                                  lambda e: self.seq_canvas.configure(
                                      scrollregion=self.seq_canvas.bbox("all")))
        self.seq_canvas.bind("<Configure>",
                              lambda e: self.seq_canvas.itemconfig("inner", width=e.width))

    # ── Interaction — grid mode ───────────────

    def _set_active(self, idx):
        T = DARK
        for i, slot in self.slot_frames.items():
            slot.configure(highlightbackground=T["accent"] if i == idx else T["border"])

    def _focus_next(self):
        for i in range(self.array_count):
            if i not in self.scanned:
                self.entries[i].focus_set()
                return

    def _on_scan_grid(self, idx):
        T = DARK
        raw = self.entries[idx].get().strip()
        if not raw:
            return

        if idx in self.scanned:
            # Re-scan: undo previous accept
            self.scanned_set.discard(self.scanned[idx])
            del self.scanned[idx]
            self.entries[idx].configure(state="normal", bg=T["scan_bg"])

        ok, reason = self._validate_array(raw)
        slot = self.slot_frames[idx]

        if ok:
            self.scanned[idx] = raw
            self.scanned_set.add(raw)
            slot.configure(highlightbackground=T["pass"])
            self.status_labels[idx].configure(
                text=f"  ✅  {raw}", fg=T["pass"], bg=T["pass_dim"])
            self.entries[idx].configure(
                state="disabled", bg=T["pass_dim"], disabledforeground=T["pass"])
        else:
            slot.configure(highlightbackground=T["fail"])
            self.status_labels[idx].configure(
                text=f"  ❌  {reason}", fg=T["fail"], bg=T["fail_dim"])
            self.entries[idx].delete(0, "end")
            self.entries[idx].configure(highlightbackground=T["fail"])
            if self.on_fail:
                self.on_fail(raw, idx, reason)

        self._update_progress()
        if len(self.scanned) < self.array_count:
            self.after(50, self._focus_next)

    # ── Interaction — sequential mode ─────────

    def _next_unscanned_index(self):
        for i in range(self.array_count):
            if i not in self.scanned:
                return i
        return None

    def _on_scan_sequential(self, event=None):
        T = DARK
        raw = self.seq_entry.get().strip()
        if not raw:
            return

        idx = self._next_unscanned_index()
        if idx is None:
            return

        ok, reason = self._validate_array(raw)
        self.seq_entry.delete(0, "end")

        if ok:
            self.scanned[idx] = raw
            self.scanned_set.add(raw)
            # Add row to completion list
            row_bg = T["pass_dim"]
            row = tk.Frame(self.seq_list_inner, bg=row_bg,
                           highlightthickness=1, highlightbackground=T["border"])
            row.pack(fill="x", pady=1, padx=2)
            tk.Label(row, text=f"  A{idx+1:02d}", font=("Segoe UI", 9, "bold"),
                     bg=row_bg, fg=T["muted"], width=5).pack(side="left")
            tk.Label(row, text=raw, font=("Consolas", 9),
                     bg=row_bg, fg=T["pass"]).pack(side="left", padx=8)
            self.seq_canvas.update_idletasks()
            self.seq_canvas.yview_moveto(1.0)

            self.seq_status_lbl.configure(text=f"✅  Accepted: {raw}", fg=T["pass"])
            next_idx = self._next_unscanned_index()
            if next_idx is not None:
                self.seq_slot_lbl.configure(
                    text=f"Scan  Array {next_idx+1:02d}  of  {self.array_count}")
        else:
            self.seq_status_lbl.configure(text=f"❌  {reason}", fg=T["fail"])
            self.seq_slot_lbl.configure(
                text=f"Retry  Array {idx+1:02d}  of  {self.array_count}")
            if self.on_fail:
                self.on_fail(raw, idx, reason)

        self._update_progress()

    # ── Shared validation ─────────────────────

    def _validate_array(self, barcode: str):
        """Runs format rules, intra-panel dedup, and optional DB dedup."""
        ok, reason = self.validator.validate(barcode, self.array_rules)
        if not ok:
            return False, reason

        # Intra-panel duplicate (two arrays on same panel cannot share a SN)
        if barcode in self.scanned_set:
            return False, "Duplicate — this barcode already scanned on this panel"

        # Cross-panel / DB duplicate (optional per project config)
        if self.arr_dup_check:
            if self.db.is_duplicate(barcode, self.project_id, self.arr_dup_scope):
                return False, f"Duplicate — barcode already passed ({self.arr_dup_scope} scope)"

        return True, ""

    # ── Progress ──────────────────────────────

    def _update_progress(self):
        T = DARK
        n = len(self.scanned)
        self.pbar_var.set(n)
        self.prog_lbl.configure(text=f"{n} / {self.array_count}")

        if self.array_count > self.GRID_THRESHOLD:
            self.seq_count_lbl.configure(text=f"{n} / {self.array_count}")

        if n == self.array_count:
            self.submit_btn.configure(state="normal", bg=T["pass_dim"], fg=T["pass"])
            self.bot_hint.configure(text="All arrays scanned ✓", fg=T["pass"])
        else:
            remaining = self.array_count - n
            self.bot_hint.configure(text=f"{remaining} array(s) remaining", fg=T["muted"])

    def _submit(self):
        if len(self.scanned) < self.array_count:
            messagebox.showwarning("Incomplete", "Not all arrays have been scanned.", parent=self)
            return
        self.on_complete(self.panel_sn, dict(self.scanned))
        self.destroy()

    def _abort(self):
        if messagebox.askyesno("Abort", "Abort this panel scan? Data will not be saved.", parent=self):
            self.on_complete(self.panel_sn, None)
            self.destroy()


# ──────────────────────────────────────────────
# RECORDS VIEWER
# ──────────────────────────────────────────────

class RecordsPanel(tk.Toplevel):
    """Searchable, filterable traceability log with CSV export."""

    COLS = ("id","project","gr_id","barcode","panel_sn",
            "timestamp","operator","station","result","reason")
    HEADS = ["#","Project","GR ID","Barcode","Panel SN",
             "Timestamp","Operator","Station","Result","Failure Reason"]
    WIDTHS = [44, 130, 90, 180, 110, 140, 90, 110, 65, 220]
    PAGE_SIZE = 200

    def __init__(self, parent, db: "Database"):
        super().__init__(parent)
        self.title("IQC Records & Traceability")
        self.geometry("1280x720")
        self.configure(bg=DARK["bg"])
        self.db = db
        self._page       = 0   # current 0-based page index
        self._total_rows = 0   # total matching rows from last count()
        self._build()
        self._load()

    def _build(self):
        T = DARK

        # ── Row 1: keyword filters ──
        sf1 = tk.Frame(self, bg=T["surface"], pady=6, padx=10)
        sf1.pack(fill="x")
        tk.Label(sf1, text="🔍", bg=T["surface"], fg=T["muted"],
                 font=("Segoe UI", 12)).pack(side="left")

        self.filter_vars = {}
        for label, key, w in [
            ("Barcode",  "barcode",      14),
            ("GR ID",    "gr_id",        10),
            ("Operator", "operator",     10),
            ("Project",  "project_name", 12),
            ("Result",   "result",        7),
        ]:
            tk.Label(sf1, text=label+":", bg=T["surface"], fg=T["muted"],
                     font=("Segoe UI", 9)).pack(side="left", padx=(8, 2))
            v = tk.StringVar()
            e = tk.Entry(sf1, textvariable=v, width=w, bg=T["scan_bg"],
                         fg=T["text"], insertbackground=T["accent"],
                         relief="flat", highlightthickness=1,
                         highlightbackground=T["border"], font=("Segoe UI", 9))
            e.pack(side="left", ipady=4)
            e.bind("<Return>", lambda evt: self._load())
            self.filter_vars[key] = v

        # ── Row 2: date range ──
        sf2 = tk.Frame(self, bg=T["surface"], pady=4, padx=10)
        sf2.pack(fill="x")

        tk.Label(sf2, text="📅 From:", bg=T["surface"], fg=T["muted"],
                 font=("Segoe UI", 9)).pack(side="left", padx=(4, 2))
        self.date_from_var = tk.StringVar()
        self._e_from = tk.Entry(sf2, textvariable=self.date_from_var, width=11,
                                 bg=T["scan_bg"], fg=T["text"],
                                 insertbackground=T["accent"], relief="flat",
                                 highlightthickness=1, highlightbackground=T["border"],
                                 font=("Segoe UI", 9))
        self._e_from.pack(side="left", ipady=4)
        self._e_from.bind("<Return>", lambda e: self._load())
        tk.Button(sf2, text="▼", width=2,
                  command=lambda: self._pick_date(self.date_from_var),
                  bg=T["surface2"], fg=T["muted"], relief="flat",
                  font=("Segoe UI", 8)).pack(side="left", padx=(0, 8))

        tk.Label(sf2, text="To:", bg=T["surface"], fg=T["muted"],
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 2))
        self.date_to_var = tk.StringVar()
        self._e_to = tk.Entry(sf2, textvariable=self.date_to_var, width=11,
                               bg=T["scan_bg"], fg=T["text"],
                               insertbackground=T["accent"], relief="flat",
                               highlightthickness=1, highlightbackground=T["border"],
                               font=("Segoe UI", 9))
        self._e_to.pack(side="left", ipady=4)
        self._e_to.bind("<Return>", lambda e: self._load())
        tk.Button(sf2, text="▼", width=2,
                  command=lambda: self._pick_date(self.date_to_var),
                  bg=T["surface2"], fg=T["muted"], relief="flat",
                  font=("Segoe UI", 8)).pack(side="left", padx=(0, 8))

        for label, days in [("Today", 0), ("This week", 6), ("This month", 29)]:
            tk.Button(sf2, text=label,
                      command=lambda d=days: self._quick_range(d),
                      bg=T["surface2"], fg=T["muted"], relief="flat",
                      font=("Segoe UI", 8), padx=6, pady=3).pack(side="left", padx=2)

        tk.Button(sf2, text="  Search  ", command=self._load,
                  bg=T["accent"], fg="white", relief="flat",
                  font=DARK["font_hd"], padx=8, pady=4).pack(side="left", padx=(10, 4))
        tk.Button(sf2, text="Clear All", command=self._clear,
                  bg=T["surface2"], fg=T["muted"], relief="flat",
                  font=("Segoe UI", 9), padx=6, pady=4).pack(side="left")
        tk.Button(sf2, text="⬇  Export CSV", command=self._export,
                  bg=T["surface2"], fg=T["text"], relief="flat",
                  font=DARK["font_hd"], padx=10, pady=4).pack(side="right", padx=6)
        self.count_lbl = tk.Label(sf2, text="", font=("Segoe UI", 9),
                                   bg=T["surface"], fg=T["muted"])
        self.count_lbl.pack(side="right", padx=8)

        # ── Tree ──
        tree_frame = tk.Frame(self, bg=T["bg"])
        tree_frame.pack(fill="both", expand=True, padx=8, pady=6)

        style = ttk.Style()
        style.configure("Records.Treeview",
                        background=T["surface"], foreground=T["text"],
                        fieldbackground=T["surface"], rowheight=24,
                        font=("Segoe UI", 9))
        style.configure("Records.Treeview.Heading",
                        background=T["surface2"], foreground=T["muted"],
                        font=("Segoe UI", 9, "bold"), relief="flat")
        style.map("Records.Treeview",
                  background=[("selected", T["accent"])],
                  foreground=[("selected", "white")])

        self.tree = ttk.Treeview(tree_frame, columns=self.COLS,
                                  show="headings", style="Records.Treeview",
                                  selectmode="browse")
        for col, head, width in zip(self.COLS, self.HEADS, self.WIDTHS):
            self.tree.heading(col, text=head)
            self.tree.column(col, width=width, anchor="w", minwidth=30)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.tree.tag_configure("PASS", foreground=DARK["pass"])
        self.tree.tag_configure("FAIL", foreground=DARK["fail"])

        # ── Pagination bar ──
        pg = tk.Frame(self, bg=T["surface2"], pady=5, padx=10)
        pg.pack(fill="x", side="bottom")

        self.btn_first = tk.Button(pg, text="⏮ First", command=self._go_first,
                                    bg=T["surface"], fg=T["muted"], relief="flat",
                                    font=("Segoe UI", 9), padx=6, pady=3)
        self.btn_first.pack(side="left", padx=2)
        self.btn_prev = tk.Button(pg, text="◀ Prev", command=self._go_prev,
                                   bg=T["surface"], fg=T["muted"], relief="flat",
                                   font=("Segoe UI", 9), padx=6, pady=3)
        self.btn_prev.pack(side="left", padx=2)
        self.page_lbl = tk.Label(pg, text="Page 1 of 1",
                                  font=("Segoe UI", 9, "bold"),
                                  bg=T["surface2"], fg=T["text"])
        self.page_lbl.pack(side="left", padx=10)
        self.btn_next = tk.Button(pg, text="Next ▶", command=self._go_next,
                                   bg=T["surface"], fg=T["muted"], relief="flat",
                                   font=("Segoe UI", 9), padx=6, pady=3)
        self.btn_next.pack(side="left", padx=2)
        self.btn_last = tk.Button(pg, text="Last ⏭", command=self._go_last,
                                   bg=T["surface"], fg=T["muted"], relief="flat",
                                   font=("Segoe UI", 9), padx=6, pady=3)
        self.btn_last.pack(side="left", padx=2)

        # Jump-to-page
        tk.Label(pg, text="Go to page:", bg=T["surface2"], fg=T["muted"],
                 font=("Segoe UI", 9)).pack(side="left", padx=(16, 2))
        self.jump_var = tk.StringVar()
        jump_entry = tk.Entry(pg, textvariable=self.jump_var, width=5,
                               bg=T["scan_bg"], fg=T["text"],
                               insertbackground=T["accent"], relief="flat",
                               highlightthickness=1, highlightbackground=T["border"],
                               font=("Segoe UI", 9))
        jump_entry.pack(side="left", ipady=3)
        jump_entry.bind("<Return>", lambda e: self._go_jump())
        tk.Button(pg, text="Go", command=self._go_jump,
                  bg=T["accent"], fg="white", relief="flat",
                  font=("Segoe UI", 9), padx=6, pady=3).pack(side="left", padx=4)

        self.rows_lbl = tk.Label(pg, text="",
                                  font=("Segoe UI", 9),
                                  bg=T["surface2"], fg=T["muted"])
        self.rows_lbl.pack(side="right", padx=8)

    # ── Pagination helpers ────────────────────

    def _total_pages(self):
        return max(1, -(-self._total_rows // self.PAGE_SIZE))  # ceiling division

    def _go_first(self):
        self._page = 0;                  self._load_page()

    def _go_prev(self):
        self._page = max(0, self._page - 1); self._load_page()

    def _go_next(self):
        self._page = min(self._total_pages() - 1, self._page + 1); self._load_page()

    def _go_last(self):
        self._page = self._total_pages() - 1; self._load_page()

    def _go_jump(self):
        try:
            p = int(self.jump_var.get()) - 1
            self._page = max(0, min(self._total_pages() - 1, p))
            self._load_page()
        except ValueError:
            pass
        self.jump_var.set("")

    def _update_pagination_controls(self):
        T = DARK
        total_pages = self._total_pages()
        self.page_lbl.configure(text=f"Page {self._page + 1} of {total_pages}")
        start = self._page * self.PAGE_SIZE + 1
        end   = min(start + self.PAGE_SIZE - 1, self._total_rows)
        self.rows_lbl.configure(
            text=f"Rows {start}–{end} of {self._total_rows:,}")
        # Enable/disable nav buttons
        can_back = self._page > 0
        can_fwd  = self._page < total_pages - 1
        for btn in (self.btn_first, self.btn_prev):
            btn.configure(fg=T["text"] if can_back else T["muted"],
                          state="normal" if can_back else "disabled")
        for btn in (self.btn_next, self.btn_last):
            btn.configure(fg=T["text"] if can_fwd  else T["muted"],
                          state="normal" if can_fwd  else "disabled")

    # ── Load helpers ──────────────────────────

    def _current_filters(self):
        fvals     = {k: v.get().strip() for k, v in self.filter_vars.items()}
        date_from = self.date_from_var.get().strip()
        date_to   = self.date_to_var.get().strip()
        return {k: v for k, v in fvals.items() if v}, date_from, date_to

    def _load(self):
        """New search — reset to page 0 and refresh count."""
        filters, date_from, date_to = self._current_filters()
        for label, val in [("From date", date_from), ("To date", date_to)]:
            if val and not re.match(r"^\d{4}-\d{2}-\d{2}$", val):
                messagebox.showwarning("Date format",
                                       f"{label} must be YYYY-MM-DD  (e.g. 2025-06-01).",
                                       parent=self)
                return
        self._page = 0
        self._total_rows = self.db.count(date_from=date_from, date_to=date_to, **filters)
        self.count_lbl.configure(text=f"{self._total_rows:,} record(s) total")
        self._load_page()

    def _load_page(self):
        """Load just the current page into the tree."""
        filters, date_from, date_to = self._current_filters()
        self.tree.delete(*self.tree.get_children())
        rows = self.db.search(
            date_from=date_from, date_to=date_to,
            limit=self.PAGE_SIZE, offset=self._page * self.PAGE_SIZE,
            **filters)
        for r in rows:
            tag = r[10] if r[10] in ("PASS", "FAIL") else ""
            self.tree.insert("", "end", tags=(tag,),
                values=(r[0], r[2], r[3], r[4], r[5] or "—",
                        r[7], r[8], r[9], r[10], r[11] or ""))
        self._update_pagination_controls()

    def _clear(self):
        for v in self.filter_vars.values():
            v.set("")
        self.date_from_var.set("")
        self.date_to_var.set("")
        self._load()

    def _export(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("All files", "*.*")],
            title="Export Records as CSV",
            parent=self)
        if path:
            filters, date_from, date_to = self._current_filters()
            self.db.export_csv(path, date_from=date_from, date_to=date_to, **filters)
            messagebox.showinfo("Exported", f"Records saved to:\n{path}", parent=self)


# ──────────────────────────────────────────────
# CONFIG EDITOR  (xlsx-based)
# ──────────────────────────────────────────────

class ConfigEditor(tk.Toplevel):
    """
    Launcher for the xlsx config file.
    Opens config.xlsx in Excel / LibreOffice.
    If the file is missing, offers to generate a template.
    """

    def __init__(self, parent, config_path: Path):
        super().__init__(parent)
        T = DARK
        self.title("Configuration — config.xlsx")
        self.geometry("540x300")
        self.resizable(False, False)
        self.configure(bg=T["bg"])
        self.config_path = config_path

        tk.Label(self, text="⚙  Configuration File",
                 font=("Segoe UI", 13, "bold"),
                 bg=T["surface"], fg=T["text"]).pack(fill="x", pady=1)
        tk.Label(self, text=str(config_path),
                 font=("Consolas", 9), bg=T["surface"],
                 fg=T["accent"]).pack(fill="x", padx=14, pady=(0, 2))

        body = tk.Frame(self, bg=T["bg"], pady=16)
        body.pack(fill="both", expand=True, padx=24)

        exists = config_path.exists()
        tk.Label(body,
                 text="✔  config.xlsx found" if exists else "⚠  config.xlsx not found",
                 font=("Segoe UI", 10, "bold"),
                 bg=T["bg"],
                 fg=T["pass"] if exists else T["warn"]).pack(anchor="w", pady=(0, 10))

        tk.Label(body,
                 text=(
                     "Edit config.xlsx in Excel or LibreOffice Calc.\n"
                     "  Sheet 'App'      →  station name, operator required, timeout.\n"
                     "  Sheet 'Projects' →  one row per PCB, all validation rules.\n\n"
                     "Save the file, then restart the application to apply changes."
                 ),
                 font=("Segoe UI", 9), bg=T["bg"], fg=T["muted"],
                 justify="left").pack(anchor="w")

        bf = tk.Frame(self, bg=T["surface"], pady=10)
        bf.pack(fill="x")

        if exists:
            tk.Button(bf, text="  📂  Open in Excel / Calc  ",
                      command=self._open_xlsx,
                      bg=T["accent"], fg="white", relief="flat",
                      font=DARK["font_hd"], padx=10, pady=6).pack(side="left", padx=12)
        else:
            tk.Button(bf, text="  📄  Generate Template xlsx  ",
                      command=self._generate,
                      bg=T["warn"], fg=T["bg"], relief="flat",
                      font=DARK["font_hd"], padx=10, pady=6).pack(side="left", padx=12)

        tk.Button(bf, text="Close", command=self.destroy,
                  bg=T["surface2"], fg=T["muted"], relief="flat",
                  font=("Segoe UI", 9), padx=10, pady=6).pack(side="right", padx=12)

    def _open_xlsx(self):
        import subprocess, platform
        p = str(self.config_path)
        try:
            if platform.system() == "Windows":
                os.startfile(p)
            elif platform.system() == "Darwin":
                subprocess.call(["open", p])
            else:
                subprocess.call(["xdg-open", p])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file:\n{e}", parent=self)

    def _generate(self):
        if not _OPENPYXL_OK:
            messagebox.showerror("Missing library",
                                 "openpyxl is required.\n\nRun:  pip install openpyxl",
                                 parent=self)
            return
        generate_template_xlsx(self.config_path)
        messagebox.showinfo(
            "Template created",
            f"config.xlsx written to:\n{self.config_path}\n\n"
            "Fill in the Projects sheet, then restart the application.",
            parent=self)
        self.destroy()


# ──────────────────────────────────────────────
# MAIN APPLICATION
# ──────────────────────────────────────────────

class IQCScannerApp(tk.Tk):

    CONFIG_PATH = Path(__file__).parent / "config.xlsx"
    MAX_HISTORY = 60  # rows in the live feed

    def __init__(self):
        super().__init__()
        self.cfg       = self._load_config()
        self.db        = Database()
        self.validator = BarcodeValidator()

        # Session state
        self.current_project = None
        self.gr_id_var    = tk.StringVar()
        self.operator_var = tk.StringVar()
        self.session_pass = 0
        self.session_fail = 0
        self._scan_timer  = None

        self._apply_styles()
        self._build_window()
        self._build_menu()
        self._build_ui()
        self._select_project(self.cfg["projects"][0]["name"] if self.cfg["projects"] else "")

    # ── Bootstrap ────────────────────────────

    _EMPTY_CFG = {
        "app":      {"station": "IQC-01", "operator_required": True,
                     "scan_timeout_seconds": 30},
        "ui":       {"scan_timeout_seconds": 30},
        "projects": [],
    }

    def _load_config(self):
        if not self.CONFIG_PATH.exists():
            return dict(self._EMPTY_CFG)
        try:
            cfg = load_config_xlsx(self.CONFIG_PATH)
            return cfg
        except Exception as e:
            messagebox.showerror(
                "Config Load Error",
                f"Failed to read config.xlsx:\n\n{e}\n\n"
                "The app will start with no projects.\n"
                "Use Config → Edit Configuration to fix or regenerate the file.")
            return dict(self._EMPTY_CFG)

    def _apply_styles(self):
        T = DARK
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure(".",
                    background=T["bg"], foreground=T["text"],
                    fieldbackground=T["scan_bg"], bordercolor=T["border"],
                    troughcolor=T["surface2"], relief="flat")
        s.configure("TCombobox",
                    selectbackground=T["accent"],
                    fieldbackground=T["scan_bg"],
                    foreground=T["text"], arrowcolor=T["muted"])
        s.map("TCombobox", fieldbackground=[("readonly", T["scan_bg"])])
        s.configure("TScrollbar",
                    background=T["surface2"],
                    troughcolor=T["bg"], arrowcolor=T["muted"])

    def _build_window(self):
        app = self.cfg["app"]
        self.title(f"IQC Barcode Scanner  ·  {app['station']}")
        self.geometry("1140x800")
        self.minsize(900, 640)
        self.configure(bg=DARK["bg"])
        self.bind("<F5>",     lambda e: self._open_records())
        self.bind("<Escape>", lambda e: self._clear_scan())
        self.bind("<F2>",     lambda e: self.scan_entry.focus_set())

    # ── Menu ──────────────────────────────────

    def _build_menu(self):
        T = DARK
        mbar = tk.Menu(self, bg=T["surface"], fg=T["text"],
                       activebackground=T["accent"], activeforeground="white",
                       relief="flat", bd=0)
        self.configure(menu=mbar)

        def menu(label):
            m = tk.Menu(mbar, tearoff=0, bg=T["surface"], fg=T["text"],
                        activebackground=T["accent"], activeforeground="white")
            mbar.add_cascade(label=label, menu=m)
            return m

        fm = menu("File")
        fm.add_command(label="Export Records…",  command=self._export_records)
        fm.add_separator()
        fm.add_command(label="Exit",              command=self.quit)

        vm = menu("View")
        vm.add_command(label="Records & Search  F5", command=self._open_records)
        vm.add_command(label="Clear Session Counts", command=self._clear_session)

        cm = menu("Config")
        cm.add_command(label="Edit Configuration…", command=self._open_config_editor)

    # ── Main UI ──────────────────────────────

    def _build_ui(self):
        T = DARK

        # ──── Top bar ────
        top = tk.Frame(self, bg=T["surface"], pady=8, padx=14)
        top.pack(fill="x")

        tk.Label(top, text="⬡  IQC SCANNER",
                 font=("Segoe UI", 11, "bold"),
                 bg=T["surface"], fg=T["accent"]).pack(side="left")

        def labeled_entry(parent, label, var, width=14):
            tk.Label(parent, text=label, bg=T["surface"], fg=T["muted"],
                     font=T["font_ui"]).pack(side="left", padx=(16,3))
            e = tk.Entry(parent, textvariable=var, width=width,
                         bg=T["scan_bg"], fg=T["text"],
                         insertbackground=T["accent"], relief="flat",
                         highlightthickness=1, highlightbackground=T["border"],
                         font=T["font_ui"])
            e.pack(side="left", ipady=4)
            return e

        labeled_entry(top, "Operator:", self.operator_var, 14)
        labeled_entry(top, "GR ID:",    self.gr_id_var,    14)

        # Project dropdown
        tk.Label(top, text="Project:", bg=T["surface"], fg=T["muted"],
                 font=T["font_ui"]).pack(side="left", padx=(16,3))
        proj_names = [p["name"] for p in self.cfg["projects"]]
        self.proj_var = tk.StringVar(value=proj_names[0] if proj_names else "")
        proj_box = ttk.Combobox(top, textvariable=self.proj_var,
                                 values=proj_names, state="readonly",
                                 width=24, font=T["font_ui"])
        proj_box.pack(side="left", ipady=4)
        proj_box.bind("<<ComboboxSelected>>",
                      lambda e: self._select_project(self.proj_var.get()))

        # Session counters
        self.fail_lbl = tk.Label(top, text="✗  0",
                                  font=("Segoe UI", 11, "bold"),
                                  bg=T["surface"], fg=T["fail"])
        self.fail_lbl.pack(side="right", padx=4)
        self.pass_lbl = tk.Label(top, text="✓  0",
                                  font=("Segoe UI", 11, "bold"),
                                  bg=T["surface"], fg=T["pass"])
        self.pass_lbl.pack(side="right", padx=4)
        tk.Label(top, text="Session:", bg=T["surface"], fg=T["muted"],
                 font=("Segoe UI", 9)).pack(side="right", padx=(12,2))

        # ──── Body ────
        body = tk.Frame(self, bg=T["bg"])
        body.pack(fill="both", expand=True, padx=12, pady=8)

        # LEFT column
        left = tk.Frame(body, bg=T["bg"])
        left.pack(side="left", fill="both", expand=True)

        # Result display block
        self.res_frame = tk.Frame(left, bg=T["surface"], height=300,
                                   highlightthickness=2,
                                   highlightbackground=T["border"])
        self.res_frame.pack(fill="x", pady=(0,8))
        self.res_frame.pack_propagate(False)

        self.res_icon   = tk.Label(self.res_frame, text="⬡", font=("Segoe UI", 54),
                                    bg=T["surface"], fg=T["border"])
        self.res_icon.pack(pady=(24,2))
        self.res_text   = tk.Label(self.res_frame, text="READY TO SCAN",
                                    font=("Segoe UI", 24, "bold"),
                                    bg=T["surface"], fg=T["muted"])
        self.res_text.pack()
        self.res_detail = tk.Label(self.res_frame,
                                    text="Configure session above, then scan a barcode",
                                    font=("Segoe UI", 10),
                                    bg=T["surface"], fg=T["muted"],
                                    wraplength=480)
        self.res_detail.pack(pady=(4,0))

        # Scan entry
        scan_wrap = tk.Frame(left, bg=T["surface"],
                              highlightthickness=1, highlightbackground=T["border"])
        scan_wrap.pack(fill="x", pady=(0,8))
        tk.Label(scan_wrap, text="SCAN  /  TYPE BARCODE",
                 font=("Segoe UI", 8, "bold"),
                 bg=T["surface"], fg=T["muted"]).pack(anchor="w", padx=14, pady=(10,0))

        self.scan_var = tk.StringVar()
        self.scan_entry = tk.Entry(scan_wrap, textvariable=self.scan_var,
                                   font=("Consolas", 17),
                                   bg=T["scan_bg"], fg=T["text"],
                                   insertbackground=T["accent"],
                                   relief="flat", highlightthickness=2,
                                   highlightbackground=T["accent"])
        self.scan_entry.pack(fill="x", padx=14, pady=8, ipady=10)
        self.scan_entry.bind("<Return>", self._on_scan)
        self.scan_entry.bind("<Key>",    self._reset_scan_timer)

        hint_row = tk.Frame(scan_wrap, bg=T["surface"])
        hint_row.pack(fill="x", padx=14, pady=(0,10))
        tk.Label(hint_row, text="↵ Enter  or  scanner trigger to submit",
                 font=("Segoe UI", 9), bg=T["surface"], fg=T["muted"]).pack(side="left")
        self.timer_lbl = tk.Label(hint_row, text="",
                                   font=("Segoe UI", 9),
                                   bg=T["surface"], fg=T["warn"])
        self.timer_lbl.pack(side="right")

        # Project info card
        self.info_frame = tk.Frame(left, bg=T["surface2"],
                                    highlightthickness=1, highlightbackground=T["border"])
        self.info_frame.pack(fill="x")
        self.info_title = tk.Label(self.info_frame, text="—",
                                    font=DARK["font_hd"],
                                    bg=T["surface2"], fg=T["text"])
        self.info_title.pack(anchor="w", padx=14, pady=(10,2))
        self.info_rules = tk.Label(self.info_frame, text="",
                                    font=("Consolas", 9),
                                    bg=T["surface2"], fg=T["muted"],
                                    justify="left", wraplength=550)
        self.info_rules.pack(anchor="w", padx=14, pady=(0,10))

        # RIGHT: live history
        right = tk.Frame(body, bg=T["surface"], width=310,
                         highlightthickness=1, highlightbackground=T["border"])
        right.pack(side="right", fill="both", padx=(8,0))
        right.pack_propagate(False)

        rh = tk.Frame(right, bg=T["surface2"], pady=8, padx=10)
        rh.pack(fill="x")
        tk.Label(rh, text="Live Feed", font=DARK["font_hd"],
                 bg=T["surface2"], fg=T["text"]).pack(side="left")
        tk.Button(rh, text="Records ↗", command=self._open_records,
                  bg=T["surface2"], fg=T["accent"],
                  relief="flat", font=("Segoe UI", 9)).pack(side="right")

        # Scrollable canvas for history
        self.hist_canvas = tk.Canvas(right, bg=T["surface"],
                                      highlightthickness=0)
        hist_vsb = ttk.Scrollbar(right, orient="vertical",
                                  command=self.hist_canvas.yview)
        self.hist_canvas.configure(yscrollcommand=hist_vsb.set)
        hist_vsb.pack(side="right", fill="y")
        self.hist_canvas.pack(side="left", fill="both", expand=True)

        self.hist_inner = tk.Frame(self.hist_canvas, bg=T["surface"])
        self.hist_canvas.create_window((0,0), window=self.hist_inner,
                                        anchor="nw", tags="inner")
        self.hist_inner.bind("<Configure>", lambda e: self.hist_canvas.configure(
            scrollregion=self.hist_canvas.bbox("all")))
        self.hist_canvas.bind("<Configure>",
                               lambda e: self.hist_canvas.itemconfig(
                                   "inner", width=e.width))

        # Status bar
        sbar = tk.Frame(self, bg=T["surface2"], pady=5, padx=14)
        sbar.pack(fill="x", side="bottom")
        self.stats_lbl = tk.Label(sbar, text="No scans yet",
                                   font=("Segoe UI", 9),
                                   bg=T["surface2"], fg=T["muted"])
        self.stats_lbl.pack(side="left")
        tk.Label(sbar, text=f"F5 Records  |  F2 Focus Scan  |  Esc Clear  |  "
                             f"Station: {self.cfg['app']['station']}",
                 font=("Segoe UI", 8),
                 bg=T["surface2"], fg=T["muted"]).pack(side="right")

        self.scan_entry.focus_set()

    # ── Project Selection ────────────────────

    def _select_project(self, name: str):
        self.current_project = next(
            (p for p in self.cfg["projects"] if p["name"] == name), None)
        self._refresh_info_card()
        self._update_stats()

    def _refresh_info_card(self):
        T = DARK
        p = self.current_project
        if not p:
            self.info_title.configure(text="No project selected")
            self.info_rules.configure(text="")
            return
        mode = "🔲 Panel + Array Mode" if p.get("panel_mode") else "◈ Single Barcode Mode"
        if p.get("panel_mode"):
            pc = p.get("panel_config", {})
            panel_desc = pc.get("panel_barcode_rules", {}).get("description", "—")
            array_desc = pc.get("array_barcode_rules", {}).get("description", "—")
            count      = pc.get("array_count", "?")
            layout     = pc.get("array_layout", "?")
            rules_txt  = (f"Panel SN format : {panel_desc}\n"
                          f"Array SN format : {array_desc}\n"
                          f"Array count     : {count}  ({layout} grid)\n"
                          f"Duplicate check : {'ON' if p.get('duplicate_check') else 'OFF'}"
                          f"  ({p.get('duplicate_scope','project').upper()} scope)")
        else:
            br = p.get("barcode_rules", {})
            rules_txt = (f"Format          : {br.get('description','—')}\n"
                         f"Length          : {br.get('min_length','—')} – {br.get('max_length','—')} chars\n"
                         f"Duplicate check : {'ON' if p.get('duplicate_check') else 'OFF'}"
                         f"  ({p.get('duplicate_scope','project').upper()} scope)")

        self.info_title.configure(text=f"  {mode}  ·  {p['name']}")
        self.info_rules.configure(text=rules_txt)

    # ── Scan Dispatch ────────────────────────

    def _on_scan(self, event=None):
        raw = self.scan_var.get().strip()
        if not raw:
            return

        p  = self.current_project
        op = self.operator_var.get().strip()
        gr = self.gr_id_var.get().strip()

        if not p:
            self._flash("⚠", "No project selected", "Choose a project from the dropdown above.", None)
            return
        if self.cfg["app"].get("operator_required") and not op:
            self._flash("⚠", "Operator required", "Enter your name in the Operator field.", None)
            return
        if not gr:
            self._flash("⚠", "GR ID required", "Enter the GR ID before scanning.", None)
            return

        if p.get("panel_mode"):
            self._scan_panel(raw, p, op, gr)
        else:
            self._scan_single(raw, p, op, gr)
        self._reset_scan_timer()

    def _scan_single(self, barcode, project, operator, gr_id):
        rules  = project.get("barcode_rules", {})
        ok, reason = self.validator.validate(barcode, rules)

        if ok and project.get("duplicate_check"):
            scope = project.get("duplicate_scope", "project")
            if self.db.is_duplicate(barcode, project["id"], scope):
                ok, reason = False, "Duplicate — barcode already PASSED in this scope"

        result = "PASS" if ok else "FAIL"
        self._save_record(project, gr_id, barcode, operator, result, reason)
        self._flash(
            "✔" if ok else "✘",
            result,
            barcode + ("\n" + reason if not ok else ""),
            ok)
        self._add_to_feed(barcode, result, reason)
        self.scan_var.set("")
        self.scan_entry.focus_set()

    def _scan_panel(self, barcode, project, operator, gr_id):
        """Validate panel SN, then open array dialog."""
        pc          = project.get("panel_config", {})
        panel_rules = pc.get("panel_barcode_rules", {})
        ok, reason  = self.validator.validate(barcode, panel_rules)

        if not ok:
            self._flash("✘", "FAIL", f"Invalid panel SN\n{reason}", False)
            self._add_to_feed(barcode, "FAIL", f"Panel SN rejected: {reason}")
            self.scan_var.set("")
            return

        self.scan_var.set("")
        self._flash("◈", "PANEL ACCEPTED",
                    f"{barcode}\nScan individual arrays in the dialog →", None)

        def on_complete(panel_sn, arrays):
            if arrays is None:
                self._flash("◈", "ABORTED", f"Panel {panel_sn} scan cancelled", None)
                return

            ts     = datetime.now().isoformat(timespec="seconds")
            all_ok = True
            for idx, arr_bc in arrays.items():
                # Validation already done inside the dialog; records are clean.
                rec = {
                    "project_id":     project["id"],
                    "project_name":   project["name"],
                    "gr_id":          gr_id,
                    "barcode":        arr_bc,
                    "panel_sn":       panel_sn,
                    "array_index":    idx + 1,
                    "timestamp":      ts,
                    "operator":       operator,
                    "station":        self.cfg["app"]["station"],
                    "result":         "PASS",
                    "failure_reason": None,
                }
                self.db.insert(rec)
                self._add_to_feed(arr_bc, "PASS", "",
                                   prefix=f"[{panel_sn} A{idx+1:02d}]")

            self._flash("✔", "PANEL PASS",
                        f"{panel_sn}\nAll {len(arrays)} arrays PASSED", True)
            self.session_pass += 1
            self._refresh_session_counters()
            self._update_stats()

        def on_fail(arr_bc, idx, reason):
            rec = {
                "project_id":     project["id"],
                "project_name":   project["name"],
                "gr_id":          gr_id,
                "barcode":        arr_bc,
                "panel_sn":       barcode,
                "array_index":    idx + 1,
                "timestamp":      datetime.now().isoformat(timespec="seconds"),
                "operator":       operator,
                "station":        self.cfg["app"]["station"],
                "result":         "FAIL",
                "failure_reason": reason,
            }
            self.db.insert(rec)
            self.session_fail += 1
            self._refresh_session_counters()
            self._update_stats()
            self._add_to_feed(arr_bc, "FAIL", reason,
                               prefix=f"[{barcode} A{idx+1:02d}]")

        PanelScanDialog(
            self, barcode, pc,
            self.validator, self.db, project["id"],
            on_complete, on_fail,
        )

    # ── Record Saving ────────────────────────

    def _save_record(self, project, gr_id, barcode, operator, result, reason):
        self.db.insert({
            "project_id":     project["id"],
            "project_name":   project["name"],
            "gr_id":          gr_id,
            "barcode":        barcode,
            "panel_sn":       None,
            "array_index":    None,
            "timestamp":      datetime.now().isoformat(timespec="seconds"),
            "operator":       operator,
            "station":        self.cfg["app"]["station"],
            "result":         result,
            "failure_reason": reason if result == "FAIL" else None,
        })
        if result == "PASS":
            self.session_pass += 1
        else:
            self.session_fail += 1
        self._refresh_session_counters()
        self._update_stats()

    # ── UI Feedback ──────────────────────────

    def _flash(self, icon, title, detail, passed):
        T = DARK
        if passed is True:
            bg, fg = T["pass_dim"], T["pass"]
            hl = T["pass"]
        elif passed is False:
            bg, fg = T["fail_dim"], T["fail"]
            hl = T["fail"]
        else:
            bg, fg = T["surface"], T["accent"]
            hl = T["accent"]

        self.res_frame.configure(bg=bg, highlightbackground=hl)
        self.res_icon.configure(text=icon, bg=bg, fg=fg)
        self.res_text.configure(text=title, bg=bg, fg=fg)
        self.res_detail.configure(
            text=detail,
            bg=bg,
            fg=T["text"] if passed is not None else T["muted"])

    def _add_to_feed(self, barcode, result, reason="", prefix=""):
        T = DARK
        is_pass = result == "PASS"
        color   = T["pass"] if is_pass else T["fail"]
        icon    = "✔" if is_pass else "✘"
        ts      = datetime.now().strftime("%H:%M:%S")
        bg      = T["pass_dim"] if is_pass else T["fail_dim"]

        row = tk.Frame(self.hist_inner, bg=bg,
                       highlightthickness=1,
                       highlightbackground=T["border"])
        row.pack(fill="x", pady=2, padx=3)

        tk.Label(row, text=icon, font=("Segoe UI", 11, "bold"),
                 bg=bg, fg=color, width=2).pack(side="left", padx=(6,2))
        inner = tk.Frame(row, bg=bg)
        inner.pack(side="left", fill="x", expand=True, pady=4)
        disp = (f"{prefix}  {barcode}" if prefix else barcode)
        tk.Label(inner, text=disp, font=("Consolas", 9),
                 bg=bg, fg=T["text"], anchor="w",
                 wraplength=220).pack(fill="x")
        if reason:
            tk.Label(inner, text=reason, font=("Segoe UI", 8),
                     bg=bg, fg=T["fail"], anchor="w",
                     wraplength=220).pack(fill="x")
        tk.Label(row, text=ts, font=("Segoe UI", 8),
                 bg=bg, fg=T["muted"]).pack(side="right", padx=6)

        # Trim old rows
        kids = self.hist_inner.winfo_children()
        if len(kids) > self.MAX_HISTORY:
            kids[0].destroy()

        # Scroll to bottom
        self.hist_canvas.update_idletasks()
        self.hist_canvas.yview_moveto(1.0)

    def _refresh_session_counters(self):
        self.pass_lbl.configure(text=f"✓  {self.session_pass}")
        self.fail_lbl.configure(text=f"✗  {self.session_fail}")

    def _update_stats(self):
        p = self.current_project
        if not p:
            self.stats_lbl.configure(text="No project selected")
            return
        s    = self.db.get_stats(project_id=p["id"])
        rate = (s["passed"] / s["total"] * 100) if s["total"] else 0.0
        self.stats_lbl.configure(
            text=(f"{p['name']}  |  All-time — "
                  f"Total: {s['total']}   Pass: {s['passed']}   "
                  f"Fail: {s['failed']}   Yield: {rate:.1f}%"))

    # ── Scan Timer ───────────────────────────

    def _reset_scan_timer(self, event=None):
        if self._scan_timer:
            self.after_cancel(self._scan_timer)
        timeout = self.cfg["ui"].get("scan_timeout_seconds", 30)
        self._scan_timer = self.after(timeout * 1000, self._scan_timeout)
        self.timer_lbl.configure(text=f"⏱ {timeout}s")
        self._tick_timer(timeout)

    def _tick_timer(self, remaining):
        if remaining <= 0:
            return
        self.timer_lbl.configure(text=f"⏱ {remaining}s")
        self._scan_timer = self.after(1000, self._tick_timer, remaining - 1)

    def _scan_timeout(self):
        self.timer_lbl.configure(text="")
        self.scan_var.set("")

    def _clear_scan(self):
        self.scan_var.set("")
        if self._scan_timer:
            self.after_cancel(self._scan_timer)
        self.timer_lbl.configure(text="")
        self.scan_entry.focus_set()

    def _clear_session(self):
        self.session_pass = 0
        self.session_fail = 0
        self._refresh_session_counters()

    # ── Action Launchers ─────────────────────

    def _open_records(self):
        RecordsPanel(self, self.db)

    def _open_config_editor(self):
        ConfigEditor(self, self.CONFIG_PATH)

    def _export_records(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("All files", "*.*")],
            title="Export All Records")
        if path:
            self.db.export_csv(path)
            messagebox.showinfo("Exported", f"Records saved to:\n{path}")


# ──────────────────────────────────────────────
# ENTRY POINT
# ──────────────────────────────────────────────

if __name__ == "__main__":
    app = IQCScannerApp()
    _ico = _resource("barcode_verifier.ico")
    if pathlib.Path(_ico).exists():
        app.iconbitmap(_ico)
    app.mainloop()
