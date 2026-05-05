"""
GFEM BDF Combination Tool
Reads a Combination Excel, resolves unit case files via a List Subcases mapping,
and writes a complete Nastran solution deck (SOL 101 + SUBCASEs + LOAD entries).
"""
from __future__ import annotations

import os
import queue
import re
import threading
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

try:
    import pandas as pd
except ImportError:
    messagebox.showerror(
        "Missing Dependency",
        "pandas is not installed.\nRun: pip install pandas openpyxl",
    )
    raise SystemExit(1)


# ---------------------------------------------------------------------------
# Configuration: output requests and PARAM cards
# ---------------------------------------------------------------------------

# Each tuple: (card_text, short_name, description, default_checked)
OUTPUT_REQUEST_OPTIONS: list[tuple[str, str, str, bool]] = [
    ("DISPLACEMENT(SORT1,PLOT,REAL)=ALL",  "DISPLACEMENT", "Düğüm noktası yer değiştirmeleri",    True),
    ("FORCE(SORT1,PLOT,REAL,CENTER)=ALL",  "FORCE",        "Eleman kuvvetleri (merkez)",           True),
    ("GPFORCE(PLOT)=ALL",                   "GPFORCE",      "Grid noktası kuvvet dengesi",          True),
    ("OLOAD(PLOT)=ALL",                     "OLOAD",        "Uygulanan dış yükler",                 True),
    ("SPCFORCE(SORT1,PLOT)=ALL",            "SPCFORCE",     "SPC reaksiyon kuvvetleri",             True),
    ("STRESS(SORT1,PLOT,REAL)=ALL",         "STRESS",       "Eleman gerilmeleri",                   False),
    ("STRAIN(SORT1,PLOT,REAL)=ALL",         "STRAIN",       "Eleman şekil değiştirmeleri",          False),
    ("MPCFORCE(SORT1,PLOT)=ALL",            "MPCFORCE",     "MPC reaksiyon kuvvetleri",             False),
]

# Each tuple: (keyword, name, card_description, [(value, value_description), ...], default_value, default_checked)
PARAM_OPTIONS: list[tuple] = [
    ("PARAM", "AUTOSPC", "Singüler DOF'ları otomatik sabitleme", [
        ("NO",  "Singüler DOF'ları olduğu gibi bırak — analiz uyarı verir, sonuç güvenilir"),
        ("YES", "Otomatik SPC ekler — analizi tamamlar, sonuç doğruluğu değişebilir"),
    ], "NO", True),

    ("PARAM", "POST", "Post-processing çıktı formatı", [
        ("-2", "OP2 + XDB — her iki formatta çıktı üret"),
        ("-1", "OP2 — en yaygın MSC/NX post-processing formatı"),
        ("0",  "Yok — post-processing çıktısı üretme"),
        ("1",  "OP2 — NX Nastran uyumlu (−1 ile işlevsel olarak aynı)"),
    ], "-1", True),

    ("PARAM", "K6ROT", "Kabuk elemanları membran sondaj rijitliği", [
        ("0.",   "Rijitlik yok — membran serbestçe döner (uyumsuzluk riski)"),
        ("1.",   "Hafif rijitlik — önerilen değer"),
        ("100.", "Yüksek rijitlik — aşırı kullanımda hatalı sonuç"),
    ], "1.", True),

    ("PARAM", "OUNIT2", "İkincil çıktı dosyası Fortran birimi", [
        ("6",  "Birim 6 — standart çıktı (stdout)"),
        ("11", "Birim 11"),
        ("12", "Birim 12 — standart ikincil çıktı birimi"),
    ], "12", True),

    ("PARAM", "OMID", "Eleman gerilme/kuvvet çıktı konumu", [
        ("YES", "Orta nokta — eleman merkezindeki değerleri çıkar"),
        ("NO",  "Düğüm noktaları — extrapolation ile hesaplanır"),
    ], "YES", True),

    ("PARAM", "PRTMAXIM", "Maksimum değer özeti tablosu", [
        ("YES", "Yazdır — analiz sonunda max değerleri özetler"),
        ("NO",  "Yazdırma"),
    ], "YES", True),

    ("PARAM", "BAILOUT", "Kritik hata toleransı", [
        ("0",  "İlk kritik hatada dur — güvenli mod"),
        ("-1", "Tüm hataları görmezden gel — sadece debug için kullan"),
    ], "0", True),

    ("PARAM", "OGEOM", "Geometri verisini çıktı dosyasına yaz", [
        ("YES", "Yaz — grid ve koordinat bilgilerini OP2'ye ekle"),
        ("NO",  "Yazma — daha küçük çıktı dosyası"),
    ], "YES", True),

    ("PARAM", "PRGPST", "Grid noktası gerilmesi çıktısı", [
        ("YES", "Yazdır — grid noktası gerilmelerini çıktıya ekle"),
        ("NO",  "Yazdırma"),
    ], "YES", True),

    ("PARAM", "POSTEXT", "Genişletilmiş NX çıktısı", [
        ("YES", "Aktif — NX'e özgü genişletilmiş çıktı formatını kullan"),
        ("NO",  "Pasif"),
    ], "YES", False),

    ("PARAM", "INREL", "Inertia relief — serbest-serbest statik analiz", [
        ("0",  "Kapalı — standart kısıtlı statik analiz (default)"),
        ("-1", "Otomatik — 6 rijit cisim modunu ortadan kaldırır (serbest GFEM için)"),
        ("-2", "Kullanıcı tanımlı — SUPORT kartıyla belirtilen mesnet noktaları"),
    ], "-1", False),

    ("PARAM", "WTMASS", "Ağırlık → kütle dönüşüm katsayısı (1/g)", [
        ("1.0",      "SI N-m-kg — dönüşüm gerekmez"),
        ("0.10197",  "N-m — ağırlık N cinsinden, g = 9.807 m/s²"),
        ("1.02e-4",  "N-mm — ağırlık N cinsinden, g = 9806.65 mm/s²"),
        ("0.00259",  "lbf-in — ağırlık lbf cinsinden, g = 386.04 in/s²"),
    ], "1.02e-4", False),

    ("MDLPRM", "HDF5", "HDF5 formatında ikincil çıktı", [
        ("0", "Kapalı — HDF5 dosyası oluşturma"),
        ("1", "Açık — .h5 uzantılı HDF5 çıktı dosyası oluştur"),
    ], "1", False),
]


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------

@dataclass
class LoadCombination:
    case_id: int
    components: list = field(default_factory=list)  # [(multiplier, unit_case_id), ...]


# ---------------------------------------------------------------------------
# Core logic
# ---------------------------------------------------------------------------

def read_case_ids_from_excel(excel_path: str, log_fn=None) -> set[int]:
    """Return unique non-zero integer case IDs from all *CASE ID columns excluding THERMAL."""
    if log_fn:
        log_fn("[EXCEL] Reading Combination Excel...")
    df = pd.read_excel(excel_path, engine="openpyxl")

    cols = list(df.columns)
    case_id_cols = [
        col for col in cols
        if "CASE ID" in str(col).upper() and "THERMAL" not in str(col).upper()
    ]
    if log_fn:
        log_fn(f"[EXCEL] Unit case columns ({len(case_id_cols)}): {', '.join(str(c) for c in case_id_cols)}")

    ids: set[int] = set()
    for col in case_id_cols:
        for val in df[col].dropna():
            try:
                int_val = int(val)
                if int_val != 0:
                    ids.add(int_val)
            except (ValueError, TypeError):
                pass
    return ids


def read_load_combinations(excel_path: str, log_fn=None) -> list[LoadCombination]:
    """
    Read each row of the Combination Excel as a LoadCombination.
    Column A = Combined Load Case ID.
    Each *CASE ID column (non-thermal) and its immediately following Multiplier column
    form one component pair.
    """
    df = pd.read_excel(excel_path, engine="openpyxl")
    cols = list(df.columns)

    combined_id_col = cols[0]

    # Build (case_id_col_index, multiplier_col_index) pairs, skip thermal
    pairs: list[tuple[int, int]] = []
    for i, col in enumerate(cols):
        col_upper = str(col).upper()
        if "CASE ID" in col_upper and "THERMAL" not in col_upper:
            if i + 1 < len(cols):
                pairs.append((i, i + 1))

    if log_fn:
        log_fn(f"[COMBO] {len(pairs)} non-thermal ID/multiplier column pairs found.")

    combinations: list[LoadCombination] = []
    for _, row in df.iterrows():
        try:
            cid = int(row.iloc[0])
        except (ValueError, TypeError):
            continue

        components: list[tuple[float, int]] = []
        for id_idx, mult_idx in pairs:
            try:
                unit_id = int(row.iloc[id_idx])
                mult = float(row.iloc[mult_idx])
                if unit_id != 0:
                    components.append((mult, unit_id))
            except (ValueError, TypeError):
                pass

        if components:
            combinations.append(LoadCombination(case_id=cid, components=components))

    if log_fn:
        log_fn(f"[COMBO] {len(combinations)} combined load cases read.")
    return combinations


def read_subcase_mapping(list_excel_path: str, log_fn=None) -> dict[int, str]:
    """
    Read the List Subcases Excel.
    Column A = FILE (relative path), Column B = SUBCASE_ID.
    Returns {subcase_id: relative_file_path}.
    """
    if log_fn:
        log_fn("[SUBCASES] Reading List Subcases Excel...")
    df = pd.read_excel(list_excel_path, engine="openpyxl")

    cols = list(df.columns)
    file_col = cols[0]
    id_col = cols[1]
    if log_fn:
        log_fn(f"[SUBCASES] FILE='{file_col}', SUBCASE_ID='{id_col}'")

    mapping: dict[int, str] = {}
    for _, row in df.iterrows():
        try:
            sid = int(row[id_col])
            fpath = str(row[file_col]).strip()
            mapping[sid] = fpath
        except (ValueError, TypeError):
            pass

    if log_fn:
        log_fn(f"[SUBCASES] {len(mapping)} subcase → file mappings loaded.")
    return mapping


def _build_bdf_file_list(base_dir: str, log_fn=None) -> list[str]:
    """Walk base_dir recursively and return full paths of all .bdf files."""
    paths: list[str] = []
    for root, _, files in os.walk(base_dir):
        for fname in files:
            if fname.lower().endswith(".bdf"):
                paths.append(os.path.join(root, fname))
    if log_fn:
        log_fn(f"[INDEX] {len(paths)} BDF files indexed under {base_dir}")
    return paths


def _extract_search_term(filename: str) -> str:
    """
    Determine what to search for inside base_dir given a List Subcases filename.

    Rules:
      MASTER_MANOEUVRE_28999121.bdf  →  '28999121'   (trailing number after last _)
      MASTER_CABIN.bdf               →  'CABIN'      (word after MASTER_ when no trailing number)
    """
    stem = os.path.splitext(filename)[0]
    # Trailing number: MASTER_TYPE_123456  →  '123456'
    m = re.search(r'_(\d+)$', stem)
    if m:
        return m.group(1)
    # No trailing number: extract token after MASTER_ (or MASTER)
    m2 = re.match(r'(?i)MASTER_?([A-Z]+)', stem)
    if m2:
        return m2.group(1).upper()
    return stem


def resolve_include_paths(
    case_ids: set[int],
    subcase_mapping: dict[int, str],
    base_dir: str,
    log_fn=None,
) -> tuple[list[str], list[int]]:
    """
    For each unit case ID:
      - extract a search term from its List Subcases filename
        (trailing number if present, otherwise the keyword after MASTER_)
      - find ALL .bdf files in base_dir whose name contains that term
      - add them to the INCLUDE list (deduplicated)
    Returns (ordered_unique_paths, missing_ids).
    """
    all_bdf = _build_bdf_file_list(base_dir, log_fn=log_fn)

    seen: set[str] = set()
    ordered_paths: list[str] = []
    missing_ids: list[int] = []

    for cid in sorted(case_ids):
        if cid not in subcase_mapping:
            missing_ids.append(cid)
            if log_fn:
                log_fn(f"[ERROR] Case ID {cid} not found in List Subcases Excel.")
            continue

        raw = subcase_mapping[cid].replace("\\", "/").split("/")[-1]
        term = _extract_search_term(raw)
        term_lower = term.lower()

        matches = [p for p in all_bdf if term_lower in os.path.basename(p).lower()]

        if not matches:
            missing_ids.append(cid)
            if log_fn:
                log_fn(f"[ERROR] {cid} ('{raw}') → no files found for search term '{term}'")
            continue

        new_count = 0
        for p in matches:
            norm = os.path.normpath(p)
            if norm not in seen:
                seen.add(norm)
                ordered_paths.append(norm)
                new_count += 1

        if log_fn:
            dup = len(matches) - new_count
            dup_str = f", {dup} already seen" if dup else ""
            log_fn(f"[OK]    {cid} (search='{term}') → {new_count} files added{dup_str}")

    return ordered_paths, missing_ids


def _resolve_load_collector_ids(
    combinations: list,
    subcase_mapping: dict,
) -> None:
    """
    Fix unit case IDs in LOAD entries to match actual load collector IDs in BDF files.

    When the List Subcases filename contains a trailing number
    (e.g. MASTER_MANOEUVRE_28999121.bdf → 28999121), the Combination Excel
    stores a sequential ID (e.g. 10079) that does NOT match the load collector
    in the BDF.  Replace it with the real number extracted from the filename.

    When there is no trailing number (e.g. MASTER_CABIN.bdf), the original
    unit case ID already matches the load collector ID in the BDF — keep it.
    """
    for combo in combinations:
        new_components = []
        for mult, unit_id in combo.components:
            raw = subcase_mapping.get(unit_id, "")
            filename = raw.replace("\\", "/").split("/")[-1]
            term = _extract_search_term(filename)
            if re.match(r'^\d+$', term):
                load_id = int(term)
            else:
                load_id = unit_id
            new_components.append((mult, load_id))
        combo.components = new_components


def _format_load_entry(case_id: int, components: list) -> list[str]:
    """
    Format a Nastran LOAD bulk entry with continuation lines.
    First line:  LOAD, SID, 1.0, S1,L1, S2,L2, S3,L3, [+]
    Cont. lines: +, S4,L4, S5,L5, S6,L6, S7,L7, [+]
    """
    FIRST_LINE_PAIRS = 3
    CONT_LINE_PAIRS = 4

    lines: list[str] = []
    remaining = list(components)

    chunk = remaining[:FIRST_LINE_PAIRS]
    remaining = remaining[FIRST_LINE_PAIRS:]

    parts = ["LOAD", str(case_id), "1.0"]
    for mult, uid in chunk:
        parts += [str(mult), str(uid)]
    if remaining:
        parts.append("+")
    lines.append(",".join(parts))

    while remaining:
        chunk = remaining[:CONT_LINE_PAIRS]
        remaining = remaining[CONT_LINE_PAIRS:]
        parts = ["+"]
        for mult, uid in chunk:
            parts += [str(mult), str(uid)]
        if remaining:
            parts.append("+")
        lines.append(",".join(parts))

    return lines


def _read_spc_id_from_bdf(path: str) -> int | None:
    """Return the first SPC/SPC1 set ID found in the BDF file."""
    with open(path, encoding="utf-8", errors="ignore") as f:
        for line in f:
            stripped = line.strip()
            if not stripped or stripped.startswith("$"):
                continue
            if "," in stripped:
                parts = [p.strip() for p in stripped.split(",")]
                if parts[0].upper() in ("SPC", "SPC1", "SPCD", "SPCADD"):
                    try:
                        return int(parts[1])
                    except (ValueError, IndexError):
                        pass
            else:
                kw = stripped[:8].strip().upper()
                if kw in ("SPC", "SPC1", "SPCD", "SPCADD"):
                    try:
                        return int(stripped[8:16].strip())
                    except (ValueError, IndexError):
                        pass
    return None


_NASTRAN_MAX_ID = 9_999_999   # Nastran SUBCASE/LOAD SID max 7 digits


def _ensure_bulk_only(source_path: str, output_dir: str, log_fn=None) -> str:
    """
    If source_path is a complete Nastran deck (contains a BEGIN BULK line),
    extract everything between BEGIN BULK and ENDDATA into a sidecar file
    placed in output_dir, then return the sidecar path.
    If the file is already bulk-only (no BEGIN BULK found), return source_path
    unchanged so it is INCLUDE'd as-is.
    """
    bulk_lines: list[str] = []
    in_bulk = False
    found = False

    with open(source_path, encoding="utf-8", errors="ignore") as fh:
        for raw in fh:
            line = raw.rstrip("\r\n")
            tokens = line.upper().split()
            if not in_bulk:
                # Match "BEGIN BULK" (with or without trailing SUPER/AUXMODEL args)
                if len(tokens) >= 2 and tokens[0] == "BEGIN" and tokens[1] == "BULK":
                    in_bulk = True
                    found = True
                    continue
            else:
                if tokens and tokens[0] == "ENDDATA":
                    break
                bulk_lines.append(line)

    if not found:
        return source_path  # already bulk-only

    basename = os.path.splitext(os.path.basename(source_path))[0]
    sidecar = os.path.join(output_dir, f"{basename}_bulk.bdf")
    with open(sidecar, "w", encoding="utf-8") as fh:
        fh.write("\n".join(bulk_lines) + "\n")

    if log_fn:
        log_fn(
            f"[BULK] '{os.path.basename(source_path)}' tam deck — "
            f"{len(bulk_lines)} satır bulk data çıkarıldı → '{os.path.basename(sidecar)}'"
        )
    return sidecar


def _format_include_lines(path: str, max_len: int = 64) -> list[str]:
    """
    Format an INCLUDE statement splitting at '/' boundaries so that
    every output line is at most max_len characters.
    Trailing slash is placed at the END of each intermediate line so
    continuation lines always start with a directory/file name.
    A single segment that is itself longer than max_len is written as-is
    on its own line (cannot split within a name).
    """
    path = path.replace("\\", "/")
    if len(f"INCLUDE '{path}'") <= max_len:
        return [f"INCLUDE '{path}'"]

    parts = path.split("/")
    # Each non-last part keeps its trailing slash
    segments = [p + "/" for p in parts[:-1]] + [parts[-1]]

    result: list[str] = []
    current = "INCLUDE '"

    for seg in segments:
        if current == "INCLUDE '":
            # First segment — always place it regardless of length
            current += seg
        elif len(current + seg) <= max_len:
            current += seg
        else:
            result.append(current)
            current = seg

    result.append(current + "'")
    return result


def write_output_bdf(
    gfem_path: str,
    include_paths: list[str],
    combinations: list[LoadCombination],
    solver: str,
    output_path: str,
    output_requests: list[str],
    param_cards: list[str],
    spc_bdf_path: str = "",
    spc_id: int = 0,
    log_fn=None,
) -> int:
    """Write a complete Nastran solution deck. Returns total INCLUDE count."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    solver_label = "NX Nastran" if solver == "NX" else "MSC Nastran"
    lines: list[str] = []

    # SUBCASEs must be in ascending order (Nastran rule)
    sorted_combos = sorted(combinations, key=lambda c: c.case_id)

    # --- Build 7-digit SUBCASE/LOAD SID map ---
    id_map: dict[int, int] = {}
    reverse_map: dict[int, int] = {}  # remapped → original (for collision detection)
    for combo in sorted_combos:
        orig = combo.case_id
        remapped = orig if orig <= _NASTRAN_MAX_ID else orig % 10_000_000
        if remapped in reverse_map and reverse_map[remapped] != orig:
            if log_fn:
                log_fn(
                    f"[WARN] ID collision: {reverse_map[remapped]} and {orig} "
                    f"both remap to {remapped} — manual fix required"
                )
        id_map[orig] = remapped
        reverse_map[remapped] = orig

    # --- Ensure all included files contain only bulk data ---
    output_dir = os.path.dirname(os.path.abspath(output_path))
    gfem_inc     = _ensure_bulk_only(gfem_path,    output_dir, log_fn)
    spc_inc      = _ensure_bulk_only(spc_bdf_path, output_dir, log_fn) if spc_bdf_path else ""
    unit_incs    = [_ensure_bulk_only(p, output_dir, log_fn) for p in include_paths]

    # --- File header ---
    lines += [
        f"$ Generated by GFEM BDF Combination Tool",
        f"$ Date: {timestamp}",
        f"$ Solver: {solver_label}",
        "$",
    ]

    # --- Executive Control ---
    lines += [
        "$ === EXECUTIVE CONTROL ===",
        "SOL 101",
        "CEND",
        "$",
    ]

    # --- Case Control ---
    lines += [
        "$ === CASE CONTROL ===",
        "TITLE = GFEM Combined Load Cases",
        "$",
    ]

    # Global output requests (before first SUBCASE)
    if output_requests:
        lines.append("$ Global Output Requests")
        lines.extend(output_requests)
        lines.append("$")

    for combo in sorted_combos:
        orig = combo.case_id
        sid = id_map[orig]
        title_comment = f"  $ Original: {orig}" if sid != orig else ""
        lines.append(f"SUBCASE {sid}")
        lines.append(f"  TITLE = Combined Case {sid}{title_comment}")
        if spc_id:
            lines.append(f"  SPC = {spc_id}")
        lines.append(f"  LOAD = {sid}")
        lines.append("$")

    # --- Bulk Data ---
    lines += [
        "BEGIN BULK",
        "$",
    ]

    # User-selected PARAM cards
    if param_cards:
        lines.append("$ PARAM Cards")
        lines.extend(param_cards)
        lines.append("$")

    # GFEM model
    lines.append("$ GFEM Model")
    lines.extend(_format_include_lines(gfem_inc))
    lines.append("$")

    # SPC BDF (optional)
    if spc_inc:
        lines.append("$ SPC Constraints")
        lines.extend(_format_include_lines(spc_inc))
        lines.append("$")

    # Unit case INCLUDEs
    lines.append("$ Unit Case Loads")
    for p in unit_incs:
        lines.extend(_format_include_lines(p))
    lines.append("$")

    # LOAD entries (same sorted order)
    lines.append("$ Combined LOAD Entries")
    for combo in sorted_combos:
        orig = combo.case_id
        sid = id_map[orig]
        orig_comment = f" (orig: {orig})" if sid != orig else ""
        lines.append(f"$ --- Case {sid}{orig_comment} ---")
        lines.extend(_format_load_entry(sid, combo.components))
        lines.append("$")

    lines.append("ENDDATA")

    with open(output_path, "w") as f:
        f.write("\n".join(lines) + "\n")

    include_count = 1 + len(include_paths) + (1 if spc_bdf_path else 0)
    if log_fn:
        log_fn(
            f"[WRITE] {output_path}  "
            f"({include_count} INCLUDEs, {len(sorted_combos)} SUBCASEs, {len(sorted_combos)} LOAD entries)"
        )
    return include_count


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("GFEM BDF Combination Tool")
        self.resizable(True, True)
        self._log_queue: queue.Queue = queue.Queue()
        self._status_queue: queue.Queue = queue.Queue()
        self._build_ui()
        self.after(100, self._poll_queues)

    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}
        frame = ttk.Frame(self, padding=10)
        frame.grid(row=0, column=0, sticky="nsew")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)

        fields = [
            ("GFEM BDF File:",        "gfem_var",     self._browse_gfem),
            ("Combination Excel:",    "excel_var",    self._browse_excel),
            ("List Subcases Excel:",  "subcases_var", self._browse_subcases),
            ("Unit Case Base Dir:",   "dir_var",      self._browse_dir),
            ("Output BDF:",           "output_var",   self._browse_output),
        ]
        for row, (label, attr, cmd) in enumerate(fields):
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="e", **pad)
            var = tk.StringVar()
            setattr(self, attr, var)
            ttk.Entry(frame, textvariable=var, width=60).grid(row=row, column=1, sticky="ew", **pad)
            ttk.Button(frame, text="Browse", command=cmd).grid(row=row, column=2, **pad)

        # --- SPC ---
        spc_row = len(fields)
        ttk.Label(frame, text="SPC BDF File:").grid(row=spc_row, column=0, sticky="e", **pad)
        self.spc_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.spc_var, width=60).grid(
            row=spc_row, column=1, sticky="ew", **pad)
        ttk.Button(frame, text="Browse", command=self._browse_spc).grid(
            row=spc_row, column=2, **pad)
        self._spc_id_label = ttk.Label(
            frame, text="(opsiyonel)", foreground="#888", font=("Courier", 8))
        self._spc_id_label.grid(row=spc_row, column=3, sticky="w", padx=4)

        _SPC_ROWS = 1

        # --- Solver selection ---
        _R = len(fields) + _SPC_ROWS
        solver_frame = ttk.LabelFrame(frame, text="Nastran Solver", padding=6)
        solver_frame.grid(row=_R, column=0, columnspan=3, sticky="ew", padx=8, pady=4)
        self.solver_var = tk.StringVar(value="NX")
        ttk.Radiobutton(solver_frame, text="NX Nastran",  variable=self.solver_var, value="NX").pack(side="left", padx=12)
        ttk.Radiobutton(solver_frame, text="MSC Nastran", variable=self.solver_var, value="MSC").pack(side="left", padx=12)

        # --- Global Output Requests ---
        out_frame = ttk.LabelFrame(frame, text="Global Output Requests", padding=6)
        out_frame.grid(row=_R + 1, column=0, columnspan=3, sticky="ew", padx=8, pady=4)
        self.output_req_vars: list[tk.BooleanVar] = []
        for i, (card, name, desc, default) in enumerate(OUTPUT_REQUEST_OPTIONS):
            var = tk.BooleanVar(value=default)
            self.output_req_vars.append(var)
            row_f = ttk.Frame(out_frame)
            row_f.grid(row=i // 2, column=(i % 2) * 2, sticky="w", padx=(0, 16))
            ttk.Checkbutton(row_f, text=name, variable=var, width=14).pack(side="left")
            ttk.Label(row_f, text=desc, foreground="#555", font=("", 7)).pack(side="left")

        # --- PARAM Cards ---
        param_frame = ttk.LabelFrame(frame, text="PARAM Cards", padding=6)
        param_frame.grid(row=_R + 2, column=0, columnspan=3, sticky="ew", padx=8, pady=4)
        param_frame.columnconfigure(4, weight=1)
        self.param_vars: list[tk.BooleanVar] = []
        self.param_value_vars: list[tk.StringVar] = []
        for i, (kw, name, _card_desc, value_opts, default_val, default_checked) in enumerate(PARAM_OPTIONS):
            include_var = tk.BooleanVar(value=default_checked)
            value_var = tk.StringVar(value=default_val)
            self.param_vars.append(include_var)
            self.param_value_vars.append(value_var)

            ttk.Checkbutton(param_frame, variable=include_var).grid(
                row=i, column=0, sticky="w", padx=(0, 2))
            ttk.Label(param_frame, text=f"{kw}  {name}", width=16,
                      anchor="e", font=("Courier", 8, "bold")).grid(row=i, column=1, sticky="e")
            ttk.Label(param_frame, text="=").grid(row=i, column=2, padx=2)

            combo = ttk.Combobox(
                param_frame, textvariable=value_var,
                values=[v for v, _ in value_opts],
                state="readonly", width=7,
            )
            combo.grid(row=i, column=3, sticky="w", padx=(0, 8))

            desc_lbl = ttk.Label(param_frame, text="", foreground="#555", font=("", 7))
            desc_lbl.grid(row=i, column=4, sticky="w")

            vdict = dict(value_opts)
            def _make_updater(vv=value_var, vd=vdict, lbl=desc_lbl):
                def _upd(*_):
                    lbl.configure(text=vd.get(vv.get(), ""))
                return _upd
            upd = _make_updater()
            value_var.trace_add("write", upd)
            upd()

        # --- Generate button ---
        self._gen_btn = ttk.Button(frame, text="Generate BDF", command=self._on_generate)
        self._gen_btn.grid(row=_R + 3, column=0, columnspan=3, pady=8)

        # --- Progress bar ---
        self._progress = ttk.Progressbar(frame, mode="indeterminate")
        self._progress.grid(row=_R + 4, column=0, columnspan=3, sticky="ew", padx=8)

        # --- Status label ---
        self._status_var = tk.StringVar(value="Ready.")
        ttk.Label(
            frame, textvariable=self._status_var,
            foreground="#005580", font=("Courier", 8), anchor="w"
        ).grid(row=_R + 5, column=0, columnspan=3, sticky="ew", padx=8, pady=(2, 0))

        # --- Log ---
        ttk.Label(frame, text="Log:").grid(row=_R + 6, column=0, sticky="w", padx=8)
        self._log_widget = scrolledtext.ScrolledText(
            frame, height=12, state="disabled", wrap="word", font=("Courier", 8)
        )
        self._log_widget.grid(row=_R + 7, column=0, columnspan=3, sticky="nsew", padx=8, pady=4)
        frame.rowconfigure(_R + 7, weight=1)

    # --- Browse helpers ---

    def _browse_gfem(self):
        p = filedialog.askopenfilename(
            title="Select GFEM BDF File",
            filetypes=[("BDF files", "*.bdf *.BDF"), ("All files", "*.*")],
        )
        if p:
            self.gfem_var.set(p)

    def _browse_excel(self):
        p = filedialog.askopenfilename(
            title="Select Combination Excel",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if p:
            self.excel_var.set(p)

    def _browse_subcases(self):
        p = filedialog.askopenfilename(
            title="Select List Subcases Excel",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if p:
            self.subcases_var.set(p)

    def _browse_dir(self):
        p = filedialog.askdirectory(title="Select Unit Case Base Directory")
        if p:
            self.dir_var.set(p)

    def _browse_output(self):
        p = filedialog.asksaveasfilename(
            title="Save Output BDF As",
            defaultextension=".bdf",
            filetypes=[("BDF files", "*.bdf"), ("All files", "*.*")],
        )
        if p:
            self.output_var.set(p)

    def _browse_spc(self):
        p = filedialog.askopenfilename(
            title="Select SPC BDF File",
            filetypes=[("BDF files", "*.bdf *.BDF"), ("All files", "*.*")],
        )
        if p:
            self.spc_var.set(p)
            spc_id = _read_spc_id_from_bdf(p)
            if spc_id is not None:
                self._spc_id_label.configure(
                    text=f"→ SPC ID: {spc_id}", foreground="#005580")
            else:
                self._spc_id_label.configure(
                    text="⚠ SPC ID bulunamadı", foreground="#cc0000")

    # --- Generate ---

    def _on_generate(self):
        gfem     = self.gfem_var.get().strip()
        excel    = self.excel_var.get().strip()
        subcases = self.subcases_var.get().strip()
        base_dir = self.dir_var.get().strip()
        output   = self.output_var.get().strip()
        solver   = self.solver_var.get()
        spc_bdf  = self.spc_var.get().strip()

        errors = []
        if not gfem or not os.path.isfile(gfem):
            errors.append("GFEM BDF file not found.")
        if not excel or not os.path.isfile(excel):
            errors.append("Combination Excel file not found.")
        if not subcases or not os.path.isfile(subcases):
            errors.append("List Subcases Excel file not found.")
        if not base_dir or not os.path.isdir(base_dir):
            errors.append("Unit case base directory not found.")
        if not output:
            errors.append("Please specify an output BDF path.")
        if spc_bdf and not os.path.isfile(spc_bdf):
            errors.append("SPC BDF file not found.")

        # Auto-read SPC ID from file
        spc_id = 0
        if spc_bdf and os.path.isfile(spc_bdf):
            spc_id = _read_spc_id_from_bdf(spc_bdf) or 0
            if not spc_id:
                errors.append("SPC BDF file selected but no SPC/SPC1 card found inside.")
        if errors:
            messagebox.showerror("Input Error", "\n".join(errors))
            return

        sel_outputs = [
            card
            for (card, _, _, _), var in zip(OUTPUT_REQUEST_OPTIONS, self.output_req_vars)
            if var.get()
        ]
        sel_params = [
            f"{kw},{name},{value_var.get()}"
            for (kw, name, *_rest), include_var, value_var
            in zip(PARAM_OPTIONS, self.param_vars, self.param_value_vars)
            if include_var.get()
        ]

        self._gen_btn.configure(state="disabled")
        self._progress.start(10)

        threading.Thread(
            target=self._run_generation,
            args=(gfem, excel, subcases, base_dir, output, solver,
                  sel_outputs, sel_params, spc_bdf, spc_id),
            daemon=True,
        ).start()

    def _run_generation(self, gfem, excel, subcases, base_dir, output, solver,
                        sel_outputs, sel_params, spc_bdf, spc_id):
        try:
            self._status("Reading Combination Excel...")
            case_ids = read_case_ids_from_excel(excel, log_fn=self._log)
            self._log(f"[INFO] {len(case_ids)} unique unit case IDs (thermal excluded).")

            self._status("Reading load combinations...")
            combinations = read_load_combinations(excel, log_fn=self._log)

            self._status("Reading List Subcases Excel...")
            subcase_mapping = read_subcase_mapping(subcases, log_fn=self._log)

            self._status("Resolving load collector IDs...")
            _resolve_load_collector_ids(combinations, subcase_mapping)

            self._status("Resolving file paths...")
            self._log("-" * 60)
            include_paths, missing = resolve_include_paths(
                case_ids, subcase_mapping, base_dir, log_fn=self._log
            )
            self._log("-" * 60)
            self._log(
                f"[INFO] {len(include_paths)} unique files to include. "
                f"{len(missing)} IDs not found in subcases list."
            )
            if missing:
                self._log(f"[ERROR] Missing IDs: {sorted(missing)}")

            solver_label = "NX Nastran" if solver == "NX" else "MSC Nastran"
            self._status(f"Writing {solver_label} deck...")
            write_output_bdf(
                gfem, include_paths, combinations, solver, output,
                output_requests=sel_outputs,
                param_cards=sel_params,
                spc_bdf_path=spc_bdf,
                spc_id=spc_id,
                log_fn=self._log,
            )
            self._log("[DONE] Generation complete.")
            self._status("Done.")

        except Exception as exc:
            self._log(f"[ERROR] {exc}")
            self._status(f"ERROR: {exc}")
            self.after(0, lambda: messagebox.showerror("Error", str(exc)))
        finally:
            self.after(0, self._generation_done)

    def _generation_done(self):
        self._progress.stop()
        self._gen_btn.configure(state="normal")

    # --- Thread-safe log & status ---

    def _log(self, msg: str):
        self._log_queue.put(msg + "\n")

    def _status(self, msg: str):
        self._status_queue.put(msg)

    def _log_write(self, msg: str):
        self._log_widget.configure(state="normal")
        self._log_widget.insert(tk.END, msg)
        self._log_widget.see(tk.END)
        self._log_widget.configure(state="disabled")

    def _poll_queues(self):
        while not self._log_queue.empty():
            self._log_write(self._log_queue.get_nowait())
        while not self._status_queue.empty():
            self._status_var.set(self._status_queue.get_nowait())
        self.after(100, self._poll_queues)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app = App()
    app.mainloop()
