#!/usr/bin/env python3
"""
SQL Server Parallel Striped Database Restore GUI (`sql_restore_gui.py`)
======================================================================
A self-contained graphical utility for restoring a batch of single-file or
multi-stream striped SQL Server databases (`.bak` sets and `.zip` archives)
as quickly as possible.

Key Capabilities:
- Connects via Windows Authentication or SQL Server Authentication.
- Queries SQL Server instance default Data & Log paths (`InstanceDefaultDataPath`,
  `InstanceDefaultLogPath`) and Instant File Initialization (`IFI`) status.
- Hybrid Backup Discovery:
  - Scans a source folder (recursively) for single `.bak` files, multi-file striped
    `.bak` sets (`_stripe01of04.bak`, `_part1of4.bak`, etc.), and `.zip` archives
    produced by `sql_backup_gui`.
  - Automatically groups striped files into unified restore sets and optionally filters
    to only the newest backup timestamp per database.
  - Inspects SQL backup headers (`RESTORE HEADERONLY` & `RESTORE FILELISTONLY`) to
    verify stripe completeness and discover exact logical Data/Log file names.
- High-Speed Parallel Execution ("As Quickly As Possible"):
  - Largest-Database-First (LPT) scheduling so multi-stripe databases begin immediately.
  - Parallel database restore worker pool (default 3 concurrent DBs, up to 16).
  - Multi-threaded fast `.zip` extraction (`7z.exe x -mmt=on` or 16 MB streaming buffer)
    into a staging folder with automatic post-restore `.bak` cleanup.
  - High-throughput T-SQL options: `MAXTRANSFERSIZE = 4194304` (4 MB) and auto-scaled
    `BUFFERCOUNT` (`max(24, streams * 8)`), with automatic retry fallback if SQL Server
    encounters contiguous virtual memory pressure.
  - Safe overwrite handling: `SINGLE_USER WITH ROLLBACK IMMEDIATE`, `WITH REPLACE`,
    automatic `WITH MOVE` file relocation, and optional Target DB Prefix/Suffix or per-DB rename.
"""

from __future__ import annotations

import concurrent.futures
import datetime
import json
import math
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import time
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

# Optional pyodbc detection
try:
    import pyodbc  # type: ignore

    HAS_PYODBC = True
except ImportError:
    HAS_PYODBC = False


APP_TITLE = "SQL Server Parallel Striped Database Restore Utility"
CONFIG_FILENAME = ".sql_restore_config.json"
DEFAULT_MAX_PARALLEL_DBS = 3
DEFAULT_MAX_TRANSFER_SIZE = 4194304  # 4 MB (SQL Server max transfer size for striped restores)
FAST_UNZIP_CHUNK_BYTES = 16 * 1024 * 1024  # 16 MB streaming buffer for fast I/O
STAGING_SUBFOLDER_NAME = "_RestoreStaging"


# ============================================================================
# Data Models
# ============================================================================


@dataclass
class SqlConnectionConfig:
    server: str = "localhost"
    auth_mode: str = "windows"  # "windows" or "sql"
    username: str = ""
    password: str = ""
    trust_server_certificate: bool = True
    connect_timeout: int = 15

    def to_dotnet_connection_string(self, database: str = "master") -> str:
        parts = [
            f"Server={self.server.strip()}",
            f"Database={database}",
            f"Connection Timeout={self.connect_timeout}",
            "Application Name=SqlParallelStripedRestoreGui",
        ]
        if self.trust_server_certificate:
            parts.append("TrustServerCertificate=True")
        if self.auth_mode.lower() == "windows":
            parts.append("Integrated Security=True")
        else:
            uid = self.username.strip()
            pwd = self.password.replace(";", ";;")
            parts.append("Integrated Security=False")
            parts.append(f"User ID={uid}")
            parts.append(f"Password={pwd}")
        return ";".join(parts) + ";"


@dataclass
class SqlInstanceInfo:
    server_name: str = "localhost"
    edition: str = "SQL Server"
    product_version: str = ""
    default_data_path: str = r"C:\SqlData"
    default_log_path: str = r"C:\SqlLogs"
    ifi_enabled: str = "Unknown"
    existing_databases: List[str] = field(default_factory=list)


@dataclass
class LogicalFileInfo:
    logical_name: str
    physical_name: str
    file_type: str = "D"  # "D" = Data (.mdf/.ndf), "L" = Log (.ldf)
    size_bytes: int = 0
    file_id: int = 1


@dataclass
class DiscoveredBackupSet:
    set_id: str
    source_db_name: str
    timestamp_label: str
    source_type: str  # "bak" or "zip"
    stripe_paths: List[str] = field(default_factory=list)
    zip_path: str = ""
    expected_stripes: int = 1
    found_stripes: int = 1
    total_size_bytes: int = 0
    mtime_epoch: float = 0.0
    logical_files: List[LogicalFileInfo] = field(default_factory=list)
    header_verified: bool = False
    custom_target_name: str = ""

    @property
    def size_mb(self) -> float:
        return round(self.total_size_bytes / (1024.0 * 1024.0), 2)

    @property
    def size_gb(self) -> float:
        return round(self.total_size_bytes / (1024.0 * 1024.0 * 1024.0), 3)

    @property
    def stream_count(self) -> int:
        return max(1, len(self.stripe_paths), self.expected_stripes)

    @property
    def is_complete(self) -> bool:
        return self.found_stripes >= self.expected_stripes and self.found_stripes > 0


@dataclass
class RestoreOptions:
    source_folder: str
    staging_folder: str = ""
    target_data_folder: str = r"C:\SqlData"
    target_log_folder: str = r"C:\SqlLogs"
    relocate_files: bool = True
    target_db_prefix: str = ""
    target_db_suffix: str = ""
    max_parallel_dbs: int = DEFAULT_MAX_PARALLEL_DBS
    with_replace: bool = True
    kill_connections: bool = True
    recovery_mode: str = "RECOVERY"  # "RECOVERY" or "NORECOVERY"
    high_speed_tuning: bool = True
    max_transfer_size: int = DEFAULT_MAX_TRANSFER_SIZE
    custom_buffer_count: int = 0  # 0 = Auto (max(24, streams * 8))
    keep_cdc: bool = False
    cleanup_staging_bak: bool = True
    latest_only_per_db: bool = True
    stats_interval: int = 5


@dataclass
class RestoreJobResult:
    set_id: str
    source_db_name: str
    target_db_name: str
    success: bool
    stream_count: int
    source_type: str = "bak"
    elapsed_seconds: float = 0.0
    unzip_elapsed_seconds: float = 0.0
    used_high_speed_buffers: bool = True
    cleaned_staging: bool = False
    error_message: str = ""


# ============================================================================
# Pure Helpers: Stripe Grouping, Fast-Unzip, WITH MOVE & T-SQL Generation
# ============================================================================

# Matches patterns like:
#   DbName_20261021_100000_stripe01of04.bak
#   DbName_20261021_100000.bak
#   DbName_stripe1of4.bak / DbName_part02of08.bak / DbName_1of4.bak
_STRIPE_REGEX = re.compile(
    r"^(?P<stem>.+?)(?:_(?P<ts>\d{8}_\d{6}|\d{8}|\d{14}))?"
    r"(?:_(?:stripe|part|stream|file|s)?(?P<idx>\d{1,3})of(?P<total>\d{1,3}))?$",
    re.IGNORECASE,
)


def sanitize_filename(name: str) -> str:
    """Sanitize a database name for safe Windows file and folder naming."""
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name.strip())
    cleaned = cleaned.rstrip(". ")
    return cleaned or "database"


def join_windows_or_native_path(base_folder: str, *parts: str) -> str:
    """
    Join path segments preserving Windows backslash conventions when base_folder
    looks like a Windows/UNC path (e.g. D:\\Backups or \\\\server\\share).
    """
    base = base_folder.strip().rstrip("\\/")
    sep = "\\" if ("\\" in base or (len(base) >= 2 and base[1] == ":")) else os.sep
    cleaned_parts = [p.strip("\\/") for p in parts if p.strip("\\/")]
    if not cleaned_parts:
        return base
    return base + sep + sep.join(cleaned_parts)


def parse_backup_filename(filename: str) -> Tuple[str, str, int, int]:
    """
    Parse a `.bak` or `.zip` filename into `(db_name, timestamp_label, stripe_index, stripe_total)`.
    Examples:
      - `EnterpriseWarehouseDB_20261021_100000_stripe01of04.bak`
        -> `("EnterpriseWarehouseDB", "20261021_100000", 1, 4)`
      - `SalesDB_20261021_100000.bak`
        -> `("SalesDB", "20261021_100000", 1, 1)`
      - `OrdersDB_20261021_110000.zip`
        -> `("OrdersDB", "20261021_110000", 1, 1)`
      - `Reporting_part2of3.bak`
        -> `("Reporting", "", 2, 3)`
    """
    base = os.path.basename(filename)
    stem, _ext = os.path.splitext(base)
    m = _STRIPE_REGEX.match(stem)
    if not m:
        return stem, "", 1, 1

    db_name = (m.group("stem") or stem).strip("_- ")
    ts = m.group("ts") or ""
    idx_str = m.group("idx")
    total_str = m.group("total")
    if idx_str and total_str:
        idx = max(1, int(idx_str))
        total = max(idx, int(total_str))
    else:
        idx, total = 1, 1
    return db_name or stem, ts, idx, total


def resolve_staging_folder(source_folder: str, custom_staging_folder: str = "") -> str:
    """Return the effective staging folder for extracting `.zip` backup archives."""
    if custom_staging_folder and custom_staging_folder.strip():
        return custom_staging_folder.strip()
    return join_windows_or_native_path(source_folder, STAGING_SUBFOLDER_NAME)


def resolve_target_db_name(
    source_db_name: str,
    prefix: str = "",
    suffix: str = "",
    custom_target_name: str = "",
) -> str:
    """Compute the effective target database name on the destination SQL Server."""
    if custom_target_name and custom_target_name.strip():
        return custom_target_name.strip()
    return f"{prefix.strip()}{source_db_name}{suffix.strip()}"


def calculate_optimal_buffercount(stream_count: int, custom_buffer_count: int = 0) -> int:
    """
    Calculate high-speed `BUFFERCOUNT` for striped SQL Server restores.
    - If `custom_buffer_count > 0`, returns `custom_buffer_count`.
    - Otherwise scales with stripe count: `min(128, max(24, stream_count * 8))`.
    """
    if custom_buffer_count > 0:
        return int(custom_buffer_count)
    streams = max(1, int(stream_count))
    return min(128, max(24, streams * 8))


def default_logical_files_for_db(db_name: str) -> List[LogicalFileInfo]:
    """Provide default Data & Log logical file metadata prior to SQL `RESTORE FILELISTONLY`."""
    return [
        LogicalFileInfo(
            logical_name=db_name,
            physical_name=f"{db_name}.mdf",
            file_type="D",
            file_id=1,
        ),
        LogicalFileInfo(
            logical_name=f"{db_name}_log",
            physical_name=f"{db_name}_log.ldf",
            file_type="L",
            file_id=2,
        ),
    ]


def scan_backup_source_folder(
    source_folder: str,
    latest_only_per_db: bool = True,
    recursive: bool = True,
) -> List[DiscoveredBackupSet]:
    """
    Scan `source_folder` for raw `.bak` sets (single or striped) and `.zip` archives
    containing `.bak` files. Groups stripes together and optionally filters to the
    latest backup set per database.
    """
    if not source_folder or not os.path.isdir(source_folder):
        return []

    raw_bak_groups: Dict[Tuple[str, str, str, int], List[Tuple[int, str, int, float]]] = {}
    discovered_sets: List[DiscoveredBackupSet] = []

    walker = os.walk(source_folder) if recursive else [(source_folder, [], os.listdir(source_folder))]
    for root, dirs, files in walker:
        # Never scan the temporary staging directory
        dirs[:] = [d for d in dirs if d.lower() != STAGING_SUBFOLDER_NAME.lower()]

        for fname in sorted(files):
            full_path = os.path.join(root, fname)
            if not os.path.isfile(full_path):
                continue
            lower_name = fname.lower()

            if lower_name.endswith(".bak"):
                db_name, ts, idx, total = parse_backup_filename(fname)
                try:
                    st = os.stat(full_path)
                    fsize = st.st_size
                    mtime = st.st_mtime
                except OSError:
                    fsize, mtime = 0, 0.0

                key = (os.path.normcase(root), db_name, ts, total)
                raw_bak_groups.setdefault(key, []).append((idx, full_path, fsize, mtime))

            elif lower_name.endswith(".zip"):
                try:
                    st = os.stat(full_path)
                    mtime = st.st_mtime
                    with zipfile.ZipFile(full_path, "r") as zf:
                        bak_infos = [
                            info
                            for info in zf.infolist()
                            if not info.is_dir() and info.filename.lower().endswith(".bak")
                        ]
                except Exception:
                    continue

                if not bak_infos:
                    continue

                # Group .bak entries inside the zip archive
                zip_db_name, zip_ts, _, _ = parse_backup_filename(fname)
                parsed_entries: List[Tuple[int, str, int, int]] = []
                for info in bak_infos:
                    entry_db, entry_ts, idx, total = parse_backup_filename(info.filename)
                    if entry_db:
                        zip_db_name = entry_db
                    if entry_ts:
                        zip_ts = entry_ts
                    parsed_entries.append((idx, info.filename, info.file_size, total))

                parsed_entries.sort(key=lambda x: (x[0], x[1]))
                expected = max(x[3] for x in parsed_entries)
                stripe_names = [x[1] for x in parsed_entries]
                total_uncompressed = sum(x[2] for x in parsed_entries)
                ts_label = zip_ts or datetime.datetime.fromtimestamp(mtime).strftime("%Y%m%d_%H%M%S")
                set_id = f"zip::{full_path}"
                discovered_sets.append(
                    DiscoveredBackupSet(
                        set_id=set_id,
                        source_db_name=zip_db_name,
                        timestamp_label=ts_label,
                        source_type="zip",
                        stripe_paths=stripe_names,
                        zip_path=full_path,
                        expected_stripes=expected,
                        found_stripes=len(stripe_names),
                        total_size_bytes=total_uncompressed,
                        mtime_epoch=mtime,
                        logical_files=default_logical_files_for_db(zip_db_name),
                    )
                )

    for (_folder_norm, db_name, ts, expected_total), items in raw_bak_groups.items():
        items.sort(key=lambda x: (x[0], x[1]))
        # Deduplicate identical stripe index if needed, preserving order
        stripe_paths = [x[1] for x in items]
        total_bytes = sum(x[2] for x in items)
        latest_mtime = max((x[3] for x in items), default=0.0)
        ts_label = ts or (
            datetime.datetime.fromtimestamp(latest_mtime).strftime("%Y%m%d_%H%M%S")
            if latest_mtime > 0
            else "Unknown"
        )
        set_id = f"bak::{stripe_paths[0]}"
        discovered_sets.append(
            DiscoveredBackupSet(
                set_id=set_id,
                source_db_name=db_name,
                timestamp_label=ts_label,
                source_type="bak",
                stripe_paths=stripe_paths,
                zip_path="",
                expected_stripes=expected_total,
                found_stripes=len(stripe_paths),
                total_size_bytes=total_bytes,
                mtime_epoch=latest_mtime,
                logical_files=default_logical_files_for_db(db_name),
            )
        )

    # Sort sets by database name ascending, then newest timestamp/mtime descending
    discovered_sets.sort(
        key=lambda s: (
            s.source_db_name.lower(),
            -(1 if s.is_complete else 0),
            s.timestamp_label,
            s.mtime_epoch,
        ),
        reverse=True,
    )
    # Re-sort primary key alphabetically by source_db_name while keeping newest first per DB
    discovered_sets.sort(key=lambda s: s.source_db_name.lower())

    if latest_only_per_db:
        best_by_db: Dict[str, DiscoveredBackupSet] = {}
        for bset in discovered_sets:
            key_db = bset.source_db_name.lower()
            if key_db not in best_by_db:
                best_by_db[key_db] = bset
            else:
                curr = best_by_db[key_db]
                # Prefer complete sets, newer timestamp_label, then raw .bak over .zip if same timestamp
                cand_rank = (
                    1 if bset.is_complete else 0,
                    bset.timestamp_label,
                    1 if bset.source_type == "bak" else 0,
                    bset.mtime_epoch,
                )
                curr_rank = (
                    1 if curr.is_complete else 0,
                    curr.timestamp_label,
                    1 if curr.source_type == "bak" else 0,
                    curr.mtime_epoch,
                )
                if cand_rank > curr_rank:
                    best_by_db[key_db] = bset
        discovered_sets = sorted(best_by_db.values(), key=lambda s: s.source_db_name.lower())

    return discovered_sets


def sort_backup_sets_largest_first(sets: List[DiscoveredBackupSet]) -> List[DiscoveredBackupSet]:
    """
    Order a batch of backup sets by `total_size_bytes` descending (Largest-Processing-Time-First)
    so multi-stripe large databases begin restoring immediately on parallel worker threads.
    """
    return sorted(
        sets,
        key=lambda s: (s.total_size_bytes, s.stream_count, s.source_db_name.lower()),
        reverse=True,
    )


def find_7zip_executable() -> Optional[str]:
    """Locate 7z.exe if installed for multi-threaded fastest extraction (`7z x -mmt=on`)."""
    which_7z = shutil.which("7z") or shutil.which("7za")
    if which_7z:
        return which_7z
    candidates = [
        r"C:\Program Files\7-Zip\7z.exe",
        r"C:\Program Files (x86)\7-Zip\7z.exe",
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c
    return None


def ensure_local_dir_permissions(folder_path: str) -> None:
    """
    Create `folder_path` locally and grant Modify/Read permissions to Builtin\\Users (*S-1-5-32-545)
    and Authenticated Users (*S-1-5-11) so the SQL Server service account can read staged `.bak`
    files or write relocated `.mdf`/`.ldf` files.
    """
    if not folder_path:
        return
    try:
        os.makedirs(folder_path, exist_ok=True)
    except OSError:
        return

    if os.name == "nt" and os.path.isdir(folder_path):
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        try:
            subprocess.run(
                [
                    "icacls",
                    folder_path,
                    "/grant",
                    "*S-1-5-32-545:(OI)(CI)M",
                    "/grant",
                    "*S-1-5-11:(OI)(CI)M",
                    "/Q",
                ],
                capture_output=True,
                creationflags=creationflags,
                timeout=10,
            )
        except Exception:
            pass


def fast_unzip_backup_set(
    zip_path: str,
    staging_root: str,
    db_name: str,
    timestamp_label: str = "",
    cancel_event: Optional[threading.Event] = None,
    status_callback: Optional[Callable[[str], None]] = None,
    prefer_7zip: bool = True,
) -> Tuple[str, List[str], float]:
    """
    Extract `.bak` stripe file(s) from `zip_path` into `<staging_root>\\<SafeDb>_<Timestamp>\\`
    prioritizing maximum extraction speed (`7z.exe x -y -mmt=on` or 16 MB buffered `zipfile`).
    Returns `(extract_dir, sorted_extracted_bak_paths, elapsed_seconds)`.
    """
    t0 = time.time()
    if not os.path.isfile(zip_path):
        raise FileNotFoundError(f"Zip archive not found: {zip_path}")

    safe_db = sanitize_filename(db_name)
    ts_part = sanitize_filename(timestamp_label) if timestamp_label else "latest"
    extract_dir = join_windows_or_native_path(staging_root, f"{safe_db}_{ts_part}")
    os.makedirs(extract_dir, exist_ok=True)
    ensure_local_dir_permissions(staging_root)
    ensure_local_dir_permissions(extract_dir)

    seven_zip = find_7zip_executable() if prefer_7zip else None
    if seven_zip:
        if status_callback:
            status_callback("Fast-Unzipping via 7-Zip (-mmt=on)...")
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        cmd = [seven_zip, "e", "-y", "-mmt=on", f"-o{extract_dir}", zip_path, "*.bak"]
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            creationflags=creationflags,
        )
        while proc.poll() is None:
            if cancel_event and cancel_event.is_set():
                proc.terminate()
                raise RuntimeError("Unzip operation cancelled by user.")
            time.sleep(0.15)
        if proc.returncode != 0:
            out = proc.stdout.read() if proc.stdout else ""
            raise RuntimeError(f"7-Zip extraction failed (exit {proc.returncode}): {out.strip()}")
    else:
        with zipfile.ZipFile(zip_path, "r") as zf:
            bak_members = [
                m for m in zf.infolist() if not m.is_dir() and m.filename.lower().endswith(".bak")
            ]
            if not bak_members:
                raise RuntimeError(f"No .bak files found inside archive {zip_path}")

            total_members = len(bak_members)
            for idx, member in enumerate(bak_members, start=1):
                if cancel_event and cancel_event.is_set():
                    raise RuntimeError("Unzip operation cancelled by user.")
                target_fname = os.path.basename(member.filename)
                dest_path = os.path.join(extract_dir, target_fname)
                if status_callback:
                    status_callback(f"Fast-Unzipping stripe {idx}/{total_members} (16 MB buffer)...")
                with zf.open(member, "r") as src_fp, open(dest_path, "wb") as dst_fp:
                    while True:
                        if cancel_event and cancel_event.is_set():
                            raise RuntimeError("Unzip operation cancelled by user.")
                        chunk = src_fp.read(FAST_UNZIP_CHUNK_BYTES)
                        if not chunk:
                            break
                        dst_fp.write(chunk)

    extracted_baks: List[Tuple[int, str]] = []
    for fname in os.listdir(extract_dir):
        if fname.lower().endswith(".bak"):
            full_p = os.path.join(extract_dir, fname)
            if os.path.isfile(full_p):
                _, _, s_idx, _ = parse_backup_filename(fname)
                extracted_baks.append((s_idx, full_p))

    extracted_baks.sort(key=lambda x: (x[0], x[1]))
    final_paths = [x[1] for x in extracted_baks]
    if not final_paths:
        raise RuntimeError(f"Unzip finished, but no .bak files were found in {extract_dir}")

    return extract_dir, final_paths, (time.time() - t0)


def cleanup_extracted_staging_files(extract_dir: str, file_paths: List[str]) -> bool:
    """Delete temporary extracted `.bak` files and remove `extract_dir` if empty."""
    cleaned = False
    for p in file_paths:
        if p and os.path.isfile(p):
            try:
                os.remove(p)
                cleaned = True
            except OSError:
                pass
    if extract_dir and os.path.isdir(extract_dir):
        try:
            if not os.listdir(extract_dir):
                os.rmdir(extract_dir)
        except OSError:
            pass
    return cleaned


def escape_tsql_identifier(identifier: str) -> str:
    """Escape a SQL Server bracketed identifier: [My]Db] -> [My]]Db]."""
    return f"[{identifier.replace(']', ']]')}]"


def escape_tsql_string(value: str) -> str:
    """Escape a SQL Server Unicode string literal: N'path'."""
    return "N'" + value.replace("'", "''") + "'"


def build_create_subdir_tsql(folder_path: str) -> str:
    """Build safe TRY/CATCH wrapped `xp_create_subdir` statement for Data/Log target folders."""
    return (
        f"BEGIN TRY EXEC master.dbo.xp_create_subdir {escape_tsql_string(folder_path)}; "
        f"END TRY BEGIN CATCH END CATCH;"
    )


def build_move_clauses(
    source_db_name: str,
    target_db_name: str,
    logical_files: List[LogicalFileInfo],
    target_data_folder: str,
    target_log_folder: str,
) -> List[Tuple[str, str]]:
    """
    Generate `(logical_name, target_physical_path)` pairs for `RESTORE DATABASE ... WITH MOVE`.
    - Data files (`file_type != 'L'`) go into `target_data_folder`.
    - Log files (`file_type == 'L'`) go into `target_log_folder`.
    - Automatically prefixes or adapts physical filenames when `target_db_name != source_db_name`
      so side-by-side restores never collide with original database files.
    """
    files = logical_files if logical_files else default_logical_files_for_db(source_db_name)
    safe_target_db = sanitize_filename(target_db_name)
    safe_source_db = sanitize_filename(source_db_name)
    is_renamed = safe_target_db.lower() != safe_source_db.lower()

    used_filenames: set[str] = set()
    moves: List[Tuple[str, str]] = []

    for idx, lf in enumerate(files, start=1):
        is_log = lf.file_type.upper() == "L"
        dest_folder = target_log_folder if is_log else target_data_folder
        orig_basename = os.path.basename(lf.physical_name.replace("\\", os.sep)) if lf.physical_name else ""
        orig_stem, orig_ext = os.path.splitext(orig_basename)
        if not orig_ext:
            orig_ext = ".ldf" if is_log else (".mdf" if idx == 1 else ".ndf")

        if not orig_stem:
            orig_stem = f"{safe_target_db}_log" if is_log else safe_target_db
        elif is_renamed:
            # Replace leading source_db_name with target_db_name if present, otherwise prefix
            pattern = re.compile(r"^" + re.escape(safe_source_db), re.IGNORECASE)
            if pattern.search(orig_stem):
                orig_stem = pattern.sub(safe_target_db, orig_stem, count=1)
            else:
                orig_stem = f"{safe_target_db}_{orig_stem}"

        candidate_name = f"{orig_stem}{orig_ext}"
        if candidate_name.lower() in used_filenames:
            candidate_name = f"{orig_stem}_{idx}{orig_ext}"
        used_filenames.add(candidate_name.lower())

        target_full_path = join_windows_or_native_path(dest_folder, candidate_name)
        moves.append((lf.logical_name, target_full_path))

    return moves


def build_restore_tsql(
    source_db_name: str,
    target_db_name: str,
    stripe_file_paths: List[str],
    logical_files: Optional[List[LogicalFileInfo]] = None,
    relocate_files: bool = True,
    target_data_folder: str = r"C:\SqlData",
    target_log_folder: str = r"C:\SqlLogs",
    with_replace: bool = True,
    kill_connections: bool = True,
    recovery_mode: str = "RECOVERY",
    high_speed_tuning: bool = True,
    max_transfer_size: int = DEFAULT_MAX_TRANSFER_SIZE,
    custom_buffer_count: int = 0,
    keep_cdc: bool = False,
    stats_interval: int = 5,
) -> str:
    """
    Build the full T-SQL batch to restore a single or multi-stream striped database as quickly as possible.
    """
    if not stripe_file_paths:
        raise ValueError("stripe_file_paths must contain at least one .bak file path.")

    escaped_target_db = escape_tsql_identifier(target_db_name)
    lit_target_db = escape_tsql_string(target_db_name)
    disk_clauses = [f"    DISK = {escape_tsql_string(p)}" for p in stripe_file_paths]
    from_clause = ",\n".join(disk_clauses)

    with_options: List[str] = []
    if relocate_files and target_data_folder and target_log_folder:
        moves = build_move_clauses(
            source_db_name=source_db_name,
            target_db_name=target_db_name,
            logical_files=logical_files or default_logical_files_for_db(source_db_name),
            target_data_folder=target_data_folder,
            target_log_folder=target_log_folder,
        )
        for logical_name, phys_path in moves:
            with_options.append(
                f"MOVE {escape_tsql_string(logical_name)} TO {escape_tsql_string(phys_path)}"
            )

    if with_replace:
        with_options.append("REPLACE")
    if keep_cdc:
        with_options.append("KEEP_CDC")

    rec_upper = recovery_mode.strip().upper()
    if rec_upper not in ("RECOVERY", "NORECOVERY"):
        rec_upper = "RECOVERY"
    with_options.append(rec_upper)

    if high_speed_tuning:
        xfer = max(65536, min(4194304, int(max_transfer_size)))
        buf_cnt = calculate_optimal_buffercount(len(stripe_file_paths), custom_buffer_count)
        with_options.append(f"MAXTRANSFERSIZE = {xfer}")
        with_options.append(f"BUFFERCOUNT = {buf_cnt}")

    stats_val = max(1, min(100, int(stats_interval)))
    with_options.append(f"STATS = {stats_val}")

    with_str = ",\n    ".join(with_options)
    restore_stmt = (
        f"RESTORE DATABASE {escaped_target_db}\n"
        f"FROM\n"
        f"{from_clause}\n"
        f"WITH\n"
        f"    {with_str};"
    )

    statements: List[str] = []
    if kill_connections:
        statements.append(
            f"IF DB_ID({lit_target_db}) IS NOT NULL AND DATABASEPROPERTYEX({lit_target_db}, 'Status') = 'ONLINE'\n"
            f"BEGIN\n"
            f"    ALTER DATABASE {escaped_target_db} SET SINGLE_USER WITH ROLLBACK IMMEDIATE;\n"
            f"END;"
        )

    statements.append(restore_stmt)

    if kill_connections and rec_upper == "RECOVERY":
        statements.append(
            f"IF DB_ID({lit_target_db}) IS NOT NULL AND DATABASEPROPERTYEX({lit_target_db}, 'Status') = 'ONLINE'\n"
            f"BEGIN\n"
            f"    ALTER DATABASE {escaped_target_db} SET MULTI_USER;\n"
            f"END;"
        )

    return "\n\n".join(statements)


# ============================================================================
# SQL Server Restore Execution Engine (.NET SqlClient PowerShell Bridge)
# ============================================================================


class SqlRestoreEngine:
    """
    Executes SQL Server instance discovery, `RESTORE HEADERONLY` / `RESTORE FILELISTONLY`
    metadata queries, and parallel high-speed `RESTORE DATABASE` commands.
    """

    def __init__(self, config: SqlConnectionConfig):
        self.config = config

    def fetch_instance_info(self) -> SqlInstanceInfo:
        """
        Connect to the target SQL Server instance and retrieve:
        - ServerName, Edition, ProductVersion
        - InstanceDefaultDataPath & InstanceDefaultLogPath
        - Instant File Initialization (IFI) status from `sys.dm_server_services`
        - List of existing databases on the server
        """
        ps_script = r"""
$ErrorActionPreference = 'Stop'
$connStr = $env:SQL_RS_CONN_STR

$conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
try {
    $conn.Open()

    $cmd = $conn.CreateCommand()
    $cmd.CommandTimeout = 30
    $cmd.CommandText = @"
DECLARE @ifi NVARCHAR(32) = N'Unknown';
BEGIN TRY
    SELECT TOP (1) @ifi = CASE WHEN instant_file_initialization_enabled = 'Y' THEN N'Enabled' ELSE N'Disabled' END
    FROM sys.dm_server_services
    WHERE servicename LIKE N'SQL Server (%';
END TRY
BEGIN CATCH
    SET @ifi = N'Unknown';
END CATCH;

SELECT
    CAST(SERVERPROPERTY('ServerName') AS NVARCHAR(256)) AS server_name,
    CAST(SERVERPROPERTY('Edition') AS NVARCHAR(256)) AS edition,
    CAST(SERVERPROPERTY('ProductVersion') AS NVARCHAR(128)) AS product_version,
    CAST(ISNULL(SERVERPROPERTY('InstanceDefaultDataPath'), N'C:\SqlData') AS NVARCHAR(512)) AS default_data_path,
    CAST(ISNULL(SERVERPROPERTY('InstanceDefaultLogPath'), N'C:\SqlLogs') AS NVARCHAR(512)) AS default_log_path,
    @ifi AS ifi_status;
"@
    $r = $cmd.ExecuteReader()
    $info = @{}
    if ($r.Read()) {
        $info['server_name'] = [string]$r['server_name']
        $info['edition'] = [string]$r['edition']
        $info['product_version'] = [string]$r['product_version']
        $info['default_data_path'] = ([string]$r['default_data_path']).TrimEnd('\', '/')
        $info['default_log_path'] = ([string]$r['default_log_path']).TrimEnd('\', '/')
        $info['ifi_status'] = [string]$r['ifi_status']
    }
    $r.Close()

    $dbCmd = $conn.CreateCommand()
    $dbCmd.CommandTimeout = 30
    $dbCmd.CommandText = "SELECT name FROM sys.databases WHERE database_id > 4 ORDER BY name ASC;"
    $dr = $dbCmd.ExecuteReader()
    $dbs = @()
    while ($dr.Read()) {
        $dbs += [string]$dr['name']
    }
    $dr.Close()

    $info['existing_databases'] = $dbs
    Write-Output ($info | ConvertTo-Json -Depth 4 -Compress)
}
finally {
    if ($conn.State -ne 'Closed') { $conn.Close() }
    $conn.Dispose()
}
"""
        env = os.environ.copy()
        env["SQL_RS_CONN_STR"] = self.config.to_dotnet_connection_string("master")
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        proc = subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                ps_script,
            ],
            env=env,
            capture_output=True,
            text=True,
            creationflags=creationflags,
            timeout=self.config.connect_timeout + 30,
        )
        if proc.returncode != 0:
            err = (proc.stderr or proc.stdout or "Unknown SQL connection error").strip()
            raise RuntimeError(err)

        raw_out = proc.stdout.strip()
        json_line = ""
        for line in reversed(raw_out.splitlines()):
            line = line.strip()
            if line.startswith("{") and line.endswith("}"):
                json_line = line
                break
        if not json_line:
            raise RuntimeError(f"Unexpected response from SQL metadata query: {raw_out}")

        payload = json.loads(json_line)
        existing = payload.get("existing_databases", [])
        if isinstance(existing, str):
            existing = [existing]

        return SqlInstanceInfo(
            server_name=str(payload.get("server_name") or self.config.server),
            edition=str(payload.get("edition") or "SQL Server"),
            product_version=str(payload.get("product_version") or "?"),
            default_data_path=str(payload.get("default_data_path") or r"C:\SqlData"),
            default_log_path=str(payload.get("default_log_path") or r"C:\SqlLogs"),
            ifi_enabled=str(payload.get("ifi_status") or "Unknown"),
            existing_databases=[str(x) for x in existing],
        )

    def inspect_backup_files_metadata(
        self, first_stripe_path: str
    ) -> Tuple[Optional[str], int, List[LogicalFileInfo]]:
        """
        Execute `RESTORE HEADERONLY` and `RESTORE FILELISTONLY` on `first_stripe_path`
        to retrieve `(database_name, media_family_count, logical_files)`.
        """
        ps_script = r"""
$ErrorActionPreference = 'Stop'
$connStr = $env:SQL_RS_CONN_STR
$bakPath = $env:SQL_RS_BAK_PATH
$escapedPath = "N'" + $bakPath.Replace("'", "''") + "'"

$conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
try {
    $conn.Open()

    $hdrCmd = $conn.CreateCommand()
    $hdrCmd.CommandTimeout = 60
    $hdrCmd.CommandText = "RESTORE HEADERONLY FROM DISK = $escapedPath;"
    $hr = $hdrCmd.ExecuteReader()
    $dbName = ""
    $familyCount = 1
    if ($hr.Read()) {
        $dbName = [string]$hr['DatabaseName']
        try {
            if ($null -ne $hr['FamilyCount'] -and $hr['FamilyCount'] -isnot [System.DBNull]) {
                $familyCount = [int]$hr['FamilyCount']
            }
        } catch {}
    }
    $hr.Close()

    $flCmd = $conn.CreateCommand()
    $flCmd.CommandTimeout = 60
    $flCmd.CommandText = "RESTORE FILELISTONLY FROM DISK = $escapedPath;"
    $fr = $flCmd.ExecuteReader()
    $files = @()
    $idx = 1
    while ($fr.Read()) {
        $fId = $idx
        try {
            if ($null -ne $fr['FileId'] -and $fr['FileId'] -isnot [System.DBNull]) {
                $fId = [int]$fr['FileId']
            }
        } catch {}
        $fSize = [int64]0
        try {
            if ($null -ne $fr['Size'] -and $fr['Size'] -isnot [System.DBNull]) {
                $fSize = [int64]$fr['Size']
            }
        } catch {}
        $files += @{
            logical_name = [string]$fr['LogicalName']
            physical_name = [string]$fr['PhysicalName']
            file_type = [string]$fr['Type']
            size_bytes = $fSize
            file_id = $fId
        }
        $idx++
    }
    $fr.Close()

    $res = @{
        database_name = $dbName
        family_count = $familyCount
        logical_files = $files
    }
    Write-Output ($res | ConvertTo-Json -Depth 5 -Compress)
}
finally {
    if ($conn.State -ne 'Closed') { $conn.Close() }
    $conn.Dispose()
}
"""
        env = os.environ.copy()
        env["SQL_RS_CONN_STR"] = self.config.to_dotnet_connection_string("master")
        env["SQL_RS_BAK_PATH"] = first_stripe_path
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        proc = subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                ps_script,
            ],
            env=env,
            capture_output=True,
            text=True,
            creationflags=creationflags,
            timeout=90,
        )
        if proc.returncode != 0:
            err = (proc.stderr or proc.stdout or "RESTORE FILELISTONLY failed").strip()
            raise RuntimeError(err)

        raw_out = proc.stdout.strip()
        json_line = ""
        for line in reversed(raw_out.splitlines()):
            line = line.strip()
            if line.startswith("{") and line.endswith("}"):
                json_line = line
                break
        if not json_line:
            raise RuntimeError(f"Unexpected response from RESTORE FILELISTONLY: {raw_out}")

        payload = json.loads(json_line)
        db_name = str(payload.get("database_name") or "").strip() or None
        family_count = max(1, int(payload.get("family_count") or 1))
        raw_files = payload.get("logical_files") or []
        if isinstance(raw_files, dict):
            raw_files = [raw_files]

        logical_files: List[LogicalFileInfo] = []
        for item in raw_files:
            logical_files.append(
                LogicalFileInfo(
                    logical_name=str(item.get("logical_name", "")),
                    physical_name=str(item.get("physical_name", "")),
                    file_type=str(item.get("file_type", "D")),
                    size_bytes=int(item.get("size_bytes", 0)),
                    file_id=int(item.get("file_id", 1)),
                )
            )
        return db_name, family_count, logical_files

    def run_single_database_restore(
        self,
        backup_set: DiscoveredBackupSet,
        options: RestoreOptions,
        cancel_event: threading.Event,
        progress_callback: Optional[Callable[[str, int, str], None]] = None,
    ) -> RestoreJobResult:
        """
        Restore a single database (from 1 or N striped `.bak` files, or a `.zip` archive):
        1. If `backup_set.source_type == "zip"`, fast-unzip the `.bak` stripes into the staging folder.
        2. Query `RESTORE FILELISTONLY` on stripe 1 if logical file names haven't been verified yet.
        3. Ensure Target Data & Log directories exist on the SQL Server host.
        4. Execute high-speed `RESTORE DATABASE` with tuned `MAXTRANSFERSIZE` and `BUFFERCOUNT`
           (with automatic fallback if SQL Server reports virtual memory/buffer pressure).
        5. Clean up temporary extracted `.bak` files from the staging directory if enabled.
        """
        start_time = time.time()
        target_db_name = resolve_target_db_name(
            source_db_name=backup_set.source_db_name,
            prefix=options.target_db_prefix,
            suffix=options.target_db_suffix,
            custom_target_name=backup_set.custom_target_name,
        )

        if cancel_event.is_set():
            return RestoreJobResult(
                set_id=backup_set.set_id,
                source_db_name=backup_set.source_db_name,
                target_db_name=target_db_name,
                success=False,
                stream_count=backup_set.stream_count,
                source_type=backup_set.source_type,
                error_message="Cancelled before start.",
            )

        extract_dir = ""
        active_stripe_paths = list(backup_set.stripe_paths)
        unzip_elapsed = 0.0

        try:
            # Step 1: Fast-Unzip if restoring from a .zip archive
            if backup_set.source_type == "zip":
                staging_root = resolve_staging_folder(options.source_folder, options.staging_folder)
                if progress_callback:
                    progress_callback(backup_set.set_id, 2, "Fast-Unzipping archive to staging...")
                extract_dir, active_stripe_paths, unzip_elapsed = fast_unzip_backup_set(
                    zip_path=backup_set.zip_path,
                    staging_root=staging_root,
                    db_name=backup_set.source_db_name,
                    timestamp_label=backup_set.timestamp_label,
                    cancel_event=cancel_event,
                    status_callback=lambda msg: (
                        progress_callback(backup_set.set_id, 8, msg) if progress_callback else None
                    ),
                )
            else:
                # Ensure SQL Server service account can read the source folder if local
                if active_stripe_paths:
                    ensure_local_dir_permissions(os.path.dirname(active_stripe_paths[0]))

            stream_count = len(active_stripe_paths)

            # Step 2: Inspect logical files via RESTORE FILELISTONLY if needed for WITH MOVE
            logical_files = list(backup_set.logical_files)
            if options.relocate_files and not backup_set.header_verified and active_stripe_paths:
                try:
                    if progress_callback:
                        progress_callback(
                            backup_set.set_id,
                            12,
                            "Reading logical file list (RESTORE FILELISTONLY)...",
                        )
                    _, _, inspected_files = self.inspect_backup_files_metadata(active_stripe_paths[0])
                    if inspected_files:
                        logical_files = inspected_files
                        backup_set.logical_files = inspected_files
                        backup_set.header_verified = True
                except Exception:
                    # Fall back to default_logical_files_for_db if FILELISTONLY fails prior to restore
                    pass

            # Step 3: Build T-SQL (High-Speed Tuned + Standard Buffer Fallback)
            ensure_local_dir_permissions(options.target_data_folder)
            ensure_local_dir_permissions(options.target_log_folder)

            create_dirs_sql = ""
            if options.relocate_files:
                create_dirs_sql = (
                    build_create_subdir_tsql(options.target_data_folder)
                    + "\n"
                    + build_create_subdir_tsql(options.target_log_folder)
                )

            restore_sql_tuned = build_restore_tsql(
                source_db_name=backup_set.source_db_name,
                target_db_name=target_db_name,
                stripe_file_paths=active_stripe_paths,
                logical_files=logical_files,
                relocate_files=options.relocate_files,
                target_data_folder=options.target_data_folder,
                target_log_folder=options.target_log_folder,
                with_replace=options.with_replace,
                kill_connections=options.kill_connections,
                recovery_mode=options.recovery_mode,
                high_speed_tuning=options.high_speed_tuning,
                max_transfer_size=options.max_transfer_size,
                custom_buffer_count=options.custom_buffer_count,
                keep_cdc=options.keep_cdc,
                stats_interval=options.stats_interval,
            )
            restore_sql_fallback = build_restore_tsql(
                source_db_name=backup_set.source_db_name,
                target_db_name=target_db_name,
                stripe_file_paths=active_stripe_paths,
                logical_files=logical_files,
                relocate_files=options.relocate_files,
                target_data_folder=options.target_data_folder,
                target_log_folder=options.target_log_folder,
                with_replace=options.with_replace,
                kill_connections=options.kill_connections,
                recovery_mode=options.recovery_mode,
                high_speed_tuning=False,
                keep_cdc=options.keep_cdc,
                stats_interval=options.stats_interval,
            )

            ps_runner = r"""
$ErrorActionPreference = 'Stop'
$connStr = $env:SQL_RS_CONN_STR
$createDirsSql = $env:SQL_RS_CREATE_DIRS_SQL
$restoreSql = $env:SQL_RS_RESTORE_SQL
$restoreFallbackSql = $env:SQL_RS_RESTORE_FALLBACK_SQL
$allowBufferFallback = ($env:SQL_RS_ALLOW_BUFFER_FALLBACK -eq '1')

$conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
$conn.FireInfoMessageEventOnUserErrors = $false

$msgHandler = [System.Data.SqlClient.SqlInfoMessageEventHandler]{
    param($sender, $eventArgs)
    foreach ($err in $eventArgs.Errors) {
        if ($null -ne $err -and -not [string]::IsNullOrWhiteSpace($err.Message)) {
            [Console]::Out.WriteLine("SQL_INFO: " + $err.Message.Trim())
            [Console]::Out.Flush()
        }
    }
}
$conn.add_InfoMessage($msgHandler)

try {
    $conn.Open()

    if (-not [string]::IsNullOrWhiteSpace($createDirsSql)) {
        try {
            $dirCmd = $conn.CreateCommand()
            $dirCmd.CommandTimeout = 60
            $dirCmd.CommandText = $createDirsSql
            [void]$dirCmd.ExecuteNonQuery()
        }
        catch {}
    }

    $usedHighSpeed = $allowBufferFallback
    try {
        $resCmd = $conn.CreateCommand()
        $resCmd.CommandTimeout = 0
        $resCmd.CommandText = $restoreSql
        [void]$resCmd.ExecuteNonQuery()
    }
    catch {
        $exMsg = $_.Exception.GetBaseException().Message
        if ($allowBufferFallback -and ($exMsg -match 'MAXTRANSFERSIZE' -or $exMsg -match 'BUFFERCOUNT' -or $exMsg -match 'memory' -or $exMsg -match '3204' -or $exMsg -match '701' -or $exMsg -match '802')) {
            [Console]::Out.WriteLine("SQL_WARN: Buffer memory pressure detected; retrying with default SQL Server buffers...")
            [Console]::Out.Flush()
            $usedHighSpeed = $false
            $resCmd2 = $conn.CreateCommand()
            $resCmd2.CommandTimeout = 0
            $resCmd2.CommandText = $restoreFallbackSql
            [void]$resCmd2.ExecuteNonQuery()
        }
        else {
            throw
        }
    }

    [Console]::Out.WriteLine("SQL_DONE: USED_HIGH_SPEED=" + $usedHighSpeed)
    [Console]::Out.Flush()
}
catch {
    $cleanErr = $_.Exception.GetBaseException().Message -replace "(\r|\n)+", " "
    [Console]::Out.WriteLine("SQL_ERROR: " + $cleanErr)
    [Console]::Out.Flush()
    exit 1
}
finally {
    if ($conn.State -ne 'Closed') { $conn.Close() }
    $conn.Dispose()
}
"""
            env = os.environ.copy()
            env["SQL_RS_CONN_STR"] = self.config.to_dotnet_connection_string("master")
            env["SQL_RS_CREATE_DIRS_SQL"] = create_dirs_sql
            env["SQL_RS_RESTORE_SQL"] = restore_sql_tuned
            env["SQL_RS_RESTORE_FALLBACK_SQL"] = restore_sql_fallback
            env["SQL_RS_ALLOW_BUFFER_FALLBACK"] = "1" if options.high_speed_tuning else "0"

            base_pct = 15 if backup_set.source_type == "zip" else 0
            if progress_callback:
                progress_callback(
                    backup_set.set_id,
                    base_pct,
                    f"Restoring [{target_db_name}] ({stream_count} {'stripe' if stream_count == 1 else 'stripes'})...",
                )

            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            proc = subprocess.Popen(
                [
                    "powershell.exe",
                    "-NoProfile",
                    "-NonInteractive",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    ps_runner,
                ],
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                creationflags=creationflags,
            )

            used_high_speed = options.high_speed_tuning
            output_lines: List[str] = []
            sql_error_msg = ""
            pct_regex = re.compile(r"(\d+)\s+percent\s+processed", re.IGNORECASE)

            try:
                assert proc.stdout is not None
                for raw_line in proc.stdout:
                    if cancel_event.is_set():
                        proc.terminate()
                        return RestoreJobResult(
                            set_id=backup_set.set_id,
                            source_db_name=backup_set.source_db_name,
                            target_db_name=target_db_name,
                            success=False,
                            stream_count=stream_count,
                            source_type=backup_set.source_type,
                            elapsed_seconds=time.time() - start_time,
                            unzip_elapsed_seconds=unzip_elapsed,
                            error_message="Cancelled by user.",
                        )

                    line = raw_line.strip()
                    if not line:
                        continue
                    output_lines.append(line)

                    if line.startswith("SQL_INFO:"):
                        info_msg = line[len("SQL_INFO:") :].strip()
                        m = pct_regex.search(info_msg)
                        if m:
                            pct = int(m.group(1))
                            scaled_pct = (
                                int(15 + pct * 0.83)
                                if backup_set.source_type == "zip"
                                else pct
                            )
                            if progress_callback:
                                progress_callback(
                                    backup_set.set_id,
                                    scaled_pct,
                                    f"SQL Restore ({pct}% | {stream_count} {'stripe' if stream_count == 1 else 'stripes'})",
                                )
                        else:
                            if progress_callback:
                                progress_callback(backup_set.set_id, -1, info_msg)
                    elif line.startswith("SQL_WARN:"):
                        warn_msg = line[len("SQL_WARN:") :].strip()
                        if progress_callback:
                            progress_callback(backup_set.set_id, -1, warn_msg)
                    elif line.startswith("SQL_ERROR:"):
                        sql_error_msg = line[len("SQL_ERROR:") :].strip()
                    elif line.startswith("SQL_DONE:"):
                        if "USED_HIGH_SPEED=False" in line:
                            used_high_speed = False

                proc.wait()
            finally:
                if proc.poll() is None:
                    proc.kill()

            if proc.returncode != 0:
                elapsed = time.time() - start_time
                err_detail = (
                    sql_error_msg
                    if sql_error_msg
                    else ("\n".join(output_lines[-6:]) if output_lines else f"Exit code {proc.returncode}")
                )
                return RestoreJobResult(
                    set_id=backup_set.set_id,
                    source_db_name=backup_set.source_db_name,
                    target_db_name=target_db_name,
                    success=False,
                    stream_count=stream_count,
                    source_type=backup_set.source_type,
                    elapsed_seconds=elapsed,
                    unzip_elapsed_seconds=unzip_elapsed,
                    used_high_speed_buffers=used_high_speed,
                    error_message=err_detail,
                )

            cleaned_staging = False
            if backup_set.source_type == "zip" and options.cleanup_staging_bak and extract_dir:
                cleaned_staging = cleanup_extracted_staging_files(extract_dir, active_stripe_paths)

            elapsed = time.time() - start_time
            unzip_note = f", unzipped in {unzip_elapsed:.1f}s" if unzip_elapsed > 0 else ""
            done_msg = (
                f"Restored in {elapsed:.1f}s ({stream_count} "
                f"{'stripe' if stream_count == 1 else 'stripes'}{unzip_note})"
            )
            if progress_callback:
                progress_callback(backup_set.set_id, 100, done_msg)

            return RestoreJobResult(
                set_id=backup_set.set_id,
                source_db_name=backup_set.source_db_name,
                target_db_name=target_db_name,
                success=True,
                stream_count=stream_count,
                source_type=backup_set.source_type,
                elapsed_seconds=elapsed,
                unzip_elapsed_seconds=unzip_elapsed,
                used_high_speed_buffers=used_high_speed,
                cleaned_staging=cleaned_staging,
            )

        except Exception as exc:
            elapsed = time.time() - start_time
            if backup_set.source_type == "zip" and options.cleanup_staging_bak and extract_dir:
                cleanup_extracted_staging_files(extract_dir, active_stripe_paths)
            return RestoreJobResult(
                set_id=backup_set.set_id,
                source_db_name=backup_set.source_db_name,
                target_db_name=target_db_name,
                success=False,
                stream_count=backup_set.stream_count,
                source_type=backup_set.source_type,
                elapsed_seconds=elapsed,
                unzip_elapsed_seconds=unzip_elapsed,
                error_message=str(exc),
            )


# ============================================================================
# Tkinter Graphical User Interface
# ============================================================================


class SqlRestoreApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1240x880")
        self.root.minsize(1040, 740)

        self.config_file_path = Path(__file__).resolve().parent / CONFIG_FILENAME

        # 1. SQL Connection state variables
        self.var_server = tk.StringVar(value="localhost")
        self.var_auth_mode = tk.StringVar(value="windows")
        self.var_username = tk.StringVar(value="sa")
        self.var_password = tk.StringVar(value="")
        self.var_show_password = tk.BooleanVar(value=False)
        self.var_trust_cert = tk.BooleanVar(value=True)
        self.var_server_status = tk.StringVar(
            value="Not connected — enter target SQL Server instance and click 'Connect & Detect Paths'"
        )

        # 2. Source Backup & Target Relocation state variables
        default_source = r"C:\SqlBackups"
        self.var_source_folder = tk.StringVar(value=default_source)
        self.var_latest_only = tk.BooleanVar(value=True)
        self.var_cleanup_staging = tk.BooleanVar(value=True)

        self.var_relocate_files = tk.BooleanVar(value=True)
        self.var_target_data_folder = tk.StringVar(value=r"C:\SqlData")
        self.var_target_log_folder = tk.StringVar(value=r"C:\SqlLogs")
        self.var_db_prefix = tk.StringVar(value="")
        self.var_db_suffix = tk.StringVar(value="")

        # 3. High-Speed Restore & Overwrite Options
        self.var_max_parallel_dbs = tk.IntVar(value=DEFAULT_MAX_PARALLEL_DBS)
        self.var_with_replace = tk.BooleanVar(value=True)
        self.var_kill_connections = tk.BooleanVar(value=True)
        self.var_recovery_mode = tk.StringVar(value="RECOVERY")
        self.var_high_speed_tuning = tk.BooleanVar(value=True)
        self.var_custom_buffercount = tk.IntVar(value=0)  # 0 = Auto (streams * 8)
        self.var_keep_cdc = tk.BooleanVar(value=False)

        # 4. Selection & Progress variables
        self.var_filter_text = tk.StringVar(value="")
        self.var_selection_summary = tk.StringVar(
            value="Selected: 0 backup sets (0.00 GB | 0 total stripes)"
        )
        self.var_overall_progress = tk.DoubleVar(value=0.0)
        self.var_progress_label = tk.StringVar(value="Idle")

        # Discovered backup inventory & state maps
        self.all_backup_sets: List[DiscoveredBackupSet] = []
        self.selected_set_ids: set[str] = set()
        self.set_status_map: Dict[str, str] = {}
        self.existing_server_dbs: set[str] = set()

        # Concurrency & UI event queue
        self.ui_queue: "queue.Queue[Tuple[str, object]]" = queue.Queue()
        self.cancel_event = threading.Event()
        self.is_running_restores = False

        self._configure_styles()
        self._build_ui()
        self._load_saved_config(silent=True)
        self._on_auth_mode_change()
        self._on_relocate_toggle_change()
        self.scan_source_backups(silent_if_missing=True)
        self.update_tsql_preview()
        self._poll_ui_queue()

    def _configure_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            if "vista" in style.theme_names():
                style.theme_use("vista")
            elif "clam" in style.theme_names():
                style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("Header.TLabel", font=("Segoe UI", 13, "bold"))
        style.configure("SubHeader.TLabel", font=("Segoe UI", 9))
        style.configure("Bold.TLabel", font=("Segoe UI", 9, "bold"))
        style.configure("Treeview", rowheight=25, font=("Segoe UI", 9))
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

    def _build_ui(self) -> None:
        main_frame = ttk.Frame(self.root, padding="10 8 10 8")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. Top Banner
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(
            header_frame,
            text="SQL Server Parallel Striped Database Restore Utility",
            style="Header.TLabel",
        ).pack(side=tk.LEFT)
        ttk.Label(
            header_frame,
            text="Largest-First Queue  •  Multi-Stripe .bak & Fast-Unzip  •  4 MB Transfer + Auto BUFFERCOUNT",
            style="SubHeader.TLabel",
        ).pack(side=tk.RIGHT)

        # 2. Target SQL Server Connection Panel
        conn_box = ttk.LabelFrame(
            main_frame, text=" 1. Target SQL Server Connection ", padding="8 6 8 6"
        )
        conn_box.pack(fill=tk.X, pady=(0, 6))

        row0 = ttk.Frame(conn_box)
        row0.pack(fill=tk.X, pady=2)

        ttk.Label(row0, text="Target Server / Instance:").pack(side=tk.LEFT, padx=(0, 4))
        self.ent_server = ttk.Entry(row0, textvariable=self.var_server, width=26)
        self.ent_server.pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(row0, text="Authentication:").pack(side=tk.LEFT, padx=(0, 4))
        self.cmb_auth = ttk.Combobox(
            row0,
            textvariable=self.var_auth_mode,
            values=["windows", "sql"],
            state="readonly",
            width=11,
        )
        self.cmb_auth.pack(side=tk.LEFT, padx=(0, 12))
        self.cmb_auth.bind("<<ComboboxSelected>>", lambda _e: self._on_auth_mode_change())

        ttk.Label(row0, text="SQL Login:").pack(side=tk.LEFT, padx=(0, 4))
        self.ent_user = ttk.Entry(row0, textvariable=self.var_username, width=15)
        self.ent_user.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(row0, text="Password:").pack(side=tk.LEFT, padx=(0, 4))
        self.ent_pass = ttk.Entry(row0, textvariable=self.var_password, width=15, show="*")
        self.ent_pass.pack(side=tk.LEFT, padx=(0, 4))

        ttk.Checkbutton(
            row0,
            text="Show",
            variable=self.var_show_password,
            command=self._toggle_password_visibility,
        ).pack(side=tk.LEFT, padx=(0, 10))

        self.btn_connect = ttk.Button(
            row0,
            text="Connect & Detect Paths",
            command=self.connect_and_detect_instance,
        )
        self.btn_connect.pack(side=tk.LEFT, padx=(2, 6))

        ttk.Button(row0, text="Save Config", command=self.save_config).pack(
            side=tk.RIGHT, padx=(4, 0)
        )
        ttk.Button(
            row0, text="Load Config", command=lambda: self._load_saved_config(silent=False)
        ).pack(side=tk.RIGHT, padx=(4, 0))

        row1 = ttk.Frame(conn_box)
        row1.pack(fill=tk.X, pady=(4, 0))
        ttk.Checkbutton(
            row1,
            text="Trust Server Certificate",
            variable=self.var_trust_cert,
        ).pack(side=tk.LEFT, padx=(0, 16))

        ttk.Label(row1, textvariable=self.var_server_status, foreground="#1e56a0").pack(
            side=tk.LEFT, fill=tk.X, expand=True
        )

        # 3. Backup Source, File Relocation (WITH MOVE) & High-Speed Restore Options
        settings_box = ttk.LabelFrame(
            main_frame,
            text=" 2. Source Backup Folder (.bak / .zip), File Relocation (WITH MOVE) & High-Speed Tuning ",
            padding="8 6 8 6",
        )
        settings_box.pack(fill=tk.X, pady=(0, 6))

        # Row A: Source Folder + Scan & Header Inspection
        srow0 = ttk.Frame(settings_box)
        srow0.pack(fill=tk.X, pady=2)
        ttk.Label(srow0, text="Source Backup Folder (.bak/.zip):", width=26, anchor=tk.W).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        self.ent_source = ttk.Entry(srow0, textvariable=self.var_source_folder)
        self.ent_source.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

        ttk.Button(srow0, text="Browse Source...", command=self.browse_source_folder).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        self.btn_scan = ttk.Button(
            srow0,
            text="Scan Backups",
            command=lambda: self.scan_source_backups(silent_if_missing=False),
        )
        self.btn_scan.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_inspect_headers = ttk.Button(
            srow0,
            text="Verify SQL Headers",
            command=self.inspect_selected_sql_headers,
        )
        self.btn_inspect_headers.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Checkbutton(
            srow0,
            text="Latest backup per DB only",
            variable=self.var_latest_only,
            command=lambda: self.scan_source_backups(silent_if_missing=True),
        ).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Checkbutton(
            srow0,
            text="Delete staged .bak after .zip restore",
            variable=self.var_cleanup_staging,
            command=self._on_settings_changed,
        ).pack(side=tk.LEFT)

        # Row B: Target Data & Log Folders (WITH MOVE) + DB Name Prefix/Suffix
        srow_move = ttk.Frame(settings_box)
        srow_move.pack(fill=tk.X, pady=3)
        ttk.Checkbutton(
            srow_move,
            text="Relocate Files (WITH MOVE) — Data Folder:",
            variable=self.var_relocate_files,
            command=self._on_relocate_toggle_change,
        ).pack(side=tk.LEFT, padx=(0, 4))

        self.ent_data_folder = ttk.Entry(
            srow_move, textvariable=self.var_target_data_folder, width=26
        )
        self.ent_data_folder.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        self.ent_data_folder.bind("<KeyRelease>", lambda _e: self._on_settings_changed())

        self.btn_browse_data = ttk.Button(
            srow_move, text="Browse...", width=9, command=self.browse_data_folder
        )
        self.btn_browse_data.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(srow_move, text="Log Folder (.ldf):").pack(side=tk.LEFT, padx=(0, 4))
        self.ent_log_folder = ttk.Entry(
            srow_move, textvariable=self.var_target_log_folder, width=26
        )
        self.ent_log_folder.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        self.ent_log_folder.bind("<KeyRelease>", lambda _e: self._on_settings_changed())

        self.btn_browse_log = ttk.Button(
            srow_move, text="Browse...", width=9, command=self.browse_log_folder
        )
        self.btn_browse_log.pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(srow_move, text="DB Prefix:").pack(side=tk.LEFT, padx=(0, 2))
        ent_prefix = ttk.Entry(srow_move, textvariable=self.var_db_prefix, width=9)
        ent_prefix.pack(side=tk.LEFT, padx=(0, 8))
        ent_prefix.bind("<KeyRelease>", lambda _e: self._on_settings_changed())

        ttk.Label(srow_move, text="Suffix:").pack(side=tk.LEFT, padx=(0, 2))
        ent_suffix = ttk.Entry(srow_move, textvariable=self.var_db_suffix, width=9)
        ent_suffix.pack(side=tk.LEFT)
        ent_suffix.bind("<KeyRelease>", lambda _e: self._on_settings_changed())

        # Row C: Parallelism, Overwrite, Recovery & High-Speed Tuning
        srow_perf = ttk.Frame(settings_box)
        srow_perf.pack(fill=tk.X, pady=(5, 2))

        ttk.Label(srow_perf, text="Max Parallel DB Restores:").pack(side=tk.LEFT, padx=(0, 4))
        spn_parallel = ttk.Spinbox(
            srow_perf,
            from_=1,
            to=16,
            textvariable=self.var_max_parallel_dbs,
            width=4,
            command=self._on_settings_changed,
        )
        spn_parallel.pack(side=tk.LEFT, padx=(0, 14))
        spn_parallel.bind("<KeyRelease>", lambda _e: self._on_settings_changed())

        ttk.Checkbutton(
            srow_perf,
            text="WITH REPLACE (Overwrite)",
            variable=self.var_with_replace,
            command=self._on_settings_changed,
        ).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Checkbutton(
            srow_perf,
            text="Kill Active Connections (SINGLE_USER)",
            variable=self.var_kill_connections,
            command=self._on_settings_changed,
        ).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(srow_perf, text="State:").pack(side=tk.LEFT, padx=(0, 4))
        cmb_rec = ttk.Combobox(
            srow_perf,
            textvariable=self.var_recovery_mode,
            values=["RECOVERY", "NORECOVERY"],
            state="readonly",
            width=12,
        )
        cmb_rec.pack(side=tk.LEFT, padx=(0, 12))
        cmb_rec.bind("<<ComboboxSelected>>", lambda _e: self._on_settings_changed())

        ttk.Checkbutton(
            srow_perf,
            text="High-Speed I/O (4MB Transfer + BUFFERCOUNT):",
            variable=self.var_high_speed_tuning,
            command=self._on_settings_changed,
        ).pack(side=tk.LEFT, padx=(0, 4))

        ttk.Label(srow_perf, text="BufferCount (0=Auto):").pack(side=tk.LEFT, padx=(0, 4))
        spn_buf = ttk.Spinbox(
            srow_perf,
            from_=0,
            to=512,
            increment=8,
            textvariable=self.var_custom_buffercount,
            width=5,
            command=self._on_settings_changed,
        )
        spn_buf.pack(side=tk.LEFT, padx=(0, 12))
        spn_buf.bind("<KeyRelease>", lambda _e: self._on_settings_changed())

        ttk.Checkbutton(
            srow_perf,
            text="KEEP_CDC",
            variable=self.var_keep_cdc,
            command=self._on_settings_changed,
        ).pack(side=tk.LEFT)

        # 4. PanedWindow: Discovered Backup Sets Table (Top) + T-SQL Preview & Execution Log (Bottom)
        paned = ttk.PanedWindow(main_frame, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True)

        # 4A. Backup Set Selection Frame
        db_box = ttk.LabelFrame(
            paned,
            text=" 3. Select Striped Backup Sets to Restore (Click/Space to toggle • Right-click or button to rename Target DB) ",
            padding="8 6 8 6",
        )
        paned.add(db_box, weight=3)

        toolbar = ttk.Frame(db_box)
        toolbar.pack(fill=tk.X, pady=(0, 6))

        ttk.Button(toolbar, text="Select All", command=self.select_all_sets).pack(
            side=tk.LEFT, padx=(0, 4)
        )
        ttk.Button(toolbar, text="Deselect All", command=self.deselect_all_sets).pack(
            side=tk.LEFT, padx=(0, 4)
        )
        ttk.Button(
            toolbar, text="Select Multi-Stripe Only", command=self.select_striped_sets
        ).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(toolbar, text="Invert Selection", command=self.invert_selection).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Button(
            toolbar,
            text="Rename Highlighted Target DB...",
            command=self.prompt_rename_selected_db,
        ).pack(side=tk.LEFT, padx=(0, 14))

        ttk.Label(toolbar, text="Filter DBs:").pack(side=tk.LEFT, padx=(0, 4))
        ent_filter = ttk.Entry(toolbar, textvariable=self.var_filter_text, width=20)
        ent_filter.pack(side=tk.LEFT, padx=(0, 6))
        self.var_filter_text.trace_add("write", lambda *_: self.refresh_backup_tree())

        ttk.Button(
            toolbar,
            text="Clear Filter",
            command=lambda: self.var_filter_text.set(""),
        ).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(
            toolbar,
            textvariable=self.var_selection_summary,
            style="Bold.TLabel",
        ).pack(side=tk.RIGHT)

        # Treeview
        tree_frame = ttk.Frame(db_box)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        columns = (
            "selected",
            "source_db",
            "target_db",
            "size_gb",
            "stripes",
            "format_ts",
            "status",
            "source_location",
        )
        self.tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show="headings",
            selectmode="browse",
        )
        self.tree.heading("selected", text="Restore?")
        self.tree.heading("source_db", text="Backup DB Name")
        self.tree.heading("target_db", text="Target DB Name")
        self.tree.heading("size_gb", text="Size (GB)")
        self.tree.heading("stripes", text="Stripes / Buffers")
        self.tree.heading("format_ts", text="Format & Timestamp")
        self.tree.heading("status", text="Status / Progress")
        self.tree.heading("source_location", text="Source File(s) / Target Relocation")

        self.tree.column("selected", width=70, anchor=tk.CENTER, stretch=False)
        self.tree.column("source_db", width=165, anchor=tk.W)
        self.tree.column("target_db", width=175, anchor=tk.W)
        self.tree.column("size_gb", width=95, anchor=tk.E, stretch=False)
        self.tree.column("stripes", width=155, anchor=tk.CENTER, stretch=False)
        self.tree.column("format_ts", width=155, anchor=tk.CENTER, stretch=False)
        self.tree.column("status", width=215, anchor=tk.W)
        self.tree.column("source_location", width=310, anchor=tk.W)

        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.tree.bind("<ButtonRelease-1>", self._on_tree_click)
        self.tree.bind("<Double-1>", self._on_tree_double_click)
        self.tree.bind("<space>", self._on_tree_space)
        self.tree.bind("<Button-3>", self._on_tree_right_click)
        self.tree.bind("<<TreeviewSelect>>", lambda _e: self.update_tsql_preview())

        # 4B. Bottom Panel: T-SQL Preview + Action Controls + Execution Log
        bottom_box = ttk.Frame(paned)
        paned.add(bottom_box, weight=2)

        split_bottom = ttk.PanedWindow(bottom_box, orient=tk.HORIZONTAL)
        split_bottom.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        preview_frame = ttk.LabelFrame(
            split_bottom,
            text=" Live High-Speed RESTORE DATABASE T-SQL Preview ",
            padding="6 4 6 4",
        )
        split_bottom.add(preview_frame, weight=1)

        preview_btn_bar = ttk.Frame(preview_frame)
        preview_btn_bar.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(
            preview_btn_bar,
            text="Shows exact multi-stripe RESTORE DATABASE T-SQL (Largest-First Execution Order):",
        ).pack(side=tk.LEFT)
        ttk.Button(
            preview_btn_bar,
            text="Copy T-SQL",
            command=self.copy_tsql_to_clipboard,
        ).pack(side=tk.RIGHT)

        self.txt_tsql = tk.Text(
            preview_frame,
            height=8,
            font=("Consolas", 9),
            bg="#f8fafc",
            fg="#0f172a",
            wrap=tk.NONE,
        )
        self.txt_tsql.pack(fill=tk.BOTH, expand=True)

        log_frame = ttk.LabelFrame(
            split_bottom,
            text=" Parallel Restore & Fast-Unzip Execution Log ",
            padding="6 4 6 4",
        )
        split_bottom.add(log_frame, weight=1)

        log_btn_bar = ttk.Frame(log_frame)
        log_btn_bar.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(log_btn_bar, textvariable=self.var_progress_label, style="Bold.TLabel").pack(
            side=tk.LEFT
        )
        ttk.Button(log_btn_bar, text="Clear Log", command=self.clear_log).pack(side=tk.RIGHT)

        self.txt_log = tk.Text(
            log_frame,
            height=8,
            font=("Consolas", 9),
            bg="#0f172a",
            fg="#e2e8f0",
            wrap=tk.WORD,
        )
        self.txt_log.pack(fill=tk.BOTH, expand=True)

        # Bottom Action Bar
        action_bar = ttk.Frame(bottom_box)
        action_bar.pack(fill=tk.X)

        self.btn_start = ttk.Button(
            action_bar,
            text="▶  Start Parallel Striped Restores",
            command=self.start_parallel_restores,
        )
        self.btn_start.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_cancel = ttk.Button(
            action_bar,
            text="■  Cancel Restores",
            command=self.cancel_restores,
            state="disabled",
        )
        self.btn_cancel.pack(side=tk.LEFT, padx=(0, 14))

        self.progress_bar = ttk.Progressbar(
            action_bar,
            variable=self.var_overall_progress,
            maximum=100.0,
            mode="determinate",
        )
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)

    # ========================================================================
    # UI Event Handlers & Helpers
    # ========================================================================

    def log(self, message: str) -> None:
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        self.txt_log.insert(tk.END, f"[{ts}] {message}\n")
        self.txt_log.see(tk.END)

    def clear_log(self) -> None:
        self.txt_log.delete("1.0", tk.END)

    def _toggle_password_visibility(self) -> None:
        self.ent_pass.configure(show="" if self.var_show_password.get() else "*")

    def _on_auth_mode_change(self) -> None:
        is_sql = self.var_auth_mode.get().lower() == "sql"
        state = "normal" if is_sql else "disabled"
        self.ent_user.configure(state=state)
        self.ent_pass.configure(state=state)

    def _on_relocate_toggle_change(self) -> None:
        enabled = bool(self.var_relocate_files.get())
        state = "normal" if enabled else "disabled"
        self.ent_data_folder.configure(state=state)
        self.btn_browse_data.configure(state=state)
        self.ent_log_folder.configure(state=state)
        self.btn_browse_log.configure(state=state)
        self._on_settings_changed()

    def browse_source_folder(self) -> None:
        initial = self.var_source_folder.get().strip() or str(Path.home())
        chosen = filedialog.askdirectory(
            parent=self.root,
            title="Select Source Backup Folder (.bak / .zip)",
            initialdir=initial if os.path.isdir(initial) else str(Path.home()),
        )
        if chosen:
            if os.name == "nt":
                chosen = chosen.replace("/", "\\")
            self.var_source_folder.set(chosen)
            self.scan_source_backups(silent_if_missing=False)

    def browse_data_folder(self) -> None:
        initial = self.var_target_data_folder.get().strip() or str(Path.home())
        chosen = filedialog.askdirectory(
            parent=self.root,
            title="Select Target SQL Data Folder (.mdf / .ndf)",
            initialdir=initial if os.path.isdir(initial) else str(Path.home()),
        )
        if chosen:
            if os.name == "nt":
                chosen = chosen.replace("/", "\\")
            self.var_target_data_folder.set(chosen)
            self._on_settings_changed()

    def browse_log_folder(self) -> None:
        initial = self.var_target_log_folder.get().strip() or str(Path.home())
        chosen = filedialog.askdirectory(
            parent=self.root,
            title="Select Target SQL Log Folder (.ldf)",
            initialdir=initial if os.path.isdir(initial) else str(Path.home()),
        )
        if chosen:
            if os.name == "nt":
                chosen = chosen.replace("/", "\\")
            self.var_target_log_folder.set(chosen)
            self._on_settings_changed()

    def _get_safe_spinbox_values(self) -> Tuple[int, int]:
        try:
            max_parallel = int(self.var_max_parallel_dbs.get())
            if max_parallel < 1:
                max_parallel = 1
        except (tk.TclError, ValueError):
            max_parallel = DEFAULT_MAX_PARALLEL_DBS

        try:
            custom_buf = int(self.var_custom_buffercount.get())
            if custom_buf < 0:
                custom_buf = 0
        except (tk.TclError, ValueError):
            custom_buf = 0

        return max_parallel, custom_buf

    def _on_settings_changed(self) -> None:
        self.refresh_backup_tree()
        self.update_tsql_preview()

    def _build_current_connection_config(self) -> SqlConnectionConfig:
        return SqlConnectionConfig(
            server=self.var_server.get().strip() or "localhost",
            auth_mode=self.var_auth_mode.get().strip() or "windows",
            username=self.var_username.get().strip(),
            password=self.var_password.get(),
            trust_server_certificate=bool(self.var_trust_cert.get()),
        )

    def _build_current_restore_options(self) -> RestoreOptions:
        max_parallel, custom_buf = self._get_safe_spinbox_values()
        source_folder = self.var_source_folder.get().strip() or r"C:\SqlBackups"
        return RestoreOptions(
            source_folder=source_folder,
            staging_folder=resolve_staging_folder(source_folder),
            target_data_folder=self.var_target_data_folder.get().strip() or r"C:\SqlData",
            target_log_folder=self.var_target_log_folder.get().strip() or r"C:\SqlLogs",
            relocate_files=bool(self.var_relocate_files.get()),
            target_db_prefix=self.var_db_prefix.get().strip(),
            target_db_suffix=self.var_db_suffix.get().strip(),
            max_parallel_dbs=max_parallel,
            with_replace=bool(self.var_with_replace.get()),
            kill_connections=bool(self.var_kill_connections.get()),
            recovery_mode=self.var_recovery_mode.get().strip() or "RECOVERY",
            high_speed_tuning=bool(self.var_high_speed_tuning.get()),
            max_transfer_size=DEFAULT_MAX_TRANSFER_SIZE,
            custom_buffer_count=custom_buf,
            keep_cdc=bool(self.var_keep_cdc.get()),
            cleanup_staging_bak=bool(self.var_cleanup_staging.get()),
            latest_only_per_db=bool(self.var_latest_only.get()),
        )

    # ========================================================================
    # Config Load / Save
    # ========================================================================

    def save_config(self) -> None:
        max_parallel, custom_buf = self._get_safe_spinbox_values()
        data = {
            "server": self.var_server.get().strip(),
            "auth_mode": self.var_auth_mode.get().strip(),
            "username": self.var_username.get().strip(),
            "trust_server_certificate": bool(self.var_trust_cert.get()),
            "source_folder": self.var_source_folder.get().strip(),
            "latest_only_per_db": bool(self.var_latest_only.get()),
            "cleanup_staging_bak": bool(self.var_cleanup_staging.get()),
            "relocate_files": bool(self.var_relocate_files.get()),
            "target_data_folder": self.var_target_data_folder.get().strip(),
            "target_log_folder": self.var_target_log_folder.get().strip(),
            "target_db_prefix": self.var_db_prefix.get().strip(),
            "target_db_suffix": self.var_db_suffix.get().strip(),
            "max_parallel_dbs": max_parallel,
            "with_replace": bool(self.var_with_replace.get()),
            "kill_connections": bool(self.var_kill_connections.get()),
            "recovery_mode": self.var_recovery_mode.get().strip(),
            "high_speed_tuning": bool(self.var_high_speed_tuning.get()),
            "custom_buffer_count": custom_buf,
            "keep_cdc": bool(self.var_keep_cdc.get()),
        }
        try:
            self.config_file_path.write_text(json.dumps(data, indent=2), encoding="utf-8")
            self.log(f"Saved configuration to {self.config_file_path.name}")
        except Exception as exc:
            messagebox.showerror("Save Config Failed", str(exc), parent=self.root)

    def _load_saved_config(self, silent: bool = False) -> None:
        if not self.config_file_path.is_file():
            if not silent:
                messagebox.showinfo(
                    "No Saved Config",
                    f"No config file found at {self.config_file_path}",
                    parent=self.root,
                )
            return
        try:
            data = json.loads(self.config_file_path.read_text(encoding="utf-8"))
            self.var_server.set(data.get("server", "localhost"))
            self.var_auth_mode.set(data.get("auth_mode", "windows"))
            self.var_username.set(data.get("username", "sa"))
            self.var_trust_cert.set(bool(data.get("trust_server_certificate", True)))
            if data.get("source_folder"):
                self.var_source_folder.set(data["source_folder"])
            if "latest_only_per_db" in data:
                self.var_latest_only.set(bool(data["latest_only_per_db"]))
            if "cleanup_staging_bak" in data:
                self.var_cleanup_staging.set(bool(data["cleanup_staging_bak"]))
            if "relocate_files" in data:
                self.var_relocate_files.set(bool(data["relocate_files"]))
            if data.get("target_data_folder"):
                self.var_target_data_folder.set(data["target_data_folder"])
            if data.get("target_log_folder"):
                self.var_target_log_folder.set(data["target_log_folder"])
            self.var_db_prefix.set(data.get("target_db_prefix", ""))
            self.var_db_suffix.set(data.get("target_db_suffix", ""))
            self.var_max_parallel_dbs.set(
                int(data.get("max_parallel_dbs", DEFAULT_MAX_PARALLEL_DBS))
            )
            self.var_with_replace.set(bool(data.get("with_replace", True)))
            self.var_kill_connections.set(bool(data.get("kill_connections", True)))
            self.var_recovery_mode.set(data.get("recovery_mode", "RECOVERY"))
            self.var_high_speed_tuning.set(bool(data.get("high_speed_tuning", True)))
            self.var_custom_buffercount.set(int(data.get("custom_buffer_count", 0)))
            self.var_keep_cdc.set(bool(data.get("keep_cdc", False)))
            self._on_auth_mode_change()
            self._on_relocate_toggle_change()
            if not silent:
                self.log(f"Loaded configuration from {self.config_file_path.name}")
        except Exception as exc:
            if not silent:
                messagebox.showerror("Load Config Failed", str(exc), parent=self.root)

    # ========================================================================
    # SQL Instance Connection & Backup Discovery
    # ========================================================================

    def connect_and_detect_instance(self) -> None:
        if self.is_running_restores:
            return
        cfg = self._build_current_connection_config()
        self.btn_connect.configure(state="disabled")
        self.var_server_status.set(f"Connecting to {cfg.server}...")
        self.log(f"Connecting to target SQL Server '{cfg.server}' ({cfg.auth_mode} auth)...")

        def worker() -> None:
            try:
                engine = SqlRestoreEngine(cfg)
                inst_info = engine.fetch_instance_info()
                self.ui_queue.put(("INSTANCE_LOADED", inst_info))
            except Exception as exc:
                self.ui_queue.put(("INSTANCE_ERROR", str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def scan_source_backups(self, silent_if_missing: bool = False) -> None:
        if self.is_running_restores:
            return
        folder = self.var_source_folder.get().strip()
        if not folder or not os.path.isdir(folder):
            if not silent_if_missing:
                messagebox.showwarning(
                    "Source Folder Not Found",
                    f"Source folder does not exist:\n{folder}",
                    parent=self.root,
                )
            return

        sets = scan_backup_source_folder(
            source_folder=folder,
            latest_only_per_db=bool(self.var_latest_only.get()),
            recursive=True,
        )
        self.all_backup_sets = sets
        self.selected_set_ids = {s.set_id for s in sets if s.is_complete}
        self.set_status_map = {
            s.set_id: ("Ready" if s.is_complete else f"⚠ Missing stripes ({s.found_stripes}/{s.expected_stripes})")
            for s in sets
        }
        self.log(
            f"Scanned '{folder}': discovered {len(sets)} backup set(s) "
            f"({'latest per DB' if self.var_latest_only.get() else 'all timestamps'})."
        )
        self.refresh_backup_tree()
        self.update_tsql_preview()

    def inspect_selected_sql_headers(self) -> None:
        if self.is_running_restores:
            return
        raw_sets = [
            s
            for s in self.all_backup_sets
            if s.set_id in self.selected_set_ids and s.source_type == "bak" and s.stripe_paths
        ]
        if not raw_sets:
            messagebox.showinfo(
                "Verify SQL Headers",
                "Select at least one raw .bak backup set to query SQL Server RESTORE HEADERONLY / FILELISTONLY.\n"
                "(Note: .zip sets are automatically inspected after fast-unzip staging during restore.)",
                parent=self.root,
            )
            return

        cfg = self._build_current_connection_config()
        self.btn_inspect_headers.configure(state="disabled")
        self.log(f"Querying RESTORE HEADERONLY & FILELISTONLY for {len(raw_sets)} backup set(s)...")

        def worker() -> None:
            engine = SqlRestoreEngine(cfg)
            inspected_count = 0
            for bset in raw_sets:
                try:
                    db_name, fam_count, lfiles = engine.inspect_backup_files_metadata(
                        bset.stripe_paths[0]
                    )
                    if db_name:
                        bset.source_db_name = db_name
                    bset.expected_stripes = max(bset.expected_stripes, fam_count)
                    if lfiles:
                        bset.logical_files = lfiles
                    bset.header_verified = True
                    inspected_count += 1
                    self.ui_queue.put(
                        (
                            "HEADER_INSPECTED",
                            (
                                bset.set_id,
                                f"✓ Verified ({len(lfiles)} files, {bset.found_stripes}/{bset.expected_stripes} stripes)",
                            ),
                        )
                    )
                except Exception as exc:
                    self.ui_queue.put(("HEADER_ERROR", (bset.set_id, str(exc))))
            self.ui_queue.put(("HEADER_BATCH_DONE", inspected_count))

        threading.Thread(target=worker, daemon=True).start()

    def refresh_backup_tree(self) -> None:
        focused_iid = self.tree.focus()
        for item in self.tree.get_children():
            self.tree.delete(item)

        options = self._build_current_restore_options()
        filter_q = self.var_filter_text.get().strip().lower()

        for bset in self.all_backup_sets:
            target_db = resolve_target_db_name(
                bset.source_db_name,
                options.target_db_prefix,
                options.target_db_suffix,
                bset.custom_target_name,
            )
            if filter_q and (
                filter_q not in bset.source_db_name.lower()
                and filter_q not in target_db.lower()
            ):
                continue

            is_sel = bset.set_id in self.selected_set_ids
            check_mark = "[✓] Yes" if is_sel else "[  ] No"

            buf_cnt = calculate_optimal_buffercount(
                bset.stream_count, options.custom_buffer_count
            )
            if bset.stream_count == 1:
                stripe_label = f"1 stripe (Buf={buf_cnt})"
            else:
                stripe_label = f"⚡ {bset.found_stripes}/{bset.expected_stripes} stripes (Buf={buf_cnt})"

            fmt_badge = "ZIP Archive" if bset.source_type == "zip" else "RAW .bak"
            format_ts = f"{fmt_badge} • {bset.timestamp_label}"

            exists_note = (
                " (Overwrite)" if target_db.lower() in self.existing_server_dbs else ""
            )
            target_display = f"{target_db}{exists_note}"

            src_path_display = (
                bset.zip_path
                if bset.source_type == "zip"
                else (bset.stripe_paths[0] if bset.stripe_paths else "")
            )
            if options.relocate_files:
                loc_display = f"{os.path.basename(src_path_display)} ➜ {options.target_data_folder}"
            else:
                loc_display = src_path_display

            status_txt = self.set_status_map.get(bset.set_id, "Ready")

            self.tree.insert(
                "",
                tk.END,
                iid=bset.set_id,
                values=(
                    check_mark,
                    bset.source_db_name,
                    target_display,
                    f"{bset.size_gb:,.2f} GB",
                    stripe_label,
                    format_ts,
                    status_txt,
                    loc_display,
                ),
            )

        if focused_iid and self.tree.exists(focused_iid):
            self.tree.focus(focused_iid)
            self.tree.selection_set(focused_iid)

        self._update_selection_summary()

    def _update_selection_summary(self) -> None:
        selected_sets = [s for s in self.all_backup_sets if s.set_id in self.selected_set_ids]
        total_gb = sum(s.size_gb for s in selected_sets)
        total_stripes = sum(s.stream_count for s in selected_sets)
        striped_count = sum(1 for s in selected_sets if s.stream_count > 1)
        zip_count = sum(1 for s in selected_sets if s.source_type == "zip")
        zip_badge = f"  |  {zip_count} .zip" if zip_count > 0 else ""
        self.var_selection_summary.set(
            f"Selected: {len(selected_sets)} DBs ({total_gb:,.2f} GB  |  "
            f"{striped_count} multi-stripe  |  {total_stripes} total streams{zip_badge})"
        )

    def _on_tree_click(self, event: tk.Event) -> None:
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        if row_id and col == "#1":
            self._toggle_set_selection(row_id)

    def _on_tree_double_click(self, event: tk.Event) -> None:
        row_id = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not row_id:
            return
        if col == "#3":
            self.prompt_rename_selected_db()
        else:
            self._toggle_set_selection(row_id)

    def _on_tree_right_click(self, event: tk.Event) -> None:
        row_id = self.tree.identify_row(event.y)
        if row_id:
            self.tree.focus(row_id)
            self.tree.selection_set(row_id)
            self.prompt_rename_selected_db()

    def _on_tree_space(self, _event: tk.Event) -> None:
        focused = self.tree.focus()
        if focused:
            self._toggle_set_selection(focused)

    def _toggle_set_selection(self, set_id: str) -> None:
        if set_id in self.selected_set_ids:
            self.selected_set_ids.remove(set_id)
        else:
            self.selected_set_ids.add(set_id)
        self.refresh_backup_tree()
        self.update_tsql_preview()

    def prompt_rename_selected_db(self) -> None:
        focused = self.tree.focus()
        if not focused:
            messagebox.showinfo(
                "Rename Target DB",
                "Highlight a backup row first to customize its target database name.",
                parent=self.root,
            )
            return
        bset = next((s for s in self.all_backup_sets if s.set_id == focused), None)
        if not bset:
            return
        options = self._build_current_restore_options()
        current_target = resolve_target_db_name(
            bset.source_db_name,
            options.target_db_prefix,
            options.target_db_suffix,
            bset.custom_target_name,
        )
        new_name = simpledialog.askstring(
            "Rename Target Database",
            f"Enter custom target database name for backup '{bset.source_db_name}'\n"
            "(leave blank to revert to default name + prefix/suffix):",
            initialvalue=current_target,
            parent=self.root,
        )
        if new_name is not None:
            bset.custom_target_name = new_name.strip()
            self.refresh_backup_tree()
            self.update_tsql_preview()

    def select_all_sets(self) -> None:
        filter_q = self.var_filter_text.get().strip().lower()
        for s in self.all_backup_sets:
            if not filter_q or filter_q in s.source_db_name.lower():
                self.selected_set_ids.add(s.set_id)
        self.refresh_backup_tree()
        self.update_tsql_preview()

    def deselect_all_sets(self) -> None:
        self.selected_set_ids.clear()
        self.refresh_backup_tree()
        self.update_tsql_preview()

    def select_striped_sets(self) -> None:
        self.selected_set_ids = {s.set_id for s in self.all_backup_sets if s.stream_count > 1}
        self.refresh_backup_tree()
        self.update_tsql_preview()

    def invert_selection(self) -> None:
        all_ids = {s.set_id for s in self.all_backup_sets}
        self.selected_set_ids = all_ids - self.selected_set_ids
        self.refresh_backup_tree()
        self.update_tsql_preview()

    # ========================================================================
    # Live High-Speed RESTORE DATABASE T-SQL Preview
    # ========================================================================

    def update_tsql_preview(self) -> None:
        options = self._build_current_restore_options()
        focused = self.tree.focus()
        set_lookup = {s.set_id: s for s in self.all_backup_sets}

        selected_ordered = sort_backup_sets_largest_first(
            [s for s in self.all_backup_sets if s.set_id in self.selected_set_ids]
        )

        preview_sets: List[DiscoveredBackupSet] = []
        if focused and focused in set_lookup:
            preview_sets.append(set_lookup[focused])
        for s in selected_ordered:
            if s not in preview_sets:
                preview_sets.append(s)

        if not preview_sets:
            sample_striped = DiscoveredBackupSet(
                set_id="sample_1",
                source_db_name="EnterpriseWarehouseDB",
                timestamp_label="20261021_100000",
                source_type="bak",
                stripe_paths=[
                    join_windows_or_native_path(
                        options.source_folder,
                        "EnterpriseWarehouseDB",
                        f"EnterpriseWarehouseDB_20261021_100000_stripe0{i}of04.bak",
                    )
                    for i in range(1, 5)
                ],
                expected_stripes=4,
                found_stripes=4,
                total_size_bytes=28 * 1024 * 1024 * 1024,
                logical_files=default_logical_files_for_db("EnterpriseWarehouseDB"),
            )
            sample_zip = DiscoveredBackupSet(
                set_id="sample_2",
                source_db_name="AppConfigDB",
                timestamp_label="20261021_100000",
                source_type="zip",
                stripe_paths=["AppConfigDB_20261021_100000.bak"],
                zip_path=join_windows_or_native_path(
                    options.source_folder, "Zipped", "AppConfigDB_20261021_100000.zip"
                ),
                expected_stripes=1,
                found_stripes=1,
                total_size_bytes=2 * 1024 * 1024 * 1024,
                logical_files=default_logical_files_for_db("AppConfigDB"),
            )
            preview_sets = [sample_striped, sample_zip]
            header_comment = (
                "-- [Preview Example] Scan a backup folder above to see exact T-SQL for your backup sets.\n"
                f"-- Parallel Workers: {options.max_parallel_dbs} concurrent DBs | "
                "Queue Order: Largest-Database-First (LPT)\n"
                f"-- High-Speed Tuning: {'ENABLED (MAXTRANSFERSIZE = 4194304 + Auto BUFFERCOUNT)' if options.high_speed_tuning else 'DISABLED'}\n\n"
            )
        else:
            header_comment = (
                f"-- Parallel Restore Plan: Up to {options.max_parallel_dbs} databases restored concurrently (Largest-DB-First order)\n"
                f"-- High-Speed Tuning: {'ENABLED (MAXTRANSFERSIZE = 4194304 + Auto BUFFERCOUNT)' if options.high_speed_tuning else 'DISABLED'}\n\n"
            )

        blocks: List[str] = [header_comment]
        if options.relocate_files:
            blocks.append(
                f"-- Ensure Target Data & Log Folders exist on SQL Server host\n"
                f"{build_create_subdir_tsql(options.target_data_folder)}\n"
                f"{build_create_subdir_tsql(options.target_log_folder)}\n"
            )

        for bset in preview_sets[:6]:
            target_db = resolve_target_db_name(
                bset.source_db_name,
                options.target_db_prefix,
                options.target_db_suffix,
                bset.custom_target_name,
            )
            if bset.source_type == "zip":
                staging_sub = join_windows_or_native_path(
                    options.staging_folder,
                    f"{sanitize_filename(bset.source_db_name)}_{sanitize_filename(bset.timestamp_label)}",
                )
                effective_paths = [
                    join_windows_or_native_path(staging_sub, os.path.basename(p))
                    for p in bset.stripe_paths
                ]
                unzip_comment = (
                    f"-- Fast-Unzip Archive: {bset.zip_path} -> {staging_sub}"
                    + (" (auto-cleanup after restore)\n" if options.cleanup_staging_bak else "\n")
                )
            else:
                effective_paths = bset.stripe_paths
                unzip_comment = ""

            tsql = build_restore_tsql(
                source_db_name=bset.source_db_name,
                target_db_name=target_db,
                stripe_file_paths=effective_paths,
                logical_files=bset.logical_files,
                relocate_files=options.relocate_files,
                target_data_folder=options.target_data_folder,
                target_log_folder=options.target_log_folder,
                with_replace=options.with_replace,
                kill_connections=options.kill_connections,
                recovery_mode=options.recovery_mode,
                high_speed_tuning=options.high_speed_tuning,
                max_transfer_size=options.max_transfer_size,
                custom_buffer_count=options.custom_buffer_count,
                keep_cdc=options.keep_cdc,
                stats_interval=options.stats_interval,
            )
            block = (
                f"-- Database: [{bset.source_db_name}] -> [{target_db}] "
                f"({bset.size_gb:,.2f} GB | {bset.stream_count} {'stripe' if bset.stream_count == 1 else 'stripes'})\n"
                f"{unzip_comment}"
                f"{tsql}\n"
            )
            blocks.append(block)

        if len(preview_sets) > 6:
            blocks.append(f"-- ... plus {len(preview_sets) - 6} more selected backup set(s) ...\n")

        self.txt_tsql.delete("1.0", tk.END)
        self.txt_tsql.insert("1.0", "\n".join(blocks))

    def copy_tsql_to_clipboard(self) -> None:
        tsql_text = self.txt_tsql.get("1.0", tk.END).strip()
        if tsql_text:
            self.root.clipboard_clear()
            self.root.clipboard_append(tsql_text)
            self.log("Copied RESTORE DATABASE T-SQL preview to clipboard.")

    # ========================================================================
    # Parallel Restore Execution Coordinator (Largest-Database-First)
    # ========================================================================

    def start_parallel_restores(self) -> None:
        if self.is_running_restores:
            return

        selected_sets = [s for s in self.all_backup_sets if s.set_id in self.selected_set_ids]
        if not selected_sets:
            messagebox.showwarning(
                "No Backup Sets Selected",
                "Please select at least one backup set from the table before starting restores.",
                parent=self.root,
            )
            return

        incomplete = [s for s in selected_sets if not s.is_complete]
        if incomplete:
            names = ", ".join(
                f"{s.source_db_name} ({s.found_stripes}/{s.expected_stripes})"
                for s in incomplete[:5]
            )
            if not messagebox.askyesno(
                "Incomplete Stripe Sets Detected",
                f"The following backup sets appear to be missing stripe files:\n{names}\n\n"
                "Do you still want to attempt restoring the selected sets?",
                parent=self.root,
            ):
                return

        options = self._build_current_restore_options()
        # Schedule largest databases first (LPT) so multi-stripe DBs start immediately
        ordered_sets = sort_backup_sets_largest_first(selected_sets)

        cfg = self._build_current_connection_config()
        self.is_running_restores = True
        self.cancel_event.clear()
        self.btn_start.configure(state="disabled")
        self.btn_connect.configure(state="disabled")
        self.btn_scan.configure(state="disabled")
        self.btn_inspect_headers.configure(state="disabled")
        self.btn_cancel.configure(state="normal")
        self.var_overall_progress.set(0.0)

        for s in ordered_sets:
            self.set_status_map[s.set_id] = "Queued (LPT)"
        self.refresh_backup_tree()

        self.log(
            f"Starting parallel striped restore batch for {len(ordered_sets)} database(s) "
            f"(Max Parallel DBs: {options.max_parallel_dbs}, Largest-First order, "
            f"High-Speed 4MB Buffers: {options.high_speed_tuning})..."
        )

        def coordinator() -> None:
            engine = SqlRestoreEngine(cfg)
            results: List[RestoreJobResult] = []
            progress_pct: Dict[str, int] = {s.set_id: 0 for s in ordered_sets}
            lock = threading.Lock()

            def on_set_progress(set_id: str, pct: int, status_msg: str) -> None:
                with lock:
                    if pct >= 0:
                        progress_pct[set_id] = pct
                    avg_pct = sum(progress_pct.values()) / max(1, len(ordered_sets))
                self.ui_queue.put(("SET_PROGRESS", (set_id, pct, status_msg, avg_pct)))

            with concurrent.futures.ThreadPoolExecutor(
                max_workers=options.max_parallel_dbs
            ) as executor:
                future_to_set = {
                    executor.submit(
                        engine.run_single_database_restore,
                        bset,
                        options,
                        self.cancel_event,
                        on_set_progress,
                    ): bset
                    for bset in ordered_sets
                }
                for future in concurrent.futures.as_completed(future_to_set):
                    bset = future_to_set[future]
                    target_db = resolve_target_db_name(
                        bset.source_db_name,
                        options.target_db_prefix,
                        options.target_db_suffix,
                        bset.custom_target_name,
                    )
                    try:
                        res = future.result()
                    except Exception as exc:
                        res = RestoreJobResult(
                            set_id=bset.set_id,
                            source_db_name=bset.source_db_name,
                            target_db_name=target_db,
                            success=False,
                            stream_count=bset.stream_count,
                            source_type=bset.source_type,
                            error_message=str(exc),
                        )
                    results.append(res)
                    with lock:
                        progress_pct[bset.set_id] = 100
                        avg_pct = sum(progress_pct.values()) / max(1, len(ordered_sets))
                    self.ui_queue.put(("SET_FINISHED", (res, avg_pct)))

            self.ui_queue.put(("BATCH_FINISHED", results))

        threading.Thread(target=coordinator, daemon=True).start()

    def cancel_restores(self) -> None:
        if self.is_running_restores:
            self.cancel_event.set()
            self.log("Cancellation requested — stopping active and queued restore workers...")
            self.btn_cancel.configure(state="disabled")

    def _poll_ui_queue(self) -> None:
        try:
            while True:
                msg_type, payload = self.ui_queue.get_nowait()

                if msg_type == "INSTANCE_LOADED":
                    info: SqlInstanceInfo = payload  # type: ignore[assignment]
                    self.btn_connect.configure(state="normal")
                    self.existing_server_dbs = {d.lower() for d in info.existing_databases}
                    if info.default_data_path:
                        self.var_target_data_folder.set(info.default_data_path)
                    if info.default_log_path:
                        self.var_target_log_folder.set(info.default_log_path)

                    ifi_badge = f"IFI: {info.ifi_enabled}"
                    status_line = (
                        f"Connected: {info.server_name} — {info.edition} "
                        f"(v{info.product_version}) | {ifi_badge} | {len(info.existing_databases)} existing DBs"
                    )
                    self.var_server_status.set(status_line)
                    self.log(
                        f"Connected to {info.server_name} ({info.edition}). "
                        f"Default Data: '{info.default_data_path}', Default Log: '{info.default_log_path}', "
                        f"Instant File Initialization: {info.ifi_enabled}."
                    )
                    if info.ifi_enabled.lower() == "disabled":
                        self.log(
                            "TIP: Instant File Initialization (IFI) is Disabled on this SQL Server service account. "
                            "Granting 'Perform volume maintenance tasks' (SE_MANAGE_VOLUME_NAME) can speed up data file creation by 2x-5x."
                        )
                    self.refresh_backup_tree()
                    self.update_tsql_preview()

                elif msg_type == "INSTANCE_ERROR":
                    err_str = str(payload)
                    self.btn_connect.configure(state="normal")
                    self.var_server_status.set(
                        "Connection failed — check Server/Instance and credentials"
                    )
                    self.log(f"ERROR connecting to SQL Server: {err_str}")
                    messagebox.showerror("SQL Server Connection Error", err_str, parent=self.root)

                elif msg_type == "HEADER_INSPECTED":
                    set_id, status_msg = payload  # type: ignore[misc]
                    self.set_status_map[set_id] = status_msg
                    self.refresh_backup_tree()
                    self.update_tsql_preview()

                elif msg_type == "HEADER_ERROR":
                    set_id, err_msg = payload  # type: ignore[misc]
                    self.set_status_map[set_id] = f"Header Error: {err_msg[:45]}"
                    self.log(f"Header inspection error ({set_id}): {err_msg}")
                    self.refresh_backup_tree()

                elif msg_type == "HEADER_BATCH_DONE":
                    count = int(payload)  # type: ignore[arg-type]
                    self.btn_inspect_headers.configure(state="normal")
                    self.log(f"Completed SQL header & logical file verification for {count} backup set(s).")

                elif msg_type == "SET_PROGRESS":
                    set_id, pct, status_msg, avg_pct = payload  # type: ignore[misc]
                    self.set_status_map[set_id] = status_msg
                    if self.tree.exists(set_id):
                        vals = list(self.tree.item(set_id, "values"))
                        vals[6] = status_msg
                        self.tree.item(set_id, values=vals)
                    self.var_overall_progress.set(avg_pct)
                    self.var_progress_label.set(f"Restoring... ({avg_pct:.0f}%)")
                    if pct in (0, 2, 15, -1):
                        self.log(f"[{set_id.split('::')[0]}] {status_msg}")

                elif msg_type == "SET_FINISHED":
                    res, avg_pct = payload  # type: ignore[misc]
                    self.var_overall_progress.set(avg_pct)
                    if res.success:
                        unzip_note = (
                            f", unzipped {res.unzip_elapsed_seconds:.1f}s"
                            if res.unzip_elapsed_seconds > 0
                            else ""
                        )
                        status_str = (
                            f"✓ Restored in {res.elapsed_seconds:.1f}s "
                            f"({res.stream_count} {'stripe' if res.stream_count == 1 else 'stripes'}{unzip_note})"
                        )
                        self.set_status_map[res.set_id] = status_str
                        self.existing_server_dbs.add(res.target_db_name.lower())
                        self.log(
                            f"✓ [{res.source_db_name} -> {res.target_db_name}] Restored in {res.elapsed_seconds:.1f}s "
                            f"({res.stream_count} stripe(s){unzip_note})"
                        )
                    else:
                        status_str = f"✗ Failed: {res.error_message[:60]}"
                        self.set_status_map[res.set_id] = status_str
                        self.log(
                            f"✗ [{res.source_db_name} -> {res.target_db_name}] FAILED: {res.error_message}"
                        )

                    if self.tree.exists(res.set_id):
                        vals = list(self.tree.item(res.set_id, "values"))
                        vals[6] = status_str
                        self.tree.item(res.set_id, values=vals)

                elif msg_type == "BATCH_FINISHED":
                    results: List[RestoreJobResult] = payload  # type: ignore[assignment]
                    self.is_running_restores = False
                    self.btn_start.configure(state="normal")
                    self.btn_connect.configure(state="normal")
                    self.btn_scan.configure(state="normal")
                    self.btn_inspect_headers.configure(state="normal")
                    self.btn_cancel.configure(state="disabled")

                    succeeded = sum(1 for r in results if r.success)
                    failed = len(results) - succeeded
                    summary = f"Restore batch finished: {succeeded} succeeded, {failed} failed."
                    self.var_progress_label.set(summary)
                    self.log("=" * 68)
                    self.log(summary)
                    if failed == 0 and succeeded > 0:
                        self.var_overall_progress.set(100.0)
        except queue.Empty:
            pass
        finally:
            self.root.after(120, self._poll_ui_queue)


def main() -> None:
    root = tk.Tk()
    SqlRestoreApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
