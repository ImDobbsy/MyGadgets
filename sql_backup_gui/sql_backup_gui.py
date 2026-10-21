#!/usr/bin/env python3
"""
SQL Server Parallel Copy-Only Backup GUI (`sql_backup_gui.py`)
==============================================================
A self-contained graphical utility for backing up user-selected SQL Server
databases in `COPY_ONLY` mode to a target folder, with fast post-backup ZIP archiving.

Key Capabilities:
- Connects via Windows Authentication or SQL Server Authentication.
- Lists all online databases with their size (GB/MB), recovery model, and auto-calculated stream count.
- Runs multiple database backups in parallel (configurable worker pool, default 3 concurrent DBs).
- Automatically splits large databases (> 8 GB by default) into parallel striped backup streams
  (`TO DISK = '..._stripe01of0N.bak', DISK = '..._stripe02of0N.bak', ...`).
- Places each database's backup files into its own subfolder (`<TargetFolder>\\<DbName>\\`) and creates
  the subfolder both locally (if reachable) and on the SQL Server host via `master.dbo.xp_create_subdir`.
- Fast Post-Backup Zipping into a target Zip folder prioritizing **speed of zip over compression ratio**
  (`compresslevel=1` / `Z_BEST_SPEED` with 16 MB streaming buffers, or `7z.exe -tzip -mx=1 -mmt=on` if installed).
- Enables `COPY_ONLY`, `COMPRESSION` (with automatic fallback for SQL Server Express/Web editions),
  `CHECKSUM`, and real-time progress reporting via `STATS = 5`.
- Zero mandatory third-party `pip` dependencies on Windows: uses `pyodbc` if installed, or falls back
  automatically to a built-in `.NET System.Data.SqlClient` PowerShell bridge.
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
from tkinter import filedialog, messagebox, ttk

# Optional pyodbc detection
try:
    import pyodbc  # type: ignore

    HAS_PYODBC = True
except ImportError:
    HAS_PYODBC = False


APP_TITLE = "SQL Server Parallel Copy-Only Backup & Fast-Zip Utility"
CONFIG_FILENAME = ".sql_backup_config.json"
DEFAULT_SPLIT_THRESHOLD_GB = 8.0
DEFAULT_MAX_STREAMS = 8
DEFAULT_MAX_PARALLEL_DBS = 3
FAST_ZIP_CHUNK_BYTES = 16 * 1024 * 1024  # 16 MB streaming buffer for fast I/O


# ============================================================================
# Data Models & Pure T-SQL / Striping / Fast-Zip Helpers
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
            "Application Name=SqlParallelCopyOnlyBackupGui",
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
class DatabaseInfo:
    name: str
    size_mb: float
    recovery_model: str = "FULL"
    state_desc: str = "ONLINE"
    is_system_db: bool = False

    @property
    def size_gb(self) -> float:
        return round(self.size_mb / 1024.0, 3)


@dataclass
class BackupOptions:
    target_folder: str
    split_threshold_gb: float = DEFAULT_SPLIT_THRESHOLD_GB
    max_streams_per_db: int = DEFAULT_MAX_STREAMS
    max_parallel_dbs: int = DEFAULT_MAX_PARALLEL_DBS
    copy_only: bool = True
    compression: bool = True
    checksum: bool = True
    verify_only_after: bool = False
    stats_interval: int = 5
    zip_after_backup: bool = True
    zip_output_folder: str = ""
    delete_bak_after_zip: bool = False


@dataclass
class BackupJobResult:
    db_name: str
    success: bool
    stream_count: int
    file_paths: List[str] = field(default_factory=list)
    zip_path: str = ""
    elapsed_seconds: float = 0.0
    zip_elapsed_seconds: float = 0.0
    used_compression: bool = True
    error_message: str = ""


def calculate_stream_count(
    size_gb: float,
    threshold_gb: float = DEFAULT_SPLIT_THRESHOLD_GB,
    max_streams: int = DEFAULT_MAX_STREAMS,
) -> int:
    """
    Calculate the number of parallel backup streams (striped .bak files) for a database.
    - Databases <= threshold_gb (default 8.0 GB) use 1 stream.
    - Databases > threshold_gb scale dynamically: ceil(size_gb / threshold_gb),
      at least 2 streams, capped at max_streams.
    """
    if threshold_gb <= 0:
        threshold_gb = DEFAULT_SPLIT_THRESHOLD_GB
    if max_streams < 1:
        max_streams = 1
    if size_gb <= threshold_gb or max_streams == 1:
        return 1
    raw_streams = math.ceil(size_gb / threshold_gb)
    return min(max_streams, max(2, raw_streams))


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


def build_backup_file_paths(
    target_folder: str,
    db_name: str,
    stream_count: int,
    timestamp: Optional[str] = None,
) -> Tuple[str, List[str]]:
    """
    Build the per-database target subfolder and list of .bak file paths.
    Returns (db_subfolder, [file_path_1, ...]).
    - If stream_count == 1: <TargetFolder>\\<SafeDb>\\<SafeDb>_<Timestamp>.bak
    - If stream_count > 1:  <TargetFolder>\\<SafeDb>\\<SafeDb>_<Timestamp>_stripe01of0N.bak ...
    """
    if not timestamp:
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_db = sanitize_filename(db_name)
    db_subfolder = join_windows_or_native_path(target_folder, safe_db)

    stream_count = max(1, int(stream_count))
    if stream_count == 1:
        filename = f"{safe_db}_{timestamp}.bak"
        return db_subfolder, [join_windows_or_native_path(db_subfolder, filename)]

    width = max(2, len(str(stream_count)))
    paths: List[str] = []
    for idx in range(1, stream_count + 1):
        filename = f"{safe_db}_{timestamp}_stripe{idx:0{width}d}of{stream_count:0{width}d}.bak"
        paths.append(join_windows_or_native_path(db_subfolder, filename))
    return db_subfolder, paths


def resolve_zip_output_folder(target_folder: str, custom_zip_folder: str = "") -> str:
    """Return the effective Zip output folder (defaults to <TargetFolder>\\Zipped)."""
    if custom_zip_folder and custom_zip_folder.strip():
        return custom_zip_folder.strip()
    return join_windows_or_native_path(target_folder, "Zipped")


def build_zip_file_path(
    target_folder: str,
    db_name: str,
    timestamp: Optional[str] = None,
    custom_zip_folder: str = "",
) -> Tuple[str, str]:
    """
    Build the destination Zip folder and full `.zip` archive path for a database's backup set.
    Returns `(zip_folder, zip_file_path)`.
    """
    if not timestamp:
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_db = sanitize_filename(db_name)
    zip_folder = resolve_zip_output_folder(target_folder, custom_zip_folder)
    zip_filename = f"{safe_db}_{timestamp}.zip"
    return zip_folder, join_windows_or_native_path(zip_folder, zip_filename)


def find_7zip_executable() -> Optional[str]:
    """Locate 7z.exe if installed for multi-threaded fastest-mode (`-mx=1 -mmt=on`) zipping."""
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


def fast_zip_backup_files(
    file_paths: List[str],
    zip_path: str,
    cancel_event: Optional[threading.Event] = None,
    status_callback: Optional[Callable[[str], None]] = None,
    delete_source_after_zip: bool = False,
    source_subfolder: Optional[str] = None,
    prefer_7zip: bool = True,
) -> Tuple[str, float]:
    """
    Archive one or more `.bak` files into `zip_path` prioritizing **speed of zip over compression ratio**.
    - Uses `7z.exe a -tzip -mx=1 -mmt=on` if 7-Zip is installed and `prefer_7zip=True`.
    - Otherwise uses Python `zipfile.ZipFile` with `ZIP_DEFLATED`, `compresslevel=1` (`Z_BEST_SPEED`),
      and a 16 MB streaming buffer (`allowZip64=True`).
    - Optionally deletes the source `.bak` files after verifying the `.zip` archive was created.
    Returns `(zip_path, zip_elapsed_seconds)`.
    """
    t0 = time.time()
    existing_files = [p for p in file_paths if os.path.isfile(p)]
    if not existing_files:
        raise FileNotFoundError(
            f"Cannot create Zip archive: backup file(s) not found on accessible filesystem ({file_paths[0] if file_paths else 'none'})."
        )

    zip_dir = os.path.dirname(zip_path)
    if zip_dir:
        os.makedirs(zip_dir, exist_ok=True)

    # Remove stale zip if overwriting
    if os.path.isfile(zip_path):
        try:
            os.remove(zip_path)
        except OSError:
            pass

    seven_zip = find_7zip_executable() if prefer_7zip else None
    if seven_zip:
        if status_callback:
            status_callback("Fast-Zipping via 7-Zip (-mx=1 -mmt=on)...")
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        cmd = [seven_zip, "a", "-tzip", "-mx=1", "-mmt=on", zip_path] + existing_files
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
                raise RuntimeError("Zip operation cancelled by user.")
            time.sleep(0.15)
        if proc.returncode != 0:
            out = proc.stdout.read() if proc.stdout else ""
            raise RuntimeError(f"7-Zip failed (exit {proc.returncode}): {out.strip()}")
    else:
        total_files = len(existing_files)
        with zipfile.ZipFile(
            zip_path,
            mode="w",
            compression=zipfile.ZIP_DEFLATED,
            compresslevel=1,  # Z_BEST_SPEED: fastest possible zip speed over compression ratio
            allowZip64=True,
        ) as zf:
            for idx, src_path in enumerate(existing_files, start=1):
                if cancel_event and cancel_event.is_set():
                    raise RuntimeError("Zip operation cancelled by user.")
                arcname = os.path.basename(src_path)
                if status_callback:
                    status_callback(f"Fast-Zipping stripe {idx}/{total_files} (compresslevel=1)...")

                st = os.stat(src_path)
                zinfo = zipfile.ZipInfo.from_file(src_path, arcname=arcname)
                zinfo.compress_type = zipfile.ZIP_DEFLATED
                zinfo.compress_level = 1  # type: ignore[attr-defined]
                with open(src_path, "rb") as src_fp, zf.open(zinfo, mode="w", force_zip64=True) as dst_fp:
                    while True:
                        if cancel_event and cancel_event.is_set():
                            raise RuntimeError("Zip operation cancelled by user.")
                        chunk = src_fp.read(FAST_ZIP_CHUNK_BYTES)
                        if not chunk:
                            break
                        dst_fp.write(chunk)

    if not os.path.isfile(zip_path) or os.path.getsize(zip_path) == 0:
        raise RuntimeError(f"Zip archive verification failed: {zip_path} is missing or empty.")

    if delete_source_after_zip:
        for src_path in existing_files:
            try:
                os.remove(src_path)
            except OSError:
                pass
        if source_subfolder and os.path.isdir(source_subfolder):
            try:
                if not os.listdir(source_subfolder):
                    os.rmdir(source_subfolder)
            except OSError:
                pass

    return zip_path, (time.time() - t0)


def escape_tsql_identifier(identifier: str) -> str:
    """Escape a SQL Server bracketed identifier: [My]Db] -> [My]]Db]."""
    return f"[{identifier.replace(']', ']]')}]"


def escape_tsql_string(value: str) -> str:
    """Escape a SQL Server Unicode string literal: N'path'."""
    return "N'" + value.replace("'", "''") + "'"


def build_create_subdir_tsql(subfolder_path: str) -> str:
    """Build T-SQL statement to create the target subfolder on the SQL Server host."""
    return f"EXEC master.dbo.xp_create_subdir {escape_tsql_string(subfolder_path)};"


def build_backup_tsql(
    db_name: str,
    file_paths: List[str],
    copy_only: bool = True,
    compression: bool = True,
    checksum: bool = True,
    stats_interval: int = 5,
) -> str:
    """
    Build the full T-SQL `BACKUP DATABASE` statement supporting 1 or N striped streams.
    """
    if not file_paths:
        raise ValueError("file_paths must contain at least one destination path.")

    escaped_db = escape_tsql_identifier(db_name)
    disk_clauses = [f"    DISK = {escape_tsql_string(p)}" for p in file_paths]
    to_clause = ",\n".join(disk_clauses)

    with_options: List[str] = []
    if copy_only:
        with_options.append("COPY_ONLY")
    if compression:
        with_options.append("COMPRESSION")
    if checksum:
        with_options.append("CHECKSUM")
    with_options.append("INIT")
    stats_val = max(1, min(100, int(stats_interval)))
    with_options.append(f"STATS = {stats_val}")

    with_str = ",\n    ".join(with_options)
    return (
        f"BACKUP DATABASE {escaped_db}\n"
        f"TO\n"
        f"{to_clause}\n"
        f"WITH\n"
        f"    {with_str};"
    )


def build_verify_tsql(file_paths: List[str], checksum: bool = True) -> str:
    """Build `RESTORE VERIFYONLY` T-SQL statement for the single or striped backup set."""
    disk_clauses = [f"    DISK = {escape_tsql_string(p)}" for p in file_paths]
    to_clause = ",\n".join(disk_clauses)
    with_clause = "\nWITH CHECKSUM;" if checksum else ";"
    return f"RESTORE VERIFYONLY\nFROM\n{to_clause}{with_clause}"


def format_size_gb(size_gb: float) -> str:
    if size_gb < 0.01 and size_gb > 0:
        return f"{size_gb * 1024.0:.1f} MB ({size_gb:.3f} GB)"
    return f"{size_gb:,.2f} GB"


# ============================================================================
# SQL Server Execution Engine (pyodbc + Built-In .NET SqlClient Bridge)
# ============================================================================


class SqlServerEngine:
    """
    Executes SQL Server queries, streaming `BACKUP DATABASE` commands, and
    fast post-backup ZIP archiving.
    """

    def __init__(self, config: SqlConnectionConfig):
        self.config = config

    def fetch_databases(
        self, include_system_dbs: bool = False
    ) -> Tuple[str, List[DatabaseInfo]]:
        """
        Connects to the SQL Server instance and returns (server_summary, list_of_databases).
        """
        ps_script = r"""
$ErrorActionPreference = 'Stop'
$connStr = $env:SQL_BG_CONN_STR
$includeSys = ($env:SQL_BG_INCLUDE_SYS -eq '1')

$conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
try {
    $conn.Open()

    $metaCmd = $conn.CreateCommand()
    $metaCmd.CommandTimeout = 30
    $metaCmd.CommandText = @"
SELECT
    CAST(SERVERPROPERTY('ServerName') AS NVARCHAR(256)) AS server_name,
    CAST(SERVERPROPERTY('Edition') AS NVARCHAR(256)) AS edition,
    CAST(SERVERPROPERTY('ProductVersion') AS NVARCHAR(128)) AS product_version;
"@
    $metaReader = $metaCmd.ExecuteReader()
    $serverInfo = @{}
    if ($metaReader.Read()) {
        $serverInfo['server_name'] = [string]$metaReader['server_name']
        $serverInfo['edition'] = [string]$metaReader['edition']
        $serverInfo['product_version'] = [string]$metaReader['product_version']
    }
    $metaReader.Close()

    $dbCmd = $conn.CreateCommand()
    $dbCmd.CommandTimeout = 60
    $dbCmd.CommandText = @"
SELECT
    d.name AS database_name,
    CAST( ROUND(SUM(CAST(mf.size AS BIGINT)) * 8.0 / 1024.0, 2) AS FLOAT ) AS size_mb,
    d.recovery_model_desc,
    d.state_desc,
    CASE WHEN d.database_id <= 4 THEN 1 ELSE 0 END AS is_system_db
FROM sys.databases AS d
LEFT JOIN sys.master_files AS mf
    ON d.database_id = mf.database_id
WHERE d.state_desc = 'ONLINE'
  AND d.name <> 'tempdb'
  AND d.source_database_id IS NULL
GROUP BY
    d.name,
    d.recovery_model_desc,
    d.state_desc,
    d.database_id
ORDER BY
    d.name ASC;
"@
    $reader = $dbCmd.ExecuteReader()
    $dbs = @()
    while ($reader.Read()) {
        $isSys = ([int]$reader['is_system_db'] -eq 1)
        if (-not $includeSys -and $isSys) { continue }
        $dbs += @{
            name = [string]$reader['database_name']
            size_mb = [double]$reader['size_mb']
            recovery_model = [string]$reader['recovery_model_desc']
            state_desc = [string]$reader['state_desc']
            is_system_db = $isSys
        }
    }
    $reader.Close()

    $result = @{
        server = $serverInfo
        databases = $dbs
    }
    Write-Output ($result | ConvertTo-Json -Depth 5 -Compress)
}
finally {
    if ($conn.State -ne 'Closed') { $conn.Close() }
    $conn.Dispose()
}
"""
        env = os.environ.copy()
        env["SQL_BG_CONN_STR"] = self.config.to_dotnet_connection_string("master")
        env["SQL_BG_INCLUDE_SYS"] = "1" if include_system_dbs else "0"

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
        srv = payload.get("server", {})
        server_summary = (
            f"{srv.get('server_name', self.config.server)} — "
            f"{srv.get('edition', 'SQL Server')} (v{srv.get('product_version', '?')})"
        )
        raw_dbs = payload.get("databases", [])
        if isinstance(raw_dbs, dict):
            raw_dbs = [raw_dbs]

        databases: List[DatabaseInfo] = []
        for item in raw_dbs:
            databases.append(
                DatabaseInfo(
                    name=str(item.get("name", "")),
                    size_mb=float(item.get("size_mb", 0.0)),
                    recovery_model=str(item.get("recovery_model", "FULL")),
                    state_desc=str(item.get("state_desc", "ONLINE")),
                    is_system_db=bool(item.get("is_system_db", False)),
                )
            )
        return server_summary, databases

    def run_single_database_backup(
        self,
        db_info: DatabaseInfo,
        options: BackupOptions,
        timestamp: str,
        cancel_event: threading.Event,
        progress_callback: Optional[Callable[[str, int, str], None]] = None,
    ) -> BackupJobResult:
        """
        Backs up a single database (using 1 or N striped streams depending on size_gb > split_threshold_gb),
        and optionally zips the resulting `.bak` file(s) into the target Zip folder using fastest compression.
        """
        start_time = time.time()
        stream_count = calculate_stream_count(
            size_gb=db_info.size_gb,
            threshold_gb=options.split_threshold_gb,
            max_streams=options.max_streams_per_db,
        )
        db_subfolder, file_paths = build_backup_file_paths(
            target_folder=options.target_folder,
            db_name=db_info.name,
            stream_count=stream_count,
            timestamp=timestamp,
        )
        _, zip_file_path = build_zip_file_path(
            target_folder=options.target_folder,
            db_name=db_info.name,
            timestamp=timestamp,
            custom_zip_folder=options.zip_output_folder,
        )

        if cancel_event.is_set():
            return BackupJobResult(
                db_name=db_info.name,
                success=False,
                stream_count=stream_count,
                file_paths=file_paths,
                error_message="Cancelled before start.",
            )

        # Best-effort local subfolder creation
        try:
            os.makedirs(db_subfolder, exist_ok=True)
        except OSError:
            pass

        create_dir_tsql = build_create_subdir_tsql(db_subfolder)
        backup_tsql_compressed = build_backup_tsql(
            db_name=db_info.name,
            file_paths=file_paths,
            copy_only=options.copy_only,
            compression=options.compression,
            checksum=options.checksum,
            stats_interval=options.stats_interval,
        )
        backup_tsql_uncompressed = build_backup_tsql(
            db_name=db_info.name,
            file_paths=file_paths,
            copy_only=options.copy_only,
            compression=False,
            checksum=options.checksum,
            stats_interval=options.stats_interval,
        )
        verify_tsql = (
            build_verify_tsql(file_paths, checksum=options.checksum)
            if options.verify_only_after
            else ""
        )

        ps_runner = r"""
$ErrorActionPreference = 'Stop'
$connStr = $env:SQL_BG_CONN_STR
$createDirSql = $env:SQL_BG_CREATE_DIR_SQL
$backupSql = $env:SQL_BG_BACKUP_SQL
$backupFallbackSql = $env:SQL_BG_BACKUP_FALLBACK_SQL
$verifySql = $env:SQL_BG_VERIFY_SQL
$allowFallback = ($env:SQL_BG_ALLOW_COMPRESSION_FALLBACK -eq '1')

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

    # 1. Create per-database subfolder on SQL Server host
    $dirCmd = $conn.CreateCommand()
    $dirCmd.CommandTimeout = 60
    $dirCmd.CommandText = $createDirSql
    [void]$dirCmd.ExecuteNonQuery()

    # 2. Execute BACKUP DATABASE (with automatic compression fallback if unsupported by edition)
    $usedCompression = $allowFallback
    try {
        $bakCmd = $conn.CreateCommand()
        $bakCmd.CommandTimeout = 0
        $bakCmd.CommandText = $backupSql
        [void]$bakCmd.ExecuteNonQuery()
    }
    catch {
        $exMsg = $_.Exception.Message
        if ($allowFallback -and ($exMsg -match 'COMPRESSION is not supported' -or $exMsg -match '1844')) {
            [Console]::Out.WriteLine("SQL_WARN: Compression not supported by SQL edition; retrying without COMPRESSION...")
            [Console]::Out.Flush()
            $usedCompression = $false
            $bakCmd2 = $conn.CreateCommand()
            $bakCmd2.CommandTimeout = 0
            $bakCmd2.CommandText = $backupFallbackSql
            [void]$bakCmd2.ExecuteNonQuery()
        }
        else {
            throw
        }
    }

    # 3. Optional RESTORE VERIFYONLY
    if (-not [string]::IsNullOrWhiteSpace($verifySql)) {
        [Console]::Out.WriteLine("SQL_INFO: Verifying backup set (RESTORE VERIFYONLY)...")
        [Console]::Out.Flush()
        $verCmd = $conn.CreateCommand()
        $verCmd.CommandTimeout = 0
        $verCmd.CommandText = $verifySql
        [void]$verCmd.ExecuteNonQuery()
    }

    [Console]::Out.WriteLine("SQL_DONE: USED_COMPRESSION=" + $usedCompression)
    [Console]::Out.Flush()
}
finally {
    if ($conn.State -ne 'Closed') { $conn.Close() }
    $conn.Dispose()
}
"""
        env = os.environ.copy()
        env["SQL_BG_CONN_STR"] = self.config.to_dotnet_connection_string("master")
        env["SQL_BG_CREATE_DIR_SQL"] = create_dir_tsql
        env["SQL_BG_BACKUP_SQL"] = backup_tsql_compressed
        env["SQL_BG_BACKUP_FALLBACK_SQL"] = backup_tsql_uncompressed
        env["SQL_BG_VERIFY_SQL"] = verify_tsql
        env["SQL_BG_ALLOW_COMPRESSION_FALLBACK"] = "1" if options.compression else "0"

        if progress_callback:
            progress_callback(
                db_info.name,
                0,
                f"Starting ({stream_count} {'stream' if stream_count == 1 else 'streams'})...",
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

        used_compression = options.compression
        output_lines: List[str] = []
        pct_regex = re.compile(r"(\d+)\s+percent\s+processed", re.IGNORECASE)

        try:
            assert proc.stdout is not None
            for raw_line in proc.stdout:
                if cancel_event.is_set():
                    proc.terminate()
                    return BackupJobResult(
                        db_name=db_info.name,
                        success=False,
                        stream_count=stream_count,
                        file_paths=file_paths,
                        elapsed_seconds=time.time() - start_time,
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
                        scaled_pct = int(pct * 0.85) if options.zip_after_backup else pct
                        if progress_callback:
                            progress_callback(
                                db_info.name,
                                scaled_pct,
                                f"SQL Backup ({pct}% | {stream_count} {'stream' if stream_count == 1 else 'streams'})",
                            )
                    else:
                        if progress_callback:
                            progress_callback(db_info.name, -1, info_msg)
                elif line.startswith("SQL_WARN:"):
                    warn_msg = line[len("SQL_WARN:") :].strip()
                    if progress_callback:
                        progress_callback(db_info.name, -1, warn_msg)
                elif line.startswith("SQL_DONE:"):
                    if "USED_COMPRESSION=False" in line:
                        used_compression = False

            proc.wait()
        finally:
            if proc.poll() is None:
                proc.kill()

        if proc.returncode != 0:
            elapsed = time.time() - start_time
            err_detail = "\n".join(output_lines[-6:]) if output_lines else f"Exit code {proc.returncode}"
            return BackupJobResult(
                db_name=db_info.name,
                success=False,
                stream_count=stream_count,
                file_paths=file_paths,
                elapsed_seconds=elapsed,
                used_compression=used_compression,
                error_message=err_detail,
            )

        # Step 4: Fast Post-Backup Zip (Speed over Compression Ratio)
        final_zip_path = ""
        zip_elapsed = 0.0
        if options.zip_after_backup:
            try:
                if progress_callback:
                    progress_callback(db_info.name, 90, "Fast-Zipping .bak file(s)...")
                final_zip_path, zip_elapsed = fast_zip_backup_files(
                    file_paths=file_paths,
                    zip_path=zip_file_path,
                    cancel_event=cancel_event,
                    status_callback=lambda msg: (
                        progress_callback(db_info.name, 92, msg) if progress_callback else None
                    ),
                    delete_source_after_zip=options.delete_bak_after_zip,
                    source_subfolder=db_subfolder,
                )
            except Exception as zip_exc:
                elapsed = time.time() - start_time
                return BackupJobResult(
                    db_name=db_info.name,
                    success=False,
                    stream_count=stream_count,
                    file_paths=file_paths,
                    elapsed_seconds=elapsed,
                    used_compression=used_compression,
                    error_message=f"SQL backup succeeded, but Fast-Zip failed: {zip_exc}",
                )

        elapsed = time.time() - start_time
        done_msg = (
            f"Completed in {elapsed:.1f}s ({stream_count} {'stream' if stream_count == 1 else 'streams'}"
            + (f", zipped in {zip_elapsed:.1f}s)" if options.zip_after_backup else ")")
        )
        if progress_callback:
            progress_callback(db_info.name, 100, done_msg)

        return BackupJobResult(
            db_name=db_info.name,
            success=True,
            stream_count=stream_count,
            file_paths=file_paths,
            zip_path=final_zip_path,
            elapsed_seconds=elapsed,
            zip_elapsed_seconds=zip_elapsed,
            used_compression=used_compression,
        )


# ============================================================================
# Tkinter Graphical User Interface
# ============================================================================


class SqlBackupApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1200x860")
        self.root.minsize(1000, 720)

        self.config_file_path = Path(__file__).resolve().parent / CONFIG_FILENAME

        # State variables
        self.var_server = tk.StringVar(value="localhost")
        self.var_auth_mode = tk.StringVar(value="windows")
        self.var_username = tk.StringVar(value="sa")
        self.var_password = tk.StringVar(value="")
        self.var_show_password = tk.BooleanVar(value=False)
        self.var_trust_cert = tk.BooleanVar(value=True)
        self.var_include_system = tk.BooleanVar(value=False)

        default_target = str(Path.home() / "SqlBackups")
        default_zip = str(Path.home() / "SqlBackups" / "Zipped")
        self.var_target_folder = tk.StringVar(value=default_target)
        self.var_zip_after_backup = tk.BooleanVar(value=True)
        self.var_zip_folder = tk.StringVar(value=default_zip)
        self.var_delete_bak_after_zip = tk.BooleanVar(value=False)
        self._user_edited_zip_folder = False

        self.var_max_parallel_dbs = tk.IntVar(value=DEFAULT_MAX_PARALLEL_DBS)
        self.var_split_threshold_gb = tk.DoubleVar(value=DEFAULT_SPLIT_THRESHOLD_GB)
        self.var_max_streams = tk.IntVar(value=DEFAULT_MAX_STREAMS)

        self.var_copy_only = tk.BooleanVar(value=True)
        self.var_compression = tk.BooleanVar(value=True)
        self.var_checksum = tk.BooleanVar(value=True)
        self.var_verify_after = tk.BooleanVar(value=False)

        self.var_filter_text = tk.StringVar(value="")
        self.var_server_status = tk.StringVar(
            value="Not connected — enter SQL Server instance and click 'Connect & Load Databases'"
        )
        self.var_selection_summary = tk.StringVar(
            value="Selected: 0 databases (0.00 GB | 0 total streams)"
        )
        self.var_overall_progress = tk.DoubleVar(value=0.0)
        self.var_progress_label = tk.StringVar(value="Idle")

        # Database inventory & selection tracking
        self.all_databases: List[DatabaseInfo] = []
        self.selected_db_names: set[str] = set()
        self.db_status_map: Dict[str, str] = {}

        # Concurrency & UI event queue
        self.ui_queue: "queue.Queue[Tuple[str, object]]" = queue.Queue()
        self.cancel_event = threading.Event()
        self.is_running_backups = False

        self._configure_styles()
        self._build_ui()
        self._load_saved_config(silent=True)
        self._on_auth_mode_change()
        self._on_zip_toggle_change()
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
            text="SQL Server Parallel Copy-Only Backup & Fast-Zip Utility",
            style="Header.TLabel",
        ).pack(side=tk.LEFT)
        ttk.Label(
            header_frame,
            text="Parallel DB Queue  •  Auto Multi-Stream Striping (> 8 GB)  •  Fastest-Speed Zip Archive",
            style="SubHeader.TLabel",
        ).pack(side=tk.RIGHT)

        # 2. SQL Server Connection Panel
        conn_box = ttk.LabelFrame(main_frame, text=" 1. SQL Server Connection ", padding="8 6 8 6")
        conn_box.pack(fill=tk.X, pady=(0, 6))

        row0 = ttk.Frame(conn_box)
        row0.pack(fill=tk.X, pady=2)

        ttk.Label(row0, text="Server / Instance:").pack(side=tk.LEFT, padx=(0, 4))
        self.ent_server = ttk.Entry(row0, textvariable=self.var_server, width=28)
        self.ent_server.pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(row0, text="Authentication:").pack(side=tk.LEFT, padx=(0, 4))
        self.cmb_auth = ttk.Combobox(
            row0,
            textvariable=self.var_auth_mode,
            values=["windows", "sql"],
            state="readonly",
            width=12,
        )
        self.cmb_auth.pack(side=tk.LEFT, padx=(0, 12))
        self.cmb_auth.bind("<<ComboboxSelected>>", lambda _e: self._on_auth_mode_change())

        ttk.Label(row0, text="SQL Login:").pack(side=tk.LEFT, padx=(0, 4))
        self.ent_user = ttk.Entry(row0, textvariable=self.var_username, width=16)
        self.ent_user.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(row0, text="Password:").pack(side=tk.LEFT, padx=(0, 4))
        self.ent_pass = ttk.Entry(row0, textvariable=self.var_password, width=16, show="*")
        self.ent_pass.pack(side=tk.LEFT, padx=(0, 4))

        ttk.Checkbutton(
            row0,
            text="Show",
            variable=self.var_show_password,
            command=self._toggle_password_visibility,
        ).pack(side=tk.LEFT, padx=(0, 12))

        self.btn_connect = ttk.Button(
            row0,
            text="Connect & Load Databases",
            command=self.connect_and_load_databases,
        )
        self.btn_connect.pack(side=tk.LEFT, padx=(4, 6))

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
        ).pack(side=tk.LEFT, padx=(0, 14))
        ttk.Checkbutton(
            row1,
            text="Include System Databases (master, model, msdb)",
            variable=self.var_include_system,
            command=self._on_include_system_changed,
        ).pack(side=tk.LEFT, padx=(0, 18))

        ttk.Label(row1, textvariable=self.var_server_status, foreground="#1e56a0").pack(
            side=tk.LEFT, fill=tk.X, expand=True
        )

        # 3. Backup Target, Parallelism & Fast-Zip Settings
        settings_box = ttk.LabelFrame(
            main_frame,
            text=" 2. Target Folders, Parallelism, Multi-Stream Striping & Fast-Zip Options ",
            padding="8 6 8 6",
        )
        settings_box.pack(fill=tk.X, pady=(0, 6))

        srow0 = ttk.Frame(settings_box)
        srow0.pack(fill=tk.X, pady=2)
        ttk.Label(srow0, text="Target Backup Folder (.bak):", width=24, anchor=tk.W).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        self.ent_target = ttk.Entry(srow0, textvariable=self.var_target_folder)
        self.ent_target.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        self.var_target_folder.trace_add("write", lambda *_: self._on_target_folder_changed())

        ttk.Button(srow0, text="Browse Folder...", command=self.browse_target_folder).pack(
            side=tk.LEFT
        )

        # Fast-Zip Row
        srow_zip = ttk.Frame(settings_box)
        srow_zip.pack(fill=tk.X, pady=3)
        ttk.Checkbutton(
            srow_zip,
            text="Zip Backups into Folder (Fastest Speed Mode):",
            variable=self.var_zip_after_backup,
            command=self._on_zip_toggle_change,
        ).pack(side=tk.LEFT, padx=(0, 6))

        self.ent_zip_folder = ttk.Entry(srow_zip, textvariable=self.var_zip_folder)
        self.ent_zip_folder.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        self.ent_zip_folder.bind("<KeyRelease>", lambda _e: self._on_zip_folder_key())

        self.btn_browse_zip = ttk.Button(
            srow_zip, text="Browse Zip Folder...", command=self.browse_zip_folder
        )
        self.btn_browse_zip.pack(side=tk.LEFT, padx=(0, 10))

        self.chk_del_bak = ttk.Checkbutton(
            srow_zip,
            text="Delete raw .bak files after Zip",
            variable=self.var_delete_bak_after_zip,
            command=self._on_settings_changed,
        )
        self.chk_del_bak.pack(side=tk.LEFT)

        srow1 = ttk.Frame(settings_box)
        srow1.pack(fill=tk.X, pady=(6, 2))

        ttk.Label(srow1, text="Max Parallel Databases:").pack(side=tk.LEFT, padx=(0, 4))
        spn_parallel = ttk.Spinbox(
            srow1,
            from_=1,
            to=16,
            textvariable=self.var_max_parallel_dbs,
            width=5,
            command=self._on_settings_changed,
        )
        spn_parallel.pack(side=tk.LEFT, padx=(0, 16))

        ttk.Label(srow1, text="Split Threshold (GB):").pack(side=tk.LEFT, padx=(0, 4))
        spn_thresh = ttk.Spinbox(
            srow1,
            from_=1.0,
            to=512.0,
            increment=1.0,
            textvariable=self.var_split_threshold_gb,
            width=6,
            command=self._on_settings_changed,
        )
        spn_thresh.pack(side=tk.LEFT, padx=(0, 16))
        spn_thresh.bind("<KeyRelease>", lambda _e: self._on_settings_changed())

        ttk.Label(srow1, text="Max Streams per Large DB:").pack(side=tk.LEFT, padx=(0, 4))
        spn_streams = ttk.Spinbox(
            srow1,
            from_=1,
            to=64,
            textvariable=self.var_max_streams,
            width=5,
            command=self._on_settings_changed,
        )
        spn_streams.pack(side=tk.LEFT, padx=(0, 18))
        spn_streams.bind("<KeyRelease>", lambda _e: self._on_settings_changed())

        chk_copy_only = ttk.Checkbutton(
            srow1,
            text="COPY_ONLY (Always On)",
            variable=self.var_copy_only,
            state="disabled",
        )
        chk_copy_only.pack(side=tk.LEFT, padx=(0, 12))

        ttk.Checkbutton(
            srow1,
            text="BACKUP COMPRESSION",
            variable=self.var_compression,
            command=self._on_settings_changed,
        ).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Checkbutton(
            srow1,
            text="CHECKSUM",
            variable=self.var_checksum,
            command=self._on_settings_changed,
        ).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Checkbutton(
            srow1,
            text="Verify After Backup (RESTORE VERIFYONLY)",
            variable=self.var_verify_after,
            command=self._on_settings_changed,
        ).pack(side=tk.LEFT)

        # 4. PanedWindow: Database Selection Table (Top) + T-SQL Preview & Execution Log (Bottom)
        paned = ttk.PanedWindow(main_frame, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True)

        # 4A. Database Selection Frame
        db_box = ttk.LabelFrame(
            paned,
            text=" 3. Select Databases to Backup (Click or Spacebar to toggle selection) ",
            padding="8 6 8 6",
        )
        paned.add(db_box, weight=3)

        toolbar = ttk.Frame(db_box)
        toolbar.pack(fill=tk.X, pady=(0, 6))

        ttk.Button(toolbar, text="Select All", command=self.select_all_databases).pack(
            side=tk.LEFT, padx=(0, 4)
        )
        ttk.Button(toolbar, text="Deselect All", command=self.deselect_all_databases).pack(
            side=tk.LEFT, padx=(0, 4)
        )
        ttk.Button(
            toolbar,
            text="Select Large (> Threshold)",
            command=self.select_large_databases,
        ).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(toolbar, text="Invert Selection", command=self.invert_selection).pack(
            side=tk.LEFT, padx=(0, 14)
        )

        ttk.Label(toolbar, text="Filter DBs:").pack(side=tk.LEFT, padx=(0, 4))
        ent_filter = ttk.Entry(toolbar, textvariable=self.var_filter_text, width=22)
        ent_filter.pack(side=tk.LEFT, padx=(0, 6))
        self.var_filter_text.trace_add("write", lambda *_: self.refresh_database_tree())

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
            "db_name",
            "size_gb",
            "streams",
            "recovery",
            "status",
            "subfolder",
        )
        self.tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show="headings",
            selectmode="browse",
        )
        self.tree.heading("selected", text="Backup?")
        self.tree.heading("db_name", text="Database Name")
        self.tree.heading("size_gb", text="Size (GB)")
        self.tree.heading("streams", text="Streams (Stripes)")
        self.tree.heading("recovery", text="Recovery Model")
        self.tree.heading("status", text="Status / Progress")
        self.tree.heading("subfolder", text="Target .bak / .zip Output")

        self.tree.column("selected", width=70, anchor=tk.CENTER, stretch=False)
        self.tree.column("db_name", width=210, anchor=tk.W)
        self.tree.column("size_gb", width=105, anchor=tk.E, stretch=False)
        self.tree.column("streams", width=140, anchor=tk.CENTER, stretch=False)
        self.tree.column("recovery", width=110, anchor=tk.CENTER, stretch=False)
        self.tree.column("status", width=220, anchor=tk.W)
        self.tree.column("subfolder", width=290, anchor=tk.W)

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
        self.tree.bind("<<TreeviewSelect>>", lambda _e: self.update_tsql_preview())

        # 4B. Bottom Panel: T-SQL Preview + Action Controls + Execution Log
        bottom_box = ttk.Frame(paned)
        paned.add(bottom_box, weight=2)

        split_bottom = ttk.PanedWindow(bottom_box, orient=tk.HORIZONTAL)
        split_bottom.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        preview_frame = ttk.LabelFrame(
            split_bottom,
            text=" Live T-SQL & Fast-Zip Pipeline Preview ",
            padding="6 4 6 4",
        )
        split_bottom.add(preview_frame, weight=1)

        preview_btn_bar = ttk.Frame(preview_frame)
        preview_btn_bar.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(
            preview_btn_bar,
            text="Shows exact T-SQL & Fast-Zip target for highlighted/selected DB(s):",
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
            text=" Parallel Backup & Fast-Zip Execution Log ",
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

        # Action Bar at very bottom
        action_bar = ttk.Frame(bottom_box)
        action_bar.pack(fill=tk.X)

        self.btn_start = ttk.Button(
            action_bar,
            text="▶  Start Parallel Copy-Only Backups",
            command=self.start_parallel_backups,
        )
        self.btn_start.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_cancel = ttk.Button(
            action_bar,
            text="■  Cancel Backups",
            command=self.cancel_backups,
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

    def _on_zip_toggle_change(self) -> None:
        enabled = bool(self.var_zip_after_backup.get())
        state = "normal" if enabled else "disabled"
        self.ent_zip_folder.configure(state=state)
        self.btn_browse_zip.configure(state=state)
        self.chk_del_bak.configure(state=state)
        self._on_settings_changed()

    def _on_target_folder_changed(self) -> None:
        if not self._user_edited_zip_folder:
            base = self.var_target_folder.get().strip()
            if base:
                self.var_zip_folder.set(join_windows_or_native_path(base, "Zipped"))
        self._on_settings_changed()

    def _on_zip_folder_key(self) -> None:
        self._user_edited_zip_folder = True
        self._on_settings_changed()

    def _on_include_system_changed(self) -> None:
        if self.all_databases:
            self.connect_and_load_databases()

    def browse_target_folder(self) -> None:
        initial = self.var_target_folder.get().strip() or str(Path.home())
        chosen = filedialog.askdirectory(
            parent=self.root,
            title="Select Target Backup Folder (.bak)",
            initialdir=initial if os.path.isdir(initial) else str(Path.home()),
        )
        if chosen:
            if os.name == "nt":
                chosen = chosen.replace("/", "\\")
            self.var_target_folder.set(chosen)

    def browse_zip_folder(self) -> None:
        initial = self.var_zip_folder.get().strip() or self.var_target_folder.get().strip() or str(Path.home())
        chosen = filedialog.askdirectory(
            parent=self.root,
            title="Select Target Zip Archive Folder",
            initialdir=initial if os.path.isdir(initial) else str(Path.home()),
        )
        if chosen:
            if os.name == "nt":
                chosen = chosen.replace("/", "\\")
            self._user_edited_zip_folder = True
            self.var_zip_folder.set(chosen)
            self._on_settings_changed()

    def _get_safe_threshold_and_streams(self) -> Tuple[float, int, int]:
        try:
            thresh = float(self.var_split_threshold_gb.get())
            if thresh <= 0:
                thresh = DEFAULT_SPLIT_THRESHOLD_GB
        except (tk.TclError, ValueError):
            thresh = DEFAULT_SPLIT_THRESHOLD_GB

        try:
            max_streams = int(self.var_max_streams.get())
            if max_streams < 1:
                max_streams = 1
        except (tk.TclError, ValueError):
            max_streams = DEFAULT_MAX_STREAMS

        try:
            max_parallel = int(self.var_max_parallel_dbs.get())
            if max_parallel < 1:
                max_parallel = 1
        except (tk.TclError, ValueError):
            max_parallel = DEFAULT_MAX_PARALLEL_DBS

        return thresh, max_streams, max_parallel

    def _on_settings_changed(self) -> None:
        self.refresh_database_tree()
        self.update_tsql_preview()

    def _build_current_connection_config(self) -> SqlConnectionConfig:
        return SqlConnectionConfig(
            server=self.var_server.get().strip() or "localhost",
            auth_mode=self.var_auth_mode.get().strip() or "windows",
            username=self.var_username.get().strip(),
            password=self.var_password.get(),
            trust_server_certificate=bool(self.var_trust_cert.get()),
        )

    def _build_current_backup_options(self) -> BackupOptions:
        thresh, max_streams, max_parallel = self._get_safe_threshold_and_streams()
        target_folder = self.var_target_folder.get().strip() or r"C:\SqlBackups"
        zip_folder = resolve_zip_output_folder(target_folder, self.var_zip_folder.get())
        return BackupOptions(
            target_folder=target_folder,
            split_threshold_gb=thresh,
            max_streams_per_db=max_streams,
            max_parallel_dbs=max_parallel,
            copy_only=True,
            compression=bool(self.var_compression.get()),
            checksum=bool(self.var_checksum.get()),
            verify_only_after=bool(self.var_verify_after.get()),
            zip_after_backup=bool(self.var_zip_after_backup.get()),
            zip_output_folder=zip_folder,
            delete_bak_after_zip=bool(self.var_delete_bak_after_zip.get()),
        )

    # ========================================================================
    # Config Load / Save
    # ========================================================================

    def save_config(self) -> None:
        thresh, max_streams, max_parallel = self._get_safe_threshold_and_streams()
        data = {
            "server": self.var_server.get().strip(),
            "auth_mode": self.var_auth_mode.get().strip(),
            "username": self.var_username.get().strip(),
            "trust_server_certificate": bool(self.var_trust_cert.get()),
            "include_system_dbs": bool(self.var_include_system.get()),
            "target_folder": self.var_target_folder.get().strip(),
            "zip_after_backup": bool(self.var_zip_after_backup.get()),
            "zip_output_folder": self.var_zip_folder.get().strip(),
            "delete_bak_after_zip": bool(self.var_delete_bak_after_zip.get()),
            "max_parallel_dbs": max_parallel,
            "split_threshold_gb": thresh,
            "max_streams_per_db": max_streams,
            "compression": bool(self.var_compression.get()),
            "checksum": bool(self.var_checksum.get()),
            "verify_only_after": bool(self.var_verify_after.get()),
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
            self.var_include_system.set(bool(data.get("include_system_dbs", False)))
            if data.get("target_folder"):
                self.var_target_folder.set(data["target_folder"])
            if "zip_after_backup" in data:
                self.var_zip_after_backup.set(bool(data["zip_after_backup"]))
            if data.get("zip_output_folder"):
                self._user_edited_zip_folder = True
                self.var_zip_folder.set(data["zip_output_folder"])
            if "delete_bak_after_zip" in data:
                self.var_delete_bak_after_zip.set(bool(data["delete_bak_after_zip"]))
            self.var_max_parallel_dbs.set(int(data.get("max_parallel_dbs", DEFAULT_MAX_PARALLEL_DBS)))
            self.var_split_threshold_gb.set(
                float(data.get("split_threshold_gb", DEFAULT_SPLIT_THRESHOLD_GB))
            )
            self.var_max_streams.set(int(data.get("max_streams_per_db", DEFAULT_MAX_STREAMS)))
            self.var_compression.set(bool(data.get("compression", True)))
            self.var_checksum.set(bool(data.get("checksum", True)))
            self.var_verify_after.set(bool(data.get("verify_only_after", False)))
            self._on_auth_mode_change()
            self._on_zip_toggle_change()
            if not silent:
                self.log(f"Loaded configuration from {self.config_file_path.name}")
        except Exception as exc:
            if not silent:
                messagebox.showerror("Load Config Failed", str(exc), parent=self.root)

    # ========================================================================
    # Database Discovery & Selection Table
    # ========================================================================

    def connect_and_load_databases(self) -> None:
        if self.is_running_backups:
            return
        cfg = self._build_current_connection_config()
        include_sys = bool(self.var_include_system.get())
        self.btn_connect.configure(state="disabled")
        self.var_server_status.set(f"Connecting to {cfg.server}...")
        self.log(f"Connecting to SQL Server '{cfg.server}' ({cfg.auth_mode} auth)...")

        def worker() -> None:
            try:
                engine = SqlServerEngine(cfg)
                server_summary, dbs = engine.fetch_databases(include_system_dbs=include_sys)
                self.ui_queue.put(("DB_LIST_LOADED", (server_summary, dbs)))
            except Exception as exc:
                self.ui_queue.put(("DB_LIST_ERROR", str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def refresh_database_tree(self) -> None:
        focused_iid = self.tree.focus()
        for item in self.tree.get_children():
            self.tree.delete(item)

        options = self._build_current_backup_options()
        filter_q = self.var_filter_text.get().strip().lower()

        for db in self.all_databases:
            if filter_q and filter_q not in db.name.lower():
                continue

            is_sel = db.name in self.selected_db_names
            check_mark = "[✓] Yes" if is_sel else "[  ] No"
            streams = calculate_stream_count(
                db.size_gb, options.split_threshold_gb, options.max_streams_per_db
            )
            if streams == 1:
                stream_label = "1 stream"
            else:
                stream_label = f"⚡ {streams} striped streams"

            subfolder, _ = build_backup_file_paths(
                options.target_folder, db.name, streams, "YYYYMMDD"
            )
            if options.zip_after_backup:
                _, zip_dest = build_zip_file_path(
                    options.target_folder,
                    db.name,
                    "YYYYMMDD",
                    options.zip_output_folder,
                )
                dest_display = f"{subfolder}  ➜  {zip_dest}"
            else:
                dest_display = subfolder

            status_txt = self.db_status_map.get(db.name, "Ready")

            self.tree.insert(
                "",
                tk.END,
                iid=db.name,
                values=(
                    check_mark,
                    db.name,
                    f"{db.size_gb:,.2f} GB",
                    stream_label,
                    db.recovery_model,
                    status_txt,
                    dest_display,
                ),
            )

        if focused_iid and self.tree.exists(focused_iid):
            self.tree.focus(focused_iid)
            self.tree.selection_set(focused_iid)

        self._update_selection_summary()

    def _update_selection_summary(self) -> None:
        thresh, max_streams, _ = self._get_safe_threshold_and_streams()
        selected_dbs = [d for d in self.all_databases if d.name in self.selected_db_names]
        total_gb = sum(d.size_gb for d in selected_dbs)
        total_streams = sum(
            calculate_stream_count(d.size_gb, thresh, max_streams) for d in selected_dbs
        )
        large_count = sum(1 for d in selected_dbs if d.size_gb > thresh)
        zip_badge = " + Fast-Zip" if self.var_zip_after_backup.get() else ""
        self.var_selection_summary.set(
            f"Selected: {len(selected_dbs)} DBs ({total_gb:,.2f} GB  |  "
            f"{large_count} > {thresh:g} GB  |  {total_streams} streams{zip_badge})"
        )

    def _on_tree_click(self, event: tk.Event) -> None:
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        if row_id and col == "#1":
            self._toggle_db_selection(row_id)

    def _on_tree_double_click(self, event: tk.Event) -> None:
        row_id = self.tree.identify_row(event.y)
        if row_id:
            self._toggle_db_selection(row_id)

    def _on_tree_space(self, _event: tk.Event) -> None:
        focused = self.tree.focus()
        if focused:
            self._toggle_db_selection(focused)

    def _toggle_db_selection(self, db_name: str) -> None:
        if db_name in self.selected_db_names:
            self.selected_db_names.remove(db_name)
        else:
            self.selected_db_names.add(db_name)
        self.refresh_database_tree()
        self.update_tsql_preview()

    def select_all_databases(self) -> None:
        filter_q = self.var_filter_text.get().strip().lower()
        for db in self.all_databases:
            if not filter_q or filter_q in db.name.lower():
                self.selected_db_names.add(db.name)
        self.refresh_database_tree()
        self.update_tsql_preview()

    def deselect_all_databases(self) -> None:
        self.selected_db_names.clear()
        self.refresh_database_tree()
        self.update_tsql_preview()

    def select_large_databases(self) -> None:
        thresh, _, _ = self._get_safe_threshold_and_streams()
        self.selected_db_names = {d.name for d in self.all_databases if d.size_gb > thresh}
        self.refresh_database_tree()
        self.update_tsql_preview()

    def invert_selection(self) -> None:
        all_names = {d.name for d in self.all_databases}
        self.selected_db_names = all_names - self.selected_db_names
        self.refresh_database_tree()
        self.update_tsql_preview()

    # ========================================================================
    # Live T-SQL Command & Fast-Zip Preview
    # ========================================================================

    def update_tsql_preview(self) -> None:
        options = self._build_current_backup_options()
        timestamp_preview = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

        focused = self.tree.focus()
        db_lookup = {d.name: d for d in self.all_databases}

        preview_dbs: List[DatabaseInfo] = []
        if focused and focused in db_lookup:
            preview_dbs.append(db_lookup[focused])
        for d in self.all_databases:
            if d.name in self.selected_db_names and d not in preview_dbs:
                preview_dbs.append(d)

        zip_mode_note = (
            f"-- Post-Backup Fast-Zip: ENABLED (Fastest Speed / Level 1 -> {options.zip_output_folder})\n"
            if options.zip_after_backup
            else "-- Post-Backup Fast-Zip: DISABLED\n"
        )

        if not preview_dbs:
            sample_small = DatabaseInfo(name="AppConfigDB", size_mb=2048.0)
            sample_large = DatabaseInfo(name="EnterpriseWarehouseDB", size_mb=28672.0)  # 28 GB
            preview_dbs = [sample_small, sample_large]
            header_comment = (
                "-- [Preview Example] Connect & select databases above to see their exact T-SQL.\n"
                f"-- Parallel Workers: {options.max_parallel_dbs} concurrent DBs | "
                f"Split Threshold: > {options.split_threshold_gb:g} GB (up to {options.max_streams_per_db} streams)\n"
                + zip_mode_note
                + "\n"
            )
        else:
            header_comment = (
                f"-- Parallel Execution Plan: Up to {options.max_parallel_dbs} databases backed up concurrently\n"
                f"-- Split Threshold: > {options.split_threshold_gb:g} GB split into up to {options.max_streams_per_db} parallel striped streams\n"
                + zip_mode_note
                + "\n"
            )

        blocks: List[str] = [header_comment]
        for db in preview_dbs[:6]:
            streams = calculate_stream_count(
                db.size_gb, options.split_threshold_gb, options.max_streams_per_db
            )
            subfolder, paths = build_backup_file_paths(
                options.target_folder, db.name, streams, timestamp_preview
            )
            mkdir_sql = build_create_subdir_tsql(subfolder)
            bak_sql = build_backup_tsql(
                db_name=db.name,
                file_paths=paths,
                copy_only=True,
                compression=options.compression,
                checksum=options.checksum,
                stats_interval=options.stats_interval,
            )
            block = (
                f"-- Database: [{db.name}] ({db.size_gb:,.2f} GB) -> {streams} "
                f"{'stream' if streams == 1 else 'parallel striped streams'}\n"
                f"{mkdir_sql}\n"
                f"{bak_sql}\n"
            )
            if options.verify_only_after:
                block += build_verify_tsql(paths, checksum=options.checksum) + "\n"
            if options.zip_after_backup:
                _, zip_dest = build_zip_file_path(
                    options.target_folder,
                    db.name,
                    timestamp_preview,
                    options.zip_output_folder,
                )
                del_note = " (then delete raw .bak)" if options.delete_bak_after_zip else ""
                block += f"-- Fast-Zip Output (Level 1 / Fastest): {zip_dest}{del_note}\n"
            blocks.append(block)

        if len(preview_dbs) > 6:
            blocks.append(f"-- ... plus {len(preview_dbs) - 6} more selected database(s) ...\n")

        self.txt_tsql.delete("1.0", tk.END)
        self.txt_tsql.insert("1.0", "\n".join(blocks))

    def copy_tsql_to_clipboard(self) -> None:
        tsql_text = self.txt_tsql.get("1.0", tk.END).strip()
        if tsql_text:
            self.root.clipboard_clear()
            self.root.clipboard_append(tsql_text)
            self.log("Copied T-SQL preview to clipboard.")

    # ========================================================================
    # Parallel Backup Execution Coordinator
    # ========================================================================

    def start_parallel_backups(self) -> None:
        if self.is_running_backups:
            return

        selected_dbs = [d for d in self.all_databases if d.name in self.selected_db_names]
        if not selected_dbs:
            messagebox.showwarning(
                "No Databases Selected",
                "Please select at least one database from the list before starting backups.",
                parent=self.root,
            )
            return

        options = self._build_current_backup_options()
        if not options.target_folder:
            messagebox.showwarning(
                "Target Folder Required",
                "Please specify a target backup folder.",
                parent=self.root,
            )
            return

        cfg = self._build_current_connection_config()
        self.is_running_backups = True
        self.cancel_event.clear()
        self.btn_start.configure(state="disabled")
        self.btn_connect.configure(state="disabled")
        self.btn_cancel.configure(state="normal")
        self.var_overall_progress.set(0.0)

        batch_timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        for d in selected_dbs:
            self.db_status_map[d.name] = "Queued"
        self.refresh_database_tree()

        total_count = len(selected_dbs)
        zip_info = (
            f", Fast-Zip -> {options.zip_output_folder}" if options.zip_after_backup else ""
        )
        self.log(
            f"Starting parallel COPY_ONLY backup batch for {total_count} database(s) "
            f"(Max Parallel DBs: {options.max_parallel_dbs}, Split Threshold: > {options.split_threshold_gb:g} GB{zip_info})..."
        )

        def coordinator() -> None:
            engine = SqlServerEngine(cfg)
            results: List[BackupJobResult] = []
            db_progress_pct: Dict[str, int] = {d.name: 0 for d in selected_dbs}
            lock = threading.Lock()

            def on_db_progress(db_name: str, pct: int, status_msg: str) -> None:
                with lock:
                    if pct >= 0:
                        db_progress_pct[db_name] = pct
                    avg_pct = sum(db_progress_pct.values()) / max(1, len(selected_dbs))
                self.ui_queue.put(("DB_PROGRESS", (db_name, pct, status_msg, avg_pct)))

            with concurrent.futures.ThreadPoolExecutor(
                max_workers=options.max_parallel_dbs
            ) as executor:
                future_to_db = {
                    executor.submit(
                        engine.run_single_database_backup,
                        db,
                        options,
                        batch_timestamp,
                        self.cancel_event,
                        on_db_progress,
                    ): db
                    for db in selected_dbs
                }
                for future in concurrent.futures.as_completed(future_to_db):
                    db = future_to_db[future]
                    try:
                        res = future.result()
                    except Exception as exc:
                        res = BackupJobResult(
                            db_name=db.name,
                            success=False,
                            stream_count=1,
                            error_message=str(exc),
                        )
                    results.append(res)
                    with lock:
                        db_progress_pct[db.name] = 100
                        avg_pct = sum(db_progress_pct.values()) / max(1, len(selected_dbs))
                    self.ui_queue.put(("DB_FINISHED", (res, avg_pct)))

            self.ui_queue.put(("BATCH_FINISHED", results))

        threading.Thread(target=coordinator, daemon=True).start()

    def cancel_backups(self) -> None:
        if self.is_running_backups:
            self.cancel_event.set()
            self.log("Cancellation requested — stopping active and queued backup workers...")
            self.btn_cancel.configure(state="disabled")

    def _poll_ui_queue(self) -> None:
        try:
            while True:
                msg_type, payload = self.ui_queue.get_nowait()

                if msg_type == "DB_LIST_LOADED":
                    server_summary, dbs = payload  # type: ignore[misc]
                    self.btn_connect.configure(state="normal")
                    self.all_databases = dbs
                    self.selected_db_names = {d.name for d in dbs if not d.is_system_db}
                    self.db_status_map = {d.name: "Ready" for d in dbs}
                    self.var_server_status.set(
                        f"Connected: {server_summary} ({len(dbs)} online DBs)"
                    )
                    self.log(
                        f"Connected to {server_summary}. Loaded {len(dbs)} online database(s)."
                    )
                    self.refresh_database_tree()
                    self.update_tsql_preview()

                elif msg_type == "DB_LIST_ERROR":
                    err_str = str(payload)
                    self.btn_connect.configure(state="normal")
                    self.var_server_status.set(
                        "Connection failed — check Server/Instance and credentials"
                    )
                    self.log(f"ERROR connecting to SQL Server: {err_str}")
                    messagebox.showerror("SQL Server Connection Error", err_str, parent=self.root)

                elif msg_type == "DB_PROGRESS":
                    db_name, pct, status_msg, avg_pct = payload  # type: ignore[misc]
                    self.db_status_map[db_name] = status_msg
                    if self.tree.exists(db_name):
                        vals = list(self.tree.item(db_name, "values"))
                        vals[5] = status_msg
                        self.tree.item(db_name, values=vals)
                    self.var_overall_progress.set(avg_pct)
                    self.var_progress_label.set(f"Processing... ({avg_pct:.0f}%)")
                    if pct in (0, 90, -1):
                        self.log(f"[{db_name}] {status_msg}")

                elif msg_type == "DB_FINISHED":
                    res, avg_pct = payload  # type: ignore[misc]
                    self.var_overall_progress.set(avg_pct)
                    if res.success:
                        zip_suffix = (
                            f", zipped in {res.zip_elapsed_seconds:.1f}s"
                            if res.zip_path
                            else ""
                        )
                        status_str = (
                            f"✓ Done in {res.elapsed_seconds:.1f}s "
                            f"({res.stream_count} {'stream' if res.stream_count == 1 else 'streams'}{zip_suffix})"
                        )
                        self.db_status_map[res.db_name] = status_str
                        dest_target = res.zip_path if res.zip_path else res.file_paths[0]
                        self.log(
                            f"✓ [{res.db_name}] Completed in {res.elapsed_seconds:.1f}s "
                            f"({res.stream_count} stream(s){zip_suffix}) -> {dest_target}"
                        )
                    else:
                        status_str = f"✗ Failed: {res.error_message[:60]}"
                        self.db_status_map[res.db_name] = status_str
                        self.log(f"✗ [{res.db_name}] FAILED: {res.error_message}")

                    if self.tree.exists(res.db_name):
                        vals = list(self.tree.item(res.db_name, "values"))
                        vals[5] = status_str
                        self.tree.item(res.db_name, values=vals)

                elif msg_type == "BATCH_FINISHED":
                    results: List[BackupJobResult] = payload  # type: ignore[assignment]
                    self.is_running_backups = False
                    self.btn_start.configure(state="normal")
                    self.btn_connect.configure(state="normal")
                    self.btn_cancel.configure(state="disabled")

                    succeeded = sum(1 for r in results if r.success)
                    failed = len(results) - succeeded
                    summary = f"Batch finished: {succeeded} succeeded, {failed} failed."
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
    SqlBackupApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
