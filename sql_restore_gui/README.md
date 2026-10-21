# SQL Server Parallel Striped Database Restore GUI (`sql_restore_gui`)

A self-contained graphical utility for restoring a batch of single-file or multi-stream striped SQL Server databases (`.bak` sets and `.zip` archives) **as quickly as possible**.

Designed as the direct companion to [`sql_backup_gui`](../sql_backup_gui/README.md) and built to run seamlessly on **both user desktops (with Python)** and **Windows Servers (without Python)**.

## Files Included

| File | Purpose |
| :--- | :--- |
| `Launch_SQL_Restore.bat` | **Universal Launcher** — Automatically runs `sql_restore_gui.py` when Python is available, or falls back to `sql_restore_gui.ps1` on Windows Servers without Python. |
| `sql_restore_gui.py` | **Python Desktop GUI (`tkinter`)** — Zero mandatory third-party `pip` dependencies on Windows (uses a built-in `.NET System.Data.SqlClient` bridge and `7z.exe` / 16 MB buffered `zipfile`). |
| `sql_restore_gui.ps1` | **Native Windows Server GUI (PowerShell / .NET WinForms)** — Runs on Windows Servers with **no Python installation required**, using PowerShell `RunspacePool`, `.NET SqlClient`, and `7z.exe` / `System.IO.Compression`. |
| `build_standalone_exe.bat` | Optional helper script that packages `sql_restore_gui.py` into a single-file `dist\sql_restore_gui.exe` via PyInstaller. |
| `test_sql_restore_gui.py` | Unit tests for stripe filename & `.zip` discovery, latest-backup filtering, `WITH MOVE` path generation, high-speed `RESTORE DATABASE` T-SQL formatting, LPT scheduling, fast unzip + staging cleanup, and GUI initialization. |

## Key Features

1. **Hybrid Discovery of Striped `.bak` Sets & `.zip` Archives**:
   - Recursively scans a source folder for:
     - Single `.bak` files (`<DbName>_<Timestamp>.bak`)
     - Multi-stream striped `.bak` sets (`<DbName>_<Timestamp>_stripe01of04.bak` ... `stripe04of04.bak`, `_part1of4.bak`, etc.)
     - Fast-Zipped `.zip` archives produced by `sql_backup_gui` (`<ZipFolder>\<DbName>_<Timestamp>.zip`)
   - Automatically groups stripes together, checks stripe completeness (`found_stripes / expected_stripes`), and defaults to **Latest backup per DB only** so historical backups in the same folder are automatically filtered out.
   - Includes one-click **Verify SQL Headers** (`RESTORE HEADERONLY` & `RESTORE FILELISTONLY`) to validate backup headers and logical file layouts prior to execution.

2. **Maximum Restore Speed ("As Quickly As Possible")**:
   - **Largest-Database-First (LPT) Queue Scheduling**: Orders the selected restore batch by total backup size descending so multi-stripe databases begin immediately on parallel worker threads instead of holding up the tail end of the batch.
   - **Parallel Database Restore Workers**: Restores multiple databases concurrently (default **3 parallel database workers**, configurable from `1` to `16`).
   - **Tuned High-Speed I/O (`MAXTRANSFERSIZE` & `BUFFERCOUNT`)**:
     - Configures `MAXTRANSFERSIZE = 4194304` (4 MB, the maximum supported by SQL Server) and auto-scales `BUFFERCOUNT` (`min(128, max(24, stream_count * 8))` or user-specified).
     - Includes **automatic retry fallback** to default SQL Server buffers if the target server experiences contiguous virtual memory pressure (`SQL Errors 3013 / 3204 / 701 / 802`).
   - **Instant File Initialization (IFI) Detection**: Automatically checks `sys.dm_server_services` upon connecting and displays whether `instant_file_initialization_enabled` is active on the SQL Server service account.
   - **Fast Multi-Threaded `.zip` Extraction & Auto-Cleanup**: Extracts `.zip` backup archives into `<SourceFolder>\_RestoreStaging\` using `7z.exe e -y -mmt=on` (or 16 MB streaming buffers) and automatically deletes staged `.bak` files once each database restore completes.

3. **Smart File Relocation (`WITH MOVE`), Overwrite & Target DB Renaming**:
   - Queries `SERVERPROPERTY('InstanceDefaultDataPath')` and `SERVERPROPERTY('InstanceDefaultLogPath')` upon connection to auto-populate the **Target Data Folder (`.mdf`/`.ndf`)** and **Target Log Folder (`.ldf`)**.
   - Generates `WITH MOVE N'<LogicalName>' TO N'<TargetFolder>\<PhysicalFile>'` clauses from `RESTORE FILELISTONLY`.
   - Supports **Target DB Prefix / Suffix** (e.g., `RESTORED_` or `_UAT`) as well as per-database target rename (double-click Target DB column or right-click a row) while automatically adapting physical `.mdf`/`.ndf`/`.ldf` filenames to prevent file collisions.
   - Defaults to `WITH REPLACE` and **Kill Active Connections (`ALTER DATABASE [...] SET SINGLE_USER WITH ROLLBACK IMMEDIATE`)**, returning the database to `MULTI_USER` once restored.

### Generated High-Speed T-SQL Example

```sql
IF DB_ID(N'EnterpriseWarehouseDB') IS NOT NULL AND DATABASEPROPERTYEX(N'EnterpriseWarehouseDB', 'Status') = 'ONLINE'
BEGIN
    ALTER DATABASE [EnterpriseWarehouseDB] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
END;

RESTORE DATABASE [EnterpriseWarehouseDB]
FROM
    DISK = N'D:\SqlBackups\EnterpriseWarehouseDB\EnterpriseWarehouseDB_20261021_100000_stripe01of04.bak',
    DISK = N'D:\SqlBackups\EnterpriseWarehouseDB\EnterpriseWarehouseDB_20261021_100000_stripe02of04.bak',
    DISK = N'D:\SqlBackups\EnterpriseWarehouseDB\EnterpriseWarehouseDB_20261021_100000_stripe03of04.bak',
    DISK = N'D:\SqlBackups\EnterpriseWarehouseDB\EnterpriseWarehouseDB_20261021_100000_stripe04of04.bak'
WITH
    MOVE N'EnterpriseWarehouseDB' TO N'E:\SqlData\EnterpriseWarehouseDB.mdf',
    MOVE N'EnterpriseWarehouseDB_log' TO N'F:\SqlLogs\EnterpriseWarehouseDB_log.ldf',
    REPLACE,
    RECOVERY,
    MAXTRANSFERSIZE = 4194304,
    BUFFERCOUNT = 32,
    STATS = 5;

IF DB_ID(N'EnterpriseWarehouseDB') IS NOT NULL AND DATABASEPROPERTYEX(N'EnterpriseWarehouseDB', 'Status') = 'ONLINE'
BEGIN
    ALTER DATABASE [EnterpriseWarehouseDB] SET MULTI_USER;
END;
```

## Quick Start

- **Any Windows Machine (Desktop or Server)**:
  Double-click `Launch_SQL_Restore.bat`
- **Run Python Edition directly**:
  ```powershell
  python sql_restore_gui\sql_restore_gui.py
  ```
- **Run Native Server PowerShell Edition directly (No Python needed)**:
  ```powershell
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\sql_restore_gui\sql_restore_gui.ps1
  ```
