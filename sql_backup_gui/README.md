# SQL Server Parallel Copy-Only Backup & Fast-Zip GUI

A self-contained graphical utility for backing up user-selected SQL Server databases in **`COPY_ONLY`** mode to a target folder, featuring **parallel multi-database execution**, **automatic multi-stream striped backups for databases larger than 8 GB**, and **fastest-speed post-backup ZIP archiving**.

Designed to run seamlessly on **both user desktops (with Python)** and **Windows Servers (without Python)**.

## Files Included

| File | Purpose |
| :--- | :--- |
| `Launch_SQL_Backup.bat` | **Universal Launcher** — Automatically runs `sql_backup_gui.py` when Python is available, or falls back to `sql_backup_gui.ps1` on Windows Servers without Python. |
| `sql_backup_gui.py` | **Python Desktop GUI (`tkinter`)** — Zero mandatory third-party `pip` dependencies on Windows (uses a built-in `.NET System.Data.SqlClient` bridge and `zipfile` `compresslevel=1` / `7z.exe`). |
| `sql_backup_gui.ps1` | **Native Windows Server GUI (PowerShell / .NET WinForms)** — Runs on Windows Servers with **no Python installation required**, using PowerShell RunspacePools, `.NET SqlClient`, and `System.IO.Compression.CompressionLevel::Fastest`. |
| `build_standalone_exe.bat` | Optional helper script that packages `sql_backup_gui.py` into a single-file `dist\sql_backup_gui.exe` via PyInstaller. |
| `test_sql_backup_gui.py` | Unit tests for `> 8 GB` stream striping calculation, per-database subfolder path generation, `BACKUP DATABASE` T-SQL formatting, Fast-Zip archiving, and GUI initialization. |

## Key Features

1. **Always-Safe `COPY_ONLY` Backups**:
   - Every backup executes `WITH COPY_ONLY` so existing production full/differential/transaction log backup chains are never disrupted.
2. **Parallel Database Execution**:
   - Backs up multiple databases concurrently (default **3 parallel database workers**, configurable from `1` to `16` in the UI).
   - Streams live `STATS = 5` progress percentages from SQL Server (`InfoMessage` events) directly into each database's row and the overall batch progress bar.
3. **Automatic Parallel Stream Striping for Large Databases (`> 8 GB`)**:
   - Databases $\le 8\text{ GB}$ back up to a single `.bak` file:
     `<TargetFolder>\<DbName>\<DbName>_<Timestamp>.bak`
   - Databases $> 8\text{ GB}$ are automatically split into parallel striped streams (`ceil(Size_GB / 8.0)`, up to `8` streams by default, configurable in the UI):
     ```sql
     BACKUP DATABASE [EnterpriseWarehouseDB]
     TO
         DISK = N'D:\SqlBackups\EnterpriseWarehouseDB\EnterpriseWarehouseDB_20261021_100000_stripe01of04.bak',
         DISK = N'D:\SqlBackups\EnterpriseWarehouseDB\EnterpriseWarehouseDB_20261021_100000_stripe02of04.bak',
         DISK = N'D:\SqlBackups\EnterpriseWarehouseDB\EnterpriseWarehouseDB_20261021_100000_stripe03of04.bak',
         DISK = N'D:\SqlBackups\EnterpriseWarehouseDB\EnterpriseWarehouseDB_20261021_100000_stripe04of04.bak'
     WITH
         COPY_ONLY,
         COMPRESSION,
         CHECKSUM,
         INIT,
         STATS = 5;
     ```
4. **Fast Post-Backup Zip Archiving (Speed Over Compression Ratio)**:
   - Immediately after each database's `.bak` stripe(s) finish writing, its worker thread zips the `.bak` file(s) into `<ZipFolder>\<DbName>_<Timestamp>.zip` (default `<TargetFolder>\Zipped\`).
   - Prioritizes **speed of zip over compression ratio**:
     - Uses `7z.exe a -tzip -mx=1 -mmt=on` (multi-threaded fastest level) if 7-Zip is installed.
     - Otherwise uses Python `zipfile.ZipFile(..., compresslevel=1)` (`Z_BEST_SPEED` with 16 MB streaming buffers) or .NET `[System.IO.Compression.CompressionLevel]::Fastest`.
   - Includes an optional checkbox to **delete the raw `.bak` files** once the `.zip` archive is verified.
5. **Automatic Subfolder Creation & Compression Fallback**:
   - Creates `<TargetFolder>\<DbName>` on the SQL Server host via `EXEC master.dbo.xp_create_subdir` prior to each backup.
   - Enables `COMPRESSION` and `CHECKSUM` by default, with automatic retry without `COMPRESSION` if the target SQL Server edition (e.g. SQL Server Express) does not support backup compression.

## Quick Start

- **Any Windows Machine (Desktop or Server)**:
  Double-click `Launch_SQL_Backup.bat`
- **Run Python Edition directly**:
  ```powershell
  python sql_backup_gui\sql_backup_gui.py
  ```
- **Run Native Server PowerShell Edition directly (No Python needed)**:
  ```powershell
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\sql_backup_gui\sql_backup_gui.ps1
  ```
