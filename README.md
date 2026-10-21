# AWS S3 CP Transfer GUI

A self-contained local graphical user interface for transferring files and folders to and from Amazon S3 (`aws s3 cp`). Designed to run seamlessly on **both user desktops (with Python)** and **Windows servers (without Python)**.

## Files Included

| File | Purpose |
| :--- | :--- |
| `Launch_AWS_S3_CP.bat` | **Universal Launcher** — Automatically runs `aws_cp_gui.py` when Python is available, or falls back to `aws_cp_gui.ps1` on Windows servers without Python. |
| `aws_cp_gui.py` | **Python Desktop GUI (`tkinter`)** — Zero mandatory third-party `pip` dependencies. Uses `boto3` if installed, or a built-in pure-stdlib AWS Signature Version 4 S3 client, plus `aws s3 cp` CLI execution. |
| `aws_cp_gui.ps1` | **Native Windows Server GUI (PowerShell / .NET WinForms)** — Runs on Windows Servers with **no Python installation required**. Includes a built-in .NET AWS SigV4 S3 client and `aws s3 cp` CLI integration. |
| `build_standalone_exe.bat` | Optional helper script that packages `aws_cp_gui.py` into a single-file `dist\aws_cp_gui.exe` via PyInstaller. |
| `test_aws_cp_gui.py` | Unit tests for S3 URI parsing, AWS SigV4 request signing, CLI command generation, and GUI initialization. |

## Key Features

1. **AWS Configuration Attributes**:
   - **AWS Access Key ID** (`AWS_ACCESS_KEY_ID`)
   - **AWS Secret Access Key** (`AWS_SECRET_ACCESS_KEY`) with Show/Hide toggle
   - **Default Region** (`AWS_DEFAULT_REGION`, e.g., `us-east-1`, `eu-west-2`, plus automatic bucket region detection)
   - **Default Output Format** (`AWS_DEFAULT_OUTPUT`: `json`, `text`, `table`, `yaml`, `yaml-stream`)
   - Optional Session Token & one-click Load/Save for `~/.aws/credentials`, `~/.aws/config`, or local `.aws_cp_config.json`.
2. **Bidirectional Transfers**:
   - **Copy TO S3 Bucket (Upload)**: Pick a local source file or folder and a target S3 bucket/folder.
   - **Copy FROM S3 Bucket (Download)**: Pick a source S3 object or folder (prefix) and a target local folder/file.
3. **Dual Explorers**:
   - **Local File & Folder Explorer**: Interactive folder/file browser plus native OS `Browse Local File...` and `Browse Local Folder...` dialogs.
   - **S3 Bucket & Folder Explorer**: Connect and list buckets, or type a bucket name directly (if your IAM policy restricts `s3:ListAllMyBuckets`), browse prefixes, and create S3 folders.
4. **Live `aws s3 cp` Command Preview & Dual Execution Engines**:
   - Displays the exact `aws s3 cp` command with `--recursive`, `--dryrun`, `--storage-class`, `--region`, and `--output`.
   - Executes via `aws s3 cp` CLI (when installed) or the built-in zero-dependency S3 HTTPS engine.

## Quick Start

- **Any Windows Machine (Desktop or Server)**:
  Double-click `Launch_AWS_S3_CP.bat`
- **Run Python Edition directly**:
  ```powershell
  python aws_cp_gui.py
  ```
- **Run Native Server PowerShell Edition directly (No Python needed)**:
  ```powershell
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\aws_cp_gui.ps1
  ```
