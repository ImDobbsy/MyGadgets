#!/usr/bin/env python3
"""
AWS S3 CP Transfer GUI (`aws_cp_gui.py`)
========================================
A self-contained local graphical application for copying files and folders
to and from Amazon S3 buckets (`aws s3 cp`).

Features:
- Pass AWS Access Key ID, Secret Access Key, Default Region, and Default Output Format
- Dual-mode transfers:
    1. Copy TO S3 Bucket (Local File/Folder -> S3 Bucket/Folder)
    2. Copy FROM S3 Bucket (S3 Object/Folder -> Local Folder/File)
- Built-in Local Explorer + Native OS File/Folder pickers
- Built-in S3 Bucket & Folder Explorer (supports typing bucket name directly or listing all buckets)
- Works on bare Python installs (zero mandatory third-party packages) via a built-in
  pure-stdlib AWS Signature Version 4 (SigV4) S3 client, while also supporting
  `boto3` (if installed) and the `aws s3 cp` CLI (if installed).
- Live `aws s3 cp` CLI command preview & one-click copy.
- Optional credential loading/saving for `~/.aws/credentials` & `~/.aws/config` or local JSON.
"""

from __future__ import annotations

import configparser
import datetime
import hashlib
import hmac
import io
import json
import os
import platform
import queue
import shlex
import shutil
import ssl
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Optional boto3 detection
try:
    import boto3
    from botocore.config import Config as BotoConfig
    from botocore.exceptions import BotoCoreError, ClientError

    HAS_BOTO3 = True
except ImportError:
    HAS_BOTO3 = False


# ============================================================================
# Constants & Defaults
# ============================================================================

APP_TITLE = "AWS S3 CP Transfer Utility"
APP_VERSION = "1.0.0"
LOCAL_CONFIG_FILENAME = ".aws_cp_config.json"

AWS_REGIONS = [
    "us-east-1",
    "us-east-2",
    "us-west-1",
    "us-west-2",
    "eu-west-1",
    "eu-west-2",
    "eu-west-3",
    "eu-central-1",
    "eu-central-2",
    "eu-north-1",
    "eu-south-1",
    "ap-south-1",
    "ap-southeast-1",
    "ap-southeast-2",
    "ap-southeast-3",
    "ap-northeast-1",
    "ap-northeast-2",
    "ap-northeast-3",
    "ap-east-1",
    "ca-central-1",
    "sa-east-1",
    "me-south-1",
    "me-central-1",
    "af-south-1",
    "il-central-1",
]

AWS_OUTPUT_FORMATS = [
    "json",
    "text",
    "table",
    "yaml",
    "yaml-stream",
]

STORAGE_CLASSES = [
    "(Default)",
    "STANDARD",
    "REDUCED_REDUNDANCY",
    "STANDARD_IA",
    "ONEZONE_IA",
    "INTELLIGENT_TIERING",
    "GLACIER",
    "DEEP_ARCHIVE",
    "GLACIER_IR",
]


def format_size(num_bytes: Optional[int]) -> str:
    """Return human-readable byte size."""
    if num_bytes is None:
        return ""
    size = float(num_bytes)
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if size < 1024.0 or unit == "TB":
            return f"{int(size)} B" if unit == "B" else f"{size:.1f} {unit}"
        size /= 1024.0
    return f"{size:.1f} PB"


def find_aws_cli() -> Optional[str]:
    """Locate the `aws` CLI executable on PATH or standard Windows install paths."""
    aws_bin = shutil.which("aws")
    if aws_bin:
        return aws_bin
    if os.name == "nt":
        candidates = [
            r"C:\Program Files\Amazon\AWSCLIV2\aws.exe",
            r"C:\Program Files (x86)\Amazon\AWSCLIV2\aws.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Programs\Amazon\AWSCLIV2\aws.exe"),
        ]
        for candidate in candidates:
            if os.path.isfile(candidate):
                return candidate
    return None


def parse_s3_uri(uri: str) -> Tuple[str, str]:
    """
    Parse an S3 URI like `s3://my-bucket/path/to/key` into `(bucket, key)`.
    Raises ValueError if invalid.
    """
    cleaned = uri.strip()
    if cleaned.startswith("s3://"):
        cleaned = cleaned[5:]
    cleaned = cleaned.lstrip("/")
    if not cleaned:
        raise ValueError("S3 URI must specify at least a bucket name (e.g. s3://my-bucket/).")
    parts = cleaned.split("/", 1)
    bucket = parts[0].strip()
    key = parts[1] if len(parts) > 1 else ""
    if not bucket:
        raise ValueError("Bucket name cannot be empty in S3 URI.")
    return bucket, key


def build_s3_uri(bucket: str, key: str = "") -> str:
    """Construct a normalized `s3://bucket/key` string."""
    bucket = bucket.strip().strip("/")
    key = key.lstrip("/")
    if not bucket:
        return "s3://"
    return f"s3://{bucket}/{key}" if key else f"s3://{bucket}/"


# ============================================================================
# Data Models
# ============================================================================


@dataclass
class AWSCredentialsConfig:
    access_key_id: str = ""
    secret_access_key: str = ""
    region: str = "us-east-1"
    output_format: str = "json"
    session_token: str = ""
    profile_name: str = "default"

    def validate(self) -> None:
        if not self.access_key_id.strip():
            raise ValueError("AWS Access Key ID is required.")
        if not self.secret_access_key.strip():
            raise ValueError("AWS Secret Access Key is required.")
        if not self.region.strip():
            raise ValueError("AWS Region is required.")
        if not self.output_format.strip():
            raise ValueError("Default Output Format is required.")

    def to_env_dict(self) -> Dict[str, str]:
        """Return environment variables dictionary for AWS CLI / SDK processes."""
        env = os.environ.copy()
        env["AWS_ACCESS_KEY_ID"] = self.access_key_id.strip()
        env["AWS_SECRET_ACCESS_KEY"] = self.secret_access_key.strip()
        env["AWS_DEFAULT_REGION"] = self.region.strip()
        env["AWS_REGION"] = self.region.strip()
        env["AWS_DEFAULT_OUTPUT"] = self.output_format.strip()
        if self.session_token.strip():
            env["AWS_SESSION_TOKEN"] = self.session_token.strip()
        elif "AWS_SESSION_TOKEN" in env:
            del env["AWS_SESSION_TOKEN"]
        return env


@dataclass
class S3Entry:
    name: str
    key: str
    is_dir: bool
    size: Optional[int] = None
    last_modified: str = ""
    storage_class: str = ""


# ============================================================================
# Pure-Stdlib AWS Signature Version 4 (SigV4) S3 Client
# ============================================================================


class _ProgressFileReader(io.RawIOBase):
    """File-like wrapper that reports read progress and exposes Content-Length for urllib."""

    def __init__(
        self,
        file_path: str,
        progress_cb: Optional[Callable[[int, int], None]] = None,
        cancel_event: Optional[threading.Event] = None,
    ):
        self._fp = open(file_path, "rb")
        self._total = os.path.getsize(file_path)
        self._sent = 0
        self._progress_cb = progress_cb
        self._cancel_event = cancel_event

    def __len__(self) -> int:
        return self._total

    def read(self, size: int = -1) -> bytes:
        if self._cancel_event and self._cancel_event.is_set():
            raise InterruptedError("Transfer cancelled by user.")
        chunk = self._fp.read(size if size > 0 else 262144)
        if chunk:
            self._sent += len(chunk)
            if self._progress_cb:
                self._progress_cb(self._sent, self._total)
        return chunk

    def close(self) -> None:
        if not self._fp.closed:
            self._fp.close()
        super().close()

    def __enter__(self) -> "_ProgressFileReader":
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.close()


class StdlibS3Client:
    """
    Zero-dependency S3 client using Python standard library (`urllib.request`, `hmac`, `hashlib`, `xml`).
    Supports:
    - Listing all buckets (`ListBuckets`)
    - Listing objects and common prefixes (`ListObjectsV2`)
    - Uploading files (`PutObject` with `UNSIGNED-PAYLOAD` streaming over HTTPS)
    - Downloading files (`GetObject` with chunked streaming)
    - Automatic bucket region discovery via `x-amz-bucket-region`
    """

    EMPTY_SHA256 = "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"

    def __init__(self, creds: AWSCredentialsConfig):
        self.creds = creds
        self._bucket_region_cache: Dict[str, str] = {}

    @staticmethod
    def _sign(key: bytes, msg: str) -> bytes:
        return hmac.new(key, msg.encode("utf-8"), hashlib.sha256).digest()

    @classmethod
    def _get_signature_key(cls, secret_key: str, date_stamp: str, region: str, service: str = "s3") -> bytes:
        k_date = cls._sign(("AWS4" + secret_key).encode("utf-8"), date_stamp)
        k_region = cls._sign(k_date, region)
        k_service = cls._sign(k_region, service)
        k_signing = cls._sign(k_service, "aws4_request")
        return k_signing

    @staticmethod
    def _uri_encode(value: str, encode_slash: bool = True) -> str:
        safe = "-_.~" if encode_slash else "/-_.~"
        return urllib.parse.quote(value, safe=safe)

    def _build_signed_request(
        self,
        method: str,
        bucket: Optional[str],
        key: str,
        query_params: Optional[Dict[str, str]] = None,
        headers: Optional[Dict[str, str]] = None,
        payload_hash: str = EMPTY_SHA256,
        region_override: Optional[str] = None,
        data: Optional[object] = None,
    ) -> urllib.request.Request:
        region = (region_override or self._bucket_region_cache.get(bucket or "") or self.creds.region or "us-east-1").strip()
        query_params = query_params or {}
        extra_headers = dict(headers or {})

        # Use virtual-hosted style unless bucket name contains dots (SSL wildcard limitation)
        if not bucket:
            host = f"s3.{region}.amazonaws.com"
            canonical_uri = "/"
        elif "." in bucket:
            host = f"s3.{region}.amazonaws.com"
            encoded_key = self._uri_encode(key.lstrip("/"), encode_slash=False)
            canonical_uri = f"/{bucket}/{encoded_key}" if encoded_key else f"/{bucket}"
        else:
            host = f"{bucket}.s3.{region}.amazonaws.com"
            encoded_key = self._uri_encode(key.lstrip("/"), encode_slash=False)
            canonical_uri = f"/{encoded_key}"

        now = datetime.datetime.now(datetime.timezone.utc)
        amz_date = now.strftime("%Y%m%dT%H%M%SZ")
        date_stamp = now.strftime("%Y%m%d")

        # Canonical query string sorted by parameter key
        sorted_q = sorted(query_params.items(), key=lambda kv: kv[0])
        canonical_querystring = "&".join(
            f"{self._uri_encode(k, True)}={self._uri_encode(str(v), True)}" for k, v in sorted_q
        )

        # Prepare headers to sign
        sign_headers: Dict[str, str] = {
            "host": host,
            "x-amz-content-sha256": payload_hash,
            "x-amz-date": amz_date,
        }
        if self.creds.session_token.strip():
            sign_headers["x-amz-security-token"] = self.creds.session_token.strip()
        for hk, hv in extra_headers.items():
            sign_headers[hk.lower()] = str(hv).strip()

        sorted_header_keys = sorted(sign_headers.keys())
        canonical_headers = "".join(f"{k}:{sign_headers[k]}\n" for k in sorted_header_keys)
        signed_headers = ";".join(sorted_header_keys)

        canonical_request = "\n".join(
            [
                method.upper(),
                canonical_uri,
                canonical_querystring,
                canonical_headers,
                signed_headers,
                payload_hash,
            ]
        )

        algorithm = "AWS4-HMAC-SHA256"
        credential_scope = f"{date_stamp}/{region}/s3/aws4_request"
        string_to_sign = "\n".join(
            [
                algorithm,
                amz_date,
                credential_scope,
                hashlib.sha256(canonical_request.encode("utf-8")).hexdigest(),
            ]
        )

        signing_key = self._get_signature_key(
            self.creds.secret_access_key.strip(), date_stamp, region, "s3"
        )
        signature = hmac.new(signing_key, string_to_sign.encode("utf-8"), hashlib.sha256).hexdigest()

        authorization_header = (
            f"{algorithm} Credential={self.creds.access_key_id.strip()}/{credential_scope}, "
            f"SignedHeaders={signed_headers}, Signature={signature}"
        )

        url = f"https://{host}{canonical_uri}"
        if canonical_querystring:
            url = f"{url}?{canonical_querystring}"

        req_headers = {k: v for k, v in sign_headers.items() if k != "host"}
        req_headers["Authorization"] = authorization_header
        req_headers["Host"] = host

        return urllib.request.Request(url, data=data, headers=req_headers, method=method.upper())

    def _execute(
        self,
        method: str,
        bucket: Optional[str],
        key: str = "",
        query_params: Optional[Dict[str, str]] = None,
        headers: Optional[Dict[str, str]] = None,
        payload_hash: str = EMPTY_SHA256,
        data: Optional[object] = None,
        timeout: int = 60,
    ):
        """Execute signed S3 request with automatic region discovery and retry on region mismatch."""
        self.creds.validate()
        ctx = ssl.create_default_context()
        region_override = self._bucket_region_cache.get(bucket or "")

        for attempt in range(2):
            req = self._build_signed_request(
                method=method,
                bucket=bucket,
                key=key,
                query_params=query_params,
                headers=headers,
                payload_hash=payload_hash,
                region_override=region_override,
                data=data,
            )
            try:
                return urllib.request.urlopen(req, context=ctx, timeout=timeout)
            except urllib.error.HTTPError as err:
                discovered_region = err.headers.get("x-amz-bucket-region")
                err_body = ""
                try:
                    err_body = err.read().decode("utf-8", errors="replace")
                except Exception:
                    pass

                if not discovered_region and "<Region>" in err_body:
                    start = err_body.find("<Region>") + len("<Region>")
                    end = err_body.find("</Region>", start)
                    if start > 7 and end > start:
                        discovered_region = err_body[start:end].strip()

                if (
                    attempt == 0
                    and bucket
                    and discovered_region
                    and discovered_region != (region_override or self.creds.region)
                    and data is None
                ):
                    self._bucket_region_cache[bucket] = discovered_region
                    region_override = discovered_region
                    continue

                msg = self._parse_s3_error_xml(err.code, err.reason, err_body)
                raise RuntimeError(msg) from err
            except urllib.error.URLError as err:
                raise RuntimeError(f"Network error connecting to S3: {err.reason}") from err

    @staticmethod
    def _parse_s3_error_xml(code: int, reason: str, xml_text: str) -> str:
        if xml_text:
            try:
                root = ET.fromstring(xml_text)
                err_code = root.findtext("Code") or root.findtext("{*}Code") or ""
                err_msg = root.findtext("Message") or root.findtext("{*}Message") or ""
                if err_code or err_msg:
                    return f"S3 Error ({code} {err_code}): {err_msg}"
            except ET.ParseError:
                pass
        return f"S3 HTTP Error {code}: {reason}"

    def list_buckets(self) -> List[str]:
        """Return list of bucket names visible to the credentials."""
        with self._execute("GET", bucket=None) as resp:
            xml_bytes = resp.read()
        root = ET.fromstring(xml_bytes)
        buckets: List[str] = []
        for elem in root.iter():
            if elem.tag.endswith("Bucket"):
                for child in elem:
                    if child.tag.endswith("Name") and child.text:
                        buckets.append(child.text.strip())
        return sorted(buckets)

    def list_objects_v2(self, bucket: str, prefix: str = "", delimiter: str = "/") -> List[S3Entry]:
        """List folders (CommonPrefixes) and objects (Contents) at a given bucket/prefix."""
        entries: List[S3Entry] = []
        continuation_token: Optional[str] = None

        while True:
            params: Dict[str, str] = {"list-type": "2", "max-keys": "1000"}
            if prefix:
                params["prefix"] = prefix
            if delimiter:
                params["delimiter"] = delimiter
            if continuation_token:
                params["continuation-token"] = continuation_token

            with self._execute("GET", bucket=bucket, query_params=params) as resp:
                xml_bytes = resp.read()

            root = ET.fromstring(xml_bytes)
            is_truncated = False
            next_token = None

            for child in root:
                tag = child.tag.split("}")[-1]
                if tag == "CommonPrefixes":
                    p_val = ""
                    for sub in child:
                        if sub.tag.endswith("Prefix") and sub.text:
                            p_val = sub.text
                    if p_val:
                        rel_name = p_val[len(prefix) :].rstrip("/") if p_val.startswith(prefix) else p_val.rstrip("/")
                        entries.append(
                            S3Entry(
                                name=rel_name or p_val,
                                key=p_val,
                                is_dir=True,
                            )
                        )
                elif tag == "Contents":
                    key_val = ""
                    size_val = 0
                    last_mod = ""
                    storage_cls = ""
                    for sub in child:
                        stag = sub.tag.split("}")[-1]
                        if stag == "Key" and sub.text:
                            key_val = sub.text
                        elif stag == "Size" and sub.text:
                            size_val = int(sub.text)
                        elif stag == "LastModified" and sub.text:
                            last_mod = sub.text[:19].replace("T", " ")
                        elif stag == "StorageClass" and sub.text:
                            storage_cls = sub.text
                    # Skip folder placeholder object that matches prefix itself
                    if key_val and key_val != prefix:
                        rel_name = key_val[len(prefix) :] if key_val.startswith(prefix) else key_val
                        entries.append(
                            S3Entry(
                                name=rel_name,
                                key=key_val,
                                is_dir=False,
                                size=size_val,
                                last_modified=last_mod,
                                storage_class=storage_cls,
                            )
                        )
                elif tag == "IsTruncated" and child.text:
                    is_truncated = child.text.strip().lower() == "true"
                elif tag == "NextContinuationToken" and child.text:
                    next_token = child.text.strip()

            if is_truncated and next_token:
                continuation_token = next_token
            else:
                break

        # Sort directories first, then files alphabetically
        entries.sort(key=lambda e: (not e.is_dir, e.name.lower()))
        return entries

    def create_folder_placeholder(self, bucket: str, prefix: str) -> None:
        """Create a zero-byte S3 folder placeholder (`prefix/`)."""
        norm_prefix = prefix.strip("/") + "/"
        with self._execute(
            "PUT",
            bucket=bucket,
            key=norm_prefix,
            headers={"Content-Length": "0"},
            payload_hash=self.EMPTY_SHA256,
            data=b"",
        ):
            pass

    def upload_file(
        self,
        local_path: str,
        bucket: str,
        key: str,
        storage_class: Optional[str] = None,
        progress_cb: Optional[Callable[[int, int], None]] = None,
        cancel_event: Optional[threading.Event] = None,
    ) -> None:
        """Stream upload a local file to `s3://bucket/key` using `UNSIGNED-PAYLOAD`."""
        file_size = os.path.getsize(local_path)
        headers: Dict[str, str] = {"Content-Length": str(file_size)}
        if storage_class and storage_class != "(Default)":
            headers["x-amz-storage-class"] = storage_class

        # Ensure bucket region is cached before sending streaming body so retry isn't needed mid-stream
        if bucket not in self._bucket_region_cache:
            try:
                with self._execute("GET", bucket=bucket, query_params={"list-type": "2", "max-keys": "1"}) as resp:
                    reg = resp.headers.get("x-amz-bucket-region")
                    if reg:
                        self._bucket_region_cache[bucket] = reg
            except Exception:
                pass

        if file_size == 0:
            with self._execute(
                "PUT",
                bucket=bucket,
                key=key,
                headers=headers,
                payload_hash=self.EMPTY_SHA256,
                data=b"",
                timeout=120,
            ):
                if progress_cb:
                    progress_cb(0, 0)
            return

        with _ProgressFileReader(local_path, progress_cb=progress_cb, cancel_event=cancel_event) as reader:
            with self._execute(
                "PUT",
                bucket=bucket,
                key=key,
                headers=headers,
                payload_hash="UNSIGNED-PAYLOAD",
                data=reader,
                timeout=300,
            ):
                pass

    def download_file(
        self,
        bucket: str,
        key: str,
        local_path: str,
        progress_cb: Optional[Callable[[int, int], None]] = None,
        cancel_event: Optional[threading.Event] = None,
    ) -> None:
        """Stream download `s3://bucket/key` to `local_path`."""
        os.makedirs(os.path.dirname(os.path.abspath(local_path)), exist_ok=True)
        with self._execute("GET", bucket=bucket, key=key, timeout=300) as resp:
            total = int(resp.headers.get("Content-Length") or 0)
            downloaded = 0
            with open(local_path, "wb") as out_fp:
                while True:
                    if cancel_event and cancel_event.is_set():
                        raise InterruptedError("Transfer cancelled by user.")
                    chunk = resp.read(262144)
                    if not chunk:
                        break
                    out_fp.write(chunk)
                    downloaded += len(chunk)
                    if progress_cb:
                        progress_cb(downloaded, total)

    def delete_object(self, bucket: str, key: str) -> None:
        """Delete a single object `s3://bucket/key`."""
        with self._execute("DELETE", bucket=bucket, key=key, timeout=60):
            pass


# ============================================================================
# Unified S3 Service (Boto3 if available, otherwise Pure-Stdlib SigV4)
# ============================================================================


class S3Service:
    """Provides S3 exploration and transfers using either `boto3` or `StdlibS3Client`."""

    def __init__(self, creds: AWSCredentialsConfig):
        self.creds = creds
        self.stdlib_client = StdlibS3Client(creds)

    def _get_boto3_client(self):
        self.creds.validate()
        kwargs = {
            "aws_access_key_id": self.creds.access_key_id.strip(),
            "aws_secret_access_key": self.creds.secret_access_key.strip(),
            "region_name": self.creds.region.strip(),
            "config": BotoConfig(signature_version="s3v4"),
        }
        if self.creds.session_token.strip():
            kwargs["aws_session_token"] = self.creds.session_token.strip()
        return boto3.client("s3", **kwargs)

    def list_buckets(self) -> List[str]:
        if HAS_BOTO3:
            try:
                client = self._get_boto3_client()
                resp = client.list_buckets()
                return sorted(b["Name"] for b in resp.get("Buckets", []))
            except Exception:
                pass
        return self.stdlib_client.list_buckets()

    def list_objects(self, bucket: str, prefix: str = "", delimiter: str = "/") -> List[S3Entry]:
        # Stdlib client handles automatic cross-region discovery cleanly
        return self.stdlib_client.list_objects_v2(bucket=bucket, prefix=prefix, delimiter=delimiter)

    def create_folder(self, bucket: str, prefix: str) -> None:
        self.stdlib_client.create_folder_placeholder(bucket=bucket, prefix=prefix)

    def upload_file(
        self,
        local_path: str,
        bucket: str,
        key: str,
        storage_class: Optional[str] = None,
        progress_cb: Optional[Callable[[int, int], None]] = None,
        cancel_event: Optional[threading.Event] = None,
    ) -> None:
        if HAS_BOTO3:
            try:
                client = self._get_boto3_client()
                total = os.path.getsize(local_path)
                sent = 0

                def _cb(bytes_amount: int):
                    nonlocal sent
                    if cancel_event and cancel_event.is_set():
                        raise InterruptedError("Transfer cancelled by user.")
                    sent += bytes_amount
                    if progress_cb:
                        progress_cb(sent, total)

                extra_args = {}
                if storage_class and storage_class != "(Default)":
                    extra_args["StorageClass"] = storage_class
                client.upload_file(local_path, bucket, key, ExtraArgs=extra_args or None, Callback=_cb)
                return
            except InterruptedError:
                raise
            except Exception:
                # Fallback to stdlib client (which handles cross-region automatically)
                pass

        self.stdlib_client.upload_file(
            local_path=local_path,
            bucket=bucket,
            key=key,
            storage_class=storage_class,
            progress_cb=progress_cb,
            cancel_event=cancel_event,
        )

    def download_file(
        self,
        bucket: str,
        key: str,
        local_path: str,
        progress_cb: Optional[Callable[[int, int], None]] = None,
        cancel_event: Optional[threading.Event] = None,
    ) -> None:
        self.stdlib_client.download_file(
            bucket=bucket,
            key=key,
            local_path=local_path,
            progress_cb=progress_cb,
            cancel_event=cancel_event,
        )

    def delete_object(self, bucket: str, key: str) -> None:
        self.stdlib_client.delete_object(bucket=bucket, key=key)


def filter_zip_entries(entries: List[S3Entry]) -> List[S3Entry]:
    """Return only non-directory S3 entries whose key ends with `.zip` (case-insensitive)."""
    return [e for e in entries if (not e.is_dir) and e.key.lower().endswith(".zip")]


def build_aws_rm_zip_command_args(
    s3_uri: str,
    region: str,
    output_format: str,
    recursive: bool = False,
    dryrun: bool = False,
    aws_executable: str = "aws",
) -> List[str]:
    """Build the argument list for `aws s3 rm` targeting only `.zip` files."""
    args = [
        aws_executable,
        "s3",
        "rm",
        s3_uri.strip(),
        "--recursive",
        "--exclude",
        "*",
        "--include",
        "*.zip" if recursive else "*.zip",
    ]
    if dryrun:
        args.append("--dryrun")
    if region.strip():
        args.extend(["--region", region.strip()])
    if output_format.strip():
        args.extend(["--output", output_format.strip()])
    return args


# ============================================================================
# Command Builder & Credential Persistence Helpers
# ============================================================================


def build_aws_cp_command_args(
    source: str,
    target: str,
    region: str,
    output_format: str,
    recursive: bool = False,
    dryrun: bool = False,
    only_show_errors: bool = False,
    storage_class: str = "(Default)",
    exclude_pattern: str = "",
    include_pattern: str = "",
    aws_executable: str = "aws",
) -> List[str]:
    """Build the argument list for `aws s3 cp`."""
    args = [aws_executable, "s3", "cp", source.strip(), target.strip()]
    if recursive:
        args.append("--recursive")
    if dryrun:
        args.append("--dryrun")
    if only_show_errors:
        args.append("--only-show-errors")
    if storage_class and storage_class != "(Default)":
        args.extend(["--storage-class", storage_class])
    if exclude_pattern.strip():
        args.extend(["--exclude", exclude_pattern.strip()])
    if include_pattern.strip():
        args.extend(["--include", include_pattern.strip()])
    if region.strip():
        args.extend(["--region", region.strip()])
    if output_format.strip():
        args.extend(["--output", output_format.strip()])
    return args


def format_command_preview(args: List[str]) -> str:
    """Format a CLI command list into a readable shell command string."""
    formatted: List[str] = []
    for a in args:
        if " " in a or "\t" in a or ('"' in a):
            escaped = a.replace('"', '\\"')
            formatted.append(f'"{escaped}"')
        else:
            formatted.append(a)
    return " ".join(formatted)


def load_aws_profile_files(profile: str = "default") -> AWSCredentialsConfig:
    """Load credentials and config from `~/.aws/credentials` and `~/.aws/config`."""
    home_aws = Path.home() / ".aws"
    cred_file = home_aws / "credentials"
    conf_file = home_aws / "config"

    cfg = AWSCredentialsConfig(profile_name=profile)

    if cred_file.is_file():
        cp = configparser.RawConfigParser()
        cp.read(cred_file, encoding="utf-8")
        if cp.has_section(profile):
            cfg.access_key_id = cp.get(profile, "aws_access_key_id", fallback="").strip()
            cfg.secret_access_key = cp.get(profile, "aws_secret_access_key", fallback="").strip()
            cfg.session_token = cp.get(profile, "aws_session_token", fallback="").strip()

    if conf_file.is_file():
        cp = configparser.RawConfigParser()
        cp.read(conf_file, encoding="utf-8")
        sec = profile if profile == "default" else f"profile {profile}"
        if not cp.has_section(sec) and cp.has_section(profile):
            sec = profile
        if cp.has_section(sec):
            cfg.region = cp.get(sec, "region", fallback=cfg.region).strip() or cfg.region
            cfg.output_format = cp.get(sec, "output", fallback=cfg.output_format).strip() or cfg.output_format

    return cfg


def save_aws_profile_files(cfg: AWSCredentialsConfig) -> Tuple[Path, Path]:
    """Save credentials and config to `~/.aws/credentials` and `~/.aws/config`."""
    home_aws = Path.home() / ".aws"
    home_aws.mkdir(parents=True, exist_ok=True)
    cred_file = home_aws / "credentials"
    conf_file = home_aws / "config"
    profile = cfg.profile_name.strip() or "default"

    cred_cp = configparser.RawConfigParser()
    if cred_file.is_file():
        cred_cp.read(cred_file, encoding="utf-8")
    if not cred_cp.has_section(profile):
        cred_cp.add_section(profile)
    cred_cp.set(profile, "aws_access_key_id", cfg.access_key_id.strip())
    cred_cp.set(profile, "aws_secret_access_key", cfg.secret_access_key.strip())
    if cfg.session_token.strip():
        cred_cp.set(profile, "aws_session_token", cfg.session_token.strip())
    elif cred_cp.has_option(profile, "aws_session_token"):
        cred_cp.remove_option(profile, "aws_session_token")

    with open(cred_file, "w", encoding="utf-8") as fp:
        cred_cp.write(fp)

    conf_cp = configparser.RawConfigParser()
    if conf_file.is_file():
        conf_cp.read(conf_file, encoding="utf-8")
    sec = profile if profile == "default" else f"profile {profile}"
    if not conf_cp.has_section(sec):
        conf_cp.add_section(sec)
    conf_cp.set(sec, "region", cfg.region.strip())
    conf_cp.set(sec, "output", cfg.output_format.strip())

    with open(conf_file, "w", encoding="utf-8") as fp:
        conf_cp.write(fp)

    return cred_file, conf_file


# ============================================================================
# Graphical User Interface (Tkinter / ttk)
# ============================================================================


class AwsCpTransferApp:
    """Main Tkinter GUI application for AWS S3 CP transfers."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"{APP_TITLE} v{APP_VERSION}")
        self.root.geometry("1220x840")
        self.root.minsize(980, 700)

        self.ui_queue: "queue.Queue[Callable[[], None]]" = queue.Queue()
        self.cancel_event = threading.Event()
        self.active_process: Optional[subprocess.Popen] = None
        self.is_transferring = False

        # State variables
        self.var_access_key = tk.StringVar()
        self.var_secret_key = tk.StringVar()
        self.var_show_secret = tk.BooleanVar(value=False)
        self.var_region = tk.StringVar(value="us-east-1")
        self.var_output_format = tk.StringVar(value="json")
        self.var_session_token = tk.StringVar()
        self.var_profile = tk.StringVar(value="default")

        # Transfer direction: "upload" (Local -> S3) or "download" (S3 -> Local)
        self.var_direction = tk.StringVar(value="upload")
        self.var_source_path = tk.StringVar()
        self.var_target_path = tk.StringVar()

        # Options
        self.var_recursive = tk.BooleanVar(value=False)
        self.var_dryrun = tk.BooleanVar(value=False)
        self.var_only_errors = tk.BooleanVar(value=False)
        self.var_storage_class = tk.StringVar(value="(Default)")
        self.var_exclude = tk.StringVar()
        self.var_include = tk.StringVar()
        self.var_engine = tk.StringVar(value="Auto (AWS CLI or Built-in S3)")
        self.var_cmd_preview = tk.StringVar()

        # Explorer state
        self.var_local_dir = tk.StringVar(value=str(Path.cwd()))
        self.var_s3_bucket = tk.StringVar()
        self.var_s3_prefix = tk.StringVar(value="")

        # Progress state
        self.var_status = tk.StringVar(value="Ready.")
        self.var_progress = tk.DoubleVar(value=0.0)

        self._configure_styles()
        self._build_ui()
        self._bind_traces()

        # Populate initial local explorer and try loading existing environment/profile defaults gently
        self._autoload_initial_credentials()
        self.refresh_local_explorer()
        self.update_command_preview()

        # Start UI queue poller
        self.root.after(100, self._process_ui_queue)

    def _configure_styles(self) -> None:
        style = ttk.Style(self.root)
        available = style.theme_names()
        for preferred in ("vista", "clam", "alt", "default"):
            if preferred in available:
                style.theme_use(preferred)
                break

        style.configure("Header.TLabel", font=("Segoe UI", 11, "bold"))
        style.configure("SubHeader.TLabel", font=("Segoe UI", 9, "bold"))
        style.configure("Badge.TLabel", font=("Segoe UI", 9), foreground="#1f5f8b")
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))

    def _build_ui(self) -> None:
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --------------------------------------------------------------------
        # 1. AWS Credentials & Attributes Section
        # --------------------------------------------------------------------
        cred_box = ttk.LabelFrame(
            main_frame,
            text=" 1. AWS Configuration Attributes (Access Key ID, Secret Key, Region, Default Format) ",
            padding="10",
        )
        cred_box.pack(fill=tk.X, pady=(0, 8))

        # Row 0: Access Key, Secret Key, Show checkbox
        ttk.Label(cred_box, text="AWS Access Key ID:").grid(row=0, column=0, sticky=tk.W, padx=(0, 6), pady=3)
        ent_ak = ttk.Entry(cred_box, textvariable=self.var_access_key, width=32)
        ent_ak.grid(row=0, column=1, sticky=tk.EW, padx=(0, 14), pady=3)

        ttk.Label(cred_box, text="AWS Secret Access Key:").grid(row=0, column=2, sticky=tk.W, padx=(0, 6), pady=3)
        self.ent_sk = ttk.Entry(cred_box, textvariable=self.var_secret_key, show="*", width=36)
        self.ent_sk.grid(row=0, column=3, sticky=tk.EW, padx=(0, 6), pady=3)

        chk_show = ttk.Checkbutton(
            cred_box,
            text="Show Secret",
            variable=self.var_show_secret,
            command=self._toggle_secret_visibility,
        )
        chk_show.grid(row=0, column=4, sticky=tk.W, padx=(0, 14), pady=3)

        ttk.Label(cred_box, text="Session Token (Optional):").grid(row=0, column=5, sticky=tk.W, padx=(0, 6), pady=3)
        ent_token = ttk.Entry(cred_box, textvariable=self.var_session_token, show="*", width=22)
        ent_token.grid(row=0, column=6, sticky=tk.EW, pady=3)

        # Row 1: Region, Default Output Format, Profile & Persistence Controls
        ttk.Label(cred_box, text="Default Region:").grid(row=1, column=0, sticky=tk.W, padx=(0, 6), pady=3)
        cmb_region = ttk.Combobox(cred_box, textvariable=self.var_region, values=AWS_REGIONS, width=20)
        cmb_region.grid(row=1, column=1, sticky=tk.W, padx=(0, 14), pady=3)

        ttk.Label(cred_box, text="Default Output Format:").grid(row=1, column=2, sticky=tk.W, padx=(0, 6), pady=3)
        cmb_format = ttk.Combobox(
            cred_box,
            textvariable=self.var_output_format,
            values=AWS_OUTPUT_FORMATS,
            width=18,
            state="readonly",
        )
        cmb_format.grid(row=1, column=3, sticky=tk.W, padx=(0, 6), pady=3)

        # Buttons for loading/saving credentials & testing S3 connection
        btn_bar = ttk.Frame(cred_box)
        btn_bar.grid(row=1, column=4, columnspan=3, sticky=tk.E, pady=3)

        ttk.Label(btn_bar, text="Profile:").pack(side=tk.LEFT, padx=(0, 4))
        ttk.Entry(btn_bar, textvariable=self.var_profile, width=10).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_bar, text="Load ~/.aws", command=self.on_load_aws_profile).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_bar, text="Save ~/.aws", command=self.on_save_aws_profile).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_bar, text="Save Local Config", command=self.on_save_local_config).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_bar, text="Connect & List S3 Buckets", command=self.refresh_s3_buckets).pack(
            side=tk.LEFT, padx=(8, 0)
        )

        cred_box.columnconfigure(1, weight=1)
        cred_box.columnconfigure(3, weight=1)
        cred_box.columnconfigure(6, weight=1)

        # --------------------------------------------------------------------
        # 2. Transfer Direction Bar
        # --------------------------------------------------------------------
        dir_box = ttk.LabelFrame(main_frame, text=" 2. Transfer Direction ", padding="8")
        dir_box.pack(fill=tk.X, pady=(0, 8))

        rb_up = ttk.Radiobutton(
            dir_box,
            text="Copy TO S3 Bucket  (Upload: Local Source File/Folder  ->  S3 Target Folder)",
            variable=self.var_direction,
            value="upload",
            command=self._on_direction_changed,
        )
        rb_up.pack(side=tk.LEFT, padx=(8, 24))

        rb_down = ttk.Radiobutton(
            dir_box,
            text="Copy FROM S3 Bucket  (Download: S3 Source Object/Folder  ->  Local Target Folder)",
            variable=self.var_direction,
            value="download",
            command=self._on_direction_changed,
        )
        rb_down.pack(side=tk.LEFT, padx=8)

        aws_cli_status = "Detected" if find_aws_cli() else "Not on PATH (Built-in S3 Engine active)"
        boto_status = "Detected" if HAS_BOTO3 else "Stdlib SigV4 active"
        lbl_env = ttk.Label(
            dir_box,
            text=f"AWS CLI: {aws_cli_status}  |  S3 API: {boto_status}",
            style="Badge.TLabel",
        )
        lbl_env.pack(side=tk.RIGHT, padx=8)

        # --------------------------------------------------------------------
        # 3. Dual Explorer Pane (Local File/Folder Explorer & S3 Bucket Explorer)
        # --------------------------------------------------------------------
        paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        # --- Left Pane: Local Explorer ---
        local_frame = ttk.LabelFrame(
            paned, text=" Local File & Folder Explorer ", padding="8"
        )
        paned.add(local_frame, weight=1)

        local_nav = ttk.Frame(local_frame)
        local_nav.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(local_nav, text="Folder:").pack(side=tk.LEFT, padx=(0, 4))
        ent_local_dir = ttk.Entry(local_nav, textvariable=self.var_local_dir)
        ent_local_dir.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        ent_local_dir.bind("<Return>", lambda _e: self.refresh_local_explorer())
        ttk.Button(local_nav, text="Go", width=4, command=self.refresh_local_explorer).pack(side=tk.LEFT, padx=2)
        ttk.Button(local_nav, text="Up", width=4, command=self.local_navigate_up).pack(side=tk.LEFT, padx=2)
        ttk.Button(local_nav, text="Refresh", width=8, command=self.refresh_local_explorer).pack(side=tk.LEFT, padx=2)

        local_pick_bar = ttk.Frame(local_frame)
        local_pick_bar.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(
            local_pick_bar,
            text="Browse Local File (OS Dialog)...",
            command=self.pick_local_file_dialog,
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(
            local_pick_bar,
            text="Browse Local Folder (OS Dialog)...",
            command=self.pick_local_folder_dialog,
        ).pack(side=tk.LEFT, padx=(0, 6))
        self.btn_use_local = ttk.Button(
            local_pick_bar,
            text="Use Selected as Source",
            command=self.apply_selected_local_item,
        )
        self.btn_use_local.pack(side=tk.RIGHT)

        local_tree_frame = ttk.Frame(local_frame)
        local_tree_frame.pack(fill=tk.BOTH, expand=True)

        self.local_tree = ttk.Treeview(
            local_tree_frame,
            columns=("type", "size", "modified"),
            selectmode="browse",
        )
        self.local_tree.heading("#0", text="Name", anchor=tk.W)
        self.local_tree.heading("type", text="Type", anchor=tk.W)
        self.local_tree.heading("size", text="Size", anchor=tk.E)
        self.local_tree.heading("modified", text="Modified", anchor=tk.W)

        self.local_tree.column("#0", width=230, anchor=tk.W)
        self.local_tree.column("type", width=70, anchor=tk.W)
        self.local_tree.column("size", width=85, anchor=tk.E)
        self.local_tree.column("modified", width=130, anchor=tk.W)

        local_scroll = ttk.Scrollbar(local_tree_frame, orient=tk.VERTICAL, command=self.local_tree.yview)
        self.local_tree.configure(yscrollcommand=local_scroll.set)
        self.local_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        local_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.local_tree.bind("<Double-1>", self.on_local_tree_double_click)
        self.local_tree.bind("<<TreeviewSelect>>", self.on_local_tree_select)

        # --- Right Pane: S3 Bucket & Folder Explorer ---
        s3_frame = ttk.LabelFrame(
            paned, text=" S3 Bucket & Folder Explorer ", padding="8"
        )
        paned.add(s3_frame, weight=1)

        s3_top_nav = ttk.Frame(s3_frame)
        s3_top_nav.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(s3_top_nav, text="Bucket:").pack(side=tk.LEFT, padx=(0, 4))
        self.cmb_bucket = ttk.Combobox(s3_top_nav, textvariable=self.var_s3_bucket, width=22)
        self.cmb_bucket.pack(side=tk.LEFT, padx=(0, 4))
        self.cmb_bucket.bind("<<ComboboxSelected>>", lambda _e: self.on_bucket_selected())
        self.cmb_bucket.bind("<Return>", lambda _e: self.on_bucket_selected())
        ttk.Button(s3_top_nav, text="List Buckets", command=self.refresh_s3_buckets).pack(side=tk.LEFT, padx=2)

        ttk.Label(s3_top_nav, text="Prefix:").pack(side=tk.LEFT, padx=(8, 4))
        ent_prefix = ttk.Entry(s3_top_nav, textvariable=self.var_s3_prefix)
        ent_prefix.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        ent_prefix.bind("<Return>", lambda _e: self.refresh_s3_objects())
        ttk.Button(s3_top_nav, text="Browse", width=7, command=self.refresh_s3_objects).pack(side=tk.LEFT, padx=2)
        ttk.Button(s3_top_nav, text="Up", width=4, command=self.s3_navigate_up).pack(side=tk.LEFT, padx=2)

        s3_action_bar = ttk.Frame(s3_frame)
        s3_action_bar.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(s3_action_bar, text="New S3 Folder...", command=self.create_s3_folder_dialog).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(
            s3_action_bar,
            text="Use Current S3 Folder",
            command=self.use_current_s3_folder,
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(
            s3_action_bar,
            text="Delete .zip Files...",
            command=self.delete_s3_zip_files_dialog,
        ).pack(side=tk.LEFT, padx=(0, 6))
        self.btn_use_s3 = ttk.Button(
            s3_action_bar,
            text="Use Selected S3 Target",
            command=self.apply_selected_s3_item,
        )
        self.btn_use_s3.pack(side=tk.RIGHT)

        s3_tree_frame = ttk.Frame(s3_frame)
        s3_tree_frame.pack(fill=tk.BOTH, expand=True)

        self.s3_tree = ttk.Treeview(
            s3_tree_frame,
            columns=("type", "size", "modified"),
            selectmode="browse",
        )
        self.s3_tree.heading("#0", text="S3 Object / Prefix", anchor=tk.W)
        self.s3_tree.heading("type", text="Type", anchor=tk.W)
        self.s3_tree.heading("size", text="Size", anchor=tk.E)
        self.s3_tree.heading("modified", text="Last Modified", anchor=tk.W)

        self.s3_tree.column("#0", width=240, anchor=tk.W)
        self.s3_tree.column("type", width=70, anchor=tk.W)
        self.s3_tree.column("size", width=85, anchor=tk.E)
        self.s3_tree.column("modified", width=130, anchor=tk.W)

        s3_scroll = ttk.Scrollbar(s3_tree_frame, orient=tk.VERTICAL, command=self.s3_tree.yview)
        self.s3_tree.configure(yscrollcommand=s3_scroll.set)
        self.s3_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        s3_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.s3_tree.bind("<Double-1>", self.on_s3_tree_double_click)
        self.s3_tree.bind("<<TreeviewSelect>>", self.on_s3_tree_select)

        # --------------------------------------------------------------------
        # 4. Transfer Paths, CLI Command Preview & Execution Controls
        # --------------------------------------------------------------------
        transfer_box = ttk.LabelFrame(
            main_frame, text=" 3. Transfer Paths, Options & AWS CLI Command Preview ", padding="10"
        )
        transfer_box.pack(fill=tk.X, pady=(0, 8))

        # Source & Target rows
        self.lbl_source = ttk.Label(transfer_box, text="Source (Local File/Folder):", style="SubHeader.TLabel")
        self.lbl_source.grid(row=0, column=0, sticky=tk.W, padx=(0, 8), pady=3)
        ent_src = ttk.Entry(transfer_box, textvariable=self.var_source_path)
        ent_src.grid(row=0, column=1, sticky=tk.EW, padx=(0, 6), pady=3)
        self.btn_quick_src = ttk.Button(transfer_box, text="Pick File...", command=self.quick_pick_source)
        self.btn_quick_src.grid(row=0, column=2, sticky=tk.EW, pady=3)

        self.lbl_target = ttk.Label(transfer_box, text="Target (S3 Bucket/Folder):", style="SubHeader.TLabel")
        self.lbl_target.grid(row=1, column=0, sticky=tk.W, padx=(0, 8), pady=3)
        ent_dst = ttk.Entry(transfer_box, textvariable=self.var_target_path)
        ent_dst.grid(row=1, column=1, sticky=tk.EW, padx=(0, 6), pady=3)
        self.btn_quick_dst = ttk.Button(
            transfer_box, text="Use Current S3 Folder", command=self.quick_pick_target
        )
        self.btn_quick_dst.grid(row=1, column=2, sticky=tk.EW, pady=3)

        # Options bar
        opt_frame = ttk.Frame(transfer_box)
        opt_frame.grid(row=2, column=0, columnspan=3, sticky=tk.EW, pady=6)

        ttk.Checkbutton(opt_frame, text="Recursive (--recursive)", variable=self.var_recursive).pack(
            side=tk.LEFT, padx=(0, 12)
        )
        ttk.Checkbutton(opt_frame, text="Dry Run (--dryrun)", variable=self.var_dryrun).pack(
            side=tk.LEFT, padx=(0, 12)
        )
        ttk.Checkbutton(opt_frame, text="Only Show Errors", variable=self.var_only_errors).pack(
            side=tk.LEFT, padx=(0, 14)
        )

        ttk.Label(opt_frame, text="Storage Class:").pack(side=tk.LEFT, padx=(0, 4))
        ttk.Combobox(
            opt_frame,
            textvariable=self.var_storage_class,
            values=STORAGE_CLASSES,
            width=17,
            state="readonly",
        ).pack(side=tk.LEFT, padx=(0, 14))

        ttk.Label(opt_frame, text="Exclude:").pack(side=tk.LEFT, padx=(0, 4))
        ttk.Entry(opt_frame, textvariable=self.var_exclude, width=12).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(opt_frame, text="Include:").pack(side=tk.LEFT, padx=(0, 4))
        ttk.Entry(opt_frame, textvariable=self.var_include, width=12).pack(side=tk.LEFT, padx=(0, 14))

        ttk.Label(opt_frame, text="Engine:").pack(side=tk.LEFT, padx=(0, 4))
        ttk.Combobox(
            opt_frame,
            textvariable=self.var_engine,
            values=[
                "Auto (AWS CLI or Built-in S3)",
                "AWS CLI (aws s3 cp)",
                "Built-in S3 API",
            ],
            width=25,
            state="readonly",
        ).pack(side=tk.LEFT)

        # CLI Preview row
        ttk.Label(transfer_box, text="CLI Command:").grid(row=3, column=0, sticky=tk.W, padx=(0, 8), pady=3)
        ent_cmd = ttk.Entry(transfer_box, textvariable=self.var_cmd_preview, state="readonly")
        ent_cmd.grid(row=3, column=1, sticky=tk.EW, padx=(0, 6), pady=3)
        ttk.Button(transfer_box, text="Copy Command", command=self.copy_command_to_clipboard).grid(
            row=3, column=2, sticky=tk.EW, pady=3
        )

        # Action & Progress Bar
        act_frame = ttk.Frame(transfer_box)
        act_frame.grid(row=4, column=0, columnspan=3, sticky=tk.EW, pady=(8, 0))

        self.btn_start = ttk.Button(
            act_frame,
            text="Start S3 CP Transfer",
            style="Accent.TButton",
            command=self.start_transfer,
        )
        self.btn_start.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_cancel = ttk.Button(
            act_frame,
            text="Cancel",
            state=tk.DISABLED,
            command=self.cancel_transfer,
        )
        self.btn_cancel.pack(side=tk.LEFT, padx=(0, 14))

        self.progress_bar = ttk.Progressbar(
            act_frame, variable=self.var_progress, maximum=100.0, mode="determinate", length=260
        )
        self.progress_bar.pack(side=tk.LEFT, padx=(0, 12))

        lbl_status = ttk.Label(act_frame, textvariable=self.var_status)
        lbl_status.pack(side=tk.LEFT, fill=tk.X, expand=True)

        transfer_box.columnconfigure(1, weight=1)

        # --------------------------------------------------------------------
        # 5. Output & Transfer Log
        # --------------------------------------------------------------------
        log_box = ttk.LabelFrame(main_frame, text=" 4. Transfer Activity & Output Log ", padding="6")
        log_box.pack(fill=tk.BOTH, expand=False)

        log_top = ttk.Frame(log_box)
        log_top.pack(fill=tk.X)
        self.txt_log = tk.Text(
            log_top,
            height=7,
            wrap=tk.WORD,
            font=("Consolas", 9),
            bg="#1e1e1e",
            fg="#d4d4d4",
            insertbackground="#ffffff",
        )
        log_scroll = ttk.Scrollbar(log_top, orient=tk.VERTICAL, command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=log_scroll.set)
        self.txt_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        btn_clear_log = ttk.Button(log_box, text="Clear Log", command=lambda: self.txt_log.delete("1.0", tk.END))
        btn_clear_log.pack(anchor=tk.E, pady=(4, 0))

    def _bind_traces(self) -> None:
        for var in (
            self.var_source_path,
            self.var_target_path,
            self.var_region,
            self.var_output_format,
            self.var_recursive,
            self.var_dryrun,
            self.var_only_errors,
            self.var_storage_class,
            self.var_exclude,
            self.var_include,
        ):
            var.trace_add("write", lambda *_args: self.update_command_preview())

    # ========================================================================
    # Helpers & Event Handlers
    # ========================================================================

    def _process_ui_queue(self) -> None:
        try:
            while True:
                fn = self.ui_queue.get_nowait()
                fn()
        except queue.Empty:
            pass
        finally:
            self.root.after(80, self._process_ui_queue)

    def post_ui(self, fn: Callable[[], None]) -> None:
        self.ui_queue.put(fn)

    def log(self, message: str) -> None:
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        line = f"[{timestamp}] {message}\n"

        def _append():
            self.txt_log.insert(tk.END, line)
            self.txt_log.see(tk.END)

        if threading.current_thread() is threading.main_thread():
            _append()
        else:
            self.post_ui(_append)

    def get_credentials_config(self) -> AWSCredentialsConfig:
        return AWSCredentialsConfig(
            access_key_id=self.var_access_key.get().strip(),
            secret_access_key=self.var_secret_key.get().strip(),
            region=self.var_region.get().strip() or "us-east-1",
            output_format=self.var_output_format.get().strip() or "json",
            session_token=self.var_session_token.get().strip(),
            profile_name=self.var_profile.get().strip() or "default",
        )

    def _toggle_secret_visibility(self) -> None:
        self.ent_sk.configure(show="" if self.var_show_secret.get() else "*")

    def _autoload_initial_credentials(self) -> None:
        """Load credentials from local config file, environment variables, or `~/.aws` if present."""
        local_cfg = Path(LOCAL_CONFIG_FILENAME)
        if local_cfg.is_file():
            try:
                data = json.loads(local_cfg.read_text(encoding="utf-8"))
                self.var_access_key.set(data.get("aws_access_key_id", ""))
                self.var_secret_key.set(data.get("aws_secret_access_key", ""))
                self.var_region.set(data.get("region", "us-east-1"))
                self.var_output_format.set(data.get("output_format", "json"))
                self.var_session_token.set(data.get("session_token", ""))
                if data.get("last_bucket"):
                    self.var_s3_bucket.set(data["last_bucket"])
                self.log(f"Loaded settings from {local_cfg.resolve()}")
                return
            except Exception as exc:
                self.log(f"Note: Could not read {LOCAL_CONFIG_FILENAME}: {exc}")

        # Otherwise check environment variables or ~/.aws
        cfg = load_aws_profile_files("default")
        ak = os.environ.get("AWS_ACCESS_KEY_ID", cfg.access_key_id)
        sk = os.environ.get("AWS_SECRET_ACCESS_KEY", cfg.secret_access_key)
        reg = os.environ.get("AWS_DEFAULT_REGION", os.environ.get("AWS_REGION", cfg.region))
        fmt = os.environ.get("AWS_DEFAULT_OUTPUT", cfg.output_format)
        tok = os.environ.get("AWS_SESSION_TOKEN", cfg.session_token)

        if ak:
            self.var_access_key.set(ak)
        if sk:
            self.var_secret_key.set(sk)
        if reg:
            self.var_region.set(reg)
        if fmt in AWS_OUTPUT_FORMATS:
            self.var_output_format.set(fmt)
        if tok:
            self.var_session_token.set(tok)

    def on_load_aws_profile(self) -> None:
        profile = self.var_profile.get().strip() or "default"
        cfg = load_aws_profile_files(profile)
        if not cfg.access_key_id and not cfg.secret_access_key:
            messagebox.showinfo(
                "Load ~/.aws",
                f"No credentials found for profile '{profile}' in ~/.aws/credentials.",
            )
            return
        self.var_access_key.set(cfg.access_key_id)
        self.var_secret_key.set(cfg.secret_access_key)
        self.var_region.set(cfg.region)
        if cfg.output_format in AWS_OUTPUT_FORMATS:
            self.var_output_format.set(cfg.output_format)
        self.var_session_token.set(cfg.session_token)
        self.log(f"Loaded AWS profile '{profile}' from ~/.aws.")

    def on_save_aws_profile(self) -> None:
        try:
            cfg = self.get_credentials_config()
            cfg.validate()
            cred_file, conf_file = save_aws_profile_files(cfg)
            self.log(f"Saved profile '{cfg.profile_name}' to {cred_file} and {conf_file}.")
            messagebox.showinfo(
                "Saved to ~/.aws",
                f"Saved profile '{cfg.profile_name}' to:\n- {cred_file}\n- {conf_file}",
            )
        except Exception as exc:
            messagebox.showerror("Save Error", str(exc))

    def on_save_local_config(self) -> None:
        try:
            cfg = self.get_credentials_config()
            data = {
                "aws_access_key_id": cfg.access_key_id,
                "aws_secret_access_key": cfg.secret_access_key,
                "region": cfg.region,
                "output_format": cfg.output_format,
                "session_token": cfg.session_token,
                "last_bucket": self.var_s3_bucket.get().strip(),
            }
            out_path = Path(LOCAL_CONFIG_FILENAME).resolve()
            out_path.write_text(json.dumps(data, indent=2), encoding="utf-8")
            self.log(f"Saved configuration to {out_path}")
            messagebox.showinfo("Saved Local Config", f"Saved configuration to:\n{out_path}")
        except Exception as exc:
            messagebox.showerror("Save Error", str(exc))

    def _on_direction_changed(self) -> None:
        direction = self.var_direction.get()
        old_src = self.var_source_path.get().strip()
        old_dst = self.var_target_path.get().strip()

        if direction == "upload":
            self.lbl_source.configure(text="Source (Local File/Folder):")
            self.lbl_target.configure(text="Target (S3 Bucket/Folder):")
            self.btn_use_local.configure(text="Use Selected as Source")
            self.btn_use_s3.configure(text="Use Selected as S3 Target")
            self.btn_quick_src.configure(text="Pick Local File...")
            self.btn_quick_dst.configure(text="Use Current S3 Folder")
            if old_src.startswith("s3://") or old_dst and not old_dst.startswith("s3://"):
                self.var_source_path.set(old_dst if not old_dst.startswith("s3://") else "")
                self.var_target_path.set(old_src if old_src.startswith("s3://") else "")
        else:
            self.lbl_source.configure(text="Source (S3 Object/Folder):")
            self.lbl_target.configure(text="Target (Local Folder/File):")
            self.btn_use_local.configure(text="Use Selected as Local Target")
            self.btn_use_s3.configure(text="Use Selected as S3 Source")
            self.btn_quick_src.configure(text="Use Current S3 Folder")
            self.btn_quick_dst.configure(text="Pick Local Folder...")
            if (old_src and not old_src.startswith("s3://")) or old_dst.startswith("s3://"):
                self.var_source_path.set(old_dst if old_dst.startswith("s3://") else "")
                self.var_target_path.set(old_src if not old_src.startswith("s3://") else "")

        self.update_command_preview()

    def update_command_preview(self) -> None:
        src = self.var_source_path.get().strip() or "<source>"
        dst = self.var_target_path.get().strip() or "<target>"
        args = build_aws_cp_command_args(
            source=src,
            target=dst,
            region=self.var_region.get(),
            output_format=self.var_output_format.get(),
            recursive=self.var_recursive.get(),
            dryrun=self.var_dryrun.get(),
            only_show_errors=self.var_only_errors.get(),
            storage_class=self.var_storage_class.get(),
            exclude_pattern=self.var_exclude.get(),
            include_pattern=self.var_include.get(),
        )
        self.var_cmd_preview.set(format_command_preview(args))

    def copy_command_to_clipboard(self) -> None:
        cmd = self.var_cmd_preview.get()
        self.root.clipboard_clear()
        self.root.clipboard_append(cmd)
        self.log(f"Copied CLI command to clipboard: {cmd}")

    # ========================================================================
    # Local Explorer Logic
    # ========================================================================

    def refresh_local_explorer(self) -> None:
        raw_dir = self.var_local_dir.get().strip() or str(Path.cwd())
        target_dir = Path(raw_dir).expanduser()
        if not target_dir.is_dir():
            messagebox.showwarning("Invalid Directory", f"Local directory does not exist:\n{target_dir}")
            return

        resolved = str(target_dir.resolve())
        self.var_local_dir.set(resolved)

        for item in self.local_tree.get_children():
            self.local_tree.delete(item)

        try:
            entries = list(os.scandir(resolved))
        except PermissionError as exc:
            self.log(f"Permission denied listing {resolved}: {exc}")
            return

        # Sort directories first, then files alphabetically
        entries.sort(key=lambda e: (not e.is_dir(follow_symlinks=False), e.name.lower()))

        for entry in entries:
            try:
                is_dir = entry.is_dir(follow_symlinks=False)
                stat = entry.stat(follow_symlinks=False)
                mtime = datetime.datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M")
                size_str = "" if is_dir else format_size(stat.st_size)
                type_str = "Folder" if is_dir else "File"
                self.local_tree.insert(
                    "",
                    tk.END,
                    iid=entry.path,
                    text=entry.name,
                    values=(type_str, size_str, mtime),
                )
            except OSError:
                continue

    def local_navigate_up(self) -> None:
        curr = Path(self.var_local_dir.get().strip() or str(Path.cwd())).resolve()
        parent = curr.parent
        if parent != curr:
            self.var_local_dir.set(str(parent))
            self.refresh_local_explorer()

    def on_local_tree_double_click(self, _event) -> None:
        sel = self.local_tree.selection()
        if not sel:
            return
        path_str = sel[0]
        if os.path.isdir(path_str):
            self.var_local_dir.set(path_str)
            self.refresh_local_explorer()
        else:
            self.apply_selected_local_item()

    def on_local_tree_select(self, _event) -> None:
        self.apply_selected_local_item()

    def apply_selected_local_item(self) -> None:
        sel = self.local_tree.selection()
        chosen_path = sel[0] if sel else self.var_local_dir.get().strip()
        if not chosen_path:
            return

        if self.var_direction.get() == "upload":
            self.var_source_path.set(chosen_path)
            self.var_recursive.set(os.path.isdir(chosen_path))
        else:
            # In download mode, target should be a local folder (or file path)
            if os.path.isfile(chosen_path):
                chosen_path = os.path.dirname(chosen_path)
            self.var_target_path.set(chosen_path)

    def pick_local_file_dialog(self) -> None:
        initial = self.var_local_dir.get().strip() or str(Path.cwd())
        if self.var_direction.get() == "upload":
            chosen = filedialog.askopenfilename(title="Select Local File to Upload", initialdir=initial)
            if chosen:
                norm = os.path.normpath(chosen)
                self.var_local_dir.set(os.path.dirname(norm))
                self.refresh_local_explorer()
                self.var_source_path.set(norm)
                self.var_recursive.set(False)
        else:
            chosen = filedialog.asksaveasfilename(title="Select Target Local File Path", initialdir=initial)
            if chosen:
                norm = os.path.normpath(chosen)
                self.var_local_dir.set(os.path.dirname(norm))
                self.refresh_local_explorer()
                self.var_target_path.set(norm)

    def pick_local_folder_dialog(self) -> None:
        initial = self.var_local_dir.get().strip() or str(Path.cwd())
        title = (
            "Select Local Folder to Upload Recursively"
            if self.var_direction.get() == "upload"
            else "Select Target Local Folder for Download"
        )
        chosen = filedialog.askdirectory(title=title, initialdir=initial)
        if chosen:
            norm = os.path.normpath(chosen)
            self.var_local_dir.set(norm)
            self.refresh_local_explorer()
            if self.var_direction.get() == "upload":
                self.var_source_path.set(norm)
                self.var_recursive.set(True)
            else:
                self.var_target_path.set(norm)

    def quick_pick_source(self) -> None:
        if self.var_direction.get() == "upload":
            self.pick_local_file_dialog()
        else:
            self.use_current_s3_folder()

    def quick_pick_target(self) -> None:
        if self.var_direction.get() == "upload":
            self.use_current_s3_folder()
        else:
            self.pick_local_folder_dialog()

    # ========================================================================
    # S3 Explorer Logic
    # ========================================================================

    def refresh_s3_buckets(self) -> None:
        try:
            creds = self.get_credentials_config()
            creds.validate()
        except ValueError as exc:
            messagebox.showwarning("AWS Credentials Required", str(exc))
            return

        self.var_status.set("Connecting to S3 and listing buckets...")
        self.log("Listing S3 buckets...")

        def _worker():
            try:
                svc = S3Service(creds)
                buckets = svc.list_buckets()

                def _on_done():
                    self.cmb_bucket["values"] = buckets
                    if buckets and not self.var_s3_bucket.get().strip():
                        self.var_s3_bucket.set(buckets[0])
                        self.refresh_s3_objects()
                    self.var_status.set(f"Found {len(buckets)} bucket(s).")
                    self.log(f"Successfully listed {len(buckets)} S3 bucket(s).")

                self.post_ui(_on_done)
            except Exception as exc:
                self.log(f"Error listing buckets: {exc}")
                self.post_ui(
                    lambda: (
                        self.var_status.set("Could not list buckets (you can still type a bucket name directly)."),
                        messagebox.showerror(
                            "S3 List Buckets Failed",
                            f"{exc}\n\nTip: If your IAM user only has permission for a specific bucket, "
                            "type the bucket name directly in the 'Bucket' box and click 'Browse'.",
                        ),
                    )
                )

        threading.Thread(target=_worker, daemon=True).start()

    def on_bucket_selected(self) -> None:
        raw = self.var_s3_bucket.get().strip()
        if raw.startswith("s3://"):
            try:
                b, p = parse_s3_uri(raw)
                self.var_s3_bucket.set(b)
                self.var_s3_prefix.set(p)
            except ValueError:
                pass
        else:
            self.var_s3_prefix.set("")
        self.refresh_s3_objects()

    def refresh_s3_objects(self) -> None:
        bucket = self.var_s3_bucket.get().strip()
        if bucket.startswith("s3://"):
            try:
                bucket, extracted_prefix = parse_s3_uri(bucket)
                self.var_s3_bucket.set(bucket)
                if extracted_prefix and not self.var_s3_prefix.get().strip():
                    self.var_s3_prefix.set(extracted_prefix)
            except ValueError as exc:
                messagebox.showwarning("Invalid S3 URI", str(exc))
                return

        if not bucket:
            messagebox.showwarning("Bucket Required", "Please select or enter an S3 bucket name first.")
            return

        try:
            creds = self.get_credentials_config()
            creds.validate()
        except ValueError as exc:
            messagebox.showwarning("AWS Credentials Required", str(exc))
            return

        prefix = self.var_s3_prefix.get().strip().lstrip("/")
        if prefix and not prefix.endswith("/"):
            prefix += "/"
        self.var_s3_prefix.set(prefix)

        uri_display = build_s3_uri(bucket, prefix)
        self.var_status.set(f"Listing {uri_display} ...")
        self.log(f"Browsing {uri_display}")

        def _worker():
            try:
                svc = S3Service(creds)
                entries = svc.list_objects(bucket=bucket, prefix=prefix, delimiter="/")

                def _populate():
                    for item in self.s3_tree.get_children():
                        self.s3_tree.delete(item)
                    for e in entries:
                        type_str = "Folder" if e.is_dir else "Object"
                        size_str = "" if e.is_dir else format_size(e.size)
                        self.s3_tree.insert(
                            "",
                            tk.END,
                            iid=e.key,
                            text=f"[DIR] {e.name}/" if e.is_dir else e.name,
                            values=(type_str, size_str, e.last_modified),
                        )
                    self.var_status.set(f"Loaded {len(entries)} item(s) in {uri_display}")
                    # If upload mode and target is empty, default target to current S3 folder
                    if self.var_direction.get() == "upload" and not self.var_target_path.get().strip():
                        self.var_target_path.set(uri_display)

                self.post_ui(_populate)
            except Exception as exc:
                self.log(f"Error browsing {uri_display}: {exc}")
                self.post_ui(
                    lambda: (
                        self.var_status.set("Error listing S3 objects."),
                        messagebox.showerror("S3 Browse Error", str(exc)),
                    )
                )

        threading.Thread(target=_worker, daemon=True).start()

    def s3_navigate_up(self) -> None:
        prefix = self.var_s3_prefix.get().strip().strip("/")
        if not prefix:
            return
        if "/" in prefix:
            parent = prefix.rsplit("/", 1)[0] + "/"
        else:
            parent = ""
        self.var_s3_prefix.set(parent)
        self.refresh_s3_objects()

    def on_s3_tree_double_click(self, _event) -> None:
        sel = self.s3_tree.selection()
        if not sel:
            return
        key = sel[0]
        item_vals = self.s3_tree.item(key, "values")
        is_dir = bool(item_vals and item_vals[0] == "Folder")
        if is_dir:
            self.var_s3_prefix.set(key)
            self.refresh_s3_objects()
        else:
            self.apply_selected_s3_item()

    def on_s3_tree_select(self, _event) -> None:
        self.apply_selected_s3_item()

    def use_current_s3_folder(self) -> None:
        bucket = self.var_s3_bucket.get().strip()
        if not bucket:
            messagebox.showwarning("No S3 Bucket", "Please enter or select an S3 bucket first.")
            return
        prefix = self.var_s3_prefix.get().strip().lstrip("/")
        if prefix and not prefix.endswith("/"):
            prefix += "/"
        uri = build_s3_uri(bucket, prefix)
        if self.var_direction.get() == "upload":
            self.var_target_path.set(uri)
        else:
            self.var_source_path.set(uri)
            self.var_recursive.set(True)

    def apply_selected_s3_item(self) -> None:
        bucket = self.var_s3_bucket.get().strip()
        if not bucket:
            return
        sel = self.s3_tree.selection()
        if not sel:
            self.use_current_s3_folder()
            return

        key = sel[0]
        item_vals = self.s3_tree.item(key, "values")
        is_dir = bool(item_vals and item_vals[0] == "Folder")
        uri = build_s3_uri(bucket, key)

        if self.var_direction.get() == "upload":
            # Target for upload should be a folder/prefix (or specific object key)
            if is_dir:
                self.var_target_path.set(uri)
            else:
                parent_prefix = key.rsplit("/", 1)[0] + "/" if "/" in key else ""
                self.var_target_path.set(build_s3_uri(bucket, parent_prefix))
        else:
            # Source for download can be either a file (object) or a folder (prefix)
            self.var_source_path.set(uri)
            self.var_recursive.set(is_dir)

    def create_s3_folder_dialog(self) -> None:
        bucket = self.var_s3_bucket.get().strip()
        if not bucket:
            messagebox.showwarning("Bucket Required", "Please select or enter an S3 bucket first.")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("Create S3 Folder (Prefix)")
        dlg.geometry("380x130")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        frame = ttk.Frame(dlg, padding="12")
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text="New folder name:").pack(anchor=tk.W)
        var_name = tk.StringVar()
        ent = ttk.Entry(frame, textvariable=var_name, width=40)
        ent.pack(fill=tk.X, pady=8)
        ent.focus_set()

        def _create():
            folder_name = var_name.get().strip().strip("/")
            if not folder_name:
                return
            curr_prefix = self.var_s3_prefix.get().strip().lstrip("/")
            if curr_prefix and not curr_prefix.endswith("/"):
                curr_prefix += "/"
            new_prefix = f"{curr_prefix}{folder_name}/"
            dlg.destroy()

            def _worker():
                try:
                    svc = S3Service(self.get_credentials_config())
                    svc.create_folder(bucket, new_prefix)
                    self.log(f"Created S3 folder s3://{bucket}/{new_prefix}")
                    self.post_ui(self.refresh_s3_objects)
                except Exception as exc:
                    self.log(f"Failed creating folder: {exc}")
                    self.post_ui(lambda: messagebox.showerror("Create Folder Error", str(exc)))

            threading.Thread(target=_worker, daemon=True).start()

        btn_row = ttk.Frame(frame)
        btn_row.pack(fill=tk.X)
        ttk.Button(btn_row, text="Create", command=_create).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(btn_row, text="Cancel", command=dlg.destroy).pack(side=tk.RIGHT)
        ent.bind("<Return>", lambda _e: _create())

    def delete_s3_zip_files_dialog(self) -> None:
        """Open a modal dialog that scans for `.zip` files in the current S3 bucket/prefix and deletes them."""
        bucket = self.var_s3_bucket.get().strip()
        if bucket.startswith("s3://"):
            try:
                bucket, extracted_prefix = parse_s3_uri(bucket)
                self.var_s3_bucket.set(bucket)
                if extracted_prefix and not self.var_s3_prefix.get().strip():
                    self.var_s3_prefix.set(extracted_prefix)
            except ValueError as exc:
                messagebox.showwarning("Invalid S3 URI", str(exc))
                return

        if not bucket:
            messagebox.showwarning("Bucket Required", "Please select or enter an S3 bucket name first.")
            return

        try:
            creds = self.get_credentials_config()
            creds.validate()
        except ValueError as exc:
            messagebox.showwarning("AWS Credentials Required", str(exc))
            return

        prefix = self.var_s3_prefix.get().strip().lstrip("/")
        if prefix and not prefix.endswith("/"):
            prefix += "/"
        target_uri = build_s3_uri(bucket, prefix)

        dlg = tk.Toplevel(self.root)
        dlg.title(f"Delete .zip Files in {target_uri}")
        dlg.geometry("760x500")
        dlg.minsize(640, 400)
        dlg.transient(self.root)
        dlg.grab_set()

        frame = ttk.Frame(dlg, padding="12")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            frame,
            text=f"Target S3 Location: {target_uri}",
            style="SubHeader.TLabel",
        ).pack(anchor=tk.W, pady=(0, 6))

        opt_row = ttk.Frame(frame)
        opt_row.pack(fill=tk.X, pady=(0, 6))

        var_inc_subfolders = tk.BooleanVar(value=False)
        var_zip_dryrun = tk.BooleanVar(value=self.var_dryrun.get())
        var_summary = tk.StringVar(value="Scanning for .zip files...")
        var_rm_cmd = tk.StringVar()

        matching_entries: List[S3Entry] = []

        def _update_rm_preview() -> None:
            args = build_aws_rm_zip_command_args(
                s3_uri=target_uri,
                region=creds.region,
                output_format=creds.output_format,
                recursive=var_inc_subfolders.get(),
                dryrun=var_zip_dryrun.get(),
            )
            var_rm_cmd.set(format_command_preview(args))

        chk_sub = ttk.Checkbutton(
            opt_row,
            text="Include subfolders recursively",
            variable=var_inc_subfolders,
            command=lambda: (_update_rm_preview(), _scan_zips()),
        )
        chk_sub.pack(side=tk.LEFT, padx=(0, 16))

        chk_dry = ttk.Checkbutton(
            opt_row,
            text="Dry Run (--dryrun)",
            variable=var_zip_dryrun,
            command=_update_rm_preview,
        )
        chk_dry.pack(side=tk.LEFT, padx=(0, 16))

        btn_rescan = ttk.Button(opt_row, text="Rescan", command=lambda: _scan_zips())
        btn_rescan.pack(side=tk.RIGHT)

        # CLI Preview inside dialog
        cmd_row = ttk.Frame(frame)
        cmd_row.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(cmd_row, text="Equivalent CLI:").pack(side=tk.LEFT, padx=(0, 6))
        ent_rm = ttk.Entry(cmd_row, textvariable=var_rm_cmd, state="readonly")
        ent_rm.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Matching .zip files table
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        zip_tree = ttk.Treeview(
            tree_frame,
            columns=("size", "modified"),
            selectmode="extended",
        )
        zip_tree.heading("#0", text=".zip Object Key", anchor=tk.W)
        zip_tree.heading("size", text="Size", anchor=tk.E)
        zip_tree.heading("modified", text="Last Modified", anchor=tk.W)
        zip_tree.column("#0", width=460, anchor=tk.W)
        zip_tree.column("size", width=100, anchor=tk.E)
        zip_tree.column("modified", width=140, anchor=tk.W)

        zip_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=zip_tree.yview)
        zip_tree.configure(yscrollcommand=zip_scroll.set)
        zip_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        zip_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Bottom summary & action bar
        bottom_bar = ttk.Frame(frame)
        bottom_bar.pack(fill=tk.X)

        ttk.Label(bottom_bar, textvariable=var_summary, style="SubHeader.TLabel").pack(side=tk.LEFT)

        btn_close = ttk.Button(bottom_bar, text="Close", command=dlg.destroy)
        btn_close.pack(side=tk.RIGHT, padx=(6, 0))

        btn_delete = ttk.Button(
            bottom_bar,
            text="Delete Matching .zip Files",
            style="Accent.TButton",
            state=tk.DISABLED,
        )
        btn_delete.pack(side=tk.RIGHT)

        def _scan_zips() -> None:
            btn_delete.configure(state=tk.DISABLED)
            btn_rescan.configure(state=tk.DISABLED)
            var_summary.set("Scanning S3 for .zip files...")
            for item in zip_tree.get_children():
                zip_tree.delete(item)

            delim = "" if var_inc_subfolders.get() else "/"

            def _worker():
                try:
                    svc = S3Service(creds)
                    entries = svc.list_objects(bucket=bucket, prefix=prefix, delimiter=delim)
                    zips = filter_zip_entries(entries)

                    def _populate():
                        if not dlg.winfo_exists():
                            return
                        matching_entries.clear()
                        matching_entries.extend(zips)
                        total_bytes = sum(e.size or 0 for e in zips)
                        for e in zips:
                            zip_tree.insert(
                                "",
                                tk.END,
                                iid=e.key,
                                text=e.key,
                                values=(format_size(e.size), e.last_modified),
                            )
                        scope_str = "including subfolders" if delim == "" else "in current folder"
                        var_summary.set(
                            f"Found {len(zips)} .zip file(s) {scope_str} (Total: {format_size(total_bytes)})"
                        )
                        btn_rescan.configure(state=tk.NORMAL)
                        if zips:
                            btn_delete.configure(state=tk.NORMAL)

                    self.post_ui(_populate)
                except Exception as exc:
                    self.log(f"Error scanning for .zip files in {target_uri}: {exc}")
                    self.post_ui(
                        lambda: (
                            var_summary.set(f"Scan failed: {exc}"),
                            btn_rescan.configure(state=tk.NORMAL),
                        )
                    )

            threading.Thread(target=_worker, daemon=True).start()

        def _execute_delete() -> None:
            if not matching_entries:
                return
            count = len(matching_entries)
            total_bytes = sum(e.size or 0 for e in matching_entries)
            is_dry = var_zip_dryrun.get()

            if not is_dry:
                confirm = messagebox.askyesno(
                    "Confirm Delete .zip Files",
                    f"Are you sure you want to permanently delete {count} .zip file(s) "
                    f"({format_size(total_bytes)}) from {target_uri}?",
                    parent=dlg,
                )
                if not confirm:
                    return

            btn_delete.configure(state=tk.DISABLED)
            btn_rescan.configure(state=tk.DISABLED)
            chk_sub.configure(state=tk.DISABLED)
            chk_dry.configure(state=tk.DISABLED)

            to_delete = list(matching_entries)

            def _del_worker():
                svc = S3Service(creds)
                deleted_count = 0
                try:
                    self.log(
                        f"Starting {'DRY RUN ' if is_dry else ''}deletion of {len(to_delete)} .zip file(s) in {target_uri}"
                    )
                    for idx, entry in enumerate(to_delete, start=1):
                        obj_uri = build_s3_uri(bucket, entry.key)
                        if is_dry:
                            self.log(f"(dryrun) delete: {obj_uri}")
                        else:
                            self.log(f"[{idx}/{len(to_delete)}] delete: {obj_uri}")
                            svc.delete_object(bucket=bucket, key=entry.key)
                        deleted_count += 1
                        pct = (idx / len(to_delete)) * 100.0
                        msg = f"{'Simulated' if is_dry else 'Deleted'} [{idx}/{len(to_delete)}]: {entry.name}"
                        self.post_ui(
                            lambda p=pct, m=msg: (
                                self.var_progress.set(p),
                                self.var_status.set(m),
                                var_summary.set(m) if dlg.winfo_exists() else None,
                            )
                        )

                    done_msg = (
                        f"Dry run complete: {deleted_count} .zip file(s) would be deleted."
                        if is_dry
                        else f"Successfully deleted {deleted_count} .zip file(s) from {target_uri}."
                    )
                    self.log(done_msg)

                    def _on_complete():
                        self.var_status.set(done_msg)
                        self.refresh_s3_objects()
                        if dlg.winfo_exists():
                            messagebox.showinfo("Delete .zip Files", done_msg, parent=dlg)
                            dlg.destroy()

                    self.post_ui(_on_complete)
                except Exception as exc:
                    self.log(f"Error deleting .zip files: {exc}")
                    self.post_ui(
                        lambda: (
                            messagebox.showerror("Delete Error", str(exc), parent=dlg if dlg.winfo_exists() else self.root),
                            btn_rescan.configure(state=tk.NORMAL) if dlg.winfo_exists() else None,
                            chk_sub.configure(state=tk.NORMAL) if dlg.winfo_exists() else None,
                            chk_dry.configure(state=tk.NORMAL) if dlg.winfo_exists() else None,
                        )
                    )

            threading.Thread(target=_del_worker, daemon=True).start()

        btn_delete.configure(command=_execute_delete)
        _update_rm_preview()
        _scan_zips()

    # ========================================================================
    # Transfer Execution (AWS CLI & Built-in S3 Engine)
    # ========================================================================

    def start_transfer(self) -> None:
        if self.is_transferring:
            return

        try:
            creds = self.get_credentials_config()
            creds.validate()
        except ValueError as exc:
            messagebox.showwarning("Missing AWS Configuration", str(exc))
            return

        source = self.var_source_path.get().strip()
        target = self.var_target_path.get().strip()
        if not source or not target:
            messagebox.showwarning("Missing Paths", "Please specify both a Source path and a Target path.")
            return

        direction = self.var_direction.get()
        if direction == "upload":
            if not os.path.exists(source):
                messagebox.showerror("Invalid Local Source", f"Local source does not exist:\n{source}")
                return
            if not target.startswith("s3://"):
                messagebox.showerror("Invalid S3 Target", "Target path must start with s3:// when uploading.")
                return
            if os.path.isdir(source) and not self.var_recursive.get():
                if messagebox.askyesno(
                    "Directory Selected",
                    "Source is a folder. Enable '--recursive' to upload all files inside it?",
                ):
                    self.var_recursive.set(True)
                else:
                    return
        else:
            if not source.startswith("s3://"):
                messagebox.showerror("Invalid S3 Source", "Source path must start with s3:// when downloading.")
                return

        self.is_transferring = True
        self.cancel_event.clear()
        self.btn_start.configure(state=tk.DISABLED)
        self.btn_cancel.configure(state=tk.NORMAL)
        self.var_progress.set(0.0)
        self.var_status.set("Starting transfer...")

        engine_choice = self.var_engine.get()
        aws_bin = find_aws_cli()

        use_cli = False
        if engine_choice == "AWS CLI (aws s3 cp)":
            if not aws_bin:
                messagebox.showerror(
                    "AWS CLI Not Found",
                    "The 'aws' CLI executable was not found on PATH.\n\n"
                    "Switch Engine to 'Auto' or 'Built-in S3 API' to transfer without AWS CLI.",
                )
                self._finish_transfer_ui("Aborted: AWS CLI not installed.")
                return
            use_cli = True
        elif engine_choice.startswith("Auto"):
            use_cli = aws_bin is not None

        if use_cli and aws_bin:
            threading.Thread(
                target=self._run_cli_transfer_worker,
                args=(creds, aws_bin, source, target),
                daemon=True,
            ).start()
        else:
            threading.Thread(
                target=self._run_builtin_transfer_worker,
                args=(creds, direction, source, target),
                daemon=True,
            ).start()

    def cancel_transfer(self) -> None:
        if not self.is_transferring:
            return
        self.cancel_event.set()
        if self.active_process and self.active_process.poll() is None:
            try:
                self.active_process.terminate()
            except Exception:
                pass
        self.log("Cancellation requested by user...")
        self.var_status.set("Cancelling transfer...")

    def _finish_transfer_ui(self, status_text: str) -> None:
        def _reset():
            self.is_transferring = False
            self.active_process = None
            self.btn_start.configure(state=tk.NORMAL)
            self.btn_cancel.configure(state=tk.DISABLED)
            self.var_status.set(status_text)

        self.post_ui(_reset)

    def _run_cli_transfer_worker(
        self,
        creds: AWSCredentialsConfig,
        aws_bin: str,
        source: str,
        target: str,
    ) -> None:
        args = build_aws_cp_command_args(
            source=source,
            target=target,
            region=creds.region,
            output_format=creds.output_format,
            recursive=self.var_recursive.get(),
            dryrun=self.var_dryrun.get(),
            only_show_errors=self.var_only_errors.get(),
            storage_class=self.var_storage_class.get(),
            exclude_pattern=self.var_exclude.get(),
            include_pattern=self.var_include.get(),
            aws_executable=aws_bin,
        )
        cmd_str = format_command_preview(args)
        self.log(f"Executing AWS CLI: {cmd_str}")

        env = creds.to_env_dict()
        creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0

        try:
            self.post_ui(lambda: self.progress_bar.configure(mode="indeterminate"))
            self.post_ui(lambda: self.progress_bar.start(12))

            proc = subprocess.Popen(
                args,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                env=env,
                text=True,
                bufsize=1,
                universal_newlines=True,
                creationflags=creationflags,
            )
            self.active_process = proc

            assert proc.stdout is not None
            buffer = ""
            while True:
                if self.cancel_event.is_set():
                    proc.terminate()
                    break
                ch = proc.stdout.read(1)
                if not ch:
                    if proc.poll() is not None:
                        break
                    continue
                if ch in ("\r", "\n"):
                    line = buffer.strip()
                    buffer = ""
                    if line:
                        if ch == "\r" and "Completed" in line:
                            self.post_ui(lambda l=line: self.var_status.set(l))
                        else:
                            self.log(line)
                else:
                    buffer += ch

            if buffer.strip():
                self.log(buffer.strip())

            rc = proc.wait()
            self.post_ui(lambda: self.progress_bar.stop())
            self.post_ui(lambda: self.progress_bar.configure(mode="determinate"))

            if self.cancel_event.is_set():
                self.log("AWS CLI transfer cancelled.")
                self._finish_transfer_ui("Transfer cancelled.")
            elif rc == 0:
                self.post_ui(lambda: self.var_progress.set(100.0))
                self.log("AWS CLI transfer completed successfully.")
                self._finish_transfer_ui("Transfer completed successfully.")
                self._refresh_explorers_after_transfer()
            else:
                self.log(f"AWS CLI exited with code {rc}.")
                self._finish_transfer_ui(f"Transfer failed (exit code {rc}).")
        except Exception as exc:
            self.post_ui(lambda: self.progress_bar.stop())
            self.post_ui(lambda: self.progress_bar.configure(mode="determinate"))
            self.log(f"AWS CLI execution error: {exc}")
            self._finish_transfer_ui("Transfer failed.")

    def _run_builtin_transfer_worker(
        self,
        creds: AWSCredentialsConfig,
        direction: str,
        source: str,
        target: str,
    ) -> None:
        svc = S3Service(creds)
        recursive = self.var_recursive.get()
        dryrun = self.var_dryrun.get()
        storage_class = self.var_storage_class.get()
        start_time = time.time()

        try:
            if direction == "upload":
                bucket, target_prefix = parse_s3_uri(target)
                local_path = os.path.abspath(source)

                if os.path.isfile(local_path):
                    # If target ends with '/' or is empty key, append filename
                    if not target_prefix or target_prefix.endswith("/"):
                        dest_key = f"{target_prefix}{os.path.basename(local_path)}"
                    else:
                        dest_key = target_prefix
                    files_to_upload = [(local_path, dest_key)]
                else:
                    if not recursive:
                        raise ValueError("Source is a directory; enable Recursive (--recursive) to upload folders.")
                    base_prefix = target_prefix.rstrip("/")
                    if base_prefix:
                        base_prefix += "/"
                    files_to_upload = []
                    for root_dir, _dirs, files in os.walk(local_path):
                        for fname in files:
                            full_p = os.path.join(root_dir, fname)
                            rel_p = os.path.relpath(full_p, local_path).replace("\\", "/")
                            files_to_upload.append((full_p, f"{base_prefix}{rel_p}"))

                total_files = len(files_to_upload)
                if total_files == 0:
                    self.log("No files found to upload.")
                    self._finish_transfer_ui("No files found to upload.")
                    return

                self.log(
                    f"Starting built-in S3 upload of {total_files} file(s) to s3://{bucket}/ "
                    f"(Region: {creds.region}, Format: {creds.output_format})"
                )

                for idx, (l_file, s3_key) in enumerate(files_to_upload, start=1):
                    if self.cancel_event.is_set():
                        raise InterruptedError("Transfer cancelled by user.")

                    dest_uri = build_s3_uri(bucket, s3_key)
                    if dryrun:
                        self.log(f"(dryrun) upload: {l_file} to {dest_uri}")
                        pct = (idx / total_files) * 100.0
                        self.post_ui(lambda p=pct: self.var_progress.set(p))
                        continue

                    self.log(f"[{idx}/{total_files}] upload: {l_file} -> {dest_uri}")

                    def _prog(sent: int, total: int, file_index: int = idx) -> None:
                        file_frac = (sent / total) if total > 0 else 1.0
                        overall_pct = ((file_index - 1 + file_frac) / total_files) * 100.0
                        elapsed = max(0.001, time.time() - start_time)
                        speed = format_size(int(sent / elapsed)) + "/s"
                        status = (
                            f"Uploading [{file_index}/{total_files}]: "
                            f"{format_size(sent)} / {format_size(total)} ({overall_pct:.1f}%) - {speed}"
                        )
                        self.post_ui(lambda p=overall_pct, s=status: (self.var_progress.set(p), self.var_status.set(s)))

                    svc.upload_file(
                        local_path=l_file,
                        bucket=bucket,
                        key=s3_key,
                        storage_class=storage_class,
                        progress_cb=_prog,
                        cancel_event=self.cancel_event,
                    )

            else:
                # Download: S3 -> Local
                bucket, src_key = parse_s3_uri(source)
                local_target = os.path.abspath(target)

                if not recursive and not src_key.endswith("/"):
                    if not src_key:
                        raise ValueError("Specify an S3 object key or enable Recursive (--recursive) to download a bucket/prefix.")
                    if os.path.isdir(local_target) or target.endswith(("/", "\\")):
                        dest_file = os.path.join(local_target, os.path.basename(src_key))
                    else:
                        dest_file = local_target
                    items_to_download = [(src_key, dest_file)]
                else:
                    # List all objects recursively under src_key prefix
                    entries = svc.list_objects(bucket=bucket, prefix=src_key, delimiter="")
                    items_to_download = []
                    for entry in entries:
                        if entry.is_dir or entry.key.endswith("/"):
                            continue
                        rel = entry.key[len(src_key) :].lstrip("/") if entry.key.startswith(src_key) else entry.key
                        if not rel:
                            rel = os.path.basename(entry.key)
                        dest_file = os.path.join(local_target, *rel.split("/"))
                        items_to_download.append((entry.key, dest_file))

                total_files = len(items_to_download)
                if total_files == 0:
                    self.log(f"No matching S3 objects found at {source}")
                    self._finish_transfer_ui("No matching S3 objects found.")
                    return

                self.log(
                    f"Starting built-in S3 download of {total_files} object(s) from s3://{bucket}/ "
                    f"(Region: {creds.region}, Format: {creds.output_format})"
                )

                for idx, (s3_key, l_file) in enumerate(items_to_download, start=1):
                    if self.cancel_event.is_set():
                        raise InterruptedError("Transfer cancelled by user.")

                    src_uri = build_s3_uri(bucket, s3_key)
                    if dryrun:
                        self.log(f"(dryrun) download: {src_uri} to {l_file}")
                        pct = (idx / total_files) * 100.0
                        self.post_ui(lambda p=pct: self.var_progress.set(p))
                        continue

                    self.log(f"[{idx}/{total_files}] download: {src_uri} -> {l_file}")

                    def _prog(recvd: int, total: int, file_index: int = idx) -> None:
                        file_frac = (recvd / total) if total > 0 else 1.0
                        overall_pct = ((file_index - 1 + file_frac) / total_files) * 100.0
                        elapsed = max(0.001, time.time() - start_time)
                        speed = format_size(int(recvd / elapsed)) + "/s"
                        status = (
                            f"Downloading [{file_index}/{total_files}]: "
                            f"{format_size(recvd)} / {format_size(total)} ({overall_pct:.1f}%) - {speed}"
                        )
                        self.post_ui(lambda p=overall_pct, s=status: (self.var_progress.set(p), self.var_status.set(s)))

                    svc.download_file(
                        bucket=bucket,
                        key=s3_key,
                        local_path=l_file,
                        progress_cb=_prog,
                        cancel_event=self.cancel_event,
                    )

            self.post_ui(lambda: self.var_progress.set(100.0))
            elapsed_total = time.time() - start_time
            self.log(f"Transfer finished successfully in {elapsed_total:.1f}s.")
            self._finish_transfer_ui(f"Transfer completed ({total_files} file(s) in {elapsed_total:.1f}s).")
            self._refresh_explorers_after_transfer()

        except InterruptedError as exc:
            self.log(str(exc))
            self._finish_transfer_ui("Transfer cancelled.")
        except Exception as exc:
            self.log(f"Transfer failed: {exc}")
            self._finish_transfer_ui("Transfer failed.")
            self.post_ui(lambda: messagebox.showerror("Transfer Error", str(exc)))

    def _refresh_explorers_after_transfer(self) -> None:
        def _do():
            self.refresh_local_explorer()
            if self.var_s3_bucket.get().strip():
                self.refresh_s3_objects()

        self.post_ui(_do)


def main() -> None:
    root = tk.Tk()
    app = AwsCpTransferApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
