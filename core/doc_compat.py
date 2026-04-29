"""
Legacy .doc → .docx auto-conversion.

`python-docx` only reads the modern Open-XML `.docx` format. Older binary
`.doc` files (Word 97-2003) need to be converted first. We delegate that to
LibreOffice headless — same engine used by `pdf_converter`.

Cache: converted output goes to a stable path under `tempfile.gettempdir()`
keyed by the input file's absolute path + mtime, so repeated tool calls on
the same .doc don't reconvert.
"""
from __future__ import annotations

import hashlib
import os
import shutil
import subprocess
import tempfile
import unicodedata
import uuid

from core.pdf_converter import (
    DOCKER_IMAGE,
    LibreOfficeNotInstalled,
    ConversionFailed,
    _docker_image_present,
    _local_soffice,
)


def resolve_path(path: str) -> str:
    """Resolve a path that may differ from the filesystem in Unicode form.

    macOS commonly stores filenames in NFD (decomposed); Linux APIs return
    them as bytes, and Python users pass NFC. The two forms compare unequal
    byte-wise. If the literal path doesn't exist, scan the parent directory
    for an entry that matches under NFC normalization.
    """
    if os.path.exists(path):
        return path
    parent = os.path.dirname(path) or "."
    target_nfc = unicodedata.normalize("NFC", os.path.basename(path))
    try:
        for entry in os.listdir(parent):
            if unicodedata.normalize("NFC", entry) == target_nfc:
                return os.path.join(parent, entry)
    except FileNotFoundError:
        pass
    # No match — let the caller raise the canonical FileNotFoundError.
    return path


def _cache_path(input_path: str) -> str:
    abs_in = os.path.abspath(input_path)
    mtime = int(os.path.getmtime(abs_in))
    key = hashlib.sha256(f"{abs_in}|{mtime}".encode("utf-8")).hexdigest()[:16]
    base = os.path.splitext(os.path.basename(abs_in))[0]
    return os.path.join(tempfile.gettempdir(), f"mcp-contracts-doc2docx-{key}-{base}.docx")


def _convert_local(soffice: str, doc_path: str, output_path: str, timeout: int) -> None:
    with tempfile.TemporaryDirectory(prefix="soffice-doc-") as tmpdir:
        profile_dir = os.path.join(tmpdir, "profile")
        out_dir = os.path.join(tmpdir, "out")
        os.makedirs(out_dir)
        proc = subprocess.run(
            [
                soffice,
                f"-env:UserInstallation=file://{profile_dir}",
                "--headless", "--convert-to", "docx",
                "--outdir", out_dir, doc_path,
            ],
            capture_output=True, text=True, timeout=timeout,
        )
        if proc.returncode != 0:
            raise ConversionFailed(
                f"soffice exit {proc.returncode}: stderr={proc.stderr.strip()}"
            )
        produced = os.path.join(
            out_dir, os.path.splitext(os.path.basename(doc_path))[0] + ".docx"
        )
        if not os.path.exists(produced):
            raise ConversionFailed(
                f"soffice succeeded but no DOCX produced. stdout={proc.stdout.strip()}"
            )
        shutil.move(produced, output_path)


def _convert_docker(doc_path: str, output_path: str, timeout: int) -> None:
    """Run conversion via docker create + cp (same pattern as pdf_converter)."""
    name = f"mcp-doc2docx-{uuid.uuid4().hex[:12]}"
    container_out = "/work/out/in.docx"
    script = (
        "mkdir -p /work/out /work/profile && "
        "/usr/bin/soffice "
        "-env:UserInstallation=file:///work/profile "
        "--headless --convert-to docx --outdir /work/out /tmp/in.doc"
    )

    create = subprocess.run(
        ["docker", "create", "--name", name,
         "--entrypoint", "/bin/sh",
         DOCKER_IMAGE, "-c", script],
        capture_output=True, text=True,
    )
    if create.returncode != 0:
        raise ConversionFailed(f"docker create failed: {create.stderr.strip()}")

    try:
        cp_in = subprocess.run(
            ["docker", "cp", doc_path, f"{name}:/tmp/in.doc"],
            capture_output=True, text=True,
        )
        if cp_in.returncode != 0:
            raise ConversionFailed(f"docker cp in failed: {cp_in.stderr.strip()}")

        run = subprocess.run(
            ["docker", "start", "-a", name],
            capture_output=True, text=True, timeout=timeout,
        )
        if run.returncode != 0:
            raise ConversionFailed(
                f"soffice (in container) exit {run.returncode}: "
                f"stderr={run.stderr.strip()} stdout={run.stdout.strip()}"
            )

        cp_out = subprocess.run(
            ["docker", "cp", f"{name}:{container_out}", output_path],
            capture_output=True, text=True,
        )
        if cp_out.returncode != 0:
            raise ConversionFailed(
                f"failed to extract DOCX from container: {cp_out.stderr.strip()}. "
                f"soffice stdout: {run.stdout.strip()}"
            )
    finally:
        subprocess.run(["docker", "rm", "-f", name],
                       capture_output=True, text=True)


def ensure_docx(file_path: str, *, timeout: int = 120) -> str:
    """If `file_path` is a legacy `.doc`, convert it to a cached `.docx` and
    return that path. Otherwise return the input untouched.

    Resolves NFC/NFD Unicode mismatches in the input path (common on
    Cyrillic filenames copied from macOS to Linux). Idempotent — repeated
    calls hit the mtime-keyed cache without invoking soffice. Cache lives
    in `tempfile.gettempdir()`, no cleanup needed (OS reclaims on reboot).
    """
    file_path = resolve_path(file_path)
    if not file_path.lower().endswith(".doc"):
        return file_path

    cached = _cache_path(file_path)
    if os.path.exists(cached) and os.path.getmtime(cached) >= os.path.getmtime(file_path):
        return cached

    soffice = _local_soffice()
    if soffice:
        _convert_local(soffice, file_path, cached, timeout)
        return cached

    if _docker_image_present():
        _convert_docker(file_path, cached, timeout)
        return cached

    raise LibreOfficeNotInstalled(
        f"Cannot convert {file_path}: LibreOffice is not available locally "
        f"and the docker image {DOCKER_IMAGE} is not built."
    )
