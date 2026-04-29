"""
DOCX → PDF conversion via LibreOffice headless.

Tries local `soffice` first, falls back to a docker image.

Build the docker image once (paths resolved relative to this file so the
hint works regardless of where the project is checked out):
    docker build -t mcp-docx-soffice:latest <project>/docker
"""
import os
import shutil
import subprocess
import tempfile
import uuid

# Project root = parent of `core/`; docker context lives at <root>/docker.
_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_DOCKER_CONTEXT = os.path.join(_PROJECT_ROOT, "docker")

DOCKER_IMAGE = "mcp-docx-soffice:latest"
DOCKER_BUILD_HINT = f"docker build -t {DOCKER_IMAGE} {_DOCKER_CONTEXT}"


class LibreOfficeNotInstalled(RuntimeError):
    pass


class ConversionFailed(RuntimeError):
    pass


def _local_soffice() -> str | None:
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    return None


def _docker_image_present() -> bool:
    if not shutil.which("docker"):
        return False
    proc = subprocess.run(
        ["docker", "image", "inspect", DOCKER_IMAGE],
        capture_output=True, text=True,
    )
    return proc.returncode == 0


def _convert_local(soffice: str, docx_path: str, output_path: str,
                   timeout: int) -> None:
    with tempfile.TemporaryDirectory(prefix="soffice-") as tmpdir:
        # Isolated UserInstallation lets concurrent conversions coexist.
        profile_dir = os.path.join(tmpdir, "profile")
        out_dir = os.path.join(tmpdir, "out")
        os.makedirs(out_dir)
        proc = subprocess.run(
            [
                soffice,
                f"-env:UserInstallation=file://{profile_dir}",
                "--headless", "--convert-to", "pdf",
                "--outdir", out_dir, docx_path,
            ],
            capture_output=True, text=True, timeout=timeout,
        )
        if proc.returncode != 0:
            raise ConversionFailed(
                f"soffice exit {proc.returncode}: stderr={proc.stderr.strip()}"
            )
        produced = os.path.join(
            out_dir, os.path.splitext(os.path.basename(docx_path))[0] + '.pdf'
        )
        if not os.path.exists(produced):
            raise ConversionFailed(
                f"soffice succeeded but no PDF produced. stdout={proc.stdout.strip()}"
            )
        shutil.move(produced, output_path)


def _convert_docker(docx_path: str, output_path: str, timeout: int) -> None:
    """Run conversion via docker create + cp pattern.

    Avoids bind mounts entirely — works regardless of host/daemon filesystem
    visibility (e.g. sandboxed daemons that can't see /tmp).
    """
    name = f"mcp-docx-soffice-{uuid.uuid4().hex[:12]}"
    # Input is renamed to /tmp/in.docx in the container, so output is in.pdf.
    container_pdf = "/work/out/in.pdf"
    script = (
        "mkdir -p /work/out /work/profile && "
        "/usr/bin/soffice "
        "-env:UserInstallation=file:///work/profile "
        "--headless --convert-to pdf --outdir /work/out /tmp/in.docx"
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
            ["docker", "cp", docx_path, f"{name}:/tmp/in.docx"],
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
            ["docker", "cp", f"{name}:{container_pdf}", output_path],
            capture_output=True, text=True,
        )
        if cp_out.returncode != 0:
            raise ConversionFailed(
                f"failed to extract PDF from container: {cp_out.stderr.strip()}. "
                f"soffice stdout: {run.stdout.strip()}"
            )
    finally:
        subprocess.run(["docker", "rm", "-f", name],
                       capture_output=True, text=True)


def convert_docx_to_pdf(docx_path: str, output_path: str | None = None,
                        timeout: int = 180) -> str:
    """Convert DOCX to PDF. Returns path to generated PDF."""
    if not os.path.exists(docx_path):
        raise FileNotFoundError(docx_path)

    if output_path is None:
        output_path = os.path.splitext(docx_path)[0] + '.pdf'

    os.makedirs(os.path.dirname(os.path.abspath(output_path)) or '.',
                exist_ok=True)

    soffice = _local_soffice()
    if soffice:
        _convert_local(soffice, docx_path, output_path, timeout)
        return output_path

    if _docker_image_present():
        _convert_docker(docx_path, output_path, timeout)
        return output_path

    raise LibreOfficeNotInstalled(
        "LibreOffice is not available locally and the docker image "
        f"{DOCKER_IMAGE} is not built. Build it with: {DOCKER_BUILD_HINT}"
    )
