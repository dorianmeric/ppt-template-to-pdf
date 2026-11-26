#!/usr/bin/env python3
"""Build the release assets: the lean source zip and the Nuitka binary for this machine.

Running it with no arguments produces both:

    dist/ppt-template-to-pdf-<version>.zip        the five source files a user needs
    dist/ppt-template-to-pdf-<version>-<platform>-<arch>     one self-contained executable

The zip ships the five files a user needs to run or install the tool, and
nothing else:

    ppt_template_to_pdf.py
    README.md
    pyproject.toml
    requirements.txt
    LICENSE

Tests, the assets folder, and requirements-dev.txt stay in the repository. Every
file sits inside one top-level folder named after the package and the version,
so unzipping cannot scatter files into the current directory.

The binary is built with Nuitka in one-file mode, under the same name and with
the same options the release workflow passes to Nuitka-Action. Nuitka cannot
cross-compile, so this builds the binary for the machine it runs on; the
workflow runs it once per platform and architecture instead.

Usage:

    python tools/make_github_release.py
    python tools/make_github_release.py --zip-only
    python tools/make_github_release.py --binary-only
    python tools/make_github_release.py --expect-version 1.9.1

`--expect-version` is what the release workflow passes the tag to, so a tag that
does not match the sources fails the build instead of shipping the wrong number.
"""

from __future__ import annotations

import argparse
import platform
import re
import subprocess
import sys
from importlib.util import find_spec
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

PROJECT_ROOT = Path(__file__).resolve().parent.parent
PACKAGE_NAME = "ppt-template-to-pdf"
BINARY_PREFIX = "ppt-template-to-pdf"
SCRIPT = PROJECT_ROOT / "ppt_template_to_pdf.py"
ICON = PROJECT_ROOT / "assets" / "logo.ico"

RELEASE_FILES = (
    "ppt_template_to_pdf.py",
    "README.md",
    "pyproject.toml",
    "requirements.txt",
    "LICENSE",
)

PROJECT_VERSION_PATTERN = re.compile(r'^version = "([^"]+)"', re.MULTILINE)
SCRIPT_VERSION_PATTERN = re.compile(r'^__version__ = "([^"]+)"', re.MULTILINE)

# A fixed timestamp and mode keep the archive byte-for-byte reproducible.
FIXED_DATE_TIME = (1980, 1, 1, 0, 0, 0)
FILE_MODE = 0o644

MACHINE_SLUGS = {
    "amd64": "x86_64",
    "x86_64": "x86_64",
    "i386": "x86",
    "i686": "x86",
    "arm64": "arm64",
    "aarch64": "aarch64",
}


def read_version(path: Path, pattern: re.Pattern[str], label: str) -> str:
    """Read a version string out of a source file."""
    match = pattern.search(path.read_text(encoding="utf-8"))
    if match is None:
        raise SystemExit(f"Cannot find {label} in {path.name}")
    return match.group(1)


def host_target() -> tuple[str, str]:
    """Return the platform and architecture slug for this machine.

    The pair matches the asset names the release workflow publishes:
    windows/x64, windows/x86, macos/arm64, macos/x86_64, linux/x86_64.
    """
    if sys.platform == "win32":
        return "windows", "x64" if sys.maxsize > 2**32 else "x86"
    machine = platform.machine().lower()
    arch = MACHINE_SLUGS.get(machine, machine or "unknown")
    if sys.platform == "darwin":
        return "macos", "arm64" if arch in {"arm64", "aarch64"} else "x86_64"
    return "linux", arch


def release_binary_name(version: str, os_slug: str, arch: str) -> str:
    """Name a binary the way the released assets are named."""
    suffix = ".exe" if os_slug == "windows" else ""
    return f"{BINARY_PREFIX}-{version}-{os_slug}-{arch}{suffix}"


def nuitka_command(version: str, os_slug: str, arch: str, output_dir: Path) -> list[str]:
    """Build the Nuitka command line, mirroring the release workflow's options."""
    command = [
        sys.executable,
        "-m",
        "nuitka",
        "--onefile",
        "--assume-yes-for-downloads",
        "--remove-output",
        f"--output-dir={output_dir}",
        f"--output-filename={release_binary_name(version, os_slug, arch)}",
        "--include-package-data=pptx",
    ]
    if os_slug == "windows":
        command += [
            f"--windows-icon-from-ico={ICON}",
            "--windows-console-mode=force",
            "--include-package=win32com",
            "--company-name=ppt-template-to-pdf",
            "--product-name=ppt-template-to-pdf",
            "--file-description=Fill a PowerPoint template from YAML or JSON and export it to PDF",
            f"--file-version={version}",
            f"--product-version={version}",
            "--copyright=MIT licensed",
        ]
    command.append(str(SCRIPT))
    return command


def build_archive(version: str, output_dir: Path) -> Path:
    """Write the source zip and return its path."""
    missing = [name for name in RELEASE_FILES if not (PROJECT_ROOT / name).is_file()]
    if missing:
        raise SystemExit(f"Missing release file(s): {', '.join(missing)}")

    output_dir.mkdir(parents=True, exist_ok=True)
    archive_path = output_dir / f"{PACKAGE_NAME}-{version}.zip"
    prefix = f"{PACKAGE_NAME}-{version}"

    with ZipFile(archive_path, "w", ZIP_DEFLATED, compresslevel=9) as archive:
        for name in RELEASE_FILES:
            info = ZipInfo(f"{prefix}/{name}", date_time=FIXED_DATE_TIME)
            info.compress_type = ZIP_DEFLATED
            info.create_system = 3
            info.external_attr = FILE_MODE << 16
            archive.writestr(info, (PROJECT_ROOT / name).read_bytes())

    # Reopen and check the result, so a serializer change cannot quietly add one.
    with ZipFile(archive_path) as archive:
        written = archive.namelist()
    expected = [f"{prefix}/{name}" for name in RELEASE_FILES]
    if written != expected:
        raise SystemExit(f"The zip holds {written} instead of {expected}")

    print(f"Built {archive_path}")
    for name in RELEASE_FILES:
        print(f"  contains {prefix}/{name}")
    return archive_path


def build_binary(version: str, output_dir: Path) -> Path:
    """Freeze the script for this machine with Nuitka and smoke test the result."""
    if find_spec("nuitka") is None:
        raise SystemExit(
            "Nuitka is not installed. Run: python -m pip install -r requirements-build.txt"
        )

    os_slug, arch = host_target()
    output_dir.mkdir(parents=True, exist_ok=True)
    binary_path = output_dir / release_binary_name(version, os_slug, arch)
    command = nuitka_command(version, os_slug, arch, output_dir)

    print(f"Building the {os_slug} {arch} binary with Nuitka")
    result = subprocess.run(command, cwd=PROJECT_ROOT, check=False)
    if result.returncode != 0:
        raise SystemExit(f"Nuitka failed with exit code {result.returncode}")
    if not binary_path.is_file():
        raise SystemExit(f"Nuitka did not produce {binary_path}")

    # A binary that cannot report its version is not worth uploading.
    check = subprocess.run(
        [str(binary_path), "--version"],
        cwd=PROJECT_ROOT,
        capture_output=True,
        text=True,
        check=False,
    )
    if check.returncode != 0 or version not in check.stdout:
        raise SystemExit(
            f"{binary_path.name} did not report version {version}: "
            f"exit {check.returncode}, output {check.stdout.strip()!r} {check.stderr.strip()!r}"
        )
    return binary_path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=PROJECT_ROOT / "dist",
        help="Folder for the assets. Defaults to dist/ in the repository.",
    )
    parser.add_argument(
        "--expect-version",
        default=None,
        metavar="VERSION",
        help="Fail unless this is the version in the sources. The release tag goes here.",
    )
    parser.add_argument(
        "--zip-only",
        action="store_true",
        help="Build the source zip and skip the Nuitka binary.",
    )
    parser.add_argument(
        "--binary-only",
        action="store_true",
        help="Build the Nuitka binary for this machine and skip the source zip.",
    )
    arguments = parser.parse_args(argv)
    if arguments.zip_only and arguments.binary_only:
        parser.error("--zip-only and --binary-only cannot be used together")

    project_version = read_version(
        PROJECT_ROOT / "pyproject.toml", PROJECT_VERSION_PATTERN, "the project version"
    )
    script_version = read_version(SCRIPT, SCRIPT_VERSION_PATTERN, "__version__")
    if project_version != script_version:
        raise SystemExit(
            f"Version mismatch: pyproject.toml says {project_version}, "
            f"the script says {script_version}."
        )
    if arguments.expect_version and arguments.expect_version != project_version:
        raise SystemExit(
            f"The tag says {arguments.expect_version}, the sources say {project_version}. "
            "Bump pyproject.toml and __version__ before tagging."
        )

    output_dir = arguments.output_dir.resolve()
    built: list[Path] = []
    if not arguments.binary_only:
        built.append(build_archive(project_version, output_dir))
    if not arguments.zip_only:
        built.append(build_binary(project_version, output_dir))

    print()
    print(f"Release assets for {project_version} in {output_dir}")
    for path in built:
        size = path.stat().st_size
        shown = f"{size / (1024 * 1024):.1f} MB" if size >= 1024 * 1024 else f"{size} bytes"
        print(f"  {path.name}  ({shown})")
    if not arguments.zip_only:
        print()
        print("Upload them with: gh release upload v" + project_version + f" {output_dir}/*")
    return 0


if __name__ == "__main__":
    sys.exit(main())
