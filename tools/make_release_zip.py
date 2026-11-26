#!/usr/bin/env python3
"""Build the lean release zip that is attached to a GitHub release.

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

Usage:

    python tools/make_release_zip.py
    python tools/make_release_zip.py --expect-version 1.7.0
    python tools/make_release_zip.py --output-dir dist

`--expect-version` is what the release workflow passes the tag to, so a tag that
does not match the sources fails the build instead of shipping the wrong number.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

PROJECT_ROOT = Path(__file__).resolve().parent.parent
PACKAGE_NAME = "ppt-template-to-pdf"

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


def read_version(path: Path, pattern: re.Pattern[str], label: str) -> str:
    """Read a version string out of a source file."""
    match = pattern.search(path.read_text(encoding="utf-8"))
    if match is None:
        raise SystemExit(f"Cannot find {label} in {path.name}")
    return match.group(1)


def build_archive(version: str, output_dir: Path) -> Path:
    """Write the zip and return its path."""
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
    return archive_path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=PROJECT_ROOT / "dist",
        help="Folder for the zip. Defaults to dist/ in the repository.",
    )
    parser.add_argument(
        "--expect-version",
        default=None,
        metavar="VERSION",
        help="Fail unless this is the version in the sources. The release tag goes here.",
    )
    arguments = parser.parse_args(argv)

    project_version = read_version(
        PROJECT_ROOT / "pyproject.toml", PROJECT_VERSION_PATTERN, "the project version"
    )
    script_version = read_version(
        PROJECT_ROOT / "ppt_template_to_pdf.py", SCRIPT_VERSION_PATTERN, "__version__"
    )
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

    archive_path = build_archive(project_version, arguments.output_dir)
    print(f"Built {archive_path}")
    print(f"  version: {project_version}")
    print(f"  size:    {archive_path.stat().st_size} bytes")
    for name in RELEASE_FILES:
        print(f"  contains {PACKAGE_NAME}-{project_version}/{name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
