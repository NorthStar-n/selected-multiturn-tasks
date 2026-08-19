"""Detect only the exact, untouched solver-visible checkpoint output."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any


DEFAULT_MANIFEST = Path(__file__).with_name("checkpoint_baseline.json")


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load_checkpoint_baseline(path: Path = DEFAULT_MANIFEST) -> dict[str, Any]:
    baseline = json.loads(path.read_text(errors="strict"))
    if baseline.get("schema_version") != "afo.checkpoint-output-baseline.v1":
        raise ValueError(f"Unsupported checkpoint baseline schema in {path}")
    if not isinstance(baseline.get("files"), dict) or not baseline["files"]:
        raise ValueError(f"Checkpoint baseline has no files: {path}")
    return baseline


def is_exact_checkpoint_output(
    output_dir: Path,
    manifest_path: Path = DEFAULT_MANIFEST,
) -> bool:
    """Return true only for the frozen checkpoint's unchanged Office-file set."""

    baseline = load_checkpoint_baseline(manifest_path)
    extensions = {str(value).lower() for value in baseline["monitored_extensions"]}
    expected = baseline["files"]

    actual: dict[str, Path] = {}
    if output_dir.exists():
        for path in sorted(output_dir.rglob("*")):
            if not path.is_file() or path.suffix.lower() not in extensions:
                continue
            relative = path.relative_to(output_dir)
            if any(part.startswith(".") for part in relative.parts):
                continue
            actual[relative.as_posix()] = path

    if set(actual) != set(expected):
        return False

    for relative, metadata in expected.items():
        path = actual[relative]
        if path.stat().st_size != int(metadata["bytes"]):
            return False
        if _sha256(path) != str(metadata["sha256"]):
            return False
    return True
