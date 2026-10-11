"""Names and legacy lookup for files in a single comparison run directory."""

import re
from pathlib import Path


def artifact_name(name: str, prefix: str = "") -> str:
    if not prefix:
        return name
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._-]*", prefix):
        raise ValueError(f"Invalid comparison artifact prefix: {prefix!r}")
    return name if name.startswith(f"{prefix}-") else f"{prefix}-{name}"


def find_artifact(directory: Path, name: str) -> Path:
    """Return a legacy or prefixed artifact; reject ambiguous run directories."""
    directory = Path(directory)
    legacy = directory / name
    matches = [legacy] if legacy.is_file() else []
    matches += sorted(path for path in directory.glob(f"*-{name}") if path.is_file())
    if len(matches) > 1:
        raise ValueError(f"Multiple comparison artifacts match {name}: {matches}")
    return matches[0] if matches else legacy
