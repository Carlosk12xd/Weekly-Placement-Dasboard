"""Atomic local run records and immutable report snapshots."""
from dataclasses import dataclass
from datetime import date
from pathlib import Path
import hashlib
import json
import shutil

ROOT = Path(__file__).resolve().parents[1]


def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def write_json(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(path.suffix + '.tmp')
    temp.write_text(json.dumps(value, indent=2, default=str) + '\n')
    temp.replace(path)


def atomic_copy(source, destination):
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = destination.with_suffix(destination.suffix + '.tmp')
    shutil.copy2(source, temporary)
    temporary.replace(destination)


def verify_file(path, expected):
    if not Path(path).is_file() or sha256(path) != expected:
        raise RuntimeError(f'Artifact missing or changed: {Path(path).name}')


@dataclass
class RunOptions:
    dry_run: bool = True
    offline: bool = False
    run_root: Path = ROOT / 'runs'
    retry: bool = False
    test_recipient: str = ''
    resend_uncertain: bool = False
    dashboard_only: bool = False
    audiences: tuple = ()
    source_root: Path = ROOT
