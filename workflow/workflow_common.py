"""Shared helpers for the local, no-API paper-analysis workflow."""

from __future__ import annotations

import csv
import os
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable


ROOT = Path(__file__).resolve().parents[1]
PROMPT_FILE = ROOT / "prompt" / "analysis_prompt.md"
OUTPUT_DIR = ROOT / "output"
PROGRESS_FILE = ROOT / "progress.csv"

FIELDNAMES = [
    "task_id",
    "paper_name",
    "source_path",
    "output_folder",
    "status",
    "chat_url",
    "docx_status",
    "result_source",
    "quality_check",
    "error_message",
    "updated_at",
]

VALID_STATUSES = {"pending", "processing", "completed", "failed", "needs_review"}

def utc_now() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def safe_folder_component(value: str, max_length: int = 96) -> str:
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value)
    value = re.sub(r"\s+", " ", value).strip().rstrip(". ")
    return (value[:max_length].rstrip(". ") or "unnamed_paper")


def read_progress() -> list[dict[str, str]]:
    if not PROGRESS_FILE.exists():
        return []
    with PROGRESS_FILE.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        return [{field: row.get(field, "") for field in FIELDNAMES} for row in reader]


def write_progress(rows: Iterable[dict[str, str]]) -> None:
    PROGRESS_FILE.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary_name = tempfile.mkstemp(
        dir=PROGRESS_FILE.parent,
        prefix=f".{PROGRESS_FILE.name}.",
        suffix=".tmp",
    )
    try:
        with os.fdopen(fd, "w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=FIELDNAMES)
            writer.writeheader()
            writer.writerows(rows)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary_name, PROGRESS_FILE)
    except BaseException:
        Path(temporary_name).unlink(missing_ok=True)
        raise


def task_folder(row: dict[str, str]) -> Path:
    return ROOT / row["output_folder"]


def task_base_name(row: dict[str, str]) -> str:
    return safe_folder_component(Path(row["paper_name"]).stem, max_length=125)


def task_pdf_path(row: dict[str, str]) -> Path:
    return task_folder(row) / f"{task_base_name(row)}.pdf"


def task_docx_path(row: dict[str, str]) -> Path:
    return task_folder(row) / f"{task_base_name(row)}.docx"


def adopt_single_task_docx(row: dict[str, str]) -> Path:
    """Rename a sole manually added Word file to the required paper name."""
    expected = task_docx_path(row)
    if expected.is_file():
        return expected
    folder = task_folder(row)
    if not folder.is_dir():
        return expected
    candidates = [
        path for path in folder.glob("*.docx")
        if path.is_file() and not path.name.startswith(("~$", "."))
    ]
    if len(candidates) == 1:
        candidates[0].replace(expected)
    return expected


def find_task(rows: list[dict[str, str]], task_id: str) -> dict[str, str]:
    wanted = task_id.zfill(3) if task_id.isdigit() else task_id
    for row in rows:
        if row["task_id"] == wanted:
            return row
    raise ValueError(f"找不到任务 {task_id}。请先通过浏览器扩展导入 PDF。")


def next_numeric_task_id(rows: list[dict[str, str]]) -> int:
    ids = [int(row["task_id"]) for row in rows if row["task_id"].isdigit()]
    return max(ids, default=0) + 1
