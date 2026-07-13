"""Archive a downloaded Word file, optionally save the chat URL, then run checks."""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

from check_results import inspect_docx
from workflow_common import find_task, read_progress, task_docx_path, task_folder, utc_now, write_progress


def main() -> int:
    parser = argparse.ArgumentParser(description="归档下载的 Word 文件并更新任务")
    parser.add_argument("task_id", help="任务编号，例如 001")
    parser.add_argument("download_path", type=Path, help="刚下载的 .docx 文件的完整路径")
    parser.add_argument("--chat-url", default="", help="可选：当前 ChatGPT 对话链接")
    parser.add_argument("--replace", action="store_true", help="允许替换已有 analysis.docx")
    parser.add_argument("--min-size-kb", type=int, default=20)
    parser.add_argument("--min-headings", type=int, default=3)
    parser.add_argument("--required-section", action="append", default=[])
    args = parser.parse_args()

    source = args.download_path.expanduser().resolve()
    if not source.is_file() or source.suffix.lower() != ".docx":
        print(f"下载文件不是有效的 .docx：{source}", file=sys.stderr)
        return 2

    rows = read_progress()
    try:
        row = find_task(rows, args.task_id)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 2

    destination = task_docx_path(row)
    if destination.exists() and not args.replace:
        print(f"目标已存在：{destination}。确认替换请添加 --replace。", file=sys.stderr)
        return 2

    # Validate a replacement before touching the previously accepted result.
    ok, message = inspect_docx(source, args.min_size_kb, args.min_headings, args.required_section)
    if destination.exists() and args.replace and not ok and source != destination:
        print(f"新文件未通过检查，已保留原结果和下载文件：{message}", file=sys.stderr)
        return 1

    destination.parent.mkdir(parents=True, exist_ok=True)
    if source != destination:
        shutil.move(str(source), str(destination))

    row["chat_url"] = args.chat_url or row["chat_url"]
    row["docx_status"] = "valid" if ok else "invalid"
    row["result_source"] = "chatgpt_docx"
    row["quality_check"] = message
    row["status"] = "completed" if ok else "needs_review"
    row["error_message"] = "" if ok else message
    row["updated_at"] = utc_now()
    write_progress(rows)
    print(f"任务 {row['task_id']}：{'完成' if ok else '需要复核'} — {message}")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
