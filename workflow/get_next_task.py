"""Reserve the next pending task and print its absolute paths as JSON."""

from __future__ import annotations

import argparse
import json
import sys

from workflow_common import PROMPT_FILE, read_progress, task_folder, task_pdf_path, utc_now, write_progress


def main() -> int:
    parser = argparse.ArgumentParser(description="取得并锁定下一篇待处理论文")
    parser.add_argument("--include-failed", action="store_true", help="失败任务也可重新领取")
    parser.add_argument("--include-needs-review", action="store_true", help="需要复核的任务也可重新领取")
    args = parser.parse_args()

    rows = read_progress()
    allowed = {"pending"} | ({"failed"} if args.include_failed else set())
    if args.include_needs_review:
        allowed.add("needs_review")
    row = next((item for item in rows if item["status"] in allowed), None)
    if row is None:
        print("没有可领取的任务。", file=sys.stderr)
        return 1

    folder = task_folder(row)
    pdf_path = task_pdf_path(row)
    if not pdf_path.exists():
        row["status"] = "failed"
        row["error_message"] = "论文 PDF 不存在"
        row["updated_at"] = utc_now()
        write_progress(rows)
        print(f"任务 {row['task_id']} 的论文 PDF 不存在。", file=sys.stderr)
        return 2

    row["status"] = "processing"
    row["error_message"] = ""
    row["updated_at"] = utc_now()
    write_progress(rows)
    print(json.dumps({
        "task_id": row["task_id"],
        "paper_name": row["paper_name"],
        "pdf_path": str(pdf_path.resolve()),
        "output_folder": str(folder.resolve()),
        "request_path": str(PROMPT_FILE.resolve()),
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
