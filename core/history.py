# 历史留档：每次排版的产物持久化到 out/history/<run_id>/，供演示界面回看与下载。
# 结构：out/history/<run_id>/{formatted.docx, formatted_tracked.docx,
#       formatted_report.docx, report.md, meta.json}
# 旧记录可能只有 formatted.docx / report.md；读取与展示时必须降级而不是报错。
# 公文场景话术：全程留痕、可追溯（审计友好）。

import json
import os
import shutil
from datetime import datetime

from core.runtime import data_path

HISTORY_ROOT = str(data_path("out", "history"))

# 四件套的归档名与 UI 下载信息只维护一份，save_run / list_runs / 页面共用。
DELIVERABLES = (
    {
        "key": "docx",
        "archive": "formatted.docx",
        "label": "DOCX",
        "download": "formatted_{run_id}.docx",
        "missing": "DOCX 不可用",
    },
    {
        "key": "tracked",
        "archive": "formatted_tracked.docx",
        "label": "修订版",
        "download": "formatted_tracked_{run_id}.docx",
        "missing": "修订版不可用",
    },
    {
        "key": "report_docx",
        "archive": "formatted_report.docx",
        "label": "报告 DOCX",
        "download": "format_report_{run_id}.docx",
        "missing": "报告 DOCX 不可用",
    },
    {
        "key": "report",
        "archive": "report.md",
        "label": "报告 MD",
        "download": "format_report_{run_id}.md",
        "missing": "报告 MD 不可用",
    },
)


def _copy_if_file(source, destination):
    """源文件存在才归档；返回是否真的写入了归档文件。"""
    if source and os.path.isfile(source):
        shutil.copy2(source, destination)
        return True
    return False


def save_run(out_path, report_path, meta, *, tracked_path=None,
             report_docx_path=None, root=None):
    """把本次排版四件套拷入历史目录，写入 meta.json，返回 run_id。

    tracked_path / report_docx_path 缺失时只归档存在的文件，不伪造占位；
    meta 里的 files 记录每个交付物最终是否归档。
    """
    base = str(root) if root is not None else HISTORY_ROOT
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    run_id = stamp
    n = 1
    while os.path.exists(os.path.join(base, run_id)):  # 同一秒内重复运行避免撞目录
        n += 1
        run_id = "%s-%d" % (stamp, n)
    run_dir = os.path.join(base, run_id)
    os.makedirs(run_dir)

    sources = {
        "docx": out_path,
        "tracked": tracked_path,
        "report_docx": report_docx_path,
        "report": report_path,
    }
    files = {}
    for item in DELIVERABLES:
        archive_path = os.path.join(run_dir, item["archive"])
        archived = _copy_if_file(sources[item["key"]], archive_path)
        files[item["key"]] = item["archive"] if archived else None

    record = dict(meta or {})
    record["run_id"] = run_id
    record["time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    record["files"] = files
    with open(os.path.join(run_dir, "meta.json"), "w", encoding="utf-8") as handle:
        json.dump(record, handle, ensure_ascii=False, indent=2)
    return run_id


def list_runs(root=None):
    """按时间倒序返回历史记录，暴露四件套路径与 available 标记。

    只有 meta.json 可读的记录才会返回；四件套全部丢失的记录跳过，
    部分丢失的记录保留并标记 available=False，由 UI 降级展示。
    """
    base = str(root) if root is not None else HISTORY_ROOT
    if not os.path.isdir(base):
        return []
    runs = []
    for dirname in os.listdir(base):
        meta_path = os.path.join(base, dirname, "meta.json")
        if not os.path.isfile(meta_path):
            continue
        try:
            with open(meta_path, encoding="utf-8") as handle:
                record = json.load(handle)
        except (OSError, ValueError):
            continue
        if not isinstance(record, dict):
            continue
        record = dict(record)
        available = {}
        for item in DELIVERABLES:
            path = os.path.join(base, dirname, item["archive"])
            record[item["key"]] = path
            available[item["key"]] = os.path.isfile(path)
        record["available"] = available
        if not any(available.values()):
            continue  # 没有任何可下载交付物的记录不展示
        runs.append(record)
    runs.sort(key=lambda item: item.get("time", ""), reverse=True)
    return runs


def deliverable_entries(record):
    """把一条历史记录映射成 UI 下载项；缺失文件只标记不可用，绝不打开。"""
    if not record:
        return []
    run_id = str(record.get("run_id") or "run")
    entries = []
    for item in DELIVERABLES:
        path = record.get(item["key"])
        available = bool(path) and os.path.isfile(path)
        entries.append(
            {
                "key": item["key"],
                "label": item["label"],
                "missing": item["missing"],
                "path": path if available else None,
                "available": available,
                "file_name": item["download"].format(run_id=run_id),
            }
        )
    return entries
