"""历史留档（四件套）与 app 结果页持久化的回归测试。

所有落盘都用 tmp_path，不写仓库 out/。
"""

import json
import os
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO_ROOT)

from core import history


def _write(path, payload):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(payload)
    return path


def _write_text(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")
    return path


def _make_sources(tmp_path):
    src = tmp_path / "sources"
    return {
        "docx": _write(src / "formatted.docx", b"docx-bytes"),
        "tracked": _write(src / "formatted_tracked.docx", b"tracked-bytes"),
        "report_docx": _write(src / "formatted_report.docx", b"report-docx-bytes"),
        "report": _write_text(src / "report.md", "# 修改报告"),
    }


def _save_four(tmp_path, root, meta=None):
    sources = _make_sources(tmp_path)
    run_id = history.save_run(
        str(sources["docx"]),
        str(sources["report"]),
        meta or {"source_name": "示例.docx", "spec_mode": "文字说明"},
        tracked_path=str(sources["tracked"]),
        report_docx_path=str(sources["report_docx"]),
        root=str(root),
    )
    return run_id, sources


def test_save_run_archives_all_four_deliverables_and_writes_meta(tmp_path):
    root = tmp_path / "history"
    run_id, _ = _save_four(tmp_path, root)
    run_dir = root / run_id

    assert (run_dir / "formatted.docx").read_bytes() == b"docx-bytes"
    assert (run_dir / "formatted_tracked.docx").read_bytes() == b"tracked-bytes"
    assert (run_dir / "formatted_report.docx").read_bytes() == b"report-docx-bytes"
    assert (run_dir / "report.md").read_text(encoding="utf-8") == "# 修改报告"

    meta = json.loads((run_dir / "meta.json").read_text(encoding="utf-8"))
    assert meta["run_id"] == run_id
    assert meta["source_name"] == "示例.docx"
    assert meta["files"] == {
        "docx": "formatted.docx",
        "tracked": "formatted_tracked.docx",
        "report_docx": "formatted_report.docx",
        "report": "report.md",
    }

    runs = history.list_runs(root=str(root))
    assert len(runs) == 1
    record = runs[0]
    assert record["run_id"] == run_id
    assert record["docx"] == str(run_dir / "formatted.docx")
    assert record["tracked"] == str(run_dir / "formatted_tracked.docx")
    assert record["report_docx"] == str(run_dir / "formatted_report.docx")
    assert record["report"] == str(run_dir / "report.md")
    assert record["available"] == {
        "docx": True,
        "tracked": True,
        "report_docx": True,
        "report": True,
    }


def test_legacy_two_file_record_is_listed_with_degraded_entries(tmp_path):
    root = tmp_path / "history"
    run_dir = root / "20260101-090000"
    _write(run_dir / "formatted.docx", b"legacy-docx")
    _write_text(run_dir / "report.md", "# legacy report")
    (run_dir / "meta.json").write_text(
        json.dumps(
            {"run_id": "20260101-090000", "time": "2026-01-01 09:00:00",
             "source_name": "旧记录.docx"},
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    runs = history.list_runs(root=str(root))
    assert len(runs) == 1
    record = runs[0]
    assert record["source_name"] == "旧记录.docx"
    assert record["available"] == {
        "docx": True,
        "tracked": False,
        "report_docx": False,
        "report": True,
    }

    entries = history.deliverable_entries(record)
    assert [entry["key"] for entry in entries] == [
        "docx", "tracked", "report_docx", "report"]
    assert [entry["available"] for entry in entries] == [True, False, False, True]
    for entry in entries:
        if entry["available"]:
            assert os.path.isfile(entry["path"])
            assert entry["file_name"]
        else:
            # 降级项没有任何可打开路径，UI 只能显示“不可用”。
            assert entry["path"] is None or not os.path.isfile(entry["path"])


def test_missing_archived_files_do_not_break_list_runs(tmp_path):
    root = tmp_path / "history"
    run_id, _ = _save_four(tmp_path, root)
    run_dir = root / run_id
    (run_dir / "formatted_tracked.docx").unlink()
    (run_dir / "formatted_report.docx").unlink()

    runs = history.list_runs(root=str(root))
    assert len(runs) == 1
    record = runs[0]
    assert record["available"] == {
        "docx": True,
        "tracked": False,
        "report_docx": False,
        "report": True,
    }
    entries = {entry["key"]: entry for entry in history.deliverable_entries(record)}
    assert entries["tracked"]["available"] is False
    assert entries["report_docx"]["available"] is False
    assert entries["docx"]["available"] is True
    assert entries["report"]["available"] is True


def test_record_without_any_deliverable_is_skipped(tmp_path):
    root = tmp_path / "history"
    run_dir = root / "20260102-090000"
    run_dir.mkdir(parents=True)
    (run_dir / "meta.json").write_text(
        json.dumps({"run_id": "20260102-090000", "time": "2026-01-02 09:00:00"}),
        encoding="utf-8",
    )

    assert history.list_runs(root=str(root)) == []


def _fake_result(tmp_path):
    files = {key: _write(tmp_path / name, payload) for key, name, payload in (
        ("out_path", "formatted.docx", b"docx"),
        ("tracked_path", "formatted_tracked.docx", b"tracked"),
        ("report_docx_path", "formatted_report.docx", b"report-docx"),
    )}
    files["report_path"] = _write_text(tmp_path / "report.md", "# 报告")
    return {
        "out_path": str(files["out_path"]),
        "tracked_path": str(files["tracked_path"]),
        "report_docx_path": str(files["report_docx_path"]),
        "report_path": str(files["report_path"]),
        "spec": {"page": {"margin_mm": 30}},
        "rolemap": {0: "标题", 1: "正文"},
        "stylemap": {"标题": "Title", "正文": "Body"},
        "paragraphs": [
            {"idx": 0, "text": "标题"},
            {"idx": 1, "text": "正文"},
        ],
        "changelog": [{"idx": 0, "style_name": "Title"}],
        "issues": [],
    }


_RESULT_LABELS = [
    "下载排版后的 DOCX",
    "下载修订模式 DOCX（审阅视图可见改动）",
    "下载修改报告 DOCX",
    "下载修改报告 Markdown",
]


def test_result_area_and_four_downloads_survive_rerun_without_agent_call(
        tmp_path, monkeypatch):
    """结果区从 session_state 渲染：任意 rerun 后四件套下载仍在，Agent 不被再次调用。"""
    import core.agent as agent_module

    calls = []

    class _StubAgent:
        def __init__(self, *args, **kwargs):
            calls.append("init")

        def run(self, *args, **kwargs):
            calls.append("run")
            raise AssertionError("结果页重跑不应再次调用 Agent/模型")

    monkeypatch.setattr(agent_module, "Agent", _StubAgent)

    from streamlit.testing.v1 import AppTest

    at = AppTest.from_file(os.path.join(REPO_ROOT, "app.py"), default_timeout=120)
    at.session_state["last_result"] = {
        "result": _fake_result(tmp_path),
        "source_name": "持久化测试.docx",
        "run_id": "20260101-000000",
        "compare": {"before": [], "after": [], "error": None, "legacy": True},
    }
    at.run()
    assert list(at.exception) == []
    labels = [element.proto.label for element in at.get("download_button")]
    result_buttons = [
        element for element in at.get("download_button")
        if element.proto.label in _RESULT_LABELS
    ]
    assert labels.count("下载排版后的 DOCX") == 1
    assert len(result_buttons) == 4
    assert all(element.proto.ignore_rerun for element in result_buttons)

    # 模拟任意 rerun：切换主题/tab、点击其他控件、或下载后触发的重跑。
    at.run()
    assert list(at.exception) == []
    assert calls == []
    assert at.session_state["last_result"]["source_name"] == "持久化测试.docx"
    labels_after = [element.proto.label for element in at.get("download_button")]
    for label in _RESULT_LABELS:
        assert labels_after.count(label) == 1, f"rerun 后缺少下载按钮：{label}"


def test_real_run_stores_result_and_keeps_downloads_after_rerun(tmp_path, monkeypatch):
    """真实执行一次（Agent 用桩）：结果写入 session_state，rerun 后仍可下载且不二次调用。"""
    import core.agent as agent_module
    import core.history as history_module
    import core.render as render_module

    # 历史归档指向 tmp；对比图渲染用桩，避免拉起本机渲染器。
    monkeypatch.setattr(history_module, "HISTORY_ROOT", str(tmp_path / "history"))
    monkeypatch.setattr(render_module, "render_docx_to_png", lambda *args, **kwargs: [])

    calls = []

    class _StubAgent:
        def __init__(self, on_event=None):
            self.on_event = on_event
            calls.append("init")

        def run(self, out_path, report_path, **kwargs):
            calls.append("run")
            base = os.path.splitext(out_path)[0]
            tracked_path = base + "_tracked.docx"
            report_docx_path = base + "_report.docx"
            for path, payload in (
                (out_path, b"docx"),
                (tracked_path, b"tracked"),
                (report_docx_path, b"report-docx"),
            ):
                with open(path, "wb") as handle:
                    handle.write(payload)
            with open(report_path, "w", encoding="utf-8") as handle:
                handle.write("# 报告")
            return {
                "out_path": out_path,
                "tracked_path": tracked_path,
                "report_docx_path": report_docx_path,
                "report_path": report_path,
                "spec": {"page": {"margin_mm": 30}},
                "rolemap": {0: "标题"},
                "stylemap": {"标题": "Title"},
                "paragraphs": [{"idx": 0, "text": "标题"}],
                "changelog": [{"idx": 0, "style_name": "Title"}],
                "issues": [],
            }

    monkeypatch.setattr(agent_module, "Agent", _StubAgent)

    from streamlit.testing.v1 import AppTest

    at = AppTest.from_file(os.path.join(REPO_ROOT, "app.py"), default_timeout=120)
    at.run()
    assert list(at.exception) == []
    at.toggle[0].set_value(True).run()
    assert list(at.exception) == []
    run_buttons = [button for button in at.button if button.label == "开始排版 →"]
    assert len(run_buttons) == 1
    assert not run_buttons[0].proto.disabled
    run_buttons[0].click().run()
    assert list(at.exception) == []
    assert calls == ["init", "run"]

    stored = at.session_state["last_result"]
    assert stored["result"]["out_path"].endswith("formatted.docx")
    labels = [element.proto.label for element in at.get("download_button")]
    for label in _RESULT_LABELS:
        assert labels.count(label) == 1, f"结果区缺少下载按钮：{label}"
    # 历史区也归档并展示四件套（含只有新记录才有的修订版/报告 DOCX）。
    assert labels.count("修订版") == 1
    assert labels.count("报告 DOCX") == 1
    assert labels.count("报告 MD") == 1
    history_dir = tmp_path / "history"
    archived = list(history_dir.glob("*/"))
    assert len(archived) == 1
    assert (archived[0] / "formatted_tracked.docx").is_file()
    assert (archived[0] / "formatted_report.docx").is_file()

    # 模拟任意 rerun：结果区和四件套下载仍在，Agent/模型没有被再次调用。
    at.run()
    assert list(at.exception) == []
    assert calls == ["init", "run"]
    labels_after = [element.proto.label for element in at.get("download_button")]
    for label in _RESULT_LABELS:
        assert labels_after.count(label) == 1, f"rerun 后缺少下载按钮：{label}"


def test_legacy_history_ui_degrades_without_opening_missing_files(tmp_path, monkeypatch):
    """旧记录只有 DOCX/Markdown：历史区仍可列出并下载已有文件，缺失项显示不可用。"""
    import core.history as history_module

    monkeypatch.setattr(history_module, "HISTORY_ROOT", str(tmp_path / "history"))
    run_dir = tmp_path / "history" / "20260101-090000"
    run_dir.mkdir(parents=True)
    (run_dir / "formatted.docx").write_bytes(b"legacy-docx")
    (run_dir / "report.md").write_text("# legacy", encoding="utf-8")
    (run_dir / "meta.json").write_text(
        json.dumps(
            {"run_id": "20260101-090000", "time": "2026-01-01 09:00:00",
             "source_name": "旧记录.docx"},
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    from streamlit.testing.v1 import AppTest

    at = AppTest.from_file(os.path.join(REPO_ROOT, "app.py"), default_timeout=120)
    at.run()
    assert list(at.exception) == []
    expander_labels = [element.label for element in at.expander]
    assert any("查看历史任务（1）" in label for label in expander_labels)
    captions = [element.value for element in at.caption]
    assert "修订版不可用" in captions
    assert "报告 DOCX 不可用" in captions
    labels = [element.proto.label for element in at.get("download_button")]
    assert labels.count("DOCX") == 1
    assert labels.count("报告 MD") == 1
    assert labels.count("修订版") == 0
    assert labels.count("报告 DOCX") == 0
