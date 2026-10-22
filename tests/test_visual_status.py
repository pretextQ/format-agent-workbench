"""视觉复核状态（verification_status）与渲染器探测的回归测试。"""

import os
import sys
import types
from contextlib import contextmanager
from types import SimpleNamespace

import pytest

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO_ROOT)

from core import agent as agent_module
from core import input_conversion as conversion_module
from core import preflight as preflight_module
from core import render as render_module
from core import report_docx as report_docx_module
from core import text_integrity as integrity_module
from core import verify_visual as visual_module

SPEC = {
    "page": {"margin_mm": 30},
    "roles": {"body": {"font_eastasia": "宋体", "size_pt": 12,
                       "alignment": "justify"}},
}
PARAGRAPHS = [
    {"idx": 0, "text": "正文", "in_table": False, "story": "main", "editable": True},
]
CHANGELOG = [
    {"idx": 0, "changed_fields": ["size_pt"], "style_name": "body",
     "allowed_additions": [], "stripped_prefixes": []},
]
LEGACY_RESULT_KEYS = {
    "spec", "paragraphs", "rolemap", "stylemap", "changelog", "issues",
    "applied_fixes", "text_integrity", "field_refresh", "preflight",
    "input_conversion", "out_path", "report_path", "report_docx_path",
    "tracked_path",
}


class _FakeTransaction:
    def __init__(self, paths):
        self.final_paths = dict(paths)
        self.temp_paths = {key: value + ".tmp" for key, value in paths.items()}

    def temp(self, key):
        return self.temp_paths[key]

    def commit(self, main_key=None):
        return None


@contextmanager
def _fake_converted(path):
    yield SimpleNamespace(
        docx_path=path, lossy=False, as_dict=lambda: {"lossy": False})


def _patch_pipeline(monkeypatch):
    """把重流程替换成内存桩，只保留视觉复核分支的真实逻辑。"""
    monkeypatch.setattr(agent_module, "validate_output_paths", lambda *a, **k: None)
    monkeypatch.setattr(agent_module, "validate_spec", lambda spec: None)
    monkeypatch.setattr(agent_module, "extract_paragraphs",
                        lambda path: [dict(item) for item in PARAGRAPHS])
    monkeypatch.setattr(agent_module, "style_name_for_role",
                        lambda role, rule: "Body")
    monkeypatch.setattr(agent_module, "AtomicOutputSet", _FakeTransaction)
    monkeypatch.setattr(agent_module, "cleanup_temps", lambda *a, **k: None)

    def fake_apply_format(target, spec, rolemap, output, **kwargs):
        with open(output, "wb") as handle:
            handle.write(b"docx")
        return [dict(change) for change in CHANGELOG]

    monkeypatch.setattr(agent_module, "apply_format", fake_apply_format)
    monkeypatch.setattr(agent_module, "write_report", lambda *a, **k: None)
    monkeypatch.setattr(conversion_module, "converted_input", _fake_converted)
    monkeypatch.setattr(report_docx_module, "build_report_docx",
                        lambda *a, **k: None)
    monkeypatch.setattr(
        integrity_module, "check_text_integrity",
        lambda *a, **k: {"ok": True, "added": [], "removed": [],
                         "changed": [], "story_differences": []})
    monkeypatch.setattr(
        preflight_module, "preflight_docx",
        lambda *a, **k: {"warnings": [], "blockers": []})
    monkeypatch.setattr(
        preflight_module, "merge_preflight_reports",
        lambda target, template_report=None: {
            "warnings": [], "blockers": [], "section_count": 1,
            "story_parts": [{"kind": "main"}]})
    monkeypatch.setattr(preflight_module, "raise_for_preflight", lambda report: None)


def _run_agent(tmp_path, monkeypatch, verify, visual_behavior,
               repair_fixes=None):
    _patch_pipeline(monkeypatch)
    calls = []

    def fake_verify(*args, **kwargs):
        calls.append("verify")
        return visual_behavior()

    monkeypatch.setattr(visual_module, "verify_visual", fake_verify)
    monkeypatch.setattr(visual_module, "apply_fixes",
                        lambda spec, failed: (spec, repair_fixes or []))
    events = []
    result = agent_module.Agent(llm=object(), on_event=events.append).run(
        target_path=str(tmp_path / "input.docx"),
        out_path=str(tmp_path / "out.docx"),
        spec=dict(SPEC),
        rolemap={0: "body"},
        verify=verify,
    )
    return result, calls, events


def test_verify_false_returns_not_run_without_visual_work(tmp_path, monkeypatch):
    result, calls, _ = _run_agent(
        tmp_path, monkeypatch, False, lambda: pytest.fail("不应调用视觉复核"))
    assert result["verification_status"] == "not_run"
    assert result["verification_error"] is None
    assert result["issues"] == []
    assert calls == []


def test_visual_pass_sets_passed(tmp_path, monkeypatch):
    issues = [{"role": "body", "field": "size_pt", "pass": True,
               "observed": "12", "expected": "12"}]
    result, calls, _ = _run_agent(
        tmp_path, monkeypatch, True, lambda: [dict(item) for item in issues])
    assert result["verification_status"] == "passed"
    assert result["verification_error"] is None
    assert result["issues"] == issues
    assert calls == ["verify"]


def test_visual_mismatch_sets_failed_with_reason(tmp_path, monkeypatch):
    failing = [{"role": "body", "field": "size_pt", "pass": False,
                "observed": "12", "expected": "14"}]
    result, _, events = _run_agent(
        tmp_path, monkeypatch, True, lambda: [dict(item) for item in failing])
    assert result["verification_status"] == "failed"
    assert "body.size_pt" in result["verification_error"]
    assert result["issues"] == failing
    assert result["applied_fixes"] == []
    assert any(event["step"] == "视觉自检" and event["status"] == "warn"
               for event in events)


def test_successful_repair_still_failed_and_notes_not_rechecked(
        tmp_path, monkeypatch):
    failing = [{"role": "body", "field": "size_pt", "pass": False,
                "observed": "12", "expected": "14"}]
    result, _, _ = _run_agent(
        tmp_path, monkeypatch, True,
        lambda: [dict(item) for item in failing],
        repair_fixes=[dict(failing[0])])
    assert result["verification_status"] == "failed"
    assert result["applied_fixes"] == [dict(failing[0])]
    assert "已定向修复 1 项" in result["verification_error"]
    assert "未重新渲染复核" in result["verification_error"]


@pytest.mark.parametrize(
    "exc_type, expected_status, expected_text",
    [
        (visual_module.VisualRenderError, "failed", "渲染阶段失败"),
        (visual_module.VisualModelError, "failed", "多模态请求失败"),
        (visual_module.VisualResponseError, "failed", "模型 JSON/结构校验失败"),
        (visual_module.VisualInconclusiveError, "incomplete", "模型无法下结论"),
    ],
)
def test_visual_errors_map_to_status_and_reason(
        tmp_path, monkeypatch, exc_type, expected_status, expected_text):
    def boom():
        raise exc_type("原始错误")

    result, _, _ = _run_agent(tmp_path, monkeypatch, True, boom)
    assert result["verification_status"] == expected_status
    assert result["verification_error"] is not None
    assert expected_text in result["verification_error"]
    assert "原始错误" in result["verification_error"]
    assert result["issues"] == []


def test_result_only_adds_verification_fields(tmp_path, monkeypatch):
    result, _, _ = _run_agent(tmp_path, monkeypatch, False, list)
    assert LEGACY_RESULT_KEYS <= set(result)
    # 封闭集合守卫：除下列键外 result 不得新增任何字段（R2-2 经 Lead 裁定
    # 授权纳入 post_refresh_integrity——刷新后复检结论；守卫强度不变）。
    assert set(result) - LEGACY_RESULT_KEYS == {
        "verification_status", "verification_error",
        "post_refresh_integrity"}


def test_describe_not_run_never_claims_zero():
    display = agent_module.describe_verification("not_run", None, [])
    assert display["status"] == "not_run"
    assert display["tone"] == "info"
    assert display["metric"] == "未运行"
    assert "不能据此认为视觉问题为 0" in display["message"]


def test_describe_passed_is_the_only_zero():
    issues = [{"role": "body", "field": "size_pt", "pass": True}]
    display = agent_module.describe_verification("passed", None, issues)
    assert display["status"] == "passed"
    assert display["tone"] == "success"
    assert display["metric"] == "0"
    assert display["problems"] == 0
    assert "通过" in display["message"]


def test_describe_failed_render_error_shows_reason():
    display = agent_module.describe_verification(
        "failed", "渲染阶段失败：Word COM 超时", [])
    assert display["tone"] == "error"
    assert display["metric"] == "失败"
    assert display["metric"] != "0"
    assert "原因" in display["message"]
    assert "渲染阶段失败" in display["message"]


@pytest.mark.parametrize("reason, expected", [
    ("多模态请求失败：HTTP 500", "多模态请求失败"),
    ("模型 JSON/结构校验失败：缺少 issues 数组", "模型 JSON/结构校验失败"),
])
def test_describe_failed_model_and_json_errors(reason, expected):
    display = agent_module.describe_verification("failed", reason, [])
    assert display["tone"] == "error"
    assert display["metric"] != "0"
    assert expected in display["message"]


def test_describe_failed_issues_counts_problems():
    issues = [
        {"role": "body", "field": "size_pt", "pass": False},
        {"role": "body", "field": "alignment", "pass": True},
    ]
    display = agent_module.describe_verification("failed", "body.size_pt", issues)
    assert display["problems"] == 1
    assert display["metric"] == "1 项"
    assert "1 项不符合规范" in display["message"]
    assert "不符项：body.size_pt" in display["message"]


def test_describe_incomplete_and_unknown_status():
    display = agent_module.describe_verification(
        "incomplete", "模型无法下结论：图片空白", [])
    assert display["status"] == "incomplete"
    assert display["tone"] == "warning"
    assert display["metric"] == "未完成"
    assert "模型无法下结论" in display["message"]

    unknown = agent_module.describe_verification("something-else", None, [])
    assert unknown["status"] == "incomplete"
    assert unknown["metric"] == "未完成"

    missing = agent_module.describe_verification(None, None, [])
    assert missing["status"] == "not_run"
    assert missing["metric"] == "未运行"


# ---------------- renderer_status ----------------

def _fake_winreg(monkeypatch, registered):
    """registered: {subkey: 默认值}；未列出的 subkey 视为未注册。"""
    module = types.ModuleType("winreg")
    module.HKEY_CLASSES_ROOT = object()

    class _Key:
        def __init__(self, value):
            self.value = value

        def __enter__(self):
            return self

        def __exit__(self, *exc_info):
            return False

    def open_key(root, path):
        if path not in registered:
            raise OSError("未注册")
        return _Key(registered[path])

    def query_value(key, name):
        return key.value, 1

    module.OpenKey = open_key
    module.QueryValueEx = query_value
    monkeypatch.setitem(sys.modules, "winreg", module)


def test_word_progid_requires_registry_entry(monkeypatch):
    _fake_winreg(monkeypatch, {})
    assert render_module._word_com_progid() is None

    _fake_winreg(monkeypatch, {r"Word.Application\CLSID": "{WORD-CLSID}"})
    assert render_module._word_com_progid() == "Word.Application"

    _fake_winreg(monkeypatch, {r"Kwps.Application\CurVer": "Kwps.Application.1"})
    assert render_module._word_com_progid() == "Kwps.Application"


def test_renderer_status_does_not_trust_win32com_import_alone(monkeypatch):
    # pywin32 可导入但 Word ProgID 未注册：不得再返回 available=True。
    monkeypatch.setattr(render_module, "_win32com_available", lambda: True)
    monkeypatch.setattr(render_module, "_word_com_progid", lambda: None)
    monkeypatch.setattr(
        render_module, "libreoffice_status",
        lambda: {"available": False, "path": None, "version": None,
                 "searched": []})
    status = render_module.renderer_status()
    assert status["available"] is False
    assert status["confirmed"] is False
    assert status["name"] != "Word COM"


def test_renderer_status_word_com_uses_registry_without_launch(monkeypatch):
    monkeypatch.setattr(render_module, "_win32com_available", lambda: True)
    monkeypatch.setattr(render_module, "_word_com_progid",
                        lambda: "Word.Application")

    def _must_not_probe_soffice():
        raise AssertionError("Word COM 可用时不应再探测 LibreOffice")

    monkeypatch.setattr(render_module, "libreoffice_status",
                        _must_not_probe_soffice)
    status = render_module.renderer_status()
    assert status["available"] is True
    assert status["name"] == "Word COM"
    assert status["confirmed"] is False
    assert "registry" in status["probe"]


def test_renderer_status_detects_wps_com_without_launch(monkeypatch):
    monkeypatch.setattr(render_module, "_win32com_available", lambda: True)
    monkeypatch.setattr(render_module, "_word_com_progid",
                        lambda: "Kwps.Application")

    def _must_not_probe_soffice():
        raise AssertionError("WPS COM 可用时不应再探测 LibreOffice")

    monkeypatch.setattr(render_module, "libreoffice_status",
                        _must_not_probe_soffice)
    status = render_module.renderer_status()
    assert status["available"] is True
    assert status["name"] == "WPS COM"
    assert status["prog_id"] == "Kwps.Application"


def test_renderer_status_falls_back_to_libreoffice(monkeypatch):
    monkeypatch.setattr(render_module, "_win32com_available", lambda: False)
    monkeypatch.setattr(render_module, "_word_com_progid", lambda: None)
    monkeypatch.setattr(
        render_module, "libreoffice_status",
        lambda: {"available": True, "path": "soffice.exe",
                 "version": "LibreOffice 24", "searched": []})
    status = render_module.renderer_status()
    assert status["available"] is True
    assert status["name"] == "LibreOffice"
    assert status["confirmed"] is True


def test_renderer_status_reports_unavailable_with_reason(monkeypatch):
    monkeypatch.setattr(render_module, "_win32com_available", lambda: False)
    monkeypatch.setattr(render_module, "_word_com_progid", lambda: None)
    monkeypatch.setattr(
        render_module, "libreoffice_status",
        lambda: {"available": False, "path": None, "version": None,
                 "searched": ["soffice"]})
    status = render_module.renderer_status()
    assert status["available"] is False
    assert status["confirmed"] is False
    assert status["reason"]


# ---------------- app 结果区文案 ----------------

def _app_result(tmp_path, status, error=None, issues=None):
    files = {}
    for key, name, payload in (
        ("out_path", "formatted.docx", b"docx"),
        ("tracked_path", "formatted_tracked.docx", b"tracked"),
        ("report_docx_path", "formatted_report.docx", b"report-docx"),
    ):
        path = tmp_path / name
        path.write_bytes(payload)
        files[key] = str(path)
    report = tmp_path / "report.md"
    report.write_text("# 报告", encoding="utf-8")
    files["report_path"] = str(report)
    return {
        **files,
        "spec": {}, "rolemap": {}, "stylemap": {},
        "paragraphs": [], "changelog": [],
        "issues": list(issues or []),
        "verification_status": status,
        "verification_error": error,
    }


def test_app_result_never_shows_zero_for_unverified_or_failed(tmp_path, monkeypatch):
    import core.history as history_module

    monkeypatch.setattr(history_module, "HISTORY_ROOT", str(tmp_path / "history"))
    from streamlit.testing.v1 import AppTest

    at = AppTest.from_file(os.path.join(REPO_ROOT, "app.py"), default_timeout=120)
    cases = [
        ("not_run", None, [], "未运行", "未运行视觉复核"),
        ("failed", "渲染阶段失败：Word COM 超时", [], "失败", "渲染阶段失败"),
        ("failed", "多模态请求失败：HTTP 500", [], "失败", "多模态请求失败"),
        ("failed", "模型 JSON/结构校验失败：缺少 issues 数组", [], "失败",
         "模型 JSON/结构校验失败"),
        ("failed", "body.size_pt", [{"role": "body", "field": "size_pt",
                                    "pass": False}], "1 项", "1 项不符合规范"),
        ("incomplete", "模型无法下结论：图片空白", [], "未完成", "模型无法下结论"),
        ("passed", None, [{"role": "body", "field": "size_pt",
                           "pass": True}], "0", "视觉复核通过"),
    ]
    for status, error, issues, metric, text in cases:
        at.session_state["last_result"] = {
            "result": _app_result(tmp_path, status, error, issues),
            "source_name": "状态测试.docx",
            "run_id": "20260101-000000",
            "compare": {"before": [], "after": [], "error": None},
        }
        at.run()
        assert list(at.exception) == [], (status, error)
        metrics = [item for item in at.metric if item.label == "视觉问题"]
        assert len(metrics) == 1, (status, error)
        assert metrics[0].value == metric, (status, error, metrics[0].value)
        messages = [element.value for element in
                    list(at.info) + list(at.warning) + list(at.error)
                    + list(at.success)]
        assert any(text in message for message in messages), (
            status, error, messages)
