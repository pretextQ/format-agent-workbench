"""R2-3 回归：AtomicOutputSet 整组提交失败必须整组回滚（审计 R-12/B6）。

旧实现逐个 os.replace、成功失败都无回滚：第 k 个替换失败时，前 k-1 个
目标已是新内容、其余是旧内容（半新半旧），且临时文件可能残留。
本文件用 monkeypatch 在第 k 个 os.replace 注入一次性失败，断言：
- 异常原样抛出（同一异常对象）；
- 四个最终产物全部保持旧字节（整组回滚）；
- 无任何临时/回滚副本残留（目录只剩四个最终产物）。
同时覆盖成功路径与替换顺序（主产物最后替换）。
"""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core import safe_output
from core.safe_output import AtomicOutputSet

# commit 顺序 = 非 main 键按插入序 + main（out_path）最后
KEYS = ["report_path", "tracked_path", "report_docx_path", "out_path"]
FILENAMES = {
    "report_path": "result_report.md",
    "tracked_path": "result_tracked.docx",
    "report_docx_path": "result_report.docx",
    "out_path": "result.docx",
}


def _old_bytes():
    return {key: f"OLD-{key}".encode() for key in KEYS}


def _new_bytes():
    return {key: f"NEW-{key}".encode() for key in KEYS}


def _build_group(tmp_path):
    """先落一组旧产物，再准备一个已写入新内容的提交组。"""
    final_paths = {key: str(tmp_path / FILENAMES[key]) for key in KEYS}
    for key, path in final_paths.items():
        Path(path).write_bytes(_old_bytes()[key])
    group = AtomicOutputSet(final_paths)
    for key, temp in group.temp_paths.items():
        Path(temp).write_bytes(_new_bytes()[key])
    return group, final_paths


def _inject_nth_replace_failure(monkeypatch, n, error):
    """让第 n 次 os.replace 抛出 error（只失败这一次，之后恢复原行为）。

    通过 safe_output.os 注入（safe_output.os 即全局 os 模块），
    monkeypatch 结束自动还原。
    """
    real_replace = safe_output.os.replace
    calls = {"count": 0}

    def flaky_replace(src, dst):
        calls["count"] += 1
        if calls["count"] == n:
            raise error
        return real_replace(src, dst)

    monkeypatch.setattr(safe_output.os, "replace", flaky_replace)
    return calls


def _assert_all_old_and_no_residue(final_paths):
    """四份最终产物必须全部是旧字节；目录里不得残留临时/回滚副本。"""
    old = _old_bytes()
    for key in KEYS:
        actual = Path(final_paths[key]).read_bytes()
        assert actual == old[key], (
            f"{key} 半新半旧：期望 {old[key]!r}，实际 {actual!r}")
    directory = Path(final_paths[KEYS[0]]).parent
    leftovers = sorted(p.name for p in directory.iterdir())
    expected = sorted(FILENAMES[key] for key in KEYS)
    assert leftovers == expected, f"残留多余文件：{leftovers} != {expected}"


def test_first_replace_failure_rolls_back_entire_group(tmp_path, monkeypatch):
    """第 1 个 os.replace 失败：四份旧产物必须原样、无临时残留。"""
    group, final_paths = _build_group(tmp_path)
    injected = OSError("injected failure at replace #1")
    _inject_nth_replace_failure(monkeypatch, 1, injected)

    with pytest.raises(OSError) as excinfo:
        group.commit()

    assert excinfo.value is injected, "异常必须原样抛出（同一对象）"
    _assert_all_old_and_no_residue(final_paths)


def test_last_replace_failure_rolls_back_entire_group(tmp_path, monkeypatch):
    """最后一个（主产物 out_path）os.replace 失败：前几个已替换也必须回滚。"""
    group, final_paths = _build_group(tmp_path)
    injected = OSError(f"injected failure at replace #{len(KEYS)}")
    _inject_nth_replace_failure(monkeypatch, len(KEYS), injected)

    with pytest.raises(OSError) as excinfo:
        group.commit()

    assert excinfo.value is injected, "异常必须原样抛出（同一对象）"
    _assert_all_old_and_no_residue(final_paths)


def test_success_replaces_all_in_order_without_residue(tmp_path, monkeypatch):
    """成功路径行为不变：全部换成新内容、主产物最后替换、无残留。"""
    group, final_paths = _build_group(tmp_path)

    real_replace = safe_output.os.replace
    replaced_destinations = []

    def recording_replace(src, dst):
        replaced_destinations.append(dst)
        return real_replace(src, dst)

    monkeypatch.setattr(safe_output.os, "replace", recording_replace)
    assert group.commit() is None

    for key in KEYS:
        assert Path(final_paths[key]).read_bytes() == _new_bytes()[key]
    # 顺序不变：非 main 键按插入序，主产物最后（完成标志）
    assert replaced_destinations == [final_paths[key] for key in KEYS]
    leftovers = sorted(p.name for p in tmp_path.iterdir())
    assert leftovers == sorted(FILENAMES[key] for key in KEYS), (
        f"成功路径残留多余文件：{leftovers}")
