"""Office COM 공통 유틸: 경로 보호, COM 세션, 값 변환, 에러 처리."""
from __future__ import annotations

import functools
import os
from contextlib import contextmanager
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

try:  # Windows + pywin32 에서만 사용 가능
    import pythoncom
    import pywintypes
    import win32com.client
except ImportError:  # pragma: no cover
    pythoncom = pywintypes = win32com = None  # type: ignore


class OfficeError(Exception):
    """LLM에게 그대로 전달되는 사용자 친화적 오류."""


# ---------------------------------------------------------------- 설정
WORKDIR = Path(os.environ.get("OFFICE_AGENT_WORKDIR") or os.getcwd()).resolve()
MAX_CELLS = int(os.environ.get("OFFICE_AGENT_MAX_CELLS", "5000"))
PPT_WINDOW = os.environ.get("OFFICE_AGENT_PPT_WINDOW", "0") == "1"


def _dirs(env: str) -> list[Path]:
    raw = os.environ.get(env, "")
    return [(WORKDIR / x.strip()).resolve() for x in raw.split(";") if x.strip()]


ALLOWED_DIRS = [WORKDIR] + _dirs("OFFICE_AGENT_ALLOWED_DIRS")
READONLY_DIRS = _dirs("OFFICE_AGENT_READONLY_DIRS")

EXCEL_EXT = {".xls", ".xlsx", ".xlsm", ".xlsb", ".csv"}
PPT_EXT = {".ppt", ".pptx", ".pptm", ".potx"}

EXCEL_FORMATS = {".xlsx": 51, ".xlsm": 52, ".xlsb": 50, ".xls": 56, ".csv": 6}
PPT_FORMATS = {".pptx": 24, ".pptm": 25, ".potx": 26, ".ppt": 1, ".pdf": 32}

# xlChartType 상수 (Excel/PowerPoint 공통)
CHART_TYPES = {
    "column": 51,
    "column_stacked": 52,
    "bar": 57,
    "bar_stacked": 58,
    "line": 4,
    "line_markers": 65,
    "pie": 5,
    "doughnut": -4120,
    "area": 1,
    "area_stacked": 76,
    "scatter": -4169,
}


def chart_type_id(name: str) -> int:
    try:
        return CHART_TYPES[name]
    except KeyError:
        raise OfficeError(f"지원하지 않는 차트 종류: {name}. 가능: {sorted(CHART_TYPES)}") from None


# ---------------------------------------------------------------- 경로
def _under(p: Path, d: Path) -> bool:
    ps, ds = os.path.normcase(str(p)), os.path.normcase(str(d)).rstrip("\\/")
    return ps == ds or ps.startswith(ds + os.sep)


def safe_path(path: str, *, must_exist: bool = False, write: bool = False) -> Path:
    """작업 폴더(및 허용 폴더) 밖의 접근을 차단하고 절대경로를 반환."""
    p = Path(path)
    if not p.is_absolute():
        p = WORKDIR / p
    p = p.resolve()
    if not any(_under(p, d) for d in ALLOWED_DIRS):
        raise OfficeError(f"허용되지 않은 경로입니다: {p} (허용: {[str(d) for d in ALLOWED_DIRS]})")
    if write and any(_under(p, d) for d in READONLY_DIRS):
        raise OfficeError(f"읽기 전용 폴더입니다. output 폴더에 복사본을 만들어 작업하세요: {p}")
    if must_exist and not p.exists():
        raise OfficeError(f"파일이 없습니다: {p}")
    return p


def check_ext(p: Path, allowed: set[str]) -> None:
    if p.suffix.lower() not in allowed:
        raise OfficeError(f"지원하지 않는 확장자 {p.suffix}: {sorted(allowed)}")


def require_windows() -> None:
    if pythoncom is None:
        raise OfficeError("Windows + pywin32 + 설치된 Microsoft Office 가 필요합니다.")


# ---------------------------------------------------------------- 값 변환
_ERR = {
    -2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?",
    -2146826288: "#NULL!", -2146826252: "#NUM!", -2146826265: "#REF!",
    -2146826273: "#VALUE!",
}


def to_py(v: Any) -> Any:
    if v is None or isinstance(v, (bool, str)):
        return v
    if isinstance(v, (datetime, date)):
        return v.isoformat()
    if isinstance(v, Decimal):
        return float(v)
    if isinstance(v, float):
        return int(v) if v.is_integer() and abs(v) < 1e15 else v
    if isinstance(v, int):
        return _ERR.get(v, v)
    return str(v)


def grid(v: Any) -> list[list[Any]]:
    """Range.Value/Formula 결과를 항상 2차원 리스트로."""
    if isinstance(v, tuple):
        return [[to_py(c) for c in row] if isinstance(row, tuple) else [to_py(row)] for row in v]
    return [[to_py(v)]]


def rgb(hex_color: str) -> int:
    h = hex_color.lstrip("#")
    if len(h) != 6:
        raise OfficeError(f"색상은 RRGGBB 형식이어야 합니다: {hex_color}")
    return int(h[0:2], 16) + int(h[2:4], 16) * 256 + int(h[4:6], 16) * 65536


def rect_rows(values: list[list[Any]]) -> tuple[tuple[Any, ...], int, int]:
    """2차원 리스트를 직사각형으로 패딩하여 COM 에 넣을 튜플로 변환."""
    if not values or not any(values):
        raise OfficeError("values 가 비어 있습니다.")
    rows = [list(r) if isinstance(r, (list, tuple)) else [r] for r in values]
    width = max(len(r) for r in rows)
    data = tuple(tuple("" if c is None else c for c in r) + ("",) * (width - len(r)) for r in rows)
    return data, len(rows), width


# ---------------------------------------------------------------- 에러 래퍼
def _com_msg(e: Exception) -> str:
    try:
        info = e.excepinfo  # type: ignore[attr-defined]
        if info and info[2]:
            return f"{info[2]} (0x{e.hresult & 0xFFFFFFFF:08X})"  # type: ignore[attr-defined]
    except Exception:
        pass
    return str(e)


def office_tool(fn):
    """COM 예외를 OfficeError 로 변환 (MCP 가 오류 메시지를 LLM에 전달)."""

    @functools.wraps(fn)
    def wrapper(*a, **kw):
        try:
            return fn(*a, **kw)
        except OfficeError:
            raise
        except Exception as e:  # noqa: BLE001
            if pywintypes is not None and isinstance(e, pywintypes.com_error):
                raise OfficeError(f"COM 오류: {_com_msg(e)}") from None
            raise OfficeError(f"{type(e).__name__}: {e}") from e

    return wrapper


# ---------------------------------------------------------------- Excel
@contextmanager
def excel_app():
    """독립된 Excel 프로세스 (사용자가 열어둔 Excel 과 분리)."""
    require_windows()
    pythoncom.CoInitialize()
    app = None
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.ScreenUpdating = False
        app.EnableEvents = False
        app.AskToUpdateLinks = False
        yield app
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        app = None
        pythoncom.CoUninitialize()


@contextmanager
def excel_workbook(path: str, mode: str = "r", overwrite: bool = False):
    """mode: r=읽기전용, w=열어서 수정 후 저장, new=새 통합문서 생성 후 저장."""
    p = safe_path(path, must_exist=(mode != "new"), write=(mode != "r"))
    check_ext(p, EXCEL_EXT)
    if mode == "new" and p.exists() and not overwrite:
        raise OfficeError(f"이미 존재합니다(overwrite=true 로 덮어쓰기): {p}")
    with excel_app() as app:
        wb = app.Workbooks.Add() if mode == "new" else app.Workbooks.Open(str(p), 0, mode == "r")
        try:
            yield app, wb
            if mode == "w":
                wb.Save()
            elif mode == "new":
                p.parent.mkdir(parents=True, exist_ok=True)
                wb.SaveAs(str(p), EXCEL_FORMATS[p.suffix.lower()])
        finally:
            try:
                wb.Close(False)
            except Exception:
                pass


# ---------------------------------------------------------------- PowerPoint
@contextmanager
def ppt_app():
    require_windows()
    pythoncom.CoInitialize()
    app = None
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        try:
            app.DisplayAlerts = 1  # ppAlertsNone
        except Exception:
            pass
        yield app
    finally:
        if app is not None:
            try:  # PowerPoint 는 단일 인스턴스: 사용자가 열어 둔 파일이 없을 때만 종료
                if app.Presentations.Count == 0:
                    app.Quit()
            except Exception:
                pass
        app = None
        pythoncom.CoUninitialize()


@contextmanager
def ppt_presentation(path: str, mode: str = "r", template: str | None = None, overwrite: bool = False):
    p = safe_path(path, must_exist=(mode != "new"), write=(mode != "r"))
    check_ext(p, PPT_EXT)
    if mode == "new" and p.exists() and not overwrite:
        raise OfficeError(f"이미 존재합니다(overwrite=true 로 덮어쓰기): {p}")
    win = -1 if PPT_WINDOW else 0
    with ppt_app() as app:
        if mode == "new":
            if template:
                tp = safe_path(template, must_exist=True)
                pres = app.Presentations.Open(str(tp), -1, -1, win)  # ReadOnly, Untitled
            else:
                pres = app.Presentations.Add(win)
        else:
            pres = app.Presentations.Open(str(p), -1 if mode == "r" else 0, 0, win)
        try:
            yield app, pres
            if mode == "w":
                pres.Save()
            elif mode == "new":
                p.parent.mkdir(parents=True, exist_ok=True)
                pres.SaveAs(str(p), PPT_FORMATS[p.suffix.lower()])
        finally:
            try:
                pres.Close()
            except Exception:
                pass
