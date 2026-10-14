"""파일 관리/변환 도구."""
from __future__ import annotations

import shutil

from .com import (
    EXCEL_EXT, EXCEL_FORMATS, PPT_EXT, PPT_FORMATS, WORKDIR, OfficeError, excel_workbook,
    office_tool, ppt_presentation, safe_path,
)


@office_tool
def list_office_files(directory: str = ".") -> dict:
    """폴더 안의 Excel/PowerPoint 파일 목록(하위 폴더 포함)을 반환한다."""
    base = safe_path(directory, must_exist=True)
    exts = EXCEL_EXT | PPT_EXT
    files = []
    for p in sorted(base.rglob("*")):
        if p.is_file() and p.suffix.lower() in exts and not p.name.startswith("~$"):
            files.append({"path": str(p.relative_to(WORKDIR)) if WORKDIR in p.parents else str(p),
                          "size_kb": round(p.stat().st_size / 1024, 1)})
    return {"files": files}


@office_tool
def copy_file(src: str, dst: str, overwrite: bool = False) -> dict:
    """파일 복사. 원본(input)을 보호하려면 수정 전에 output 폴더로 복사한 뒤 복사본에서 작업한다."""
    s = safe_path(src, must_exist=True)
    d = safe_path(dst, write=True)
    if d.exists() and not overwrite:
        raise OfficeError(f"이미 존재합니다(overwrite=true 로 덮어쓰기): {d}")
    d.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(s, d)
    return {"copied": str(d)}


@office_tool
def convert_file(src: str, dst: str, overwrite: bool = False) -> dict:
    """형식 변환(확장자로 결정). 예: a.xls→a.xlsx, a.ppt→a.pptx, 엑셀/PPT→PDF, xlsx→csv."""
    s = safe_path(src, must_exist=True)
    d = safe_path(dst, write=True)
    if d.exists() and not overwrite:
        raise OfficeError(f"이미 존재합니다(overwrite=true 로 덮어쓰기): {d}")
    d.parent.mkdir(parents=True, exist_ok=True)
    se, de = s.suffix.lower(), d.suffix.lower()
    if se in EXCEL_EXT:
        with excel_workbook(str(s)) as (app, wb):
            if de == ".pdf":
                wb.ExportAsFixedFormat(0, str(d))
            elif de in EXCEL_FORMATS:
                wb.SaveAs(str(d), EXCEL_FORMATS[de])
            else:
                raise OfficeError(f"변환 불가 대상 확장자: {de}")
    elif se in PPT_EXT:
        if de not in PPT_FORMATS:
            raise OfficeError(f"변환 불가 대상 확장자: {de}")
        with ppt_presentation(str(s)) as (app, pres):
            pres.SaveAs(str(d), PPT_FORMATS[de])
    else:
        raise OfficeError(f"지원하지 않는 원본 형식: {se}")
    return {"converted": str(d)}
