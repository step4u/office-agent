"""PowerPoint(.ppt/.pptx) COM 도구. 각 함수는 MCP tool 로 노출된다."""
from __future__ import annotations

from typing import Any

from .com import (
    OfficeError, chart_type_id, grid, office_tool, ppt_presentation, rect_rows, rgb, safe_path, to_py,
)

Cell = str | int | float | bool | None

_SHAPE_TYPES = {1: "autoshape", 3: "chart", 6: "group", 13: "picture", 14: "placeholder",
                17: "textbox", 19: "table", 24: "smartart", 16: "media"}


def _items(coll):
    for i in range(1, coll.Count + 1):
        yield coll.Item(i)


def _slide(pres, idx: int):
    n = pres.Slides.Count
    if not 1 <= idx <= n:
        raise OfficeError(f"슬라이드 번호 범위 오류: {idx} (1~{n})")
    return pres.Slides.Item(idx)


def _dims(pres) -> tuple[float, float]:
    return pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight


def _find_shape(slide, name: str):
    for shp in _items(slide.Shapes):
        if shp.Name == name:
            return shp
    raise OfficeError(f"도형 '{name}' 없음. 존재: {[s.Name for s in _items(slide.Shapes)]}")


def _placeholder(slide, types: tuple[int, ...]):
    for shp in _items(slide.Shapes):
        try:
            if shp.Type == 14 and shp.PlaceholderFormat.Type in types:
                return shp
        except Exception:
            continue
    return None


def _table_rows(tbl, cap: int = 40) -> list[list[str]]:
    rows = []
    for r in range(1, min(tbl.Rows.Count, cap) + 1):
        rows.append([tbl.Cell(r, c).Shape.TextFrame.TextRange.Text.replace("\r", "\n")
                     for c in range(1, tbl.Columns.Count + 1)])
    return rows


def _chart_info(ch) -> dict:
    info: dict[str, Any] = {"chart_type": ch.ChartType}
    try:
        if ch.HasTitle:
            info["title"] = ch.ChartTitle.Text
    except Exception:
        pass
    series = []
    try:
        for s in _items(ch.SeriesCollection()):
            series.append({"name": s.Name,
                           "categories": [to_py(x) for x in list(s.XValues)[:50]],
                           "values": [to_py(x) for x in list(s.Values)[:50]]})
    except Exception:
        pass
    info["series"] = series
    return info


def _shape_info(shp, depth: int = 0) -> dict:
    info: dict[str, Any] = {
        "name": shp.Name, "type": _SHAPE_TYPES.get(shp.Type, shp.Type),
        "left": round(shp.Left, 1), "top": round(shp.Top, 1),
        "width": round(shp.Width, 1), "height": round(shp.Height, 1),
    }
    try:
        if shp.HasTable:
            info["table"] = _table_rows(shp.Table)
    except Exception:
        pass
    try:
        if shp.HasChart:
            info["chart"] = _chart_info(shp.Chart)
    except Exception:
        pass
    try:
        if shp.HasTextFrame and shp.TextFrame.HasText:
            info["text"] = shp.TextFrame.TextRange.Text.replace("\r", "\n")
    except Exception:
        pass
    if shp.Type == 6 and depth < 3:
        info["children"] = [_shape_info(c, depth + 1) for c in _items(shp.GroupItems)]
    return info


@office_tool
def ppt_read(path: str) -> dict:
    """PPT 전체 내용을 읽는다(슬라이드별 레이아웃, 도형 위치/크기, 텍스트, 표, 차트 데이터, 노트, 슬라이드 크기)."""
    with ppt_presentation(path) as (app, pres):
        w, h = _dims(pres)
        slides = []
        for s in _items(pres.Slides):
            notes = ""
            try:
                notes = s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text
            except Exception:
                pass
            slides.append({"index": s.SlideIndex, "layout": s.CustomLayout.Name,
                           "shapes": [_shape_info(x) for x in _items(s.Shapes)], "notes": notes})
        return {"file": path, "slide_width": w, "slide_height": h, "slide_count": len(slides), "slides": slides}


@office_tool
def ppt_list_layouts(path: str) -> dict:
    """슬라이드 마스터의 레이아웃 목록(번호, 이름, 플레이스홀더). ppt_add_slide 의 layout_index 선택용."""
    with ppt_presentation(path) as (app, pres):
        out = []
        for i, lay in enumerate(_items(pres.SlideMaster.CustomLayouts), 1):
            phs = []
            for shp in _items(lay.Shapes):
                try:
                    if shp.Type == 14:
                        phs.append({"name": shp.Name, "placeholder_type": shp.PlaceholderFormat.Type})
                except Exception:
                    pass
            out.append({"index": i, "name": lay.Name, "placeholders": phs})
        return {"layouts": out}


@office_tool
def ppt_create(path: str, template_path: str | None = None, overwrite: bool = False) -> dict:
    """새 프레젠테이션 생성. template_path(.potx/.pptx)를 주면 회사 템플릿(마스터/테마)을 그대로 사용한다."""
    with ppt_presentation(path, "new", template=template_path, overwrite=overwrite) as (app, pres):
        w, h = _dims(pres)
        n = pres.Slides.Count
    return {"created": path, "slide_width": w, "slide_height": h, "existing_slides": n}


@office_tool
def ppt_add_slide(path: str, layout_index: int = 2, title: str | None = None,
                  bullets: list[str] | None = None, notes: str | None = None,
                  position: int | None = None) -> dict:
    """슬라이드 추가. layout_index 는 ppt_list_layouts 번호(기본 Office 테마: 1=제목,2=제목+내용,6=제목만,7=빈 화면).
    bullets 는 본문(또는 부제목) 플레이스홀더에 들어가며, 앞에 탭을 붙이면 들여쓰기 수준이 올라간다."""
    with ppt_presentation(path, "w") as (app, pres):
        layouts = pres.SlideMaster.CustomLayouts
        if not 1 <= layout_index <= layouts.Count:
            raise OfficeError(f"layout_index 범위 오류: 1~{layouts.Count}")
        idx = position or pres.Slides.Count + 1
        slide = pres.Slides.AddSlide(idx, layouts.Item(layout_index))
        if title:
            t = _placeholder(slide, (1, 3))
            if t is None:
                raise OfficeError("이 레이아웃에는 제목 플레이스홀더가 없습니다.")
            t.TextFrame.TextRange.Text = title
        if bullets:
            body = _placeholder(slide, (2, 4, 7))
            if body is None:
                raise OfficeError("이 레이아웃에는 본문 플레이스홀더가 없습니다. ppt_add_text_box 를 사용하세요.")
            clean, levels = [], []
            for b in bullets:
                s = b.lstrip("\t")
                clean.append(s)
                levels.append(len(b) - len(s) + 1)
            tr = body.TextFrame.TextRange
            tr.Text = "\r".join(clean)
            for i, lv in enumerate(levels, 1):
                if lv > 1:
                    tr.Paragraphs(i, 1).IndentLevel = lv
        if notes:
            slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = notes
        return {"slide_index": slide.SlideIndex, "shapes": [s.Name for s in _items(slide.Shapes)]}


@office_tool
def ppt_add_text_box(path: str, slide_index: int, text: str, left: float = 40, top: float = 110,
                     width: float | None = None, height: float = 60, font_size: float = 18,
                     bold: bool = False, color_hex: str | None = None, align: str = "left",
                     font_name: str | None = None, shape_name: str | None = None) -> dict:
    """텍스트 상자 추가(단위: pt). width 생략 시 슬라이드 폭에서 좌우 여백 제외. align: left|center|right."""
    with ppt_presentation(path, "w") as (app, pres):
        slide = _slide(pres, slide_index)
        sw, _ = _dims(pres)
        shp = slide.Shapes.AddTextbox(1, left, top, width or sw - 2 * left, height)
        tr = shp.TextFrame.TextRange
        tr.Text = text
        tr.Font.Size = font_size
        tr.Font.Bold = bold
        if color_hex:
            tr.Font.Color.RGB = rgb(color_hex)
        if font_name:
            tr.Font.Name = font_name
        tr.ParagraphFormat.Alignment = {"left": 1, "center": 2, "right": 3}[align]
        shp.TextFrame.WordWrap = True
        if shape_name:
            shp.Name = shape_name
        return {"shape": shp.Name}


@office_tool
def ppt_update_text(path: str, slide_index: int, shape_name: str, text: str) -> dict:
    """기존 도형(플레이스홀더 포함)의 텍스트를 교체한다. 도형 이름은 ppt_read 로 확인."""
    with ppt_presentation(path, "w") as (app, pres):
        _find_shape(_slide(pres, slide_index), shape_name).TextFrame.TextRange.Text = text
    return {"updated": shape_name}


def _walk_ranges(shapes):
    for shp in _items(shapes):
        try:
            if shp.Type == 6:
                yield from _walk_ranges(shp.GroupItems)
                continue
            if shp.HasTable:
                t = shp.Table
                for r in range(1, t.Rows.Count + 1):
                    for c in range(1, t.Columns.Count + 1):
                        yield t.Cell(r, c).Shape.TextFrame.TextRange
            elif shp.HasTextFrame and shp.TextFrame.HasText:
                yield shp.TextFrame.TextRange
        except Exception:
            continue


@office_tool
def ppt_replace_text(path: str, find: str, replace: str) -> dict:
    """전체 슬라이드(표 포함)에서 텍스트를 찾아 바꾼다. 템플릿의 {{자리표시}} 채우기에 유용."""
    if not find:
        raise OfficeError("find 가 비어 있습니다.")
    count = 0
    with ppt_presentation(path, "w") as (app, pres):
        for s in _items(pres.Slides):
            for tr in _walk_ranges(s.Shapes):
                for _ in range(200):
                    if tr.Replace(find, replace) is None:
                        break
                    count += 1
    return {"replaced": count}


@office_tool
def ppt_add_table(path: str, slide_index: int, rows: list[list[Cell]], left: float = 40, top: float = 110,
                  width: float | None = None, height: float | None = None, font_size: float = 14,
                  header_fill_hex: str = "1F4E79", header_font_hex: str = "FFFFFF") -> dict:
    """표 추가. rows[0] 은 헤더 행. 숫자는 오른쪽 정렬된다."""
    data, nr, nc = rect_rows(rows)
    with ppt_presentation(path, "w") as (app, pres):
        slide = _slide(pres, slide_index)
        sw, _ = _dims(pres)
        shp = slide.Shapes.AddTable(nr, nc, left, top, width or sw - 2 * left, height or nr * (font_size + 14))
        tbl = shp.Table
        for r in range(nr):
            for c in range(nc):
                v = data[r][c]
                cell = tbl.Cell(r + 1, c + 1)
                tr = cell.Shape.TextFrame.TextRange
                tr.Text = "" if v == "" else (f"{v:,}" if isinstance(v, int) and not isinstance(v, bool) else str(v))
                tr.Font.Size = font_size
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    tr.ParagraphFormat.Alignment = 3
                if r == 0:
                    cell.Shape.Fill.ForeColor.RGB = rgb(header_fill_hex)
                    tr.Font.Color.RGB = rgb(header_font_hex)
                    tr.Font.Bold = True
        return {"shape": shp.Name}


@office_tool
def ppt_add_chart(path: str, slide_index: int, data: list[list[Cell]], chart_type: str = "column",
                  title: str | None = None, left: float = 40, top: float = 110,
                  width: float | None = None, height: float | None = None,
                  series_in_columns: bool = True, data_labels: bool = False, legend: bool = True,
                  shape_name: str | None = None) -> dict:
    """편집 가능한 네이티브 PPT 차트 추가. data[0] 은 헤더(첫 칸은 빈 문자열 가능), 첫 열은 카테고리.
    예: [["", "매출", "비용"], ["1월", 100, 60], ["2월", 120, 70]].
    chart_type: column|column_stacked|bar|bar_stacked|line|line_markers|pie|doughnut|area|area_stacked|scatter."""
    ct = chart_type_id(chart_type)
    vals, nr, nc = rect_rows(data)
    if nr < 2 or nc < 2:
        raise OfficeError("차트 data 는 최소 2행 2열(헤더+데이터)이어야 합니다.")
    with ppt_presentation(path, "w") as (app, pres):
        slide = _slide(pres, slide_index)
        sw, sh = _dims(pres)
        shp = slide.Shapes.AddChart2(-1, ct, left, top, width or sw - 2 * left, height or sh - top - 40)
        ch = shp.Chart
        try:
            ch.ChartData.ActivateChartDataWindow()
        except Exception:
            ch.ChartData.Activate()
        wb = ch.ChartData.Workbook
        try:
            ws = wb.Worksheets.Item(1)
            ws.Range("A1:Z60").ClearContents()
            rng = ws.Range("A1").Resize(nr, nc)
            rng.Value = vals
            try:
                ws.ListObjects.Item(1).Resize(rng)
            except Exception:
                pass
            ch.SetSourceData(f"='{ws.Name}'!{rng.Address}", 2 if series_in_columns else 1)
        finally:
            try:
                wb.Close()
            except Exception:
                pass
        if title:
            ch.HasTitle = True
            ch.ChartTitle.Text = title
        else:
            ch.HasTitle = False
        ch.HasLegend = legend
        if data_labels:
            for i in range(1, ch.SeriesCollection().Count + 1):
                try:
                    ch.SeriesCollection(i).HasDataLabels = True
                except Exception:
                    pass
        if shape_name:
            shp.Name = shape_name
        return {"shape": shp.Name}


@office_tool
def ppt_add_picture(path: str, slide_index: int, image_path: str, left: float = 40, top: float = 110,
                    width: float = -1, height: float = -1) -> dict:
    """이미지(PNG 등) 삽입. width/height 를 -1 로 두면 원본 비율 크기."""
    img = safe_path(image_path, must_exist=True)
    with ppt_presentation(path, "w") as (app, pres):
        shp = _slide(pres, slide_index).Shapes.AddPicture(str(img), 0, -1, left, top, width, height)
        return {"shape": shp.Name, "width": shp.Width, "height": shp.Height}


@office_tool
def ppt_set_notes(path: str, slide_index: int, notes: str) -> dict:
    """슬라이드 발표자 노트 설정."""
    with ppt_presentation(path, "w") as (app, pres):
        _slide(pres, slide_index).NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = notes
    return {"slide_index": slide_index}


@office_tool
def ppt_delete_slide(path: str, slide_index: int) -> dict:
    """슬라이드 삭제."""
    with ppt_presentation(path, "w") as (app, pres):
        _slide(pres, slide_index).Delete()
        return {"remaining": pres.Slides.Count}
