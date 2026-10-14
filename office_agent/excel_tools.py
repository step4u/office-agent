"""Excel(.xls/.xlsx/.xlsm/.csv) COM 도구. 각 함수는 MCP tool 로 노출된다."""
from __future__ import annotations

from typing import Any

from .com import (
    MAX_CELLS, OfficeError, chart_type_id, excel_workbook, grid, office_tool,
    rect_rows, rgb, safe_path,
)

Cell = str | int | float | bool | None


def _sheet(wb, name: str | None = None, create: bool = False):
    if not name:
        return wb.ActiveSheet
    names = [wb.Worksheets.Item(i).Name for i in range(1, wb.Worksheets.Count + 1)]
    if name in names:
        return wb.Worksheets(name)
    if create:
        ws = wb.Worksheets.Add(After=wb.Worksheets.Item(wb.Worksheets.Count))
        ws.Name = name
        return ws
    raise OfficeError(f"시트 '{name}' 없음. 사용 가능: {names}")


@office_tool
def excel_inspect(path: str, preview_rows: int = 5) -> dict:
    """Excel 파일의 구조를 요약한다(시트 목록, 사용 범위, 미리보기, 차트/표 이름). 데이터 작업 전에 먼저 호출할 것."""
    with excel_workbook(path) as (app, wb):
        sheets = []
        for i in range(1, wb.Worksheets.Count + 1):
            ws = wb.Worksheets.Item(i)
            ur = ws.UsedRange
            rows, cols = ur.Rows.Count, ur.Columns.Count
            prev = ur.Resize(min(rows, preview_rows), min(cols, 20))
            sheets.append({
                "name": ws.Name,
                "visible": ws.Visible == -1,
                "used_range": ur.Address(False, False),
                "rows": rows,
                "cols": cols,
                "preview": grid(prev.Value),
                "charts": [ws.ChartObjects().Item(k).Name for k in range(1, ws.ChartObjects().Count + 1)],
                "tables": [ws.ListObjects.Item(k).Name for k in range(1, ws.ListObjects.Count + 1)],
            })
        return {"file": path, "sheets": sheets}


@office_tool
def excel_read_range(path: str, sheet: str | None = None, cell_range: str | None = None,
                     include_formulas: bool = False, max_cells: int = 0) -> dict:
    """셀 범위 값을 2차원 배열로 읽는다. cell_range 생략 시 사용 범위 전체(max_cells 초과 시 위쪽 행만 반환, truncated=true)."""
    with excel_workbook(path) as (app, wb):
        ws = _sheet(wb, sheet)
        rng = ws.Range(cell_range) if cell_range else ws.UsedRange
        rows, cols = rng.Rows.Count, rng.Columns.Count
        limit = max_cells or MAX_CELLS
        read_rows = rows if rows * cols <= limit else max(1, limit // cols)
        sub = rng.Resize(read_rows, cols) if read_rows < rows else rng
        out: dict[str, Any] = {
            "sheet": ws.Name,
            "address": sub.Address(False, False),
            "total_rows": rows,
            "returned_rows": read_rows,
            "truncated": read_rows < rows,
            "values": grid(sub.Value),
        }
        if include_formulas:
            out["formulas"] = grid(sub.Formula)
        return out


@office_tool
def excel_evaluate(path: str, formula: str, sheet: str | None = None) -> dict:
    """Excel 수식을 파일 변경 없이 계산한다. 예: 'SUMIFS(B:B,A:A,"서울")', 'AVERAGE(B2:B100)'. 앞의 '=' 는 생략 가능."""
    with excel_workbook(path) as (app, wb):
        ws = _sheet(wb, sheet)
        return {"formula": formula, "result": grid(ws.Evaluate(formula.lstrip("=")))}


@office_tool
def excel_create_workbook(path: str, sheet_name: str = "Sheet1", overwrite: bool = False) -> dict:
    """새 통합문서(.xlsx/.xls/.xlsm)를 만든다."""
    with excel_workbook(path, "new", overwrite=overwrite) as (app, wb):
        wb.Worksheets.Item(1).Name = sheet_name
    return {"created": path, "sheet": sheet_name}


@office_tool
def excel_write_range(path: str, sheet: str, top_left: str, values: list[list[Cell]],
                      create_sheet: bool = False, autofit: bool = True) -> dict:
    """top_left 셀부터 2차원 배열을 쓴다. '=' 로 시작하는 문자열은 수식으로 입력된다(예: '=SUM(B2:B5)').
    새 시트가 필요하면 create_sheet=true. 저장까지 자동 수행."""
    data, nr, nc = rect_rows(values)
    with excel_workbook(path, "w") as (app, wb):
        ws = _sheet(wb, sheet, create=create_sheet)
        rng = ws.Range(top_left).Resize(nr, nc)
        rng.Formula = data
        if autofit:
            rng.Columns.AutoFit()
        app.Calculate()
        return {"sheet": ws.Name, "written": rng.Address(False, False)}


@office_tool
def excel_format_range(path: str, sheet: str, cell_range: str, number_format: str | None = None,
                       bold: bool | None = None, fill_hex: str | None = None, font_hex: str | None = None,
                       border: bool = False, horizontal_align: str | None = None,
                       autofit_columns: bool = False, column_width: float | None = None) -> dict:
    """범위 서식 지정. number_format 예: '#,##0', '0.0%', 'yyyy-mm-dd'. 색상은 RRGGBB. horizontal_align: left|center|right."""
    align = {"left": -4131, "center": -4108, "right": -4152}
    with excel_workbook(path, "w") as (app, wb):
        ws = _sheet(wb, sheet)
        rng = ws.Range(cell_range)
        if number_format:
            rng.NumberFormat = number_format
        if bold is not None:
            rng.Font.Bold = bold
        if fill_hex:
            rng.Interior.Color = rgb(fill_hex)
        if font_hex:
            rng.Font.Color = rgb(font_hex)
        if border:
            rng.Borders.LineStyle = 1
        if horizontal_align:
            rng.HorizontalAlignment = align[horizontal_align]
        if autofit_columns:
            rng.Columns.AutoFit()
        if column_width:
            rng.ColumnWidth = column_width
        return {"formatted": rng.Address(False, False)}


@office_tool
def excel_add_chart(path: str, sheet: str, source_range: str, chart_type: str = "column",
                    title: str | None = None, anchor_cell: str = "H2", width: float = 480, height: float = 300,
                    x_title: str | None = None, y_title: str | None = None, series_in_columns: bool = True,
                    data_labels: bool = False, legend: bool = True, chart_name: str | None = None,
                    chart_sheet: str | None = None) -> dict:
    """데이터 범위(헤더 포함, 예: 'A1:C13')로 Excel 차트를 만든다.
    chart_type: column|column_stacked|bar|bar_stacked|line|line_markers|pie|doughnut|area|area_stacked|scatter.
    series_in_columns=true 면 열이 계열(첫 행=계열명, 첫 열=카테고리)."""
    ct = chart_type_id(chart_type)
    with excel_workbook(path, "w") as (app, wb):
        ws = _sheet(wb, sheet)
        wc = _sheet(wb, chart_sheet) if chart_sheet else ws
        anchor = wc.Range(anchor_cell)
        shape = wc.Shapes.AddChart2(-1, ct, anchor.Left, anchor.Top, width, height)
        ch = shape.Chart
        ch.SetSourceData(ws.Range(source_range), 2 if series_in_columns else 1)
        name = chart_name or f"chart_{wc.ChartObjects().Count}"
        shape.Name = name
        if title:
            ch.HasTitle = True
            ch.ChartTitle.Text = title
        else:
            ch.HasTitle = False
        for axis, text in ((1, x_title), (2, y_title)):
            if text:
                try:
                    ch.Axes(axis).HasTitle = True
                    ch.Axes(axis).AxisTitle.Text = text
                except Exception:
                    pass
        ch.HasLegend = legend
        if legend:
            try:
                ch.Legend.Position = -4107  # 아래
            except Exception:
                pass
        if data_labels:
            for i in range(1, ch.SeriesCollection().Count + 1):
                try:
                    ch.SeriesCollection(i).HasDataLabels = True
                except Exception:
                    pass
        return {"sheet": wc.Name, "chart_name": name}


@office_tool
def excel_export_chart(path: str, sheet: str, chart_name: str, out_path: str) -> dict:
    """Excel 차트를 PNG 이미지로 내보낸다(PPT 삽입용: ppt_add_picture 와 함께 사용)."""
    out = safe_path(out_path, write=True)
    out.parent.mkdir(parents=True, exist_ok=True)
    with excel_workbook(path) as (app, wb):
        ws = _sheet(wb, sheet)
        ws.ChartObjects(chart_name).Chart.Export(str(out), "PNG")
    return {"exported": str(out)}
