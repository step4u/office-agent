"""Windows + Office 환경 검증: python tests\\smoke_test.py (프로젝트 루트에서 실행)."""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from office_agent import excel_tools as ex  # noqa: E402
from office_agent import ppt_tools as pp  # noqa: E402

X = "output/demo.xlsx"
P = "output/demo.pptx"

print("1) Excel 생성/쓰기")
ex.excel_create_workbook(X, "Sales", overwrite=True)
ex.excel_write_range(X, "Sales", "A1", [
    ["월", "매출", "비용", "이익"],
    ["1월", 100, 60, "=B2-C2"],
    ["2월", 120, 70, "=B3-C3"],
    ["3월", 90, 65, "=B4-C4"],
    ["4월", 150, 80, "=B5-C5"],
    ["합계", "=SUM(B2:B5)", "=SUM(C2:C5)", "=SUM(D2:D5)"],
])
ex.excel_format_range(X, "Sales", "A1:D1", bold=True, fill_hex="1F4E79", font_hex="FFFFFF")
ex.excel_format_range(X, "Sales", "B2:D6", number_format="#,##0")

print("2) 수식 계산/읽기")
print(json.dumps(ex.excel_evaluate(X, "AVERAGE(B2:B5)", "Sales"), ensure_ascii=False))
print(json.dumps(ex.excel_read_range(X, "Sales", include_formulas=True), ensure_ascii=False))

print("3) Excel 차트 + PNG 내보내기")
ex.excel_add_chart(X, "Sales", "A1:C5", "column", title="월별 매출/비용", x_title="월",
                   y_title="금액(백만원)", chart_name="sales_chart")
ex.excel_export_chart(X, "Sales", "sales_chart", "output/sales_chart.png")

print("4) PPT 생성")
pp.ppt_create(P, overwrite=True)
pp.ppt_add_slide(P, 1, title="월별 실적 보고", bullets=["2026년 1~4월"])
pp.ppt_add_slide(P, 6, title="4월 매출이 가장 높았다", notes="발표자 노트 예시")
pp.ppt_add_chart(P, 2, [["", "매출", "비용"], ["1월", 100, 60], ["2월", 120, 70],
                        ["3월", 90, 65], ["4월", 150, 80]], "column", title="월별 매출/비용")
pp.ppt_add_slide(P, 6, title="요약 표")
pp.ppt_add_table(P, 3, [["월", "매출", "이익"], ["1월", 100, 40], ["2월", 120, 50]])
pp.ppt_add_slide(P, 2, title="핵심 요약", bullets=["매출 4월 최고", "\t전월 대비 +67%", "비용은 안정적"])
pp.ppt_add_picture(P, 3, "output/sales_chart.png", left=500, top=110, width=400)

print("5) PPT 읽기")
print(json.dumps(pp.ppt_read(P), ensure_ascii=False, indent=1)[:3000])
print("\nOK: output/demo.xlsx, output/demo.pptx 를 열어 확인하세요.")
