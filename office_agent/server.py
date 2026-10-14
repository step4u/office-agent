"""MCP 서버 진입점: python -m office_agent.server  (stdio)
주의: stdout 은 MCP 프로토콜 전용이므로 print() 를 사용하지 말 것."""
from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from . import excel_tools as ex
from . import files_tools as ft
from . import ppt_tools as pp

mcp = FastMCP("office")

TOOLS = (
    ft.list_office_files, ft.copy_file, ft.convert_file,
    ex.excel_inspect, ex.excel_read_range, ex.excel_evaluate, ex.excel_create_workbook,
    ex.excel_write_range, ex.excel_format_range, ex.excel_add_chart, ex.excel_export_chart,
    pp.ppt_read, pp.ppt_list_layouts, pp.ppt_create, pp.ppt_add_slide, pp.ppt_add_text_box,
    pp.ppt_update_text, pp.ppt_replace_text, pp.ppt_add_table, pp.ppt_add_chart,
    pp.ppt_add_picture, pp.ppt_set_notes, pp.ppt_delete_slide,
)

for _fn in TOOLS:
    mcp.tool()(_fn)


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
