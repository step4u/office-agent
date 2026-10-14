# Office Agent (Claude Code + Office COM MCP)

사내 LLM이 연결된 Claude Code 가 xls/xlsx/ppt/pptx 를 읽고, 계산·차트·PPT 작성 후 파일로 저장하도록 하는 에이전트.

```
사용자 지시 ─► Claude Code (사내 LLM) ─MCP(stdio)─► office_agent.server (Python/pywin32)
                                                      └─ COM ─► Excel / PowerPoint (설치된 Office)
```

- 대상 OS: Windows 10/11 Pro, 설치된 Microsoft Office(2013 이상, `AddChart2` 필요), Python 3.10+
- LLM 연결 설정(사내 LLM 엔드포인트/토큰)은 기존 Claude Code 설정을 그대로 사용한다.
- Excel 은 **독립 프로세스**로 실행되어 사용자가 열어 둔 Excel 과 섞이지 않는다. PowerPoint 는 단일 인스턴스라
  사용자가 열어 둔 파일이 없을 때만 종료한다.

## 설치
```bat
setup.bat
.venv\Scripts\python tests\smoke_test.py
claude
```
`setup.bat` 은 venv 생성, 패키지 설치, `input/`·`output/` 생성, 프로젝트 범위 MCP 서버 `office` 등록을 수행한다.
사내망이면 pip 사내 미러 설정이 필요할 수 있다. Claude Code 안에서 `/mcp` 로 `office` 가 connected 인지 확인.

## 사용 예
`input/` 에 파일을 넣고 프로젝트 폴더에서 `claude` 실행 후:

- `input/sales.xlsx 를 분석해서 지역별 월 매출 합계 시트를 만들고 라인 차트를 추가해 output/sales_report.xlsx 로 저장해줘`
- `input/sales.xlsx 의 요약을 10장 이내 PPT 로 만들어 output/summary.pptx 로 저장. 핵심 차트 포함`
- `input/old.ppt 를 pptx 로 변환하고 3번 슬라이드의 수치를 input/data.xlsx 의 최신 값으로 갱신해줘`

배치 실행: `run_agent.bat "지시문"` (`claude -p` 사용).

## 제공 도구
| 구분 | 도구 |
|---|---|
| 파일 | `list_office_files`, `copy_file`, `convert_file`(xls↔xlsx, ppt↔pptx, PDF, csv) |
| Excel | `excel_inspect`, `excel_read_range`, `excel_evaluate`, `excel_create_workbook`, `excel_write_range`, `excel_format_range`, `excel_add_chart`, `excel_export_chart` |
| PPT | `ppt_read`, `ppt_list_layouts`, `ppt_create`(템플릿 지원), `ppt_add_slide`, `ppt_add_text_box`, `ppt_update_text`, `ppt_replace_text`, `ppt_add_table`, `ppt_add_chart`, `ppt_add_picture`, `ppt_set_notes`, `ppt_delete_slide` |

## 설계 원칙 / 보안
- 호출마다 파일을 열고 → 작업 → 저장 → 닫기(상태 없음). 중간에 오류가 나도 Office 프로세스가 남지 않는다.
- 작업 폴더 밖 경로 차단. `input/` 은 읽기 전용(`OFFICE_AGENT_READONLY_DIRS`), 수정은 `output/` 복사본에서만.
- 임의 VBA/명령 실행 도구는 의도적으로 제공하지 않는다.

## 환경변수
| 변수 | 기본 | 설명 |
|---|---|---|
| `OFFICE_AGENT_WORKDIR` | 현재 폴더 | 허용 작업 루트 |
| `OFFICE_AGENT_ALLOWED_DIRS` | - | 추가 허용 폴더(`;` 구분) |
| `OFFICE_AGENT_READONLY_DIRS` | `input` | 쓰기 금지 폴더 |
| `OFFICE_AGENT_MAX_CELLS` | 5000 | 한 번에 읽는 최대 셀 수 |
| `OFFICE_AGENT_PPT_WINDOW` | 0 | 1이면 PowerPoint 창을 띄워 작업(차트 추가 오류 시 사용) |

## 문제 해결
- **파일이 잠김/저장 실패**: 해당 파일을 Excel/PowerPoint 에서 닫는다.
- **PPT 차트 추가 오류**: `OFFICE_AGENT_PPT_WINDOW=1` 로 설정(`claude mcp add` 의 `--env` 에 추가하거나 `.mcp.json` 수정).
- **좀비 프로세스**: 작업 관리자에서 `EXCEL.EXE`/`POWERPNT.EXE` 종료. 반복되면 이슈로 기록.
- **`office` 가 connected 아님**: `.venv\Scripts\python -m office_agent.server` 를 직접 실행해 import 오류 확인.
- 사용자가 Office 를 한 번도 실행하지 않았거나 라이선스 팝업이 뜨는 PC에서는 먼저 수동으로 1회 실행해 활성화한다.
- 서버 가동 중 Office 를 서비스 계정/세션 0 에서 돌리는 환경(작업 스케줄러 등)은 COM 이 불안정하므로 로그인한 사용자 세션에서 실행한다.
