# Office 데이터/문서 에이전트 지침

이 프로젝트는 `office` MCP 서버(Excel/PowerPoint COM 제어)로 xls/xlsx/ppt/pptx 를 읽고,
계산·차트·PPT 작성 후 파일로 저장한다. 항상 한국어로 답한다.

## 폴더 규칙
- `input/` : 원본 (읽기 전용으로 취급, 절대 직접 수정하지 않는다)
- `output/` : 결과물 저장 위치. 기존 파일을 고쳐야 하면 `copy_file` 로 output 에 복사한 뒤 복사본을 수정한다.
- 경로는 프로젝트 루트 기준 상대경로를 사용한다.

## 작업 순서
1. `list_office_files` 로 대상 파일 확인 → `excel_inspect` / `ppt_read` 로 구조 파악. 파악 없이 쓰기 금지.
2. 데이터가 크면 `excel_read_range` 의 `truncated` 를 확인하고, 합계/평균/조건집계는 직접 암산하지 말고
   `excel_evaluate`(수식 계산) 또는 결과 시트에 수식으로 기록해 Excel 이 계산하게 한다.
3. 계산 결과 시트는 가능하면 값 대신 수식(`=SUM(...)`)으로 작성해 추적 가능하게 한다.
4. 쓰기 후에는 `excel_read_range` / `ppt_read` 로 결과를 다시 읽어 검증한다(오류값 `#DIV/0!` 등 확인).
5. 완료 보고: 만든 파일 경로, 핵심 수치 2~3개, 가정/주의사항만 간단히.

## 차트 선택
- 시간 추이: `line` / `line_markers`, 항목 비교: `column` 또는 `bar`, 구성비(항목 6개 이하): `pie`
- 차트에는 제목, 축 제목(단위 포함)을 넣는다. 계열이 1개면 범례 생략(`legend=false`).
- PPT 에는 편집 가능한 `ppt_add_chart` 를 우선 사용한다. Excel 차트 이미지가 꼭 필요할 때만
  `excel_export_chart` → `ppt_add_picture`.

## PPT 작성 원칙
- 새 덱은 회사 템플릿이 있으면 `ppt_create(template_path=...)`, 없으면 기본 테마.
- 레이아웃 번호는 `ppt_list_layouts` 로 확인(기본: 1 제목, 2 제목+내용, 6 제목만, 7 빈 화면).
- 슬라이드 1장 = 메시지 1개. 제목은 결론형 문장, 불릿은 6개 이하, 한 줄 한 문장.
- 차트/표 슬라이드는 레이아웃 6(제목만)에 제목을 넣고 차트/표를 추가한다.
- 위치 단위는 pt. 슬라이드 크기는 `ppt_read`/`ppt_create` 결과의 slide_width/height 를 따른다.
- 발표자 설명이 필요하면 `ppt_set_notes`.

## 주의
- 숫자는 숫자형으로 넣는다(문자열 "1,000" 금지). 서식은 `excel_format_range` 의 number_format 사용.
- 같은 파일을 사용자가 Excel/PowerPoint 에서 열어 두면 저장이 실패할 수 있다. 오류 시 닫아 달라고 안내한다.
- 도구가 오류를 반환하면 원인을 읽고 인자를 고쳐 재시도하며, 2회 실패 시 사용자에게 상황을 보고한다.
- 이 서버는 지정된 작업 폴더 밖의 파일에 접근할 수 없다.
