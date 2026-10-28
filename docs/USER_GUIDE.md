# OPEN_EXCELL 사용자 가이드 (User Manual)

> **지원 플랫폼**: macOS (Apple Silicon / Intel), Windows 10/11 (x64 / ARM64), Linux, CLI, MCP  
> **지원 모드**: 
> 1. VS Code 익스텐션 (VS Code Extension Custom Editor)  
> 2. 네이티브 데스크톱 앱 (Rust Native Desktop App - macOS & Windows)  
> 3. CLI 무헤드 실행기 (`gridcraft-cli`)  
> 4. MCP 인공지능 에이전트 연동 (`gridcraft-cli mcp`)

---

## 1. VS Code 익스텐션으로 사용하기

OPEN_EXCELL은 Visual Studio Code 내부에서 `.xlsx`, `.xls`, `.csv`, `.tsv` 파일을 직접 열고 수정할 수 있는 커스텀 에디터를 제공합니다.

### 1.1 스프레드시트 파일 열기
- 탐색기(Explorer)에서 `.xlsx`, `.xls`, `.csv`, `.tsv` 파일을 **더블클릭**하면 자동으로 Open Excel 에디터로 열립니다.
- 다른 에디터와 충돌 시: 파일을 우클릭하고 **`연결 프로그램... (Open With...)` -> `Open Excel Editor`**를 선택하세요.

### 1.2 주요 편집 및 그리드 인터랙션
- **셀 편집**: 셀을 더블클릭하거나 타이핑하여 값을 입력하고, `Enter` 또는 `Tab`으로 다음 셀로 이동합니다.
- **자동 채우기 핸들 (Auto-Fill Drag Handle)**: 선택 영역 우하단의 작은 녹색 사각형 핸들을 마우스로 드래그하면 연속 숫자(`1, 2, 3...`) 또는 수식이 상대참조 오프셋과 함께 자동으로 채워집니다.
- **행 높이 / 열 너비 리사이징**: 컬럼 또는 행 헤더 사이의 경계선을 마우스로 드래그하여 원하는 폭과 높이로 자유롭게 조절할 수 있습니다.
- **셀 병합 (Merge & Center)**: 원하는 범위를 드래그한 후 리본바의 **Merge** 버튼을 누르면 하나로 결합되며, 다시 누르면 분할 해제됩니다. OpenXML `.xlsx` 파일 저장 시 `!merges`로 무손실 보존됩니다.

### 1.3 고급 수식 엔진 (50+ Enterprise Functions)
수식 입력줄(`fx`)에 타이핑하면 실시간으로 함수 목록과 매개변수 시그니처 힌트가 팝업됩니다.
- **데이터 조회/참조**:
  - `=VLOOKUP(102, A2:C10, 2, FALSE)` : 첫 열에서 키값을 찾아 대응 열 반환
  - `=XLOOKUP("Apple", A2:A10, C2:C10)` : 강력한 양방향 검색
  - `=INDEX(B2:B10, MATCH("Apple", A2:A10, 0))` : 유연한 행/열 교차 조회
- **조건부 집계 및 논리**:
  - `=SUMIF(C2:C10, ">1000")` : 조건에 맞는 행만 합산
  - `=COUNTIF(C2:C10, ">=2500")` : 조건에 맞는 셀 개수 집계
  - `=IFERROR(A1/B1, 0)` : 에러 발생 시 대체값 안전 반환
  - `=IFS(A1>=90, "A", A1>=80, "B", TRUE, "C")` : 다중 조건 판정
- **수학, 텍스트, 날짜**:
  - `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `ABS`, `MOD`, `POWER`, `SQRT`
  - `CONCAT`, `LEFT`, `RIGHT`, `MID`, `LEN`, `TRIM`, `UPPER`, `LOWER`
  - `TODAY()`, `NOW()`, `DATE`, `YEAR`, `MONTH`, `DAY`

### 1.4 데이터 필터링 및 정렬 (AutoFilter)
- 리본바의 **Filter** 버튼을 누르면 표 헤더 열에 `▾` 드롭다운 필터 아이콘이 생성됩니다.
- 필터 버튼 클릭 시 프로스티드 글래스 팝업이 나타나며, 열 내 고유값 체크박스를 검색 및 선택 해제하여 일치하는 행만 실시간으로 표시(나머지 행 숨김)할 수 있습니다.
- 팝업 내에서 '오름차순(A→Z)' 및 '내림차순(Z→A)' 정렬도 원클릭으로 수행할 수 있습니다.

### 1.5 전역 찾기 및 바꾸기 (Find & Replace)
- **단축키**: `Ctrl + F` (또는 `Cmd + F`)를 누르면 상단 우측에 플로팅 검색바가 나타납니다.
- **바꾸기**: `Ctrl + H` (또는 `Cmd + H`)를 누르면 바꾸기 모드로 전환됩니다.
- 실시간으로 일치 건수(`1 of 5`)가 카운트되며, Enter(다음), Shift+Enter(이전) 키로 일치 셀을 순차 탐색할 수 있습니다. Replace(단일) 또는 Replace All(전체) 교체가 가능합니다.

### 1.6 다중 시트 관리 (Multi-Sheet)
- 하단 시트 탭 바에서 `+` 버튼으로 새 워크시트를 추가할 수 있습니다.
- 탭을 **더블클릭**하면 인라인 입력창으로 전환되어 시트 이름을 즉시 변경할 수 있습니다.
- 탭을 **우클릭**하면 컨텍스트 메뉴가 나타나 '이름 바꾸기', '시트 복제(Duplicate)', '시트 삭제(Delete)'를 수행할 수 있습니다.

### 1.7 저장 (Save) 및 네이티브 전환
- **저장 단축키**: `Ctrl + S` (Windows) / `Cmd + S` (macOS)
- 상단 리본바의 **Native App** 버튼을 누르면 현재 보고 있는 파일이 독립 네이티브 앱 창(122 FPS Metal / DirectX)으로 즉시 전환되어 열립니다.

---

## 2. 네이티브 데스크톱 앱으로 사용하기 (Windows & macOS)

GridCraft 엔진 기반의 초고속 순수 네이티브 앱을 독립 실행할 수 있습니다.

### 2.1 실행 명령어

#### macOS 환경:
```bash
# 샘플 대시보드(Sales Dashboard)로 네이티브 앱 실행
cargo run --release -p gridcraft -- --sample sales

# 예산 샘플(Budget)로 실행
cargo run --release -p gridcraft -- --sample budget

# 로컬 엑셀 파일 직접 열기
cargo run --release -p gridcraft -- /path/to/your/file.xlsx
```

#### Windows 환경 (PowerShell / CMD):
```powershell
# 샘플 대시보드로 네이티브 앱 실행
cargo run --release -p gridcraft -- --sample sales

# 로컬 엑셀 파일 열기
cargo run --release -p gridcraft -- C:\path\to\your\file.xlsx
```

### 2.2 독립 바이너리로 실행하기
- **macOS**: `./target/release/gridcraft [파일명]`
- **Windows**: `.\target\release\gridcraft.exe [파일명]`

---

## 3. CLI 무헤드 실행기 (`gridcraft-cli`)

터미널이나 CI/CD 스크립트에서 GUI 창 없이 엑셀 파일을 일괄 처리할 수 있습니다.

```bash
# 1. 수식 단독 평가
gridcraft-cli eval '=SUM(10, 20, 30)'
# 출력: 60

# 2. 엑셀 파일 로드 및 배치 명령 적용 후 저장
gridcraft-cli run --in input.xlsx --cmd 'cell.set={"cell":"A1","input":"Title"}' --out output.xlsx

# 3. JSON-RPC 원격 제어 포트 바인딩 실행
gridcraft-cli send 7979 engine.execute '{"command":"cell.set","params":{"cell":"B2","input":"=SUM(A1:A9)"}}'
```

---

## 4. MCP (Model Context Protocol) AI 에이전트 연동

Claude Code, Cursor, Windsurf, Antigravity 등 최신 AI 에이전트와 연동하여 자연어로 엑셀 스프레드시트를 생성하고 분석할 수 있습니다.

```bash
# MCP 서버 실행
gridcraft-cli mcp
```
제공 툴 규약: `read_cell`, `set_cell`, `get_range`, `set_range`, `eval_formula`, `format_cells`, `create_chart`, `list_sheets`, `save_workbook`

---

## 5. QA 및 화면 캡처 검증 (Proof of QA)
- 네이티브 렌더러 오프스크린 스냅샷 생성:
```bash
cargo run -p gridcraft-ui-egui --example snapshot -- --sample sales docs/images/qa_proof_native.png
```
생성된 실제 2880x1800 고해상도 UI 캡처는 `docs/images/qa_proof_native.png` 및 `docs/QA_PROOF.md`에서 확인하실 수 있습니다.
