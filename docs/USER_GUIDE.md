# OPEN_EXCELL 사용자 가이드 (User Manual)

> **지원 플랫폼**: macOS (Apple Silicon / Intel), Windows 10/11 (x64 / ARM64), Linux  
> **실행 모드**: 
> 1. VS Code 익스텐션 (VS Code Extension Custom Editor)  
> 2. 네이티브 데스크톱 앱 (Rust Native Desktop App)

---

## 1. VS Code 익스텐션으로 사용하기

OPEN_EXCELL은 Visual Studio Code 내부에서 `.xlsx`, `.xls`, `.csv`, `.tsv` 파일을 직접 열고 수정할 수 있는 커스텀 에디터를 제공합니다.

### 1.1 스프레드시트 파일 열기
- 탐색기(Explorer)에서 `.xlsx`, `.xls`, `.csv`, `.tsv` 파일을 **더블클릭**하면 자동으로 Open Excel 에디터로 열립니다.
- 다른 에디터와 충돌 시: 파일을 우클릭하고 **`연결 프로그램... (Open With...)` -> `Open Excel Editor`**를 선택하세요.

### 1.2 주요 편집 기능
- **셀 편집**: 셀을 더블클릭하거나 타이핑하여 값을 입력하고, `Enter` 또는 `Tab`으로 다음 셀로 이동합니다.
- **수식 계산 (Formula Engine)**:
  - `=SUM(A1:A10)` : 지정 범위의 합계 자동 계산
  - `=AVERAGE(B1:B20)` : 지정 범위의 평균 자동 계산
  - `=COUNT(C1:C10)` : 지정 범위의 숫자 개수 계산
  - `=A1` : 다른 셀의 값을 참조
- **서식 및 툴바**:
  - `B` (Bold): 글씨 굵게
  - `I` (Italic): 글씨 기울임
  - 정렬 (Align): 좌측, 가운데, 우측 정렬
  - `+ Row`, `+ Col`: 선택한 위치에 행/열 삽입
  - `A→Z`, `Z→A`: 지정된 열 기준 오름차순/내림차순 정렬
  - `Export CSV`: 현재 시트를 CSV 파일로 즉시 다운로드
- **저장 (Save)**:
  - 단축키: `Ctrl + S` (Windows) / `Cmd + S` (macOS)
  - 또는 상단 리본바의 **Save** 버튼 클릭 시 파일 시스템에 즉시 동기화 저장됩니다.
- **시트 관리**: 하단 탭의 `+` 버튼으로 시트를 추가하거나 전환할 수 있습니다.
- **원클릭 네이티브 전환**: 상단 리본바의 **Native App** 버튼을 누르면 현재 보고 있는 파일이 독립 네이티브 앱 창으로 즉시 전환되어 열립니다.

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
릴리즈 빌드 완료 후 생성된 바이너리를 더블클릭하거나 PATH에 등록하여 즉시 실행할 수 있습니다:
- **macOS**: `./target/release/gridcraft [파일명]`
- **Windows**: `.\target\release\gridcraft.exe [파일명]`

---

## 3. QA 및 화면 캡처 검증 (Proof of QA)
- 네이티브 렌더러 오프스크린 스냅샷 생성:
```bash
cargo run -p gridcraft-ui-egui --example snapshot -- --sample sales docs/images/qa_proof_native.png
```
생성된 실제 2880x1800 고해상도 UI 캡처는 `docs/images/qa_proof_native.png` 및 `docs/QA_PROOF.md`에서 확인하실 수 있습니다.
