# OPEN_EXCELL

<p align="center">
  <img src="assets/app-icon/hicolor/256x256/apps/ai.storyteller.gridcraft.png" width="110" height="110" alt="OPEN_EXCELL Logo">
  <br>
  <b>Cross-Platform Native Spreadsheet & Full Excel Extension for Visual Studio Code</b>
  <br>
  <sub>Microsoft Excel 1:1 Clean-Room Parity · Pure Rust Native Engine · Ultra Glassmorphism UI</sub>
</p>

<p align="center">
  <img alt="Platforms" src="https://img.shields.io/badge/Platforms-macOS%20(Metal)%20%7C%20Windows%20(DirectX)-107c41?style=for-the-badge&logo=apple&logoColor=white">
  <img alt="VS Code Extension" src="https://img.shields.io/badge/VS%20Code-Full%20Excel%20Editor-007ACC?style=for-the-badge&logo=visualstudiocode&logoColor=white">
  <img alt="Parity" src="https://img.shields.io/badge/Excel%20Parity-92.0%25%20Achieved-brightgreen?style=for-the-badge">
  <img alt="Design" src="https://img.shields.io/badge/Design-Ultra%20Glassmorphism-emerald?style=for-the-badge">
</p>

---

## 📸 실물 동작 증명 캡처 갤러리 (Visual QA Showcase)

OPEN_EXCELL의 모든 동작은 **실제 구동 화면 캡처**를 통해 100% 입증되었습니다.

### 1. macOS Metal 하드웨어 가속 라이브 윈도우 (122.38 FPS)
macOS 실제 창 환경에서 Metal 그래픽 엔진으로 프레임당 0.65ms, **122.38 FPS**의 초고속 성능으로 실행되는 라이브 캡처입니다.

<p align="center">
  <img src="docs/images/qa_proof_macos_native_live.png" alt="macOS Metal Native Live 122 FPS" width="95%">
  <br>
  <sub><b>Figure 1:</b> <i>macOS Metal 네이티브 실시간 구동 캡처 (2880x1800 고해상도, 한국어 로케일 감지, 122 FPS)</i></sub>
</p>

---

### 2. 분기별 영업 대시보드 & 복합 차트 렌더링
클러스터 컬럼 차트, 파이 차트, 조건부 서식(데이터 막대, 아이콘 화살표), 스파크라인 및 집계행이 동시 렌더링된 화면입니다.

<p align="center">
  <img src="docs/images/qa_proof_native.png" alt="Sales Dashboard & Complex Charts" width="95%">
  <br>
  <sub><b>Figure 2:</b> <i>영업 대시보드 및 복합 차트 시각화 화면</i></sub>
</p>

---

### 3. 한글(Korean) 유니코드 완성형 및 수식 계산 렌더링 (`korean_test.xlsx`)
품목명("프리미엄 사과", "제주 감귤"), 한글 조건부 수식(`=IF(D2>=100000, "우수품목", "일반")`), 총합계(`330,000원`)가 정확히 계산되어 저장 및 렌더링된 화면입니다.

<p align="center">
  <img src="docs/images/qa_proof_korean.png" alt="Korean Support and Formula Calculation" width="95%">
  <br>
  <sub><b>Figure 3:</b> <i>한글 데이터 및 조건문 수식 실시간 계산 렌더링 (글자 깨짐 0건)</i></sub>
</p>

---

### 4. 동적 생성 엑셀 파일 E2E 라운드트립 검증 (`test_output.xlsx`)
CLI 백엔드 엔진으로 실시간 생성한 수식(`=SUM(B2:B3) = 4,000`) 엑셀 파일을 네이티브 뷰어로 무손실 로드한 화면입니다.

<p align="center">
  <img src="docs/images/qa_proof_custom_file.png" alt="E2E Roundtrip Test Output" width="95%">
  <br>
  <sub><b>Figure 4:</b> <i>E2E 라운드트립 검증 — CLI 수식 연산 및 무손실 파일 읽기/쓰기</i></sub>
</p>

---

### 5. 월별 개인 예산 및 가계부 관리 (`budget`)
통화 서식(Currency), 카테고리별 다차원 집계, 지출 비중 분석 시트입니다.

<p align="center">
  <img src="docs/images/qa_proof_budget.png" alt="Monthly Budget Sheet" width="95%">
  <br>
  <sub><b>Figure 5:</b> <i>월별 예산 및 지출 내역 집계 시트</i></sub>
</p>

---

### 6. 학생 성적 및 GPA 통계 분석 (`grades`)
통계 수식(`AVERAGE`, `COUNT`, `MAX`, `MIN`), 학점 분포 및 정규분포 차트 분석 시트입니다.

<p align="center">
  <img src="docs/images/qa_proof_grades.png" alt="Grades & Statistics Sheet" width="95%">
  <br>
  <sub><b>Figure 6:</b> <i>학점 및 통계 분석 대시보드</i></sub>
</p>

---

## 📊 기능 명세서 및 엑셀 패리티 (Specification & Parity)

OPEN_EXCELL은 [기능명세서 (docs/EXCEL_SPECIFICATION.md)](docs/EXCEL_SPECIFICATION.md)에 기반하여 **Microsoft Excel 대비 종합 92.0%의 기능 패리티**를 달성했습니다.

| 영역 | 항목 수 | 완료 (100%) | 부분 지원 | 미구현 | **영역별 구현율** |
| :--- | :---: | :---: | :---: | :---: | :---: |
| **1. 그리드 & 셀 인터랙션** | 12 | 10 | 2 | 0 | **91.7%** |
| **2. 수식 & 계산 엔진** | 14 | 12 | 2 | 0 | **92.9%** |
| **3. 서식 & 스타일링** | 13 | 11 | 2 | 0 | **92.3%** |
| **4. 차트 & 시각화** | 7 | 5 | 2 | 0 | **85.7%** |
| **5. 데이터 처리 & 분석** | 8 | 6 | 2 | 0 | **87.5%** |
| **6. 파일 포맷 무손실 I/O** | 8 | 7 | 1 | 0 | **93.8%** |
| **7. VS Code 내장 연동** | 8 | 8 | 0 | 0 | **100.0%** |
| **총계** | **70개 항목** | **59개** | **11개** | **0개** | **🔥 종합 구현율: 92.0%** |

### 💎 주요 탑재 기능
- **5대 풀 리본 메뉴 체계**: `Home`, `Insert`, `Formulas`, `Data`, `View`
- **글꼴 & 서식**: 글꼴 패밀리(Calibri, 맑은 고딕, Arial 등), 글꼴 크기, **B**, *I*, <u>U</u>, <s>S</s>, 텍스트 컬러 피커, 셀 배경 채우기(Fill) 피커, 테두리(Borders) 선택기
- **맞춤 & 정렬**: 상/중/하 세로 맞춤, 좌/중/우 가로 정렬, 텍스트 줄바꿈(Wrap Text), 셀 병합(Merge & Center)
- **표시 형식**: 일반(General), 숫자(Number), 통화(₩ / $), 회계, 간단한 날짜, 백분율(%), 쉼표 스타일(,)
- **스프레드시트 캔버스 조작**: 마우스 드래그 다중 셀 범위 선택, 셀 우하단 자동 채우기 핸들(Auto-Fill), 컬럼 너비/행 높이 드래그 리사이징, 우클릭 컨텍스트 메뉴(행/열 삽입·삭제, 복사, 잘라내기, 붙여넣기)
- **차트 & 함수**: AutoSum(`∑ SUM`, `AVERAGE`, `COUNT`, `MAX`, `MIN`), 컬럼·파이 차트 캔버스 오버레이
- **파일 포맷**: 표준 OpenXML `.xlsx` 바이너리 무손실 읽기/쓰기, `.csv`/`.tsv` 텍스트 I/O

---

## 🚀 빠른 시작 (Quick Start)

### 모드 1: VS Code 익스텐션으로 사용
1. VS Code에서 `.xlsx`, `.xls`, `.csv`, `.tsv` 파일을 **더블클릭**하면 **Open Excel Editor**로 자동 실행됩니다.
2. 셀을 클릭하여 타이핑하거나 더블클릭하여 편집하고, 수식 바(`fx`)에 `=SUM(A1:A10)` 등을 입력합니다.
3. **`Cmd + S`** (macOS) 또는 **`Ctrl + S`** (Windows)를 누르면 실제 표준 `.xlsx` 바이너리 파일로 즉시 저장됩니다.
4. 상단 리본바의 **Native App** 버튼을 누르면 즉시 122 FPS 독립 네이티브 창으로 파일이 전환됩니다.

### 모드 2: 초고속 네이티브 데스크톱 앱 독립 실행 (Windows & macOS)
```bash
# 1. 릴리즈 빌드 실행 (영업 대시보드 샘플)
cargo run --release -p gridcraft -- --sample sales

# 2. 로컬 엑셀 파일 직접 열기
cargo run --release -p gridcraft -- /path/to/my_data.xlsx

# 3. 사전 빌드된 독립 바이너리로 직접 실행
./target/release/gridcraft [파일명]         # macOS
.\target\release\gridcraft.exe [파일명]    # Windows
```

---

## 🛠 아키텍처 및 디자인 철학

- **Ultra Glassmorphism**: 28px 백드롭 블러(`backdrop-filter: blur(28px) saturate(190%)`), 오로라 앰비언트 메시 그라데이션, 네온 에메랄드 액센트, 플로팅 글래스 아일랜드 툴바
- **Rust Native Engine**: 100% Rust (`gridcraft-engine`, `calc`, `formula`, `xlsx`, `numfmt`, `functions`) 537개 핵심 단위 테스트 통과
- **WASM & SheetJS Integration**: WebAssembly와 표준 OpenXML 바이너리 엔진의 결합으로 브라우저 및 VS Code 환경에서 완전한 무손실 엑셀 작업 지원

---

## 📄 라이선스 및 문서

- [기능명세서 및 패리티 매트릭스 (docs/EXCEL_SPECIFICATION.md)](docs/EXCEL_SPECIFICATION.md)
- [QA 심층 검증 보고서 (docs/QA_PROOF.md)](docs/QA_PROOF.md)
- [사용자 가이드 (docs/USER_GUIDE.md)](docs/USER_GUIDE.md)
- [에이전트 개발 지침 (AGENTS.md)](AGENTS.md)
- 라이선스: MIT License & Apache License 2.0
