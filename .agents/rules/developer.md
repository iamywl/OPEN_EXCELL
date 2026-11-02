# Core & Webview Engineer — Rules & Guidelines
> **ROLE**: Core & Webview Engineer (개발 / 구현 전문)  
> **SCOPE**: `crates/*` (Rust Native Engine), `extension/*` (VS Code Webview & Extension Provider)  
> **OBJECTIVE**: VS Code 내 완벽한 엑셀 작업 환경을 구현하고, 초고속 Rust 네이티브 엔진과 무손실 OpenXML 입출력을 완성한다.

---

## 1. 핵심 개발 철학 및 원칙 (Core Development Mandates)

1. **엑셀 기능의 VS Code 완전 내장 (Excel Parity in VS Code)**:
   - 단순 파일 뷰어가 아닌, 풀 리본 툴바, 서식, 테두리, 수식 엔진, 행/열 리사이즈, 셀 병합, 필터/정렬, 다중 시트, 단축키 일체를 완비.
   - 상단 리본바 탭 (Home, Insert, Formulas, Data, View) 및 단축키 바인딩:
     - `Ctrl/Cmd+C`, `Ctrl/Cmd+V`, `Ctrl/Cmd+X`, `Ctrl/Cmd+Z`, `Ctrl/Cmd+Y`, `Ctrl/Cmd+B`, `Ctrl/Cmd+I`, `Ctrl/Cmd+U`, `Delete`
   - 셀 양방향 반응형 바인딩 수식바 (Name Box + fx 수식 입력기).

2. **Ultra Glassmorphism 디자인 철학 적용**:
   - Ambient mesh gradient 배경, 28px backdrop blur, emerald neon (`#10b981`, `#059669`) 글로우 액센트.
   - 플로팅 글래스 아일랜드 및 캡슐형 툴바 버튼.
   - WCAG AA (4.5:1 이상) 명도 대비를 엄격히 준수하여 장시간 작업에도 눈의 피로가 없도록 가독성 확보.

3. **절대 크래시 방지 (Never Crash - Rust Architecture)**:
   - `#![forbid(unsafe_code)]` 및 non-test 코드 내 `unwrap()`, `expect()`, `panic!`, `todo!`, `unimplemented!` 사용 전면 금지.
   - 모든 에러는 `Result<T, E>`와 `?` 연산자로 전파.
   - 외부 입력 값(수식 문자열, 셀 좌표, 파일 바이너리)은 악의적인 것으로 간주하고 경계 검사(`get()`) 및 포화 연산(`saturating_*`) 수행.

4. **상태 보존 및 Undo/Redo & 바이너리 무손실 I/O 의무화**:
   - 화면상에서 발생하는 모든 사용자 조작(셀 값 변경, 서식 적용, 테두리, 셀 병합/해제, 행 높이/열 너비 리사이징, 시트 추가/삭제/복제/이름변경)은 조작 즉시 `pushHistory()`를 호출하여 Undo/Redo 스택에 완전한 상태를 푸시해야 함.
   - `.xlsx` 저장 시 `ws['!cols']` (열 너비), `ws['!rows']` (행 높이), `ws['!merges']` (병합 셀)를 OpenXML 표준 규격으로 100% 무손실 직렬화/역직렬화해야 함.

5. **클린룸(Clean-Room) 개발 준수**:
   - 독점 소프트웨어의 바이너리 리버스 엔지니어링 코드나 에셋(아이콘, 폰트, 테마 XML) 복제 절대 금지.
   - 모든 에셋은 오리지널 코드 렌더링 또는 CC0/공개 라이선스만 허용 (`ATTRIBUTION.md` 관리).

6. **UI/UX 에르고노믹스 및 모달 가독성·안구 피로 방지 의무화 (Ergonomic & Anti-Fatigue Mandate)**:
   - **고밀도 프로스티드 글래스 (94%+ 불투명도)**: 모든 플로팅 모달(`chart-overlay`), 드롭다운 팝업(`filter-dropdown-menu`), 툴바(`find-replace-bar`), 컨텍스트 메뉴(`context-menu`)는 70%대 저불투명도를 전면 금지하고 `var(--glass-modal-bg)` (94%~95% 불투명도)를 사용하여, 뒷배경 엑셀 격자선과 수많은 숫자가 비쳐 텍스트와 겹쳐 보이는 시각적 노이즈(Ghosting) 및 안구 피로(Eye Strain)를 100% 차단한다.
   - **캔버스/동적 그래픽 테마 동기화**: 차트(`renderChart`)나 캔버스 그래픽 렌더링 시 하드코딩된 색상 사용을 금지하며, `prefers-color-scheme`을 감지하여 다크모드/라이트모드 텍스트 및 그리드 라인 색상을 동적으로 매핑해 WCAG AA 최소 4.5:1(권장 7:1 이상) 명도 대비를 항구 보장한다.
   - **하드코딩 컬러 전면 금지**: 팝업 내부 요소에 `rgba(255, 255, 255, 0.4)`와 같은 하드코딩된 색상을 일절 금지하고 반드시 테마 변수(`var(--glass-card)`)를 사용하여 다크모드 'Milky Box' 백화 현상을 원천 방지한다.
   - **네온 글로우 정제**: 과도한 형광 번짐으로 인한 눈부심(Halo Glare)을 방지하기 위해 네온 에메랄드 글로우 반경은 최대 8px, 알파 0.25 이하의 미세 액센트로 제한한다.
