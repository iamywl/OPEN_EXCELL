# OPEN_EXCELL / GridCraft — Instructions for Agents & Loop Engineering

> **PROJECT**: OPEN_EXCELL (Cross-Platform Native Spreadsheet & VS Code Extension)  
> **REMOTE REPO**: https://github.com/iamywl/OPEN_EXCELL.git  
> **PLATFORMS**: macOS (Apple Silicon & Intel), Windows 10/11 (x64, ARM64), Linux  
> **DESIGN PHILOSOPHY**: **Ultra Glassmorphism / Frosted Liquid Glass** (Ambient mesh gradient, 28px backdrop blur, emerald neon accents, floating glass islands & pills)  

---

## ⭐️ [핵심 기본 철학] 엑셀 기능의 VS Code 완전 내장 (Excel Parity in VS Code)
**"단순한 파일 뷰어나 간이 테이블이 아니라, Microsoft Excel의 핵심 기능 일체(풀 리본 툴바, 셀 포맷, 테두리, 수식 엔진, 행/열 리사이즈, 셀 병합, 필터/정렬, 다중 시트, 단축키, 바이너리 저장 등)를 VS Code 익스텐션 내에 온전히 구현하여 완벽한 엑셀 작업 환경을 제공하는 것이 본 프로젝트의 최우선 불변 철학이다."**

1. **완전한 엑셀 사용자 경험**:
   - 상단 리본바 탭 체계 (Home, Insert, Formulas, Data, View) 완비
   - 폰트(글꼴 패밀리, 크기), 글꼴색/배경색(컬러 피커), 테두리(Borders), 정렬(수직/수평), 텍스트 줄바꿈(Wrap Text), 셀 병합(Merge & Center)
   - 숫자 서식 드롭다운 (일반, 숫자, 통화, 회계, 백분율, 날짜, 지수)
   - 엑셀 표준 수식 바 (Name Box + fx 수식 입력 + 셀 양방향 반응형 바인딩)
   - 스프레드시트 그리드 인터랙션:
     - 마우스 드래그 다중 셀 범위 선택 (Range Selection: `A1:C5`)
     - 셀 우하단 자동 채우기 핸들 (Auto Fill Drag Handle)
     - 행 높이 / 열 너비 마우스 드래그 리사이징
     - 컨텍스트 메뉴 (행/열 삽입, 삭제, 복사, 잘라내기, 붙여넣기)
     - 엑셀 표준 단축키: `Ctrl/Cmd+C`, `Ctrl/Cmd+V`, `Ctrl/Cmd+X`, `Ctrl/Cmd+Z`, `Ctrl/Cmd+Y`, `Ctrl/Cmd+B`, `Ctrl/Cmd+I`, `Ctrl/Cmd+U`, `Delete`
   - 표준 바이너리 `.xlsx` (OpenXML) 무손실 읽기 및 쓰기 저장

2. **크로스 플랫폼 네이티브 연동**:
   - macOS(Metal) 및 Windows(DirectX) 초고속 네이티브 데스크톱 앱과의 원클릭 양방향 연동

---

## 📋 [명세서 기반 정량적 개발 원칙] (Craft Specification & Parity Tracking)
**"모든 개발은 `docs/EXCEL_SPECIFICATION.md`의 기능명세서와 패리티 매트릭스(Parity Matrix)를 기준으로 진행하며, 카테고리별 구현율(%)을 수치화하여 객관적으로 측정하고 갭(Gap)을 체계적으로 메워 나간다."**
1. **기능명세서 기반 관리**: Excel 7대 핵심 영역(그리드 인터랙션, 수식/함수 엔진, 서식/스타일, 차트/시각화, 데이터 처리, 파일 무손실 I/O, VSC 내장 연동)별 1:1 세부 명세 수립.
2. **정량적 구현율 산출**: 각 항목별 [완료(100%) / 부분 지원(50%) / 미구현(0%)]을 엄격하게 채점하여 종합 달성률 및 카테고리별 달성률을 지속적으로 공시.
3. **P0 우선순위 집중 개발**: 사용자가 일상적인 업무에서 필수적으로 사용하는 엑셀 핵심 기능(서식, 드래그 채우기, 수식, 행/열 리사이즈, 바이너리 저장 등)부터 결손 없이 100% 완성.

---

## ⚖️ [허위 기재 방지 및 미검증 전수 발굴 원칙] (Zero Untested Claims & Defect Rectification Mandate)
1. **[명세서 무결성 & 허위/과장 기재 절대 금지]**:
   - 명세서의 상태를 '완료(100%)'로 표시하려면, 해당 기능이 실제 소스 코드(Webview 및 Native)에 완전히 작동하는 상태로 구현되어 있어야 하며, 1줄짜리 placeholder나 미동작 더미 함수는 절대 '완료'로 기재할 수 없다.
   - 명세서에 기재된 모든 항목은 테스트 또는 QA 시나리오와 1:1로 매핑되어 실질적으로 검증되어야 한다. 검증되지 않은 기능은 '완료'로 표시할 수 없고 '미검증 완료(Unverified)' 또는 '부분 지원(50%)'으로 분류하여 즉시 QA 시나리오를 추가해야 한다.
2. **[기능 및 UI/UX 실질 구현율 감사]**:
   - UI 컨트롤(버튼, 드롭다운, 인풋, 모달 등)은 겉모양만 있어서는 안 되며, 데이터 모델 변경, 2-Pass 재계산, 실행취소(Undo) 스택 저장, 바이너리 저장 연동까지 100% 완전하게 동작해야 한다.
   - 글래스모피즘(Ultra Glassmorphism) 테마는 가독성, 명도 대비(WCAG AA 4.5:1), 백드롭 블러(28px) 및 네온 에메랄드 액센트가 전 컴포넌트에 일관되게 적용되어야 한다.
3. **[QA 미포함 기능 전수 발굴 및 100% 라이브 검증]**:
   - 명세서에 있는 기능 중 기존 QA 스위트에서 누락되었거나 직접 테스트되지 않은 기능(셀 서식 변경, 글꼴 크기/색상, 테두리 적용, 정렬, 셀 병합/해제, 숫자 서식 드롭다운, 컬럼/행 리사이징, 다중 시트 복제/삭제 등)을 전수 발굴하여 테스트 시나리오로 편입하고 자동화 E2E 스크립트 및 실물 캡처로 100% 검증한다.
4. **[로컬 시스템 정식 앱 번들링 및 Launchpad/Spotlight 통합 필수 원칙]**:
   - **단순 터미널 바이너리 배포 절대 금지**: `cargo build` 후 `~/.cargo/bin`에 바이너리 파일만 복사하고 설치를 마치는 것은 엄격히 금지된다. 일반 사용자는 Finder, Launchpad, Spotlight를 통해 GUI 앱을 탐색하고 실행하므로, 반드시 정식 macOS Application Bundle (`/Applications/OPEN_EXCELL.app`, `/Applications/GridCraft.app`)로 패키징해야 한다.
   - **필수 번들 규격 구축**:
     - `Contents/MacOS/<Executable>` (실행 권한 `chmod +x` 부여된 네이티브 바이너리)
     - `Contents/Info.plist` (앱 이름, 카테고리, 문서 확장자 `.xlsx`, `.csv` 연결 정의)
     - `Contents/PkgInfo` (`APPL????` 매직 시그니처)
     - `Contents/Resources/AppIcon.icns` (512x512 고해상도 아이콘 번들)
   - **고유 CFBundleIdentifier 준수**: 복수의 앱 배포 시 번들 식별자(Bundle ID)는 절대 중복되어서는 안 된다 (예: `ai.storyteller.openexcell`, `ai.storyteller.gridcraft`). 번들 ID가 중복되면 macOS Launchpad SQLite DB에서 충돌이 발생하여 앱이 비활성 고아 그룹(orphan group)으로 숨겨져 런치패드에서 보이지 않는 치명적 결함이 발생한다.
   - **시스템 등록 및 캐시 갱신 파이프라인 자동화**:
     1) `codesign --force --deep --sign - /Applications/<AppName>.app` (로컬 ad-hoc 코드 서명)
     2) `/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister -f /Applications/<AppName>.app` (LaunchServices 강제 재등록)
     3) `mdimport /Applications/<AppName>.app` (Spotlight 메타데이터 인덱싱)
     4) `killall Dock` (Launchpad SQLite 캐시 및 독 프로세스 리로드)
   - **사용자 위치 및 실행 안내 필수화**: macOS Launchpad 첫 페이지가 35개 기본 앱으로 꽉 차 있는 경우 신규 앱은 2번째 페이지(또는 마지막 페이지)에 등록된다. 사용자에게 단순 "설치 완료"만 전달하지 말고, 1) 런치패드 2번째 페이지 위치, 2) 상단 검색창 검색법, 3) Spotlight (`Cmd + Space`), 4) Finder 응용 프로그램 폴더, 5) 터미널 `open /Applications/<AppName>.app` 실행법을 종합 안내한다.

5. **[Undo/Redo 스택 & 바이너리 무손실 I/O 누락 방지]**:
   - 화면상에서 동작하는 모든 UI 인터랙션(셀 서식, 글꼴/배경색, 테두리, 셀 병합/해제, 행 높이/열 너비 리사이징, 다중 시트 생성/삭제/복제/이름변경)은 조작 즉시 `pushHistory()`를 호출하여 Undo/Redo 스택에 완전한 상태가 저장되어야 한다.
   - 열 너비(`!cols`) 및 행 높이(`!rows`), 병합 범위(`!merges`)는 바이너리 저장 시 누락 없이 OpenXML 표준 규격으로 직렬화되어야 하며, 재오픈 시 1px의 오차도 없이 복원되어야 한다.

6. **[백그라운드 프로세스 및 좀비 태스크 정리 원칙]**:
   - GUI 애플리케이션 및 백그라운드 커맨드는 실행 후 방치되어 백그라운드 태스크나 포트를 점유해서는 안 된다. 테스트 및 스크린샷 캡처 후에는 즉시 정상 종료시키거나 `manage_task(Action='kill')`로 정리하여 클린 상태를 유지한다.

7. **[전수 문서화 및 증명 스냅샷 실시간 동기화 원칙]**:
   - `README.md`, `docs/EXCEL_SPECIFICATION.md`, `docs/QA_PROOF.md`, `docs/AUDIT_AND_IMPROVEMENT_REPORT.md` 4대 문서의 패리티 수치, 감사 내용, 라이브 증명 스크린샷 매핑(Fig. 1 ~ Fig. 27)은 실시간으로 100% 일치해야 하며, 캡처 파일은 반드시 디스크에 실존하는 비어있지 않은 PNG 파일이어야 한다.

---

## Loop Engineering Protocol (루프 엔지니어링 원칙)
모든 개발 및 배포는 다음의 5단계 검증 루프를 통과해야 하며 100% 완수를 목표로 함:
1. **[Requirement & Plan Loop]**: 사용자의 요구사항(기본철학: 엑셀 기능을 VSC에 완전 내장, 윈도우/맥 호환, 글래스모피즘 디자인, QA 캡처, 깃허브 연동)을 분석하고 반영.
2. **[Implementation Loop]**:
   - 엑셀 완전 기능(Full Excel Feature Parity) VS Code Extension 내장 구현 (`extension/web/grid.js`, `grid.css`, `index.html`, `editorProvider.ts`).
   - 글래스모피즘 디자인 철학 적용 (Frosted Liquid Glass, Neon Emerald Glow, Capsule Inputs).
   - 네이티브 데스크톱 빌드 및 실행 체계 완비 (macOS Metal 122 FPS 검증 완료, Windows 호환).
   - 크로스 플랫폼 파일 입출력 및 인코딩/경로 호환성 보장.
3. **[Testing & Build Loop]**:
   - Rust crates 585개 핵심 단위/통합 테스트 100% 통과.
   - VS Code Extension 빌드 및 VSIX 패키징 (`open-excell-1.0.0.vsix`) 완료 및 VS Code 실제 설치 검증 (`code --install-extension open-excell-1.0.0.vsix --force`).
   - 사용자 macOS 로컬 시스템에 정식 네이티브 `.app` 번들 생성 및 `/Applications` 배포 (`/Applications/OPEN_EXCELL.app`, `/Applications/GridCraft.app`), `codesign`, `lsregister`, `killall Dock`을 통한 Launchpad/Spotlight 등록 및 실행 검증.
   - 백그라운드 테스트 및 데몬 프로세스 종료 및 좀비 태스크 정리 확인.
4. **[Proof of QA & Capture Loop]**:
   - 모든 구현 기능에 대해 결손 없이 라이브 스크린샷 캡처 및 E2E 테스트 통과 증명.
   - 오프스크린/스냅샷 렌더링 스냅샷 증명.
   - 사용자 사용 가이드 및 증명 문서 작성.
5. **[Sync & Deliver Loop]**:
   - `git add`, `git commit`, `git push origin main` 원격 저장소 동기화.

---

# GridCraft — Base Architecture & Clean-room Rules

GridCraft is a clean-room, open-source, Rust-native spreadsheet targeting Microsoft Excel parity — and superiority (speed, openness, agent control). It runs natively on macOS, Windows, Linux and BSD, and on the web via WASM. Siblings with the same conventions: `../designcraft` (InDesign-class, the closest reference), `../photocraft` (Photoshop), `../vectorcraft` (Illustrator), `../pdfcraft` (Acrobat), `../filmcraft` (Premiere), `../lightcraft` (Lightroom), `../effectcraft` (After Effects).

Standards and learnings shared across the crafting apps live in `../../craftrules` (checked out next to the craft apps, or `storytold/craftrules`). Read its `AGENTS.md` at the start of a session, follow its standards, and contribute reusable learnings back there. Never code: repos don't share code.

## Start every session here
1. Read `plan/STATUS.md` (current milestone, next task), then the task in `plan/execution-plan.md` and the relevant `plan/architecture.md` section. Behaviour reference: `plan/excel/*.md` (`observed-ui.md` holds measured observations of the running app).
2. Follow the autonomous operation protocol (`plan/execution-plan.md` §7). Don't stop to ask unless §7 lists the decision as the user's.

`plan/` is gitignored (local only).

## Assets and iconography — absolute rule
**No Microsoft, Adobe, Avid, Autodesk (or any other vendor's) iconography, images, artwork, fonts, templates, themes or sample files — ever.** This is the most important rule in this repository; breaking it is a serious failure.
- Every asset in the repo is original work by GridCraft contributors, or public domain / CC0, or open source / Creative Commons that allows redistribution (or licensed open source by the contributor who created it).
- Every asset file has a row in [`ATTRIBUTION.md`](ATTRIBUTION.md) (author, source, licence). `cargo xtask assets` enforces it. Adding an asset without a row is a bug.
- Prefer art generated in code: the UI icon set (`crates/ui-egui/src/icons.rs`), table and cell style palettes, themes and sample workbooks are drawn/defined in code and original.
- Never copy Office's icons, ribbon artwork, theme XML, table-style definitions, templates, clip art or wording beyond feature names. Excel feature *names* (menu labels, function names) are fine.
- Screenshots of Microsoft software are never committed (they live only in `plan/excel/screenshots/`).
- The ArtCraft brand files in `docs/brand/` are ArtCraft trademarks under `docs/brand/LICENSE-brand.txt`, the one exception.

## Never crash
People trust GridCraft with their numbers; a crash loses their work. **This outranks feature work.** Standard: [`craftrules/standards/never-crash.md`](https://github.com/storytold/craftrules/blob/main/standards/never-crash.md).
- **No panics in non-test code:** no `unwrap()`, `expect()`, `panic!`, `unreachable!`, `todo!`, `unimplemented!`; no `unsafe` (`unsafe_code = "forbid"`).
- **Errors are `Result<T, E>`** through the crate's error type and `?`.
- **Input-derived numbers are hostile** (files, formulas, commands, MCP/control params): `get()` instead of indexing, checked/saturating arithmetic, cap input-sized allocations (huge ranges, `REPT`, `SEQUENCE`, whole-column references).
- **Bound recursion** (formula nesting, dependency chains, names referring to names).
- **Last-resort guard:** `Session::execute` catches escaped panics and restores the workbook. It's a safety net, not a licence.
- **Prove it:** every crash fix lands with a regression test (`tests::every_command_survives_empty_params` fuzzes every command; `crates/engine/tests/it/command_sequence_fuzz.rs` fuzzes random command sequences against undo/redo, XLSX round-trip and full-recalc invariants; soak it with `GRIDCRAFT_FUZZ_SEEDS`/`GRIDCRAFT_FUZZ_STEPS`, replay with `GRIDCRAFT_FUZZ_SEED`).
- Crates carry `#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic, clippy::unimplemented, clippy::todo, clippy::unreachable)]`.

## Non-negotiables
- **Clean-room.** Microsoft Excel is installed on the dev machine and may be *observed* black-box: run it, use its UI with synthetic workbooks, take screenshots (by window id) stored only under `plan/excel/screenshots/`. Never read, disassemble or copy anything inside the Office bundle, never commit files produced by Excel, never copy GPL/AGPL/LGPL code (LibreOffice, Gnumeric…). Function behaviour comes from public documentation and observation.
- **Fonts live in [`storytold/craft-fonts`](https://github.com/storytold/craft-fonts)**, never in this repo. It is an optional build input (`CRAFT_FONTS_DIR`); without it the app uses system fonts and egui's defaults. Standard: craftrules `standards/fonts.md`.
- **Test corpora** (real Excel-authored files) live in a separate corpus repo, fetched pinned and sha256-verified (craftrules `standards/test-corpora.md`). Never commit large binary fixtures; build test workbooks in code.
- **Everything is a command.** User-visible behaviour = a command in `crates/engine/src/cmd/*` (id, label, ribbon path, shortcut, params doc, `enabled`, `run`) + tests. Ids follow Excel's ribbon (`home.bold`, `insert.chart`, `data.sortAscending`) plus primitives (`cell.set`, `selection.set`, `range.setValues`). UI-only commands live in `crates/ui-egui`. The control channel, CLI and MCP reach all of them.
- **Programmatic calls never open dialogs**: a command with empty params runs with defaults or returns an error; only menu-style invocation opens a dialog (`UiRequest::Dialog`).
- **Layering** is enforced by `cargo xtask layers`: `core`/`numfmt` (L0) → `formula`/`functions` (L1) → `model`/`calc`/`xlsx` (L2) → `chart` (L3) → `engine` (L5) → `ui-egui`/`mcp` (L6) → `apps/*` (L7). Nothing below L6 depends on egui/eframe/winit/rfd. **The UI crate is swappable.**
- **The UI is thin**: panels read engine state and act through `app.run(id, params)`. Colours come from `theme::Tokens`.
- **Time goes through `chrono`** (workspace dependency). Never hand-roll UTC/calendar conversion (days-to-date algorithms, leap-year math). A crate owns only its output format string. The one exception is range extension: past the years chrono can represent, shift the date by whole 400-year Gregorian cycles (146,097 days, after which the calendar repeats), let chrono compute the month and day, and add the cycles back to the year (splitting seconds into hours, minutes and seconds is plain arithmetic).
- **Rust only** (no handwritten JS/TS). **Never break wasm** (`cargo xtask wasm`).
- **Quality gates** before every commit: `cargo xtask ci` (fmt, clippy -D warnings, tests, assets, layers, wasm). One task id per commit (`M2.1: borders gallery`).

## Running and looking at the app
- `cargo run --release -p gridcraft -- --sample sales --control 7979` (sample workbook + control channel).
- Drive it: `gridcraft-cli send 7979 <method> '<json>'` (reads the token file the app printed on stderr), or JSON lines on `127.0.0.1:7979` whose first line is `{"id":0,"method":"auth","params":{"token":"…"}}` (token: `$XDG_RUNTIME_DIR/gridcraft/control-7979.token`, or your own `GRIDCRAFT_CONTROL_TOKEN`), e.g. `{"id":1,"method":"engine.execute","params":{"command":"cell.set","params":{"cell":"B2","input":"=SUM(A1:A9)"}}}` then `{"id":2,"method":"ui.screenshot","params":{"path":"/tmp/shot.png"}}`. Methods: `crates/ui-egui/src/control.rs`, docs: `docs/control-protocol.md`.
- **Offscreen UI render** (no window, from a source checkout): `cargo run -p gridcraft-ui-egui --example snapshot -- --sample sales out.png`.
- Headless: `gridcraft-cli eval '=SUM(1,2,3)'`, `gridcraft-cli run --in book.xlsx --cmd 'home.bold={"range":"A1:C1"}' --out book.xlsx`, `gridcraft-cli mcp`.
- **For UI work, look at the result** (snapshot PNG) and compare with `plan/excel/observed-ui.md`.
- Shell gotcha: `mv`/`cp` are aliased interactive here — use `/bin/mv -f` / `/bin/cp -f`. macOS has no `timeout`; use `perl -e 'alarm 60; exec @ARGV' cmd`.
- Parallel agents: separate `CARGO_TARGET_DIR` per agent; edit only the crates you own; delete your target dir when done.

## Roadmap
`ROADMAP.md` (committed) is the one-page summary: stage, headline numbers, dimensions, languages, progress log. It follows craftrules `standards/progress-docs.md` and points at the detail in `docs/`: `target-app-parity.md` (authoritative assessment), `gaps.md` (ranked work list: pick from it), `roadmap.md` (current focus, beta gates, milestones), `architecture.md`, and the area checklists `function-parity.md`, `chart-parity.md`, `file-format-parity.md`, `ui-parity.md`, `hardware-parity.md`, `localization-parity.md`. Update the affected docs (and their timestamp lines and revision history) whenever work lands. `cargo xtask parity` regenerates `docs/parity-checklist.md` from `crates/engine/src/catalog.rs`.

## Contributor credits (About window)

- About ▸ Contributors/Models are compiled into the binary from `contributors/contributors.json`
  (commit stats; generated, never hand-edit) and `contributors/people.toml` (names people chose for
  themselves). See `docs/contributors.md`.
- **Agents working for a contributor:** when you prepare a PR, check whether your human's GitHub
  username has a `[people.<username>]` entry in `contributors/people.toml`. If not, ask them once
  whether they want to be credited by more than their username: a real name, a display name, and/or
  their public GitHub profile name (`sync_github_name = true`). If yes, add **only their own** entry
  (copy the template at the top of the file, or run
  `python3 ../../craftrules/scripts/contributors.py --add-me . --real-name "…" --sync-github-name`)
  and include it in their PR, committed as them. If no, change nothing: they are credited as
  `@username` anyway.
- Never add, edit, guess or copy anyone else's entry or name (not from git config, commit authors or
  GitHub profiles). Never hand-edit `contributors.json`.
- Maintainers refresh the stats with `python3 ../../craftrules/scripts/contributors.py .` (it also
  re-verifies who wrote each `people.toml` entry; `--check` only verifies).
