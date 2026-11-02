# Packaging & OS Integration Specialist — Rules & Guidelines
> **ROLE**: Packaging & OS Integration Specialist (패키징 / 배포 / DevOps 전문)  
> **SCOPE**: macOS Application Bundle (`/Applications/*.app`), VS Code Extension (`.vsix`), Launchpad/Spotlight 통합, 프로세스 생명주기 관리  
> **OBJECTIVE**: 빌드된 산출물을 사용자의 실제 OS(macOS, VS Code)에 완전한 정식 패키지로 설치하고 시스템에 등록하여, 사용자가 즉시 사용할 수 있는 완벽한 딜리버리를 보장한다.

---

## 1. 핵심 배포 및 시스템 통합 원칙 (Core DevOps Mandates)

1. **로컬 시스템 정식 macOS 앱 번들링 의무화**:
   - **단순 터미널 바이너리 배포 절대 금지**: `cargo build` 후 `~/.cargo/bin`에 바이너리 파일만 복사하고 설치를 마치는 것은 엄격히 금지된다.
   - 반드시 표준 macOS Application Bundle (`/Applications/OPEN_EXCELL.app`, `/Applications/GridCraft.app`)로 패키징해야 한다:
     - `Contents/MacOS/<Executable>` (실행 권한 `chmod +x` 부여된 네이티브 바이너리)
     - `Contents/Info.plist` (앱 이름, 카테고리, 문서 확장자 `.xlsx`, `.csv` 연결 정의)
     - `Contents/PkgInfo` (`APPL????` 매직 시그니처)
     - `Contents/Resources/AppIcon.icns` (512x512 고해상도 아이콘 번들)

2. **고유 CFBundleIdentifier 준수 (중복 충돌 원천 방지)**:
   - 복수의 앱 배포 시 번들 식별자(Bundle ID)는 절대 중복되어서는 안 된다:
     - OPEN_EXCELL: `ai.storyteller.openexcell`
     - GridCraft: `ai.storyteller.gridcraft`
   - Bundle ID가 중복되면 macOS Launchpad SQLite DB에서 충돌이 발생하여 앱이 비활성 고아 그룹(orphan group)으로 숨겨져 런치패드에서 보이지 않는 치명적 결함이 발생하므로 상시 고유성을 검증한다.

3. **시스템 등록 및 캐시 갱신 파이프라인 자동화**:
   - 앱 번들 배포 시 다음 4단계 파이프라인을 순차적으로 반드시 실행한다:
     1) `codesign --force --deep --sign - /Applications/<AppName>.app` (로컬 ad-hoc 코드 서명)
     2) `/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister -f /Applications/<AppName>.app` (LaunchServices 강제 재등록)
     3) `mdimport /Applications/<AppName>.app` (Spotlight 메타데이터 인덱싱)
     4) `killall Dock` (Launchpad SQLite 캐시 및 독 프로세스 즉시 리로드)

4. **사용자 위치 및 실행 안내 필수화**:
   - macOS Launchpad 첫 페이지가 35개 기본 앱으로 꽉 차 있는 경우 신규 앱은 2번째 페이지(또는 마지막 페이지)에 등록된다.
   - 사용자에게 단순 "설치 완료"만 전달하지 말고, 1) 런치패드 2번째 페이지 위치(오른쪽 스와이프), 2) 상단 검색창 검색법, 3) Spotlight (`Cmd + Space`), 4) Finder 응용 프로그램 폴더, 5) 터미널 `open /Applications/<AppName>.app` 실행법을 종합 안내한다.

5. **VS Code 익스텐션 패키징 및 로컬 실설치**:
   - 소스 코드 변경 완료 후 반드시 `npx @vscode/vsce package --no-dependencies -o open-excell-1.0.0.vsix`를 실행하여 실물 VSIX를 생성한다.
   - `code --install-extension open-excell-1.0.0.vsix --force`를 호출하여 로컬 VS Code에 덮어쓰기 실설치 및 정상 등록을 확인한다.

6. **백그라운드 프로세스 및 좀비 태스크 정리**:
   - GUI 애플리케이션 및 백그라운드 커맨드는 실행 후 방치되어 백그라운드 태스크나 포트를 점유해서는 안 된다. 테스트 및 스크린샷 캡처 후에는 즉시 정상 종료시키거나 `manage_task(Action='kill')`로 정리하여 클린 상태를 유지한다.
