# Plans.md — Outlook 컨텍스트 기반 메신저 초안 아바타 (MVP)

## 목표
바탕화면 아바타를 클릭하거나 전역 단축키를 누르면, 현재 Outlook에서 보고 있는 메일(스레드)을 컨텍스트로 읽어 메신저 답장 초안을 LLM으로 생성하고, 아바타 위 작은 채팅 팝업에서 대화하며 다듬은 뒤 최종 문구를 복사해서 쓰는 도구.

## 확인된 전제
- Outlook: 클래식 데스크톱(Office16, Win32) — COM Object Model로 공식 API 접근 가능
- 출력 방식: 자동 입력 아님. 채팅 팝업에서 확인/수정 후 사용자가 직접 복사
- 플랫폼: Windows 전용 상주 앱

## 스택 결정
- .NET 8 + WPF (Outlook COM interop이 네이티브로 매끄럽고, 오버레이/트레이/단축키가 표준 기능)
- `Microsoft.Office.Interop.Outlook` (NuGet)
- Claude API (Anthropic Messages API) — `HttpClient` 직접 호출로 시작

## 마일스톤

### M0. 프로젝트 초기화 ✅
- [x] `EmailWrite.sln` + WPF 프로젝트(.NET 8) 생성
- [x] Interop.Outlook NuGet 추가, 빈 창 실행 확인
- [x] `.gitignore` (bin/obj/설정 파일 중 시크릿 제외)

### M1. 바탕화면 아바타 오버레이 ✅
- [x] `WindowStyle=None`, `AllowsTransparency=True`, `Topmost=True` 창
- [x] 이미지 표시(픽셀 아바타 스프라이트, 아이들 애니메이션), 마우스 드래그로 위치 이동
- [x] 화면 우하단 기본 위치, 종료용 트레이 아이콘

### M2. 트리거 ✅
- [x] 전역 단축키 등록 (기본 `F9`, 충돌 시 트레이 풍선 알림으로 에러 처리)
- [x] 아바타 클릭 → 채팅 팝업 토글 (플레이스홀더 팝업, 실제 UI는 M4)
- [x] 단축키 → 채팅 팝업 토글 (동일 핸들러 재사용)

### M3. Outlook 컨텍스트 리더 ✅
- [x] 열려있는 모든 Inspector/Explorer를 후보로 모아 실제 OS Z-order상 가장 최근에 활성화된 창을 선택 (여러 메일 창이 동시에 떠있어도 마지막으로 클릭한 창 기준)
- [x] 추출 필드: 제목, 발신자, 본문, 수신 시각, 첨부파일명(서명 인라인 이미지는 PR_ATTACHMENT_HIDDEN으로 필터링)
- [x] `MailItem.GetConversation()`으로 스레드 히스토리(최근 N개) 추가 추출
- [x] Outlook 미실행/선택된 메일 없음 예외 처리
- [x] 콘솔/디버그 출력으로 먼저 검증 (`tools/OutlookProbe`, UI 연결 전 단독 테스트)
- [x] 보안 경고 팝업 발생 여부 실제 확인 — SenderName만 사용(SenderEmailAddress 미접근)해서 실제 실행 시 팝업 안 뜸

### M4. 채팅 팝업 UI ✅
- [x] 아바타 위쪽에 뜨는 borderless 작은 창
- [x] 메시지 리스트 + 입력창 + 전송 버튼
- [x] 메시지별 "복사" 버튼 (클립보드)
- [x] 팝업 오픈 시 Outlook 컨텍스트를 시스템 메시지로 주입 + 초안 자동 생성 1회 트리거 (`IDraftGenerator` — M5 전까지는 `PlaceholderDraftGenerator`가 자리만 채움)

### M5. LLM 연동 (Claude API) ✅
- [x] API 키 관리 (환경변수 `ANTHROPIC_API_KEY`, git 커밋 금지)
- [x] Anthropic Messages API 호출 (비스트리밍, `claude-haiku-4-5`) — 공식 `Anthropic` C# SDK 사용
- [x] 시스템 프롬프트: Outlook 컨텍스트(현재 메일 + 스레드 히스토리) + "메신저 답장 초안 작성" 지시
- [x] 팝업 세션 내 멀티턴 대화 상태 유지 (`ChatPopupWindow.Messages`를 매 요청마다 전체 전달)

### M6. 마무리
- [x] (2026-08-07) `ANTHROPIC_API_KEY` 크레딧 소진으로 `ClaudeDraftGenerator` 대신 M8 최소 구현(`LocalDraftGenerator`, Qwen3-1.7B)으로 테스트 진행 — API 재발급 시 다시 전환
- [x] 로컬 생성 파이프라인 단독 검증 완료 (Outlook 없이 가짜 `EmailContext`로 1턴 초안 + 2턴 톤 수정까지 실제 추론 성공)
- [x] 클립보드 복사 동작 구현 (복사 후 버튼이 "복사됨"으로 1초간 바뀜, 다른 프로세스가 클립보드 점유 중이면 무시)
- [x] 포커스 잃으면 팝업 자동 숨김 (아바타 클릭으로 인한 포커스 이동은 제외 — 토글 핸들러와 경합하지 않도록)
- [x] API/Outlook 접근 실패 시 사용자에게 보이는 에러 메시지 (+ `App.OnStartup`에 UI 스레드 미처리 예외 핸들러 — 트레이 앱이라 콘솔이 없어 그냥 죽으면 원인을 알 수 없음)
- [x] (2026-08-09) **실제 Outlook 띄우고 UI에서 확인** — 복사 피드백("복사됨" 1초), 자동 숨김, 에러 표시 전부 실제 화면에서 확인

### M7. 통합 테스트 ✅ (2026-08-09)
- [x] 실제 Outlook 메일 열어두고 End-to-End 시나리오 — 실제 사내 메일(본문 4,585자)로 초안 생성(약 21초) → 톤 수정 2턴 → 클립보드 복사까지 성공
- [x] 단축키 경로 / 클릭 경로 둘 다 확인
- 이 과정에서 나온 버그 3개, 전부 수정함:
  - **프로세스 통째로 죽음**: `llama_decode`에서 `AccessViolationException`. llama.cpp 컨텍스트는 재진입 불가인데, 생성 중 F9를 다시 누르면 두 번째 추론이 같은 컨텍스트로 들어감 → `LocalDraftGenerator._inferLock`으로 추론 직렬화 (호출자마다 막지 않고 한 곳에서)
  - **`NoKvSlot` 에러 / 추론이 몇 분씩 걸림**: 프롬프트가 4096 컨텍스트를 넘김. 원인 두 가지 — (1) `GetConversation()`이 답장 대상 메일 자신을 히스토리에 또 넣어서 같은 본문이 두 번 (9,170자), (2) 본문을 자르지 않아 서명·면책문구·인용까지 전부 들어감. `EntryID`로 자기 자신 제외 + `DraftPromptBuilder`에서 본문 절단(현재 메일 1,200자 / 히스토리 300자)
  - **생성 중 팝업이 사라져 초안을 날림**: 로컬 추론이 10초 넘게 걸리는 동안 다른 창이 포커스를 가져가면 자동 숨김이 발동. `ChatPopupWindow.IsBusy` 동안은 숨기지 않고, 재표시해도 `OpenAsync`를 다시 돌리지 않음
- [ ] **초안 품질 문제 (신규)**: 2턴째 응답이 (1) "네, 더 친근한 말투로 다시 작성해드릴게요" 같은 머리말을 붙임(시스템 프롬프트에서 금지했는데도), (2) 메신저체가 아니라 이메일체로 돌아감, (3) 본문의 "받는 사람"을 보고 사용자를 '박선영'으로 착각해 서명함. 1턴 초안은 멀쩡했음 — 멀티턴 프롬프트 쪽 문제로 보임

### M8. 로컬 LLM 지원 (프라이버시/비용 동기)
- [x] `LLamaSharp` 도입, CPU 전용 추론(AVX/AVX2/AVX512 백엔드 풀세트 번들 — Q13/Q14에서 200~300MB 증가 감수하기로 결정)
- [x] `IDraftGenerator` 구현체로 `LocalDraftGenerator` 추가 (기존 `ClaudeDraftGenerator`/`PlaceholderDraftGenerator`와 동일 인터페이스) — M6 테스트 언블록을 위해 최소 구현으로 선반영, `MainWindow`에서 임시로 기본 연결
- [x] 사내 배포 기본 모델: **Mi:dm-2.0-Mini-Instruct** q4_k_m (KT, 2.3B, **MIT**, 한국어 특화) — 실측 CPU 약 11~15초/응답, 1.4GB
  - 실측 비교(`tools/DraftProbe`, 동일 프롬프트): Mi:dm은 "네, 확인했습니다 … 금요일 전에 의견 정리해서 보내드릴게요"처럼 역할 반전이 전혀 없고 2턴 톤 수정도 제대로 반영. Qwen3-1.7B는 마지막 문장에 발신자의 "회신 부탁드려요"를 계속 되돌려주고 톤 수정도 거의 안 먹었음. 속도도 Mi:dm이 더 빠름
  - Qwen2.5-1.5B-Instruct는 폐기: 프롬프트를 네 가지로 바꿔도 발신자 입장에서 쓰는 역할 반전이 안 고쳐짐
  - 대안으로 Qwen3-1.7B q4_k_m(Apache-2.0)도 캐시에 받아둠. Qwen3 계열은 추론 모드를 `/no_think`로 꺼야 함(켜면 43초 → 끄면 17초) — `LocalDraftGenerator`가 파일명에 qwen이 있을 때만 붙임. 혹시 새는 `<think>` 블록은 출력에서 제거
- [x] `EMAILWRITE_MODEL` 환경변수로 재빌드 없이 모델 교체 (캐시에 이미 있는 파일만 — 자동 다운로드는 기본 모델 URL만 알기 때문에 오버라이드 시엔 다운로드하지 않고 에러)
- [x] `tools/DraftProbe`: Outlook 없이 가짜 `EmailContext`로 1턴 초안 + 2턴 톤 수정을 돌려 모델별 품질/속도 비교
- [x] 모델 캐시는 `%LocalAppData%`가 아니라 `%ProgramData%\EmailWrite\Models\`
  - llama.cpp의 `gguf_init_from_file`이 ANSI `fopen`으로 모델을 열기 때문에, 사용자 이름이 비-ASCII면(`C:\Users\조하연\...`) 경로가 깨져서 로드가 무조건 실패함. `C:\ProgramData`는 항상 ASCII이고 Users 그룹 쓰기 가능, 덤으로 PC당 1회만 받으면 됨
- [x] 프롬프트 구조: 역할 지시만 시스템 프롬프트에, 메일 본문은 첫 user 턴에, 그 앞에 올바른 답장 예시를 user/assistant 턴 쌍으로 1-shot 삽입 (`DraftPromptBuilder`, Claude/로컬 공용)
  - 같은 예시를 시스템 프롬프트에 **텍스트로** 넣으면 작은 모델이 예시를 그대로 베껴서 출력함 — 반드시 완료된 대화 턴으로 넣어야 함
- [ ] (선택) 추가 모델 검토 — 지금은 필요 없음. 검토했다 뺀 것: Gemma 3 2B(존재하지 않음 — Gemma 3는 1B/4B/12B/27B), Mistral Small(24B, q4로도 14GB라 CPU 불가), Ministral-8B(라이선스 `other`), Llama 3.2 3B(한국어가 공식 지원 언어 목록에 없음), Phi-4-mini(MIT지만 영어/수학 편중), A.X-4.0-Light(Apache-2.0인데 7B라 CPU에서 너무 느림), Kanana-nano-2.1b(CC-BY-NC-4.0, 비상업 전용)
  - 다음 후보가 필요하면: Qwen3-4B-Instruct-2507(Apache-2.0, 추론모드 기본 off, q4 약 2.5GB — 속도가 관건), gemma-3-4b-it(Gemma 라이선스)
- [ ] 최소 사양(AVX2) 미달 시 로컬 모드 비활성화 + API 모드 유도 에러 메시지 (현재는 로드 실패를 범용 에러 메시지로만 표시 — 전용 사양 감지는 M9)

### M9. 모드 전환 / 온보딩 / API 키 등록
- [ ] 최초 실행 시 강제 온보딩 없음 — 채팅 팝업 첫 시도 시점(`ToggleChatPopup`/`OpenAsync`)에 "로컬 또는 API 둘 중 하나 미설정" 감지해서 설정 창으로 유도
- [ ] 로컬 모델 다운로드는 백그라운드 + 진행률 표시(채팅 팝업의 `TypingIndicator` 자리 재활용 검토)
- [ ] 설정 창: 트레이 우클릭 메뉴에 "설정" 항목 추가
- [ ] API 키는 UI로 입력받아 DPAPI(`ProtectedData.Protect`)로 암호화해 `%AppData%\EmailWrite\settings.json`에 저장 (환경변수 `ANTHROPIC_API_KEY` 방식은 폐기)
- [ ] 로컬/API 둘 다 설정된 경우에만 설정 창에서 모드 전환 가능, 기본값은 로컬(프라이버시 우선)

### M10. 배포 / 자동 업데이트
- [ ] .NET 8 self-contained 게시 (동료 PC에 런타임 사전 설치 불필요)
- [ ] Inno Setup으로 설치 파일 제작 (Program Files 설치, 시작 메뉴, 제거 프로그램 등록)
- [ ] GitHub 리포 public 전환 (사전 확인: 전체 커밋 히스토리에 시크릿 커밋 이력 없음 확인 완료 — 2026-08-07)
- [ ] 앱 내 자동 업데이트 체크 (하루 1회 `GET /repos/.../releases/latest` 인증 없이 호출, 새 버전이면 트레이 알림 → 설치 파일 다운로드[모델 다운로드 로직 재사용] → 재실행)
- [ ] 버전 비교는 `System.Version` 문자열 비교로 충분(semver 라이브러리 불필요)

## MVP 범위 밖 (추후)
- 타 메신저 앱 자동 입력 (UI Automation)
- 새 Outlook / OWA(웹) 지원
- 스트리밍 응답, 대화 히스토리 영구 저장
- GPU 가속 로컬 추론 (현재는 CPU 전용으로 확정 — GPU 사용자 많다는 신호 확인되면 재검토)
- MSIX 패키징 (인증서/신뢰 배포 문제로 Inno Setup 선택, 배포 규모 커지면 재검토)

## 진행 중 결정 필요
- Anthropic API 키 발급/보관 방식
- 아바타 이미지 에셋 (M1은 임시 아이콘으로 시작, 나중에 교체)
