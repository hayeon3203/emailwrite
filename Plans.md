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
- [ ] 클립보드 복사 동작 확인
- [ ] 포커스 잃으면 팝업 자동 숨김(옵션)
- [ ] API/Outlook 접근 실패 시 사용자에게 보이는 에러 메시지

### M7. 통합 테스트
- [ ] 실제 Outlook 메일 열어두고 End-to-End 시나리오
- [ ] 단축키 경로 / 클릭 경로 둘 다 확인

## MVP 범위 밖 (추후)
- 타 메신저 앱 자동 입력 (UI Automation)
- 새 Outlook / OWA(웹) 지원
- 자동 시작, 설치 프로그램 패키징
- 스트리밍 응답, 대화 히스토리 영구 저장

## 진행 중 결정 필요
- Anthropic API 키 발급/보관 방식
- 아바타 이미지 에셋 (M1은 임시 아이콘으로 시작, 나중에 교체)
