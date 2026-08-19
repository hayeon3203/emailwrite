# Requirements.md — 실행 하드웨어 요구사항

측정은 전부 아래 "기준 머신"에서 나온 값이다. 수치의 출처와 재현 방법은 `Experiments.md`.

## 기준 머신 (모든 측정의 출처)

| | |
|---|---|
| CPU | Intel Core Ultra 7 258V (Lunar Lake, 8코어 — P 4 + LP-E 4, HT 없음) |
| GPU | Intel Arc 140V (16GB 공유), 드라이버 32.0.101.8724 |
| OS | Windows 11 Pro 26200 |
| 메모리 | 통합 메모리 (GPU가 시스템 메모리를 나눠 씀) |

## 최소 요구사항 — 로컬 모델로 쓸 때

| 항목 | 요구 | 안 맞으면 |
|---|---|---|
| OS | Windows 10/11 **x64** | 실행 안 됨 (WPF + Outlook COM) |
| Outlook | **클래식 데스크톱** (Office16 Win32) | 메일을 못 읽음. 초안은 "자유 작성 모드"로만 동작 |
| CPU | **AVX2 지원** | 앱이 시작 시 거부하고 API 모드를 안내한다 (`Avx2.IsSupported`) |
| 메모리 | 여유 **3GB** 이상 | GPU 없는 PC에서 모델 1.08GB(mmap) + KV 384MB + 컴퓨트 427MB ≈ **1.9GB**를 상시 잡는다. GPU가 붙으면 이 중 대부분이 GPU로 넘어간다 (아래) |
| 디스크 | **1.2GB** (`C:\ProgramData\EmailWrite\Models\`) + 앱 | 첫 실행이 다운로드에서 실패 |
| 네트워크 | **첫 실행 1회만** — 모델 1.06GB 다운로드 | 그 뒤로는 완전 오프라인 동작 |

AVX2는 대략 2013년(Haswell) 이후 인텔 / 2015년(Excavator) 이후 AMD면 있다. 그보다 오래된 PC는
로컬 모델을 못 쓰고, **API 모드(Solar)로는 정상 동작한다** — 아래 참조.

## 권장 — 30초 밑으로 쓰려면 GPU가 필요하다

로컬 초안의 비용은 **93%가 프리필**(프롬프트를 읽는 시간)이고, 그건 GPU가 제일 잘하는 일이다.
같은 프롬프트(1,931자) 1턴 실측 — 2026-08-14, 현재 기본 모델(CLOVA 1.5B):

| 구성 | 프리필 | 1턴(모델 로드 포함) | 2·3턴 |
|---|---|---|---|
| 레이어를 안 올림 (`GPU_LAYERS=0`) | 6.3s | 9.2s | 1.5s / 1.5s |
| **레이어를 GPU에 올림 (기본값)** | **3.7s** | **8.9s** | 0.9s / 0.8s |

**긴 메일에서 벌어진다.** 위 표의 1턴 차이가 작아 보이는 것은 가짜 짧은 메일이라 프리필 자체가
싸기 때문이다. 실제 메일(프롬프트 3,366자)에서는 프리필이 턴 시간의 대부분을 먹는다 —
`Experiments.md` §A5·§D.

| 항목 | 요구 |
|---|---|
| GPU | **Vulkan 1.x** 지원 (Intel Arc·Iris Xe, AMD, NVIDIA 모두 해당) |
| 드라이버 | 그래픽 드라이버에 Vulkan 런타임이 딸려온다. 별도 설치 없음 |
| GPU 메모리 | **1.6GB** (전용 VRAM 또는 공유 메모리) — 가중치 956MB + KV 384MB + 컴퓨트 220MB |

**GPU가 없어도 앱은 그대로 돈다.** CPU 백엔드가 같이 들어있어서, Vulkan 장치가 응답하지 않으면
llama.cpp가 알아서 CPU로 떨어진다 — 느려질 뿐 실패하지 않는다. 대신 위 최소 요구사항의
메모리를 전부 시스템 메모리에서 낸다.

### 지금 어느 쪽으로 돌고 있는지 확인

팝업 상태줄의 맨 앞을 본다:

```
커널 avx2 544K +vulkan · 프리필 3.7s · 17토큰 42.1t/s · 프롬프트 1931자
      ^^^^          ^^^^^^^
      CPU 커널       Vulkan 백엔드가 로드됨
```

- `+vulkan`이 **없으면** Vulkan 장치가 없는 것이다. 이 표시는 `ggml-vulkan.dll`이 매핑됐는지만
  보고, **레이어를 실제로 올렸는지는 말해주지 않는다** — `EMAILWRITE_GPU_LAYERS=0`으로 돌려도
  `+vulkan`은 그대로 붙는다. 백엔드가 등록만 돼 있어도 큰 배치 matmul(=프리필)은 GPU가 가져가고,
  컴퓨트 버퍼도 GPU에 남는다. 올렸는지 아닌지는 토큰 속도로 갈린다 (위 표: 42t/s 대 9t/s)
- `win-x64 427K`처럼 나오면 네이티브 백엔드 폴더가 평탄화돼 **noavx 커널**로 도는 것이다.
  8배 느려진다 — `Experiments.md` §8

## API 모드 — 로컬 요구사항 없음

`EMAILWRITE_BACKEND=solar`(또는 팝업 드롭다운에서 선택)이면 초안은 서버에서 만들어진다.

| 항목 | 요구 |
|---|---|
| CPU / GPU / 메모리 | **해당 없음.** AVX2도 필요 없다 |
| 디스크 | 모델을 안 받는다 |
| 네트워크 | **상시 연결 필요** |
| 키 | 팝업에서 입력받아 DPAPI로 암호화해 `%LocalAppData%\EmailWrite\keys.json`에 저장한다. 설치 시 미리 넣어둘 게 없다. 환경변수(`UPSTAGE_API_KEY` / `ANTHROPIC_API_KEY`)는 폴백으로 남아있고, 파일이 이긴다 |

턴당 0.5~1.2s. **대신 메일 본문이 외부로 전송된다** — 사내 정책에 따라 못 쓸 수 있고, 그래서
기본값이 아니다.

구형 PC에 배포할 때의 실질적인 답은 "AVX2가 없으면 API 모드"다.

## 조정 노브

기본값(미설정)이 위 표의 측정 조건이다.

| 변수 | 용도 |
|---|---|
| `EMAILWRITE_GPU_LAYERS` | GPU에 올릴 레이어 수. 기본 99(전부), `0`이면 CPU만 |
| `EMAILWRITE_THREADS` | 추론 스레드 수. 기본은 논리 코어 전부 |
| `EMAILWRITE_MODEL` | 다른 `.gguf`로 교체 (캐시 폴더에 직접 받아둔 파일만) |
| `EMAILWRITE_BACKEND` | `solar` / `claude` / 미설정(로컬) |

## 모델 파일

`C:\ProgramData\EmailWrite\Models\` (`%LocalAppData%`가 **아니다** — llama.cpp가 ANSI `fopen`으로
파일을 열어서 `C:\Users\조하연\...` 같은 비ASCII 경로가 깨진다)

| 파일 | 크기 | 라이선스 |
|---|---|---|
| `hyperclovax-seed-text-instruct-1.5b-q4_k_m.gguf` (기본) | 1.06GB | HyperCLOVA X SEED — **사내 배포 가능** (확인 2026-08-14). 조건은 유사 서비스로의 상업화 금지고, 이 앱은 사내 초안 도구라 해당하지 않는다. 약관 사본을 exe 옆에 함께 배포한다 (3.1(ii)). 같은 자리에 `NOTICE.txt`가 저작권 문구를 담고(3.1(iv)), 채팅 팝업 최하단에 "Powered by HyperCLOVA X"를 표시한다(3.1(v)) |
| `midm-2.0-mini-instruct-q4_k_m.gguf` | 1.33GB | MIT |

기본 모델만 자동으로 내려받는다. 나머지는 직접 받아 넣어야 `EMAILWRITE_MODEL`이 먹는다.

기본 모델은 개인이 변환해 올린 파일이라 **다운로드 뒤 SHA-256을 검증한다**
(`LocalDraftGenerator.ModelSha256`). 업스트림이 force-push하면 검증에서 걸린다 —
모델을 바꿀 때 파일명·URL·해시 셋을 같이 옮겨야 한다.
