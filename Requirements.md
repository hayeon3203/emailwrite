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
| 메모리 | 여유 **4GB** 이상 | 모델 1.33GB(mmap) + KV 768MB + 컴퓨트 287MB ≈ **2.5GB**를 상시 잡는다 |
| 디스크 | **1.4GB** (`C:\ProgramData\EmailWrite\Models\`) + 앱 | 첫 실행이 다운로드에서 실패 |
| 네트워크 | **첫 실행 1회만** — 모델 1.4GB 다운로드 | 그 뒤로는 완전 오프라인 동작 |

AVX2는 대략 2013년(Haswell) 이후 인텔 / 2015년(Excavator) 이후 AMD면 있다. 그보다 오래된 PC는
로컬 모델을 못 쓰고, **API 모드(Solar)로는 정상 동작한다** — 아래 참조.

## 권장 — 30초 밑으로 쓰려면 GPU가 필요하다

로컬 초안의 비용은 **93%가 프리필**(프롬프트를 읽는 시간)이고, 그건 GPU가 제일 잘하는 일이다.
같은 프롬프트(2,974자) 1턴 실측:

| 구성 | 1턴 | 2·3턴 |
|---|---|---|
| CPU만 (AVX2) | **34.6s** | 3.1s / 1.7s |
| **Vulkan GPU (기본값)** | **13.4s** | 2.1s / 1.8s |

| 항목 | 요구 |
|---|---|
| GPU | **Vulkan 1.x** 지원 (Intel Arc·Iris Xe, AMD, NVIDIA 모두 해당) |
| 드라이버 | 그래픽 드라이버에 Vulkan 런타임이 딸려온다. 별도 설치 없음 |
| GPU 메모리 | **2.5GB** (전용 VRAM 또는 공유 메모리) |

**GPU가 없어도 앱은 그대로 돈다.** CPU 백엔드가 같이 들어있어서, Vulkan 장치가 응답하지 않으면
llama.cpp가 알아서 CPU로 떨어진다 — 느려질 뿐 실패하지 않는다.

### 지금 어느 쪽으로 돌고 있는지 확인

팝업 상태줄의 맨 앞을 본다:

```
커널 avx2 544K +vulkan · 프리필 8.9s · 24토큰 24.4t/s · 프롬프트 2974자
      ^^^^          ^^^^^^^
      CPU 커널       GPU가 레이어를 받았음
```

- `+vulkan`이 **없으면** GPU를 못 쓰고 있는 것 (장치 없음, 또는 `EMAILWRITE_GPU_LAYERS=0`)
- `win-x64 427K`처럼 나오면 네이티브 백엔드 폴더가 평탄화돼 **noavx 커널**로 도는 것이다.
  8배 느려진다 — `Experiments.md` §8

## API 모드 — 로컬 요구사항 없음

`EMAILWRITE_BACKEND=solar`(또는 팝업 드롭다운에서 선택)이면 초안은 서버에서 만들어진다.

| 항목 | 요구 |
|---|---|
| CPU / GPU / 메모리 | **해당 없음.** AVX2도 필요 없다 |
| 디스크 | 모델을 안 받는다 |
| 네트워크 | **상시 연결 필요** |
| 키 | `UPSTAGE_API_KEY` 환경변수 (Claude는 `ANTHROPIC_API_KEY`) |

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
| `midm-2.0-mini-instruct-q4_k_m.gguf` (기본) | 1.33GB | MIT |
| `hyperclovax-1.5b-instruct-q4_k_m.gguf` | 1.06GB | 조건부 상업 이용 가능 — 배포 전 확인 필요 |

기본 모델만 자동으로 내려받는다. 나머지는 직접 받아 넣어야 `EMAILWRITE_MODEL`이 먹는다.
