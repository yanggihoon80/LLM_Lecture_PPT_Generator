# linkvalue_llm_lecture_ppt_generator

강의 커리큘럼과 프롬프트를 바탕으로 교시별 강의용 PPT를 자동 생성하는 프로젝트입니다.  
출력은 표준 `.pptx` 파일이며, 템플릿 PPT의 레이아웃을 참고해 슬라이드를 만듭니다.

## 개요
프로그램은 아래 순서로 동작합니다.

1. `template/` 폴더의 템플릿 PPT를 분석합니다.
2. `prompts/curriculum.txt`에서 교시별 주제와 핵심 내용을 읽습니다.
3. `prompts/lecture_prompt.txt`의 공통 프롬프트를 교시별 내용으로 채웁니다.
4. OpenAI API를 호출해 슬라이드 구조 JSON을 생성합니다.
5. JSON을 바탕으로 교시별 PPT를 생성합니다.
6. 전체 교시를 실행하면 마지막에 병합 PPT도 함께 생성합니다.

## 폴더 구조
```text
linkvalue_llm_lecture_ppt_generator/
├─ app.py
├─ run.ps1
├─ requirements.txt
├─ .env
├─ .env.example
├─ template/
├─ prompts/
│  ├─ lecture_prompt.txt
│  └─ curriculum.txt
└─ output/
```

## 템플릿 작성 규칙
`template/` 폴더에는 강의 장표의 기본 레이아웃으로 사용할 PPT 파일을 넣습니다.

현재 프로젝트는 아래 구조를 기준으로 가장 안정적으로 동작합니다.

- 슬라이드 상단: 제목 영역
- 슬라이드 좌측 상단: 본문 텍스트 영역
- 슬라이드 우측 하단: 이미지 영역
- 필요 시 좌측 하단: 표 또는 간단한 다이어그램 영역

권장 사항:

- 템플릿은 너무 복잡하지 않게 유지합니다.
- 본문, 이미지, 표/다이어그램 영역이 서로 겹치지 않게 잡습니다.
- 기본 글꼴은 `맑은 고딕` 사용을 권장합니다.

## 프롬프트 파일
### lecture_prompt.txt
모든 교시에 공통으로 적용되는 프롬프트 템플릿입니다.  
강의 스타일, 슬라이드 수, 출력 구조, 문장 톤 등의 공통 규칙을 넣습니다.

이 파일 안에는 보통 아래와 같은 공통 지시가 들어갑니다.

- 강의 주제별 장표를 만들 것
- 슬라이드 흐름이 자연스럽게 이어질 것
- 불릿, 표, 다이어그램 중 적절한 형식을 선택할 것
- 발표용 문장 톤으로 작성할 것

### curriculum.txt
교시별 제목과 핵심 내용을 적는 파일입니다.  
교시 수만큼 PPT가 생성되며, 각 교시의 제목과 핵심 내용이 여기서 결정됩니다.

예시:

```text
1교시. 생성형 AI 개요 및 업무 변화 이해
- 생성형 AI란 무엇인가
- 기존 자동화와 생성형 AI 차이
- 기업 활용 사례
- 한계와 주의점

2교시 : 업무 자동화의 본질 IPO 구조 재정의
- 입력, 처리, 출력 구조 이해
- 기존 자동화와 생성형 AI 자동화 비교
- 업무 흐름 재설계 관점
```

### curriculum.txt 파싱 규칙
- 교시 제목 줄은 아래 형식 중 하나를 사용합니다.
  - `1교시. 제목`
  - `1교시 : 제목`
  - `1교시: 제목`
- 교시 아래 줄들은 느슨한 파서로 읽습니다.
- 아래 형식은 모두 핵심 내용으로 인식합니다.
  - `- 내용`
  - `• 내용`
  - `1. 내용`
  - `내용`
- 줄 앞의 `-`, `•`, `1.` 같은 표시는 자동으로 제거됩니다.
- 다음 형식의 메타 정보 줄은 `core_points`가 아니라 `meta`로 자동 분리됩니다.
  - 지원 라벨: `메모`, `참고`, `실습`, `주의`, `비고`, `목표`, `준비물`, `과제`
  - `메모: 내용`
  - `참고: 내용`
  - `실습: 내용`
  - `[메모] 내용`
  - `[참고] 내용`
  - `[실습] 내용`
- 라벨만 한 줄로 쓰고 다음 줄에 내용을 쓰는 방식도 허용합니다.
  - `메모:`
  - `이 장표는 실습 전에 설명`
- 교시 수만큼 PPT가 생성됩니다.
- `curriculum.txt`에서 교시가 하나도 파싱되지 않으면, 프로그램은 교시별 생성이 아니라 단일 PPT 생성 경로로 동작할 수 있습니다.

### curriculum.txt 생성을 위한 LLM 프롬프트 예시
아래 프롬프트를 그대로 LLM에 넣으면 `curriculum.txt` 초안을 만드는 데 사용할 수 있습니다.

```text
너는 기업 교육 과정 설계자다.
아래 주제를 바탕으로 curriculum.txt 형식의 강의 커리큘럼 초안을 작성하라.

[주제]
생성형 AI 개요 및 업무 활용

[작성 조건]
1. 총 8교시로 구성하라.
2. 각 교시는 제목 1줄과 핵심 내용 3~5줄로 작성하라.
3. 각 핵심 내용은 curriculum.txt에 바로 넣을 수 있게 한 줄씩 작성하라.
4. 실습, 참고, 주의가 필요하면 메타 형식으로 작성하라.
5. 추상적인 표현보다 실제 강의에서 설명할 개념과 활동 중심으로 작성하라.

[출력 형식]
1교시. 제목
- 핵심 내용
- 핵심 내용
메모: 내용

2교시 : 제목
- 핵심 내용
- 핵심 내용
[실습] 내용

출력은 설명 없이 curriculum.txt 본문만 작성하라.
```

메타 정보까지 같이 생성하고 싶다면 아래처럼 더 구체적으로 요청하면 됩니다.

```text
각 교시마다 필요하면 아래 메타 라벨도 사용할 수 있다.
- 메모:
- 참고:
- 실습:
- 주의:
- 비고:
- 목표:
- 준비물:
- 과제:

출력은 반드시 curriculum.txt에서 바로 읽을 수 있는 평문 형식으로 작성하라.
```

## 환경 변수
`.env.example`을 복사해서 `.env`를 만든 뒤 값을 채워넣습니다.

예시:

```env
OPENAI_API_KEY=
OPENAI_MODEL=gpt-5.4
OPENAI_ENABLE_IMAGE_GENERATION=false
# OPENAI_IMAGE_MODEL=gpt-image-1
# OPENAI_BASE_URL=https://api.openai.com/v1
OPENAI_TIMEOUT_SECONDS=300
CONTINUE_ON_SESSION_ERROR=true
```

설명:

- `OPENAI_API_KEY`: OpenAI API 키
- `OPENAI_MODEL`: 슬라이드 구조 생성에 사용할 모델
- `OPENAI_ENABLE_IMAGE_GENERATION`: 이미지 생성 사용 여부
- `OPENAI_IMAGE_MODEL`: 이미지 생성 모델
- `OPENAI_BASE_URL`: 필요 시 사용자 지정 API 엔드포인트
- `OPENAI_TIMEOUT_SECONDS`: API 대기 시간
- `CONTINUE_ON_SESSION_ERROR`: 특정 교시 실패 시 다음 교시 계속 진행할지 여부

## 실행 방법
먼저 프로젝트 폴더로 이동합니다.

```powershell
cd D:\dev\linkvalue_llm_lecture_ppt_generator
```

가상환경과 의존성 설치까지 포함한 기본 실행은 아래 스크립트를 사용합니다.

```powershell
.\run.ps1
```

### 전체 교시 실행
전체 교시를 대상으로 LLM을 호출하고, 교시별 PPT와 최종 병합 PPT를 생성합니다.

```powershell
.\run.ps1
```

### Mock 실행
LLM을 호출하지 않고 내장된 샘플 슬라이드 데이터로 PPT 생성 흐름만 테스트합니다.

```powershell
.\run.ps1 -Mock
```

참고:

- `-Mock`은 실제 교시별 LLM 생성이 아닙니다.
- 구조 테스트, 레이아웃 확인, 이미지 재사용 확인용입니다.

### Google Slides 호환성
현재 기본 생성 방식은 Google Slides import 호환성을 우선하는 단순 표준 PPTX 경로를 사용합니다.

즉, 별도 옵션 없이도 PowerPoint와 Google Slides에서 모두 열릴 수 있도록 아래 요소를 보수적으로 처리합니다.

- 저수준 XML 기반 불릿 대신 일반 텍스트 불릿 사용
- 표 border 커스텀 XML 생략
- 발표자 노트 생략

필요하면 기존 실행 명령을 그대로 사용하면 됩니다.

```powershell
.\run.ps1
.\run.ps1 -lecture 1
.\run.ps1 -Mock -lecture 1 -page 4
```

기존 `-GoogleSafe` 옵션도 그대로 남아 있지만, 현재는 기본 동작과 동일합니다.

```powershell
.\run.ps1 -GoogleSafe
.\run.ps1 -lecture 1 -GoogleSafe
.\run.ps1 -Mock -lecture 1 -page 4 -GoogleSafe
```

### 특정 교시만 실행
```powershell
.\run.ps1 -lecture 1
.\run.ps1 -lecture 1,3
```

동작:

- `-lecture 1`: 1교시만 생성
- `-lecture 1,3`: 1교시와 3교시만 생성

### 특정 페이지(슬라이드)만 실행
`-page` 옵션은 반드시 `-lecture`와 함께 사용해야 합니다.

```powershell
.\run.ps1 -lecture 1 -page 4
.\run.ps1 -lecture 1 -page 1,3
.\run.ps1 -Mock -lecture 1 -page 4
```

동작:

- `-lecture 1 -page 4`: 1교시의 4번 슬라이드만 생성
- `-lecture 1 -page 1,3`: 1교시의 1번, 3번 슬라이드만 생성
- 페이지 단위 실행 시 결과 파일은 예를 들어 `lecture1_page4.pptx`처럼 별도 이름으로 저장됩니다.
- 같은 교시의 전체 `*_slide_plan.json`이 이미 있으면, 페이지 실행 시 해당 계획을 재사용해 LLM을 다시 호출하지 않습니다.
- 같은 교시의 전체 `*_slide_plan.json`, `*_llm_raw_response.txt`가 있으면 기본 실행에서도 해당 캐시를 우선 재사용합니다.
- 처음부터 다시 생성하려면 해당 교시의 캐시 파일을 직접 삭제하면 됩니다.

### 템플릿 분석만 실행
```powershell
.\run.ps1 -AnalyzeOnly
```

## 출력 파일
출력 파일은 `output/` 폴더에 저장됩니다.

예시:

- `template_analysis.json`
- `01_교시명_prompt.txt`
- `01_교시명_llm_raw_response.txt`
- `01_교시명_slide_plan.json`
- `01_교시명.pptx`
- `lecture1_page4.pptx`
- `final_merged_curriculum.pptx`

교시별 출력 예시:

- `01_생성형_AI_개요_및_업무_변화_이해.pptx`
- `02_업무_자동화의_본질_IPO_구조_재정의.pptx`

이미지 생성이 켜져 있으면 교시별 이미지도 저장됩니다.

- `output/images/<교시별_폴더>/slide_01.png`

## 이미지 캐시 동작
- 이미지 생성 전에 해당 교시 폴더 안에 같은 슬라이드 번호의 이미지가 있으면 재사용합니다.
- 이미지 생성이 꺼져 있어도 기존 이미지 파일이 있으면 PPT에 삽입합니다.
- 기존 이미지가 없으면 템플릿 이미지 또는 빈 영역 상태로 유지될 수 있습니다.

## 슬라이드 규칙
현재 코드 기준 기본 규칙은 아래와 같습니다.

- 본문 글꼴: `맑은 고딕`
- 슬라이드 상단 제목과 본문 제목: Bold, `20pt`
- 본문 줄 간격: 기본 `1.5`
- 본문은 불릿 중심으로 구성
- 예시가 필요하면 `ex>` 형식 사용 가능
- 필요 시 하위 불릿, 표, 간단한 다이어그램 사용
- `왜 이 내용이 필요한가`, `구체적 예시`, `이미지 프롬프트` 등은 발표자 노트로 보낼 수 있음

## 오류 및 로그
오류 발생 시 아래 파일들을 확인합니다.

- 교시별 오류: `output/<prefix>_error.txt`
- 전체 오류 로그: `output/error.log`
- 병합 오류: `output/merge_error.log`

참고:

- 콘솔 로그는 `[교시 x/y]`, `[슬라이드 x/y]` 형식으로 진행 상태를 출력합니다.
- `Ctrl+C`로 실행 중단 시 현재 작업이 중단 로그와 함께 종료됩니다.

## Git 제외 권장
버전 관리에서 제외하는 것을 권장하는 항목:

- `.env`
- `output/`
- `template/`
- `prompts/`
- `.venv/`
- `.venv_test/`
- `.tmp/`
- `__pycache__/`

이미 Git에 올라간 파일은 `.gitignore`만 추가해서는 제외되지 않으므로, 필요하면 `git rm --cached`로 추적 해제해야 합니다.
