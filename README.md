# LLM Lecture PPT Generator

강의 커리큘럼과 프롬프트를 바탕으로 교시별 강의용 PPT를 자동 생성하는 프로젝트입니다.  
출력은 표준 `.pptx` 파일이며, 템플릿 PPT의 레이아웃을 참고해 슬라이드를 만듭니다.

## 개요

기본 흐름은 다음과 같습니다.

1. `template/` 폴더의 템플릿 PPT를 읽습니다.
2. `prompts/curriculum.txt`에서 교시별 주제와 핵심 내용을 읽습니다.
3. `prompts/lecture_prompt.txt`의 공통 프롬프트를 교시별 내용으로 채웁니다.
4. LLM이 슬라이드 구조 JSON을 생성합니다.
5. JSON을 바탕으로 교시별 PPT를 생성합니다.
6. 전체 교시를 실행하면 마지막에 병합 PPT도 함께 생성합니다.

## 프로젝트 구조

```text
linkvalue_llm_lecture_ppt_generator/
├─ app.py
├─ run.ps1
├─ activate_venv.ps1
├─ merge_ppt.ps1
├─ requirements.txt
├─ .env
├─ .env.example
├─ .gitignore
├─ template/
├─ prompts/
└─ output/
```

## template 작성 방법

`template/` 폴더에는 기준이 되는 `.pptx` 템플릿 파일을 넣습니다.

현재 프로젝트는 비교적 단순한 템플릿을 기준으로 동작합니다.

- 상단에 강의 제목 영역이 있음
- 본문은 왼쪽 상단 불릿 텍스트 영역을 사용함
- 오른쪽 하단은 이미지 영역으로 사용함
- 필요 시 왼쪽 하단에 표 또는 다이어그램이 들어갈 수 있음

권장 템플릿 조건:

- 본문 슬라이드 레이아웃이 너무 복잡하지 않을 것
- 본문 텍스트와 이미지 영역이 명확히 분리되어 있을 것
- 글꼴과 기본 테마가 템플릿에 미리 적용되어 있을 것

예시:

```text
template/templates.pptx
```

## prompts 작성 방법

`prompts/` 폴더에는 생성 규칙과 커리큘럼을 넣습니다.

- `lecture_prompt.txt`
- `curriculum.txt`

### lecture_prompt.txt

모든 교시에 공통으로 적용되는 프롬프트 템플릿입니다.  
슬라이드 수, 문장 톤, 본문 구조, 표/다이어그램 판단 기준 등을 정의합니다.

여기에는 보통 아래와 같은 내용이 들어갑니다.

- 강의 목적
- 슬라이드 수
- 장표 작성 규칙
- 출력 JSON 구조
- 불릿/표/다이어그램 판단 기준
- 이미지 프롬프트 작성 규칙

### curriculum.txt

전체 강의 커리큘럼 파일입니다.  
교시 수만큼 PPT가 생성되며, 각 교시의 제목과 핵심 내용이 여기서 결정됩니다.

예시 형식:

```text
[강의 개요]
- 생성형 AI 활용과 업무 자동화 흐름을 이해하는 강의

1교시. 생성형 AI 개요 및 업무 변화 이해
- 생성형 AI란 무엇인가
- 기존 자동화와 생성형 AI 차이
- 기업 활용 사례
- 한계와 주의점

2교시. 업무 자동화의 본질 IPO 구조 재정의
- Input, Process, Output 구조 이해
- 자동화 대상 선정 기준
- 실무 적용 예시
```

작성 규칙:

- 반드시 `1교시. 제목` 형식을 사용합니다.
- 교시 아래 줄들은 해당 교시의 핵심 내용으로 사용됩니다.
- 교시 수만큼 PPT가 생성됩니다.

## 환경 변수

`.env.example`을 복사해서 `.env`를 만든 뒤 값을 채웁니다.

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
- `OPENAI_MODEL`: 슬라이드 구조 생성을 위한 모델
- `OPENAI_ENABLE_IMAGE_GENERATION`: 이미지 생성 사용 여부
- `OPENAI_IMAGE_MODEL`: 이미지 생성 모델
- `OPENAI_BASE_URL`: 별도 API 엔드포인트 사용 시 설정
- `OPENAI_TIMEOUT_SECONDS`: API 타임아웃
- `CONTINUE_ON_SESSION_ERROR`: 특정 교시 실패 시 다음 교시 계속 진행할지 여부

## 실행 방법

프로젝트 폴더에서 실행합니다.

```powershell
cd D:\dev\linkvalue_llm_lecture_ppt_generator
.\run.ps1
```

### 전체 생성

```powershell
.\run.ps1
```

전체 교시를 대상으로 LLM을 호출하고, 교시별 PPT와 최종 병합 PPT를 생성합니다.

### Mock 테스트

```powershell
.\run.ps1 -Mock
```

OpenAI API를 호출하지 않고 내장 mock 데이터로 구조와 렌더링만 테스트합니다.

### 특정 교시만 실행

```powershell
.\run.ps1 -lecture 1
.\run.ps1 -lecture 1,3
.\run.ps1 -Mock -lecture 1,3
```

- `-lecture 1`: 1교시만 생성
- `-lecture 1,3`: 1교시와 3교시만 생성
- `-lecture` 사용 시 최종 병합은 수행하지 않음

### 특정 페이지 테스트

`-page` 옵션은 반드시 `-lecture`와 함께 사용해야 합니다.

```powershell
.\run.ps1 -Mock -lecture 1 -page 4
.\run.ps1 -lecture 1 -page 1,3
```

- `-lecture 1 -page 4`: 1교시의 4번 슬라이드만 생성
- `-lecture 1 -page 1,3`: 1교시의 1번, 3번 슬라이드만 생성

페이지 테스트 결과 파일은 별도 이름으로 저장됩니다.

- `lecture1_page4.pptx`
- `lecture1_page1_3.pptx`

### 템플릿 분석만 실행

```powershell
.\run.ps1 -AnalyzeOnly
```

### 가상환경만 활성화

```powershell
.\activate_venv.ps1
```

## 출력 결과

주요 출력 파일:

- `output/template_analysis.json`
- `output/error.log`
- `output/merge_error.log`
- `output/final_merged_curriculum.pptx`

교시별 출력 예시:

- `output/01_생성형_AI_개요_및_업무_변화_이해_prompt.txt`
- `output/01_생성형_AI_개요_및_업무_변화_이해_llm_raw_response.txt`
- `output/01_생성형_AI_개요_및_업무_변화_이해_slide_plan.json`
- `output/01_생성형_AI_개요_및_업무_변화_이해.pptx`

페이지 테스트 출력 예시:

- `output/lecture1_page4.pptx`
- `output/lecture1_page4_slide_plan.json`

이미지 출력 예시:

- `output/images/01_생성형_AI_개요_및_업무_변화_이해/slide_01.png`
- `output/images/lecture1_page4/slide_04.png`

## 슬라이드 규칙

현재 프로젝트의 기본 생성 규칙은 다음과 같습니다.

- 본문은 발표용 문장 중심
- 필요 시 `ex>` 예시 문장 포함
- 계층 구조 설명은 하위 불릿 지원
- 비교 구조는 표로 생성 가능
- 프로세스/비교/계층/순환/관계형 다이어그램 지원
- 본문 타이틀은 `Bold`, `20pt`
- 본문은 `맑은 고딕` 기준

## 문제 해결

오류 로그 위치:

- 교시별 오류: `output/<prefix>_error.txt`
- 전체 오류 로그: `output/error.log`
- 병합 오류: `output/merge_error.log`

참고:

- PPT 병합은 Windows PowerPoint COM을 사용합니다.
- 따라서 병합 기능은 Microsoft PowerPoint가 설치된 Windows 환경에서 동작합니다.
- PowerPoint가 열려 있거나 `output/` 폴더가 잠겨 있으면 생성/저장이 실패할 수 있습니다.

## Git 제외 대상

로컬 전용 파일은 Git에 포함하지 않는 것을 권장합니다.

- `.env`
- `.venv/`
- `.venv_test/`
- `.tmp/`
- `output/`
- `template/`
- `prompts/`

## 요구 사항

- Windows
- Python 3.12 이상
- Microsoft PowerPoint 설치
- OpenAI API 키
