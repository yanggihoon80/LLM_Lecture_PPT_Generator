# LLM Lecture PPT Generator

템플릿 PPT를 기준으로 교시별 강의용 PPT를 자동 생성하고, 마지막에 하나의 최종 PPT로 병합하는 프로젝트입니다.

## 개요

이 프로젝트는 아래 순서로 동작합니다.

1. `template/` 폴더의 PPT 템플릿을 분석합니다.
2. `prompts/curriculum.txt`에서 전체 강의 맥락과 교시 목록을 읽습니다.
3. `prompts/lecture_prompt.txt`를 공통 프롬프트 템플릿으로 사용합니다.
4. 각 교시별로 LLM을 호출해 슬라이드 구조 JSON을 생성합니다.
5. 템플릿 레이아웃을 유지한 교시별 PPT를 생성합니다.
6. 모든 교시 PPT를 하나의 최종 PPT로 병합합니다.

## 폴더 구조

```text
linkvalue_ppt_template_slide/
├─ app.py
├─ run.ps1
├─ activate_venv.ps1
├─ merge_ppt.ps1
├─ requirements.txt
├─ .env
├─ .env.example
├─ template/
├─ prompts/
└─ output/
```

## 준비 파일 작성 방법

### template 폴더

`template/`에는 기준이 되는 `.pptx` 템플릿 파일을 넣습니다.

권장 구조:

- 슬라이드 상단에 섹션 타이틀 영역
- 본문 타이틀 영역
- 왼쪽 상단에 불릿형 본문 텍스트 영역
- 오른쪽 하단에 이미지가 들어갈 영역

현재 프로젝트는 템플릿의 전체 디자인, 마스터, 테마, 레이아웃을 유지한 채 본문만 채우는 방식입니다.  
따라서 템플릿은 너무 복잡하기보다, 반복 가능한 강의 본문 슬라이드 구조가 있는 편이 좋습니다.

권장 예시:

- 1개 이상의 본문용 슬라이드 포함
- 본문 텍스트 박스가 최소 3개 이상 존재
  - 섹션 라벨
  - 슬라이드 제목
  - 본문 불릿 영역
- 오른쪽 하단에 대표 이미지 영역 존재

기본 파일명 예시:

```text
template/templates.pptx
```

템플릿을 여러 개 둘 수도 있지만, 기본적으로는 가장 최근 템플릿 또는 지정한 템플릿을 사용합니다.

### prompts 폴더

`prompts/`에는 두 개의 텍스트 파일을 둡니다.

- `lecture_prompt.txt`
- `curriculum.txt`

#### lecture_prompt.txt

교시마다 공통으로 유지할 출력 규칙을 넣는 파일입니다.  
여기에는 PPT 톤, 장수, 강의 스타일, 슬라이드 구성 규칙을 작성합니다.

예시:

```text
너는 교육용 강의 설계자다.
보고서가 아니라 교실에서 사용하는 PPT를 만들어라.

강의 시간: 1시간
슬라이드 수: 20장

[출력 조건]
1. 모든 슬라이드는 하나의 학습 흐름 안에 있어야 한다.
2. 각 장표는 한 개의 핵심 메시지만 가져야 한다.
3. 문장은 발표용으로 바로 읽을 수 있어야 한다.
4. 서술형 에세이 문장은 금지한다.
5. 각 장표에는 핵심 개념 설명 중심의 불릿만 작성한다.
6. 이미지 프롬프트도 함께 제시한다.
```

실제 실행 시에는 이 공통 프롬프트에 각 교시의 주제와 핵심 내용이 자동으로 붙어서 LLM에 전달됩니다.

#### curriculum.txt

전체 강의 맥락과 교시별 주제를 적는 파일입니다.  
앱은 이 파일을 읽고 교시 수만큼 PPT를 생성합니다.

권장 형식:

```text
[과정 개요]
- 생성형 AI 개요부터 업무 자동화 프로젝트 설계까지 단계적으로 학습

1교시. 생성형 AI 개요 및 업무 변화 이해
- 생성형 AI란 무엇인가
- 기존 자동화와 생성형 AI 차이
- 기업 활용 사례
- 한계와 주의점

2교시. 업무 자동화의 본질 IPO 구조 재정의
- Input, Process, Output 구조 이해
- 생성형 AI 적용 지점 찾기
- 자동화 가능 업무 재정의
```

작성 팁:

- 교시 제목은 `1교시. 제목` 형식을 권장합니다.
- 각 교시 아래에는 핵심 내용을 불릿으로 적습니다.
- 교시 수가 최종 생성 PPT 수가 됩니다.
- 교시가 14개면 교시별 PPT 14개와 최종 병합 PPT 1개가 생성됩니다.

## 환경 변수 설정

`.env.example`을 참고해 `.env` 파일을 준비합니다.

```env
OPENAI_API_KEY=your_api_key_here
OPENAI_MODEL=gpt-4.1
OPENAI_ENABLE_IMAGE_GENERATION=false
OPENAI_IMAGE_MODEL=gpt-image-1
OPENAI_TIMEOUT_SECONDS=300
CONTINUE_ON_SESSION_ERROR=true
```

설명:

- `OPENAI_API_KEY`: OpenAI API 키
- `OPENAI_MODEL`: 슬라이드 구조 생성용 모델
- `OPENAI_ENABLE_IMAGE_GENERATION`: 이미지 생성 사용 여부
- `OPENAI_IMAGE_MODEL`: 이미지 생성 모델
- `OPENAI_TIMEOUT_SECONDS`: LLM 응답 타임아웃
- `CONTINUE_ON_SESSION_ERROR`: 특정 교시 실패 시 다음 교시 계속 진행 여부

## 실행 방법

### 기본 실행

```powershell
.\run.ps1
```

실제 LLM을 호출해 교시별 PPT와 최종 병합 PPT를 생성합니다.

### Mock 실행

```powershell
.\run.ps1 -Mock
```

API 호출 없이 샘플 데이터로 전체 흐름만 테스트합니다.

### 템플릿 분석만 실행

```powershell
.\run.ps1 -AnalyzeOnly
```

### 가상환경만 활성화

```powershell
.\activate_venv.ps1
```

### 특정 템플릿 또는 프롬프트 지정

```powershell
.\run.ps1 -TemplateFile ".\template\templates.pptx"
.\run.ps1 -PromptFile ".\prompts\lecture_prompt.txt"
```

## 출력 결과

### 공통 출력

- `output/template_analysis.json`
- `output/final_merged_curriculum.pptx`
- `output/error.log`
- `output/merge_error.log`

### 교시별 출력

각 교시마다 아래 파일이 생성됩니다.

- `output/01_교시명_prompt.txt`
- `output/01_교시명_llm_raw_response.txt`
- `output/01_교시명_slide_plan.json`
- `output/01_교시명.pptx`

### 이미지 출력

이미지 생성이 켜져 있으면 아래처럼 저장됩니다.

- `output/images/01_교시명/slide_01.png`
- `output/images/01_교시명/slide_02.png`

같은 교시 폴더 안에 해당 슬라이드 이미지가 이미 있으면 재생성하지 않고 재사용합니다.

## 슬라이드 규칙

- 템플릿 레이아웃 유지
- 본문은 핵심 개념 설명 불릿만 사용
- 발표자 노트에는 아래 정보 저장
  - 왜 이 내용이 필요한가
  - 구체적 예시 또는 수치
  - 실습 프롬프트 예시
  - 이미지 프롬프트
  - 다음 단계 연결 문장
- 폰트: `맑은 고딕`
- 상단 타이틀과 본문 타이틀: `Bold`, `20pt`
- 줄 간격: `1.5`
- 본문 불릿은 템플릿 기본 불릿 형식으로 통일

## 로그와 에러 처리

실행 중에는 교시와 슬라이드 기준으로 진행 로그가 출력됩니다.

예시:

```text
[교시 3/14] 생성 시작
[교시 3/14] LLM 호출 시작
[5/20] 슬라이드 생성 중: 생성형 AI란 무엇인가
[5/20] 기존 이미지 재사용
[교시 3/14] 완료
[병합] 교시별 PPT 14개를 하나로 병합 중
```

에러 파일:

- 교시별 에러: `output/<교시명>_error.txt`
- 전체 에러 누적: `output/error.log`
- 병합 에러: `output/merge_error.log`

## 병합 방식

최종 병합은 Windows PowerPoint COM을 사용합니다.  
따라서 최종 병합 기능을 사용하려면 PowerPoint가 설치된 Windows 환경이어야 합니다.

병합 결과 파일:

- `output/final_merged_curriculum.pptx`

## Git 커밋 전 참고

현재 `.gitignore`에는 아래 폴더와 파일이 제외되도록 설정하는 것을 권장합니다.

- `.env`
- `.venv/`
- `output/`
- `template/`
- `prompts/`

즉 템플릿 파일과 실제 강의 프롬프트는 로컬에서만 관리하고, 코드와 문서 중심으로 저장소를 운영하는 방식에 맞춰져 있습니다.

## 요구 사항

- Windows
- Python 3.12 이상 권장
- Microsoft PowerPoint 설치
- OpenAI API 키

