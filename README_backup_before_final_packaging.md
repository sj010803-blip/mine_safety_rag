# MineSafe AI

광산 안전 지침 및 중대재해처벌법 대응을 위한 LLM-RAG 기반 가상 안전관리자 개발

광산 안전 지침과 관련 법령을 검색하고, 검색 근거를 바탕으로 현장 안전조치와 KRAS식 위험성평가 기록 초안을 제공하는 Streamlit 애플리케이션입니다.

## 프로젝트 개요

이 프로젝트는 ChromaDB Vector DB와 Gemini API를 결합한 광산 안전 분야 RAG(Retrieval-Augmented Generation) 시스템입니다. 외부 LLM API가 불안정한 상황에서도 검색 근거 기반 안정 답변을 제공하며, 질문 시나리오 평가와 비교 실험 결과를 함께 관리합니다.

> 본 시스템은 안전관리 업무를 지원하기 위한 연구용 도구입니다. 최종 법령 해석, 위험등급 결정 및 작업 재개 판단은 현장 안전관리자와 관계 전문가가 확인해야 합니다.

## 주요 기능

- 광산 안전 문서 기반 ChromaDB 검색
- 문서명, `chunk_id`, 거리값 및 메타데이터 표시
- 안정 모드, Gemini 모드, 하이브리드 모드
- Gemini 실패 시 검색 근거 기반 답변 제공
- 상황 유형별 즉시 판단과 현장 조치 체크리스트
- KRAS식 위험성평가 기록 초안 생성
- Q001~Q110 질문 시나리오 테스트 및 평가 현황 관리
- 4개 평가 기준, 100점 만점 자동평가 결과 표시
- ChatGPT, Gemini, 도메인 특화 RAG 비교 실험 결과 제공

## 시스템 흐름

```text
사용자 질문
  -> 질문 유형 분류
  -> ChromaDB 관련 문서 검색
  -> 검색 결과 후처리 및 재정렬
  -> 검색 근거 기반 안전 답변 생성
  -> 선택 모드에 따라 Gemini 보조 답변 생성
  -> 근거 문서와 KRAS식 위험성평가 초안 표시
  -> 시나리오 평가 및 결과 저장
```

## 사용 데이터

- `02_질문시나리오/question_scenarios_110.tsv`: Q001~Q110 평가 질문
- `08_chunks/chunks.jsonl`: 전처리된 검색 단위 문서
- `10_vector_db/`: ChromaDB Vector DB
- `09_answer_tests/`: 자동평가 결과와 교수자용 결과 파일
- `12_compare_experiment/`: 범용 LLM과 RAG 비교 실험 파일

주요 근거 범위에는 광산안전법, 광산안전기술기준, 산업안전보건법, 중대재해처벌법, 위험성평가 안내서 및 안전보건관리체계 구축 가이드북 등이 포함됩니다.

## 평가 결과

Q001~Q110은 다음 4개 기준을 각 25점, 총 100점으로 통합 평가합니다.

1. 검색 적합성
2. 근거 기반성
3. 안전·법령 판단 정확성
4. 실무성

20개 동일 질문을 이용한 규칙 기반 1차 비교평가 평균은 다음과 같습니다.

| 비교 대상 | 총점 평균 |
|---|---:|
| ChatGPT | 63.00 |
| Gemini | 74.50 |
| 내 사이트 RAG | 94.90 |

비교 결과는 답변의 키워드, 출처 표현, 안전조치 및 구조화 요소를 분석한 자동 1차 평가입니다. 최종 평가는 대표 문항의 수동 검토가 필요합니다.

## 실행 방법

```powershell
cd C:\Users\USER\Desktop\mine_safety_rag
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

`python` 명령을 찾지 못하는 환경에서는 설치된 Python 실행 파일의 전체 경로 또는 Windows의 `py` 런처를 사용하세요.

## 환경변수 설정

프로젝트 루트에 `.env` 파일을 만들고 다음 형식으로 설정합니다.

```dotenv
GEMINI_API_KEY=your_api_key_here
```

Streamlit Secrets를 사용하는 경우 `.streamlit/secrets.toml`에 동일한 키를 설정할 수 있습니다.

```toml
GEMINI_API_KEY = "your_api_key_here"
```

`.env`와 `.streamlit/secrets.toml`은 `.gitignore`에 포함되어야 하며 GitHub에 업로드하면 안 됩니다.

## 주요 폴더 구조

```text
mine_safety_rag/
├─ app.py
├─ requirements.txt
├─ README.md
├─ 02_질문시나리오/
├─ 04_documents/
├─ 05_metadata/
├─ 06_processed_texts/
├─ 08_chunks/
├─ 09_answer_tests/
├─ 10_vector_db/
├─ 12_compare_experiment/
└─ scripts/
```

## 주의사항

- API 키를 코드, README, 보고서 또는 화면 캡처에 포함하지 마세요.
- `.env`, `secrets.toml`, 백업 파일, 캐시 및 임시 파일을 GitHub에 올리지 마세요.
- 검색 답변은 원문 법령과 현장 조건을 대체하지 않습니다.
- 문서에 명확히 없는 조문, 수치 또는 처벌 기준을 단정하지 마세요.
- Gemini API 오류가 발생해도 안정 모드의 검색 근거 기반 답변을 사용할 수 있습니다.
- GitHub 공유 전 `python check_github_share_ready.py`로 최종 점검하세요.
