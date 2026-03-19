# KR Chart Rater - 사용 가이드

## 개요

매일 자동으로 관심 종목의 차트를 AI가 분석하고, 결과를 Notion에 저장하는 시스템입니다.

- 매일 **평일 오후 4시** (KRX 정규장 마감 30분 후)에 자동 실행
- Notion에 등록한 종목을 읽어와서 일봉 차트를 AI에게 분석시킴
- 매력도 A-1, A-2로 선정된 종목만 Notion 결과 DB에 저장
- 원하면 수동으로도 즉시 실행 가능

---

## 전체 구조 (한눈에 보기)

```
[Notion DB1: 종목 리스트]      ← 여기서 종목 추가/삭제
        ↓ (매일 자동 읽기)
[GitHub Actions: 차트 생성 + AI 분석]
        ↓ (결과 자동 저장)
[Notion DB2: 분석 결과]        ← 여기서 결과 확인
```

---

## 1. 종목 관리 (Notion DB1)

![image.png](attachment:f4dfe3ff-8153-4af5-9087-17b38c90b0b5:image.png)

### 종목 추가하기

DB1에서 새 행을 추가하고:

1. **종목명** 칸에 정확한 종목명 입력 (예: `삼성전자`, `SK하이닉스`)
2. **리스트** 칸에 태그 선택 (예: `BIO`, `반도체`, `전체 종목`)

### 종목 삭제하기

해당 행을 삭제하면 됩니다. 다음 실행부터 반영됩니다.

### 리스트 태그란?

종목을 그룹으로 묶는 태그입니다. 한 종목이 여러 리스트에 속할 수 있습니다.

| 종목명 | 리스트 |
| --- | --- |
| 삼천당제약 | `전체 종목`, `BIO` |
| SK하이닉스 | `전체 종목`, `반도체` |
| 파두 | `전체 종목` |

---

## 2. 분석 결과 확인 (Notion DB2)

자동 실행 후 DB2에 새 행이 생깁니다.

![image.png](attachment:9d550c8e-7fe5-4103-9e53-1811eb80b3e6:image.png)

### 각 열의 의미

- **A-1**: 속도형 매력 종목 (5/20일선 이격 확대 가능). 추천순으로 나열
- **A-2**: 완만추세 지속형 종목 (20일선 중심 흐름). 추천순으로 나열
- **리스트**: 어떤 리스트를 분석했는지
- **종목수**: 총 분석한 종목 수
- **선정**: A-1 + A-2로 선정된 종목 수
- **비용**: AI API 사용료 (달러)

### 매력도 등급 설명

| 등급 | 의미 | 결과에 포함? |
| --- | --- | --- |
| A-1 | 속도형 매력 (강한 상승 추세) | O |
| A-2 | 완만추세 지속형 (안정적 상승) | O |
| B | 보유 가능 (신규 진입 조건부) | X |
| C | 관망 | X |
| D | 리스크 구간 | X |

---

## 3. 수동으로 실행하기

원할 때 즉시 분석을 돌릴 수 있습니다.

GitHub 저장소 접속: `github.com/NP-INSIGHT/KR_Chart_Rater`

[GitHub - NP-INSIGHT/KR_Chart_Rater](https://github.com/NP-INSIGHT/KR_Chart_Rater/)

1. 상단 **Actions** 탭 클릭

![image.png](attachment:0da8458f-e870-4048-a635-6735f67bb7e6:image.png)

1. 왼쪽 목록에서 **Notion Daily Analysis** 클릭

![image.png](attachment:115e1116-1c8f-494f-b8b2-3dfa3ebb174b:image.png)

1. 오른쪽 **Run workflow** 버튼 클릭

![image.png](attachment:1dd674f2-506f-4226-bc97-613267531f6f:image.png)

1. 옵션 선택:
    - **Provider**: `claude` 또는 `gemini` (AI 모델 선택)
    - **Filter by list name**: 특정 리스트만 분석할 경우 입력 (예: `BIO`). 비워두면 `.github/workflows/notion_analysis.yml`  에 설정된 기본 설정으로 실행
2. 초록색 **Run workflow** 클릭

### 실행 상태 확인

- Actions 탭에서 실행 중인 항목 클릭
- `analyze` 단계 클릭하면 실시간 로그 확인 가능
- 초록색 체크 = 성공, 빨간색 X = 실패

---

## 4. 여러 리스트를 각각 분석하고 싶을 때

기본 설정은 전체 종목을 한 번에 분석합니다. 리스트별로 따로 결과를 만들고 싶으면:

1. GitHub에서 `.github/workflows/notion_analysis.yml` 파일을 엽니다
2. 연필 아이콘 (Edit) 클릭
3. `run:` 부분을 아래처럼 수정합니다:

```yaml
run: |
  python run_notion.py --provider claude --list "BIO"
  python run_notion.py --provider claude --list "반도체"
  python run_notion.py --provider gemini --list "테마주"
```

1. **Commit changes** 클릭

이렇게 하면 DB2에 리스트별로 별도 행이 생깁니다:

| 날짜 | 리스트 | A-1 | A-2 |
| --- | --- | --- | --- |
| 03/19 | BIO | 한올바이오파마(009420) | ... |
| 03/19 | 반도체 | SK하이닉스(000660) | ... |

---

## 5. 자동 실행 시간 변경하기

1. GitHub에서 `.github/workflows/notion_analysis.yml` 파일을 엽니다
2. 연필 아이콘 클릭
3. `cron:` 줄을 수정합니다

### 시간 계산법

한국시간(KST)에서 **9를 빼면** UTC 시간입니다.

```
cron: '분 시 * * 요일'
```

| 원하는 시간 | cron 값 | 설명 |
| --- | --- | --- |
| 평일 16:00 | `0 7 * * 1-5` | 현재 설정 |
| 평일 15:40 | `40 6 * * 1-5` | 장 마감 직후 |
| 매일 18:00 | `0 9 * * *` | 주말 포함 매일 |
| 평일 09:00 | `0 0 * * 1-5` | 장 시작 전 |
1. **Commit changes** 클릭

---

## 6. AI 모델(Provider) 변경하기

### Claude vs Gemini 차이

| 항목 | Claude | Gemini |
| --- | --- | --- |
| 분석 품질 | 높음 | 보통 |
| 비용 (16종목 기준) | ~$0.38 (약 530원) | ~$0.02 (약 28원) |
| 속도 | 보통 | 빠름 |

### 변경 방법

- **수동 실행 시**: Run workflow에서 Provider를 선택하면 됨
- **자동 실행 기본값 변경**: GitHub에서 `config.txt` 파일 열고 아래 줄 수정:

```
LLM_PROVIDER=claude
```

`claude` 또는 `gemini` 중 하나를 입력합니다.

---

## 7. 자동 실행 끄기/켜기

1. GitHub → **Actions** 탭
2. 왼쪽에서 **Notion Daily Analysis** 클릭
3. 오른쪽 `...` (점 세 개) 클릭
4. **Disable workflow** 클릭 → 자동 실행 중지
5. 다시 켜려면 같은 위치에서 **Enable workflow** 클릭

---

## 8. 비용 안내

| 항목 | 비용 |
| --- | --- |
| GitHub Actions | 무료 (공개 저장소) |
| Notion | 무료 |
| Claude API | 종목당 약 $0.024 |
| Gemini API | 종목당 약 $0.001 |

> 예시: 매일 20종목을 Claude로 분석하면 → 하루 약 $0.48 (약 670원), 한 달 약 $10 (약 14,000원)
> 

---

## 9. 문제가 생겼을 때

### 실행이 실패했어요

1. Actions 탭에서 실패한 실행 클릭
2. `analyze` 단계 클릭해서 로그 확인
3. 흔한 원인:
    - **Notion API 오류**: Notion 통합이 페이지에 연결되어 있는지 확인
    - **API 키 만료**: Settings → Secrets에서 키 업데이트
    - **종목명 오타**: Notion DB1에서 종목명이 정확한지 확인

### 특정 종목만 분석 실패해요

- 종목명이 정확한지 확인 (예: `SK하이닉스` O, `sk하이닉스` X)
- 최근 상장한 종목은 데이터가 부족할 수 있음

### 

---

## 10. 사용한 스택:

- GUI 버전: customtkinter
- 서버 활용: Github Actions (코드를 공개하면 무료 사용)
- Github과의 코드 연동: Git (Claude Code 사용 시에 설치되어 있음)
- LLM: Claude, Sonnet 4.6 API
- 차트 데이터: 야후파이낸스(Python 라이브러리:yfinance), 최근 1년간 일봉+거래량+이평선
- 야후파이낸스 티커↔ 이름 매칭: DB화 이후, DB에서 매칭
