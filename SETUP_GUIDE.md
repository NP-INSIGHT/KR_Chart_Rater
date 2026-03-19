# KR Chart Rater - 무료 자동화 셋업 가이드

매일 자동으로 종목 차트를 AI로 분석하고, 결과를 Notion/이메일로 받는 방법입니다.
GitHub Actions(Public 저장소) + Notion + 이메일을 조합하여 **서버 비용 0원**으로 운영합니다.

> LLM API 비용만 발생합니다 (Gemini Flash 기준 하루 약 0.2달러 / Claude Sonnet 기준 하루 약 15달러, 100종목 기준)

---

## 1단계: GitHub 저장소 만들기

1. GitHub.com에서 **New repository** 클릭
2. Repository name: `KR_Chart_Rater` (또는 원하는 이름)
3. **Public** 선택 (Actions 무제한 무료)
4. **Create repository** 클릭
5. 로컬 코드를 push:
   ```
   cd KR_Chart_Rater
   git init
   git add .
   git commit -m "initial commit"
   git remote add origin https://github.com/내계정/KR_Chart_Rater.git
   git push -u origin main
   ```

---

## 2단계: API 키 발급

### Gemini API (추천 - 저렴)
1. https://aistudio.google.com/apikey 접속
2. **Create API Key** 클릭
3. 키 복사해두기

### Claude API (선택)
1. https://console.anthropic.com/ 접속
2. API Keys → Create Key
3. 키 복사해두기

---

## 3단계: Notion 설정

### 3-1. Notion 통합 만들기
1. https://www.notion.so/my-integrations 접속
2. **새 통합 만들기** 클릭
3. 이름: `KR Chart Rater`
4. **제출** → 토큰 복사해두기 (`ntn_` 또는 `secret_`으로 시작)

### 3-2. 한 페이지에 데이터베이스 2개 만들기

하나의 Notion 페이지에 아래 두 데이터베이스를 만듭니다:

**DB 1: 종목 리스트 (Watchlist)**
1. 페이지에서 `/database` 입력 → **데이터베이스 - 인라인** 선택
2. 속성 설정:
   - `종목명` (제목) — 이미 있음
   - `상태` (선택) — 옵션: `활성`, `제외` (선택사항)
3. 종목명 열에 분석할 종목 입력 (예: 삼성전자, SK하이닉스, ...)

**DB 2: 분석 결과 (Report)**
1. 같은 페이지에서 다시 `/database` 입력 → **데이터베이스 - 인라인** 선택
2. 속성 설정:
   - `날짜` (제목) — 이미 있음
   - `분석일` (날짜)
   - `종목수` (숫자)
   - `선정` (숫자) — A-1/A-2 선정 종목 수
   - `비용` (숫자)

### 3-3. 통합 연결하기
페이지 오른쪽 상단 `···` 클릭 → **연결** → `KR Chart Rater` 선택
(한 번만 연결하면 페이지 내 모든 DB에 적용됩니다)

### 3-4. DB ID 확인하기
각 DB를 **전체 페이지로 열기** 후 URL에서 ID를 추출합니다:
```
https://www.notion.so/내이름/abc123def456...?v=...
                            ^^^^^^^^^^^^^^^^
                            이 부분이 DB ID (32자)
```

---

## 4단계: GitHub Secrets 설정

1. GitHub 저장소 → **Settings** → **Secrets and variables** → **Actions**
2. **New repository secret** 으로 아래 추가:

| Name | 값 |
|------|-----|
| `GEMINI_API_KEY` | Gemini API 키 |
| `ANTHROPIC_API_KEY` | Claude API 키 (선택) |
| `NOTION_API_KEY` | Notion 통합 토큰 |

---

## 5단계: config.txt 설정

`config.txt` 파일에서 아래 주석을 해제하고 값을 입력합니다:

```
# Notion 연동
NOTION_WATCHLIST_DB=종목리스트DB의ID
NOTION_REPORT_DB=보고서DB의ID

# GitHub (차트 이미지를 Notion에 삽입하기 위함)
GITHUB_REPO=내계정/KR_Chart_Rater
```

변경 후 commit & push:
```
git add config.txt
git commit -m "config: Notion 설정 추가"
git push
```

---

## 6단계: 워치리스트 설정 (다중 리스트)

GUI에서 또는 CLI에서 여러 워치리스트를 만들 수 있습니다:

### GUI에서
1. 종목 분석 탭 → 좌측 상단 **+** 버튼
2. 리스트 이름 입력 (예: "관심 리스트1")
3. **⚙** 버튼 → Notion DB ID, 이메일 수신자 등 설정

### CLI에서
```bash
python cli.py watchlist --create-list "관심 리스트1"
python cli.py watchlist --list "관심 리스트1" --add 삼성전자 SK하이닉스
python cli.py watchlist --lists  # 전체 리스트 확인
```

---

## 7단계: 이메일 설정 (선택)

Gmail 앱 비밀번호를 사용합니다:

1. Google 계정 → 보안 → 2단계 인증 활성화
2. 앱 비밀번호 생성 (https://myaccount.google.com/apppasswords)
3. `secrets/email_password.txt` 파일에 앱 비밀번호 저장
4. `config.txt`에서 이메일 설정 활성화:
   ```
   EMAIL_SMTP_HOST=smtp.gmail.com
   EMAIL_SMTP_PORT=587
   EMAIL_FROM=내이메일@gmail.com
   EMAIL_TO=받을이메일@gmail.com
   ```

또는 GUI에서 **이메일 설정** 버튼 → 값 입력 → 저장

---

## 8단계: 첫 실행 테스트

### 로컬 테스트
```bash
# 3개 종목으로 간단 테스트 (Notion 저장 안 함)
python run_notion.py --provider gemini --dry-run

# 정상이면 실제 Notion 저장
python run_notion.py --provider gemini
```

### GitHub Actions 수동 실행
1. GitHub 저장소 → **Actions** 탭
2. 좌측에서 **Notion Daily Analysis** 선택
3. **Run workflow** 클릭
4. Provider 선택 → **Run workflow**

---

## 9단계: 자동화 확인

설정이 완료되면 매일 평일 16:00(KST)에 자동으로:

1. Notion 종목 리스트 DB에서 종목명 읽기
2. 각 종목 차트 생성 + AI 분석 (프롬프트 내부 3회 합의)
3. A-1/A-2 선정 종목만 Notion 보고서 DB에 저장 (종목명 + 티커코드)
4. 이메일로 요약 발송 (설정된 경우)

---

## 분석 등급 체계 (v2)

| 등급 | 의미 | Notion 보고서 |
|------|------|--------------|
| **A-1** | 속도형 매력 (5/20일선 이격 확대 가능) | 저장됨 (종목명 + 티커) |
| **A-2** | 완만추세 지속형 매력 (20일선 중심) | 저장됨 (종목명 + 티커) |
| **B** | 보유 가능 (신규 진입 조건부) | 미저장 |
| **C** | 관망 | 미저장 |
| **D** | 리스크 구간 | 미저장 |

- 프롬프트가 내부적으로 3회 독립 분석을 수행하고, 2회 이상 일치하는 결론만 채택합니다
- 각 분석 결과에는 신뢰도 등급(High/Medium/Low)과 합의 횟수가 포함됩니다

---

## 운영 팁

- **종목 추가/삭제**: Notion 종목 리스트 DB에서 행 추가/삭제만 하면 됨
- **종목 제외**: "상태" 속성을 "제외"로 설정하면 분석에서 제외됨
- **비용 절약**: `LLM_PROVIDER=gemini` + `GEMINI_MODEL=gemini-2.5-flash` 설정 (하루 ~0.2달러)
- **정확도 우선**: `LLM_PROVIDER=claude` 설정 (하루 ~15달러, 100종목 기준)
- **수동 실행**: GitHub Actions → Run workflow로 언제든 수동 실행 가능
- **로그 확인**: Actions 실행 기록에서 상세 로그 확인 가능

---

## 주의사항

- GitHub 저장소를 **Public**으로 만들어야 Actions가 무료입니다
- API 키는 반드시 **GitHub Secrets**에 저장하세요 (코드에 직접 입력 금지)
- Notion 통합이 DB 페이지에 **연결**되어 있어야 읽기/쓰기가 가능합니다
- Claude와 Gemini 모두 사용 가능하며, `config.txt`의 `LLM_PROVIDER`로 전환합니다
