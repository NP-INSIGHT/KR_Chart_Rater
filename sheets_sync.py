"""
KR Chart Rater - Google Spreadsheet 자동 생성 모듈

분석 완료 후 수익률 추적용 스프레드시트를 Google Sheets에 생성한다.

인증: secrets/google_service_account.json (서비스 계정 키 JSON)
필요 API: Google Sheets API, Google Drive API (Google Cloud Console에서 활성화)
"""

import logging
from datetime import datetime, timedelta
from pathlib import Path

logger = logging.getLogger(__name__)

# 서비스 계정 키 경로
_BASE_DIR = Path(__file__).parent
_CREDS_PATH = _BASE_DIR / "secrets" / "google_service_account.json"

# 분석 시작일 기준 사전 생성할 거래일 수 (weekday 기준, 한국 공휴일 미반영)
_FUTURE_TRADING_DAYS = 5


def _get_next_weekdays(start_date, count):
    """start_date 다음 평일 count개를 반환 (토/일 제외)."""
    result = []
    d = start_date
    while len(result) < count:
        d = d + timedelta(days=1)
        if d.weekday() < 5:  # 0=월 ~ 4=금
            result.append(d)
    return result


def _make_ticker_str(result):
    """종목 결과 dict → 'name (MARKET:CODE)' 형식 문자열."""
    name = result.get("ticker_name", "")
    code = result.get("code", "")
    market = result.get("market", "")

    if market == "KOSPI":
        ticker = f"KRX:{code}"
    elif market == "KOSDAQ":
        ticker = f"KOSDAQ:{code}"
    else:
        ticker = code

    return f"{name} ({ticker})"


def create_performance_spreadsheet(a_results, analysis_date, log=None):
    """
    분석 결과로 Google Spreadsheet를 생성하고 URL을 반환한다.

    Parameters
    ----------
    a_results : list
        A-1/A-2 분석 결과 dict 리스트
    analysis_date : datetime
        분석 기준일 (KST)
    log : callable, optional
        로그 출력 콜백

    Returns
    -------
    str or None
        생성된 스프레드시트 URL. 인증 파일 없거나 오류 시 None 반환.
    """
    def _log(msg):
        logger.info(msg)
        if log:
            log(msg)

    if not _CREDS_PATH.exists():
        _log("[Spreadsheet] secrets/google_service_account.json 없음 - 스프레드시트 생성 스킵")
        return None

    # JSON 유효성 사전 검증
    import json as _json
    try:
        _json.loads(_CREDS_PATH.read_text(encoding="utf-8"))
    except Exception:
        _log("[Spreadsheet] google_service_account.json이 유효한 JSON이 아님 - 시크릿 값을 확인하세요")
        return None

    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        _log("[Spreadsheet] gspread / google-auth 패키지 없음 - pip install gspread google-auth")
        return None

    # ── 인증 ──────────────────────────────────────────────────────
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(str(_CREDS_PATH), scopes=scopes)
    gc = gspread.authorize(creds)

    # ── 스프레드시트 생성 ─────────────────────────────────────────
    # GOOGLE_DRIVE_FOLDER_ID가 있으면 해당 폴더에 생성 (Drive 할당량 문제 해결)
    import os
    folder_id = os.environ.get("GOOGLE_DRIVE_FOLDER_ID") or ""

    # 환경변수 없으면 engine config에서 조회 시도
    if not folder_id:
        try:
            import engine
            folder_id = engine.CONFIG.get("GOOGLE_DRIVE_FOLDER_ID", "")
        except Exception:
            pass

    date_label = analysis_date.strftime("%Y-%m-%d")
    title = f"KR Chart Rater - {date_label}"
    sh = gc.create(title, folder_id=folder_id if folder_id else None)
    ws = sh.get_worksheet(0)
    ws.update_title("수익률 추적")

    # ── 종목 행 빌드 ──────────────────────────────────────────────
    a1_results = [r for r in a_results if r.get("grade") == "A-1"]
    a2_results = [r for r in a_results if r.get("grade") == "A-2"]

    # 미래 평일 날짜 (D, F, H ... 열용)
    future_days = _get_next_weekdays(analysis_date, _FUTURE_TRADING_DAYS)

    # 헤더 행 (Row 1): 그룹 | 종목 | 분석일 | 거래일1 | 수익률 | 거래일2 | 수익률 | ...
    start_date_str = analysis_date.strftime("%-m/%d") if hasattr(analysis_date, "strftime") else analysis_date.strftime("%m/%d").lstrip("0")
    # Windows 호환 날짜 포맷 (%-m/%d 는 Linux/Mac 전용)
    try:
        start_date_str = analysis_date.strftime("%-m/%d")
    except ValueError:
        # Windows: %#m/%d
        start_date_str = analysis_date.strftime("%#m/%d")

    header_row = ["그룹", "종목", start_date_str]
    for fd in future_days:
        try:
            date_str = fd.strftime("%-m/%d")
        except ValueError:
            date_str = fd.strftime("%#m/%d")
        header_row.append(date_str)
        header_row.append("수익률")

    # 종목 행 데이터 수집
    stock_data_rows = []  # (A_value, B_value)
    all_groups = [("A-1", a1_results), ("A-2", a2_results)]
    for grade_label, grade_results in all_groups:
        if not grade_results:
            continue
        for idx, r in enumerate(grade_results):
            a_val = f"{grade_label} ({len(grade_results)}종목)" if idx == 0 else ""
            b_val = _make_ticker_str(r)
            stock_data_rows.append((a_val, b_val))

    # ── 시트 쓰기 ────────────────────────────────────────────────
    all_rows = [header_row]
    for a_val, b_val in stock_data_rows:
        row = [a_val, b_val] + [""] * (len(header_row) - 2)
        all_rows.append(row)

    ws.update(values=all_rows, range_name="A1")

    # ── 수식 삽입 ────────────────────────────────────────────────
    # C열(분석일 종가), D/F/H... 열(거래일 종가), E/G/I... 열(수익률)
    # C=3, D=4, E=5, F=6, G=7, H=8, I=9, ...
    # 열 매핑: 분석일→col 3, 거래일1→col 4, 수익률1→col 5, 거래일2→col 6, ...

    cell_updates = []  # (row, col, value)

    for data_row_idx, _ in enumerate(stock_data_rows):
        sheet_row = data_row_idx + 2  # 헤더가 row 1이므로 data는 row 2부터

        # C열: 분석일 종가
        c_formula = (
            f'=INDEX(GOOGLEFINANCE(REGEXEXTRACT($B{sheet_row},"\\\\((.*)\\\\)"),'
            f'"close",C$1),2,2)'
        )
        cell_updates.append((sheet_row, 3, c_formula))

        # D/E, F/G, H/I ... 열: 거래일 종가 + 수익률
        for i, _ in enumerate(future_days):
            date_col = 4 + i * 2      # D=4, F=6, H=8 ...
            ret_col = date_col + 1    # E=5, G=7, I=9 ...

            date_col_letter = _col_letter(date_col)
            close_formula = (
                f'=INDEX(GOOGLEFINANCE(REGEXEXTRACT($B{sheet_row},"\\\\((.*)\\\\)"),'
                f'"close",{date_col_letter}$1),2,2)'
            )
            cell_updates.append((sheet_row, date_col, close_formula))

            ret_formula = f"=({date_col_letter}{sheet_row}-$C{sheet_row})/$C{sheet_row}"
            cell_updates.append((sheet_row, ret_col, ret_formula))

    # batch update로 수식 삽입
    if cell_updates:
        ws.update_cells(
            [gspread.Cell(r, c, v) for r, c, v in cell_updates],
            value_input_option="USER_ENTERED",
        )

    # ── 수익률 열 서식 적용 ──────────────────────────────────────
    if stock_data_rows:
        _apply_return_formatting(sh, ws, len(stock_data_rows), len(future_days))

    # ── 링크 공유 설정 ───────────────────────────────────────────
    try:
        sh.share("", perm_type="anyone", role="reader")
    except Exception as e:
        _log(f"[Spreadsheet] 링크 공유 설정 실패 (계속 진행): {e}")

    url = sh.url
    _log(f"[Spreadsheet] 생성 완료: {title}")
    return url


def _col_letter(col_idx):
    """1-based 열 인덱스 → 열 문자 (A, B, ..., Z, AA, ...)."""
    result = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _apply_return_formatting(sh, ws, num_stock_rows, num_future_days):
    """
    수익률 열(E, G, I ...) 에 퍼센트 형식 + 볼드 + 조건부 서식 적용.
    음수 → 파란 배경, 양수 → 빨간 배경.
    실패 시 경고 후 스킵.
    """
    try:
        # 수익률 열 인덱스 목록 (1-based)
        ret_col_indices = [5 + i * 2 for i in range(num_future_days)]  # E=5, G=7, I=9 ...

        # 데이터 범위: row 2 ~ (num_stock_rows + 1)
        start_row = 2
        end_row = num_stock_rows + 1

        # 각 수익률 열에 서식 적용
        fmt_requests = []
        for col_idx in ret_col_indices:
            col_ltr = _col_letter(col_idx)
            range_notation = f"{col_ltr}{start_row}:{col_ltr}{end_row}"

            # 퍼센트 + 볼드 기본 서식
            ws.format(range_notation, {
                "numberFormat": {"type": "PERCENT", "pattern": "0.00%"},
                "textFormat": {"bold": True},
            })

        # 조건부 서식 (batch_update)
        sheet_id = ws.id  # gspread 6.x: Worksheet.id = sheetId(gid)
        rules = []

        for col_idx in ret_col_indices:
            col_ltr = _col_letter(col_idx)
            grid_range = {
                "sheetId": sheet_id,
                "startRowIndex": start_row - 1,
                "endRowIndex": end_row,
                "startColumnIndex": col_idx - 1,
                "endColumnIndex": col_idx,
            }

            # 음수 → 파란 배경
            rules.append({
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [grid_range],
                        "booleanRule": {
                            "condition": {
                                "type": "NUMBER_LESS",
                                "values": [{"userEnteredValue": "0"}],
                            },
                            "format": {
                                "backgroundColor": {"red": 0.53, "green": 0.73, "blue": 0.98}
                            },
                        },
                    },
                    "index": 0,
                }
            })

            # 양수 → 빨간 배경
            rules.append({
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [grid_range],
                        "booleanRule": {
                            "condition": {
                                "type": "NUMBER_GREATER",
                                "values": [{"userEnteredValue": "0"}],
                            },
                            "format": {
                                "backgroundColor": {"red": 0.98, "green": 0.60, "blue": 0.60}
                            },
                        },
                    },
                    "index": 0,
                }
            })

        if rules:
            sh.batch_update({"requests": rules})

    except Exception as e:
        logger.warning(f"[Spreadsheet] 서식 적용 실패 (스킵): {e}")
