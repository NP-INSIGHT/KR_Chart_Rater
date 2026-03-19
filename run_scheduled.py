"""
KR Chart Rater - 스케줄 자동 분석 + 이메일 보고서 발송

사용법:
  python run_scheduled.py                    # 활성 리스트 분석 + 이메일 발송
  python run_scheduled.py --all-lists        # 모든 리스트 순회 분석 + 발송
  python run_scheduled.py --list "관심 리스트1" # 특정 리스트만 분석
  python run_scheduled.py --no-email         # 분석만 (이메일 없이)
  python run_scheduled.py --provider claude  # Claude 사용

cron 설정 예시 (Linux):
  40 15 * * 1-5 cd /path/to/KR_Chart_Rater && python run_scheduled.py

GitHub Actions cron:
  40 6 * * 1-5  (UTC 06:40 = KST 15:40)
"""

import sys
import argparse

import engine


def _run_for_list(list_name, provider, no_email, log):
    """단일 리스트에 대해 분석 + 이메일 발송 수행."""
    # Provider: 리스트별 설정 > 인자 > 글로벌
    list_provider = engine.get_list_config(list_name, "provider") or provider

    tickers = engine.get_active_watchlist(list_name)
    if not tickers:
        log(f"[{list_name}] 활성 종목 없음, 건너뜁니다.")
        return 0

    log(f"\n[{list_name}] {len(tickers)}개 종목 | LLM: {list_provider}")
    log(f"종목: {', '.join(tickers)}")

    # 분석 실행
    try:
        result = engine.run_stock_analysis(
            ticker_names=tickers,
            provider=list_provider,
            log_callback=log,
        )
    except Exception as e:
        log(f"\n[X] [{list_name}] 분석 실패: {e}")
        return 2

    # Word 보고서 생성
    try:
        docx_path = engine.save_results_docx(result, "scheduled")
        log(f"\n[Word] 보고서 저장: {docx_path}")
    except Exception as e:
        log(f"\n[X] Word 보고서 생성 실패: {e}")
        docx_path = None

    # 이메일 발송
    if not no_email:
        to_addr = engine.get_list_config(list_name, "email_to")
        if not to_addr:
            log(f"\n[!] [{list_name}] 이메일 수신자 미설정, 발송 건너뜁니다.")
        else:
            try:
                subject, body = engine.build_email_body(result, list_provider)
                engine.send_report_email(
                    to_addr=to_addr,
                    subject=subject,
                    body_text=body,
                    attachment_path=str(docx_path) if docx_path else None,
                )
                log(f"\n[이메일] [{list_name}] {to_addr}로 발송 완료")
            except Exception as e:
                log(f"\n[X] [{list_name}] 이메일 발송 실패: {e}")

    return 0


def main():
    parser = argparse.ArgumentParser(description="KR Chart Rater 스케줄 실행")
    parser.add_argument(
        "--provider", default=None, choices=["claude", "gemini"],
        help="LLM provider (기본: config.txt의 SCHEDULED_PROVIDER 또는 LLM_PROVIDER)",
    )
    parser.add_argument(
        "--no-email", action="store_true",
        help="이메일 발송 없이 분석만 수행",
    )
    parser.add_argument(
        "--list", default=None,
        help="분석할 워치리스트 이름 (기본: 활성 리스트)",
    )
    parser.add_argument(
        "--all-lists", action="store_true",
        help="모든 워치리스트 순회 분석",
    )
    args = parser.parse_args()

    # Provider 결정: 인자 > SCHEDULED_PROVIDER > LLM_PROVIDER
    provider = args.provider
    if not provider:
        provider = engine.CONFIG.get("SCHEDULED_PROVIDER", engine.LLM_PROVIDER)

    def log(msg):
        print(msg, flush=True)

    if args.all_lists:
        # 모든 리스트 순회
        list_names = engine.get_list_names()
        log(f"[스케줄 분석] {len(list_names)}개 리스트 순회 | LLM 기본: {provider}")
        for name in list_names:
            _run_for_list(name, provider, args.no_email, log)
    else:
        # 단일 리스트
        list_name = args.list  # None이면 active_list 사용
        resolved = list_name or engine.load_watchlist_data().get("active_list", "전체 종목")
        log(f"[스케줄 분석] 리스트: {resolved} | LLM: {provider}")
        return _run_for_list(resolved, provider, args.no_email, log)

    return 0


if __name__ == "__main__":
    sys.exit(main())
