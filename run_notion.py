"""
KR Chart Rater - Notion 기반 완전자동화 분석 스크립트
GitHub Actions에서 매일 실행.

사용법:
  python run_notion.py                      # 기본 (전체 종목)
  python run_notion.py --provider claude    # Claude 사용
  python run_notion.py --list "BIO"         # 특정 리스트만 필터
  python run_notion.py --dry-run            # 분석만, Notion 저장 안 함
"""

import sys
import argparse
from datetime import datetime

import engine
from notion_sync import NotionSync


def _format_stock_list(results):
    """A-1 또는 A-2 결과를 '종목명(티커), ...' 형식으로 변환."""
    return ", ".join(f"{r['ticker_name']}({r.get('code', '')})" for r in results)


def _run(list_name, provider, notion, dry_run, github_repo, log):
    """단일 리스트(또는 전체)에 대해 Notion 기반 분석 수행."""

    # Notion DB 설정 (config.txt에서 글로벌)
    notion_wl_db = engine.CONFIG.get("NOTION_WATCHLIST_DB", "")
    notion_rp_db = engine.CONFIG.get("NOTION_REPORT_DB", "")

    if not notion_wl_db:
        log("[X] NOTION_WATCHLIST_DB가 config.txt에 설정되지 않았습니다.")
        return

    # Notion에서 종목 읽기 (리스트 필터 적용)
    log(f"Notion에서 종목 읽기... (리스트: {list_name or '전체'})")
    ticker_names = notion.read_watchlist(notion_wl_db, list_name=list_name)

    if not ticker_names:
        log("분석할 종목 없음, 종료합니다.")
        return

    log(f"\n{len(ticker_names)}개 종목 | LLM: {provider}")

    # 종목별 분석
    a_results = []
    total_usage = {
        "input_tokens": 0, "output_tokens": 0, "total_tokens": 0,
        "input_cost_usd": 0, "output_cost_usd": 0, "total_cost_usd": 0,
        "api_calls": 0,
    }
    total_analyzed = 0
    total_errors = 0

    for name in ticker_names:
        log(f"\n--- {name} ---")

        # 1. 티커 해석 + 데이터 + 차트
        try:
            yf_ticker, code, market = engine.resolve_ticker(name)
            df = engine.fetch_ohlcv(name)
            chart_path = engine.generate_chart(name, df)
        except Exception as e:
            log(f"  [X] {name} 데이터/차트 오류: {e}")
            total_errors += 1
            continue

        # 2. LLM 분석 (1회 — 프롬프트 내부에서 3회 합의 수행)
        try:
            result = engine.analyze_chart_with_llm(
                chart_path=str(chart_path),
                ticker_name=name,
                provider=provider,
            )
            result["code"] = code
            result["market"] = market
            result["chart_path"] = str(chart_path)

            grade = result.get("grade", "N/A")
            reliability = result.get("reliability", "")
            consensus_count = result.get("consensus_count", "")
            log(f"  [분석] {grade} (신뢰도: {reliability}, {consensus_count})")

            usage = result.get("token_usage", {})
            if usage:
                engine._accumulate_usage(total_usage, usage)
            total_analyzed += 1

            if grade in ("A-1", "A-2"):
                a_results.append(result)
                log(f"  *** {grade} 선정 ***")

        except Exception as e:
            log(f"  [X] {name} 분석 오류: {e}")
            total_errors += 1

    # A-1 우선, A-2 다음, confidence 내림차순 정렬
    a_results.sort(key=lambda x: (0 if x.get("grade") == "A-1" else 1, -x.get("confidence", 0)))

    # A-1 / A-2 분리
    a1_results = [r for r in a_results if r.get("grade") == "A-1"]
    a2_results = [r for r in a_results if r.get("grade") == "A-2"]

    # 요약
    cost = total_usage["total_cost_usd"]
    log(f"\n{'='*50}")
    log(f"분석 완료: {total_analyzed}개 분석, {total_errors}개 오류")
    log(f"A-1/A-2 선정: {len(a_results)}개")
    log(f"비용: ${cost:.4f} (약 {cost*1400:.0f}원)")

    if a_results:
        log("\nA-1/A-2 선정 종목:")
        for i, r in enumerate(a_results, 1):
            log(f"  {i}. [{r.get('grade', '')}] {r['ticker_name']} ({r.get('code', '')})")

    # Notion 보고서 작성
    if not dry_run and notion_rp_db:
        log(f"\nNotion 보고서 작성 중...")
        try:
            dt = datetime.now(engine.KST)
            title = dt.strftime("%m/%d")
            meta = {
                "total_analyzed": total_analyzed,
                "a_count": len(a_results),
                "errors": total_errors,
                "provider": provider,
                "cost_usd": cost,
            }
            summary_props = {
                "분석일": dt.strftime("%Y-%m-%d"),
                "리스트": list_name or "전체",
                "종목수": total_analyzed,
                "선정": len(a_results),
                "비용": round(cost, 4),
                "A-1": _format_stock_list(a1_results),
                "A-2": _format_stock_list(a2_results),
            }
            blocks = notion.build_report_blocks(a_results, meta, github_repo)
            notion.create_report_page(notion_rp_db, title, dt.strftime("%Y-%m-%d"), summary_props, blocks)
            log(f"Notion 보고서 작성 완료: {title}")
        except Exception as e:
            log(f"[X] Notion 보고서 작성 실패: {e}")
    elif dry_run:
        log("\n[dry-run] Notion 저장 건너뜀")

    # 이메일 발송
    email_to = engine.CONFIG.get("EMAIL_TO", "")
    if email_to and not dry_run:
        try:
            email_result = {
                "total_analyzed": total_analyzed,
                "total_errors": total_errors,
                "a_rated": a_results,
                "token_usage": total_usage,
            }
            subject, body = engine.build_email_body(email_result, provider)
            engine.send_report_email(to_addr=email_to, subject=subject, body_text=body)
            log(f"[이메일] {email_to}로 발송 완료")
        except Exception as e:
            log(f"[X] 이메일 발송 실패: {e}")


def main():
    parser = argparse.ArgumentParser(description="KR Chart Rater - Notion 자동화 분석")
    parser.add_argument(
        "--provider", default=None, choices=["claude", "gemini"],
        help="LLM provider (기본: config.txt 설정값)",
    )
    parser.add_argument(
        "--list", default=None,
        help="Notion DB 'リスト' multi-select 필터 (예: 'BIO'). 미지정 시 전체.",
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="분석만 수행, Notion 저장 안 함",
    )
    args = parser.parse_args()

    provider = args.provider or engine.CONFIG.get("LLM_PROVIDER", engine.LLM_PROVIDER)
    github_repo = engine.CONFIG.get("GITHUB_REPO", "")

    def log(msg):
        print(msg, flush=True)

    # Notion 클라이언트 초기화
    try:
        notion_token = engine.read_secret("notion_api_key.txt")
        notion = NotionSync(notion_token)
    except Exception as e:
        log(f"[X] Notion 인증 실패: {e}")
        log("NOTION_API_KEY 환경변수 또는 secrets/notion_api_key.txt를 확인하세요.")
        return 2

    _run(args.list, provider, notion, args.dry_run, github_repo, log)
    return 0


if __name__ == "__main__":
    sys.exit(main())
