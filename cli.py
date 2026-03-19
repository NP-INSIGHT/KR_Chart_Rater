"""
KR Chart Rater - CLI 진입점
GitHub Actions 및 터미널 환경에서 사용.

사용법:
  python cli.py stocks --tickers 삼성전자 SK하이닉스 --provider claude
  python cli.py stocks --watchlist --provider gemini
  python cli.py watchlist --add 삼성전자 카카오
  python cli.py themes --all --provider claude --top-n 10
  python cli.py themes --themes 2차전지 반도체 --provider gemini
"""

import sys
import argparse
import engine


def cmd_stocks(args):
    """종목 분석 실행"""
    list_name = getattr(args, "list", None)

    # 워치리스트 모드 vs 직접 지정 모드
    if getattr(args, "watchlist", False):
        tickers = engine.get_active_watchlist(list_name)
        if not tickers:
            print("오류: 워치리스트에 활성 종목이 없습니다. 먼저 종목을 추가하세요.")
            return 2
        label = f" ({list_name})" if list_name else ""
        print(f"[워치리스트{label}] {len(tickers)}개 활성 종목: {', '.join(tickers)}", flush=True)
    else:
        tickers = args.tickers
        if not tickers:
            print("오류: --tickers에 종목명을 1개 이상 입력하세요.")
            return 2

    def log_print(msg):
        print(msg, flush=True)

    try:
        result = engine.run_stock_analysis(
            ticker_names=tickers,
            provider=args.provider,
            log_callback=log_print,
        )
    except Exception as e:
        print(f"\n치명적 오류: {e}")
        return 2

    # Word 보고서 내보내기
    if getattr(args, "docx", False):
        try:
            docx_path = engine.save_results_docx(result, "stocks")
            print(f"\n[Word] 보고서 저장: {docx_path}", flush=True)
        except Exception as e:
            print(f"\n[X] Word 내보내기 실패: {e}", flush=True)

    # 종료 코드 결정
    total = result.get("total_analyzed", 0)
    errors = result.get("total_errors", 0)

    if errors == 0:
        return 0  # 성공
    elif total > errors:
        return 1  # 부분 실패
    else:
        return 2  # 전체 실패


def cmd_themes(args):
    """테마 분석 실행"""
    if not args.all and not args.themes:
        print("오류: --themes로 테마명을 지정하거나 --all을 사용하세요.")
        return 2

    theme_names = None if args.all else args.themes

    def log_print(msg):
        print(msg, flush=True)

    try:
        result = engine.run_theme_analysis(
            theme_names=theme_names,
            top_n=args.top_n,
            provider=args.provider,
            log_callback=log_print,
        )
    except Exception as e:
        print(f"\n치명적 오류: {e}")
        return 2

    # Word 보고서 내보내기
    if getattr(args, "docx", False):
        try:
            docx_path = engine.save_results_docx(result, "themes")
            print(f"\n[Word] 보고서 저장: {docx_path}", flush=True)
        except Exception as e:
            print(f"\n[X] Word 내보내기 실패: {e}", flush=True)

    # 종료 코드: 테마 중 하나라도 성공이면 0
    themes = result.get("themes", [])
    has_success = any(t.get("holdings_analyzed") for t in themes)
    return 0 if has_success else 2


def cmd_refresh_cache():
    """종목 캐시 갱신"""
    print("KIND API에서 전체 종목 목록을 가져오는 중...", flush=True)
    try:
        count = engine.refresh_ticker_cache()
        print(f"[OK] ticker_cache.json 갱신 완료: {count}개 종목", flush=True)
        return 0
    except Exception as e:
        print(f"[X] 캐시 갱신 실패: {e}", flush=True)
        return 2


def cmd_refresh_themes(args):
    """테마 ETF 캐시 갱신 (LLM 기반)"""
    def log_print(msg):
        print(msg, flush=True)

    try:
        count = engine.refresh_theme_cache(
            provider=args.provider,
            log_callback=log_print,
        )
        print(f"\n[OK] theme_cache.json 갱신 완료: {count}개 테마", flush=True)
        return 0
    except Exception as e:
        print(f"\n[X] 테마 캐시 갱신 실패: {e}", flush=True)
        return 2


def cmd_watchlist(args):
    """워치리스트 관리"""
    list_name = getattr(args, "list", None)

    # 리스트 목록 출력
    if getattr(args, "lists", False):
        names = engine.get_list_names()
        data = engine.load_watchlist_data()
        active = data.get("active_list", "")
        print(f"\n[워치리스트 목록] {len(names)}개", flush=True)
        print("-" * 40, flush=True)
        for name in names:
            marker = " *" if name == active else ""
            lst = engine.get_list(name)
            count = len(lst["stocks"]) if lst else 0
            print(f"  {name} ({count}개 종목){marker}", flush=True)
        return 0

    # 리스트 생성
    if getattr(args, "create_list", None):
        name = args.create_list
        if engine.create_list(name):
            print(f"[OK] 워치리스트 '{name}' 생성 완료", flush=True)
        else:
            print(f"[!] '{name}' 리스트가 이미 존재합니다.", flush=True)
        return 0

    # 리스트 삭제
    if getattr(args, "delete_list", None):
        name = args.delete_list
        if engine.delete_list(name):
            print(f"[OK] 워치리스트 '{name}' 삭제 완료", flush=True)
        else:
            print(f"[!] '{name}' 리스트를 삭제할 수 없습니다. (마지막 리스트 또는 존재하지 않음)", flush=True)
        return 0

    # 추가
    if args.add:
        added = engine.add_to_watchlist(args.add, list_name)
        print(f"[OK] {added}개 종목 추가", flush=True)
        _print_watchlist(list_name)
        return 0

    # 삭제
    if args.remove:
        removed = engine.remove_from_watchlist(args.remove, list_name)
        print(f"[OK] {removed}개 종목 삭제", flush=True)
        _print_watchlist(list_name)
        return 0

    # 활성화
    if args.on:
        changed = engine.set_watchlist_active(args.on, True, list_name)
        print(f"[OK] {changed}개 종목 활성화", flush=True)
        _print_watchlist(list_name)
        return 0

    # 비활성화
    if args.off:
        changed = engine.set_watchlist_active(args.off, False, list_name)
        print(f"[OK] {changed}개 종목 비활성화", flush=True)
        _print_watchlist(list_name)
        return 0

    # 인자 없으면 현재 목록 출력
    _print_watchlist(list_name)
    return 0


def _print_watchlist(list_name=None):
    """워치리스트 출력"""
    stocks = engine.load_watchlist(list_name)
    resolved_name = list_name or engine.load_watchlist_data().get("active_list", "전체 종목")

    if not stocks:
        print(f"\n워치리스트 '{resolved_name}'이(가) 비어있습니다.", flush=True)
        return

    active = [s for s in stocks if s.get("active", True)]
    inactive = [s for s in stocks if not s.get("active", True)]

    print(f"\n[워치리스트: {resolved_name}] 총 {len(stocks)}개 (활성: {len(active)}, 비활성: {len(inactive)})", flush=True)
    print("-" * 40, flush=True)

    for s in stocks:
        status = "O" if s.get("active", True) else "X"
        added = s.get("added", "")
        print(f"  [{status}] {s['name']}  (추가: {added})", flush=True)


def main():
    parser = argparse.ArgumentParser(
        description="KR Chart Rater - 차트 기반 종목 추천 프로그램",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
예시:
  python cli.py stocks --tickers 삼성전자 SK하이닉스 LG에너지솔루션
  python cli.py stocks --watchlist --provider gemini
  python cli.py watchlist --add 삼성전자 카카오
  python cli.py watchlist --remove 카카오
  python cli.py watchlist --on 삼성전자 --off SK하이닉스
  python cli.py refresh-themes --provider gemini
  python cli.py themes --all
  python cli.py themes --themes 2차전지 반도체 --top-n 5
        """,
    )
    subparsers = parser.add_subparsers(dest="command", help="실행 모드")

    # stocks 서브커맨드
    stock_parser = subparsers.add_parser("stocks", help="개별 종목 차트 분석")
    stock_source = stock_parser.add_mutually_exclusive_group(required=True)
    stock_source.add_argument(
        "--tickers", nargs="+",
        help="분석할 종목명 (한글, 공백 구분)",
    )
    stock_source.add_argument(
        "--watchlist", action="store_true",
        help="워치리스트의 활성 종목 분석",
    )
    stock_parser.add_argument(
        "--provider", default=None, choices=["claude", "gemini"],
        help="LLM provider (기본: config.txt 설정값)",
    )
    stock_parser.add_argument(
        "--docx", action="store_true",
        help="분석 완료 후 Word 보고서(.docx) 자동 생성",
    )
    stock_parser.add_argument(
        "--list", default=None,
        help="워치리스트 이름 (기본: 활성 리스트)",
    )

    # watchlist 서브커맨드
    wl_parser = subparsers.add_parser("watchlist", help="워치리스트 관리")
    wl_parser.add_argument(
        "--list", default=None,
        help="워치리스트 이름 (기본: 활성 리스트)",
    )
    wl_parser.add_argument(
        "--lists", action="store_true",
        help="모든 워치리스트 목록 출력",
    )
    wl_parser.add_argument(
        "--create-list", default=None,
        help="새 워치리스트 생성",
    )
    wl_parser.add_argument(
        "--delete-list", default=None,
        help="워치리스트 삭제",
    )
    wl_parser.add_argument(
        "--add", nargs="+",
        help="종목 추가",
    )
    wl_parser.add_argument(
        "--remove", nargs="+",
        help="종목 삭제",
    )
    wl_parser.add_argument(
        "--on", nargs="+",
        help="종목 활성화",
    )
    wl_parser.add_argument(
        "--off", nargs="+",
        help="종목 비활성화",
    )

    # themes 서브커맨드
    theme_parser = subparsers.add_parser("themes", help="테마별 ETF 보유종목 분석")
    theme_group = theme_parser.add_mutually_exclusive_group(required=True)
    theme_group.add_argument(
        "--themes", nargs="+",
        help="분석할 테마명 (공백 구분)",
    )
    theme_group.add_argument(
        "--all", action="store_true",
        help="themes.json의 전체 테마 분석",
    )
    theme_parser.add_argument(
        "--top-n", type=int, default=None,
        help="ETF 보유종목 상위 N개 (기본: config.txt 설정값)",
    )
    theme_parser.add_argument(
        "--provider", default=None, choices=["claude", "gemini"],
        help="LLM provider (기본: config.txt 설정값)",
    )
    theme_parser.add_argument(
        "--docx", action="store_true",
        help="분석 완료 후 Word 보고서(.docx) 자동 생성",
    )

    # refresh-cache 서브커맨드
    subparsers.add_parser("refresh-cache", help="KIND API에서 전체 종목 캐시 갱신")

    # refresh-themes 서브커맨드
    rt_parser = subparsers.add_parser("refresh-themes", help="LLM으로 테마별 ETF 캐시 갱신")
    rt_parser.add_argument(
        "--provider", default=None, choices=["claude", "gemini"],
        help="LLM provider (기본: config.txt 설정값)",
    )

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return 0

    if args.command == "stocks":
        return cmd_stocks(args)
    elif args.command == "watchlist":
        return cmd_watchlist(args)
    elif args.command == "themes":
        return cmd_themes(args)
    elif args.command == "refresh-cache":
        return cmd_refresh_cache()
    elif args.command == "refresh-themes":
        return cmd_refresh_themes(args)

    return 0


if __name__ == "__main__":
    sys.exit(main())
