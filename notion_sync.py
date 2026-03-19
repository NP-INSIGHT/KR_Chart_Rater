"""
KR Chart Rater - Notion API 래퍼
Notion DB에서 종목 리스트 읽기 / 분석 보고서 쓰기
httpx로 직접 Notion API 호출 (안정적인 API 버전 사용)
"""

import time
import logging
import httpx
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
logger = logging.getLogger("kr_chart_rater")

NOTION_BASE = "https://api.notion.com/v1"
NOTION_VERSION = "2022-06-28"


def _to_uuid(id_str):
    """32자 hex를 UUID 형식(하이픈 포함)으로 변환."""
    s = id_str.replace("-", "").strip()
    if len(s) == 32:
        return f"{s[:8]}-{s[8:12]}-{s[12:16]}-{s[16:20]}-{s[20:]}"
    return id_str


class NotionSync:
    def __init__(self, token):
        self._token = token
        self._headers = {
            "Authorization": f"Bearer {token}",
            "Notion-Version": NOTION_VERSION,
            "Content-Type": "application/json",
        }
        self._last = 0

    def _throttle(self):
        """Notion rate limit: 3 req/sec"""
        gap = time.time() - self._last
        if gap < 0.35:
            time.sleep(0.35 - gap)
        self._last = time.time()

    def _post(self, path, body=None):
        """Notion API POST 요청."""
        self._throttle()
        resp = httpx.post(
            f"{NOTION_BASE}/{path}",
            headers=self._headers,
            json=body or {},
            timeout=30,
        )
        resp.raise_for_status()
        return resp.json()

    # ── 종목 리스트 읽기 ──

    def read_watchlist(self, db_id, list_name=None):
        """종목 리스트 DB에서 종목명 읽기 (페이지네이션 포함).

        Args:
            db_id: Notion 데이터베이스 ID
            list_name: 필터링할 리스트 이름 (multi-select "리스트" 속성).
                       None이면 전체 종목 반환.
        """
        db_id = _to_uuid(db_id)
        names = []
        has_more = True
        start_cursor = None

        while has_more:
            body = {"page_size": 100}
            if start_cursor:
                body["start_cursor"] = start_cursor

            # 리스트 필터 (multi-select "리스트" 속성)
            if list_name:
                body["filter"] = {
                    "property": "리스트",
                    "multi_select": {"contains": list_name},
                }

            resp = self._post(f"databases/{db_id}/query", body)
            for page in resp.get("results", []):
                props = page.get("properties", {})

                # "종목명" title 추출
                title_prop = props.get("종목명", {})
                if title_prop.get("type") == "title":
                    titles = title_prop.get("title", [])
                    if titles:
                        name = titles[0].get("plain_text", "").strip()
                        if name:
                            names.append(name)

            has_more = resp.get("has_more", False)
            start_cursor = resp.get("next_cursor")

        logger.info(f"Notion 종목 리스트 읽기: {len(names)}개 종목 (DB: {db_id[:8]}..., 리스트: {list_name or '전체'})")
        return names

    # ── 보고서 쓰기 ──

    def create_report_page(self, db_id, title, date_str, summary_props, blocks):
        """보고서 DB에 새 페이지 생성."""
        db_id = _to_uuid(db_id)

        properties = {
            "날짜": {"title": [{"text": {"content": title}}]},
        }

        # 선택적 속성 (DB에 있을 때만)
        if "분석일" in summary_props:
            properties["분석일"] = {"date": {"start": summary_props["분석일"]}}
        if "종목수" in summary_props:
            properties["종목수"] = {"number": summary_props["종목수"]}
        if "선정" in summary_props:
            properties["선정"] = {"number": summary_props["선정"]}
        if "비용" in summary_props:
            properties["비용"] = {"number": summary_props["비용"]}

        # 리스트명 (rich_text)
        if "리스트" in summary_props:
            properties["리스트"] = {"rich_text": [{"text": {"content": summary_props["리스트"]}}]}

        # A-1 / A-2 종목 리스트 (rich_text, 추천 순)
        if "A-1" in summary_props:
            properties["A-1"] = {"rich_text": [{"text": {"content": summary_props["A-1"][:2000]}}]}
        if "A-2" in summary_props:
            properties["A-2"] = {"rich_text": [{"text": {"content": summary_props["A-2"][:2000]}}]}

        # 페이지 생성 (최대 100 블록)
        first_blocks = blocks[:100]
        remaining_blocks = blocks[100:]

        page_body = {
            "parent": {"database_id": db_id},
            "properties": properties,
            "children": first_blocks,
        }
        page = self._post("pages", page_body)
        page_id = page["id"]

        # 100개 초과 블록은 append로 분할 추가
        while remaining_blocks:
            batch = remaining_blocks[:100]
            remaining_blocks = remaining_blocks[100:]
            self._post(f"blocks/{page_id}/children", {"children": batch})

        logger.info(f"Notion 보고서 페이지 생성: {title} (DB: {db_id[:8]}...)")
        return page_id

    def build_report_blocks(self, a_results, meta, github_repo=None):
        """A-1/A-2 종목 리스트를 간단한 Notion 블록으로 변환."""
        blocks = []

        total = meta.get("total_analyzed", 0)
        selected = len(a_results)
        provider = meta.get("provider", "")
        cost = meta.get("cost_usd", 0)

        blocks.append(self._callout(
            f"분석 종목: {total}개 | 선정: {selected}개 | "
            f"LLM: {provider.upper()} | 비용: ${cost:.4f}"
        ))

        if not a_results:
            blocks.append(self._paragraph("A-1/A-2 선정 종목 없음"))
            return blocks

        # A-1 / A-2 분리
        a1 = [r for r in a_results if r.get("grade") == "A-1"]
        a2 = [r for r in a_results if r.get("grade") == "A-2"]

        if a1:
            blocks.append(self._heading3("A-1 (속도형 매력)"))
            for idx, r in enumerate(a1, 1):
                name = r.get("ticker_name", "")
                code = r.get("code", "")
                blocks.append(self._paragraph(f"{idx}. {name} ({code})"))

        if a2:
            blocks.append(self._heading3("A-2 (완만추세 지속형)"))
            for idx, r in enumerate(a2, 1):
                name = r.get("ticker_name", "")
                code = r.get("code", "")
                blocks.append(self._paragraph(f"{idx}. {name} ({code})"))

        blocks.append(self._divider())
        blocks.append(self._paragraph(
            f"비용: ${cost:.4f} (약 {cost * 1400:.0f}원) | "
            f"Generated by KR Chart Rater"
        ))

        return blocks

    # ── Notion block 헬퍼 ──

    @staticmethod
    def _heading2(text):
        return {"type": "heading_2", "heading_2": {"rich_text": [{"type": "text", "text": {"content": text}}]}}

    @staticmethod
    def _heading3(text):
        return {"type": "heading_3", "heading_3": {"rich_text": [{"type": "text", "text": {"content": text}}]}}

    @staticmethod
    def _paragraph(text):
        # Notion rich_text content 최대 2000자
        chunks = [text[i:i+2000] for i in range(0, len(text), 2000)]
        return {"type": "paragraph", "paragraph": {"rich_text": [{"type": "text", "text": {"content": c}} for c in chunks]}}

    @staticmethod
    def _bulleted(text):
        return {"type": "bulleted_list_item", "bulleted_list_item": {"rich_text": [{"type": "text", "text": {"content": text}}]}}

    @staticmethod
    def _divider():
        return {"type": "divider", "divider": {}}

    @staticmethod
    def _callout(text):
        return {"type": "callout", "callout": {"rich_text": [{"type": "text", "text": {"content": text}}], "icon": {"type": "emoji", "emoji": "\U0001f4ca"}}}

    @staticmethod
    def _image(url):
        return {"type": "image", "image": {"type": "external", "external": {"url": url}}}
