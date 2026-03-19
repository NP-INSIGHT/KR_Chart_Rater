"""
KR Chart Rater - GUI (customtkinter)
3탭: 종목 분석 / 테마 분석 / 결과 조회
"""

import os
import sys
import json
import threading
import queue
from datetime import datetime
from pathlib import Path

import customtkinter as ctk
from PIL import Image

import engine

ctk.set_appearance_mode("light")

# ── NP 오렌지 테마 색상 ──
ACCENT = "#E8864A"
ACCENT_HOVER = "#D0743A"
ACCENT_LIGHT = "#FFF3EB"
ACCENT_DARK = "#C2652E"
TEXT_DARK = "#3D2B1F"
BORDER_COLOR = "#F0D0B0"


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("KR Chart Rater - 차트 기반 종목 추천")
        self.geometry("1200x800")
        self.minsize(1000, 650)

        # 상태
        self.log_queue = queue.Queue()
        self._running = False
        self._stock_entries = []        # 종목 분석 탭: 입력된 종목 리스트
        self._theme_checkboxes = {}     # 테마 분석 탭: 테마 체크박스 dict
        self._result_files = []         # 결과 조회 탭: 결과 파일 리스트

        # 상단 바로가기 버튼
        top_bar = ctk.CTkFrame(self, fg_color="transparent")
        top_bar.pack(fill="x", padx=10, pady=(5, 0))

        self._orange_btn(
            top_bar, text="config.txt 열기", width=120,
            command=lambda: os.startfile(str(engine.BASE_DIR / "config.txt")),
        ).pack(side="left", padx=(0, 5))
        self._orange_btn(
            top_bar, text="테마 설정 열기", width=120,
            command=lambda: os.startfile(str(engine.BASE_DIR / "themes.json")),
        ).pack(side="left", padx=(0, 5))
        self._orange_btn(
            top_bar, text="프롬프트 편집", width=120,
            command=lambda: os.startfile(str(engine.BASE_DIR / "chart_prompt.txt")),
        ).pack(side="left", padx=(0, 5))
        self._orange_btn(
            top_bar, text="차트 폴더 열기", width=120,
            command=lambda: os.startfile(str(engine.CHARTS_DIR)),
        ).pack(side="left", padx=(0, 5))
        self._orange_btn(
            top_bar, text="이메일 설정", width=100,
            fg_color="#2E7D32", hover_color="#1B5E20",
            command=self._open_email_settings,
        ).pack(side="left", padx=(0, 5))

        # 탭뷰
        self.tabview = ctk.CTkTabview(
            self, anchor="nw",
            segmented_button_fg_color=ACCENT_LIGHT,
            segmented_button_selected_color=ACCENT,
            segmented_button_selected_hover_color=ACCENT_HOVER,
            segmented_button_unselected_color=ACCENT_LIGHT,
            segmented_button_unselected_hover_color=BORDER_COLOR,
            text_color=TEXT_DARK,
            text_color_disabled="gray",
        )
        self.tabview.pack(fill="both", expand=True, padx=10, pady=(5, 10))

        self._build_stock_tab()
        self._build_theme_tab()
        self._build_result_tab()

        # 로그 폴링
        self._poll_log()

    # ============================================================
    # 유틸
    # ============================================================
    def _orange_btn(self, parent, **kwargs):
        kwargs.setdefault("fg_color", ACCENT)
        kwargs.setdefault("hover_color", ACCENT_HOVER)
        kwargs.setdefault("text_color", "white")
        return ctk.CTkButton(parent, **kwargs)

    def _section_label(self, parent, text):
        return ctk.CTkLabel(
            parent, text=text,
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=ACCENT_DARK,
        )

    def _open_email_settings(self):
        """글로벌 SMTP 설정 다이얼로그"""
        win = ctk.CTkToplevel(self)
        win.title("이메일 설정 (SMTP)")
        win.geometry("400x320")
        win.resizable(False, False)
        win.grab_set()

        pad = {"padx": 15, "pady": (8, 0)}

        fields = [
            ("SMTP 호스트", "EMAIL_SMTP_HOST", "smtp.gmail.com"),
            ("SMTP 포트", "EMAIL_SMTP_PORT", "587"),
            ("발신 이메일", "EMAIL_FROM", ""),
            ("비밀번호 파일", "EMAIL_PASSWORD_FILE", "email_password.txt"),
            ("수신 이메일 (기본)", "EMAIL_TO", ""),
        ]

        entries = {}
        for label, key, default in fields:
            ctk.CTkLabel(win, text=label, font=ctk.CTkFont(weight="bold")).pack(anchor="w", **pad)
            entry = ctk.CTkEntry(win, border_color=BORDER_COLOR)
            entry.pack(fill="x", padx=15, pady=(2, 0))
            current = engine.CONFIG.get(key, default)
            if current:
                entry.insert(0, current)
            entries[key] = entry

        def _save():
            config_path = engine.BASE_DIR / "config.txt"
            lines = []
            if config_path.exists():
                lines = config_path.read_text(encoding="utf-8").splitlines()

            for key, entry in entries.items():
                val = entry.get().strip()
                # config.txt에서 해당 키 업데이트 또는 추가
                found = False
                for i, line in enumerate(lines):
                    stripped = line.strip()
                    if stripped.startswith(f"{key}=") or stripped.startswith(f"# {key}="):
                        lines[i] = f"{key}={val}" if val else f"# {key}="
                        found = True
                        break
                if not found and val:
                    lines.append(f"{key}={val}")
                # 즉시 CONFIG에도 반영
                if val:
                    engine.CONFIG[key] = val

            config_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
            win.destroy()

        self._orange_btn(win, text="저장", width=100, command=_save).pack(pady=(15, 10))

    def _append_log(self, textbox, msg):
        """텍스트박스에 로그 추가"""
        textbox.configure(state="normal")
        textbox.insert("end", msg + "\n")
        textbox.see("end")
        textbox.configure(state="disabled")

    def _clear_log(self, textbox):
        """텍스트박스 내용 초기화"""
        textbox.configure(state="normal")
        textbox.delete("1.0", "end")
        textbox.configure(state="disabled")

    def _poll_log(self):
        """로그 큐 폴링 (100ms 간격)"""
        try:
            while True:
                tag, msg = self.log_queue.get_nowait()
                if tag == "__done_stock__":
                    self._on_stock_done(msg)
                elif tag == "__done_theme__":
                    self._on_theme_done(msg)
                elif tag == "__done_refresh_themes__":
                    self._on_refresh_themes_done(msg)
                elif tag == "stock_log":
                    self._append_log(self.stock_log_box, msg)
                elif tag == "theme_log":
                    self._append_log(self.theme_log_box, msg)
        except queue.Empty:
            pass
        self.after(100, self._poll_log)

    # ============================================================
    # 탭 1: 종목 분석
    # ============================================================
    def _build_stock_tab(self):
        tab = self.tabview.add("종목 분석")

        content = ctk.CTkFrame(tab, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=5, pady=5)

        # ── 좌측: 종목 입력 ──
        left = ctk.CTkFrame(content, width=280, fg_color="transparent")
        left.pack(side="left", fill="y", padx=(0, 10))
        left.pack_propagate(False)

        # 리스트 셀렉터
        list_selector_frame = ctk.CTkFrame(left, fg_color="transparent")
        list_selector_frame.pack(fill="x", pady=(0, 8))

        self._section_label(list_selector_frame, "워치리스트").pack(side="left")

        self._orange_btn(
            list_selector_frame, text="+", width=30, height=28,
            command=self._create_new_list,
        ).pack(side="right", padx=(2, 0))
        self._orange_btn(
            list_selector_frame, text="\u2699", width=30, height=28,
            fg_color="#4A90D9", hover_color="#3A78C2",
            command=self._open_list_config,
        ).pack(side="right", padx=(2, 0))

        list_names = engine.get_list_names()
        data = engine.load_watchlist_data()
        active_list = data.get("active_list", list_names[0] if list_names else "전체 종목")
        self._current_list_name = active_list

        self.list_selector_var = ctk.StringVar(value=active_list)
        self.list_selector = ctk.CTkOptionMenu(
            list_selector_frame,
            variable=self.list_selector_var,
            values=list_names,
            fg_color=ACCENT, button_color=ACCENT_DARK,
            button_hover_color=ACCENT_HOVER,
            dropdown_fg_color="white",
            dropdown_hover_color=ACCENT_LIGHT,
            dropdown_text_color=TEXT_DARK,
            width=130,
            command=self._on_list_changed,
        )
        self.list_selector.pack(side="right")

        self._section_label(left, "분석 종목").pack(anchor="w", pady=(0, 5))

        # 종목 입력 필드
        input_frame = ctk.CTkFrame(left, fg_color="transparent")
        input_frame.pack(fill="x", pady=(0, 5))

        self.stock_entry = ctk.CTkEntry(
            input_frame, placeholder_text="종목명 (콤마로 여러개 입력)",
            border_color=BORDER_COLOR,
        )
        self.stock_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.stock_entry.bind("<Return>", lambda e: self._add_stock())

        self._orange_btn(
            input_frame, text="추가", width=60,
            command=self._add_stock,
        ).pack(side="left")

        # 종목 리스트 (스크롤)
        self.stock_list_frame = ctk.CTkScrollableFrame(
            left, fg_color="transparent",
        )
        self.stock_list_frame.pack(fill="both", expand=True, pady=(0, 5))

        # 전체 선택/해제 버튼
        toggle_frame = ctk.CTkFrame(left, fg_color="transparent")
        toggle_frame.pack(fill="x", pady=(0, 3))

        self._orange_btn(
            toggle_frame, text="전체 선택", width=90,
            command=lambda: self._toggle_all_stocks(True),
        ).pack(side="left", padx=(0, 5))
        self._orange_btn(
            toggle_frame, text="전체 해제", width=90,
            command=lambda: self._toggle_all_stocks(False),
        ).pack(side="left")

        # 삭제/초기화 버튼
        btn_frame = ctk.CTkFrame(left, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(0, 10))

        self._orange_btn(
            btn_frame, text="선택 삭제", width=100,
            command=self._remove_selected_stocks,
        ).pack(side="left", padx=(0, 5))
        self._orange_btn(
            btn_frame, text="전체 삭제", width=100,
            command=self._clear_all_stocks,
        ).pack(side="left")

        # LLM 선택
        self._section_label(left, "LLM 설정").pack(anchor="w", pady=(10, 5))

        self.stock_provider_var = ctk.StringVar(value=engine.LLM_PROVIDER)
        for label, val in [("Claude", "claude"), ("Gemini", "gemini")]:
            ctk.CTkRadioButton(
                left, text=label, variable=self.stock_provider_var, value=val,
                fg_color=ACCENT, hover_color=ACCENT_HOVER,
            ).pack(anchor="w", pady=2)

        # 실행 버튼
        self.stock_run_btn = self._orange_btn(
            left, text="분석 실행", height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._run_stock_analysis,
        )
        self.stock_run_btn.pack(fill="x", pady=(15, 5))

        # Word 내보내기 버튼
        self.stock_docx_btn = self._orange_btn(
            left, text="Word 보고서 내보내기", height=32,
            fg_color="#4A90D9", hover_color="#3A78C2",
            command=self._export_stock_docx,
            state="disabled",
        )
        self.stock_docx_btn.pack(fill="x", pady=(5, 5))

        # 메일 발송 버튼
        self.stock_email_btn = self._orange_btn(
            left, text="메일 발송", height=32,
            fg_color="#2E7D32", hover_color="#1B5E20",
            command=self._send_stock_email,
            state="disabled",
        )
        self.stock_email_btn.pack(fill="x", pady=(0, 5))

        # 상태 표시
        self.stock_status = ctk.CTkLabel(
            left, text="대기 중",
            text_color=ACCENT_DARK,
            font=ctk.CTkFont(size=12),
        )
        self.stock_status.pack(anchor="w")

        # 마지막 분석 결과 저장
        self._last_stock_result = None

        # ── 워치리스트에서 기존 종목 로드 ──
        self._load_watchlist_to_gui()

        # ── 우측: 결과 + 로그 ──
        right = ctk.CTkFrame(content, fg_color="transparent")
        right.pack(side="left", fill="both", expand=True)

        # 결과 영역
        self._section_label(right, "차트 분석 결과").pack(anchor="w", pady=(0, 5))

        self.stock_result_frame = ctk.CTkScrollableFrame(
            right, height=300, fg_color="transparent",
        )
        self.stock_result_frame.pack(fill="both", expand=True, pady=(0, 10))

        # 로그 영역
        self._section_label(right, "실행 로그").pack(anchor="w", pady=(0, 5))

        self.stock_log_box = ctk.CTkTextbox(
            right, height=200, state="disabled",
            font=ctk.CTkFont(family="Consolas", size=11),
            wrap="word",
        )
        self.stock_log_box.pack(fill="both", expand=True)

    def _load_watchlist_to_gui(self):
        """워치리스트에서 종목 로드하여 GUI에 표시"""
        stocks = engine.load_watchlist(self._current_list_name)
        for s in stocks:
            name = s["name"]
            active = s.get("active", True)

            entry_data = {"name": name, "var": ctk.BooleanVar(value=active)}
            self._stock_entries.append(entry_data)

            cb = ctk.CTkCheckBox(
                self.stock_list_frame, text=name,
                variable=entry_data["var"],
                fg_color=ACCENT, hover_color=ACCENT_HOVER,
                command=self._on_stock_check_changed,
            )
            cb.pack(anchor="w", pady=2)
            entry_data["widget"] = cb

    def _save_gui_to_watchlist(self):
        """GUI 상태를 watchlist.json에 저장"""
        stocks = []
        today = datetime.now(engine.KST).strftime("%Y-%m-%d")
        for e in self._stock_entries:
            stocks.append({
                "name": e["name"],
                "active": e["var"].get(),
                "added": today,
            })
        engine.save_watchlist(stocks, self._current_list_name)

    def _toggle_all_stocks(self, value):
        """전체 종목 선택/해제"""
        for e in self._stock_entries:
            e["var"].set(value)
        self._save_gui_to_watchlist()

    def _on_stock_check_changed(self):
        """체크박스 상태 변경 시 워치리스트에 반영"""
        self._save_gui_to_watchlist()

    def _add_stock(self):
        """종목 추가 (콤마/공백 구분으로 여러 종목 동시 입력 가능)"""
        raw = self.stock_entry.get().strip()
        if not raw:
            return

        # 콤마 또는 공백으로 분리
        names = [n.strip() for n in raw.replace(",", " ").split() if n.strip()]
        existing = {e["name"] for e in self._stock_entries}

        added_any = False
        for name in names:
            if name in existing:
                continue  # 중복 방지
            existing.add(name)

            entry_data = {"name": name, "var": ctk.BooleanVar(value=True)}
            self._stock_entries.append(entry_data)

            cb = ctk.CTkCheckBox(
                self.stock_list_frame, text=name,
                variable=entry_data["var"],
                fg_color=ACCENT, hover_color=ACCENT_HOVER,
                command=self._on_stock_check_changed,
            )
            cb.pack(anchor="w", pady=2)
            entry_data["widget"] = cb
            added_any = True

        self.stock_entry.delete(0, "end")
        self.stock_entry.focus()

        if added_any:
            self._save_gui_to_watchlist()

    def _remove_selected_stocks(self):
        """선택된(체크된) 종목 삭제"""
        to_remove = [e for e in self._stock_entries if e["var"].get()]
        if not to_remove:
            return
        for e in to_remove:
            e["widget"].destroy()
            self._stock_entries.remove(e)
        self._save_gui_to_watchlist()

    def _clear_all_stocks(self):
        """전체 종목 삭제"""
        for e in self._stock_entries:
            e["widget"].destroy()
        self._stock_entries.clear()
        self._save_gui_to_watchlist()

    def _run_stock_analysis(self):
        """종목 분석 실행 (백그라운드 스레드)"""
        if self._running:
            return

        names = [e["name"] for e in self._stock_entries if e["var"].get()]
        if not names:
            self._append_log(self.stock_log_box, "[!] 분석할 종목을 추가하거나 체크(활성화)하세요.")
            return

        self._running = True
        self.stock_run_btn.configure(state="disabled")
        self.stock_status.configure(text="분석 중...")
        self._clear_log(self.stock_log_box)

        # 결과 영역 초기화
        for w in self.stock_result_frame.winfo_children():
            w.destroy()

        provider = self.stock_provider_var.get()

        def worker():
            try:
                result = engine.run_stock_analysis(
                    ticker_names=names,
                    provider=provider,
                    log_callback=lambda msg: self.log_queue.put(("stock_log", msg)),
                )
                self.log_queue.put(("__done_stock__", result))
            except Exception as e:
                self.log_queue.put(("stock_log", f"\n치명적 오류: {e}"))
                self.log_queue.put(("__done_stock__", None))

        threading.Thread(target=worker, daemon=True).start()

    def _on_stock_done(self, result):
        """종목 분석 완료 콜백"""
        self._running = False
        self.stock_run_btn.configure(state="normal")

        if not result:
            self.stock_status.configure(text="오류 발생")
            return

        total = result.get("total_analyzed", 0)
        a_count = len(result.get("a_rated", []))
        errors = result.get("total_errors", 0)
        usage = result.get("token_usage", {})
        cost_str = ""
        if usage.get("api_calls", 0) > 0:
            cost_usd = usage["total_cost_usd"]
            cost_krw = cost_usd * 1400
            cost_str = f" | 비용: ${cost_usd:.4f} ({cost_krw:.0f}원)"
        self.stock_status.configure(text=f"완료: {total}개 분석 | A-1/A-2 선정: {a_count}개 | 오류: {errors}개{cost_str}")

        # Word 내보내기 / 메일 발송 활성화
        self._last_stock_result = result
        self.stock_docx_btn.configure(state="normal")
        self.stock_email_btn.configure(state="normal")

        # 결과 카드 표시
        self._display_stock_results(result)

    def _display_stock_results(self, result):
        """분석 결과를 카드 형태로 표시"""
        parent = self.stock_result_frame

        # 결과 초기화
        for w in parent.winfo_children():
            w.destroy()

        all_results = result.get("results", [])
        if not all_results:
            ctk.CTkLabel(parent, text="분석 결과 없음", text_color="gray").pack(pady=20)
            return

        # A-1 우선, A-2 다음, 나머지 등급 순
        grade_order = {"A-1": 0, "A-2": 1, "B": 2, "C": 3, "D": 4, "N/A": 5}
        all_results.sort(key=lambda x: (grade_order.get(x.get("grade", "N/A"), 5), -x.get("confidence", 0)))

        for r in all_results:
            self._create_result_card(parent, r)

    def _create_result_card(self, parent, r):
        """개별 종목 결과 카드 생성 (차트 이미지 + 분석 결과)"""
        grade = r.get("grade", "N/A")
        grade_colors = {"A-1": "#C62828", "A-2": "#E65100", "B": "#1565C0", "C": "#757575", "D": "#9E9E9E", "N/A": "#9E9E9E"}
        color = grade_colors.get(grade, "#9E9E9E")

        card = ctk.CTkFrame(parent, fg_color=ACCENT_LIGHT, corner_radius=8)
        card.pack(fill="x", pady=4, padx=2)

        # 상단: 종목명 + 등급
        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", padx=10, pady=(8, 3))

        name = r.get("ticker_name", r.get("ticker", ""))
        ctk.CTkLabel(
            header, text=name,
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color=TEXT_DARK,
        ).pack(side="left")

        confidence = r.get("confidence", 0)
        ctk.CTkLabel(
            header, text=f"[{grade}] {confidence}%",
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color=color,
        ).pack(side="right")

        # 본문: 좌측 차트 이미지 + 우측 분석 텍스트
        body = ctk.CTkFrame(card, fg_color="transparent")
        body.pack(fill="x", padx=10, pady=(3, 8))

        # 좌측: 차트 이미지
        chart_path = r.get("chart_path", "")
        if chart_path and Path(chart_path).exists():
            try:
                pil_img = Image.open(chart_path)
                # 카드 내 적절한 크기로 축소
                chart_img = ctk.CTkImage(light_image=pil_img, size=(360, 240))
                img_label = ctk.CTkLabel(body, image=chart_img, text="")
                img_label.pack(side="left", padx=(0, 10), pady=2)
                # 참조 유지 (GC 방지)
                img_label._chart_img = chart_img
            except Exception:
                pass

        # 우측: 분석 텍스트
        info = ctk.CTkFrame(body, fg_color="transparent")
        info.pack(side="left", fill="both", expand=True)

        # 추세 정보
        trend = r.get("trend", {})
        if trend:
            direction = trend.get("direction", "")
            ma = trend.get("ma_arrangement", "")
            strength = trend.get("strength", "")
            trend_text = f"추세: {direction} | MA: {ma} | 강도: {strength}"
            ctk.CTkLabel(
                info, text=trend_text,
                font=ctk.CTkFont(size=11),
                text_color=TEXT_DARK,
            ).pack(anchor="w", pady=(0, 3))

        # 신호
        signals = r.get("signals", [])
        if signals:
            for sig in signals[:4]:
                ctk.CTkLabel(
                    info, text=f"  + {sig}",
                    font=ctk.CTkFont(size=11),
                    text_color="#2E7D32",
                    wraplength=400,
                ).pack(anchor="w", pady=0)

        # 리스크
        risks = r.get("risk_factors", [])
        if risks:
            for risk in risks[:3]:
                ctk.CTkLabel(
                    info, text=f"  - {risk}",
                    font=ctk.CTkFont(size=11),
                    text_color="#C62828",
                    wraplength=400,
                ).pack(anchor="w", pady=0)

        # 종합 의견 (전문)
        reasoning = r.get("reasoning", "")
        if reasoning:
            ctk.CTkLabel(
                info, text=reasoning,
                font=ctk.CTkFont(size=11),
                text_color="gray",
                wraplength=400,
                justify="left",
            ).pack(anchor="w", pady=(5, 0))

        # 목표/손절 구간
        target = r.get("price_target_zone", "")
        stoploss = r.get("stop_loss_zone", "")
        if target or stoploss:
            zones = ctk.CTkFrame(info, fg_color="transparent")
            zones.pack(anchor="w", pady=(3, 0))
            if target:
                ctk.CTkLabel(
                    zones, text=f"목표: {target}",
                    font=ctk.CTkFont(size=10),
                    text_color="#1565C0",
                ).pack(anchor="w")
            if stoploss:
                ctk.CTkLabel(
                    zones, text=f"손절: {stoploss}",
                    font=ctk.CTkFont(size=10),
                    text_color="#C62828",
                ).pack(anchor="w")

    def _export_stock_docx(self):
        """종목 분석 결과 Word 내보내기"""
        if not self._last_stock_result:
            return
        try:
            path = engine.save_results_docx(self._last_stock_result, "stocks")
            self._append_log(self.stock_log_box, f"\n[Word] 보고서 저장: {path.name}")
            self.stock_status.configure(text=f"Word 보고서 저장 완료: {path.name}")
            os.startfile(str(path))
        except Exception as e:
            self._append_log(self.stock_log_box, f"\n[X] Word 내보내기 실패: {e}")

    # ── 워치리스트 관리 ──

    def _on_list_changed(self, new_list_name):
        """드롭다운에서 리스트 변경 시"""
        # 현재 리스트 상태 저장
        self._save_gui_to_watchlist()
        # 새 리스트로 전환
        self._current_list_name = new_list_name
        engine.set_active_list(new_list_name)
        # GUI 종목 초기화 후 다시 로드
        self._reload_stock_list_gui()

    def _reload_stock_list_gui(self):
        """종목 리스트 GUI를 현재 리스트로 새로고침"""
        for e in self._stock_entries:
            e["widget"].destroy()
        self._stock_entries.clear()
        self._load_watchlist_to_gui()

    def _refresh_list_selector(self):
        """드롭다운 옵션 갱신"""
        names = engine.get_list_names()
        self.list_selector.configure(values=names)
        if self._current_list_name not in names:
            self._current_list_name = names[0] if names else "전체 종목"
        self.list_selector_var.set(self._current_list_name)

    def _create_new_list(self):
        """새 리스트 생성 다이얼로그"""
        dialog = ctk.CTkInputDialog(
            text="새 워치리스트 이름을 입력하세요:",
            title="워치리스트 생성",
        )
        name = dialog.get_input()
        if not name or not name.strip():
            return
        name = name.strip()
        if engine.create_list(name):
            self._save_gui_to_watchlist()
            self._current_list_name = name
            engine.set_active_list(name)
            self._refresh_list_selector()
            self._reload_stock_list_gui()
            self._append_log(self.stock_log_box, f"[리스트] '{name}' 생성 완료")
        else:
            self._append_log(self.stock_log_box, f"[!] '{name}' 리스트가 이미 존재합니다.")

    def _open_list_config(self):
        """현재 리스트의 설정 다이얼로그"""
        lst = engine.get_list(self._current_list_name)
        if not lst:
            return

        win = ctk.CTkToplevel(self)
        win.title(f"워치리스트 설정 - {self._current_list_name}")
        win.geometry("420x420")
        win.resizable(False, False)
        win.grab_set()

        config = lst.get("config", {})

        pad = {"padx": 15, "pady": (8, 0)}

        # 리스트 이름
        ctk.CTkLabel(win, text="리스트 이름", font=ctk.CTkFont(weight="bold")).pack(anchor="w", **pad)
        name_entry = ctk.CTkEntry(win, border_color=BORDER_COLOR)
        name_entry.pack(fill="x", padx=15, pady=(2, 0))
        name_entry.insert(0, self._current_list_name)

        # LLM 프로바이더
        ctk.CTkLabel(win, text="LLM 프로바이더", font=ctk.CTkFont(weight="bold")).pack(anchor="w", **pad)
        provider_var = ctk.StringVar(value=config.get("provider") or "기본값")
        ctk.CTkOptionMenu(
            win, variable=provider_var,
            values=["기본값", "claude", "gemini"],
            fg_color=ACCENT, button_color=ACCENT_DARK,
        ).pack(fill="x", padx=15, pady=(2, 0))

        # 이메일 수신자
        ctk.CTkLabel(win, text="이메일 수신자", font=ctk.CTkFont(weight="bold")).pack(anchor="w", **pad)
        email_entry = ctk.CTkEntry(win, placeholder_text="recipient@example.com", border_color=BORDER_COLOR)
        email_entry.pack(fill="x", padx=15, pady=(2, 0))
        if config.get("email_to"):
            email_entry.insert(0, config["email_to"])

        # Notion Watchlist DB
        ctk.CTkLabel(win, text="Notion 종목 DB ID", font=ctk.CTkFont(weight="bold")).pack(anchor="w", **pad)
        notion_wl_entry = ctk.CTkEntry(win, placeholder_text="Notion DB ID (선택)", border_color=BORDER_COLOR)
        notion_wl_entry.pack(fill="x", padx=15, pady=(2, 0))
        if config.get("notion_watchlist_db"):
            notion_wl_entry.insert(0, config["notion_watchlist_db"])

        # Notion Report DB
        ctk.CTkLabel(win, text="Notion 보고서 DB ID", font=ctk.CTkFont(weight="bold")).pack(anchor="w", **pad)
        notion_rp_entry = ctk.CTkEntry(win, placeholder_text="Notion DB ID (선택)", border_color=BORDER_COLOR)
        notion_rp_entry.pack(fill="x", padx=15, pady=(2, 0))
        if config.get("notion_report_db"):
            notion_rp_entry.insert(0, config["notion_report_db"])

        # 버튼 영역
        btn_frame = ctk.CTkFrame(win, fg_color="transparent")
        btn_frame.pack(fill="x", padx=15, pady=(20, 10))

        def _save_config():
            new_name = name_entry.get().strip()
            if not new_name:
                return

            # 이름 변경
            if new_name != self._current_list_name:
                if engine.rename_list(self._current_list_name, new_name):
                    self._current_list_name = new_name
                    self._refresh_list_selector()

            # 설정 저장
            prov = provider_var.get()
            engine.update_list_config(self._current_list_name, "provider", None if prov == "기본값" else prov)
            engine.update_list_config(self._current_list_name, "email_to", email_entry.get().strip() or None)
            engine.update_list_config(self._current_list_name, "notion_watchlist_db", notion_wl_entry.get().strip() or None)
            engine.update_list_config(self._current_list_name, "notion_report_db", notion_rp_entry.get().strip() or None)

            win.destroy()
            self._append_log(self.stock_log_box, f"[설정] '{self._current_list_name}' 설정 저장 완료")

        def _delete_list():
            if engine.delete_list(self._current_list_name):
                self._current_list_name = engine.get_list_names()[0]
                engine.set_active_list(self._current_list_name)
                self._refresh_list_selector()
                self._reload_stock_list_gui()
                win.destroy()
                self._append_log(self.stock_log_box, "[리스트] 삭제 완료")
            else:
                self._append_log(self.stock_log_box, "[!] 마지막 리스트는 삭제할 수 없습니다.")

        self._orange_btn(btn_frame, text="저장", width=100, command=_save_config).pack(side="left", padx=(0, 10))
        self._orange_btn(
            btn_frame, text="리스트 삭제", width=100,
            fg_color="#C62828", hover_color="#B71C1C",
            command=_delete_list,
        ).pack(side="right")

    def _send_stock_email(self):
        """분석 결과를 이메일로 발송"""
        if not self._last_stock_result:
            return

        to_addr = engine.get_list_config(self._current_list_name, "email_to")
        if not to_addr:
            self._append_log(self.stock_log_box, "[!] 이메일 수신자가 설정되지 않았습니다. 리스트 설정(⚙)에서 지정하세요.")
            return

        self.stock_email_btn.configure(state="disabled")
        self._append_log(self.stock_log_box, f"\n[메일] {to_addr}로 발송 중...")

        def worker():
            try:
                provider = self.stock_provider_var.get()
                subject, body = engine.build_email_body(self._last_stock_result, provider)

                # Word 보고서 생성
                docx_path = engine.save_results_docx(self._last_stock_result, "stocks")

                engine.send_report_email(
                    to_addr=to_addr,
                    subject=subject,
                    body_text=body,
                    attachment_path=str(docx_path),
                )
                self.log_queue.put(("stock_log", f"[메일] {to_addr}로 발송 완료"))
            except Exception as e:
                self.log_queue.put(("stock_log", f"[X] 메일 발송 실패: {e}"))
            finally:
                self.after(0, lambda: self.stock_email_btn.configure(state="normal"))

        threading.Thread(target=worker, daemon=True).start()

    # ============================================================
    # 탭 2: 테마 분석
    # ============================================================
    def _build_theme_tab(self):
        tab = self.tabview.add("테마 분석")

        content = ctk.CTkFrame(tab, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=5, pady=5)

        # ── 좌측: 테마 선택 ──
        left = ctk.CTkFrame(content, width=280, fg_color="transparent")
        left.pack(side="left", fill="y", padx=(0, 10))
        left.pack_propagate(False)

        self._section_label(left, "투자 테마").pack(anchor="w", pady=(0, 5))

        # 전체 선택/해제
        sel_frame = ctk.CTkFrame(left, fg_color="transparent")
        sel_frame.pack(fill="x", pady=(0, 5))

        self._orange_btn(
            sel_frame, text="전체 선택", width=90,
            command=lambda: self._toggle_themes(True),
        ).pack(side="left", padx=(0, 5))
        self._orange_btn(
            sel_frame, text="전체 해제", width=90,
            command=lambda: self._toggle_themes(False),
        ).pack(side="left")

        # 테마 체크박스 리스트
        theme_scroll = ctk.CTkScrollableFrame(left, fg_color="transparent")
        theme_scroll.pack(fill="both", expand=True, pady=(0, 10))

        try:
            themes = engine.load_themes()
        except Exception:
            themes = []

        for t in themes:
            var = ctk.BooleanVar(value=True)
            desc = t.get("description", "")
            cb = ctk.CTkCheckBox(
                theme_scroll, text=f"{t['name']} ({desc})",
                variable=var,
                fg_color=ACCENT, hover_color=ACCENT_HOVER,
            )
            cb.pack(anchor="w", pady=3)
            self._theme_checkboxes[t["name"]] = var

        # 설정
        self._section_label(left, "설정").pack(anchor="w", pady=(10, 5))

        # Top N 설정
        topn_frame = ctk.CTkFrame(left, fg_color="transparent")
        topn_frame.pack(fill="x", pady=(0, 5))

        ctk.CTkLabel(topn_frame, text="보유종목 수:", text_color=TEXT_DARK).pack(side="left")
        self.theme_topn_entry = ctk.CTkEntry(
            topn_frame, width=60,
            border_color=BORDER_COLOR,
        )
        self.theme_topn_entry.insert(0, str(engine.ETF_TOP_N))
        self.theme_topn_entry.pack(side="left", padx=5)

        # LLM 선택
        self.theme_provider_var = ctk.StringVar(value=engine.LLM_PROVIDER)
        for label, val in [("Claude", "claude"), ("Gemini", "gemini")]:
            ctk.CTkRadioButton(
                left, text=label, variable=self.theme_provider_var, value=val,
                fg_color=ACCENT, hover_color=ACCENT_HOVER,
            ).pack(anchor="w", pady=2)

        # 테마 캐시 갱신 버튼
        self.theme_refresh_btn = self._orange_btn(
            left, text="테마 캐시 갱신 (LLM)", height=32,
            fg_color="#6B8E23", hover_color="#556B2F",
            command=self._run_refresh_themes,
        )
        self.theme_refresh_btn.pack(fill="x", pady=(10, 5))

        # 실행 버튼
        self.theme_run_btn = self._orange_btn(
            left, text="테마 분석 실행", height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._run_theme_analysis,
        )
        self.theme_run_btn.pack(fill="x", pady=(5, 5))

        # Word 내보내기 버튼
        self.theme_docx_btn = self._orange_btn(
            left, text="Word 보고서 내보내기", height=32,
            fg_color="#4A90D9", hover_color="#3A78C2",
            command=self._export_theme_docx,
            state="disabled",
        )
        self.theme_docx_btn.pack(fill="x", pady=(5, 5))

        # 상태
        self.theme_status = ctk.CTkLabel(
            left, text="대기 중",
            text_color=ACCENT_DARK,
            font=ctk.CTkFont(size=12),
        )
        self.theme_status.pack(anchor="w")

        # 마지막 테마 분석 결과 저장
        self._last_theme_result = None

        # ── 우측: 결과 + 로그 ──
        right = ctk.CTkFrame(content, fg_color="transparent")
        right.pack(side="left", fill="both", expand=True)

        self._section_label(right, "테마별 분석 결과").pack(anchor="w", pady=(0, 5))

        self.theme_result_frame = ctk.CTkScrollableFrame(
            right, height=300, fg_color="transparent",
        )
        self.theme_result_frame.pack(fill="both", expand=True, pady=(0, 10))

        self._section_label(right, "실행 로그").pack(anchor="w", pady=(0, 5))

        self.theme_log_box = ctk.CTkTextbox(
            right, height=200, state="disabled",
            font=ctk.CTkFont(family="Consolas", size=11),
            wrap="word",
        )
        self.theme_log_box.pack(fill="both", expand=True)

    def _toggle_themes(self, value):
        for var in self._theme_checkboxes.values():
            var.set(value)

    def _run_theme_analysis(self):
        """테마 분석 실행 (백그라운드 스레드)"""
        if self._running:
            return

        selected = [name for name, var in self._theme_checkboxes.items() if var.get()]
        if not selected:
            self._append_log(self.theme_log_box, "[!] 분석할 테마를 선택하세요.")
            return

        self._running = True
        self.theme_run_btn.configure(state="disabled")
        self.theme_status.configure(text="분석 중...")
        self._clear_log(self.theme_log_box)

        for w in self.theme_result_frame.winfo_children():
            w.destroy()

        provider = self.theme_provider_var.get()
        try:
            top_n = int(self.theme_topn_entry.get())
        except ValueError:
            top_n = engine.ETF_TOP_N

        def worker():
            try:
                result = engine.run_theme_analysis(
                    theme_names=selected,
                    top_n=top_n,
                    provider=provider,
                    log_callback=lambda msg: self.log_queue.put(("theme_log", msg)),
                )
                self.log_queue.put(("__done_theme__", result))
            except Exception as e:
                self.log_queue.put(("theme_log", f"\n치명적 오류: {e}"))
                self.log_queue.put(("__done_theme__", None))

        threading.Thread(target=worker, daemon=True).start()

    def _on_theme_done(self, result):
        """테마 분석 완료 콜백"""
        self._running = False
        self.theme_run_btn.configure(state="normal")

        if not result:
            self.theme_status.configure(text="오류 발생")
            return

        themes = result.get("themes", [])
        total_a = sum(len(t.get("a_rated", [])) for t in themes)
        self.theme_status.configure(text=f"완료: {len(themes)}개 테마 | 전체 A-1/A-2 선정: {total_a}개")

        # Word 내보내기 활성화
        self._last_theme_result = result
        self.theme_docx_btn.configure(state="normal")

        # 테마별 결과 표시
        self._display_theme_results(result)

    def _display_theme_results(self, result):
        """테마별 분석 결과 표시"""
        parent = self.theme_result_frame

        for w in parent.winfo_children():
            w.destroy()

        themes = result.get("themes", [])
        if not themes:
            ctk.CTkLabel(parent, text="분석 결과 없음", text_color="gray").pack(pady=20)
            return

        for tr in themes:
            theme_name = tr.get("theme", "")
            etf = tr.get("etf", {})
            a_rated = tr.get("a_rated", [])
            all_analyzed = tr.get("holdings_analyzed", [])
            error = tr.get("error", "")

            # 테마 카드
            card = ctk.CTkFrame(parent, fg_color=ACCENT_LIGHT, corner_radius=8)
            card.pack(fill="x", pady=4, padx=2)

            # 테마 헤더
            header = ctk.CTkFrame(card, fg_color="transparent")
            header.pack(fill="x", padx=10, pady=(8, 3))

            ctk.CTkLabel(
                header, text=f"테마: {theme_name}",
                font=ctk.CTkFont(size=14, weight="bold"),
                text_color=ACCENT_DARK,
            ).pack(side="left")

            if a_rated:
                ctk.CTkLabel(
                    header, text=f"A-1/A-2 {len(a_rated)}개",
                    font=ctk.CTkFont(size=13, weight="bold"),
                    text_color="#C62828",
                ).pack(side="right")

            # ETF 정보
            if etf:
                etf_name = etf.get("etf_name", "")
                etf_code = etf.get("ticker_code", "")
                ctk.CTkLabel(
                    card, text=f"ETF: {etf_name} ({etf_code})",
                    font=ctk.CTkFont(size=11),
                    text_color=TEXT_DARK,
                ).pack(anchor="w", padx=10, pady=(0, 2))

            if error:
                ctk.CTkLabel(
                    card, text=f"오류: {error}",
                    text_color="#C62828",
                    font=ctk.CTkFont(size=11),
                ).pack(anchor="w", padx=10, pady=(0, 5))
                continue

            # 등급별 종목 표시
            grade_groups = {"A-1": [], "A-2": [], "B": [], "C": [], "D": [], "N/A": []}
            for r in all_analyzed:
                g = r.get("grade", "N/A")
                name = r.get("ticker_name", r.get("ticker", ""))
                if g in grade_groups:
                    grade_groups[g].append(name)
                else:
                    grade_groups["N/A"].append(name)

            grade_colors = {"A-1": "#C62828", "A-2": "#E65100", "B": "#1565C0", "C": "#757575", "D": "#9E9E9E"}
            for g in ["A-1", "A-2", "B", "C", "D"]:
                names = grade_groups[g]
                if names:
                    ctk.CTkLabel(
                        card,
                        text=f"{g}: {', '.join(names)}",
                        font=ctk.CTkFont(size=11),
                        text_color=grade_colors[g],
                        wraplength=600,
                    ).pack(anchor="w", padx=10, pady=(0, 2))

            # 하단 패딩
            ctk.CTkFrame(card, fg_color="transparent", height=5).pack()

    def _run_refresh_themes(self):
        """테마 ETF 캐시 갱신 (백그라운드 스레드)"""
        if self._running:
            return

        self._running = True
        self.theme_refresh_btn.configure(state="disabled")
        self.theme_run_btn.configure(state="disabled")
        self.theme_status.configure(text="테마 캐시 갱신 중...")
        self._clear_log(self.theme_log_box)

        provider = self.theme_provider_var.get()

        def worker():
            try:
                count = engine.refresh_theme_cache(
                    provider=provider,
                    log_callback=lambda msg: self.log_queue.put(("theme_log", msg)),
                )
                self.log_queue.put(("__done_refresh_themes__", count))
            except Exception as e:
                self.log_queue.put(("theme_log", f"\n치명적 오류: {e}"))
                self.log_queue.put(("__done_refresh_themes__", None))

        threading.Thread(target=worker, daemon=True).start()

    def _on_refresh_themes_done(self, count):
        """테마 캐시 갱신 완료 콜백"""
        self._running = False
        self.theme_refresh_btn.configure(state="normal")
        self.theme_run_btn.configure(state="normal")

        if count is not None:
            self.theme_status.configure(text=f"테마 캐시 갱신 완료: {count}개 테마")
        else:
            self.theme_status.configure(text="테마 캐시 갱신 실패")

    def _export_theme_docx(self):
        """테마 분석 결과 Word 내보내기"""
        if not self._last_theme_result:
            return
        try:
            path = engine.save_results_docx(self._last_theme_result, "themes")
            self._append_log(self.theme_log_box, f"\n[Word] 보고서 저장: {path.name}")
            self.theme_status.configure(text=f"Word 보고서 저장 완료: {path.name}")
            os.startfile(str(path))
        except Exception as e:
            self._append_log(self.theme_log_box, f"\n[X] Word 내보내기 실패: {e}")

    # ============================================================
    # 탭 3: 결과 조회
    # ============================================================
    def _build_result_tab(self):
        tab = self.tabview.add("결과 조회")

        content = ctk.CTkFrame(tab, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=5, pady=5)

        # ── 좌측: 결과 파일 목록 ──
        left = ctk.CTkFrame(content, width=280, fg_color="transparent")
        left.pack(side="left", fill="y", padx=(0, 10))
        left.pack_propagate(False)

        header = ctk.CTkFrame(left, fg_color="transparent")
        header.pack(fill="x", pady=(0, 5))

        self._section_label(header, "분석 결과").pack(side="left")
        self._orange_btn(
            header, text="새로고침", width=80,
            command=self._refresh_result_list,
        ).pack(side="right")

        # Word 내보내기 버튼
        self.result_docx_btn = self._orange_btn(
            left, text="선택 결과 Word 내보내기", height=32,
            fg_color="#4A90D9", hover_color="#3A78C2",
            command=self._export_result_docx,
            state="disabled",
        )
        self.result_docx_btn.pack(fill="x", pady=(5, 0))

        self.result_list_frame = ctk.CTkScrollableFrame(
            left, fg_color="transparent",
        )
        self.result_list_frame.pack(fill="both", expand=True)

        # ── 우측: 결과 상세 ──
        right = ctk.CTkFrame(content, fg_color="transparent")
        right.pack(side="left", fill="both", expand=True)

        # 서브탭
        self.result_subtab = ctk.CTkTabview(
            right,
            segmented_button_fg_color=ACCENT_LIGHT,
            segmented_button_selected_color=ACCENT,
            segmented_button_selected_hover_color=ACCENT_HOVER,
            segmented_button_unselected_color=ACCENT_LIGHT,
            segmented_button_unselected_hover_color=BORDER_COLOR,
            text_color=TEXT_DARK,
        )
        self.result_subtab.pack(fill="both", expand=True)

        # A-1/A-2 선정 종목 탭
        a_tab = self.result_subtab.add("A-1/A-2 선정")
        self.result_a_frame = ctk.CTkScrollableFrame(a_tab, fg_color="transparent")
        self.result_a_frame.pack(fill="both", expand=True)

        # 전체 결과 탭
        all_tab = self.result_subtab.add("전체 결과")
        self.result_all_frame = ctk.CTkScrollableFrame(all_tab, fg_color="transparent")
        self.result_all_frame.pack(fill="both", expand=True)

        # 원본 JSON 탭
        json_tab = self.result_subtab.add("원본 JSON")
        self.result_json_box = ctk.CTkTextbox(
            json_tab, state="disabled",
            font=ctk.CTkFont(family="Consolas", size=11),
            wrap="none",
        )
        self.result_json_box.pack(fill="both", expand=True)

        # 현재 로드된 결과 데이터
        self._loaded_result_data = None

        # 초기 로드
        self._refresh_result_list()

    def _refresh_result_list(self):
        """결과 파일 목록 새로고침"""
        for w in self.result_list_frame.winfo_children():
            w.destroy()

        if not engine.RESULTS_DIR.exists():
            return

        files = sorted(engine.RESULTS_DIR.glob("*.json"), reverse=True)
        self._result_files = files

        if not files:
            ctk.CTkLabel(
                self.result_list_frame, text="결과 파일 없음",
                text_color="gray",
            ).pack(pady=20)
            return

        for f in files[:50]:  # 최대 50개 표시
            btn = ctk.CTkButton(
                self.result_list_frame,
                text=f.stem,
                fg_color="transparent",
                text_color=TEXT_DARK,
                hover_color=ACCENT_LIGHT,
                anchor="w",
                command=lambda path=f: self._load_result_file(path),
            )
            btn.pack(fill="x", pady=1)

    def _load_result_file(self, path):
        """결과 파일 로드 및 표시"""
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            self._clear_log(self.result_json_box)
            self._append_log(self.result_json_box, f"파일 로드 오류: {e}")
            return

        # 로드된 데이터 저장 + Word 버튼 활성화
        self._loaded_result_data = data
        self.result_docx_btn.configure(state="normal")

        # 원본 JSON 표시
        self.result_json_box.configure(state="normal")
        self.result_json_box.delete("1.0", "end")
        self.result_json_box.insert("1.0", json.dumps(data, ensure_ascii=False, indent=2))
        self.result_json_box.configure(state="disabled")

        # A-1/A-2 선정 종목 표시
        for w in self.result_a_frame.winfo_children():
            w.destroy()
        for w in self.result_all_frame.winfo_children():
            w.destroy()

        # stocks 결과인지 themes 결과인지 판별
        if "results" in data:
            # 종목 분석 결과
            all_results = data.get("results", [])
            a_rated = data.get("a_rated", [])
        elif "themes" in data:
            # 테마 분석 결과 - 모든 테마의 결과 취합
            all_results = []
            a_rated = []
            for t in data.get("themes", []):
                all_results.extend(t.get("holdings_analyzed", []))
                a_rated.extend(t.get("a_rated", []))
        else:
            all_results = []
            a_rated = []

        # A-1/A-2 선정 탭
        if a_rated:
            for idx, r in enumerate(a_rated, 1):
                name = r.get("ticker_name", r.get("ticker", ""))
                grade = r.get("grade", "")
                reliability = r.get("reliability", "")
                reasoning = r.get("reasoning", "")[:80]

                text = f"{idx}. [{grade}] {name} (신뢰도: {reliability}) - {reasoning}..."
                ctk.CTkLabel(
                    self.result_a_frame, text=text,
                    font=ctk.CTkFont(size=12),
                    text_color=TEXT_DARK,
                    wraplength=700,
                    anchor="w", justify="left",
                ).pack(anchor="w", pady=3, padx=5)
        else:
            ctk.CTkLabel(
                self.result_a_frame, text="A-1/A-2 선정 종목 없음",
                text_color="gray",
            ).pack(pady=20)

        # 전체 결과 탭
        if all_results:
            for r in all_results:
                self._create_result_card(self.result_all_frame, r)
        else:
            ctk.CTkLabel(
                self.result_all_frame, text="분석 결과 없음",
                text_color="gray",
            ).pack(pady=20)


    def _export_result_docx(self):
        """결과 조회 탭에서 Word 내보내기"""
        if not self._loaded_result_data:
            return
        try:
            prefix = "themes" if "themes" in self._loaded_result_data else "stocks"
            path = engine.save_results_docx(self._loaded_result_data, prefix)
            os.startfile(str(path))
        except Exception as e:
            self.result_json_box.configure(state="normal")
            self.result_json_box.insert("end", f"\n\n[X] Word 내보내기 실패: {e}")
            self.result_json_box.configure(state="disabled")


if __name__ == "__main__":
    app = App()
    app.mainloop()
