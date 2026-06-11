"""
gmarket_air_gui.py
G마켓 항공료 자동 추출 - GUI 실행기
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import os
import datetime
import sys

# 런처(exe)로 실행 시 scraper_core 경로 확보
_here = os.path.dirname(os.path.abspath(__file__)) if "__file__" in dir() else ""
if getattr(sys, "frozen", False):
    _exe_dir = os.path.dirname(sys.executable)
    if _exe_dir not in sys.path:
        sys.path.insert(0, _exe_dir)
elif _here and _here not in sys.path:
    sys.path.insert(0, _here)

# scraper_core가 같은 폴더에 있어야 함
try:
    from scraper_core import (
        AIRPORTS, AIRPORT_CODES, LCC_ALL, FSC_ALL, LCC_TIER1, LCC_TIER2,
        FOREIGN_ALL, collect_monthly, save_excel
    )
    SCRAPER_OK = True
except ImportError:
    SCRAPER_OK = False
    AIRPORTS = {
        "국내": {"인천":"ICN","부산":"PUS","김포":"GMP","청주":"CJJ","대구":"TAE"},
        "일본": {"삿포로":"CTS","도쿄":"TYO","오사카":"OSA","후쿠오카":"FUK","오키나와":"OKA"},
    }
    LCC_ALL = FSC_ALL = LCC_TIER1 = LCC_TIER2 = FOREIGN_ALL = []


# ─────────────────────────────────────────────
#  색상 / 폰트 상수
# ─────────────────────────────────────────────
C_BG       = "#F5F7FA"
C_PANEL    = "#FFFFFF"
C_ACCENT   = "#1F4E79"
C_ACCENT2  = "#2E75B6"
C_BTN      = "#2E75B6"
C_BTN_FG   = "#FFFFFF"
C_BTN_RUN  = "#217346"
C_WARN     = "#C00000"
C_BORDER   = "#D0D7E3"
C_LOG_BG   = "#1E1E2E"
C_LOG_FG   = "#CDD6F4"
C_LOG_OK   = "#A6E3A1"
C_LOG_ERR  = "#F38BA8"
C_LOG_INFO = "#89B4FA"

FONT_TITLE  = ("맑은 고딕", 15, "bold")
FONT_SUB    = ("맑은 고딕", 10, "bold")
FONT_BODY   = ("맑은 고딕", 10)
FONT_SMALL  = ("맑은 고딕", 9)
FONT_LOG    = ("Consolas", 9)

AIRLINE_MODES = [
    "LCC 우선 → FSC 대체",
    "LCC만",
    "FSC만 (아시아나/대한항공)",
    "외항사만",
    "특정 항공사 지정",
]

TIME_SLOT_LABELS = ["새벽 (00~06시)", "오전 (06~12시)", "오후 (12~18시)", "야간 (18~24시)"]
TIME_SLOT_KEYS   = ["새벽", "오전", "오후", "야간"]

SPECIFIC_AIRLINES = [
    "대한항공", "아시아나항공", "에어프레미아", "진에어", "제주항공",
    "티웨이항공", "이스타항공", "에어부산", "에어서울", "에어로케이",
    "파라타항공", "타이항공", "필리핀항공", "베트남항공",
]


# ─────────────────────────────────────────────
#  유틸
# ─────────────────────────────────────────────
def get_default_output_dir():
    return os.path.join(os.path.expanduser("~"), "Desktop")


def roundup_label(widget, text, bg, font=FONT_BODY, fg="#333333"):
    f = tk.Frame(widget, bg=bg)
    tk.Label(f, text=text, bg=bg, fg=fg, font=font).pack(side="left")
    return f


# ─────────────────────────────────────────────
#  메인 앱
# ─────────────────────────────────────────────
class GmarketAirApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("G마켓 항공료 자동 추출기  |  투어로 2팀")
        self.geometry("1180x820")
        self.resizable(True, True)
        self.configure(bg=C_BG)
        self._running = False
        self._thread  = None
        self._build_ui()

    # ──────── UI 구성 ────────
    def _build_ui(self):
        # 타이틀 배너
        banner = tk.Frame(self, bg=C_ACCENT, height=54)
        banner.pack(fill="x")
        banner.pack_propagate(False)
        tk.Label(banner, text="✈  G마켓 항공료 자동 추출기",
                 bg=C_ACCENT, fg="white", font=FONT_TITLE).pack(side="left", padx=18, pady=10)
        tk.Label(banner, text="투어로 2팀",
                 bg=C_ACCENT, fg="#BDD7EE", font=FONT_SMALL).pack(side="right", padx=18)

        # 스크롤 가능 메인 영역 (세로 + 가로)
        canvas_wrap = tk.Frame(self, bg=C_BG)
        canvas_wrap.pack(fill="both", expand=True)

        main_canvas = tk.Canvas(canvas_wrap, bg=C_BG, highlightthickness=0)
        scrollbar   = ttk.Scrollbar(canvas_wrap, orient="vertical", command=main_canvas.yview)
        scrollbar_x = ttk.Scrollbar(canvas_wrap, orient="horizontal", command=main_canvas.xview)
        self._scroll_frame = tk.Frame(main_canvas, bg=C_BG)
        self._scroll_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )
        main_canvas.create_window((0, 0), window=self._scroll_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set, xscrollcommand=scrollbar_x.set)
        scrollbar.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        main_canvas.pack(side="left", fill="both", expand=True)
        main_canvas.bind_all("<MouseWheel>", lambda e: main_canvas.yview_scroll(-1*(e.delta//120), "units"))
        main_canvas.bind_all("<Shift-MouseWheel>", lambda e: main_canvas.xview_scroll(-1*(e.delta//120), "units"))

        pad = {"padx": 16, "pady": 6}

        # ── 섹션 1: 노선 ──
        self._add_section_header("📍  1. 노선 설정")
        row1 = self._panel()

        # 출발지
        tk.Label(row1, text="출발지", bg=C_PANEL, font=FONT_SUB).grid(row=0, column=0, sticky="w", **pad)
        self.origin_country_var = tk.StringVar(value="국내")
        origin_country_cb = ttk.Combobox(row1, textvariable=self.origin_country_var,
                     values=list(AIRPORTS.keys()), state="readonly", width=8)
        origin_country_cb.grid(row=0, column=1, **pad)
        self.origin_city_var = tk.StringVar(value="인천")
        self._origin_city_cb = ttk.Combobox(row1, textvariable=self.origin_city_var,
                     values=list(AIRPORTS["국내"].keys()), state="readonly", width=10)
        self._origin_city_cb.grid(row=0, column=2, **pad)
        origin_country_cb.bind("<<ComboboxSelected>>", lambda e: self._on_country_change("origin"))

        tk.Label(row1, text="→", bg=C_PANEL, font=("맑은 고딕", 14)).grid(row=0, column=3)

        # 목적지
        tk.Label(row1, text="목적지", bg=C_PANEL, font=FONT_SUB).grid(row=0, column=4, sticky="w", **pad)
        self.dest_country_var = tk.StringVar(value="일본")
        dest_country_cb = ttk.Combobox(row1, textvariable=self.dest_country_var,
                     values=list(AIRPORTS.keys()), state="readonly", width=8)
        dest_country_cb.grid(row=0, column=5, **pad)
        self.dest_city_var = tk.StringVar(value="삿포로")
        self._dest_city_cb = ttk.Combobox(row1, textvariable=self.dest_city_var,
                     values=list(AIRPORTS["일본"].keys()), state="readonly", width=10)
        self._dest_city_cb.grid(row=0, column=6, **pad)
        dest_country_cb.bind("<<ComboboxSelected>>", lambda e: self._on_country_change("dest"))

        # ── 섹션 2: 기간 ──
        self._add_section_header("📅  2. 출발 월 / 체류 일수")
        row2 = self._panel()
        now = datetime.date.today()
        # 출발일 범위 (YYYY-MM-DD)
        default_y = now.year
        default_m = now.month + 1 if now.month < 12 else 1
        tk.Label(row2, text="출발일", bg=C_PANEL, font=FONT_SUB).grid(row=0, column=0, sticky="w", **pad)
        self.date_from_var = tk.StringVar(value=f"{default_y}-{default_m:02d}-01")
        tk.Entry(row2, textvariable=self.date_from_var, width=12, font=FONT_BODY,
                 relief="solid", bd=1, bg="#F8F9FB").grid(row=0, column=1, **pad)
        tk.Label(row2, text="~", bg=C_PANEL, font=("맑은 고딕", 12)).grid(row=0, column=2)
        self.date_to_var = tk.StringVar(value=f"{default_y}-{default_m:02d}-28")
        tk.Entry(row2, textvariable=self.date_to_var, width=12, font=FONT_BODY,
                 relief="solid", bd=1, bg="#F8F9FB").grid(row=0, column=3, **pad)

        # 귀국: 현지출발 출발 +N일
        tk.Label(row2, text="여행 일수", bg=C_PANEL, font=FONT_SUB).grid(row=0, column=4, sticky="w", **pad)
        self.return_offset_var = tk.StringVar(value="3")
        ro_spin = ttk.Spinbox(row2, from_=1, to=20, textvariable=self.return_offset_var,
                    width=4, font=FONT_BODY)
        ro_spin.grid(row=0, column=5, padx=2)
        tk.Label(row2, text="일", bg=C_PANEL, font=FONT_BODY).grid(row=0, column=6, sticky="w")

        # 추가 귀국일 검색 (동남아 새벽도착 대응)
        self.extra_return_var = tk.BooleanVar(value=False)
        tk.Checkbutton(row2, text="-1일도 함께 검색 (귀국 현지출발이 하루 빠른 새벽도착편 비교용)",
                       variable=self.extra_return_var, bg=C_PANEL, font=FONT_SMALL,
                       activebackground=C_PANEL, command=self._update_return_hint
                       ).grid(row=1, column=0, columnspan=9, sticky="w", padx=14, pady=(2,0))

        # 실시간 안내문 (검색되는 귀국 ADATE 미리보기)
        self._return_hint = tk.Label(row2, text="", bg=C_PANEL, fg=C_ACCENT2,
                                     font=("맑은 고딕", 8))
        self._return_hint.grid(row=2, column=0, columnspan=9, sticky="w", padx=16, pady=(0,2))

        # 안내문
        tk.Label(row2, text="(귀국일은 현지 출발일 기준. 같은 한국일정도 새벽도착편은 현지출발이 하루 빨라 -1일 체크 권장. 날짜형식: 2026-07-10)",
                 bg=C_PANEL, fg="#999999", font=("맑은 고딕", 8)
                 ).grid(row=3, column=0, columnspan=9, sticky="w", padx=16, pady=(0,4))

        # 입력 변경 시 안내 갱신
        self.return_offset_var.trace_add("write", lambda *a: self._update_return_hint())
        self.date_from_var.trace_add("write", lambda *a: self._update_return_hint())
        self.after(100, self._update_return_hint)

        # ── 섹션 2-1: 대체 조건 자동 확장 ──
        self._add_section_header("🔄  2-1. 대체 조건 자동 확장 (선택)")
        expand_panel = self._panel()

        self.use_expand_var = tk.BooleanVar(value=False)
        expand_toggle = tk.Checkbutton(
            expand_panel, text="사용 (체크 시 아래 순위 조건으로 검색, 위 일수 설정 무시)",
            variable=self.use_expand_var, bg=C_PANEL, font=FONT_BODY,
            activebackground=C_PANEL, command=self._on_expand_toggle
        )
        expand_toggle.grid(row=0, column=0, sticky="w", padx=16, pady=(8,4))

        # 순위 테이블 프레임
        self._expand_frame = tk.Frame(expand_panel, bg=C_PANEL)
        self._expand_frame.grid(row=1, column=0, sticky="ew", padx=16, pady=(0,8))

        # 헤더
        for ci, hdr in enumerate(["순위", "귀국일수(일)", "도착 ~(상한)", "귀국출발 (이후~이전)", "도착 시간대(상한 비울 때)", ""]):
            tk.Label(self._expand_frame, text=hdr, bg=C_PANEL, font=FONT_SUB,
                     fg=C_ACCENT).grid(row=0, column=ci, padx=8, pady=(4,2), sticky="w")

        self._expand_rows = []   # list of dict per priority row
        self._expand_row_frames = []

        # 기본 순위 2개
        self._add_expand_row(nights_default="2", arr_to="10:10", ret_dep_from="19:00", ret_dep_to="", arr_band_default="전체")
        self._add_expand_row(nights_default="3", arr_to="10:10", ret_dep_from="19:00", ret_dep_to="", arr_band_default="전체")
        self._add_expand_row(nights_default="3", arr_to="",      ret_dep_from="19:00", ret_dep_to="", arr_band_default="오후+야간")

        # + 순위 추가 버튼
        self._add_row_btn = tk.Button(
            expand_panel, text="＋ 순위 추가", command=self._add_expand_row,
            bg="#E8F0FE", fg=C_ACCENT, font=FONT_SMALL, relief="flat", cursor="hand2", padx=8
        )
        self._add_row_btn.grid(row=2, column=0, sticky="w", padx=16, pady=(0,8))

        # 처음엔 숨김
        self._expand_frame.grid_remove()
        self._add_row_btn.grid_remove()

        # ── 섹션 3: 항공사 ──
        self._add_section_header("✈  3. 항공사 조건")
        row3 = self._panel()
        tk.Label(row3, text="항공사 모드", bg=C_PANEL, font=FONT_SUB).grid(row=0, column=0, sticky="w", **pad)
        self.airline_mode_var = tk.StringVar(value=AIRLINE_MODES[0])
        mode_cb = ttk.Combobox(row3, textvariable=self.airline_mode_var,
                               values=AIRLINE_MODES, state="readonly", width=26)
        mode_cb.grid(row=0, column=1, **pad)
        mode_cb.bind("<<ComboboxSelected>>", self._on_mode_change)

        # 특정 항공사 선택 프레임 (여러 줄로 배치)
        self._specific_frame = tk.Frame(row3, bg=C_PANEL)
        self._specific_frame.grid(row=1, column=0, columnspan=6, sticky="w", padx=16, pady=4)
        tk.Label(self._specific_frame, text="항공사 선택 (복수 가능):",
                 bg=C_PANEL, font=FONT_SMALL).grid(row=0, column=0, columnspan=6, sticky="w", pady=(0,2))
        self._airline_vars = {}
        PER_ROW = 6  # 한 줄에 6개씩
        for i, al in enumerate(SPECIFIC_AIRLINES):
            v = tk.BooleanVar(value=False)
            self._airline_vars[al] = v
            cb = tk.Checkbutton(self._specific_frame, text=al, variable=v,
                                bg=C_PANEL, font=FONT_SMALL, activebackground=C_PANEL,
                                anchor="w", width=13)
            r = 1 + (i // PER_ROW)
            c = i % PER_ROW
            cb.grid(row=r, column=c, sticky="w", padx=2, pady=1)
        self._specific_frame.grid_remove()

        # ── 섹션 4: 시간대 조건 ──
        self._add_section_header("🕐  4. 시간대 조건")
        row4 = self._panel()
        band_cols = [
            ("가는편 출발",  "dep_band"),
            ("가는편 도착",  "arr_band"),
            ("오는편 출발",  "ret_dep_band"),
            ("오는편 도착",  "ret_arr_band"),
        ]
        self._band_vars = {}
        self._band_all_vars = {}
        self._band_custom_from = {}   # 직접입력 시작
        self._band_custom_to   = {}   # 직접입력 종료

        # ── 헤더 ──
        for col, (lbl, key) in enumerate(band_cols):
            tk.Label(row4, text=lbl, bg=C_PANEL, font=FONT_SUB,
                     width=13, anchor="center").grid(row=0, column=col, padx=10, pady=(8,2))

        # ── 직접 입력 행 ──
        tk.Label(row4, text="", bg=C_PANEL, width=2).grid(row=1, column=0, padx=(10,0))
        for col, (lbl, key) in enumerate(band_cols):
            frame_direct = tk.Frame(row4, bg=C_PANEL)
            frame_direct.grid(row=1, column=col, padx=10, pady=(2,4), sticky="w")
            tk.Label(frame_direct, text="직접입력:", bg=C_PANEL, fg="#555", font=FONT_SMALL).pack(side="left")
            from_var = tk.StringVar(value="")
            to_var   = tk.StringVar(value="")
            self._band_custom_from[key] = from_var
            self._band_custom_to[key]   = to_var
            tk.Entry(frame_direct, textvariable=from_var, width=6, font=FONT_SMALL,
                     relief="solid", bd=1, bg="#F8F9FB").pack(side="left", padx=(2,0))
            tk.Label(frame_direct, text="~", bg=C_PANEL, font=FONT_SMALL).pack(side="left", padx=1)
            tk.Entry(frame_direct, textvariable=to_var, width=6, font=FONT_SMALL,
                     relief="solid", bd=1, bg="#F8F9FB").pack(side="left")
            tk.Label(frame_direct, text="(비우면 아래 조건 사용)", bg=C_PANEL,
                     fg="#AAAAAA", font=("맑은 고딕", 7)).pack(side="left", padx=(3,0))

        # ── 구분선 ──
        sep = tk.Frame(row4, bg=C_BORDER, height=1)
        sep.grid(row=2, column=0, columnspan=4, sticky="ew", padx=10, pady=2)

        # ── 전체 체크박스 ──
        for col, (lbl, key) in enumerate(band_cols):
            av = tk.BooleanVar(value=True)
            self._band_all_vars[key] = av
            def _make_all_cb(k, v):
                def _toggle():
                    all_on = v.get()
                    for sv in self._band_vars[k]:
                        sv.set(all_on)
                return _toggle
            cb = tk.Checkbutton(row4, text="전체", variable=av, bg=C_PANEL,
                                font=FONT_SMALL, activebackground=C_PANEL,
                                command=_make_all_cb(key, av))
            cb.grid(row=3, column=col, padx=10, sticky="w")
            slot_vars = []
            for r, slot_lbl in enumerate(TIME_SLOT_LABELS):
                v2 = tk.BooleanVar(value=True)
                slot_vars.append(v2)
            self._band_vars[key] = slot_vars

        # ── 시간대별 체크박스 ──
        for col, (lbl, key) in enumerate(band_cols):
            for r, slot_lbl in enumerate(TIME_SLOT_LABELS):
                v2 = self._band_vars[key][r]
                def _make_slot_cb(k, slot_vs, all_v):
                    def _toggle():
                        all_v.set(all(sv.get() for sv in slot_vs))
                    return _toggle
                cb = tk.Checkbutton(row4, text=slot_lbl, variable=v2, bg=C_PANEL,
                                    font=FONT_SMALL, activebackground=C_PANEL,
                                    command=_make_slot_cb(key, self._band_vars[key], self._band_all_vars[key]))
                cb.grid(row=r+4, column=col, padx=10, sticky="w", pady=1)
        tk.Label(row4, text="", bg=C_PANEL).grid(row=9, column=0)

        # ── 섹션 5: 저장 위치 ──
        self._add_section_header("💾  5. 저장 위치")
        row5 = self._panel()
        self.out_dir_var = tk.StringVar(value=get_default_output_dir())
        tk.Entry(row5, textvariable=self.out_dir_var, width=46, font=FONT_BODY,
                 bg="#F8F9FB", relief="solid", bd=1).grid(row=0, column=0, **pad)
        tk.Button(row5, text="폴더 선택", command=self._choose_dir,
                  bg=C_BTN, fg=C_BTN_FG, font=FONT_SMALL,
                  relief="flat", cursor="hand2", padx=8).grid(row=0, column=1, **pad)

        # ── 기타 옵션 ──
        self._add_section_header("⚙  6. 기타 옵션")
        row6 = self._panel()
        self.show_browser_var = tk.BooleanVar(value=False)
        tk.Checkbutton(row6, text="브라우저 창 표시 (문제 확인용)",
                       variable=self.show_browser_var,
                       bg=C_PANEL, font=FONT_BODY, activebackground=C_PANEL
                       ).grid(row=0, column=0, sticky="w", **pad)

        # ── 실행 버튼 ──
        btn_frame = tk.Frame(self._scroll_frame, bg=C_BG)
        btn_frame.pack(fill="x", padx=16, pady=10)
        self._run_btn = tk.Button(
            btn_frame, text="▶  추출 시작", command=self._start,
            bg=C_BTN_RUN, fg="white", font=("맑은 고딕", 12, "bold"),
            relief="flat", cursor="hand2", padx=24, pady=8
        )
        self._run_btn.pack(side="left")
        self._stop_btn = tk.Button(
            btn_frame, text="■  중지", command=self._stop,
            bg=C_WARN, fg="white", font=("맑은 고딕", 11, "bold"),
            relief="flat", cursor="hand2", padx=16, pady=8, state="disabled"
        )
        self._stop_btn.pack(side="left", padx=10)

        # ── 진행률 ──
        prog_frame = tk.Frame(self._scroll_frame, bg=C_BG)
        prog_frame.pack(fill="x", padx=16, pady=4)
        self._prog_label = tk.Label(prog_frame, text="대기 중...", bg=C_BG,
                                    fg="#555555", font=FONT_SMALL)
        self._prog_label.pack(side="left")
        self._progressbar = ttk.Progressbar(prog_frame, orient="horizontal",
                                            length=400, mode="determinate")
        self._progressbar.pack(side="left", padx=10)
        self._prog_pct = tk.Label(prog_frame, text="", bg=C_BG, fg=C_ACCENT2, font=FONT_SMALL)
        self._prog_pct.pack(side="left")

        # ── 로그 ──
        log_hdr = tk.Frame(self._scroll_frame, bg=C_BG)
        log_hdr.pack(fill="x", padx=16, pady=(8, 2))
        tk.Label(log_hdr, text="📋  실행 로그", bg=C_BG, fg=C_ACCENT, font=FONT_SUB).pack(side="left")
        tk.Button(log_hdr, text="로그 지우기", command=self._clear_log,
                  bg="#EEEEEE", fg="#555555", font=FONT_SMALL,
                  relief="flat", cursor="hand2", padx=6).pack(side="right")

        log_frame = tk.Frame(self._scroll_frame, bg=C_LOG_BG,
                             highlightbackground=C_BORDER, highlightthickness=1)
        log_frame.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        self._log = tk.Text(log_frame, bg=C_LOG_BG, fg=C_LOG_FG, font=FONT_LOG,
                            height=14, wrap="word", state="disabled",
                            relief="flat", padx=8, pady=6)
        log_sb = ttk.Scrollbar(log_frame, command=self._log.yview)
        self._log.configure(yscrollcommand=log_sb.set)
        log_sb.pack(side="right", fill="y")
        self._log.pack(side="left", fill="both", expand=True)
        self._log.tag_configure("ok",   foreground=C_LOG_OK)
        self._log.tag_configure("err",  foreground=C_LOG_ERR)
        self._log.tag_configure("info", foreground=C_LOG_INFO)

        self._log_msg("시스템 준비 완료. 조건을 설정하고 '추출 시작'을 눌러주세요.", "info")
        if not SCRAPER_OK:
            self._log_msg("⚠ scraper_core.py를 찾을 수 없습니다. 같은 폴더에 있는지 확인하세요.", "err")

    # ──────── 섹션 헤더 / 패널 ────────
    def _add_section_header(self, text):
        f = tk.Frame(self._scroll_frame, bg=C_BG)
        f.pack(fill="x", padx=16, pady=(10, 2))
        tk.Label(f, text=text, bg=C_BG, fg=C_ACCENT, font=FONT_SUB).pack(side="left")

    def _panel(self):
        f = tk.Frame(self._scroll_frame, bg=C_PANEL,
                     highlightbackground=C_BORDER, highlightthickness=1)
        f.pack(fill="x", padx=16, pady=2)
        return f

    # ──────── 이벤트 ────────
    def _on_expand_toggle(self):
        if self.use_expand_var.get():
            self._expand_frame.grid()
            self._add_row_btn.grid()
        else:
            self._expand_frame.grid_remove()
            self._add_row_btn.grid_remove()

    def _add_expand_row(self, nights_default="3", arr_to="", ret_dep_from="",
                        ret_dep_to="", arr_band_default="전체"):
        row_idx = len(self._expand_rows) + 1  # 헤더 행 다음부터
        frame = tk.Frame(self._expand_frame, bg=C_PANEL)
        frame.grid(row=row_idx, column=0, columnspan=6, sticky="ew", pady=1)

        # 순위 번호
        tk.Label(frame, text=f"{row_idx}순위", bg=C_PANEL, font=FONT_SMALL,
                 width=5).pack(side="left", padx=(0,4))

        # 일수
        nights_var = tk.StringVar(value=nights_default)
        tk.Entry(frame, textvariable=nights_var, width=3, font=FONT_SMALL,
                 relief="solid", bd=1, bg="#F8F9FB").pack(side="left")
        tk.Label(frame, text="일", bg=C_PANEL, font=FONT_SMALL).pack(side="left", padx=(1,10))

        # 도착 상한
        arr_to_var = tk.StringVar(value=arr_to)
        tk.Label(frame, text="도착~", bg=C_PANEL, font=FONT_SMALL).pack(side="left")
        tk.Entry(frame, textvariable=arr_to_var, width=7, font=FONT_SMALL,
                 relief="solid", bd=1, bg="#F8F9FB").pack(side="left", padx=(2,10))

        # 귀국출발 하한 ~ 상한
        ret_dep_from_var = tk.StringVar(value=ret_dep_from)
        tk.Label(frame, text="귀국출발", bg=C_PANEL, font=FONT_SMALL).pack(side="left")
        tk.Entry(frame, textvariable=ret_dep_from_var, width=6, font=FONT_SMALL,
                 relief="solid", bd=1, bg="#F8F9FB").pack(side="left", padx=(2,1))
        tk.Label(frame, text="이후~", bg=C_PANEL, font=FONT_SMALL).pack(side="left")
        ret_dep_to_var = tk.StringVar(value=ret_dep_to)
        tk.Entry(frame, textvariable=ret_dep_to_var, width=6, font=FONT_SMALL,
                 relief="solid", bd=1, bg="#F8F9FB").pack(side="left", padx=(1,1))
        tk.Label(frame, text="이전", bg=C_PANEL, font=FONT_SMALL).pack(side="left", padx=(0,10))

        # 도착 시간대 (상한 없을 때)
        arr_band_choices = ["전체", "오전", "오후", "야간", "오후+야간"]
        arr_band_var = tk.StringVar(value=arr_band_default)
        tk.Label(frame, text="도착시간대:", bg=C_PANEL, font=FONT_SMALL).pack(side="left")
        ttk.Combobox(frame, textvariable=arr_band_var, values=arr_band_choices,
                     state="readonly", width=8, font=FONT_SMALL).pack(side="left", padx=(2,10))

        # 삭제 버튼
        row_data = {
            "frame": frame,
            "nights": nights_var,
            "arr_to": arr_to_var,
            "ret_dep_from": ret_dep_from_var,
            "ret_dep_to": ret_dep_to_var,
            "arr_band": arr_band_var,
        }
        def _del(rd=row_data):
            rd["frame"].destroy()
            self._expand_rows.remove(rd)
            self._relabel_expand_rows()
        tk.Button(frame, text="✕", command=_del, bg="#FFEEEE", fg=C_WARN,
                  font=FONT_SMALL, relief="flat", cursor="hand2", padx=4).pack(side="left")

        self._expand_rows.append(row_data)

    def _relabel_expand_rows(self):
        for i, rd in enumerate(self._expand_rows):
            for widget in rd["frame"].winfo_children():
                if isinstance(widget, tk.Label) and "순위" in (widget.cget("text") or ""):
                    widget.configure(text=f"{i+1}순위")
                    break

    def _update_return_hint(self):
        import datetime as _dt
        try:
            df = _dt.datetime.strptime(self.date_from_var.get().strip(), "%Y-%m-%d").date()
            off = int(self.return_offset_var.get()) - 1  # 출발일 포함 총 일수 → 오프셋
            d1 = df + _dt.timedelta(days=off)
            if self.extra_return_var.get():
                d2 = df + _dt.timedelta(days=off - 1)
                txt = f"  → {df.strftime('%m/%d')} 출발 시 귀국 현지출발일: {d1.strftime('%m/%d')}, {d2.strftime('%m/%d')} 둘 다 검색"
            else:
                txt = f"  → {df.strftime('%m/%d')} 출발 시 귀국 현지출발일: {d1.strftime('%m/%d')} 검색"
            self._return_hint.configure(text=txt)
        except Exception:
            self._return_hint.configure(text="")

    def _on_country_change(self, which: str):
        if which == "origin":
            country = self.origin_country_var.get()
            cities = list(AIRPORTS[country].keys())
            self._origin_city_cb["values"] = cities
            self.origin_city_var.set(cities[0])
        else:
            country = self.dest_country_var.get()
            cities = list(AIRPORTS[country].keys())
            self._dest_city_cb["values"] = cities
            self.dest_city_var.set(cities[0])

    def _on_mode_change(self, _=None):
        if self.airline_mode_var.get() == "특정 항공사 지정":
            self._specific_frame.grid()
        else:
            self._specific_frame.grid_remove()

    def _choose_dir(self):
        d = filedialog.askdirectory(initialdir=self.out_dir_var.get())
        if d:
            self.out_dir_var.set(d)

    def _clear_log(self):
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

    # ──────── 로그 ────────
    def _log_msg(self, msg: str, tag: str = ""):
        def _do():
            self._log.configure(state="normal")
            ts = datetime.datetime.now().strftime("%H:%M:%S")
            self._log.insert("end", f"[{ts}] {msg}\n", tag or "")
            self._log.see("end")
            self._log.configure(state="disabled")
        self.after(0, _do)

    # ──────── 진행률 업데이트 ────────
    def _update_progress(self, cur: int, total: int):
        def _do():
            pct = int(cur / total * 100)
            self._progressbar["value"] = pct
            self._prog_label.configure(text=f"{cur} / {total} 일")
            self._prog_pct.configure(text=f"{pct}%")
        self.after(0, _do)

    # ──────── 입력 검증 ────────
    def _validate(self):
        import datetime as _dt
        try:
            date_from = _dt.datetime.strptime(self.date_from_var.get().strip(), "%Y-%m-%d").date()
            date_to   = _dt.datetime.strptime(self.date_to_var.get().strip(), "%Y-%m-%d").date()
            return_offset = int(self.return_offset_var.get()) - 1  # 출발일 포함 총 일수 → 오프셋
            assert date_from <= date_to
            assert 1 <= return_offset <= 20
        except Exception:
            messagebox.showerror("입력 오류",
                "출발일 형식(2026-07-10), 시작일≤종료일, 귀국일수(1~20)를 올바르게 입력하세요.")
            return None

        origin_city = self.origin_city_var.get()
        dest_city   = self.dest_city_var.get()
        origin_country = self.origin_country_var.get()
        dest_country   = self.dest_country_var.get()
        if origin_country == dest_country and origin_city == dest_city:
            messagebox.showerror("입력 오류", "출발지와 목적지가 같습니다.")
            return None

        origin = AIRPORTS[origin_country][origin_city]
        dest   = AIRPORTS[dest_country][dest_city]

        mode_disp = self.airline_mode_var.get()
        mode_map  = {
            "LCC 우선 → FSC 대체": "LCC우선_FSC대체",
            "LCC만":               "LCC만",
            "FSC만 (아시아나/대한항공)": "FSC만",
            "외항사만":            "외항사만",
            "특정 항공사 지정":    "특정항공사",
        }
        airline_mode = mode_map[mode_disp]

        specific = []
        if airline_mode == "특정항공사":
            specific = [al for al, v in self._airline_vars.items() if v.get()]
            if not specific:
                messagebox.showerror("입력 오류", "특정 항공사를 하나 이상 선택하세요.")
                return None

        def parse_time_str(s):
            """HH:MM 문자열 → 분(int), 실패시 None"""
            s = s.strip()
            if not s:
                return None
            import re
            m = re.match(r"^(\d{1,2}):?(\d{2})$", s.replace(":", "").zfill(4))
            if m:
                h, mn = int(s.split(":")[0]), int(s.split(":")[1]) if ":" in s else (int(s[:2]), int(s[2:]))
                return h * 60 + mn
            return None

        def get_band_condition(key):
            """직접입력 있으면 (from_min, to_min) 튜플, 없으면 슬롯 리스트"""
            f_str = self._band_custom_from[key].get().strip()
            t_str = self._band_custom_to[key].get().strip()
            if f_str or t_str:
                # 직접입력 우선
                try:
                    fh, fm = (int(x) for x in (f_str or "00:00").split(":"))
                    th, tm = (int(x) for x in (t_str or "24:00").split(":"))
                    to_min = th * 60 + tm
                    if to_min == 0: to_min = 1440  # 00:00 = 자정(24:00)
                    return {"type": "range", "from": fh * 60 + fm, "to": to_min}
                except Exception:
                    messagebox.showwarning("입력 오류", key + " 직접입력 형식이 잘못됐습니다.\n예) 06:00 ~ 09:00")
                    return None
            # 체크박스
            slots = [TIME_SLOT_KEYS[i] for i, v in enumerate(self._band_vars[key]) if v.get()]
            return {"type": "slots", "slots": slots if slots else TIME_SLOT_KEYS[:]}

        dep_cond     = get_band_condition("dep_band")
        arr_cond     = get_band_condition("arr_band")
        ret_dep_cond = get_band_condition("ret_dep_band")
        ret_arr_cond = get_band_condition("ret_arr_band")
        if None in (dep_cond, arr_cond, ret_dep_cond, ret_arr_cond):
            return None

        config = {
            "airline_mode":      airline_mode,
            "specific_airlines": specific,
            "dep_band":     dep_cond,
            "arr_band":     arr_cond,
            "ret_dep_band": ret_dep_cond,
            "ret_arr_band": ret_arr_cond,
        }

        # ── 대체 조건 파싱 ──
        expand_priorities = None
        if self.use_expand_var.get():
            if not self._expand_rows:
                messagebox.showerror("입력 오류", "대체 조건 순위를 1개 이상 추가하세요.")
                return None
            expand_priorities = []
            for i, rd in enumerate(self._expand_rows):
                try:
                    n = int(rd["nights"].get()) - 1  # 출발일 포함 총 일수 → 오프셋
                    assert 1 <= n <= 14
                except Exception:
                    messagebox.showerror("입력 오류", f"{i+1}순위 일수를 올바르게 입력하세요.")
                    return None

                arr_to_str      = rd["arr_to"].get().strip()
                ret_dep_str     = rd["ret_dep_from"].get().strip()
                ret_dep_to_str  = rd["ret_dep_to"].get().strip()
                arr_band_choice = rd["arr_band"].get()

                # 도착 상한 파싱
                if arr_to_str:
                    try:
                        ah, am = (int(x) for x in arr_to_str.split(":"))
                        arr_to_min = ah * 60 + am
                        if arr_to_min == 0: arr_to_min = 1440  # 00:00 = 자정
                        arr_cond_exp = {"type": "range", "from": 0, "to": arr_to_min}
                    except Exception:
                        messagebox.showerror("입력 오류", f"{i+1}순위 도착 상한 형식 오류 (예: 10:10)")
                        return None
                else:
                    band_map2 = {
                        "전체": ["새벽","오전","오후","야간"],
                        "오전": ["오전"], "오후": ["오후"], "야간": ["야간"],
                        "오후+야간": ["오후","야간"],
                    }
                    arr_cond_exp = {"type": "slots", "slots": band_map2.get(arr_band_choice, ["새벽","오전","오후","야간"])}

                # 귀국출발 하한~상한 파싱
                if ret_dep_str or ret_dep_to_str:
                    try:
                        if ret_dep_str:
                            rh, rm = (int(x) for x in ret_dep_str.split(":"))
                            r_from = rh * 60 + rm
                        else:
                            r_from = 0
                        if ret_dep_to_str:
                            th, tm = (int(x) for x in ret_dep_to_str.split(":"))
                            r_to = th * 60 + tm
                            if r_to == 0: r_to = 1440  # 00:00 = 자정
                        else:
                            r_to = 24 * 60
                        ret_dep_cond_exp = {"type": "range", "from": r_from, "to": r_to}
                    except Exception:
                        messagebox.showerror("입력 오류", f"{i+1}순위 귀국출발 형식 오류 (예: 19:00)")
                        return None
                else:
                    ret_dep_cond_exp = {"type": "slots", "slots": ["새벽","오전","오후","야간"]}

                expand_priorities.append({
                    "nights": n,
                    "arr_cond": arr_cond_exp,
                    "ret_dep_cond": ret_dep_cond_exp,
                })

        return {
            "origin": origin, "dest": dest,
            "date_from": date_from, "date_to": date_to,
            "return_offset": return_offset,
            "extra_return": self.extra_return_var.get(),
            "config": config,
            "expand_priorities": expand_priorities,
            "show_browser": self.show_browser_var.get(),
            "out_dir": self.out_dir_var.get(),
        }

    # ──────── 실행 ────────
    def _start(self):
        if not SCRAPER_OK:
            messagebox.showerror("오류", "scraper_core.py가 없어 실행할 수 없습니다.")
            return
        params = self._validate()
        if params is None:
            return
        self._running = True
        self._run_btn.configure(state="disabled")
        self._stop_btn.configure(state="normal")
        self._progressbar["value"] = 0
        self._prog_label.configure(text="실행 중...")
        self._thread = threading.Thread(target=self._run_scrape, args=(params,), daemon=True)
        self._thread.start()

    def _stop(self):
        self._running = False
        self._log_msg("⏹ 중지 요청됨. 현재 날짜 완료 후 종료됩니다.", "err")

    def _run_scrape(self, p):
        origin   = p["origin"]
        dest     = p["dest"]
        date_from = p["date_from"]
        date_to   = p["date_to"]
        config   = p["config"]
        show     = p["show_browser"]
        out_dir  = p["out_dir"]
        expand_priorities = p.get("expand_priorities")
        extra_return = p.get("extra_return", False)

        from datetime import timedelta
        from scraper_core import (init_driver, build_url, fetch_flights,
                                  select_best, parse_price, calc_per_person, in_band, parse_hour)

        # 귀국 오프셋 = 출발 +N일 (사용자 직접 입력)
        base_offset = p["return_offset"]
        # 추가 귀국일 검색 시 [기본, 기본+1] 둘 다 검색
        offsets = [base_offset, base_offset - 1] if extra_return else [base_offset]

        # 출발일 리스트 생성
        date_list = []
        d = date_from
        while d <= date_to:
            date_list.append(d)
            d += timedelta(days=1)
        total_days = len(date_list)

        fname = f"gmarket_{origin}_{dest}_{date_from.strftime('%Y%m%d')}_{date_to.strftime('%Y%m%d')}.xlsx"
        out_path = os.path.join(out_dir, fname)

        self._log_msg(
            f"🛫 추출 시작: {origin} → {dest}  |  {date_from} ~ {date_to}  |  귀국 출발+{base_offset}일  |  총 {total_days}일",
            "info"
        )
        if expand_priorities:
            self._log_msg(f"  🔄 대체 조건 모드: {len(expand_priorities)}개 순위 적용", "info")

        try:
            driver = init_driver(show=show)
            rows = []

            for idx, dep_date in enumerate(date_list, 1):
                if not self._running:
                    self._log_msg("⏹ 중지됨.", "err")
                    break

                self._log_msg(
                    f"  {dep_date.strftime('%Y-%m-%d')} ({dep_date.strftime('%a')}) 검색 중..."
                )

                best = None
                used_offset = base_offset
                used_priority = None

                if expand_priorities:
                    for pri_idx, pri in enumerate(expand_priorities):
                        # 순위별 일수를 귀국 오프셋으로 사용 (n일 → 출발+n일 귀국)
                        n = pri["nights"]
                        off = n
                        arr_date_try = dep_date + timedelta(days=off)
                        url = build_url(origin, dest,
                                        dep_date.strftime("%Y%m%d"),
                                        arr_date_try.strftime("%Y%m%d"))
                        flights = fetch_flights(driver, url, self._log_msg)
                        if not flights:
                            continue
                        trial_config = dict(config)
                        trial_config["arr_band"]     = pri["arr_cond"]
                        trial_config["ret_dep_band"] = pri["ret_dep_cond"]
                        trial_config["ret_arr_band"] = {"type":"slots","slots":["새벽","오전","오후","야간"]}
                        candidate = select_best(flights, trial_config)
                        if candidate:
                            best = candidate
                            used_offset = off
                            used_priority = pri_idx + 1
                            break
                    arr_date = dep_date + timedelta(days=used_offset)
                else:
                    # 여러 귀국 오프셋으로 각각 검색 후 최저가 선택
                    best = None
                    best_arr_date = dep_date + timedelta(days=base_offset)
                    best_total = None
                    for off in offsets:
                        arr_try = dep_date + timedelta(days=off)
                        url = build_url(origin, dest,
                                        dep_date.strftime("%Y%m%d"),
                                        arr_try.strftime("%Y%m%d"))
                        flights = fetch_flights(driver, url, self._log_msg)
                        # -1일 오프셋 검색 시 오는편 출발일이 arr_try이고 한국 도착이 10:30 이전인 편만 유효
                        if off < base_offset and flights:
                            arr_try_str = arr_try.strftime("%Y-%m-%d")
                            def _rArr_min(f):
                                t = f.get("rArr", "")
                                if ":" not in t:
                                    return 9999
                                h, m = t.split(":")
                                return int(h) * 60 + int(m)
                            flights = [f for f in flights
                                       if f.get("rDepDate", "") == arr_try_str
                                       and _rArr_min(f) <= 630]
                        cand = select_best(flights, config) if flights else None
                        if cand:
                            cand_total = parse_price(cand.get("cardPrice", "0"))
                            if best is None or (cand_total and cand_total < best_total):
                                best = cand
                                best_total = cand_total
                                best_arr_date = arr_try
                    arr_date = best_arr_date
                    # -1일 오프셋 선택 시 한국 도착일은 +1일 (현지 출발 익일 도착)
                    if best_arr_date < dep_date + timedelta(days=base_offset):
                        arr_date = best_arr_date + timedelta(days=1)

                if best:
                    actual_price = best.get("cardPrice", "")
                    best["price"] = actual_price
                    total4 = parse_price(actual_price)
                    per1   = calc_per_person(total4)
                    pri_note = f" [{used_priority}순위]" if used_priority else ""
                    rows.append({
                        "dep_date": dep_date.strftime("%Y-%m-%d"),
                        "arr_date": arr_date.strftime("%Y-%m-%d"),
                        "airline":  best["airline"],
                        "dep":  best["dep"],  "arr":  best["arr"],
                        "rDep": best.get("rDep",""), "rArr": best.get("rArr",""),
                        "total4": total4, "per1": per1,
                        "seller": pri_note.strip(),
                        "found": True,
                    })
                    self._log_msg(
                        f"    ✔{pri_note} {best['airline']}  {best['dep']}→{best['arr']}  4인:{total4:,}원  1인:{per1:,}원",
                        "ok"
                    )
                else:
                    rows.append({
                        "dep_date": dep_date.strftime("%Y-%m-%d"),
                        "arr_date": (dep_date + timedelta(days=base_offset)).strftime("%Y-%m-%d"),
                        "airline":"", "dep":"","arr":"","rDep":"","rArr":"",
                        "total4":0,"per1":0,"seller":"","found":False,
                    })
                    self._log_msg("    ✗ 조건에 맞는 항공편 없음")

                self._update_progress(idx, total_days)

            driver.quit()

            if rows:
                save_excel(rows, origin, dest, date_from.year, date_from.month, out_path)
                found_cnt = sum(1 for r in rows if r["found"])
                self._log_msg(
                    f"\n✅ 완료! 총 {len(rows)}일 중 {found_cnt}일 확인 → {fname}",
                    "ok"
                )
                self.after(0, lambda: messagebox.showinfo(
                    "완료",
                    f"추출 완료!\n{date_from} ~ {date_to}\n{found_cnt}일 확인\n저장: {out_path}"
                ))
                try:
                    from plyer import notification
                    notification.notify(
                        title="✈ G마켓 항공료 추출 완료",
                        message=f"{origin} → {dest}  {found_cnt}일 확인 완료",
                        app_name="G마켓 항공료 추출기",
                        timeout=8,
                    )
                except Exception:
                    pass
            else:
                self._log_msg("⚠ 수집된 데이터 없음", "err")

        except Exception as e:
            self._log_msg(f"❌ 오류 발생: {e}", "err")
            self.after(0, lambda: messagebox.showerror("오류", str(e)))

        finally:
            self._running = False
            self.after(0, lambda: self._run_btn.configure(state="normal"))
            self.after(0, lambda: self._stop_btn.configure(state="disabled"))
            self.after(0, lambda: self._prog_label.configure(text="완료"))


# ─────────────────────────────────────────────
#  진입점
# ─────────────────────────────────────────────
if __name__ == "__main__":
    app = GmarketAirApp()
    app.mainloop()
