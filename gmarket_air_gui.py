"""
gmarket_air_gui.py
G마켓 항공료 자동 추출 - GUI 실행기
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import os
import re
import datetime
import sys
import shutil
import urllib.request
import ctypes

# ─────────────────────────────────────────────
#  자동 업데이트 (GitHub)
# ─────────────────────────────────────────────
CURRENT_VERSION = "2.0"
GITHUB_USER     = "rlawlsah22"
GITHUB_REPO     = "gmarket-air-tool"
RAW_BASE        = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/main"
UPDATE_FILES    = ["gmarket_air_gui.py", "scraper_core.py", "airpremia_fare_checker.py"]

def check_and_update():
    try:
        url = f"{RAW_BASE}/version.txt"
        with urllib.request.urlopen(url, timeout=5) as r:
            latest = r.read().decode().strip()

        if latest <= CURRENT_VERSION:
            return

        # 현재 exe/py 위치
        if getattr(sys, "frozen", False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))

        for fname in UPDATE_FILES:
            file_url = f"{RAW_BASE}/{fname}"
            dst = os.path.join(base_dir, fname)
            with urllib.request.urlopen(file_url, timeout=10) as r:
                with open(dst, "wb") as f:
                    f.write(r.read())

        ctypes.windll.user32.MessageBoxW(
            0,
            f"새 버전({latest})으로 업데이트되었습니다.\n프로그램을 닫고 다시 실행해주세요.",
            "업데이트 완료",
            0x40
        )
        sys.exit()

    except Exception:
        pass

check_and_update()

# ─────────────────────────────────────────────
#  프리셋 저장 (로컬 JSON, 노선/항공사모드/시간대)
# ─────────────────────────────────────────────
import json

def _preset_file_path():
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, "presets.json")

def load_presets():
    path = _preset_file_path()
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_presets(data):
    path = _preset_file_path()
    _unhide_file_windows(path)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    _hide_file_windows(path)

def _hide_file_windows(path):
    try:
        if os.name == "nt":
            FILE_ATTRIBUTE_HIDDEN = 0x02
            ctypes.windll.kernel32.SetFileAttributesW(str(path), FILE_ATTRIBUTE_HIDDEN)
    except Exception:
        pass

def _unhide_file_windows(path):
    try:
        if os.name == "nt" and os.path.exists(path):
            FILE_ATTRIBUTE_NORMAL = 0x80
            ctypes.windll.kernel32.SetFileAttributesW(str(path), FILE_ATTRIBUTE_NORMAL)
    except Exception:
        pass

# scraper_core가 같은 폴더에 있어야 함
try:
    from scraper_core import (
        AIRPORTS, AIRPORT_CODES, LCC_ALL, FSC_ALL, LCC_TIER1, LCC_TIER2,
        FOREIGN_ALL, collect_monthly, save_excel, save_excel_multi
    )
    SCRAPER_OK = True
except ImportError:
    SCRAPER_OK = False
    AIRPORTS = {
        "국내": {"인천":"ICN","부산":"PUS","김포":"GMP","청주":"CJJ","대구":"TAE"},
        "일본": {"삿포로":"CTS","도쿄":"TYO","오사카":"OSA","후쿠오카":"FUK","오키나와":"OKA"},
    }
    LCC_ALL = FSC_ALL = LCC_TIER1 = LCC_TIER2 = FOREIGN_ALL = []

try:
    from airpremia_fare_checker import (
        init_driver as airpremia_init_driver,
        check_one_date as airpremia_check_one_date,
        write_styled_sheet as airpremia_write_styled_sheet,
    )
    AIRPREMIA_OK = True
except ImportError:
    AIRPREMIA_OK = False

AIRPREMIA_DESTINATIONS = {
    "도쿄/나리타 (NRT)": "NRT",
    "방콕 (BKK)": "BKK",
}


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
C_FSC_BG   = "#EBF3FB"
C_LCC_BG   = "#EDFBEE"

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
    "중국국제항공", "중국남방항공", "중국동방항공", "산동항공",
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
#  검색 조건 블록 (여행일수 + 항공사모드 + 시간대 + 대체조건)
#  단일 모드/세트 모드 양쪽에서 재사용
# ─────────────────────────────────────────────
class ConditionBlock:
    """
    하나의 검색 조건 세트를 표현하는 위젯.
    - 여행 일수 + (-1일 검색)
    - 항공사 모드 + 특정항공사 선택
    - 시간대 조건 (4칸)
    - 대체 조건 자동 확장 (선택)

    parent: 부모 tk 프레임
    bg: 배경색
    show_set_header: True면 "세트 N" 헤더 + 삭제버튼 표시 (세트 모드용)
    set_index: 세트 번호 (헤더 표시용)
    on_delete: 삭제 버튼 콜백 (세트 모드용)
    """

    def __init__(self, parent, bg=None, show_set_header=False,
                 set_index=1, on_delete=None, default_nights="3",
                 default_mode=None):
        self.bg = bg if bg else C_PANEL
        self.frame = tk_module = None  # placeholder, set below

        import tkinter as tk
        from tkinter import ttk
        self._tk = tk
        self._ttk = ttk

        self.outer = tk.Frame(parent, bg=self.bg,
                              highlightbackground=C_BORDER, highlightthickness=1)
        self.outer.pack(fill="x", padx=16, pady=8)

        pad = {"padx": 12, "pady": 4}
        row_idx = 0

        # ── 세트 헤더 (세트 모드일 때만) ──
        if show_set_header:
            hdr = tk.Frame(self.outer, bg=self.bg)
            hdr.grid(row=row_idx, column=0, columnspan=8, sticky="ew", padx=18, pady=(14, 4))
            badge = tk.Label(hdr, text=str(set_index), bg=C_PANEL, fg=C_ACCENT,
                             font=FONT_SUB, width=2,
                             highlightbackground=C_BORDER, highlightthickness=1)
            badge.pack(side="left", padx=(0, 6))
            tk.Label(hdr, text=f"세트 {set_index}", bg=self.bg, fg=C_ACCENT,
                     font=FONT_SUB).pack(side="left")
            if on_delete:
                tk.Button(hdr, text="✕ 세트 삭제", command=on_delete,
                          bg="#FFEEEE", fg=C_WARN, font=FONT_SMALL,
                          relief="flat", cursor="hand2", padx=6).pack(side="right")
            row_idx += 1

        # ── 여행 일수 + 항공사 모드 (한 줄) ──
        top_row = tk.Frame(self.outer, bg=self.bg)
        top_row.grid(row=row_idx, column=0, columnspan=8, sticky="w", padx=18, pady=(14, 8))
        row_idx += 1

        tk.Label(top_row, text="여행 일수", bg=self.bg, font=FONT_SUB).pack(side="left", pady=4)
        self.return_offset_var = tk.StringVar(value=default_nights)
        ttk.Spinbox(top_row, from_=1, to=20, textvariable=self.return_offset_var,
                    width=4, font=FONT_BODY).pack(side="left", padx=(6, 4))
        tk.Label(top_row, text="일", bg=self.bg, font=FONT_BODY).pack(side="left", padx=(0, 24))

        tk.Label(top_row, text="항공사 모드", bg=self.bg, font=FONT_SUB).pack(side="left", pady=4)
        self.airline_mode_var = tk.StringVar(value=default_mode or AIRLINE_MODES[0])
        mode_cb = ttk.Combobox(top_row, textvariable=self.airline_mode_var,
                               values=AIRLINE_MODES, state="readonly", width=30)
        mode_cb.pack(side="left", padx=(6, 0), pady=4)
        mode_cb.bind("<<ComboboxSelected>>", self._on_mode_change)

        # ── 특정 항공사 선택 ──
        self._specific_frame = tk.Frame(self.outer, bg=self.bg)
        self._specific_frame.grid(row=row_idx, column=0, columnspan=8, sticky="w", padx=18, pady=6)
        row_idx += 1
        tk.Label(self._specific_frame, text="항공사 선택 (복수 가능):",
                 bg=self.bg, font=FONT_SMALL).grid(row=0, column=0, columnspan=6, sticky="w", pady=(0, 2))
        self._airline_vars = {}
        PER_ROW = 6
        for i, al in enumerate(SPECIFIC_AIRLINES):
            v = tk.BooleanVar(value=False)
            self._airline_vars[al] = v
            cb = tk.Checkbutton(self._specific_frame, text=al, variable=v,
                                bg=self.bg, font=FONT_SMALL, activebackground=self.bg,
                                anchor="w", width=13)
            cb.grid(row=1 + (i // PER_ROW), column=i % PER_ROW, sticky="w", padx=2, pady=1)
        self._specific_frame.grid_remove()

        # ── -1일도 함께 검색 ──
        self.extra_return_var = tk.BooleanVar(value=False)
        tk.Checkbutton(self.outer, text="-1일도 함께 검색 (귀국 현지출발이 하루 빠른 것도 함께 비교)",
                       variable=self.extra_return_var, bg=self.bg, font=FONT_SMALL,
                       activebackground=self.bg, command=self._update_return_hint
                       ).grid(row=row_idx, column=0, columnspan=8, sticky="w", padx=18, pady=(4, 0))
        row_idx += 1

        self._return_hint = tk.Label(self.outer, text="", bg=self.bg, fg=C_ACCENT2,
                                     font=("맑은 고딕", 8))
        self._return_hint.grid(row=row_idx, column=0, columnspan=8, sticky="w", padx=20, pady=(2, 8))
        row_idx += 1

        self.return_offset_var.trace_add("write", lambda *a: self._update_return_hint())
        self._date_from_getter = lambda: None  # 외부에서 주입 (date_from 참조용)

        # ── 구분선 ──
        sep_line = tk.Frame(self.outer, bg=C_BORDER, height=1)
        sep_line.grid(row=row_idx, column=0, columnspan=8, sticky="ew", padx=18, pady=(4, 0))
        row_idx += 1

        # ── 시간대 조건 ──
        band_outer = tk.Frame(self.outer, bg=self.bg)
        band_outer.grid(row=row_idx, column=0, columnspan=8, sticky="ew", padx=18, pady=(14, 8))
        row_idx += 1

        tk.Label(band_outer, text="시간대 조건", bg=self.bg, fg=C_ACCENT,
                 font=FONT_SUB).pack(anchor="w", pady=(0, 8))

        band_grid = tk.Frame(band_outer, bg=self.bg)
        band_grid.pack(fill="x")
        for _col in range(4):
            band_grid.columnconfigure(_col, weight=1)

        band_cols = [
            ("가는편 출발", "dep_band"),
            ("가는편 도착", "arr_band"),
            ("오는편 출발", "ret_dep_band"),
            ("오는편 도착", "ret_arr_band"),
        ]
        self._band_vars = {}
        self._band_all_vars = {}
        self._band_custom_from = {}
        self._band_custom_to   = {}

        for col, (lbl, key) in enumerate(band_cols):
            tk.Label(band_grid, text=lbl, bg=self.bg, font=FONT_SMALL,
                     anchor="center").grid(row=0, column=col, padx=10, pady=(2, 4), sticky="ew")

        for col, (lbl, key) in enumerate(band_cols):
            fd = tk.Frame(band_grid, bg=self.bg)
            fd.grid(row=1, column=col, padx=10, pady=(0, 4), sticky="ew")
            tk.Label(fd, text="직접입력:", bg=self.bg, fg="#555", font=("맑은 고딕", 8)).pack(side="left")
            fv = tk.StringVar(value="")
            tv = tk.StringVar(value="")
            self._band_custom_from[key] = fv
            self._band_custom_to[key]   = tv
            tk.Entry(fd, textvariable=fv, width=5, font=("맑은 고딕", 8),
                     relief="solid", bd=1, bg="#F8F9FB").pack(side="left", padx=(2, 0))
            tk.Label(fd, text="~", bg=self.bg, font=("맑은 고딕", 8)).pack(side="left", padx=1)
            tk.Entry(fd, textvariable=tv, width=5, font=("맑은 고딕", 8),
                     relief="solid", bd=1, bg="#F8F9FB").pack(side="left")

        for col in range(4):
            av = tk.BooleanVar(value=True)
            key = band_cols[col][1]
            self._band_all_vars[key] = av
            def _make_all_cb(k, v):
                def _toggle():
                    for sv in self._band_vars[k]:
                        sv.set(v.get())
                return _toggle
            tk.Checkbutton(band_grid, text="전체", variable=av, bg=self.bg,
                           font=FONT_SMALL, activebackground=self.bg,
                           command=_make_all_cb(key, av)).grid(row=2, column=col, padx=10, sticky="ew", pady=(2,2))
            slot_vars = []
            for r, slot_lbl in enumerate(TIME_SLOT_LABELS):
                v2 = tk.BooleanVar(value=True)
                slot_vars.append(v2)
            self._band_vars[key] = slot_vars

        for col, (lbl, key) in enumerate(band_cols):
            for r, slot_lbl in enumerate(TIME_SLOT_LABELS):
                v2 = self._band_vars[key][r]
                def _make_slot_cb(k, slot_vs, all_v):
                    def _toggle():
                        all_v.set(all(sv.get() for sv in slot_vs))
                    return _toggle
                tk.Checkbutton(band_grid, text=slot_lbl, variable=v2, bg=self.bg,
                              font=FONT_SMALL, activebackground=self.bg,
                              command=_make_slot_cb(key, self._band_vars[key], self._band_all_vars[key])
                              ).grid(row=r + 3, column=col, padx=10, sticky="ew", pady=2)

        # ── 구분선 ──
        sep_line2 = tk.Frame(self.outer, bg=C_BORDER, height=1)
        sep_line2.grid(row=row_idx, column=0, columnspan=8, sticky="ew", padx=18, pady=(6, 0))
        row_idx += 1

        # ── 대체 조건 자동 확장 ──
        expand_outer = tk.Frame(self.outer, bg=self.bg)
        expand_outer.grid(row=row_idx, column=0, columnspan=8, sticky="ew", padx=18, pady=(14, 14))
        row_idx += 1

        self.use_expand_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            expand_outer, text="대체 조건 자동 확장 사용 (체크 시 아래 순위 조건으로 검색, 위 일수 설정 무시)",
            variable=self.use_expand_var, bg=self.bg, font=FONT_SMALL,
            activebackground=self.bg, command=self._on_expand_toggle
        ).pack(anchor="w")

        self._expand_frame = tk.Frame(expand_outer, bg=self.bg)
        self._expand_frame.pack(fill="x", pady=(4, 0))

        hdr_row = tk.Frame(self._expand_frame, bg=self.bg)
        hdr_row.pack(fill="x")
        for hdr in ["순위", "귀국일수(일)", "도착~(상한)", "귀국출발(이후~이전)", "도착시간대", ""]:
            tk.Label(hdr_row, text=hdr, bg=self.bg, font=("맑은 고딕", 8, "bold"),
                     fg=C_ACCENT, width=10, anchor="w").pack(side="left", padx=2)

        self._expand_rows_frame = tk.Frame(self._expand_frame, bg=self.bg)
        self._expand_rows_frame.pack(fill="x")
        self._expand_rows = []

        self._add_row_btn = tk.Button(
            expand_outer, text="＋ 순위 추가", command=self._add_expand_row,
            bg="#E8F0FE", fg=C_ACCENT, font=FONT_SMALL, relief="flat", cursor="hand2", padx=8
        )
        self._add_row_btn.pack(anchor="w", pady=(4, 0))

        self._expand_frame.pack_forget()
        self._add_row_btn.pack_forget()

        self.after_id_holder = None  # set by caller for hint updates

    # ──────── 프리셋 export / import (항공사 모드 + 특정항공사 + 시간대) ────────
    def export_preset_fields(self):
        return {
            "airline_mode": self.airline_mode_var.get(),
            "specific_airlines": [al for al, v in self._airline_vars.items() if v.get()],
            "bands": {
                key: [sv.get() for sv in slot_vars]
                for key, slot_vars in self._band_vars.items()
            },
            "band_custom": {
                key: [self._band_custom_from[key].get(), self._band_custom_to[key].get()]
                for key in self._band_custom_from
            },
        }

    def import_preset_fields(self, data):
        mode = data.get("airline_mode")
        if mode in AIRLINE_MODES:
            self.airline_mode_var.set(mode)
            self._on_mode_change()

        specific = set(data.get("specific_airlines", []))
        for al, v in self._airline_vars.items():
            v.set(al in specific)

        bands = data.get("bands", {})
        for key, slot_vals in bands.items():
            if key not in self._band_vars:
                continue
            for sv, val in zip(self._band_vars[key], slot_vals):
                sv.set(val)
            if key in self._band_all_vars:
                self._band_all_vars[key].set(all(slot_vals))

        band_custom = data.get("band_custom", {})
        for key, (from_val, to_val) in band_custom.items():
            if key in self._band_custom_from:
                self._band_custom_from[key].set(from_val)
                self._band_custom_to[key].set(to_val)

    # ──────── 항공사 모드 토글 ────────
    def _on_mode_change(self, _=None):
        if self.airline_mode_var.get() == "특정 항공사 지정":
            self._specific_frame.grid()
        else:
            self._specific_frame.grid_remove()

    # ──────── 대체조건 토글 ────────
    def _on_expand_toggle(self):
        if self.use_expand_var.get():
            self._expand_frame.pack(fill="x", pady=(4, 0))
            self._add_row_btn.pack(anchor="w", pady=(4, 0))
            if not self._expand_rows:
                self._add_expand_row(nights_default="2", arr_to="10:10", ret_dep_from="19:00")
                self._add_expand_row(nights_default="3", arr_to="10:10", ret_dep_from="19:00")
        else:
            self._expand_frame.pack_forget()
            self._add_row_btn.pack_forget()

    def _add_expand_row(self, nights_default="3", arr_to="", ret_dep_from="",
                        ret_dep_to="", arr_band_default="전체"):
        tk = self._tk
        ttk = self._ttk
        row_idx = len(self._expand_rows) + 1
        frame = tk.Frame(self._expand_rows_frame, bg=self.bg)
        frame.pack(fill="x", pady=1)

        tk.Label(frame, text=f"{row_idx}순위", bg=self.bg, font=FONT_SMALL,
                 width=5).pack(side="left", padx=(0, 2))
        nights_var = tk.StringVar(value=nights_default)
        tk.Entry(frame, textvariable=nights_var, width=3, font=FONT_SMALL,
                 relief="solid", bd=1, bg="#F8F9FB").pack(side="left")
        tk.Label(frame, text="일", bg=self.bg, font=FONT_SMALL).pack(side="left", padx=(1, 8))

        arr_to_var = tk.StringVar(value=arr_to)
        tk.Label(frame, text="도착~", bg=self.bg, font=FONT_SMALL).pack(side="left")
        tk.Entry(frame, textvariable=arr_to_var, width=6, font=FONT_SMALL,
                 relief="solid", bd=1, bg="#F8F9FB").pack(side="left", padx=(2, 8))

        ret_dep_from_var = tk.StringVar(value=ret_dep_from)
        tk.Label(frame, text="귀국출발", bg=self.bg, font=FONT_SMALL).pack(side="left")
        tk.Entry(frame, textvariable=ret_dep_from_var, width=5, font=FONT_SMALL,
                 relief="solid", bd=1, bg="#F8F9FB").pack(side="left", padx=(2, 1))
        tk.Label(frame, text="~", bg=self.bg, font=FONT_SMALL).pack(side="left")
        ret_dep_to_var = tk.StringVar(value=ret_dep_to)
        tk.Entry(frame, textvariable=ret_dep_to_var, width=5, font=FONT_SMALL,
                 relief="solid", bd=1, bg="#F8F9FB").pack(side="left", padx=(1, 8))

        arr_band_var = tk.StringVar(value=arr_band_default)
        ttk.Combobox(frame, textvariable=arr_band_var,
                     values=["전체", "오전", "오후", "야간", "오후+야간"],
                     state="readonly", width=8, font=FONT_SMALL).pack(side="left", padx=(2, 8))

        row_data = {
            "frame": frame, "nights": nights_var, "arr_to": arr_to_var,
            "ret_dep_from": ret_dep_from_var, "ret_dep_to": ret_dep_to_var,
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

    # ──────── 힌트 갱신 (외부에서 date_from 전달받아 호출) ────────
    def update_return_hint(self, date_from_str):
        import datetime as _dt
        try:
            df = _dt.datetime.strptime(date_from_str.strip(), "%Y-%m-%d").date()
            off = int(self.return_offset_var.get()) - 1
            d1 = df + _dt.timedelta(days=off)
            if self.extra_return_var.get() and off > 0:
                d2 = df + _dt.timedelta(days=off - 1)
                txt = f"  → {df.strftime('%m/%d')} 출발 시 귀국 현지출발일: {d1.strftime('%m/%d')}, {d2.strftime('%m/%d')} 둘 다 검색"
            else:
                txt = f"  → {df.strftime('%m/%d')} 출발 시 귀국 현지출발일: {d1.strftime('%m/%d')} 검색"
            self._return_hint.configure(text=txt)
        except Exception:
            self._return_hint.configure(text="")

    def _update_return_hint(self):
        # 외부에서 등록한 date_from getter 사용
        getter = getattr(self, "_date_from_getter", None)
        if getter:
            val = getter()
            if val:
                self.update_return_hint(val)

    def bind_date_from(self, getter_fn):
        """외부(메인 앱)의 date_from StringVar 값을 가져오는 함수를 등록"""
        self._date_from_getter = getter_fn

    # ──────── 검증 + config 반환 ────────
    def validate_and_get(self, messagebox, label_prefix=""):
        """
        반환: (return_offset:int, extra_return:bool, config:dict, expand_priorities:list|None) 또는 None(오류)
        """
        try:
            return_offset = int(self.return_offset_var.get()) - 1
            assert 1 <= return_offset <= 20
        except Exception:
            messagebox.showerror("입력 오류", f"{label_prefix}여행 일수를 올바르게 입력하세요 (1~20일).")
            return None

        mode_disp = self.airline_mode_var.get()
        mode_map = {
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
                messagebox.showerror("입력 오류", f"{label_prefix}특정 항공사를 하나 이상 선택하세요.")
                return None

        def get_band_condition(key):
            f_str = self._band_custom_from[key].get().strip()
            t_str = self._band_custom_to[key].get().strip()
            if f_str or t_str:
                try:
                    fh, fm = (int(x) for x in (f_str or "00:00").split(":"))
                    th, tm = (int(x) for x in (t_str or "24:00").split(":"))
                    to_min = th * 60 + tm
                    if to_min == 0:
                        to_min = 1440
                    return {"type": "range", "from": fh * 60 + fm, "to": to_min}
                except Exception:
                    messagebox.showwarning("입력 오류", f"{label_prefix}{key} 직접입력 형식이 잘못됐습니다.\n예) 06:00 ~ 09:00")
                    return None
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

        expand_priorities = None
        if self.use_expand_var.get():
            if not self._expand_rows:
                messagebox.showerror("입력 오류", f"{label_prefix}대체 조건 순위를 1개 이상 추가하세요.")
                return None
            expand_priorities = []
            for i, rd in enumerate(self._expand_rows):
                try:
                    n = int(rd["nights"].get()) - 1
                    assert 1 <= n <= 14
                except Exception:
                    messagebox.showerror("입력 오류", f"{label_prefix}{i+1}순위 일수를 올바르게 입력하세요.")
                    return None

                arr_to_str      = rd["arr_to"].get().strip()
                ret_dep_str     = rd["ret_dep_from"].get().strip()
                ret_dep_to_str  = rd["ret_dep_to"].get().strip()
                arr_band_choice = rd["arr_band"].get()

                if arr_to_str:
                    try:
                        ah, am = (int(x) for x in arr_to_str.split(":"))
                        arr_to_min = ah * 60 + am
                        if arr_to_min == 0:
                            arr_to_min = 1440
                        arr_cond_exp = {"type": "range", "from": 0, "to": arr_to_min}
                    except Exception:
                        messagebox.showerror("입력 오류", f"{label_prefix}{i+1}순위 도착 상한 형식 오류 (예: 10:10)")
                        return None
                else:
                    band_map2 = {
                        "전체": ["새벽", "오전", "오후", "야간"],
                        "오전": ["오전"], "오후": ["오후"], "야간": ["야간"],
                        "오후+야간": ["오후", "야간"],
                    }
                    arr_cond_exp = {"type": "slots", "slots": band_map2.get(arr_band_choice, ["새벽", "오전", "오후", "야간"])}

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
                            if r_to == 0:
                                r_to = 1440
                        else:
                            r_to = 24 * 60
                        ret_dep_cond_exp = {"type": "range", "from": r_from, "to": r_to}
                    except Exception:
                        messagebox.showerror("입력 오류", f"{label_prefix}{i+1}순위 귀국출발 형식 오류 (예: 19:00)")
                        return None
                else:
                    ret_dep_cond_exp = {"type": "slots", "slots": ["새벽", "오전", "오후", "야간"]}

                expand_priorities.append({
                    "nights": n,
                    "arr_cond": arr_cond_exp,
                    "ret_dep_cond": ret_dep_cond_exp,
                })

        extra_return = self.extra_return_var.get()
        return (return_offset, extra_return, config, expand_priorities)


# ─────────────────────────────────────────────
#  검색 조건 블록 (에어프레미아 전용: 목적지 + 박수 + 인원수)
# ─────────────────────────────────────────────
class AirpremiaConditionBlock:
    def __init__(self, parent, bg=None, show_set_header=False, set_index=1,
                 on_delete=None, default_nights="3"):
        self.bg = bg if bg else C_PANEL

        self.outer = tk.Frame(parent, bg=self.bg,
                              highlightbackground=C_BORDER, highlightthickness=1)
        self.outer.pack(fill="x", padx=16, pady=8)

        if show_set_header:
            hdr = tk.Frame(self.outer, bg=self.bg)
            hdr.pack(fill="x", padx=16, pady=(12, 2))
            badge = tk.Label(hdr, text=str(set_index), bg=C_PANEL, fg=C_ACCENT,
                             font=FONT_SUB, width=2,
                             highlightbackground=C_BORDER, highlightthickness=1)
            badge.pack(side="left", padx=(0, 6))
            tk.Label(hdr, text=f"세트 {set_index}", bg=self.bg, fg=C_ACCENT,
                     font=FONT_SUB).pack(side="left")
            if on_delete:
                tk.Button(hdr, text="✕ 세트 삭제", command=lambda: self._on_delete(),
                          bg="#FFEEEE", fg=C_WARN, font=FONT_SMALL,
                          relief="flat", cursor="hand2", padx=6).pack(side="right")

        self._on_delete = on_delete or (lambda: None)

        row = tk.Frame(self.outer, bg=self.bg)
        row.pack(fill="x", padx=16, pady=(4, 14))

        tk.Label(row, text="목적지", bg=self.bg, font=FONT_SUB).pack(side="left")
        self.dest_var = tk.StringVar(value=list(AIRPREMIA_DESTINATIONS.keys())[1])  # 기본: 방콕
        ttk.Combobox(row, textvariable=self.dest_var, values=list(AIRPREMIA_DESTINATIONS.keys()),
                     state="readonly", width=16).pack(side="left", padx=(6, 20))

        tk.Label(row, text="박수", bg=self.bg, font=FONT_SUB).pack(side="left")
        self.nights_var = tk.StringVar(value=default_nights)
        ttk.Spinbox(row, from_=1, to=20, textvariable=self.nights_var, width=4,
                    font=FONT_BODY).pack(side="left", padx=(6, 2))
        tk.Label(row, text="박", bg=self.bg, font=FONT_BODY).pack(side="left", padx=(0, 20))

        tk.Label(row, text="인원수", bg=self.bg, font=FONT_SUB).pack(side="left")
        self.adt_var = tk.StringVar(value="4")
        ttk.Spinbox(row, from_=1, to=9, textvariable=self.adt_var, width=4,
                    font=FONT_BODY).pack(side="left", padx=(6, 2))
        tk.Label(row, text="명", bg=self.bg, font=FONT_BODY).pack(side="left", padx=(0, 20))

        self._hint = tk.Label(row, text="", bg=self.bg, fg=C_ACCENT2, font=("맑은 고딕", 8))
        self._hint.pack(side="left", padx=(10, 0))

        self._date_from_getter = lambda: None

    # ── 프리셋 export / import ──
    def export_preset_fields(self):
        return {
            "dest": self.dest_var.get(),
            "nights": self.nights_var.get(),
            "adt": self.adt_var.get(),
        }

    def import_preset_fields(self, data):
        if data.get("dest") in AIRPREMIA_DESTINATIONS:
            self.dest_var.set(data["dest"])
        if data.get("nights"):
            self.nights_var.set(data["nights"])
        if data.get("adt"):
            self.adt_var.set(data["adt"])

    # ── 힌트 (출발일 기준 귀국일 미리보기) ──
    def bind_date_from(self, getter_fn):
        self._date_from_getter = getter_fn
        self.nights_var.trace_add("write", lambda *a: self.update_hint())

    def update_hint(self):
        df_str = self._date_from_getter() if self._date_from_getter else None
        if not df_str:
            self._hint.configure(text="")
            return
        try:
            df = datetime.datetime.strptime(df_str.strip(), "%Y-%m-%d").date()
            n = int(self.nights_var.get())
            ret = df + datetime.timedelta(days=n)
            self._hint.configure(text=f"→ {df.strftime('%m/%d')} 출발 시 귀국: {ret.strftime('%m/%d')}")
        except Exception:
            self._hint.configure(text="")

    # 기존 ConditionBlock과 호환되게 update_return_hint 이름도 지원
    def update_return_hint(self, date_from_str):
        self._date_from_getter = lambda: date_from_str
        self.update_hint()

    # ── 검증 ──
    def validate_and_get(self, label_prefix=""):
        dest_label = self.dest_var.get()
        if dest_label not in AIRPREMIA_DESTINATIONS:
            messagebox.showerror("입력 오류", f"{label_prefix}목적지를 선택하세요.")
            return None
        try:
            nights = int(self.nights_var.get())
            assert 1 <= nights <= 20
        except Exception:
            messagebox.showerror("입력 오류", f"{label_prefix}박수를 올바르게 입력하세요 (1~20).")
            return None
        try:
            adt = int(self.adt_var.get())
            assert 1 <= adt <= 9
        except Exception:
            messagebox.showerror("입력 오류", f"{label_prefix}인원수를 올바르게 입력하세요 (1~9).")
            return None
        return {
            "dest_label": dest_label,
            "dest": AIRPREMIA_DESTINATIONS[dest_label],
            "nights": nights,
            "adt": adt,
        }


# ─────────────────────────────────────────────
#  메인 앱
# ─────────────────────────────────────────────
class GmarketAirApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("항공료 자동 추출기  |  투어로 김진모")
        self.geometry("1280x940")
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
        tk.Label(banner, text="✈  항공료 자동 추출기",
                 bg=C_ACCENT, fg="white", font=FONT_TITLE).pack(side="left", padx=18, pady=10)
        tk.Label(banner, text="투어로 김진모",
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
        canvas_window = main_canvas.create_window((0, 0), window=self._scroll_frame, anchor="nw")
        main_canvas.bind(
            "<Configure>",
            lambda e: main_canvas.itemconfig(canvas_window, width=e.width)
        )
        main_canvas.configure(yscrollcommand=scrollbar.set, xscrollcommand=scrollbar_x.set)
        scrollbar.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        main_canvas.pack(side="left", fill="both", expand=True)
        main_canvas.bind_all("<MouseWheel>", lambda e: main_canvas.yview_scroll(-1*(e.delta//120), "units"))
        main_canvas.bind_all("<Shift-MouseWheel>", lambda e: main_canvas.xview_scroll(-1*(e.delta//120), "units"))

        pad = {"padx": 16, "pady": 6}

        # ── 검색 대상 선택 ──
        target_row = tk.Frame(self._scroll_frame, bg=C_BG)
        target_row.pack(fill="x", padx=16, pady=(10, 0))
        tk.Label(target_row, text="🔎 검색 대상", bg=C_BG, fg=C_ACCENT, font=FONT_SUB).pack(side="left", padx=(0, 10))
        self.search_target_var = tk.StringVar(value="gmarket")
        tk.Radiobutton(target_row, text="G마켓 (전체 노선)", variable=self.search_target_var,
                       value="gmarket", bg=C_BG, font=FONT_BODY,
                       command=self._on_search_target_change).pack(side="left", padx=(0, 16))
        tk.Radiobutton(target_row, text="에어프레미아 (나리타/방콕)", variable=self.search_target_var,
                       value="airpremia", bg=C_BG, font=FONT_BODY,
                       command=self._on_search_target_change).pack(side="left")

        # ── 섹션 1: 노선 ──
        self._add_section_header("📍  1. 노선 설정")
        row1 = self._panel()

        # 출발지 (G마켓용 - 자유 선택)
        self._origin_label = tk.Label(row1, text="출발지", bg=C_PANEL, font=FONT_SUB)
        self._origin_label.grid(row=0, column=0, sticky="w", **pad)
        self.origin_country_var = tk.StringVar(value="국내")
        self._origin_country_cb = ttk.Combobox(row1, textvariable=self.origin_country_var,
                     values=list(AIRPORTS.keys()), state="readonly", width=8)
        self._origin_country_cb.grid(row=0, column=1, **pad)
        self.origin_city_var = tk.StringVar(value="인천")
        self._origin_city_cb = ttk.Combobox(row1, textvariable=self.origin_city_var,
                     values=list(AIRPORTS["국내"].keys()), state="readonly", width=10)
        self._origin_city_cb.grid(row=0, column=2, **pad)
        self._origin_country_cb.bind("<<ComboboxSelected>>", lambda e: self._on_country_change("origin"))

        # 출발지 (에어프레미아용 - 인천 고정 표시, 처음엔 숨김)
        self._origin_fixed_label = tk.Label(row1, text="출발지: 인천(ICN) 고정", bg=C_PANEL,
                                            font=FONT_SUB, fg=C_ACCENT2)

        tk.Label(row1, text="→", bg=C_PANEL, font=("맑은 고딕", 14)).grid(row=0, column=3)

        # 목적지 (G마켓용 - 자유 선택)
        self._dest_label = tk.Label(row1, text="목적지", bg=C_PANEL, font=FONT_SUB)
        self._dest_label.grid(row=0, column=4, sticky="w", **pad)
        self.dest_country_var = tk.StringVar(value="일본")
        self._dest_country_cb = ttk.Combobox(row1, textvariable=self.dest_country_var,
                     values=list(AIRPORTS.keys()), state="readonly", width=8)
        self._dest_country_cb.grid(row=0, column=5, **pad)
        self.dest_city_var = tk.StringVar(value="삿포로")
        self._dest_city_cb = ttk.Combobox(row1, textvariable=self.dest_city_var,
                     values=list(AIRPORTS["일본"].keys()), state="readonly", width=10)
        self._dest_city_cb.grid(row=0, column=6, **pad)
        self._dest_country_cb.bind("<<ComboboxSelected>>", lambda e: self._on_country_change("dest"))

        # 에어프레미아 모드 안내 (목적지는 각 조건 블록에서 선택)
        self._airpremia_dest_note = tk.Label(
            row1, text="목적지: 아래 검색 조건에서 세트별로 선택",
            bg=C_PANEL, font=FONT_SUB, fg=C_ACCENT2
        )

        # 인원수 (G마켓 전용 - 에어프레미아는 세트마다 각각 지정)
        self._adults_label = tk.Label(row1, text="인원수", bg=C_PANEL, font=FONT_SUB)
        self._adults_label.grid(row=0, column=7, sticky="w", **pad)
        self.adults_var = tk.IntVar(value=4)
        self._adults_spin = tk.Spinbox(row1, from_=1, to=9, textvariable=self.adults_var, width=4,
                   font=FONT_BODY, justify="center")
        self._adults_spin.grid(row=0, column=8, **pad)
        self._adults_unit_label = tk.Label(row1, text="명", bg=C_PANEL, font=FONT_SUB)
        self._adults_unit_label.grid(row=0, column=9, sticky="w")

        # ── 프리셋 (노선 + 항공사모드 + 시간대 저장/불러오기) ──
        preset_row = tk.Frame(row1, bg=C_PANEL)
        preset_row.grid(row=1, column=0, columnspan=10, sticky="w", padx=12, pady=(0, 8))
        tk.Label(preset_row, text="💾 프리셋", bg=C_PANEL, fg=C_ACCENT2,
                 font=FONT_SMALL).pack(side="left", padx=(0, 6))
        self._presets = load_presets()
        self.preset_var = tk.StringVar(value="")
        self._preset_cb = ttk.Combobox(preset_row, textvariable=self.preset_var,
                                       values=list(self._presets.keys()),
                                       state="normal", width=20)
        self._preset_cb.pack(side="left", padx=(0, 6))
        self._preset_cb.configure(
            postcommand=lambda: self._preset_cb.configure(values=list(self._presets.keys()))
        )
        tk.Button(preset_row, text="불러오기", command=self._load_preset,
                  bg=C_BTN, fg=C_BTN_FG, font=FONT_SMALL,
                  relief="flat", cursor="hand2", padx=8).pack(side="left", padx=2)
        tk.Button(preset_row, text="저장", command=self._save_preset,
                  bg="#E8F0FE", fg=C_ACCENT, font=FONT_SMALL,
                  relief="flat", cursor="hand2", padx=8).pack(side="left", padx=2)
        tk.Button(preset_row, text="삭제", command=self._delete_preset,
                  bg="#FFEEEE", fg=C_WARN, font=FONT_SMALL,
                  relief="flat", cursor="hand2", padx=8).pack(side="left", padx=2)
        tk.Button(preset_row, text="＋ 새 프리셋", command=lambda: self.preset_var.set(""),
                  bg="#EEEEEE", fg="#555555", font=FONT_SMALL,
                  relief="flat", cursor="hand2", padx=8).pack(side="left", padx=2)

        # ── 섹션 2: 출발 월 ──
        self._add_section_header("📅  2. 출발 월")
        date_panel = self._panel()
        now = datetime.date.today()
        default_y = now.year
        default_m = now.month + 1 if now.month < 12 else 1
        tk.Label(date_panel, text="출발일", bg=C_PANEL, font=FONT_SUB).grid(row=0, column=0, sticky="w", **pad)
        self.date_from_var = tk.StringVar(value=f"{default_y}-{default_m:02d}-01")
        tk.Entry(date_panel, textvariable=self.date_from_var, width=12, font=FONT_BODY,
                 relief="solid", bd=1, bg="#F8F9FB").grid(row=0, column=1, **pad)
        tk.Label(date_panel, text="~", bg=C_PANEL, font=("맑은 고딕", 12)).grid(row=0, column=2)
        self.date_to_var = tk.StringVar(value=f"{default_y}-{default_m:02d}-28")
        tk.Entry(date_panel, textvariable=self.date_to_var, width=12, font=FONT_BODY,
                 relief="solid", bd=1, bg="#F8F9FB").grid(row=0, column=3, **pad)
        tk.Label(date_panel, text="(날짜형식: 2026-07-10)", bg=C_PANEL, fg="#999999",
                 font=("맑은 고딕", 8)).grid(row=0, column=4, sticky="w", padx=(4,0))
        self.date_from_var.trace_add("write", lambda *a: self._refresh_all_hints())

        # ── 섹션 3: 검색 조건 (단일 모드 / 세트 모드 토글) ──
        self._add_section_header("✈  3. 검색 조건")

        # 단일 모드 컨테이너
        self._single_container = tk.Frame(self._scroll_frame, bg=C_BG)
        self._single_container.pack(fill="x")
        self._single_block = ConditionBlock(self._single_container, bg=C_PANEL,
                                            show_set_header=False, default_nights="3")
        self._single_block.bind_date_from(lambda: self.date_from_var.get())

        # 세트 모드 안내 + 전환 버튼 바
        self._toggle_bar = tk.Frame(self._scroll_frame, bg=C_BG)
        self._toggle_bar.pack(fill="x", padx=16, pady=(6, 0))
        self._toggle_label = tk.Label(
            self._toggle_bar,
            text="여러 조건의 항공료를 동시에 추출하고 싶다면 검색 조건 세트를 활성화하세요",
            bg=C_BG, fg=C_ACCENT2, font=FONT_SMALL
        )
        self._toggle_label.pack(side="left")
        self._toggle_btn = tk.Button(
            self._toggle_bar, text="📑  검색 조건 세트 활성화", command=self._toggle_set_mode,
            bg="#E8F0FE", fg=C_ACCENT, font=FONT_SMALL, relief="flat", cursor="hand2", padx=10, pady=4
        )
        self._toggle_btn.pack(side="right")

        # 세트 모드 컨테이너 (처음엔 숨김)
        self._set_mode = False
        self._set_blocks = []  # ConditionBlock 리스트
        self._sets_container = tk.Frame(self._scroll_frame, bg=C_BG)

        sets_hdr = tk.Frame(self._sets_container, bg=C_BG)
        sets_hdr.pack(fill="x", padx=16, pady=(8, 2))
        tk.Label(sets_hdr, text="세트별로 결과가 시트로 분리되어 같은 엑셀 파일에 저장됩니다.",
                 bg=C_BG, fg="#888888", font=("맑은 고딕", 8)).pack(side="left")
        tk.Button(sets_hdr, text="＋ 세트 추가", command=self._add_condition_set,
                  bg="#E8F0FE", fg=C_ACCENT, font=FONT_SMALL, relief="flat", cursor="hand2", padx=8
                  ).pack(side="right")

        self._sets_list_frame = tk.Frame(self._sets_container, bg=C_BG)
        self._sets_list_frame.pack(fill="x", padx=16)

        self._sets_container.pack_forget()  # 처음엔 단일 모드

        # ── 섹션 4: 저장 위치 ──
        self._add_section_header("💾  4. 저장 위치")
        row5 = self._panel()
        self.out_dir_var = tk.StringVar(value=get_default_output_dir())
        tk.Entry(row5, textvariable=self.out_dir_var, width=46, font=FONT_BODY,
                 bg="#F8F9FB", relief="solid", bd=1).grid(row=0, column=0, **pad)
        tk.Button(row5, text="폴더 선택", command=self._choose_dir,
                  bg=C_BTN, fg=C_BTN_FG, font=FONT_SMALL,
                  relief="flat", cursor="hand2", padx=8).grid(row=0, column=1, **pad)

        tk.Label(row5, text="파일명 (비워두면 자동 생성):", bg=C_PANEL, fg=C_ACCENT2,
                 font=FONT_SMALL).grid(row=1, column=0, sticky="w", padx=6, pady=(4, 0))
        self.custom_filename_var = tk.StringVar(value="")
        tk.Entry(row5, textvariable=self.custom_filename_var, width=46, font=FONT_BODY,
                 bg="#F8F9FB", relief="solid", bd=1
                 ).grid(row=2, column=0, columnspan=2, sticky="w", **pad)

        # ── 기타 옵션 ──
        self._add_section_header("⚙  5. 기타 옵션")
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
        self._log_msg("오류 등 문의 사항은 진모에게 문의 주세요.", "info")
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

    # ──────── 단일/세트 모드 전환 ────────
    def _toggle_set_mode(self):
        self._set_mode = not self._set_mode
        if self._set_mode:
            self._single_container.pack_forget()
            self._toggle_label.configure(text="단일 조건으로 검색하고 싶다면 세트 모드를 비활성화하세요")
            self._toggle_btn.configure(text="📑  단일 모드로 전환")
            self._sets_container.pack(fill="x", after=self._toggle_bar)
            if not self._set_blocks:
                self._add_condition_set()
        else:
            self._sets_container.pack_forget()
            self._toggle_label.configure(text="여러 조건의 항공료를 동시에 추출하고 싶다면 검색 조건 세트를 활성화하세요")
            self._toggle_btn.configure(text="📑  검색 조건 세트 활성화")
            self._single_container.pack(fill="x", before=self._toggle_bar)

    def _add_condition_set(self):
        idx = len(self._set_blocks) + 1
        def _make_delete(block_ref):
            def _del():
                block_ref.outer.destroy()
                self._set_blocks.remove(block_ref)
                self._relabel_condition_sets()
            return _del
        block_cls = ConditionBlock if self.search_target_var.get() == "gmarket" else AirpremiaConditionBlock
        block = block_cls(self._sets_list_frame, bg=C_LCC_BG if idx % 2 else C_FSC_BG,
                          show_set_header=True, set_index=idx,
                          on_delete=lambda: None, default_nights=str(2 + idx))
        block.on_delete_fn = _make_delete(block)
        self._rebind_delete_button(block)
        block.bind_date_from(lambda: self.date_from_var.get())
        self._set_blocks.append(block)

    def _rebind_delete_button(self, block):
        # 헤더 안의 삭제버튼을 찾아 command 재설정
        def find_and_bind(widget):
            for child in widget.winfo_children():
                if isinstance(child, tk.Button) and "세트 삭제" in (child.cget("text") or ""):
                    child.configure(command=block.on_delete_fn)
                    return True
                if find_and_bind(child):
                    return True
            return False
        find_and_bind(block.outer)

    def _relabel_condition_sets(self):
        for i, block in enumerate(self._set_blocks, 1):
            def find_and_relabel(widget, new_idx):
                for child in widget.winfo_children():
                    if isinstance(child, tk.Label):
                        txt = child.cget("text") or ""
                        if txt.startswith("세트 ") and len(txt) <= 5:
                            child.configure(text=f"세트 {new_idx}")
                        elif txt.isdigit():
                            child.configure(text=str(new_idx))
                    find_and_relabel(child, new_idx)
            find_and_relabel(block.outer, i)

    def _refresh_all_hints(self):
        df = self.date_from_var.get()
        self._single_block.update_return_hint(df)
        for block in self._set_blocks:
            block.update_return_hint(df)

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

    def _on_search_target_change(self):
        target = self.search_target_var.get()
        if target == "gmarket":
            self._origin_fixed_label.grid_remove()
            self._airpremia_dest_note.grid_remove()
            self._origin_label.grid()
            self._origin_country_cb.grid()
            self._origin_city_cb.grid()
            self._dest_label.grid()
            self._dest_country_cb.grid()
            self._dest_city_cb.grid()
            self._adults_label.grid()
            self._adults_spin.grid()
            self._adults_unit_label.grid()
        else:
            self._origin_label.grid_remove()
            self._origin_country_cb.grid_remove()
            self._origin_city_cb.grid_remove()
            self._dest_label.grid_remove()
            self._dest_country_cb.grid_remove()
            self._dest_city_cb.grid_remove()
            self._adults_label.grid_remove()
            self._adults_spin.grid_remove()
            self._adults_unit_label.grid_remove()
            self._origin_fixed_label.grid(row=0, column=0, columnspan=3, sticky="w", padx=16, pady=6)
            self._airpremia_dest_note.grid(row=0, column=4, columnspan=3, sticky="w", padx=16, pady=6)
        self._toggle_condition_ui_for_target()

    def _toggle_condition_ui_for_target(self):
        """검색 대상(G마켓/에어프레미아)에 맞는 조건 블록 클래스로 단일/세트 블록을 재구성"""
        target = self.search_target_var.get()
        block_cls = ConditionBlock if target == "gmarket" else AirpremiaConditionBlock

        # 단일 블록 재생성
        if hasattr(self, "_single_block"):
            self._single_block.outer.destroy()
        if target == "gmarket":
            self._single_block = ConditionBlock(self._single_container, bg=C_PANEL,
                                                show_set_header=False, default_nights="3")
        else:
            self._single_block = AirpremiaConditionBlock(self._single_container, bg=C_PANEL,
                                                          show_set_header=False, default_nights="3")
        self._single_block.bind_date_from(lambda: self.date_from_var.get())

        # 세트 블록 전부 재생성 (세트 개수는 유지)
        prev_count = max(len(self._set_blocks), 1)
        for block in list(self._set_blocks):
            block.outer.destroy()
        self._set_blocks = []
        for _ in range(prev_count):
            self._add_condition_set()

    # ──────── 프리셋 저장/불러오기/삭제 ────────
    def _current_route_and_condition(self):
        target = self.search_target_var.get()
        data = {
            "search_target": target,
            "adults": self.adults_var.get(),
            "set_mode": self._set_mode,
        }
        if target == "gmarket":
            data["origin_country"] = self.origin_country_var.get()
            data["origin_city"] = self.origin_city_var.get()
            data["dest_country"] = self.dest_country_var.get()
            data["dest_city"] = self.dest_city_var.get()
            if self._set_mode:
                data["sets"] = [
                    {"nights": block.return_offset_var.get(), **block.export_preset_fields()}
                    for block in self._set_blocks
                ]
            else:
                data.update(self._single_block.export_preset_fields())
        else:
            if self._set_mode:
                data["sets"] = [block.export_preset_fields() for block in self._set_blocks]
            else:
                data.update(self._single_block.export_preset_fields())
        return data

    def _save_preset(self):
        name = self.preset_var.get().strip()
        if not name:
            from tkinter import simpledialog
            name = simpledialog.askstring("프리셋 저장", "프리셋 이름을 입력하세요:", parent=self)
            if not name:
                return
            name = name.strip()
        if name in self._presets:
            if not messagebox.askyesno(
                "프리셋 덮어쓰기",
                f"'{name}' 프리셋이 이미 있습니다.\n현재 설정으로 덮어쓸까요?"
            ):
                return
        self._presets[name] = self._current_route_and_condition()
        save_presets(self._presets)
        self._preset_cb["values"] = list(self._presets.keys())
        self.preset_var.set(name)
        messagebox.showinfo("프리셋 저장", f"'{name}' 프리셋이 저장되었습니다.")

    def _load_preset(self):
        name = self.preset_var.get()
        if not name or name not in self._presets:
            messagebox.showwarning("프리셋 불러오기", "불러올 프리셋을 선택하세요.")
            return
        data = self._presets[name]

        preset_target = data.get("search_target", "gmarket")
        if preset_target != self.search_target_var.get():
            self.search_target_var.set(preset_target)
            self._on_search_target_change()

        if preset_target == "gmarket":
            oc = data.get("origin_country")
            if oc in AIRPORTS:
                self.origin_country_var.set(oc)
                self._origin_city_cb["values"] = list(AIRPORTS[oc].keys())
                oct_ = data.get("origin_city")
                if oct_ in AIRPORTS[oc]:
                    self.origin_city_var.set(oct_)

            dc = data.get("dest_country")
            if dc in AIRPORTS:
                self.dest_country_var.set(dc)
                self._dest_city_cb["values"] = list(AIRPORTS[dc].keys())
                dct_ = data.get("dest_city")
                if dct_ in AIRPORTS[dc]:
                    self.dest_city_var.set(dct_)

        adults = data.get("adults")
        if isinstance(adults, int) and 1 <= adults <= 9:
            self.adults_var.set(adults)

        preset_is_set_mode = data.get("set_mode", False)

        if preset_is_set_mode:
            if not self._set_mode:
                self._toggle_set_mode()
            sets_data = data.get("sets", [])
            # 기존 세트 블록 전부 제거 후 프리셋 개수만큼 새로 생성
            for block in list(self._set_blocks):
                block.outer.destroy()
            self._set_blocks = []
            for set_data in sets_data:
                self._add_condition_set()
                new_block = self._set_blocks[-1]
                if preset_target == "gmarket":
                    nights = set_data.get("nights")
                    if nights:
                        new_block.return_offset_var.set(nights)
                new_block.import_preset_fields(set_data)
            if not sets_data:
                self._add_condition_set()
        else:
            if self._set_mode:
                self._toggle_set_mode()
            self._single_block.import_preset_fields(data)

        messagebox.showinfo("프리셋 불러오기", f"'{name}' 프리셋을 불러왔습니다. (날짜는 직접 입력해주세요)")

    def _delete_preset(self):
        name = self.preset_var.get()
        if not name or name not in self._presets:
            messagebox.showwarning("프리셋 삭제", "삭제할 프리셋을 선택하세요.")
            return
        if not messagebox.askyesno("프리셋 삭제", f"'{name}' 프리셋을 삭제할까요?"):
            return
        del self._presets[name]
        save_presets(self._presets)
        self._preset_cb["values"] = list(self._presets.keys())
        self.preset_var.set("")

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
        if self.search_target_var.get() == "airpremia":
            return self._validate_airpremia()
        return self._validate_gmarket()

    def _validate_airpremia(self):
        import datetime as _dt
        try:
            date_from = _dt.datetime.strptime(self.date_from_var.get().strip(), "%Y-%m-%d").date()
            date_to   = _dt.datetime.strptime(self.date_to_var.get().strip(), "%Y-%m-%d").date()
            assert date_from <= date_to
        except Exception:
            messagebox.showerror("입력 오류", "출발일 형식(2026-07-10), 시작일≤종료일을 올바르게 입력하세요.")
            return None

        sets = []  # [{"label": str, "dest": str, "nights": int, "adt": int}, ...]

        if self._set_mode:
            if not self._set_blocks:
                messagebox.showerror("입력 오류", "검색 조건 세트를 1개 이상 추가하세요.")
                return None
            for i, block in enumerate(self._set_blocks, 1):
                result = block.validate_and_get(label_prefix=f"[세트{i}] ")
                if result is None:
                    return None
                label = f"세트{i}_{result['dest_label']}_{result['nights']}박"
                sets.append({
                    "label": label,
                    "dest": result["dest"],
                    "nights": result["nights"],
                    "adt": result["adt"],
                })
        else:
            result = self._single_block.validate_and_get()
            if result is None:
                return None
            sets.append({
                "label": "결과",
                "dest": result["dest"],
                "nights": result["nights"],
                "adt": result["adt"],
            })

        return {
            "search_target": "airpremia",
            "date_from": date_from, "date_to": date_to,
            "sets": sets,
            "multi_mode": self._set_mode,
            "show_browser": self.show_browser_var.get(),
            "out_dir": self.out_dir_var.get(),
            "custom_filename": self.custom_filename_var.get().strip(),
        }

    def _validate_gmarket(self):
        import datetime as _dt
        try:
            date_from = _dt.datetime.strptime(self.date_from_var.get().strip(), "%Y-%m-%d").date()
            date_to   = _dt.datetime.strptime(self.date_to_var.get().strip(), "%Y-%m-%d").date()
            assert date_from <= date_to
        except Exception:
            messagebox.showerror("입력 오류", "출발일 형식(2026-07-10), 시작일≤종료일을 올바르게 입력하세요.")
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

        sets = []  # [{"label": str, "return_offset": int, "extra_return": bool, "config": dict, "expand_priorities": list|None}, ...]

        if self._set_mode:
            if not self._set_blocks:
                messagebox.showerror("입력 오류", "검색 조건 세트를 1개 이상 추가하세요.")
                return None
            for i, block in enumerate(self._set_blocks, 1):
                result = block.validate_and_get(messagebox, label_prefix=f"[세트{i}] ")
                if result is None:
                    return None
                return_offset, extra_return, config, expand_priorities = result
                mode_label = config["airline_mode"]
                label = f"세트{i}_{return_offset+1}일_{mode_label}"
                sets.append({
                    "label": label,
                    "return_offset": return_offset,
                    "extra_return": extra_return,
                    "config": config,
                    "expand_priorities": expand_priorities,
                })
        else:
            result = self._single_block.validate_and_get(messagebox)
            if result is None:
                return None
            return_offset, extra_return, config, expand_priorities = result
            sets.append({
                "label": "결과",
                "return_offset": return_offset,
                "extra_return": extra_return,
                "config": config,
                "expand_priorities": expand_priorities,
            })

        return {
            "search_target": "gmarket",
            "origin": origin, "dest": dest,
            "date_from": date_from, "date_to": date_to,
            "sets": sets,
            "multi_mode": self._set_mode,
            "show_browser": self.show_browser_var.get(),
            "out_dir": self.out_dir_var.get(),
            "custom_filename": self.custom_filename_var.get().strip(),
            "adults": self.adults_var.get(),
        }

    # ──────── 실행 ────────
    def _start(self):
        target = self.search_target_var.get()
        if target == "gmarket" and not SCRAPER_OK:
            messagebox.showerror("오류", "scraper_core.py가 없어 실행할 수 없습니다.")
            return
        if target == "airpremia" and not AIRPREMIA_OK:
            messagebox.showerror("오류", "airpremia_fare_checker.py가 없어 실행할 수 없습니다.")
            return
        params = self._validate()
        if params is None:
            return
        self._running = True
        self._run_btn.configure(state="disabled")
        self._stop_btn.configure(state="normal")
        self._progressbar["value"] = 0
        self._prog_label.configure(text="실행 중...")
        target_fn = self._run_scrape if target == "gmarket" else self._run_airpremia_scrape
        self._thread = threading.Thread(target=target_fn, args=(params,), daemon=True)
        self._thread.start()

    def _stop(self):
        self._running = False
        self._log_msg("⏹ 중지 요청됨. 현재 날짜 완료 후 종료됩니다.", "err")

    def _run_scrape(self, p):
        origin    = p["origin"]
        dest      = p["dest"]
        date_from = p["date_from"]
        date_to   = p["date_to"]
        show      = p["show_browser"]
        out_dir   = p["out_dir"]
        sets      = p["sets"]
        multi_mode = p["multi_mode"]
        custom_filename = p.get("custom_filename", "")
        adults    = p.get("adults", 4)

        from datetime import timedelta
        from scraper_core import (init_driver, build_url, fetch_flights,
                                  select_best, parse_price, calc_per_person, in_band,
                                  reset_debug, save_excel, save_excel_multi)
        reset_debug()

        date_list = []
        d = date_from
        while d <= date_to:
            date_list.append(d)
            d += timedelta(days=1)
        total_days = len(date_list)
        total_work = total_days * len(sets)

        if custom_filename:
            safe_name = re.sub(r'[\\/:*?"<>|]', "_", custom_filename)
            if not safe_name.lower().endswith(".xlsx"):
                safe_name += ".xlsx"
            fname = safe_name
        else:
            fname = f"gmarket_{origin}_{dest}_{date_from.strftime('%Y%m%d')}_{date_to.strftime('%Y%m%d')}.xlsx"
        out_path = os.path.join(out_dir, fname)

        self._log_msg(
            f"🛫 추출 시작: {origin} → {dest}  |  {date_from} ~ {date_to}  |  세트 {len(sets)}개  |  총 {total_days}일",
            "info"
        )

        try:
            driver = init_driver(show=show)
            rows_by_set = []
            work_done = 0

            for set_idx, s in enumerate(sets, 1):
                config             = s["config"]
                expand_priorities  = s["expand_priorities"]
                base_offset        = s["return_offset"]
                extra_return       = s["extra_return"]
                offsets            = [base_offset, base_offset - 1] if extra_return and base_offset > 0 else [base_offset]
                label              = s["label"]
                set_prefix         = f"[{label}] " if multi_mode else ""

                if multi_mode:
                    self._log_msg(f"\n▶ {label} 검색 시작 (귀국 출발+{base_offset}일)", "info")
                if expand_priorities:
                    self._log_msg(f"  🔄 {set_prefix}대체 조건 모드: {len(expand_priorities)}개 순위 적용", "info")

                rows = []

                for idx, dep_date in enumerate(date_list, 1):
                    if not self._running:
                        self._log_msg("⏹ 중지됨.", "err")
                        break

                    self._log_msg(
                        f"  {set_prefix}{dep_date.strftime('%Y-%m-%d')} ({dep_date.strftime('%a')}) 검색 중..."
                    )

                    best = None
                    used_offset = base_offset
                    used_priority = None

                    if expand_priorities:
                        for pri_idx, pri in enumerate(expand_priorities):
                            n = pri["nights"]
                            off = n
                            arr_date_try = dep_date + timedelta(days=off)
                            url = build_url(origin, dest,
                                            dep_date.strftime("%Y%m%d"),
                                            arr_date_try.strftime("%Y%m%d"),
                                            adults=adults)
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
                        best = None
                        best_arr_date = dep_date + timedelta(days=base_offset)
                        best_total = None
                        for off in offsets:
                            arr_try = dep_date + timedelta(days=off)
                            url = build_url(origin, dest,
                                            dep_date.strftime("%Y%m%d"),
                                            arr_try.strftime("%Y%m%d"),
                                            adults=adults)
                            flights = fetch_flights(driver, url, self._log_msg)
                            cand = select_best(flights, config) if flights else None
                            if cand:
                                cand_total = parse_price(cand.get("cardPrice", "0"))
                                if best is None or (cand_total and cand_total < best_total):
                                    best = cand
                                    best_total = cand_total
                                    best_arr_date = arr_try
                        arr_date = best_arr_date

                    if best:
                        actual_price = best.get("cardPrice", "")
                        best["price"] = actual_price
                        total4 = parse_price(actual_price)
                        per1   = calc_per_person(total4, adults=adults)
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
                            f"    ✔{pri_note} {set_prefix}{best['airline']}  {best['dep']}→{best['arr']}  {adults}인:{total4:,}원  1인:{per1:,}원",
                            "ok"
                        )
                    else:
                        rows.append({
                            "dep_date": dep_date.strftime("%Y-%m-%d"),
                            "arr_date": (dep_date + timedelta(days=base_offset)).strftime("%Y-%m-%d"),
                            "airline":"", "dep":"","arr":"","rDep":"","rArr":"",
                            "total4":0,"per1":0,"seller":"","found":False,
                        })
                        self._log_msg(f"    ✗ {set_prefix}조건에 맞는 항공편 없음")

                    work_done += 1
                    self._update_progress(work_done, total_work)

                rows_by_set.append({"label": label, "rows": rows})

                if not self._running:
                    break

            driver.quit()

            total_rows = sum(len(entry["rows"]) for entry in rows_by_set)
            if total_rows > 0:
                if multi_mode:
                    save_excel_multi(rows_by_set, origin, dest, out_path, adults=adults)
                else:
                    only = rows_by_set[0]
                    save_excel(only["rows"], origin, dest, date_from.year, date_from.month, out_path, adults=adults)

                summary_parts = []
                for entry in rows_by_set:
                    found_cnt = sum(1 for r in entry["rows"] if r["found"])
                    summary_parts.append(f"{entry['label']}: {found_cnt}/{len(entry['rows'])}일")
                summary = "  /  ".join(summary_parts)

                self._log_msg(f"\n✅ 완료! {summary} → {fname}", "ok")
                self.after(0, lambda: messagebox.showinfo(
                    "완료",
                    f"추출 완료!\n{date_from} ~ {date_to}\n{summary}\n저장: {out_path}"
                ))
                try:
                    from plyer import notification
                    notification.notify(
                        title="✈ G마켓 항공료 추출 완료",
                        message=f"{origin} → {dest}  완료",
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

    def _run_airpremia_scrape(self, p):
        from datetime import timedelta as _timedelta
        from airpremia_fare_checker import init_driver as _ap_init_driver, check_one_date as _ap_check_one_date

        date_from = p["date_from"]
        date_to   = p["date_to"]
        sets      = p["sets"]
        show      = p["show_browser"]
        out_dir   = p["out_dir"]
        custom_filename = p.get("custom_filename", "")

        date_list = []
        d = date_from
        while d <= date_to:
            date_list.append(d)
            d += _timedelta(days=1)
        total_work = len(date_list) * len(sets)

        if custom_filename:
            safe_name = re.sub(r'[\\/:*?"<>|]', "_", custom_filename)
            if not safe_name.lower().endswith(".xlsx"):
                safe_name += ".xlsx"
            fname = safe_name
        else:
            dests = "_".join(sorted({s["dest"] for s in sets}))
            fname = f"airpremia_{dests}_{date_from.strftime('%Y%m%d')}_{date_to.strftime('%Y%m%d')}.xlsx"
        out_path = os.path.join(out_dir, fname)

        self._log_msg(
            f"🛫 추출 시작(에어프레미아): {len(sets)}개 세트  |  {date_from} ~ {date_to}  |  총 {len(date_list)}일",
            "info",
        )

        rows_by_set = []
        driver = None
        try:
            driver = _ap_init_driver(headless=not show)

            work_done = 0
            for s in sets:
                label = s["label"]
                self._log_msg(f"\n▶ {label} 검색 시작 (목적지={s['dest']}, {s['nights']}박, {s['adt']}명)", "info")
                rows = []

                for dep_date in date_list:
                    if not self._running:
                        self._log_msg("⏹ 중지됨.", "err")
                        break

                    self._log_msg(f"  {dep_date.strftime('%Y-%m-%d')} ({dep_date.strftime('%a')}) 조회 중...")
                    result = _ap_check_one_date(driver, s["dest"], dep_date, s["nights"], s["adt"])
                    rows.append(result)

                    if result.status == "OK":
                        self._log_msg(
                            f"    ✔ {result.departure_date}→{result.return_date}  "
                            f"총액:{result.total_amount:,}원  1인:{result.per_person_rounded:,}원",
                            "ok",
                        )
                    else:
                        self._log_msg(
                            f"    ✗ {result.departure_date}→{result.return_date}  "
                            f"{result.status} {result.note}"
                        )

                    work_done += 1
                    self._update_progress(work_done, total_work)

                rows_by_set.append({"label": label, "rows": rows, "adt": s["adt"]})
                if not self._running:
                    break

            self._save_airpremia_excel_multi(rows_by_set, out_path)

            summary = "  /  ".join(
                f"{e['label']}: {sum(1 for r in e['rows'] if r.status == 'OK')}/{len(e['rows'])}일"
                for e in rows_by_set
            )
            self._log_msg(f"\n✅ 완료! {summary} → {fname}", "ok")
            self.after(0, lambda: messagebox.showinfo(
                "완료", f"추출 완료!\n{date_from} ~ {date_to}\n{summary}\n저장: {out_path}"
            ))

        except Exception as e:
            import traceback
            self._log_msg(f"❌ 오류 발생: {e}\n{traceback.format_exc()}", "err")
            self.after(0, lambda: messagebox.showerror("오류", str(e)))

        finally:
            if driver is not None:
                try:
                    driver.quit()
                except Exception:
                    pass
            self._running = False
            self.after(0, lambda: self._run_btn.configure(state="normal"))
            self.after(0, lambda: self._stop_btn.configure(state="disabled"))
            self.after(0, lambda: self._prog_label.configure(text="완료"))

    @staticmethod
    def _save_airpremia_excel_multi(rows_by_set, out_path: str):
        from openpyxl import Workbook
        from airpremia_fare_checker import write_styled_sheet as _ap_write_styled_sheet

        wb = Workbook()
        wb.remove(wb.active)

        for entry in rows_by_set:
            sheet_name = re.sub(r"[\\/*?:\[\]]", "_", entry["label"])[:31] or "결과"
            ws = wb.create_sheet(title=sheet_name)
            _ap_write_styled_sheet(ws, entry["rows"], entry["adt"])
        wb.save(out_path)


# ─────────────────────────────────────────────
#  진입점
# ─────────────────────────────────────────────
if __name__ == "__main__":
    app = GmarketAirApp()
    app.mainloop()
