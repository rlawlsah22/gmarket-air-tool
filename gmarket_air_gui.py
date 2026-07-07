"""
gmarket_air_gui.py  —  리디자인: 좌측 사이드바 + 우측 메인 (customtkinter)
"""
import sys, subprocess, os

# ── customtkinter 자동 설치 (없을 경우) ──────────────────────
def _ensure_customtkinter():
    try:
        import customtkinter
    except ImportError:
        import ctypes
        # 설치 중 안내 (tkinter 대신 ctypes 메시지박스 사용)
        ctypes.windll.user32.MessageBoxW(
            0,
            "필요한 패키지(customtkinter)를 설치합니다.\n확인을 누르면 설치가 시작됩니다.\n설치 완료 후 프로그램을 다시 실행해 주세요.",
            "초기 설정",
            0x40  # MB_OK | MB_ICONINFORMATION
        )
        try:
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", "customtkinter",
                 "--quiet", "--disable-pip-version-check"],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            ctypes.windll.user32.MessageBoxW(
                0,
                "설치가 완료되었습니다.\n프로그램을 다시 실행해 주세요.",
                "설치 완료",
                0x40
            )
        except Exception as e:
            ctypes.windll.user32.MessageBoxW(
                0,
                f"설치 중 오류가 발생했습니다.\n진모에게 문의해주세요.\n\n오류: {e}",
                "오류",
                0x10
            )
        sys.exit(0)

_ensure_customtkinter()

import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import threading, os, datetime, sys

# ── 자동 업데이트 ─────────────────────────────────────────
CURRENT_VERSION = "1.0"
REPO_RAW = "https://raw.githubusercontent.com/touroro-team3/gmarket-air/main"

def check_and_update():
    try:
        import urllib.request
        with urllib.request.urlopen(f"{REPO_RAW}/version.txt", timeout=4) as r:
            latest = r.read().decode().strip()
        if latest <= CURRENT_VERSION:
            return
        for fn in ["gmarket_air_gui.py", "scraper_core.py"]:
            with urllib.request.urlopen(f"{REPO_RAW}/{fn}", timeout=10) as r:
                data = r.read()
            base = os.path.dirname(sys.executable if getattr(sys,'frozen',False) else os.path.abspath(__file__))
            with open(os.path.join(base, fn), "wb") as f:
                f.write(data)
        messagebox.showinfo("업데이트 완료", f"v{latest}로 업데이트되었습니다.\n프로그램을 재실행해주세요.")
        sys.exit(0)
    except Exception:
        pass

check_and_update()

# ── scraper_core import ───────────────────────────────────
_here = os.path.dirname(os.path.abspath(__file__)) if "__file__" in dir() else ""
if getattr(sys,"frozen",False):
    _d = os.path.dirname(sys.executable)
    if _d not in sys.path: sys.path.insert(0, _d)
elif _here and _here not in sys.path:
    sys.path.insert(0, _here)

try:
    from scraper_core import (
        AIRPORTS, LCC_ALL, FSC_ALL, LCC_TIER1, LCC_TIER2, FOREIGN_ALL,
        save_excel, save_excel_multi
    )
    SCRAPER_OK = True
except ImportError:
    SCRAPER_OK = False
    AIRPORTS = {
        "국내": {"인천":"ICN","부산":"PUS","김포":"GMP","청주":"CJJ","대구":"TAE"},
        "일본": {"삿포로":"CTS","도쿄(나리타)":"NRT","도쿄(하네다)":"HND","오사카":"OSA","후쿠오카":"FUK","오키나와":"OKA"},
        "태국": {"방콕":"BKK","치앙마이":"CNX","푸켓":"HKT"},
        "베트남": {"하노이":"HAN","다낭":"DAD","나트랑":"CXR","호치민":"SGN","푸꾸옥":"PQC"},
        "필리핀": {"마닐라":"MNL","클락":"CRK","세부":"CEB","보라카이":"MPH"},
        "중국": {"북경":"BJS","천진":"TSN","청도":"TAO","연태":"YNT","대련":"DLC"},
    }
    LCC_ALL=FSC_ALL=LCC_TIER1=LCC_TIER2=FOREIGN_ALL=[]

# ── CTk 테마 ──────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ── 팔레트 ────────────────────────────────────────────────
SB_BG    = "#111318"   # 사이드바
MAIN_BG  = "#16181d"   # 메인 배경
CARD_BG  = "#1e2028"   # 카드
CARD2_BG = "#252830"   # 인풋/서브카드
BORDER   = "#2e3140"
ACCENT   = "#3b82f6"
ACCENT_H = "#60a5fa"
TEXT1    = "#f1f3f9"
TEXT2    = "#c9d1de"
TEXT3    = "#8b95a5"
OK       = "#34d399"
ERR      = "#f87171"
INFO     = "#60a5fa"
WARN     = "#fbbf24"

FNT      = "맑은 고딕"

# ── 상수 ──────────────────────────────────────────────────
AIRLINE_MODES = [
    "LCC 우선 → FSC 대체","LCC만",
    "FSC만 (아시아나/대한항공)","외항사만","특정 항공사 지정",
]
TIME_SLOT_LABELS = ["새벽 (00~06시)","오전 (06~12시)","오후 (12~18시)","야간 (18~24시)"]
TIME_SLOT_KEYS   = ["새벽","오전","오후","야간"]
SPECIFIC_AIRLINES = [
    "대한항공","아시아나항공","에어프레미아","진에어","제주항공",
    "티웨이항공","이스타항공","에어부산","에어서울","에어로케이",
    "파라타항공","타이항공","필리핀항공","베트남항공",
    "중국국제항공","중국남방항공","중국동방항공","산동항공",
]

def get_default_output_dir():
    return os.path.join(os.path.expanduser("~"), "Desktop")

# ══════════════════════════════════════════════════════════
#  헬퍼 위젯
# ══════════════════════════════════════════════════════════
def Label(parent, text, size=13, bold=False, color=None, **kw):
    return ctk.CTkLabel(parent, text=text,
                        font=(FNT, size, "bold" if bold else "normal"),
                        text_color=color or TEXT2, **kw)

def Entry(parent, var, w=120, **kw):
    return ctk.CTkEntry(parent, textvariable=var, width=w, height=32,
                        fg_color=CARD2_BG, border_color=BORDER,
                        text_color=TEXT1, font=(FNT,13), **kw)

def Combo(parent, var, values, w=200, cmd=None, **kw):
    return ctk.CTkComboBox(parent, variable=var, values=values,
                           width=w, height=32, state="readonly",
                           fg_color=CARD2_BG, border_color=BORDER,
                           button_color=ACCENT, button_hover_color=ACCENT_H,
                           dropdown_fg_color=CARD_BG, dropdown_hover_color=CARD2_BG,
                           text_color=TEXT1, dropdown_text_color=TEXT1,
                           font=(FNT,13), command=cmd or (lambda v: None), **kw)

def Check(parent, var, text, cmd=None, **kw):
    return ctk.CTkCheckBox(parent, variable=var, text=text,
                           font=(FNT,13), text_color=TEXT2,
                           fg_color=ACCENT, hover_color=ACCENT_H,
                           border_color=BORDER, checkmark_color="white",
                           command=cmd or (lambda: None), **kw)

def Divider(parent, **kw):
    return ctk.CTkFrame(parent, height=1, fg_color=BORDER, **kw)

def Card(parent, **kw):
    return ctk.CTkFrame(parent, fg_color=CARD_BG, corner_radius=10,
                        border_width=1, border_color=BORDER, **kw)

def SmallBtn(parent, text, cmd, color=CARD2_BG, tc=TEXT2, hc=BORDER,
             w=80, h=28, **kw):
    return ctk.CTkButton(parent, text=text, command=cmd,
                         fg_color=color, hover_color=hc,
                         text_color=tc, border_width=1, border_color=BORDER,
                         font=(FNT,12), width=w, height=h, **kw)

# ══════════════════════════════════════════════════════════
#  ConditionBlock
# ══════════════════════════════════════════════════════════
class ConditionBlock:
    def __init__(self, parent, show_set_header=False, set_index=1,
                 on_delete=None, default_nights="3", default_mode=None):
        self.outer = Card(parent)
        self.outer.pack(fill="x", pady=(0,8))

        if show_set_header:
            hdr = ctk.CTkFrame(self.outer, fg_color="transparent")
            hdr.pack(fill="x", padx=16, pady=(12,0))
            Label(hdr, f"세트 {set_index}", size=13, bold=True, color=ACCENT).pack(side="left")
            if on_delete:
                ctk.CTkButton(hdr, text="삭제", command=on_delete,
                              width=52, height=24, fg_color="transparent",
                              border_width=1, border_color="#3d1515",
                              text_color=ERR, hover_color="#1a0808",
                              font=(FNT,12)).pack(side="right")

        # 여행일수 + 항공사 모드
        row1 = ctk.CTkFrame(self.outer, fg_color="transparent")
        row1.pack(fill="x", padx=16, pady=(14,8))
        Label(row1, "여행 일수", size=12, color=TEXT2).pack(side="left")
        self.return_offset_var = tk.StringVar(value=default_nights)
        Entry(row1, self.return_offset_var, w=56).pack(side="left", padx=(8,4))
        Label(row1, "일", size=12, color=TEXT3).pack(side="left", padx=(0,20))
        Label(row1, "항공사 모드", size=12, color=TEXT2).pack(side="left")
        self.airline_mode_var = tk.StringVar(value=default_mode or AIRLINE_MODES[0])
        Combo(row1, self.airline_mode_var, AIRLINE_MODES, w=220,
              cmd=lambda v: self._on_mode_change()).pack(side="left", padx=(8,0))

        # 특정 항공사
        self._specific_frame = ctk.CTkFrame(self.outer, fg_color="transparent")
        self._specific_frame.pack(fill="x", padx=16, pady=(0,6))
        Label(self._specific_frame, "항공사 선택 (복수 가능):", size=11, color=TEXT3
              ).grid(row=0, column=0, columnspan=6, sticky="w", pady=(0,4))
        self._airline_vars = {}
        for i, al in enumerate(SPECIFIC_AIRLINES):
            v = tk.BooleanVar(value=False)
            self._airline_vars[al] = v
            Check(self._specific_frame, v, al, width=110
                  ).grid(row=1+(i//6), column=i%6, sticky="w", padx=2, pady=1)
        self._specific_frame.pack_forget()

        # +1일
        self.extra_return_var = tk.BooleanVar(value=False)
        Check(self.outer, self.extra_return_var,
              "+1일도 함께 검색 (기내1박 새벽도착편 비교용)",
              cmd=self._update_return_hint
              ).pack(anchor="w", padx=16, pady=(0,2))
        self._return_hint = Label(self.outer, "", size=10, color=INFO)
        self._return_hint.pack(anchor="w", padx=20, pady=(0,8))
        self.return_offset_var.trace_add("write", lambda *a: self._update_return_hint())
        self._date_from_getter = lambda: None

        Divider(self.outer).pack(fill="x", padx=16, pady=(0,12))

        # 시간대 조건
        Label(self.outer, "시간대 조건", size=12, color=ACCENT
              ).pack(anchor="w", padx=16, pady=(0,10))

        band_frame = ctk.CTkFrame(self.outer, fg_color="transparent")
        band_frame.pack(fill="x", padx=16, pady=(0,8))
        for c in range(4): band_frame.columnconfigure(c, weight=1)

        band_cols = [("가는편 출발","dep_band"),("가는편 도착","arr_band"),
                     ("오는편 출발","ret_dep_band"),("오는편 도착","ret_arr_band")]
        self._band_vars={} ; self._band_all_vars={}
        self._band_custom_from={} ; self._band_custom_to={}

        for col,(lbl,key) in enumerate(band_cols):
            Label(band_frame, lbl, size=11, color=TEXT2, anchor="center"
                  ).grid(row=0, column=col, padx=8, pady=(0,6), sticky="ew")

        for col,(lbl,key) in enumerate(band_cols):
            fd = ctk.CTkFrame(band_frame, fg_color="transparent")
            fd.grid(row=1, column=col, padx=8, pady=(0,4), sticky="ew")
            Label(fd, "직접입력:", size=10, color=TEXT3).pack(side="left")
            fv, tv = tk.StringVar(), tk.StringVar()
            self._band_custom_from[key]=fv ; self._band_custom_to[key]=tv
            for var in (fv, tv):
                ctk.CTkEntry(fd, textvariable=var, width=48, height=24,
                             fg_color=CARD2_BG, border_color=BORDER,
                             text_color=TEXT1, font=(FNT,11)
                             ).pack(side="left", padx=2)
                if var is fv:
                    Label(fd, "~", size=10, color=TEXT3).pack(side="left")

        Divider(band_frame).grid(row=2, column=0, columnspan=4,
                                 sticky="ew", padx=8, pady=6)

        for col,(lbl,key) in enumerate(band_cols):
            av = tk.BooleanVar(value=True)
            self._band_all_vars[key]=av
            def _make_all_cb(k,v):
                def _t():
                    for sv in self._band_vars[k]: sv.set(v.get())
                return _t
            Check(band_frame, av, "전체", cmd=_make_all_cb(key,av)
                  ).grid(row=3, column=col, padx=8, sticky="ew", pady=(0,2))
            slot_vars=[]
            for r2,slot_lbl in enumerate(TIME_SLOT_LABELS):
                v2=tk.BooleanVar(value=True) ; slot_vars.append(v2)
                def _make_slot_cb(k,svs,av2):
                    def _t(): av2.set(all(sv.get() for sv in svs))
                    return _t
                Check(band_frame, v2, slot_lbl,
                      cmd=_make_slot_cb(key,slot_vars,av)
                      ).grid(row=r2+4, column=col, padx=8, sticky="ew", pady=2)
            self._band_vars[key]=slot_vars

        Divider(self.outer).pack(fill="x", padx=16, pady=(4,12))

        # 대체 조건
        exp_wrap = ctk.CTkFrame(self.outer, fg_color="transparent")
        exp_wrap.pack(fill="x", padx=16, pady=(0,14))
        self.use_expand_var = tk.BooleanVar(value=False)
        Check(exp_wrap, self.use_expand_var,
              "대체 조건 자동 확장 사용 (체크 시 아래 순위 조건으로 검색, 위 일수 설정 무시)",
              cmd=self._on_expand_toggle).pack(anchor="w")

        self._expand_frame = ctk.CTkFrame(exp_wrap, fg_color="transparent")
        self._expand_frame.pack(fill="x", pady=(8,0))
        hdr_row = ctk.CTkFrame(self._expand_frame, fg_color="transparent")
        hdr_row.pack(fill="x")
        for h in ["순위","여행 일수(일)","도착~(상한)","귀국출발(이후~이전)","도착시간대",""]:
            Label(hdr_row, h, size=10, color=ACCENT, width=10, anchor="w"
                  ).pack(side="left", padx=2)
        self._expand_rows_frame = ctk.CTkFrame(self._expand_frame, fg_color="transparent")
        self._expand_rows_frame.pack(fill="x")
        self._expand_rows=[]
        self._add_row_btn = ctk.CTkButton(
            exp_wrap, text="＋ 순위 추가", command=self._add_expand_row,
            width=100, height=26, fg_color="transparent",
            border_width=1, border_color=ACCENT, text_color=ACCENT,
            hover_color="#0d2137", font=(FNT,12))
        self._add_row_btn.pack(anchor="w", pady=(6,0))
        self._expand_frame.pack_forget()
        self._add_row_btn.pack_forget()

    def _on_mode_change(self, _=None):
        if self.airline_mode_var.get() == "특정 항공사 지정":
            self._specific_frame.pack(fill="x", padx=16, pady=(0,6))
        else:
            self._specific_frame.pack_forget()

    def _on_expand_toggle(self):
        if self.use_expand_var.get():
            self._expand_frame.pack(fill="x", pady=(4,0))
            self._add_row_btn.pack(anchor="w", pady=(6,0))
            if not self._expand_rows:
                self._add_expand_row("2","10:10","19:00")
                self._add_expand_row("3","10:10","19:00")
        else:
            self._expand_frame.pack_forget()
            self._add_row_btn.pack_forget()

    def _add_expand_row(self, nights_default="3", arr_to="", ret_dep_from="",
                        ret_dep_to="", arr_band_default="전체"):
        row_idx = len(self._expand_rows)+1
        frame = ctk.CTkFrame(self._expand_rows_frame, fg_color="transparent")
        frame.pack(fill="x", pady=2)
        Label(frame, f"{row_idx}순위", size=10, color=TEXT3, width=5).pack(side="left",padx=(0,2))
        nights_var = tk.StringVar(value=nights_default)
        ctk.CTkEntry(frame, textvariable=nights_var, width=36, height=24,
                     fg_color=CARD2_BG, border_color=BORDER,
                     text_color=TEXT1, font=(FNT,11)).pack(side="left")
        Label(frame,"일",size=10,color=TEXT3).pack(side="left",padx=(2,8))
        arr_to_var=tk.StringVar(value=arr_to)
        Label(frame,"도착~",size=10,color=TEXT3).pack(side="left")
        ctk.CTkEntry(frame,textvariable=arr_to_var,width=52,height=24,
                     fg_color=CARD2_BG,border_color=BORDER,
                     text_color=TEXT1,font=(FNT,11)).pack(side="left",padx=(2,8))
        ret_dep_from_var=tk.StringVar(value=ret_dep_from)
        Label(frame,"귀국출발",size=10,color=TEXT3).pack(side="left")
        ctk.CTkEntry(frame,textvariable=ret_dep_from_var,width=46,height=24,
                     fg_color=CARD2_BG,border_color=BORDER,
                     text_color=TEXT1,font=(FNT,11)).pack(side="left",padx=(2,1))
        Label(frame,"~",size=10,color=TEXT3).pack(side="left")
        ret_dep_to_var=tk.StringVar(value=ret_dep_to)
        ctk.CTkEntry(frame,textvariable=ret_dep_to_var,width=46,height=24,
                     fg_color=CARD2_BG,border_color=BORDER,
                     text_color=TEXT1,font=(FNT,11)).pack(side="left",padx=(1,8))
        arr_band_var=tk.StringVar(value=arr_band_default)
        Combo(frame,arr_band_var,["전체","오전","오후","야간","오후+야간"],w=100
              ).pack(side="left",padx=(2,8))
        row_data={"frame":frame,"nights":nights_var,"arr_to":arr_to_var,
                  "ret_dep_from":ret_dep_from_var,"ret_dep_to":ret_dep_to_var,
                  "arr_band":arr_band_var}
        def _del(rd=row_data):
            rd["frame"].destroy()
            self._expand_rows.remove(rd)
            for i,r in enumerate(self._expand_rows):
                for w in r["frame"].winfo_children():
                    if isinstance(w,ctk.CTkLabel) and "순위" in (w.cget("text") or ""):
                        w.configure(text=f"{i+1}순위"); break
        ctk.CTkButton(frame,text="✕",command=_del,width=28,height=24,
                      fg_color="transparent",border_width=1,border_color="#3d1515",
                      text_color=ERR,hover_color="#1a0808",font=(FNT,11)
                      ).pack(side="left")
        self._expand_rows.append(row_data)

    def update_return_hint(self, date_from_str):
        import datetime as _dt
        try:
            df=_dt.datetime.strptime(date_from_str.strip(),"%Y-%m-%d").date()
            off=int(self.return_offset_var.get())-1
            d1=df+_dt.timedelta(days=off)
            if self.extra_return_var.get():
                d2=df+_dt.timedelta(days=off+1)
                txt=f"  → {df.strftime('%m/%d')} 출발: {d1.strftime('%m/%d')}, {d2.strftime('%m/%d')} 둘 다 검색"
            else:
                txt=f"  → {df.strftime('%m/%d')} 출발: {d1.strftime('%m/%d')} 검색"
            self._return_hint.configure(text=txt)
        except: self._return_hint.configure(text="")

    def _update_return_hint(self):
        getter=getattr(self,"_date_from_getter",None)
        if getter:
            val=getter()
            if val: self.update_return_hint(val)

    def bind_date_from(self, getter_fn):
        self._date_from_getter=getter_fn

    def validate_and_get(self, label_prefix=""):
        try:
            ro=int(self.return_offset_var.get())-1; assert 1<=ro<=20
        except:
            messagebox.showerror("입력 오류",f"{label_prefix}여행 일수를 올바르게 입력하세요.")
            return None
        mode_map={"LCC 우선 → FSC 대체":"LCC우선_FSC대체","LCC만":"LCC만",
                  "FSC만 (아시아나/대한항공)":"FSC만","외항사만":"외항사만",
                  "특정 항공사 지정":"특정항공사"}
        am=mode_map[self.airline_mode_var.get()]
        specific=[]
        if am=="특정항공사":
            specific=[al for al,v in self._airline_vars.items() if v.get()]
            if not specific:
                messagebox.showerror("입력 오류",f"{label_prefix}특정 항공사를 하나 이상 선택하세요.")
                return None
        def _band(key):
            fs=self._band_custom_from[key].get().strip()
            ts=self._band_custom_to[key].get().strip()
            if fs or ts:
                try:
                    fh,fm=(int(x) for x in (fs or "00:00").split(":"))
                    th,tm=(int(x) for x in (ts or "24:00").split(":"))
                    tm2=th*60+tm; tm2=tm2 or 1440
                    return {"type":"range","from":fh*60+fm,"to":tm2}
                except:
                    messagebox.showwarning("입력 오류",f"{label_prefix}{key} 형식 오류\n예) 06:00 ~ 09:00")
                    return None
            slots=[TIME_SLOT_KEYS[i] for i,v in enumerate(self._band_vars[key]) if v.get()]
            return {"type":"slots","slots":slots or TIME_SLOT_KEYS[:]}
        dc=_band("dep_band"); ac=_band("arr_band")
        rdc=_band("ret_dep_band"); rac=_band("ret_arr_band")
        if None in (dc,ac,rdc,rac): return None
        config={"airline_mode":am,"specific_airlines":specific,
                "dep_band":dc,"arr_band":ac,"ret_dep_band":rdc,"ret_arr_band":rac}
        ep=None
        if self.use_expand_var.get():
            if not self._expand_rows:
                messagebox.showerror("입력 오류",f"{label_prefix}순위를 1개 이상 추가하세요.")
                return None
            ep=[]
            bmap={"전체":["새벽","오전","오후","야간"],"오전":["오전"],"오후":["오후"],
                  "야간":["야간"],"오후+야간":["오후","야간"]}
            for i,rd in enumerate(self._expand_rows):
                try: n=int(rd["nights"].get())-1; assert 1<=n<=14
                except:
                    messagebox.showerror("입력 오류",f"{label_prefix}{i+1}순위 일수 오류.")
                    return None
                at=rd["arr_to"].get().strip()
                rf=rd["ret_dep_from"].get().strip(); rt=rd["ret_dep_to"].get().strip()
                abc=rd["arr_band"].get()
                if at:
                    try:
                        ah,am2=(int(x) for x in at.split(":")); atm=ah*60+am2 or 1440
                        ace={"type":"range","from":0,"to":atm}
                    except:
                        messagebox.showerror("입력 오류",f"{label_prefix}{i+1}순위 도착 상한 오류"); return None
                else: ace={"type":"slots","slots":bmap.get(abc,["새벽","오전","오후","야간"])}
                if rf or rt:
                    try:
                        rfrom=0
                        if rf: rh,rm=(int(x) for x in rf.split(":")); rfrom=rh*60+rm
                        rto=24*60
                        if rt: th,tm=(int(x) for x in rt.split(":")); rto=th*60+tm or 1440
                        rce={"type":"range","from":rfrom,"to":rto}
                    except:
                        messagebox.showerror("입력 오류",f"{label_prefix}{i+1}순위 귀국출발 오류"); return None
                else: rce={"type":"slots","slots":["새벽","오전","오후","야간"]}
                ep.append({"nights":n,"arr_cond":ace,"ret_dep_cond":rce})
        return (ro, self.extra_return_var.get(), config, ep)


# ══════════════════════════════════════════════════════════
#  메인 앱 — 사이드바 레이아웃
# ══════════════════════════════════════════════════════════
class GmarketAirApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("G마켓 항공료 자동 추출기  |  투어로 김진모")
        self.geometry("1320x860")
        self.minsize(1100, 700)
        self.configure(fg_color=MAIN_BG)
        self._running=False ; self._thread=None
        self._set_mode=False ; self._set_blocks=[]
        try:
            base = os.path.dirname(sys.executable if getattr(sys,'frozen',False)
                                   else os.path.abspath(__file__)) if "__file__" in dir() else ""
            ico = os.path.join(base,"icon.ico")
            if os.path.exists(ico): self.iconbitmap(ico)
        except: pass
        self._build_ui()

    def _build_ui(self):
        # ── 루트 레이아웃: 사이드바 | 메인 ──
        root_f = ctk.CTkFrame(self, fg_color="transparent")
        root_f.pack(fill="both", expand=True)
        root_f.columnconfigure(1, weight=1)
        root_f.rowconfigure(0, weight=1)

        # ════════════════════════════════════
        #  사이드바
        # ════════════════════════════════════
        sb = ctk.CTkFrame(root_f, width=220, fg_color=SB_BG,
                          corner_radius=0)
        sb.grid(row=0, column=0, sticky="nsew")
        sb.pack_propagate(False)

        # 로고
        logo_f = ctk.CTkFrame(sb, fg_color="transparent")
        logo_f.pack(fill="x", padx=16, pady=(20,24))
        logo_box = ctk.CTkFrame(logo_f, width=36, height=36,
                                fg_color=ACCENT, corner_radius=8)
        logo_box.pack(side="left")
        logo_box.pack_propagate(False)
        ctk.CTkLabel(logo_box, text="✈", font=(FNT,20),
                     text_color="white").place(relx=.5,rely=.5,anchor="center")
        ctk.CTkLabel(logo_f, text="  항공료 추출기",
                     font=(FNT,14,"bold"), text_color=TEXT1).pack(side="left")

        Divider(sb).pack(fill="x", padx=14, pady=(0,16))

        # 메뉴
        self._nav_items = {}
        nav_defs = [
            ("search",  "🔍  검색 조건"),
            ("log",     "📋  실행 로그"),
            ("settings","⚙   설정"),
        ]
        for key, label in nav_defs:
            f = ctk.CTkFrame(sb, fg_color="transparent", cursor="hand2")
            f.pack(fill="x", padx=10, pady=2)
            lbl = ctk.CTkLabel(f, text=label, font=(FNT,13),
                               text_color=TEXT2, anchor="w")
            lbl.pack(fill="x", padx=10, pady=8)
            self._nav_items[key] = (f, lbl)
            f.bind("<Button-1>", lambda e, k=key: self._nav_to(k))
            lbl.bind("<Button-1>", lambda e, k=key: self._nav_to(k))

        # 하단 정보
        sb_btm = ctk.CTkFrame(sb, fg_color="transparent")
        sb_btm.pack(side="bottom", fill="x", padx=14, pady=16)
        Divider(sb_btm).pack(fill="x", pady=(0,12))
        Label(sb_btm, f"v{CURRENT_VERSION}", size=11, color=ACCENT).pack(anchor="w")
        Label(sb_btm, "투어로 김진모", size=11, color=TEXT3).pack(anchor="w")
        Label(sb_btm, "오류 문의: 진모에게", size=10, color=TEXT3).pack(anchor="w",pady=(4,0))

        # ════════════════════════════════════
        #  메인 영역 (탭처럼 전환)
        # ════════════════════════════════════
        main_f = ctk.CTkFrame(root_f, fg_color=MAIN_BG, corner_radius=0)
        main_f.grid(row=0, column=1, sticky="nsew")
        main_f.columnconfigure(0, weight=1)
        main_f.rowconfigure(0, weight=1)

        # 상단 타이틀 바
        title_bar = ctk.CTkFrame(main_f, fg_color=SB_BG, height=52, corner_radius=0)
        title_bar.pack(fill="x")
        title_bar.pack_propagate(False)
        self._page_title = ctk.CTkLabel(
            title_bar, text="검색 조건", font=(FNT,15,"bold"), text_color=TEXT1)
        self._page_title.pack(side="left", padx=24, pady=14)

        # 실행 버튼 (타이틀 바 우측)
        btn_area = ctk.CTkFrame(title_bar, fg_color="transparent")
        btn_area.pack(side="right", padx=16, pady=8)
        self._stop_btn = ctk.CTkButton(
            btn_area, text="■  중지", command=self._stop,
            width=90, height=34, fg_color="#7f1d1d", hover_color="#991b1b",
            font=(FNT,12,"bold"), state="disabled")
        self._stop_btn.pack(side="right", padx=(6,0))
        self._run_btn = ctk.CTkButton(
            btn_area, text="▶  추출 시작", command=self._start,
            width=120, height=34, fg_color="#14532d", hover_color="#166534",
            font=(FNT,12,"bold"))
        self._run_btn.pack(side="right")

        # ── 페이지: 검색 조건 ──
        self._page_search = ctk.CTkScrollableFrame(
            main_f, fg_color=MAIN_BG, corner_radius=0,
            scrollbar_button_color=CARD2_BG,
            scrollbar_button_hover_color=BORDER)
        self._page_search.columnconfigure(0, weight=1)

        self._build_search_page()

        # ── 페이지: 실행 로그 ──
        self._page_log = ctk.CTkFrame(main_f, fg_color=MAIN_BG, corner_radius=0)
        self._build_log_page()

        # ── 페이지: 설정 ──
        self._page_settings = ctk.CTkFrame(main_f, fg_color=MAIN_BG, corner_radius=0)
        self._build_settings_page()

        self._nav_to("search")

    # ── 네비게이션 ──────────────────────────────────────────
    def _nav_to(self, key):
        titles = {"search":"검색 조건","log":"실행 로그","settings":"설정"}
        pages  = {"search":self._page_search,"log":self._page_log,
                  "settings":self._page_settings}
        for k,(f,lbl) in self._nav_items.items():
            if k==key:
                f.configure(fg_color=CARD2_BG)
                lbl.configure(text_color=TEXT1, font=(FNT,12,"bold"))
            else:
                f.configure(fg_color="transparent")
                lbl.configure(text_color=TEXT2, font=(FNT,13))
        for k,p in pages.items():
            if k==key: p.pack(fill="both", expand=True)
            else: p.pack_forget()
        self._page_title.configure(text=titles[key])

    # ── 검색 조건 페이지 ────────────────────────────────────
    def _build_search_page(self):
        p = self._page_search
        now = datetime.date.today()
        dy = now.year; dm = now.month+1 if now.month<12 else 1

        # 노선 카드
        c1 = Card(p); c1.pack(fill="x", padx=20, pady=(16,0))
        Label(c1, "노선", size=12, color=TEXT2).pack(anchor="w", padx=16, pady=(12,8))
        r1 = ctk.CTkFrame(c1, fg_color="transparent")
        r1.pack(fill="x", padx=16, pady=(0,14))

        Label(r1, "출발지", size=12, color=TEXT2).pack(side="left")
        self.origin_country_var = tk.StringVar(value="국내")
        self._occ = Combo(r1, self.origin_country_var, list(AIRPORTS.keys()), w=100,
                          cmd=lambda v: self._on_country_change("origin"))
        self._occ.pack(side="left", padx=(8,4))
        self.origin_city_var = tk.StringVar(value="인천")
        self._occ2 = Combo(r1, self.origin_city_var, list(AIRPORTS["국내"].keys()), w=120)
        self._occ2.pack(side="left", padx=(0,20))

        Label(r1, "→", size=15, color=TEXT3).pack(side="left", padx=(0,20))

        Label(r1, "목적지", size=12, color=TEXT2).pack(side="left")
        self.dest_country_var = tk.StringVar(value="일본")
        self._dcc = Combo(r1, self.dest_country_var, list(AIRPORTS.keys()), w=100,
                          cmd=lambda v: self._on_country_change("dest"))
        self._dcc.pack(side="left", padx=(8,4))
        self.dest_city_var = tk.StringVar(value="삿포로")
        self._dcc2 = Combo(r1, self.dest_city_var, list(AIRPORTS["일본"].keys()), w=120)
        self._dcc2.pack(side="left")

        # 출발 기간 카드
        c2 = Card(p); c2.pack(fill="x", padx=20, pady=(10,0))
        Label(c2, "출발 기간", size=12, color=TEXT2).pack(anchor="w", padx=16, pady=(12,8))
        r2 = ctk.CTkFrame(c2, fg_color="transparent")
        r2.pack(fill="x", padx=16, pady=(0,14))
        self.date_from_var = tk.StringVar(value=f"{dy}-{dm:02d}-01")
        Entry(r2, self.date_from_var, w=120).pack(side="left")
        Label(r2, "~", size=13, color=TEXT3).pack(side="left", padx=8)
        self.date_to_var = tk.StringVar(value=f"{dy}-{dm:02d}-28")
        Entry(r2, self.date_to_var, w=120).pack(side="left")
        Label(r2, "(형식: 2026-07-10)", size=10, color=TEXT3).pack(side="left", padx=(12,0))
        self.date_from_var.trace_add("write", lambda *a: self._refresh_all_hints())

        # 검색 조건
        cond_hdr = ctk.CTkFrame(p, fg_color="transparent")
        cond_hdr.pack(fill="x", padx=20, pady=(14,4))
        Label(cond_hdr, "검색 조건", size=13, bold=True, color=TEXT1).pack(side="left")

        # 단일 모드 컨테이너
        self._single_container = ctk.CTkFrame(p, fg_color="transparent")
        self._single_block = ConditionBlock(self._single_container, default_nights="3")
        self._single_block.bind_date_from(lambda: self.date_from_var.get())

        # 세트 모드 토글 바
        self._toggle_bar = ctk.CTkFrame(p, fg_color="transparent")
        self._toggle_bar.pack(fill="x", padx=20, pady=(0,8))
        Label(self._toggle_bar,
              "여러 조건을 동시에 추출하고 싶다면 검색 조건 세트를 활성화하세요",
              size=11, color=TEXT3).pack(side="left")
        self._toggle_btn = ctk.CTkButton(
            self._toggle_bar, text="📑  검색 조건 세트 활성화",
            command=self._toggle_set_mode, width=190, height=28,
            fg_color="transparent", border_width=1, border_color=ACCENT,
            text_color=ACCENT, hover_color="#0d2137", font=(FNT,12))
        self._toggle_btn.pack(side="right")

        # 세트 컨테이너
        self._sets_container = ctk.CTkFrame(p, fg_color="transparent")
        sets_hdr = ctk.CTkFrame(self._sets_container, fg_color="transparent")
        sets_hdr.pack(fill="x", pady=(0,6))
        Label(sets_hdr, "세트별 결과는 시트로 분리되어 같은 엑셀에 저장됩니다.",
              size=10, color=TEXT3).pack(side="left")
        ctk.CTkButton(sets_hdr, text="＋ 세트 추가", command=self._add_condition_set,
                      width=90, height=26, fg_color="transparent",
                      border_width=1, border_color=ACCENT, text_color=ACCENT,
                      hover_color="#0d2137", font=(FNT,12)).pack(side="right")
        self._sets_list_frame = ctk.CTkFrame(self._sets_container, fg_color="transparent")
        self._sets_list_frame.pack(fill="x")

        self._single_container.pack(fill="x", padx=20, before=self._toggle_bar)
        self._sets_container.pack_forget()

        # 진행률
        prog_f = ctk.CTkFrame(p, fg_color="transparent")
        prog_f.pack(fill="x", padx=20, pady=(8,0))
        self._prog_label = Label(prog_f, "대기 중...", size=11, color=TEXT3)
        self._prog_label.pack(side="left")
        self._progressbar = ctk.CTkProgressBar(prog_f, height=6,
                                               fg_color=CARD2_BG, progress_color=ACCENT)
        self._progressbar.pack(side="left", fill="x", expand=True, padx=10)
        self._progressbar.set(0)
        self._prog_pct = Label(prog_f, "", size=11, color=INFO)
        self._prog_pct.pack(side="left")

        ctk.CTkFrame(p, height=20, fg_color="transparent").pack()

    # ── 로그 페이지 ─────────────────────────────────────────
    def _build_log_page(self):
        p = self._page_log
        top = ctk.CTkFrame(p, fg_color="transparent")
        top.pack(fill="x", padx=20, pady=(14,8))
        Label(top, "실행 로그", size=13, bold=True, color=TEXT1).pack(side="left")
        SmallBtn(top, "로그 지우기", self._clear_log,
                 w=90, h=28).pack(side="right")

        log_card = ctk.CTkFrame(p, fg_color="#0a0c10", corner_radius=10,
                                border_width=1, border_color=BORDER)
        log_card.pack(fill="both", expand=True, padx=20, pady=(0,20))
        self._log = tk.Text(log_card, bg="#0a0c10", fg=TEXT1,
                            font=("Consolas",9), height=14, wrap="word",
                            state="disabled", relief="flat",
                            padx=14, pady=10, insertbackground=TEXT1,
                            highlightthickness=0, selectbackground=ACCENT)
        sb2 = ttk.Scrollbar(log_card, command=self._log.yview)
        self._log.configure(yscrollcommand=sb2.set)
        sb2.pack(side="right", fill="y")
        self._log.pack(side="left", fill="both", expand=True)
        self._log.tag_configure("ok",   foreground=OK)
        self._log.tag_configure("err",  foreground=ERR)
        self._log.tag_configure("info", foreground=INFO)
        self._log.tag_configure("warn", foreground=WARN)
        self._log_msg("시스템 준비 완료. 조건을 설정하고 '추출 시작'을 눌러주세요.", "info")
        self._log_msg("오류 등 문의 사항은 진모에게 문의 주세요.", "info")
        if not SCRAPER_OK:
            self._log_msg("⚠ scraper_core.py를 찾을 수 없습니다.", "err")

    # ── 설정 페이지 ─────────────────────────────────────────
    def _build_settings_page(self):
        p = self._page_settings
        Label(p, "설정", size=13, bold=True, color=TEXT1).pack(anchor="w", padx=20, pady=(20,12))

        c = Card(p); c.pack(fill="x", padx=20)
        Label(c, "저장 위치", size=12, color=TEXT2).pack(anchor="w", padx=16, pady=(12,6))
        r = ctk.CTkFrame(c, fg_color="transparent")
        r.pack(fill="x", padx=16, pady=(0,14))
        self.out_dir_var = tk.StringVar(value=get_default_output_dir())
        Entry(r, self.out_dir_var, w=400).pack(side="left", fill="x", expand=True, padx=(0,10))
        ctk.CTkButton(r, text="폴더 선택", command=self._choose_dir,
                      width=100, height=32, fg_color=ACCENT, hover_color=ACCENT_H,
                      font=(FNT,13)).pack(side="left")

        c_fname = Card(p); c_fname.pack(fill="x", padx=20, pady=(10,0))
        Label(c_fname, "엑셀 파일명", size=12, color=TEXT2).pack(anchor="w", padx=16, pady=(12,4))
        Label(c_fname, "비워두면 자동 생성됩니다  (예: gmarket_ICN_CTS_20260801_20260828.xlsx)",
              size=10, color=TEXT3).pack(anchor="w", padx=16, pady=(0,6))
        fn_row = ctk.CTkFrame(c_fname, fg_color="transparent")
        fn_row.pack(fill="x", padx=16, pady=(0,14))
        self.custom_fname_var = tk.StringVar(value="")
        Entry(fn_row, self.custom_fname_var, w=400).pack(side="left", fill="x", expand=True, padx=(0,10))
        Label(fn_row, ".xlsx", size=12, color=TEXT3).pack(side="left")

        c2 = Card(p); c2.pack(fill="x", padx=20, pady=(10,0))
        Label(c2, "기타 옵션", size=12, color=TEXT2).pack(anchor="w", padx=16, pady=(12,6))
        self.show_browser_var = tk.BooleanVar(value=False)
        Check(c2, self.show_browser_var, "브라우저 창 표시 (문제 확인용)"
              ).pack(anchor="w", padx=16, pady=(0,14))

    # ── 세트 모드 ───────────────────────────────────────────
    def _toggle_set_mode(self):
        self._set_mode = not self._set_mode
        if self._set_mode:
            self._single_container.pack_forget()
            self._sets_container.pack(fill="x", padx=20, before=self._toggle_bar)
            self._toggle_label_upd("단일 모드로 전환","📑  단일 모드로 전환")
            if not self._set_blocks: self._add_condition_set()
        else:
            self._sets_container.pack_forget()
            self._single_container.pack(fill="x", padx=20, before=self._toggle_bar)
            self._toggle_label_upd(
                "여러 조건을 동시에 추출하고 싶다면 검색 조건 세트를 활성화하세요",
                "📑  검색 조건 세트 활성화")

    def _toggle_label_upd(self, hint, btn_text):
        for w in self._toggle_bar.winfo_children():
            if isinstance(w, ctk.CTkLabel): w.configure(text=hint)
            if isinstance(w, ctk.CTkButton): w.configure(text=btn_text)

    def _add_condition_set(self):
        idx = len(self._set_blocks)+1
        def _make_delete(block_ref):
            def _del():
                block_ref.outer.destroy()
                self._set_blocks.remove(block_ref)
                for i,b in enumerate(self._set_blocks,1):
                    for w in b.outer.winfo_children():
                        for c in w.winfo_children():
                            if isinstance(c,ctk.CTkLabel) and "세트" in (c.cget("text") or ""):
                                c.configure(text=f"세트 {i}"); break
            return _del
        block = ConditionBlock(self._sets_list_frame, show_set_header=True,
                               set_index=idx, on_delete=lambda: None,
                               default_nights=str(2+idx))
        block.on_delete_fn = _make_delete(block)
        self._rebind_delete(block)
        block.bind_date_from(lambda: self.date_from_var.get())
        self._set_blocks.append(block)

    def _rebind_delete(self, block):
        def _find(w):
            for c in w.winfo_children():
                if isinstance(c,ctk.CTkButton) and "삭제" in (c.cget("text") or ""):
                    c.configure(command=block.on_delete_fn); return
                _find(c)
        _find(block.outer)

    def _refresh_all_hints(self):
        df = self.date_from_var.get()
        self._single_block.update_return_hint(df)
        for b in self._set_blocks: b.update_return_hint(df)

    def _on_country_change(self, which):
        if which=="origin":
            c=self.origin_country_var.get(); cities=list(AIRPORTS[c].keys())
            self._occ2.configure(values=cities); self.origin_city_var.set(cities[0]); self._occ2.set(cities[0])
        else:
            c=self.dest_country_var.get(); cities=list(AIRPORTS[c].keys())
            self._dcc2.configure(values=cities); self.dest_city_var.set(cities[0]); self._dcc2.set(cities[0])

    def _choose_dir(self):
        d=filedialog.askdirectory(initialdir=self.out_dir_var.get())
        if d: self.out_dir_var.set(d)

    def _clear_log(self):
        self._log.configure(state="normal"); self._log.delete("1.0","end")
        self._log.configure(state="disabled")

    def _log_msg(self, msg, tag=""):
        def _do():
            self._log.configure(state="normal")
            ts=datetime.datetime.now().strftime("%H:%M:%S")
            self._log.insert("end",f"[{ts}] {msg}\n",tag or "")
            self._log.see("end"); self._log.configure(state="disabled")
        self.after(0,_do)

    def _update_progress(self, cur, total):
        def _do():
            pct=cur/total if total else 0
            self._progressbar.set(pct)
            self._prog_label.configure(text=f"{cur} / {total} 일")
            self._prog_pct.configure(text=f"{int(pct*100)}%")
        self.after(0,_do)

    def _validate(self):
        import datetime as _dt
        try:
            df=_dt.datetime.strptime(self.date_from_var.get().strip(),"%Y-%m-%d").date()
            dt=_dt.datetime.strptime(self.date_to_var.get().strip(),"%Y-%m-%d").date()
            assert df<=dt
        except:
            messagebox.showerror("입력 오류","출발일 형식(2026-07-10), 시작일≤종료일을 확인하세요."); return None
        oc=self.origin_country_var.get(); ocity=self.origin_city_var.get()
        dc=self.dest_country_var.get(); dcity=self.dest_city_var.get()
        if oc==dc and ocity==dcity:
            messagebox.showerror("입력 오류","출발지와 목적지가 같습니다."); return None
        origin=AIRPORTS[oc][ocity]; dest=AIRPORTS[dc][dcity]
        sets=[]
        if self._set_mode:
            if not self._set_blocks:
                messagebox.showerror("입력 오류","세트를 1개 이상 추가하세요."); return None
            for i,block in enumerate(self._set_blocks,1):
                res=block.validate_and_get(label_prefix=f"[세트{i}] ")
                if res is None: return None
                ro,er,config,ep=res
                sets.append({"label":f"세트{i}_{ro+1}일_{config['airline_mode']}",
                             "return_offset":ro,"extra_return":er,"config":config,"expand_priorities":ep})
        else:
            res=self._single_block.validate_and_get()
            if res is None: return None
            ro,er,config,ep=res
            sets.append({"label":"결과","return_offset":ro,"extra_return":er,
                         "config":config,"expand_priorities":ep})
        return {"origin":origin,"dest":dest,"date_from":df,"date_to":dt,
                "sets":sets,"multi_mode":self._set_mode,
                "show_browser":self.show_browser_var.get(),
                "out_dir":self.out_dir_var.get(),
                "custom_fname":self.custom_fname_var.get().strip()}

    def _start(self):
        if not SCRAPER_OK:
            messagebox.showerror("오류","scraper_core.py가 없어 실행할 수 없습니다."); return
        params=self._validate()
        if params is None: return
        self._running=True
        self._run_btn.configure(state="disabled")
        self._stop_btn.configure(state="normal")
        self._progressbar.set(0)
        self._prog_label.configure(text="실행 중...")
        self._nav_to("log")
        self._thread=threading.Thread(target=self._run_scrape,args=(params,),daemon=True)
        self._thread.start()

    def _stop(self):
        self._running=False
        self._log_msg("⏹ 중지 요청됨. 현재 날짜 완료 후 종료됩니다.","err")

    def _run_scrape(self, p):
        from datetime import timedelta
        from scraper_core import (init_driver, build_url, fetch_flights, select_best,
                                  parse_price, calc_per_person, reset_debug,
                                  save_excel, save_excel_multi)
        reset_debug()
        origin=p["origin"]; dest=p["dest"]
        date_from=p["date_from"]; date_to=p["date_to"]
        sets=p["sets"]; multi_mode=p["multi_mode"]
        show=p["show_browser"]; out_dir=p["out_dir"]

        date_list=[]
        d=date_from
        while d<=date_to: date_list.append(d); d+=timedelta(days=1)
        total_days=len(date_list); total_work=total_days*len(sets)
        custom_fname = p.get("custom_fname", "").strip()
        if custom_fname:
            fname = custom_fname if custom_fname.endswith(".xlsx") else custom_fname + ".xlsx"
        else:
            fname = f"gmarket_{origin}_{dest}_{date_from.strftime('%Y%m%d')}_{date_to.strftime('%Y%m%d')}.xlsx"
        out_path=os.path.join(out_dir,fname)

        self._log_msg(f"🛫 추출 시작: {origin} → {dest}  |  {date_from} ~ {date_to}  |  세트 {len(sets)}개  |  총 {total_days}일","info")
        try:
            driver=init_driver(show=show); rows_by_set=[]; work_done=0
            for s in sets:
                config=s["config"]; ep=s["expand_priorities"]
                base_offset=s["return_offset"]; er=s["extra_return"]
                offsets=[base_offset,base_offset+1] if er else [base_offset]
                label=s["label"]; pfx=f"[{label}] " if multi_mode else ""
                if multi_mode: self._log_msg(f"\n▶ {label} 검색 시작","info")
                rows=[]
                for idx,dep_date in enumerate(date_list,1):
                    if not self._running: self._log_msg("⏹ 중지됨.","err"); break
                    self._log_msg(f"  {pfx}{dep_date.strftime('%Y-%m-%d')} ({dep_date.strftime('%a')}) 검색 중...")
                    best=None; used_offset=base_offset; used_priority=None
                    specific=config.get("specific_airlines") if config.get("airline_mode")=="특정항공사" else None
                    if ep:
                        for pi,pri in enumerate(ep):
                            n=pri["nights"]
                            url=build_url(origin,dest,dep_date.strftime("%Y%m%d"),(dep_date+timedelta(days=n)).strftime("%Y%m%d"))
                            flights=fetch_flights(driver,url,self._log_msg,specific_airlines=specific)
                            if not flights: continue
                            tc=dict(config); tc["arr_band"]=pri["arr_cond"]; tc["ret_dep_band"]=pri["ret_dep_cond"]
                            tc["ret_arr_band"]={"type":"slots","slots":["새벽","오전","오후","야간"]}
                            cand=select_best(flights,tc)
                            if cand: best=cand; used_offset=n; used_priority=pi+1; break
                        arr_date=dep_date+timedelta(days=used_offset)
                    else:
                        best_arr=dep_date+timedelta(days=base_offset); best_total=None
                        for off in offsets:
                            arr_try=dep_date+timedelta(days=off)
                            url=build_url(origin,dest,dep_date.strftime("%Y%m%d"),arr_try.strftime("%Y%m%d"))
                            flights=fetch_flights(driver,url,self._log_msg,specific_airlines=specific)
                            cand=select_best(flights,config) if flights else None
                            if cand:
                                ct=parse_price(cand.get("cardPrice","0"))
                                if best is None or (ct and ct<best_total): best=cand; best_total=ct; best_arr=arr_try
                        arr_date=best_arr
                    if best:
                        ap=best.get("cardPrice",""); best["price"]=ap
                        t4=parse_price(ap); p1=calc_per_person(t4)
                        pn=f" [{used_priority}순위]" if used_priority else ""
                        rows.append({"dep_date":dep_date.strftime("%Y-%m-%d"),"arr_date":arr_date.strftime("%Y-%m-%d"),
                                     "airline":best["airline"],"dep":best["dep"],"arr":best["arr"],
                                     "rDep":best.get("rDep",""),"rArr":best.get("rArr",""),
                                     "total4":t4,"per1":p1,"seller":pn.strip(),"found":True})
                        self._log_msg(f"    ✔{pn} {pfx}{best['airline']}  {best['dep']}→{best['arr']}  4인:{t4:,}원  1인:{p1:,}원","ok")
                    else:
                        rows.append({"dep_date":dep_date.strftime("%Y-%m-%d"),
                                     "arr_date":(dep_date+timedelta(days=base_offset)).strftime("%Y-%m-%d"),
                                     "airline":"","dep":"","arr":"","rDep":"","rArr":"",
                                     "total4":0,"per1":0,"seller":"","found":False})
                        self._log_msg(f"    ✗ {pfx}조건에 맞는 항공편 없음")
                    work_done+=1; self._update_progress(work_done,total_work)
                rows_by_set.append({"label":label,"rows":rows})
                if not self._running: break
            driver.quit()
            tr=sum(len(e["rows"]) for e in rows_by_set)
            if tr>0:
                if multi_mode: save_excel_multi(rows_by_set,origin,dest,out_path)
                else: save_excel(rows_by_set[0]["rows"],origin,dest,date_from.year,date_from.month,out_path)
                sm=" / ".join(f"{e['label']}: {sum(1 for r in e['rows'] if r['found'])}/{len(e['rows'])}일" for e in rows_by_set)
                self._log_msg(f"\n✅ 완료! {sm} → {fname}","ok")
                self.after(0,lambda: messagebox.showinfo("완료",f"추출 완료!\n{date_from} ~ {date_to}\n{sm}\n저장: {out_path}"))
                try:
                    from plyer import notification
                    notification.notify(title="✈ G마켓 항공료 추출 완료",message=f"{origin} → {dest} 완료",app_name="G마켓 항공료 추출기",timeout=8)
                except: pass
            else: self._log_msg("⚠ 수집된 데이터 없음","err")
        except Exception as e:
            self._log_msg(f"❌ 오류 발생: {e}","err")
            self.after(0,lambda: messagebox.showerror("오류",str(e)))
        finally:
            self._running=False
            self.after(0,lambda: self._run_btn.configure(state="normal"))
            self.after(0,lambda: self._stop_btn.configure(state="disabled"))
            self.after(0,lambda: self._prog_label.configure(text="완료"))


if __name__=="__main__":
    app=GmarketAirApp()
    app.mainloop()
