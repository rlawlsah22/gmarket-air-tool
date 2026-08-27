"""
gmarket_air_gui.py
항공료 자동 추출기 - pywebview 기반 GUI
"""

import os
import re
import sys
import json
import ctypes
import threading
import datetime
import urllib.request
import webview

# ─────────────────────────────────────────────
#  자동 업데이트 (GitHub)
# ─────────────────────────────────────────────
CURRENT_VERSION = "3.0"
GITHUB_USER     = "rlawlsah22"
GITHUB_REPO     = "gmarket-air-tool"
RAW_BASE        = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/main"
UPDATE_FILES    = ["gmarket_air_gui.py", "scraper_core.py", "airpremia_fare_checker.py",
                   "web/index.html", "web/style.css", "web/app.js"]


def _hide_file_windows(path):
    try:
        if os.name == "nt":
            ctypes.windll.kernel32.SetFileAttributesW(str(path), 0x02)
    except Exception:
        pass


def _unhide_file_windows(path):
    try:
        if os.name == "nt" and os.path.exists(path):
            ctypes.windll.kernel32.SetFileAttributesW(str(path), 0x80)
    except Exception:
        pass


def check_and_update():
    try:
        url = f"{RAW_BASE}/version.txt"
        with urllib.request.urlopen(url, timeout=5) as r:
            latest = r.read().decode().strip()
        if latest <= CURRENT_VERSION:
            return
        base_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))
        for fname in UPDATE_FILES:
            file_url = f"{RAW_BASE}/{fname}"
            dst = os.path.join(base_dir, fname)
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            _unhide_file_windows(dst)
            with urllib.request.urlopen(file_url, timeout=10) as r:
                with open(dst, "wb") as f:
                    f.write(r.read())
            if not fname.startswith("web/"):
                _hide_file_windows(dst)
        ctypes.windll.user32.MessageBoxW(
            0, f"새 버전({latest})으로 업데이트되었습니다.\n프로그램을 닫고 다시 실행해주세요.",
            "업데이트 완료", 0x40
        )
        sys.exit()
    except SystemExit:
        raise
    except Exception:
        pass


# ─────────────────────────────────────────────
#  프리셋 저장 (로컬 JSON)
# ─────────────────────────────────────────────
def _preset_file_path(kind="gmarket"):
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    fname = "presets.json" if kind == "gmarket" else f"presets_{kind}.json"
    return os.path.join(base_dir, fname)


def load_presets(kind="gmarket"):
    path = _preset_file_path(kind)
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_presets(kind, data):
    path = _preset_file_path(kind)
    _unhide_file_windows(path)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    _hide_file_windows(path)


def get_default_output_dir():
    return os.path.join(os.path.expanduser("~"), "Desktop")


# ─────────────────────────────────────────────
#  백엔드 모듈 로드
# ─────────────────────────────────────────────
try:
    from scraper_core import (
        AIRPORTS, init_driver, build_url, fetch_flights, select_best,
        parse_price, calc_per_person, reset_debug, save_excel, save_excel_multi,
    )
    SCRAPER_OK = True
except ImportError:
    SCRAPER_OK = False
    AIRPORTS = {
        "국내": {"인천": "ICN", "부산": "PUS", "김포": "GMP", "청주": "CJJ", "대구": "TAE"},
        "일본": {"삿포로": "CTS", "도쿄(나리타)": "NRT", "오사카": "OSA", "후쿠오카": "FUK", "오키나와": "OKA"},
    }

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
    "방콕 (BKK)": "BKK",
}

AIRLINE_MODES = [
    "LCC 우선 → FSC 대체",
    "LCC만",
    "FSC만 (아시아나/대한항공)",
    "외항사만",
    "특정 항공사 지정",
]

TIME_SLOT_LABELS = ["새벽 (00~06시)", "오전 (06~12시)", "오후 (12~18시)", "야간 (18~24시)"]
TIME_SLOT_KEYS = ["새벽", "오전", "오후", "야간"]

SPECIFIC_AIRLINES = [
    "대한항공", "아시아나항공", "에어프레미아", "진에어", "제주항공",
    "티웨이항공", "이스타항공", "에어부산", "에어서울", "에어로케이",
    "파라타항공", "타이항공", "필리핀항공", "베트남항공",
    "중국국제항공", "중국남방항공", "중국동방항공", "산동항공",
]

_AIRLINE_MODE_MAP = {
    "LCC 우선 → FSC 대체": "LCC우선_FSC대체",
    "LCC만": "LCC만",
    "FSC만 (아시아나/대한항공)": "FSC만",
    "외항사만": "외항사만",
    "특정 항공사 지정": "특정항공사",
}


def compute_actual_arrival_date(ret_dep_date, rDep: str, rDuration):
    if not rDep or not rDuration:
        return ret_dep_date
    try:
        h, m = rDep.strip().split(":")
        dep_minutes = int(h) * 60 + int(m)
        arrival_minutes = dep_minutes + int(rDuration)
        extra_days = arrival_minutes // (24 * 60)
        return ret_dep_date + datetime.timedelta(days=extra_days)
    except Exception:
        return ret_dep_date


def _band_from_js(cond):
    """JS에서 온 {type:'range'|'slots', ...} 형식을 백엔드 config 형식으로 변환"""
    if not cond:
        return {"type": "slots", "slots": TIME_SLOT_KEYS[:]}
    if cond.get("type") == "range":
        try:
            fh, fm = (int(x) for x in cond["from"].split(":"))
            th, tm = (int(x) for x in cond["to"].split(":"))
            to_min = th * 60 + tm
            if to_min == 0:
                to_min = 1440
            return {"type": "range", "from": fh * 60 + fm, "to": to_min}
        except Exception:
            return {"type": "slots", "slots": TIME_SLOT_KEYS[:]}
    return {"type": "slots", "slots": cond.get("slots") or TIME_SLOT_KEYS[:]}


class Api:
    def __init__(self):
        self._running = {"gmarket": False, "airpremia": False}
        self._window = None

    def set_window(self, window):
        self._window = window

    # ── 초기 메타데이터 ──
    def get_meta(self):
        return {
            "airports": AIRPORTS,
            "airline_modes": AIRLINE_MODES,
            "specific_airlines": SPECIFIC_AIRLINES,
            "airpremia_destinations": AIRPREMIA_DESTINATIONS,
            "time_slot_labels": TIME_SLOT_LABELS,
            "time_slot_keys": TIME_SLOT_KEYS,
            "default_out_dir": get_default_output_dir(),
            "presets_gmarket": load_presets("gmarket"),
            "presets_airpremia": load_presets("airpremia"),
        }

    def choose_dir(self):
        result = self._window.create_file_dialog(webview.FOLDER_DIALOG)
        if result:
            return result[0]
        return None

    def save_presets(self, kind, data):
        save_presets(kind, data)
        return True

    # ── 로그/진행률 전송 (kind별로 JS에 전달) ──
    def _log(self, kind, text, log_kind=""):
        if self._window:
            safe = json.dumps(text)
            self._window.evaluate_js(f"window.onBackendLog({json.dumps(kind)}, {safe}, {json.dumps(log_kind)})")

    def _progress(self, kind, cur, total):
        if self._window:
            self._window.evaluate_js(f"window.onBackendProgress({json.dumps(kind)}, {cur}, {total})")

    def _done(self, kind):
        if self._window:
            self._window.evaluate_js(f"window.onBackendDone({json.dumps(kind)})")

    def stop_scraping_gmarket(self):
        self._running["gmarket"] = False
        return True

    def stop_scraping_airpremia(self):
        self._running["airpremia"] = False
        return True

    # ── 실행 진입점 (JS에서 호출) ──
    def start_scraping_gmarket(self, config):
        if self._running["gmarket"]:
            return {"error": "이미 실행 중입니다."}
        if not SCRAPER_OK:
            return {"error": "scraper_core.py가 없어 실행할 수 없습니다."}
        try:
            date_from = datetime.datetime.strptime(config["date_from"].strip(), "%Y-%m-%d").date()
            date_to = datetime.datetime.strptime(config["date_to"].strip(), "%Y-%m-%d").date()
            if date_from > date_to:
                return {"error": "출발일 시작일이 종료일보다 늦습니다."}
        except Exception:
            return {"error": "출발일 형식이 올바르지 않습니다. (예: 2026-07-10)"}

        self._running["gmarket"] = True
        thread = threading.Thread(target=self._run_gmarket, args=(config, date_from, date_to), daemon=True)
        thread.start()
        return {"ok": True}

    def start_scraping_airpremia(self, config):
        if self._running["airpremia"]:
            return {"error": "이미 실행 중입니다."}
        if not AIRPREMIA_OK:
            return {"error": "airpremia_fare_checker.py가 없어 실행할 수 없습니다."}
        try:
            date_from = datetime.datetime.strptime(config["date_from"].strip(), "%Y-%m-%d").date()
            date_to = datetime.datetime.strptime(config["date_to"].strip(), "%Y-%m-%d").date()
            if date_from > date_to:
                return {"error": "출발일 시작일이 종료일보다 늦습니다."}
        except Exception:
            return {"error": "출발일 형식이 올바르지 않습니다. (예: 2026-07-10)"}

        self._running["airpremia"] = True
        thread = threading.Thread(target=self._run_airpremia, args=(config, date_from, date_to), daemon=True)
        thread.start()
        return {"ok": True}

    # ── G마켓 실행 ──
    def _run_gmarket(self, config, date_from, date_to):
        try:
            origin = AIRPORTS[config["origin_country"]][config["origin_city"]]
            dest = AIRPORTS[config["dest_country"]][config["dest_city"]]
            adults = int(config.get("adults", 4))
            out_dir = config.get("out_dir") or get_default_output_dir()
            custom_filename = config.get("custom_filename", "")
            multi_mode = bool(config.get("set_mode"))

            raw_sets = config.get("sets", [])
            sets = []
            for i, s in enumerate(raw_sets, 1):
                mode_disp = s.get("airline_mode", AIRLINE_MODES[0])
                airline_mode = _AIRLINE_MODE_MAP.get(mode_disp, "LCC우선_FSC대체")
                bands = s.get("bands", {})
                cfg = {
                    "airline_mode": airline_mode,
                    "specific_airlines": s.get("specific_airlines", []),
                    "dep_band": _band_from_js(bands.get("dep_band")),
                    "arr_band": _band_from_js(bands.get("arr_band")),
                    "ret_dep_band": _band_from_js(bands.get("ret_dep_band")),
                    "ret_arr_band": _band_from_js(bands.get("ret_arr_band")),
                }
                nights = int(s.get("nights", 3))
                label = f"세트{i}_{nights}일_{mode_disp}" if multi_mode else "결과"
                sets.append({
                    "label": label,
                    "return_offset": nights - 1,
                    "config": cfg,
                    "expand_priorities": None,
                })

            date_list = []
            d = date_from
            while d <= date_to:
                date_list.append(d)
                d += datetime.timedelta(days=1)
            total_work = len(date_list) * len(sets)

            if custom_filename:
                safe_name = re.sub(r'[\\/:*?"<>|]', "_", custom_filename)
                if not safe_name.lower().endswith(".xlsx"):
                    safe_name += ".xlsx"
                fname = safe_name
            else:
                fname = f"gmarket_{origin}_{dest}_{date_from.strftime('%Y%m%d')}_{date_to.strftime('%Y%m%d')}.xlsx"
            out_path = os.path.join(out_dir, fname)

            reset_debug()
            self._log(
                "gmarket",
                f"🛫 추출 시작: {origin} → {dest}  |  {date_from} ~ {date_to}  |  세트 {len(sets)}개  |  총 {len(date_list)}일",
                "info",
            )

            driver = init_driver(show=config.get("show_browser", False))
            rows_by_set = []
            work_done = 0

            for s in sets:
                cfg = s["config"]
                base_offset = s["return_offset"]
                label = s["label"]
                set_prefix = f"[{label}] " if multi_mode else ""

                if multi_mode:
                    self._log("gmarket", f"\n▶ {label} 검색 시작 (여행 {base_offset+1}일, 한국 도착일 기준)", "info")

                rows = []
                for dep_date in date_list:
                    if not self._running["gmarket"]:
                        self._log("gmarket", "⏹ 중지됨.", "err")
                        break

                    self._log("gmarket", f"  {set_prefix}{dep_date.strftime('%Y-%m-%d')} ({dep_date.strftime('%a')}) 검색 중...")

                    target_arrival_date = dep_date + datetime.timedelta(days=base_offset)
                    search_candidates = [target_arrival_date, target_arrival_date - datetime.timedelta(days=1)]

                    best = None
                    best_total = None
                    specific_airlines = cfg.get("specific_airlines", []) if cfg.get("airline_mode") == "특정항공사" else None
                    for search_date in search_candidates:
                        url = build_url(origin, dest, dep_date.strftime("%Y%m%d"), search_date.strftime("%Y%m%d"), adults=adults)
                        flights = fetch_flights(driver, url, lambda t, k="": self._log("gmarket", t, k), specific_airlines=specific_airlines)
                        if not flights:
                            continue
                        valid_flights = [
                            f for f in flights
                            if compute_actual_arrival_date(search_date, f.get("rDep", ""), f.get("rDuration")) == target_arrival_date
                        ]
                        if not valid_flights:
                            continue
                        cand = select_best(valid_flights, cfg)
                        if cand:
                            cand_total = parse_price(cand.get("cardPrice", cand.get("price", "0")))
                            if best is None or (cand_total and cand_total < best_total):
                                best = cand
                                best_total = cand_total
                    arr_date = target_arrival_date

                    if best:
                        actual_price = best.get("cardPrice", "")
                        best["price"] = actual_price
                        total4 = parse_price(actual_price)
                        per1 = calc_per_person(total4, adults=adults)
                        r_airline = best.get("rAirline", best["airline"])
                        airline_display = best["airline"] if r_airline == best["airline"] else f"{best['airline']}/{r_airline}"
                        rows.append({
                            "dep_date": dep_date.strftime("%Y-%m-%d"),
                            "arr_date": arr_date.strftime("%Y-%m-%d"),
                            "airline": airline_display,
                            "dep": best["dep"], "arr": best["arr"],
                            "rDep": best.get("rDep", ""), "rArr": best.get("rArr", ""),
                            "total4": total4, "per1": per1,
                            "seller": "", "found": True,
                        })
                        self._log(
                            "gmarket",
                            f"    ✔ {set_prefix}{airline_display}  {best['dep']}→{best['arr']}  {adults}인:{total4:,}원  1인:{per1:,}원",
                            "ok",
                        )
                    else:
                        rows.append({
                            "dep_date": dep_date.strftime("%Y-%m-%d"),
                            "arr_date": arr_date.strftime("%Y-%m-%d"),
                            "airline": "", "dep": "", "arr": "", "rDep": "", "rArr": "",
                            "total4": 0, "per1": 0, "seller": "", "found": False,
                        })
                        self._log("gmarket", f"    ✗ {set_prefix}조건에 맞는 항공편 없음")

                    work_done += 1
                    self._progress("gmarket", work_done, total_work)

                rows_by_set.append({"label": label, "rows": rows})
                if not self._running["gmarket"]:
                    break

            driver.quit()

            total_rows = sum(len(e["rows"]) for e in rows_by_set)
            if total_rows > 0:
                if multi_mode:
                    save_excel_multi(rows_by_set, origin, dest, out_path, adults=adults)
                else:
                    only = rows_by_set[0]
                    save_excel(only["rows"], origin, dest, date_from.year, date_from.month, out_path, adults=adults)

                summary = "  /  ".join(
                    f"{e['label']}: {sum(1 for r in e['rows'] if r['found'])}/{len(e['rows'])}일"
                    for e in rows_by_set
                )
                self._log("gmarket", f"\n✅ 완료! {summary} → {fname}", "ok")
            else:
                self._log("gmarket", "⚠ 수집된 데이터 없음", "err")

        except Exception as e:
            self._log("gmarket", f"❌ 오류 발생: {e}", "err")
        finally:
            self._running["gmarket"] = False
            self._done("gmarket")

    # ── 에어프레미아 실행 ──
    def _run_airpremia(self, config, date_from, date_to):
        try:
            out_dir = config.get("out_dir") or get_default_output_dir()
            custom_filename = config.get("custom_filename", "")
            raw_sets = config.get("sets", [])
            sets = []
            for i, s in enumerate(raw_sets, 1):
                dest_label = s.get("dest", list(AIRPREMIA_DESTINATIONS.keys())[0])
                sets.append({
                    "label": f"세트{i}_{dest_label}_{s.get('nights', 3)}박" if config.get("set_mode") else "결과",
                    "dest": AIRPREMIA_DESTINATIONS.get(dest_label, "BKK"),
                    "nights": int(s.get("nights", 3)),
                    "adt": int(s.get("adt", 4)),
                })

            date_list = []
            d = date_from
            while d <= date_to:
                date_list.append(d)
                d += datetime.timedelta(days=1)
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

            self._log(
                "airpremia",
                f"🛫 추출 시작(에어프레미아): {len(sets)}개 세트  |  {date_from} ~ {date_to}  |  총 {len(date_list)}일",
                "info",
            )

            driver = airpremia_init_driver(headless=not config.get("show_browser", False))
            rows_by_set = []
            work_done = 0

            for s in sets:
                label = s["label"]
                self._log("airpremia", f"\n▶ {label} 검색 시작 (목적지={s['dest']}, {s['nights']}박, {s['adt']}명)", "info")
                rows = []
                for dep_date in date_list:
                    if not self._running["airpremia"]:
                        self._log("airpremia", "⏹ 중지됨.", "err")
                        break
                    self._log("airpremia", f"  {dep_date.strftime('%Y-%m-%d')} ({dep_date.strftime('%a')}) 조회 중...")
                    result = airpremia_check_one_date(driver, s["dest"], dep_date, s["nights"], s["adt"])
                    rows.append(result)
                    if result.status == "OK":
                        self._log(
                            "airpremia",
                            f"    ✔ {result.departure_date}→{result.return_date}  총액:{result.total_amount:,}원  1인:{result.per_person_rounded:,}원",
                            "ok",
                        )
                    else:
                        self._log("airpremia", f"    ✗ {result.departure_date}→{result.return_date}  {result.status} {result.note}")
                    work_done += 1
                    self._progress("airpremia", work_done, total_work)
                rows_by_set.append({"label": label, "rows": rows, "adt": s["adt"]})
                if not self._running["airpremia"]:
                    break

            driver.quit()

            from openpyxl import Workbook
            wb = Workbook()
            wb.remove(wb.active)
            for entry in rows_by_set:
                sheet_name = re.sub(r"[\\/*?:\[\]]", "_", entry["label"])[:31] or "결과"
                ws = wb.create_sheet(title=sheet_name)
                airpremia_write_styled_sheet(ws, entry["rows"], entry["adt"])
            wb.save(out_path)

            summary = "  /  ".join(
                f"{e['label']}: {sum(1 for r in e['rows'] if r.status == 'OK')}/{len(e['rows'])}일"
                for e in rows_by_set
            )
            self._log("airpremia", f"\n✅ 완료! {summary} → {fname}", "ok")

        except Exception as e:
            import traceback
            self._log("airpremia", f"❌ 오류 발생: {e}\n{traceback.format_exc()}", "err")
        finally:
            self._running["airpremia"] = False
            self._done("airpremia")


# ─────────────────────────────────────────────
#  진입점
# ─────────────────────────────────────────────
def run_app():
    """pywebview 창을 띄우고 앱을 실행한다. launcher.py가 이 함수를 호출한다."""
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    index_path = os.path.join(base_dir, "web", "index.html")
    if not os.path.exists(index_path):
        ctypes.windll.user32.MessageBoxW(
            0,
            "web/index.html 파일을 찾을 수 없습니다.\n"
            "gmarket_air_gui.py와 같은 폴더에 web 폴더(index.html, style.css, app.js)가 있는지 확인하세요.",
            "오류", 0x10
        )
        sys.exit(1)

    icon_path = os.path.join(base_dir, "icon.ico")
    if not os.path.exists(icon_path):
        icon_path = None

    api = Api()
    window = webview.create_window(
        "항공료 자동 추출기",
        index_path,
        js_api=api,
        width=1280,
        height=940,
        min_size=(1000, 700),
    )
    api.set_window(window)
    webview.start(icon=icon_path)


if __name__ == "__main__":
    # gmarket_air_gui.py를 직접 실행할 때만 자체 업데이트 체크 (launcher를 거치면 launcher가 담당)
    check_and_update()
    run_app()
