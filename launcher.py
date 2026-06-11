"""
launcher.py
G마켓 항공료 자동 추출기 - 런처
※ 이 파일만 exe로 빌드합니다. 한 번 빌드 후 다시 빌드할 필요 없습니다.
※ 코드 수정은 GitHub에 gmarket_air_gui.py / scraper_core.py 올리면 자동 적용됩니다.
"""

import os
import sys
import ctypes
import importlib.util
import urllib.request
import tempfile

# ─────────────────────────────────────────────
#  설정
# ─────────────────────────────────────────────
GITHUB_USER  = "rlawlsah22"
GITHUB_REPO  = "gmarket-air-tool"
RAW_BASE     = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/main"
UPDATE_FILES = ["gmarket_air_gui.py", "scraper_core.py"]

# exe 실행 시 BASE_DIR = exe가 있는 폴더
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# ─────────────────────────────────────────────
#  GitHub에서 최신 파일 다운로드
# ─────────────────────────name───────────────
def download_latest():
    """GitHub에서 항상 최신 py 파일을 받아서 BASE_DIR에 저장"""
    try:
        for fname in UPDATE_FILES:
            url = f"{RAW_BASE}/{fname}"
            dst = os.path.join(BASE_DIR, fname)
            with urllib.request.urlopen(url, timeout=10) as r:
                with open(dst, "wb") as f:
                    f.write(r.read())
    except Exception:
        pass  # 다운로드 실패 시 기존 파일로 실행


# ─────────────────────────────────────────────
#  GUI 로드 및 실행
# ─────────────────────────────────────────────
def run_gui():
    gui_path = os.path.join(BASE_DIR, "gmarket_air_gui.py")

    if not os.path.exists(gui_path):
        ctypes.windll.user32.MessageBoxW(
            0,
            "프로그램 파일을 찾을 수 없습니다.\n인터넷 연결을 확인하거나 진모에게 문의하세요.",
            "오류",
            0x10
        )
        sys.exit(1)

    # BASE_DIR를 sys.path 맨 앞에 추가
    if BASE_DIR in sys.path:
        sys.path.remove(BASE_DIR)
    sys.path.insert(0, BASE_DIR)

    # py 파일을 importlib으로 동적 로드 → Tcl/Tk는 exe 기준으로 정상 동작
    spec = importlib.util.spec_from_file_location("gmarket_air_gui", gui_path)
    mod  = importlib.util.module_from_spec(spec)
    sys.modules["gmarket_air_gui"] = mod
    spec.loader.exec_module(mod)

    app = mod.GmarketAirApp()
    app.mainloop()


# ─────────────────────────────────────────────
#  진입점
# ─────────────────────────────────────────────
if __name__ == "__main__":
    download_latest()  # 항상 최신 파일 다운로드
    run_gui()
