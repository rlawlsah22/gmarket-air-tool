"""
launcher.py
항공료 자동 추출기 - 런처 (pywebview 버전, 자동 설치 지원)
※ 이 파일만 exe로 빌드합니다. 한 번 빌드 후 다시 빌드할 필요 없습니다.
※ 코드/화면 수정은 GitHub에 파일을 올리면 자동으로 적용됩니다.
※ 빌드 전 `pip install pywin32` 필요 (바탕화면 바로가기 생성용)
"""

import os
import sys
import ctypes
import importlib.util
import urllib.request

# ─────────────────────────────────────────────
#  설정
# ─────────────────────────────────────────────
GITHUB_USER  = "rlawlsah22"
GITHUB_REPO  = "gmarket-air-tool"
RAW_BASE     = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/main"
UPDATE_FILES = [
    "gmarket_air_gui.py",
    "scraper_core.py",
    "airpremia_fare_checker.py",
    "web/index.html",
    "web/style.css",
    "web/app.js",
    "icon.ico",
]

APP_NAME = "항공료 자동 추출기"

# 실제 프로그램 파일이 설치되는 고정 위치 (사용자 눈에 안 띄는 표준 위치)
# 예: C:\Users\사용자명\AppData\Local\항공료추출기
INSTALL_DIR = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "항공료추출기")


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


# ─────────────────────────────────────────────
#  GitHub에서 최신 파일 다운로드
# ─────────────────────────────────────────────
def download_latest():
    """GitHub에서 항상 최신 파일을 받아서 INSTALL_DIR에 저장 (.py, icon.ico는 숨김 처리)"""
    os.makedirs(INSTALL_DIR, exist_ok=True)
    for fname in UPDATE_FILES:
        try:
            url = f"{RAW_BASE}/{fname}"
            dst = os.path.join(INSTALL_DIR, fname)
            dst_dir = os.path.dirname(dst)
            if dst_dir:
                os.makedirs(dst_dir, exist_ok=True)
            _unhide_file_windows(dst)
            with urllib.request.urlopen(url, timeout=10) as r:
                with open(dst, "wb") as f:
                    f.write(r.read())
            if fname.endswith(".py") or fname == "icon.ico":
                _hide_file_windows(dst)
        except Exception:
            pass  # 개별 파일 다운로드 실패 시 기존 파일로 계속 진행


# ─────────────────────────────────────────────
#  바탕화면 바로가기 자동 생성 (최초 1회)
# ─────────────────────────────────────────────
def create_desktop_shortcut_if_needed():
    """
    바탕화면에 이미 바로가기가 있으면 아무 것도 하지 않는다.
    없으면(최초 실행) launcher exe를 가리키는 바로가기를 새로 만든다.
    pywin32가 없는 환경(예: 개발 중 python으로 직접 실행)에서는 조용히 건너뛴다.
    알집 등 압축 프로그램 안에서 바로 실행되어 임시(Temp) 폴더에서 도는 경우,
    그 임시 경로를 가리키는 깨진 바로가기가 만들어지는 것을 막기 위해
    실행 위치가 임시 폴더로 보이면 바로가기 생성을 건너뛴다.
    """
    if os.name != "nt":
        return
    try:
        target = sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__)

        # 임시 폴더(예: ...\AppData\Local\Temp\_AZTMP1_\...)에서 실행 중이면
        # 그 경로를 바로가기에 저장해봐야 곧 사라져서 깨지므로 생성을 건너뛴다.
        temp_dir = os.environ.get("TEMP", "")
        if temp_dir and os.path.normcase(target).startswith(os.path.normcase(temp_dir)):
            return

        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        shortcut_path = os.path.join(desktop, f"{APP_NAME}.lnk")
        if os.path.exists(shortcut_path):
            return  # 이미 있으면 다시 안 만듦 (사용자가 지웠으면 그 의사를 존중)

        import win32com.client  # pywin32
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = target
        shortcut.WorkingDirectory = os.path.dirname(target)
        icon_path = os.path.join(INSTALL_DIR, "icon.ico")
        if os.path.exists(icon_path):
            shortcut.IconLocation = icon_path
        shortcut.save()
    except Exception:
        pass  # 바로가기 생성 실패해도 프로그램 실행 자체는 계속 진행


# ─────────────────────────────────────────────
#  앱 로드 및 실행
# ─────────────────────────────────────────────
def run_app():
    gui_path = os.path.join(INSTALL_DIR, "gmarket_air_gui.py")

    if not os.path.exists(gui_path):
        ctypes.windll.user32.MessageBoxW(
            0,
            "프로그램 파일을 찾을 수 없습니다.\n인터넷 연결을 확인하거나 진모에게 문의하세요.",
            "오류",
            0x10
        )
        sys.exit(1)

    # INSTALL_DIR을 sys.path 맨 앞에 추가 (gmarket_air_gui.py가 scraper_core 등을 import할 수 있도록)
    if INSTALL_DIR in sys.path:
        sys.path.remove(INSTALL_DIR)
    sys.path.insert(0, INSTALL_DIR)

    # py 파일을 importlib으로 동적 로드
    spec = importlib.util.spec_from_file_location("gmarket_air_gui", gui_path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules["gmarket_air_gui"] = mod
    spec.loader.exec_module(mod)

    # gmarket_air_gui.py의 run_app()을 호출해 pywebview 창을 띄움
    mod.run_app()


# ─────────────────────────────────────────────
#  진입점
# ─────────────────────────────────────────────
if __name__ == "__main__":
    download_latest()                     # 항상 최신 파일 다운로드 (INSTALL_DIR로)
    create_desktop_shortcut_if_needed()   # 최초 실행이면 바탕화면 바로가기 생성
    run_app()
