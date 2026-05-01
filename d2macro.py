import time
import subprocess
import os
import psutil
import sys
import json
import threading
import traceback
import ctypes
import ctypes.wintypes
import shutil
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk


def _register_mei_cleanup():
    import tempfile
    mei = getattr(sys, '_MEIPASS', None)
    if not mei:
        return
    tmp = tempfile.gettempdir()
    running_meis = set()
    for proc in psutil.process_iter(['pid', 'exe']):
        try:
            exe = proc.info['exe']
            if exe:
                running_meis.add(os.path.normcase(os.path.dirname(exe)))
        except Exception:
            pass
    for name in os.listdir(tmp):
        if not name.startswith('_MEI'):
            continue
        path = os.path.join(tmp, name)
        if os.path.normcase(path) == os.path.normcase(mei):
            continue
        if os.path.normcase(path) in running_meis:
            continue
        shutil.rmtree(path, ignore_errors=True)
    # _MEI 삭제는 _on_close()에서 cmd 지연 삭제로 처리

_register_mei_cleanup()

_KEYEVENTF_SCANCODE = 0x0008
_KEYEVENTF_KEYUP = 0x0002

_SCAN_CODES = {
    'a': 0x1E, 'c': 0x2E, 'd': 0x20, 'e': 0x12, 'f': 0x21,
    'g': 0x22, 'i': 0x17, 'q': 0x10, 's': 0x1F, 'w': 0x11,
    'x': 0x2D, 'space': 0x39, 'tab': 0x0F, 'escape': 0x01,
}

ULONG_PTR = ctypes.c_ulonglong


class _MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", ctypes.c_long), ("dy", ctypes.c_long),
        ("mouseData", ctypes.c_ulong), ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong), ("dwExtraInfo", ULONG_PTR),
    ]


class _KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", ctypes.c_ushort), ("wScan", ctypes.c_ushort),
        ("dwFlags", ctypes.c_ulong), ("time", ctypes.c_ulong),
        ("dwExtraInfo", ULONG_PTR),
    ]


class _INPUT(ctypes.Structure):
    class _U(ctypes.Union):
        _fields_ = [("mi", _MOUSEINPUT), ("ki", _KEYBDINPUT)]
    _fields_ = [("type", ctypes.c_ulong), ("ii", _U)]


_SendInput = ctypes.windll.user32.SendInput
_SendInput.argtypes = [ctypes.c_uint, ctypes.c_void_p, ctypes.c_int]
_SendInput.restype = ctypes.c_uint


def _send_scan(scan, key_up=False):
    flags = _KEYEVENTF_SCANCODE | (_KEYEVENTF_KEYUP if key_up else 0)
    inp = _INPUT()
    inp.type = 1
    inp.ii.ki.wVk = 0
    inp.ii.ki.wScan = scan
    inp.ii.ki.dwFlags = flags
    inp.ii.ki.time = 0
    inp.ii.ki.dwExtraInfo = 0
    _SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp))


_gdi32 = ctypes.windll.gdi32
_user32 = ctypes.windll.user32


def _get_pixel(x, y):
    hdc = _user32.GetDC(0)
    color = _gdi32.GetPixel(hdc, x, y)
    _user32.ReleaseDC(0, hdc)
    return (color & 0xFF, (color >> 8) & 0xFF, (color >> 16) & 0xFF)


def di_key_down(key):
    scan = _SCAN_CODES.get(key.lower())
    if scan is None:
        raise KeyError(f"'{key}' 키가 _SCAN_CODES에 없습니다")
    _send_scan(scan)


def di_key_up(key):
    scan = _SCAN_CODES.get(key.lower())
    if scan is None:
        raise KeyError(f"'{key}' 키가 _SCAN_CODES에 없습니다")
    _send_scan(scan, key_up=True)


if getattr(sys, 'frozen', False):
    _BASE_DIR = os.path.dirname(sys.executable)
else:
    _BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_PATH = os.path.join(_BASE_DIR, "config.json")

DEFAULT_CONFIG = {
    "pixel_timeout": 60,
    "tick_interval": 1,
    "pixel_tolerance_pct": 5,
    "launcher_type": "uplay",
    "division2_path": "",
    "pixel_colors": {
        "create_character": {"x": 0, "y": 0, "color": None},
        "ingame_loaded":    {"x": 0, "y": 0, "color": None},
        "login_screen":     {"x": 0, "y": 0, "color": None},
    }
}


def load_config():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                saved = json.load(f)
        except (json.JSONDecodeError, OSError):
            return json.loads(json.dumps(DEFAULT_CONFIG))
        for k, v in DEFAULT_CONFIG.items():
            if k not in saved:
                saved[k] = v
        px = saved.get("pixel_colors", {})
        if "title_screen" in px and "create_character" not in px:
            px["create_character"] = px.pop("title_screen")
        ordered = {}
        for name in DEFAULT_CONFIG["pixel_colors"]:
            ordered[name] = px.get(name, DEFAULT_CONFIG["pixel_colors"][name].copy())
        saved["pixel_colors"] = ordered
        return saved
    return json.loads(json.dumps(DEFAULT_CONFIG))


def save_config(cfg):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


class MacroEngine:
    _UBISOFT_PROCS = frozenset({
        "thedivision2.exe",
        "upc.exe",
        "uplaywebcore.exe",
        "ubisoftconnectwebcore.exe",
        "ubisoftgamelauncher.exe",
    })

    def __init__(self, config, log_callback=None, clear_log_callback=None):
        self.config = config
        self.running = False
        self.thread = None
        self.log = log_callback or print
        self.clear_log = clear_log_callback or (lambda: None)
        self.loop_count = 0

    def stop(self):
        self.running = False

    def start(self):
        if self.thread and self.thread.is_alive():
            return
        self.running = True
        self.loop_count = 0
        self.thread = threading.Thread(target=self._run_loop, daemon=True)
        self.thread.start()

    def _wait(self, ms):
        end = time.time() + ms / 1000.0
        while time.time() < end and self.running:
            time.sleep(max(0.01, min(0.1, end - time.time())))

    def _color_match(self, c1, c2):
        pct = self.config["pixel_tolerance_pct"]
        tol = 255 * pct / 100
        return all(abs(a - b) <= tol for a, b in zip(c1, c2))

    def _read_pixels(self, x, y):
        rx = max(x - 1, 0)
        ry = max(y - 1, 0)
        return [_get_pixel(rx + px, ry + py) for py in range(4) for px in range(4)]

    def _wait_for_pixel(self, name):
        info = self.config["pixel_colors"][name]
        x, y, target = info["x"], info["y"], info["color"]

        if target is None:
            self.log(f"[픽셀 오류] '{name}' 색상 미설정 — 캡처를 먼저 하세요")
            self.running = False
            return False

        tr, tg, tb = target
        timeout = self.config["pixel_timeout"]
        interval = self.config["tick_interval"]
        self.log(f"[픽셀 대기] '{name}' 좌표=({x},{y}) 목표=#{tr:02X}{tg:02X}{tb:02X} 타임아웃={timeout}초 틱={interval}초")
        start = time.time()

        tick_count = 0
        while self.running and (time.time() - start < timeout):
            try:
                pixels = self._read_pixels(x, y)
                if any(self._color_match(p, target) for p in pixels):
                    self.log(f"[픽셀 감지] '{name}' 일치 — 경과={time.time()-start:.1f}초")
                    return True
                tick_count += 1
                if tick_count % 5 == 1:
                    cr, cg, cb = pixels[0]
                    self.log(f"[픽셀 폴링] '{name}' 현재=#{cr:02X}{cg:02X}{cb:02X} 목표=#{tr:02X}{tg:02X}{tb:02X} 경과={time.time()-start:.0f}초")
            except Exception as e:
                self.log(f"[픽셀 읽기 오류] {e}")
            time.sleep(interval)

        if self.running:
            self.log(f"[픽셀 타임아웃] '{name}' {timeout}초 초과 — 다음 단계 진행 불가")
        return False

    def _focus_game(self):
        found = ctypes.c_ulong(0)
        def _cb(h, _):
            pid = ctypes.c_ulong(0)
            _user32.GetWindowThreadProcessId(h, ctypes.byref(pid))
            try:
                p = psutil.Process(pid.value)
                if p.name().lower() == "thedivision2.exe" and _user32.IsWindowVisible(h):
                    found.value = h
                    return False
            except Exception:
                pass
            return True
        _WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.c_int)
        _user32.EnumWindows(_WNDENUMPROC(_cb), 0)
        if found.value:
            self.log(f"TheDivision2.exe 창 발견 (hwnd={found.value:#010x})")
            _user32.SetForegroundWindow(found.value)
            _user32.BringWindowToTop(found.value)
            time.sleep(1.5)
        else:
            self.log("TheDivision2.exe 창 없음 — 포커스 생략")

    def _press(self, key, hold_ms=60):
        if not self.running:
            return
        di_key_down(key)
        self._wait(hold_ms)
        di_key_up(key)

    def _tap(self, key):
        self._press(key, 60)

    def _click(self):
        if not self.running:
            return
        ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0)
        ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)

    def _kill_process(self, name):
        killed = []
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'] and proc.info['name'].lower() == name.lower():
                try:
                    proc.kill()
                    killed.append(proc)
                    self.log(f"{name} PID={proc.info['pid']} — kill 신호 전송")
                except (psutil.NoSuchProcess, psutil.AccessDenied) as e:
                    self.log(f"{name} PID={proc.info['pid']} — {e}")
        if killed:
            psutil.wait_procs(killed, timeout=8)
            self.log(f"{name} — {len(killed)}개 소멸 확인 완료")
        else:
            self.log(f"{name} — 실행 중인 프로세스 없음")

    def _kill_division2(self):
        self._kill_process("TheDivision2.exe")

    def _kill_all(self):
        found_procs = []
        for proc in psutil.process_iter(['pid', 'name']):
            n = (proc.info.get('name') or '').lower()
            if n in self._UBISOFT_PROCS:
                found_procs.append(proc)

        if not found_procs:
            self.log("종료 대상 Ubisoft 프로세스 없음")
            return

        self.log(f"종료된 프로세스 {len(found_procs)}개: " +
                 ", ".join(f"{p.info['name']}(PID={p.info['pid']})" for p in found_procs))

        killed = []
        for proc in found_procs:
            try:
                proc.kill()
                killed.append(proc)
            except (psutil.NoSuchProcess, psutil.AccessDenied) as e:
                self.log(f"{proc.info['name']} PID={proc.info['pid']} — {e}")

        if killed:
            psutil.wait_procs(killed, timeout=8)

        deadline = time.time() + 10
        while self.running and time.time() < deadline:
            still_alive = [
                p for p in psutil.process_iter(['name'])
                if (p.info.get('name') or '').lower() in self._UBISOFT_PROCS
            ]
            if not still_alive:
                self.log(f"모든 프로세스 소멸 확인 — 대기={10-(deadline-time.time()):.1f}")
                break
            time.sleep(0.5)
        else:
            if self.running:
                still = [p.info.get('name') for p in psutil.process_iter(['name'])
                         if (p.info.get('name') or '').lower() in self._UBISOFT_PROCS]
                self.log(f"잔존 프로세스: {still}")

    def _post_kill_settle(self):
        launcher = self.config.get("launcher_type", "uplay")
        if launcher == "steam":
            self.log("Steam 게임 종료 처리 완료 대기 중 (최대 15초)")
            deadline = time.time() + 15
            while self.running and time.time() < deadline:
                still = [p for p in psutil.process_iter(['name'])
                         if (p.info.get('name') or '').lower() == 'thedivision2.exe']
                if not still:
                    elapsed = 15 - (deadline - time.time())
                    self.log(f"TheDivision2.exe 소멸 확인 (경과={elapsed:.1f}초) — 5초 추가 정착")
                    self._wait(5000)
                    return
                time.sleep(0.5)
            self.log("소멸 미확인 — 강제 진행")
        else:
            self._wait(2000)

    def _run_program(self, path):
        self.log(f"Popen: {path}")
        subprocess.Popen([path])

    def _launch_game(self):
        launcher = self.config.get("launcher_type", "uplay")
        if launcher == "steam":
            self.log("Steam으로 실행")
            os.startfile("steam://rungameid/2221490")
        else:
            self.log(f"Uplay으로 실행")
            self._run_program(self.config["division2_path"])

    def _run_loop(self):
        try:
            while self.running:
                self.loop_count += 1
                self.clear_log()
                self.log(f"=== 루프 #{self.loop_count} 시작 ===")
                self._run_macro()
        except Exception as e:
            self.log(f"[치명적 오류] {type(e).__name__}: {e}\n{traceback.format_exc()}")
        finally:
            self.running = False
            self.log("=== 매크로 종료 ===")

    def _run_macro(self):
        w = self._wait
        tap = self._tap
        press = self._press

        self._focus_game()
        w(60); self._click()
        w(60); self._click()
        w(400)
        for _ in range(3):
            tap('c'); w(60)

        w(1200); tap('space'); w(60)
        di_key_down('space'); w(1200); di_key_up('space')

        self.log("캐릭터 생성 화면 픽셀 대기")
        w(10000)
        if not self._wait_for_pixel("create_character"):
            self.log("create_character 픽셀 미감지")
            return
        w(500)

        self.log("TheDivision2.exe 종료 후 게임 재실행")
        self._kill_division2()
        self._post_kill_settle()
        sw = ctypes.windll.user32.GetSystemMetrics(0)
        sh = ctypes.windll.user32.GetSystemMetrics(1)
        ctypes.windll.user32.SetCursorPos(sw // 2, sh // 2)
        self._launch_game()
        w(50000)

        self.log("로그인 화면 픽셀 대기")
        if not self._wait_for_pixel("login_screen"):
            if not self.running:
                return
            self.log("로그인 화면 미감지 — 전체 종료 후 재실행")
            self._kill_all()
            self._post_kill_settle()
            self._launch_game()
            w(60000)
            if not self._wait_for_pixel("login_screen"):
                self.log("재시도 후에도 로그인 화면 미감지")
                return
        w(500)
        self._focus_game()
        self.log("로그인 화면 감지됨")
        tap('space'); w(10000)

        self.log("인게임 로드 픽셀 대기")
        if not self._wait_for_pixel("ingame_loaded"):
            self.log("ingame_loaded 픽셀 미감지")
            return
        w(500)
        self._focus_game()
        self.log("인게임 확인 — 이동 시퀀스 시작")
        press('w', 2000); w(120)
        press('a', 1000); w(120)
        press('w', 3000); w(120)
        press('a', 1000); w(120)
        press('w', 1000); w(120)

        press('f', 600); w(600)
        tap('e'); w(600)
        tap('d'); w(600)
        tap('space'); w(600)
        tap('x'); w(600)
        tap('s'); w(600)
        tap('space'); w(1200)

        for _ in range(4):
            tap('f'); w(60)
        w(600)

        tap('q'); w(60)
        tap('q'); w(1200)

        press('tab', 1800); w(600)
        tap('space'); w(600)
        tap('i'); w(1200)
        tap('escape'); w(1200)
        tap('g'); w(600)
        tap('space'); w(5000)

        self.log("메인메뉴 복귀")
        if not self._wait_for_pixel("login_screen"):
            self.log("로그인 화면 복귀 미감지")
            return


class MacroApp:
    PICK_KEY = "f7"
    PICK_COUNTDOWN = 3

    def __init__(self):
        self.config = load_config()
        self.engine = None
        self.picking = False

        self.root = tk.Tk()
        self._apply_dpi_scaling()
        self.root.title("D2 Macro")
        self.root.resizable(True, True)
        self.root.minsize(560, 600)

        if getattr(sys, 'frozen', False):
            icon_path = os.path.join(getattr(sys, '_MEIPASS'), "images", "div2_icon.ico")
        else:
            icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "images", "div2_icon.ico")
        if os.path.exists(icon_path):
            self.root.iconbitmap(icon_path)

        self._build_ui()
        self._start_hotkey_listener()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()

    def _apply_dpi_scaling(self):
        try:
            dpi = self.root.winfo_fpixels('1i')
            self._dpi_scale = dpi / 96.0
            scale = dpi / 72.0
            self.root.tk.call('tk', 'scaling', scale)
        except Exception:
            self._dpi_scale = 1.0

    def _build_ui(self):
        root = self.root
        PX, PY = 8, 4

        warn = tk.Label(root, text=(
            "본 프로그램은 AI로 제작되어 불안정할 수 있습니다. "
            "사용에 따른 모든 책임은 사용자 본인에게 있습니다."
        ), foreground="red", font=("맑은 고딕", 8), wraplength=int(540 * self._dpi_scale), justify="center")
        warn.pack(padx=PX, pady=(4, 0))

        tk.Label(root, text="Made by Vepley.AMD (2026-04), v1.2",
                 foreground="gray", font=("맑은 고딕", 8)).pack(padx=PX, pady=(0, 2))

        frame_hk = ttk.LabelFrame(root, text="핫키")
        frame_hk.pack(fill="x", padx=PX, pady=PY)
        ttk.Label(frame_hk, text="F5: 시작    F6: 중단    F7: 픽셀 찍기 확인",
                  font=("맑은 고딕", 10, "bold")).pack(padx=PX, pady=6)

        frame_path = ttk.LabelFrame(root, text="경로 설정")
        frame_path.pack(fill="x", padx=PX, pady=PY)

        ttk.Label(frame_path, text="Division2.exe:").grid(row=0, column=0, sticky="w", padx=PX, pady=PY)
        self.var_d2path = tk.StringVar(value=self.config["division2_path"])
        self.entry_d2path = ttk.Entry(frame_path, textvariable=self.var_d2path, width=48)
        self.entry_d2path.grid(row=0, column=1, padx=PX, pady=PY)
        self.btn_d2path = ttk.Button(frame_path, text="찾기", width=5,
                   command=lambda: self._browse_file(self.var_d2path, "TheDivision2.exe", [("Executable", "*.exe")]))
        self.btn_d2path.grid(row=0, column=2, padx=PX, pady=PY)

        ttk.Label(frame_path, text="게임 종료:").grid(row=1, column=0, sticky="w", padx=PX, pady=PY)
        ttk.Label(frame_path, text="psutil.terminate — TheDivision2.exe + upc.exe  (자동)",
                  foreground="gray").grid(row=1, column=1, sticky="w", padx=PX, pady=PY)

        frame_launcher = ttk.LabelFrame(root, text="런처 선택")
        frame_launcher.pack(fill="x", padx=PX, pady=PY)

        self.var_launcher = tk.StringVar(value=self.config.get("launcher_type", "uplay"))
        frame_rb = ttk.Frame(frame_launcher)
        frame_rb.pack(anchor="center", pady=PY)
        ttk.Radiobutton(frame_rb, text="Uplay (Ubisoft Connect)",
                        variable=self.var_launcher, value="uplay",
                        command=self._on_launcher_change).pack(side="left", padx=PX*2)
        ttk.Radiobutton(frame_rb, text="Steam",
                        variable=self.var_launcher, value="steam",
                        command=self._on_launcher_change).pack(side="left", padx=PX*2)

        self._on_launcher_change()

        frame_tm = ttk.LabelFrame(root, text="타이밍 설정")
        frame_tm.pack(fill="x", padx=PX, pady=PY)

        ttk.Label(frame_tm, text="픽셀 타임아웃 (초):").grid(row=0, column=0, sticky="w", padx=PX, pady=PY)
        self.var_timeout = tk.IntVar(value=self.config["pixel_timeout"])
        ttk.Spinbox(frame_tm, from_=10, to=9999, textvariable=self.var_timeout, width=8).grid(row=0, column=1, padx=PX, pady=PY)

        ttk.Label(frame_tm, text="틱 간격 (초):").grid(row=0, column=2, sticky="w", padx=PX, pady=PY)
        self.var_tick = tk.DoubleVar(value=self.config["tick_interval"])
        ttk.Spinbox(frame_tm, from_=0.1, to=10, increment=0.1, textvariable=self.var_tick, width=8).grid(row=0, column=3, padx=PX, pady=PY)

        ttk.Label(frame_tm, text="색상 오차 (%):").grid(row=1, column=0, sticky="w", padx=PX, pady=PY)
        self.var_tol = tk.IntVar(value=self.config["pixel_tolerance_pct"])
        ttk.Spinbox(frame_tm, from_=0, to=100, textvariable=self.var_tol, width=8).grid(row=1, column=1, padx=PX, pady=PY)

        frame_px = ttk.LabelFrame(root, text="픽셀 체크포인트")
        frame_px.pack(fill="x", padx=PX, pady=PY)

        if getattr(sys, 'frozen', False):
            img_dir = os.path.join(getattr(sys, '_MEIPASS'), "images")
        else:
            img_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "images")
        help_images = {
            "create_character": os.path.join(img_dir, "create_character_image.png"),
            "ingame_loaded":    os.path.join(img_dir, "ingame_loaded_image.png"),
            "login_screen":     os.path.join(img_dir, "login_screen_image.png"),
        }

        self.pixel_widgets = {}
        for i, (name, info) in enumerate(self.config["pixel_colors"].items()):
            ttk.Label(frame_px, text=f"{name}:").grid(row=i, column=0, sticky="w", padx=PX, pady=PY)

            var_x = tk.IntVar(value=info["x"])
            ttk.Label(frame_px, text="X:").grid(row=i, column=1, padx=(PX, 0), pady=PY)
            ttk.Entry(frame_px, textvariable=var_x, width=6).grid(row=i, column=2, padx=(0, PX), pady=PY)

            var_y = tk.IntVar(value=info["y"])
            ttk.Label(frame_px, text="Y:").grid(row=i, column=3, padx=(PX, 0), pady=PY)
            ttk.Entry(frame_px, textvariable=var_y, width=6).grid(row=i, column=4, padx=(0, PX), pady=PY)

            img_path = help_images.get(name, "")
            ttk.Button(frame_px, text="?", width=2,
                       command=lambda p=img_path, n=name: self._show_help_image(p, n)).grid(row=i, column=5, padx=PX, pady=PY)

            color_frame = ttk.Frame(frame_px)
            color_frame.grid(row=i, column=6, padx=PX, pady=PY)

            color_swatch = tk.Label(color_frame, width=2, height=1, relief="solid", borderwidth=1)
            color_swatch.pack(side="left", padx=(0, 2))

            ttk.Label(color_frame, text="#").pack(side="left")
            hex_str = ""
            if info["color"]:
                r, g, b = info["color"]
                hex_str = f"{r:02X}{g:02X}{b:02X}"
                color_swatch.config(background=f"#{hex_str}")
            else:
                color_swatch.config(background="#cccccc")
            var_hex = tk.StringVar(value=hex_str)
            hex_entry = ttk.Entry(color_frame, textvariable=var_hex, width=8)
            hex_entry.pack(side="left")
            var_hex.trace_add("write", lambda *_, n=name: self._on_hex_change(n))

            ttk.Button(frame_px, text="찍기", width=5,
                       command=lambda n=name: self._start_pick(n)).grid(row=i, column=7, padx=PX, pady=PY)

            self.pixel_widgets[name] = {"var_x": var_x, "var_y": var_y, "var_hex": var_hex, "swatch": color_swatch}

        frame_btn = ttk.Frame(root)
        frame_btn.pack(fill="x", padx=PX, pady=PY)

        ttk.Button(frame_btn, text="설정 불러오기", command=self._load).pack(side="left", padx=PX, pady=PY)
        ttk.Button(frame_btn, text="설정 저장", command=self._save).pack(side="left", padx=PX, pady=PY)

        self.btn_start = ttk.Button(frame_btn, text="시작 (F5)", command=self._start_macro)
        self.btn_start.pack(side="left", padx=PX, pady=PY)

        self.btn_stop = ttk.Button(frame_btn, text="중단 (F6)", command=self._stop_macro, state="disabled")
        self.btn_stop.pack(side="left", padx=PX, pady=PY)

        self.var_status = tk.StringVar(value="대기중")
        ttk.Label(root, textvariable=self.var_status, font=("맑은 고딕", 10, "bold")).pack(padx=PX, pady=PY)

        frame_log = ttk.LabelFrame(root, text="로그")
        frame_log.pack(fill="both", expand=True, padx=PX, pady=PY)

        self.log_text = tk.Text(frame_log, height=6, state="disabled", font=("Consolas", 9))
        scrollbar = ttk.Scrollbar(frame_log, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)

        github_link = tk.Label(root, text="GitHub: YoungHoon02", foreground="blue",
                               font=("맑은 고딕", 8), cursor="hand2")
        github_link.pack(pady=(2, 4))
        github_link.bind("<Button-1>", lambda _: os.startfile("https://github.com/YoungHoon02"))

    def _show_help_image(self, img_path, name):
        if not os.path.exists(img_path):
            messagebox.showinfo("예시 이미지", f"'{name}' 예시 이미지가 없습니다.\n경로: {img_path}")
            return

        popup = tk.Toplevel(self.root)
        popup.title(f"예시 위치 - {name}")
        popup.resizable(False, False)

        img = Image.open(img_path)
        max_w, max_h = 800, 600
        if img.width > max_w or img.height > max_h:
            img.thumbnail((max_w, max_h), Image.Resampling.LANCZOS)

        photo = ImageTk.PhotoImage(img)
        label = tk.Label(popup, image=photo)
        setattr(label, '_photo', photo)
        label.pack()

        popup.grab_set()
        popup.focus_set()

    def _on_launcher_change(self):
        state = "disabled" if self.var_launcher.get() == "steam" else "normal"
        self.entry_d2path.config(state=state)
        self.btn_d2path.config(state=state)

    def _on_hex_change(self, name):
        w = self.pixel_widgets[name]
        hex_val = w["var_hex"].get().strip().lstrip("#")
        if len(hex_val) == 6:
            try:
                r, g, b = int(hex_val[0:2], 16), int(hex_val[2:4], 16), int(hex_val[4:6], 16)
                w["swatch"].config(background=f"#{hex_val}")
                self.config["pixel_colors"][name]["color"] = [r, g, b]
            except ValueError:
                pass

    def _browse_file(self, var, default_name, filetypes):
        path = filedialog.askopenfilename(
            title=f"{default_name} 선택",
            filetypes=filetypes + [("All files", "*.*")]
        )
        if path:
            var.set(path)

    def _start_hotkey_listener(self):
        try:
            import keyboard
            def _on_f5(_): self.root.after(0, self._start_macro)
            def _on_f6(_): self.root.after(0, self._stop_macro)
            keyboard.on_press_key("f5", _on_f5)
            keyboard.on_press_key("f6", _on_f6)
            self._log_msg("[핫키 등록] F5=시작 / F6=중단 / F7=픽셀 찍기 확인")
        except ImportError:
            self._log_msg("[핫키 오류] 'keyboard' 미설치 — pip install keyboard")

    def _start_pick(self, name):
        if self.picking:
            return
        self.picking = True
        self._pick_target = name
        self._log_msg(f"[픽셀 찍기] '{name}' — 원하는 위치로 마우스를 이동 후 F7을 누르세요")
        self.var_status.set(f"찍기 모드: 마우스 이동 후 F7")
        self.root.iconify()

        try:
            import keyboard
            def _on_f7(_):
                keyboard.unhook(self._f7_hook)
                self.root.after(0, self._finish_pick)
            self._f7_hook = keyboard.on_press_key("f7", _on_f7)
        except ImportError:
            self.root.deiconify()
            self.picking = False
            messagebox.showerror("오류", "keyboard 라이브러리가 필요합니다")

    def _finish_pick(self):
        name = self._pick_target
        try:
            pt = ctypes.wintypes.POINT()
            ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
            x, y = pt.x, pt.y
            color = list(_get_pixel(x, y))

            self.config["pixel_colors"][name] = {"x": x, "y": y, "color": color}

            w = self.pixel_widgets[name]
            w["var_x"].set(x)
            w["var_y"].set(y)
            r, g, b = color
            w["var_hex"].set(f"{r:02X}{g:02X}{b:02X}")

            self._log_msg(f"[픽셀 찍기 완료] {name}: 좌표=({x},{y}) 색상=#{r:02X}{g:02X}{b:02X}")
        except Exception as e:
            self._log_msg(f"[픽셀 찍기 오류] {e}")
        finally:
            self.picking = False
            self.var_status.set("대기중")
            self.root.deiconify()
            self.root.lift()

    def _load(self):
        path = filedialog.askopenfilename(
            title="설정 파일 선택",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                saved = json.load(f)
            for k, v in DEFAULT_CONFIG.items():
                if k not in saved:
                    saved[k] = v
            self.config = saved
            self.var_timeout.set(self.config["pixel_timeout"])
            self.var_tick.set(self.config["tick_interval"])
            self.var_tol.set(self.config["pixel_tolerance_pct"])
            self.var_launcher.set(self.config.get("launcher_type", "uplay"))
            self.var_d2path.set(self.config["division2_path"])
            for name, w in self.pixel_widgets.items():
                info = self.config["pixel_colors"].get(name, {})
                w["var_x"].set(info.get("x", 0))
                w["var_y"].set(info.get("y", 0))
                color = info.get("color")
                if color:
                    r, g, b = color
                    w["var_hex"].set(f"{r:02X}{g:02X}{b:02X}")
                else:
                    w["var_hex"].set("")
                    w["swatch"].config(background="#cccccc")
            self._log_msg(f"[설정 불러오기] {os.path.basename(path)} 로드 완료")
        except Exception as e:
            messagebox.showerror("불러오기 실패", str(e))

    def _apply_to_config(self):
        self.config["pixel_timeout"] = self.var_timeout.get()
        self.config["tick_interval"] = self.var_tick.get()
        self.config["pixel_tolerance_pct"] = self.var_tol.get()
        self.config["launcher_type"] = self.var_launcher.get()
        self.config["division2_path"] = self.var_d2path.get()
        for name, w in self.pixel_widgets.items():
            self.config["pixel_colors"][name]["x"] = w["var_x"].get()
            self.config["pixel_colors"][name]["y"] = w["var_y"].get()
            hex_val = w["var_hex"].get().strip().lstrip("#")
            if len(hex_val) == 6:
                try:
                    r, g, b = int(hex_val[0:2], 16), int(hex_val[2:4], 16), int(hex_val[4:6], 16)
                    self.config["pixel_colors"][name]["color"] = [r, g, b]
                except ValueError:
                    pass

    def _save(self):
        self._apply_to_config()
        path = filedialog.asksaveasfilename(
            title="설정 파일 저장",
            defaultextension=".json",
            initialfile="config.json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.config, f, indent=2, ensure_ascii=False)
        self._log_msg(f"[설정 저장] {os.path.basename(path)} 저장 완료")

    def _start_macro(self):
        if self.engine and self.engine.running:
            return
        if self.picking:
            return
        missing = []

        d2 = self.var_d2path.get()
        if self.var_launcher.get() == "uplay" and (not d2 or not os.path.exists(d2)):
            missing.append("Division2.exe 경로")

        for name, w in self.pixel_widgets.items():
            x, y = w["var_x"].get(), w["var_y"].get()
            color = self.config["pixel_colors"][name].get("color")
            if x == 0 and y == 0:
                missing.append(f"{name} 좌표")
            if color is None:
                missing.append(f"{name} 색상")

        if missing:
            msg = "다음 항목을 설정해주세요:\n\n" + "\n".join(f"  - {m}" for m in missing)
            messagebox.showwarning("설정 미완료", msg)
            return

        self._apply_to_config()
        self._log_msg(f"매크로 시작 — 런처={self.config.get('launcher_type','uplay')} "
                      f"타임아웃={self.config['pixel_timeout']}초 "
                      f"틱={self.config['tick_interval']}초 "
                      f"오차={self.config['pixel_tolerance_pct']}%")
        self.engine = MacroEngine(
            self.config,
            log_callback=self._log_msg,
            clear_log_callback=self._clear_log
        )
        self.engine.start()
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.var_status.set("실행중...")
        self._monitor_engine()

    def _stop_macro(self):
        if self.engine:
            self.engine.stop()
        self._log_msg("매크로 중단 — 사용자 요청")
        self.btn_start.config(state="normal")
        self.btn_stop.config(state="disabled")
        self.var_status.set("대기중")

    def _monitor_engine(self):
        if self.engine and self.engine.running:
            self.var_status.set(f"실행중... (루프 #{self.engine.loop_count})")
            self.root.after(500, self._monitor_engine)
        else:
            self.btn_start.config(state="normal")
            self.btn_stop.config(state="disabled")
            self.var_status.set("대기중")

    def _log_msg(self, msg):
        def _append():
            self.log_text.config(state="normal")
            self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")
            self.log_text.config(state="disabled")
        self.root.after(0, _append)

    def _clear_log(self):
        def _clear():
            self.log_text.config(state="normal")
            self.log_text.delete("1.0", "end")
            self.log_text.config(state="disabled")
        self.root.after(0, _clear)

    def _on_close(self):
        self._log_msg("창 닫힘 — 매크로 및 핫키 해제")
        if self.engine:
            self.engine.stop()
        try:
            import keyboard
            keyboard.unhook_all()
        except Exception:
            pass
        mei = getattr(sys, '_MEIPASS', None)
        if mei and os.path.isdir(mei):
            subprocess.Popen(
                f'cmd /c ping 127.0.0.1 -n 2 >nul 2>&1 & rmdir /s /q "{mei}"',
                shell=True,
                creationflags=subprocess.CREATE_NO_WINDOW | subprocess.DETACHED_PROCESS,
            )
        self.root.destroy()


def _ensure_admin():
    if getattr(sys, 'frozen', False):
        return
    import ctypes as _ct
    if _ct.windll.shell32.IsUserAnAdmin():
        return
    _ct.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, subprocess.list2cmdline(sys.argv), None, 1
    )
    sys.exit()


if __name__ == "__main__":
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
    _ensure_admin()
    MacroApp()
