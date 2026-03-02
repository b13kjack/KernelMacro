import sys
import os
import ctypes
from ctypes import wintypes
import math
import random
import threading
import time

from pynput import keyboard, mouse

# Ensure DLLs can be found when packaged with PyInstaller or running locally
if hasattr(os, 'add_dll_directory'):
    try:
        if getattr(sys, 'frozen', False):
            # Running as compiled executable (add sys._MEIPASS and the folder containing the .exe)
            os.add_dll_directory(sys._MEIPASS)
            os.add_dll_directory(os.path.dirname(sys.executable))
        else:
            # Running from source (add current script dir and the dist folder)
            os.add_dll_directory(os.path.abspath(os.path.dirname(__file__)))
            dist_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'dist')
            if os.path.exists(dist_path):
                os.add_dll_directory(dist_path)
    except Exception:
        pass

# Try to load interception-python for kernel-level input injection.
# Falls back to ctypes SendInput if the driver is not installed.
_USE_INTERCEPTION = False
try:
    import interception as _interception
    _interception.auto_capture_devices()
    _USE_INTERCEPTION = True
    print("[MacroPlayer] Interception driver loaded - using kernel-level input.")
except Exception as _interception_err:
    print(f"[MacroPlayer] Interception driver not available ({_interception_err}).")
    print("[MacroPlayer] Falling back to SendInput (may not work in some games).")


DEFAULT_CONTROL_KEY_NAMES = ("f11", "f6", "f8", "f9")

KEYBOARD_ACTIONS = {"press", "release"}
MOUSE_ACTIONS = {"move", "click", "scroll"}
CLICK_STATES = {"press", "release"}

# ---------------------------------------------------------------------------
# pynput key name → interception key name translation
# ---------------------------------------------------------------------------
_PYNPUT_TO_INTERCEPTION = {
    "Key.shift": "shift", "Key.shift_r": "shiftright",
    "Key.ctrl": "ctrl", "Key.ctrl_l": "ctrlleft", "Key.ctrl_r": "ctrlright",
    "Key.alt": "alt", "Key.alt_l": "altleft", "Key.alt_r": "altright",
    "Key.alt_gr": "altright",
    "Key.cmd": "win", "Key.cmd_l": "winleft", "Key.cmd_r": "winright",
    "Key.backspace": "backspace", "Key.enter": "enter", "Key.space": "space",
    "Key.tab": "tab", "Key.caps_lock": "capslock",
    "Key.esc": "esc", "Key.escape": "escape",
    "Key.insert": "insert", "Key.delete": "delete",
    "Key.home": "home", "Key.end": "end",
    "Key.page_up": "pageup", "Key.page_down": "pagedown",
    "Key.up": "up", "Key.down": "down", "Key.left": "left", "Key.right": "right",
    "Key.f1": "f1", "Key.f2": "f2", "Key.f3": "f3", "Key.f4": "f4",
    "Key.f5": "f5", "Key.f6": "f6", "Key.f7": "f7", "Key.f8": "f8",
    "Key.f9": "f9", "Key.f10": "f10", "Key.f11": "f11", "Key.f12": "f12",
    "Key.num_lock": "numlock", "Key.scroll_lock": "scrolllock",
    "Key.print_screen": "printscreen", "Key.pause": "pause",
    "Key.menu": "apps",
}


def _translate_key_for_interception(key_str: str) -> str:
    """Convert a pynput-style key string to an interception key name."""
    # Direct mapping hit
    mapped = _PYNPUT_TO_INTERCEPTION.get(key_str)
    if mapped:
        return mapped

    # Strip 'Key.' prefix for any remaining special keys
    if key_str.startswith("Key."):
        stripped = key_str[4:]
        # Convert underscores (e.g. shift_r → shiftright)
        return stripped.replace("_l", "left").replace("_r", "right").replace("_", "")

    # Regular character – interception expects lowercase single chars
    return key_str.lower()


# ---------------------------------------------------------------------------
# pynput mouse button name → interception button name
# ---------------------------------------------------------------------------
def _translate_button_for_interception(button_str: str) -> str:
    """Convert a pynput Button string to interception button name."""
    b = button_str.lower()
    if "right" in b:
        return "right"
    if "middle" in b:
        return "middle"
    # mouse4, mouse5 etc.
    if "x2" in b or "button9" in b:
        return "mouse5"
    if "x1" in b or "button8" in b:
        return "mouse4"
    return "left"


# ---------------------------------------------------------------------------
# SendInput fallback (kept for systems without the Interception driver)
# ---------------------------------------------------------------------------
user32 = ctypes.WinDLL('user32', use_last_error=True)

INPUT_MOUSE    = 0
INPUT_KEYBOARD = 1
INPUT_HARDWARE = 2

KEYEVENTF_EXTENDEDKEY = 0x0001
KEYEVENTF_KEYUP       = 0x0002
KEYEVENTF_UNICODE     = 0x0004
KEYEVENTF_SCANCODE    = 0x0008

MOUSEEVENTF_MOVE       = 0x0001
MOUSEEVENTF_LEFTDOWN   = 0x0002
MOUSEEVENTF_LEFTUP     = 0x0004
MOUSEEVENTF_RIGHTDOWN  = 0x0008
MOUSEEVENTF_RIGHTUP    = 0x0010
MOUSEEVENTF_MIDDLEDOWN = 0x0020
MOUSEEVENTF_MIDDLEUP   = 0x0040
MOUSEEVENTF_XDOWN      = 0x0080
MOUSEEVENTF_XUP        = 0x0100
MOUSEEVENTF_WHEEL      = 0x0800
MOUSEEVENTF_HWHEEL     = 0x1000
MOUSEEVENTF_ABSOLUTE   = 0x8000

class MOUSEINPUT(ctypes.Structure):
    _fields_ = (("dx",          wintypes.LONG),
                ("dy",          wintypes.LONG),
                ("mouseData",   wintypes.DWORD),
                ("dwFlags",     wintypes.DWORD),
                ("time",        wintypes.DWORD),
                ("dwExtraInfo", ctypes.POINTER(wintypes.ULONG)))

class KEYBDINPUT(ctypes.Structure):
    _fields_ = (("wVk",         wintypes.WORD),
                ("wScan",       wintypes.WORD),
                ("dwFlags",     wintypes.DWORD),
                ("time",        wintypes.DWORD),
                ("dwExtraInfo", ctypes.POINTER(wintypes.ULONG)))

class HARDWAREINPUT(ctypes.Structure):
    _fields_ = (("uMsg",    wintypes.DWORD),
                ("wParamL", wintypes.WORD),
                ("wParamH", wintypes.WORD))

class INPUT(ctypes.Structure):
    class _INPUT(ctypes.Union):
        _fields_ = (("ki", KEYBDINPUT),
                    ("mi", MOUSEINPUT),
                    ("hi", HARDWAREINPUT))
    _anonymous_ = ("_input",)
    _fields_ = (("type",   wintypes.DWORD),
                ("_input", _INPUT))

SCAN_CODES = {
    "a": 0x1E, "b": 0x30, "c": 0x2E, "d": 0x20, "e": 0x12, "f": 0x21, "g": 0x22, "h": 0x23, "i": 0x17, "j": 0x24,
    "k": 0x25, "l": 0x26, "m": 0x32, "n": 0x31, "o": 0x18, "p": 0x19, "q": 0x10, "r": 0x13, "s": 0x1F, "t": 0x14,
    "u": 0x16, "v": 0x2F, "w": 0x11, "x": 0x2D, "y": 0x15, "z": 0x2C,
    "1": 0x02, "2": 0x03, "3": 0x04, "4": 0x05, "5": 0x06, "6": 0x07, "7": 0x08, "8": 0x09, "9": 0x0A, "0": 0x0B,
    "-": 0x0C, "=": 0x0D, "[": 0x1A, "]": 0x1B, "\\": 0x2B, ";": 0x27, "'": 0x28, ",": 0x33, ".": 0x34, "/": 0x35,
    "`": 0x29,
    "Key.shift": 0x2A, "Key.shift_r": 0x36, "Key.ctrl": 0x1D, "Key.ctrl_r": 0x1D,
    "Key.alt": 0x38, "Key.alt_r": 0x38,
    "Key.cmd": 0x5B, "Key.cmd_r": 0x5C,
    "Key.backspace": 0x0E, "Key.enter": 0x1C, "Key.space": 0x39, "Key.tab": 0x0F, "Key.caps_lock": 0x3A,
    "Key.esc": 0x01, "Key.insert": 0x52, "Key.delete": 0x53, "Key.home": 0x47, "Key.end": 0x4F, "Key.page_up": 0x49,
    "Key.page_down": 0x51, "Key.up": 0x48, "Key.down": 0x50, "Key.left": 0x4B, "Key.right": 0x4D,
    "Key.f1": 0x3B, "Key.f2": 0x3C, "Key.f3": 0x3D, "Key.f4": 0x3E, "Key.f5": 0x3F, "Key.f6": 0x40, "Key.f7": 0x41,
    "Key.f8": 0x42, "Key.f9": 0x43, "Key.f10": 0x44, "Key.f11": 0x57, "Key.f12": 0x58
}

EXTENDED_KEYS = {"Key.up", "Key.down", "Key.left", "Key.right", "Key.insert", "Key.delete", "Key.home", "Key.end", "Key.page_up", "Key.page_down"}

def _get_virtual_screen_metrics():
    """Return (x, y, width, height) of the full virtual screen spanning all monitors."""
    SM_XVIRTUALSCREEN = 76
    SM_YVIRTUALSCREEN = 77
    SM_CXVIRTUALSCREEN = 78
    SM_CYVIRTUALSCREEN = 79
    vx = user32.GetSystemMetrics(SM_XVIRTUALSCREEN)
    vy = user32.GetSystemMetrics(SM_YVIRTUALSCREEN)
    vw = user32.GetSystemMetrics(SM_CXVIRTUALSCREEN)
    vh = user32.GetSystemMetrics(SM_CYVIRTUALSCREEN)
    if vw <= 0 or vh <= 0:
        vx, vy = 0, 0
        vw = user32.GetSystemMetrics(0)  # SM_CXSCREEN
        vh = user32.GetSystemMetrics(1)  # SM_CYSCREEN
    return vx, vy, vw, vh

def _send_input(inputs):
    nInputs = len(inputs)
    LPINPUT = INPUT * nInputs
    pInputs = LPINPUT(*inputs)
    cbSize = ctypes.c_int(ctypes.sizeof(INPUT))
    user32.SendInput(nInputs, pInputs, cbSize)

# --- end of ctypes definitions ---


def _coerce_non_negative_float(value, field_name, index):
    try:
        parsed = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"Event {index + 1}: '{field_name}' must be a number.") from exc

    if not math.isfinite(parsed):
        raise ValueError(f"Event {index + 1}: '{field_name}' must be finite.")

    if parsed < 0:
        raise ValueError(f"Event {index + 1}: '{field_name}' must be 0 or greater.")
    return parsed


def _coerce_int(value, field_name, index):
    try:
        return int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"Event {index + 1}: '{field_name}' must be an integer.") from exc


def normalize_events(events):
    if not isinstance(events, list):
        raise ValueError("Macro data must be a list of events.")

    normalized = []

    for index, raw_event in enumerate(events):
        if not isinstance(raw_event, dict):
            raise ValueError(f"Event {index + 1}: each event must be an object.")

        event_type = str(raw_event.get("type", "")).strip().lower()
        action = str(raw_event.get("action", "")).strip().lower()
        event_time = _coerce_non_negative_float(raw_event.get("time", 0), "time", index)

        if event_type == "keyboard":
            if action not in KEYBOARD_ACTIONS:
                raise ValueError(
                    f"Event {index + 1}: keyboard action must be one of {sorted(KEYBOARD_ACTIONS)}."
                )

            key_value = raw_event.get("key", "")
            if key_value in (None, ""):
                raise ValueError(f"Event {index + 1}: keyboard events require a key.")

            normalized.append(
                {
                    "time": event_time,
                    "type": "keyboard",
                    "action": action,
                    "key": str(key_value),
                }
            )
            continue

        if event_type == "timeout":
            normalized.append(
                {
                    "time": event_time,
                    "type": "timeout",
                    "action": "pause",
                }
            )
            continue

        if event_type != "mouse":
            raise ValueError(f"Event {index + 1}: unsupported event type '{event_type}'.")

        if action not in MOUSE_ACTIONS:
            raise ValueError(
                f"Event {index + 1}: mouse action must be one of {sorted(MOUSE_ACTIONS)}."
            )

        event = {
            "time": event_time,
            "type": "mouse",
            "action": action,
            "x": _coerce_int(raw_event.get("x", 0), "x", index),
            "y": _coerce_int(raw_event.get("y", 0), "y", index),
        }

        if action == "click":
            click_state = str(raw_event.get("action_type", "")).strip().lower()
            if click_state not in CLICK_STATES:
                raise ValueError(
                    f"Event {index + 1}: mouse click state must be one of {sorted(CLICK_STATES)}."
                )
            button_value = raw_event.get("button", "")
            if button_value in (None, ""):
                raise ValueError(f"Event {index + 1}: mouse click events require a button.")

            event["button"] = str(button_value)
            event["action_type"] = click_state

        if action == "scroll":
            event["dx"] = _coerce_int(raw_event.get("dx", 0), "dx", index)
            event["dy"] = _coerce_int(raw_event.get("dy", 0), "dy", index)

        normalized.append(event)

    normalized.sort(key=lambda event: event["time"])
    return normalized


def resolve_control_keys(control_key_names):
    resolved_keys = set()

    for key_name in control_key_names or DEFAULT_CONTROL_KEY_NAMES:
        normalized_name = str(key_name).strip().lower()
        key = getattr(keyboard.Key, normalized_name, None)
        if key is not None:
            resolved_keys.add(key)

    return resolved_keys


def _resolve_key_string(key):
    """Extract a usable key string from a pynput key object."""
    try:
        char = key.char
        if char is not None:
            return char
    except AttributeError:
        pass
    return str(key)


class MacroRecorder:
    def __init__(self, control_key_names=None):
        self.events = []
        self._lock = threading.Lock()
        self._start_time = None
        self.running = False
        self.mouse_listener = None
        self.control_key_names = tuple(control_key_names or DEFAULT_CONTROL_KEY_NAMES)
        self.control_keys = resolve_control_keys(self.control_key_names)

    def set_control_keys(self, control_key_names):
        self.control_key_names = tuple(control_key_names or DEFAULT_CONTROL_KEY_NAMES)
        self.control_keys = resolve_control_keys(self.control_key_names)

    def start(self):
        if self.running:
            return

        self.stop()
        with self._lock:
            self.events = []
            self._start_time = time.perf_counter()
        self.running = True

        # Keyboard events are forwarded by the app's global keyboard listener.
        # Only the mouse listener is owned by the recorder.
        self.mouse_listener = mouse.Listener(
            on_move=self.on_move,
            on_click=self.on_click,
            on_scroll=self.on_scroll,
        )

        try:
            self.mouse_listener.start()
        except Exception:
            self.stop()
            raise

    def stop(self):
        self.running = False
        self._stop_listener(self.mouse_listener)
        self.mouse_listener = None
        with self._lock:
            self._start_time = None

    @staticmethod
    def _stop_listener(listener):
        if listener is None:
            return

        try:
            listener.stop()
        except RuntimeError:
            return

    def _record_event(self, event_type, action, **kwargs):
        if not self.running:
            return

        with self._lock:
            start = self._start_time
            if start is None:
                return
            elapsed = time.perf_counter() - start

        event = {
            "time": elapsed,
            "type": event_type,
            "action": action,
            **kwargs,
        }
        with self._lock:
            self.events.append(event)

    def on_press(self, key):
        if key in self.control_keys:
            return
        self._record_event("keyboard", "press", key=_resolve_key_string(key))

    def on_release(self, key):
        if key in self.control_keys:
            return
        self._record_event("keyboard", "release", key=_resolve_key_string(key))

    def on_move(self, x, y):
        self._record_event("mouse", "move", x=x, y=y)

    def on_click(self, x, y, button, pressed):
        click_state = "press" if pressed else "release"
        self._record_event(
            "mouse",
            "click",
            x=x,
            y=y,
            button=str(button),
            action_type=click_state,
        )

    def on_scroll(self, x, y, dx, dy):
        self._record_event("mouse", "scroll", x=x, y=y, dx=dx, dy=dy)


class MacroPlayer:
    def __init__(self):
        self.running = False
        self.thread = None
        self.stop_event = threading.Event()
        self._vscreen_x = 0
        self._vscreen_y = 0
        self._vscreen_w = 1
        self._vscreen_h = 1

    def play(self, events, speed=1.0, loops=1, action_delay=0.0, on_finish=None):
        if self.running:
            return False

        normalized_events = normalize_events(events)
        if not normalized_events:
            if on_finish:
                on_finish()
            return False

        try:
            parsed_speed = float(speed)
        except (TypeError, ValueError):
            parsed_speed = 1.0

        if not math.isfinite(parsed_speed) or parsed_speed <= 0:
            parsed_speed = 1.0

        try:
            parsed_loops = int(loops)
        except (TypeError, ValueError):
            parsed_loops = 1

        if parsed_loops < 1:
            parsed_loops = 1

        try:
            parsed_action_delay = float(action_delay)
        except (TypeError, ValueError):
            parsed_action_delay = 0.0

        if not math.isfinite(parsed_action_delay) or parsed_action_delay < 0:
            parsed_action_delay = 0.0

        self.stop_event.clear()
        self.running = True
        self.thread = threading.Thread(
            target=self._play_thread,
            args=(normalized_events, parsed_speed, parsed_loops, parsed_action_delay, on_finish),
            daemon=True,
        )
        self.thread.start()
        return True

    def stop(self):
        self.stop_event.set()
        self.running = False

        if self.thread and self.thread.is_alive() and self.thread is not threading.current_thread():
            self.thread.join(timeout=1.0)

    def _play_thread(self, events, speed, loops, action_delay, on_finish):
        try:
            # Only needed for SendInput fallback path
            if not _USE_INTERCEPTION:
                vx, vy, vw, vh = _get_virtual_screen_metrics()
                self._vscreen_x = vx
                self._vscreen_y = vy
                self._vscreen_w = vw
                self._vscreen_h = vh

            for _ in range(loops):
                if self.stop_event.is_set():
                    break

                previous_event_time = 0.0

                for event_index, event in enumerate(events):
                    if self.stop_event.is_set():
                        break

                    recorded_delay = max(0.0, (event["time"] - previous_event_time) / speed)
                    previous_event_time = event["time"]

                    extra_delay = 0.0
                    if event_index > 0 and event.get("action") != "move":
                        if action_delay > 0:
                            extra_delay = random.uniform(0.0, action_delay)

                    event_delay = recorded_delay + extra_delay

                    if event_delay and self.stop_event.wait(event_delay):
                        break

                    if _USE_INTERCEPTION:
                        self._execute_event_interception(event)
                    else:
                        self._execute_event_sendinput(event)
        finally:
            self.running = False
            self.thread = None
            self.stop_event.clear()
            if on_finish:
                on_finish()

    # ------------------------------------------------------------------
    # Interception driver path (kernel-level, appears as real hardware)
    # ------------------------------------------------------------------
    def _execute_event_interception(self, event):
        try:
            if event["type"] == "timeout":
                return

            if event["type"] == "keyboard":
                key_str = str(event["key"])
                ic_key = _translate_key_for_interception(key_str)

                if event["action"] == "press":
                    _interception.key_down(ic_key, delay=0)
                elif event["action"] == "release":
                    _interception.key_up(ic_key, delay=0)
                return

            # Mouse events
            x, y = int(event["x"]), int(event["y"])

            if event["action"] == "move":
                _interception.move_to(x, y)
                return

            if event["action"] == "click":
                _interception.move_to(x, y)
                button = _translate_button_for_interception(str(event["button"]))
                if event["action_type"] == "press":
                    _interception.mouse_down(button, delay=0)
                else:
                    _interception.mouse_up(button, delay=0)
                return

            if event["action"] == "scroll":
                _interception.move_to(x, y)
                sdy = event.get("dy", 0)
                # Scroll one unit per recorded event
                if sdy > 0:
                    for _ in range(abs(int(sdy))):
                        _interception.scroll("up")
                elif sdy < 0:
                    for _ in range(abs(int(sdy))):
                        _interception.scroll("down")

        except Exception as exc:
            print(f"[Interception] Error executing event {event}: {exc}")

    # ------------------------------------------------------------------
    # SendInput fallback path (virtual input, may be blocked by games)
    # ------------------------------------------------------------------
    def _to_absolute_coords(self, px, py):
        """Convert pixel coordinates to MOUSEEVENTF_ABSOLUTE values (0-65535 over virtual screen)."""
        ax = int((px - self._vscreen_x) * 65536 / self._vscreen_w)
        ay = int((py - self._vscreen_y) * 65536 / self._vscreen_h)
        return ax, ay

    def _execute_event_sendinput(self, event):
        try:
            if event["type"] == "timeout":
                return

            if event["type"] == "keyboard":
                key_str = str(event["key"])

                scan_code = SCAN_CODES.get(key_str.lower(), 0)
                if scan_code == 0 and key_str:
                    vk = user32.VkKeyScanW(ord(key_str[0])) & 0xFF
                    if vk:
                        scan_code = user32.MapVirtualKeyW(vk, 0)

                if not scan_code:
                    return

                flags = KEYEVENTF_SCANCODE
                if key_str in EXTENDED_KEYS:
                    flags |= KEYEVENTF_EXTENDEDKEY

                if event["action"] == "release":
                    flags |= KEYEVENTF_KEYUP

                inp = INPUT(type=INPUT_KEYBOARD,
                            ki=KEYBDINPUT(wVk=0,
                                          wScan=scan_code,
                                          dwFlags=flags,
                                          time=0,
                                          dwExtraInfo=None))
                _send_input([inp])
                return

            ax, ay = self._to_absolute_coords(int(event["x"]), int(event["y"]))

            if event["action"] == "move":
                inp = INPUT(type=INPUT_MOUSE,
                            mi=MOUSEINPUT(dx=ax, dy=ay, mouseData=0,
                                          dwFlags=MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE,
                                          time=0, dwExtraInfo=None))
                _send_input([inp])
                return

            if event["action"] == "click":
                inp_move = INPUT(type=INPUT_MOUSE,
                                 mi=MOUSEINPUT(dx=ax, dy=ay, mouseData=0,
                                               dwFlags=MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE,
                                               time=0, dwExtraInfo=None))

                button_str = str(event["button"]).lower()
                if "right" in button_str:
                    btn_flag = MOUSEEVENTF_RIGHTDOWN if event["action_type"] == "press" else MOUSEEVENTF_RIGHTUP
                elif "middle" in button_str:
                    btn_flag = MOUSEEVENTF_MIDDLEDOWN if event["action_type"] == "press" else MOUSEEVENTF_MIDDLEUP
                else:
                    btn_flag = MOUSEEVENTF_LEFTDOWN if event["action_type"] == "press" else MOUSEEVENTF_LEFTUP

                inp_click = INPUT(type=INPUT_MOUSE,
                                  mi=MOUSEINPUT(dx=ax, dy=ay, mouseData=0, dwFlags=btn_flag,
                                                time=0, dwExtraInfo=None))

                _send_input([inp_move, inp_click])
                return

            if event["action"] == "scroll":
                w_delta = 120
                sdx = event.get("dx", 0)
                sdy = event.get("dy", 0)

                inp_move = INPUT(type=INPUT_MOUSE,
                                 mi=MOUSEINPUT(dx=ax, dy=ay, mouseData=0,
                                               dwFlags=MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE,
                                               time=0, dwExtraInfo=None))
                inputs = [inp_move]

                if sdy != 0:
                    inputs.append(INPUT(type=INPUT_MOUSE,
                                        mi=MOUSEINPUT(dx=0, dy=0, mouseData=int(sdy * w_delta),
                                                      dwFlags=MOUSEEVENTF_WHEEL,
                                                      time=0, dwExtraInfo=None)))
                if sdx != 0:
                    inputs.append(INPUT(type=INPUT_MOUSE,
                                        mi=MOUSEINPUT(dx=0, dy=0, mouseData=int(sdx * w_delta),
                                                      dwFlags=MOUSEEVENTF_HWHEEL,
                                                      time=0, dwExtraInfo=None)))
                _send_input(inputs)

        except Exception as exc:
            print(f"[SendInput] Error executing event {event}: {exc}")

