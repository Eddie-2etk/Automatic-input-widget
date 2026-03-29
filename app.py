import subprocess
import sys
import time
import random
import math
import tkinter as tk
from pathlib import Path
from threading import Lock, Thread
from tkinter import messagebox, ttk

VENDOR_PATH = Path(__file__).with_name(".vendor")
if VENDOR_PATH.exists():
    sys.path.insert(0, str(VENDOR_PATH))

from pynput import keyboard


APP_TITLE = "Word 自动输入小工具"
COMMON_CJK_TYPO_CHARS = (
    "的是了一我在人有他这中大来上个们到说国和地也子时道出而要于就下得可你年生"
)
ASCII_TYPO_CHARS = "abcdefghijklmnopqrstuvwxyz"
DIGIT_TYPO_CHARS = "0123456789"
PUNCTUATION_TYPO_CHARS = ",.;:!?-_=+"
SENTENCE_END_CHARS = "。！？.!?"
CLAUSE_BREAK_CHARS = "，；：,;:"
OPENING_CHARS = "([{\"'“‘（【《"
CLOSING_CHARS = ")]}\"'”’）】》"
STYLE_CONFIGS = {
    "自然型": {
        "thought": 0.85,
        "slip": 0.75,
        "phrase_typo": 0.18,
        "burst_pause": 0.85,
    },
    "思考型": {
        "thought": 1.35,
        "slip": 0.65,
        "phrase_typo": 0.16,
        "burst_pause": 1.25,
    },
    "手滑型": {
        "thought": 0.75,
        "slip": 1.35,
        "phrase_typo": 0.42,
        "burst_pause": 0.95,
    },
    "结合型": {
        "thought": 1.22,
        "slip": 1.18,
        "phrase_typo": 0.38,
        "burst_pause": 1.12,
    },
}


class UserCancelledError(Exception):
    pass


def build_applescript_command(
    lines: list[str], args: list[str] | None = None
) -> list[str]:
    command = ["osascript"]
    for line in lines:
        command.extend(["-e", line])
    if args:
        command.extend(args)
    return command


class WordTyperApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("760x560")
        self.root.minsize(680, 500)

        self.status_var = tk.StringVar(
            value="把内容贴进下方文本框，然后点击按钮即可。"
        )
        self.state_var = tk.StringVar(value="空闲")
        self.state_color = "#6b7280"
        self.speed_var = tk.DoubleVar(value=0.05)
        self.speed_text_var = tk.StringVar(value="0.05 秒/字")
        self.human_mode_var = tk.BooleanVar(value=True)
        self.style_var = tk.StringVar(value="结合型")
        self.thought_var = tk.DoubleVar(value=1.00)
        self.thought_text_var = tk.StringVar(value="1.00x")
        self.typo_rate_var = tk.DoubleVar(value=0.06)
        self.typo_rate_text_var = tk.StringVar(value="6%")
        self.estimate_var = tk.StringVar(value="预计时间：请先输入内容")
        self.action_buttons: list[tk.Widget] = []
        self.current_process: subprocess.Popen[str] | None = None
        self.process_lock = Lock()
        self.state_lock = Lock()
        self.stop_requested = False
        self.pause_requested = False
        self.typing_active = False
        self.last_start_mode: str | None = None
        self.pending_restart_mode: str | None = None
        self.hotkey_listener: keyboard.Listener | None = None
        self.last_space_press = 0.0
        self.ignore_space_until = 0.0
        self.session_speed_factor = 1.0
        self.fatigue_bias = 0.0
        self.focus_wave_depth = 0.0
        self.focus_wave_cycles = 1.0

        self._build_ui()
        self.set_state("空闲", "#6b7280")
        self.update_estimate()
        self.start_hotkey_listener()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=18)
        container.pack(fill="both", expand=True)

        ttk.Label(
            container,
            text="把一段文字自动打到 Word 里",
            font=("PingFang SC", 20, "bold"),
        ).pack(anchor="w")

        ttk.Label(
            container,
            text="内容不会直接粘贴，而是会像键盘输入一样，一个字一个字打出来。",
            foreground="#4b5563",
        ).pack(anchor="w", pady=(6, 14))

        ttk.Label(
            container,
            text="输入过程中可在任意窗口按空格暂停/继续，按 Esc 停止。",
            foreground="#b91c1c",
        ).pack(anchor="w", pady=(0, 12))

        state_row = ttk.Frame(container)
        state_row.pack(fill="x", pady=(0, 12))

        ttk.Label(
            state_row,
            text="当前状态",
            font=("PingFang SC", 11, "bold"),
        ).pack(side="left")
        self.state_label = tk.Label(
            state_row,
            textvariable=self.state_var,
            font=("PingFang SC", 11, "bold"),
            bg="#6b7280",
            fg="white",
            padx=12,
            pady=6,
        )
        self.state_label.pack(side="left", padx=(10, 0))

        options_row = ttk.Frame(container)
        options_row.pack(fill="x", pady=(0, 12))

        ttk.Label(options_row, text="输入速度").pack(side="left")
        ttk.Scale(
            options_row,
            variable=self.speed_var,
            from_=0.01,
            to=0.30,
            orient="horizontal",
            command=self.on_speed_change,
        ).pack(side="left", fill="x", expand=True, padx=(10, 12))
        ttk.Label(options_row, textvariable=self.speed_text_var, width=12).pack(
            side="left"
        )

        human_row = ttk.Frame(container)
        human_row.pack(fill="x", pady=(0, 12))

        ttk.Checkbutton(
            human_row,
            text="模拟真人失误",
            variable=self.human_mode_var,
            command=self.update_estimate,
        ).pack(side="left")
        ttk.Label(human_row, text="真人风格").pack(side="left", padx=(14, 0))
        style_box = ttk.Combobox(
            human_row,
            textvariable=self.style_var,
            values=tuple(STYLE_CONFIGS.keys()),
            state="readonly",
            width=8,
        )
        style_box.pack(side="left", padx=(10, 12))
        style_box.bind("<<ComboboxSelected>>", lambda _event: self.update_estimate())
        ttk.Label(human_row, text="失误概率").pack(side="left", padx=(14, 0))
        ttk.Scale(
            human_row,
            variable=self.typo_rate_var,
            from_=0.0,
            to=0.18,
            orient="horizontal",
            command=self.on_typo_rate_change,
        ).pack(side="left", fill="x", expand=True, padx=(10, 12))
        ttk.Label(
            human_row,
            textvariable=self.typo_rate_text_var,
            width=8,
        ).pack(side="left")

        thought_row = ttk.Frame(container)
        thought_row.pack(fill="x", pady=(0, 12))

        ttk.Label(thought_row, text="思考停顿").pack(side="left")
        ttk.Scale(
            thought_row,
            variable=self.thought_var,
            from_=0.20,
            to=2.00,
            orient="horizontal",
            command=self.on_thought_change,
        ).pack(side="left", fill="x", expand=True, padx=(10, 12))
        ttk.Label(
            thought_row,
            textvariable=self.thought_text_var,
            width=8,
        ).pack(side="left")

        estimate_row = ttk.Frame(container)
        estimate_row.pack(fill="x", pady=(0, 12))

        ttk.Label(
            estimate_row,
            textvariable=self.estimate_var,
            foreground="#0f766e",
            font=("PingFang SC", 11, "bold"),
        ).pack(side="left")

        actions_frame = ttk.LabelFrame(container, text="开始逐字输入", padding=14)
        actions_frame.pack(fill="x", pady=(0, 12))

        primary_row = ttk.Frame(actions_frame)
        primary_row.pack(fill="x")

        self._add_primary_button(
            primary_row,
            "新建 Word 并逐字输入",
            self.type_into_new_word_document,
        ).pack(side="left", fill="x", expand=True)
        self._add_primary_button(
            primary_row,
            "3 秒后开始逐字输入",
            self.type_into_current_cursor,
        ).pack(side="left", fill="x", expand=True, padx=(12, 0))

        secondary_row = ttk.Frame(actions_frame)
        secondary_row.pack(fill="x", pady=(10, 0))

        self._add_button(secondary_row, "从头开始", self.restart_from_beginning).pack(
            side="left"
        )
        self._add_button(secondary_row, "从剪贴板读取", self.load_from_clipboard).pack(
            side="left", padx=(8, 0)
        )
        self._add_button(secondary_row, "清空内容", self.clear_text).pack(
            side="left", padx=8
        )

        self.text_input = tk.Text(
            container,
            wrap="word",
            font=("PingFang SC", 14),
            height=14,
            undo=True,
            padx=14,
            pady=14,
        )
        self.text_input.pack(fill="both", expand=True)
        self.text_input.bind("<<Modified>>", self.on_text_modified)

        ttk.Label(
            container,
            textvariable=self.status_var,
            foreground="#2563eb",
        ).pack(anchor="w", pady=(4, 0))

    def _add_button(self, parent: ttk.Frame, text: str, command) -> ttk.Button:
        button = ttk.Button(parent, text=text, command=command)
        self.action_buttons.append(button)
        return button

    def _add_primary_button(self, parent: ttk.Frame, text: str, command) -> tk.Button:
        button = tk.Button(
            parent,
            text=text,
            command=command,
            font=("PingFang SC", 14, "bold"),
            bg="#2563eb",
            fg="white",
            activebackground="#1d4ed8",
            activeforeground="white",
            relief="flat",
            padx=18,
            pady=14,
        )
        self.action_buttons.append(button)
        return button

    def get_text(self) -> str:
        return self.text_input.get("1.0", "end-1c")

    def ensure_text(self) -> str | None:
        text = self.get_text()
        if not text.strip():
            messagebox.showwarning("缺少内容", "请先输入或粘贴需要写入 Word 的文字。")
            return None
        return text

    def set_status(self, message: str) -> None:
        self.status_var.set(message)

    def set_state(self, state_text: str, color: str) -> None:
        self.state_var.set(state_text)
        self.state_color = color
        self.state_label.configure(bg=color)

    def on_speed_change(self, _value: str) -> None:
        self.speed_text_var.set(f"{self.speed_var.get():.2f} 秒/字")
        self.update_estimate()

    def on_thought_change(self, _value: str) -> None:
        self.thought_text_var.set(f"{self.thought_var.get():.2f}x")
        self.update_estimate()

    def on_typo_rate_change(self, _value: str) -> None:
        self.typo_rate_text_var.set(f"{round(self.typo_rate_var.get() * 100)}%")
        self.update_estimate()

    def on_text_modified(self, _event=None) -> None:
        if self.text_input.edit_modified():
            self.text_input.edit_modified(False)
            self.update_estimate()

    def get_key_delay(self) -> str:
        return f"{self.speed_var.get():.2f}"

    def get_typo_rate(self) -> float:
        return self.typo_rate_var.get()

    def get_style_config(self) -> dict[str, float]:
        return STYLE_CONFIGS.get(self.style_var.get(), STYLE_CONFIGS["结合型"])

    def get_thought_multiplier(self) -> float:
        return self.thought_var.get()

    def format_duration(self, seconds: float) -> str:
        total_seconds = max(0, int(round(seconds)))
        minutes, remaining_seconds = divmod(total_seconds, 60)
        if minutes == 0:
            return f"{remaining_seconds} 秒"
        return f"{minutes} 分 {remaining_seconds:02d} 秒"

    def estimate_typing_seconds(self, text: str, initial_delay: float) -> float:
        if not text:
            return 0.0

        base_delay = float(self.get_key_delay())
        style = self.get_style_config()
        thought = style["thought"] * self.get_thought_multiplier()
        slip = style["slip"] if self.human_mode_var.get() else 0.0
        typo_rate = self.get_typo_rate() if self.human_mode_var.get() else 0.0

        char_count = len(text)
        typo_candidate_count = sum(1 for char in text if self.is_typo_candidate(char))
        clause_break_count = sum(1 for char in text if char in CLAUSE_BREAK_CHARS)
        sentence_end_count = sum(1 for char in text if char in SENTENCE_END_CHARS)
        newline_count = text.count("\n") + text.count("\r")
        space_count = text.count(" ")

        total = initial_delay + char_count * base_delay
        if not self.human_mode_var.get():
            return total

        total *= 1.03 + 0.05 * thought
        total += clause_break_count * 0.18 * thought
        total += sentence_end_count * 0.42 * thought
        total += newline_count * 0.65 * thought
        total += space_count * 0.03
        total += max(0, char_count // 6) * 0.05 * style["burst_pause"]
        total += max(0, char_count // 20) * 0.07 * thought

        typo_events = typo_candidate_count * typo_rate * 0.75 * slip
        phrase_typo_events = typo_candidate_count * typo_rate * style["phrase_typo"] * 0.18
        total += typo_events * (0.28 + base_delay * 2.8)
        total += phrase_typo_events * (0.70 + base_delay * 5.2)
        return total

    def update_estimate(self) -> None:
        text = self.get_text().strip()
        if not text:
            self.estimate_var.set("预计时间：请先输入内容")
            return

        current_cursor_seconds = self.estimate_typing_seconds(text, 3.0)
        new_word_seconds = self.estimate_typing_seconds(text, 1.2)
        self.estimate_var.set(
            "预计时间：当前光标约 "
            f"{self.format_duration(current_cursor_seconds)}，"
            "新 Word 约 "
            f"{self.format_duration(new_word_seconds)}"
        )

    def reset_human_profile(self, text_length: int) -> None:
        style = self.get_style_config()
        thought = style["thought"] * self.get_thought_multiplier()
        self.session_speed_factor = random.uniform(0.92, 1.05)
        self.fatigue_bias = random.uniform(0.06, 0.18) * thought
        self.focus_wave_depth = random.uniform(0.02, 0.08) * thought
        self.focus_wave_cycles = random.uniform(1.0, 2.4)

        if text_length < 40:
            self.fatigue_bias *= 0.65
            self.focus_wave_depth *= 0.7

    def get_progress_ratio(self, index: int, total_length: int) -> float:
        if total_length <= 1:
            return 0.0
        return index / (total_length - 1)

    def get_progress_multiplier(self, progress: float) -> float:
        if not self.human_mode_var.get():
            return 1.0

        warmup = 1.18 - min(progress / 0.10, 1.0) * 0.18
        fatigue = 1.0 + max(progress - 0.58, 0.0) * self.fatigue_bias
        focus_wave = 1.0 + math.sin(progress * math.pi * self.focus_wave_cycles) * self.focus_wave_depth
        return max(0.72, min(1.45, self.session_speed_factor * warmup * fatigue * focus_wave))

    def get_effective_typo_rate(self, progress: float) -> float:
        style = self.get_style_config()
        rate = self.get_typo_rate()
        if not self.human_mode_var.get():
            return rate
        rate *= style["slip"]
        if progress < 0.12:
            rate *= 0.7
        if progress > 0.55:
            rate *= 1.0 + (progress - 0.55) * 0.7
        return min(rate, 0.28)

    def set_buttons_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        for button in self.action_buttons:
            button.configure(state=state)

    def start_hotkey_listener(self) -> None:
        self.hotkey_listener = keyboard.Listener(on_press=self.on_global_key_press)
        self.hotkey_listener.daemon = True
        self.hotkey_listener.start()

    def on_global_key_press(self, key: keyboard.Key | keyboard.KeyCode) -> None:
        with self.state_lock:
            is_active = self.typing_active
        if not is_active:
            return
        if key == keyboard.Key.space:
            now = time.monotonic()
            if now < self.ignore_space_until:
                return
            if now - self.last_space_press < 0.25:
                return
            self.last_space_press = now
            self.root.after(0, self.toggle_pause)
        elif key == keyboard.Key.esc:
            self.root.after(0, self.stop_typing)

    def ensure_not_stopped(self) -> None:
        with self.state_lock:
            should_stop = self.stop_requested
        if should_stop:
            raise UserCancelledError()

    def wait_if_paused(self) -> None:
        while True:
            self.ensure_not_stopped()
            with self.state_lock:
                is_paused = self.pause_requested
            if not is_paused:
                return
            time.sleep(0.05)

    def delay_with_controls(self, seconds: float) -> None:
        deadline = time.monotonic() + seconds
        while True:
            self.wait_if_paused()
            self.ensure_not_stopped()
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                return
            time.sleep(min(0.05, remaining))

    def run_applescript(
        self,
        lines: list[str],
        args: list[str] | None = None,
        timeout: int = 30,
    ) -> None:
        self.ensure_not_stopped()
        command = build_applescript_command(lines, args)
        process = subprocess.Popen(
            command,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        with self.process_lock:
            self.current_process = process

        try:
            _stdout, stderr = process.communicate(timeout=timeout)
        except subprocess.TimeoutExpired as exc:
            process.kill()
            process.communicate()
            raise RuntimeError("执行超时，请缩短内容或调快输入速度。") from exc
        finally:
            with self.process_lock:
                if self.current_process is process:
                    self.current_process = None

        if self.stop_requested or process.returncode in {-15, -9}:
            raise UserCancelledError()
        if process.returncode != 0:
            raise RuntimeError(stderr.strip() or "AppleScript 执行失败。")

    def stop_typing(self) -> None:
        with self.state_lock:
            if self.stop_requested:
                return
            self.stop_requested = True
            self.pause_requested = False
            is_active = self.typing_active
        if not is_active:
            return
        self.set_state("正在停止", "#dc2626")
        self.set_status("正在彻底停止输入...")
        with self.process_lock:
            process = self.current_process
        if process is not None and process.poll() is None:
            process.terminate()

    def toggle_pause(self) -> None:
        with self.state_lock:
            if not self.typing_active or self.stop_requested:
                return
            self.pause_requested = not self.pause_requested
            is_paused = self.pause_requested

        if is_paused:
            self.set_state("已暂停", "#d97706")
            self.set_status("已暂停。再按一次空格继续，按 Esc 停止。")
        else:
            self.set_state("输入中", "#2563eb")
            self.set_status("已继续输入。按空格可再次暂停，按 Esc 可停止。")

    def run_in_background(self, worker, busy_message: str) -> None:
        with self.state_lock:
            self.stop_requested = False
            self.pause_requested = False
            self.typing_active = True
        self.set_buttons_enabled(False)
        self.set_state("准备中", "#7c3aed")
        self.set_status(busy_message)

        def runner() -> None:
            try:
                worker()
            except UserCancelledError:
                self.root.after(0, self.on_cancelled)
                return
            except Exception as exc:  # noqa: BLE001
                self.root.after(0, self.on_error, str(exc))
                return
            self.root.after(0, self.on_success)

        Thread(target=runner, daemon=True).start()

    def on_success(self) -> None:
        with self.state_lock:
            self.typing_active = False
            self.pause_requested = False
        self.set_buttons_enabled(True)
        self.pending_restart_mode = None
        self.set_state("已完成", "#16a34a")

    def on_cancelled(self) -> None:
        with self.state_lock:
            self.typing_active = False
            self.pause_requested = False
        self.set_buttons_enabled(True)
        self.set_state("已停止", "#dc2626")
        self.set_status("输入已停止。")
        if self.pending_restart_mode is not None:
            restart_mode = self.pending_restart_mode
            self.pending_restart_mode = None
            self.root.after(200, lambda: self.begin_typing(restart_mode))

    def on_error(self, message: str) -> None:
        with self.state_lock:
            self.typing_active = False
            self.pause_requested = False
        self.set_buttons_enabled(True)
        self.pending_restart_mode = None
        self.set_state("出错", "#dc2626")
        self.set_status("执行失败，请检查 Word 和系统权限设置。")
        messagebox.showerror("执行失败", message)

    def load_from_clipboard(self) -> None:
        try:
            clipboard_text = self.root.clipboard_get()
        except tk.TclError:
            messagebox.showwarning("剪贴板为空", "当前剪贴板里没有可读取的文本。")
            return

        self.text_input.delete("1.0", "end")
        self.text_input.insert("1.0", clipboard_text)
        self.set_state("空闲", "#6b7280")
        self.set_status("已从剪贴板载入内容。")
        self.update_estimate()

    def clear_text(self) -> None:
        self.text_input.delete("1.0", "end")
        self.set_state("空闲", "#6b7280")
        self.set_status("内容已清空。")
        self.update_estimate()

    def begin_typing(self, mode: str) -> None:
        text = self.ensure_text()
        if text is None:
            return

        self.last_start_mode = mode
        if mode == "new_word":
            self._start_new_word_typing(text)
        else:
            self._start_current_cursor_typing(text)

    def restart_from_beginning(self) -> None:
        if self.last_start_mode is None:
            messagebox.showinfo("还不能重启", "请先执行一次输入，再使用“从头开始”。")
            return

        if self.typing_active:
            self.pending_restart_mode = self.last_start_mode
            self.stop_typing()
            self.set_status("正在停止，并准备从头开始...")
            return

        self.begin_typing(self.last_start_mode)

    def type_character(self, char: str) -> None:
        if char == " ":
            self.ignore_space_until = time.monotonic() + 0.35
        self.run_applescript(
            [
                "on run argv",
                "set typedChar to item 1 of argv",
                'tell application "System Events"',
                "if typedChar is return or typedChar is linefeed then",
                "key code 36",
                "else if typedChar is tab then",
                "key code 48",
                "else",
                "keystroke typedChar",
                "end if",
                "end tell",
                "end run",
            ],
            args=[char],
            timeout=10,
        )

    def press_backspace(self) -> None:
        self.run_applescript(
            [
                'tell application "System Events"',
                "key code 51",
                "end tell",
            ],
            timeout=10,
        )

    def is_cjk_character(self, char: str) -> bool:
        return "\u4e00" <= char <= "\u9fff"

    def is_word_character(self, char: str) -> bool:
        return char.isascii() and (char.isalpha() or char.isdigit())

    def is_typo_candidate(self, char: str) -> bool:
        return bool(char.strip()) and char not in {"\n", "\r", "\t"}

    def should_simulate_typo(self, char: str, progress: float) -> bool:
        if not self.human_mode_var.get():
            return False
        if not self.is_typo_candidate(char):
            return False
        return random.random() < self.get_effective_typo_rate(progress)

    def choose_wrong_character(self, expected_char: str) -> str:
        if expected_char.isalpha() and expected_char.isascii():
            pool = ASCII_TYPO_CHARS.upper() if expected_char.isupper() else ASCII_TYPO_CHARS
        elif expected_char.isdigit():
            pool = DIGIT_TYPO_CHARS
        elif self.is_cjk_character(expected_char):
            pool = COMMON_CJK_TYPO_CHARS
        elif expected_char in PUNCTUATION_TYPO_CHARS:
            pool = PUNCTUATION_TYPO_CHARS
        else:
            pool = ASCII_TYPO_CHARS

        choices = [char for char in pool if char != expected_char]
        if not choices:
            return expected_char
        return random.choice(choices)

    def get_typo_sequence(self, expected_char: str) -> list[str]:
        typo_count = 1 if random.random() < 0.75 else 2
        return [self.choose_wrong_character(expected_char) for _ in range(typo_count)]

    def get_typo_segment(self, text: str, start_index: int) -> str:
        segment_chars = [text[start_index]]
        expected_class = self.get_character_class(text[start_index])
        max_length = 3 if expected_class == "cjk" else 4

        for index in range(start_index + 1, min(len(text), start_index + max_length)):
            current_char = text[index]
            if not self.is_typo_candidate(current_char):
                break
            if self.get_character_class(current_char) != expected_class:
                break
            segment_chars.append(current_char)
        return "".join(segment_chars)

    def get_character_class(self, char: str) -> str:
        if self.is_cjk_character(char):
            return "cjk"
        if char.isascii() and char.isalpha():
            return "alpha"
        if char.isdigit():
            return "digit"
        if char in PUNCTUATION_TYPO_CHARS:
            return "punctuation"
        return "other"

    def should_simulate_phrase_typo(self, text: str, start_index: int, progress: float) -> bool:
        if not self.human_mode_var.get():
            return False
        style = self.get_style_config()
        segment = self.get_typo_segment(text, start_index)
        if len(segment) < 2:
            return False
        if self.get_character_class(segment[0]) == "punctuation":
            return False
        return random.random() < self.get_effective_typo_rate(progress) * style["phrase_typo"]

    def get_humanized_delay(self, base_delay: float, char: str) -> float:
        delay = base_delay
        if self.human_mode_var.get():
            delay *= random.uniform(0.65, 1.45)
            if char in "，。！？；：,.!?;:":
                delay += random.uniform(0.12, 0.28)
        return max(0.01, delay)

    def sample_burst_profile(self) -> tuple[int, float]:
        style = self.get_style_config()
        if self.style_var.get() == "思考型":
            return random.randint(2, 6), random.uniform(0.88, 1.18)
        if self.style_var.get() == "手滑型":
            return random.randint(4, 9), random.uniform(0.74, 1.02)
        if self.style_var.get() == "结合型":
            return random.randint(3, 8), random.uniform(0.76, 1.12)
        return random.randint(4, 9), random.uniform(0.82, 1.05)

    def get_pre_character_delay(
        self,
        current_char: str,
        previous_char: str,
        streak: int,
        progress_multiplier: float,
    ) -> float:
        if not self.human_mode_var.get():
            return 0.0

        style = self.get_style_config()
        delay = 0.0
        if previous_char in {"\n", "\r"}:
            delay += random.uniform(0.20, 0.45) * style["thought"]
        elif previous_char in SENTENCE_END_CHARS:
            delay += random.uniform(0.14, 0.30) * style["thought"]
        elif previous_char in CLAUSE_BREAK_CHARS:
            delay += random.uniform(0.05, 0.14) * style["thought"]

        if current_char in CLOSING_CHARS:
            delay *= 0.6

        if streak > 0 and streak % random.randint(14, 22) == 0:
            delay += random.uniform(0.08, 0.22) * style["thought"]

        if current_char.strip() and random.random() < 0.03 * style["thought"]:
            delay += random.uniform(0.06, 0.18) * style["thought"]

        if previous_char == " " and self.is_word_character(current_char):
            delay += random.uniform(0.03, 0.10) * style["thought"]
        if previous_char in OPENING_CHARS:
            delay *= 0.7

        return delay * progress_multiplier

    def get_post_character_delay(
        self,
        base_delay: float,
        current_char: str,
        next_char: str,
        burst_factor: float,
        streak: int,
        progress_multiplier: float,
    ) -> float:
        style = self.get_style_config()
        delay = self.get_humanized_delay(base_delay, current_char)
        if not self.human_mode_var.get():
            return delay

        delay *= burst_factor * progress_multiplier

        if self.is_word_character(current_char) and self.is_word_character(next_char):
            delay *= random.uniform(0.72, 0.92)
        elif self.is_cjk_character(current_char) and self.is_cjk_character(next_char):
            delay *= random.uniform(0.80, 0.96)

        if current_char == " ":
            delay += random.uniform(0.02, 0.08)
        if current_char in CLAUSE_BREAK_CHARS:
            delay += random.uniform(0.12, 0.28) * style["thought"]
        if current_char in SENTENCE_END_CHARS:
            delay += random.uniform(0.28, 0.62) * style["thought"]
        if current_char in {"\n", "\r"}:
            delay += random.uniform(0.35, 0.80) * style["thought"]
        if current_char in OPENING_CHARS:
            delay *= 0.75
        if next_char in CLOSING_CHARS:
            delay *= 0.82
        if streak > 0 and streak % random.randint(9, 15) == 0:
            delay += random.uniform(0.04, 0.14) * style["burst_pause"]

        return max(0.01, delay)

    def simulate_typo_and_fix(
        self,
        expected_char: str,
        base_delay: float,
        progress_multiplier: float,
    ) -> None:
        style = self.get_style_config()
        typo_chars = self.get_typo_sequence(expected_char)

        for wrong_char in typo_chars:
            self.wait_if_paused()
            self.ensure_not_stopped()
            self.type_character(wrong_char)
            self.delay_with_controls(
                self.get_humanized_delay(base_delay, wrong_char) * progress_multiplier
            )

        if style["slip"] > 1.0 and random.random() < 0.28:
            self.type_character(expected_char)
            self.delay_with_controls(random.uniform(0.04, 0.08) * progress_multiplier)
            self.press_backspace()
            self.delay_with_controls(random.uniform(0.04, 0.08) * progress_multiplier)

        self.delay_with_controls(
            random.uniform(0.08, 0.18) * progress_multiplier * style["slip"]
        )

        for _wrong_char in reversed(typo_chars):
            self.wait_if_paused()
            self.ensure_not_stopped()
            self.press_backspace()
            self.delay_with_controls(
                random.uniform(0.05, 0.12) * progress_multiplier * style["slip"]
            )

        self.delay_with_controls(
            random.uniform(0.05, 0.12) * progress_multiplier * style["thought"]
        )

    def simulate_phrase_typo_and_fix(
        self,
        expected_segment: str,
        base_delay: float,
        progress_multiplier: float,
    ) -> None:
        style = self.get_style_config()
        wrong_segment = "".join(
            self.choose_wrong_character(char) for char in expected_segment
        )

        for wrong_char in wrong_segment:
            self.wait_if_paused()
            self.ensure_not_stopped()
            self.type_character(wrong_char)
            self.delay_with_controls(
                self.get_humanized_delay(base_delay, wrong_char)
                * progress_multiplier
                * random.uniform(0.90, 1.18)
                * style["slip"]
            )

        if style["slip"] > 1.05 and random.random() < 0.24:
            self.delay_with_controls(random.uniform(0.10, 0.18) * progress_multiplier)

        self.delay_with_controls(
            random.uniform(0.18, 0.38) * progress_multiplier * style["thought"]
        )

        for _wrong_char in reversed(wrong_segment):
            self.wait_if_paused()
            self.ensure_not_stopped()
            self.press_backspace()
            self.delay_with_controls(
                random.uniform(0.04, 0.09) * progress_multiplier * style["slip"]
            )

        self.delay_with_controls(
            random.uniform(0.08, 0.18) * progress_multiplier * style["thought"]
        )

    def type_text(self, text: str, initial_delay: str) -> None:
        key_delay = float(self.get_key_delay())
        self.reset_human_profile(len(text))
        self.root.after(0, self.set_state, "等待开始", "#7c3aed")
        self.delay_with_controls(float(initial_delay))
        self.root.after(0, self.set_state, "输入中", "#2563eb")
        self.root.after(
            0,
            self.set_status,
            "正在输入。按空格暂停/继续，按 Esc 停止。",
        )
        burst_remaining, burst_factor = self.sample_burst_profile()
        streak = 0

        for index, char in enumerate(text):
            self.wait_if_paused()
            self.ensure_not_stopped()
            previous_char = text[index - 1] if index > 0 else ""
            next_char = text[index + 1] if index < len(text) - 1 else ""
            progress = self.get_progress_ratio(index, len(text))
            progress_multiplier = self.get_progress_multiplier(progress)

            if self.human_mode_var.get():
                if burst_remaining <= 0:
                    self.delay_with_controls(
                        random.uniform(0.05, 0.16) * progress_multiplier
                    )
                    burst_remaining, burst_factor = self.sample_burst_profile()

                pre_delay = self.get_pre_character_delay(
                    char,
                    previous_char,
                    streak,
                    progress_multiplier,
                )
                if pre_delay > 0:
                    self.delay_with_controls(pre_delay)

            if self.should_simulate_phrase_typo(text, index, progress):
                self.simulate_phrase_typo_and_fix(
                    self.get_typo_segment(text, index),
                    key_delay,
                    progress_multiplier,
                )
            elif self.should_simulate_typo(char, progress):
                self.simulate_typo_and_fix(char, key_delay, progress_multiplier)

            self.type_character(char)
            burst_remaining -= 1
            streak = streak + 1 if char not in {"\n", "\r"} else 0

            if index < len(text) - 1 and key_delay > 0:
                self.delay_with_controls(
                    self.get_post_character_delay(
                        key_delay,
                        char,
                        next_char,
                        burst_factor,
                        streak,
                        progress_multiplier,
                    )
                )

    def _start_new_word_typing(self, text: str) -> None:
        def worker() -> None:
            self.run_applescript(
                [
                    'tell application "Microsoft Word"',
                    "activate",
                    "make new document",
                    "end tell",
                    'tell application "System Events"',
                    'tell process "Microsoft Word"',
                    "set frontmost to true",
                    "end tell",
                ]
            )
            self.ensure_not_stopped()
            self.type_text(text, "1.2")
            self.root.after(
                0, self.set_status, "已在新的 Word 文档中逐字输入完成。"
            )

        self.run_in_background(worker, "正在打开 Word，并准备逐字输入...")

    def _start_current_cursor_typing(self, text: str) -> None:
        def worker() -> None:
            self.type_text(text, "3")
            self.root.after(
                0, self.set_status, "已按当前光标位置逐字输入完成。"
            )

        self.run_in_background(
            worker,
            "请在 3 秒内切到 Word，并把光标放到目标位置。空格暂停/继续，Esc 彻底停止。",
        )

    def type_into_new_word_document(self) -> None:
        self.begin_typing("new_word")

    def type_into_current_cursor(self) -> None:
        self.begin_typing("current_cursor")

    def on_close(self) -> None:
        with self.state_lock:
            self.stop_requested = True
            self.pause_requested = False
        with self.process_lock:
            process = self.current_process
        if process is not None and process.poll() is None:
            process.terminate()
        if self.hotkey_listener is not None:
            self.hotkey_listener.stop()
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    style = ttk.Style(root)
    if "clam" in style.theme_names():
        style.theme_use("clam")

    app = WordTyperApp(root)
    icon_path = Path(__file__).with_name("icon.ico")
    if icon_path.exists():
        try:
            root.iconbitmap(default=str(icon_path))
        except tk.TclError:
            pass

    app.text_input.focus_set()
    root.mainloop()


if __name__ == "__main__":
    main()
