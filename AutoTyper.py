"""
Auto Typer
==========
Types your text into any window, with adjustable speed and optional
human-like typos that get backspaced and corrected.

Just run this file (or double-click run_windows.bat). The one extra library it
needs (pynput) is installed automatically the first time.

No API, no keys, no internet needed after that.
"""

import random
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


# ----------------------------------------------------------------------------
# Auto-install the keyboard library so the app "just works"
# ----------------------------------------------------------------------------
def _pip_install(package):
    subprocess.check_call(
        [sys.executable, "-m", "pip", "install", "--quiet", package],
        stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )


def load_pynput():
    def _import():
        from pynput.keyboard import Controller, Key, Listener
        from pynput.mouse import Button, Controller as MouseController
        return Controller, Key, Listener, MouseController, Button

    try:
        return _import()
    except ImportError:
        pass
    try:
        _pip_install("pynput")
        return _import()
    except Exception:
        return None


def load_docx():
    """Loads python-docx (used to read Word files), installing it if needed."""
    try:
        import docx
        return docx
    except ImportError:
        pass
    try:
        _pip_install("python-docx")
        import docx
        return docx
    except Exception:
        return None


# ----------------------------------------------------------------------------
# Look and feel
# ----------------------------------------------------------------------------
BG = "#12141a"
PANEL = "#1b1e27"
PANEL_2 = "#252937"
FG = "#e9ebf2"
MUTED = "#8a90a3"
ACCENT = "#6c8cff"
ACCENT_HOVER = "#8aa3ff"
DANGER = "#ff5c6c"
DANGER_HOVER = "#ff7d8a"
OK = "#4ade80"

if sys.platform.startswith("win"):
    FONT = "Segoe UI"
elif sys.platform == "darwin":
    FONT = "Helvetica Neue"
else:
    FONT = "DejaVu Sans"

PLACEHOLDER = "Paste or type the text you want auto-typed here..."

# Neighbouring keys on a QWERTY keyboard, used to make realistic typos
NEIGHBORS = {
    "q": "wa", "w": "qes", "e": "wrd", "r": "etf", "t": "ryg", "y": "tuh",
    "u": "yij", "i": "uok", "o": "ipl", "p": "ol",
    "a": "qsz", "s": "awdx", "d": "sefc", "f": "drgv", "g": "fthb",
    "h": "gyjn", "j": "hukm", "k": "jilm", "l": "kop",
    "z": "asx", "x": "zsc", "c": "xdv", "v": "cfb", "b": "vgn",
    "n": "bhm", "m": "njk",
}


class FlatButton(tk.Label):
    """A label-based button so colours look the same on Windows, Mac and Linux."""

    def __init__(self, parent, text, command, bg, hover, fg="#ffffff",
                 size=10, bold=True, padx=14, pady=8):
        super().__init__(
            parent, text=text, bg=bg, fg=fg, cursor="hand2", padx=padx, pady=pady,
            font=(FONT, size, "bold" if bold else "normal"),
        )
        self.command = command
        self.base = bg
        self.hover = hover
        self.bind("<Button-1>", lambda e: self.command())
        self.bind("<Enter>", lambda e: self.config(bg=self.hover))
        self.bind("<Leave>", lambda e: self.config(bg=self.base))

    def restyle(self, text, bg, hover):
        self.base, self.hover = bg, hover
        self.config(text=text, bg=bg)


class Toggle(tk.Canvas):
    """An on/off switch."""

    def __init__(self, parent, value, command, bg=PANEL):
        super().__init__(parent, width=46, height=26, bg=bg, highlightthickness=0, cursor="hand2")
        self.value = value
        self.command = command
        self.bind("<Button-1>", self._click)
        self._draw()

    def _draw(self):
        self.delete("all")
        color = ACCENT if self.value else "#3a3f52"
        self.create_oval(2, 2, 26, 24, fill=color, outline=color)
        self.create_oval(20, 2, 44, 24, fill=color, outline=color)
        self.create_rectangle(14, 2, 32, 24, fill=color, outline=color)
        x = 33 if self.value else 13
        self.create_oval(x - 9, 4, x + 9, 22, fill="#ffffff", outline="#ffffff")

    def _click(self, _):
        self.value = not self.value
        self._draw()
        self.command(self.value)


class AutoTyperApp:
    def __init__(self, root, pynput_parts):
        Controller, self.Key, Listener, MouseController, self.MouseButton = pynput_parts
        self.keyboard = Controller()
        self.mouse = MouseController()
        self.click_pos = None
        self.click_enabled = False
        self.root = root

        self.wpm_val = 65
        self.err_val = 6
        self.errors_enabled = True
        self.stop_event = threading.Event()
        self.running = False
        self.placeholder_active = False

        root.title("Auto Typer")
        root.configure(bg=BG)
        root.geometry("640x900")
        root.minsize(560, 840)

        self._setup_styles()
        self._build_ui()
        self._show_placeholder()
        self._update_stats()

        self.listener = Listener(on_press=self._on_global_key)
        self.listener.daemon = True
        self.listener.start()
        root.protocol("WM_DELETE_WINDOW", self._close)

    # ------------------------------------------------------------------ UI
    def _setup_styles(self):
        s = ttk.Style()
        s.theme_use("clam")
        s.configure(
            "Horizontal.TScale", background=ACCENT, troughcolor=PANEL_2,
            lightcolor=ACCENT, darkcolor=ACCENT, bordercolor=PANEL, focuscolor=PANEL,
        )
        s.configure(
            "Horizontal.TProgressbar", background=ACCENT, troughcolor=PANEL_2,
            bordercolor=PANEL_2, lightcolor=ACCENT, darkcolor=ACCENT, thickness=8,
        )
        s.configure("TSpinbox", fieldbackground=PANEL_2, background=PANEL_2,
                    foreground=FG, bordercolor=PANEL_2, arrowcolor=FG)

    def _card(self, parent, **pack):
        card = tk.Frame(parent, bg=PANEL, padx=16, pady=14)
        card.pack(fill="x", pady=(0, 12), **pack)
        return card

    def _label(self, parent, text, size=10, color=FG, bold=False, bg=PANEL):
        return tk.Label(parent, text=text, bg=bg, fg=color,
                        font=(FONT, size, "bold" if bold else "normal"))

    def _build_ui(self):
        outer = tk.Frame(self.root, bg=BG, padx=20, pady=18)
        outer.pack(fill="both", expand=True)

        # Header
        self._label(outer, "Auto Typer", 20, FG, True, BG).pack(anchor="w")
        self._label(
            outer,
            "1. Paste your text    2. Set speed    3. Press Start, then click where you want it typed",
            9, MUTED, False, BG,
        ).pack(anchor="w", pady=(2, 14))

        # Text card
        card = self._card(outer)
        top = tk.Frame(card, bg=PANEL)
        top.pack(fill="x", pady=(0, 8))
        self._label(top, "Text to type", 11, FG, True).pack(side="left")
        FlatButton(top, "Clear", self._clear_text, PANEL_2, "#30354a", FG, 9, False, 10, 4).pack(side="right")
        FlatButton(top, "Paste", self._paste_clipboard, PANEL_2, "#30354a", FG, 9, False, 10, 4).pack(
            side="right", padx=(0, 6))
        FlatButton(top, "Open Word doc", self._open_docx, PANEL_2, "#30354a", FG, 9, False, 10, 4).pack(
            side="right", padx=(0, 6))

        box = tk.Frame(card, bg=PANEL_2)
        box.pack(fill="both", expand=True)
        self.text = tk.Text(
            box, wrap="word", height=9, bg=PANEL_2, fg=FG, insertbackground=FG,
            relief="flat", padx=10, pady=10, font=(FONT, 10), undo=True,
            highlightthickness=0, selectbackground=ACCENT,
        )
        scroll = ttk.Scrollbar(box, command=self.text.yview)
        self.text.configure(yscrollcommand=scroll.set)
        self.text.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.text.bind("<FocusIn>", self._hide_placeholder)
        self.text.bind("<FocusOut>", lambda e: self._show_placeholder())
        self.text.bind("<<Modified>>", self._on_modified)

        info = tk.Frame(card, bg=PANEL)
        info.pack(fill="x", pady=(8, 0))
        self.stats_label = self._label(info, "", 9, MUTED)
        self.stats_label.pack(side="left")
        self.eta_label = self._label(info, "", 9, MUTED)
        self.eta_label.pack(side="right")

        # Speed card
        card = self._card(outer)
        row = tk.Frame(card, bg=PANEL)
        row.pack(fill="x")
        self._label(row, "Typing speed", 11, FG, True).pack(side="left")
        self.wpm_label = self._label(row, "", 11, ACCENT, True)
        self.wpm_label.pack(side="right")
        self.wpm_scale = ttk.Scale(card, from_=10, to=300, orient="horizontal",
                                   command=self._on_wpm)
        self.wpm_scale.set(self.wpm_val)
        self.wpm_scale.pack(fill="x", pady=(10, 8))

        presets = tk.Frame(card, bg=PANEL)
        presets.pack(fill="x")
        for name, value in (("Slow", 35), ("Normal", 65), ("Fast", 110), ("Turbo", 180)):
            FlatButton(presets, f"{name} {value}", lambda v=value: self.wpm_scale.set(v),
                       PANEL_2, "#30354a", FG, 9, False, 10, 4).pack(side="left", padx=(0, 6))

        # Errors card
        card = self._card(outer)
        row = tk.Frame(card, bg=PANEL)
        row.pack(fill="x")
        left = tk.Frame(row, bg=PANEL)
        left.pack(side="left")
        self._label(left, "Add typing mistakes", 11, FG, True).pack(anchor="w")
        self._label(left, "Wrong keys get typed, then backspaced and fixed", 9, MUTED).pack(anchor="w")
        Toggle(row, True, self._on_toggle_errors).pack(side="right")

        row = tk.Frame(card, bg=PANEL)
        row.pack(fill="x", pady=(12, 0))
        self._label(row, "Mistake amount", 10, FG).pack(side="left")
        self.err_label = self._label(row, "", 10, ACCENT, True)
        self.err_label.pack(side="right")
        self.err_scale = ttk.Scale(card, from_=0, to=30, orient="horizontal", command=self._on_err)
        self.err_scale.set(self.err_val)
        self.err_scale.pack(fill="x", pady=(8, 0))

        # Options card
        card = self._card(outer)
        row = tk.Frame(card, bg=PANEL)
        row.pack(fill="x")
        self._label(row, "Start delay", 10, FG).pack(side="left")
        self._label(row, "seconds", 9, MUTED).pack(side="right")
        self.delay_var = tk.IntVar(value=5)
        ttk.Spinbox(row, from_=1, to=60, width=4, textvariable=self.delay_var).pack(side="right", padx=8)

        row = tk.Frame(card, bg=PANEL)
        row.pack(fill="x", pady=(12, 0))
        self._label(row, "Keep this window on top", 10, FG).pack(side="left")
        Toggle(row, True, lambda v: self.root.attributes("-topmost", v)).pack(side="right")
        self.root.attributes("-topmost", True)

        # Mouse card
        card = self._card(outer)
        row = tk.Frame(card, bg=PANEL)
        row.pack(fill="x")
        left = tk.Frame(row, bg=PANEL)
        left.pack(side="left")
        self._label(left, "Click a spot before typing", 11, FG, True).pack(anchor="w")
        self._label(left, "Moves the mouse there and clicks, then starts typing", 9, MUTED).pack(anchor="w")
        Toggle(row, False, self._on_toggle_click).pack(side="right")
        spot_row = tk.Frame(card, bg=PANEL)
        spot_row.pack(fill="x", pady=(10, 0))
        FlatButton(spot_row, "Set click spot (3 second countdown)", self._pick_spot,
                   PANEL_2, "#30354a", FG, 9, False, 10, 4).pack(side="left")
        self.spot_label = self._label(spot_row, "No spot set", 9, MUTED)
        self.spot_label.pack(side="left", padx=10)

        # Start / stop
        self.main_btn = FlatButton(outer, "Start typing", self._toggle_run, ACCENT, ACCENT_HOVER,
                                   size=13, pady=12)
        self.main_btn.pack(fill="x", pady=(2, 10))

        self.progress = ttk.Progressbar(outer, maximum=100, mode="determinate")
        self.progress.pack(fill="x")
        self.status = self._label(outer, "Ready. Press Esc at any time to stop.", 10, MUTED, False, BG)
        self.status.pack(anchor="w", pady=(8, 0))

        self._refresh_labels()

    # ------------------------------------------------------ placeholder text
    def _show_placeholder(self):
        if not self.text.get("1.0", "end-1c").strip() and not self.placeholder_active:
            self.placeholder_active = True
            self.text.insert("1.0", PLACEHOLDER)
            self.text.config(fg=MUTED)
            self.text.edit_modified(False)

    def _hide_placeholder(self, _=None):
        if self.placeholder_active:
            self.placeholder_active = False
            self.text.delete("1.0", "end")
            self.text.config(fg=FG)
            self.text.edit_modified(False)

    def _get_content(self):
        return "" if self.placeholder_active else self.text.get("1.0", "end-1c")

    def _on_modified(self, _=None):
        if self.text.edit_modified():
            self._update_stats()
            self.text.edit_modified(False)

    def _paste_clipboard(self):
        try:
            clip = self.root.clipboard_get()
        except tk.TclError:
            self._set_status("Clipboard is empty.", MUTED)
            return
        self._hide_placeholder()
        self.text.delete("1.0", "end")
        self.text.insert("1.0", clip)
        self._update_stats()

    def _open_docx(self):
        path = filedialog.askopenfilename(
            title="Choose a Word document", filetypes=[("Word documents", "*.docx")])
        if not path:
            return
        docx = load_docx()
        if docx is None:
            messagebox.showerror(
                "Auto Typer",
                "Couldn't install the Word reader automatically.\n\n"
                "Open a terminal and run:\n    pip install python-docx")
            return
        try:
            from docx.table import Table
            from docx.text.paragraph import Paragraph
            document = docx.Document(path)
            lines = []
            for child in document.element.body.iterchildren():
                if child.tag.endswith("}p"):
                    lines.append(Paragraph(child, document).text)
                elif child.tag.endswith("}tbl"):
                    for row in Table(child, document).rows:
                        lines.append("\t".join(cell.text for cell in row.cells))
        except Exception as err:
            messagebox.showerror("Auto Typer", f"Couldn't open that file.\n\n{err}")
            return
        content = "\n".join(lines).strip()
        self._hide_placeholder()
        self.text.delete("1.0", "end")
        self.text.insert("1.0", content)
        if not content:
            self._show_placeholder()
        self._update_stats()
        name = path.replace("\\", "/").split("/")[-1]
        self._set_status(f"Loaded {name}", OK)

    # ---------------------------------------------------------- mouse spot
    def _on_toggle_click(self, on):
        self.click_enabled = on
        if on and self.click_pos is None:
            self._set_status("Now press 'Set click spot' to choose where to click.", MUTED)

    def _pick_spot(self, seconds=3):
        if seconds > 0:
            self.spot_label.config(text=f"Move your mouse to the spot...  {seconds}", fg=ACCENT)
            self.root.after(1000, lambda: self._pick_spot(seconds - 1))
        else:
            self.click_pos = self.mouse.position
            x, y = self.click_pos
            self.spot_label.config(text=f"Saved: {int(x)}, {int(y)}", fg=OK)

    def _move_and_click(self):
        """Glide the mouse to the saved spot and click. Returns True if stopped."""
        tx, ty = self.click_pos
        sx, sy = self.mouse.position
        steps = 25
        for step in range(1, steps + 1):
            t = step / steps
            t = t * t * (3 - 2 * t)  # ease in and out
            self.mouse.position = (sx + (tx - sx) * t, sy + (ty - sy) * t)
            if self._pause(0.01):
                return True
        if self._pause(0.15):
            return True
        self.mouse.click(self.MouseButton.left, 1)
        return self._pause(0.3)

    def _clear_text(self):
        self._hide_placeholder()
        self.text.delete("1.0", "end")
        self._show_placeholder()
        self._update_stats()

    # ------------------------------------------------------------ labels
    def _on_wpm(self, value):
        self.wpm_val = int(round(float(value)))
        self._refresh_labels()

    def _on_err(self, value):
        self.err_val = int(round(float(value)))
        self._refresh_labels()

    def _on_toggle_errors(self, on):
        self.errors_enabled = on
        self.err_scale.state(["!disabled"] if on else ["disabled"])
        self._refresh_labels()

    def _refresh_labels(self):
        self.wpm_label.config(text=f"{self.wpm_val} WPM")
        if self.errors_enabled:
            self.err_label.config(text=f"{self.err_val}%", fg=ACCENT)
        else:
            self.err_label.config(text="Off", fg=MUTED)
        self._update_stats()

    def _update_stats(self):
        content = self._get_content()
        chars = len(content)
        words = len(content.split())
        self.stats_label.config(text=f"{chars:,} characters  |  {words:,} words")
        if chars:
            seconds = chars * 60.0 / (self.wpm_val * 5)
            if self.errors_enabled:
                seconds *= 1 + self.err_val / 100.0 * 3
            m, s = divmod(int(seconds), 60)
            self.eta_label.config(text=f"About {m}m {s:02d}s" if m else f"About {s}s")
        else:
            self.eta_label.config(text="")

    def _set_status(self, msg, color=MUTED):
        self.root.after(0, lambda: self.status.config(text=msg, fg=color))

    def _set_progress(self, pct):
        self.root.after(0, lambda: self.progress.configure(value=pct))

    # ------------------------------------------------------------ control
    def _on_global_key(self, key):
        if key == self.Key.esc and self.running:
            self.stop_event.set()

    def _toggle_run(self):
        if self.running:
            self.stop_event.set()
        else:
            self._start()

    def _start(self):
        content = self._get_content()
        if not content.strip():
            self._set_status("Paste some text first.", DANGER)
            return
        if self.click_enabled and self.click_pos is None:
            self._set_status("Set a click spot first, or turn the click option off.", DANGER)
            return
        self.stop_event.clear()
        self.running = True
        self.main_btn.restyle("Stop  (Esc)", DANGER, DANGER_HOVER)
        self.progress.configure(value=0)
        try:
            delay = max(1, int(self.delay_var.get()))
        except (tk.TclError, ValueError):
            delay = 5
        threading.Thread(target=self._run, args=(content, delay), daemon=True).start()

    def _finish(self, msg, color):
        def done():
            self.running = False
            self.main_btn.restyle("Start typing", ACCENT, ACCENT_HOVER)
            self.status.config(text=msg, fg=color)
        self.root.after(0, done)

    def _close(self):
        self.stop_event.set()
        self.listener.stop()
        self.root.destroy()

    # ------------------------------------------------------ typing engine
    def _pause(self, seconds):
        """Sleep, but wake instantly if Stop is pressed. Returns True if stopped."""
        return self.stop_event.wait(seconds)

    def _char_delay(self):
        base = 60.0 / (max(1, self.wpm_val) * 5)  # 5 characters = 1 word
        return base * random.uniform(0.7, 1.3)

    def _press(self, ch):
        if ch == "\n":
            self.keyboard.press(self.Key.enter)
            self.keyboard.release(self.Key.enter)
        elif ch == "\t":
            self.keyboard.press(self.Key.tab)
            self.keyboard.release(self.Key.tab)
        else:
            self.keyboard.type(ch)

    def _backspace(self, n):
        for _ in range(n):
            self.keyboard.press(self.Key.backspace)
            self.keyboard.release(self.Key.backspace)
            if self._pause(self._char_delay() * 0.6):
                return True
        return False

    @staticmethod
    def _wrong_char(ch):
        options = NEIGHBORS.get(ch.lower())
        if not options:
            return None
        wrong = random.choice(options)
        return wrong.upper() if ch.isupper() else wrong

    def _run(self, content, delay):
        stopped = ("Stopped.", DANGER)

        for remaining in range(delay, 0, -1):
            self._set_status(f"Click into the window you want to type in...  starting in {remaining}", ACCENT)
            if self._pause(1):
                return self._finish(*stopped)

        if self.click_enabled and self.click_pos is not None:
            self._set_status("Moving the mouse and clicking...", ACCENT)
            if self._move_and_click():
                return self._finish(*stopped)

        i, total = 0, len(content)
        while i < total:
            if self.stop_event.is_set():
                return self._finish(*stopped)

            ch = content[i]
            self._set_progress(i * 100 / total)
            self._set_status(f"Typing...  {i * 100 // total}%   (Esc to stop)", ACCENT)

            make_error = (
                self.errors_enabled and ch.isalpha()
                and random.random() < self.err_val / 100.0
            )
            wrong = self._wrong_char(ch) if make_error else None

            if wrong:
                # Type a wrong key, maybe carry on for 0-2 more characters before
                # "noticing", then backspace and retype it correctly.
                lag = min(random.choice([0, 0, 1, 2]), total - i - 1)
                self._press(wrong)
                if self._pause(self._char_delay()):
                    return self._finish(*stopped)
                for extra in content[i + 1: i + 1 + lag]:
                    self._press(extra)
                    if self._pause(self._char_delay()):
                        return self._finish(*stopped)
                if self._pause(random.uniform(0.15, 0.4)):  # the "oops" moment
                    return self._finish(*stopped)
                if self._backspace(lag + 1):
                    return self._finish(*stopped)
                for correct in content[i: i + 1 + lag]:
                    self._press(correct)
                    if self._pause(self._char_delay()):
                        return self._finish(*stopped)
                i += 1 + lag
            else:
                self._press(ch)
                pause = self._char_delay()
                if ch in ".,!?;:\n":
                    pause *= 1.5  # slight natural pause after punctuation
                if self._pause(pause):
                    return self._finish(*stopped)
                i += 1

        self._set_progress(100)
        self._finish("Done.", OK)


def main():
    root = tk.Tk()
    parts = load_pynput()
    if parts is None:
        root.withdraw()
        messagebox.showerror(
            "Auto Typer",
            "Couldn't install the keyboard library automatically.\n\n"
            "Open a terminal and run:\n    pip install pynput\n\nThen start Auto Typer again.",
        )
        return
    AutoTyperApp(root, parts)
    root.mainloop()


if __name__ == "__main__":
    main()
