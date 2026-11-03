"""Capture the real Tk UI into docs/screens/ for Q5.

Uses a short-lived ready hook in app.gui. No AI/cloud call and no Excel write:
the preview dialog is opened, captured, then cancelled. Screenshots are visual
evidence only; a human still has to confirm readability and layout.
"""
import ctypes
import sys
from pathlib import Path
from unittest.mock import patch

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    pass

from PIL import ImageGrab
import tkinter as tk
from tkinter import messagebox

import app
import ai_client

ROOT = Path(__file__).resolve().parent
OUT = ROOT / 'docs' / 'screens'
errors = []
SAFE_CONFIG = {
    'provider': 'heuristic',
    'provider_keys': {},
    'openai_api_key': '',
    'openai_model': 'gpt-4o-mini',
    'openai_base_url': 'https://api.openai.com/v1',
}


def capture(widget, path):
    if widget is None or not widget.winfo_exists():
        print(f'missing widget for {path.name}')
        return
    widget.update_idletasks()
    widget.update()
    widget.lift()
    widget.attributes('-topmost', True)
    widget.update()
    x, y = widget.winfo_rootx(), widget.winfo_rooty()
    w, h = widget.winfo_width(), widget.winfo_height()
    image = ImageGrab.grab(bbox=(x, y, x + w, y + h), all_screens=True)
    path.parent.mkdir(parents=True, exist_ok=True)
    image.save(path)
    widget.attributes('-topmost', False)
    print(f'saved {path.relative_to(ROOT)} ({w}x{h})')


def dialog_of(window):
    for child in window.winfo_children():
        if isinstance(child, tk.Toplevel) and child.winfo_exists():
            return child
    return None


def close_dialog(window):
    dialog = dialog_of(window)
    if dialog is not None:
        dialog.destroy()


def ready(window):
    ui = window.controller
    ui.work = lambda title, function, callback: callback(function())
    ui.fields['source'][0].set(str(ROOT / 'examples' / 'real_like' / 'sinh_vien.xlsx'))
    ui.load('source', auto_detect=True)
    if errors:
        raise RuntimeError(f'load failed: {errors}')
    ui.action.set('clean')
    ui.vars['key'].set('Họ và tên')
    ui.labels()
    window.update()
    capture(window, OUT / '01-main.png')

    def step2():
        capture(dialog_of(window), OUT / '02-preview.png')
        close_dialog(window)
        window.after(400, step3)

    def step3():
        window.after(700, step4)
        ui.open_ai_settings()

    def step4():
        capture(dialog_of(window), OUT / '03-settings.png')
        close_dialog(window)
        if errors:
            print(f'ui messages: {errors}')
        window.after(200, window.destroy)

    window.after(600, step2)
    ui.run()


def main():
    if not (ROOT / 'examples' / 'real_like' / 'sinh_vien.xlsx').exists():
        raise SystemExit('missing examples/real_like/sinh_vien.xlsx')
    with patch.object(ai_client, 'get_config', return_value=dict(SAFE_CONFIG, provider_keys={})), \
         patch.object(messagebox, 'showerror', side_effect=lambda *a, **k: errors.append(a)), \
         patch.object(messagebox, 'showinfo', side_effect=lambda *a, **k: errors.append(a)):
        app.gui(on_ready=ready)
    return 0


if __name__ == '__main__':
    sys.exit(main())
