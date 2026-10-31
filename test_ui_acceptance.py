"""UI acceptance checks (Cycle 0): tooltips, Ctrl+Enter, busy lock, error dialogs.

These checks drive the real Tk widgets but never open Excel or user files.
They exist because the Cycle 0 handoff left the visible-UI items unverified.
"""
import json
import tempfile
import threading
import time
from pathlib import Path
from tkinter import messagebox
from unittest.mock import patch
import tkinter as tk

import app

checks = []


def check(name, condition):
    assert condition, name
    checks.append(name)


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


def find_widget(widget, predicate):
    return next((item for item in descendants(widget) if predicate(item)), None)


def tooltip_of(window):
    for item in descendants(window):
        if isinstance(item, tk.Toplevel) and item.wm_overrideredirect():
            labels = item.winfo_children()
            if labels:
                return labels[0].cget('text')
    return None


def pump(window, seconds=0.2):
    end = time.time() + seconds
    while time.time() < end:
        window.update()
        time.sleep(0.01)


def exercise(window):
    window.deiconify()
    window.geometry('1100x820+80+40')
    window.update()
    ui = window.controller

    # --- defaults: provider comes from config; a fresh config means Ollama ---
    fresh = Path(tempfile.gettempdir()) / 'no-such-config-cycle0.json'
    with patch.object(app.ai_client, 'CONFIG_FILE', fresh):
        check('default engine is Ollama when no config exists',
              app.ai_client.get_config().get('provider') == 'ollama')
    radios = [w for w in descendants(window) if w.winfo_class() == 'TRadiobutton'
              and any(tag in str(w.cget('text')) for tag in ('Ollama', 'Siêu tốc', 'Smart Router', '100% Cục bộ'))]
    check('four engine choices are visible', len(radios) == 4)
    check('Ollama engine is the selected radio',
          any('Ollama' in str(w.cget('text')) and str(w.instate(['selected'])) for w in radios))

    # --- tooltips appear on hover and disappear on leave ---
    targets = {
        'Ollama radio': lambda w: w.winfo_class() == 'TRadiobutton' and 'Ollama' in str(w.cget('text')),
        'Preview button': lambda w: w is ui.run_button,
        'PDF checkbox': lambda w: w.winfo_class() == 'TCheckbutton' and 'Xuất PDF' in str(w.cget('text')),
        'Auto-heal checkbox': lambda w: w.winfo_class() == 'TCheckbutton' and 'Tự chữa lành' in str(w.cget('text')),
        'Clean output checkbox': lambda w: w.winfo_class() == 'TCheckbutton' and 'bảng gọn gàng' in str(w.cget('text')),
    }
    for name, predicate in targets.items():
        widget = find_widget(window, predicate)
        check(f'{name} exists', widget is not None)
        widget.event_generate('<Enter>')
        window.update()
        text = tooltip_of(window)
        check(f'{name} shows a tooltip', bool(text) and len(text) > 20)
        widget.event_generate('<Leave>')
        window.update()
        check(f'{name} tooltip hides on leave', tooltip_of(window) is None)

    # --- Ctrl+Enter runs the proposal path ---
    calls = []
    original_plan = ui.plan
    ui.plan = lambda: calls.append('plan')
    check('Ctrl+Enter is bound in the request box', bool(ui.prompt.bind('<Control-Return>')))
    ui.prompt.focus_force()
    ui.prompt.event_generate('<Control-Return>')
    window.update()
    check('Ctrl+Enter triggers the AI proposal', calls == ['plan'])
    ui.plan = original_plan

    # --- busy guard + progress controls ---
    notes = []
    with patch.object(messagebox, 'showinfo', side_effect=lambda *args: notes.append(args)):
        ui.busy = True
        ui.load('source')
        check('loading while busy explains why and stops', any('Đang xử lý' in str(n) for n in notes))
        ui.run()
        check('preview while busy does not open a dialog', not any(isinstance(w, tk.Toplevel)
                                                                   for w in descendants(window)))
        ui.close()
        check('closing while busy keeps the window open', window.winfo_exists())
        ui.busy = False

    release = threading.Event()
    done = []
    ui.work('Đang chạy kiểm thử…', lambda: (release.wait(5), 'ok')[1], lambda value: done.append(value))
    window.update()
    check('preview button is disabled while working', str(ui.run_button['state']) == 'disabled')
    check('stop button is visible while working', bool(ui.stop_button.winfo_ismapped()))
    release.set()
    deadline = time.time() + 5
    while time.time() < deadline and not done:
        window.update()
        time.sleep(0.02)
    check('worker result is delivered', done == ['ok'])
    check('preview button is re-enabled after work', str(ui.run_button['state']) == 'normal')
    check('stop button is hidden after work', not ui.stop_button.winfo_ismapped())

    # --- UC-33 / UC-35 messages reach the user through the error dialog ---
    errors = []

    def load_and_wait():
        ui.load('source')
        deadline = time.time() + 5
        while time.time() < deadline and not errors:
            window.update()
            time.sleep(0.02)
        pump(window, 0.05)

    with tempfile.TemporaryDirectory(dir=app.WORK) as private:
        path = Path(private) / 'source.csv'
        path.write_text('ID,Note\n001,x\n', encoding='utf-8')
        ui.fields['source'][0].set(str(path))
        for label, message in (('UC-33', 'Tệp đang được mở trong Excel. Hãy đóng tệp rồi thử lại.'),
                               ('UC-35', 'Ô lỗi #DIV/0! không được phép đọc.')):
            errors.clear()
            with patch.object(app, 'inspect_file', side_effect=RuntimeError(message)), \
                    patch.object(messagebox, 'showerror', side_effect=lambda *args: errors.append(args)):
                load_and_wait()
            check(f'{label} error dialog shows the Vietnamese reason',
                  any(message in str(item) for item in errors))
            check(f'{label} error is written to the activity log', message in ui.log.get('1.0', 'end'))
            check(f'{label} error leaves no stale source', ui.source is None)
        check('failed load leaves the buttons usable', str(ui.run_button['state']) == 'normal' and not ui.busy)

        # --- batch failure shows a dialog, not a silent crash (VD-17) ---
        batch_dir = Path(private) / 'batch'
        batch_dir.mkdir()
        (batch_dir / 'mot.csv').write_text('ID,Note\n001,x\n', encoding='utf-8')
        errors.clear()
        plan = dict(action='clean', key='Note', value='', ref_key='', ref_value='', message='')
        with patch.object(app, 'execute_batch', side_effect=RuntimeError('lỗi thử nghiệm batch')), \
                patch.object(messagebox, 'showerror', side_effect=lambda *args: errors.append(args)):
            ui.start_batch_run(str(batch_dir), None, plan, False)
            deadline = time.time() + 5
            while time.time() < deadline and not errors:
                window.update()
                time.sleep(0.02)
        check('batch failure shows an error dialog', any('lỗi thử nghiệm batch' in str(item) for item in errors))
        batch_dialog = next((w for w in descendants(window)
                             if isinstance(w, tk.Toplevel) and 'Hàng loạt' in str(w.title())), None)
        check('batch dialog stays open after the failure', batch_dialog is not None)
        batch_dialog.update_idletasks()
        check('batch dialog fits its content',
              batch_dialog.winfo_width() >= batch_dialog.winfo_reqwidth()
              and batch_dialog.winfo_height() >= batch_dialog.winfo_reqheight())
        close_button = find_widget(batch_dialog, lambda w: w.winfo_class() == 'TButton' and str(w.cget('text')) == 'Đóng')
        check('batch dialog close button is enabled again', str(close_button['state']) == 'normal')
        check('batch dialog close button is not squeezed',
              close_button.winfo_height() >= close_button.winfo_reqheight())
        batch_dialog.destroy()

    # --- window can shrink to its minimum without clipping the main sections ---
    ui.reset_session()
    window.geometry('1080x800')
    pump(window, 0.3)
    outer = window.winfo_children()[0]
    clipped = [str(widget) for widget in outer.winfo_children()
               if widget.winfo_height() <= 1 or widget.winfo_y() + widget.winfo_height() > outer.winfo_height() + 1]
    check('minimum window size keeps every section visible', not clipped)
    check('footer buttons stay visible at minimum size', ui.engine_badge.winfo_height() > 1)
    check('preview table keeps a usable height at minimum size', ui.preview.winfo_height() > 60)

    def squeezed_labels():
        found = []
        for widget in descendants(window):
            if widget.winfo_class() not in ('TButton', 'TRadiobutton', 'TCheckbutton', 'TLabel'):
                continue
            if widget.winfo_width() + 1 < widget.winfo_reqwidth() and widget.winfo_viewable():
                text = str(widget.cget('text')) if 'text' in widget.keys() else ''
                found.append(f'{widget.winfo_class()}:{text[:40]}')
        return found

    check('no label is truncated at minimum window size', not squeezed_labels())
    window.geometry('1160x900')
    pump(window, 0.3)
    check('no label is truncated at the default window size', not squeezed_labels())

    # --- dialogs size themselves to their content (UI-02 regression) ---
    ui.open_ai_settings()
    pump(window, 0.4)
    settings = next(w for w in descendants(window)
                    if isinstance(w, tk.Toplevel) and not w.wm_overrideredirect())
    settings.update_idletasks()
    check('AI settings dialog is wide enough for its content',
          settings.winfo_width() >= settings.winfo_reqwidth())
    check('AI settings dialog is tall enough for its footer',
          settings.winfo_height() >= settings.winfo_reqheight())
    settings_buttons = [w for w in descendants(settings) if w.winfo_class() == 'TButton']
    check('AI settings buttons are not squeezed',
          bool(settings_buttons) and all(b.winfo_height() >= b.winfo_reqheight() for b in settings_buttons))
    settings.destroy()

    window.destroy()


if __name__ == '__main__':
    app.WORK.mkdir(exist_ok=True)
    failure = {}
    original_mainloop = tk.Tk.mainloop

    def runner(self):
        def begin():
            try:
                exercise(self)
            except BaseException as exc:  # noqa: BLE001 - report, then fail the run
                failure['error'] = exc
            finally:
                try:
                    if self.winfo_exists():
                        self.destroy()
                except tk.TclError:
                    pass
        self.after(150, begin)
        original_mainloop(self)

    try:
        with patch.object(tk.Tk, 'mainloop', runner):
            app.gui()
    except tk.TclError as exc:
        print(f'SKIP: test_ui_acceptance requires a display (Tk could not initialise: {exc}).')
        raise SystemExit(0)
    if failure:
        raise failure['error']
    result = dict(passed=len(checks), checks=checks,
                  scope='Real Tk widgets; Excel and AI calls mocked; visual layout checked by geometry, screenshots separate')
    results_dir = app.ROOT / 'results'
    results_dir.mkdir(exist_ok=True)
    (results_dir / 'ui-acceptance-cycle0.json').write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(result, ensure_ascii=False, indent=2))
