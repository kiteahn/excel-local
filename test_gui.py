"""Test the actual Tk approval controls without executing Excel or using user files."""
import json
from pathlib import Path
import tempfile
import tkinter as tk
from tkinter import messagebox
from unittest.mock import patch
import app

checks=[]


def check(name,condition):
    assert condition,name
    checks.append(name)


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


def exercise(window):
    window.withdraw()
    ui=window.controller
    errors=[]
    performed=[]
    with tempfile.TemporaryDirectory(dir=app.WORK) as private, patch.object(messagebox,'showerror',side_effect=lambda *args:errors.append(args)), patch.object(messagebox,'showinfo'):
        path=Path(private)/'source.csv'
        path.write_text('#0,Note\n001,  Đoàn  \n',encoding='utf-8')
        ui.fields['source'][0].set(str(path))
        ui.work=lambda title,function,callback:callback(function())
        ui.load('source')
        check('header #0 safely displayed',ui.preview.heading('c0')['text']=='#0')
        ui.action.set('clean')
        ui.vars['key'].set('Note')
        ui.perform=lambda *args:performed.append(args)
        def button(name):
            return next(widget for widget in descendants(window) if widget.winfo_name()==name)
        ui.run()
        check('manual preview waits without executing',not performed)
        button('cancel').invoke()
        check('cancel produces no execution',not performed)
        ui.run()
        button('approve').invoke()
        check('explicit acceptance executes exactly once',len(performed)==1 and len(performed[0][-1])==64)
        performed.clear()
        plan=dict(action='clean',key='Note',value='',ref_key='',ref_value='',message='Làm sạch ghi chú.')
        with patch.object(app,'ai_plan',return_value=plan):
            ui.plan()
        check('AI proposal waits for acceptance',not performed and button('approve').winfo_exists())
        button('cancel').invoke()
        with patch.object(app,'ai_plan',return_value=dict(action='clarify',message='Bạn muốn cộng cột nào?')):
            ui.plan()
        check('AI clarification waits without execution',not performed and 'bổ sung' in ui.status.get())
        ui.run()
        path.write_text('#0,Note\n001,changed\n',encoding='utf-8')
        button('approve').invoke()
        check('changed file blocks acceptance',not performed and bool(errors))
        ui.fields['source'][2].set('invalid')
        ui.run()
        check('invalid header input reports an error',len(errors)>=2)
        ui.busy=True
        ui.close()
        check('busy window keeps worker alive',window.winfo_exists())
        ui.busy=False
        def broken(value): raise ValueError('callback regression')
        ui.jobs.put((broken,None,None))
        ui.poll()
        check('callback error is handled',any('callback regression' in str(error) for error in errors))
        ui.fields['source'][2].set(1)
        path.write_text('Report,\nID,Note\n001,  note  \n',encoding='utf-8')
        ui.load('source')
        check('header error shows raw row numbers and blocks execution',ui.source is None and ui.preview.heading('c0')['text']=='Dòng Excel' and bool(ui.preview.get_children()))
        check('sheet selector remains available after header error',ui.source_sheet['values']==('CSV',))
        ui.fields['source'][2].set(2)
        ui.load('source')
        check('correcting header row recovers same file',ui.source['headers']==['ID','Note'] and ui.source['header_row']==2)
        ui.reset_session()
        check('reset session clears all inputs and returns to step 1', ui.source is None and ui.fields['source'][0].get()=='' and ui.current_step==1)
    window.destroy()


if __name__=='__main__':
    app.WORK.mkdir(exist_ok=True)
    with patch.object(tk.Tk,'mainloop',exercise):
        app.gui()
    result=dict(passed=len(checks),checks=checks,scope='Actual Tk controls; Excel execution mocked')
    (app.ROOT/'gui-test-results.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(result,ensure_ascii=False,indent=2))
