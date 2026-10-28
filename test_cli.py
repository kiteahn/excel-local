"""Verify CLI proposals cannot execute without acceptance of the current token."""
import contextlib
import io
import json
from pathlib import Path
import sys
import tempfile
from unittest.mock import patch
import app

checks=[]


def invoke(arguments):
    output=io.StringIO()
    with patch.object(sys,'argv',['app.py',*arguments]),contextlib.redirect_stdout(output):
        app.main()
    return json.loads(output.getvalue())


with tempfile.TemporaryDirectory(dir=app.WORK) as private:
    directory=Path(private)
    source=directory/'source.csv'
    plan=directory/'plan.json'
    source.write_text('ID,Note\n001,  note  \n',encoding='utf-8')
    plan.write_text(json.dumps(dict(action='clean',key='Note')),encoding='utf-8')
    arguments=['--source',str(source),'--plan',str(plan)]
    with patch.object(app,'execute',return_value={'ok':True}) as execute:
        proposal=invoke(arguments)
        assert proposal['status']=='awaiting_approval' and not execute.called
        checks.append('default CLI previews without execution')
        token=proposal['preview']['approval_token']
        with contextlib.redirect_stderr(io.StringIO()):
            try: invoke(arguments+['--approve','invalid'])
            except SystemExit as error: assert error.code==2
            else: raise AssertionError('Invalid approval accepted')
        assert not execute.called
        checks.append('wrong token blocks execution')
        assert invoke(arguments+['--approve',token])=={'ok':True}
        assert execute.call_count==1 and execute.call_args.kwargs['approval_token']==token
        checks.append('matching token executes exactly once')
        source.write_text('ID,Note\n001,changed\n',encoding='utf-8')
        execute.reset_mock()
        with contextlib.redirect_stderr(io.StringIO()):
            try: invoke(arguments+['--approve',token])
            except SystemExit as error: assert error.code==2
            else: raise AssertionError('Stale approval accepted')
        assert not execute.called
        checks.append('changed source blocks previous token')

    # Test CLI Batch Processing with --batch-dir and --plan
    batch_dir = directory / 'batch_folder'
    batch_dir.mkdir()
    (batch_dir / 'f1.csv').write_text('ID,Note\n001,  note 1  \n', encoding='utf-8')
    (batch_dir / 'f2.csv').write_text('ID,Note\n002,  note 2  \n', encoding='utf-8')
    with patch.object(app, 'execute', return_value={'status': 'success', 'output': 'out.xlsx', 'stats': {}}):
        batch_out = invoke(['--batch-dir', str(batch_dir), '--plan', str(plan)])
        assert batch_out['total_files'] == 2 and batch_out['succeeded'] == 2
        checks.append('CLI --batch-dir with --plan executes batch successfully')

    # Test CLI Batch Processing with --batch-dir and --ai-prompt (0 tokens)
    with patch.object(app, 'execute', return_value={'status': 'success', 'output': 'out.xlsx', 'stats': {}}):
        batch_prompt_out = invoke(['--batch-dir', str(batch_dir), '--ai-prompt', 'Làm sạch khoảng trắng cột Note', '--provider', 'smart_local'])
        assert batch_prompt_out['total_files'] == 2 and batch_prompt_out['succeeded'] == 2
        checks.append('CLI --batch-dir with --ai-prompt infers plan and executes batch')

result=dict(passed=len(checks),checks=checks,scope='CLI gate; execution mocked')
(app.ROOT/'cli-test-results.json').write_text(json.dumps(result,indent=2),encoding='utf-8')
print(json.dumps(result,indent=2))

