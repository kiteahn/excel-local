"""File swaps must be rejected before native Excel sees the copied workbook."""
import json
from pathlib import Path
import tempfile
from unittest.mock import patch
import app
from test_app import fixture

checks=[]


def rejected(name,function):
    try: function()
    except ValueError: checks.append(name)
    else: raise AssertionError(name)


with tempfile.TemporaryDirectory(dir=app.WORK) as private:
    directory=Path(private)
    source=directory/'source.xlsx'
    fixture(source,{'Data':[['ID'],['001']]})
    with patch.object(app,'bridge',side_effect=AssertionError('Unverified workbook reached Excel')):
        def changed_copy(incoming,outgoing):
            fixture(outgoing,{'Data':[['ID'],['changed']]})
        with patch.object(app.shutil,'copyfile',side_effect=changed_copy):
            rejected('Changed copy rejected before Excel opens',lambda:app.inspect_file(source))
        fixture(source,{'Data':[['ID'],[{'formula':'WEBSERVICE("https://example.invalid")','cached':0}]]})
        preflight=app.preflight
        def swapped_preflight(path):
            if Path(path)==source: return {'sheets':{'Data':{'formats':{}}}}
            return preflight(path)
        with patch.object(app,'preflight',side_effect=swapped_preflight):
            rejected('Exact copied XML is checked before Excel opens',lambda:app.inspect_file(source))
    fixture(source,{'Data':[['ID'],['001']]})
    snapshot=app.inspect_file(source)
    assert snapshot['headers']==['ID'] and snapshot['data']==[['001']] and snapshot['hash']==app.file_hash(source)
    checks.append('Verified one-column workbook reads in native Excel')
result=dict(passed=len(checks),checks=checks,scope='Two adversarial swaps + one real Excel read; synthetic files only')
(app.ROOT/'read-boundary-test-results.json').write_text(json.dumps(result,indent=2),encoding='utf-8')
print(json.dumps(result,indent=2))
