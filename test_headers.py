"""Header-row diagnostics, empty margins and native Excel number formats."""
import json
from pathlib import Path
import tempfile
import app
from test_app import fixture

checks=[]


def check(name,condition):
    assert condition,name
    checks.append(name)


def fails(snapshot,row,expected):
    try: app.select_header(snapshot,row)
    except app.HeaderValidationError as error:
        assert all(value in str(error) for value in expected),str(error)
        assert error.snapshot is snapshot
        return error
    raise AssertionError('Invalid header accepted')


raw=dict(sheet='Data',sheets=['Data','Other'],rows=[['Report',None,None,None],
    [None,'ID','Name',None],[None,123,'An',None]],metadata={'sheets':{'Data':{'formats':{'B3':'00000'}}}})
fails(raw,1,['B1','C1','dòng tiêu đề 1'])
checks.append('title row reports actual empty cell addresses')
selected=app.select_header(raw,2)
check('empty table margins excluded',selected['headers']==['ID','Name'] and selected['data']==[[123,'An']] and selected['column_offset']==1)
check('original cell format offset preserved',app.format_at(selected,0,0)=='00000' and app.normalized_keys(selected,0)==['00123'])
check('raw source rows unchanged',raw['rows'][1]==[None,'ID','Name',None])
middle=dict(sheet='Data',rows=[['ID',None,'Name'],['1','populated','An']])
fails(middle,1,['B1'])
checks.append('populated unnamed middle column is not removed')
duplicate=dict(sheet='Data',rows=[['ID','ID'],['1','2']])
fails(duplicate,1,['Tên cột trùng','A1','B1'])
checks.append('duplicate header identifies both addresses')
fails(raw,100,['100','1 đến 3'])
checks.append('out-of-range row keeps raw inspection available')
with tempfile.TemporaryDirectory(dir=app.WORK) as private:
    directory=Path(private)
    source=directory/'source.xlsx'
    fixture(source,{'Data':[['Report',None,None],[None,'ID','Note'],[None,{'value':123,'format':'00000'},'  An  ']],'Other':[['ID'],['00123']]})
    try: app.inspect_file(source)
    except app.HeaderValidationError as error:
        check('native Excel sheet choices retained on header error',error.snapshot['sheets']==['Data','Other'] and 'B1' in str(error))
    else: raise AssertionError('Title row accepted')
    snapshot=app.inspect_file(source,'Data',2)
    check('native table with leading margin reads correctly',snapshot['headers']==['ID','Note'] and app.normalized_keys(snapshot,0)==['00123'])
    plan=dict(action='clean',key='Note')
    preview=app.preview_plan(snapshot,None,plan,output_dir=directory/'output')
    report=app.execute(snapshot,None,plan,output_dir=directory/'output',approval_token=preview['approval_token'])
    result=app.inspect_file(report['output'],report['result_sheet'])
    check('native output keeps code format and source sheets',result['data']==[[123,'An']] and app.format_at(result,0,0)=='00000' and report['preservation']['unchanged'])
    reference=app.inspect_file(source,'Other')
    lookup=dict(action='lookup',key='ID',ref_key='ID',ref_value='ID')
    proposal=app.preview_plan(snapshot,reference,lookup,output_dir=directory/'output')
    linked=app.execute(snapshot,reference,lookup,output_dir=directory/'output',approval_token=proposal['approval_token'])
    linked_result=app.inspect_file(linked['output'],linked['result_sheet'])
    check('lookup reads zero-padded code from original offset',linked['matched']==1 and linked_result['data'][0][0]=='00123' and linked_result['data'][0][2]=='00123')
result=dict(passed=len(checks),checks=checks,scope='Header parser + native Excel read/write; synthetic files only',app_sha256=app.file_hash(app.ROOT/'app.py'))
(app.ROOT/'header-test-results.json').write_text(json.dumps(result,indent=2),encoding='utf-8')
print(json.dumps(result,indent=2))
