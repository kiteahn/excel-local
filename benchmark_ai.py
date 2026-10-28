"""Measure one original request, one current request, and one RAM-cache hit."""
import importlib.machinery
import importlib.util
import json
import time
import app
import ai_client

loader=importlib.machinery.SourceFileLoader('before',str(app.ROOT/'app-v0.1.py.bak'))
before=importlib.util.module_from_spec(importlib.util.spec_from_loader('before',loader))
loader.exec_module(before)
source={'headers':['Mã hồ sơ','Nhóm','Số tiền','Ghi chú']}
reference={'headers':['Mã hồ sơ','Họ tên']}
prompt='Đối chiếu Mã hồ sơ với Mã hồ sơ trong danh mục, lấy Họ tên.'
results={}
ai_client._request_plan.cache_clear()
for name,function in [('before',before.ai_plan),('after',app.ai_plan),('repeat',app.ai_plan)]:
    started=time.perf_counter()
    plan=function(prompt,source,reference)
    elapsed=round(time.perf_counter()-started,3)
    assert plan['action']=='lookup' and plan['key']=='Mã hồ sơ' and plan['ref_value']=='Họ tên'
    results[name]=elapsed
    print(name,elapsed,flush=True)
report=dict(seconds=results,model=ai_client.MODEL,ai_client_sha256=app.file_hash(app.ROOT/'ai_client.py'),scope='One request per case, CPU; cache is RAM-only and runtime/model checks still run')
(app.ROOT/'benchmark-ai.json').write_text(json.dumps(report,indent=2),encoding='utf-8')
