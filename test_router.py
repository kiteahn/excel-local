"""Comprehensive unit tests for Smart Multi-Model Router (9Router-style cascade)."""
import sys
from pathlib import Path

# UTF-8 stdout configuration for Windows
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception: pass
if sys.stderr and hasattr(sys.stderr, 'reconfigure'):
    try: sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception: pass

import json
import tempfile
from unittest.mock import patch
import ai_client

def run_tests():
    # Hermetic: redirect the config file to a temp directory so the tests never
    # touch (or depend on) the real user config next to the code, on any machine.
    _tmp_config_dir = tempfile.TemporaryDirectory()
    ai_client.CONFIG_FILE = Path(_tmp_config_dir.name) / 'config.json'
    has_dpapi = ai_client.dpapi_protect('probe').startswith('dpapi:')
    dpapi_simulated = False
    if not has_dpapi:
        # save_config() now refuses to persist unencrypted keys, so off Windows
        # (where DPAPI does not exist) simulate a faithful reversible DPAPI
        # cipher to exercise the same save/read contract. The real cipher is
        # only verified on Windows.
        import base64
        ai_client.dpapi_protect = lambda s: '' if not s else 'dpapi:' + base64.b64encode(s.encode('utf-8')).decode('ascii')
        ai_client.dpapi_unprotect = lambda c: base64.b64decode(c[6:].encode('ascii')).decode('utf-8') if c and c.startswith('dpapi:') else (c or '')
        dpapi_simulated = True
        print("NOTE: No Windows DPAPI on this platform; using a simulated reversible cipher for the config save/read checks.")

    print("=== 1. Test DPAPI Encryption for Provider Keys ===")
    raw_keys = {
        'gemini': 'AIzaSyFakeGeminiKey123',
        'groq': 'gsk_FakeGroqKey456',
        'openai': 'sk-FakeOpenAI789'
    }
    saved_cfg = ai_client.save_config({
        'provider': 'smart',
        'provider_keys': raw_keys,
        'openai_api_key': 'sk-LegacyKey'
    })
    # Read the raw file on disk to verify no plaintext keys exist:
    raw_file = ai_client.CONFIG_FILE.read_text(encoding='utf-8')
    assert 'AIzaSyFakeGeminiKey123' not in raw_file, "Gemini key must not be in plaintext on disk!"
    assert 'gsk_FakeGroqKey456' not in raw_file, "Groq key must not be in plaintext on disk!"
    assert 'sk-FakeOpenAI789' not in raw_file, "OpenAI key must not be in plaintext on disk!"
    print("PASS: All keys encrypted on disk!" + (" (simulated cipher, non-Windows platform)" if dpapi_simulated else " (Windows DPAPI)"))

    reloaded_cfg = ai_client.get_config()
    assert reloaded_cfg['provider_keys']['gemini'] == 'AIzaSyFakeGeminiKey123'
    assert reloaded_cfg['provider_keys']['groq'] == 'gsk_FakeGroqKey456'
    assert reloaded_cfg['provider_keys']['openai'] == 'sk-FakeOpenAI789'
    print("PASS: Decryption on the fly works perfectly!")

    print("\n=== 2. Test Smart Router Candidate Generation ===")
    candidates = ai_client.get_active_candidates(reloaded_cfg)
    names = [c['name'] for c in candidates]
    print("Active Candidates in priority order:", names)
    assert names[0] == 'Google Gemini (Free Tier)'
    assert names[1] == 'Groq Cloud (Free Llama 70B)'
    assert names[2] == 'OpenAI Official'
    print("PASS: Priority order is Gemini ➔ Groq ➔ OpenAI!")

    print("\n=== 3. Test Q&A / Semantic Instant Answering (0ms, 0 Network Calls) ===")
    qa_plan = ai_client.request_plan("Tại sao bị lỗi #N/A vậy?", ["MSSV", "Họ tên"], ["MSSV", "Khoa"])
    assert qa_plan['action'] == 'clarify'
    assert 'Lỗi #N/A' in qa_plan['message']
    print("PASS: Instant Q&A response for error explanation!")

    col_plan = ai_client.request_plan("Bảng này có bao nhiêu cột?", ["MSSV", "Họ tên", "Lớp"])
    assert col_plan['action'] == 'clarify'
    assert '3 cột' in col_plan['message']
    print("PASS: Instant response for column count inquiries!")

    print("\n=== 4. Test Auto-Rotation Cascade (Gemini Fails ➔ Groq Takes Over) ===")
    ai_client._PLAN_CACHE.clear()

    call_sequence = []
    def mock_call_openai(prompt, src_cols, ref_cols, api_key=None, model=None, base_url=None):
        call_sequence.append(base_url)
        if 'generativelanguage.googleapis.com' in base_url:
            # Simulate Gemini 429 Rate Limit
            raise RuntimeError("Google Gemini HTTP 429: Resource has been exhausted (quota exceeded)")
        if 'api.groq.com' in base_url:
            # Simulate Groq succeeds
            return {
                'action': 'lookup',
                'key': 'MSSV',
                'value': '',
                'ref_key': 'MSSV',
                'ref_value': 'Quê quán',
                'message': 'Đã đối chiếu thành công qua Groq Cloud.'
            }
        raise RuntimeError("Unexpected call")

    with patch('ai_client.call_openai_plan', side_effect=mock_call_openai):
        plan = ai_client.request_plan("Lấy quê quán cho sinh viên theo mã", ["MSSV", "Tên"], ["MSSV", "Quê quán"])
        print("Plan returned:", plan)
        assert plan['action'] == 'lookup'
        assert plan['engine_used'] == 'Groq Cloud (Free Llama 70B)'
        assert any('api.groq.com' in url for url in call_sequence)
        print("PASS: Automatic failover from Gemini to Groq succeeded seamlessly!")

    print("\n=== 5. Test Fallback to Ollama Local When All Cloud Quotas Fail ===")
    ai_client._PLAN_CACHE.clear()

    def mock_call_all_fail(*args, **kwargs):
        raise RuntimeError("Cloud HTTP 429 / Offline")

    with patch('ai_client.call_openai_plan', side_effect=mock_call_all_fail), \
         patch('ai_client.ensure_runtime'), \
         patch('ai_client._request_plan_ollama', return_value=json.dumps({
             'action': 'summary',
             'key': 'Lớp',
             'value': 'Học phí',
             'ref_key': '',
             'ref_value': '',
             'message': 'Tổng hợp học phí theo lớp qua Ollama.'
         })):
        plan = ai_client.request_plan("Tính tổng học phí từng lớp", ["Lớp", "Học phí"])
        print("Plan returned:", plan)
        assert plan['action'] == 'summary'
        assert 'Ollama Cục bộ' in plan['engine_used']
        print("PASS: Automatic failover to local offline Ollama succeeded!")

    print("\n=== 6. Test Semantic Heuristic Rescue When Everything Is Down ===")
    ai_client._PLAN_CACHE.clear()

    with patch('ai_client.call_openai_plan', side_effect=mock_call_all_fail), \
         patch('ai_client.ensure_runtime', side_effect=RuntimeError("Ollama daemon down")):
        plan = ai_client.request_plan("Cộng tổng tiền của từng lớp", ["Lớp", "Tiền"])
        print("Plan returned:", plan)
        assert plan['action'] == 'summary'
        assert plan['key'] == 'Lớp'
        assert plan['value'] == 'Tiền'
        assert plan['engine_used'] == 'Smart Heuristic Engine'
        print("PASS: Semantic Heuristic Rescue rescued the plan with zero network and zero daemons!")

    print("\n=== 7. Test save_config Refuses To Persist Unencrypted Keys (VD-11) ===")
    before = ai_client.CONFIG_FILE.read_text(encoding='utf-8') if ai_client.CONFIG_FILE.exists() else None
    real_protect = ai_client.dpapi_protect
    ai_client.dpapi_protect = lambda s: s  # simulate DPAPI failure: returns plaintext
    try:
        try:
            ai_client.save_config({'provider_keys': {'gemini': 'AIzaSyShouldNeverHitDisk'}})
            raise AssertionError("save_config must refuse to save when encryption fails!")
        except RuntimeError as exc:
            assert 'DPAPI' in str(exc)
        after = ai_client.CONFIG_FILE.read_text(encoding='utf-8') if ai_client.CONFIG_FILE.exists() else None
        assert after == before, "Config file must be left untouched when saving is refused!"
        assert after is None or 'AIzaSyShouldNeverHitDisk' not in after
        print("PASS: save_config raised a clear error and wrote nothing when DPAPI failed!")
    finally:
        ai_client.dpapi_protect = real_protect

    # Cleanup: the config used here lives in a temp directory and is discarded;
    # the real user config next to the code was never touched.
    _tmp_config_dir.cleanup()
    print("\nAll 7 Smart Multi-Model Router architecture tests PASSED brilliantly!")

if __name__ == '__main__':
    run_tests()
