"""Offline checks for the cloud preview + cancel confirmation (Q2).

No network call is made: cloud_preview() only applies PrivacyGuard and reads
the local routing configuration. The dialog itself is injected via ``ask``.
"""
import sys
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

from unittest.mock import patch
import ai_client
import app

CHECKS = []


def check(name, condition):
    assert condition, name
    CHECKS.append(name)


def test_exact_masked_payload():
    prompt = 'Gọi 0912345678 cho Nguyễn Văn An, lọc cột Điểm'
    preview = ai_client.cloud_preview(prompt, ['Họ tên', 'Điểm'], [], provider='openai', api_key='test-key-123')
    check('preview says a cloud call would happen', preview['will_send'] is True)
    check('preview prompt equals PrivacyGuard output',
          preview['prompt'] == ai_client.PrivacyGuard.anonymize_prompt(prompt))
    check('phone is masked in preview', '0912345678' not in preview['prompt'])
    check('name is masked in preview', 'Nguyễn Văn An' not in preview['prompt'])
    check('column names are the ones sent', preview['columns'] == ['Họ tên', 'Điểm'])
    check('provider name is listed', preview['providers'] == ['openai'])


def test_local_modes_never_send():
    preview = ai_client.cloud_preview('Lọc các dòng có Điểm > 8', ['Điểm'], [], provider='ollama')
    check('ollama never needs the cloud dialog', preview['will_send'] is False and preview['reason'] == 'ollama')

    preview = ai_client.cloud_preview('Lọc các dòng có Điểm > 8', ['Điểm'], [], provider='heuristic')
    check('heuristic never needs the cloud dialog', preview['will_send'] is False)


def test_no_key_means_no_send():
    cfg = {'provider': 'openai', 'provider_keys': {}, 'openai_api_key': ''}
    with patch.object(ai_client, 'get_config', return_value=cfg):
        preview = ai_client.cloud_preview('Lọc các dòng có Điểm > 8', ['Điểm'], [], provider='openai')
    check('missing key means no cloud send', preview['will_send'] is False and preview['reason'] == 'no_cloud_key')


def test_preview_never_calls_cloud_code():
    def explode(*args, **kwargs):
        raise AssertionError('cloud_preview must not call the network layer')

    with patch.object(ai_client, 'call_openai_plan', side_effect=explode):
        preview = ai_client.cloud_preview('Lọc các dòng có Điểm > 8', ['Điểm'], [], provider='openai', api_key='test-key-123')
    check('preview does not call call_openai_plan', preview['will_send'] is True)


def test_cancel_blocks_and_shows_masked_text():
    seen = {}

    def deny(title, body, parent=None):
        seen['title'] = title
        seen['body'] = body
        return False

    source = {'headers': ['Họ tên', 'Điểm']}
    approved, preview = app.confirm_cloud_preview(
        'Gọi 0912345678 cho Nguyễn Văn An', source, None, 'openai', ask=deny, api_key='test-key-123')
    check('cancel returns not approved', approved is False)
    check('cancel closes over the cloud preview', preview['will_send'] is True)
    body = seen['body']
    check('dialog title mentions Cloud', 'Cloud' in seen['title'])
    check('dialog states the exact destination/provider', 'openai' in body)
    check('dialog shows masked phone', '[SDT_AN_DANH]' in body and '0912345678' not in body)
    check('dialog shows masked name', '[TEN_AN_DANH]' in body and 'Nguyễn Văn An' not in body)
    check('dialog lists the columns sent', 'Họ tên' in body and 'Điểm' in body)
    check('dialog explains cancellation means no data leaves', 'không có dữ liệu nào rời khỏi máy' in body)


def test_accept_allows_and_local_skips_dialog():
    approved, _ = app.confirm_cloud_preview(
        'Lọc các dòng có Điểm > 8', {'headers': ['Điểm']}, None, 'openai',
        ask=lambda *args, **kwargs: True, api_key='test-key-123')
    check('accept returns approved', approved is True)

    def should_not_ask(*args, **kwargs):
        raise AssertionError('local mode must not open the cloud dialog')

    approved, preview = app.confirm_cloud_preview(
        'Lọc các dòng có Điểm > 8', {'headers': ['Điểm']}, None, 'ollama', ask=should_not_ask)
    check('local mode skips the dialog', approved is True and preview['will_send'] is False)


def main():
    test_exact_masked_payload()
    test_local_modes_never_send()
    test_no_key_means_no_send()
    test_preview_never_calls_cloud_code()
    test_cancel_blocks_and_shows_masked_text()
    test_accept_allows_and_local_skips_dialog()
    print(f'All {len(CHECKS)} cloud-preview checks passed (no network used).')


if __name__ == '__main__':
    main()
