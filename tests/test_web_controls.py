# -*- coding: utf-8 -*-
"""Web control-plane security and validation-state regression tests."""

import asyncio
import json
import queue
import subprocess
from pathlib import Path
from types import SimpleNamespace

import pytest


async def _asgi_request_async(app, method, path, *, json_body=None, body=None,
                              headers=None, client_host='127.0.0.1',
                              include_content_length=True):
    if body is None:
        body = b'' if json_body is None else json.dumps(json_body).encode('utf-8')
    request_headers = {k.lower(): v for k, v in (headers or {}).items()}
    if json_body is not None:
        request_headers.setdefault('content-type', 'application/json')
    if include_content_length:
        request_headers.setdefault('content-length', str(len(body)))
    scope = {
        'type': 'http',
        'asgi': {'version': '3.0'},
        'http_version': '1.1',
        'method': method,
        'scheme': 'http',
        'path': path,
        'raw_path': path.encode('ascii'),
        'query_string': b'',
        'headers': [(k.encode('latin-1'), v.encode('latin-1'))
                    for k, v in request_headers.items()],
        'client': (client_host, 45678),
        'server': ('127.0.0.1', 8000),
        'root_path': '',
    }
    messages = []
    delivered = False

    async def receive():
        nonlocal delivered
        if not delivered:
            delivered = True
            return {'type': 'http.request', 'body': body, 'more_body': False}
        return {'type': 'http.disconnect'}

    async def send(message):
        messages.append(message)

    await app(scope, receive, send)
    start = next(m for m in messages if m['type'] == 'http.response.start')
    response_body = b''.join(
        m.get('body', b'') for m in messages if m['type'] == 'http.response.body'
    )
    response_headers = {
        k.decode('latin-1').lower(): v.decode('latin-1')
        for k, v in start.get('headers', [])
    }
    return SimpleNamespace(
        status_code=start['status'], content=response_body, headers=response_headers,
        json=lambda: json.loads(response_body.decode('utf-8')),
    )


def _request(app, method, path, **kwargs):
    return asyncio.run(_asgi_request_async(app, method, path, **kwargs))


@pytest.fixture
def web_server(tmp_path, monkeypatch):
    from web import server

    monkeypatch.setattr(server, '_jobs', {})
    monkeypatch.setattr(server, '_queue', queue.Queue())
    monkeypatch.setattr(server, 'DATA_DIR', str(tmp_path))
    monkeypatch.setattr(server, 'OUT_DIR', str(tmp_path / 'out'))
    monkeypatch.setattr(server, 'JOBS_JSON', str(tmp_path / 'jobs.json'))
    monkeypatch.setattr(server, 'MAX_JOBS', 100)
    monkeypatch.setattr(server, 'MAX_ACTIVE_JOBS', 8, raising=False)
    monkeypatch.setattr(server, 'MAX_REQUEST_BYTES', 64 * 1024, raising=False)
    monkeypatch.delenv('WEB_API_TOKEN', raising=False)
    monkeypatch.delenv('WEB_ALLOW_LLM', raising=False)
    monkeypatch.delenv('WEB_ALLOW_UNVERIFIED_DOWNLOAD', raising=False)
    return server


def test_verify_result_is_structured_and_preserves_nonzero_returncode(web_server, tmp_path,
                                                                      monkeypatch):
    monkeypatch.setattr(web_server.shutil, 'which', lambda name: '/usr/bin/soffice')

    def fake_verify(cmd, **kwargs):
        if '--json-summary' in cmd:
            Path(cmd[cmd.index('--json-summary') + 1]).write_text(
                '{"schema_version":1,"verdict":"FAIL","exit_code":1,"controls":[]}',
                encoding='utf-8',
            )
        return SimpleNamespace(returncode=7, stdout='CHECK FAIL\n', stderr='')

    monkeypatch.setattr(web_server.subprocess, 'run', fake_verify)

    result = web_server._run_verify('model.xlsx', 'model.addr.json', str(tmp_path / 'job.log'))

    assert result['status'] == 'failed_validation'
    assert result['returncode'] == 7
    assert result['verdict'] == 'FAIL'
    assert result['verification']['exit_code'] == 1
    assert result['summary'] == 'CHECK FAIL\n'


def test_verify_requires_passing_json_verdict_even_when_process_returns_zero(web_server,
                                                                              tmp_path,
                                                                              monkeypatch):
    monkeypatch.setattr(web_server.shutil, 'which', lambda name: '/usr/bin/soffice')

    def fake_verify(cmd, **kwargs):
        if '--json-summary' in cmd:
            Path(cmd[cmd.index('--json-summary') + 1]).write_text(
                '{"schema_version":1,"verdict":"REVIEW","exit_code":2,"controls":[]}',
                encoding='utf-8',
            )
        return SimpleNamespace(returncode=0, stdout='REVIEW\n', stderr='')

    monkeypatch.setattr(web_server.subprocess, 'run', fake_verify)

    result = web_server._run_verify('model.xlsx', 'model.addr.json', str(tmp_path / 'job.log'))

    assert result['status'] == 'failed_validation'
    assert result['returncode'] == 0
    assert result['verdict'] == 'REVIEW'


@pytest.mark.parametrize(
    ('soffice', 'verify_returncode', 'expected_status'),
    [
        ('/usr/bin/soffice', 0, 'verified'),
        ('/usr/bin/soffice', 3, 'failed_validation'),
        (None, None, 'built_unverified'),
    ],
)
def test_run_job_persists_validation_outcome(web_server, tmp_path, monkeypatch,
                                             soffice, verify_returncode, expected_status):
    monkeypatch.setattr(web_server.shutil, 'which', lambda name: soffice)

    def fake_run(cmd, **kwargs):
        if cmd[1].endswith('build_model.py'):
            xlsx = Path(cmd[cmd.index('--out') + 1])
            addr = Path(cmd[cmd.index('--addr') + 1])
            xlsx.write_bytes(b'xlsx')
            addr.write_text('{"meta":{"name":"测试公司"}}', encoding='utf-8')
            return SimpleNamespace(returncode=0, stdout='', stderr='')
        if '--json-summary' in cmd:
            verdict = 'PASS' if verify_returncode == 0 else 'FAIL'
            exit_code = 0 if verify_returncode == 0 else 1
            Path(cmd[cmd.index('--json-summary') + 1]).write_text(
                json.dumps({'schema_version': 1, 'verdict': verdict,
                            'exit_code': exit_code, 'controls': []}),
                encoding='utf-8',
            )
        return SimpleNamespace(
            returncode=verify_returncode,
            stdout='PASS\n' if verify_returncode == 0 else 'FAIL\n',
            stderr='',
        )

    monkeypatch.setattr(web_server.subprocess, 'run', fake_run)
    job = {
        'id': 'a1b2c3d4', 'code': '300476', 'name': '300476',
        'status': 'queued', 'error': None, 'created_at': '2026-08-18 10:00:00',
        'started_at': None, 'finished_at': None,
        'params': {'config': None, 'dr': None, 'consensus': None,
                   'announcements': False, 'llm': 'off'},
        'xlsx': None, 'log_file': None,
    }
    web_server._jobs[job['id']] = job

    web_server._run_job(job)

    assert job['status'] == expected_status
    assert job['verify_status'] == expected_status
    assert job['verify_returncode'] == verify_returncode
    assert job['verify_verdict'] == ('PASS' if verify_returncode == 0 else
                                      'FAIL' if verify_returncode else None)
    persisted = json.loads(Path(web_server.JOBS_JSON).read_text(encoding='utf-8'))[0]
    assert persisted['status'] == expected_status
    assert persisted['verify_status'] == expected_status
    assert persisted['verify_returncode'] == verify_returncode
    assert persisted['verify_verdict'] == job['verify_verdict']


def test_api_requires_bearer_token_when_configured(web_server, monkeypatch):
    monkeypatch.setenv('WEB_API_TOKEN', 'test-secret-token')

    missing = _request(web_server.app, 'GET', '/api/jobs')
    wrong = _request(
        web_server.app, 'GET', '/api/jobs',
        headers={'authorization': 'Bearer wrong-token'},
    )
    valid = _request(
        web_server.app, 'GET', '/api/jobs',
        headers={'authorization': 'Bearer test-secret-token'},
    )

    assert missing.status_code == 401
    assert missing.headers['www-authenticate'] == 'Bearer'
    assert wrong.status_code == 401
    assert valid.status_code == 200
    assert valid.json() == []


def test_api_rejects_non_loopback_client_without_configured_token(web_server):
    response = _request(
        web_server.app, 'GET', '/api/jobs', client_host='198.51.100.23',
    )

    assert response.status_code == 503
    assert 'WEB_API_TOKEN' in response.json()['detail']


def test_submission_rejected_when_active_queue_limit_reached(web_server, monkeypatch):
    monkeypatch.setattr(web_server, 'MAX_ACTIVE_JOBS', 1)
    web_server._jobs['11111111'] = {
        'id': '11111111', 'status': 'running', 'created_at': '2026-08-18 10:00:00',
    }

    response = _request(
        web_server.app, 'POST', '/api/jobs', json_body={'code': '300476'},
    )

    assert response.status_code == 429
    assert len(web_server._jobs) == 1
    assert web_server._queue.qsize() == 0


def test_web_llm_requires_explicit_opt_in(web_server):
    response = _request(
        web_server.app, 'POST', '/api/jobs',
        json_body={'code': '300476', 'llm': 'codex'},
    )

    assert response.status_code == 403
    assert response.json()['detail'].startswith('Web LLM 已禁用')
    assert web_server._jobs == {}


def test_oversized_request_body_is_rejected_before_json_parsing(web_server, monkeypatch):
    monkeypatch.setattr(web_server, 'MAX_REQUEST_BYTES', 1024)
    body = json.dumps({'code': '300476', 'unused': 'x' * 2048}).encode('utf-8')

    response = _request(web_server.app, 'POST', '/api/jobs', body=body,
                        headers={'content-type': 'application/json'})

    assert response.status_code == 413
    assert web_server._jobs == {}


def test_oversized_request_without_content_length_is_still_rejected(web_server, monkeypatch):
    monkeypatch.setattr(web_server, 'MAX_REQUEST_BYTES', 1024)
    body = json.dumps({'code': '300476', 'unused': 'x' * 2048}).encode('utf-8')

    response = _request(
        web_server.app, 'POST', '/api/jobs', body=body,
        headers={'content-type': 'application/json'}, include_content_length=False,
    )

    assert response.status_code == 413
    assert web_server._jobs == {}


def test_download_allows_verified_and_never_failed_validation(web_server, tmp_path,
                                                               monkeypatch):
    xlsx = tmp_path / 'model.xlsx'
    xlsx.write_bytes(b'xlsx-content')

    def install_job(status):
        web_server._jobs.clear()
        web_server._jobs['a1b2c3d4'] = {
            'id': 'a1b2c3d4', 'code': '300476', 'name': '测试公司',
            'status': status, 'xlsx': str(xlsx), 'has_xlsx': True,
            'params': {}, 'error': None, 'created_at': '2026-08-18 10:00:00',
            'started_at': None, 'finished_at': '2026-08-18 10:01:00',
            'verify_status': status, 'verify_returncode': 0 if status == 'verified' else None,
        }

    install_job('verified')
    verified = _request(web_server.app, 'GET', '/api/jobs/a1b2c3d4/download')
    install_job('built_unverified')
    unverified_default = _request(web_server.app, 'GET', '/api/jobs/a1b2c3d4/download')
    monkeypatch.setenv('WEB_ALLOW_UNVERIFIED_DOWNLOAD', '1')
    unverified_opt_in = _request(web_server.app, 'GET', '/api/jobs/a1b2c3d4/download')
    install_job('failed_validation')
    failed_validation = _request(web_server.app, 'GET', '/api/jobs/a1b2c3d4/download')

    assert verified.status_code == 200
    assert verified.content == b'xlsx-content'
    assert unverified_default.status_code == 409
    assert unverified_opt_in.status_code == 200
    assert failed_validation.status_code == 409


def test_frontend_stores_token_and_sends_bearer_header():
    html = Path(__file__).parents[1] / 'web' / 'static' / 'index.html'
    harness = r"""
const fs = require('fs');
const html = fs.readFileSync(process.argv[1], 'utf8');
const script = html.match(/<script>([\s\S]*?)<\/script>/)[1];
const ids = new Set([...html.matchAll(/\bid="([^"]+)"/g)].map(m => m[1]));
const elements = {};
for (const id of ids) {
  elements[id] = {
    id, value: '', checked: false, disabled: false, textContent: '', innerHTML: '',
    handlers: {},
    addEventListener(type, fn) { this.handlers[type] = fn; },
  };
}
global.document = {
  hidden: true,
  getElementById(id) { return elements[id] || null; },
  addEventListener() {},
};
const storage = {'astock-dcf-api-token': 'saved-token'};
global.localStorage = {
  getItem(k) { return storage[k] || null; },
  setItem(k, v) { storage[k] = String(v); },
  removeItem(k) { delete storage[k]; },
};
const calls = [];
global.fetch = async (path, opts = {}) => {
  calls.push({path, opts});
  return {ok: true, status: 200, statusText: 'OK', json: async () => []};
};
global.confirm = () => true;
global.alert = () => {};
eval(script);
(async () => {
  await new Promise(resolve => setImmediate(resolve));
  if (!elements.apiToken || !elements.saveToken) throw new Error('token controls missing');
  if (elements.apiToken.value !== 'saved-token') throw new Error('stored token not restored');
  if (calls[0].opts.headers.Authorization !== 'Bearer saved-token') {
    throw new Error('initial API request missing saved Bearer token');
  }
  elements.apiToken.value = 'new-token';
  elements.saveToken.handlers.click();
  await new Promise(resolve => setImmediate(resolve));
  if (storage['astock-dcf-api-token'] !== 'new-token') throw new Error('token not stored');
  const last = calls[calls.length - 1];
  if (last.opts.headers.Authorization !== 'Bearer new-token') {
    throw new Error('API request missing updated Bearer token');
  }
})().catch(err => { console.error(err.stack || err); process.exitCode = 1; });
"""

    result = subprocess.run(
        ['node', '-e', harness, str(html)], capture_output=True, text=True, timeout=10,
    )

    assert result.returncode == 0, result.stderr
