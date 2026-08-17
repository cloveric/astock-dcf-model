# -*- coding: utf-8 -*-
"""
Web 服务模式: FastAPI + 单页前端 (web/static/index.html, 零构建)

启动:
    python -m web.server            # 默认 127.0.0.1:8000
    uvicorn web.server:app --port 8000

接口:
    POST /api/jobs                 提交建模任务 {code, config?, dr?, consensus?, announcements?, llm?}
    GET  /api/jobs                 历史任务列表 (按创建时间倒序)
    GET  /api/jobs/{id}            任务详情 (状态/日志尾部/产物路径)
    GET  /api/jobs/{id}/download   下载 xlsx 产物

实现要点:
  - 建模逻辑完全复用 build_model.py (子进程调用, 不重写); 单 worker 线程串行执行;
  - 任务状态落盘 web/.data/jobs.json (原子写), xlsx/addr/日志存 web/.data/out/<job_id>/;
  - 服务重启时, 此前 running/queued 的任务标记为 failed(中断);
  - config/dr/consensus 仅接受仓库内已存在的文件路径, 防路径穿越。
"""
import json
import os
import queue
import re
import subprocess
import sys
import threading
import time
import uuid
from typing import Optional

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

WEB_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(WEB_DIR)
DATA_DIR = os.path.join(WEB_DIR, '.data')
OUT_DIR = os.path.join(DATA_DIR, 'out')
JOBS_JSON = os.path.join(DATA_DIR, 'jobs.json')
STATIC_DIR = os.path.join(WEB_DIR, 'static')

CODE_RE = re.compile(r'^\d{5,6}$')          # 6位A股 或 5位港股
VALID_LLM = {'off', 'auto', 'claude', 'codex'}

_lock = threading.Lock()
_jobs = {}                                   # id -> job dict
_queue = queue.Queue()


def _now():
    return time.strftime('%Y-%m-%d %H:%M:%S')


def _save():
    tmp = JOBS_JSON + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(list(_jobs.values()), f, ensure_ascii=False, indent=1)
    os.replace(tmp, JOBS_JSON)


def _load():
    global _jobs
    if os.path.exists(JOBS_JSON):
        with open(JOBS_JSON, encoding='utf-8') as f:
            for j in json.load(f):
                if j['status'] in ('queued', 'running'):
                    j['status'] = 'failed'
                    j['error'] = '服务重启, 任务中断'
                    j['finished_at'] = _now()
                _jobs[j['id']] = j


def _repo_file(rel, must_exist=True):
    """校验仓库内相对路径, 返回绝对路径; 不合法抛 HTTP 400"""
    if not rel:
        return None
    p = os.path.realpath(os.path.join(REPO_ROOT, rel))
    if not p.startswith(REPO_ROOT + os.sep):
        raise HTTPException(400, f'路径必须位于仓库内: {rel}')
    if must_exist and not os.path.isfile(p):
        raise HTTPException(400, f'文件不存在: {rel}')
    return p


def _read_log(job, tail=4000):
    lf = job.get('log_file')
    if lf and os.path.exists(lf):
        with open(lf, encoding='utf-8', errors='replace') as f:
            return f.read()[-tail:]
    return ''


def _run_job(job):
    jid = job['id']
    job_dir = os.path.join(OUT_DIR, jid)
    os.makedirs(job_dir, exist_ok=True)
    xlsx = os.path.join(job_dir, 'model.xlsx')
    addr = os.path.join(job_dir, 'model.addr.json')
    log_file = os.path.join(job_dir, 'build.log')
    cmd = [sys.executable, os.path.join(REPO_ROOT, 'build_model.py'),
           '--code', job['code'], '--out', xlsx, '--addr', addr]
    if job['params'].get('config'):
        cmd += ['--config', job['params']['config']]
    for flag, arg in (('dr', '--dr'), ('consensus', '--consensus'), ('llm', '--llm')):
        if job['params'].get(flag):
            cmd += [arg, job['params'][flag]]
    if job['params'].get('announcements'):
        cmd.append('--announcements')
    with _lock:
        job['status'] = 'running'
        job['started_at'] = _now()
        job['log_file'] = log_file
        _save()
    try:
        with open(log_file, 'w', encoding='utf-8') as lf:
            lf.write('$ ' + ' '.join(cmd) + '\n\n')
            lf.flush()
            proc = subprocess.run(cmd, cwd=REPO_ROOT, stdout=lf, stderr=subprocess.STDOUT,
                                  timeout=600)
        if proc.returncode == 0 and os.path.exists(xlsx):
            name = job['code']
            try:
                meta = json.load(open(addr, encoding='utf-8')).get('meta', {})
                name = meta.get('name', name)
            except Exception:
                pass
            with _lock:
                job['status'] = 'done'
                job['name'] = name
                job['xlsx'] = xlsx
                job['finished_at'] = _now()
                _save()
        else:
            with _lock:
                job['status'] = 'failed'
                job['error'] = f'build_model.py 退出码 {proc.returncode}'
                job['finished_at'] = _now()
                _save()
    except Exception as e:                                   # 含超时
        with _lock:
            job['status'] = 'failed'
            job['error'] = f'{type(e).__name__}: {e}'
            job['finished_at'] = _now()
            _save()


def _worker():
    while True:
        jid = _queue.get()
        with _lock:
            job = _jobs.get(jid)
        if job and job['status'] == 'queued':
            _run_job(job)
        _queue.task_done()


class JobIn(BaseModel):
    code: str
    config: Optional[str] = None
    dr: Optional[str] = None
    consensus: Optional[str] = None
    announcements: bool = False
    llm: str = 'off'


def _public(job, with_log=False):
    d = {k: job.get(k) for k in ('id', 'code', 'name', 'status', 'error',
                                 'created_at', 'started_at', 'finished_at')}
    d['params'] = {k: v for k, v in job['params'].items() if v not in (None, False, 'off')}
    d['has_xlsx'] = bool(job.get('xlsx') and os.path.exists(job['xlsx']))
    if with_log:
        d['log_tail'] = _read_log(job)
    return d


def create_app():
    os.makedirs(OUT_DIR, exist_ok=True)
    _load()
    threading.Thread(target=_worker, daemon=True).start()

    app = FastAPI(title='astock-dcf-model web', docs_url=None, redoc_url=None)

    @app.post('/api/jobs')
    def submit(body: JobIn):
        code = body.code.strip()
        if not CODE_RE.match(code):
            raise HTTPException(400, f'代码须为6位A股或5位港股数字: {body.code}')
        if body.llm not in VALID_LLM:
            raise HTTPException(400, f'llm 仅支持 {sorted(VALID_LLM)}')
        params = {'config': _repo_file(body.config) if body.config else None,
                  'dr': _repo_file(body.dr) if body.dr else None,
                  'consensus': _repo_file(body.consensus) if body.consensus else None,
                  'announcements': bool(body.announcements),
                  'llm': body.llm}
        job = {'id': uuid.uuid4().hex[:8], 'code': code, 'name': code,
               'status': 'queued', 'error': None, 'created_at': _now(),
               'started_at': None, 'finished_at': None, 'params': params,
               'xlsx': None, 'log_file': None}
        with _lock:
            _jobs[job['id']] = job
            _save()
        _queue.put(job['id'])
        return _public(job)

    @app.get('/api/jobs')
    def list_jobs():
        with _lock:
            jobs = sorted(_jobs.values(), key=lambda j: j['created_at'], reverse=True)
            return [_public(j) for j in jobs]

    @app.get('/api/jobs/{jid}')
    def get_job(jid: str):
        job = _jobs.get(jid)
        if not job:
            raise HTTPException(404, '任务不存在')
        return _public(job, with_log=True)

    @app.get('/api/jobs/{jid}/download')
    def download(jid: str):
        job = _jobs.get(jid)
        if not job:
            raise HTTPException(404, '任务不存在')
        if job['status'] != 'done' or not job.get('xlsx') or not os.path.exists(job['xlsx']):
            raise HTTPException(409, '产物不可用(任务未完成或文件缺失)')
        fname = f"{job['code']}_{job.get('name') or job['code']}_估值模型.xlsx"
        return FileResponse(job['xlsx'], filename=fname,
                            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    app.mount('/', StaticFiles(directory=STATIC_DIR, html=True), name='static')
    return app


app = create_app()


def main():
    import uvicorn
    host = os.environ.get('HOST', '127.0.0.1')
    port = int(os.environ.get('PORT', '8000'))
    print(f'astock-dcf-model web: http://{host}:{port}')
    uvicorn.run(app, host=host, port=port, log_level='warning')


if __name__ == '__main__':
    main()
