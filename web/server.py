# -*- coding: utf-8 -*-
"""
Web 服务模式: FastAPI + 单页前端 (web/static/index.html, 零构建)

启动:
    python -m web.server            # 默认 127.0.0.1:8000
    uvicorn web.server:app --port 8000

接口:
    POST   /api/jobs                 提交建模任务 {code, config?, dr?, consensus?, announcements?, llm?}
    GET    /api/jobs                 历史任务列表 (按创建时间倒序, 最多保留100个)
    GET    /api/jobs/{id}            任务详情 (状态/日志尾部/验收摘要)
    GET    /api/jobs/{id}/download   下载 xlsx 产物
    DELETE /api/jobs/{id}            删除任务并清理产物 (进行中任务拒绝)

实现要点:
  - 建模逻辑完全复用 build_model.py (子进程调用, 不重写); 单 worker 线程串行执行;
  - 任务状态落盘 web/.data/jobs.json (原子写), xlsx/addr/日志存 web/.data/out/<job_id>/;
  - 服务重启时, 此前 running/queued 的任务标记为 failed(中断);
  - config/dr/consensus 仅接受仓库内已存在的文件路径, 防路径穿越;
  - 任务保留上限 MAX_JOBS=100, 超限时淘汰最早完成的任务并同步清理产物目录;
  - 构建成功后若本机有 LibreOffice 自动附跑 verify_model.py, 摘要随详情返回;
  - 无鉴权: 仅绑定可信网络/本机使用, 公网暴露前须自行加反向代理鉴权。
"""
import fcntl
import json
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import time
import traceback
import uuid
from typing import Optional
from urllib.parse import quote

from fastapi import FastAPI, HTTPException
from fastapi.responses import Response
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

WEB_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(WEB_DIR)
DATA_DIR = os.path.join(WEB_DIR, '.data')
OUT_DIR = os.path.join(DATA_DIR, 'out')
JOBS_JSON = os.path.join(DATA_DIR, 'jobs.json')
STATIC_DIR = os.path.join(WEB_DIR, 'static')
LOCK_FILE = os.path.join(DATA_DIR, 'server.lock')

if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)
from fetch_data import em_code, is_hk  # noqa: E402  (代码语义校验复用数据层规则)

# _repo_file 路径校验的基准: REPO_ROOT 本身也可能位于符号链接下(如 macOS /var -> /private/var),
# 统一 realpath 后再比较, 否则合法路径会被误判 400
_REPO_REAL = os.path.realpath(REPO_ROOT)

CODE_RE = re.compile(r'^\d{5,6}$')          # 6位A股 或 5位港股
VALID_LLM = {'off', 'auto', 'claude', 'codex'}
MAX_JOBS = 100                              # 任务保留上限, 超限淘汰最早完成的任务(含产物)

_lock = threading.Lock()
_jobs = {}                                   # id -> job dict
_queue = queue.Queue()
_instance_lock_fh = None                     # flock 句柄, 进程存活期间持有不释放


def _acquire_instance_lock():
    """强制单进程: _jobs/_queue/jobs.json 均为进程内状态, 多进程 (如 uvicorn --workers>1
    或重复启动) 会各持一份内存态并竞写 jobs.json, 造成脑裂/互相覆盖.
    这里用 fcntl.flock 对 web/.data/server.lock 加排它非阻塞锁, 句柄存模块级变量,
    进程退出时由内核自动释放; 拿不到锁说明已有实例在跑, 直接拒绝启动."""
    global _instance_lock_fh
    if _instance_lock_fh is not None:        # 同进程内重复调用 (如测试多次 create_app)
        return
    os.makedirs(DATA_DIR, exist_ok=True)
    fh = open(LOCK_FILE, 'w')
    try:
        fcntl.flock(fh, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError:
        fh.close()
        raise RuntimeError('检测到另一个 web 服务实例(仅支持单进程,勿用 --workers>1)')
    _instance_lock_fh = fh


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
                # 启动时一次性核对产物是否仍在, 运行期 _public 直接读该字段 (避免轮询 stat)
                j['has_xlsx'] = bool(j.get('xlsx') and os.path.exists(j['xlsx']))
                _jobs[j['id']] = j


def _repo_file(rel, must_exist=True):
    """校验仓库内相对路径, 返回绝对路径; 不合法抛 HTTP 400 (两侧均 realpath 后比较)"""
    if not rel:
        return None
    p = os.path.realpath(os.path.join(REPO_ROOT, rel))
    try:
        inside = os.path.commonpath([p, _REPO_REAL]) == _REPO_REAL
    except ValueError:
        inside = False
    if not inside:
        raise HTTPException(400, f'路径必须位于仓库内: {rel}')
    if must_exist and not os.path.isfile(p):
        raise HTTPException(400, f'文件不存在: {rel}')
    return p


def _read_log(job, tail=4000):
    """只读日志尾部 tail 字节 (seek 定位), 避免大日志整读进内存"""
    lf = job.get('log_file')
    if not lf:
        return ''
    try:
        size = os.path.getsize(lf)
        with open(lf, 'rb') as f:
            f.seek(max(0, size - tail))
            return f.read().decode('utf-8', errors='replace')
    except OSError:                          # 文件不存在/已被清理
        return ''


def _job_dir(jid):
    return os.path.join(OUT_DIR, jid)


def _drop_job(jid):
    """移除任务记录并同步清理产物目录 (调用方须持锁)"""
    if _jobs.pop(jid, None) is not None:
        shutil.rmtree(_job_dir(jid), ignore_errors=True)


def _enforce_retention():
    """任务总数超MAX_JOBS时, 淘汰最早完成的任务(产物目录一并删除); 进行中任务不淘汰 (调用方须持锁)"""
    while len(_jobs) > MAX_JOBS:
        finished = [j for j in _jobs.values() if j['status'] in ('done', 'failed')]
        if not finished:
            break
        victim = min(finished, key=lambda j: j['finished_at'] or j['created_at'])
        _drop_job(victim['id'])


def _run_verify(xlsx, addr, log_file):
    """构建成功后附跑LibreOffice重算验收; 本机无soffice时跳过并说明. 返回摘要文本"""
    if not shutil.which('soffice'):
        return '本机未检测到LibreOffice(soffice), 跳过自动验收; 可手工执行: python verify_model.py <xlsx>'
    try:
        proc = subprocess.run([sys.executable, os.path.join(REPO_ROOT, 'verify_model.py'),
                               xlsx, '--addr', addr,
                               '--workdir', os.path.join(os.path.dirname(xlsx), '.lo')],
                              cwd=REPO_ROOT, capture_output=True, text=True, timeout=600)
        tail = ((proc.stdout or '') + (proc.stderr or ''))[-2000:]
    except Exception as e:
        tail = f'验收执行失败: {type(e).__name__}: {e}'
    try:
        with open(log_file, 'a', encoding='utf-8') as lf:
            lf.write('\n$ verify_model.py (自动验收)\n' + tail + '\n')
    except Exception:
        pass
    return tail


def _run_job(job):
    jid = job['id']
    try:
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
            verify_tail = _run_verify(xlsx, addr, log_file)
            with _lock:
                job['status'] = 'done'
                job['name'] = name
                job['xlsx'] = xlsx
                job['has_xlsx'] = True
                job['verify_tail'] = verify_tail
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
    finally:
        # 兜底: 无论上面哪一步失败 (含 except 分支里 _save 再抛), 任务都不得停留在
        # running/queued, 否则前端永远显示进行中且 DELETE 拒删
        with _lock:
            if job['status'] in ('queued', 'running'):
                job['status'] = 'failed'
                job['error'] = job.get('error') or '任务异常终止(服务内部错误)'
                job['finished_at'] = _now()
                try:
                    _save()
                except Exception:
                    traceback.print_exc()


def _worker():
    while True:
        try:
            jid = _queue.get()
            with _lock:
                job = _jobs.get(jid)
            if job and job['status'] == 'queued':
                _run_job(job)
            _queue.task_done()
        except Exception:
            # worker 是唯一执行线程, 任何异常都只记录日志后继续取下一个任务, 线程永不退出
            traceback.print_exc()
            continue


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
    # 读缓存字段而非每次 stat (列表接口会被前端轮询, N个任务N次stat); download 仍做真实检查
    d['has_xlsx'] = bool(job.get('has_xlsx'))
    if with_log:
        d['log_tail'] = _read_log(job)
        if job.get('verify_tail'):
            d['verify_tail'] = job['verify_tail']
    return d


def create_app():
    _acquire_instance_lock()                 # 单进程强约束, 见函数注释
    os.makedirs(OUT_DIR, exist_ok=True)
    _load()
    threading.Thread(target=_worker, daemon=True).start()

    app = FastAPI(title='astock-dcf-model web', docs_url=None, redoc_url=None)

    @app.post('/api/jobs')
    def submit(body: JobIn):
        code = body.code.strip()
        if not CODE_RE.match(code):
            raise HTTPException(400, f'代码须为6位A股或5位港股数字: {body.code}')
        if not is_hk(code):
            try:
                em_code(code)                # 语义校验: 6位但非沪深北合法号段 → 拒绝
            except Exception:
                raise HTTPException(400, f'无法识别的交易所/代码: {code}')
        if body.llm not in VALID_LLM:
            raise HTTPException(400, f'llm 仅支持 {sorted(VALID_LLM)}')
        params = {'config': _repo_file(body.config) if body.config else None,
                  'dr': _repo_file(body.dr) if body.dr else None,
                  'consensus': _repo_file(body.consensus) if body.consensus else None,
                  'announcements': bool(body.announcements),
                  'llm': body.llm}
        with _lock:
            jid = uuid.uuid4().hex[:8]
            while jid in _jobs:              # 8位hex偶发碰撞会静默覆盖旧任务, 碰撞则重生成
                jid = uuid.uuid4().hex[:8]
            job = {'id': jid, 'code': code, 'name': code,
                   'status': 'queued', 'error': None, 'created_at': _now(),
                   'started_at': None, 'finished_at': None, 'params': params,
                   'xlsx': None, 'log_file': None}
            _jobs[jid] = job
            _enforce_retention()
            _save()
        _queue.put(jid)
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

    @app.delete('/api/jobs/{jid}')
    def delete_job(jid: str):
        with _lock:
            job = _jobs.get(jid)
            if not job:
                raise HTTPException(404, '任务不存在')
            if job['status'] in ('queued', 'running'):
                raise HTTPException(409, '任务进行中, 不可删除')
            _drop_job(jid)
            _save()
        return {'deleted': jid}

    @app.get('/api/jobs/{jid}/download')
    def download(jid: str):
        # 持锁把文件读成 bytes 再返回: FileResponse 是响应期异步读盘, 与 _drop_job 的
        # rmtree (删除/淘汰) 存在竞态会致下载中途 500; xlsx 仅数MB, 整读可接受
        with _lock:
            job = _jobs.get(jid)
            if not job:
                raise HTTPException(404, '任务不存在')
            if job['status'] != 'done' or not job.get('xlsx'):
                raise HTTPException(409, '产物不可用(任务未完成或文件缺失)')
            try:
                with open(job['xlsx'], 'rb') as f:
                    content = f.read()
            except OSError:
                raise HTTPException(404, '产物文件不存在')
            fname = f"{job['code']}_{job.get('name') or job['code']}_估值模型.xlsx"
        disp = "attachment; filename=\"model.xlsx\"; filename*=UTF-8''" + quote(fname)
        return Response(content=content,
                        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        headers={'Content-Disposition': disp})

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
