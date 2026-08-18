"""LLM研究备忘录: 调用本机 codex / claude / kimi CLI; 不存在或失败则优雅降级。"""

import shutil
import subprocess


_AUTO_ORDER = ('codex', 'claude', 'kimi')


def _which(cli):
    return shutil.which(cli) is not None


def gen_memo(kind, prompt, timeout=180):
    """kind ∈ auto|codex|claude|kimi → (memo文本|None, 说明note)。"""
    candidates = _AUTO_ORDER if kind == 'auto' else (kind,)
    failures = []
    for cli in candidates:
        if cli not in _AUTO_ORDER or not _which(cli):
            failures.append(f'{cli}:无可用CLI')
            continue
        if cli == 'codex':
            cmd = ['codex', 'exec', '--skip-git-repo-check', prompt]
        elif cli == 'claude':
            cmd = ['claude', '-p', prompt]
        else:
            cmd = ['kimi', '-p', prompt]
        try:
            out = subprocess.run(
                cmd, capture_output=True, timeout=timeout, text=True, check=False,
            )
        except (OSError, subprocess.SubprocessError) as e:
            failures.append(f'{cli}:调用失败({e})')
            continue
        txt = (out.stdout or '').strip()
        if out.returncode != 0:
            # 非0退出码: stdout可能是CLI报错横幅, 绝不能当备忘录。
            err = ' '.join((out.stderr or '').strip().split())[:160] or '无stderr输出'
            print(f'警告: llm={cli} CLI退出码{out.returncode}, stderr摘要: {err}')
            failures.append(f'{cli}:退出码{out.returncode}({err[:80]})')
            continue
        if not txt:
            failures.append(f'{cli}:CLI无输出')
            continue
        route = f'auto→{cli}' if kind == 'auto' else cli
        return txt, f'llm={route}: 备忘录已生成({len(txt)}字)'
    detail = '; '.join(failures) or '无可用CLI'
    return None, f'llm={kind}: {detail}, 已跳过(不影响建模)'
