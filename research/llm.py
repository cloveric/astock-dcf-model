# -*- coding: utf-8 -*-
"""LLM研究备忘录: 调用本机 claude / codex CLI 生成一段研究文本; 不存在或失败则优雅降级"""
import shutil
import subprocess


def _which(cli):
    return shutil.which(cli) is not None


def gen_memo(kind, prompt, timeout=180):
    """kind ∈ auto|claude|codex → (memo文本|None, 说明note)"""
    cli = None
    if kind == 'auto':
        for c in ('claude', 'codex'):
            if _which(c):
                cli = c
                break
    elif kind in ('claude', 'codex'):
        cli = kind if _which(kind) else None
    if not cli:
        return None, f'llm={kind}: 本机无可用CLI, 已跳过(不影响建模)'
    cmd = (['claude', '-p', prompt] if cli == 'claude'
           else ['codex', 'exec', '--skip-git-repo-check', prompt])
    try:
        out = subprocess.run(cmd, capture_output=True, timeout=timeout, text=True)
        txt = (out.stdout or '').strip()
        if txt:
            return txt, f'llm={cli}: 备忘录已生成({len(txt)}字)'
        return None, f'llm={cli}: CLI无输出, 已跳过(不影响建模)'
    except Exception as e:
        return None, f'llm={cli}: 调用失败({e}), 已跳过(不影响建模)'
