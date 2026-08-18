"""LLM CLI routing regression tests (no network, no real model calls)."""

import build_model
from research import llm as llm_mod


def _fake_cli(bin_dir, name, script):
    path = bin_dir / name
    path.write_text(f'#!/bin/sh\n{script}\n', encoding='utf-8')
    path.chmod(0o755)


def test_auto_llm_falls_back_in_codex_claude_kimi_order(tmp_path, monkeypatch):
    """Wrong provider order or stopping on failed/empty output must fail this test."""
    bin_dir = tmp_path / 'bin'
    bin_dir.mkdir()
    log = tmp_path / 'calls.log'
    monkeypatch.setenv('PATH', str(bin_dir))
    monkeypatch.setenv('LLM_TEST_LOG', str(log))
    _fake_cli(bin_dir, 'codex', 'echo codex >> "$LLM_TEST_LOG"; exit 7')
    _fake_cli(bin_dir, 'claude', 'echo claude >> "$LLM_TEST_LOG"; exit 0')
    _fake_cli(bin_dir, 'kimi', 'echo kimi >> "$LLM_TEST_LOG"; echo "Kimi memo"')

    memo, note = llm_mod.gen_memo('auto', 'test prompt')

    assert memo == 'Kimi memo'
    assert log.read_text(encoding='utf-8').splitlines() == ['codex', 'claude', 'kimi']
    assert 'kimi' in note


def test_bare_llm_flag_means_auto_but_omission_remains_off():
    """A bare --llm must opt in to auto routing without changing the safe no-flag path."""
    parser = build_model._argument_parser()

    assert parser.parse_args(['--code', '300476']).llm == 'off'
    assert parser.parse_args(['--code', '300476', '--llm']).llm == 'auto'
    assert parser.parse_args(['--code', '300476', '--llm', 'kimi']).llm == 'kimi'
