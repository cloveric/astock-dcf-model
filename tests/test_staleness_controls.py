import copy
from pathlib import Path

import pytest
import yaml

import build_model as bm


ROOT = Path(__file__).resolve().parents[1]


def _write_config(tmp_path, *, valuation_date="2026-01-01", interim=None):
    raw = yaml.safe_load((ROOT / "configs" / "300476.yaml").read_text(encoding="utf-8"))
    raw["model"]["valuation_date"] = valuation_date
    if interim is None:
        raw.pop("interim", None)
    else:
        raw["interim"] = interim
    path = tmp_path / "config.yaml"
    path.write_text(yaml.safe_dump(raw, allow_unicode=True, sort_keys=False), encoding="utf-8")
    return path


def test_stale_config_warns_by_default_and_can_be_a_hard_gate(tmp_path, capsys):
    path = _write_config(tmp_path)

    bm.load_config(
        config_path=str(path), allow_fallback=False, as_of="2026-08-18",
        stale_after_days=30,
    )
    assert "配置已过期" in capsys.readouterr().err

    with pytest.raises(ValueError, match="配置已过期"):
        bm.load_config(
            config_path=str(path), allow_fallback=False, as_of="2026-08-18",
            stale_after_days=30, fail_on_stale=True,
        )


def test_as_of_rejects_look_ahead_configuration(tmp_path):
    path = _write_config(tmp_path, valuation_date="2026-08-19")
    with pytest.raises(ValueError, match="晚于 --as-of"):
        bm.load_config(
            config_path=str(path), allow_fallback=False, as_of="2026-08-18",
        )


def test_as_of_requires_canonical_calendar_date(tmp_path):
    path = _write_config(tmp_path)
    with pytest.raises(ValueError, match="YYYY-MM-DD"):
        bm.load_config(
            config_path=str(path), allow_fallback=False, as_of="20260818",
        )


def test_require_interim_distinguishes_missing_or_failed_anchor(tmp_path):
    missing = _write_config(tmp_path)
    with pytest.raises(ValueError, match="require-interim"):
        bm.load_config(
            config_path=str(missing), allow_fallback=False, require_interim=True,
        )

    failed = _write_config(
        tmp_path,
        interim={"anchor": False, "fetch_status": "error", "basis": "upstream failed"},
    )
    with pytest.raises(ValueError, match="require-interim"):
        bm.load_config(
            config_path=str(failed), allow_fallback=False, require_interim=True,
        )


def test_refresh_explicitly_rebuilds_public_fallback_in_memory(monkeypatch):
    expected = yaml.safe_load(
        (ROOT / "configs" / "300476.yaml").read_text(encoding="utf-8")
    )
    expected["company"]["name"] = "fresh"
    monkeypatch.setattr("fetch_data.build_fallback_config", lambda code: copy.deepcopy(expected))

    cfg = bm.load_config(code="300476", refresh=True, allow_fallback=False)

    assert cfg.raw == expected
