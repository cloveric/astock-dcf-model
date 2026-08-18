# -*- coding: utf-8 -*-
"""Focused verification controls added after the full model/CI audit."""

import hashlib

import pytest

import verify_model as vm


def _missing_control(*_args, **_kwargs):
    return {"status": "MISSING", "detail": "control is not implemented"}


def _control(name):
    """Keep the first RED run an assertion failure when the API is absent."""
    return getattr(vm, name, _missing_control)


def test_default_golden_hash_drift_is_failure_unless_explicitly_accepted(tmp_path, monkeypatch):
    cfg = tmp_path / "configs" / "999999.yaml"
    cfg.parent.mkdir()
    cfg.write_text("version: changed\n", encoding="utf-8")
    old_hash = hashlib.sha256(b"version: frozen\n").hexdigest()
    current_hash = hashlib.sha256(cfg.read_bytes()).hexdigest()
    monkeypatch.setitem(vm.GOLDEN_DCF_PS, "999999", {
        "value": 12.5,
        "config_sha256": old_hash,
    })
    meta = {
        "config_path": str(cfg),
        "config_sha256": current_hash,
        "is_default_config": True,
    }

    result = _control("_golden_control")(
        "999999", 12.5, meta, repo_root=str(tmp_path), accept_config_change=False,
    )
    assert result["status"] == "FAIL"

    accepted = _control("_golden_control")(
        "999999", 12.5, meta, repo_root=str(tmp_path), accept_config_change=True,
    )
    assert accepted["status"] == "WAIVED"


def test_custom_config_never_uses_default_golden_even_when_code_matches(tmp_path, monkeypatch):
    custom = tmp_path / "custom.yaml"
    custom.write_text("custom: true\n", encoding="utf-8")
    digest = hashlib.sha256(custom.read_bytes()).hexdigest()
    monkeypatch.setitem(vm.GOLDEN_DCF_PS, "999999", {
        "value": 12.5,
        "config_sha256": "0" * 64,
    })

    result = _control("_golden_control")(
        "999999",
        987.6,
        {"config_path": str(custom), "config_sha256": digest, "is_default_config": False},
        repo_root=str(tmp_path),
    )
    assert result["status"] == "NOT_APPLICABLE"


def test_golden_value_mismatch_fails_when_default_fingerprint_matches(tmp_path, monkeypatch):
    cfg = tmp_path / "configs" / "999999.yaml"
    cfg.parent.mkdir()
    cfg.write_text("version: frozen\n", encoding="utf-8")
    digest = hashlib.sha256(cfg.read_bytes()).hexdigest()
    monkeypatch.setitem(vm.GOLDEN_DCF_PS, "999999", {
        "value": 12.5,
        "config_sha256": digest,
    })

    result = _control("_golden_control")(
        "999999",
        12.0,
        {"config_path": str(cfg), "config_sha256": digest, "is_default_config": True},
        repo_root=str(tmp_path),
    )
    assert result["status"] == "FAIL"


def test_default_golden_rejects_addr_fingerprint_that_disagrees_with_source_file(tmp_path, monkeypatch):
    cfg = tmp_path / "configs" / "999999.yaml"
    cfg.parent.mkdir()
    cfg.write_text("actual: changed\n", encoding="utf-8")
    frozen_hash = hashlib.sha256(b"actual: frozen\n").hexdigest()
    monkeypatch.setitem(vm.GOLDEN_DCF_PS, "999999", {
        "value": 12.5,
        "config_sha256": frozen_hash,
    })

    result = _control("_golden_control")(
        "999999",
        12.5,
        {"config_path": str(cfg), "config_sha256": frozen_hash, "is_default_config": True},
        repo_root=str(tmp_path),
    )
    assert result["status"] == "FAIL"


@pytest.mark.parametrize(
    ("wacc", "terminal_g", "want"),
    [
        (0.08, 0.03, "PASS"),
        (0.03, 0.03, "FAIL"),
        (0.034, 0.03, "FAIL"),
        (None, 0.03, "FAIL"),
    ],
)
def test_wacc_requires_positive_values_and_at_least_50bp_terminal_spread(wacc, terminal_g, want):
    result = _control("_wacc_terminal_control")(wacc, terminal_g, min_spread=0.005)
    assert result["status"] == want


def test_scenario_values_must_be_monotonic_and_same_engine_when_declared():
    control = _control("_scenario_control")
    assert control(80.0, 100.0, 120.0, "simplified_same_engine_v1")["status"] == "PASS"
    assert control(105.0, 100.0, 120.0, "simplified_same_engine_v1")["status"] == "FAIL"
    assert control(80.0, 100.0, 120.0, "mixed_engine_v0")["status"] == "REVIEW"


def test_relative_median_must_equal_median_of_pricing_rows():
    control = _control("_relative_median_control")
    assert control(20.0, [10.0, 20.0, 30.0])["status"] == "PASS"
    assert control(30.0, [10.0, 20.0, 30.0])["status"] == "FAIL"
    assert control(None, [10.0, 20.0, 30.0])["status"] == "FAIL"
    assert control(None, [])["status"] == "NOT_APPLICABLE"


def test_relative_display_uses_true_median_only_when_pricing_rows_exist():
    choose = _control("_relative_median_ref")
    assert choose({
        "relative_price_rows": ["Relative_Val!F5"],
        "rel_median_cell": "Relative_Val!F9",
        "rel_med_pe": "Relative_Val!C12",
    }) == "Relative_Val!F9"
    assert choose({
        "relative_price_rows": [],
        "rel_median_cell": "Relative_Val!F5",
        "rel_med_pe": "Relative_Val!C9",
    }) == "Relative_Val!C9"


def test_gross_margin_hard_bounds_reject_impossible_percentages():
    control = _control("_gross_margin_control")
    assert control([-0.2, 0.25, 1.0])["status"] == "PASS"
    assert control([0.2, 1.01])["status"] == "FAIL"
    assert control([-1.01, 0.2])["status"] == "FAIL"
    assert control([0.2, None])["status"] == "FAIL"


def test_historical_plug_verifier_uses_same_rounding_threshold_as_builder():
    control = _control("_historical_plug_control")
    tiny = [{"year": 2025, "diff": 0.1, "assets": 20000.0, "ratio": 0.000005}]
    impossible = [{"year": 2025, "diff": 5.0, "assets": 2000.0, "ratio": 0.0025}]
    assert control(tiny)["status"] == "PASS"
    assert control(impossible)["status"] == "FAIL"
    assert control([])["status"] == "PASS"


@pytest.mark.parametrize(
    ("meta", "want"),
    [
        ({"currency_code": "CNY", "per_share_unit": "元/股"}, "元/股"),
        ({"currency_code": "USD", "per_share_unit": "美元/股"}, "美元/股"),
        ({"currency_code": "HKD"}, "港元/股"),
        ({"currency_note": "美元(财报口径) / 百万元"}, "美元/股"),
    ],
)
def test_currency_output_unit_uses_addr_metadata(meta, want):
    assert _control("_per_share_unit")(meta) == want


def test_fcfe_divergence_is_review_gate_and_explicit_waiver_is_visible():
    control = _control("_fcfe_divergence_control")
    assert control(0.299, waived=False)["status"] == "PASS"
    assert control(0.30, waived=False)["status"] == "REVIEW"
    assert control(-0.45, waived=True)["status"] == "WAIVED"


def test_fcfe_divergence_accepts_only_nonempty_reviewed_metadata_waiver():
    control = _control("_fcfe_divergence_control")
    reviewed = control(-0.65, waiver_reason="已复核次新股融资路径; FCFE仅作辅助视图")
    blank = control(-0.65, waiver_reason="  ")
    assert reviewed["status"] == "WAIVED"
    assert "次新股融资路径" in reviewed["detail"]
    assert blank["status"] == "REVIEW"


def test_summary_exit_codes_distinguish_fail_review_and_all_pass():
    summarize = _control("_summarize_controls")
    passed = summarize([{"status": "PASS"}, {"status": "NOT_APPLICABLE"}])
    review = summarize([{"status": "PASS"}, {"status": "REVIEW"}])
    failed = summarize([{"status": "REVIEW"}, {"status": "FAIL"}])
    assert (passed["verdict"], passed["exit_code"], passed["label"]) == ("PASS", 0, "ALL PASS")
    assert (review["verdict"], review["exit_code"], review["label"]) == ("REVIEW", 2, "需要人工复核")
    assert (failed["verdict"], failed["exit_code"], failed["label"]) == ("FAIL", 1, "存在FAIL项")


def test_verification_payload_is_stable_for_web_consumers():
    controls = [{"name": "WACC", "status": "PASS", "detail": "ok"}]
    payload = _control("_verification_payload")(
        "model.xlsx", {"code": "300476", "currency_code": "CNY"}, controls,
    )
    assert payload == {
        "schema_version": 1,
        "file": "model.xlsx",
        "model": {"code": "300476", "currency_code": "CNY"},
        "verdict": "PASS",
        "exit_code": 0,
        "label": "ALL PASS",
        "controls": controls,
    }
