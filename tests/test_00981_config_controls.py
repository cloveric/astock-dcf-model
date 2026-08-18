from pathlib import Path

import yaml


ROOT = Path(__file__).resolve().parents[1]


def test_00981_history_contains_only_verified_complete_years():
    cfg = yaml.safe_load((ROOT / "configs" / "00981.yaml").read_text(encoding="utf-8"))
    years = cfg["model"]["hist_years"]

    assert years == [2024, 2025]
    assert len(cfg["segments"][0]["hist_share"]) == len(years)
    assert len(cfg["segments"][0]["hist_gm"]) == len(years)

    for statement in ("is", "bs", "cf"):
        for name, values in cfg["hist"][statement].items():
            assert len(values) == len(years), f"{statement}.{name} length"

    # Negative 'other' residuals are evidence of a missing source subtotal, not
    # valid balancing items.  The previously included 2023 row had -8.4bn and
    # -14.4bn residuals and therefore must never return under a new label.
    for name in ("oca", "onca", "ocl", "oncl", "oeq"):
        assert min(cfg["hist"]["bs"][name]) >= 0, name
