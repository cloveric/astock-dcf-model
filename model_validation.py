"""Fail-fast configuration and historical balance-sheet controls."""

from __future__ import annotations

import math
from pathlib import Path
from typing import Any, Optional, Union


BS_ASSET_KEYS = (
    "cash", "tfa", "ar", "pre", "orec", "inv", "oca", "ppe", "rou",
    "ia", "gw", "lpe", "dta", "oei", "onca",
)
BS_LIABILITY_KEYS = (
    "stl", "ap", "contract", "staff", "taxp", "opay", "cur1y", "ocl",
    "ltl", "lease", "oncl",
)
BS_EQUITY_KEYS = ("sc", "cr", "oci", "sr", "re", "oeq", "mi")
BS_RESIDUAL_KEYS = ("oca", "onca", "ocl", "oncl", "oeq")
BS_ROUNDING_ABS = 0.5
BS_ROUNDING_REL = 1e-5
WACC_G_SAFETY_MARGIN = 0.005


def load_yaml_strict(path: Union[str, Path]) -> dict[str, Any]:
    """Load YAML after rejecting duplicate mapping keys at any depth."""
    import yaml
    from yaml.nodes import MappingNode, ScalarNode, SequenceNode

    text = Path(path).read_text(encoding="utf-8")
    root = yaml.compose(text, Loader=yaml.SafeLoader)

    def visit(node: Any, parts: list[str]) -> None:
        if isinstance(node, MappingNode):
            seen: set[str] = set()
            for key_node, value_node in node.value:
                key = key_node.value if isinstance(key_node, ScalarNode) else str(key_node.value)
                full_path = ".".join(parts + [key])
                if key in seen:
                    raise ValueError(f"配置错误: duplicate YAML key at {full_path}")
                seen.add(key)
                visit(value_node, parts + [key])
        elif isinstance(node, SequenceNode):
            for index, value_node in enumerate(node.value):
                visit(value_node, parts + [f"[{index}]"])

    if root is not None:
        visit(root, [])
    loaded = yaml.safe_load(text)
    if not isinstance(loaded, dict):
        raise ValueError("配置错误: YAML root must be a mapping")
    return loaded


def _get(raw: dict[str, Any], path: str, default: Any = None) -> Any:
    node: Any = raw
    for part in path.split("."):
        if not isinstance(node, dict) or part not in node:
            return default
        node = node[part]
    return node


def _entry(entry: Any) -> tuple[Any, str]:
    if isinstance(entry, dict) and ("value" in entry or "values" in entry):
        return entry.get("values", entry.get("value")), str(entry.get("basis") or "")
    return entry, ""


def _numbers(value: Any, path: str) -> list[float]:
    values = list(value) if isinstance(value, (list, tuple)) else [value]
    if not values:
        raise ValueError(f"配置错误: {path} must not be empty")
    for item in values:
        if isinstance(item, dict):
            item = item.get("value")
        if isinstance(item, bool) or not isinstance(item, (int, float)):
            raise ValueError(f"配置错误: {path} must contain finite numbers")
        if not math.isfinite(float(item)):
            raise ValueError(f"配置错误: finite number required at {path}")
    return [float(item.get("value")) if isinstance(item, dict) else float(item) for item in values]


def _validate_finite_tree(node: Any, path: str = "config") -> None:
    if isinstance(node, dict):
        for key, value in node.items():
            _validate_finite_tree(value, f"{path}.{key}")
    elif isinstance(node, (list, tuple)):
        for index, value in enumerate(node):
            _validate_finite_tree(value, f"{path}[{index}]")
    elif isinstance(node, float) and not math.isfinite(node):
        raise ValueError(f"配置错误: finite number required at {path}")


def _validate_years(raw: dict[str, Any]) -> tuple[list[int], list[int]]:
    hist = _get(raw, "model.hist_years")
    fcst = _get(raw, "model.fcst_years")
    if not isinstance(hist, list) or not isinstance(fcst, list):
        raise ValueError("配置错误: model.hist_years/model.fcst_years must be lists")
    if len(hist) < 2 or not 3 <= len(fcst) <= 5:
        raise ValueError("配置错误: model.hist_years需至少2年且model.fcst_years需3-5年")
    if any(isinstance(year, bool) or not isinstance(year, int) for year in hist + fcst):
        raise ValueError("配置错误: model.hist_years/model.fcst_years must contain integer years")
    if hist != list(range(hist[0], hist[0] + len(hist))):
        raise ValueError("配置错误: model.hist_years must be sorted and contiguous")
    if fcst != list(range(fcst[0], fcst[0] + len(fcst))):
        raise ValueError("配置错误: model.fcst_years must be sorted and contiguous")
    if fcst[0] != hist[-1] + 1:
        raise ValueError("配置错误: model.fcst_years must start immediately after model.hist_years")
    return hist, fcst


def _validate_vector(raw: dict[str, Any], path: str, length: int, *, required: bool = True) -> list[float]:
    entry = _get(raw, path)
    if entry is None:
        if required:
            raise ValueError(f"配置错误: missing {path}")
        return []
    value, _ = _entry(entry)
    if isinstance(value, (list, tuple)) and len(value) != length:
        raise ValueError(
            f"配置错误: {path} vector length must be exactly {length}; "
            "use a scalar for explicit broadcasting"
        )
    return _numbers(value, path)


def _validate_basis(raw: dict[str, Any], path: str) -> None:
    entry = _get(raw, path)
    _, basis = _entry(entry)
    if not basis.strip():
        raise ValueError(f"配置错误: {path}.basis must be non-empty")


def _validate_range(raw: dict[str, Any], path: str, low: float, high: float) -> None:
    value, _ = _entry(_get(raw, path))
    for number in _numbers(value, path):
        if not low <= number <= high:
            raise ValueError(f"配置错误: {path} must be within [{low:g}, {high:g}]")


def historical_balance_totals(bs: dict[str, Any], years: Optional[list[int]] = None) -> list[dict[str, float]]:
    if years is None:
        length = len(bs[BS_ASSET_KEYS[0]])
        years = list(range(length))
    totals = []
    for index, year in enumerate(years):
        assets = sum(float(bs[key][index]) for key in BS_ASSET_KEYS)
        liabilities_equity = (
            sum(float(bs[key][index]) for key in BS_LIABILITY_KEYS)
            + sum(float(bs[key][index]) for key in BS_EQUITY_KEYS)
            - float(bs["ts"][index])
        )
        totals.append({
            "year": year,
            "assets": assets,
            "liabilities_equity": liabilities_equity,
            "diff": assets - liabilities_equity,
        })
    return totals


def prepare_historical_bs(raw: dict[str, Any]) -> tuple[dict[str, Any], list[dict[str, float]]]:
    """Return a normalized BS copy, allowing only immaterial rounding plugs."""
    years = list(_get(raw, "model.hist_years") or [])
    source = _get(raw, "hist.bs")
    if not isinstance(source, dict):
        raise ValueError("配置错误: missing hist.bs")
    bs = {key: list(value) if isinstance(value, (list, tuple)) else value for key, value in source.items()}
    bs.setdefault("mi", [0.0] * len(years))
    required = set(BS_ASSET_KEYS + BS_LIABILITY_KEYS + BS_EQUITY_KEYS + ("ts",))
    for key in sorted(required):
        if key not in bs:
            raise ValueError(f"配置错误: missing hist.bs.{key}")
        values = _numbers(bs[key], f"hist.bs.{key}")
        if len(values) != len(years):
            raise ValueError(f"配置错误: hist.bs.{key} vector length must be exactly {len(years)}")

    for index, year in enumerate(years):
        assets = sum(float(bs[key][index]) for key in BS_ASSET_KEYS)
        threshold = max(BS_ROUNDING_ABS, abs(assets) * BS_ROUNDING_REL)
        for key in BS_RESIDUAL_KEYS:
            value = float(bs[key][index])
            if value < -threshold:
                raise ValueError(
                    f"配置错误: hist.bs.{key}[{year}]={value:g} is a materially negative residual bucket"
                )
            if value < 0:
                bs[key][index] = 0.0

    plugs: list[dict[str, float]] = []
    for item in historical_balance_totals(bs, years):
        diff = float(item["diff"])
        assets = float(item["assets"])
        threshold = max(BS_ROUNDING_ABS, abs(assets) * BS_ROUNDING_REL)
        ratio = abs(diff) / max(abs(assets), 1.0)
        if abs(diff) > threshold:
            raise ValueError(
                "配置错误: material balance-sheet difference at "
                f"{item['year']}A: assets-liabilities-equity={diff:+.4f}, "
                f"threshold={threshold:.4f}"
            )
        if abs(diff) > 1e-9:
            index = years.index(item["year"])
            bs["oeq"][index] = round(float(bs["oeq"][index]) + diff, 10)
            plugs.append({
                "year": int(item["year"]),
                "diff": diff,
                "assets": assets,
                "ratio": ratio,
                "threshold": threshold,
            })
    remaining = historical_balance_totals(bs, years)
    if any(abs(item["diff"]) > 1e-7 for item in remaining):
        raise ValueError("配置错误: historical balance-sheet rounding plug did not clear")
    return bs, plugs


def validate_config(raw: dict[str, Any]) -> None:
    if not isinstance(raw, dict):
        raise ValueError("配置错误: config root must be a mapping")
    _validate_finite_tree(raw)
    checks = raw.get("checks")
    if isinstance(checks, dict) and "fcfe_divergence_waiver" in checks:
        waiver = checks["fcfe_divergence_waiver"]
        if not isinstance(waiver, str) or not waiver.strip():
            raise ValueError("配置错误: checks.fcfe_divergence_waiver must be a non-empty reviewed reason")
    hist, fcst = _validate_years(raw)
    nh, nf = len(hist), len(fcst)

    forecast_vectors = (
        "opex.tax_add_rate", "opex.sale_rate", "opex.adm_rate", "opex.rd_rate",
        "opex.oth_op", "opex.nonop", "opex.tax_rate", "working_capital.dso",
        "working_capital.dio", "working_capital.dpo", "working_capital.pre_rate",
        "working_capital.staff_rate", "working_capital.taxp_rate", "capex.capex_rate",
        "dividend.payout", "financing.rep_st", "financing.rep_cur",
        "financing.rep_lt", "financing.rep_lease",
    )
    for path in forecast_vectors:
        _validate_vector(raw, path, nf)
    for path in ("consensus.rev", "consensus.np"):
        _validate_vector(raw, path, min(3, nf))

    hist_sections = ("hist.is", "hist.cf")
    for section in hist_sections:
        node = _get(raw, section)
        if not isinstance(node, dict):
            raise ValueError(f"配置错误: missing {section}")
        for key, value in node.items():
            if isinstance(value, (list, tuple)):
                if len(value) != nh:
                    raise ValueError(f"配置错误: {section}.{key} vector length must be exactly {nh}")
                _numbers(value, f"{section}.{key}")

    segments = _get(raw, "segments")
    if not isinstance(segments, list) or not segments:
        raise ValueError("配置错误: segments must be non-empty")
    for index, segment in enumerate(segments):
        for field in ("hist_share", "hist_gm"):
            values = segment.get(field)
            if not isinstance(values, list) or len(values) != nh:
                raise ValueError(f"配置错误: segments[{index}].{field} vector length must be exactly {nh}")
            _numbers(values, f"segments[{index}].{field}")
        for field in ("vol", "gm"):
            _validate_vector({"segment": segment}, f"segment.{field}", nf)
        if segment.get("driver") == "vol_asp" or segment.get("asp") is not None:
            _validate_vector({"segment": segment}, "segment.asp", nf)
        gm, _ = _entry(segment.get("gm"))
        for number in _numbers(gm, f"segments[{index}].gm"):
            if number < 0 or number > 1:
                raise ValueError(f"配置错误: segments[{index}].gm must be within [0, 1]")

    price_path = "market.price_hkd" if _get(raw, "market.price_hkd") is not None else "market.price"
    price, _ = _entry(_get(raw, price_path))
    shares, _ = _entry(_get(raw, "market.shares"))
    if _numbers(price, price_path)[0] <= 0:
        raise ValueError(f"配置错误: {price_path} must be positive")
    if _numbers(shares, "market.shares")[0] <= 0:
        raise ValueError("配置错误: market.shares must be positive")

    _validate_range(raw, "opex.tax_rate", 0, 0.6)
    for path in (
        "financing.rate_st", "financing.rate_cur", "financing.rate_lt",
        "financing.rate_lease", "financing.cash_yield", "wacc.rf", "wacc.erp",
        "wacc.srp", "wacc.kd",
    ):
        _validate_range(raw, path, 0, 0.5)
    _validate_range(raw, "wacc.tg", -0.05, 0.1)

    scenarios = _get(raw, "scenarios") or {}
    for field in ("rev_adj", "npm_adj"):
        ordered = []
        for key in ("bear", "base", "bull"):
            scenario = scenarios.get(key) or {}
            ordered.extend(_numbers(scenario.get(field), f"scenarios.{key}.{field}"))
            if not str(scenario.get("logic") or "").strip():
                raise ValueError(f"配置错误: scenarios.{key}.logic must be non-empty")
        if ordered != sorted(ordered):
            raise ValueError(f"配置错误: scenarios.*.{field} must satisfy bear <= base <= bull")

    override, _ = _entry(_get(raw, "wacc.override"))
    terminal_g, _ = _entry(_get(raw, "wacc.tg"))
    terminal_g = _numbers(terminal_g, "wacc.tg")[0]
    if override is not None:
        override = _numbers(override, "wacc.override")[0]
        if override <= terminal_g + WACC_G_SAFETY_MARGIN:
            raise ValueError(
                f"配置错误: wacc.override must exceed wacc.tg by more than {WACC_G_SAFETY_MARGIN:.1%}"
            )

    basis_paths = (
        price_path, "market.shares", "consensus.rev", "consensus.np",
        *forecast_vectors, "capex.trans_rate", "capex.dep_rate", "capex.dep_new_rate",
        "capex.disp_rate", "capex.amort_rate", "dividend.surplus_rate",
        "financing.min_cash_pct", "financing.rate_st", "financing.rate_cur",
        "financing.rate_lt", "financing.rate_lease", "financing.cash_yield",
        "financing.fin_oth", "wacc.rf", "wacc.erp", "wacc.srp", "wacc.kd",
        "wacc.tg", "relative_val.target_pe_lo",
    )
    for path in basis_paths:
        _validate_basis(raw, path)
    for index, segment in enumerate(segments):
        for field in ("vol", "gm") + (("asp",) if segment.get("asp") is not None else ()):
            _, basis = _entry(segment.get(field))
            if not basis.strip():
                raise ValueError(f"配置错误: segments[{index}].{field}.basis must be non-empty")
    comps = _get(raw, "relative_val.comps") or []
    if not any(comp.get("beta_l") is not None for comp in comps):
        if _get(raw, "wacc.beta_unlevered_input") is None:
            raise ValueError("配置错误: wacc.beta_unlevered_input required without beta comparables")
        if not str(_get(raw, "wacc.beta_unlevered_basis") or "").strip():
            raise ValueError("配置错误: wacc.beta_unlevered_basis must be non-empty")

    prepare_historical_bs(raw)
