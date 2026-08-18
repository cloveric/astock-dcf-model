"""Currency and provenance labels shared by workbook outputs."""

from __future__ import annotations

from collections.abc import Mapping
from typing import Any


def derive_currency_labels(company: dict[str, Any]) -> dict[str, str]:
    explicit = str(company.get("currency") or company.get("currency_code") or "").upper()
    description = " ".join(str(company.get(key) or "") for key in ("unit", "currency_note"))
    if explicit in {"USD", "US$"} or "美元" in description:
        return {"currency_code": "USD", "per_share_unit": "美元/股", "denomination": "美元"}
    if explicit in {"HKD", "HK$"} or "港元" in description or "港币" in description:
        return {"currency_code": "HKD", "per_share_unit": "港元/股", "denomination": "港元"}
    return {"currency_code": "CNY", "per_share_unit": "元/股", "denomination": "元"}


def consensus_is_verified(basis: str) -> bool:
    """Return whether a configuration basis names a reviewed external source.

    Ambiguous provenance such as ``待核`` or ``占位`` must never enter the
    verified valuation envelope merely because it also contains the phrase
    ``一致预期``.
    """
    normalized = str(basis or "").strip().lower()
    if not normalized:
        return False
    if any(marker in normalized for marker in (
        "自动推导", "自动外推", "无一致预期", "建议手工填入",
        "待核", "占位", "未验证", "平推", "沿用原配置",
    )):
        return False
    return any(marker in normalized for marker in (
        "gildata", "聚源", "wind", "万得", "一致预期", "券商", "外部", "手工", "manual",
    ))


def consensus_year_is_verified(
    basis: str,
    year: int,
    external_values: Mapping[object, object] | None = None,
) -> bool:
    """Verify one series/year, not an entire partially supplied research file.

    ``None`` means no external file was active, so the checked-in basis decides.
    A mapping (including an empty one) means an external file was active and the
    requested year must be explicitly present in that exact series.
    """
    if external_values is None:
        return consensus_is_verified(basis)
    for raw_year, value in external_values.items():
        try:
            matches = int(raw_year) == int(year)
        except (TypeError, ValueError):
            matches = False
        if matches and value is not None:
            return True
    return False


def consensus_label(verified: bool) -> str:
    return "一致预期" if verified else "模型自动外推"


def forward_earnings_verified(comp: dict[str, Any], inherited_source: str = "") -> bool:
    """Whether a comparable has a formal FY1/NTM earnings basis for pricing."""
    if comp.get("ref_only"):
        return False
    if "earnings_basis" in comp or "earnings_verified" in comp:
        basis = str(comp.get("earnings_basis") or "").strip().upper()
        return comp.get("earnings_verified") is True and basis in {
            "FY1", "F0", "NTM", "FORWARD", "FORWARD_FY1",
        }
    source = str(comp.get("src") or "").strip()
    if source in {"同上", "同上。"}:
        source = inherited_source
    normalized = source.lower()
    if any(marker in normalized for marker in ("ttm", "占位", "待核", "估算")):
        return False
    return any(marker in normalized for marker in (
        "一致预期", "gildata", "聚源", "wind", "万得", "券商",
    ))
