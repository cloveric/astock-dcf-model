"""Currency and provenance labels shared by workbook outputs."""

from __future__ import annotations

import datetime
import ipaddress
import math
import re
from collections.abc import Mapping
from typing import Any
from urllib.parse import urlsplit


FORWARD_EARNINGS_BASES = frozenset({"FY1", "F0", "NTM", "FORWARD", "FORWARD_FY1"})
FORWARD_PE_RECONCILIATION_TOLERANCE = 0.15
FORWARD_SOURCE_MAX_AGE_DAYS = 90
_STRICT_DATE_RE = re.compile(r"\d{4}-\d{2}-\d{2}\Z")
_DNS_LABEL_RE = re.compile(r"[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\Z")


def _public_https_url(value: Any) -> bool:
    """Accept a credential-free public HTTPS URL without local-network literals."""
    source_url = str(value or "")
    if (
        not source_url
        or source_url != source_url.strip()
        or any(ord(char) < 32 for char in source_url)
    ):
        return False
    try:
        parsed = urlsplit(source_url)
        hostname = parsed.hostname
    except ValueError:
        return False
    if parsed.scheme.lower() != "https" or parsed.username or parsed.password or not hostname:
        return False
    try:
        parsed.port
    except ValueError:
        return False
    host = hostname.rstrip(".").lower()
    if host == "localhost" or host.endswith(".localhost"):
        return False
    try:
        address = ipaddress.ip_address(host)
    except ValueError:
        try:
            ascii_host = host.encode("idna").decode("ascii")
        except UnicodeError:
            return False
        labels = ascii_host.split(".")
        return (
            len(ascii_host) <= 253
            and len(labels) >= 2
            and all(_DNS_LABEL_RE.fullmatch(label) for label in labels)
        )
    return address.is_global


def _strict_date(value: Any) -> datetime.date | None:
    rendered = str(value or "")
    if not _STRICT_DATE_RE.fullmatch(rendered):
        return None
    try:
        return datetime.date.fromisoformat(rendered)
    except ValueError:
        return None


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


def forward_earnings_review(
    comp: dict[str, Any],
    inherited_source: str = "",
    valuation_date: object = None,
) -> tuple[bool, str]:
    """Return whether a comparable is safe for formal pricing and why.

    Explicitly reviewed comparables must carry structured provenance and their
    market-cap/earnings PE must reconcile to the provider's published forward
    PE. This prevents a self-attested boolean from bypassing data controls.
    """
    if comp.get("ref_only"):
        return False, "ref_only comparable"
    if "earnings_basis" in comp or "earnings_verified" in comp:
        if comp.get("earnings_verified") is not True:
            return False, "earnings_verified is not true"
        basis = str(comp.get("earnings_basis") or "").strip().upper()
        if basis not in FORWARD_EARNINGS_BASES:
            return False, "earnings_basis must be FY1/F0/NTM/forward"
        source = str(comp.get("src") or "").strip()
        if not source:
            return False, "src must be non-empty for verified earnings"
        if not _public_https_url(comp.get("source_url")):
            return False, "source_url must be a credential-free public HTTPS URL"
        source_as_of = _strict_date(comp.get("source_as_of"))
        if source_as_of is None:
            return False, "source_as_of must be YYYY-MM-DD"
        if valuation_date is not None:
            valuation = _strict_date(valuation_date)
            if valuation is None:
                return False, "valuation_date must be YYYY-MM-DD"
            if source_as_of > valuation:
                return False, "source_as_of must not be after model.valuation_date"
            age_days = (valuation - source_as_of).days
            if age_days > FORWARD_SOURCE_MAX_AGE_DAYS:
                return False, (
                    f"source_as_of is {age_days} days old; maximum is "
                    f"{FORWARD_SOURCE_MAX_AGE_DAYS} days"
                )
        values: dict[str, float] = {}
        for field in ("mcap", "np_f0", "source_forward_pe"):
            value = comp.get(field)
            if isinstance(value, bool) or not isinstance(value, (int, float)):
                return False, f"{field} must be a positive finite number"
            number = float(value)
            if not math.isfinite(number) or number <= 0:
                return False, f"{field} must be a positive finite number"
            values[field] = number
        calculated_pe = values["mcap"] / values["np_f0"]
        mismatch = abs(calculated_pe / values["source_forward_pe"] - 1.0)
        if mismatch > FORWARD_PE_RECONCILIATION_TOLERANCE:
            return False, (
                "source_forward_pe mismatch: "
                f"calculated={calculated_pe:.2f}x, source={values['source_forward_pe']:.2f}x, "
                f"gap={mismatch:.1%} exceeds {FORWARD_PE_RECONCILIATION_TOLERANCE:.0%}"
            )
        return True, f"forward PE reconciled within {mismatch:.1%}"
    return False, "structured forward provenance is required for formal pricing"


def forward_earnings_verified(
    comp: dict[str, Any],
    inherited_source: str = "",
    valuation_date: object = None,
) -> bool:
    """Whether a comparable has a formal, reconciled FY1/NTM pricing basis."""
    return forward_earnings_review(comp, inherited_source, valuation_date)[0]
