# -*- coding: utf-8 -*-
"""Data-layer controls added after the 2026-08 comprehensive audit.

Each test names a production failure it must catch: silent interim fallback,
economically impossible interim margins, materially negative HK residuals, and
missing request-lineage metadata.
"""
import hashlib
import json
import os
import sys
from pathlib import Path

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import fetch_data as fd


def test_fetch_interim_distinguishes_transport_error(monkeypatch, capsys):
    """A broken endpoint must not be indistinguishable from no filing."""
    status = {}

    def broken(_url):
        raise RuntimeError("upstream unavailable")

    monkeypatch.setattr(fd, "curl_get", broken)
    assert fd.fetch_interim("688825", 2025, status_out=status) is None
    assert status["status"] == "error"
    assert status["source"] == "eastmoney_interim"
    assert "upstream unavailable" in status["error"]
    assert "中期数据抓取失败" in capsys.readouterr().err


def test_fetch_interim_distinguishes_no_filing(monkeypatch):
    """A valid empty response must be recorded as not_available, not error."""
    status = {}
    monkeypatch.setattr(fd, "curl_get", lambda _url: json.dumps({"data": []}).encode())
    assert fd.fetch_interim("688825", 2025, status_out=status) is None
    assert status == {"source": "eastmoney_interim", "status": "not_available"}


def test_interim_gross_margin_above_100_percent_is_rejected():
    """A negative cost mapping must not create a >100% forecast gross margin."""
    interim = {"label": "2026Q1", "rev": 100.0, "cost": -20.0}
    with pytest.raises(ValueError, match="毛利率"):
        fd._interim_gm_path(interim, [0.2, 0.3, 0.4])


def test_material_negative_hk_residual_is_rejected():
    """Missing HK totals must not surface as a large negative residual bucket."""
    with pytest.raises(RuntimeError, match="oncl"):
        fd._guard_hk_residual("oncl", -8360.9, 47787.4, 2023)
    assert fd._guard_hk_residual("oncl", -0.1, 47787.4, 2023) == -0.1


def test_curl_get_records_request_lineage(monkeypatch):
    """Every public-data response needs timestamp/hash/size lineage."""
    payload = b'{"ok":true}'

    class Result:
        stdout = payload

    monkeypatch.setattr(fd.subprocess, "run", lambda *a, **k: Result())
    fd.reset_fetch_manifest()
    url = "https://example.test/data?q=1"
    assert fd.curl_get(url) == payload
    manifest = fd.get_fetch_manifest()
    assert len(manifest) == 1
    assert manifest[0]["url"] == url
    assert manifest[0]["sha256"] == hashlib.sha256(payload).hexdigest()
    assert manifest[0]["bytes"] == len(payload)
    assert manifest[0]["fetched_at"].endswith("Z")


def test_curl_get_can_persist_immutable_raw_snapshot(monkeypatch, tmp_path):
    payload = b'{"data":[1,2,3]}'

    class Result:
        stdout = payload

    monkeypatch.setattr(fd.subprocess, "run", lambda *a, **k: Result())
    fd.reset_fetch_manifest(raw_dir=tmp_path)
    fd.curl_get("https://example.test/public-data")

    item = fd.get_fetch_manifest()[0]
    snapshot = Path(item["snapshot"])
    assert snapshot.parent == tmp_path
    assert snapshot.read_bytes() == payload
    assert item["sha256"] in snapshot.name
