"""Audience Data loader — verbatim Reach/Index from the GTM xlsx."""

from __future__ import annotations

import io

import pytest
from openpyxl import Workbook

from ingestion.audience_data import (
    AudienceData,
    AudienceRow,
    match_segment_names,
    rank_segments_by_index,
)


def _xlsx_bytes(rows: list[tuple], *, headers: list[str] | None = None) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Audience Data"
    ws.append(headers or ["Audience Segment", "Reach", "Index"])
    for row in rows:
        ws.append(list(row))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


_SAMPLE = [
    ("Chief Executive Officer", "1.1M", 154),
    ("C-suite", "3.6M", 172),
    ("C-Suite in energy sector", "164K", 129),
    ("Chief Financial Officer", "422K", 146),
]


@pytest.fixture
def audience() -> AudienceData:
    return AudienceData.from_xlsx_bytes(_xlsx_bytes(_SAMPLE))


def test_lookup_verbatim_reach_and_index(audience: AudienceData):
    row = audience.lookup("Chief Executive Officer")
    assert row.segment == "Chief Executive Officer"
    assert row.reach == "1.1M"
    assert row.index == "154"


def test_lookup_case_insensitive(audience: AudienceData):
    row = audience.lookup("chief executive officer")
    assert row.segment == "Chief Executive Officer"


def test_lookup_missing_fails_loud(audience: AudienceData):
    with pytest.raises(ValueError, match="No Audience Data row for segment 'Interns'"):
        audience.lookup("Interns")


def test_match_ignores_unrecognized_prose(audience: AudienceData):
    rows = audience.match_targeting(
        "US enterprise tech decision-makers, Chief Executive Officer, C-suite"
    )
    names = [r.segment for r in rows]
    assert names == ["Chief Executive Officer", "C-suite"]
    assert all(r.reach and r.index for r in rows)


def test_match_prefers_longer_segment_when_both_would_hit(audience: AudienceData):
    rows = audience.match_targeting("C-Suite in energy sector and other buyers")
    names = [r.segment for r in rows]
    assert names == ["C-Suite in energy sector"]
    assert "C-suite" not in names


def test_match_short_name_alone(audience: AudienceData):
    rows = audience.match_targeting("Need C-suite coverage in the US")
    assert [r.segment for r in rows] == ["C-suite"]


def test_missing_sheet_raises():
    wb = Workbook()
    wb.active.title = "Product Tags"
    buf = io.BytesIO()
    wb.save(buf)
    with pytest.raises(ValueError, match="missing 'Audience Data'"):
        AudienceData.from_xlsx_bytes(buf.getvalue())


def test_missing_columns_raises():
    data = _xlsx_bytes(
        [("Chief Executive Officer", "1.1M")],
        headers=["Audience Segment", "Reach"],
    )
    with pytest.raises(ValueError, match="missing columns"):
        AudienceData.from_xlsx_bytes(data)


def test_empty_reach_row_skipped_lookup_fails():
    data = _xlsx_bytes(
        [
            ("Chief Executive Officer", "1.1M", 154),
            ("Ghost Segment", "", 100),
        ]
    )
    audience = AudienceData.from_xlsx_bytes(data)
    audience.lookup("Chief Executive Officer")
    with pytest.raises(ValueError, match="No Audience Data row"):
        audience.lookup("Ghost Segment")


def test_conflicting_duplicate_segment_raises():
    data = _xlsx_bytes(
        [
            ("C-suite", "3.6M", 172),
            ("C-suite", "9.9M", 1),
        ]
    )
    with pytest.raises(ValueError, match="conflicting Reach/Index"):
        AudienceData.from_xlsx_bytes(data)


def test_rank_segments_by_index_does_not_invent_metrics():
    rows = [
        AudienceRow("A", "1M", "100"),
        AudienceRow("B", "2M", "200"),
        AudienceRow("C", "3M", "150"),
    ]
    ranked = rank_segments_by_index(rows)
    assert [r.segment for r in ranked] == ["B", "C", "A"]
    assert ranked[0].reach == "2M"


def test_match_segment_names_empty_targeting():
    assert match_segment_names("", ["C-suite"]) == []


def test_match_curly_apostrophe_in_targeting():
    data = AudienceData(
        [
            AudienceRow("Influences Other People's Investments", "4.0M", "161"),
            AudienceRow("C-suite", "3.6M", "172"),
        ]
    )
    rows = data.match_targeting("Influences Other People’s Investments in the US")
    assert [r.segment for r in rows] == ["Influences Other People's Investments"]
    assert rows[0].reach == "4.0M"
