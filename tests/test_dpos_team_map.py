"""Tests for DPOS codename mapping."""

from app.config.dpos_team_map_parser import (
    extract_dpos_codename,
    load_dpos_team_map,
)
from app.config.team_matching import find_team_by_platform_name
from app.models import MonthData, TeamMetrics


def test_extract_dpos_codename():
    assert extract_dpos_codename("Byron\n(Account, Identity)") == "Byron"


def test_load_dpos_team_map():
    cfg = load_dpos_team_map()
    assert "Goldberg" in cfg.teams
    assert cfg.teams["Goldberg"].platform_team == "SOLO"


def test_dpos_mapping_resolves_platform_team():
    month = MonthData(
        month_name="March",
        teams=[
            TeamMetrics("SOLO", {"Total Defects Found in PROD (P0-P4)": 3}),
        ],
    )
    cfg = load_dpos_team_map()
    platform_name = cfg.teams["Goldberg"].platform_team
    team = find_team_by_platform_name(platform_name, month)
    assert team is not None
    assert team.values["Total Defects Found in PROD (P0-P4)"] == 3
