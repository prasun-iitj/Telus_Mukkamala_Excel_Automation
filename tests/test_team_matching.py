"""Tests for team name matching between Excel and PPT."""

from app.config.team_matching import (
    build_team_lookup,
    find_team_for_ppt_row,
    normalize_team,
    teams_match,
)
from app.models import MonthData, TeamMetrics


def test_normalize_team_strips_newlines():
    assert normalize_team("Effie\n(Account, Identity)") == "effie"


def test_apps_domain_services_mapping():
    assert teams_match("DataServices -Core", "Domain Services")


def test_platforms_cats_mapping():
    assert teams_match("Charter Abuse Tracking System (CATS)", "CATS")


def test_gateways_merge():
    month = MonthData(
        month_name="March",
        teams=[
            TeamMetrics("Gateways - Legacy", {"# of deployments": 2}),
            TeamMetrics("Gateways - Blue Cloud", {"# of deployments": 3}),
        ],
    )
    lookup = build_team_lookup(month)
    merged = lookup["gateways - dc"]
    assert merged.values["# of deployments"] == 5


def test_find_team_for_ppt_gateways_dc():
    month = MonthData(
        month_name="March",
        teams=[
            TeamMetrics("Gateways - Legacy", {"# of deployments": 1}),
            TeamMetrics("Gateways - Blue Cloud", {"# of deployments": 4}),
        ],
    )
    team = find_team_for_ppt_row("Gateways - DC", month)
    assert team is not None
    assert team.values["# of deployments"] == 5
