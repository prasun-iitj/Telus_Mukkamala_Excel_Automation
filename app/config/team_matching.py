"""Team-name matching between Excel data and PPT template rows."""

from __future__ import annotations

import re
from typing import Optional

from app.models import MonthData, TeamMetrics

# Excel team name (normalized) -> PPT template aliases (normalized substrings)
EXCEL_TO_PPT_ALIASES: dict[str, list[str]] = {
    # Apps
    "seamless entertainment": ["seamless entertainment"],
    "auth id": ["auth id"],
    "dsb": ["dsb"],
    "dataservices -core": ["domain services", "dataservices -core", "dataservices-core"],
    "enterprise data services": ["enterprise data services"],
    # Platforms
    "effie": ["effie"],
    "charter abuse tracking system (cats)": ["cats"],
    "enterprise data platforms (edp)": ["edp"],
    "solo": ["solo"],
    "ping esd": ["ping"],
    "gateways aws": ["gateways - aws", "gateways aws"],
    "gateways - legacy": ["gateways - dc", "gateways - legacy"],
    "gateways - blue cloud": ["gateways - dc", "gateways - blue cloud"],
}

GATEWAY_EXCEL_TEAMS = {"gateways - legacy", "gateways - blue cloud"}


def normalize_team(name: str) -> str:
    """Normalize a team label for fuzzy comparison."""
    if not name:
        return ""
    first_line = str(name).split("\n")[0].strip().lower()
    first_line = first_line.replace("\xa0", " ")
    first_line = re.sub(r"\s+", " ", first_line)
    return first_line.strip()


def _alias_match(excel_norm: str, ppt_norm: str) -> bool:
    """Return True if excel team maps to the ppt team via alias table."""
    if excel_norm == ppt_norm:
        return True
    if excel_norm in ppt_norm or ppt_norm in excel_norm:
        return True

    aliases = EXCEL_TO_PPT_ALIASES.get(excel_norm, [])
    for alias in aliases:
        alias_norm = normalize_team(alias)
        if alias_norm == ppt_norm or alias_norm in ppt_norm or ppt_norm in alias_norm:
            return True

    # Reverse lookup: ppt label may be a short form (e.g. "cats")
    for _excel_key, ppt_aliases in EXCEL_TO_PPT_ALIASES.items():
        if excel_norm == _excel_key or _excel_key in excel_norm:
            for alias in ppt_aliases:
                alias_norm = normalize_team(alias)
                if alias_norm == ppt_norm or alias_norm in ppt_norm or ppt_norm in alias_norm:
                    return True
    return False


def teams_match(excel_team: str, ppt_team: str) -> bool:
    """Return True when an Excel team corresponds to a PPT template team row."""
    return _alias_match(normalize_team(excel_team), normalize_team(ppt_team))


def _merge_team_values(teams: list[TeamMetrics]) -> dict:
    """Merge metric dicts from multiple teams (sum numerics, keep first text)."""
    merged: dict = {}
    for team in teams:
        for key, val in team.values.items():
            if isinstance(val, (int, float)):
                existing = merged.get(key)
                if isinstance(existing, (int, float)):
                    merged[key] = existing + val
                else:
                    merged[key] = val
            elif key not in merged or merged[key] is None:
                merged[key] = val
    return merged


def build_team_lookup(month_data: MonthData) -> dict[str, TeamMetrics]:
    """Build a normalized lookup for month teams, merging gateway rows into one."""
    lookup: dict[str, TeamMetrics] = {}
    gateway_teams: list[TeamMetrics] = []

    for team in month_data.teams:
        key = normalize_team(team.team_name)
        if key in GATEWAY_EXCEL_TEAMS:
            gateway_teams.append(team)
            continue
        lookup[key] = team

    if gateway_teams:
        lookup["gateways - dc"] = TeamMetrics(
            team_name="Gateways - DC",
            values=_merge_team_values(gateway_teams),
        )

    return lookup


def find_team_by_platform_name(
    platform_team_name: str,
    month_data: MonthData,
) -> Optional[TeamMetrics]:
    """Find team metrics by exact or alias-matched Master-Platforms team name."""
    target = normalize_team(platform_team_name)
    lookup = build_team_lookup(month_data)

    if target in lookup:
        return lookup[target]

    for team in month_data.teams:
        if teams_match(team.team_name, platform_team_name):
            return team

    if "gateways" in target and "dc" in target:
        return lookup.get("gateways - dc")

    return None


def find_team_for_ppt_row(ppt_team: str, month_data: MonthData) -> Optional[TeamMetrics]:
    """Find the Excel TeamMetrics that corresponds to a PPT table row label."""
    lookup = build_team_lookup(month_data)
    ppt_norm = normalize_team(ppt_team)

    if ppt_norm in lookup:
        return lookup[ppt_norm]

    for team in month_data.teams:
        if teams_match(team.team_name, ppt_team):
            return team

    if "gateways" in ppt_norm and ("dc" in ppt_norm or "legacy" in ppt_norm):
        return lookup.get("gateways - dc")

    return None
