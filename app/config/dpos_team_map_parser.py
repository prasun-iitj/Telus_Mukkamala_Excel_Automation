"""Parser for dpos_team_map.yaml — DPOS codename to platform team mapping."""

from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Optional

import yaml

from app.exceptions import SlideMapError


@dataclass
class DposTeamMapping:
    """Maps a DPOS PPT codename to a Master-Platforms team."""

    codename: str
    platform_team: str
    notes: Optional[str] = None


@dataclass
class DposTeamMapConfig:
    """Full DPOS team mapping configuration."""

    teams: dict[str, DposTeamMapping]


def load_dpos_team_map(path: Optional[str] = None) -> DposTeamMapConfig:
    """Load and validate dpos_team_map.yaml."""
    if path is None:
        path = os.path.join(os.path.dirname(__file__), "dpos_team_map.yaml")

    if not os.path.isfile(path):
        raise SlideMapError(f"DPOS team map not found: {path}")

    with open(path, encoding="utf-8") as fh:
        raw = yaml.safe_load(fh)

    if not raw or "dpos_teams" not in raw:
        raise SlideMapError("dpos_team_map.yaml must contain a 'dpos_teams' key")

    teams: dict[str, DposTeamMapping] = {}
    for codename, entry in raw["dpos_teams"].items():
        if not isinstance(entry, dict) or "platform_team" not in entry:
            raise SlideMapError(
                f"DPOS mapping for '{codename}' must include 'platform_team'"
            )
        teams[str(codename).strip()] = DposTeamMapping(
            codename=str(codename).strip(),
            platform_team=str(entry["platform_team"]).strip(),
            notes=entry.get("notes"),
        )

    return DposTeamMapConfig(teams=teams)


def extract_dpos_codename(ppt_team_cell: str) -> str:
    """Extract codename prefix from a PPT team cell (before first newline)."""
    return ppt_team_cell.split("\n")[0].strip()
