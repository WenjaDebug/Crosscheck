"""Einstellungen, per Umgebungsvariable überschreibbar."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class Settings:
    # Der Datenordner: Übungsmappen, Prüfberichte, eigene Regeln. VOICEAGENT_DATA_DIR=data/test nimmt einen leeren.
    data_dir: Path = field(default_factory=lambda: Path(os.environ.get("VOICEAGENT_DATA_DIR", "data")))
