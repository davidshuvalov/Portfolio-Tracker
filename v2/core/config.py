"""
Application configuration — replaces all Excel named ranges.
Loaded from config/default_settings.yaml and saved to
~/.portfolio_tracker/settings.yaml (per-user overrides).
"""

from __future__ import annotations
from pathlib import Path
from typing import Literal

import yaml
from pydantic import BaseModel, Field, field_validator


CONFIG_DIR = Path.home() / ".portfolio_tracker"
USER_CONFIG_FILE = CONFIG_DIR / "settings.yaml"
DEFAULT_CONFIG_FILE = Path(__file__).parent.parent / "config" / "default_settings.yaml"


class MCConfig(BaseModel):
    simulations: int = 10_000
    period: Literal["IS", "OOS", "IS+OOS"] = "OOS"
    risk_ruin_target: float = 0.10
    risk_ruin_tolerance: float = 0.01
    trade_adjustment: float = 0.0
    trade_option: Literal["Closed", "M2M"] = "Closed"


class EligibilityConfig(BaseModel):
    days_threshold_oos: int = 0
    oos_dd_vs_is_cap: float = 1.5       # 0 = disabled
    status_include: list[str] = Field(default_factory=lambda: ["Live"])
    efficiency_ratio: float = 0.5
    date_type: Literal["OOS Start Date", "Incubation Pass Date"] = "OOS Start Date"
    enable_sector_analysis: bool = True
    enable_symbol_analysis: bool = True
    max_horizon: int = 12


class PortfolioConfig(BaseModel):
    period_years: float = 3.0
    cutoff_date: str | None = None
    use_cutoff: bool = False
    buy_and_hold_status: str = "Buy&Hold"
    live_status: str = "Live"
    pass_status: str = "Pass"


class AppConfig(BaseModel):
    folders: list[Path] = Field(default_factory=list)
    date_format: Literal["DMY", "MDY"] = "DMY"
    portfolio: PortfolioConfig = Field(default_factory=PortfolioConfig)
    monte_carlo: MCConfig = Field(default_factory=MCConfig)
    eligibility: EligibilityConfig = Field(default_factory=EligibilityConfig)
    corr_normal_threshold: float = 0.70
    corr_negative_threshold: float = 0.30
    corr_drawdown_threshold: float = 0.70
    symbol_margins: dict[str, float] = Field(
        default_factory=dict,
        description="Per-symbol margin requirement in $ (e.g. {'ES': 12000, 'NQ': 18000})",
    )
    default_margin: float = 5000.0

    @field_validator("folders", mode="before")
    @classmethod
    def coerce_paths(cls, v: list) -> list[Path]:
        return [Path(p) for p in v]

    # ── Persistence ────────────────────────────────────────────

    @classmethod
    def load(cls) -> "AppConfig":
        """Load config: defaults first, then merge user overrides."""
        data: dict = {}

        if DEFAULT_CONFIG_FILE.exists():
            with open(DEFAULT_CONFIG_FILE) as f:
                data = yaml.safe_load(f) or {}

        if USER_CONFIG_FILE.exists():
            with open(USER_CONFIG_FILE) as f:
                overrides = yaml.safe_load(f) or {}
            data = _deep_merge(data, overrides)

        return cls.model_validate(data)

    def save(self) -> None:
        """Persist current config to the user config file."""
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        data = self.model_dump(mode="json")
        # Convert Path objects to strings for YAML serialisation
        data["folders"] = [str(p) for p in self.folders]
        with open(USER_CONFIG_FILE, "w") as f:
            yaml.dump(data, f, default_flow_style=False, sort_keys=False)

    def add_folder(self, folder: Path) -> None:
        if folder not in self.folders:
            self.folders.append(folder)
            self.save()

    def remove_folder(self, folder: Path) -> None:
        self.folders = [f for f in self.folders if f != folder]
        self.save()


def _deep_merge(base: dict, overrides: dict) -> dict:
    """Recursively merge overrides into base dict."""
    result = base.copy()
    for key, val in overrides.items():
        if key in result and isinstance(result[key], dict) and isinstance(val, dict):
            result[key] = _deep_merge(result[key], val)
        else:
            result[key] = val
    return result
