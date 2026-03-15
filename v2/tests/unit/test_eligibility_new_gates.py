"""
Tests for Sprint A eligibility gates:
  - exclude_buy_and_hold
  - exclude_previously_quit
  - backtest_data_scope (OOS vs IS+OOS window slicing)

Also tests the new PortfolioContractConfig (Sprint B) round-trips through YAML.
"""

import sys
from pathlib import Path
from datetime import date

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent.parent))

import pandas as pd

from core.config import AppConfig, EligibilityConfig, PortfolioContractConfig
from core.portfolio.summary import apply_eligibility_rules, _compute_dynamic_metrics


# ── Helpers ───────────────────────────────────────────────────────────────────

def _summary(strategies: dict[str, dict]) -> pd.DataFrame:
    """Build a minimal summary DataFrame."""
    return pd.DataFrame(strategies).T


def _bare_elig(**kwargs) -> EligibilityConfig:
    """EligibilityConfig with all profit/efficiency rules disabled — tests gates in isolation."""
    return EligibilityConfig(
        profit_1m=False, profit_3m=False, profit_6m=False, profit_3or6m=False,
        profit_9m=False, profit_12m=False, profit_oos=False,
        efficiency_1m=False, efficiency_3m=False, efficiency_6m=False,
        efficiency_9m=False, efficiency_12m=False, efficiency_oos=False,
        loss_1m=False, loss_3m=False, loss_6m=False,
        efficiency_loss_1m=False, efficiency_loss_3m=False, efficiency_loss_6m=False,
        use_incubation=False, use_quitting=False, use_count_monthly_profits=False,
        **kwargs,
    )


# ══════════════════════════════════════════════════════════════════════════════
# Buy & Hold exclusion gate
# ══════════════════════════════════════════════════════════════════════════════

class TestExcludeBuyAndHold:
    def test_bh_excluded_when_flag_true(self):
        df = _summary({
            "TrendA":    {"status": "Live",      "quitting_date": None},
            "BH_Bench":  {"status": "Buy&Hold",  "quitting_date": None},
        })
        elig = _bare_elig(exclude_buy_and_hold=True)
        result = apply_eligibility_rules(df, elig)
        assert result["TrendA"] == True
        assert result["BH_Bench"] == False

    def test_bh_included_when_flag_false(self):
        df = _summary({
            "TrendA":   {"status": "Live",     "quitting_date": None},
            "BH_Bench": {"status": "Buy&Hold", "quitting_date": None},
        })
        elig = _bare_elig(exclude_buy_and_hold=False)
        result = apply_eligibility_rules(df, elig)
        assert result["BH_Bench"] == True

    def test_case_insensitive_buy_hold_detection(self):
        df = _summary({
            "A": {"status": "buy and hold",    "quitting_date": None},
            "B": {"status": "BUY & HOLD",      "quitting_date": None},
            "C": {"status": "Buy&Hold",         "quitting_date": None},
        })
        elig = _bare_elig(exclude_buy_and_hold=True)
        result = apply_eligibility_rules(df, elig)
        assert result["A"] == False
        assert result["B"] == False
        assert result["C"] == False

    def test_live_strategy_not_excluded(self):
        df = _summary({"A": {"status": "Live", "quitting_date": None}})
        elig = _bare_elig(exclude_buy_and_hold=True)
        result = apply_eligibility_rules(df, elig)
        assert result["A"] == True

    def test_no_status_column_no_crash(self):
        """No 'status' column → gate silently skips."""
        df = _summary({"A": {"profit_last_1_month": 100.0}})
        elig = _bare_elig(exclude_buy_and_hold=True)
        result = apply_eligibility_rules(df, elig)
        assert result["A"] == True


# ══════════════════════════════════════════════════════════════════════════════
# Previously-quit exclusion gate
# ══════════════════════════════════════════════════════════════════════════════

class TestExcludePreviouslyQuit:
    def test_strategy_with_quit_date_excluded(self):
        df = _summary({
            "Clean":  {"status": "Live", "quitting_date": None},
            "Quit1":  {"status": "Live", "quitting_date": date(2022, 6, 1)},
        })
        elig = _bare_elig(exclude_previously_quit=True)
        result = apply_eligibility_rules(df, elig)
        assert result["Clean"] == True
        assert result["Quit1"] == False

    def test_recovered_strategy_still_excluded(self):
        """'Recovered' status but quitting_date set → still excluded."""
        df = _summary({
            "Recovered": {"status": "Live", "quitting_date": date(2021, 1, 1)},
        })
        elig = _bare_elig(exclude_previously_quit=True)
        result = apply_eligibility_rules(df, elig)
        assert result["Recovered"] == False

    def test_flag_false_includes_all(self):
        df = _summary({
            "A": {"status": "Live", "quitting_date": date(2022, 6, 1)},
        })
        elig = _bare_elig(exclude_previously_quit=False)
        result = apply_eligibility_rules(df, elig)
        assert result["A"] == True

    def test_no_quitting_date_column_no_crash(self):
        df = _summary({"A": {"status": "Live"}})
        elig = _bare_elig(exclude_previously_quit=True)
        result = apply_eligibility_rules(df, elig)
        assert result["A"] == True


# ══════════════════════════════════════════════════════════════════════════════
# Data scope: OOS vs IS+OOS
# ══════════════════════════════════════════════════════════════════════════════

class TestDataScope:
    def _make_pnl(self, is_values: list[float], oos_values: list[float],
                  is_start="2019-01-02", oos_start="2021-01-04"):
        """Build a PnL series with IS and OOS segments."""
        is_dates = pd.bdate_range(is_start, periods=len(is_values))
        oos_dates = pd.bdate_range(oos_start, periods=len(oos_values))
        idx = is_dates.append(oos_dates)
        vals = is_values + oos_values
        return (
            pd.Series(vals, index=idx),
            is_dates[-1].date(),
            pd.Timestamp(oos_start).date(),
            oos_dates[-1].date(),
        )

    def test_oos_scope_uses_only_oos_data(self):
        """With data_scope=OOS, 1M window only counts OOS data (all negative)."""
        # IS: 2 years of +200/day so IS window would be very positive
        # OOS: 65 days of -50/day (enough for a 1m window to land within OOS)
        is_pnl = [200.0] * 504           # 2 years IS
        oos_pnl = [-50.0] * 65           # ~3 months OOS at -50/day
        pnl, _, oos_start, oos_end = self._make_pnl(is_pnl, oos_pnl)

        result = _compute_dynamic_metrics(
            pnl=pnl,
            oos_begin=oos_start, oos_end=oos_end,
            expected_annual_profit=50_000.0, annual_sd_is=1000.0,
            is_max_drawdown=5_000.0, days_threshold=0,
            incubation_months=6, min_incubation_ratio=1.0,
            eligibility_months=12, data_scope="OOS",
        )
        # 1M window is entirely in OOS (-50/day) → negative profit
        assert result["profit_last_1_month"] is not None
        assert result["profit_last_1_month"] < 0

    def test_isoos_scope_includes_is_data_in_windows(self):
        """With data_scope=IS+OOS, the 12M window reaches into profitable IS data."""
        # IS: 2 years of +200/day, OOS: 65 days of -50/day
        # OOS-only 12M: only 65 OOS days → sum = 65 × -50 = -3250
        # IS+OOS 12M: ~252 days of mix (+200 for most, -50 tail) → positive
        is_pnl = [200.0] * 504
        oos_pnl = [-50.0] * 65
        pnl, _, oos_start, oos_end = self._make_pnl(is_pnl, oos_pnl)

        result_oos = _compute_dynamic_metrics(
            pnl=pnl,
            oos_begin=oos_start, oos_end=oos_end,
            expected_annual_profit=50_000.0, annual_sd_is=1000.0,
            is_max_drawdown=5_000.0, days_threshold=0,
            incubation_months=6, min_incubation_ratio=1.0,
            eligibility_months=12, data_scope="OOS",
        )
        result_isoos = _compute_dynamic_metrics(
            pnl=pnl,
            oos_begin=oos_start, oos_end=oos_end,
            expected_annual_profit=50_000.0, annual_sd_is=1000.0,
            is_max_drawdown=5_000.0, days_threshold=0,
            incubation_months=6, min_incubation_ratio=1.0,
            eligibility_months=12, data_scope="IS+OOS",
        )
        # IS+OOS 12M window reaches into IS (+200/day) → much higher profit
        p_oos   = result_oos.get("profit_last_12_months") or 0.0
        p_isoos = result_isoos.get("profit_last_12_months") or 0.0
        assert p_isoos > p_oos


# ══════════════════════════════════════════════════════════════════════════════
# PortfolioContractConfig (Sprint B)
# ══════════════════════════════════════════════════════════════════════════════

class TestPortfolioContractConfig:
    def test_defaults(self):
        cfg = PortfolioContractConfig()
        assert cfg.starting_equity == pytest.approx(705_000.0)
        assert cfg.cease_type == "Percentage"
        assert cfg.cease_trading_threshold == pytest.approx(0.25)
        assert cfg.contract_ratio_margin_atr == pytest.approx(0.50)
        assert cfg.contract_size_pct_equity == pytest.approx(0.01)
        assert cfg.atr_window == "ATR Last 3 Months"
        assert cfg.reweight_scope == "All"
        assert cfg.reweight_gain == pytest.approx(1.0)

    def test_appconfig_has_contract_sizing_field(self):
        cfg = AppConfig()
        assert hasattr(cfg, "contract_sizing")
        assert isinstance(cfg.contract_sizing, PortfolioContractConfig)

    def test_mc_config_new_fields(self):
        from core.config import MCConfig
        cfg = MCConfig()
        assert cfg.output_samples == 50
        assert cfg.remove_best_pct == pytest.approx(0.02)
        assert cfg.solve_for_ror is False

    def test_config_round_trips_yaml(self, tmp_path):
        """New fields survive save/load cycle."""
        import yaml

        cfg = AppConfig()
        cfg.contract_sizing.starting_equity = 500_000.0
        cfg.contract_sizing.cease_trading_threshold = 0.30
        cfg.contract_sizing.atr_window = "ATR Last 6 Months"
        cfg.contract_sizing.reweight_scope = "Index Only"
        cfg.contract_sizing.reweight_gain  = 1.25
        cfg.monte_carlo.output_samples = 100
        cfg.monte_carlo.solve_for_ror = True
        cfg.eligibility.exclude_buy_and_hold = True
        cfg.eligibility.exclude_previously_quit = True
        cfg.eligibility.backtest_data_scope = "IS+OOS"
        cfg.ranking.metric = "sharpe_isoos"
        cfg.ranking.group_by_sector = False
        cfg.ranking.eligible_only = False

        # Dump and reload via raw YAML
        data = cfg.model_dump(mode="json")
        data["folders"] = [str(p) for p in cfg.folders]
        yaml_str = yaml.dump(data, default_flow_style=False)
        reloaded_data = yaml.safe_load(yaml_str)
        cfg2 = AppConfig.model_validate(reloaded_data)

        assert cfg2.contract_sizing.starting_equity == pytest.approx(500_000.0)
        assert cfg2.contract_sizing.atr_window == "ATR Last 6 Months"
        assert cfg2.contract_sizing.reweight_scope == "Index Only"
        assert cfg2.contract_sizing.reweight_gain  == pytest.approx(1.25)
        assert cfg2.monte_carlo.output_samples == 100
        assert cfg2.monte_carlo.solve_for_ror is True
        assert cfg2.eligibility.exclude_buy_and_hold is True
        assert cfg2.eligibility.exclude_previously_quit is True
        assert cfg2.eligibility.backtest_data_scope == "IS+OOS"
        assert cfg2.ranking.metric == "sharpe_isoos"
        assert cfg2.ranking.group_by_sector is False
        assert cfg2.ranking.eligible_only is False
