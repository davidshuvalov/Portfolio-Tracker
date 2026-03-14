#!/usr/bin/env python3
"""
Golden fixture capture script.

Run this script against a real dataset to record "blessed" reference values
that regression tests can compare against.  The captured fixtures are saved
to tests/fixtures/golden_<name>.json.

Usage:
    cd v2
    python tests/fixtures/capture_golden.py \
        --folder /path/to/MultiWalk/base/folder \
        --date-format DMY \
        [--cutoff 2024-12-31]

After running, commit the generated JSON files so tests/unit/test_golden.py
can validate future runs against them.

What gets captured:
    golden_import.json      — strategy names, date range, trade count
    golden_summary.json     — per-strategy key metrics from compute_summary
    golden_portfolio.json   — portfolio aggregate metrics
    golden_mc.json          — MC summary metrics (with fixed RNG seed=42)
    golden_correlations.json — top-5 most-correlated pairs

The fixture schema is intentionally narrow: only the fields that are stable
across refactors are captured.  Volatile fields (timestamps, exact floats
with platform rounding) are excluded.
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import date, datetime
from pathlib import Path

# Add v2/ to path so we can import from core
sys.path.insert(0, str(Path(__file__).parent.parent.parent))


def _json_default(obj):
    if isinstance(obj, (date, datetime)):
        return obj.isoformat()
    if hasattr(obj, "item"):  # numpy scalar
        return obj.item()
    raise TypeError(f"Object of type {type(obj)} is not JSON serializable")


def capture_import(strategy_folders, date_format, use_cutoff, cutoff_date) -> dict:
    from core.ingestion.csv_importer import import_all
    print(f"  Importing {len(strategy_folders)} strategies…")
    imported, warnings = import_all(
        strategy_folders, date_format=date_format,
        use_cutoff=use_cutoff, cutoff_date=cutoff_date,
    )
    start, end = imported.date_range
    return {
        "strategy_names":  sorted(imported.strategy_names),
        "n_strategies":    len(imported.strategy_names),
        "date_start":      str(start),
        "date_end":        str(end),
        "n_trading_days":  len(imported.daily_m2m),
        "n_trades":        len(imported.trades),
        "n_warnings":      len(warnings),
    }


def capture_summary(imported, strategy_folders, date_format) -> dict:
    from core.portfolio.summary import compute_summary
    print("  Computing summary metrics…")
    df = compute_summary(imported, strategy_folders, date_format=date_format)
    out = {}
    _KEY_COLS = [
        "expected_annual_profit", "profit_since_oos_start",
        "max_drawdown_isoos", "sharpe_isoos", "oos_win_rate",
    ]
    for name in df.index:
        row = {}
        for col in _KEY_COLS:
            if col in df.columns:
                v = df.loc[name, col]
                try:
                    row[col] = None if (v != v) else round(float(v), 4)
                except (TypeError, ValueError):
                    row[col] = str(v)
        out[name] = row
    return out


def capture_portfolio(imported, strategy_folders, date_format, config) -> dict:
    from core.portfolio.aggregator import build_portfolio
    print("  Building portfolio…")
    portfolio = build_portfolio(imported, config)
    pnl = portfolio.daily_pnl.sum(axis=1)
    equity = pnl.cumsum()
    rolling_max = equity.cummax()
    dd = equity - rolling_max
    max_dd_pct = float((dd / rolling_max[dd.idxmin()]).min()) if len(dd) > 0 else 0.0
    return {
        "n_active_strategies": len(portfolio.strategies),
        "total_pnl":  round(float(pnl.sum()), 2),
        "max_dd_pct": round(max_dd_pct, 6),
        "n_days":     len(pnl),
    }


def capture_mc(imported, config) -> dict:
    import numpy as np
    from core.analytics.monte_carlo import run_monte_carlo
    from core.config import MCConfig
    from core.portfolio.aggregator import build_portfolio
    print("  Running MC with seed=42…")
    portfolio = build_portfolio(imported, config)
    from core.portfolio.aggregator import portfolio_total_pnl
    m2m_total = portfolio_total_pnl(portfolio)

    np.random.seed(42)  # deterministic for golden capture
    mc_cfg = MCConfig(simulations=5_000, period="IS+OOS", risk_ruin_target=0.10)
    result = run_monte_carlo(
        daily_m2m=m2m_total,
        config=mc_cfg,
        margin_threshold=5_000.0,
        return_scenarios=False,
    )
    return {
        "starting_equity":  round(result.starting_equity, 0),
        "expected_profit":  round(result.expected_profit, 0),
        "risk_of_ruin":     round(result.risk_of_ruin, 4),
        "max_drawdown_pct": round(result.max_drawdown_pct, 4),
        "sharpe_ratio":     round(result.sharpe_ratio, 4),
    }


def capture_correlations(imported, config) -> dict:
    from core.analytics.correlations import (
        CorrelationMode, compute_correlation_matrix, get_correlation_pairs
    )
    from core.portfolio.aggregator import build_portfolio
    print("  Computing correlations…")
    portfolio = build_portfolio(imported, config)
    daily_pnl = portfolio.daily_pnl
    corr_matrix = compute_correlation_matrix(daily_pnl, CorrelationMode.NORMAL)
    pairs = get_correlation_pairs(corr_matrix)
    top5 = pairs.nlargest(5, "correlation")[["strategy_a", "strategy_b", "correlation"]].round(4)
    return {
        "n_strategies":    len(corr_matrix),
        "top5_pairs": top5.to_dict(orient="records"),
    }


def main():
    parser = argparse.ArgumentParser(description="Capture golden fixtures")
    parser.add_argument("--folder",       required=True, help="MultiWalk base folder(s)", nargs="+")
    parser.add_argument("--date-format",  default="DMY",  choices=["DMY", "MDY"])
    parser.add_argument("--cutoff",       default=None,   help="YYYY-MM-DD cutoff date")
    parser.add_argument("--output-dir",   default=str(Path(__file__).parent),
                        help="Directory to write fixture JSON files")
    args = parser.parse_args()

    from pathlib import Path as P
    from core.config import AppConfig
    from core.ingestion.folder_scanner import scan_folders

    config = AppConfig(folders=[P(f) for f in args.folder], date_format=args.date_format)
    use_cutoff = args.cutoff is not None
    cutoff_date = date.fromisoformat(args.cutoff) if args.cutoff else None

    print("Scanning folders…")
    scan = scan_folders(config.folders)
    if scan.errors:
        for e in scan.errors:
            print(f"  ERROR: {e}")
        sys.exit(1)
    print(f"Found {len(scan.strategies)} strategy folders")

    from core.ingestion.csv_importer import import_all
    imported, _ = import_all(
        scan.strategies, date_format=args.date_format,
        use_cutoff=use_cutoff, cutoff_date=cutoff_date,
    )

    out_dir = P(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    fixtures = {
        "golden_import":       capture_import(scan.strategies, args.date_format, use_cutoff, cutoff_date),
        "golden_summary":      capture_summary(imported, scan.strategies, args.date_format),
        "golden_portfolio":    capture_portfolio(imported, scan.strategies, args.date_format, config),
        "golden_mc":           capture_mc(imported, config),
        "golden_correlations": capture_correlations(imported, config),
    }

    for name, data in fixtures.items():
        path = out_dir / f"{name}.json"
        path.write_text(json.dumps(data, indent=2, default=_json_default))
        print(f"  Wrote {path}")

    print("\nGolden fixtures captured successfully.")
    print("Commit these files and run: pytest tests/unit/test_golden.py")


if __name__ == "__main__":
    main()
