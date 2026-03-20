# Portfolio Tracker v2

**Analytics for systematic futures traders using MultiWalk / MultiCharts.**

Portfolio Tracker aggregates strategy data from MultiWalk, tracks in-sample and out-of-sample performance, estimates portfolio margin from historical positions, and provides a full analytics suite: correlations, diversification, Monte Carlo simulations, walk-forward eligibility backtests, and more.

---

## Features

| Feature | Description |
|---|---|
| Strategy import | Scans your MultiWalk CSV export folder automatically |
| Portfolio summary | Aggregated P&L, profit factor, Sharpe, drawdown metrics |
| Monte Carlo | Randomised trade-sequence simulation with confidence bands |
| Correlations | Pearson correlation heatmap across strategies |
| Diversification | Concentration and diversification score analytics |
| Leave-one-out | Isolate the impact of removing any single strategy |
| Eligibility backtest | Walk-forward rule validation across your strategy history |
| Portfolio optimiser | ATR-based contract sizing recommendation |
| Market analysis | ATR percentiles and sector correlation by market |
| Margin tracking | Live and historical portfolio margin estimation |
| PDF / Excel export | One-click report generation |
| Cloud sync | Settings and portfolio config persist across reinstalls |

---

## Installing (End Users)

See **[INSTALL.md](INSTALL.md)** for the full installation guide, including:

- Windows installer (recommended)
- Portable bundle (no install required)
- First-time setup walkthrough
- Troubleshooting

**Quick start:**

1. Run `PortfolioTracker-v2.0.0-Setup.exe`
2. Launch Portfolio Tracker from the desktop shortcut
3. Sign in and point the app at your MultiWalk data folder
4. Click **Import** — your strategies load automatically

---

## Building from Source (Developers)

### Prerequisites

- Windows 10+ (64-bit)
- Python 3.11+
- [Inno Setup 6](https://jrsoftware.org/isinfo.php) (for the installer)

### Setup

```bat
git clone https://github.com/davidshuvalov/Portfolio-Tracker.git
cd Portfolio-Tracker
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

### Run in development mode

```bat
streamlit run app.py
```

The app opens at `http://localhost:8501`.

### Build the Windows installer

```bat
pip install pyinstaller pyinstaller-hooks-contrib
build_windows.bat
```

This produces:
- `dist\PortfolioTracker\PortfolioTracker.exe` — portable bundle
- `installer\Output\PortfolioTracker-v2.0.0-Setup.exe` — installer (requires Inno Setup)

### Run tests

```bat
pytest
```

691+ tests covering all analytics modules.

---

## Project Structure

```
Portfolio-Tracker/
├── app.py                      # Streamlit entry point
├── run.py                      # PyInstaller launcher
├── build_windows.bat           # Build script (PyInstaller + Inno Setup)
├── portfolio_tracker.spec      # PyInstaller spec
├── installer/
│   └── setup.iss               # Inno Setup installer script
├── core/                       # Analytics engine (no UI dependencies)
│   ├── ingestion/              # CSV / XLSB import pipeline
│   ├── portfolio/              # Portfolio aggregation & metrics
│   ├── analytics/              # Monte Carlo, correlations, diversification, …
│   ├── reporting/              # PDF & Excel export
│   └── licensing/              # Licence validation
├── ui/                         # Streamlit pages and components
├── auth/                       # Supabase authentication
├── backend/                    # FastAPI billing backend (optional cloud)
├── config/                     # Default settings YAML
├── tests/                      # 691+ unit & integration tests
└── INSTALL.md                  # End-user installation guide
```

---

## Architecture

```
Browser (localhost:8501)
    │
    ▼
Streamlit UI  (ui/)
    │
    ▼
Analytics Engine  (core/)
    ├── Ingestion   — reads MultiWalk CSV exports
    ├── Portfolio   — aggregates metrics & equity curves
    ├── Analytics   — MC, correlations, diversification, LOO, ATR
    └── Reporting   — PDF / Excel output
    │
    ▼
Cloud (optional)
    ├── Supabase    — auth + settings sync
    └── Stripe      — subscription billing
```

The core analytics engine has **no UI dependencies** and can be used as a standalone library.

---

## Deployment (Cloud)

See **[DEPLOYMENT.md](DEPLOYMENT.md)** for the full cloud deployment guide covering Supabase, Stripe, Render, and Streamlit Community Cloud.

---

## Legacy VBA Version

The original Excel/VBA version (v1.24) is included as
`Portfolio Tracker - A Tool for MultiWalk v1.24.xlsb`.
All functionality has been ported and extended in v2. The VBA source is available
as individual `.bas` / `.cls` files in the repo root for reference.

---

## Licence

Portfolio Tracker is commercial software. See INSTALL.md for licence details.
