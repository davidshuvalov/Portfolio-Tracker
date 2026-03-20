# Portfolio Tracker v2 — Installation Guide

> Analytics for systematic futures traders using MultiWalk / MultiCharts

---

## System Requirements

| Requirement | Minimum |
|---|---|
| Operating System | Windows 10 (64-bit) or later |
| RAM | 4 GB (8 GB recommended for large portfolios) |
| Disk Space | 600 MB free |
| MultiCharts / MultiWalk | Any version with CSV export enabled |
| Internet connection | Required for first-time account activation only |

---

## Installation

### Option A — Windows Installer (Recommended)

This is the easiest option. It installs Portfolio Tracker like any other Windows application, creates a desktop shortcut, and supports clean uninstall.

1. **Download** `PortfolioTracker-v2.0.0-Setup.exe` from the link provided.

2. **Run the installer.**
   - If Windows shows a SmartScreen warning, click **More info → Run anyway**.
     *(The app is not yet code-signed; this is expected.)*
   - Accept the license agreement and choose an install folder (the default is fine).
   - Tick **"Create a desktop shortcut"** for easy access.
   - Click **Install**, then **Finish**.

3. **Launch.** Double-click the **Portfolio Tracker** desktop icon.
   - A terminal window will open briefly, then your default browser will open to
     `http://localhost:8501`
   - Leave the terminal window open while using the app — closing it stops the app.

---

### Option B — Portable Bundle (No Installation)

Use this if you cannot run an installer (e.g., on a locked-down corporate machine).

1. **Download** `PortfolioTracker-portable.zip` from the link provided.

2. **Extract** the ZIP to any folder (e.g., `C:\Tools\PortfolioTracker\`).

3. **Launch** by double-clicking `PortfolioTracker.exe` inside the extracted folder.
   - The app will open in your browser at `http://localhost:8501`.

> **Note:** Do not move the `.exe` outside its folder — it depends on the files next to it.

---

## First-Time Setup

When you launch Portfolio Tracker for the first time, you will be guided through a short setup wizard:

1. **Sign in or create an account.**
   An account is needed to activate your licence. Use the email address your licence was issued to.

2. **Set your MultiWalk data folder.**
   This is the folder where MultiWalk exports its strategy CSV files.
   Typical path: `C:\MultiCharts .NET\MultiWalk\Results\`

3. **Import your strategies.**
   Click **Import** on the sidebar to scan your MultiWalk folder and load your strategies.

4. **Explore your portfolio.**
   Navigate the pages in the left sidebar: Portfolio, Monte Carlo, Correlations, and more.

---

## Updating

To update to a new version, simply run the new installer — it will replace the old version automatically. Your settings and data are preserved.

For the portable bundle, extract the new ZIP to the same folder (overwrite when prompted).

---

## Uninstalling

**If installed via the installer:**

1. Open **Windows Settings → Apps** (or **Control Panel → Programs → Uninstall a program**).
2. Find **Portfolio Tracker** in the list and click **Uninstall**.

**If using the portable bundle:**

Simply delete the folder you extracted the ZIP into.

---

## Troubleshooting

### The app does not open in the browser

- Wait 10–15 seconds after launching — startup takes a moment on the first run.
- If nothing opens, manually navigate to `http://localhost:8501` in your browser.
- If that page is blank or gives an error, restart the app.

### Windows SmartScreen blocks the installer

This happens because the installer is not yet code-signed.

1. Click **More info** on the SmartScreen dialog.
2. Click **Run anyway**.

### The app crashes on startup / shows a Numba error

Numba (used for Monte Carlo simulations) compiles on the first run and may take up to a minute on older hardware. If it crashes:

1. Close the app.
2. Delete the Numba cache: `%LOCALAPPDATA%\numba_cache` (you can paste this path in File Explorer).
3. Relaunch the app.

### Port 8501 is already in use

Another application is using port 8501. Either close it, or set a different port:

1. Create a file called `.streamlit\config.toml` in the app install folder.
2. Add:
   ```toml
   [server]
   port = 8502
   ```
3. Relaunch the app, then navigate to `http://localhost:8502`.

### My MultiWalk data folder is not found

Make sure MultiWalk has exported at least one strategy to CSV. In MultiWalk, right-click a strategy result and choose **Export to CSV**. Then re-run the import in Portfolio Tracker.

### I forgot my password

On the login screen, click **Forgot password?** and follow the instructions sent to your email.

---

## Licence

Portfolio Tracker is commercial software. Your licence key is tied to the email address used during purchase. Each licence allows use on **one machine** at a time.

For licence questions, contact support at the email address where you received this installer.

---

## What's New in v2.0

- Rebuilt from scratch in Python with a modern browser-based interface
- Monte Carlo simulation with Numba JIT acceleration
- Walk-forward eligibility backtest
- Portfolio optimiser (contract sizing by ATR)
- Market analysis: ATR percentiles, sector correlations
- Leave-one-out impact analysis
- PDF and Excel export
- Cloud account sync (settings saved across reinstalls)
