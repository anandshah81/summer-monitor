# Summer Strength Monitor

India weather tracking dashboard for consumer equity research. Tracks temperature patterns across 30 major Indian cities to assess summer season intensity for beverage, cooling, and FMCG demand.

## Live Dashboard

**[View Dashboard →](https://anandshah81.github.io/summer-monitor/)**

Auto-refreshes daily at 6 AM IST via GitHub Actions.

## Features

- **30 cities** across North, West, Central, South, and East India
- **Jan–Jun full pre-monsoon tracking** with Mar+ summary metrics
- **Summer Strength Index (SSI)** — YoY composite comparing temperature, hot days, and rainfall
- **5yr baseline deviation** — how current temps compare to 2020–2024 normals
- **Summer Onset Tracker** — first date each city crosses 30°C and 35°C, with YoY comparison
- **Heating Rate** — weekly temperature climb showing how fast summer is building
- **Cumulative Hot Days** — running total of city-days ≥30°C and ≥35°C
- **Monthly Heatmap** — city × month deviation from normal
- **Excel export** — download all data for custom analysis

## Data Sources

- **Open-Meteo** (29 cities) — ERA5 reanalysis + high-resolution weather models. Free, no API key needed.
- **Visual Crossing** (Mumbai only) — METAR station data from Santa Cruz airport. Required because Open-Meteo's 14km grid cannot resolve Mumbai's narrow peninsula.

## Setup

### Run locally

```bash
pip install requests
python summer_monitor.py --vc-key YOUR_VISUAL_CROSSING_KEY
```

### GitHub Pages (auto-refresh)

1. Fork/create this repo
2. Go to **Settings → Secrets → Actions** → Add `VC_KEY` with your Visual Crossing API key
3. Go to **Settings → Pages** → Set source to `gh-pages` branch
4. Go to **Actions** → Run "Refresh Summer Monitor" manually for first build
5. Dashboard auto-refreshes daily at 6 AM IST

Get a free Visual Crossing key at [visualcrossing.com/sign-up](https://www.visualcrossing.com/sign-up)

## Local Usage

```bash
# Standard run
python summer_monitor.py

# With validation report
python summer_monitor.py --validate

# With explicit VC key
python summer_monitor.py --vc-key YOUR_KEY

# Or set environment variable
set VC_KEY=YOUR_KEY          # Windows
export VC_KEY=YOUR_KEY       # Mac/Linux
python summer_monitor.py
```

Output: `summer_strength_monitor.html` — opens automatically in browser.
