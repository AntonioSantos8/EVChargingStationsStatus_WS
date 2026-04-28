# EV Scraper — Charging Station Data Collector

Automated data collector for EV charging stations using TupiMob's internal API. Runs on a 15-minute schedule, capturing connector type, power output, and real-time availability per station. Data is stored in a structured Excel file with utilization metrics, financial estimates, and an AI-generated analysis report.

---

## Installation

Python 3.10+ required.

```bash
pip install -r requirements.txt
```

---

## Usage

**Collect data** (runs every 15 minutes automatically):
```bash
python scraper.py
```

**Open the analysis dashboard:**
```bash
python run_server.py
```

Opens `http://localhost:8765/dashboard.html` in your browser automatically.

**Run analysis only** (without the dashboard server):
```bash
python analyzer.py
```

---

## Configuration

At the top of `scraper.py`:

```python
SEARCH_CITY = "Blumenau"  # city to search
INTERVAL    = 15          # collection interval in minutes
EXCEL_FILE  = "charging_stations.xlsx"
```

---

## AI Analysis Setup

The analyzer calls the Anthropic API to generate a written analysis. To enable it, add your API key in `analyzer.py`:

```python
ANTHROPIC_API_KEY = "YOUR_API_KEY_HERE"
```

Get a key at [console.anthropic.com](https://console.anthropic.com). Each analysis call costs roughly $0.003.

---

## Output

After running the dashboard, the following files are generated:

```
output/
  charts/
    utilization_over_time.png
    financials.png
  ai_analysis.txt
  dashboard_data.json
charging_stations.xlsx   ← Records / Stations / Analysis / Financial tabs
```

---

## Financial Model

Each "In Use" reading represents 15 minutes of active charging. Estimated values:

- **Revenue** = power (kW) × (15 min / 60) × price per kWh
- **Cost** = power (kW) × (15 min / 60) × energy cost per kWh
- **Profit** = Revenue − Cost

Price and cost per kWh are adjustable via sliders in the dashboard.

---

## Notes

- The scraper uses TupiMob's internal API discovered via browser DevTools. If the API changes or requires authentication in the future, the request headers in `scraper.py` may need to be updated.
- If the API returns an empty response, a `debug.json` file is saved automatically for inspection.
- Logs are written to `scraper.log`.
