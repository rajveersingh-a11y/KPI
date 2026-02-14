# KPI Dashboards (9 Dashboards, 105+ KPIs)

## Quick start

1. **Generate data** (if not already done):
   ```bash
   python export_dashboard_data.py
   ```
   This creates `dashboards.json` with calculated KPI values and chart data.

2. **Run the frontend**:
   ```bash
   python serve.py
   ```
   This starts a local server and opens **http://localhost:8080** in your browser.

   Alternatively:
   ```bash
   python -m http.server 8080
   ```
   Then open http://localhost:8080 in your browser.

## What you get

- **9 dashboards** (D-1 … D-9) in the top nav. Each dashboard shows:
  - **KPI cards**: all KPIs for that dashboard with name, value, and unit.
  - **Charts**: where it fits (loss trend, SAIDI/SAIFI, outage metrics, tamper by type, voltage/power quality, theft/revenue, communication, mapping, anomaly/phase).

- **Dashboard-1 (Finance)**: Feeder/DT/LT loss %, Billing/Collection efficiency, AT&C loss, top high-loss counts; **Loss % trend** (line), **Billing & collection** (doughnut).

- **Dashboard-2 (Operation + Analytics)**: SAIDI, SAIFI, CAIDI, CAIFI, MAIFI, outages, MTTR, reliability scores; **SAIDI/SAIFI trend** (line), **Outage & response** (bar).

- **Dashboard-3 (Technical)**: DT loading %, loading bands, overloaded DTs/feeders, load violation; **Loading metrics** (bar).

- **Dashboard-4 (Operation)**: Voltage deviation, VDI, FDI, power factor, unbalance; **Voltage & power quality** (bar).

- **Dashboard-5 (Analytics)**: Tamper alerts by type, anomalies, energy gap; **Tamper by type** (bar).

- **Dashboard-6 (Analytics + Finance)**: Theft flags, reduction %, revenue recovery; **Theft & revenue** (bar).

- **Dashboard-7 (Technical + Analytics)**: Signal strength, packet loss, retry counts, non-reporting meters; **Communication health** (bar).

- **Dashboard-8 (Advanced Analytics)**: Mapping accuracy, assets tracked, verification, correction cycle; **Mapping & verification** (bar).

- **Dashboard-9 (Analytics)**: Tamper sequence, phase imbalance, overload/MD risk, anomalies; **Anomaly & phase metrics** (bar).

Data is dummy/calculated from the same logic as `generate_kpi_data.py` and the Excel export.
