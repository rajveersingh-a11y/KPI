"""
Export KPI data per dashboard to JSON for frontend (values + trend data for charts).
"""
import json
import random
from pathlib import Path

# Reuse same value logic as generate_kpi_data
def generate_value(vtype, lo, hi, unit):
    if vtype == "pct":
        return round(random.uniform(lo, hi), 2)
    if vtype == "count":
        return random.randint(int(lo), int(hi))
    if vtype == "index":
        return round(random.uniform(lo, hi), 3)
    if vtype == "minutes":
        return round(random.uniform(lo, hi), 1)
    if vtype == "kva":
        return random.choice([25, 63, 100, 160, 200, 250, 315])
    if vtype == "score":
        return random.randint(int(lo), int(hi))
    return round(random.uniform(lo, hi), 2)

# KPI_SPECS: (Dashboard, Department, KPI Name, value_type, min_val, max_val, unit)
KPI_SPECS = [
    ("Dashboard-1", "Finance", "Feeder Loss (%)", "pct", 3, 18, "%"),
    ("Dashboard-1", "Finance", "DT (Distribution Transformer) Loss (%)", "pct", 2, 12, "%"),
    ("Dashboard-1", "Finance", "LT Loss (%)", "pct", 1, 8, "%"),
    ("Dashboard-1", "Finance", "Billing Efficiency (%)", "pct", 78, 98, "%"),
    ("Dashboard-1", "Finance", "Collection Efficiency (%)", "pct", 72, 96, "%"),
    ("Dashboard-1", "Finance", "AT&C Loss (%)", "pct", 8, 28, "%"),
    ("Dashboard-1", "Finance", "Top X Best/Worst Feeders/DTs", "count", 5, 20, "count"),
    ("Dashboard-1", "Finance", "Top High Loss DTs / Feeders", "count", 8, 35, "count"),
    ("Dashboard-1", "Finance", "Top High-Loss Feeders / DTs", "count", 6, 28, "count"),
    ("Dashboard-2", "Operation", "SAIDI", "minutes", 45, 380, "min"),
    ("Dashboard-2", "Operation", "SAIFI", "index", 2, 25, "interruptions"),
    ("Dashboard-2", "Operation", "CAIDI", "minutes", 25, 95, "min"),
    ("Dashboard-2", "Operation", "CAIFI", "index", 1, 18, "interruptions"),
    ("Dashboard-2", "Operation", "MAIFI", "index", 0.2, 8, "interruptions"),
    ("Dashboard-2", "Operation", "Number of Outages (Frequency)", "count", 12, 450, "count"),
    ("Dashboard-2", "Operation", "Duration of Outages (Minutes)", "minutes", 120, 7200, "min"),
    ("Dashboard-2", "Operation", "DT/Feeder Reliability Trends (Monthly/Yearly)", "pct", 85, 99.5, "%"),
    ("Dashboard-2", "Operation", "DTs with High Failure Rate", "count", 3, 45, "count"),
    ("Dashboard-2", "Operation", "Detection Accuracy", "pct", 82, 98, "%"),
    ("Dashboard-2", "Operation", "False Positive Rate", "pct", 1, 15, "%"),
    ("Dashboard-2", "Operation", "Field inspection hit-rate", "pct", 65, 92, "%"),
    ("Dashboard-2", "Operation", "MTTI", "minutes", 8, 95, "min"),
    ("Dashboard-2", "Operation", "MTTR", "minutes", 25, 180, "min"),
    ("Dashboard-2", "Operation", "Alert response time", "minutes", 5, 45, "min"),
    ("Dashboard-2", "Operation", "Planned outage suppression rate", "pct", 70, 98, "%"),
    ("Dashboard-2", "Analytics", "Low-voltage pockets", "count", 2, 28, "count"),
    ("Dashboard-2", "Operation", "Feeders with Maximum Outages", "count", 4, 22, "count"),
    ("Dashboard-2", "Operation", "Reliability Improvement Trend", "pct", 2, 18, "%"),
    ("Dashboard-2", "Operation", "Consumer Service Reliability Score", "score", 72, 95, "score"),
    ("Dashboard-2", "Operation", "Composite Reliability Score", "score", 68, 94, "score"),
    ("Dashboard-2", "Operation", "Composite Efficiency Score", "score", 65, 92, "score"),
    ("Dashboard-3", "Technical", "% DT Peak Loading", "pct", 45, 98, "%"),
    ("Dashboard-3", "Technical", "% DT Loading", "pct", 38, 92, "%"),
    ("Dashboard-3", "Technical", "DT Load (kVA)", "kva", 25, 315, "kVA"),
    ("Dashboard-3", "Technical", "% Loading Bands", "pct", 0, 100, "%"),
    ("Dashboard-3", "Technical", "Top Overloaded DTs / Feeders", "count", 5, 30, "count"),
    ("Dashboard-3", "Technical", "Load Rise Trend", "pct", 2, 22, "%"),
    ("Dashboard-3", "Technical", "Consumers exceeding sanctioned load", "count", 15, 380, "count"),
    ("Dashboard-3", "Technical", "% Consumers with Load Violation", "pct", 0.5, 12, "%"),
    ("Dashboard-3", "Technical", "Load Duration Curve & Asset Loading Spread", "pct", 55, 88, "%"),
    ("Dashboard-3", "Technical", "DT Failure Rate (%)", "pct", 0.2, 5.5, "%"),
    ("Dashboard-3", "Technical", "Top Overloaded Assets", "count", 8, 42, "count"),
    ("Dashboard-3", "Technical", "Top Power Quality Issues", "count", 6, 35, "count"),
    ("Dashboard-4", "Operation", "Voltage Deviation (%)", "pct", 1, 12, "%"),
    ("Dashboard-4", "Operation", "Voltage Deviation Index (VDI)", "index", 0.02, 0.95, "index"),
    ("Dashboard-4", "Operation", "Frequency Deviation Index (FDI)", "index", 0.01, 0.35, "index"),
    ("Dashboard-4", "Operation", "Voltage Fluctuation Index", "index", 0.01, 0.45, "index"),
    ("Dashboard-4", "Operation", "Voltage Unbalance Index", "index", 0.02, 0.28, "index"),
    ("Dashboard-4", "Operation", "Voltage Drop (V)", "count", 5, 45, "V"),
    ("Dashboard-4", "Operation", "Low Power Factor (%) by DT/Feeder", "pct", 60, 92, "%"),
    ("Dashboard-4", "Operation", "Meter Current Unbalance (%)", "pct", 2, 18, "%"),
    ("Dashboard-4", "Operation", "% Time beyond voltage tolerance band", "pct", 0.5, 15, "%"),
    ("Dashboard-4", "Operation", "% Time with unacceptable current imbalance (>10%)", "pct", 1, 22, "%"),
    ("Dashboard-5", "Analytics", "Number of Tamper Alerts (Cover Open)", "count", 2, 85, "count"),
    ("Dashboard-5", "Analytics", "Number of Tamper Alerts (External Magnet)", "count", 0, 42, "count"),
    ("Dashboard-5", "Analytics", "Number of Tamper Alerts (Neutral Disturbance)", "count", 1, 38, "count"),
    ("Dashboard-5", "Analytics", "Number of Tamper Alerts (Neutral Missing)", "count", 0, 25, "count"),
    ("Dashboard-5", "Analytics", "Consumption Comparison - Energy Gap (kWh)", "count", 120, 8500, "kWh"),
    ("Dashboard-5", "Analytics", "Total anomalies detected (by time period)", "count", 25, 420, "count"),
    ("Dashboard-5", "Analytics", "Anomalies by type", "count", 3, 12, "types"),
    ("Dashboard-5", "Analytics", "Anomalies by severity", "count", 2, 5, "levels"),
    ("Dashboard-5", "Analytics", "Anomalies by geography", "count", 5, 45, "zones"),
    ("Dashboard-5", "Analytics", "Anomaly trends (daily/weekly/monthly)", "pct", -15, 25, "%"),
    ("Dashboard-5", "Analytics", "Repeat anomaly tracking", "count", 3, 65, "count"),
    ("Dashboard-6", "Analytics", "Theft Suspect Flags", "count", 8, 120, "count"),
    ("Dashboard-6", "Analytics", "% Reduction in Theft Events (monthly trend)", "pct", 5, 45, "%"),
    ("Dashboard-6", "Analytics", "Theft / Load diversion", "count", 2, 55, "count"),
    ("Dashboard-6", "Analytics", "Areas with Highest Theft Risk", "count", 3, 28, "count"),
    ("Dashboard-6", "Finance", "Revenue Recovery Improvement (%)", "pct", 3, 28, "%"),
    ("Dashboard-7", "Analytics", "Communication health issues", "count", 5, 95, "count"),
    ("Dashboard-7", "Technical", "Signal strength statistics", "pct", 72, 98, "%"),
    ("Dashboard-7", "Technical", "Packet loss percentage", "pct", 0.2, 8, "%"),
    ("Dashboard-7", "Technical", "Communication retry counts", "count", 50, 850, "count"),
    ("Dashboard-7", "Technical", "Non-reporting meters (>24 hours)", "count", 12, 220, "count"),
    ("Dashboard-7", "Technical", "Communication technology performance (RF/GPRS/PLC)", "pct", 85, 99, "%"),
    ("Dashboard-7", "Technical", "Weak Signal Percentage", "pct", 2, 18, "%"),
    ("Dashboard-8", "Advanced Analytics", "Auto-indexing consumers and DTRs for correct mapping", "count", 1200, 45000, "count"),
    ("Dashboard-8", "Advanced Analytics", "Track updated tag of DTs to Feeders", "count", 85, 1200, "count"),
    ("Dashboard-8", "Advanced Analytics", "Track updated tag of consumers to DTs", "count", 250, 8500, "count"),
    ("Dashboard-8", "Advanced Analytics", "Re-index consumer/DTR data for correct past-period T&D loss", "count", 500, 12000, "count"),
    ("Dashboard-8", "Advanced Analytics", "Mapping Accuracy (95%)", "pct", 88, 98, "%"),
    ("Dashboard-8", "Advanced Analytics", "DT-to-meter mapping accuracy", "pct", 90, 99, "%"),
    ("Dashboard-8", "Advanced Analytics", "% meters pending field verification (<5%)", "pct", 0.8, 6, "%"),
    ("Dashboard-8", "Advanced Analytics", "Confidence scoring (High/Medium/Low)", "pct", 75, 95, "%"),
    ("Dashboard-8", "Advanced Analytics", "Total assets tracked (Meters/Feeders/DTs)", "count", 5000, 85000, "count"),
    ("Dashboard-8", "Advanced Analytics", "Overloaded DTs identified and monitored", "count", 15, 180, "count"),
    ("Dashboard-8", "Advanced Analytics", "Mismatch analysis (Feeder→DT, DT→Meter)", "count", 20, 450, "count"),
    ("Dashboard-8", "Advanced Analytics", "Correctly mapped meters (%)", "pct", 88, 99, "%"),
    ("Dashboard-8", "Advanced Analytics", "Incorrectly mapped meters requiring correction (%)", "pct", 0.5, 8, "%"),
    ("Dashboard-8", "Advanced Analytics", "Verification pending count", "count", 50, 1200, "count"),
    ("Dashboard-8", "Advanced Analytics", "Correction cycle time (avg days)", "count", 2, 18, "days"),
    ("Dashboard-8", "Advanced Analytics", "Transformer utilization rate (% of rated capacity)", "pct", 45, 88, "%"),
    ("Dashboard-8", "Advanced Analytics", "Field verification completion rate", "pct", 82, 99, "%"),
    ("Dashboard-9", "Analytics", "Tamper sequence detection", "count", 5, 75, "count"),
    ("Dashboard-9", "Analytics", "Voltage/Current imbalance", "count", 8, 95, "count"),
    ("Dashboard-9", "Analytics", "Power factor deterioration", "count", 3, 42, "count"),
    ("Dashboard-9", "Analytics", "Overload / MD breach risk", "count", 12, 88, "count"),
    ("Dashboard-9", "Analytics", "Hidden outage pockets", "count", 2, 35, "count"),
    ("Dashboard-9", "Analytics", "Data quality issues", "count", 15, 120, "count"),
    ("Dashboard-9", "Analytics", "Reverse flow", "count", 0, 28, "count"),
    ("Dashboard-9", "Analytics", "Consumption spikes/drops", "count", 20, 180, "count"),
    ("Dashboard-9", "Analytics", "Phase-level mapping accuracy", "pct", 82, 98, "%"),
    ("Dashboard-9", "Analytics", "Phase imbalance reduced by minimum 30%", "pct", 28, 55, "%"),
    ("Dashboard-9", "Analytics", "Real-time phase load monitoring per transformer", "pct", 85, 99, "%"),
    ("Dashboard-9", "Analytics", "Imbalance alerts when threshold exceeded", "count", 10, 95, "count"),
    ("Dashboard-9", "Analytics", "Phase transfer recommendations (what-if)", "count", 5, 65, "count"),
]

MONTHS = ["Aug", "Sep", "Oct", "Nov", "Dec", "Jan"]

def main():
    random.seed(42)
    dashboards = {}
    for d in range(1, 10):
        key = f"Dashboard-{d}"
        dashboards[key] = { "title": f"Dashboard-{d}", "departments": [], "kpis": [], "charts": {} }

    for spec in KPI_SPECS:
        dashboard, dept, name, vtype, lo, hi, unit = spec
        value = generate_value(vtype, lo, hi, unit)
        if dept not in dashboards[dashboard]["departments"]:
            dashboards[dashboard]["departments"].append(dept)
        kpi = { "name": name, "department": dept, "value": value, "unit": unit }
        # Add trend (last 6 months) for chart-friendly KPIs
        if vtype in ("pct", "minutes", "index", "score") and "trend" not in kpi:
            trend = []
            for _ in MONTHS:
                trend.append(generate_value(vtype, max(lo, value * 0.7), min(hi, value * 1.3), unit))
            trend[-1] = value  # current month = value
            kpi["trend"] = trend
        dashboards[dashboard]["kpis"].append(kpi)

    # Chart-specific datasets
    for dkey, dval in dashboards.items():
        kpis = dval["kpis"]
        labels = MONTHS
        if dkey == "Dashboard-1":
            loss_kpis = [k for k in kpis if "Loss" in k["name"] and k["unit"] == "%"]
            dval["charts"]["lossTrend"] = {
                "labels": labels,
                "datasets": [{"name": k["name"].replace(" (%)", ""), "data": k.get("trend", [k["value"]]*6)} for k in loss_kpis[:4]]
            }
            eff_kpis = [k for k in kpis if "Efficiency" in k["name"] or "AT&C" in k["name"]]
            dval["charts"]["efficiency"] = [{"name": k["name"], "value": k["value"], "unit": k["unit"]} for k in eff_kpis]
        if dkey == "Dashboard-2":
            saidi = next((k for k in kpis if k["name"] == "SAIDI"), None)
            saifi = next((k for k in kpis if k["name"] == "SAIFI"), None)
            if saidi and saifi:
                dval["charts"]["reliabilityTrend"] = {
                    "labels": labels,
                    "SAIDI": saidi.get("trend", [saidi["value"]]*6),
                    "SAIFI": saifi.get("trend", [saifi["value"]]*6)
                }
            outage_kpis = [k for k in kpis if "Outage" in k["name"] or "MTTR" in k["name"] or "MTTI" in k["name"]][:6]
            dval["charts"]["outageMetrics"] = [{"name": k["name"], "value": k["value"], "unit": k["unit"]} for k in outage_kpis]
        if dkey == "Dashboard-3":
            loading = [k for k in kpis if "Loading" in k["name"] or "Load" in k["name"]][:5]
            dval["charts"]["loadingBands"] = [{"name": k["name"], "value": k["value"], "unit": k["unit"]} for k in loading]
        if dkey == "Dashboard-4":
            voltage = [k for k in kpis if "Voltage" in k["name"] or "Power Factor" in k["name"] or "Unbalance" in k["name"]][:8]
            dval["charts"]["voltageQuality"] = [{"name": k["name"], "value": k["value"], "unit": k["unit"]} for k in voltage]
        if dkey == "Dashboard-5":
            tamper = [k for k in kpis if "Tamper" in k["name"]]
            dval["charts"]["tamperByType"] = [{"name": k["name"].replace("Number of Tamper Alerts (", "").replace(")", ""), "value": k["value"]} for k in tamper]
        if dkey == "Dashboard-6":
            theft_rev = [k for k in kpis if "Theft" in k["name"] or "Revenue" in k["name"] or "Risk" in k["name"]]
            dval["charts"]["theftRevenue"] = [{"name": k["name"], "value": k["value"], "unit": k["unit"]} for k in theft_rev]
        if dkey == "Dashboard-7":
            comm = [k for k in kpis if "Signal" in k["name"] or "Packet" in k["name"] or "retry" in k["name"] or "Non-reporting" in k["name"]]
            dval["charts"]["communication"] = [{"name": k["name"], "value": k["value"], "unit": k["unit"]} for k in comm]
        if dkey == "Dashboard-8":
            mapping = [k for k in kpis if "Mapping" in k["name"] or "mapped" in k["name"] or "Accuracy" in k["name"]]
            dval["charts"]["mappingAccuracy"] = [{"name": k["name"], "value": k["value"], "unit": k["unit"]} for k in mapping[:8]]
        if dkey == "Dashboard-9":
            anomaly_phase = kpis[:10]
            dval["charts"]["anomalyPhase"] = [{"name": k["name"], "value": k["value"], "unit": k["unit"]} for k in anomaly_phase]

    out_path = Path(__file__).resolve().parent / "dashboards.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(dashboards, f, indent=2)
    print(f"Exported: {out_path}")

if __name__ == "__main__":
    main()
