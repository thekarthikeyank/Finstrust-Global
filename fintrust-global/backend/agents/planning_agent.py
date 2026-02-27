"""
Agent 3 — Planning Agent
Maps scraped financial data to model assumptions
Generates Bull / Base / Bear scenarios
Identifies missing data and asks user
"""

import asyncio
from typing import Optional


class PlanningAgent:

    def __init__(self, session: dict):
        self.session = session

    def _log(self, msg: str, status: str = "info"):
        self.session["logs"].append({
            "agent": "Planning Agent",
            "message": msg,
            "status": status,
            "timestamp": __import__("datetime").datetime.now().isoformat()
        })

    async def plan(self) -> dict:
        """Map company data to model assumptions"""
        data = self.session.get("company_data", {})
        model_type = self.session.get("model_recommendation", "DCF")
        company_name = data.get("company_name", "Company")

        self._log(f"Mapping {company_name} data to {model_type} assumptions...", "thinking")
        await asyncio.sleep(0.3)

        # Calculate core metrics from historical data
        metrics = self._calculate_metrics(data)
        self._log(f"Revenue CAGR: {metrics['rev_cagr']:.1%}", "info")
        self._log(f"Avg EBITDA Margin: {metrics['avg_ebitda_margin']:.1%}", "info")

        # Build assumptions based on model type
        if model_type == "DCF":
            assumptions = self._build_dcf_assumptions(data, metrics)
        elif model_type == "LBO":
            assumptions = self._build_lbo_assumptions(data, metrics)
        elif model_type == "3-Statement":
            assumptions = self._build_3stmt_assumptions(data, metrics)
        else:
            assumptions = self._build_fpa_assumptions(data, metrics)

        # Generate scenarios
        self._log("Generating Bull / Base / Bear scenarios...", "thinking")
        await asyncio.sleep(0.2)
        scenarios = self._generate_scenarios(assumptions, metrics, model_type)
        assumptions["scenarios"] = scenarios
        self._log("Scenarios generated: Bull +30% / Base / Bear -25%", "success")

        # Narrator notes
        self._log("Writing assumptions narrative...", "thinking")
        from agents.analyst_agent import AnalystAgent
        analyst = AnalystAgent(self.session)
        narrator_notes = await analyst.generate_narrator_notes(assumptions, data)
        self._log(f"Narrative written for {len(narrator_notes)} assumptions", "success")

        # Check missing fields
        missing = self._identify_missing(assumptions, data)

        self._log(f"Planning complete — {len(assumptions)} assumptions mapped", "success")

        return {
            "assumptions": assumptions,
            "narrator_notes": narrator_notes,
            "missing_fields": missing,
            "scenarios": scenarios,
        }

    def _calculate_metrics(self, data: dict) -> dict:
        """Calculate key metrics from historical data"""
        revenue = [r for r in data.get("revenue_history", []) if r and r > 0]
        ebitda = [e for e in data.get("ebitda_history", []) if e and e != 0]

        # Revenue CAGR
        rev_cagr = 0.08  # default
        if len(revenue) >= 2:
            try:
                rev_cagr = (revenue[0] / revenue[-1]) ** (1 / max(len(revenue) - 1, 1)) - 1
                rev_cagr = max(min(rev_cagr, 0.40), -0.10)  # cap between -10% and 40%
            except:
                rev_cagr = 0.08

        # Average EBITDA margin
        avg_ebitda_margin = 0.22  # default
        if ebitda and revenue:
            try:
                margins = [e/r for e, r in zip(ebitda[:len(revenue)], revenue) if r > 0]
                avg_ebitda_margin = sum(margins) / len(margins)
                avg_ebitda_margin = max(min(avg_ebitda_margin, 0.60), 0.05)
            except:
                avg_ebitda_margin = 0.22

        # Base revenue (most recent year)
        base_revenue = revenue[0] if revenue else 100.0

        # D&A estimation (typically 4-6% of revenue for tech, 6-10% for manufacturing)
        da_pct = 0.05

        # Capex estimation
        capex_pct = 0.06

        return {
            "rev_cagr": rev_cagr,
            "avg_ebitda_margin": avg_ebitda_margin,
            "base_revenue": base_revenue,
            "da_pct": da_pct,
            "capex_pct": capex_pct,
            "revenue_history": revenue,
            "ebitda_history": ebitda,
        }

    def _build_dcf_assumptions(self, data: dict, metrics: dict) -> dict:
        """Build DCF model assumptions from real data"""
        cagr = metrics["rev_cagr"]
        is_indian = data.get("is_indian", False)

        # Decelerate growth over projection period
        growth_rates = [
            round(cagr, 3),
            round(cagr * 0.95, 3),
            round(cagr * 0.90, 3),
            round(cagr * 0.85, 3),
            round(cagr * 0.80, 3),
        ]

        # Risk-free rate based on market
        rfr = 0.070 if is_indian else 0.045
        erp = 0.055 if not is_indian else 0.065
        beta = float(data.get("beta") or 1.10)

        # Cost of debt
        cod = 0.090 if is_indian else 0.065

        # Terminal growth
        tg = 0.055 if is_indian else 0.025

        # Net debt
        net_debt = float(data.get("total_debt") or 0) - float(data.get("cash") or 0)

        self._log(f"WACC inputs: RFR={rfr:.1%}, ERP={erp:.1%}, Beta={beta:.2f}", "info")
        self._log(f"Growth rates Y1-Y5: {[f'{g:.0%}' for g in growth_rates]}", "info")

        return {
            "company_name": data.get("company_name", "Company"),
            "base_year": 2024,
            "projection_years": 5,
            "currency": "INR" if is_indian else "USD",
            "base_revenue": round(metrics["base_revenue"], 1),
            "rev_growth_y1": growth_rates[0],
            "rev_growth_y2": growth_rates[1],
            "rev_growth_y3": growth_rates[2],
            "rev_growth_y4": growth_rates[3],
            "rev_growth_y5": growth_rates[4],
            "gross_margin": round(min(metrics["avg_ebitda_margin"] + 0.30, 0.80), 2),
            "ebitda_margin": round(metrics["avg_ebitda_margin"], 2),
            "da_pct": metrics["da_pct"],
            "tax_rate": 0.25 if not is_indian else 0.25,
            "capex_pct": metrics["capex_pct"],
            "nwc_pct": 0.02,
            "risk_free_rate": rfr,
            "erp": erp,
            "beta": beta,
            "cost_of_debt": cod,
            "debt_weight": 0.30,
            "equity_weight": 0.70,
            "terminal_growth": tg,
            "exit_multiple": round(metrics["avg_ebitda_margin"] * 60 + 6, 1),
            "shares_out": round(float(data.get("shares_outstanding") or 100), 1),
            "net_debt": round(max(net_debt, 0), 1),
            "model_date": __import__("datetime").datetime.now().strftime("%B %Y"),
            "analyst_name": "Financial Analyst",
            "peers": data.get("peers", []),
        }

    def _build_lbo_assumptions(self, data: dict, metrics: dict) -> dict:
        """Build LBO model assumptions"""
        is_indian = data.get("is_indian", False)
        base_ebitda = metrics["ebitda_history"][0] if metrics.get("ebitda_history") else metrics["base_revenue"] * metrics["avg_ebitda_margin"]

        return {
            "company_name": data.get("company_name", "Company"),
            "base_year": 2024,
            "currency": "INR" if is_indian else "USD",
            "entry_ebitda": round(base_ebitda, 1),
            "entry_multiple": 8.5,
            "txn_fees_pct": 0.02,
            "hold_period": 5,
            "senior_leverage": 4.0,
            "senior_rate": 0.090 if is_indian else 0.070,
            "sub_leverage": 1.5,
            "sub_rate": 0.12 if is_indian else 0.10,
            "amort_pct": 0.05,
            "cash_sweep_pct": 0.50,
            "base_revenue": round(metrics["base_revenue"], 1),
            "rev_growth": round(metrics["rev_cagr"], 3),
            "ebitda_margin": round(metrics["avg_ebitda_margin"], 2),
            "da_pct": metrics["da_pct"],
            "capex_pct": metrics["capex_pct"],
            "tax_rate": 0.25,
            "nwc_pct": 0.02,
            "exit_multiple": 9.5,
            "model_date": __import__("datetime").datetime.now().strftime("%B %Y"),
            "analyst_name": "Financial Analyst",
            "peers": data.get("peers", []),
        }

    def _build_3stmt_assumptions(self, data: dict, metrics: dict) -> dict:
        """Build 3-Statement model assumptions"""
        is_indian = data.get("is_indian", False)
        return {
            "company_name": data.get("company_name", "Company"),
            "base_year": 2024,
            "projection_years": 5,
            "currency": "INR" if is_indian else "USD",
            "base_revenue": round(metrics["base_revenue"], 1),
            "rev_growth": round(metrics["rev_cagr"], 3),
            "gross_margin": round(min(metrics["avg_ebitda_margin"] + 0.30, 0.80), 2),
            "ebitda_margin": round(metrics["avg_ebitda_margin"], 2),
            "da_pct": metrics["da_pct"],
            "interest_rate": 0.08 if is_indian else 0.06,
            "tax_rate": 0.25,
            "dividend_pct": 0.20,
            "capex_pct": metrics["capex_pct"],
            "dso": 45,
            "dio": 60,
            "dpo": 30,
            "open_cash": round(float(data.get("cash") or 20), 1),
            "open_debt": round(float(data.get("total_debt") or 50), 1),
            "open_ppe": round(metrics["base_revenue"] * 0.40, 1),
            "open_equity": round(metrics["base_revenue"] * 0.60, 1),
            "model_date": __import__("datetime").datetime.now().strftime("%B %Y"),
            "analyst_name": "Financial Analyst",
            "peers": data.get("peers", []),
        }

    def _build_fpa_assumptions(self, data: dict, metrics: dict) -> dict:
        """Build FP&A model assumptions"""
        return {
            "company_name": data.get("company_name", "Company"),
            "base_year": 2024,
            "currency": "USD",
            "total_revenue_budget": round(metrics["base_revenue"] * 1.10, 1),
            "rev_growth": round(metrics["rev_cagr"], 3),
            "q1_weight": 0.22, "q2_weight": 0.25,
            "q3_weight": 0.24, "q4_weight": 0.29,
            "gross_margin": round(min(metrics["avg_ebitda_margin"] + 0.30, 0.80), 2),
            "sm_pct": 0.15, "rd_pct": 0.12, "ga_pct": 0.08,
            "da_amount": round(metrics["base_revenue"] * 0.05, 1),
            "hc_budget": 150, "hc_cost_k": 120, "hc_growth": 0.08,
            "actual_rev_ytd": round(metrics["base_revenue"] * 0.48, 1),
            "budget_rev_ytd": round(metrics["base_revenue"] * 0.50, 1),
            "actual_ebitda_ytd": round(metrics["base_revenue"] * metrics["avg_ebitda_margin"] * 0.48, 1),
            "budget_ebitda_ytd": round(metrics["base_revenue"] * metrics["avg_ebitda_margin"] * 0.50, 1),
            "model_date": __import__("datetime").datetime.now().strftime("%B %Y"),
            "peers": data.get("peers", []),
        }

    def _generate_scenarios(self, base_assumptions: dict, metrics: dict, model_type: str) -> dict:
        """Generate Bull / Base / Bear scenario assumptions"""
        base_growth = metrics["rev_cagr"]
        base_margin = metrics["avg_ebitda_margin"]

        bull = dict(base_assumptions)
        bear = dict(base_assumptions)

        # Bull case — 30% better growth, 200bps margin expansion
        bull_growth = min(base_growth * 1.30, 0.40)
        for i in range(1, 6):
            key = f"rev_growth_y{i}"
            if key in bull:
                bull[key] = round(base_assumptions[key] * 1.30, 3)
        if "ebitda_margin" in bull:
            bull["ebitda_margin"] = round(min(base_margin + 0.02, 0.60), 2)
        if "terminal_growth" in bull:
            bull["terminal_growth"] = round(base_assumptions.get("terminal_growth", 0.025) + 0.005, 3)

        # Bear case — 25% worse growth, 200bps margin compression
        for i in range(1, 6):
            key = f"rev_growth_y{i}"
            if key in bear:
                bear[key] = round(max(base_assumptions[key] * 0.75, 0.01), 3)
        if "ebitda_margin" in bear:
            bear["ebitda_margin"] = round(max(base_margin - 0.02, 0.05), 2)
        if "terminal_growth" in bear:
            bear["terminal_growth"] = round(max(base_assumptions.get("terminal_growth", 0.025) - 0.005, 0.01), 3)

        return {
            "bull": bull,
            "base": base_assumptions,
            "bear": bear,
            "descriptions": {
                "bull": f"Accelerated growth (+30% vs base), margin expansion — optimistic scenario",
                "base": f"Historical trend continuation — most likely outcome",
                "bear": f"Growth deceleration (-25% vs base), margin pressure — downside scenario",
            }
        }

    def _identify_missing(self, assumptions: dict, data: dict) -> list:
        """Identify any remaining missing fields"""
        missing = []
        critical = {
            "beta": "Beta (Market Sensitivity Factor, e.g. 1.1)",
            "shares_out": "Shares Outstanding (in Millions)",
            "net_debt": "Net Debt = Total Debt minus Cash ($ Millions)",
        }
        for field, label in critical.items():
            val = assumptions.get(field)
            if val is None or val == 0:
                missing.append({"field": field, "label": label, "reason": "Not available from data source"})
                self._log(f"Missing: {label}", "warning")
        return missing
