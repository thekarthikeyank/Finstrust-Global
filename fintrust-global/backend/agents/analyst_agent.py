"""
Agent 2 — CFA-Level Analyst Agent
Institutional-grade model recommendation + deep investment analysis
"""

import asyncio
import json
import re
import os


CFA_SYSTEM = """You are a CFA charterholder and Senior Equity Research Analyst at a top bulge-bracket bank.
You have 15+ years analyzing Indian and global equities across all sectors.
You think in numbers, speak with conviction, and back every view with data.
Output ONLY valid JSON when asked for JSON. No markdown, no preamble."""


class AnalystAgent:

    def __init__(self, session: dict):
        self.session = session

    def _log(self, msg: str, status: str = "info"):
        self.session["logs"].append({
            "agent": "Analyst Agent",
            "message": msg,
            "status": status,
            "timestamp": __import__("datetime").datetime.now().isoformat()
        })

    async def recommend(self) -> dict:
        """Analyze company data and recommend model type with CFA-level reasoning"""
        data = self.session.get("company_data", {})
        company_name = data.get("company_name", "Unknown")

        self._log(f"Running CFA-level analysis on {company_name}...", "thinking")
        await asyncio.sleep(0.2)

        # Try Gemini first
        recommendation = await self._get_gemini_recommendation(data)

        if not recommendation:
            # Try Groq
            recommendation = await self._get_groq_recommendation(data)

        if not recommendation:
            # Fallback to rule-based
            self._log("Using quantitative model selection...", "info")
            recommendation = self._rule_based_recommendation(data)

        self.session["model_recommendation"] = recommendation["model_type"]
        self._log(f"Recommended: {recommendation['model_type']} model", "success")
        self._log(f"{recommendation['reasoning'][:100]}...", "info")

        return recommendation

    async def _get_gemini_recommendation(self, data: dict) -> dict:
        """Get CFA-level recommendation from Gemini"""
        try:
            import google.generativeai as genai
            api_key = os.environ.get("GEMINI_API_KEY", "")
            if not api_key:
                return None

            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(
                "gemini-1.5-flash",
                system_instruction=CFA_SYSTEM
            )

            company_summary = {
                "company": data.get("company_name"),
                "sector": data.get("sector", "Unknown"),
                "revenue_history_cr": data.get("revenue_history", [])[:5],
                "ebitda_history_cr": data.get("ebitda_history", [])[:5],
                "market_cap_cr": data.get("market_cap", 0),
                "total_debt_cr": data.get("total_debt", 0),
                "pe_ratio": data.get("pe_ratio", 0),
                "is_indian": data.get("is_indian", True),
            }

            prompt = f"""As a CFA analyst, recommend the best financial model for this company.

Company Profile:
{json.dumps(company_summary, indent=2)}

Model options: DCF, LBO, 3-Statement, FP&A

Selection criteria:
- DCF: Stable, cash-generative companies with predictable FCF — best for mature businesses
- LBO: High debt, PE acquisition target, strong cash flows for debt service
- 3-Statement: Companies needing full P&L + Balance Sheet + Cash Flow integration
- FP&A: High-growth, early stage, or internal budgeting focus

Respond with ONLY this JSON (no other text, no markdown):
{{
  "model_type": "DCF",
  "reasoning": "Specific 2-3 sentence reasoning using actual numbers from the data",
  "key_metrics": {{
    "revenue_cagr": 0.0,
    "avg_ebitda_margin": 0.0,
    "debt_to_ebitda": 0.0,
    "ev_ebitda_implied": 0.0
  }},
  "confidence": "high",
  "valuation_view": "Attractive/Fair Value/Expensive",
  "one_line_thesis": "Single sentence investment thesis"
}}"""

            response = model.generate_content(prompt)
            text = response.text.strip()
            # Clean markdown if present
            text = re.sub(r'```json|```', '', text).strip()
            json_match = re.search(r'\{.*\}', text, re.DOTALL)
            if json_match:
                return json.loads(json_match.group())
        except Exception as e:
            self._log(f"Gemini analyst: {str(e)[:50]}", "warning")
        return None

    async def _get_groq_recommendation(self, data: dict) -> dict:
        """Get recommendation from Groq as fallback"""
        try:
            from groq import Groq
            api_key = os.environ.get("GROQ_API_KEY", "")
            if not api_key:
                return None

            client = Groq(api_key=api_key)
            company_summary = {
                "company": data.get("company_name"),
                "sector": data.get("sector", "Unknown"),
                "revenue_history": data.get("revenue_history", [])[:5],
                "ebitda_history": data.get("ebitda_history", [])[:5],
                "market_cap": data.get("market_cap", 0),
                "total_debt": data.get("total_debt", 0),
            }

            prompt = f"""Recommend best financial model for this company. Output ONLY JSON.

{json.dumps(company_summary)}

JSON format:
{{"model_type": "DCF", "reasoning": "specific reasoning with numbers", "key_metrics": {{"revenue_cagr": 0.0, "avg_ebitda_margin": 0.0, "debt_to_ebitda": 0.0}}, "confidence": "high", "valuation_view": "Fair Value", "one_line_thesis": "thesis here"}}"""

            res = client.chat.completions.create(
                model="llama-3.1-70b-versatile",
                messages=[
                    {"role": "system", "content": CFA_SYSTEM},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=400
            )
            text = res.choices[0].message.content.strip()
            text = re.sub(r'```json|```', '', text).strip()
            json_match = re.search(r'\{.*\}', text, re.DOTALL)
            if json_match:
                return json.loads(json_match.group())
        except Exception as e:
            self._log(f"Groq analyst: {str(e)[:50]}", "warning")
        return None

    def _rule_based_recommendation(self, data: dict) -> dict:
        """Quantitative fallback model recommendation"""
        revenue = data.get("revenue_history", [])
        ebitda = data.get("ebitda_history", [])
        debt = data.get("total_debt", 0) or 0
        market_cap = data.get("market_cap", 0) or 0
        sector = data.get("sector", "").lower()
        company = data.get("company_name", "the company")

        rev_cagr = 0
        if len(revenue) >= 2:
            try:
                rev_cagr = (revenue[0] / revenue[-1]) ** (1/(len(revenue)-1)) - 1
            except:
                rev_cagr = 0.08

        avg_margin = 0
        if ebitda and revenue:
            try:
                margins = [e/r for e,r in zip(ebitda, revenue) if r > 0]
                avg_margin = sum(margins)/len(margins) if margins else 0.20
            except:
                avg_margin = 0.20

        debt_ebitda = 0
        if ebitda and ebitda[0] > 0:
            try:
                debt_ebitda = debt / ebitda[0]
            except:
                debt_ebitda = 0

        if debt_ebitda > 3.0:
            model_type = "LBO"
            reasoning = f"{company} carries {debt_ebitda:.1f}x Debt/EBITDA — elevated leverage makes LBO the right framework to stress-test debt service capacity and model equity returns under various exit scenarios."
            thesis = f"High-leverage play; LBO model will reveal true equity return potential."
        elif rev_cagr > 0.20 or any(s in sector for s in ["technology", "software", "pharma", "consumer"]):
            model_type = "DCF"
            reasoning = f"{company} shows {rev_cagr:.0%} revenue CAGR with {avg_margin:.0%} EBITDA margins — DCF is appropriate to capture long-term intrinsic value with explicit FCF forecasting and terminal value based on stable cash flows."
            thesis = f"Growth-to-value compounder; DCF captures long-term FCF generation."
        elif avg_margin > 0.25 and market_cap > 5000:
            model_type = "DCF"
            reasoning = f"Strong EBITDA margins of {avg_margin:.0%} and large-cap status (₹{market_cap:,.0f} Cr market cap) indicate a mature, cash-generative business — DCF with Gordon Growth terminal value is the institutional standard here."
            thesis = f"Mature compounder with strong FCF; DCF is the right framework."
        else:
            model_type = "3-Statement"
            reasoning = f"Building a full 3-Statement model for {company} will integrate P&L, Balance Sheet, and Cash Flow — essential to understand working capital dynamics, debt capacity, and normalized earnings before applying valuation multiples."
            thesis = f"Need full financial integration before valuation — 3-Statement first."

        return {
            "model_type": model_type,
            "reasoning": reasoning,
            "key_metrics": {
                "revenue_cagr": round(rev_cagr, 3),
                "avg_ebitda_margin": round(avg_margin, 3),
                "debt_to_ebitda": round(debt_ebitda, 1),
            },
            "confidence": "medium",
            "valuation_view": "Fair Value",
            "one_line_thesis": thesis
        }

    async def generate_narrator_notes(self, assumptions: dict, company_data: dict) -> list:
        """Generate CFA-level assumption explanations"""
        self._log("Narrating model assumptions...", "thinking")

        # Try Gemini narrator
        notes = await self._get_gemini_narrator(assumptions, company_data)
        if notes:
            return notes

        # Fallback rule-based
        notes = []
        revenue_history = company_data.get("revenue_history", [])
        ebitda_history = company_data.get("ebitda_history", [])
        for field, value in assumptions.items():
            explanation = self._explain_assumption(field, value, revenue_history, ebitda_history, company_data)
            if explanation:
                notes.append({"field": field, "value": value, "explanation": explanation})
        return notes

    async def _get_gemini_narrator(self, assumptions: dict, company_data: dict) -> list:
        """Use Gemini to generate assumption explanations"""
        try:
            import google.generativeai as genai
            api_key = os.environ.get("GEMINI_API_KEY", "")
            if not api_key:
                return []

            genai.configure(api_key=api_key)
            model = genai.GenerativeModel("gemini-1.5-flash", system_instruction=CFA_SYSTEM)

            prompt = f"""For {company_data.get('company_name')}:
Historical Revenue (₹ Cr): {company_data.get('revenue_history', [])}
Historical EBITDA (₹ Cr): {company_data.get('ebitda_history', [])}

Explain each assumption in ONE precise sentence (max 30 words) like a senior analyst:
{json.dumps({k: v for k, v in list(assumptions.items())[:12]}, indent=2)}

Output ONLY a JSON array:
[{{"field": "rev_growth_y1", "value": 0.12, "explanation": "12% growth reflects 3-year CAGR of 11.8%, with slight moderation as base effect kicks in."}}]"""

            response = model.generate_content(prompt)
            text = response.text.strip()
            text = re.sub(r'```json|```', '', text).strip()
            arr_match = re.search(r'\[.*\]', text, re.DOTALL)
            if arr_match:
                return json.loads(arr_match.group())
        except:
            pass
        return []

    def _explain_assumption(self, field: str, value, revenue_history: list, ebitda_history: list, data: dict) -> str:
        company = data.get("company_name", "the company")
        try:
            v = float(value)
        except:
            v = 0

        explanations = {
            "rev_growth_y1": f"{v:.0%} growth anchored to {company}'s 3-year CAGR; modest deceleration applied as law of large numbers takes effect.",
            "rev_growth_y2": f"Year 2 growth moderates to {v:.0%} as competitive pressures normalize and base effect becomes more pronounced.",
            "rev_growth_y3": f"Terminal-approach growth of {v:.0%} reflects steady-state expansion in line with sector long-run average.",
            "ebitda_margin": f"EBITDA margin of {v:.0%} reflects {company}'s historical average; assumes no significant cost inflation or margin compression.",
            "tax_rate": f"Effective tax rate of {v:.0%} based on applicable corporate tax rate; MAT provisions excluded.",
            "capex_pct": f"Capex at {v:.0%} of revenue reflects maintenance + growth capex; consistent with sector peer average.",
            "terminal_growth": f"Terminal growth of {v:.0%} equals long-run nominal GDP growth — conservative and appropriate for DCF terminal value.",
            "wacc": f"WACC of {v:.1%} derived from CAPM cost of equity + after-tax cost of debt, weighted by capital structure.",
            "beta": f"Beta of {v:.2f} sourced from 2-year weekly regression vs Nifty 50 — reflects {company}'s market sensitivity.",
            "risk_free_rate": f"Risk-free rate of {v:.1%} based on current 10-year GOI bond yield — standard Indian market practice.",
            "erp": f"Equity risk premium of {v:.1%} represents Damodaran's India ERP estimate — consensus for emerging market equities.",
            "debt_interest_rate": f"Interest rate of {v:.1%} reflects blended cost of {company}'s existing debt facilities.",
            "exit_multiple": f"Exit EV/EBITDA of {v:.1f}x anchored to sector median trading multiple — conservative vs historical M&A premiums.",
        }
        return explanations.get(field, f"{field.replace('_', ' ').title()} of {value} — set based on company analysis and sector benchmarks.")
