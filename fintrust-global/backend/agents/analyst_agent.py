"""
Agent 2 — Analyst Agent
Uses Ollama (Llama 3.1) to analyze company data
Recommends best model type with reasoning
Smart Assumptions Narrator
"""

import asyncio
import json
import re


ANALYST_SYSTEM_PROMPT = """You are a Senior Investment Banking Analyst at a top-tier firm.
Your job is to analyze a company's financial profile and recommend the most appropriate financial model.

Rules:
- Be precise and use correct financial terminology
- Base recommendations on the actual data provided
- Explain your reasoning clearly in 2-3 sentences
- Always output valid JSON

Model selection logic:
- DCF: Stable, cash-generative public companies with predictable growth
- LBO: Companies with strong cash flows, potential for leverage, PE acquisition targets
- 3-Statement: Companies needing full financial statement integration and forecasting
- FP&A: High-growth companies, startups, or internal budgeting use cases
"""

NARRATOR_PROMPT = """You are a financial modeling expert explaining assumptions to a client.
For each assumption, write ONE clear sentence explaining:
1. What the assumption is
2. Where the number came from (historical data, industry average, or estimate)
3. Why it is reasonable for this company

Be concise. Maximum 25 words per assumption. Use professional but accessible language.
Output as JSON array of {field, value, explanation} objects.
"""


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
        """Analyze company data and recommend model type"""
        data = self.session.get("company_data", {})
        company_name = data.get("company_name", "Unknown")

        self._log(f"Reading financial profile of {company_name}...", "thinking")
        await asyncio.sleep(0.2)

        # Try Ollama first
        recommendation = await self._get_ollama_recommendation(data)

        if not recommendation:
            # Fallback to rule-based recommendation
            self._log("Using rule-based model selection...", "info")
            recommendation = self._rule_based_recommendation(data)

        self.session["model_recommendation"] = recommendation["model_type"]

        self._log(f"Recommended model: {recommendation['model_type']}", "success")
        self._log(f"Reasoning: {recommendation['reasoning'][:80]}...", "info")

        return recommendation

    async def _get_ollama_recommendation(self, data: dict) -> dict:
        """Get recommendation from Ollama Llama 3.1"""
        try:
            import httpx

            company_summary = {
                "company": data.get("company_name"),
                "sector": data.get("sector", "Unknown"),
                "revenue_history": data.get("revenue_history", []),
                "ebitda_history": data.get("ebitda_history", []),
                "market_cap_millions": data.get("market_cap", 0),
                "total_debt_millions": data.get("total_debt", 0),
                "is_public": True,
                "is_indian": data.get("is_indian", False),
            }

            prompt = f"""Analyze this company and recommend the best financial model.

Company Data:
{json.dumps(company_summary, indent=2)}

Respond with ONLY this JSON (no other text):
{{
  "model_type": "DCF",
  "reasoning": "Two to three sentence explanation of why this model fits",
  "key_metrics": {{
    "revenue_cagr": 0.0,
    "avg_ebitda_margin": 0.0,
    "debt_to_ebitda": 0.0
  }},
  "confidence": "high"
}}"""

            async with httpx.AsyncClient(timeout=30) as client:
                resp = await client.post(
                    "http://localhost:11434/api/generate",
                    json={
                        "model": "llama3.1",
                        "prompt": prompt,
                        "system": ANALYST_SYSTEM_PROMPT,
                        "stream": False
                    }
                )
                if resp.status_code == 200:
                    result = resp.json()
                    text = result.get("response", "")
                    # Extract JSON from response
                    json_match = re.search(r'\{.*\}', text, re.DOTALL)
                    if json_match:
                        return json.loads(json_match.group())
        except Exception as e:
            self._log(f"Ollama unavailable: {str(e)[:50]} — using rule-based", "warning")

        return None

    def _rule_based_recommendation(self, data: dict) -> dict:
        """Fallback rule-based model recommendation"""
        revenue = data.get("revenue_history", [])
        ebitda = data.get("ebitda_history", [])
        debt = data.get("total_debt", 0) or 0
        market_cap = data.get("market_cap", 0) or 0
        sector = data.get("sector", "").lower()

        # Calculate metrics
        rev_cagr = 0
        if len(revenue) >= 2:
            try:
                rev_cagr = (revenue[0] / revenue[-1]) ** (1 / (len(revenue) - 1)) - 1
            except:
                rev_cagr = 0.08

        avg_ebitda_margin = 0
        if ebitda and revenue:
            try:
                margins = [e/r for e, r in zip(ebitda, revenue) if r and r > 0]
                avg_ebitda_margin = sum(margins) / len(margins) if margins else 0.20
            except:
                avg_ebitda_margin = 0.20

        debt_ebitda = 0
        if ebitda:
            try:
                debt_ebitda = debt / ebitda[0] if ebitda[0] > 0 else 0
            except:
                debt_ebitda = 0

        # Decision logic
        if debt_ebitda > 3.0:
            model_type = "LBO"
            reasoning = f"{data.get('company_name')} carries {debt_ebitda:.1f}x debt/EBITDA suggesting leveraged structure. LBO model will properly account for debt service and equity returns analysis."
        elif rev_cagr > 0.20 or "technology" in sector or "software" in sector:
            model_type = "DCF"
            reasoning = f"{data.get('company_name')} shows {rev_cagr:.0%} revenue CAGR with {avg_ebitda_margin:.0%} EBITDA margins — ideal for DCF valuation with terminal value based on stable cash flows."
        elif avg_ebitda_margin > 0.25 and market_cap > 1000:
            model_type = "DCF"
            reasoning = f"Strong EBITDA margins of {avg_ebitda_margin:.0%} and large market cap indicate cash-generative business suited for DCF with Gordon Growth terminal value."
        else:
            model_type = "3-Statement"
            reasoning = f"Comprehensive 3-Statement model recommended to fully understand {data.get('company_name')}'s financial dynamics, working capital cycles, and debt capacity before valuation."

        return {
            "model_type": model_type,
            "reasoning": reasoning,
            "key_metrics": {
                "revenue_cagr": round(rev_cagr, 3),
                "avg_ebitda_margin": round(avg_ebitda_margin, 3),
                "debt_to_ebitda": round(debt_ebitda, 1),
            },
            "confidence": "medium"
        }

    async def generate_narrator_notes(self, assumptions: dict, company_data: dict) -> list:
        """Generate Smart Assumptions Narrator explanations"""
        self._log("Generating assumptions narrative...", "thinking")

        notes = []
        revenue_history = company_data.get("revenue_history", [])
        ebitda_history = company_data.get("ebitda_history", [])

        # Try Ollama narrator
        ollama_notes = await self._get_ollama_narrator(assumptions, company_data)
        if ollama_notes:
            return ollama_notes

        # Fallback rule-based narrator
        for field, value in assumptions.items():
            explanation = self._explain_assumption(field, value, revenue_history, ebitda_history, company_data)
            if explanation:
                notes.append({"field": field, "value": value, "explanation": explanation})

        return notes

    async def _get_ollama_narrator(self, assumptions: dict, company_data: dict) -> list:
        """Use Ollama to generate narrator explanations"""
        try:
            import httpx

            prompt = f"""For this company: {company_data.get('company_name')}
Historical Revenue: {company_data.get('revenue_history', [])}
Historical EBITDA: {company_data.get('ebitda_history', [])}

Explain these assumptions in one sentence each (max 25 words):
{json.dumps({k: v for k, v in list(assumptions.items())[:10]}, indent=2)}

Output ONLY a JSON array like:
[{{"field": "rev_growth_y1", "value": 0.12, "explanation": "12% growth based on 3-year historical CAGR of 11.8%, reflecting continued market expansion."}}]"""

            async with httpx.AsyncClient(timeout=25) as client:
                resp = await client.post(
                    "http://localhost:11434/api/generate",
                    json={"model": "llama3.1", "prompt": prompt,
                          "system": NARRATOR_PROMPT, "stream": False}
                )
                if resp.status_code == 200:
                    text = resp.json().get("response", "")
                    arr_match = re.search(r'\[.*\]', text, re.DOTALL)
                    if arr_match:
                        return json.loads(arr_match.group())
        except:
            pass
        return []

    def _explain_assumption(self, field: str, value, revenue_history: list, ebitda_history: list, data: dict) -> str:
        """Rule-based assumption explanation"""
        company = data.get("company_name", "the company")

        explanations = {
            "rev_growth_y1": f"Year 1 growth of {float(value):.0%} derived from {company}'s 3-year historical revenue CAGR with slight deceleration applied.",
            "ebitda_margin": f"EBITDA margin of {float(value):.0%} based on average of last 3 years of reported margins for {company}.",
            "tax_rate": f"Tax rate of {float(value):.0%} reflects standard corporate tax rate applicable to {company}'s domicile.",
            "capex_pct": f"Capex at {float(value):.0%} of revenue based on industry average maintenance capex requirements.",
            "terminal_growth": f"Terminal growth of {float(value):.0%} reflects long-run GDP growth assumption — conservative for mature business.",
            "beta": f"Beta of {float(value):.2f} sourced from market data, reflecting {company}'s systematic risk vs market index.",
            "risk_free_rate": f"Risk-free rate of {float(value):.1%} based on current 10-year government bond yield.",
            "erp": f"Equity risk premium of {float(value):.1%} represents consensus market premium over risk-free rate.",
        }
        return explanations.get(field, f"{field.replace('_', ' ').title()} set to {value} based on company analysis and industry benchmarks.")
