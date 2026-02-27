"""
Agent 1 — Research Agent
Auto-detects Indian vs Global company
Scrapes Screener.in or Yahoo Finance
Extracts 5-year financial data
"""

import re
import asyncio
from typing import Optional


INDIAN_KEYWORDS = [
    "ltd", "limited", "industries", "enterprises", "infosys", "tcs", "wipro",
    "reliance", "hdfc", "icici", "bajaj", "mahindra", "tatamotors", "hcl",
    "itc", "bharti", "airtel", "kotak", "axis", "sbi", "ongc", "ntpc",
    "powergrid", "adani", "vedanta", "hindalco", "jswsteel", "tatasteel"
]

PEER_MAP = {
    "infosys":    ["TCS.NS", "WIPRO.NS", "HCLTECH.NS", "TECHM.NS", "MPHASIS.NS"],
    "tcs":        ["INFY.NS", "WIPRO.NS", "HCLTECH.NS", "TECHM.NS", "MPHASIS.NS"],
    "wipro":      ["INFY.NS", "TCS.NS", "HCLTECH.NS", "TECHM.NS", "LTIM.NS"],
    "reliance":   ["ONGC.NS", "IOC.NS", "BPCL.NS", "HPCL.NS", "GAIL.NS"],
    "hdfc":       ["ICICIBANK.NS", "SBIN.NS", "KOTAKBANK.NS", "AXISBANK.NS", "INDUSINDBK.NS"],
    "apple":      ["MSFT", "GOOGL", "META", "AMZN", "NVDA"],
    "microsoft":  ["AAPL", "GOOGL", "META", "AMZN", "ORCL"],
    "google":     ["AAPL", "MSFT", "META", "AMZN", "NFLX"],
    "amazon":     ["AAPL", "MSFT", "GOOGL", "META", "WALMART"],
    "tesla":      ["GM", "F", "RIVN", "NIO", "STLA"],
}

class ResearchAgent:

    def __init__(self, session: dict):
        self.session = session

    def _log(self, msg: str, status: str = "info"):
        entry = {
            "agent": "Research Agent",
            "message": msg,
            "status": status,
            "timestamp": __import__("datetime").datetime.now().isoformat()
        }
        self.session["logs"].append(entry)

    def _is_indian_company(self, name: str) -> bool:
        name_lower = name.lower().replace(" ", "")
        for kw in INDIAN_KEYWORDS:
            if kw in name_lower:
                return True
        return False

    async def research(self, company_name: str) -> dict:
        self._log(f"Analyzing company origin for '{company_name}'...", "thinking")
        await asyncio.sleep(0.3)

        is_indian = self._is_indian_company(company_name)

        if is_indian:
            self._log("Indian company detected — trying Screener.in first", "info")
            data = await self._scrape_screener(company_name)
            if not data.get("found"):
                self._log("Screener.in failed — falling back to Yahoo Finance", "warning")
                data = await self._fetch_yahoo(company_name + ".NS")
        else:
            self._log("Global company detected — fetching from Yahoo Finance", "info")
            data = await self._fetch_yahoo(company_name)

        # Fetch peers
        self._log("Fetching comparable companies data...", "thinking")
        peers = await self._fetch_peers(company_name, is_indian)
        data["peers"] = peers
        data["is_indian"] = is_indian

        return data

    async def _scrape_screener(self, company_name: str) -> dict:
        """Scrape Screener.in for Indian company data"""
        try:
            import requests
            from bs4 import BeautifulSoup

            # Search for company
            search_url = f"https://www.screener.in/api/company/search/?q={company_name.replace(' ', '+')}&v=3&fts=1"
            headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

            self._log(f"Searching Screener.in for '{company_name}'...", "thinking")
            resp = requests.get(search_url, headers=headers, timeout=10)

            if resp.status_code != 200:
                return {"found": False}

            results = resp.json()
            if not results:
                return {"found": False}

            # Get first result
            company_url = results[0].get("url", "")
            slug = company_url.strip("/").split("/")[-1]

            self._log(f"Found on Screener.in: {results[0].get('name', company_name)}", "success")

            # Fetch company page
            page_url = f"https://www.screener.in/company/{slug}/consolidated/"
            page_resp = requests.get(page_url, headers=headers, timeout=15)
            soup = BeautifulSoup(page_resp.content, "lxml")

            data = self._parse_screener_page(soup, results[0].get("name", company_name))
            data["found"] = True
            data["source"] = "Screener.in"
            data["company_name"] = results[0].get("name", company_name)
            return data

        except Exception as e:
            self._log(f"Screener.in error: {str(e)[:60]}", "warning")
            return {"found": False}

    def _parse_screener_page(self, soup, company_name: str) -> dict:
        """Parse Screener.in HTML page"""
        data = {"company_name": company_name}
        missing = []

        try:
            # Market cap
            market_cap_el = soup.find("li", {"class": re.compile(".*market-cap.*", re.I)})
            if market_cap_el:
                data["market_cap"] = self._clean_number(market_cap_el.get_text())

            # Current price
            price_el = soup.find("span", {"id": "current-price"}) or \
                       soup.find("span", {"class": re.compile(".*price.*", re.I)})
            if price_el:
                data["current_price"] = self._clean_number(price_el.get_text())

            # Financial tables
            tables = soup.find_all("table")
            for table in tables:
                headers = [th.get_text(strip=True) for th in table.find_all("th")]
                if not headers:
                    continue

                rows = table.find_all("tr")
                for row in rows:
                    cells = [td.get_text(strip=True) for td in row.find_all("td")]
                    if not cells:
                        continue

                    label = cells[0].lower() if cells else ""
                    values = cells[1:] if len(cells) > 1 else []

                    if "revenue" in label or "sales" in label:
                        data["revenue_history"] = [self._clean_number(v) for v in values[-5:]]
                        self._log(f"Revenue (5yr): {data['revenue_history']}", "success")
                    elif "ebitda" in label:
                        data["ebitda_history"] = [self._clean_number(v) for v in values[-5:]]
                        self._log(f"EBITDA (5yr): {data['ebitda_history']}", "success")
                    elif "net profit" in label or "net income" in label:
                        data["net_income_history"] = [self._clean_number(v) for v in values[-5:]]
                        self._log(f"Net Income (5yr): {data['net_income_history']}", "success")
                    elif "debt" in label and "net" not in label:
                        data["total_debt"] = self._clean_number(values[-1]) if values else None
                    elif "cash" in label:
                        data["cash"] = self._clean_number(values[-1]) if values else None
                    elif "eps" in label:
                        data["eps"] = self._clean_number(values[-1]) if values else None
                    elif "dividend" in label:
                        data["dividend"] = self._clean_number(values[-1]) if values else None

            # Check for missing critical fields
            critical = ["revenue_history", "ebitda_history", "net_income_history"]
            for field in critical:
                if field not in data or not data[field]:
                    missing.append({"field": field, "label": field.replace("_", " ").title(),
                                   "reason": "Not found on Screener.in"})
                    self._log(f"{field} not found", "warning")

            # Beta always from Yahoo for Indian stocks
            missing.append({"field": "beta", "label": "Beta (Market Sensitivity)",
                           "reason": "Please provide or use 1.0 as default"})

        except Exception as e:
            self._log(f"Parse error: {str(e)[:60]}", "warning")

        data["missing_fields"] = missing
        return data

    async def _fetch_yahoo(self, ticker_or_name: str) -> dict:
        """Fetch data from Yahoo Finance via yfinance"""
        try:
            import yfinance as yf

            # Try to resolve ticker
            ticker_str = self._resolve_ticker(ticker_or_name)
            self._log(f"Fetching Yahoo Finance data for ticker: {ticker_str}", "thinking")

            stock = yf.Ticker(ticker_str)
            info = stock.info
            financials = stock.financials
            balance_sheet = stock.balance_sheet
            cash_flow = stock.cashflow

            if not info or info.get("regularMarketPrice") is None:
                self._log(f"Ticker {ticker_str} not found — trying alternate", "warning")
                return {"found": False}

            company_name = info.get("longName", ticker_or_name)
            self._log(f"Found: {company_name}", "success")

            data = {
                "found": True,
                "source": "Yahoo Finance",
                "company_name": company_name,
                "ticker": ticker_str,
                "sector": info.get("sector", "Unknown"),
                "industry": info.get("industry", "Unknown"),
                "current_price": info.get("regularMarketPrice", 0),
                "market_cap": info.get("marketCap", 0) / 1e6,  # Convert to millions
                "shares_outstanding": info.get("sharesOutstanding", 0) / 1e6,
                "beta": info.get("beta", None),
                "total_debt": info.get("totalDebt", 0) / 1e6 if info.get("totalDebt") else 0,
                "cash": info.get("totalCash", 0) / 1e6 if info.get("totalCash") else 0,
                "currency": info.get("currency", "USD"),
            }

            # Extract historical financials (last 5 years)
            missing = []

            if financials is not None and not financials.empty:
                rev_row = None
                for idx in financials.index:
                    if "revenue" in str(idx).lower() or "total revenue" in str(idx).lower():
                        rev_row = financials.loc[idx]
                        break
                if rev_row is not None:
                    data["revenue_history"] = [round(v/1e6, 1) for v in rev_row.values[:5] if v and v == v]
                    self._log(f"Revenue history: {data['revenue_history']}", "success")

                ebitda_row = None
                for idx in financials.index:
                    if "ebitda" in str(idx).lower():
                        ebitda_row = financials.loc[idx]
                        break
                if ebitda_row is not None:
                    data["ebitda_history"] = [round(v/1e6, 1) for v in ebitda_row.values[:5] if v and v == v]
                    self._log(f"EBITDA history: {data['ebitda_history']}", "success")

                ni_row = None
                for idx in financials.index:
                    if "net income" in str(idx).lower():
                        ni_row = financials.loc[idx]
                        break
                if ni_row is not None:
                    data["net_income_history"] = [round(v/1e6, 1) for v in ni_row.values[:5] if v and v == v]
                    self._log(f"Net Income history: {data['net_income_history']}", "success")

            # Check missing
            if not data.get("revenue_history"):
                missing.append({"field": "base_revenue", "label": "Base Year Revenue ($M)",
                               "reason": "Could not extract from Yahoo Finance"})
            if not data.get("beta"):
                missing.append({"field": "beta", "label": "Beta (e.g. 1.1)",
                               "reason": "Beta not available — please provide"})
                self._log("Beta not found — will ask user", "warning")

            data["missing_fields"] = missing
            return data

        except Exception as e:
            self._log(f"Yahoo Finance error: {str(e)[:80]}", "error")
            return {
                "found": False,
                "missing_fields": [
                    {"field": "base_revenue", "label": "Base Year Revenue ($M)", "reason": "Auto-fetch failed"},
                    {"field": "ebitda_margin", "label": "EBITDA Margin %", "reason": "Auto-fetch failed"},
                    {"field": "beta", "label": "Beta", "reason": "Auto-fetch failed"},
                ]
            }

    def _resolve_ticker(self, name: str) -> str:
        """Convert company name to ticker symbol"""
        name_map = {
            "apple": "AAPL", "microsoft": "MSFT", "google": "GOOGL",
            "alphabet": "GOOGL", "amazon": "AMZN", "tesla": "TSLA",
            "meta": "META", "facebook": "META", "netflix": "NFLX",
            "nvidia": "NVDA", "infosys": "INFY.NS", "tcs": "TCS.NS",
            "wipro": "WIPRO.NS", "reliance": "RELIANCE.NS", "hdfc": "HDFCBANK.NS",
            "icici": "ICICIBANK.NS", "bajaj": "BAJFINANCE.NS",
        }
        name_lower = name.lower().replace(" ", "")
        for key, ticker in name_map.items():
            if key in name_lower:
                return ticker
        # If it looks like a ticker already
        if name.upper() == name or len(name) <= 5:
            return name.upper()
        return name.upper()

    async def _fetch_peers(self, company_name: str, is_indian: bool) -> list:
        """Fetch peer company data for COMPS sheet"""
        name_lower = company_name.lower().replace(" ", "")
        peers_tickers = []

        for key, tickers in PEER_MAP.items():
            if key in name_lower:
                peers_tickers = tickers
                break

        if not peers_tickers:
            # Default peers based on region
            if is_indian:
                peers_tickers = ["TCS.NS", "INFY.NS", "WIPRO.NS", "HCLTECH.NS", "TECHM.NS"]
            else:
                peers_tickers = ["AAPL", "MSFT", "GOOGL", "AMZN", "META"]

        peers_data = []
        try:
            import yfinance as yf
            for ticker in peers_tickers[:5]:
                try:
                    stock = yf.Ticker(ticker)
                    info = stock.info
                    if info:
                        peers_data.append({
                            "name": info.get("shortName", ticker),
                            "ticker": ticker,
                            "ev_ebitda": round(info.get("enterpriseToEbitda", 0), 1),
                            "pe_ratio": round(info.get("trailingPE", 0), 1),
                            "ev_revenue": round(info.get("enterpriseToRevenue", 0), 1),
                            "market_cap": round(info.get("marketCap", 0) / 1e6, 0),
                            "revenue_growth": round(info.get("revenueGrowth", 0) * 100, 1) if info.get("revenueGrowth") else 0,
                            "ebitda_margin": round(info.get("ebitdaMargins", 0) * 100, 1) if info.get("ebitdaMargins") else 0,
                        })
                        self._log(f"Peer fetched: {info.get('shortName', ticker)}", "info")
                except:
                    pass
        except:
            self._log("Peer fetch skipped — yfinance unavailable", "warning")

        return peers_data

    def _clean_number(self, text: str) -> Optional[float]:
        """Clean and convert number strings"""
        if not text:
            return None
        text = str(text).replace(",", "").replace("₹", "").replace("$", "").replace("%", "").strip()
        # Handle Cr (crores) and Lakh
        if "cr" in text.lower():
            text = text.lower().replace("cr", "").strip()
            try:
                return round(float(text) / 100, 2)  # Convert Cr to millions (approx)
            except:
                return None
        try:
            return float(text)
        except:
            return None
