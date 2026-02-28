"""
Agent 1 — Research Agent
- Extracts clean company name from user query
- Auto-detects Indian vs Global
- Scrapes Screener.in or Yahoo Finance
- 5-year financial history
"""

import re
import asyncio
import os
from typing import Optional


INDIAN_COMPANIES = {
    "infosys": "INFY.NS", "infy": "INFY.NS",
    "tcs": "TCS.NS", "tata consultancy": "TCS.NS",
    "wipro": "WIPRO.NS",
    "hcl": "HCLTECH.NS", "hcltech": "HCLTECH.NS",
    "tech mahindra": "TECHM.NS", "techm": "TECHM.NS",
    "reliance": "RELIANCE.NS", "ril": "RELIANCE.NS",
    "hdfc bank": "HDFCBANK.NS", "hdfc": "HDFCBANK.NS",
    "icici bank": "ICICIBANK.NS", "icici": "ICICIBANK.NS",
    "sbi": "SBIN.NS", "state bank": "SBIN.NS",
    "kotak": "KOTAKBANK.NS", "kotak mahindra": "KOTAKBANK.NS",
    "axis bank": "AXISBANK.NS", "axis": "AXISBANK.NS",
    "bajaj finance": "BAJFINANCE.NS", "bajaj": "BAJFINANCE.NS",
    "tata motors": "TATAMOTORS.NS", "tatamotors": "TATAMOTORS.NS",
    "tata steel": "TATASTEEL.NS",
    "adani": "ADANIENT.NS",
    "ongc": "ONGC.NS",
    "ntpc": "NTPC.NS",
    "itc": "ITC.NS",
    "airtel": "BHARTIARTL.NS", "bharti airtel": "BHARTIARTL.NS",
    "maruti": "MARUTI.NS", "maruti suzuki": "MARUTI.NS",
    "sun pharma": "SUNPHARMA.NS", "sun pharmaceutical": "SUNPHARMA.NS",
    "dr reddy": "DRREDDY.NS",
    "cipla": "CIPLA.NS",
    "asian paints": "ASIANPAINT.NS",
    "hindustan unilever": "HINDUNILVR.NS", "hul": "HINDUNILVR.NS",
    "nestle": "NESTLEIND.NS",
    "titan": "TITAN.NS",
    "ultratech": "ULTRACEMCO.NS", "ultratech cement": "ULTRACEMCO.NS",
    "l&t": "LT.NS", "larsen": "LT.NS",
    "jsw steel": "JSWSTEEL.NS",
    "hindalco": "HINDALCO.NS",
    "vedanta": "VEDL.NS",
    "power grid": "POWERGRID.NS",
    "coal india": "COALINDIA.NS",
    "divis": "DIVISLAB.NS", "divis labs": "DIVISLAB.NS",
    "eicher motors": "EICHERMOT.NS",
    "hero motocorp": "HEROMOTOCO.NS",
    "britannia": "BRITANNIA.NS",
    "grasim": "GRASIM.NS",
    "indusind bank": "INDUSINDBK.NS",
    "shree cement": "SHREECEM.NS",
    "apollo hospitals": "APOLLOHOSP.NS",
    "dmart": "DMART.NS", "avenue supermarts": "DMART.NS",
    "pidilite": "PIDILITIND.NS",
    "berger paints": "BERGEPAINT.NS",
    "havells": "HAVELLS.NS",
    "muthoot finance": "MUTHOOTFIN.NS",
    "page industries": "PAGEIND.NS",
    "info edge": "NAUKRI.NS", "naukri": "NAUKRI.NS",
    "zomato": "ZOMATO.NS",
    "paytm": "PAYTM.NS", "one97": "PAYTM.NS",
    "nykaa": "NYKAA.NS",
}

GLOBAL_COMPANIES = {
    "apple": "AAPL", "microsoft": "MSFT", "google": "GOOGL",
    "alphabet": "GOOGL", "amazon": "AMZN", "tesla": "TSLA",
    "meta": "META", "facebook": "META", "netflix": "NFLX",
    "nvidia": "NVDA", "amd": "AMD", "intel": "INTC",
    "salesforce": "CRM", "oracle": "ORCL", "sap": "SAP",
    "jpmorgan": "JPM", "jp morgan": "JPM", "goldman": "GS",
    "morgan stanley": "MS", "bank of america": "BAC",
    "walmart": "WMT", "target": "TGT", "costco": "COST",
    "disney": "DIS", "nike": "NKE", "coca cola": "KO",
    "pepsi": "PEP", "mcdonalds": "MCD", "starbucks": "SBUX",
    "johnson": "JNJ", "pfizer": "PFE", "abbvie": "ABBV",
    "exxon": "XOM", "chevron": "CVX", "shell": "SHEL",
    "berkshire": "BRK-B", "visa": "V", "mastercard": "MA",
    "paypal": "PYPL", "uber": "UBER", "lyft": "LYFT",
    "airbnb": "ABNB", "spotify": "SPOT", "twitter": "X",
    "snap": "SNAP", "pinterest": "PINS",
}

PEER_MAP = {
    "INFY.NS":  ["TCS.NS", "WIPRO.NS", "HCLTECH.NS", "TECHM.NS", "MPHASIS.NS"],
    "TCS.NS":   ["INFY.NS", "WIPRO.NS", "HCLTECH.NS", "TECHM.NS", "LTI.NS"],
    "WIPRO.NS": ["INFY.NS", "TCS.NS", "HCLTECH.NS", "TECHM.NS", "MPHASIS.NS"],
    "RELIANCE.NS": ["ONGC.NS", "IOC.NS", "BPCL.NS", "HPCL.NS", "GAIL.NS"],
    "HDFCBANK.NS": ["ICICIBANK.NS", "SBIN.NS", "KOTAKBANK.NS", "AXISBANK.NS", "INDUSINDBK.NS"],
    "ICICIBANK.NS": ["HDFCBANK.NS", "SBIN.NS", "KOTAKBANK.NS", "AXISBANK.NS", "BANDHANBNK.NS"],
    "TATAMOTORS.NS": ["MARUTI.NS", "M&M.NS", "EICHERMOT.NS", "HEROMOTOCO.NS", "BAJAJ-AUTO.NS"],
    "AAPL": ["MSFT", "GOOGL", "META", "AMZN", "NVDA"],
    "MSFT": ["AAPL", "GOOGL", "META", "AMZN", "ORCL"],
    "GOOGL": ["AAPL", "MSFT", "META", "AMZN", "NFLX"],
    "NVDA": ["AMD", "INTC", "QCOM", "TSM", "AVGO"],
    "TSLA": ["GM", "F", "RIVN", "NIO", "STLA"],
}


class ResearchAgent:

    def __init__(self, session: dict):
        self.session = session

    def _log(self, msg: str, status: str = "info"):
        self.session["logs"].append({
            "agent": "Research Agent",
            "message": msg,
            "status": status,
            "timestamp": __import__("datetime").datetime.now().isoformat()
        })

    async def fetch(self):
        """Main entry point — extract company name and fetch data"""
        raw_query = self.session.get("company_name", "")
        company_name = self._extract_company_name(raw_query)
        self.session["company_name"] = company_name

        self._log(f"Researching '{company_name}'...", "thinking")
        await asyncio.sleep(0.2)

        data = await self._fetch_company_data(company_name)
        self.session["company_data"] = data

        if not data.get("found"):
            self.session["phase"] = "error"
            self.session["error"] = f"Could not find financial data for '{company_name}'. Please try the stock ticker directly (e.g. INFY, TCS, RELIANCE)."
            self._log(f"Company not found: {company_name}", "error")

    def _extract_company_name(self, query: str) -> str:
        """Extract clean company name from user query using AI or rules"""
        # Remove common prefixes
        query = query.strip()
        prefixes = [
            "analyse ", "analyze ", "build a dcf for ", "build dcf for ",
            "build a lbo for ", "build lbo for ", "build a 3-statement for ",
            "build 3-statement for ", "build a fpa for ", "build fpa for ",
            "build a model for ", "build model for ", "create a model for ",
            "create dcf for ", "dcf for ", "lbo for ", "model for ",
            "research ", "tell me about ", "what about ", "analyse and build ",
            "build a dcf model for ", "build dcf model for ",
            "build a ", "create a ", "analyse ", "analyze ",
            "give me a dcf for ", "give me ", "show me ",
        ]
        query_lower = query.lower()
        for prefix in sorted(prefixes, key=len, reverse=True):
            if query_lower.startswith(prefix):
                query = query[len(prefix):].strip()
                query_lower = query.lower()
                break

        # Remove trailing model type mentions
        suffixes = [
            " dcf model", " lbo model", " 3-statement model", " fpa model",
            " model", " analysis", " valuation", " and build a dcf",
            " and build dcf", " dcf", " lbo"
        ]
        for suffix in sorted(suffixes, key=len, reverse=True):
            if query_lower.endswith(suffix):
                query = query[:-len(suffix)].strip()
                query_lower = query.lower()
                break

        return query.strip()

    async def _fetch_company_data(self, company_name: str) -> dict:
        """Fetch company data from best available source"""
        name_lower = company_name.lower().strip()

        # Check Indian companies first
        ticker = None
        is_indian = False
        for key, tkr in INDIAN_COMPANIES.items():
            if key in name_lower or name_lower in key:
                ticker = tkr
                is_indian = True
                break

        # Check global companies
        if not ticker:
            for key, tkr in GLOBAL_COMPANIES.items():
                if key in name_lower or name_lower in key:
                    ticker = tkr
                    is_indian = False
                    break

        # Try Screener.in for Indian companies
        if is_indian:
            self._log("Indian company detected — trying Screener.in first", "info")
            data = await self._scrape_screener(company_name)
            if data.get("found"):
                data["is_indian"] = True
                if ticker:
                    # Also get Yahoo data for beta and market data
                    yahoo_supplement = await self._fetch_yahoo_supplement(ticker)
                    data.update({k: v for k, v in yahoo_supplement.items() if v and k not in data})
                data["peers"] = await self._fetch_peers(ticker or company_name, True)
                return data
            # Fallback to Yahoo
            self._log("Screener.in failed — trying Yahoo Finance", "warning")
            if ticker:
                data = await self._fetch_yahoo(ticker)
                if data.get("found"):
                    data["is_indian"] = True
                    data["peers"] = await self._fetch_peers(ticker, True)
                    return data

        # Global company via Yahoo
        if ticker:
            self._log(f"Fetching {company_name} ({ticker}) from Yahoo Finance", "info")
            data = await self._fetch_yahoo(ticker)
            if data.get("found"):
                data["is_indian"] = False
                data["peers"] = await self._fetch_peers(ticker, False)
                return data

        # Last resort: try as ticker directly
        self._log(f"Trying '{company_name.upper()}' as direct ticker", "thinking")
        data = await self._fetch_yahoo(company_name.upper())
        if data.get("found"):
            data["is_indian"] = ".NS" in company_name.upper() or ".BO" in company_name.upper()
            data["peers"] = await self._fetch_peers(company_name.upper(), data["is_indian"])
            return data

        return {"found": False, "company_name": company_name}

    async def _scrape_screener(self, company_name: str) -> dict:
        """Scrape Screener.in for Indian company financials"""
        try:
            import requests
            from bs4 import BeautifulSoup

            headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/91.0.4472.124 Safari/537.36"}

            # Search API
            search_url = f"https://www.screener.in/api/company/search/?q={company_name.replace(' ', '+')}&v=3&fts=1"
            self._log(f"Searching Screener.in for '{company_name}'...", "thinking")

            resp = requests.get(search_url, headers=headers, timeout=12)
            if resp.status_code != 200 or not resp.json():
                return {"found": False}

            results = resp.json()
            company_url = results[0].get("url", "")
            slug = company_url.strip("/").split("/")[-1]
            found_name = results[0].get("name", company_name)

            self._log(f"Found on Screener.in: {found_name}", "success")

            # Try consolidated first, then standalone
            for suffix in ["/consolidated/", "/"]:
                page_url = f"https://www.screener.in/company/{slug}{suffix}"
                page_resp = requests.get(page_url, headers=headers, timeout=15)
                if page_resp.status_code == 200:
                    soup = BeautifulSoup(page_resp.content, "lxml")
                    data = self._parse_screener_page(soup, found_name)
                    if data.get("revenue_history"):
                        data["found"] = True
                        data["source"] = "Screener.in"
                        return data

            return {"found": False}

        except Exception as e:
            self._log(f"Screener.in error: {str(e)[:80]}", "warning")
            return {"found": False}

    def _parse_screener_page(self, soup, company_name: str) -> dict:
        """Parse Screener.in HTML for financial data"""
        data = {"company_name": company_name}

        try:
            # Key ratios from top section
            ratios = soup.find_all("li", {"class": re.compile(".*")})
            for li in ratios:
                text = li.get_text(strip=True).lower()
                value_el = li.find("span", {"class": re.compile(".*number.*|.*value.*")})
                if not value_el:
                    continue
                val = self._clean_number(value_el.get_text(strip=True))
                if "market cap" in text and val:
                    data["market_cap"] = val
                elif "p/e" in text and val and "pe" not in data:
                    data["pe_ratio"] = val
                elif "book value" in text and val:
                    data["book_value"] = val
                elif "dividend yield" in text and val:
                    data["dividend_yield"] = val
                elif "roce" in text and val:
                    data["roce"] = val
                elif "roe" in text and val:
                    data["roe"] = val

            # Financial tables
            tables = soup.find_all("section", {"id": re.compile("profit-loss|balance-sheet|cash-flow", re.I)})
            if not tables:
                tables_raw = soup.find_all("table")
                tables = tables_raw

            for table in tables:
                rows = table.find_all("tr")
                for row in rows:
                    cells = [td.get_text(strip=True) for td in row.find_all(["td", "th"])]
                    if len(cells) < 3:
                        continue
                    label = cells[0].lower()
                    values = [self._clean_number(c) for c in cells[1:] if c and c != "TTM"]
                    values = [v for v in values if v is not None and v != 0]

                    if not values:
                        continue

                    if any(k in label for k in ["sales", "revenue", "net sales"]) and "revenue_history" not in data:
                        data["revenue_history"] = values[-5:][::-1]  # oldest to newest → reverse for recent first
                        self._log(f"Revenue (5yr ₹Cr): {data['revenue_history']}", "success")
                    elif "ebitda" in label and "ebitda_history" not in data:
                        data["ebitda_history"] = values[-5:][::-1]
                        self._log(f"EBITDA (5yr ₹Cr): {data['ebitda_history']}", "success")
                    elif any(k in label for k in ["net profit", "profit after tax", "pat"]) and "net_income_history" not in data:
                        data["net_income_history"] = values[-5:][::-1]
                        self._log(f"Net Profit (5yr ₹Cr): {data['net_income_history']}", "success")
                    elif any(k in label for k in ["total debt", "borrowings"]) and "total_debt" not in data:
                        data["total_debt"] = values[-1]
                    elif "cash" in label and "equivalents" in label and "cash" not in data:
                        data["cash"] = values[-1]
                    elif "eps" in label and "eps" not in data:
                        data["eps"] = values[-1]
                    elif "operating profit margin" in label and "ebitda_margin" not in data:
                        data["ebitda_margin"] = values[-1] / 100 if values[-1] > 1 else values[-1]

            # Estimate EBITDA from operating profit if not found
            if "ebitda_history" not in data and "net_income_history" in data:
                self._log("EBITDA not found — estimating from net profit", "warning")

            data["sector"] = data.get("sector", "Indian Equity")
            data["currency"] = "INR"
            data["missing_fields"] = []

        except Exception as e:
            self._log(f"Parse error: {str(e)[:80]}", "warning")

        return data

    async def _fetch_yahoo(self, ticker: str) -> dict:
        """Fetch company financials from Yahoo Finance"""
        try:
            import yfinance as yf

            self._log(f"Fetching Yahoo Finance: {ticker}", "thinking")
            stock = yf.Ticker(ticker)
            info = stock.info

            if not info or not info.get("regularMarketPrice") and not info.get("currentPrice"):
                return {"found": False}

            company_name = info.get("longName") or info.get("shortName") or ticker
            self._log(f"Found: {company_name}", "success")

            is_indian = ticker.endswith(".NS") or ticker.endswith(".BO")
            divisor = 1e7 if is_indian else 1e6  # Convert to Cr for Indian, M for global
            currency = "INR" if is_indian else info.get("currency", "USD")

            data = {
                "found": True,
                "source": "Yahoo Finance",
                "company_name": company_name,
                "ticker": ticker,
                "sector": info.get("sector", "Unknown"),
                "industry": info.get("industry", "Unknown"),
                "current_price": info.get("regularMarketPrice") or info.get("currentPrice", 0),
                "market_cap": round((info.get("marketCap", 0) or 0) / divisor, 1),
                "beta": info.get("beta"),
                "pe_ratio": round(info.get("trailingPE", 0) or 0, 1),
                "pb_ratio": round(info.get("priceToBook", 0) or 0, 1),
                "total_debt": round((info.get("totalDebt", 0) or 0) / divisor, 1),
                "cash": round((info.get("totalCash", 0) or 0) / divisor, 1),
                "currency": currency,
                "shares_outstanding": round((info.get("sharesOutstanding", 0) or 0) / 1e6, 1),
                "missing_fields": [],
            }

            # Historical financials
            try:
                financials = stock.financials
                if financials is not None and not financials.empty:
                    for idx in financials.index:
                        idx_str = str(idx).lower()
                        row = [round(v/divisor, 1) for v in financials.loc[idx].values if v and str(v) != 'nan']
                        if not row:
                            continue
                        if "total revenue" in idx_str or "revenue" in idx_str:
                            data["revenue_history"] = row[:5]
                            self._log(f"Revenue history: {data['revenue_history']}", "success")
                        elif "ebitda" in idx_str:
                            data["ebitda_history"] = row[:5]
                            self._log(f"EBITDA history: {data['ebitda_history']}", "success")
                        elif "net income" in idx_str:
                            data["net_income_history"] = row[:5]
                            self._log(f"Net Income history: {data['net_income_history']}", "success")
                        elif "operating income" in idx_str and "ebitda_history" not in data:
                            data["ebitda_history"] = row[:5]
            except Exception as e:
                self._log(f"Financials parse: {str(e)[:50]}", "warning")

            return data

        except Exception as e:
            self._log(f"Yahoo Finance error: {str(e)[:80]}", "error")
            return {"found": False}

    async def _fetch_yahoo_supplement(self, ticker: str) -> dict:
        """Get supplementary data from Yahoo (beta, price, ratios)"""
        try:
            import yfinance as yf
            stock = yf.Ticker(ticker)
            info = stock.info
            return {
                "beta": info.get("beta"),
                "pe_ratio": round(info.get("trailingPE", 0) or 0, 1),
                "ticker": ticker,
                "sector": info.get("sector"),
                "industry": info.get("industry"),
            }
        except:
            return {}

    async def _fetch_peers(self, ticker_or_name: str, is_indian: bool) -> list:
        """Fetch peer company data for COMPS analysis"""
        peers_tickers = PEER_MAP.get(ticker_or_name, [])

        if not peers_tickers:
            # Find by partial match
            for key, tickers in PEER_MAP.items():
                if ticker_or_name in key or key in ticker_or_name:
                    peers_tickers = tickers
                    break

        if not peers_tickers:
            peers_tickers = ["TCS.NS", "INFY.NS", "WIPRO.NS", "HCLTECH.NS", "TECHM.NS"] if is_indian else ["AAPL", "MSFT", "GOOGL", "AMZN", "META"]

        peers_data = []
        try:
            import yfinance as yf
            for ticker in peers_tickers[:5]:
                try:
                    info = yf.Ticker(ticker).info
                    if info and (info.get("regularMarketPrice") or info.get("currentPrice")):
                        is_ind = ticker.endswith(".NS")
                        div = 1e7 if is_ind else 1e6
                        peers_data.append({
                            "name": info.get("shortName", ticker),
                            "ticker": ticker,
                            "ev_ebitda": round(info.get("enterpriseToEbitda") or 0, 1),
                            "pe_ratio": round(info.get("trailingPE") or 0, 1),
                            "ev_revenue": round(info.get("enterpriseToRevenue") or 0, 1),
                            "market_cap": round((info.get("marketCap") or 0) / div, 0),
                            "revenue_growth": round((info.get("revenueGrowth") or 0) * 100, 1),
                            "ebitda_margin": round((info.get("ebitdaMargins") or 0) * 100, 1),
                            "roe": round((info.get("returnOnEquity") or 0) * 100, 1),
                        })
                        self._log(f"Peer: {info.get('shortName', ticker)}", "info")
                    await asyncio.sleep(0.5)  # Rate limit protection
                except:
                    pass
        except:
            self._log("Peer fetch skipped", "warning")

        return peers_data

    def _clean_number(self, text: str) -> Optional[float]:
        if not text:
            return None
        text = str(text).replace(",", "").replace("₹", "").replace("$", "").replace("%", "").strip()
        text = re.sub(r'[^\d.\-]', '', text)
        try:
            return float(text) if text else None
        except:
            return None
