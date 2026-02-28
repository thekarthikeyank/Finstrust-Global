"""
Fintrust Global — FastAPI Backend
CFA-Level AI Finance Analyst
"""

from fastapi import FastAPI, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse
from pydantic import BaseModel
import asyncio, json, os, uuid, sys
from typing import Optional
sys.path.insert(0, os.path.dirname(__file__))

app = FastAPI(title="Fintrust Global API", version="2.0.0")
app.add_middleware(CORSMiddleware,
    allow_origins=["*"], allow_credentials=True,
    allow_methods=["*"], allow_headers=["*"])

sessions = {}

CFA_ANALYST_SYSTEM = """You are a CFA charterholder and Senior Equity Research Analyst with 15 years of experience at a bulge-bracket investment bank. You have deep expertise in:

- Fundamental analysis: DCF, LBO, 3-Statement, Comparable Company Analysis
- Indian markets: NSE, BSE, Nifty 50, sectoral indices, SEBI regulations
- Global markets: NYSE, NASDAQ, S&P 500, sector dynamics
- Valuation frameworks: EV/EBITDA, P/E, P/B, EV/Sales, PEG, DDM
- Risk frameworks: Beta, WACC, Cost of Capital, Credit Risk
- Accounting: Ind AS, IFRS, US GAAP — you can read between the lines

Your communication style:
- Direct and confident like a senior analyst on a client call
- Use precise financial terminology but explain it when needed
- Back every claim with numbers and data
- Give actionable insights, not generic commentary
- When you don't have data, say so clearly and explain what you'd need
- Format responses with **bold** for key metrics and numbers
- Use ₹ for Indian stocks, $ for US stocks
- Always end with: "⚠️ Not financial advice — for educational purposes only"

When analyzing a company:
1. Start with the INVESTMENT THESIS in 2-3 sentences
2. Key financial metrics (revenue, margins, growth, valuation)
3. Bull case vs Bear case
4. Peer comparison if relevant
5. Key risks
6. Your view: Attractive / Fair Value / Expensive

You think in numbers. You speak like a Bloomberg terminal that can talk."""


class ChatRequest(BaseModel):
    question: str
    session_id: Optional[str] = None

class CompanyRequest(BaseModel):
    company_name: str
    session_id: Optional[str] = None

class ConfirmModel(BaseModel):
    session_id: str
    confirmed: bool
    model_type: Optional[str] = None

class MissingData(BaseModel):
    session_id: str
    field: str
    value: str


def get_session(sid: str) -> dict:
    if sid not in sessions:
        sessions[sid] = {
            "id": sid, "phase": "idle",
            "company_name": None, "company_data": {},
            "model_recommendation": None, "assumptions": {},
            "narrator_notes": [], "missing_fields": [],
            "excel_path": None, "qa_report": None, "logs": [],
        }
    return sessions[sid]

def add_log(session, agent, message, status="info"):
    import datetime
    session["logs"].append({
        "agent": agent, "message": message, "status": status,
        "timestamp": datetime.datetime.now().isoformat()
    })

@app.get("/")
def root():
    return {"status": "Fintrust Global API Running", "version": "2.0.0"}

@app.post("/api/session/new")
def new_session():
    sid = str(uuid.uuid4())
    get_session(sid)
    return {"session_id": sid}

@app.post("/api/chat")
async def finance_chat(req: ChatRequest):
    """Handle general finance questions via CFA-level AI"""
    answer = await get_finance_answer(req.question)
    return {"answer": answer, "session_id": req.session_id}


async def get_finance_answer(question: str) -> str:
    """Get CFA-level answer from Gemini or Groq"""

    # Try Gemini first
    try:
        import google.generativeai as genai
        api_key = os.environ.get("GEMINI_API_KEY", "")
        if api_key:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(
                "gemini-1.5-flash",
                system_instruction=CFA_ANALYST_SYSTEM
            )
            response = model.generate_content(question)
            return response.text
    except Exception as e:
        pass

    # Try Groq fallback
    try:
        from groq import Groq
        api_key = os.environ.get("GROQ_API_KEY", "")
        if api_key:
            client = Groq(api_key=api_key)
            res = client.chat.completions.create(
                model="llama-3.1-70b-versatile",
                messages=[
                    {"role": "system", "content": CFA_ANALYST_SYSTEM},
                    {"role": "user", "content": question}
                ],
                max_tokens=1200
            )
            return res.choices[0].message.content
    except:
        pass

    return "I'm having trouble connecting to the AI engine. Please try again in a moment."


@app.post("/api/research")
async def start_research(req: CompanyRequest, background_tasks: BackgroundTasks):
    """Start company research and model generation pipeline"""
    sid = req.session_id or str(uuid.uuid4())
    session = get_session(sid)
    session["phase"] = "researching"
    session["company_name"] = req.company_name
    session["logs"] = []

    background_tasks.add_task(run_pipeline, sid)
    return {"session_id": sid, "status": "started"}


async def run_pipeline(sid: str):
    """Run the full 6-agent pipeline"""
    session = get_session(sid)
    try:
        from agents.research_agent import ResearchAgent
        from agents.analyst_agent import AnalystAgent
        from agents.planning_agent import PlanningAgent
        from agents.build_agent import BuildAgent

        # Agent 1: Research
        research = ResearchAgent(session)
        await research.fetch()

        if session.get("phase") == "error":
            return

        # Agent 2: Analyst — CFA-level recommendation
        analyst = AnalystAgent(session)
        recommendation = await analyst.recommend()

        session["phase"] = "awaiting_confirmation"
        session["model_recommendation"] = recommendation["model_type"]
        session["analyst_reasoning"] = recommendation.get("reasoning", "")
        session["key_metrics"] = recommendation.get("key_metrics", {})

        # Generate deep analysis
        await generate_deep_analysis(session)

    except Exception as e:
        session["phase"] = "error"
        session["error"] = str(e)
        add_log(session, "System", f"Pipeline error: {str(e)}", "error")


async def generate_deep_analysis(session: dict):
    """Generate CFA-level investment analysis"""
    data = session.get("company_data", {})
    company = data.get("company_name", "the company")
    sector = data.get("sector", "")
    revenue = data.get("revenue_history", [])
    ebitda = data.get("ebitda_history", [])
    market_cap = data.get("market_cap", 0)
    pe_ratio = data.get("pe_ratio", 0)
    model_type = session.get("model_recommendation", "DCF")

    # Calculate key metrics
    rev_cagr = 0
    if len(revenue) >= 2:
        try:
            rev_cagr = (revenue[0] / revenue[-1]) ** (1/(len(revenue)-1)) - 1
        except:
            rev_cagr = 0

    avg_margin = 0
    if revenue and ebitda:
        try:
            margins = [e/r for e,r in zip(ebitda, revenue) if r > 0]
            avg_margin = sum(margins)/len(margins) if margins else 0
        except:
            avg_margin = 0

    analysis_prompt = f"""Analyze {company} ({sector} sector) as a CFA-level equity analyst.

Financial Data:
- Revenue History (most recent first, ₹ Cr): {revenue[:5] if revenue else 'Not available'}
- EBITDA History (most recent first, ₹ Cr): {ebitda[:5] if ebitda else 'Not available'}
- Revenue CAGR: {rev_cagr:.1%}
- Avg EBITDA Margin: {avg_margin:.1%}
- Market Cap: ₹{market_cap:,.0f} Cr
- P/E Ratio: {pe_ratio}x
- Recommended Model: {model_type}

Provide a structured analyst note with:
1. **Investment Thesis** (2-3 sentences, direct view)
2. **Key Financial Metrics** (growth, margins, valuation vs peers)
3. **Bull Case** (what could drive upside)
4. **Bear Case** (key risks and downside scenarios)
5. **Valuation View** (Attractive / Fair Value / Expensive — with reasoning)
6. **Why {model_type} model** (explain why this is the right framework)

Be specific with numbers. Think like a Goldman Sachs research note."""

    try:
        import google.generativeai as genai
        api_key = os.environ.get("GEMINI_API_KEY", "")
        if api_key:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(
                "gemini-1.5-flash",
                system_instruction=CFA_ANALYST_SYSTEM
            )
            response = model.generate_content(analysis_prompt)
            session["deep_analysis"] = response.text
            add_log(session, "Analyst Agent", "CFA-level analysis complete", "success")
            return
    except:
        pass

    try:
        from groq import Groq
        api_key = os.environ.get("GROQ_API_KEY", "")
        if api_key:
            client = Groq(api_key=api_key)
            res = client.chat.completions.create(
                model="llama-3.1-70b-versatile",
                messages=[
                    {"role": "system", "content": CFA_ANALYST_SYSTEM},
                    {"role": "user", "content": analysis_prompt}
                ],
                max_tokens=1500
            )
            session["deep_analysis"] = res.choices[0].message.content
            add_log(session, "Analyst Agent", "CFA-level analysis complete", "success")
    except:
        session["deep_analysis"] = session.get("analyst_reasoning", "Analysis unavailable.")


@app.post("/api/confirm-model")
async def confirm_model(req: ConfirmModel, background_tasks: BackgroundTasks):
    """User confirms model type, start building"""
    session = get_session(req.session_id)

    if req.model_type:
        session["model_recommendation"] = req.model_type

    if req.confirmed:
        session["phase"] = "building"
        background_tasks.add_task(build_model, req.session_id)
        return {"status": "building", "model_type": session["model_recommendation"]}

    return {"status": "cancelled"}


async def build_model(sid: str):
    """Build the Excel model"""
    session = get_session(sid)
    try:
        from agents.build_agent import BuildAgent
        builder = BuildAgent(session)
        await builder.build()
    except Exception as e:
        session["phase"] = "error"
        session["error"] = str(e)
        add_log(session, "Build Agent", f"Build error: {str(e)}", "error")


@app.get("/api/status/{session_id}")
async def get_status(session_id: str):
    """Get current session status"""
    session = get_session(session_id)
    return {
        "phase": session.get("phase", "idle"),
        "company_name": session.get("company_name"),
        "model_recommendation": session.get("model_recommendation"),
        "analyst_reasoning": session.get("analyst_reasoning", ""),
        "deep_analysis": session.get("deep_analysis", ""),
        "key_metrics": session.get("key_metrics", {}),
        "logs": session.get("logs", [])[-20:],
        "error": session.get("error"),
        "excel_ready": session.get("excel_path") is not None,
    }


@app.get("/api/download/{session_id}")
async def download_model(session_id: str):
    """Download the generated Excel model"""
    session = get_session(session_id)
    excel_path = session.get("excel_path")

    if not excel_path or not os.path.exists(excel_path):
        return {"error": "Model not ready yet"}

    company = session.get("company_name", "Company").replace(" ", "_")
    model_type = session.get("model_recommendation", "Model")
    filename = f"{company}_{model_type}_Fintrust.xlsx"

    return FileResponse(
        excel_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename
    )


@app.post("/api/provide-data")
async def provide_missing_data(req: MissingData):
    """User provides missing data"""
    session = get_session(req.session_id)
    session["company_data"][req.field] = req.value
    return {"status": "updated"}
