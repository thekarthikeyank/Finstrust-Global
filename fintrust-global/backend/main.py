"""
Fintrust Global — FastAPI Backend
Finance Chat AI + Model Generation Pipeline
"""

from fastapi import FastAPI, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse
from pydantic import BaseModel
import asyncio, json, os, uuid, sys
from typing import Optional
sys.path.insert(0, os.path.dirname(__file__))

app = FastAPI(title="Fintrust Global API", version="1.0.0")
app.add_middleware(CORSMiddleware,
    allow_origins=["*"], allow_credentials=True,
    allow_methods=["*"], allow_headers=["*"])

sessions = {}

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
    return {"status": "Fintrust Global API Running", "version": "1.0.0"}

@app.post("/api/session/new")
def new_session():
    sid = str(uuid.uuid4())
    get_session(sid)
    return {"session_id": sid}

@app.post("/api/chat")
async def finance_chat(req: ChatRequest):
    """Handle general finance questions via AI"""
    answer = await get_finance_answer(req.question)
    return {"answer": answer, "session_id": req.session_id}

async def get_finance_answer(question: str) -> str:
    """Get answer from Gemini or Groq"""

    FINANCE_SYSTEM = """You are Fintrust Global, an expert AI financial analyst.
You help students and retail investors understand finance, stocks, and markets.
Rules:
- Give precise, data-backed answers when possible
- Use Indian market context when relevant (NSE, BSE, Nifty, Sensex)
- Explain financial concepts clearly without jargon
- Always add disclaimer: "Not financial advice — for educational purposes"
- Keep responses concise but complete
- Use **bold** for key terms and numbers
- Format numbers properly (₹ for INR, $ for USD)"""

    # Try Gemini first
    try:
        import google.generativeai as genai
        import os
        api_key = os.environ.get("GEMINI_API_KEY", "")
        if api_key:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel("gemini-1.5-flash",
                system_instruction=FINANCE_SYSTEM)
            response = model.generate_content(question)
            return response.text
    except Exception as e:
        pass

    # Try Groq
    try:
        from groq import Groq
        import os
        api_key = os.environ.get("GROQ_API_KEY", "")
        if api_key:
            client = Groq(api_key=api_key)
            res = client.chat.completions.create(
                model="llama-3.1-70b-versatile",
                messages=[
                    {"role": "system", "content": FINANCE_SYSTEM},
                    {"role": "user", "content": question}
                ],
                max_tokens=800
            )
            return res.choices[0].message.content
    except:
        pass

    # Fallback rule-based
    return get_fallback_answer(question)

def get_fallback_answer(question: str) -> str:
    q = question.lower()
    if "dcf" in q or "discounted cash flow" in q:
        return "**DCF (Discounted Cash Flow)** values a company by estimating future free cash flows and discounting them to present value using WACC. It is the gold standard for intrinsic valuation. Key inputs: revenue growth, EBITDA margin, WACC, and terminal growth rate.\n\n*Not financial advice — for educational purposes*"
    elif "ebitda" in q:
        return "**EBITDA** = Earnings Before Interest, Taxes, Depreciation and Amortization. It measures operating profitability before financing and accounting decisions. A 25% EBITDA margin means ₹25 of operating profit for every ₹100 of revenue. Used widely in valuation multiples like EV/EBITDA.\n\n*Not financial advice — for educational purposes*"
    elif "wacc" in q:
        return "**WACC** (Weighted Average Cost of Capital) is the blended cost of a company's funding sources — equity and debt. Formula: WACC = (E/V × Ke) + (D/V × Kd × (1-t)). For Indian large-caps it typically ranges from 10-14%. Used as the discount rate in DCF models.\n\n*Not financial advice — for educational purposes*"
    elif "pe" in q or "p/e" in q:
        return "**P/E Ratio** (Price-to-Earnings) = Market Price per share ÷ EPS. Nifty 50 trades at ~22x P/E historically. IT sector trades at premium (25-35x) due to high ROCE and growth. A stock trading below its 5yr average P/E may indicate undervaluation — but context matters.\n\n*Not financial advice — for educational purposes*"
    return "I can help you analyse any listed company, explain financial concepts, or build valuation models. Try asking: 'Analyse Infosys', 'Explain DCF', or 'Is TCS overvalued?'\n\n*Not financial advice — for educational purposes*"

@app.post("/api/research")
async def research_company(req: CompanyRequest, background_tasks: BackgroundTasks):
    sid = req.session_id or str(uuid.uuid4())
    session = get_session(sid)
    session["company_name"] = req.company_name
    session["phase"] = "researching"
    session["logs"] = []
    background_tasks.add_task(run_research_pipeline, sid, req.company_name)
    return {"session_id": sid, "status": "researching"}

@app.post("/api/confirm-model")
async def confirm_model(req: ConfirmModel, background_tasks: BackgroundTasks):
    session = get_session(req.session_id)
    if req.model_type:
        session["model_recommendation"] = req.model_type
    session["phase"] = "planning"
    background_tasks.add_task(run_planning_pipeline, req.session_id)
    return {"status": "planning"}

@app.post("/api/build")
async def build_model(req: ConfirmModel, background_tasks: BackgroundTasks):
    session = get_session(req.session_id)
    session["phase"] = "building"
    background_tasks.add_task(run_build_pipeline, req.session_id)
    return {"status": "building"}

@app.get("/api/logs/{session_id}")
async def stream_logs(session_id: str):
    async def generator():
        session = get_session(session_id)
        sent = 0
        while True:
            logs = session.get("logs", [])
            for log in logs[sent:]:
                yield f"data: {json.dumps(log)}\n\n"
            sent = len(logs)
            if session.get("phase") in ["delivered", "error", "awaiting_confirmation"]:
                yield f"data: {json.dumps({'type':'done','phase':session['phase']})}\n\n"
                break
            await asyncio.sleep(0.3)
    return StreamingResponse(generator(), media_type="text/event-stream")

@app.get("/api/status/{session_id}")
def get_status(session_id: str):
    session = get_session(session_id)
    return {
        "phase": session.get("phase"),
        "model_recommendation": session.get("model_recommendation"),
        "missing_fields": session.get("missing_fields", []),
        "qa_report": session.get("qa_report"),
        "narrator_notes": session.get("narrator_notes", []),
        "assumptions_summary": session.get("assumptions", {}),
        "company_data": session.get("company_data", {}),
    }

@app.get("/api/download/{session_id}")
def download_excel(session_id: str):
    session = get_session(session_id)
    path = session.get("excel_path")
    if not path or not os.path.exists(path):
        return {"error": "File not ready"}
    company = session.get("company_name", "Company").replace(" ", "_")
    model = session.get("model_recommendation", "Model")
    return FileResponse(path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"{company}_{model}_FintrustGlobal.xlsx")

async def run_research_pipeline(session_id, company_name):
    from agents.research_agent import ResearchAgent
    from agents.analyst_agent import AnalystAgent
    session = get_session(session_id)
    try:
        add_log(session, "Research Agent", f"Researching '{company_name}'...", "thinking")
        researcher = ResearchAgent(session)
        data = await researcher.research(company_name)
        session["company_data"] = data
        add_log(session, "Research Agent", "Financial data extracted", "success")
        add_log(session, "Analyst Agent", "Analysing financial profile...", "thinking")
        analyst = AnalystAgent(session)
        rec = await analyst.recommend()
        session["model_recommendation"] = rec["model_type"]
        session["phase"] = "awaiting_confirmation"
        add_log(session, "Analyst Agent", f"Recommendation: {rec['model_type']}", "success")
    except Exception as e:
        add_log(session, "Research Agent", f"Error: {str(e)[:60]}", "error")
        session["phase"] = "error"

async def run_planning_pipeline(session_id):
    from agents.planning_agent import PlanningAgent
    session = get_session(session_id)
    try:
        add_log(session, "Planning Agent", "Mapping data to assumptions...", "thinking")
        planner = PlanningAgent(session)
        result = await planner.plan()
        session["assumptions"] = result["assumptions"]
        session["narrator_notes"] = result["narrator_notes"]
        session["phase"] = "building"
        add_log(session, "Planning Agent", "Assumptions ready — building model", "success")
        await run_build_pipeline(session_id)
    except Exception as e:
        add_log(session, "Planning Agent", f"Error: {str(e)[:60]}", "error")
        session["phase"] = "error"

async def run_build_pipeline(session_id):
    from agents.build_agent import BuildAgent, QAAgent, DeliveryAgent
    session = get_session(session_id)
    try:
        add_log(session, "Build Agent", "Building Excel model...", "thinking")
        builder = BuildAgent(session)
        path = await builder.build()
        session["excel_path"] = path
        add_log(session, "Build Agent", "Model generated", "success")
        add_log(session, "QA Agent", "Running quality checks...", "thinking")
        qa = QAAgent(session)
        report = await qa.audit(path)
        session["qa_report"] = report
        if not report["passed"]:
            fixed = await qa.auto_fix(path, report["issues"])
            if fixed:
                session["excel_path"] = fixed
        add_log(session, "QA Agent", f"{report['checks_passed']} checks passed", "success")
        add_log(session, "Delivery Agent", "Model ready for download", "success")
        session["phase"] = "delivered"
    except Exception as e:
        add_log(session, "Build Agent", f"Error: {str(e)[:60]}", "error")
        session["phase"] = "error"
