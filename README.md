# 🏦 Fintrust Global

> **AI-powered financial modeling platform for students and retail investors.**  
> Type a company name. Get an institutional-grade Excel model in seconds.

[![Live Demo](https://img.shields.io/badge/Live%20Demo-fintrust--global.vercel.app-blue?style=for-the-badge)](https://fintrust-global.vercel.app)
[![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)](LICENSE)
[![Built with Next.js](https://img.shields.io/badge/Next.js-14-black?style=for-the-badge&logo=next.js)](https://nextjs.org)
[![FastAPI](https://img.shields.io/badge/FastAPI-Python-009688?style=for-the-badge&logo=fastapi)](https://fastapi.tiangolo.com)

---

## ✨ What is Fintrust Global?

Fintrust Global is a **free SaaS platform** that lets anyone — students, retail investors, or finance enthusiasts — generate professional financial models just by typing a company name.

No finance degree required. No Excel expertise needed. Just type and download.

---

## 🎯 Features

### 💬 Finance Chat AI
- Ask any finance question in plain English
- Real-time answers with Indian market context (NSE, BSE, Nifty, Sensex)
- Explains concepts clearly: DCF, WACC, EBITDA, P/E ratios, and more

### 📊 Instant Excel Models
- **4 model types:** DCF, LBO, 3-Statement, FP&A
- **Real company data** auto-fetched from Screener.in (Indian) and Yahoo Finance (Global)
- **COMPS sheet** with 5 peer companies and trading multiples
- **Bull / Base / Bear scenarios** auto-generated from historical data
- **Smart Assumptions Narrator** explains every input in plain English

### 🤖 6-Agent AI Pipeline
| Agent | Role |
|-------|------|
| 🔍 Research Agent | Scrapes live data from Screener.in or Yahoo Finance |
| 🧠 Analyst Agent | Recommends model type, narrates assumptions |
| 📋 Planning Agent | Maps real data to scenarios (Bull/Base/Bear) |
| 🏗️ Build Agent | Generates institutional-grade Excel model |
| ✅ QA Agent | Audits formulas, charts, formatting — auto-fixes issues |
| 📦 Delivery Agent | Packages and delivers final Excel file |

### 📁 Excel Output Quality
- Institutional formatting standards (Calibri 11, color-coded inputs)
- No hardcoded numbers inside formulas
- Cover page with color legend
- Dashboard with KPI cards and charts
- Freeze panes, timeline headers, labeled units

---

## 🖥️ Tech Stack

| Layer | Technology |
|-------|-----------|
| Frontend | Next.js 14, Tailwind CSS |
| Backend | FastAPI (Python) |
| AI | Gemini 1.5 Flash (primary), Groq LLaMA 3.1 (fallback) |
| Auth & DB | Supabase |
| Excel Engine | openpyxl |
| Data Sources | Screener.in, Yahoo Finance (yfinance) |
| Hosting | Vercel (frontend) + Railway (backend) |

---

## 🚀 Deploy Your Own (Free)

### Prerequisites
- GitHub account
- [Gemini API key](https://aistudio.google.com) (free)
- [Groq API key](https://console.groq.com) (free)
- [Railway account](https://railway.app) (free)
- [Vercel account](https://vercel.com) (free)

### Step 1 — Clone & Push to GitHub
```bash
git clone https://github.com/thekarthikeyank/fintrust-global.git
cd fintrust-global
```

### Step 2 — Deploy Backend on Railway
1. Go to [railway.app](https://railway.app) → **New Project** → Deploy from GitHub
2. Select the `fintrust-global` repo → set root to `backend`
3. Add environment variables:
```
GEMINI_API_KEY=your_gemini_key_here
GROQ_API_KEY=your_groq_key_here
```
4. Settings → Domains → **Generate Domain**
5. Copy your Railway URL (e.g. `https://fintrust-global-production.up.railway.app`)

### Step 3 — Deploy Frontend on Vercel
1. Go to [vercel.com](https://vercel.com) → **Add New Project** → Import repo
2. Set root directory to `frontend`
3. Add environment variable:
```
NEXT_PUBLIC_API_URL=https://your-railway-url.up.railway.app
```
4. Click **Deploy** — your app is live!

### Step 4 — Optional: Enable Auth with Supabase
1. Go to [supabase.com](https://supabase.com) → New Project (free)
2. Enable Email Auth
3. Add to Vercel environment variables:
```
NEXT_PUBLIC_SUPABASE_URL=your_supabase_url
NEXT_PUBLIC_SUPABASE_ANON_KEY=your_anon_key
```

---

## 💰 Cost: $0/month

| Service | Free Tier |
|---------|-----------|
| Vercel | Unlimited hobby projects |
| Railway | 500 hours/month |
| Supabase | 50,000 monthly active users |
| Gemini 1.5 Flash | 15 RPM, 1M tokens/month |
| Groq LLaMA 3.1 | 14,400 requests/day |

---

## 📁 Project Structure

```
fintrust-global/
├── frontend/                  # Next.js app
│   └── src/app/
│       ├── page.jsx           # Landing page
│       ├── auth/              # Login & signup
│       └── chat/              # Main product UI
│
├── backend/                   # FastAPI app
│   ├── main.py                # API routes
│   ├── agents/
│   │   ├── research_agent.py  # Data scraping
│   │   ├── analyst_agent.py   # AI recommendations
│   │   ├── planning_agent.py  # Scenario generation
│   │   └── build_agent.py     # Excel + QA + Delivery
│   ├── builders/
│   │   ├── dcf_builder.py
│   │   ├── lbo_builder.py
│   │   ├── three_stmt_builder.py
│   │   └── fpa_builder.py
│   └── formatting/
│       └── institutional.py   # Excel formatting engine
│
└── README.md
```

---

## 🧠 How It Works

```
User types: "Build a DCF model for Infosys"
         ↓
🔍 Research Agent scrapes Screener.in for Infosys financials
         ↓
🧠 Analyst Agent recommends DCF, narrates assumptions
         ↓
📋 Planning Agent builds Bull/Base/Bear scenarios from real data
         ↓
🏗️ Build Agent generates 7-sheet Excel with COMPS
         ↓
✅ QA Agent audits formulas & charts, auto-fixes issues
         ↓
📦 User downloads institutional-grade Excel model
```

---

## 🤝 Contributing

Pull requests are welcome! For major changes, please open an issue first.

1. Fork the repo
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

---

## 📄 License

MIT License — free to use, modify, and distribute.

---

## 👤 Author

**Karthikeyan** — [@thekarthikeyank](https://github.com/thekarthikeyank)

---

<p align="center">Built with ❤️ for students and retail investors who deserve institutional-quality tools.</p>
