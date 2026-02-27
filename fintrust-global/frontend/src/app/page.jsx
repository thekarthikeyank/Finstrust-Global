'use client'
import Link from 'next/link'
import { useState } from 'react'
import {
  BarChart2, MessageSquare, FileSpreadsheet, TrendingUp,
  Shield, Globe, ChevronRight, Star, Check, Zap
} from 'lucide-react'

export default function LandingPage() {
  const [email, setEmail] = useState('')

  return (
    <div className="min-h-screen bg-white font-sans">

      {/* NAV */}
      <nav className="fixed top-0 left-0 right-0 z-50 bg-white/90 backdrop-blur-md border-b border-blue-100">
        <div className="max-w-6xl mx-auto px-6 h-16 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-[#1a3a6b] to-[#2563eb] flex items-center justify-center shadow">
              <BarChart2 size={16} className="text-white" />
            </div>
            <span className="text-lg font-bold text-[#1a3a6b] tracking-tight">
              Fintrust <span className="text-[#2563eb]">Global</span>
            </span>
          </div>
          <div className="hidden md:flex items-center gap-8">
            {['Features', 'How it Works', 'FAQ'].map(item => (
              <a key={item} href={`#${item.toLowerCase().replace(' ','-')}`}
                className="text-sm text-gray-600 hover:text-[#2563eb] font-medium transition-colors">
                {item}
              </a>
            ))}
          </div>
          <div className="flex items-center gap-3">
            <Link href="/auth/login"
              className="text-sm font-medium text-[#1a3a6b] hover:text-[#2563eb] transition-colors">
              Sign In
            </Link>
            <Link href="/auth/signup"
              className="bg-[#1a3a6b] hover:bg-[#2563eb] text-white text-sm font-semibold
              px-5 py-2.5 rounded-xl transition-colors shadow-sm">
              Get Started Free
            </Link>
          </div>
        </div>
      </nav>

      {/* HERO */}
      <section className="pt-32 pb-20 px-6 relative overflow-hidden">
        {/* Background */}
        <div className="absolute inset-0 bg-gradient-to-br from-[#f0f4ff] via-white to-[#e8f0fe] -z-10" />
        <div className="absolute top-20 right-0 w-96 h-96 bg-blue-100 rounded-full blur-3xl opacity-40 -z-10" />
        <div className="absolute bottom-0 left-0 w-64 h-64 bg-blue-200 rounded-full blur-3xl opacity-30 -z-10" />

        <div className="max-w-4xl mx-auto text-center">
          {/* Badge */}
          <div className="inline-flex items-center gap-2 bg-blue-50 border border-blue-200
            text-blue-700 text-xs font-semibold px-4 py-2 rounded-full mb-6">
            <Zap size={12} className="text-blue-500" />
            AI-Powered Financial Analysis — Free Forever
          </div>

          <h1 className="text-5xl md:text-6xl font-extrabold text-[#1a3a6b] leading-tight mb-6 tracking-tight">
            Your Personal{' '}
            <span className="relative">
              <span className="text-[#2563eb]">Finance Analyst</span>
              <svg className="absolute -bottom-2 left-0 w-full" viewBox="0 0 300 12" fill="none">
                <path d="M2 9C50 3 100 1 150 4C200 7 250 5 298 2" stroke="#2563eb" strokeWidth="3"
                  strokeLinecap="round" opacity="0.4"/>
              </svg>
            </span>
            <br />Powered by AI
          </h1>

          <p className="text-xl text-gray-600 max-w-2xl mx-auto leading-relaxed mb-10">
            Ask any finance question. Get instant answers, real market data,
            and institutional-grade Excel models — completely free.
          </p>

          {/* CTA */}
          <div className="flex flex-col sm:flex-row gap-3 justify-center mb-6">
            <Link href="/auth/signup"
              className="inline-flex items-center justify-center gap-2 bg-[#1a3a6b] hover:bg-[#2563eb]
              text-white font-bold px-8 py-4 rounded-2xl text-base transition-all shadow-lg
              hover:shadow-blue-200 hover:-translate-y-0.5 hover:shadow-xl">
              Start Free — No Credit Card
              <ChevronRight size={18} />
            </Link>
            <Link href="/auth/login"
              className="inline-flex items-center justify-center gap-2 border-2 border-[#1a3a6b]
              text-[#1a3a6b] hover:bg-blue-50 font-semibold px-8 py-4 rounded-2xl text-base transition-all">
              Sign In
            </Link>
          </div>

          <p className="text-xs text-gray-400">
            Join students & investors analysing Indian and Global markets
          </p>
        </div>

        {/* Chat Preview */}
        <div className="max-w-2xl mx-auto mt-16">
          <div className="bg-white rounded-3xl shadow-2xl border border-blue-100 overflow-hidden">
            {/* Window chrome */}
            <div className="bg-[#1a3a6b] px-5 py-3 flex items-center gap-2">
              <div className="flex gap-1.5">
                <div className="w-3 h-3 rounded-full bg-red-400 opacity-70"></div>
                <div className="w-3 h-3 rounded-full bg-yellow-400 opacity-70"></div>
                <div className="w-3 h-3 rounded-full bg-green-400 opacity-70"></div>
              </div>
              <span className="text-white/60 text-xs ml-2 font-medium">Fintrust Global — Chat</span>
            </div>
            {/* Chat messages */}
            <div className="p-5 space-y-4 bg-gray-50">
              <ChatPreviewMsg role="user" text="Analyse Infosys and build a DCF model" />
              <ChatPreviewMsg role="ai" text="Researching Infosys (INFY.NS)... Revenue CAGR: 13.2% | EBITDA Margin: 25.4% | I recommend a DCF model. Implied share price: ₹1,847 — currently trading at a slight discount to intrinsic value." />
              <ChatPreviewMsg role="user" text="Is TCS overvalued compared to peers?" />
              <ChatPreviewMsg role="ai" text="TCS trades at 28x P/E vs sector average of 24x — a 17% premium. Justified by superior ROCE of 48% and ₹52,000Cr cash on balance sheet. Not overvalued on fundamentals." />
            </div>
            {/* Input */}
            <div className="px-5 py-4 border-t border-gray-100 bg-white flex items-center gap-3">
              <input readOnly value="Ask anything about stocks, valuation, finance..."
                className="flex-1 text-sm text-gray-400 bg-transparent outline-none" />
              <div className="w-9 h-9 rounded-xl bg-[#1a3a6b] flex items-center justify-center">
                <ChevronRight size={16} className="text-white" />
              </div>
            </div>
          </div>
        </div>
      </section>

      {/* STATS */}
      <section className="py-12 bg-[#1a3a6b]">
        <div className="max-w-4xl mx-auto px-6 grid grid-cols-2 md:grid-cols-4 gap-8 text-center">
          {[
            { num: '4', label: 'Model Types', sub: 'DCF · LBO · 3-Stmt · FP&A' },
            { num: '5000+', label: 'Companies', sub: 'NSE · BSE · NYSE · NASDAQ' },
            { num: '100%', label: 'Free', sub: 'No credit card ever' },
            { num: 'Real', label: 'Live Data', sub: 'Screener.in · Yahoo Finance' },
          ].map(s => (
            <div key={s.label}>
              <div className="text-3xl font-extrabold text-white">{s.num}</div>
              <div className="text-blue-200 font-semibold text-sm mt-1">{s.label}</div>
              <div className="text-blue-400 text-xs mt-0.5">{s.sub}</div>
            </div>
          ))}
        </div>
      </section>

      {/* FEATURES */}
      <section id="features" className="py-20 px-6">
        <div className="max-w-6xl mx-auto">
          <div className="text-center mb-14">
            <div className="text-[#2563eb] font-bold text-sm uppercase tracking-widest mb-3">Features</div>
            <h2 className="text-4xl font-extrabold text-[#1a3a6b]">Everything You Need to Analyse Markets</h2>
            <p className="text-gray-500 mt-4 max-w-xl mx-auto">
              Professional-grade tools that were previously only available to investment bankers — now free for everyone.
            </p>
          </div>
          <div className="grid md:grid-cols-3 gap-6">
            {[
              {
                icon: MessageSquare, color: 'bg-blue-50 text-blue-600',
                title: 'Finance Chat AI',
                desc: 'Ask any question in plain English. Get precise answers backed by real market data, not hallucinations.',
                examples: ['Is Reliance overvalued?', 'Explain P/E ratio simply', 'Compare HDFC vs ICICI']
              },
              {
                icon: FileSpreadsheet, color: 'bg-indigo-50 text-indigo-600',
                title: 'Instant Excel Models',
                desc: 'Generate institutional-grade DCF, LBO, 3-Statement and FP&A models with one sentence.',
                examples: ['DCF model for TCS', 'LBO analysis Zomato', '3-Statement for Wipro']
              },
              {
                icon: TrendingUp, color: 'bg-emerald-50 text-emerald-600',
                title: 'Real Market Data',
                desc: 'Live financial data from NSE, BSE, NYSE and NASDAQ. 5-year historical financials auto-populated.',
                examples: ['Infosys 5yr revenue', 'HDFC Bank margins', 'Nifty 50 valuation']
              },
              {
                icon: BarChart2, color: 'bg-violet-50 text-violet-600',
                title: 'Scenario Analysis',
                desc: 'Every model includes Bull, Base and Bear scenarios automatically generated from historical data.',
                examples: ['Bull case Tata Motors', 'Bear scenario IT sector', 'Base case pharma']
              },
              {
                icon: Globe, color: 'bg-cyan-50 text-cyan-600',
                title: 'Peer Comparison',
                desc: 'Automatic comparable companies analysis with EV/EBITDA, P/E and revenue multiples.',
                examples: ['TCS vs Infosys vs Wipro', 'Apple vs Microsoft', 'HDFC vs Kotak']
              },
              {
                icon: Shield, color: 'bg-amber-50 text-amber-600',
                title: 'QA Verified Models',
                desc: 'Every Excel model passes automated audit — formula checks, chart validation, formatting standards.',
                examples: ['Zero #REF! errors', 'Institutional formatting', 'Audit trail included']
              },
            ].map(f => (
              <div key={f.title}
                className="p-6 rounded-2xl border border-gray-100 hover:border-blue-200
                hover:shadow-lg transition-all group bg-white">
                <div className={`w-11 h-11 rounded-xl ${f.color} flex items-center justify-center mb-4`}>
                  <f.icon size={20} />
                </div>
                <h3 className="font-bold text-[#1a3a6b] text-lg mb-2">{f.title}</h3>
                <p className="text-gray-500 text-sm leading-relaxed mb-4">{f.desc}</p>
                <div className="space-y-1.5">
                  {f.examples.map(ex => (
                    <div key={ex} className="flex items-center gap-2 text-xs text-gray-400">
                      <div className="w-1.5 h-1.5 rounded-full bg-blue-300"></div>
                      <span className="italic">"{ex}"</span>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* HOW IT WORKS */}
      <section id="how-it-works" className="py-20 px-6 bg-gradient-to-br from-[#f0f4ff] to-white">
        <div className="max-w-4xl mx-auto">
          <div className="text-center mb-14">
            <div className="text-[#2563eb] font-bold text-sm uppercase tracking-widest mb-3">How It Works</div>
            <h2 className="text-4xl font-extrabold text-[#1a3a6b]">Three Steps to Institutional Analysis</h2>
          </div>
          <div className="grid md:grid-cols-3 gap-8">
            {[
              { step: '01', title: 'Create Free Account', desc: 'Sign up in 30 seconds. No credit card. No trial period. Free forever.' },
              { step: '02', title: 'Type Your Question', desc: 'Ask anything — "Analyse Infosys", "Is TCS overvalued?", "Build DCF for Wipro"' },
              { step: '03', title: 'Get Model + Answer', desc: 'Receive a conversational answer plus a downloadable institutional Excel model.' },
            ].map(s => (
              <div key={s.step} className="text-center">
                <div className="w-16 h-16 rounded-2xl bg-[#1a3a6b] text-white font-extrabold text-2xl
                  flex items-center justify-center mx-auto mb-4 shadow-lg">
                  {s.step}
                </div>
                <h3 className="font-bold text-[#1a3a6b] text-lg mb-2">{s.title}</h3>
                <p className="text-gray-500 text-sm leading-relaxed">{s.desc}</p>
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* TESTIMONIALS */}
      <section className="py-20 px-6">
        <div className="max-w-4xl mx-auto">
          <div className="text-center mb-12">
            <h2 className="text-3xl font-extrabold text-[#1a3a6b]">Built for Students & Investors</h2>
          </div>
          <div className="grid md:grid-cols-3 gap-6">
            {[
              { name: 'Finance Student', role: 'MBA — IIM', text: 'I built a full DCF model for my valuation assignment in 5 minutes. The Excel output is better than what my seniors make manually.' },
              { name: 'Retail Investor', role: 'NSE Investor', text: 'Finally something that explains whether a stock is worth buying with actual numbers — not just opinions.' },
              { name: 'CFA Candidate', role: 'Level 2', text: 'The COMPS sheet and scenario analysis are exactly what you learn in CFA curriculum. This is a game-changer for practice.' },
            ].map(t => (
              <div key={t.name} className="p-5 rounded-2xl border border-gray-100 bg-white shadow-sm">
                <div className="flex gap-0.5 mb-3">
                  {[...Array(5)].map((_, i) => (
                    <Star key={i} size={13} className="text-amber-400 fill-amber-400" />
                  ))}
                </div>
                <p className="text-gray-600 text-sm leading-relaxed mb-4 italic">"{t.text}"</p>
                <div>
                  <div className="text-sm font-bold text-[#1a3a6b]">{t.name}</div>
                  <div className="text-xs text-gray-400">{t.role}</div>
                </div>
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* CTA SECTION */}
      <section className="py-20 px-6 bg-[#1a3a6b]">
        <div className="max-w-2xl mx-auto text-center">
          <h2 className="text-4xl font-extrabold text-white mb-4">
            Start Analysing Markets Today
          </h2>
          <p className="text-blue-200 mb-8 text-lg">
            Free forever. No credit card. Institutional-grade financial models in seconds.
          </p>
          <Link href="/auth/signup"
            className="inline-flex items-center gap-2 bg-white text-[#1a3a6b]
            font-bold px-10 py-4 rounded-2xl text-base hover:bg-blue-50 transition-all
            shadow-xl hover:-translate-y-0.5">
            Create Free Account
            <ChevronRight size={18} />
          </Link>
          <div className="flex items-center justify-center gap-6 mt-8">
            {['No credit card', 'Free forever', 'Cancel anytime'].map(item => (
              <div key={item} className="flex items-center gap-1.5 text-blue-300 text-sm">
                <Check size={14} className="text-green-400" />
                {item}
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* FOOTER */}
      <footer className="py-10 px-6 bg-[#0f2347] text-center">
        <div className="flex items-center justify-center gap-2 mb-4">
          <div className="w-6 h-6 rounded-md bg-[#2563eb] flex items-center justify-center">
            <BarChart2 size={12} className="text-white" />
          </div>
          <span className="text-white font-bold text-sm">Fintrust Global</span>
        </div>
        <p className="text-blue-400 text-xs">
          © 2026 Fintrust Global · AI-Powered Financial Analysis for Everyone
        </p>
        <p className="text-blue-600 text-xs mt-2">
          Not financial advice · For educational and research purposes only
        </p>
      </footer>
    </div>
  )
}

function ChatPreviewMsg({ role, text }) {
  return (
    <div className={`flex ${role === 'user' ? 'justify-end' : 'justify-start'} gap-2`}>
      {role === 'ai' && (
        <div className="w-6 h-6 rounded-full bg-[#1a3a6b] flex items-center justify-center shrink-0 mt-1">
          <BarChart2 size={10} className="text-white" />
        </div>
      )}
      <div className={`max-w-xs px-4 py-2.5 rounded-2xl text-xs leading-relaxed ${
        role === 'user'
          ? 'bg-[#1a3a6b] text-white rounded-br-sm'
          : 'bg-white text-gray-700 rounded-bl-sm border border-gray-100 shadow-sm'
      }`}>
        {text}
      </div>
    </div>
  )
}
