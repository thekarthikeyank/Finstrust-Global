'use client'
import { useState, useRef, useEffect } from 'react'
import Link from 'next/link'
import {
  BarChart2, Send, Plus, FileSpreadsheet, TrendingUp,
  MessageSquare, LogOut, Download, ChevronDown, Sparkles
} from 'lucide-react'

const API = process.env.NEXT_PUBLIC_API_URL || 'http://localhost:8000'

const SUGGESTED_PROMPTS = [
  'Analyse Infosys and build a DCF model',
  'Is TCS overvalued right now?',
  'Compare HDFC Bank vs ICICI Bank',
  'Build an LBO model for Zomato',
  'Explain EBITDA in simple terms',
  'What is the fair value of Reliance?',
  'Best IT stocks by ROCE on NSE',
  'Apple vs Microsoft valuation comparison',
]

export default function ChatPage() {
  const [messages, setMessages] = useState([])
  const [input, setInput] = useState('')
  const [loading, setLoading] = useState(false)
  const [sessionId, setSessionId] = useState(null)
  const [logs, setLogs] = useState([])
  const [showLogs, setShowLogs] = useState(false)
  const [downloadUrl, setDownloadUrl] = useState(null)
  const [modelReady, setModelReady] = useState(false)
  const [history, setHistory] = useState([])
  const bottomRef = useRef(null)
  const inputRef = useRef(null)
  const esRef = useRef(null)

  useEffect(() => {
    initSession()
    inputRef.current?.focus()
    return () => esRef.current?.close()
  }, [])

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: 'smooth' })
  }, [messages, logs])

  const initSession = async () => {
    try {
      const res = await fetch(`${API}/api/session/new`, { method: 'POST' })
      const data = await res.json()
      setSessionId(data.session_id)
    } catch { }
  }

  const addMessage = (role, content, extra = {}) => {
    setMessages(prev => [...prev, { role, content, id: Date.now(), ...extra }])
  }

  const handleSend = async (text = input) => {
    if (!text.trim() || loading) return
    const question = text.trim()
    setInput('')
    setLoading(true)
    setLogs([])
    setDownloadUrl(null)
    setModelReady(false)

    addMessage('user', question)
    setHistory(prev => [{ q: question.slice(0, 40), date: new Date().toLocaleDateString() }, ...prev.slice(0, 9)])

    try {
      // Detect if model generation is requested
      const isModelRequest = /dcf|lbo|3.statement|fpa|fp&a|model|analyse|analyze|valuat/i.test(question)

      if (isModelRequest) {
        addMessage('assistant', '', { type: 'thinking', sid: sessionId })
        setShowLogs(true)
        await handleModelGeneration(question)
      } else {
        await handleFinanceChat(question)
      }
    } catch (err) {
      addMessage('assistant', `I encountered an error: ${err.message}. Please try again.`)
    } finally {
      setLoading(false)
    }
  }

  const handleFinanceChat = async (question) => {
    const res = await fetch(`${API}/api/chat`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ question, session_id: sessionId })
    })
    const data = await res.json()
    setMessages(prev => prev.filter(m => m.type !== 'thinking'))
    addMessage('assistant', data.answer || 'I could not process that question. Please try rephrasing.', {
      sources: data.sources,
      hasModel: data.suggested_model
    })
  }

  const handleModelGeneration = async (question) => {
    // Extract company name from question
    const res = await fetch(`${API}/api/research`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ company_name: question, session_id: sessionId })
    })
    const data = await res.json()
    const sid = data.session_id || sessionId
    setSessionId(sid)

    // Stream logs
    if (esRef.current) esRef.current.close()
    const es = new EventSource(`${API}/api/logs/${sid}`)
    esRef.current = es

    es.onmessage = (e) => {
      const log = JSON.parse(e.data)
      if (log.type === 'done') {
        es.close()
        checkFinalStatus(sid)
        return
      }
      setLogs(prev => [...prev, log])
    }
    es.onerror = () => { es.close(); checkFinalStatus(sid) }
  }

  const checkFinalStatus = async (sid) => {
    try {
      const res = await fetch(`${API}/api/status/${sid}`)
      const data = await res.json()
      setMessages(prev => prev.filter(m => m.type !== 'thinking'))

      if (data.phase === 'awaiting_confirmation') {
        addMessage('assistant',
          `I've analysed **${data.company_data?.company_name || 'the company'}** and recommend a **${data.model_recommendation} model**.\n\n${data.model_recommendation === 'DCF' ? '📊 Strong cash flows and stable margins — ideal for DCF valuation.' : '🏦 Leverage characteristics suit LBO analysis.'}\n\nShall I build the model?`,
          { type: 'confirm', sid, modelType: data.model_recommendation }
        )
      } else if (data.phase === 'delivered') {
        setDownloadUrl(`${API}/api/download/${sid}`)
        setModelReady(true)
        const company = data.company_data?.company_name || 'Company'
        const model = data.model_recommendation
        addMessage('assistant',
          `✅ Your **${company} ${model} Model** is ready. All QA checks passed — ${data.qa_report?.checks_passed || 0} formula and chart validations confirmed.`,
          { type: 'download', downloadUrl: `${API}/api/download/${sid}`, company, model }
        )
      }
    } catch { }
  }

  const confirmBuild = async (sid, modelType) => {
    setLoading(true)
    setLogs([])
    addMessage('user', `Build the ${modelType} model`)

    await fetch(`${API}/api/confirm-model`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ session_id: sid, confirmed: true, model_type: modelType })
    })

    const es = new EventSource(`${API}/api/logs/${sid}`)
    esRef.current = es
    es.onmessage = (e) => {
      const log = JSON.parse(e.data)
      if (log.type === 'done') { es.close(); checkFinalStatus(sid); setLoading(false); return }
      setLogs(prev => [...prev, log])
    }
    es.onerror = () => { es.close(); checkFinalStatus(sid); setLoading(false) }
  }

  const newChat = () => {
    setMessages([])
    setLogs([])
    setInput('')
    setDownloadUrl(null)
    setModelReady(false)
    setShowLogs(false)
    initSession()
    inputRef.current?.focus()
  }

  return (
    <div className="flex h-screen bg-white overflow-hidden font-sans">

      {/* SIDEBAR */}
      <div className="w-64 shrink-0 flex flex-col bg-[#0f1f3d] border-r border-white/10">
        {/* Logo */}
        <div className="p-4 border-b border-white/10">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-[#2563eb] flex items-center justify-center">
              <BarChart2 size={16} className="text-white" />
            </div>
            <div>
              <div className="text-sm font-bold text-white">Fintrust Global</div>
              <div className="text-[10px] text-blue-400">AI Finance Assistant</div>
            </div>
          </div>
        </div>

        {/* New Chat */}
        <div className="p-3">
          <button onClick={newChat}
            className="w-full flex items-center gap-2 px-3 py-2.5 rounded-xl
            bg-[#2563eb] hover:bg-blue-500 text-white text-sm font-medium transition-colors">
            <Plus size={15} />
            New Analysis
          </button>
        </div>

        {/* History */}
        <div className="flex-1 overflow-y-auto px-3 py-2">
          {history.length > 0 && (
            <>
              <div className="text-[10px] text-gray-500 uppercase font-semibold tracking-wider px-1 mb-2">
                Recent
              </div>
              {history.map((h, i) => (
                <div key={i} className="px-3 py-2 rounded-lg hover:bg-white/5 cursor-pointer mb-1">
                  <div className="text-xs text-gray-300 truncate">{h.q}...</div>
                  <div className="text-[10px] text-gray-600 mt-0.5">{h.date}</div>
                </div>
              ))}
            </>
          )}
        </div>

        {/* Model types */}
        <div className="p-3 border-t border-white/10">
          <div className="text-[10px] text-gray-500 uppercase font-semibold tracking-wider mb-2">Models</div>
          {[['DCF', 'bg-blue-500'], ['LBO', 'bg-purple-500'], ['3-Statement', 'bg-green-500'], ["FP&A", 'bg-orange-400']].map(([m, c]) => (
            <div key={m} className="flex items-center gap-2 py-1">
              <div className={`w-1.5 h-1.5 rounded-full ${c}`}></div>
              <span className="text-xs text-gray-400">{m}</span>
            </div>
          ))}
        </div>

        {/* User */}
        <div className="p-3 border-t border-white/10 flex items-center gap-2">
          <div className="w-7 h-7 rounded-full bg-[#2563eb] flex items-center justify-center text-white text-xs font-bold">
            U
          </div>
          <div className="flex-1 min-w-0">
            <div className="text-xs text-white font-medium truncate">My Account</div>
            <div className="text-[10px] text-gray-500">Free Plan</div>
          </div>
          <Link href="/" className="text-gray-500 hover:text-gray-300">
            <LogOut size={13} />
          </Link>
        </div>
      </div>

      {/* MAIN */}
      <div className="flex-1 flex flex-col min-w-0">

        {/* Top bar */}
        <div className="h-14 border-b border-gray-100 flex items-center px-6 bg-white shrink-0">
          <span className="text-sm font-semibold text-gray-800">Finance Chat</span>
          <div className="ml-auto flex items-center gap-3">
            <div className="flex items-center gap-1.5 text-xs text-gray-400 bg-gray-50 px-3 py-1.5 rounded-full border border-gray-200">
              <Sparkles size={11} className="text-blue-500" />
              Gemini 1.5 Flash
            </div>
          </div>
        </div>

        {/* Messages */}
        <div className="flex-1 overflow-y-auto bg-white">
          {messages.length === 0 ? (
            <EmptyState onPrompt={handleSend} />
          ) : (
            <div className="max-w-3xl mx-auto px-4 py-6 space-y-6">
              {messages.map(msg => (
                <MessageItem key={msg.id} msg={msg} onConfirm={confirmBuild} />
              ))}

              {/* Live Agent Logs */}
              {logs.length > 0 && showLogs && (
                <div className="bg-gradient-to-b from-blue-50 to-white rounded-2xl border border-blue-100 overflow-hidden">
                  <button
                    onClick={() => setShowLogs(s => !s)}
                    className="w-full px-4 py-3 flex items-center justify-between text-xs font-semibold text-blue-700">
                    <span className="flex items-center gap-2">
                      <div className="flex gap-1">
                        {loading && <>
                          <span className="w-1.5 h-1.5 bg-blue-500 rounded-full animate-bounce" style={{animationDelay:'0ms'}}></span>
                          <span className="w-1.5 h-1.5 bg-blue-500 rounded-full animate-bounce" style={{animationDelay:'150ms'}}></span>
                          <span className="w-1.5 h-1.5 bg-blue-500 rounded-full animate-bounce" style={{animationDelay:'300ms'}}></span>
                        </>}
                      </div>
                      Agent Pipeline — {logs.length} events
                    </span>
                    <ChevronDown size={14} />
                  </button>
                  <div className="px-4 pb-4 space-y-1 max-h-48 overflow-y-auto">
                    {logs.map((log, i) => (
                      <div key={i} className="flex items-start gap-2 text-xs py-1">
                        <div className={`w-1.5 h-1.5 rounded-full mt-1.5 shrink-0 ${
                          log.status === 'success' ? 'bg-green-400' :
                          log.status === 'warning' ? 'bg-amber-400' :
                          log.status === 'error' ? 'bg-red-400' :
                          log.status === 'thinking' ? 'bg-blue-400 animate-pulse' : 'bg-gray-300'
                        }`}></div>
                        <span className={`${
                          log.status === 'success' ? 'text-green-700' :
                          log.status === 'warning' ? 'text-amber-700' :
                          log.status === 'error' ? 'text-red-700' :
                          log.status === 'thinking' ? 'text-blue-600' : 'text-gray-600'
                        }`}>
                          <span className="font-semibold">{log.agent}:</span> {log.message}
                        </span>
                      </div>
                    ))}
                  </div>
                </div>
              )}
              <div ref={bottomRef} />
            </div>
          )}
        </div>

        {/* Input */}
        <div className="border-t border-gray-100 bg-white p-4 shrink-0">
          <div className="max-w-3xl mx-auto">
            <div className="relative flex items-end gap-2">
              <div className="flex-1 relative">
                <textarea
                  ref={inputRef}
                  value={input}
                  onChange={e => setInput(e.target.value)}
                  onKeyDown={e => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); handleSend() } }}
                  placeholder="Ask anything — 'Analyse Infosys', 'Is TCS overvalued?', 'Build DCF for Wipro'..."
                  rows={1}
                  disabled={loading}
                  className="w-full px-5 py-4 pr-14 rounded-2xl border border-gray-200
                    text-sm text-gray-800 placeholder-gray-400 resize-none
                    focus:outline-none focus:ring-2 focus:ring-[#1a3a6b] focus:border-transparent
                    disabled:bg-gray-50 disabled:cursor-not-allowed
                    shadow-sm transition-all leading-relaxed"
                  style={{ minHeight: '56px', maxHeight: '160px' }}
                />
                <button
                  onClick={() => handleSend()}
                  disabled={!input.trim() || loading}
                  className="absolute right-3 bottom-3 w-9 h-9 rounded-xl
                    bg-[#1a3a6b] hover:bg-[#2563eb]
                    disabled:bg-gray-200 disabled:cursor-not-allowed
                    flex items-center justify-center transition-colors shadow-sm">
                  <Send size={15} className={input.trim() && !loading ? 'text-white' : 'text-gray-400'} />
                </button>
              </div>
            </div>
            <p className="text-center text-[10px] text-gray-400 mt-2">
              Fintrust Global · Not financial advice · For educational purposes
            </p>
          </div>
        </div>
      </div>
    </div>
  )
}

// ── EMPTY STATE ───────────────────────────────────────────
function EmptyState({ onPrompt }) {
  return (
    <div className="flex flex-col items-center justify-center h-full py-16 px-4">
      <div className="w-14 h-14 rounded-2xl bg-gradient-to-br from-[#1a3a6b] to-[#2563eb] flex items-center justify-center shadow-lg mb-5">
        <BarChart2 size={24} className="text-white" />
      </div>
      <h2 className="text-xl font-bold text-gray-900 mb-2">How can I help you today?</h2>
      <p className="text-sm text-gray-500 mb-8 text-center max-w-md">
        Ask any finance question or request a model for any publicly listed company
      </p>
      <div className="grid grid-cols-2 gap-2 max-w-xl w-full">
        {SUGGESTED_PROMPTS.map(p => (
          <button key={p} onClick={() => onPrompt(p)}
            className="text-left px-4 py-3 rounded-xl border border-gray-200
            hover:border-blue-300 hover:bg-blue-50 text-xs text-gray-600
            transition-all group">
            <span className="group-hover:text-blue-700">{p}</span>
          </button>
        ))}
      </div>
    </div>
  )
}

// ── MESSAGE ITEM ─────────────────────────────────────────
function MessageItem({ msg, onConfirm }) {
  const isUser = msg.role === 'user'

  if (isUser) {
    return (
      <div className="flex justify-end">
        <div className="max-w-lg bg-[#1a3a6b] text-white px-5 py-3.5 rounded-2xl rounded-br-sm text-sm leading-relaxed shadow-sm">
          {msg.content}
        </div>
      </div>
    )
  }

  return (
    <div className="flex gap-3">
      <div className="w-8 h-8 rounded-full bg-gradient-to-br from-[#1a3a6b] to-[#2563eb] flex items-center justify-center shrink-0 mt-1 shadow-sm">
        <BarChart2 size={14} className="text-white" />
      </div>
      <div className="flex-1 min-w-0">
        {msg.type === 'thinking' && (
          <div className="flex items-center gap-2 py-3">
            <div className="flex gap-1">
              {[0, 150, 300].map(d => (
                <span key={d} className="w-2 h-2 bg-blue-400 rounded-full animate-bounce" style={{animationDelay:`${d}ms`}}></span>
              ))}
            </div>
            <span className="text-sm text-gray-400">Analysing...</span>
          </div>
        )}

        {msg.content && (
          <div className="bg-gray-50 border border-gray-100 rounded-2xl rounded-tl-sm px-5 py-4 text-sm text-gray-800 leading-relaxed">
            {formatContent(msg.content)}
          </div>
        )}

        {msg.type === 'confirm' && (
          <div className="mt-3 flex gap-2">
            <button onClick={() => onConfirm(msg.sid, msg.modelType)}
              className="flex items-center gap-2 bg-[#1a3a6b] hover:bg-[#2563eb] text-white px-5 py-2.5 rounded-xl text-sm font-semibold transition-colors shadow-sm">
              <FileSpreadsheet size={15} />
              Build {msg.modelType} Model
            </button>
            <button className="border border-gray-200 hover:bg-gray-50 text-gray-600 px-4 py-2.5 rounded-xl text-sm transition-colors">
              Choose Different Model
            </button>
          </div>
        )}

        {msg.type === 'download' && msg.downloadUrl && (
          <div className="mt-3">
            <a href={msg.downloadUrl} download
              className="inline-flex items-center gap-2 bg-green-600 hover:bg-green-700
              text-white px-5 py-2.5 rounded-xl text-sm font-semibold transition-colors shadow-sm">
              <Download size={15} />
              Download {msg.company} {msg.model} Model.xlsx
            </a>
          </div>
        )}
      </div>
    </div>
  )
}

function formatContent(text) {
  const parts = text.split(/(\*\*[^*]+\*\*)/g)
  return parts.map((part, i) =>
    part.startsWith('**') && part.endsWith('**')
      ? <strong key={i} className="text-[#1a3a6b]">{part.slice(2,-2)}</strong>
      : <span key={i}>{part}</span>
  )
}
