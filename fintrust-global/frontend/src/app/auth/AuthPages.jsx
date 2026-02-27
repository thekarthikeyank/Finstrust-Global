'use client'
import { useState } from 'react'
import Link from 'next/link'
import { BarChart2, Eye, EyeOff, ArrowLeft } from 'lucide-react'

// ── SHARED AUTH LAYOUT ────────────────────────────────────
function AuthLayout({ children, title, subtitle }) {
  return (
    <div className="min-h-screen bg-gradient-to-br from-[#0f2347] via-[#1a3a6b] to-[#2563eb] flex items-center justify-center p-4">
      <div className="w-full max-w-md">
        {/* Logo */}
        <div className="text-center mb-8">
          <Link href="/" className="inline-flex items-center gap-2.5 group">
            <div className="w-10 h-10 rounded-xl bg-white/20 backdrop-blur flex items-center justify-center">
              <BarChart2 size={20} className="text-white" />
            </div>
            <span className="text-white font-bold text-xl tracking-tight">
              Fintrust <span className="text-blue-300">Global</span>
            </span>
          </Link>
          <h1 className="text-2xl font-bold text-white mt-6 mb-1">{title}</h1>
          <p className="text-blue-200 text-sm">{subtitle}</p>
        </div>

        {/* Card */}
        <div className="bg-white rounded-3xl shadow-2xl p-8">
          {children}
        </div>

        {/* Back */}
        <div className="text-center mt-6">
          <Link href="/" className="inline-flex items-center gap-1.5 text-blue-200 text-sm hover:text-white transition-colors">
            <ArrowLeft size={14} />
            Back to home
          </Link>
        </div>
      </div>
    </div>
  )
}

// ── INPUT COMPONENT ───────────────────────────────────────
function AuthInput({ label, type = 'text', value, onChange, placeholder, show, onToggle }) {
  return (
    <div>
      <label className="block text-sm font-semibold text-gray-700 mb-1.5">{label}</label>
      <div className="relative">
        <input
          type={type === 'password' ? (show ? 'text' : 'password') : type}
          value={value}
          onChange={onChange}
          placeholder={placeholder}
          className="w-full px-4 py-3 rounded-xl border border-gray-200 text-sm
            text-gray-800 placeholder-gray-400
            focus:outline-none focus:ring-2 focus:ring-[#1a3a6b] focus:border-transparent
            transition-all"
          required
        />
        {type === 'password' && (
          <button type="button" onClick={onToggle}
            className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600">
            {show ? <EyeOff size={16} /> : <Eye size={16} />}
          </button>
        )}
      </div>
    </div>
  )
}

// ── LOGIN PAGE ────────────────────────────────────────────
export function LoginPage() {
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [showPass, setShowPass] = useState(false)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')

  const handleLogin = async (e) => {
    e.preventDefault()
    setLoading(true)
    setError('')
    try {
      const res = await fetch('/api/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email, password })
      })
      const data = await res.json()
      if (data.error) throw new Error(data.error)
      window.location.href = '/chat'
    } catch (err) {
      setError(err.message || 'Login failed. Please try again.')
    } finally {
      setLoading(false)
    }
  }

  return (
    <AuthLayout title="Welcome Back" subtitle="Sign in to your Fintrust Global account">
      <form onSubmit={handleLogin} className="space-y-5">
        <AuthInput label="Email Address" type="email" value={email}
          onChange={e => setEmail(e.target.value)} placeholder="you@example.com" />
        <AuthInput label="Password" type="password" value={password}
          onChange={e => setPassword(e.target.value)} placeholder="Your password"
          show={showPass} onToggle={() => setShowPass(!showPass)} />

        {error && (
          <div className="bg-red-50 border border-red-200 text-red-600 text-sm px-4 py-3 rounded-xl">
            {error}
          </div>
        )}

        <button type="submit" disabled={loading}
          className="w-full bg-[#1a3a6b] hover:bg-[#2563eb] disabled:bg-gray-300
          text-white font-bold py-3.5 rounded-xl text-sm transition-all
          shadow-lg hover:shadow-blue-200">
          {loading ? (
            <span className="flex items-center justify-center gap-2">
              <svg className="animate-spin h-4 w-4" viewBox="0 0 24 24" fill="none">
                <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"/>
                <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"/>
              </svg>
              Signing in...
            </span>
          ) : 'Sign In'}
        </button>

        <div className="text-center">
          <span className="text-sm text-gray-500">Don't have an account? </span>
          <Link href="/auth/signup" className="text-sm font-semibold text-[#1a3a6b] hover:text-[#2563eb]">
            Create one free
          </Link>
        </div>
      </form>
    </AuthLayout>
  )
}

// ── SIGNUP PAGE ───────────────────────────────────────────
export function SignupPage() {
  const [form, setForm] = useState({ name: '', email: '', password: '' })
  const [showPass, setShowPass] = useState(false)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')
  const [success, setSuccess] = useState(false)

  const update = (field) => (e) => setForm(prev => ({ ...prev, [field]: e.target.value }))

  const handleSignup = async (e) => {
    e.preventDefault()
    if (form.password.length < 8) {
      setError('Password must be at least 8 characters')
      return
    }
    setLoading(true)
    setError('')
    try {
      const res = await fetch('/api/auth/signup', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(form)
      })
      const data = await res.json()
      if (data.error) throw new Error(data.error)
      setSuccess(true)
    } catch (err) {
      setError(err.message || 'Signup failed. Please try again.')
    } finally {
      setLoading(false)
    }
  }

  if (success) {
    return (
      <AuthLayout title="Check Your Email" subtitle="We sent you a confirmation link">
        <div className="text-center py-4">
          <div className="w-16 h-16 bg-green-100 rounded-full flex items-center justify-center mx-auto mb-4">
            <span className="text-3xl">✉️</span>
          </div>
          <p className="text-gray-600 text-sm leading-relaxed mb-6">
            We sent a confirmation link to <strong>{form.email}</strong>.
            Click the link to activate your account.
          </p>
          <Link href="/auth/login"
            className="inline-block bg-[#1a3a6b] text-white font-semibold
            px-6 py-3 rounded-xl text-sm hover:bg-[#2563eb] transition-colors">
            Go to Login
          </Link>
        </div>
      </AuthLayout>
    )
  }

  return (
    <AuthLayout title="Create Free Account" subtitle="Start analysing markets in 30 seconds">
      <form onSubmit={handleSignup} className="space-y-5">
        <AuthInput label="Full Name" value={form.name}
          onChange={update('name')} placeholder="Your full name" />
        <AuthInput label="Email Address" type="email" value={form.email}
          onChange={update('email')} placeholder="you@example.com" />
        <AuthInput label="Password" type="password" value={form.password}
          onChange={update('password')} placeholder="Min 8 characters"
          show={showPass} onToggle={() => setShowPass(!showPass)} />

        {error && (
          <div className="bg-red-50 border border-red-200 text-red-600 text-sm px-4 py-3 rounded-xl">
            {error}
          </div>
        )}

        {/* What you get */}
        <div className="bg-blue-50 rounded-xl p-3 space-y-1.5">
          {['Free DCF, LBO, 3-Statement & FP&A models',
            'Real NSE, BSE, NYSE & NASDAQ data',
            'Unlimited finance questions',
            'No credit card ever required'].map(item => (
            <div key={item} className="flex items-center gap-2 text-xs text-blue-700">
              <div className="w-4 h-4 rounded-full bg-blue-600 flex items-center justify-center shrink-0">
                <span className="text-white text-[8px] font-bold">✓</span>
              </div>
              {item}
            </div>
          ))}
        </div>

        <button type="submit" disabled={loading}
          className="w-full bg-[#1a3a6b] hover:bg-[#2563eb] disabled:bg-gray-300
          text-white font-bold py-3.5 rounded-xl text-sm transition-all shadow-lg">
          {loading ? 'Creating Account...' : 'Create Free Account →'}
        </button>

        <p className="text-center text-xs text-gray-400">
          By signing up you agree to our Terms of Service
        </p>

        <div className="text-center">
          <span className="text-sm text-gray-500">Already have an account? </span>
          <Link href="/auth/login" className="text-sm font-semibold text-[#1a3a6b] hover:text-[#2563eb]">
            Sign in
          </Link>
        </div>
      </form>
    </AuthLayout>
  )
}

export default LoginPage
