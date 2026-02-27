import './globals.css'
export const metadata = { title: 'Fintrust Global', description: 'AI-Powered Financial Analysis for Everyone' }
export default function RootLayout({ children }) {
  return <html lang="en"><body className="antialiased">{children}</body></html>
}
