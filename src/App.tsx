import { useState, useEffect } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { Mail, Calendar, LogOut, CheckCircle2, AlertCircle, ExternalLink, ShieldCheck, Terminal } from 'lucide-react';

export default function App() {
  const [isAuthenticated, setIsAuthenticated] = useState<boolean | null>(null);
  const [loading, setLoading] = useState(true);

  const checkAuth = async () => {
    try {
      const res = await fetch('/api/auth/status');
      const data = await res.json();
      setIsAuthenticated(data.isAuthenticated);
    } catch (err) {
      console.error('Failed to check auth', err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    checkAuth();

    const handleMessage = (event: MessageEvent) => {
      if (event.data?.type === 'OAUTH_AUTH_SUCCESS') {
        checkAuth();
      }
    };
    window.addEventListener('message', handleMessage);
    return () => window.removeEventListener('message', handleMessage);
  }, []);

  const handleLogin = async () => {
    try {
      const res = await fetch('/api/auth/url');
      const { url } = await res.json();
      window.open(url, 'outlook_auth', 'width=600,height=700');
    } catch (err) {
      alert('Failed to get auth URL');
    }
  };

  const handleLogout = async () => {
    await fetch('/api/auth/logout', { method: 'POST' });
    setIsAuthenticated(false);
  };

  if (loading) {
    return (
      <div className="min-h-screen bg-[#0A0A0A] flex items-center justify-center">
        <div className="animate-spin rounded-full h-8 w-8 border-t-2 border-b-2 border-blue-500"></div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#0A0A0A] text-white font-sans selection:bg-blue-500/30">
      {/* Background Decor */}
      <div className="fixed inset-0 overflow-hidden pointer-events-none">
        <div className="absolute top-[-10%] left-[-10%] w-[40%] h-[40%] bg-blue-600/10 blur-[120px] rounded-full" />
        <div className="absolute bottom-[-10%] right-[-10%] w-[40%] h-[40%] bg-indigo-600/10 blur-[120px] rounded-full" />
      </div>

      <main className="relative z-10 max-w-4xl mx-auto px-6 py-20">
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="space-y-12"
        >
          {/* Header */}
          <div className="space-y-4">
            <div className="flex items-center gap-3 text-blue-400 mb-2">
              <ShieldCheck className="w-5 h-5" />
              <span className="text-xs font-mono uppercase tracking-widest">Secure MCP Bridge</span>
            </div>
            <h1 className="text-6xl md:text-7xl font-bold tracking-tighter leading-tight">
              Outlook <span className="text-blue-500">MCP</span> Server
            </h1>
            <p className="text-xl text-zinc-400 max-w-2xl leading-relaxed">
              Connect your AI agents to Microsoft Outlook. Read emails, manage your calendar, 
              and search contacts through the Model Context Protocol.
            </p>
          </div>

          {/* Connection Card */}
          <div className="grid md:grid-cols-2 gap-6">
            <motion.div 
              whileHover={{ scale: 1.01 }}
              className="p-8 rounded-3xl bg-zinc-900/50 border border-zinc-800 backdrop-blur-xl flex flex-col justify-between"
            >
              <div className="space-y-6">
                <div className="flex items-center justify-between">
                  <div className={`px-3 py-1 rounded-full text-[10px] font-mono uppercase tracking-wider ${isAuthenticated ? 'bg-emerald-500/10 text-emerald-400 border border-emerald-500/20' : 'bg-zinc-800 text-zinc-500 border border-zinc-700'}`}>
                    {isAuthenticated ? 'Connected' : 'Disconnected'}
                  </div>
                  {isAuthenticated ? <CheckCircle2 className="w-5 h-5 text-emerald-500" /> : <AlertCircle className="w-5 h-5 text-zinc-600" />}
                </div>
                
                <div className="space-y-2">
                  <h3 className="text-2xl font-semibold">Microsoft Account</h3>
                  <p className="text-zinc-500 text-sm">
                    {isAuthenticated 
                      ? "Your account is linked and ready for MCP requests." 
                      : "Authorize this server to access your Outlook data via Microsoft Graph."}
                  </p>
                </div>
              </div>

              <div className="mt-8">
                {!isAuthenticated ? (
                  <button 
                    onClick={handleLogin}
                    className="w-full py-4 bg-white text-black rounded-2xl font-semibold flex items-center justify-center gap-2 hover:bg-zinc-200 transition-colors group"
                  >
                    Connect Outlook
                    <ExternalLink className="w-4 h-4 group-hover:translate-x-0.5 group-hover:-translate-y-0.5 transition-transform" />
                  </button>
                ) : (
                  <button 
                    onClick={handleLogout}
                    className="w-full py-4 bg-zinc-800 text-white rounded-2xl font-semibold flex items-center justify-center gap-2 hover:bg-zinc-700 transition-colors"
                  >
                    <LogOut className="w-4 h-4" />
                    Disconnect
                  </button>
                )}
              </div>
            </motion.div>

            <div className="p-8 rounded-3xl bg-blue-600/5 border border-blue-500/10 backdrop-blur-xl space-y-6">
              <div className="flex items-center gap-3 text-blue-400">
                <Terminal className="w-5 h-5" />
                <h3 className="text-lg font-semibold">MCP Endpoint</h3>
              </div>
              <div className="space-y-4">
                <div className="p-4 rounded-xl bg-black/40 border border-white/5 font-mono text-xs text-zinc-400 break-all">
                  {window.location.origin}/mcp/sse
                </div>
                <p className="text-sm text-zinc-500">
                  Use this URL in your MCP client (like Claude Desktop) to connect to this server.
                </p>
                <div className="flex flex-wrap gap-2">
                  <span className="px-2 py-1 rounded-md bg-zinc-800 text-[10px] text-zinc-400 font-mono">Mail.ReadWrite</span>
                  <span className="px-2 py-1 rounded-md bg-zinc-800 text-[10px] text-zinc-400 font-mono">Calendars.ReadWrite</span>
                  <span className="px-2 py-1 rounded-md bg-zinc-800 text-[10px] text-zinc-400 font-mono">Contacts.Read</span>
                </div>
              </div>
            </div>
          </div>

          {/* Features Grid */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-8 pt-12 border-t border-zinc-800">
            <div className="space-y-4">
              <div className="w-10 h-10 rounded-full bg-blue-500/10 flex items-center justify-center text-blue-400">
                <Mail className="w-5 h-5" />
              </div>
              <h4 className="font-semibold">Email Intelligence</h4>
              <p className="text-sm text-zinc-500 leading-relaxed">
                List, read, and send emails. Perfect for summarizing threads or drafting replies with AI context.
              </p>
            </div>
            <div className="space-y-4">
              <div className="w-10 h-10 rounded-full bg-indigo-500/10 flex items-center justify-center text-indigo-400">
                <Calendar className="w-5 h-5" />
              </div>
              <h4 className="font-semibold">Calendar Sync</h4>
              <p className="text-sm text-zinc-500 leading-relaxed">
                Check availability and schedule meetings directly. Seamlessly bridge your AI assistant with your schedule.
              </p>
            </div>
            <div className="space-y-4">
              <div className="w-10 h-10 rounded-full bg-emerald-500/10 flex items-center justify-center text-emerald-400">
                <ShieldCheck className="w-5 h-5" />
              </div>
              <h4 className="font-semibold">Enterprise Security</h4>
              <p className="text-sm text-zinc-500 leading-relaxed">
                Built on Microsoft Graph API with OAuth 2.0. Your data stays between you and Microsoft.
              </p>
            </div>
          </div>
        </motion.div>
      </main>

      {/* Footer */}
      <footer className="max-w-4xl mx-auto px-6 py-12 border-t border-zinc-800/50 text-center text-zinc-600 text-xs">
        &copy; 2024 Outlook MCP Server &bull; Powered by Microsoft Graph
      </footer>
    </div>
  );
}
