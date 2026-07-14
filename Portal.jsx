import React, { useState, useEffect, useMemo } from 'react';
import { Truck, Loader2, Check, LogOut, Package, ClipboardList } from 'lucide-react';

function fieldStyle() {
  return { width: '100%', padding: '9px 10px', borderRadius: 7, border: '1px solid #DCD9CE', fontSize: 13.5, marginBottom: 12 };
}
function labelStyle() {
  return { fontSize: 11.5, fontWeight: 700, color: '#5B6470', textTransform: 'uppercase', letterSpacing: '0.03em', display: 'block', marginBottom: 5 };
}
function th() { return { padding: '7px 10px', textAlign: 'left', fontSize: 10.5, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.03em' }; }
function td() { return { padding: '8px 10px', fontSize: 13, verticalAlign: 'middle' }; }

async function api(action, body) {
  const res = await fetch('/api/dealer-auth', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ action, ...body }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || 'Something went wrong.');
  return data;
}

export default function Portal() {
  const [token, setToken] = useState(() => localStorage.getItem('gr-dealer-token') || '');
  const [dealerName, setDealerName] = useState('');
  const [checking, setChecking] = useState(true);

  useEffect(() => {
    (async () => {
      if (!token) { setChecking(false); return; }
      try {
        const data = await api('session', { token });
        setDealerName(data.dealerName);
      } catch {
        localStorage.removeItem('gr-dealer-token');
        setToken('');
      }
      setChecking(false);
    })();
  }, [token]);

  function handleLogin(tok, name) {
    localStorage.setItem('gr-dealer-token', tok);
    setToken(tok);
    setDealerName(name);
  }

  function handleLogout() {
    api('logout', { token }).catch(() => {});
    localStorage.removeItem('gr-dealer-token');
    setToken('');
    setDealerName('');
  }

  return (
    <div style={{
      fontFamily: "'Inter', sans-serif", background: '#EEEEE9', minHeight: '100vh', color: '#1C2126'
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Oswald:wght@600;700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@600&display=swap');
        * { box-sizing: border-box; }
        table { border-collapse: collapse; width: 100%; }
        button { cursor: pointer; font-family: 'Inter', sans-serif; }
      `}</style>
      {checking ? (
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '100vh', color: '#5B6470', gap: 10 }}>
          <Loader2 size={18} className="spin" /> Loading…
          <style>{`.spin { animation: spin 1s linear infinite; } @keyframes spin { to { transform: rotate(360deg); } }`}</style>
        </div>
      ) : !token || !dealerName ? (
        <LoginScreen onLogin={handleLogin} />
      ) : (
        <PortalHome dealerName={dealerName} token={token} onLogout={handleLogout} />
      )}
    </div>
  );
}

function LoginScreen({ onLogin }) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);

  async function submit(e) {
    e.preventDefault();
    setError('');
    setLoading(true);
    try {
      const data = await api('login', { username, password });
      onLogin(data.token, data.dealerName);
    } catch (err) {
      setError(err.message);
    }
    setLoading(false);
  }

  return (
    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', minHeight: '100vh', padding: 16 }}>
      <form onSubmit={submit} style={{ background: 'white', border: '1px solid #DCD9CE', borderRadius: 12, padding: 28, width: '100%', maxWidth: 340 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 20 }}>
          <div style={{ width: 32, height: 32, borderRadius: 7, background: '#1C2126', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
            <Truck size={17} color="#E8592A" />
          </div>
          <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 700, fontSize: 18, textTransform: 'uppercase' }}>Dealer Portal</div>
        </div>
        <label style={labelStyle()}>Username</label>
        <input value={username} onChange={e => setUsername(e.target.value)} style={fieldStyle()} autoFocus />
        <label style={labelStyle()}>Password</label>
        <input type="password" value={password} onChange={e => setPassword(e.target.value)} style={fieldStyle()} />
        {error && <div style={{ color: '#B23A2E', fontSize: 12.5, marginBottom: 12 }}>{error}</div>}
        <button disabled={loading} type="submit" style={{
          width: '100%', background: '#E8592A', color: 'white', border: 'none', borderRadius: 8,
          padding: '11px', fontSize: 13.5, fontWeight: 700
        }}>{loading ? 'Signing in…' : 'Sign in'}</button>
      </form>
    </div>
  );
}

function PortalHome({ dealerName, token, onLogout }) {
  const [orders, setOrders] = useState([]);
  const [models, setModels] = useState([]);
  const [loading, setLoading] = useState(true);
  const [lines, setLines] = useState([{ sku: '', qty: '' }]);
  const [po, setPo] = useState('');
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState('');
  const [toast, setToast] = useState('');

  function updateLine(i, updates) {
    setLines(ls => ls.map((l, idx) => idx === i ? { ...l, ...updates } : l));
  }
  function addLine() {
    setLines(ls => [...ls, { sku: '', qty: '' }]);
  }
  function removeLine(i) {
    setLines(ls => ls.filter((_, idx) => idx !== i));
  }

  async function loadOrders() {
    const res = await fetch(`/api/dealer-orders?token=${encodeURIComponent(token)}`);
    const data = await res.json();
    if (res.ok) setOrders(data.orders);
  }

  useEffect(() => {
    (async () => {
      await Promise.all([
        loadOrders(),
        fetch('/api/dealer-models').then(r => r.json()).then(d => setModels(d.skus || [])),
      ]);
      setLoading(false);
    })();
  }, []);

  // Quietly re-check for updates (e.g. staff confirming or shipping an order)
  // every few seconds, so the dealer doesn't have to refresh the page.
  useEffect(() => {
    const interval = setInterval(loadOrders, 10000);
    return () => clearInterval(interval);
  }, []);

  useEffect(() => {
    if (!toast) return;
    const t = setTimeout(() => setToast(''), 2500);
    return () => clearTimeout(t);
  }, [toast]);

  async function submitOrder(e) {
    e.preventDefault();
    setError('');
    const validLines = lines.filter(l => l.sku && l.qty);
    if (validLines.length === 0) { setError('Pick at least one model and enter a quantity.'); return; }
    setSubmitting(true);
    try {
      const res = await fetch('/api/dealer-orders', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ token, po, items: validLines }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Could not place order.');
      setLines([{ sku: '', qty: '' }]); setPo('');
      setToast(validLines.length > 1 ? `Order request sent — ${validLines.length} models, pending confirmation` : 'Order request sent — pending confirmation');
      await loadOrders();
    } catch (err) {
      setError(err.message);
    }
    setSubmitting(false);
  }

  const sorted = useMemo(() => orders.slice().sort((a, b) => {
    const da = new Date(a.date).getTime() || 0, db = new Date(b.date).getTime() || 0;
    return db - da;
  }), [orders]);

  return (
    <div style={{ maxWidth: 720, margin: '0 auto', padding: '20px 18px 60px' }}>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 24 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <div style={{ width: 32, height: 32, borderRadius: 7, background: '#1C2126', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
            <Truck size={17} color="#E8592A" />
          </div>
          <div>
            <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 700, fontSize: 17, textTransform: 'uppercase', lineHeight: 1.1 }}>{dealerName}</div>
            <div style={{ fontSize: 11, color: '#5B6470' }}>Dealer Portal</div>
          </div>
        </div>
        <button onClick={onLogout} style={{
          display: 'flex', alignItems: 'center', gap: 6, background: 'white', border: '1px solid #DCD9CE',
          borderRadius: 7, padding: '7px 12px', fontSize: 12.5, fontWeight: 600, color: '#5B6470'
        }}><LogOut size={13} /> Log out</button>
      </div>

      <div style={{ background: 'white', border: '1px solid #DCD9CE', borderRadius: 12, padding: 20, marginBottom: 20 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4 }}>
          <Package size={15} color="#33546E" />
          <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 600, fontSize: 14, textTransform: 'uppercase' }}>Place an order</div>
        </div>
        <div style={{ fontSize: 11.5, color: '#8A8F97', marginBottom: 14 }}>Your order request will show as "Pending confirmation" until we confirm it on our end.</div>
        <form onSubmit={submitOrder}>
          <label style={labelStyle()}>PO # (optional, applies to the whole order)</label>
          <input value={po} onChange={e => setPo(e.target.value)} style={fieldStyle()} />

          <label style={labelStyle()}>Models</label>
          {lines.map((l, i) => (
            <div key={i} style={{ display: 'flex', gap: 8, marginBottom: 8, alignItems: 'center' }}>
              <select value={l.sku} onChange={e => updateLine(i, { sku: e.target.value })} style={{ ...fieldStyle(), marginBottom: 0, flex: 2 }}>
                <option value="">Select a model…</option>
                {models.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
              <input type="number" min="1" placeholder="Qty" value={l.qty} onChange={e => updateLine(i, { qty: e.target.value })} style={{ ...fieldStyle(), marginBottom: 0, width: 80, flex: 'none' }} />
              {lines.length > 1 && (
                <button type="button" onClick={() => removeLine(i)} style={{
                  background: '#FCEEE8', border: '1px solid #F0C4B8', color: '#B23A2E', borderRadius: 6,
                  padding: '9px 10px', display: 'flex', alignItems: 'center', flex: 'none'
                }}>✕</button>
              )}
            </div>
          ))}
          <button type="button" onClick={addLine} style={{
            display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6, background: '#FAFAF7',
            border: '1px dashed #C9C5B8', color: '#5B6470', borderRadius: 7, padding: '8px 10px',
            fontSize: 12.5, fontWeight: 600, marginBottom: 14, width: '100%'
          }}>+ Add another model</button>

          {error && <div style={{ color: '#B23A2E', fontSize: 12.5, marginBottom: 10 }}>{error}</div>}
          <button disabled={submitting} type="submit" style={{
            background: '#E8592A', color: 'white', border: 'none', borderRadius: 8,
            padding: '10px 18px', fontSize: 13, fontWeight: 700
          }}>{submitting ? 'Placing order…' : 'Place order'}</button>
        </form>
      </div>

      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10 }}>
        <ClipboardList size={15} color="#33546E" />
        <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 600, fontSize: 14, textTransform: 'uppercase' }}>My orders</div>
      </div>

      {loading ? (
        <div style={{ color: '#8A8F97', fontSize: 13 }}>Loading…</div>
      ) : sorted.length === 0 ? (
        <div style={{ background: 'white', border: '1px solid #DCD9CE', borderRadius: 12, padding: 20, color: '#8A8F97', fontSize: 13 }}>
          No orders yet — place one above.
        </div>
      ) : (
        <div style={{ background: 'white', border: '1px solid #DCD9CE', borderRadius: 12, overflow: 'hidden' }}>
          <table>
            <thead>
              <tr style={{ background: '#1C2126', color: '#F5F3EE' }}>
                <th style={th()}>Model</th>
                <th style={th()}>Date</th>
                <th style={th()}>PO</th>
                <th style={{ ...th(), textAlign: 'right' }}>Qty</th>
                <th style={{ ...th(), textAlign: 'right' }}>Owed</th>
                <th style={th()}>Status</th>
              </tr>
            </thead>
            <tbody>
              {sorted.map((o, i) => (
                <tr key={o.id} style={{ borderTop: i ? '1px solid #EFEDE4' : 'none' }}>
                  <td style={{ ...td(), fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600 }}>{o.sku}</td>
                  <td style={{ ...td(), color: '#5B6470' }}>{o.date || '—'}</td>
                  <td style={{ ...td(), color: '#8A8F97' }}>{o.po || '—'}</td>
                  <td style={{ ...td(), textAlign: 'right', fontFamily: "'IBM Plex Mono', monospace" }}>{o.qty}</td>
                  <td style={{ ...td(), textAlign: 'right', fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, color: o.backordered > 0 ? '#B23A2E' : '#C9C5B8' }}>{o.backordered}</td>
                  <td style={td()}>
                    <span style={{
                      fontSize: 10.5, fontWeight: 700, padding: '2px 7px', borderRadius: 5, textTransform: 'uppercase',
                      color: o.pending ? '#8A6D1F' : (o.backordered > 0 ? '#B58A2E' : '#3E7B4F'),
                      background: o.pending ? '#F3EFE6' : (o.backordered > 0 ? '#FBF6EC' : '#EAF4EC')
                    }}>{o.pending ? 'Pending confirmation' : (o.backordered > 0 ? 'Open' : 'Fulfilled')}</span>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {toast && (
        <div style={{
          position: 'fixed', bottom: 20, left: '50%', transform: 'translateX(-50%)',
          background: '#1C2126', color: 'white', padding: '10px 18px', borderRadius: 8,
          fontSize: 13, display: 'flex', alignItems: 'center', gap: 8, boxShadow: '0 8px 24px rgba(0,0,0,0.25)'
        }}><Check size={14} color="#7FC79A" /> {toast}</div>
      )}
    </div>
  );
}
