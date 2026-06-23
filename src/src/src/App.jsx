import { useState, useEffect, useCallback } from 'react';
import { Users, MapPin, Calendar, TrendingUp, FileText, Plus, Trash2, Edit2, X, Download, Bell, Navigation, CheckCircle2, Clock, Phone, Search, BarChart3, AlertTriangle, Target, ShoppingCart, Sprout, RefreshCw, XCircle, Crosshair, Package, LineChart as LineChartIcon, History, LogOut, Lock, RotateCcw } from 'lucide-react';
import { LineChart, Line, BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid, Legend } from 'recharts';
import * as XLSX from 'xlsx';

// ============================================================
// CONFIG SUPABASE
// ============================================================
const SUPABASE_URL = 'https://jhfomkbceuekdttekbch.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpoZm9ta2JjZXVla2R0dGVrYmNoIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODIxMzc1ODgsImV4cCI6MjA5NzcxMzU4OH0.P3A3WOZ3FcWDbmK0QaIsNQWfUNjYruMLIQWCFpyxwoY';

async function sb(path, opts = {}) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    ...opts,
    headers: {
      'apikey': SUPABASE_ANON_KEY,
      'Authorization': `Bearer ${SUPABASE_ANON_KEY}`,
      'Content-Type': 'application/json',
      'Prefer': opts.prefer || 'return=representation',
      ...(opts.headers || {}),
    },
  });
  if (!res.ok) {
    const txt = await res.text().catch(() => '');
    throw new Error(`Supabase ${res.status}: ${txt.slice(0, 200)}`);
  }
  const text = await res.text();
  return text ? JSON.parse(text) : null;
}

const TVA_DEFAULT = 22;
const COTE_TVA = [11, 22];

const CATEGORII = [
  { id: 'fermier_cm', label: 'Fermier cultură mare', short: 'Fermier CM', color: 'bg-amber-100 text-amber-800 border-amber-300', dot: 'bg-amber-400' },
  { id: 'fermier_h', label: 'Fermier horticol', short: 'Fermier H', color: 'bg-green-100 text-green-800 border-green-300', dot: 'bg-green-500' },
  { id: 'distrib_cm', label: 'Distribuitor CM', short: 'Distrib. CM', color: 'bg-blue-100 text-blue-800 border-blue-300', dot: 'bg-blue-500' },
  { id: 'distrib_h', label: 'Distribuitor Horti', short: 'Distrib. H', color: 'bg-teal-100 text-teal-800 border-teal-300', dot: 'bg-teal-500' },
  { id: 'prospect', label: 'Prospect (nonclient)', short: 'Prospect', color: 'bg-purple-100 text-purple-800 border-purple-300', dot: 'bg-purple-500' },
  { id: 'altii', label: 'Alții', short: 'Alții', color: 'bg-gray-100 text-gray-700 border-gray-300', dot: 'bg-gray-400' },
];
const catOf = (id) => CATEGORII.find(c => c.id === id) || CATEGORII[5];

const STADII = ['Prospectare', 'Ofertare', 'Negociere', 'Câștigată', 'Pierdută'];
const stadiuColor = { 'Prospectare': 'bg-slate-100 text-slate-700', 'Ofertare': 'bg-sky-100 text-sky-800', 'Negociere': 'bg-amber-100 text-amber-800', 'Câștigată': 'bg-green-100 text-green-800', 'Pierdută': 'bg-rose-100 text-rose-700' };
const OFERTA_STATUS = ['Draft', 'Vânzare', 'Respinsă'];
const ofStatusColor = { 'Draft': 'bg-slate-100 text-slate-600 border-slate-300', 'Vânzare': 'bg-emerald-100 text-emerald-700 border-emerald-300', 'Respinsă': 'bg-rose-100 text-rose-700 border-rose-300' };

const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const localISO = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
const todayStr = () => localISO(new Date());
const addDays = (d, n) => { const x = new Date(d + 'T00:00:00'); x.setDate(x.getDate() + n); return localISO(x); };
const fmt = (n) => (n || n === 0) ? Number(n).toLocaleString('ro-RO', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '0,00';
const emptyCulturi = () => Array.from({ length: 3 }, () => ({ cultura: '', suprafata: '' }));
const VALAB_DEFAULT = 30;
const prodLabel = (p) => {
  const base = [p.cod, p.nume, p.ambalaj].filter(Boolean).join(' \u00B7 ');
  const arePret = (p.pret || p.pret === 0) && String(p.pret) !== '';
  return arePret ? base + ' \u00B7 ' + fmt(p.pret) + ' lei' : base;
};

function calcLinie(l) {
  const cant = Number(l.cantitate) || 0;
  const pret = Number(l.pretLista) || 0;
  const red = Number(l.reducere) || 0;
  const redTip = l.redTip || 'pct';
  const tva = (l.tva === '' || l.tva == null) ? TVA_DEFAULT : Number(l.tva);
  const pretNet = redTip === 'lei' ? Math.max(0, pret - red) : pret * (1 - red / 100);
  const valFaraTVA = cant * pretNet;
  const valTVA = valFaraTVA * (tva / 100);
  return { pretNet, valFaraTVA, valTVA, valCuTVA: valFaraTVA + valTVA, tva };
}
function calcOferta(linii) {
  return (linii || []).reduce((acc, l) => { const c = calcLinie(l); acc.faraTVA += c.valFaraTVA; acc.tva += c.valTVA; acc.cuTVA += c.valCuTVA; return acc; }, { faraTVA: 0, tva: 0, cuTVA: 0 });
}
const isExpirat = (o) => o.ofStatus === 'Draft' && o.valabil && o.valabil < todayStr();
const zileRamase = (o) => o.valabil ? Math.ceil((new Date(o.valabil) - new Date(todayStr())) / 86400000) : null;

const clientFromDb = (r) => ({ id: r.id, nume: r.nume, categorie: r.categorie, status: r.status, contact: r.contact || '', telefon: r.telefon || '', email: r.email || '', localitate: r.localitate || '', judet: r.judet || '', lat: r.lat || '', lng: r.lng || '', culturi: r.culturi && r.culturi.length ? r.culturi : emptyCulturi(), ultimaVizita: r.ultima_vizita || '', notite: r.notite || '', atcResponsabil: r.atc_responsabil || '' });
const clientToDb = (c) => ({ nume: c.nume, categorie: c.categorie, status: c.status, contact: c.contact, telefon: c.telefon, email: c.email, localitate: c.localitate, judet: c.judet, lat: c.lat, lng: c.lng, culturi: c.culturi, ultima_vizita: c.ultimaVizita, notite: c.notite, atc_responsabil: c.atcResponsabil });

const productFromDb = (r) => ({ id: r.id, cod: r.cod || '', nume: r.nume, categorie: r.categorie || '', compozitie: r.compozitie || '', doze: r.doze || '', culturi: r.culturi || '', ambalaj: r.ambalaj || '', densitate: r.densitate || '', um: r.um || 'L', pret: r.pret || '', tva: r.tva ?? TVA_DEFAULT, istoricPret: r.istoric_pret || [], creatLa: r.created_at ? r.created_at.slice(0, 10) : todayStr() });
const productToDb = (p) => ({ cod: p.cod, nume: p.nume, categorie: p.categorie, compozitie: p.compozitie, doze: p.doze, culturi: p.culturi, ambalaj: p.ambalaj, densitate: p.densitate, um: p.um, pret: p.pret, tva: Number(p.tva) || TVA_DEFAULT, istoric_pret: p.istoricPret || [] });

const reportFromDb = (r) => ({ id: r.id, clientId: r.client_id, data: r.data, notite: r.notite || '', atc: r.atc || '', foto: r.foto || [] });
const reportToDb = (r) => ({ client_id: r.clientId, data: r.data, notite: r.notite, atc: r.atc, foto: r.foto || [] });

const oppFromDb = (r) => ({ id: r.id, nrOferta: r.nr_oferta || '', clientId: r.client_id, titlu: r.titlu || '', stadiu: r.stadiu, ofStatus: r.of_status, data: r.data, valabil: r.valabil, dataInchidere: r.data_inchidere || '', notite: r.notite || '', linii: r.linii || [], atc: r.atc || '' });
const oppToDb = (o) => ({ nr_oferta: o.nrOferta, client_id: o.clientId, titlu: o.titlu, stadiu: o.stadiu, of_status: o.ofStatus, data: o.data, valabil: o.valabil, data_inchidere: o.dataInchidere, notite: o.notite, linii: o.linii, atc: o.atc });

const planFromDb = (r) => ({ id: r.id, clientId: r.client_id, tip: r.tip, data: r.data, nota: r.nota || '', done: !!r.done, atc: r.atc || '' });
const planToDb = (p) => ({ client_id: p.clientId, tip: p.tip, data: p.data, nota: p.nota, done: p.done, atc: p.atc });

// ============================================================
// LOGIN — sesiunea se tine in localStorage (functioneaza normal aici,
// nu suntem in sandboxul de artifact Claude)
// ============================================================
function LoginScreen({ onLogin }) {
  const [users, setUsers] = useState(null);
  const [error, setError] = useState('');
  const [selected, setSelected] = useState('');
  const [pin, setPin] = useState('');
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    (async () => {
      try {
        const rows = await sb('atc_users?select=nume&order=nume.asc');
        setUsers(rows.map(r => r.nume));
      } catch (e) { setError('Nu pot contacta baza de date: ' + e.message); }
    })();
  }, []);

  const submit = async () => {
    if (!selected) { setError('Selectează-ți numele.'); return; }
    if (!pin.trim()) { setError('Introdu PIN-ul.'); return; }
    setBusy(true); setError('');
    try {
      const rows = await sb(`atc_users?nume=eq.${encodeURIComponent(selected)}&select=nume,pin`);
      if (!rows.length || rows[0].pin !== pin.trim()) { setError('PIN incorect.'); setBusy(false); return; }
      localStorage.setItem('crm-session', JSON.stringify({ nume: selected, ts: Date.now() }));
      onLogin(selected);
    } catch (e) { setError('Eroare: ' + e.message); }
    setBusy(false);
  };

  return (
    <div className="min-h-screen flex items-center justify-center p-4" style={{ background: 'linear-gradient(135deg,#15803d,#65a30d)' }}>
      <div className="bg-white rounded-2xl shadow-xl p-6 w-full max-w-sm">
        <div className="flex flex-col items-center mb-5">
          <div className="w-14 h-14 rounded-2xl bg-green-100 flex items-center justify-center mb-2"><Sprout size={28} className="text-green-600" /></div>
          <div className="font-bold text-lg text-slate-800">CRM Agro</div>
          <div className="text-xs text-slate-400">Autentificare echipă</div>
        </div>
        {users === null && !error && <div className="text-sm text-slate-400 text-center py-4">Se conectează...</div>}
        {error && <div className="text-xs bg-rose-50 text-rose-600 rounded-lg p-2 mb-3">{error}</div>}
        {users && <>
          <label className="block text-xs text-slate-500 mb-1">Nume</label>
          <select className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm mb-3" value={selected} onChange={e => setSelected(e.target.value)}>
            <option value="">Selectează...</option>
            {users.map(u => <option key={u} value={u}>{u}</option>)}
          </select>
          <label className="block text-xs text-slate-500 mb-1">PIN</label>
          <div className="relative mb-4">
            <Lock size={14} className="absolute left-3 top-3 text-slate-300" />
            <input type="password" className="w-full border border-slate-200 rounded-lg pl-8 pr-3 py-2 text-sm" value={pin} onChange={e => setPin(e.target.value)} onKeyDown={e => e.key === 'Enter' && submit()} placeholder="••••" />
          </div>
          <button onClick={submit} disabled={busy} className="w-full bg-green-600 hover:bg-green-700 disabled:opacity-50 text-white py-2.5 rounded-lg text-sm font-medium transition">{busy ? 'Se verifică...' : 'Intră în cont'}</button>
        </>}
      </div>
    </div>
  );
}

export default function App() {
  const [me, setMe] = useState(null);
  const [checking, setChecking] = useState(true);

  const [clients, setClients] = useState([]);
  const [products, setProducts] = useState([]);
  const [reports, setReports] = useState([]);
  const [opportunities, setOpportunities] = useState([]);
  const [plans, setPlans] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [syncing, setSyncing] = useState(false);
  const [loadError, setLoadError] = useState('');

  const [tab, setTab] = useState('dashboard');
  const [modal, setModal] = useState(null);
  const [q, setQ] = useState('');
  const [filterCat, setFilterCat] = useState('');
  const [perioada, setPerioada] = useState(12);

  useEffect(() => {
    try {
      const raw = localStorage.getItem('crm-session');
      if (raw) { const s = JSON.parse(raw); if (s && s.nume) setMe(s.nume); }
    } catch (e) {}
    setChecking(false);
  }, []);

  const loadAll = useCallback(async () => {
    setSyncing(true); setLoadError('');
    try {
      const [c, p, r, o, pl] = await Promise.all([
        sb('clients?select=*&order=nume.asc'),
        sb('products?select=*&order=nume.asc'),
        sb('reports?select=*&order=data.desc'),
        sb('opportunities?select=*&order=data.desc'),
        sb('plans?select=*&order=data.asc'),
      ]);
      setClients(c.map(clientFromDb));
      setProducts(p.map(productFromDb));
      setReports(r.map(reportFromDb));
      setOpportunities(o.map(oppFromDb));
      setPlans(pl.map(planFromDb));
      setLoaded(true);
    } catch (e) { setLoadError('Eroare la încărcare: ' + e.message); }
    setSyncing(false);
  }, []);

  useEffect(() => { if (me) loadAll(); }, [me, loadAll]);

  const logout = () => { localStorage.removeItem('crm-session'); setMe(null); setLoaded(false); };

  const clientById = (id) => clients.find(c => c.id === id);

  const insertRow = async (table, dbObj) => { const [row] = await sb(`${table}`, { method: 'POST', body: JSON.stringify(dbObj) }); return row; };
  const updateRow = async (table, id, dbObj) => { await sb(`${table}?id=eq.${id}`, { method: 'PATCH', body: JSON.stringify(dbObj) }); };
  const deleteRow = async (table, id) => { await sb(`${table}?id=eq.${id}`, { method: 'DELETE', prefer: 'return=minimal' }); };
  const deleteAll = async (table) => { await sb(`${table}?id=neq.00000000-0000-0000-0000-000000000000`, { method: 'DELETE', prefer: 'return=minimal' }); };

  const saveClient = async (c) => {
    try {
      if (clients.some(x => x.id === c.id)) { await updateRow('clients', c.id, clientToDb(c)); }
      else { const row = await insertRow('clients', clientToDb({ ...c, atcResponsabil: c.atcResponsabil || me })); c.id = row.id; }
      await loadAll();
    } catch (e) { alert('Eroare salvare client: ' + e.message); }
  };
  const removeClient = async (id) => { try { await deleteRow('clients', id); await loadAll(); } catch (e) { alert('Eroare ștergere: ' + e.message); } };

  const saveProduct = async (p) => {
    try {
      let out = { ...p };
      if (!out.istoricPret || out.istoricPret.length === 0) out.istoricPret = p.pret ? [{ data: todayStr(), pret: Number(p.pret) }] : [];
      if (products.some(x => x.id === p.id)) await updateRow('products', p.id, productToDb(out));
      else { const row = await insertRow('products', productToDb(out)); out.id = row.id; }
      await loadAll();
    } catch (e) { alert('Eroare salvare produs: ' + e.message); }
  };
  const removeProduct = async (id) => { try { await deleteRow('products', id); await loadAll(); } catch (e) { alert('Eroare: ' + e.message); } };

  const saveReport = async (r) => {
    try {
      const payload = { ...r, atc: r.atc || me };
      if (reports.some(x => x.id === r.id)) await updateRow('reports', r.id, reportToDb(payload));
      else { const row = await insertRow('reports', reportToDb(payload)); payload.id = row.id; }
      const cl = clientById(payload.clientId);
      if (cl) {
        const allForClient = [...reports.filter(x => x.clientId === cl.id && x.id !== payload.id), payload];
        const maxData = allForClient.map(x => x.data).sort().pop();
        if (maxData && maxData !== cl.ultimaVizita) await updateRow('clients', cl.id, { ultima_vizita: maxData });
      }
      await loadAll();
    } catch (e) { alert('Eroare salvare vizită: ' + e.message); }
  };
  const removeReport = async (id) => { try { await deleteRow('reports', id); await loadAll(); } catch (e) { alert('Eroare: ' + e.message); } };

  const saveOpp = async (o) => {
    try {
      let item = { ...o, atc: o.atc || me };
      if (!item.nrOferta) {
        const rows = await sb('oferta_seq?id=eq.1&select=seq');
        const cur = (rows[0]?.seq || 0) + 1;
        await sb('oferta_seq?id=eq.1', { method: 'PATCH', body: JSON.stringify({ seq: cur }) });
        item.nrOferta = `OF-${new Date().getFullYear()}-${String(cur).padStart(4, '0')}`;
      }
      if (opportunities.some(x => x.id === o.id)) await updateRow('opportunities', o.id, oppToDb(item));
      else { const row = await insertRow('opportunities', oppToDb(item)); item.id = row.id; }
      await loadAll();
    } catch (e) { alert('Eroare salvare ofertă: ' + e.message); }
  };
  const removeOpp = async (id) => { try { await deleteRow('opportunities', id); await loadAll(); } catch (e) { alert('Eroare: ' + e.message); } };

  const savePlan = async (p) => {
    try {
      const payload = { ...p, atc: p.atc || me };
      if (plans.some(x => x.id === p.id)) await updateRow('plans', p.id, planToDb(payload));
      else { const row = await insertRow('plans', planToDb(payload)); payload.id = row.id; }
      await loadAll();
    } catch (e) { alert('Eroare salvare activitate: ' + e.message); }
  };
  const removePlan = async (id) => { try { await deleteRow('plans', id); await loadAll(); } catch (e) { alert('Eroare: ' + e.message); } };

  if (checking) return <div className="min-h-screen flex items-center justify-center text-slate-400 text-sm">Se verifică sesiunea...</div>;
  if (!me) return <LoginScreen onLogin={setMe} />;

  const curMonth = new Date().toISOString().slice(0, 7);
  const vizLuna = reports.filter(r => (r.data || '').slice(0, 7) === curMonth).length;
  const oportActive = opportunities.filter(o => !['Câștigată', 'Pierdută'].includes(o.stadiu));
  const vanzari = opportunities.filter(o => o.ofStatus === 'Vânzare');
  const sumaVanzari = vanzari.reduce((s, o) => s + calcOferta(o.linii).cuTVA, 0);
  const valPipeline = oportActive.reduce((s, o) => s + calcOferta(o.linii).cuTVA, 0);
  const activi = clients.filter(c => c.status !== 'inactiv').length;
  const expirate = opportunities.filter(isExpirat).length;

  const planSorted = [...plans].filter(p => !p.done).sort((a, b) => (a.data || '').localeCompare(b.data || ''));
  const t = todayStr();
  const restante = planSorted.filter(p => p.data < t);
  const azi = planSorted.filter(p => p.data === t);
  const urmatoarele = planSorted.filter(p => p.data > t).slice(0, 8);

  const pipeline = STADII.map(s => ({ s, n: opportunities.filter(o => o.stadiu === s).length }));
  const maxPipe = Math.max(1, ...pipeline.map(p => p.n));
  const byCat = CATEGORII.map(c => ({ ...c, n: clients.filter(cl => cl.categorie === c.id).length }));
  const maxCat = Math.max(1, ...byCat.map(c => c.n));

  const LUNI_RO = ['ian', 'feb', 'mar', 'apr', 'mai', 'iun', 'iul', 'aug', 'sep', 'oct', 'noi', 'dec'];
  function monthLabel(m) { const parts = String(m).split('-'); const y = parts[0] || ''; const mo = Number(parts[1]) || 1; return LUNI_RO[mo - 1] + ' ' + y.slice(2); }
  function buildEvolutie(nLuni) {
    const months = []; const now = new Date();
    for (let i = nLuni - 1; i >= 0; i--) { const d = new Date(now.getFullYear(), now.getMonth() - i, 1); months.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`); }
    const map = {}; for (const m of months) map[m] = { luna: m, label: monthLabel(m), nr: 0, faraTVA: 0, cuTVA: 0, cantitate: 0 };
    for (const o of vanzari) {
      const m = String(o.dataInchidere || o.data || '').slice(0, 7);
      if (map[m]) { const tot = calcOferta(o.linii); map[m].nr += 1; map[m].faraTVA += tot.faraTVA; map[m].cuTVA += tot.cuTVA; let qty = 0; for (const l of (o.linii || [])) qty += (Number(l.cantitate) || 0); map[m].cantitate += qty; }
    }
    return months.map(m => map[m]);
  }
  const evol = buildEvolutie(perioada);
  const evolNezero = evol.filter(m => m.nr > 0);
  const totalCantitate = evol.reduce((s, m) => s + m.cantitate, 0);
  const evolNr = evol.reduce((s, m) => s + m.nr, 0);
  const evolFaraTVA = evol.reduce((s, m) => s + m.faraTVA, 0);
  const evolCuTVA = evol.reduce((s, m) => s + m.cuTVA, 0);
  const evol12 = buildEvolutie(12);
  const evol12HasData = evol12.some(m => m.nr > 0);

  const exportComplet = () => {
    const clientRows = clients.map(c => ({ Nume: c.nume, Categorie: catOf(c.categorie).label, Status: c.status === 'inactiv' ? 'Inactiv' : 'Activ', ATC: c.atcResponsabil, Persoana_contact: c.contact || '', Telefon: c.telefon || '', Localitate: c.localitate || '', Judet: c.judet || '', Culturi: (c.culturi || []).filter(x => x.cultura).map(x => `${x.cultura}${x.suprafata ? ' (' + x.suprafata + ' ha)' : ''}`).join('; '), Ultima_vizita: c.ultimaVizita || '', Observatii: c.notite || '' }));
    const repRows = reports.map(r => { const cl = clientById(r.clientId); return { Data: r.data, ATC: r.atc, Client: cl ? cl.nume : '—', Notite: r.notite || '' }; });
    const opRows = opportunities.map(o => { const cl = clientById(o.clientId); const tot = calcOferta(o.linii); return { Nr_oferta: o.nrOferta || '', ATC: o.atc, Data: o.data, Client: cl ? cl.nume : '—', Titlu: o.titlu || '', Status_oferta: o.ofStatus, Stadiu: o.stadiu, Valoare_cu_TVA: tot.cuTVA.toFixed(2) }; });
    const planRows = plans.map(p => { const cl = clientById(p.clientId); return { Data: p.data, ATC: p.atc, Tip: p.tip, Client: cl ? cl.nume : '—', Status: p.done ? 'Realizat' : 'Planificat', Nota: p.nota || '' }; });
    const wb = XLSX.utils.book_new();
    const add = (rows, name) => XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.length ? rows : [{}]), name);
    add(clientRows, 'Clienți'); add(repRows, 'Vizite'); add(opRows, 'Oferte'); add(planRows, 'Planificări');
    XLSX.writeFile(wb, `CRM_agro_export_${todayStr()}.xlsx`);
  };

  const exportOfertaPDF = (o) => {
    const cl = clientById(o.clientId); const tot = calcOferta(o.linii);
    const esc = (s) => String(s == null ? '' : s).replace(/[&<>]/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[m]));
    const linii = (o.linii || []).map((l, i) => { const c = calcLinie(l); const redTxt = (Number(l.reducere) || 0) ? ((l.redTip === 'lei') ? `${fmt(l.reducere)} lei/u` : `${l.reducere}%`) : '—'; return `<tr><td style="text-align:center">${i + 1}</td><td>${esc(l.cod)}</td><td>${esc(l.produs)}</td><td>${esc(l.ambalaj)}</td><td class=r>${esc(l.cantitate)}</td><td class=r>${fmt(l.pretLista)}</td><td class=r>${redTxt}</td><td class=r>${fmt(c.pretNet)}</td><td class=r>${c.tva}%</td><td class=r>${fmt(c.valFaraTVA)}</td><td class=r><b>${fmt(c.valCuTVA)}</b></td></tr>`; }).join('');
    const html = `<!doctype html><html lang="ro"><head><meta charset="utf-8"><title>${esc(o.nrOferta || 'Oferta')}</title><style>*{font-family:'Segoe UI',Arial,sans-serif;box-sizing:border-box}body{margin:32px;color:#1f2937;font-size:13px}.hd{display:flex;justify-content:space-between;align-items:flex-start;border-bottom:4px solid #16a34a;padding-bottom:14px}.leaf{font-size:13px;color:#16a34a;font-weight:bold}table{width:100%;border-collapse:collapse;margin-top:8px}th{background:#16a34a;color:#fff;padding:8px 6px;font-size:11px;text-align:left}.r{text-align:right}td{padding:7px 6px;border-bottom:1px solid #e5e7eb;font-size:12px}.tot{text-align:right;margin-top:14px;line-height:1.8}.big{font-size:17px;color:#15803d;font-weight:bold}.badge{display:inline-block;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:bold}@media print{body{margin:14mm}.noprint{display:none}}</style></head><body>
      <div class="hd"><div><div class="leaf">🌱 OFERTĂ COMERCIALĂ</div><div style="font-size:11px;color:#6b7280">Rețea comercială horticolă · ATC: ${esc(o.atc)}</div></div><div style="text-align:right;font-size:11px;color:#374151"><div style="font-size:15px;font-weight:bold;color:#15803d">${esc(o.nrOferta || '—')}</div><div><b>Data:</b> ${esc(o.data)}</div>${o.valabil ? `<div><b>Valabilă până:</b> ${esc(o.valabil)}</div>` : ''}<div style="margin-top:4px"><span class="badge" style="background:#dcfce7;color:#15803d">${esc(o.ofStatus)}</span></div></div></div>
      <div style="margin:18px 0;font-size:12px"><b>Către:</b> ${esc(cl ? cl.nume : '—')} · ${esc(cl?.contact || '')} · ${esc([cl?.localitate, cl?.judet].filter(Boolean).join(', '))} · ${esc(cl?.telefon || '')}<br><b>Referință:</b> ${esc(o.titlu || '—')}</div>
      <table><thead><tr><th>#</th><th>Cod</th><th>Produs</th><th>Ambalaj</th><th class=r>Cant.</th><th class=r>Preț listă</th><th class=r>Reducere</th><th class=r>Preț net</th><th class=r>TVA</th><th class=r>Val. fără TVA</th><th class=r>Val. cu TVA</th></tr></thead><tbody>${linii}</tbody></table>
      <div class="tot"><div>Total fără TVA: <b>${fmt(tot.faraTVA)} lei</b></div><div>TVA: <b>${fmt(tot.tva)} lei</b></div><div class="big">Total de plată: ${fmt(tot.cuTVA)} lei</div></div>
      ${o.notite ? `<div style="margin-top:16px;font-size:11px;color:#6b7280"><b>Observații:</b> ${esc(o.notite)}</div>` : ''}
      <div class="noprint" style="margin-top:24px;text-align:center"><button onclick="window.print()" style="background:#16a34a;color:#fff;border:none;padding:10px 20px;border-radius:8px;font-size:14px;cursor:pointer">🖨️ Printează / Salvează PDF</button></div></body></html>`;
    const blob = new Blob([html], { type: 'text/html' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = `${o.nrOferta || 'Oferta'}_${esc(cl ? cl.nume : '').replace(/[^\w]/g, '_')}.html`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  const exportListPDF = (rows, cols, title) => {
    const esc = (s) => String(s == null ? '' : s).replace(/[&<>]/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[m]));
    const thead = cols.map(c => `<th>${esc(c.label)}</th>`).join('');
    const tbody = rows.map(r => `<tr>${cols.map(c => `<td class="${c.r ? 'r' : ''}">${esc(c.fmt ? c.fmt(r) : r[c.key])}</td>`).join('')}</tr>`).join('');
    const html = `<!doctype html><html lang="ro"><head><meta charset="utf-8"><title>${esc(title)}</title><style>*{font-family:'Segoe UI',Arial,sans-serif;box-sizing:border-box}body{margin:28px;color:#1f2937;font-size:12px}h1{font-size:18px;color:#15803d;border-bottom:4px solid #16a34a;padding-bottom:10px}table{width:100%;border-collapse:collapse}th{background:#16a34a;color:#fff;padding:7px 6px;font-size:11px;text-align:left}.r{text-align:right}td{padding:6px;border-bottom:1px solid #e5e7eb;font-size:11px}@media print{body{margin:14mm}.noprint{display:none}}</style></head><body><h1>🌱 ${esc(title)}</h1><table><thead><tr>${thead}</tr></thead><tbody>${tbody}</tbody></table><div class="noprint" style="margin-top:20px;text-align:center"><button onclick="window.print()" style="background:#16a34a;color:#fff;border:none;padding:10px 20px;border-radius:8px;font-size:14px;cursor:pointer">🖨️ Printează / Salvează PDF</button></div></body></html>`;
    const blob = new Blob([html], { type: 'text/html' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = `${title.replace(/[^\w]/g, '_')}_${todayStr()}.html`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };
  const exportListExcel = (rows, cols, title) => {
    const data = rows.map(r => { const o = {}; cols.forEach(c => { o[c.label] = c.fmt ? c.fmt(r) : r[c.key]; }); return o; });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data.length ? data : [{}]), title.slice(0, 28));
    XLSX.writeFile(wb, `${title.replace(/[^\w]/g, '_')}_${todayStr()}.xlsx`);
  };
  const exportListMail = (rows, cols, title) => {
    const lines = rows.slice(0, 40).map(r => cols.map(c => c.label + ': ' + (c.fmt ? c.fmt(r) : r[c.key])).join(' | '));
    const body = title + '\n(CRM Agro, ' + todayStr() + ')\n\n' + lines.join('\n');
    window.location.href = 'mailto:?subject=' + encodeURIComponent(title) + '&body=' + encodeURIComponent(body);
  };
  const exportListWhatsApp = (rows, cols, title) => {
    const lines = rows.slice(0, 25).map(r => cols.map(c => c.label + ': ' + (c.fmt ? c.fmt(r) : r[c.key])).join(' | '));
    window.open('https://wa.me/?text=' + encodeURIComponent('*' + title + '*\n' + lines.join('\n')), '_blank');
  };

  const setCoordCurente = (setter) => {
    if (!navigator.geolocation) { alert('Geolocalizarea nu este disponibilă.'); return; }
    navigator.geolocation.getCurrentPosition(
      (pos) => setter(pos.coords.latitude.toFixed(6), pos.coords.longitude.toFixed(6)),
      (err) => alert('Nu am putut obține locația: ' + err.message),
      { enableHighAccuracy: true, timeout: 10000 }
    );
  };

  const clientsFiltered = clients.filter(c => (!filterCat || c.categorie === filterCat) && (!q || (c.nume + ' ' + (c.localitate || '') + ' ' + (c.contact || '')).toLowerCase().includes(q.toLowerCase())));

  if (!loaded) return <div className="min-h-screen flex items-center justify-center text-slate-400 text-sm">{loadError ? <span className="text-rose-500">{loadError}</span> : 'Se încarcă datele din Supabase...'}</div>;

  const tabs = [
    { id: 'dashboard', label: 'Dashboard', icon: BarChart3 },
    { id: 'clienti', label: 'Clienți', icon: Users },
    { id: 'produse', label: 'Produse', icon: Package },
    { id: 'rapoarte', label: 'Vizite', icon: FileText },
    { id: 'oportunitati', label: 'Oferte', icon: Target },
    { id: 'vanzari', label: 'Vânzări', icon: TrendingUp },
    { id: 'evolutie', label: 'Evoluție', icon: LineChartIcon },
    { id: 'planificari', label: 'Planificări', icon: Calendar },
  ];

  return (
    <div className="min-h-screen text-slate-800" style={{ fontFamily: 'system-ui, sans-serif', background: 'linear-gradient(180deg,#f0fdf4 0%,#f8fafc 220px)' }}>
      <div className="text-white px-4 sm:px-5 py-3 sm:py-4 flex items-center justify-between flex-wrap gap-2 relative overflow-hidden" style={{ background: 'linear-gradient(120deg,#15803d 0%,#16a34a 55%,#65a30d 100%)' }}>
        <div className="absolute right-0 top-0 opacity-10 text-9xl leading-none select-none">🌾</div>
        <div className="flex items-center gap-2.5 relative">
          <div className="w-10 h-10 sm:w-11 sm:h-11 rounded-2xl bg-white/20 backdrop-blur flex items-center justify-center"><Sprout size={24} className="text-white" /></div>
          <div><div className="font-bold text-base sm:text-lg leading-tight">CRM Agro</div><div className="text-[11px] sm:text-xs text-green-50/90">{me} · echipă conectată</div></div>
        </div>
        <div className="flex gap-1.5 sm:gap-2 relative flex-wrap items-center">
          <button onClick={loadAll} className="text-xs sm:text-sm bg-white/15 hover:bg-white/25 px-2.5 sm:px-3 py-1.5 rounded-lg backdrop-blur flex items-center gap-1.5 transition" title="Reîncarcă datele"><RotateCcw size={14} className={syncing ? 'animate-spin' : ''} /></button>
          <button onClick={exportComplet} className="text-xs sm:text-sm bg-white/15 hover:bg-white/25 px-2.5 sm:px-3 py-1.5 rounded-lg backdrop-blur flex items-center gap-1.5 transition"><Download size={15} /> Backup Excel</button>
          <button onClick={logout} className="text-xs sm:text-sm bg-white/15 hover:bg-white/25 px-2.5 sm:px-3 py-1.5 rounded-lg backdrop-blur flex items-center gap-1.5 transition"><LogOut size={14} /> Ieșire</button>
        </div>
      </div>

      <div className="bg-white/80 backdrop-blur border-b border-green-100 sticky top-0 z-10 flex overflow-x-auto shadow-sm">
        {tabs.map(tb => {
          const Icon = tb.icon; const active = tab === tb.id;
          return (
            <button key={tb.id} onClick={() => setTab(tb.id)} className={`px-4 py-3 text-sm flex items-center gap-1.5 whitespace-nowrap border-b-[3px] transition ${active ? 'border-green-600 text-green-700 font-semibold bg-green-50/60' : 'border-transparent text-slate-500 hover:text-green-600 hover:bg-green-50/30'}`}>
              <Icon size={16} /> {tb.label}
              {tb.id === 'planificari' && (restante.length + azi.length > 0) && <span className="ml-1 bg-rose-500 text-white text-xs rounded-full px-1.5">{restante.length + azi.length}</span>}
              {tb.id === 'oportunitati' && expirate > 0 && <span className="ml-1 bg-rose-500 text-white text-xs rounded-full px-1.5">{expirate}</span>}
            </button>
          );
        })}
      </div>

      <div className="p-4 sm:p-5 max-w-6xl mx-auto">
        {tab === 'dashboard' && (
          <div className="space-y-5">
            <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
              <Stat icon={Users} label="Clienți" value={clients.length} sub={`${activi} activi`} grad="from-blue-500 to-sky-400" />
              <Stat icon={FileText} label="Vizite (luna)" value={vizLuna} sub={`${reports.length} total`} grad="from-teal-500 to-emerald-400" />
              <Stat icon={Target} label="Oferte active" value={oportActive.length} sub={`${fmt(valPipeline)} pipeline`} grad="from-amber-500 to-yellow-400" />
              <Stat icon={TrendingUp} label="Vânzări" value={vanzari.length} sub={`${fmt(sumaVanzari)} cu TVA`} grad="from-green-600 to-lime-500" />
            </div>
            {expirate > 0 && <div className="bg-rose-50 border border-rose-200 rounded-xl p-3 flex items-center gap-2 text-sm text-rose-700"><AlertTriangle size={16} /> {expirate} {expirate === 1 ? 'ofertă expirată' : 'oferte expirate'} — verifică tab-ul Oferte.</div>}
            {(restante.length > 0 || azi.length > 0) && (
              <Card title="Memento activități" icon={Bell} iconColor="text-rose-500">
                <div className="space-y-2">{restante.map(p => <PlanLine key={p.id} p={p} client={clientById(p.clientId)} late onOpen={() => setModal({ type: 'plan', item: p })} />)}{azi.map(p => <PlanLine key={p.id} p={p} client={clientById(p.clientId)} today onOpen={() => setModal({ type: 'plan', item: p })} />)}</div>
              </Card>
            )}
            <div className="grid md:grid-cols-2 gap-4">
              <Card title="Pipeline oferte" icon={Target} iconColor="text-green-600">
                {pipeline.map(p => (<div key={p.s} className="mb-2.5"><div className="flex justify-between text-sm mb-1"><span>{p.s}</span><span className="text-slate-500">{p.n}</span></div><div className="h-2.5 bg-slate-100 rounded-full"><div className="h-2.5 rounded-full" style={{ width: `${(p.n / maxPipe) * 100}%`, background: 'linear-gradient(90deg,#16a34a,#65a30d)' }} /></div></div>))}
              </Card>
              <Card title="Clienți pe categorie" icon={Users} iconColor="text-amber-500">
                {byCat.map(c => (<div key={c.id} className="mb-2.5"><div className="flex justify-between text-sm mb-1"><span className="flex items-center gap-1.5"><span className={`w-2 h-2 rounded-full ${c.dot}`} />{c.short}</span><span className="text-slate-500">{c.n}</span></div><div className="h-2.5 bg-slate-100 rounded-full"><div className="h-2.5 rounded-full" style={{ width: `${(c.n / maxCat) * 100}%`, background: 'linear-gradient(90deg,#eab308,#facc15)' }} /></div></div>))}
              </Card>
            </div>
            <Card title="Următoarele activități planificate" icon={Calendar} iconColor="text-teal-500">
              {urmatoarele.length === 0 ? <div className="text-sm text-slate-400">Nimic planificat.</div> : <div className="space-y-2">{urmatoarele.map(p => <PlanLine key={p.id} p={p} client={clientById(p.clientId)} onOpen={() => setModal({ type: 'plan', item: p })} />)}</div>}
            </Card>
          </div>
        )}

        {tab === 'clienti' && (
          <div className="space-y-4">
            <div className="flex flex-wrap gap-2 items-center">
              <div className="relative flex-1 min-w-[180px]"><Search size={15} className="absolute left-2.5 top-2.5 text-slate-400" /><input value={q} onChange={e => setQ(e.target.value)} placeholder="Caută client..." className="w-full pl-8 pr-3 py-2 border border-green-200 rounded-lg text-sm bg-white focus:ring-2 focus:ring-green-200 outline-none" /></div>
              <select value={filterCat} onChange={e => setFilterCat(e.target.value)} className="border border-green-200 rounded-lg px-2 py-2 text-sm bg-white"><option value="">Toate categoriile</option>{CATEGORII.map(c => <option key={c.id} value={c.id}>{c.label}</option>)}</select>
              <button onClick={() => setModal({ type: 'client', item: null })} className="bg-green-600 hover:bg-green-700 text-white px-3 py-2 rounded-lg text-sm flex items-center gap-1.5 shadow-sm transition"><Plus size={16} /> Creare client</button>
              <ExportBar rows={clientsFiltered} title="Clienți" cols={[{ key: 'nume', label: 'Nume' }, { key: 'categorie', label: 'Categorie', fmt: c => catOf(c.categorie).label }, { key: 'atc', label: 'ATC', fmt: c => c.atcResponsabil }, { key: 'telefon', label: 'Telefon' }, { key: 'localitate', label: 'Localitate' }, { key: 'ultimaVizita', label: 'Ultima vizită' }]} onPDF={exportListPDF} onExcel={exportListExcel} onMail={exportListMail} onWhatsApp={exportListWhatsApp} />
              {clients.length > 0 && <DelBtn onDelete={async () => { await deleteAll('clients'); await loadAll(); }} label="Șterge tot" />}
            </div>
            {clientsFiltered.length === 0 ? <Empty text="Niciun client încă." /> :
              <div className="grid sm:grid-cols-2 gap-3">
                {clientsFiltered.map(c => (
                  <div key={c.id} className="bg-white rounded-xl border border-green-100 p-4 shadow-sm hover:shadow-md transition">
                    <div className="flex justify-between items-start gap-2">
                      <div>
                        <div className="font-semibold flex items-center gap-2">{c.nume}{c.status === 'inactiv' && <span className="text-xs bg-slate-100 text-slate-500 px-1.5 rounded">inactiv</span>}</div>
                        <div className="flex items-center gap-1.5 mt-1 flex-wrap"><span className={`text-xs px-2 py-0.5 rounded-full border ${catOf(c.categorie).color}`}>{catOf(c.categorie).short}</span>{c.atcResponsabil && <span className="text-xs px-2 py-0.5 rounded-full bg-slate-50 text-slate-500 border border-slate-200">{c.atcResponsabil}</span>}</div>
                      </div>
                      <div className="flex gap-1"><button onClick={() => setModal({ type: 'client', item: c })} className="p-1.5 text-slate-400 hover:text-green-600"><Edit2 size={15} /></button><DelBtn onDelete={() => removeClient(c.id)} /></div>
                    </div>
                    <div className="mt-2 text-sm text-slate-600 space-y-1">
                      {c.localitate && <div className="flex items-center gap-1.5"><MapPin size={13} className="text-green-500" />{c.localitate}{c.judet ? `, ${c.judet}` : ''}</div>}
                      {c.telefon && <div className="flex items-center gap-1.5"><Phone size={13} className="text-green-500" />{c.telefon}{c.contact ? ` · ${c.contact}` : ''}</div>}
                      {(c.culturi || []).filter(x => x.cultura).length > 0 && <div className="text-xs text-slate-500 pt-1 flex items-start gap-1"><Sprout size={12} className="text-green-500 mt-0.5" />{(c.culturi || []).filter(x => x.cultura).map(x => `${x.cultura}${x.suprafata ? ' (' + x.suprafata + ' ha)' : ''}`).join(' · ')}</div>}
                      <div className="flex items-center gap-1.5 text-xs text-slate-400 pt-1"><Clock size={12} /> Ultima vizită: {c.ultimaVizita || '—'}</div>
                    </div>
                    {c.lat && c.lng && (<div className="mt-3 flex gap-2"><a href={`https://waze.com/ul?ll=${c.lat},${c.lng}&navigate=yes`} target="_blank" rel="noreferrer" className="text-xs flex items-center gap-1 bg-sky-50 text-sky-700 px-2 py-1 rounded-lg hover:bg-sky-100"><Navigation size={12} /> Waze</a><a href={`https://www.google.com/maps/search/?api=1&query=${c.lat},${c.lng}`} target="_blank" rel="noreferrer" className="text-xs flex items-center gap-1 bg-emerald-50 text-emerald-700 px-2 py-1 rounded-lg hover:bg-emerald-100"><MapPin size={12} /> Maps</a></div>)}
                  </div>
                ))}
              </div>}
          </div>
        )}

        {tab === 'produse' && (
          <div className="space-y-4">
            <div className="flex flex-wrap justify-between items-center gap-2">
              <div className="text-sm text-slate-500">{products.length} produse în catalog</div>
              <div className="flex gap-2 flex-wrap items-center">
                <ExportBar rows={products} title="Catalog produse" cols={[{ key: 'cod', label: 'Cod' }, { key: 'nume', label: 'Denumire' }, { key: 'pret', label: 'Preț', r: true, fmt: p => fmt(p.pret) }, { key: 'tva', label: 'TVA %', r: true }]} onPDF={exportListPDF} onExcel={exportListExcel} onMail={exportListMail} onWhatsApp={exportListWhatsApp} />
                {products.length > 0 && <DelBtn onDelete={async () => { await deleteAll('products'); await loadAll(); }} label="Șterge tot" />}
                <button onClick={() => setModal({ type: 'product', item: null })} className="bg-green-600 hover:bg-green-700 text-white px-3 py-2 rounded-lg text-sm flex items-center gap-1.5 shadow-sm"><Plus size={16} /> Produs nou</button>
              </div>
            </div>
            {products.length === 0 ? <Empty text="Niciun produs încă." /> :
              <div className="bg-white rounded-xl border border-green-100 overflow-hidden shadow-sm">
                <div className="overflow-x-auto"><table className="w-full text-sm min-w-[680px]">
                  <thead className="text-slate-600 text-left" style={{ background: 'linear-gradient(90deg,#dcfce7,#ecfccb)' }}><tr><th className="px-3 py-2">Cod</th><th className="px-3 py-2">Denumire</th><th className="px-3 py-2">Ambalaj</th><th className="px-3 py-2 text-right">Preț</th><th className="px-3 py-2 text-right">TVA</th><th className="px-2"></th></tr></thead>
                  <tbody>{products.map(p => (<tr key={p.id} className="border-t border-green-50"><td className="px-3 py-2 font-mono text-xs text-slate-500">{p.cod}</td><td className="px-3 py-2 font-medium">{p.nume}</td><td className="px-3 py-2 text-slate-500">{p.ambalaj}</td><td className="px-3 py-2 text-right">{fmt(p.pret)}</td><td className="px-3 py-2 text-right text-slate-500">{p.tva}%</td><td className="px-1 py-2"><div className="flex gap-0.5 items-center"><button title="Evoluție preț" onClick={() => setModal({ type: 'prodHist', item: p })} className="p-1.5 text-slate-400 hover:text-green-600"><LineChartIcon size={14} /></button><button onClick={() => setModal({ type: 'product', item: p })} className="p-1.5 text-slate-400 hover:text-green-600"><Edit2 size={14} /></button><DelBtn onDelete={() => removeProduct(p.id)} size={14} /></div></td></tr>))}</tbody>
                </table></div>
              </div>}
          </div>
        )}

        {tab === 'rapoarte' && (
          <div className="space-y-4">
            <div className="flex justify-between items-center flex-wrap gap-2">
              <div className="text-sm text-slate-500">{reports.length} vizite</div>
              <div className="flex gap-2 flex-wrap items-center">
                <ExportBar rows={reports} title="Vizite" cols={[{ key: 'data', label: 'Data' }, { key: 'atc', label: 'ATC' }, { key: 'client', label: 'Client', fmt: r => (clientById(r.clientId)?.nume || '—') }, { key: 'notite', label: 'Notițe' }]} onPDF={exportListPDF} onExcel={exportListExcel} onMail={exportListMail} onWhatsApp={exportListWhatsApp} />
                {reports.length > 0 && <DelBtn onDelete={async () => { await deleteAll('reports'); await loadAll(); }} label="Șterge tot" />}
                <button onClick={() => setModal({ type: 'report', item: null })} className="bg-green-600 hover:bg-green-700 text-white px-3 py-2 rounded-lg text-sm flex items-center gap-1.5 shadow-sm"><Plus size={16} /> Raport vizită</button>
              </div>
            </div>
            {reports.length === 0 ? <Empty text="Nicio vizită înregistrată." /> :
              <div className="space-y-2">{reports.map(r => { const cl = clientById(r.clientId); return (
                <div key={r.id} className="bg-white rounded-xl border border-green-100 p-3 flex justify-between items-start gap-3 shadow-sm">
                  <div className="flex-1"><div className="flex items-center gap-2 flex-wrap"><span className="font-medium">{cl ? cl.nume : '— client șters —'}</span><span className="text-xs text-slate-400">{r.data}</span>{r.atc && <span className="text-xs bg-slate-50 text-slate-500 border border-slate-200 rounded-full px-2">{r.atc}</span>}</div>{r.notite && <div className="text-sm text-slate-500 mt-1">{r.notite}</div>}</div>
                  <div className="flex gap-1"><button onClick={() => setModal({ type: 'report', item: r })} className="p-1.5 text-slate-400 hover:text-green-600"><Edit2 size={15} /></button><DelBtn onDelete={() => removeReport(r.id)} /></div>
                </div>); })}</div>}
          </div>
        )}

        {tab === 'oportunitati' && (
          <div className="space-y-4">
            <div className="flex justify-between items-center flex-wrap gap-2">
              <div className="text-sm text-slate-500">{opportunities.length} oferte · {fmt(valPipeline)} în pipeline</div>
              <div className="flex gap-2 flex-wrap items-center">
                <ExportBar rows={opportunities} title="Oferte" cols={[{ key: 'nrOferta', label: 'Nr.' }, { key: 'atc', label: 'ATC' }, { key: 'client', label: 'Client', fmt: o => (clientById(o.clientId)?.nume || '—') }, { key: 'ofStatus', label: 'Status' }, { key: 'total', label: 'Total cu TVA', r: true, fmt: o => fmt(calcOferta(o.linii).cuTVA) }]} onPDF={exportListPDF} onExcel={exportListExcel} onMail={exportListMail} onWhatsApp={exportListWhatsApp} />
                {opportunities.length > 0 && <DelBtn onDelete={async () => { await deleteAll('opportunities'); await loadAll(); }} label="Șterge tot" />}
                <button onClick={() => setModal({ type: 'opp', item: null })} className="bg-green-600 hover:bg-green-700 text-white px-3 py-2 rounded-lg text-sm flex items-center gap-1.5 shadow-sm"><Plus size={16} /> Ofertă nouă</button>
              </div>
            </div>
            {opportunities.length === 0 ? <Empty text="Nicio ofertă încă." /> :
              <div className="space-y-2">{opportunities.map(o => {
                const cl = clientById(o.clientId); const tot = calcOferta(o.linii); const exp = isExpirat(o); const zr = zileRamase(o);
                return (
                  <div key={o.id} className={`bg-white rounded-xl border p-4 shadow-sm ${exp ? 'border-rose-300 ring-1 ring-rose-200' : 'border-green-100'}`}>
                    <div className="flex justify-between items-start gap-3">
                      <div className="flex-1">
                        <div className="flex items-center gap-2 flex-wrap"><span className="text-xs font-mono bg-slate-100 text-slate-600 px-1.5 py-0.5 rounded">{o.nrOferta || 'nou'}</span><span className="font-semibold">{cl ? cl.nume : '—'}</span>{o.titlu && <span className="text-sm text-slate-500">· {o.titlu}</span>}{o.atc && <span className="text-xs bg-slate-50 text-slate-500 border border-slate-200 rounded-full px-2">{o.atc}</span>}</div>
                        <div className="flex items-center gap-1.5 flex-wrap mt-1.5">
                          <span className={`text-xs px-2 py-0.5 rounded-full border ${ofStatusColor[o.ofStatus]}`}>{o.ofStatus}</span>
                          <span className={`text-xs px-2 py-0.5 rounded-full ${stadiuColor[o.stadiu]}`}>{o.stadiu}</span>
                          {o.valabil && (exp ? <span className="text-xs px-2 py-0.5 rounded-full bg-rose-100 text-rose-700 font-medium flex items-center gap-1"><AlertTriangle size={11} /> Expirată {o.valabil}</span> : o.ofStatus === 'Draft' ? <span className={`text-xs px-2 py-0.5 rounded-full ${zr <= 5 ? 'bg-amber-100 text-amber-700' : 'bg-slate-50 text-slate-500'}`}>Valabilă până {o.valabil}{zr != null ? ` (${zr}z)` : ''}</span> : null)}
                        </div>
                        <div className="text-sm mt-2 flex flex-wrap gap-x-4 gap-y-0.5"><span className="text-slate-500">Fără TVA: <b className="text-slate-700">{fmt(tot.faraTVA)}</b></span><span className="text-slate-500">TVA: <b className="text-slate-700">{fmt(tot.tva)}</b></span><span className="text-slate-500">Total: <b className="text-green-700">{fmt(tot.cuTVA)}</b></span></div>
                        {exp && (<div className="mt-2 flex gap-2"><button onClick={() => saveOpp({ ...o, valabil: addDays(todayStr(), VALAB_DEFAULT) })} className="text-xs bg-green-50 text-green-700 px-2 py-1 rounded-lg hover:bg-green-100 flex items-center gap-1"><RefreshCw size={12} /> Prelungește 30z</button><button onClick={() => saveOpp({ ...o, ofStatus: 'Respinsă', stadiu: 'Pierdută' })} className="text-xs bg-rose-50 text-rose-700 px-2 py-1 rounded-lg hover:bg-rose-100 flex items-center gap-1"><XCircle size={12} /> Respinge</button></div>)}
                      </div>
                      <div className="flex flex-col items-end gap-1">
                        <div className="flex gap-1"><button onClick={() => exportOfertaPDF(o)} className="text-xs bg-slate-100 text-slate-600 px-2 py-1 rounded-lg hover:bg-slate-200 flex items-center gap-1"><FileText size={13} /> PDF</button>{o.ofStatus !== 'Vânzare' && <button onClick={() => saveOpp({ ...o, ofStatus: 'Vânzare', stadiu: 'Câștigată', dataInchidere: todayStr() })} className="text-xs bg-emerald-50 text-emerald-700 px-2 py-1 rounded-lg hover:bg-emerald-100 flex items-center gap-1"><CheckCircle2 size={13} /> Vânzare</button>}</div>
                        <div className="flex gap-1"><button onClick={() => setModal({ type: 'opp', item: o })} className="p-1.5 text-slate-400 hover:text-green-600"><Edit2 size={15} /></button><DelBtn onDelete={() => removeOpp(o.id)} /></div>
                      </div>
                    </div>
                  </div>
                );
              })}</div>}
          </div>
        )}

        {tab === 'vanzari' && (
          <div className="space-y-4">
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
              <div className="bg-white rounded-xl border border-green-100 p-4 shadow-sm"><div className="text-sm text-slate-500">Nr. vânzări</div><div className="text-3xl font-bold text-green-700 mt-1">{vanzari.length}</div></div>
              <div className="bg-white rounded-xl border border-green-100 p-4 shadow-sm"><div className="text-sm text-slate-500">Sumă vânzări (cu TVA)</div><div className="text-3xl font-bold text-green-700 mt-1">{fmt(sumaVanzari)}</div></div>
              <div className="bg-white rounded-xl border border-amber-100 p-4 shadow-sm col-span-2 sm:col-span-1"><div className="text-sm text-slate-500">În pipeline</div><div className="text-3xl font-bold text-amber-600 mt-1">{fmt(valPipeline)}</div></div>
            </div>
            {evol12HasData && (
              <Card title="Evoluție vânzări — ultimele 12 luni (cu TVA)" icon={LineChartIcon} iconColor="text-green-600">
                <div className="h-56"><ResponsiveContainer width="100%" height="100%"><BarChart data={evol12} margin={{ top: 6, right: 10, left: 0, bottom: 0 }}><CartesianGrid strokeDasharray="3 3" stroke="#e5e7eb" vertical={false} /><XAxis dataKey="label" tick={{ fontSize: 10 }} interval={0} angle={-30} textAnchor="end" height={42} /><YAxis tick={{ fontSize: 10 }} /><Tooltip formatter={(val) => [fmt(val) + ' lei', 'Vânzări']} /><Bar dataKey="cuTVA" fill="#16a34a" radius={[4, 4, 0, 0]} /></BarChart></ResponsiveContainer></div>
              </Card>
            )}
            <div className="text-sm font-semibold text-slate-600 flex items-center gap-1.5"><CheckCircle2 size={15} className="text-green-600" /> Vânzări închise</div>
            {vanzari.length === 0 ? <Empty text="Nicio vânzare închisă." /> :
              <div className="bg-white rounded-xl border border-green-100 overflow-hidden shadow-sm"><div className="overflow-x-auto"><table className="w-full text-sm min-w-[560px]">
                <thead className="text-slate-600 text-left" style={{ background: 'linear-gradient(90deg,#dcfce7,#ecfccb)' }}><tr><th className="px-3 py-2">Nr.</th><th className="px-3 py-2">Data</th><th className="px-3 py-2">Client</th><th className="px-3 py-2">ATC</th><th className="px-3 py-2 text-right">Cu TVA</th></tr></thead>
                <tbody>{vanzari.map(o => { const cl = clientById(o.clientId); const tot = calcOferta(o.linii); return (<tr key={o.id} className="border-t border-green-50"><td className="px-3 py-2 font-mono text-xs text-slate-500">{o.nrOferta}</td><td className="px-3 py-2 text-slate-500">{o.dataInchidere || o.data}</td><td className="px-3 py-2 font-medium">{cl ? cl.nume : '—'}</td><td className="px-3 py-2 text-slate-500">{o.atc}</td><td className="px-3 py-2 text-right font-semibold text-green-700">{fmt(tot.cuTVA)}</td></tr>); })}</tbody>
              </table></div></div>}
          </div>
        )}

        {tab === 'evolutie' && (
          <div className="space-y-4">
            <div className="flex justify-between items-center flex-wrap gap-2">
              <div className="text-sm font-semibold text-slate-600 flex items-center gap-1.5"><LineChartIcon size={15} className="text-green-600" /> Evoluție vânzări — ultimele {perioada} luni</div>
              <div className="inline-flex rounded-lg border border-green-200 overflow-hidden bg-white shadow-sm">{[6, 12, 24].map(p => (<button key={p} onClick={() => setPerioada(p)} className={`px-3 py-1.5 text-sm transition ${perioada === p ? 'bg-green-600 text-white font-semibold' : 'text-slate-600 hover:bg-green-50'}`}>{p} luni</button>))}</div>
            </div>
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
              <Stat icon={CheckCircle2} label={`Total vânzări (${perioada}l)`} value={evolNr} sub="număr tranzacții" grad="from-green-600 to-lime-500" />
              <Stat icon={TrendingUp} label="Sumă (cu TVA)" value={fmt(evolCuTVA)} sub={`ultimele ${perioada} luni`} grad="from-emerald-600 to-green-400" />
              <Stat icon={Package} label="Cantitate totală" value={fmt(totalCantitate)} sub="unități vândute" grad="from-teal-500 to-cyan-400" />
              <Stat icon={BarChart3} label="Medie lunară" value={fmt(evolCuTVA / perioada)} sub="cu TVA / lună" grad="from-amber-500 to-yellow-400" />
            </div>
            {evolNr === 0 ? <Empty text="Nicio vânzare în perioada selectată." /> : <>
              <Card title="Sume vândute pe luni" icon={TrendingUp} iconColor="text-green-600">
                <div className="h-64"><ResponsiveContainer width="100%" height="100%"><BarChart data={evol} margin={{ top: 6, right: 10, left: 0, bottom: 0 }}><CartesianGrid strokeDasharray="3 3" stroke="#e5e7eb" vertical={false} /><XAxis dataKey="label" tick={{ fontSize: perioada > 12 ? 8 : 10 }} interval={perioada > 12 ? 1 : 0} angle={-30} textAnchor="end" height={42} /><YAxis tick={{ fontSize: 10 }} /><Tooltip formatter={(val, name) => [fmt(val) + ' lei', name === 'cuTVA' ? 'Cu TVA' : 'Fara TVA']} /><Legend formatter={(val) => val === 'cuTVA' ? 'Cu TVA' : 'Fara TVA'} wrapperStyle={{ fontSize: 12 }} /><Bar dataKey="faraTVA" fill="#86efac" radius={[3, 3, 0, 0]} /><Bar dataKey="cuTVA" fill="#16a34a" radius={[3, 3, 0, 0]} /></BarChart></ResponsiveContainer></div>
              </Card>
              <div className="flex justify-between items-center flex-wrap gap-2">
                <div className="text-sm font-semibold text-slate-600">Tabel lunar ({perioada} luni)</div>
                <ExportBar rows={evolNezero} title={`Evolutie vanzari ${perioada} luni`} cols={[{ key: 'label', label: 'Luna' }, { key: 'nr', label: 'Nr. vânzări' }, { key: 'cuTVA', label: 'Sumă cu TVA', r: true, fmt: m => fmt(m.cuTVA) }]} onPDF={exportListPDF} onExcel={exportListExcel} onMail={exportListMail} onWhatsApp={exportListWhatsApp} />
              </div>
              <div className="bg-white rounded-xl border border-green-100 overflow-hidden shadow-sm"><div className="overflow-x-auto"><table className="w-full text-sm min-w-[520px]">
                <thead className="text-slate-600 text-left" style={{ background: 'linear-gradient(90deg,#dcfce7,#ecfccb)' }}><tr><th className="px-3 py-2">Luna</th><th className="px-3 py-2 text-right">Nr.</th><th className="px-3 py-2 text-right">Cantitate</th><th className="px-3 py-2 text-right">Fără TVA</th><th className="px-3 py-2 text-right">Cu TVA</th></tr></thead>
                <tbody>{evol.map(m => (<tr key={m.luna} className={`border-t border-green-50 ${m.nr === 0 ? 'text-slate-300' : ''}`}><td className="px-3 py-2 font-medium capitalize">{m.label}</td><td className="px-3 py-2 text-right">{m.nr}</td><td className="px-3 py-2 text-right">{m.cantitate ? fmt(m.cantitate) : '—'}</td><td className="px-3 py-2 text-right">{m.faraTVA ? fmt(m.faraTVA) : '—'}</td><td className="px-3 py-2 text-right font-semibold text-green-700">{m.cuTVA ? fmt(m.cuTVA) : '—'}</td></tr>))}</tbody>
                <tfoot><tr className="border-t-2 border-green-200 bg-green-50/40 font-semibold"><td className="px-3 py-2">TOTAL</td><td className="px-3 py-2 text-right">{evolNr}</td><td className="px-3 py-2 text-right">{fmt(totalCantitate)}</td><td className="px-3 py-2 text-right">{fmt(evolFaraTVA)}</td><td className="px-3 py-2 text-right text-green-700">{fmt(evolCuTVA)}</td></tr></tfoot>
              </table></div></div>
            </>}
          </div>
        )}

        {tab === 'planificari' && (
          <div className="space-y-4">
            <div className="flex justify-between items-center flex-wrap gap-2">
              <div className="text-sm text-slate-500">{planSorted.length} activități deschise</div>
              <div className="flex gap-2 flex-wrap items-center">
                <ExportBar rows={plans} title="Planificări" cols={[{ key: 'data', label: 'Data' }, { key: 'atc', label: 'ATC' }, { key: 'tip', label: 'Tip' }, { key: 'client', label: 'Client', fmt: p => (clientById(p.clientId)?.nume || '—') }, { key: 'status', label: 'Status', fmt: p => (p.done ? 'Realizat' : 'Planificat') }]} onPDF={exportListPDF} onExcel={exportListExcel} onMail={exportListMail} onWhatsApp={exportListWhatsApp} />
                {plans.length > 0 && <DelBtn onDelete={async () => { await deleteAll('plans'); await loadAll(); }} label="Șterge tot" />}
                <button onClick={() => setModal({ type: 'plan', item: null })} className="bg-green-600 hover:bg-green-700 text-white px-3 py-2 rounded-lg text-sm flex items-center gap-1.5 shadow-sm"><Plus size={16} /> Activitate nouă</button>
              </div>
            </div>
            {restante.length > 0 && <Card title="Restante" icon={AlertTriangle} iconColor="text-rose-600">{restante.map(p => <PlanRow key={p.id} p={p} c={clientById(p.clientId)} late onDone={() => savePlan({ ...p, done: true })} onEdit={() => setModal({ type: 'plan', item: p })} onDel={() => removePlan(p.id)} />)}</Card>}
            {azi.length > 0 && <Card title="Astăzi" icon={Bell} iconColor="text-amber-500">{azi.map(p => <PlanRow key={p.id} p={p} c={clientById(p.clientId)} today onDone={() => savePlan({ ...p, done: true })} onEdit={() => setModal({ type: 'plan', item: p })} onDel={() => removePlan(p.id)} />)}</Card>}
            <Card title="Viitoare" icon={Calendar} iconColor="text-teal-500">{planSorted.filter(p => p.data > t).length === 0 ? <div className="text-sm text-slate-400 px-1">Nimic planificat.</div> : planSorted.filter(p => p.data > t).map(p => <PlanRow key={p.id} p={p} c={clientById(p.clientId)} onDone={() => savePlan({ ...p, done: true })} onEdit={() => setModal({ type: 'plan', item: p })} onDel={() => removePlan(p.id)} />)}</Card>
            {plans.some(p => p.done) && <Card title="Realizate" icon={CheckCircle2} iconColor="text-green-600">{plans.filter(p => p.done).map(p => <PlanRow key={p.id} p={p} c={clientById(p.clientId)} done onDel={() => removePlan(p.id)} />)}</Card>}
          </div>
        )}
      </div>

      {modal?.type === 'client' && <ClientModal item={modal.item} me={me} onClose={() => setModal(null)} onSave={(it) => { saveClient(it); setModal(null); }} setCoordCurente={setCoordCurente} />}
      {modal?.type === 'product' && <ProductModal item={modal.item} onClose={() => setModal(null)} onSave={(it) => { saveProduct(it); setModal(null); }} />}
      {modal?.type === 'prodHist' && <ProductHistoryModal product={modal.item} onClose={() => setModal(null)} />}
      {modal?.type === 'report' && <ReportModal item={modal.item} clients={clients} me={me} onClose={() => setModal(null)} onSave={(it) => { saveReport(it); setModal(null); }} />}
      {modal?.type === 'opp' && <OppModal item={modal.item} clients={clients} products={products} me={me} onClose={() => setModal(null)} onSave={(it) => { saveOpp(it); setModal(null); }} />}
      {modal?.type === 'plan' && <PlanModal item={modal.item} clients={clients} me={me} onClose={() => setModal(null)} onSave={(it) => { savePlan(it); setModal(null); }} />}
    </div>
  );
}

function Stat({ icon: Icon, label, value, sub, grad }) {
  return (<div className="bg-white rounded-xl border border-green-100 p-4 shadow-sm relative overflow-hidden"><div className={`absolute -right-3 -top-3 w-16 h-16 rounded-full bg-gradient-to-br ${grad} opacity-15`} /><div className="flex items-center justify-between relative"><span className="text-sm text-slate-500">{label}</span><div className={`w-8 h-8 rounded-lg bg-gradient-to-br ${grad} flex items-center justify-center text-white`}><Icon size={16} /></div></div><div className="text-2xl font-bold mt-1">{value}</div><div className="text-xs text-slate-400">{sub}</div></div>);
}
function Card({ title, icon: Icon, iconColor, children }) { return <div className="bg-white rounded-xl border border-green-100 p-4 shadow-sm"><div className="flex items-center gap-2 font-semibold text-slate-700 mb-3"><Icon size={16} className={iconColor} />{title}</div>{children}</div>; }
function Empty({ text }) { return <div className="bg-white/70 rounded-xl border-2 border-dashed border-green-200 p-8 text-center text-slate-400 text-sm">{text}</div>; }
function ExportBar({ rows, cols, title, onPDF, onExcel, onMail, onWhatsApp }) {
  const disabled = !rows || rows.length === 0;
  const btn = "text-xs px-2.5 py-1.5 rounded-lg flex items-center gap-1.5 transition disabled:opacity-40 disabled:cursor-not-allowed";
  return (<div className="flex gap-1.5 flex-wrap"><button disabled={disabled} onClick={() => onPDF(rows, cols, title)} className={btn + " bg-slate-100 text-slate-600 hover:bg-slate-200"}><FileText size={13} /> PDF</button><button disabled={disabled} onClick={() => onExcel(rows, cols, title)} className={btn + " bg-green-50 text-green-700 hover:bg-green-100"}><Download size={13} /> Excel</button><button disabled={disabled} onClick={() => onMail(rows, cols, title)} className={btn + " bg-sky-50 text-sky-700 hover:bg-sky-100"}><Phone size={13} className="rotate-90" /> Mail</button><button disabled={disabled} onClick={() => onWhatsApp(rows, cols, title)} className={btn + " bg-emerald-50 text-emerald-700 hover:bg-emerald-100"}>💬 WhatsApp</button></div>);
}
function DelBtn({ onDelete, size = 15, label }) {
  const [armed, setArmed] = useState(false);
  useEffect(() => { if (!armed) return; const t = setTimeout(() => setArmed(false), 3000); return () => clearTimeout(t); }, [armed]);
  if (armed) return (<button onClick={(e) => { e.stopPropagation(); onDelete(); setArmed(false); }} className="text-xs bg-rose-600 text-white px-2 py-1 rounded-lg hover:bg-rose-700 whitespace-nowrap">Sigur?</button>);
  return label ? <button onClick={(e) => { e.stopPropagation(); setArmed(true); }} className="bg-rose-50 text-rose-600 px-3 py-2 rounded-lg text-sm flex items-center gap-1.5 hover:bg-rose-100"><Trash2 size={16} /> {label}</button> : <button onClick={(e) => { e.stopPropagation(); setArmed(true); }} className="p-1.5 text-slate-400 hover:text-rose-500"><Trash2 size={size} /></button>;
}
function PlanLine({ p, client, late, today, onOpen }) {
  return (<button onClick={onOpen} className={`w-full text-left flex items-center gap-2 text-sm p-2 rounded-lg ${late ? 'bg-rose-50' : today ? 'bg-amber-50' : 'hover:bg-green-50/50'}`}><span className={`text-xs px-1.5 py-0.5 rounded-full ${p.tip === 'Vizită' ? 'bg-teal-100 text-teal-700' : 'bg-purple-100 text-purple-700'}`}>{p.tip}</span><span className="font-medium">{client ? client.nume : '—'}</span><span className="text-slate-400 truncate">{p.nota}</span><span className="ml-auto text-xs text-slate-500">{p.data}</span></button>);
}
function PlanRow({ p, c, late, today, done, onDone, onEdit, onDel }) {
  return (<div className={`flex items-center gap-2 p-2 rounded-lg text-sm ${late ? 'bg-rose-50' : today ? 'bg-amber-50' : ''}`}><span className={`text-xs px-1.5 py-0.5 rounded-full ${p.tip === 'Vizită' ? 'bg-teal-100 text-teal-700' : 'bg-purple-100 text-purple-700'}`}>{p.tip}</span><span className={`font-medium ${done ? 'line-through text-slate-400' : ''}`}>{c ? c.nume : '—'}</span><span className="text-slate-500 truncate">{p.nota}</span><span className="ml-auto text-xs text-slate-400">{p.data}</span>{!done && onDone && <button onClick={onDone} className="p-1 text-slate-400 hover:text-green-600"><CheckCircle2 size={16} /></button>}{onEdit && <button onClick={onEdit} className="p-1 text-slate-400 hover:text-green-600"><Edit2 size={14} /></button>}{onDel && <DelBtn onDelete={onDel} size={14} />}</div>);
}
function Modal({ title, onClose, children, onSave, saveLabel = 'Salvează', wide }) {
  return (<div className="fixed inset-0 bg-black/40 flex items-center justify-center p-3 z-50" onClick={onClose}><div className={`bg-white rounded-2xl w-full ${wide ? 'max-w-3xl' : 'max-w-lg'} max-h-[92vh] overflow-auto shadow-xl`} onClick={e => e.stopPropagation()}><div className="flex justify-between items-center px-4 py-3 border-b sticky top-0 bg-white z-10 rounded-t-2xl"><div className="font-semibold text-slate-700">{title}</div><button onClick={onClose} className="text-slate-400 hover:text-slate-700"><X size={18} /></button></div><div className="p-4 space-y-3">{children}</div><div className="px-4 py-3 border-t flex justify-end gap-2 sticky bottom-0 bg-white rounded-b-2xl"><button onClick={onClose} className="px-3 py-2 text-sm text-slate-600">Anulează</button><button onClick={onSave} className="px-4 py-2 text-sm bg-green-600 hover:bg-green-700 text-white rounded-lg">{saveLabel}</button></div></div></div>);
}
const Field = ({ label, children }) => <div><label className="block text-xs text-slate-500 mb-1">{label}</label>{children}</div>;
const inp = "w-full border border-slate-200 rounded-lg px-2.5 py-2 text-sm focus:ring-2 focus:ring-green-200 outline-none";

function ClientModal({ item, me, onClose, onSave, setCoordCurente }) {
  const [f, setF] = useState(item || { id: uid(), nume: '', categorie: 'fermier_h', status: 'activ', contact: '', telefon: '', email: '', localitate: '', judet: '', lat: '', lng: '', culturi: emptyCulturi(), ultimaVizita: '', notite: '', atcResponsabil: me });
  const setC = (i, k, v) => { const culturi = [...f.culturi]; culturi[i] = { ...culturi[i], [k]: v }; setF({ ...f, culturi }); };
  const delC = (i) => setF({ ...f, culturi: f.culturi.filter((_, idx) => idx !== i) });
  const save = () => { if (!f.nume.trim()) { alert('Numele clientului este obligatoriu.'); return; } onSave(f); };
  return (
    <Modal title={item ? 'Editează client' : 'Creare client'} onClose={onClose} onSave={save}>
      <Field label="Nume / denumire *"><input className={inp} value={f.nume} onChange={e => setF({ ...f, nume: e.target.value })} /></Field>
      <div className="grid grid-cols-2 gap-3"><Field label="Categorie"><select className={inp} value={f.categorie} onChange={e => setF({ ...f, categorie: e.target.value })}>{CATEGORII.map(c => <option key={c.id} value={c.id}>{c.label}</option>)}</select></Field><Field label="Status"><select className={inp} value={f.status} onChange={e => setF({ ...f, status: e.target.value })}><option value="activ">Activ</option><option value="inactiv">Inactiv</option></select></Field></div>
      <Field label="ATC responsabil"><input className={inp} value={f.atcResponsabil} onChange={e => setF({ ...f, atcResponsabil: e.target.value })} /></Field>
      <div className="grid grid-cols-2 gap-3"><Field label="Persoană contact"><input className={inp} value={f.contact} onChange={e => setF({ ...f, contact: e.target.value })} /></Field><Field label="Telefon"><input className={inp} value={f.telefon} onChange={e => setF({ ...f, telefon: e.target.value })} /></Field></div>
      <div className="grid grid-cols-2 gap-3"><Field label="Localitate"><input className={inp} value={f.localitate} onChange={e => setF({ ...f, localitate: e.target.value })} /></Field><Field label="Județ"><input className={inp} value={f.judet} onChange={e => setF({ ...f, judet: e.target.value })} /></Field></div>
      <Field label="Email"><input className={inp} value={f.email} onChange={e => setF({ ...f, email: e.target.value })} /></Field>
      <Field label="Ultima vizită"><input type="date" className={inp} value={f.ultimaVizita} onChange={e => setF({ ...f, ultimaVizita: e.target.value })} /></Field>
      <div className="bg-green-50/60 rounded-lg p-3 border border-green-100">
        <div className="flex items-center justify-between mb-2"><div className="text-xs text-slate-500 flex items-center gap-1"><MapPin size={13} className="text-green-500" /> Coordonate GPS</div><button onClick={() => setCoordCurente((lat, lng) => setF(prev => ({ ...prev, lat, lng })))} className="text-xs bg-green-600 text-white px-2.5 py-1.5 rounded-lg flex items-center gap-1 hover:bg-green-700"><Crosshair size={13} /> Locația mea acum</button></div>
        <div className="grid grid-cols-2 gap-3"><Field label="Latitudine"><input className={inp} value={f.lat} onChange={e => setF({ ...f, lat: e.target.value })} /></Field><Field label="Longitudine"><input className={inp} value={f.lng} onChange={e => setF({ ...f, lng: e.target.value })} /></Field></div>
      </div>
      <div>
        <div className="text-xs text-slate-500 mb-1 flex items-center gap-1"><Sprout size={13} className="text-green-500" /> Culturi și suprafețe (ha)</div>
        {f.culturi.map((c, i) => (<div key={i} className="flex gap-2 mb-1.5 items-center"><input className="flex-1 border border-slate-200 rounded-lg px-2 py-1.5 text-sm" placeholder={`Cultură ${i + 1}`} value={c.cultura} onChange={e => setC(i, 'cultura', e.target.value)} /><input className="w-24 border border-slate-200 rounded-lg px-2 py-1.5 text-sm" placeholder="ha" value={c.suprafata} onChange={e => setC(i, 'suprafata', e.target.value)} /><button onClick={() => delC(i)} className="p-1 text-slate-300 hover:text-rose-500"><X size={15} /></button></div>))}
        <button onClick={() => setF({ ...f, culturi: [...f.culturi, { cultura: '', suprafata: '' }] })} className="text-xs text-green-600 flex items-center gap-1 mt-1"><Plus size={12} /> Adaugă rând cultură</button>
      </div>
      <Field label="Observații"><textarea className={inp} rows={2} value={f.notite} onChange={e => setF({ ...f, notite: e.target.value })} /></Field>
    </Modal>
  );
}

function ProductModal({ item, onClose, onSave }) {
  const [f, setF] = useState(item || { id: uid(), creatLa: todayStr(), cod: '', nume: '', categorie: '', compozitie: '', doze: '', culturi: '', ambalaj: '', densitate: '', um: 'L', pret: '', tva: TVA_DEFAULT, istoricPret: [] });
  const [pretNou, setPretNou] = useState('');
  const save = () => { if (!f.nume.trim()) { alert('Denumirea produsului este obligatorie.'); return; } onSave(f); };
  const aplicaPretNou = () => { const p = Number(pretNou); if (!p) { alert('Introdu un preț valid.'); return; } setF({ ...f, pret: String(p), istoricPret: [...(f.istoricPret || []), { data: todayStr(), pret: p }] }); setPretNou(''); };
  const hist = (f.istoricPret || []).slice().sort((a, b) => a.data.localeCompare(b.data));
  return (
    <Modal title={item ? 'Editează produs' : 'Produs nou'} onClose={onClose} onSave={save} wide>
      <div className="grid grid-cols-2 md:grid-cols-3 gap-3"><Field label="Cod / SKU"><input className={inp} value={f.cod} onChange={e => setF({ ...f, cod: e.target.value })} /></Field><Field label="Categorie"><input className={inp} value={f.categorie} onChange={e => setF({ ...f, categorie: e.target.value })} /></Field><Field label="UM"><input className={inp} value={f.um} onChange={e => setF({ ...f, um: e.target.value })} /></Field></div>
      <Field label="Denumire produs *"><input className={inp} value={f.nume} onChange={e => setF({ ...f, nume: e.target.value })} /></Field>
      <Field label="Compoziție"><textarea className={inp} rows={2} value={f.compozitie} onChange={e => setF({ ...f, compozitie: e.target.value })} /></Field>
      <div className="grid grid-cols-2 gap-3"><Field label="Doze recomandate"><textarea className={inp} rows={2} value={f.doze} onChange={e => setF({ ...f, doze: e.target.value })} /></Field><Field label="Culturi recomandate"><textarea className={inp} rows={2} value={f.culturi} onChange={e => setF({ ...f, culturi: e.target.value })} /></Field></div>
      <div className="grid grid-cols-2 md:grid-cols-3 gap-3"><Field label="Ambalaj"><input className={inp} value={f.ambalaj} onChange={e => setF({ ...f, ambalaj: e.target.value })} /></Field><Field label="Densitate"><input className={inp} value={f.densitate} onChange={e => setF({ ...f, densitate: e.target.value })} /></Field><Field label="Cotă TVA %"><select className={inp} value={f.tva} onChange={e => setF({ ...f, tva: e.target.value })}>{COTE_TVA.map(c => <option key={c} value={c}>{c}%</option>)}</select></Field></div>
      <div className="bg-green-50/60 rounded-lg p-3 border border-green-100">
        <div className="text-xs text-slate-500 mb-2 flex items-center gap-1"><History size={13} className="text-green-500" /> Preț de listă și istoric</div>
        <div className="flex items-end gap-2 flex-wrap"><Field label="Preț curent (lei)"><input className={inp + ' w-32'} value={f.pret} onChange={e => setF({ ...f, pret: e.target.value })} /></Field><div className="flex items-end gap-1"><Field label="Preț nou (azi)"><input className={inp + ' w-32'} value={pretNou} onChange={e => setPretNou(e.target.value)} /></Field><button onClick={aplicaPretNou} className="text-xs bg-green-600 text-white px-2.5 py-2 rounded-lg hover:bg-green-700 flex items-center gap-1 mb-px"><Plus size={13} /> Adaugă</button></div></div>
        {hist.length > 0 && <div className="mt-2 flex flex-wrap gap-1.5">{hist.map((h, i) => <span key={i} className="text-xs bg-white border border-green-200 rounded-full px-2 py-0.5 text-slate-600">{h.data}: <b>{fmt(h.pret)}</b></span>)}</div>}
      </div>
    </Modal>
  );
}

function ProductHistoryModal({ product, onClose }) {
  const hist = (product.istoricPret || []).slice().sort((a, b) => a.data.localeCompare(b.data));
  const months = []; const now = new Date();
  for (let i = 11; i >= 0; i--) { const d = new Date(now.getFullYear(), now.getMonth() - i, 1); months.push(d.toISOString().slice(0, 7)); }
  let last = null;
  const serie = months.map(m => { const inLuna = hist.filter(h => h.data.slice(0, 7) <= m); if (inLuna.length) last = inLuna[inLuna.length - 1].pret; return { luna: m.slice(2).replace('-', '/'), pret: last }; }).filter(s => s.pret != null);
  return (
    <Modal title={`Evoluție preț — ${product.nume}`} onClose={onClose} onSave={onClose} saveLabel="Închide" wide>
      {serie.length < 1 ? <div className="text-sm text-slate-400">Nu există istoric de preț.</div> : <>
        <div className="h-64"><ResponsiveContainer width="100%" height="100%"><LineChart data={serie} margin={{ top: 10, right: 20, left: 0, bottom: 0 }}><CartesianGrid strokeDasharray="3 3" stroke="#e5e7eb" /><XAxis dataKey="luna" tick={{ fontSize: 11 }} /><YAxis tick={{ fontSize: 11 }} domain={['auto', 'auto']} /><Tooltip formatter={(val) => [fmt(val) + ' lei', 'Preț']} /><Line type="monotone" dataKey="pret" stroke="#16a34a" strokeWidth={2.5} dot={{ r: 3, fill: '#16a34a' }} /></LineChart></ResponsiveContainer></div>
        <div className="flex flex-wrap gap-1.5 mt-2">{hist.map((h, i) => <span key={i} className="text-xs bg-green-50 border border-green-200 rounded-full px-2 py-0.5 text-slate-600">{h.data}: <b>{fmt(h.pret)}</b> lei</span>)}</div>
      </>}
    </Modal>
  );
}

function ReportModal({ item, clients, me, onClose, onSave }) {
  const [f, setF] = useState(item || { id: uid(), clientId: clients[0]?.id || '', data: todayStr(), notite: '', foto: [], atc: me });
  const save = () => { if (!f.clientId) { alert('Selectează un client.'); return; } onSave(f); };
  return (
    <Modal title={item ? 'Editează vizită' : 'Raport de vizită'} onClose={onClose} onSave={save}>
      {clients.length === 0 ? <div className="text-sm text-amber-600">Adaugă întâi un client.</div> : <>
        <div className="grid grid-cols-2 gap-3"><Field label="Client"><select className={inp} value={f.clientId} onChange={e => setF({ ...f, clientId: e.target.value })}>{clients.map(c => <option key={c.id} value={c.id}>{c.nume}</option>)}</select></Field><Field label="Data vizitei"><input type="date" className={inp} value={f.data} onChange={e => setF({ ...f, data: e.target.value })} /></Field></div>
        <Field label="ATC"><input className={inp} value={f.atc} onChange={e => setF({ ...f, atc: e.target.value })} /></Field>
        <Field label="Notițe vizită"><textarea className={inp} rows={3} value={f.notite} onChange={e => setF({ ...f, notite: e.target.value })} /></Field>
        <div className="text-xs text-slate-400">La salvare, „Ultima vizită" a clientului se actualizează automat.</div>
      </>}
    </Modal>
  );
}

function OppModal({ item, clients, products, me, onClose, onSave }) {
  const [f, setF] = useState(item || { id: uid(), nrOferta: '', clientId: clients[0]?.id || '', titlu: '', stadiu: 'Ofertare', ofStatus: 'Draft', data: todayStr(), valabil: addDays(todayStr(), VALAB_DEFAULT), dataInchidere: '', linii: [{ id: uid(), produs: '', cod: '', ambalaj: '', um: '', cantitate: '', pretLista: '', reducere: '', redTip: 'pct', tva: TVA_DEFAULT }], notite: '', atc: me });
  const setL = (id, k, v) => setF({ ...f, linii: f.linii.map(l => l.id === id ? { ...l, [k]: v } : l) });
  const pickProduct = (id, val) => { const p = (products || []).find(x => prodLabel(x) === val || x.nume === val); setF({ ...f, linii: f.linii.map(l => l.id === id ? { ...l, produs: p ? p.nume : val, cod: p ? p.cod : (l.cod || ''), ambalaj: p ? p.ambalaj : (l.ambalaj || ''), um: p ? p.um : (l.um || ''), pretLista: p ? p.pret : l.pretLista, tva: p ? p.tva : l.tva } : l) }); };
  const addL = () => setF({ ...f, linii: [...f.linii, { id: uid(), produs: '', cod: '', ambalaj: '', um: '', cantitate: '', pretLista: '', reducere: '', redTip: 'pct', tva: TVA_DEFAULT }] });
  const delL = (id) => setF({ ...f, linii: f.linii.filter(l => l.id !== id) });
  const tot = calcOferta(f.linii);
  const save = () => { if (!f.clientId) { alert('Selectează un client.'); return; } onSave(f); };
  return (
    <Modal title={item ? `Editează ${f.nrOferta || 'ofertă'}` : 'Ofertă nouă'} onClose={onClose} onSave={save} wide>
      {clients.length === 0 ? <div className="text-sm text-amber-600">Adaugă întâi un client.</div> : <>
        {!item && <div className="text-xs text-slate-400">Numărul ofertei se generează automat la salvare.</div>}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
          <Field label="Client"><select className={inp} value={f.clientId} onChange={e => setF({ ...f, clientId: e.target.value })}>{clients.map(c => <option key={c.id} value={c.id}>{c.nume}</option>)}</select></Field>
          <Field label="Titlu / referință"><input className={inp} value={f.titlu} onChange={e => setF({ ...f, titlu: e.target.value })} /></Field>
          <Field label="Data ofertei"><input type="date" className={inp} value={f.data} onChange={e => setF({ ...f, data: e.target.value })} /></Field>
          <Field label="Valabilă până"><input type="date" className={inp} value={f.valabil} onChange={e => setF({ ...f, valabil: e.target.value })} /></Field>
        </div>
        <div className="grid grid-cols-3 gap-3">
          <Field label="Status ofertă"><select className={inp} value={f.ofStatus} onChange={e => setF({ ...f, ofStatus: e.target.value })}>{OFERTA_STATUS.map(s => <option key={s}>{s}</option>)}</select></Field>
          <Field label="Stadiu comercial"><select className={inp} value={f.stadiu} onChange={e => setF({ ...f, stadiu: e.target.value })}>{STADII.map(s => <option key={s}>{s}</option>)}</select></Field>
          <Field label="ATC"><input className={inp} value={f.atc} onChange={e => setF({ ...f, atc: e.target.value })} /></Field>
        </div>
        <div className="border border-green-100 rounded-lg overflow-hidden">
          <div className="px-3 py-2 text-xs font-semibold text-green-800 flex items-center gap-1.5" style={{ background: 'linear-gradient(90deg,#dcfce7,#ecfccb)' }}><ShoppingCart size={13} /> Linii ofertă</div>
          <div className="overflow-x-auto"><table className="w-full text-sm min-w-[820px]">
            <thead className="text-xs text-slate-500 bg-slate-50/60"><tr><th className="px-2 py-1.5 text-left font-normal">Cod</th><th className="px-2 py-1.5 text-left font-normal">Produs</th><th className="px-2 py-1.5 text-left font-normal">Ambalaj</th><th className="px-2 py-1.5 text-right font-normal">Cant.</th><th className="px-2 py-1.5 text-right font-normal">Preț listă</th><th className="px-2 py-1.5 text-right font-normal">Reducere</th><th className="px-2 py-1.5 text-right font-normal">TVA %</th><th className="px-2 py-1.5 text-right font-normal">Fără TVA</th><th className="px-2 py-1.5 text-right font-normal">Cu TVA</th><th className="px-1"></th></tr></thead>
            <tbody>{f.linii.map(l => { const c = calcLinie(l); return (
              <tr key={l.id} className="border-t">
                <td className="px-1 py-1"><span className="text-xs font-mono text-slate-500 whitespace-nowrap">{l.cod || '—'}</span></td>
                <td className="px-1 py-1"><input list={`prod-${l.id}`} className="w-full border border-slate-200 rounded px-2 py-1 text-sm min-w-[150px]" placeholder="Produs" value={l.produs} onChange={e => pickProduct(l.id, e.target.value)} />{products?.length > 0 && <datalist id={`prod-${l.id}`}>{products.map(p => <option key={p.id} value={prodLabel(p)} />)}</datalist>}</td>
                <td className="px-1 py-1"><span className="text-xs text-slate-500 whitespace-nowrap">{l.ambalaj || '—'}</span></td>
                <td className="px-1 py-1"><input className="w-16 border border-slate-200 rounded px-1.5 py-1 text-sm text-right" value={l.cantitate} onChange={e => setL(l.id, 'cantitate', e.target.value)} /></td>
                <td className="px-1 py-1"><input className="w-20 border border-slate-200 rounded px-1.5 py-1 text-sm text-right" value={l.pretLista} onChange={e => setL(l.id, 'pretLista', e.target.value)} /></td>
                <td className="px-1 py-1"><div className="flex gap-1 justify-end"><input className="w-12 border border-slate-200 rounded px-1.5 py-1 text-sm text-right" value={l.reducere} onChange={e => setL(l.id, 'reducere', e.target.value)} /><select className="border border-slate-200 rounded px-1 py-1 text-xs" value={l.redTip || 'pct'} onChange={e => setL(l.id, 'redTip', e.target.value)}><option value="pct">%</option><option value="lei">lei</option></select></div></td>
                <td className="px-1 py-1"><input className="w-14 border border-slate-200 rounded px-1.5 py-1 text-sm text-right" value={l.tva} onChange={e => setL(l.id, 'tva', e.target.value)} /></td>
                <td className="px-2 py-1 text-right text-slate-600 whitespace-nowrap">{fmt(c.valFaraTVA)}</td>
                <td className="px-2 py-1 text-right font-semibold text-green-700 whitespace-nowrap">{fmt(c.valCuTVA)}</td>
                <td className="px-1 py-1"><button onClick={() => delL(l.id)} className="p-1 text-slate-300 hover:text-rose-500"><Trash2 size={14} /></button></td>
              </tr>); })}</tbody>
          </table></div>
          <div className="px-3 py-2 border-t flex justify-between items-center bg-green-50/40"><button onClick={addL} className="text-xs text-green-600 flex items-center gap-1"><Plus size={13} /> Adaugă linie</button><div className="text-sm flex gap-4"><span className="text-slate-500">Fără TVA: <b className="text-slate-700">{fmt(tot.faraTVA)}</b></span><span className="text-slate-500">TVA: <b className="text-slate-700">{fmt(tot.tva)}</b></span><span className="text-slate-500">Total: <b className="text-green-700">{fmt(tot.cuTVA)}</b></span></div></div>
        </div>
        <Field label="Notițe"><textarea className={inp} rows={2} value={f.notite} onChange={e => setF({ ...f, notite: e.target.value })} /></Field>
      </>}
    </Modal>
  );
}

function PlanModal({ item, clients, me, onClose, onSave }) {
  const [f, setF] = useState(item || { id: uid(), clientId: clients[0]?.id || '', tip: 'Vizită', data: todayStr(), nota: '', done: false, atc: me });
  const save = () => { if (!f.clientId) { alert('Selectează un client.'); return; } onSave(f); };
  return (
    <Modal title={item ? 'Editează activitate' : 'Activitate planificată'} onClose={onClose} onSave={save}>
      {clients.length === 0 ? <div className="text-sm text-amber-600">Adaugă întâi un client.</div> : <>
        <div className="grid grid-cols-2 gap-3"><Field label="Tip"><select className={inp} value={f.tip} onChange={e => setF({ ...f, tip: e.target.value })}><option>Vizită</option><option>Ofertă</option></select></Field><Field label="Data planificată"><input type="date" className={inp} value={f.data} onChange={e => setF({ ...f, data: e.target.value })} /></Field></div>
        <Field label="Client / distribuitor"><select className={inp} value={f.clientId} onChange={e => setF({ ...f, clientId: e.target.value })}>{clients.map(c => <option key={c.id} value={c.id}>{c.nume}</option>)}</select></Field>
        <Field label="ATC"><input className={inp} value={f.atc} onChange={e => setF({ ...f, atc: e.target.value })} /></Field>
        <Field label="Notă / scop"><textarea className={inp} rows={2} value={f.nota} onChange={e => setF({ ...f, nota: e.target.value })} /></Field>
      </>}
    </Modal>
  );
}
