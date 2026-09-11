'use client';

import { useState, useEffect, useMemo } from 'react';
import { useRouter } from 'next/navigation';
import * as XLSX from 'xlsx';
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
} from 'recharts';
import {
  Building2, Receipt, LayoutDashboard, Settings as SettingsIcon,
  Plus, X, Pencil, Trash2, TrendingUp, TrendingDown, Landmark,
  ShieldCheck, Download, AlertTriangle, Home, ArrowRight,
  CalendarDays, Upload, CheckCircle2, LogOut,
} from 'lucide-react';
import { createClient } from '@/lib/supabase/client';

/* ---------------------------------------------------------------- */
/* supabase row <-> app object mapping                                */
/* ---------------------------------------------------------------- */

function propRowToObj(r) {
  return {
    id: r.id, name: r.name, type: r.type, address: r.address || '',
    purchaseValue: Number(r.purchase_value) || 0,
    currentValue: r.current_value === null || r.current_value === undefined ? null : Number(r.current_value),
    purchaseDate: r.purchase_date || '', status: r.status || 'Ativo', notes: r.notes || '',
    loans: r.loans || [], insurances: r.insurances || [],
  };
}
function propObjToRow(o, ownerId) {
  return {
    id: o.id, owner_id: ownerId, name: o.name, type: o.type, address: o.address,
    purchase_value: o.purchaseValue, current_value: o.currentValue,
    purchase_date: o.purchaseDate || null, status: o.status, notes: o.notes,
    loans: o.loans || [], insurances: o.insurances || [],
  };
}
function txRowToObj(r) {
  return {
    id: r.id, propertyId: r.property_id, type: r.type, category: r.category,
    amount: Number(r.amount) || 0, date: r.date, description: r.description || '',
    recurring: r.recurring || 'none', reservationId: r.reservation_id || undefined,
  };
}
function txObjToRow(o, ownerId) {
  return {
    id: o.id, owner_id: ownerId, property_id: o.propertyId, type: o.type, category: o.category,
    amount: o.amount, date: o.date, description: o.description, recurring: o.recurring,
    reservation_id: o.reservationId || null,
  };
}
function resRowToObj(r) {
  return {
    id: r.id, propertyId: r.property_id, guestName: r.guest_name || '', platform: r.platform || 'Airbnb',
    checkIn: r.check_in || '', checkOut: r.check_out || '', grossAmount: Number(r.gross_amount) || 0,
    platformFee: Number(r.platform_fee) || 0, cleaningFee: Number(r.cleaning_fee) || 0, notes: r.notes || '',
  };
}
function resObjToRow(o, ownerId) {
  return {
    id: o.id, owner_id: ownerId, property_id: o.propertyId, guest_name: o.guestName, platform: o.platform,
    check_in: o.checkIn || null, check_out: o.checkOut || null, gross_amount: o.grossAmount,
    platform_fee: o.platformFee, cleaning_fee: o.cleaningFee, notes: o.notes,
  };
}
function settingsRowToObj(r) {
  return {
    inflationRate: Number(r.inflation_rate) || 2.5,
    customIncomeCategories: r.custom_income_categories || [],
    customExpenseCategories: r.custom_expense_categories || [],
  };
}

/* ---------------------------------------------------------------- */
/* constants                                                          */
/* ---------------------------------------------------------------- */

const PROPERTY_TYPES = [
  { id: 'personal', label: 'Casa pessoal' },
  { id: 'long_term_rental', label: 'Arrendamento longo prazo' },
  { id: 'airbnb', label: 'Airbnb / alojamento local' },
  { id: 'informal_rental', label: 'Arrendamento sem contrato' },
  { id: 'garage_contract', label: 'Garagem com contrato' },
  { id: 'garage_no_contract', label: 'Garagem sem contrato' },
  { id: 'other', label: 'Outro' },
];

const PROPERTY_STATUS = ['Ativo', 'Vago', 'Em obras'];
const INSURANCE_TYPES = ['Multirriscos', 'Seguro de vida', 'Crédito habitação', 'Outro'];
const RESERVATION_PLATFORMS = ['Airbnb', 'Booking.com', 'Direto', 'Outro'];

const DEFAULT_INCOME_CATEGORIES = ['Renda', 'Airbnb / alojamento local', 'Outras receitas'];
const DEFAULT_EXPENSE_CATEGORIES = [
  'IMI', 'Condomínio', 'EDP (eletricidade)', 'Água', 'Net / internet / TV',
  'Manutenção', 'Impostos', 'Seguro multirriscos', 'Seguro de vida',
  'Empréstimo (prestação)', 'Comissão de angariação de inquilino',
  'Comissão de plataforma (Airbnb/Booking)', 'Limpeza entre estadias',
  'Outras despesas',
];

const CHART_COLORS = ['#A9823C', '#3E6B52', '#5C7A99', '#A23E2E', '#7A5C99', '#4B8B8B', '#C77B3F', '#8B6F4E', '#6B7280'];

const NAV_ITEMS = [
  { id: 'dashboard', label: 'Resumo', icon: LayoutDashboard },
  { id: 'properties', label: 'Imóveis', icon: Building2 },
  { id: 'shortterm', label: 'Alojamento Local', icon: CalendarDays },
  { id: 'transactions', label: 'Movimentos', icon: Receipt },
  { id: 'settings', label: 'Definições', icon: SettingsIcon },
];

/* ---------------------------------------------------------------- */
/* helpers                                                             */
/* ---------------------------------------------------------------- */

const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);

const fmtMoney = (v, decimals = 0) =>
  new Intl.NumberFormat('pt-PT', { style: 'currency', currency: 'EUR', minimumFractionDigits: decimals, maximumFractionDigits: decimals }).format(v || 0);

const fmtPct = (v) => `${(v || 0).toFixed(1)}%`;

const fmtDate = (d) => {
  if (!d) return '—';
  const dt = new Date(d);
  if (isNaN(dt)) return '—';
  return new Intl.DateTimeFormat('pt-PT', { day: '2-digit', month: 'short', year: 'numeric' }).format(dt);
};

const monthKey = (d) => {
  const dt = new Date(d);
  return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}`;
};

const monthLabel = (key) => {
  const [y, m] = key.split('-');
  const dt = new Date(Number(y), Number(m) - 1, 1);
  return new Intl.DateTimeFormat('pt-PT', { month: 'short', year: '2-digit' }).format(dt);
};

function monthlyLoanPayment(amount, ratePct, termYears) {
  const n = Math.round((termYears || 0) * 12);
  if (!n || n <= 0 || !amount) return 0;
  const r = (ratePct || 0) / 100 / 12;
  if (r === 0) return amount / n;
  return (amount * r * Math.pow(1 + r, n)) / (Math.pow(1 + r, n) - 1);
}

function outstandingBalance(amount, ratePct, termYears, startDate) {
  const n = Math.round((termYears || 0) * 12);
  if (!n || n <= 0 || !amount) return 0;
  const r = (ratePct || 0) / 100 / 12;
  const start = new Date(startDate);
  if (isNaN(start)) return amount;
  const now = new Date();
  let elapsed = (now.getFullYear() - start.getFullYear()) * 12 + (now.getMonth() - start.getMonth());
  elapsed = Math.max(0, Math.min(elapsed, n));
  if (elapsed === 0) return amount;
  if (r === 0) return Math.max(0, amount - (amount / n) * elapsed);
  const M = monthlyLoanPayment(amount, ratePct, termYears);
  const bal = amount * Math.pow(1 + r, elapsed) - M * ((Math.pow(1 + r, elapsed) - 1) / r);
  return Math.max(0, bal);
}

function last12MonthKeys() {
  const out = [];
  const now = new Date();
  for (let i = 11; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    out.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`);
  }
  return out;
}

function propertyTypeLabel(id) {
  return PROPERTY_TYPES.find((t) => t.id === id)?.label || id;
}

function nightsBetween(checkIn, checkOut) {
  const a = new Date(checkIn), b = new Date(checkOut);
  if (isNaN(a) || isNaN(b)) return 0;
  return Math.max(0, Math.round((b - a) / 86400000));
}

function normalizeText(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
}

function parseAmountValue(raw) {
  if (typeof raw === 'number') return raw;
  if (raw === null || raw === undefined || raw === '') return NaN;
  let s = String(raw).trim().replace(/[€\s]/g, '');
  if (s.includes(',') && s.includes('.')) s = s.replace(/\./g, '').replace(',', '.');
  else if (s.includes(',')) s = s.replace(',', '.');
  return parseFloat(s);
}

function parseImportDate(raw) {
  if (raw instanceof Date && !isNaN(raw)) return raw.toISOString().slice(0, 10);
  if (typeof raw === 'string') {
    const s = raw.trim();
    let m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
    if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
    m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
    const dt = new Date(s);
    if (!isNaN(dt)) return dt.toISOString().slice(0, 10);
  }
  return null;
}

/* ---------------------------------------------------------------- */
/* small ui building blocks                                           */
/* ---------------------------------------------------------------- */

function Field({ label, children, hint }) {
  return (
    <label className="block mb-3">
      <span className="block text-xs text-stone-500 mb-1">{label}</span>
      {children}
      {hint && <span className="block text-[11px] text-stone-400 mt-1">{hint}</span>}
    </label>
  );
}

const inputCls = 'w-full rounded border border-stone-300 bg-white px-3 py-2 text-sm text-stone-900 focus:outline-none focus:ring-2 focus:ring-emerald-700/40 focus:border-emerald-700';

function Badge({ children, tone = 'neutral' }) {
  const tones = {
    neutral: 'bg-stone-200 text-stone-700',
    income: 'bg-emerald-100 text-emerald-800',
    expense: 'bg-red-100 text-red-800',
    brass: 'bg-amber-100 text-amber-800',
  };
  return <span className={`inline-block rounded-full px-2 py-0.5 text-[11px] font-medium ${tones[tone]}`}>{children}</span>;
}

function EmptyState({ icon: Icon, title, message, actionLabel, onAction }) {
  return (
    <div className="border border-dashed border-stone-300 rounded bg-stone-50 p-10 text-center">
      <Icon className="mx-auto mb-3 text-stone-400" size={28} />
      <p className="font-serif-display text-lg text-stone-700">{title}</p>
      <p className="text-sm text-stone-500 mt-1 max-w-sm mx-auto">{message}</p>
      {actionLabel && (
        <button onClick={onAction} className="mt-4 inline-flex items-center gap-1.5 rounded bg-emerald-800 text-white text-sm px-4 py-2 hover:bg-emerald-900">
          <Plus size={15} /> {actionLabel}
        </button>
      )}
    </div>
  );
}

function Modal({ title, onClose, children, wide }) {
  return (
    <div className="fixed inset-0 z-50 flex items-start md:items-center justify-center bg-stone-900/50 p-3 overflow-y-auto">
      <div className={`bg-white rounded shadow-xl w-full ${wide ? 'max-w-2xl' : 'max-w-md'} my-6`}>
        <div className="flex items-center justify-between border-b border-stone-200 px-5 py-4">
          <h3 className="font-serif-display text-lg text-stone-800">{title}</h3>
          <button onClick={onClose} className="text-stone-400 hover:text-stone-700"><X size={20} /></button>
        </div>
        <div className="px-5 py-4">{children}</div>
      </div>
    </div>
  );
}

function KPICard({ label, value, sub, tone = 'neutral' }) {
  const bar = { neutral: 'border-t-stone-400', income: 'border-t-emerald-700', expense: 'border-t-red-700', brass: 'border-t-amber-600' }[tone];
  return (
    <div className={`bg-white border border-stone-200 border-t-4 ${bar} rounded p-4`}>
      <p className="text-[11px] text-stone-500">{label}</p>
      <p className="font-serif-display text-2xl text-stone-800 mt-1 tabular-nums">{value}</p>
      {sub && <p className="text-[11px] text-stone-400 mt-1">{sub}</p>}
    </div>
  );
}

/* ---------------------------------------------------------------- */
/* main app                                                            */
/* ---------------------------------------------------------------- */

export default function ImobiliarioApp() {
  const router = useRouter();
  const supabase = useMemo(() => createClient(), []);

  const [loading, setLoading] = useState(true);
  const [userId, setUserId] = useState(null);
  const [userEmail, setUserEmail] = useState('');
  const [properties, setProperties] = useState([]);
  const [transactions, setTransactions] = useState([]);
  const [settings, setSettings] = useState({ inflationRate: 2.5, customIncomeCategories: [], customExpenseCategories: [] });

  const [activeTab, setActiveTab] = useState('dashboard');
  const [propertyModal, setPropertyModal] = useState(null); // {mode, data}
  const [txModal, setTxModal] = useState(null);
  const [confirmModal, setConfirmModal] = useState(null); // {title, message, onConfirm}
  const [txFilter, setTxFilter] = useState({ propertyId: 'all', type: 'all', category: 'all' });
  const [reservations, setReservations] = useState([]);
  const [reservationModal, setReservationModal] = useState(null);
  const [importModalOpen, setImportModalOpen] = useState(false);

  useEffect(() => {
    (async () => {
      try {
        const { data: { user } } = await supabase.auth.getUser();
        if (!user) { router.push('/login'); return; }
        setUserId(user.id);
        setUserEmail(user.email || '');

        const [propsRes, txRes, resRes, settingsRes] = await Promise.all([
          supabase.from('imob_properties').select('*').order('created_at', { ascending: false }),
          supabase.from('imob_transactions').select('*').order('date', { ascending: false }),
          supabase.from('imob_reservations').select('*').order('check_in', { ascending: false }),
          supabase.from('imob_settings').select('*').eq('owner_id', user.id).maybeSingle(),
        ]);
        setProperties((propsRes.data || []).map(propRowToObj));
        setTransactions((txRes.data || []).map(txRowToObj));
        setReservations((resRes.data || []).map(resRowToObj));
        if (settingsRes.data) {
          setSettings(settingsRowToObj(settingsRes.data));
        } else {
          await supabase.from('imob_settings').insert({ owner_id: user.id });
        }
      } catch (e) {
        console.error('Erro ao carregar dados', e);
      }
      setLoading(false);
    })();
    // eslint-disable-next-line
  }, []);

  async function persistSettings(next) {
    setSettings(next);
    if (!userId) return;
    const { error } = await supabase.from('imob_settings').upsert({
      owner_id: userId,
      inflation_rate: next.inflationRate,
      custom_income_categories: next.customIncomeCategories,
      custom_expense_categories: next.customExpenseCategories,
    });
    if (error) console.error(error);
  }

  async function handleSignOut() {
    await supabase.auth.signOut();
    router.push('/login');
  }

  const incomeCategories = useMemo(() => [...DEFAULT_INCOME_CATEGORIES, ...settings.customIncomeCategories], [settings]);
  const expenseCategories = useMemo(() => [...DEFAULT_EXPENSE_CATEGORIES, ...settings.customExpenseCategories], [settings]);

  /* ---------- derived data ---------- */

  const monthlySeries = useMemo(() => {
    const keys = last12MonthKeys();
    const map = {};
    keys.forEach((k) => (map[k] = { key: k, label: monthLabel(k), Receitas: 0, Despesas: 0 }));
    transactions.forEach((t) => {
      const k = monthKey(t.date);
      if (map[k]) {
        if (t.type === 'income') map[k].Receitas += Number(t.amount) || 0;
        else map[k].Despesas += Number(t.amount) || 0;
      }
    });
    return keys.map((k) => map[k]);
  }, [transactions]);

  const netSeries = useMemo(
    () => monthlySeries.map((m) => ({ label: m.label, Saldo: m.Receitas - m.Despesas })),
    [monthlySeries]
  );

  const expenseBreakdown = useMemo(() => {
    const now = new Date();
    const yearStart = new Date(now.getFullYear(), 0, 1);
    const map = {};
    transactions
      .filter((t) => t.type === 'expense' && new Date(t.date) >= yearStart)
      .forEach((t) => { map[t.category] = (map[t.category] || 0) + (Number(t.amount) || 0); });
    return Object.entries(map).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [transactions]);

  function propertyIndicators(property) {
    const now = new Date();
    const oneYearAgo = new Date(now); oneYearAgo.setFullYear(now.getFullYear() - 1);
    const propTx = transactions.filter((t) => t.propertyId === property.id);
    const trailing = propTx.filter((t) => new Date(t.date) >= oneYearAgo);
    const incomeTotal = trailing.filter((t) => t.type === 'income').reduce((s, t) => s + (Number(t.amount) || 0), 0);
    const loanPayments = trailing.filter((t) => t.type === 'expense' && t.category === 'Empréstimo (prestação)').reduce((s, t) => s + (Number(t.amount) || 0), 0);
    const expenseExclLoan = trailing.filter((t) => t.type === 'expense' && t.category !== 'Empréstimo (prestação)').reduce((s, t) => s + (Number(t.amount) || 0), 0);
    const baseValue = Number(property.currentValue) || Number(property.purchaseValue) || 0;
    const grossYield = baseValue ? (incomeTotal / baseValue) * 100 : 0;
    const netYield = baseValue ? ((incomeTotal - expenseExclLoan) / baseValue) * 100 : 0;
    const realYield = netYield - (Number(settings.inflationRate) || 0);
    const cashFlowAfterLoan = incomeTotal - expenseExclLoan - loanPayments;
    const totalOutstandingLoans = (property.loans || []).reduce((s, l) => s + outstandingBalance(Number(l.amount), Number(l.rate), Number(l.termYears), l.startDate), 0);
    const totalMonthlyLoanPayment = (property.loans || []).reduce((s, l) => s + monthlyLoanPayment(Number(l.amount), Number(l.rate), Number(l.termYears)), 0);
    const totalAnnualInsurance = (property.insurances || []).reduce((s, i) => s + (Number(i.annualPremium) || 0), 0);
    return { incomeTotal, expenseExclLoan, loanPayments, baseValue, grossYield, netYield, realYield, cashFlowAfterLoan, totalOutstandingLoans, totalMonthlyLoanPayment, totalAnnualInsurance };
  }

  const portfolioTotals = useMemo(() => {
    if (properties.length === 0) return null;
    let income = 0, expense = 0, cashFlow = 0, weightedYield = 0, valueSum = 0;
    properties.forEach((p) => {
      const ind = propertyIndicators(p);
      income += ind.incomeTotal;
      expense += ind.expenseExclLoan + ind.loanPayments;
      cashFlow += ind.cashFlowAfterLoan;
      weightedYield += ind.netYield * ind.baseValue;
      valueSum += ind.baseValue;
    });
    return {
      count: properties.length,
      monthlyIncome: income / 12,
      monthlyExpense: expense / 12,
      monthlyCashFlow: cashFlow / 12,
      portfolioNetYield: valueSum ? weightedYield / valueSum : 0,
    };
  }, [properties, transactions, settings]);

  /* ---------- property CRUD ---------- */

  function openNewProperty() {
    setPropertyModal({
      mode: 'new',
      data: { id: uid(), name: '', type: 'long_term_rental', address: '', purchaseValue: '', currentValue: '', purchaseDate: '', status: 'Ativo', notes: '', loans: [], insurances: [] },
    });
  }
  function openEditProperty(p) { setPropertyModal({ mode: 'edit', data: JSON.parse(JSON.stringify(p)) }); }

  async function saveProperty(data) {
    const clean = {
      ...data,
      purchaseValue: Number(data.purchaseValue) || 0,
      currentValue: data.currentValue === '' ? null : Number(data.currentValue),
      loans: (data.loans || []).map((l) => ({ ...l, amount: Number(l.amount) || 0, rate: Number(l.rate) || 0, termYears: Number(l.termYears) || 0 })),
      insurances: (data.insurances || []).map((i) => ({ ...i, annualPremium: Number(i.annualPremium) || 0 })),
    };
    const { error } = await supabase.from('imob_properties').upsert(propObjToRow(clean, userId));
    if (error) { console.error(error); return; }
    const exists = properties.some((p) => p.id === clean.id);
    setProperties(exists ? properties.map((p) => (p.id === clean.id ? clean : p)) : [clean, ...properties]);
    setPropertyModal(null);
  }

  function deleteProperty(p) {
    setConfirmModal({
      title: 'Eliminar imóvel',
      message: `Eliminar "${p.name}"? Os movimentos associados deixam de estar ligados a um imóvel mas não são apagados.`,
      onConfirm: async () => {
        const { error } = await supabase.from('imob_properties').delete().eq('id', p.id);
        if (!error) setProperties(properties.filter((x) => x.id !== p.id));
        setConfirmModal(null);
      },
    });
  }

  /* ---------- transaction CRUD ---------- */

  function openNewTransaction() {
    if (properties.length === 0) return;
    setTxModal({
      mode: 'new',
      data: { id: uid(), propertyId: properties[0].id, type: 'income', category: incomeCategories[0], amount: '', date: new Date().toISOString().slice(0, 10), description: '', recurring: 'none' },
    });
  }
  function openEditTransaction(t) { setTxModal({ mode: 'edit', data: { ...t } }); }

  async function saveTransaction(data, newCategory) {
    if (newCategory) {
      let nextSettings = settings;
      if (data.type === 'income' && !incomeCategories.includes(newCategory)) {
        nextSettings = { ...settings, customIncomeCategories: [...settings.customIncomeCategories, newCategory] };
      } else if (data.type === 'expense' && !expenseCategories.includes(newCategory)) {
        nextSettings = { ...settings, customExpenseCategories: [...settings.customExpenseCategories, newCategory] };
      }
      if (nextSettings !== settings) await persistSettings(nextSettings);
    }
    const clean = { ...data, amount: Number(data.amount) || 0, category: newCategory || data.category };
    const { error } = await supabase.from('imob_transactions').upsert(txObjToRow(clean, userId));
    if (error) { console.error(error); return; }
    const exists = transactions.some((t) => t.id === clean.id);
    setTransactions(exists ? transactions.map((t) => (t.id === clean.id ? clean : t)) : [clean, ...transactions]);
    setTxModal(null);
  }

  function deleteTransaction(t) {
    setConfirmModal({
      title: 'Eliminar movimento',
      message: `Eliminar este movimento de ${fmtMoney(t.amount)} em ${fmtDate(t.date)}?`,
      onConfirm: async () => {
        const { error } = await supabase.from('imob_transactions').delete().eq('id', t.id);
        if (!error) setTransactions(transactions.filter((x) => x.id !== t.id));
        setConfirmModal(null);
      },
    });
  }

  /* ---------- alojamento local (reservas) CRUD ---------- */

  function openNewReservation() {
    if (properties.length === 0) return;
    setReservationModal({
      mode: 'new',
      data: { id: uid(), propertyId: properties[0].id, guestName: '', platform: 'Airbnb', checkIn: '', checkOut: '', grossAmount: '', platformFee: '', cleaningFee: '', notes: '' },
    });
  }
  function openEditReservation(r) { setReservationModal({ mode: 'edit', data: { ...r } }); }

  async function saveReservation(data) {
    const clean = { ...data, grossAmount: Number(data.grossAmount) || 0, platformFee: Number(data.platformFee) || 0, cleaningFee: Number(data.cleaningFee) || 0 };
    const nights = nightsBetween(clean.checkIn, clean.checkOut);
    const guestLabel = clean.guestName || clean.platform;
    const linked = [];
    if (clean.grossAmount > 0) {
      linked.push({ id: uid(), propertyId: clean.propertyId, type: 'income', category: 'Airbnb / alojamento local', amount: clean.grossAmount, date: clean.checkOut || clean.checkIn, description: `Reserva ${guestLabel} · ${nights} noite${nights !== 1 ? 's' : ''}`, recurring: 'none', reservationId: clean.id });
    }
    if (clean.platformFee > 0) {
      linked.push({ id: uid(), propertyId: clean.propertyId, type: 'expense', category: 'Comissão de plataforma (Airbnb/Booking)', amount: clean.platformFee, date: clean.checkOut || clean.checkIn, description: `Comissão ${clean.platform} · ${guestLabel}`, recurring: 'none', reservationId: clean.id });
    }
    if (clean.cleaningFee > 0) {
      linked.push({ id: uid(), propertyId: clean.propertyId, type: 'expense', category: 'Limpeza entre estadias', amount: clean.cleaningFee, date: clean.checkOut || clean.checkIn, description: `Limpeza · ${guestLabel}`, recurring: 'none', reservationId: clean.id });
    }

    await supabase.from('imob_transactions').delete().eq('reservation_id', clean.id);
    if (linked.length) {
      const { error: txErr } = await supabase.from('imob_transactions').insert(linked.map((t) => txObjToRow(t, userId)));
      if (txErr) console.error(txErr);
    }
    const { error } = await supabase.from('imob_reservations').upsert(resObjToRow(clean, userId));
    if (error) { console.error(error); return; }

    setTransactions([...transactions.filter((t) => t.reservationId !== clean.id), ...linked]);
    const exists = reservations.some((r) => r.id === clean.id);
    setReservations(exists ? reservations.map((r) => (r.id === clean.id ? clean : r)) : [clean, ...reservations]);
    setReservationModal(null);
  }

  function deleteReservation(r) {
    setConfirmModal({
      title: 'Eliminar reserva',
      message: `Eliminar a reserva de ${r.guestName || r.platform}? Os movimentos gerados por esta reserva também são removidos.`,
      onConfirm: async () => {
        await supabase.from('imob_transactions').delete().eq('reservation_id', r.id);
        const { error } = await supabase.from('imob_reservations').delete().eq('id', r.id);
        if (!error) {
          setTransactions(transactions.filter((t) => t.reservationId !== r.id));
          setReservations(reservations.filter((x) => x.id !== r.id));
        }
        setConfirmModal(null);
      },
    });
  }

  /* ---------- importação de excel ---------- */

  async function importTransactions(newTransactions, newCategories) {
    if (newCategories.income.length || newCategories.expense.length) {
      await persistSettings({
        ...settings,
        customIncomeCategories: [...new Set([...settings.customIncomeCategories, ...newCategories.income])],
        customExpenseCategories: [...new Set([...settings.customExpenseCategories, ...newCategories.expense])],
      });
    }
    if (newTransactions.length) {
      const { error } = await supabase.from('imob_transactions').insert(newTransactions.map((t) => txObjToRow(t, userId)));
      if (error) { console.error(error); return; }
    }
    setTransactions([...newTransactions, ...transactions]);
  }

  /* ---------- render ---------- */

  return (
    <div className="min-h-screen bg-stone-100 text-stone-900" style={{ fontFamily: "'IBM Plex Sans', ui-sans-serif, sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600;9..144,700&family=IBM+Plex+Sans:wght@400;500;600;700&display=swap');
        .font-serif-display { font-family: 'Fraunces', serif; }
        .tabular-nums { font-variant-numeric: tabular-nums; }
      `}</style>

      <div className="flex flex-col md:flex-row min-h-screen">
        <nav className="md:w-56 bg-emerald-950 text-emerald-50 flex flex-col shrink-0">
          <div className="p-5 border-b border-emerald-800">
            <h1 className="font-serif-display text-xl font-semibold leading-tight">Gestão Imobiliária</h1>
            <p className="text-[11px] text-emerald-300 mt-1">carteira &amp; rentabilidade</p>
          </div>
          <div className="flex md:flex-col overflow-x-auto md:overflow-visible">
            {NAV_ITEMS.map((item) => (
              <button
                key={item.id}
                onClick={() => setActiveTab(item.id)}
                className={`flex items-center gap-2.5 px-5 py-3 text-sm whitespace-nowrap border-b-2 md:border-b-0 md:border-l-4 transition-colors text-left ${
                  activeTab === item.id ? 'bg-emerald-900 border-amber-400 text-white' : 'border-transparent text-emerald-300 hover:bg-emerald-900/60'
                }`}
              >
                <item.icon size={17} /> {item.label}
              </button>
            ))}
          </div>
          <div className="mt-auto p-5 hidden md:block text-[11px] text-emerald-400 border-t border-emerald-800">
            <p className="truncate mb-2">{userEmail}</p>
            <p className="mb-3">{properties.length} imóve{properties.length === 1 ? 'l' : 'is'} · {transactions.length} movimentos</p>
            <button onClick={handleSignOut} className="inline-flex items-center gap-1.5 text-emerald-300 hover:text-white">
              <LogOut size={13} /> Sair
            </button>
          </div>
        </nav>

        <main className="flex-1 p-4 md:p-8 max-w-6xl mx-auto w-full">
          {loading ? (
            <p className="text-stone-500 text-sm">A carregar…</p>
          ) : activeTab === 'dashboard' ? (
            <Dashboard
              properties={properties} portfolioTotals={portfolioTotals} monthlySeries={monthlySeries}
              netSeries={netSeries} expenseBreakdown={expenseBreakdown} settings={settings}
              propertyIndicators={propertyIndicators} onNewProperty={openNewProperty}
            />
          ) : activeTab === 'properties' ? (
            <PropertiesView
              properties={properties} propertyIndicators={propertyIndicators}
              onNew={openNewProperty} onEdit={openEditProperty} onDelete={deleteProperty}
              onViewTx={(p) => { setTxFilter({ propertyId: p.id, type: 'all', category: 'all' }); setActiveTab('transactions'); }}
            />
          ) : activeTab === 'shortterm' ? (
            <ShortTermView
              properties={properties} reservations={reservations} transactions={transactions}
              onNew={openNewReservation} onEdit={openEditReservation} onDelete={deleteReservation}
            />
          ) : activeTab === 'transactions' ? (
            <TransactionsView
              transactions={transactions} properties={properties} filter={txFilter} setFilter={setTxFilter}
              onNew={openNewTransaction} onEdit={openEditTransaction} onDelete={deleteTransaction}
              onImportClick={() => setImportModalOpen(true)}
            />
          ) : (
            <SettingsView
              settings={settings} persistSettings={persistSettings}
              incomeCategories={incomeCategories} expenseCategories={expenseCategories}
              properties={properties} transactions={transactions} reservations={reservations}
              onClearAll={() => setConfirmModal({
                title: 'Limpar todos os dados',
                message: 'Esta ação apaga permanentemente todos os imóveis, movimentos e reservas guardados. Não é possível desfazer.',
                onConfirm: async () => {
                  await Promise.all([
                    supabase.from('imob_transactions').delete().eq('owner_id', userId),
                    supabase.from('imob_reservations').delete().eq('owner_id', userId),
                    supabase.from('imob_properties').delete().eq('owner_id', userId),
                  ]);
                  setProperties([]); setTransactions([]); setReservations([]);
                  setConfirmModal(null);
                },
              })}
            />
          )}
        </main>
      </div>

      {propertyModal && (
        <PropertyFormModal
          modal={propertyModal} onClose={() => setPropertyModal(null)} onSave={saveProperty}
        />
      )}
      {txModal && (
        <TransactionFormModal
          modal={txModal} onClose={() => setTxModal(null)} onSave={saveTransaction}
          properties={properties} incomeCategories={incomeCategories} expenseCategories={expenseCategories}
        />
      )}
      {reservationModal && (
        <ReservationFormModal modal={reservationModal} onClose={() => setReservationModal(null)} onSave={saveReservation} properties={properties} />
      )}
      {importModalOpen && (
        <ImportExcelModal
          properties={properties} incomeCategories={incomeCategories} expenseCategories={expenseCategories}
          onClose={() => setImportModalOpen(false)} onImport={importTransactions}
        />
      )}
      {confirmModal && (
        <Modal title={confirmModal.title} onClose={() => setConfirmModal(null)}>
          <p className="text-sm text-stone-600">{confirmModal.message}</p>
          <div className="flex justify-end gap-2 mt-5">
            <button onClick={() => setConfirmModal(null)} className="px-4 py-2 text-sm rounded border border-stone-300 text-stone-600 hover:bg-stone-50">Cancelar</button>
            <button onClick={confirmModal.onConfirm} className="px-4 py-2 text-sm rounded bg-red-700 text-white hover:bg-red-800">Eliminar</button>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ---------------------------------------------------------------- */
/* dashboard                                                           */
/* ---------------------------------------------------------------- */

function Dashboard({ properties, portfolioTotals, monthlySeries, netSeries, expenseBreakdown, settings, propertyIndicators, onNewProperty }) {
  const hasTx = monthlySeries.some((m) => m.Receitas || m.Despesas);

  if (properties.length === 0) {
    return (
      <div>
        <h2 className="font-serif-display text-2xl text-stone-800 mb-1">Resumo</h2>
        <p className="text-sm text-stone-500 mb-6">A vista geral da carteira aparece assim que houver imóveis registados.</p>
        <EmptyState icon={Building2} title="Ainda não há imóveis" message="Comece por adicionar a casa pessoal, um arrendamento, o Airbnb ou uma garagem." actionLabel="Adicionar imóvel" onAction={onNewProperty} />
      </div>
    );
  }

  return (
    <div>
      <h2 className="font-serif-display text-2xl text-stone-800 mb-1">Resumo</h2>
      <p className="text-sm text-stone-500 mb-6">Últimos 12 meses · inflação considerada: {fmtPct(settings.inflationRate)}</p>

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 mb-6">
        <KPICard label="Imóveis na carteira" value={portfolioTotals.count} tone="neutral" />
        <KPICard label="Receita mensal média" value={fmtMoney(portfolioTotals.monthlyIncome)} tone="income" />
        <KPICard label="Despesa mensal média" value={fmtMoney(portfolioTotals.monthlyExpense)} tone="expense" />
        <KPICard
          label="Cash flow mensal médio" value={fmtMoney(portfolioTotals.monthlyCashFlow)}
          tone={portfolioTotals.monthlyCashFlow >= 0 ? 'income' : 'expense'}
          sub={`Rentabilidade líquida da carteira: ${fmtPct(portfolioTotals.portfolioNetYield)}`}
        />
      </div>

      {!hasTx ? (
        <EmptyState icon={Receipt} title="Ainda sem movimentos" message="Os gráficos e indicadores aparecem assim que registar rendas e despesas em Movimentos." />
      ) : (
        <>
          <div className="grid md:grid-cols-2 gap-4 mb-4">
            <div className="bg-white border border-stone-200 rounded p-4">
              <p className="text-sm font-medium text-stone-700 mb-3">Receitas vs. despesas por mês</p>
              <div className="h-64">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={monthlySeries} margin={{ left: -10 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#E7E2D3" vertical={false} />
                    <XAxis dataKey="label" tick={{ fontSize: 11, fill: '#78716c' }} axisLine={{ stroke: '#D6D0BF' }} tickLine={false} />
                    <YAxis tick={{ fontSize: 11, fill: '#78716c' }} axisLine={false} tickLine={false} width={60} tickFormatter={(v) => fmtMoney(v)} />
                    <Tooltip formatter={(v) => fmtMoney(v)} contentStyle={{ fontSize: 12, borderRadius: 4 }} />
                    <Legend wrapperStyle={{ fontSize: 12 }} />
                    <Bar dataKey="Receitas" fill="#3E6B52" radius={[2, 2, 0, 0]} />
                    <Bar dataKey="Despesas" fill="#A23E2E" radius={[2, 2, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>

            <div className="bg-white border border-stone-200 rounded p-4">
              <p className="text-sm font-medium text-stone-700 mb-3">Despesas por categoria (ano corrente)</p>
              <div className="h-64">
                {expenseBreakdown.length === 0 ? (
                  <p className="text-sm text-stone-400 flex items-center justify-center h-full">Sem despesas registadas este ano.</p>
                ) : (
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie data={expenseBreakdown} dataKey="value" nameKey="name" innerRadius={45} outerRadius={80} paddingAngle={2}>
                        {expenseBreakdown.map((entry, i) => <Cell key={entry.name} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}
                      </Pie>
                      <Tooltip formatter={(v) => fmtMoney(v)} contentStyle={{ fontSize: 12, borderRadius: 4 }} />
                      <Legend wrapperStyle={{ fontSize: 11 }} layout="vertical" align="right" verticalAlign="middle" />
                    </PieChart>
                  </ResponsiveContainer>
                )}
              </div>
            </div>
          </div>

          <div className="bg-white border border-stone-200 rounded p-4 mb-6">
            <p className="text-sm font-medium text-stone-700 mb-3">Saldo líquido mensal</p>
            <div className="h-56">
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={netSeries} margin={{ left: -10 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#E7E2D3" vertical={false} />
                  <XAxis dataKey="label" tick={{ fontSize: 11, fill: '#78716c' }} axisLine={{ stroke: '#D6D0BF' }} tickLine={false} />
                  <YAxis tick={{ fontSize: 11, fill: '#78716c' }} axisLine={false} tickLine={false} width={60} tickFormatter={(v) => fmtMoney(v)} />
                  <Tooltip formatter={(v) => fmtMoney(v)} contentStyle={{ fontSize: 12, borderRadius: 4 }} />
                  <Line type="monotone" dataKey="Saldo" stroke="#A9823C" strokeWidth={2.5} dot={{ r: 3 }} />
                </LineChart>
              </ResponsiveContainer>
            </div>
          </div>
        </>
      )}

      <div className="bg-white border border-stone-200 rounded overflow-hidden">
        <p className="text-sm font-medium text-stone-700 px-4 pt-4 pb-2">Indicadores por imóvel (últimos 12 meses)</p>
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="text-left text-[11px] text-stone-500 border-t border-stone-200">
                <th className="px-4 py-2 font-medium">Imóvel</th>
                <th className="px-3 py-2 font-medium text-right">Receita</th>
                <th className="px-3 py-2 font-medium text-right">Despesa</th>
                <th className="px-3 py-2 font-medium text-right">Rent. bruta</th>
                <th className="px-3 py-2 font-medium text-right">Rent. líquida</th>
                <th className="px-3 py-2 font-medium text-right">Rent. real</th>
                <th className="px-3 py-2 font-medium text-right">Cash flow</th>
              </tr>
            </thead>
            <tbody>
              {properties.map((p) => {
                const ind = propertyIndicators(p);
                return (
                  <tr key={p.id} className="border-t border-stone-100">
                    <td className="px-4 py-2.5 text-stone-800">{p.name}</td>
                    <td className="px-3 py-2.5 text-right tabular-nums text-emerald-700">{fmtMoney(ind.incomeTotal)}</td>
                    <td className="px-3 py-2.5 text-right tabular-nums text-red-700">{fmtMoney(ind.expenseExclLoan + ind.loanPayments)}</td>
                    <td className="px-3 py-2.5 text-right tabular-nums">{fmtPct(ind.grossYield)}</td>
                    <td className="px-3 py-2.5 text-right tabular-nums font-medium">{fmtPct(ind.netYield)}</td>
                    <td className={`px-3 py-2.5 text-right tabular-nums ${ind.realYield >= 0 ? 'text-emerald-700' : 'text-red-700'}`}>{fmtPct(ind.realYield)}</td>
                    <td className={`px-3 py-2.5 text-right tabular-nums font-medium ${ind.cashFlowAfterLoan >= 0 ? 'text-emerald-700' : 'text-red-700'}`}>{fmtMoney(ind.cashFlowAfterLoan)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <p className="text-[11px] text-stone-400 px-4 py-3 border-t border-stone-100">
          Rentabilidade bruta = receita ÷ valor do imóvel. Líquida = (receita − despesas, excl. empréstimo) ÷ valor. Real = líquida − taxa de inflação. Cash flow inclui a prestação do empréstimo.
        </p>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------- */
/* properties view                                                     */
/* ---------------------------------------------------------------- */

function PropertiesView({ properties, propertyIndicators, onNew, onEdit, onDelete, onViewTx }) {
  return (
    <div>
      <div className="flex items-center justify-between mb-1">
        <h2 className="font-serif-display text-2xl text-stone-800">Imóveis</h2>
        <button onClick={onNew} className="inline-flex items-center gap-1.5 rounded bg-emerald-800 text-white text-sm px-3.5 py-2 hover:bg-emerald-900">
          <Plus size={15} /> Novo imóvel
        </button>
      </div>
      <p className="text-sm text-stone-500 mb-6">Casas, arrendamentos, Airbnb e garagens — com ou sem contrato.</p>

      {properties.length === 0 ? (
        <EmptyState icon={Building2} title="Ainda não há imóveis" message="Adicione o primeiro imóvel para começar a acompanhar rendas, despesas e rentabilidade." actionLabel="Adicionar imóvel" onAction={onNew} />
      ) : (
        <div className="grid md:grid-cols-2 lg:grid-cols-3 gap-4">
          {properties.map((p) => {
            const ind = propertyIndicators(p);
            return (
              <div key={p.id} className="bg-white border border-stone-200 rounded p-4 flex flex-col">
                <div className="flex items-start justify-between">
                  <div>
                    <p className="font-serif-display text-lg text-stone-800 leading-tight">{p.name}</p>
                    <Badge>{propertyTypeLabel(p.type)}</Badge>
                  </div>
                  <span className={`text-[11px] px-2 py-0.5 rounded-full ${p.status === 'Ativo' ? 'bg-emerald-100 text-emerald-800' : p.status === 'Vago' ? 'bg-amber-100 text-amber-800' : 'bg-stone-200 text-stone-600'}`}>{p.status}</span>
                </div>
                {p.address && <p className="text-xs text-stone-500 mt-2">{p.address}</p>}

                <div className="grid grid-cols-2 gap-2 mt-3 text-sm">
                  <div>
                    <p className="text-[11px] text-stone-400">Valor</p>
                    <p className="tabular-nums text-stone-800">{fmtMoney(p.currentValue || p.purchaseValue)}</p>
                  </div>
                  <div>
                    <p className="text-[11px] text-stone-400">Rentabilidade líquida (12m)</p>
                    <p className={`tabular-nums font-medium ${ind.netYield >= 0 ? 'text-emerald-700' : 'text-red-700'}`}>{fmtPct(ind.netYield)}</p>
                  </div>
                  <div>
                    <p className="text-[11px] text-stone-400">Cash flow (12m)</p>
                    <p className={`tabular-nums ${ind.cashFlowAfterLoan >= 0 ? 'text-emerald-700' : 'text-red-700'}`}>{fmtMoney(ind.cashFlowAfterLoan)}</p>
                  </div>
                  {(p.loans || []).length > 0 && (
                    <div>
                      <p className="text-[11px] text-stone-400">Empréstimo em dívida</p>
                      <p className="tabular-nums text-stone-800">{fmtMoney(ind.totalOutstandingLoans)}</p>
                    </div>
                  )}
                </div>

                <div className="flex flex-wrap gap-1.5 mt-3">
                  {(p.loans || []).length > 0 && <Badge tone="brass"><Landmark size={11} className="inline -mt-0.5 mr-1" />{p.loans.length} empréstimo{p.loans.length > 1 ? 's' : ''}</Badge>}
                  {(p.insurances || []).map((i) => <Badge key={i.id} tone="neutral"><ShieldCheck size={11} className="inline -mt-0.5 mr-1" />{i.type}</Badge>)}
                </div>

                <div className="mt-auto pt-3 flex items-center gap-3 border-t border-stone-100 mt-3">
                  <button onClick={() => onViewTx(p)} className="text-xs text-stone-500 hover:text-emerald-800 inline-flex items-center gap-1">Movimentos <ArrowRight size={12} /></button>
                  <button onClick={() => onEdit(p)} className="ml-auto text-stone-400 hover:text-stone-700"><Pencil size={15} /></button>
                  <button onClick={() => onDelete(p)} className="text-stone-400 hover:text-red-700"><Trash2 size={15} /></button>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

function PropertyFormModal({ modal, onClose, onSave }) {
  const [form, setForm] = useState(modal.data);
  const set = (field, value) => setForm((f) => ({ ...f, [field]: value }));

  const addLoan = () => set('loans', [...(form.loans || []), { id: uid(), name: '', amount: '', rate: '', termYears: '', startDate: '' }]);
  const updateLoan = (id, field, value) => set('loans', form.loans.map((l) => (l.id === id ? { ...l, [field]: value } : l)));
  const removeLoan = (id) => set('loans', form.loans.filter((l) => l.id !== id));

  const addInsurance = () => set('insurances', [...(form.insurances || []), { id: uid(), type: 'Multirriscos', provider: '', annualPremium: '' }]);
  const updateInsurance = (id, field, value) => set('insurances', form.insurances.map((i) => (i.id === id ? { ...i, [field]: value } : i)));
  const removeInsurance = (id) => set('insurances', form.insurances.filter((i) => i.id !== id));

  return (
    <Modal title={modal.mode === 'new' ? 'Novo imóvel' : 'Editar imóvel'} onClose={onClose} wide>
      <div className="grid md:grid-cols-2 gap-x-4">
        <Field label="Nome"><input className={inputCls} value={form.name} onChange={(e) => set('name', e.target.value)} placeholder="Ex.: T2 Rua das Flores" /></Field>
        <Field label="Tipo">
          <select className={inputCls} value={form.type} onChange={(e) => set('type', e.target.value)}>
            {PROPERTY_TYPES.map((t) => <option key={t.id} value={t.id}>{t.label}</option>)}
          </select>
        </Field>
        <Field label="Morada"><input className={inputCls} value={form.address} onChange={(e) => set('address', e.target.value)} /></Field>
        <Field label="Estado">
          <select className={inputCls} value={form.status} onChange={(e) => set('status', e.target.value)}>
            {PROPERTY_STATUS.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
        </Field>
        <Field label="Valor de compra (€)"><input type="number" step="0.01" className={inputCls} value={form.purchaseValue} onChange={(e) => set('purchaseValue', e.target.value)} /></Field>
        <Field label="Valor de mercado atual (€)" hint="opcional — usado nos indicadores se preenchido"><input type="number" step="0.01" className={inputCls} value={form.currentValue ?? ''} onChange={(e) => set('currentValue', e.target.value)} /></Field>
        <Field label="Data de compra"><input type="date" className={inputCls} value={form.purchaseDate} onChange={(e) => set('purchaseDate', e.target.value)} /></Field>
      </div>
      <Field label="Notas"><textarea className={inputCls} rows={2} value={form.notes} onChange={(e) => set('notes', e.target.value)} /></Field>

      <div className="mt-4 border-t border-stone-200 pt-4">
        <div className="flex items-center justify-between mb-2">
          <p className="text-sm font-medium text-stone-700 flex items-center gap-1.5"><Landmark size={15} /> Empréstimos bancários</p>
          <button type="button" onClick={addLoan} className="text-xs text-emerald-800 hover:underline">+ adicionar</button>
        </div>
        {(form.loans || []).map((l) => (
          <div key={l.id} className="grid grid-cols-2 md:grid-cols-5 gap-2 mb-2 items-end bg-stone-50 border border-stone-200 rounded p-2">
            <Field label="Designação"><input className={inputCls} value={l.name} onChange={(e) => updateLoan(l.id, 'name', e.target.value)} /></Field>
            <Field label="Montante (€)"><input type="number" className={inputCls} value={l.amount} onChange={(e) => updateLoan(l.id, 'amount', e.target.value)} /></Field>
            <Field label="Taxa anual (%)"><input type="number" step="0.01" className={inputCls} value={l.rate} onChange={(e) => updateLoan(l.id, 'rate', e.target.value)} /></Field>
            <Field label="Prazo (anos)"><input type="number" className={inputCls} value={l.termYears} onChange={(e) => updateLoan(l.id, 'termYears', e.target.value)} /></Field>
            <div className="flex gap-1">
              <Field label="Início"><input type="date" className={inputCls} value={l.startDate} onChange={(e) => updateLoan(l.id, 'startDate', e.target.value)} /></Field>
              <button type="button" onClick={() => removeLoan(l.id)} className="mb-3 text-stone-400 hover:text-red-700"><Trash2 size={16} /></button>
            </div>
          </div>
        ))}
      </div>

      <div className="mt-2 border-t border-stone-200 pt-4">
        <div className="flex items-center justify-between mb-2">
          <p className="text-sm font-medium text-stone-700 flex items-center gap-1.5"><ShieldCheck size={15} /> Seguros (multirriscos, vida, crédito habitação…)</p>
          <button type="button" onClick={addInsurance} className="text-xs text-emerald-800 hover:underline">+ adicionar</button>
        </div>
        {(form.insurances || []).map((i) => (
          <div key={i.id} className="grid grid-cols-2 md:grid-cols-4 gap-2 mb-2 items-end bg-stone-50 border border-stone-200 rounded p-2">
            <Field label="Tipo">
              <select className={inputCls} value={i.type} onChange={(e) => updateInsurance(i.id, 'type', e.target.value)}>
                {INSURANCE_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
              </select>
            </Field>
            <Field label="Seguradora"><input className={inputCls} value={i.provider} onChange={(e) => updateInsurance(i.id, 'provider', e.target.value)} /></Field>
            <Field label="Prémio anual (€)"><input type="number" step="0.01" className={inputCls} value={i.annualPremium} onChange={(e) => updateInsurance(i.id, 'annualPremium', e.target.value)} /></Field>
            <button type="button" onClick={() => removeInsurance(i.id)} className="mb-3 text-stone-400 hover:text-red-700 justify-self-start"><Trash2 size={16} /></button>
          </div>
        ))}
        <p className="text-[11px] text-stone-400 mt-1">Empréstimos e seguros aqui servem para calcular a prestação mensal e o valor em dívida. Os pagamentos reais continuam a registar-se em Movimentos.</p>
      </div>

      <div className="flex justify-end gap-2 mt-5 pt-4 border-t border-stone-200">
        <button onClick={onClose} className="px-4 py-2 text-sm rounded border border-stone-300 text-stone-600 hover:bg-stone-50">Cancelar</button>
        <button onClick={() => onSave(form)} disabled={!form.name} className="px-4 py-2 text-sm rounded bg-emerald-800 text-white hover:bg-emerald-900 disabled:opacity-40">Guardar</button>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------- */
/* transactions view                                                   */
/* ---------------------------------------------------------------- */

function TransactionsView({ transactions, properties, filter, setFilter, onNew, onEdit, onDelete, onImportClick }) {
  const propertyName = (id) => properties.find((p) => p.id === id)?.name || '—';

  const filtered = useMemo(() => {
    return transactions
      .filter((t) => filter.propertyId === 'all' || t.propertyId === filter.propertyId)
      .filter((t) => filter.type === 'all' || t.type === filter.type)
      .filter((t) => filter.category === 'all' || t.category === filter.category)
      .sort((a, b) => new Date(b.date) - new Date(a.date));
  }, [transactions, filter]);

  const totalIncome = filtered.filter((t) => t.type === 'income').reduce((s, t) => s + Number(t.amount), 0);
  const totalExpense = filtered.filter((t) => t.type === 'expense').reduce((s, t) => s + Number(t.amount), 0);
  const allCategories = [...new Set(transactions.map((t) => t.category))];

  return (
    <div>
      <div className="flex items-center justify-between mb-1">
        <h2 className="font-serif-display text-2xl text-stone-800">Movimentos</h2>
        <div className="flex gap-2">
          <button onClick={onImportClick} disabled={properties.length === 0} className="inline-flex items-center gap-1.5 rounded border border-stone-300 text-stone-600 text-sm px-3.5 py-2 hover:bg-stone-50 disabled:opacity-40">
            <Upload size={15} /> Importar Excel
          </button>
          <button onClick={onNew} disabled={properties.length === 0} className="inline-flex items-center gap-1.5 rounded bg-emerald-800 text-white text-sm px-3.5 py-2 hover:bg-emerald-900 disabled:opacity-40">
            <Plus size={15} /> Novo movimento
          </button>
        </div>
      </div>
      <p className="text-sm text-stone-500 mb-4">
        {properties.length === 0 ? 'Adicione um imóvel primeiro para poder registar movimentos.' : 'Rendas, IMI, condomínio, EDP, água, net, manutenção, seguros e mais.'}
      </p>

      {properties.length > 0 && (
        <div className="flex flex-wrap gap-2 mb-4">
          <select className={`${inputCls} w-auto`} value={filter.propertyId} onChange={(e) => setFilter((f) => ({ ...f, propertyId: e.target.value }))}>
            <option value="all">Todos os imóveis</option>
            {properties.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
          </select>
          <select className={`${inputCls} w-auto`} value={filter.type} onChange={(e) => setFilter((f) => ({ ...f, type: e.target.value }))}>
            <option value="all">Receitas e despesas</option>
            <option value="income">Só receitas</option>
            <option value="expense">Só despesas</option>
          </select>
          <select className={`${inputCls} w-auto`} value={filter.category} onChange={(e) => setFilter((f) => ({ ...f, category: e.target.value }))}>
            <option value="all">Todas as categorias</option>
            {allCategories.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>
      )}

      {filtered.length === 0 ? (
        <EmptyState icon={Receipt} title="Sem movimentos" message={properties.length === 0 ? 'Comece por adicionar um imóvel em Imóveis.' : 'Registe a primeira renda ou despesa para começar a ver a análise.'} actionLabel={properties.length > 0 ? 'Novo movimento' : undefined} onAction={onNew} />
      ) : (
        <div className="bg-white border border-stone-200 rounded overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-left text-[11px] text-stone-500">
                  <th className="px-4 py-2.5 font-medium">Data</th>
                  <th className="px-3 py-2.5 font-medium">Imóvel</th>
                  <th className="px-3 py-2.5 font-medium">Categoria</th>
                  <th className="px-3 py-2.5 font-medium">Descrição</th>
                  <th className="px-3 py-2.5 font-medium">Recorrência</th>
                  <th className="px-3 py-2.5 font-medium text-right">Valor</th>
                  <th className="px-3 py-2.5"></th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((t) => (
                  <tr key={t.id} className="border-t border-stone-100 hover:bg-stone-50">
                    <td className="px-4 py-2.5 text-stone-600 whitespace-nowrap">
                      {fmtDate(t.date)}
                      {t.reservationId && <span className="ml-1.5 text-[10px] text-amber-700 bg-amber-100 rounded-full px-1.5 py-0.5">reserva</span>}
                    </td>
                    <td className="px-3 py-2.5 text-stone-800">{propertyName(t.propertyId)}</td>
                    <td className="px-3 py-2.5"><Badge tone={t.type === 'income' ? 'income' : 'expense'}>{t.category}</Badge></td>
                    <td className="px-3 py-2.5 text-stone-500">{t.description || '—'}</td>
                    <td className="px-3 py-2.5 text-stone-500">{t.recurring === 'monthly' ? 'Mensal' : t.recurring === 'yearly' ? 'Anual' : '—'}</td>
                    <td className={`px-3 py-2.5 text-right tabular-nums font-medium whitespace-nowrap ${t.type === 'income' ? 'text-emerald-700' : 'text-red-700'}`}>
                      {t.type === 'income' ? '+' : '−'}{fmtMoney(Math.abs(t.amount), 2)}
                    </td>
                    <td className="px-3 py-2.5">
                      <div className="flex gap-2 justify-end">
                        <button onClick={() => onEdit(t)} className="text-stone-400 hover:text-stone-700"><Pencil size={14} /></button>
                        <button onClick={() => onDelete(t)} className="text-stone-400 hover:text-red-700"><Trash2 size={14} /></button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr className="border-t border-stone-200 bg-stone-50 font-medium">
                  <td className="px-4 py-2.5" colSpan={5}>Total ({filtered.length} movimento{filtered.length !== 1 ? 's' : ''})</td>
                  <td className="px-3 py-2.5 text-right tabular-nums" colSpan={2}>
                    <span className="text-emerald-700">{fmtMoney(totalIncome, 2)}</span>
                    {' / '}
                    <span className="text-red-700">{fmtMoney(totalExpense, 2)}</span>
                  </td>
                </tr>
              </tfoot>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

function TransactionFormModal({ modal, onClose, onSave, properties, incomeCategories, expenseCategories }) {
  const [form, setForm] = useState(modal.data);
  const [newCategory, setNewCategory] = useState('');
  const [addingCategory, setAddingCategory] = useState(false);
  const set = (field, value) => setForm((f) => ({ ...f, [field]: value }));
  const categories = form.type === 'income' ? incomeCategories : expenseCategories;

  useEffect(() => {
    if (!categories.includes(form.category)) set('category', categories[0]);
    // eslint-disable-next-line
  }, [form.type]);

  return (
    <Modal title={modal.mode === 'new' ? 'Novo movimento' : 'Editar movimento'} onClose={onClose}>
      <Field label="Imóvel">
        <select className={inputCls} value={form.propertyId} onChange={(e) => set('propertyId', e.target.value)}>
          {properties.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
        </select>
      </Field>

      <div className="flex gap-2 mb-3">
        <button type="button" onClick={() => set('type', 'income')} className={`flex-1 rounded py-2 text-sm border ${form.type === 'income' ? 'bg-emerald-800 text-white border-emerald-800' : 'border-stone-300 text-stone-600'}`}>Receita</button>
        <button type="button" onClick={() => set('type', 'expense')} className={`flex-1 rounded py-2 text-sm border ${form.type === 'expense' ? 'bg-red-700 text-white border-red-700' : 'border-stone-300 text-stone-600'}`}>Despesa</button>
      </div>

      <Field label="Categoria">
        {!addingCategory ? (
          <select className={inputCls} value={form.category} onChange={(e) => {
            if (e.target.value === '__new__') setAddingCategory(true); else set('category', e.target.value);
          }}>
            {categories.map((c) => <option key={c} value={c}>{c}</option>)}
            <option value="__new__">+ nova categoria…</option>
          </select>
        ) : (
          <div className="flex gap-2">
            <input autoFocus className={inputCls} placeholder="Nome da categoria" value={newCategory} onChange={(e) => setNewCategory(e.target.value)} />
            <button type="button" onClick={() => setAddingCategory(false)} className="text-stone-400 hover:text-stone-700"><X size={18} /></button>
          </div>
        )}
      </Field>

      <div className="grid grid-cols-2 gap-3">
        <Field label="Valor (€)"><input type="number" step="0.01" className={inputCls} value={form.amount} onChange={(e) => set('amount', e.target.value)} /></Field>
        <Field label="Data"><input type="date" className={inputCls} value={form.date} onChange={(e) => set('date', e.target.value)} /></Field>
      </div>
      <Field label="Descrição (opcional)"><input className={inputCls} value={form.description} onChange={(e) => set('description', e.target.value)} /></Field>
      <Field label="Recorrência" hint="apenas informativo — cada movimento tem de ser registado quando ocorre">
        <select className={inputCls} value={form.recurring} onChange={(e) => set('recurring', e.target.value)}>
          <option value="none">Pontual</option>
          <option value="monthly">Mensal</option>
          <option value="yearly">Anual</option>
        </select>
      </Field>

      <div className="flex justify-end gap-2 mt-5 pt-4 border-t border-stone-200">
        <button onClick={onClose} className="px-4 py-2 text-sm rounded border border-stone-300 text-stone-600 hover:bg-stone-50">Cancelar</button>
        <button
          onClick={() => onSave(form, addingCategory ? newCategory.trim() : null)}
          disabled={!form.amount || !form.date || (addingCategory && !newCategory.trim())}
          className="px-4 py-2 text-sm rounded bg-emerald-800 text-white hover:bg-emerald-900 disabled:opacity-40"
        >
          Guardar
        </button>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------- */
/* alojamento local                                                     */
/* ---------------------------------------------------------------- */

function shortTermStats(property, reservations) {
  const now = new Date();
  const oneYearAgo = new Date(now); oneYearAgo.setFullYear(now.getFullYear() - 1);
  const propRes = reservations.filter((r) => r.propertyId === property.id && new Date(r.checkOut || r.checkIn) >= oneYearAgo);
  const nights = propRes.reduce((s, r) => s + nightsBetween(r.checkIn, r.checkOut), 0);
  const grossRevenue = propRes.reduce((s, r) => s + (Number(r.grossAmount) || 0), 0);
  const fees = propRes.reduce((s, r) => s + (Number(r.platformFee) || 0) + (Number(r.cleaningFee) || 0), 0);
  const days = 365;
  return {
    count: propRes.length, nights, grossRevenue, fees,
    occupancy: (nights / days) * 100,
    adr: nights ? grossRevenue / nights : 0,
    revpar: grossRevenue / days,
  };
}

function ShortTermView({ properties, reservations, onNew, onEdit, onDelete }) {
  const stProperties = properties.filter((p) => p.type === 'airbnb');
  const [filterProperty, setFilterProperty] = useState('all');
  const propertyName = (id) => properties.find((p) => p.id === id)?.name || '—';

  const filtered = useMemo(
    () => reservations.filter((r) => filterProperty === 'all' || r.propertyId === filterProperty).sort((a, b) => new Date(b.checkIn) - new Date(a.checkIn)),
    [reservations, filterProperty]
  );

  return (
    <div>
      <div className="flex items-center justify-between mb-1">
        <h2 className="font-serif-display text-2xl text-stone-800">Alojamento Local</h2>
        <button onClick={onNew} disabled={properties.length === 0} className="inline-flex items-center gap-1.5 rounded bg-emerald-800 text-white text-sm px-3.5 py-2 hover:bg-emerald-900 disabled:opacity-40">
          <Plus size={15} /> Nova reserva
        </button>
      </div>
      <p className="text-sm text-stone-500 mb-6">Reservas, ocupação e diária média — cada reserva gera automaticamente os movimentos de receita, comissão e limpeza.</p>

      {stProperties.length === 0 ? (
        <div className="mb-6 border border-dashed border-amber-300 bg-amber-50 rounded p-4 text-sm text-amber-800">
          Ainda não há nenhum imóvel do tipo "Airbnb / alojamento local". Marque esse tipo ao editar um imóvel em Imóveis para ver aqui os indicadores de ocupação — pode registar reservas na mesma para qualquer imóvel.
        </div>
      ) : (
        <div className="grid md:grid-cols-2 lg:grid-cols-3 gap-4 mb-6">
          {stProperties.map((p) => {
            const stat = shortTermStats(p, reservations);
            return (
              <div key={p.id} className="bg-white border border-stone-200 rounded p-4">
                <p className="font-serif-display text-lg text-stone-800 mb-2">{p.name}</p>
                <div className="grid grid-cols-2 gap-2 text-sm">
                  <div><p className="text-[11px] text-stone-400">Ocupação (12m)</p><p className="tabular-nums font-medium">{fmtPct(stat.occupancy)}</p></div>
                  <div><p className="text-[11px] text-stone-400">Diária média</p><p className="tabular-nums font-medium">{fmtMoney(stat.adr, 2)}</p></div>
                  <div><p className="text-[11px] text-stone-400">RevPAR</p><p className="tabular-nums">{fmtMoney(stat.revpar, 2)}</p></div>
                  <div><p className="text-[11px] text-stone-400">Receita bruta (12m)</p><p className="tabular-nums text-emerald-700">{fmtMoney(stat.grossRevenue)}</p></div>
                </div>
                <p className="text-[11px] text-stone-400 mt-2">{stat.count} reserva{stat.count !== 1 ? 's' : ''} · {stat.nights} noites · {fmtMoney(stat.fees)} em comissões/limpeza</p>
              </div>
            );
          })}
        </div>
      )}

      {reservations.length === 0 ? (
        <EmptyState icon={CalendarDays} title="Sem reservas" message="Registe a primeira reserva para começar a acompanhar ocupação e diária média." actionLabel={properties.length > 0 ? 'Nova reserva' : undefined} onAction={onNew} />
      ) : (
        <div className="bg-white border border-stone-200 rounded overflow-hidden">
          <div className="px-4 pt-4 pb-2">
            <select className={`${inputCls} w-auto`} value={filterProperty} onChange={(e) => setFilterProperty(e.target.value)}>
              <option value="all">Todos os imóveis</option>
              {properties.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
            </select>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-left text-[11px] text-stone-500 border-t border-stone-200">
                  <th className="px-4 py-2 font-medium">Imóvel</th>
                  <th className="px-3 py-2 font-medium">Hóspede</th>
                  <th className="px-3 py-2 font-medium">Plataforma</th>
                  <th className="px-3 py-2 font-medium">Check-in</th>
                  <th className="px-3 py-2 font-medium">Check-out</th>
                  <th className="px-3 py-2 font-medium text-right">Noites</th>
                  <th className="px-3 py-2 font-medium text-right">Valor bruto</th>
                  <th className="px-3 py-2"></th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((r) => (
                  <tr key={r.id} className="border-t border-stone-100 hover:bg-stone-50">
                    <td className="px-4 py-2.5 text-stone-800">{propertyName(r.propertyId)}</td>
                    <td className="px-3 py-2.5 text-stone-600">{r.guestName || '—'}</td>
                    <td className="px-3 py-2.5"><Badge>{r.platform}</Badge></td>
                    <td className="px-3 py-2.5 text-stone-600 whitespace-nowrap">{fmtDate(r.checkIn)}</td>
                    <td className="px-3 py-2.5 text-stone-600 whitespace-nowrap">{fmtDate(r.checkOut)}</td>
                    <td className="px-3 py-2.5 text-right tabular-nums">{nightsBetween(r.checkIn, r.checkOut)}</td>
                    <td className="px-3 py-2.5 text-right tabular-nums text-emerald-700 font-medium">{fmtMoney(r.grossAmount, 2)}</td>
                    <td className="px-3 py-2.5">
                      <div className="flex gap-2 justify-end">
                        <button onClick={() => onEdit(r)} className="text-stone-400 hover:text-stone-700"><Pencil size={14} /></button>
                        <button onClick={() => onDelete(r)} className="text-stone-400 hover:text-red-700"><Trash2 size={14} /></button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

function ReservationFormModal({ modal, onClose, onSave, properties }) {
  const [form, setForm] = useState(modal.data);
  const set = (field, value) => setForm((f) => ({ ...f, [field]: value }));
  const nights = nightsBetween(form.checkIn, form.checkOut);
  const valid = form.propertyId && form.checkIn && form.checkOut && nights > 0 && Number(form.grossAmount) > 0;

  return (
    <Modal title={modal.mode === 'new' ? 'Nova reserva' : 'Editar reserva'} onClose={onClose}>
      <Field label="Imóvel">
        <select className={inputCls} value={form.propertyId} onChange={(e) => set('propertyId', e.target.value)}>
          {properties.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
        </select>
      </Field>
      <div className="grid grid-cols-2 gap-3">
        <Field label="Hóspede (opcional)"><input className={inputCls} value={form.guestName} onChange={(e) => set('guestName', e.target.value)} /></Field>
        <Field label="Plataforma">
          <select className={inputCls} value={form.platform} onChange={(e) => set('platform', e.target.value)}>
            {RESERVATION_PLATFORMS.map((p) => <option key={p} value={p}>{p}</option>)}
          </select>
        </Field>
      </div>
      <div className="grid grid-cols-2 gap-3">
        <Field label="Check-in"><input type="date" className={inputCls} value={form.checkIn} onChange={(e) => set('checkIn', e.target.value)} /></Field>
        <Field label="Check-out"><input type="date" className={inputCls} value={form.checkOut} onChange={(e) => set('checkOut', e.target.value)} /></Field>
      </div>
      {form.checkIn && form.checkOut && <p className="text-xs text-stone-400 -mt-2 mb-3">{nights} noite{nights !== 1 ? 's' : ''}</p>}
      <div className="grid grid-cols-3 gap-3">
        <Field label="Valor bruto (€)"><input type="number" step="0.01" className={inputCls} value={form.grossAmount} onChange={(e) => set('grossAmount', e.target.value)} /></Field>
        <Field label="Comissão (€)" hint="opcional"><input type="number" step="0.01" className={inputCls} value={form.platformFee} onChange={(e) => set('platformFee', e.target.value)} /></Field>
        <Field label="Limpeza (€)" hint="opcional"><input type="number" step="0.01" className={inputCls} value={form.cleaningFee} onChange={(e) => set('cleaningFee', e.target.value)} /></Field>
      </div>
      <Field label="Notas (opcional)"><input className={inputCls} value={form.notes} onChange={(e) => set('notes', e.target.value)} /></Field>
      <p className="text-[11px] text-stone-400 -mt-2">Ao guardar, esta reserva gera automaticamente os movimentos de receita, comissão e limpeza em Movimentos.</p>

      <div className="flex justify-end gap-2 mt-5 pt-4 border-t border-stone-200">
        <button onClick={onClose} className="px-4 py-2 text-sm rounded border border-stone-300 text-stone-600 hover:bg-stone-50">Cancelar</button>
        <button onClick={() => onSave(form)} disabled={!valid} className="px-4 py-2 text-sm rounded bg-emerald-800 text-white hover:bg-emerald-900 disabled:opacity-40">Guardar</button>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------- */
/* importação de excel                                                  */
/* ---------------------------------------------------------------- */

const IMPORT_FIELDS = [
  { id: 'ignore', label: 'Ignorar coluna' },
  { id: 'date', label: 'Data' },
  { id: 'property', label: 'Imóvel' },
  { id: 'type', label: 'Tipo (receita/despesa)' },
  { id: 'category', label: 'Categoria' },
  { id: 'amount', label: 'Valor' },
  { id: 'description', label: 'Descrição' },
];

function guessMapping(header) {
  const h = normalizeText(header);
  if (/data|date/.test(h)) return 'date';
  if (/imovel|propriedade|casa/.test(h)) return 'property';
  if (/tipo/.test(h)) return 'type';
  if (/categoria/.test(h)) return 'category';
  if (/valor|montante|amount/.test(h)) return 'amount';
  if (/descri|nota|obs/.test(h)) return 'description';
  return 'ignore';
}

function ImportExcelModal({ properties, incomeCategories, expenseCategories, onClose, onImport }) {
  const [step, setStep] = useState('upload'); // upload | map | done
  const [fileName, setFileName] = useState('');
  const [headers, setHeaders] = useState([]);
  const [rows, setRows] = useState([]);
  const [mapping, setMapping] = useState({});
  const [defaultPropertyId, setDefaultPropertyId] = useState(properties[0]?.id || '');
  const [typeMode, setTypeMode] = useState('column'); // column | fixed | sign
  const [fixedType, setFixedType] = useState('expense');
  const [result, setResult] = useState(null);
  const [parseError, setParseError] = useState('');

  async function handleFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setParseError('');
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array', cellDates: true });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
      if (!data.length) { setParseError('A folha parece estar vazia.'); return; }
      const [head, ...body] = data;
      const cleanHeaders = head.map((h, i) => (h === '' || h === undefined ? `Coluna ${i + 1}` : String(h)));
      const dataRows = body.filter((r) => r.some((c) => c !== '' && c !== undefined && c !== null));
      const initialMapping = {};
      cleanHeaders.forEach((h, i) => { initialMapping[i] = guessMapping(h); });
      setFileName(file.name);
      setHeaders(cleanHeaders);
      setRows(dataRows);
      setMapping(initialMapping);
      if (!Object.values(initialMapping).includes('type')) setTypeMode('fixed');
      setStep('map');
    } catch (err) {
      console.error(err);
      setParseError('Não foi possível ler este ficheiro. Confirme que é um .xlsx, .xls ou .csv válido.');
    }
  }

  function colIndexFor(field) {
    return Object.keys(mapping).find((k) => mapping[k] === field);
  }

  function resolveRow(r) {
    const dateIdx = colIndexFor('date'), propIdx = colIndexFor('property'), typeIdx = colIndexFor('type');
    const catIdx = colIndexFor('category'), amountIdx = colIndexFor('amount'), descIdx = colIndexFor('description');

    const date = dateIdx !== undefined ? parseImportDate(r[dateIdx]) : null;
    const amountRaw = amountIdx !== undefined ? parseAmountValue(r[amountIdx]) : NaN;

    let type = null;
    if (typeMode === 'fixed') type = fixedType;
    else if (typeMode === 'sign') type = amountRaw < 0 ? 'expense' : 'income';
    else if (typeIdx !== undefined) {
      const t = normalizeText(r[typeIdx]);
      if (/desp|expense|saida/.test(t)) type = 'expense';
      else if (/rece|income|entrada/.test(t)) type = 'income';
    }

    let propertyId = null, propertyLabel = '—';
    if (propIdx !== undefined && r[propIdx]) {
      const norm = normalizeText(r[propIdx]);
      const match = properties.find((p) => normalizeText(p.name) === norm);
      propertyId = match ? match.id : (defaultPropertyId || null);
      propertyLabel = match ? match.name : `${r[propIdx]} (sem correspondência → imóvel padrão)`;
    } else {
      propertyId = defaultPropertyId || null;
      propertyLabel = properties.find((p) => p.id === defaultPropertyId)?.name || '—';
    }

    const categoryList = type === 'income' ? incomeCategories : expenseCategories;
    let category = catIdx !== undefined && r[catIdx] ? String(r[catIdx]).trim() : null;
    let isNewCategory = false;
    if (category) {
      const match = categoryList.find((c) => normalizeText(c) === normalizeText(category));
      if (match) category = match; else isNewCategory = true;
    } else {
      category = type === 'income' ? 'Outras receitas' : 'Outras despesas';
    }

    const description = descIdx !== undefined ? String(r[descIdx] || '') : '';
    const amount = isNaN(amountRaw) ? null : Math.abs(amountRaw);

    const errors = [];
    if (!date) errors.push('data inválida');
    if (amount === null) errors.push('valor inválido');
    if (!type) errors.push('tipo indefinido');
    if (!propertyId) errors.push('imóvel não identificado');

    return { date, propertyId, propertyLabel, type, category, isNewCategory, description, amount, errors };
  }

  const preview = useMemo(() => (step === 'map' ? rows.slice(0, 5).map(resolveRow) : []), [rows, mapping, typeMode, fixedType, defaultPropertyId, step]);

  function confirmImport() {
    const resolved = rows.map(resolveRow);
    const valid = resolved.filter((r) => r.errors.length === 0);
    const invalid = resolved.filter((r) => r.errors.length > 0);
    const newTransactions = valid.map((r) => ({ id: uid(), propertyId: r.propertyId, type: r.type, category: r.category, amount: r.amount, date: r.date, description: r.description, recurring: 'none' }));
    const newIncomeCats = [...new Set(valid.filter((r) => r.isNewCategory && r.type === 'income').map((r) => r.category))];
    const newExpenseCats = [...new Set(valid.filter((r) => r.isNewCategory && r.type === 'expense').map((r) => r.category))];
    onImport(newTransactions, { income: newIncomeCats, expense: newExpenseCats });
    setResult({ importedCount: valid.length, skippedCount: invalid.length, errors: invalid.slice(0, 8) });
    setStep('done');
  }

  return (
    <Modal title="Importar movimentos de Excel" onClose={onClose} wide>
      {step === 'upload' && (
        <div>
          <p className="text-sm text-stone-600 mb-4">Carregue um ficheiro .xlsx, .xls ou .csv. No passo seguinte associa cada coluna a um campo (data, imóvel, tipo, categoria, valor, descrição).</p>
          <label className="flex flex-col items-center justify-center gap-2 border border-dashed border-stone-300 rounded p-10 text-center cursor-pointer hover:bg-stone-50">
            <Upload className="text-stone-400" size={26} />
            <span className="text-sm text-stone-600">Clique para escolher o ficheiro</span>
            <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleFile} />
          </label>
          {parseError && <p className="text-sm text-red-700 mt-3">{parseError}</p>}
        </div>
      )}

      {step === 'map' && (
        <div>
          <p className="text-sm text-stone-600 mb-1">Ficheiro: <strong>{fileName}</strong> · {rows.length} linhas detetadas</p>
          <p className="text-xs text-stone-400 mb-4">Associe cada coluna do ficheiro a um campo.</p>

          <div className="overflow-x-auto border border-stone-200 rounded mb-4">
            <table className="w-full text-sm">
              <thead>
                <tr className="bg-stone-50 text-left text-[11px] text-stone-500">
                  {headers.map((h, i) => (
                    <th key={i} className="px-3 py-2 font-medium min-w-[140px]">
                      <p className="mb-1 truncate">{h}</p>
                      <select className={`${inputCls} text-xs py-1`} value={mapping[i]} onChange={(e) => setMapping((m) => ({ ...m, [i]: e.target.value }))}>
                        {IMPORT_FIELDS.map((f) => <option key={f.id} value={f.id}>{f.label}</option>)}
                      </select>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.slice(0, 3).map((r, i) => (
                  <tr key={i} className="border-t border-stone-100 text-stone-500">
                    {headers.map((_, ci) => <td key={ci} className="px-3 py-1.5 truncate max-w-[160px]">{String(r[ci] ?? '')}</td>)}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="grid md:grid-cols-2 gap-3 mb-4">
            <Field label="Imóvel padrão" hint="usado quando a coluna Imóvel não tem correspondência, ou não foi mapeada">
              <select className={inputCls} value={defaultPropertyId} onChange={(e) => setDefaultPropertyId(e.target.value)}>
                {properties.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
              </select>
            </Field>
            <Field label="Como determinar o tipo (receita/despesa)">
              <select className={inputCls} value={typeMode} onChange={(e) => setTypeMode(e.target.value)}>
                <option value="column">Usar a coluna mapeada como Tipo</option>
                <option value="sign">Inferir pelo sinal do valor (negativo = despesa)</option>
                <option value="fixed">Aplicar o mesmo tipo a todas as linhas</option>
              </select>
            </Field>
            {typeMode === 'fixed' && (
              <Field label="Tipo a aplicar">
                <select className={inputCls} value={fixedType} onChange={(e) => setFixedType(e.target.value)}>
                  <option value="income">Receita</option>
                  <option value="expense">Despesa</option>
                </select>
              </Field>
            )}
          </div>

          <p className="text-xs font-medium text-stone-600 mb-2">Pré-visualização (primeiras {preview.length} linhas)</p>
          <div className="overflow-x-auto border border-stone-200 rounded mb-4">
            <table className="w-full text-xs">
              <thead>
                <tr className="bg-stone-50 text-left text-stone-500">
                  <th className="px-3 py-1.5">Data</th><th className="px-3 py-1.5">Imóvel</th><th className="px-3 py-1.5">Tipo</th>
                  <th className="px-3 py-1.5">Categoria</th><th className="px-3 py-1.5 text-right">Valor</th><th className="px-3 py-1.5">Estado</th>
                </tr>
              </thead>
              <tbody>
                {preview.map((r, i) => (
                  <tr key={i} className="border-t border-stone-100">
                    <td className="px-3 py-1.5">{r.date || '—'}</td>
                    <td className="px-3 py-1.5">{r.propertyLabel}</td>
                    <td className="px-3 py-1.5">{r.type === 'income' ? 'Receita' : r.type === 'expense' ? 'Despesa' : '—'}</td>
                    <td className="px-3 py-1.5">{r.category}{r.isNewCategory && <span className="text-amber-700"> (nova)</span>}</td>
                    <td className="px-3 py-1.5 text-right tabular-nums">{r.amount !== null ? fmtMoney(r.amount, 2) : '—'}</td>
                    <td className="px-3 py-1.5">{r.errors.length ? <span className="text-red-700">{r.errors.join(', ')}</span> : <span className="text-emerald-700">ok</span>}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="flex justify-between gap-2 pt-2 border-t border-stone-200">
            <button onClick={() => setStep('upload')} className="px-4 py-2 text-sm rounded border border-stone-300 text-stone-600 hover:bg-stone-50">Voltar</button>
            <button onClick={confirmImport} disabled={!properties.length} className="px-4 py-2 text-sm rounded bg-emerald-800 text-white hover:bg-emerald-900 disabled:opacity-40">Importar {rows.length} linhas</button>
          </div>
        </div>
      )}

      {step === 'done' && result && (
        <div className="text-center py-4">
          <CheckCircle2 className="mx-auto text-emerald-700 mb-3" size={32} />
          <p className="font-serif-display text-lg text-stone-800">{result.importedCount} movimentos importados</p>
          {result.skippedCount > 0 && (
            <div className="mt-3 text-left inline-block">
              <p className="text-sm text-amber-700 mb-1">{result.skippedCount} linha{result.skippedCount !== 1 ? 's' : ''} ignorada{result.skippedCount !== 1 ? 's' : ''}:</p>
              <ul className="text-xs text-stone-500 list-disc pl-4">
                {result.errors.map((e, i) => <li key={i}>{e.propertyLabel} · {e.errors.join(', ')}</li>)}
              </ul>
            </div>
          )}
          <button onClick={onClose} className="mt-5 px-4 py-2 text-sm rounded bg-emerald-800 text-white hover:bg-emerald-900">Concluir</button>
        </div>
      )}
    </Modal>
  );
}

/* ---------------------------------------------------------------- */
/* settings view                                                       */
/* ---------------------------------------------------------------- */

function SettingsView({ settings, persistSettings, incomeCategories, expenseCategories, properties, transactions, reservations, onClearAll }) {
  const [inflation, setInflation] = useState(settings.inflationRate);
  const [newIncomeCat, setNewIncomeCat] = useState('');
  const [newExpenseCat, setNewExpenseCat] = useState('');

  function removeCategory(list, cat) {
    const key = list === 'income' ? 'customIncomeCategories' : 'customExpenseCategories';
    persistSettings({ ...settings, [key]: settings[key].filter((c) => c !== cat) });
  }
  function addCategory(list) {
    const value = (list === 'income' ? newIncomeCat : newExpenseCat).trim();
    if (!value) return;
    const key = list === 'income' ? 'customIncomeCategories' : 'customExpenseCategories';
    if (!settings[key].includes(value)) persistSettings({ ...settings, [key]: [...settings[key], value] });
    if (list === 'income') setNewIncomeCat(''); else setNewExpenseCat('');
  }

  function exportData() {
    const payload = { properties, transactions, reservations, settings, exportedAt: new Date().toISOString() };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `gestao-imobiliaria-${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
  }

  return (
    <div>
      <h2 className="font-serif-display text-2xl text-stone-800 mb-1">Definições</h2>
      <p className="text-sm text-stone-500 mb-6">Parâmetros gerais da análise e gestão de categorias.</p>

      <div className="bg-white border border-stone-200 rounded p-4 mb-4 max-w-md">
        <p className="text-sm font-medium text-stone-700 mb-3">Análise</p>
        <Field label="Taxa de inflação anual (%)" hint="usada para calcular a rentabilidade real por imóvel">
          <input
            type="number" step="0.1" className={inputCls} value={inflation}
            onChange={(e) => setInflation(e.target.value)}
            onBlur={() => persistSettings({ ...settings, inflationRate: Number(inflation) || 0 })}
          />
        </Field>
        <p className="text-xs text-stone-400">Moeda: EUR — Euro</p>
      </div>

      <div className="grid md:grid-cols-2 gap-4 mb-4">
        <div className="bg-white border border-stone-200 rounded p-4">
          <p className="text-sm font-medium text-stone-700 mb-3">Categorias de receita</p>
          <div className="flex flex-wrap gap-1.5 mb-3">
            {incomeCategories.map((c) => (
              <span key={c} className="inline-flex items-center gap-1 rounded-full bg-emerald-50 text-emerald-800 text-[11px] px-2 py-1">
                {c}
                {settings.customIncomeCategories.includes(c) && (
                  <button onClick={() => removeCategory('income', c)} className="hover:text-red-700"><X size={11} /></button>
                )}
              </span>
            ))}
          </div>
          <div className="flex gap-2">
            <input className={inputCls} placeholder="Nova categoria" value={newIncomeCat} onChange={(e) => setNewIncomeCat(e.target.value)} />
            <button onClick={() => addCategory('income')} className="px-3 rounded bg-stone-200 text-stone-700 text-sm hover:bg-stone-300">Adicionar</button>
          </div>
        </div>

        <div className="bg-white border border-stone-200 rounded p-4">
          <p className="text-sm font-medium text-stone-700 mb-3">Categorias de despesa</p>
          <div className="flex flex-wrap gap-1.5 mb-3">
            {expenseCategories.map((c) => (
              <span key={c} className="inline-flex items-center gap-1 rounded-full bg-red-50 text-red-800 text-[11px] px-2 py-1">
                {c}
                {settings.customExpenseCategories.includes(c) && (
                  <button onClick={() => removeCategory('expense', c)} className="hover:text-red-900"><X size={11} /></button>
                )}
              </span>
            ))}
          </div>
          <div className="flex gap-2">
            <input className={inputCls} placeholder="Nova categoria" value={newExpenseCat} onChange={(e) => setNewExpenseCat(e.target.value)} />
            <button onClick={() => addCategory('expense')} className="px-3 rounded bg-stone-200 text-stone-700 text-sm hover:bg-stone-300">Adicionar</button>
          </div>
        </div>
      </div>

      <div className="bg-white border border-stone-200 rounded p-4 max-w-md">
        <p className="text-sm font-medium text-stone-700 mb-1">Dados</p>
        <p className="text-xs text-stone-400 mb-3">{properties.length} imóveis · {transactions.length} movimentos · {reservations.length} reservas guardados neste dispositivo.</p>
        <div className="flex gap-2">
          <button onClick={exportData} className="inline-flex items-center gap-1.5 text-sm px-3 py-2 rounded border border-stone-300 text-stone-600 hover:bg-stone-50"><Download size={14} /> Exportar (JSON)</button>
          <button onClick={onClearAll} className="inline-flex items-center gap-1.5 text-sm px-3 py-2 rounded border border-red-200 text-red-700 hover:bg-red-50"><AlertTriangle size={14} /> Limpar tudo</button>
        </div>
      </div>
    </div>
  );
}
