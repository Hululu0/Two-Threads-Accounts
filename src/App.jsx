import React, { useState, useEffect, useMemo, useCallback } from "react";
import {
  LayoutDashboard, FileText, Receipt, Wallet, ScrollText, Package, ClipboardList,
  Plus, Trash2, X, LogOut, ShieldCheck, User, AlertCircle,
  CheckCircle2, Loader2, PieChart, Menu
} from "lucide-react";
import { supabase } from "./supabaseClient";
import * as XLSX from "xlsx";

/* ---------------------------------------------------------
   Constants & helpers
--------------------------------------------------------- */

const ACCOUNT_TYPES = [
  { key: "asset", label: "Asset" },
  { key: "liability", label: "Liability" },
  { key: "equity", label: "Equity" },
  { key: "income", label: "Income" },
  { key: "expense", label: "Expense" },
];

const DEFAULT_ACCOUNTS = [
  { id: "acc-1000", code: "1000", name: "Cash in Hand", type: "asset" },
  { id: "acc-1010", code: "1010", name: "Bank Account", type: "asset" },
  { id: "acc-1200", code: "1200", name: "Accounts Receivable", type: "asset" },
  { id: "acc-1300", code: "1300", name: "Inventory", type: "asset" },
  { id: "acc-2000", code: "2000", name: "Accounts Payable", type: "liability" },
  { id: "acc-3000", code: "3000", name: "Owner's Equity", type: "equity" },
  { id: "acc-4000", code: "4000", name: "Sales / Service Income", type: "income" },
  { id: "acc-5000", code: "5000", name: "General Expense", type: "expense" },
  { id: "acc-5010", code: "5010", name: "Rent", type: "expense" },
  { id: "acc-5020", code: "5020", name: "Utilities", type: "expense" },
  { id: "acc-5030", code: "5030", name: "Salaries", type: "expense" },
  { id: "acc-5040", code: "5040", name: "Cost of Goods Sold", type: "expense" },
];

const PAYMENT_METHODS = ["Cash", "Bank Transfer", "Mobile Banking (bKash/Nagad/Rocket)", "Card", "Cheque", "Other"];
const PRODUCT_COLORS = ["#F4A896", "#8FBFA3", "#F2C078", "#9FB8DD", "#D9A6C2", "#B7C4A8", "#E9A178", "#A9C9D9"];

function uid(prefix = "id") {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}
function todayStr() {
  return new Date().toISOString().slice(0, 10);
}
function fmtMoney(n) {
  const v = Number(n) || 0;
  const sign = v < 0 ? "-" : "";
  return `${sign}BDT ${Math.abs(v).toLocaleString("en-IN", { maximumFractionDigits: 2 })}`;
}
function normalSideFor(type) {
  return (type === "asset" || type === "expense") ? "debit" : "credit";
}
function colorFor(seed) {
  let h = 0;
  for (let i = 0; i < seed.length; i++) h = seed.charCodeAt(i) + ((h << 5) - h);
  return PRODUCT_COLORS[Math.abs(h) % PRODUCT_COLORS.length];
}

/* ---------------------------------------------------------
   Storage layer
--------------------------------------------------------- */

// Shared storage backed by a Supabase table (see README.md for setup).
// Table: kv_store(key text primary key, value jsonb)
async function storageGet(key, fallback) {
  try {
    const { data, error } = await supabase.from("kv_store").select("value").eq("key", key).maybeSingle();
    if (error || !data) return fallback;
    return data.value;
  } catch (e) { console.error("storage get failed", key, e); return fallback; }
}
async function storageSet(key, value) {
  try {
    const { error } = await supabase.from("kv_store").upsert({ key, value });
    if (error) { console.error("storage set failed", key, error); return false; }
    return true;
  } catch (e) { console.error("storage set failed", key, e); return false; }
}

const KEYS = {
  users: "ledger-users-v2",
  accounts: "ledger-accounts-v2",
  journal: "ledger-journal-v2",
  invoices: "ledger-invoices-v2",
  expenses: "ledger-expenses-v2",
  products: "ledger-products-v2",
  bills: "ledger-bills-v2",
};

/* ---------------------------------------------------------
   Design tokens — Pinterest-style: card-forward, rounded, airy
--------------------------------------------------------- */

const PALETTE = {
  bg: "#FAFAFA",
  card: "#FFFFFF",
  ink: "#1F2123",
  inkSoft: "#767A80",
  accent: "#E0483E",
  accentSoft: "#FDEAE8",
  credit: "#2F8F76",
  creditSoft: "#E4F4EE",
  debit: "#E0483E",
  line: "#EEEEEE",
  chip: "#F4F4F4",
  sidebarActive: "#FDEAE8",
};

const FONT = {
  display: "'Poppins', 'Segoe UI', sans-serif",
  body: "'Inter', 'Segoe UI', sans-serif",
  mono: "'IBM Plex Mono', 'Courier New', monospace",
};

function GlobalStyles() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@500;600;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');
      * { box-sizing: border-box; }
      body { margin: 0; }
      input, select { font-family: ${FONT.body}; }
      table { border-collapse: collapse; width: 100%; }
      .pin-btn { cursor: pointer; border: none; transition: transform .12s ease, box-shadow .12s ease, opacity .12s ease; }
      .pin-btn:hover { transform: translateY(-1px); opacity: 0.94; }
      .pin-btn:active { transform: translateY(0); }
      .pin-card { transition: box-shadow .15s ease, transform .15s ease; }
      .pin-card:hover { box-shadow: 0 10px 28px rgba(0,0,0,0.09); transform: translateY(-2px); }
      .row-hover:hover { background: #FAFAFA; }
      ::-webkit-scrollbar { width: 8px; height: 8px; }
      ::-webkit-scrollbar-thumb { background: #E4E4E4; border-radius: 4px; }

      .mobile-topbar { display: none; }
      .sidebar-overlay { display: none; }

      @media (max-width: 860px) {
        .app-shell { flex-direction: column; }
        .app-sidebar {
          position: fixed; top: 0; left: 0; height: 100vh; z-index: 200;
          transform: translateX(-100%); transition: transform .22s ease;
          box-shadow: 0 0 0 rgba(0,0,0,0);
        }
        .app-sidebar.open { transform: translateX(0); box-shadow: 10px 0 30px rgba(0,0,0,0.15); }
        .mobile-topbar {
          display: flex; align-items: center; justify-content: space-between;
          padding: 14px 16px; background: #fff; border-bottom: 1px solid ${PALETTE.line};
          position: sticky; top: 0; z-index: 60;
        }
        .sidebar-overlay.open {
          display: block; position: fixed; inset: 0; background: rgba(0,0,0,0.35); z-index: 150;
        }
        .app-main { padding: 18px 16px !important; }
        .responsive-grid { grid-template-columns: 1fr !important; }
        .responsive-form-row { grid-template-columns: 1fr !important; }
        table { min-width: 560px; }
      }
    `}</style>
  );
}

const styles = {
  bootScreen: { minHeight: "100vh", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", background: PALETTE.bg },
  appShell: { display: "flex", minHeight: "100vh", background: PALETTE.bg, fontFamily: FONT.body, color: PALETTE.ink },
  main: { flex: 1, padding: "30px 36px", minWidth: 0, position: "relative" },
  toast: {
    position: "fixed", top: 20, right: 24, background: "#fff", padding: "10px 18px",
    borderRadius: 999, borderLeft: "4px solid", boxShadow: "0 8px 24px rgba(0,0,0,0.12)",
    display: "flex", alignItems: "center", gap: 8, fontSize: 14, zIndex: 50,
  },
};

const labelStyle = { display: "block", fontSize: 12, color: PALETTE.inkSoft, marginBottom: 5, marginTop: 14, fontWeight: 600, letterSpacing: 0.2 };
const inputStyle = {
  width: "100%", padding: "10px 13px", borderRadius: 12, border: `1px solid ${PALETTE.line}`,
  fontSize: 14, background: "#fff", color: PALETTE.ink, outline: "none",
};

function PageHeader({ title, subtitle }) {
  return (
    <div style={{ marginBottom: 22 }}>
      <h1 style={{ fontFamily: FONT.display, fontSize: 25, margin: 0, color: PALETTE.ink, fontWeight: 600 }}>{title}</h1>
      {subtitle && <p style={{ color: PALETTE.inkSoft, fontSize: 13.5, marginTop: 5 }}>{subtitle}</p>}
    </div>
  );
}
function Card({ children, style, className }) {
  return (
    <div className={`pin-card ${className || ""}`} style={{ background: PALETTE.card, border: `1px solid ${PALETTE.line}`, borderRadius: 18, padding: 20, boxShadow: "0 1px 3px rgba(0,0,0,0.04)", ...style }}>
      {children}
    </div>
  );
}
function PrimaryButton({ children, onClick, type = "button", style, disabled }) {
  return (
    <button type={type} onClick={onClick} disabled={disabled} className="pin-btn" style={{
      display: "flex", alignItems: "center", gap: 6, background: PALETTE.accent, color: "#fff",
      padding: "10px 18px", borderRadius: 999, fontSize: 13.5, fontWeight: 600, fontFamily: FONT.body,
      opacity: disabled ? 0.5 : 1, ...style,
    }}>{children}</button>
  );
}
function GhostButton({ children, onClick, style }) {
  return (
    <button type="button" onClick={onClick} className="pin-btn" style={{
      background: "transparent", color: PALETTE.accent, fontSize: 13, fontWeight: 600,
      padding: "6px 4px", ...style,
    }}>{children}</button>
  );
}
function Th({ children, align }) {
  return <th style={{ textAlign: align || "left", padding: "8px 10px", fontSize: 11, color: PALETTE.inkSoft, fontWeight: 700, letterSpacing: 0.4, textTransform: "uppercase" }}>{children}</th>;
}
function Td({ children, align, mono }) {
  return <td style={{ textAlign: align || "left", padding: "11px 10px", fontSize: 13.5, fontFamily: mono ? FONT.mono : FONT.body }}>{children}</td>;
}
function EmptyState({ text }) {
  return <div style={{ padding: "30px 10px", textAlign: "center", color: PALETTE.inkSoft, fontSize: 13.5 }}>{text}</div>;
}
function Badge({ children, tone = "neutral" }) {
  const tones = {
    neutral: { bg: PALETTE.chip, color: PALETTE.inkSoft },
    good: { bg: PALETTE.creditSoft, color: PALETTE.credit },
    bad: { bg: PALETTE.accentSoft, color: PALETTE.debit },
  };
  const t = tones[tone];
  return <span style={{ fontSize: 11, padding: "3px 10px", borderRadius: 20, fontWeight: 600, background: t.bg, color: t.color }}>{children}</span>;
}

/* ---------------------------------------------------------
   Root App
--------------------------------------------------------- */

export default function App() {
  const [booting, setBooting] = useState(true);
  const [error, setError] = useState("");

  const [users, setUsers] = useState([]);
  const [accounts, setAccounts] = useState([]);
  const [journal, setJournal] = useState([]);
  const [invoices, setInvoices] = useState([]);
  const [expenses, setExpenses] = useState([]);
  const [products, setProducts] = useState([]);
  const [bills, setBills] = useState([]);

  const [currentUser, setCurrentUser] = useState(null);
  const [tab, setTab] = useState("dashboard");
  const [toast, setToast] = useState(null);
  const [mobileNavOpen, setMobileNavOpen] = useState(false);

  useEffect(() => {
    (async () => {
      try {
        const [u, a, j, inv, exp, prod, bl] = await Promise.all([
          storageGet(KEYS.users, null),
          storageGet(KEYS.accounts, null),
          storageGet(KEYS.journal, []),
          storageGet(KEYS.invoices, []),
          storageGet(KEYS.expenses, []),
          storageGet(KEYS.products, []),
          storageGet(KEYS.bills, []),
        ]);
        let finalUsers = u;
        if (!finalUsers) { finalUsers = []; await storageSet(KEYS.users, finalUsers); }
        let finalAccounts = a;
        if (!finalAccounts) { finalAccounts = DEFAULT_ACCOUNTS; await storageSet(KEYS.accounts, finalAccounts); }

        setUsers(finalUsers);
        setAccounts(finalAccounts);
        setJournal(j || []);
        setInvoices(inv || []);
        setExpenses(exp || []);
        setProducts(prod || []);
        setBills(bl || []);
      } catch (e) {
        setError("Failed to load data. Please refresh the page and try again.");
      } finally {
        setBooting(false);
      }
    })();
  }, []);

  const showToast = useCallback((msg, kind = "ok") => {
    setToast({ msg, kind });
    setTimeout(() => setToast(null), 2600);
  }, []);

  const persistUsers = async (n) => { setUsers(n); await storageSet(KEYS.users, n); };
  const persistAccounts = async (n) => { setAccounts(n); await storageSet(KEYS.accounts, n); };
  const persistJournal = async (n) => { setJournal(n); await storageSet(KEYS.journal, n); };
  const persistInvoices = async (n) => { setInvoices(n); await storageSet(KEYS.invoices, n); };
  const persistExpenses = async (n) => { setExpenses(n); await storageSet(KEYS.expenses, n); };
  const persistProducts = async (n) => { setProducts(n); await storageSet(KEYS.products, n); };
  const persistBills = async (n) => { setBills(n); await storageSet(KEYS.bills, n); };

  const balances = useMemo(() => {
    const map = {};
    accounts.forEach((a) => (map[a.id] = 0));
    journal.forEach((entry) => {
      entry.lines.forEach((line) => {
        if (!(line.accountId in map)) map[line.accountId] = 0;
        const acc = accounts.find((a) => a.id === line.accountId);
        const side = acc ? normalSideFor(acc.type) : "debit";
        const delta = (Number(line.debit) || 0) - (Number(line.credit) || 0);
        map[line.accountId] += side === "debit" ? delta : -delta;
      });
    });
    return map;
  }, [accounts, journal]);

  if (booting) {
    return (
      <div style={styles.bootScreen}>
        <Loader2 size={28} style={{ animation: "spin 1s linear infinite" }} />
        <p style={{ marginTop: 12, fontFamily: FONT.body, color: PALETTE.ink }}>Loading…</p>
        <style>{`@keyframes spin { from { transform: rotate(0deg);} to { transform: rotate(360deg);} }`}</style>
      </div>
    );
  }
  if (error) {
    return (
      <div style={styles.bootScreen}>
        <AlertCircle size={28} color={PALETTE.debit} />
        <p style={{ marginTop: 12, color: PALETTE.ink, textAlign: "center", maxWidth: 320 }}>{error}</p>
      </div>
    );
  }
  if (!currentUser) {
    return (
      <LoginScreen
        users={users}
        onCreateFirstAdmin={async (name, pin) => {
          const admin = { id: uid("usr"), name, pin, role: "admin" };
          await persistUsers([admin]);
          setCurrentUser(admin);
        }}
        onLogin={(u) => setCurrentUser(u)}
      />
    );
  }

  const arAccount = accounts.find((a) => a.name === "Accounts Receivable");
  const apAccount = accounts.find((a) => a.name === "Accounts Payable");
  const inventoryAccount = accounts.find((a) => a.name === "Inventory");
  const cogsAccount = accounts.find((a) => a.name === "Cost of Goods Sold");
  const incomeAccount = accounts.find((a) => a.type === "income");

  return (
    <div className="app-shell" style={styles.appShell}>
      <GlobalStyles />
      <div className="mobile-topbar">
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <img src="/logo.png" alt="Two Threads" style={{ height: 32, width: "auto" }} />
        </div>
        <button className="pin-btn" onClick={() => setMobileNavOpen(true)} style={{ background: "transparent", padding: 6 }}>
          <Menu size={22} color={PALETTE.ink} />
        </button>
      </div>
      <div className={`sidebar-overlay ${mobileNavOpen ? "open" : ""}`} onClick={() => setMobileNavOpen(false)} />
      <Sidebar
        tab={tab}
        setTab={(t) => { setTab(t); setMobileNavOpen(false); }}
        currentUser={currentUser}
        onLogout={() => setCurrentUser(null)}
        mobileOpen={mobileNavOpen}
        onCloseMobile={() => setMobileNavOpen(false)}
      />
      <main className="app-main" style={styles.main}>
        {toast && (
          <div style={{ ...styles.toast, borderColor: toast.kind === "ok" ? PALETTE.credit : PALETTE.debit }}>
            {toast.kind === "ok" ? <CheckCircle2 size={16} color={PALETTE.credit} /> : <AlertCircle size={16} color={PALETTE.debit} />}
            <span>{toast.msg}</span>
          </div>
        )}

        {tab === "dashboard" && <Dashboard accounts={accounts} balances={balances} journal={journal} products={products} />}

        {tab === "inventory" && (
          <Inventory
            products={products} currentUser={currentUser}
            onAdd={async (p) => { await persistProducts([{ ...p, id: uid("prod") }, ...products]); showToast("Product added"); }}
            onDelete={async (id) => { await persistProducts(products.filter((p) => p.id !== id)); showToast("Product deleted"); }}
            onImport={async (rows) => {
              const withIds = rows.map((p) => ({ ...p, id: uid("prod") }));
              await persistProducts([...withIds, ...products]);
              showToast(`${withIds.length} products imported`);
            }}
          />
        )}

        {tab === "bills" && (
          <Bills
            accounts={accounts} products={products} bills={bills} currentUser={currentUser}
            onAdd={async (bill, journalEntry, qtyChanges) => {
              const je = { ...journalEntry, id: uid("je"), createdBy: currentUser.name };
              await persistJournal([je, ...journal]);
              const nextProducts = products.map((p) => {
                const chg = qtyChanges.find((c) => c.productId === p.id);
                if (!chg) return p;
                const newQty = (Number(p.qty) || 0) + chg.qty;
                const newCost = newQty > 0 ? (((Number(p.qty) || 0) * (Number(p.costPrice) || 0) + chg.qty * chg.cost) / newQty) : chg.cost;
                return { ...p, qty: newQty, costPrice: Math.round(newCost * 100) / 100 };
              });
              await persistProducts(nextProducts);
              await persistBills([{ ...bill, id: uid("bill"), journalId: je.id, createdBy: currentUser.name }, ...bills]);
              showToast("Bill recorded");
            }}
            onMarkPaid={async (bill, paymentAccountId) => {
              const je = {
                id: uid("je"), date: todayStr(), memo: `Payment for Bill ${bill.billNo} — ${bill.vendor}`,
                lines: [
                  { accountId: apAccount.id, debit: bill.total, credit: 0 },
                  { accountId: paymentAccountId, debit: 0, credit: bill.total },
                ],
                createdBy: currentUser.name, source: "bill-payment", refId: bill.id,
              };
              await persistJournal([je, ...journal]);
              await persistBills(bills.map((b) => (b.id === bill.id ? { ...b, status: "paid", paymentAccountId } : b)));
              showToast("Bill marked as paid");
            }}
            onDelete={async (bill) => {
              const nextProducts = products.map((p) => {
                const item = bill.items.find((it) => it.productId === p.id);
                if (!item) return p;
                return { ...p, qty: (Number(p.qty) || 0) - item.qty };
              });
              await persistProducts(nextProducts);
              await persistBills(bills.filter((b) => b.id !== bill.id));
              await persistJournal(journal.filter((j) => j.id !== bill.journalId));
              showToast("Bill deleted");
            }}
          />
        )}

        {tab === "invoices" && (
          <Invoices
            accounts={accounts} products={products} invoices={invoices} currentUser={currentUser}
            onAdd={async (invoice, journalEntry, qtyChanges) => {
              const je = { ...journalEntry, id: uid("je"), createdBy: currentUser.name };
              await persistJournal([je, ...journal]);
              if (qtyChanges.length) {
                const nextProducts = products.map((p) => {
                  const chg = qtyChanges.find((c) => c.productId === p.id);
                  if (!chg) return p;
                  return { ...p, qty: (Number(p.qty) || 0) - chg.qty };
                });
                await persistProducts(nextProducts);
              }
              await persistInvoices([{ ...invoice, id: uid("inv"), journalId: je.id, createdBy: currentUser.name }, ...invoices]);
              showToast("Invoice created");
            }}
            onMarkPaid={async (invoice, paymentAccountId, paymentMethod) => {
              const je = {
                id: uid("je"), date: todayStr(), memo: `Payment received — Invoice ${invoice.number} (${paymentMethod})`,
                lines: [
                  { accountId: paymentAccountId, debit: invoice.total, credit: 0 },
                  { accountId: arAccount ? arAccount.id : "", debit: 0, credit: invoice.total },
                ],
                createdBy: currentUser.name, source: "invoice-payment", refId: invoice.id,
              };
              await persistJournal([je, ...journal]);
              await persistInvoices(invoices.map((i) => (i.id === invoice.id ? { ...i, status: "paid", paymentMethod, paymentAccountId } : i)));
              showToast("Marked as paid");
            }}
            onDelete={async (invoice) => {
              const productItems = invoice.items.filter((it) => it.productId);
              if (productItems.length) {
                const nextProducts = products.map((p) => {
                  const item = productItems.find((it) => it.productId === p.id);
                  if (!item) return p;
                  return { ...p, qty: (Number(p.qty) || 0) + (Number(item.qty) || 0) };
                });
                await persistProducts(nextProducts);
              }
              await persistInvoices(invoices.filter((i) => i.id !== invoice.id));
              await persistJournal(journal.filter((j) => j.id !== invoice.journalId));
              showToast("Invoice deleted");
            }}
          />
        )}

        {tab === "expenses" && (
          <Expenses
            accounts={accounts} expenses={expenses} currentUser={currentUser}
            onAdd={async (expense, journalEntry) => {
              const je = { ...journalEntry, id: uid("je"), createdBy: currentUser.name };
              await persistJournal([je, ...journal]);
              await persistExpenses([{ ...expense, id: uid("exp"), journalId: je.id, createdBy: currentUser.name }, ...expenses]);
              showToast("Expense added");
            }}
            onDelete={async (expense) => {
              await persistExpenses(expenses.filter((e) => e.id !== expense.id));
              await persistJournal(journal.filter((j) => j.id !== expense.journalId));
              showToast("Expense deleted");
            }}
          />
        )}

        {tab === "accounts" && (
          <ChartOfAccounts
            accounts={accounts} balances={balances} currentUser={currentUser}
            onAdd={async (acc) => { await persistAccounts([...accounts, { ...acc, id: uid("acc") }]); showToast("Account added"); }}
            onDelete={async (id) => { await persistAccounts(accounts.filter((a) => a.id !== id)); showToast("Account deleted"); }}
          />
        )}

        {tab === "journal" && (
          <Journal
            accounts={accounts} journal={journal} currentUser={currentUser}
            onAdd={async (entry) => { await persistJournal([{ ...entry, id: uid("je"), createdBy: currentUser.name }, ...journal]); showToast("Journal entry added"); }}
            onDelete={async (id) => { await persistJournal(journal.filter((j) => j.id !== id)); showToast("Entry deleted"); }}
          />
        )}

        {tab === "reports" && <Reports accounts={accounts} balances={balances} />}

        {tab === "users" && currentUser.role === "admin" && (
          <UsersPanel
            users={users}
            onAdd={async (u) => { await persistUsers([...users, { ...u, id: uid("usr") }]); showToast("User added"); }}
            onDelete={async (id) => { if (id === currentUser.id) return; await persistUsers(users.filter((u) => u.id !== id)); showToast("User deleted"); }}
          />
        )}
      </main>
    </div>
  );
}

/* ---------------------------------------------------------
   Login
--------------------------------------------------------- */

function LoginScreen({ users, onCreateFirstAdmin, onLogin }) {
  const [name, setName] = useState("");
  const [pin, setPin] = useState("");
  const [err, setErr] = useState("");
  const isFirstRun = users.length === 0;

  const submit = (e) => {
    e.preventDefault();
    setErr("");
    if (isFirstRun) {
      if (!name.trim() || pin.trim().length < 4) { setErr("Enter a name and set a PIN of at least 4 digits."); return; }
      onCreateFirstAdmin(name.trim(), pin.trim());
      return;
    }
    const match = users.find((u) => u.name.toLowerCase() === name.trim().toLowerCase() && u.pin === pin.trim());
    if (!match) { setErr("Name or PIN doesn't match."); return; }
    onLogin(match);
  };

  return (
    <div style={{ minHeight: "100vh", background: "linear-gradient(160deg, #FDEAE8 0%, #FAFAFA 55%)", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: FONT.body }}>
      <GlobalStyles />
      <form onSubmit={submit} style={{ background: "#fff", width: 380, maxWidth: "90vw", padding: "38px 34px", borderRadius: 24, boxShadow: "0 24px 60px rgba(0,0,0,0.12)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 6 }}>
          <img src="/logo.png" alt="Two Threads" style={{ height: 46, width: "auto" }} />
        </div>
        <p style={{ color: PALETTE.inkSoft, fontSize: 13, marginTop: 0, marginBottom: 22 }}>
          {isFirstRun ? "First time here — create your Admin account" : "Sign in with your name and PIN"}
        </p>
        <label style={labelStyle}>Name</label>
        <input value={name} onChange={(e) => setName(e.target.value)} placeholder="e.g. Rahim" style={inputStyle} />
        <label style={labelStyle}>PIN</label>
        <input value={pin} onChange={(e) => setPin(e.target.value)} type="password" placeholder="4+ digits" style={inputStyle} />
        {err && <div style={{ color: PALETTE.debit, fontSize: 13, marginTop: 8, display: "flex", gap: 6, alignItems: "center" }}><AlertCircle size={14} /> {err}</div>}
        <PrimaryButton type="submit" style={{ width: "100%", marginTop: 20, justifyContent: "center", padding: "12px 0" }}>
          {isFirstRun ? "Create Admin Account" : "Sign In"}
        </PrimaryButton>
        {!isFirstRun && <p style={{ fontSize: 12, color: PALETTE.inkSoft, marginTop: 16, lineHeight: 1.5 }}>To add a new user, ask your Admin — new team members can be added from the "Users" tab after logging in.</p>}
      </form>
    </div>
  );
}

/* ---------------------------------------------------------
   Sidebar
--------------------------------------------------------- */

function Sidebar({ tab, setTab, currentUser, onLogout, mobileOpen, onCloseMobile }) {
  const items = [
    { key: "dashboard", label: "Dashboard", icon: LayoutDashboard },
    { key: "invoices", label: "Invoices", icon: FileText },
    { key: "bills", label: "Bills", icon: ClipboardList },
    { key: "expenses", label: "Expenses", icon: Wallet },
    { key: "inventory", label: "Inventory", icon: Package },
    { key: "accounts", label: "Chart of Accounts", icon: ScrollText },
    { key: "journal", label: "Journal", icon: Receipt },
    { key: "reports", label: "Reports", icon: PieChart },
  ];
  if (currentUser.role === "admin") items.push({ key: "users", label: "Users", icon: User });

  return (
    <aside className={`app-sidebar ${mobileOpen ? "open" : ""}`} style={{ width: 236, background: "#fff", borderRight: `1px solid ${PALETTE.line}`, display: "flex", flexDirection: "column", padding: "26px 14px", flexShrink: 0 }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "0 8px", marginBottom: 26 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <img src="/logo.png" alt="Two Threads" style={{ height: 34, width: "auto" }} />
        </div>
        <button className="pin-btn" onClick={onCloseMobile} style={{ background: "transparent", padding: 4, display: mobileOpen ? "block" : "none" }}>
          <X size={20} color={PALETTE.inkSoft} />
        </button>
      </div>

      <nav style={{ flex: 1, display: "flex", flexDirection: "column", gap: 2 }}>
        {items.map((it) => {
          const active = tab === it.key;
          const Icon = it.icon;
          return (
            <button key={it.key} className="pin-btn" onClick={() => setTab(it.key)} style={{
              display: "flex", alignItems: "center", gap: 10, width: "100%", padding: "10px 14px",
              background: active ? PALETTE.sidebarActive : "transparent", borderRadius: 999,
              color: active ? PALETTE.accent : PALETTE.inkSoft, fontSize: 13.5, fontWeight: active ? 600 : 500,
              textAlign: "left", fontFamily: FONT.body,
            }}>
              <Icon size={16} /> {it.label}
            </button>
          );
        })}
      </nav>

      <div style={{ padding: "14px 10px 0", borderTop: `1px solid ${PALETTE.line}`, marginTop: 10 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, marginBottom: 10, marginTop: 14, color: PALETTE.ink }}>
          {currentUser.role === "admin" ? <ShieldCheck size={15} color={PALETTE.accent} /> : <User size={15} />}
          <span>{currentUser.name}</span>
          <span style={{ fontSize: 11, color: PALETTE.inkSoft }}>({currentUser.role === "admin" ? "Admin" : "Staff"})</span>
        </div>
        <button className="pin-btn" onClick={onLogout} style={{ display: "flex", alignItems: "center", gap: 8, background: "transparent", color: PALETTE.inkSoft, fontSize: 13, padding: "6px 4px" }}>
          <LogOut size={14} /> Log out
        </button>
      </div>
    </aside>
  );
}

/* ---------------------------------------------------------
   Dashboard
--------------------------------------------------------- */

function Dashboard({ accounts, balances, journal, products }) {
  const cash = accounts.filter((a) => a.type === "asset" && (a.name.includes("Cash") || a.name.includes("Bank"))).reduce((s, a) => s + (balances[a.id] || 0), 0);
  const receivable = accounts.filter((a) => a.name === "Accounts Receivable").reduce((s, a) => s + (balances[a.id] || 0), 0);
  const income = accounts.filter((a) => a.type === "income").reduce((s, a) => s + (balances[a.id] || 0), 0);
  const expenseTotal = accounts.filter((a) => a.type === "expense").reduce((s, a) => s + (balances[a.id] || 0), 0);
  const inventoryValue = products.reduce((s, p) => s + (Number(p.qty) || 0) * (Number(p.costPrice) || 0), 0);
  const recent = journal.slice(0, 6);
  const lowStock = products.filter((p) => (Number(p.qty) || 0) <= (Number(p.reorderLevel) || 3));

  const cards = [
    { label: "Cash + Bank", value: cash, color: PALETTE.credit },
    { label: "Receivable", value: receivable, color: PALETTE.accent },
    { label: "Inventory Value", value: inventoryValue, color: "#8F6FC9" },
    { label: "Total Income", value: income, color: PALETTE.credit },
    { label: "Total Expense", value: expenseTotal, color: PALETTE.debit },
  ];

  return (
    <div>
      <PageHeader title="Dashboard" subtitle="A quick overview of your business" />
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px,1fr))", gap: 14, marginBottom: 24 }}>
        {cards.map((c) => (
          <Card key={c.label}>
            <div style={{ fontSize: 12, color: PALETTE.inkSoft, fontWeight: 600, marginBottom: 8 }}>{c.label}</div>
            <div style={{ fontFamily: FONT.mono, fontSize: 20, color: c.color }}>{fmtMoney(c.value)}</div>
          </Card>
        ))}
      </div>

      <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "1.6fr 1fr", gap: 16 }}>
        <Card>
          <div style={{ fontFamily: FONT.display, fontSize: 15.5, marginBottom: 12, fontWeight: 600 }}>Recent Journal Entries</div>
          {recent.length === 0 ? <EmptyState text="No entries yet. Add an invoice, bill, or expense to see it here." /> : (
            <div style={{ overflowX: "auto" }}><table style={{ minWidth: 560 }}>
              <thead><tr style={{ borderBottom: `1px solid ${PALETTE.line}` }}><Th>Date</Th><Th>Memo</Th><Th align="right">Amount</Th><Th>By</Th></tr></thead>
              <tbody>
                {recent.map((j) => {
                  const total = j.lines.reduce((s, l) => s + (Number(l.debit) || 0), 0);
                  return (
                    <tr key={j.id} className="row-hover" style={{ borderBottom: `1px solid ${PALETTE.line}` }}>
                      <Td>{j.date}</Td><Td>{j.memo}</Td><Td align="right" mono>{fmtMoney(total)}</Td><Td>{j.createdBy}</Td>
                    </tr>
                  );
                })}
              </tbody>
            </table></div>
          )}
        </Card>

        <Card>
          <div style={{ fontFamily: FONT.display, fontSize: 15.5, marginBottom: 12, fontWeight: 600 }}>Low Stock</div>
          {lowStock.length === 0 ? <EmptyState text="All products are well stocked." /> : (
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {lowStock.map((p) => (
                <div key={p.id} style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <div style={{ width: 30, height: 30, borderRadius: 9, background: colorFor(p.name), flexShrink: 0, display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontSize: 12, fontWeight: 700 }}>
                    {p.name.slice(0, 1).toUpperCase()}
                  </div>
                  <div style={{ flex: 1, fontSize: 13 }}>{p.name}</div>
                  <Badge tone={(Number(p.qty) || 0) === 0 ? "bad" : "neutral"}>{p.qty} {p.unit || "pcs"}</Badge>
                </div>
              ))}
            </div>
          )}
        </Card>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Inventory — Pinterest-style product grid
--------------------------------------------------------- */

function Inventory({ products, currentUser, onAdd, onDelete, onImport }) {
  const [open, setOpen] = useState(false);
  const [form, setForm] = useState({ name: "", sku: "", unit: "pcs", qty: "", costPrice: "", salePrice: "", reorderLevel: "3" });
  const [importing, setImporting] = useState(false);
  const [importMsg, setImportMsg] = useState("");
  const isAdmin = currentUser.role === "admin";

  const submit = (e) => {
    e.preventDefault();
    if (!form.name.trim()) return;
    onAdd({
      name: form.name.trim(), sku: form.sku.trim(), unit: form.unit.trim() || "pcs",
      qty: Number(form.qty) || 0, costPrice: Number(form.costPrice) || 0, salePrice: Number(form.salePrice) || 0,
      reorderLevel: Number(form.reorderLevel) || 3,
    });
    setForm({ name: "", sku: "", unit: "pcs", qty: "", costPrice: "", salePrice: "", reorderLevel: "3" });
    setOpen(false);
  };

  const pick = (row, keys) => {
    for (const k of keys) {
      for (const rk of Object.keys(row)) {
        if (rk.trim().toLowerCase() === k.toLowerCase()) return row[rk];
      }
    }
    return undefined;
  };

  const handleFile = (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    setImporting(true);
    setImportMsg("");
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const sheetName = wb.SheetNames.find((n) => n.trim().toLowerCase() === "stock") || wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

        const parsed = [];
        for (const row of rows) {
          const name = String(pick(row, ["Product Name", "Name"]) || "").trim();
          if (!name) continue;
          const sku = String(pick(row, ["Product Code", "SKU"]) || "").trim();
          const qty = Number(pick(row, ["Quantity", "Qty", "Remaing Stock", "Remaining Stock"])) || 0;
          const costPrice = Number(pick(row, ["Purchase Price", "Cost Price"])) || 0;
          const salePrice = Number(pick(row, ["Selling Price", "Sale Price"])) || 0;
          parsed.push({ name, sku, unit: "pcs", qty, costPrice, salePrice, reorderLevel: 3 });
        }

        if (parsed.length === 0) {
          setImportMsg('No products found. Make sure the sheet has "Product Name", "Quantity", "Purchase Price" and "Selling Price" columns.');
          setImporting(false);
          return;
        }

        const totalValue = parsed.reduce((s, p) => s + p.qty * p.costPrice, 0);
        const ok = window.confirm(
          `${parsed.length} products found (total stock value ${fmtMoney(totalValue)}). Import them into Inventory?`
        );
        if (ok) {
          onImport(parsed);
        }
      } catch (err) {
        setImportMsg("Could not read this file. Please check it's a valid .xlsx export.");
      } finally {
        setImporting(false);
        e.target.value = "";
      }
    };
    reader.readAsArrayBuffer(file);
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 10 }}>
        <PageHeader title="Inventory" subtitle="Your products and current stock on hand" />
        {isAdmin && (
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            <label className="pin-btn" style={{
              display: "flex", alignItems: "center", gap: 6, background: "#fff", color: PALETTE.ink,
              padding: "10px 16px", borderRadius: 999, fontSize: 13.5, fontWeight: 600, border: `1px solid ${PALETTE.line}`, cursor: "pointer",
            }}>
              {importing ? <Loader2 size={15} style={{ animation: "spin 1s linear infinite" }} /> : <FileText size={15} />}
              {importing ? "Importing…" : "Import from Excel"}
              <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} style={{ display: "none" }} />
            </label>
            <PrimaryButton onClick={() => setOpen((v) => !v)}>{open ? <X size={15} /> : <Plus size={15} />} {open ? "Cancel" : "New Product"}</PrimaryButton>
          </div>
        )}
      </div>

      {importMsg && (
        <Card style={{ marginBottom: 16, borderColor: PALETTE.debit }}>
          <div style={{ color: PALETTE.debit, fontSize: 13, display: "flex", gap: 8, alignItems: "center" }}>
            <AlertCircle size={15} /> {importMsg}
          </div>
        </Card>
      )}

      {open && (
        <Card style={{ marginBottom: 22 }}>
          <form onSubmit={submit} className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 110px", gap: 12 }}>
            <div><label style={labelStyle}>Product Name</label><input style={inputStyle} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder="e.g. Blue Cotton Dress" /></div>
            <div><label style={labelStyle}>SKU (optional)</label><input style={inputStyle} value={form.sku} onChange={(e) => setForm({ ...form, sku: e.target.value })} placeholder="SKU-001" /></div>
            <div><label style={labelStyle}>Unit</label><input style={inputStyle} value={form.unit} onChange={(e) => setForm({ ...form, unit: e.target.value })} placeholder="pcs" /></div>
            <div><label style={labelStyle}>Opening Qty</label><input style={inputStyle} type="number" value={form.qty} onChange={(e) => setForm({ ...form, qty: e.target.value })} placeholder="0" /></div>
            <div><label style={labelStyle}>Cost Price (BDT)</label><input style={inputStyle} type="number" value={form.costPrice} onChange={(e) => setForm({ ...form, costPrice: e.target.value })} placeholder="0" /></div>
            <div><label style={labelStyle}>Sale Price (BDT)</label><input style={inputStyle} type="number" value={form.salePrice} onChange={(e) => setForm({ ...form, salePrice: e.target.value })} placeholder="0" /></div>
            <div style={{ gridColumn: "1 / -1", display: "flex", justifyContent: "space-between", alignItems: "end" }}>
              <div style={{ maxWidth: 160 }}><label style={labelStyle}>Reorder Alert Below</label><input style={inputStyle} type="number" value={form.reorderLevel} onChange={(e) => setForm({ ...form, reorderLevel: e.target.value })} /></div>
              <PrimaryButton type="submit">Add Product</PrimaryButton>
            </div>
          </form>
        </Card>
      )}

      {products.length === 0 ? (
        <Card><EmptyState text="No products yet — add your first product to start tracking stock." /></Card>
      ) : (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))", gap: 16 }}>
          {products.map((p) => {
            const low = (Number(p.qty) || 0) <= (Number(p.reorderLevel) || 3);
            const out = (Number(p.qty) || 0) <= 0;
            return (
              <Card key={p.id} style={{ padding: 0, overflow: "hidden" }}>
                <div style={{ height: 84, background: `linear-gradient(135deg, ${colorFor(p.name)}, ${colorFor(p.name + "x")})`, display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontSize: 30, fontWeight: 700, fontFamily: FONT.display }}>
                  {p.name.slice(0, 1).toUpperCase()}
                </div>
                <div style={{ padding: 16 }}>
                  <div style={{ fontWeight: 600, fontSize: 14.5, marginBottom: 2 }}>{p.name}</div>
                  <div style={{ fontSize: 11.5, color: PALETTE.inkSoft, marginBottom: 10 }}>{p.sku || "No SKU"}</div>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                    <Badge tone={out ? "bad" : low ? "neutral" : "good"}>{p.qty} {p.unit} in stock</Badge>
                    {isAdmin && (
                      <button className="pin-btn" onClick={() => onDelete(p.id)} style={{ background: "transparent", color: PALETTE.debit }}><Trash2 size={14} /></button>
                    )}
                  </div>
                  <div style={{ fontFamily: FONT.mono, fontSize: 12.5, color: PALETTE.inkSoft, display: "flex", justifyContent: "space-between" }}>
                    <span>Cost {fmtMoney(p.costPrice)}</span>
                    <span>Sale {fmtMoney(p.salePrice)}</span>
                  </div>
                </div>
              </Card>
            );
          })}
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   Bills — purchases that add stock (QuickBooks-style)
--------------------------------------------------------- */

function Bills({ accounts, products, bills, currentUser, onAdd, onMarkPaid, onDelete }) {
  const [open, setOpen] = useState(false);
  const [vendor, setVendor] = useState("");
  const [date, setDate] = useState(todayStr());
  const [items, setItems] = useState([{ productId: "", qty: 1, cost: "" }]);
  const [paymentAccountId, setPaymentAccountId] = useState("");
  const [payFor, setPayFor] = useState(null);
  const [payAccount, setPayAccount] = useState("");

  const apAccount = accounts.find((a) => a.name === "Accounts Payable");
  const inventoryAccount = accounts.find((a) => a.name === "Inventory");
  const cashAccounts = accounts.filter((a) => a.name.includes("Cash") || a.name.includes("Bank"));

  const total = items.reduce((s, it) => s + (Number(it.qty) || 0) * (Number(it.cost) || 0), 0);

  const updateItem = (i, field, value) => {
    const next = [...items];
    next[i] = { ...next[i], [field]: value };
    if (field === "productId") {
      const p = products.find((pr) => pr.id === value);
      if (p) next[i].cost = p.costPrice || "";
    }
    setItems(next);
  };

  const submit = (e) => {
    e.preventDefault();
    const validItems = items.filter((it) => it.productId && Number(it.qty) > 0);
    if (!vendor.trim() || validItems.length === 0 || !paymentAccountId || !inventoryAccount) return;
    const billNo = `BILL-${String(bills.length + 1).padStart(4, "0")}`;
    const bill = { billNo, vendor: vendor.trim(), date, items: validItems.map((it) => ({ productId: it.productId, qty: Number(it.qty), cost: Number(it.cost) || 0 })), total, status: paymentAccountId === apAccount?.id ? "unpaid" : "paid", paymentAccountId };
    const journalEntry = {
      date, memo: `Bill ${billNo} — ${vendor.trim()}`,
      lines: [{ accountId: inventoryAccount.id, debit: total, credit: 0 }, { accountId: paymentAccountId, debit: 0, credit: total }],
      source: "bill",
    };
    const qtyChanges = validItems.map((it) => ({ productId: it.productId, qty: Number(it.qty), cost: Number(it.cost) || 0 }));
    onAdd(bill, journalEntry, qtyChanges);
    setVendor(""); setItems([{ productId: "", qty: 1, cost: "" }]); setPaymentAccountId(""); setOpen(false);
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 10 }}>
        <PageHeader title="Bills" subtitle="Record a purchase — stock is added to inventory automatically" />
        <PrimaryButton onClick={() => setOpen((v) => !v)}>{open ? <X size={15} /> : <Plus size={15} />} {open ? "Cancel" : "New Bill"}</PrimaryButton>
      </div>

      {open && (
        <Card style={{ marginBottom: 20 }}>
          {products.length === 0 ? (
            <EmptyState text="Add a product in Inventory first, then you can record a bill for it." />
          ) : (
            <form onSubmit={submit}>
              <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 160px", gap: 10 }}>
                <div><label style={labelStyle}>Vendor</label><input style={inputStyle} value={vendor} onChange={(e) => setVendor(e.target.value)} placeholder="Supplier name" /></div>
                <div><label style={labelStyle}>Date</label><input type="date" style={inputStyle} value={date} onChange={(e) => setDate(e.target.value)} /></div>
              </div>

              <div style={{ marginTop: 14 }}>
                <label style={labelStyle}>Items Purchased</label>
                {items.map((it, i) => (
                  <div key={i} className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 90px 120px 30px", gap: 8, marginBottom: 8 }}>
                    <select style={inputStyle} value={it.productId} onChange={(e) => updateItem(i, "productId", e.target.value)}>
                      <option value="">Select product</option>
                      {products.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
                    </select>
                    <input style={inputStyle} type="number" placeholder="Qty" value={it.qty} onChange={(e) => updateItem(i, "qty", e.target.value)} />
                    <input style={inputStyle} type="number" placeholder="Cost/unit" value={it.cost} onChange={(e) => updateItem(i, "cost", e.target.value)} />
                    {items.length > 1 && <button type="button" className="pin-btn" onClick={() => setItems(items.filter((_, idx) => idx !== i))} style={{ background: "transparent", color: PALETTE.debit }}><Trash2 size={14} /></button>}
                  </div>
                ))}
                <GhostButton onClick={() => setItems([...items, { productId: "", qty: 1, cost: "" }])}>+ Add another item</GhostButton>
              </div>

              <div style={{ marginTop: 12, maxWidth: 260 }}>
                <label style={labelStyle}>Paid From</label>
                <select style={inputStyle} value={paymentAccountId} onChange={(e) => setPaymentAccountId(e.target.value)}>
                  <option value="">Select</option>
                  {cashAccounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
                  {apAccount && <option value={apAccount.id}>Record as unpaid (Accounts Payable)</option>}
                </select>
              </div>

              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 16 }}>
                <div style={{ fontFamily: FONT.mono, fontSize: 16 }}>Total: {fmtMoney(total)}</div>
                <PrimaryButton type="submit">Save Bill</PrimaryButton>
              </div>
            </form>
          )}
        </Card>
      )}

      <Card>
        {bills.length === 0 ? <EmptyState text="No bills recorded yet" /> : (
          <div style={{ overflowX: "auto" }}><table style={{ minWidth: 560 }}>
            <thead><tr style={{ borderBottom: `1px solid ${PALETTE.line}` }}><Th>Bill No.</Th><Th>Vendor</Th><Th>Date</Th><Th align="right">Total</Th><Th>Status</Th><Th> </Th></tr></thead>
            <tbody>
              {bills.map((b) => (
                <tr key={b.id} className="row-hover" style={{ borderBottom: `1px solid ${PALETTE.line}` }}>
                  <Td mono>{b.billNo}</Td><Td>{b.vendor}</Td><Td>{b.date}</Td><Td align="right" mono>{fmtMoney(b.total)}</Td>
                  <Td><Badge tone={b.status === "paid" ? "good" : "bad"}>{b.status === "paid" ? "Paid" : "Unpaid"}</Badge></Td>
                  <Td>
                    <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                      {b.status !== "paid" && (
                        payFor === b.id ? (
                          <>
                            <select style={{ ...inputStyle, padding: "4px 6px", fontSize: 12, width: "auto" }} value={payAccount} onChange={(e) => setPayAccount(e.target.value)}>
                              <option value="">Pay from…</option>
                              {cashAccounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
                            </select>
                            <button className="pin-btn" disabled={!payAccount} onClick={() => { onMarkPaid(b, payAccount); setPayFor(null); setPayAccount(""); }} style={{ background: PALETTE.credit, color: "#fff", fontSize: 12, padding: "5px 10px", borderRadius: 999, opacity: payAccount ? 1 : 0.5 }}>Confirm</button>
                          </>
                        ) : (
                          <GhostButton onClick={() => setPayFor(b.id)}>Mark Paid</GhostButton>
                        )
                      )}
                      {currentUser.role === "admin" && <button className="pin-btn" onClick={() => onDelete(b)} style={{ background: "transparent", color: PALETTE.debit }}><Trash2 size={14} /></button>}
                    </div>
                  </Td>
                </tr>
              ))}
            </tbody>
          </table></div>
        )}
      </Card>
    </div>
  );
}

/* ---------------------------------------------------------
   Invoices — items can link to inventory products
--------------------------------------------------------- */

function Invoices({ accounts, products, invoices, currentUser, onAdd, onMarkPaid, onDelete }) {
  const [open, setOpen] = useState(false);
  const [customer, setCustomer] = useState("");
  const [date, setDate] = useState(todayStr());
  const [dueDate, setDueDate] = useState(todayStr());
  const [items, setItems] = useState([{ productId: "", desc: "", qty: 1, rate: "" }]);
  const [payFor, setPayFor] = useState(null);
  const [payAccount, setPayAccount] = useState("");
  const [payMethod, setPayMethod] = useState("Cash");

  const incomeAccount = accounts.find((a) => a.type === "income");
  const arAccount = accounts.find((a) => a.name === "Accounts Receivable");
  const inventoryAccount = accounts.find((a) => a.name === "Inventory");
  const cogsAccount = accounts.find((a) => a.name === "Cost of Goods Sold");
  const cashAccounts = accounts.filter((a) => a.name.includes("Cash") || a.name.includes("Bank"));

  const total = items.reduce((s, it) => s + (Number(it.qty) || 0) * (Number(it.rate) || 0), 0);

  const updateItem = (i, field, value) => {
    const next = [...items];
    next[i] = { ...next[i], [field]: value };
    if (field === "productId") {
      const p = products.find((pr) => pr.id === value);
      if (p) { next[i].desc = p.name; next[i].rate = p.salePrice || ""; }
      else { next[i].desc = ""; }
    }
    setItems(next);
  };

  const submit = (e) => {
    e.preventDefault();
    if (!customer.trim() || total <= 0 || !arAccount || !incomeAccount) return;
    const number = `INV-${String(invoices.length + 1).padStart(4, "0")}`;
    const cleanItems = items.filter((it) => it.desc.trim() && (Number(it.qty) || 0) > 0);
    const invoice = { number, customer: customer.trim(), date, dueDate, items: cleanItems, total, status: "unpaid" };

    const lines = [{ accountId: arAccount.id, debit: total, credit: 0 }, { accountId: incomeAccount.id, debit: 0, credit: total }];
    const qtyChanges = [];
    let totalCogs = 0;
    cleanItems.forEach((it) => {
      if (it.productId) {
        const p = products.find((pr) => pr.id === it.productId);
        if (p) {
          totalCogs += (Number(it.qty) || 0) * (Number(p.costPrice) || 0);
          qtyChanges.push({ productId: it.productId, qty: Number(it.qty) || 0 });
        }
      }
    });
    if (totalCogs > 0 && inventoryAccount && cogsAccount) {
      lines.push({ accountId: cogsAccount.id, debit: totalCogs, credit: 0 });
      lines.push({ accountId: inventoryAccount.id, debit: 0, credit: totalCogs });
    }

    const journalEntry = { date, memo: `Invoice ${number} — ${customer.trim()}`, lines, source: "invoice" };
    onAdd(invoice, journalEntry, qtyChanges);
    setCustomer(""); setItems([{ productId: "", desc: "", qty: 1, rate: "" }]); setOpen(false);
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 10 }}>
        <PageHeader title="Invoices" subtitle="Create a new invoice — inventory and journal entries update automatically" />
        <PrimaryButton onClick={() => setOpen((v) => !v)}>{open ? <X size={15} /> : <Plus size={15} />} {open ? "Cancel" : "New Invoice"}</PrimaryButton>
      </div>

      {open && (
        <Card style={{ marginBottom: 20 }}>
          <form onSubmit={submit}>
            <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 160px 160px", gap: 10 }}>
              <div><label style={labelStyle}>Customer Name</label><input style={inputStyle} value={customer} onChange={(e) => setCustomer(e.target.value)} placeholder="Customer / company" /></div>
              <div><label style={labelStyle}>Date</label><input type="date" style={inputStyle} value={date} onChange={(e) => setDate(e.target.value)} /></div>
              <div><label style={labelStyle}>Due Date</label><input type="date" style={inputStyle} value={dueDate} onChange={(e) => setDueDate(e.target.value)} /></div>
            </div>

            <div style={{ marginTop: 14 }}>
              <label style={labelStyle}>Items</label>
              {items.map((it, i) => {
                const p = products.find((pr) => pr.id === it.productId);
                const short = p && Number(it.qty) > Number(p.qty);
                return (
                  <div key={i} style={{ marginBottom: 8 }}>
                    <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "160px 1fr 80px 120px 30px", gap: 8 }}>
                      <select style={inputStyle} value={it.productId} onChange={(e) => updateItem(i, "productId", e.target.value)}>
                        <option value="">Other (no stock)</option>
                        {products.map((pr) => <option key={pr.id} value={pr.id}>{pr.name}</option>)}
                      </select>
                      <input style={inputStyle} placeholder="Description" value={it.desc} onChange={(e) => updateItem(i, "desc", e.target.value)} disabled={!!it.productId} />
                      <input style={inputStyle} type="number" placeholder="Qty" value={it.qty} onChange={(e) => updateItem(i, "qty", e.target.value)} />
                      <input style={inputStyle} type="number" placeholder="Rate (BDT)" value={it.rate} onChange={(e) => updateItem(i, "rate", e.target.value)} />
                      {items.length > 1 && <button type="button" className="pin-btn" onClick={() => setItems(items.filter((_, idx) => idx !== i))} style={{ background: "transparent", color: PALETTE.debit }}><Trash2 size={14} /></button>}
                    </div>
                    {p && <div style={{ fontSize: 11.5, color: short ? PALETTE.debit : PALETTE.inkSoft, marginTop: 3, marginLeft: 2 }}>
                      {p.qty} {p.unit} in stock{short ? " — not enough stock, this will oversell" : ""}
                    </div>}
                  </div>
                );
              })}
              <GhostButton onClick={() => setItems([...items, { productId: "", desc: "", qty: 1, rate: "" }])}>+ Add another item</GhostButton>
            </div>

            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 16 }}>
              <div style={{ fontFamily: FONT.mono, fontSize: 16 }}>Total: {fmtMoney(total)}</div>
              <PrimaryButton type="submit">Create Invoice</PrimaryButton>
            </div>
          </form>
        </Card>
      )}

      <Card>
        {invoices.length === 0 ? <EmptyState text="No invoices yet" /> : (
          <div style={{ overflowX: "auto" }}><table style={{ minWidth: 560 }}>
            <thead><tr style={{ borderBottom: `1px solid ${PALETTE.line}` }}><Th>No.</Th><Th>Customer</Th><Th>Date</Th><Th align="right">Total</Th><Th>Status</Th><Th>Payment</Th><Th> </Th></tr></thead>
            <tbody>
              {invoices.map((inv) => (
                <tr key={inv.id} className="row-hover" style={{ borderBottom: `1px solid ${PALETTE.line}` }}>
                  <Td mono>{inv.number}</Td><Td>{inv.customer}</Td><Td>{inv.date}</Td><Td align="right" mono>{fmtMoney(inv.total)}</Td>
                  <Td><Badge tone={inv.status === "paid" ? "good" : "bad"}>{inv.status === "paid" ? "Paid" : "Unpaid"}</Badge></Td>
                  <Td style={{ fontSize: 12.5, color: PALETTE.inkSoft }}>{inv.status === "paid" ? (inv.paymentMethod || "—") : "—"}</Td>
                  <Td>
                    <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                      {inv.status !== "paid" && (
                        payFor === inv.id ? (
                          <>
                            <select style={{ ...inputStyle, padding: "4px 6px", fontSize: 12, width: "auto" }} value={payMethod} onChange={(e) => setPayMethod(e.target.value)}>
                              {PAYMENT_METHODS.map((m) => <option key={m} value={m}>{m}</option>)}
                            </select>
                            <select style={{ ...inputStyle, padding: "4px 6px", fontSize: 12, width: "auto" }} value={payAccount} onChange={(e) => setPayAccount(e.target.value)}>
                              <option value="">Deposit to…</option>
                              {cashAccounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
                            </select>
                            <button className="pin-btn" disabled={!payAccount} onClick={() => { onMarkPaid(inv, payAccount, payMethod); setPayFor(null); setPayAccount(""); setPayMethod("Cash"); }} style={{ background: PALETTE.credit, color: "#fff", fontSize: 12, padding: "5px 10px", borderRadius: 999, opacity: payAccount ? 1 : 0.5 }}>Confirm</button>
                          </>
                        ) : (
                          <GhostButton onClick={() => setPayFor(inv.id)}>Record Payment</GhostButton>
                        )
                      )}
                      {currentUser.role === "admin" && <button className="pin-btn" onClick={() => onDelete(inv)} style={{ background: "transparent", color: PALETTE.debit }}><Trash2 size={14} /></button>}
                    </div>
                  </Td>
                </tr>
              ))}
            </tbody>
          </table></div>
        )}
      </Card>
    </div>
  );
}

/* ---------------------------------------------------------
   Expenses
--------------------------------------------------------- */

function Expenses({ accounts, expenses, currentUser, onAdd, onDelete }) {
  const [open, setOpen] = useState(false);
  const [date, setDate] = useState(todayStr());
  const [vendor, setVendor] = useState("");
  const [accountId, setAccountId] = useState("");
  const [paymentAccountId, setPaymentAccountId] = useState("");
  const [amount, setAmount] = useState("");
  const [note, setNote] = useState("");

  const expenseAccounts = accounts.filter((a) => a.type === "expense" && a.name !== "Cost of Goods Sold");
  const cashAccounts = accounts.filter((a) => a.name.includes("Cash") || a.name.includes("Bank"));

  const submit = (e) => {
    e.preventDefault();
    const amt = Number(amount);
    if (!vendor.trim() || !accountId || !paymentAccountId || !(amt > 0)) return;
    const expense = { date, vendor: vendor.trim(), accountId, paymentAccountId, amount: amt, note: note.trim() };
    const journalEntry = { date, memo: `Expense — ${vendor.trim()}`, lines: [{ accountId, debit: amt, credit: 0 }, { accountId: paymentAccountId, debit: 0, credit: amt }], source: "expense" };
    onAdd(expense, journalEntry);
    setVendor(""); setAmount(""); setNote(""); setOpen(false);
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 10 }}>
        <PageHeader title="Expenses" subtitle="Operating costs like rent, utilities, and salaries" />
        <PrimaryButton onClick={() => setOpen((v) => !v)}>{open ? <X size={15} /> : <Plus size={15} />} {open ? "Cancel" : "New Expense"}</PrimaryButton>
      </div>

      {open && (
        <Card style={{ marginBottom: 20 }}>
          <form onSubmit={submit}>
            <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "140px 1fr 140px", gap: 10 }}>
              <div><label style={labelStyle}>Date</label><input type="date" style={inputStyle} value={date} onChange={(e) => setDate(e.target.value)} /></div>
              <div><label style={labelStyle}>Vendor / Paid To</label><input style={inputStyle} value={vendor} onChange={(e) => setVendor(e.target.value)} placeholder="e.g. Electricity Bill" /></div>
              <div><label style={labelStyle}>Amount (BDT)</label><input style={inputStyle} type="number" value={amount} onChange={(e) => setAmount(e.target.value)} placeholder="0" /></div>
            </div>
            <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: 10 }}>
              <div><label style={labelStyle}>Expense Category</label><select style={inputStyle} value={accountId} onChange={(e) => setAccountId(e.target.value)}><option value="">Select</option>{expenseAccounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}</select></div>
              <div><label style={labelStyle}>Paid From (Account)</label><select style={inputStyle} value={paymentAccountId} onChange={(e) => setPaymentAccountId(e.target.value)}><option value="">Select</option>{cashAccounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}</select></div>
            </div>
            <div style={{ marginTop: 10 }}><label style={labelStyle}>Note (optional)</label><input style={inputStyle} value={note} onChange={(e) => setNote(e.target.value)} placeholder="Additional details" /></div>
            <PrimaryButton type="submit" style={{ marginTop: 16 }}>Save Expense</PrimaryButton>
          </form>
        </Card>
      )}

      <Card>
        {expenses.length === 0 ? <EmptyState text="No expenses added yet" /> : (
          <div style={{ overflowX: "auto" }}><table style={{ minWidth: 560 }}>
            <thead><tr style={{ borderBottom: `1px solid ${PALETTE.line}` }}><Th>Date</Th><Th>Vendor</Th><Th>Category</Th><Th>Paid From</Th><Th align="right">Amount</Th><Th>By</Th><Th> </Th></tr></thead>
            <tbody>
              {expenses.map((exp) => {
                const acc = accounts.find((a) => a.id === exp.accountId);
                const paidFrom = accounts.find((a) => a.id === exp.paymentAccountId);
                return (
                  <tr key={exp.id} className="row-hover" style={{ borderBottom: `1px solid ${PALETTE.line}` }}>
                    <Td>{exp.date}</Td><Td>{exp.vendor}</Td><Td>{acc?.name || "—"}</Td>
                    <Td style={{ color: PALETTE.inkSoft, fontSize: 12.5 }}>{paidFrom?.name || "—"}</Td>
                    <Td align="right" mono>{fmtMoney(exp.amount)}</Td><Td>{exp.createdBy}</Td>
                    <Td>{currentUser.role === "admin" && <button className="pin-btn" onClick={() => onDelete(exp)} style={{ background: "transparent", color: PALETTE.debit }}><Trash2 size={14} /></button>}</Td>
                  </tr>
                );
              })}
            </tbody>
          </table></div>
        )}
      </Card>
    </div>
  );
}

/* ---------------------------------------------------------
   Chart of Accounts
--------------------------------------------------------- */

function ChartOfAccounts({ accounts, balances, currentUser, onAdd, onDelete }) {
  const [open, setOpen] = useState(false);
  const [form, setForm] = useState({ code: "", name: "", type: "asset" });
  const isAdmin = currentUser.role === "admin";
  const grouped = ACCOUNT_TYPES.map((t) => ({ ...t, items: accounts.filter((a) => a.type === t.key) }));

  const submit = (e) => {
    e.preventDefault();
    if (!form.code.trim() || !form.name.trim()) return;
    onAdd({ code: form.code.trim(), name: form.name.trim(), type: form.type });
    setForm({ code: "", name: "", type: "asset" });
    setOpen(false);
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 10 }}>
        <PageHeader title="Chart of Accounts" subtitle="All accounts in the business and their balances" />
        {isAdmin && <PrimaryButton onClick={() => setOpen((v) => !v)}>{open ? <X size={15} /> : <Plus size={15} />} {open ? "Cancel" : "New Account"}</PrimaryButton>}
      </div>

      {open && (
        <Card style={{ marginBottom: 20 }}>
          <form onSubmit={submit} className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "100px 1fr 160px auto", gap: 10, alignItems: "end" }}>
            <div><label style={labelStyle}>Code</label><input style={inputStyle} value={form.code} onChange={(e) => setForm({ ...form, code: e.target.value })} placeholder="6000" /></div>
            <div><label style={labelStyle}>Name</label><input style={inputStyle} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder="Account name" /></div>
            <div><label style={labelStyle}>Type</label><select style={inputStyle} value={form.type} onChange={(e) => setForm({ ...form, type: e.target.value })}>{ACCOUNT_TYPES.map((t) => <option key={t.key} value={t.key}>{t.label}</option>)}</select></div>
            <PrimaryButton type="submit">Add</PrimaryButton>
          </form>
        </Card>
      )}

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(320px,1fr))", gap: 16 }}>
        {grouped.map((g) => (
          <Card key={g.key}>
            <div style={{ fontFamily: FONT.display, fontSize: 15, marginBottom: 10, fontWeight: 600 }}>{g.label}</div>
            {g.items.length === 0 ? <EmptyState text="No accounts yet" /> : (
              <div style={{ overflowX: "auto" }}><table style={{ minWidth: 560 }}>
                <thead><tr style={{ borderBottom: `1px solid ${PALETTE.line}` }}><Th>Code</Th><Th>Name</Th><Th align="right">Balance</Th>{isAdmin && <Th align="right"> </Th>}</tr></thead>
                <tbody>
                  {g.items.map((a) => (
                    <tr key={a.id} className="row-hover" style={{ borderBottom: `1px solid ${PALETTE.line}` }}>
                      <Td mono>{a.code}</Td><Td>{a.name}</Td><Td align="right" mono>{fmtMoney(balances[a.id] || 0)}</Td>
                      {isAdmin && <Td align="right"><button className="pin-btn" onClick={() => onDelete(a.id)} style={{ background: "transparent", color: PALETTE.debit }}><Trash2 size={14} /></button></Td>}
                    </tr>
                  ))}
                </tbody>
              </table></div>
            )}
          </Card>
        ))}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Journal
--------------------------------------------------------- */

function Journal({ accounts, journal, currentUser, onAdd, onDelete }) {
  const [open, setOpen] = useState(false);
  const [date, setDate] = useState(todayStr());
  const [memo, setMemo] = useState("");
  const [lines, setLines] = useState([{ accountId: accounts[0]?.id || "", debit: "", credit: "" }, { accountId: accounts[1]?.id || "", debit: "", credit: "" }]);
  const isAdmin = currentUser.role === "admin";
  const totalDebit = lines.reduce((s, l) => s + (Number(l.debit) || 0), 0);
  const totalCredit = lines.reduce((s, l) => s + (Number(l.credit) || 0), 0);
  const balanced = totalDebit > 0 && totalDebit === totalCredit;

  const updateLine = (i, field, value) => { const next = [...lines]; next[i] = { ...next[i], [field]: value }; setLines(next); };

  const submit = (e) => {
    e.preventDefault();
    if (!memo.trim() || !balanced) return;
    onAdd({ date, memo: memo.trim(), lines: lines.filter((l) => l.accountId && (Number(l.debit) || Number(l.credit))) });
    setMemo(""); setLines([{ accountId: accounts[0]?.id || "", debit: "", credit: "" }, { accountId: accounts[1]?.id || "", debit: "", credit: "" }]); setOpen(false);
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 10 }}>
        <PageHeader title="Journal" subtitle="Log of every debit/credit entry (also auto-created from invoices, bills and expenses)" />
        <PrimaryButton onClick={() => setOpen((v) => !v)}>{open ? <X size={15} /> : <Plus size={15} />} {open ? "Cancel" : "Manual Entry"}</PrimaryButton>
      </div>

      {open && (
        <Card style={{ marginBottom: 20 }}>
          <form onSubmit={submit}>
            <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "160px 1fr", gap: 10, marginBottom: 10 }}>
              <div><label style={labelStyle}>Date</label><input type="date" style={inputStyle} value={date} onChange={(e) => setDate(e.target.value)} /></div>
              <div><label style={labelStyle}>Description / Memo</label><input style={inputStyle} value={memo} onChange={(e) => setMemo(e.target.value)} placeholder="What is this entry for" /></div>
            </div>
            {lines.map((l, i) => (
              <div key={i} className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 130px 130px 30px", gap: 8, marginBottom: 8, alignItems: "center" }}>
                <select style={inputStyle} value={l.accountId} onChange={(e) => updateLine(i, "accountId", e.target.value)}>
                  <option value="">Select account</option>
                  {accounts.map((a) => <option key={a.id} value={a.id}>{a.code} — {a.name}</option>)}
                </select>
                <input style={inputStyle} type="number" placeholder="Debit" value={l.debit} onChange={(e) => updateLine(i, "debit", e.target.value)} />
                <input style={inputStyle} type="number" placeholder="Credit" value={l.credit} onChange={(e) => updateLine(i, "credit", e.target.value)} />
                {lines.length > 2 && <button type="button" className="pin-btn" onClick={() => setLines(lines.filter((_, idx) => idx !== i))} style={{ background: "transparent", color: PALETTE.debit }}><Trash2 size={14} /></button>}
              </div>
            ))}
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 10 }}>
              <GhostButton onClick={() => setLines([...lines, { accountId: "", debit: "", credit: "" }])}>+ Add another line</GhostButton>
              <div style={{ fontFamily: FONT.mono, fontSize: 13 }}>
                Dr {fmtMoney(totalDebit)} &nbsp;|&nbsp; Cr {fmtMoney(totalCredit)}
                {!balanced && <span style={{ color: PALETTE.debit, marginLeft: 8 }}>(doesn't balance)</span>}
              </div>
            </div>
            <PrimaryButton type="submit" style={{ marginTop: 14 }} disabled={!balanced}>Save Entry</PrimaryButton>
          </form>
        </Card>
      )}

      <Card>
        {journal.length === 0 ? <EmptyState text="No journal entries yet" /> : (
          <div style={{ overflowX: "auto" }}><table style={{ minWidth: 560 }}>
            <thead><tr style={{ borderBottom: `1px solid ${PALETTE.line}` }}><Th>Date</Th><Th>Memo</Th><Th>Lines</Th><Th align="right">Total</Th><Th>By</Th>{isAdmin && <Th> </Th>}</tr></thead>
            <tbody>
              {journal.map((j) => {
                const total = j.lines.reduce((s, l) => s + (Number(l.debit) || 0), 0);
                return (
                  <tr key={j.id} className="row-hover" style={{ borderBottom: `1px solid ${PALETTE.line}`, verticalAlign: "top" }}>
                    <Td>{j.date}</Td><Td>{j.memo}</Td>
                    <Td>{j.lines.map((l, idx) => {
                      const acc = accounts.find((a) => a.id === l.accountId);
                      return <div key={idx} style={{ fontSize: 12, fontFamily: FONT.mono, color: PALETTE.inkSoft }}>{acc?.name || "—"}: {l.debit ? `Dr ${fmtMoney(l.debit)}` : `Cr ${fmtMoney(l.credit)}`}</div>;
                    })}</Td>
                    <Td align="right" mono>{fmtMoney(total)}</Td><Td>{j.createdBy}</Td>
                    {isAdmin && <Td><button className="pin-btn" onClick={() => onDelete(j.id)} style={{ background: "transparent", color: PALETTE.debit }}><Trash2 size={14} /></button></Td>}
                  </tr>
                );
              })}
            </tbody>
          </table></div>
        )}
      </Card>
    </div>
  );
}

/* ---------------------------------------------------------
   Reports
--------------------------------------------------------- */

function Reports({ accounts, balances }) {
  let totalDebit = 0, totalCredit = 0;
  const rows = accounts.map((a) => {
    const bal = balances[a.id] || 0;
    const side = normalSideFor(a.type);
    const debit = side === "debit" ? Math.max(bal, 0) : Math.max(-bal, 0);
    const credit = side === "credit" ? Math.max(bal, 0) : Math.max(-bal, 0);
    totalDebit += debit; totalCredit += credit;
    return { ...a, debit, credit };
  });

  return (
    <div>
      <PageHeader title="Trial Balance" subtitle="Debit and credit balances for every account — the two columns should match when the books are correct" />
      <Card>
        <div style={{ overflowX: "auto" }}><table style={{ minWidth: 560 }}>
          <thead><tr style={{ borderBottom: `2px solid ${PALETTE.ink}` }}><Th>Code</Th><Th>Account</Th><Th align="right">Debit</Th><Th align="right">Credit</Th></tr></thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.id} className="row-hover" style={{ borderBottom: `1px solid ${PALETTE.line}` }}>
                <Td mono>{r.code}</Td><Td>{r.name}</Td><Td align="right" mono>{r.debit ? fmtMoney(r.debit) : "—"}</Td><Td align="right" mono>{r.credit ? fmtMoney(r.credit) : "—"}</Td>
              </tr>
            ))}
          </tbody>
          <tfoot><tr style={{ borderTop: `2px solid ${PALETTE.ink}` }}><Td></Td><Td><b>Total</b></Td><Td align="right" mono><b>{fmtMoney(totalDebit)}</b></Td><Td align="right" mono><b>{fmtMoney(totalCredit)}</b></Td></tr></tfoot>
        </table></div>
        <div style={{ marginTop: 14, fontSize: 13, color: totalDebit === totalCredit ? PALETTE.credit : PALETTE.debit, display: "flex", alignItems: "center", gap: 6 }}>
          {totalDebit === totalCredit ? <CheckCircle2 size={15} /> : <AlertCircle size={15} />}
          {totalDebit === totalCredit ? "Books balance (Debit = Credit)" : "Books don't balance — check journal entries"}
        </div>
      </Card>
    </div>
  );
}

/* ---------------------------------------------------------
   Users panel
--------------------------------------------------------- */

function UsersPanel({ users, onAdd, onDelete }) {
  const [open, setOpen] = useState(false);
  const [name, setName] = useState("");
  const [pin, setPin] = useState("");
  const [role, setRole] = useState("staff");

  const submit = (e) => {
    e.preventDefault();
    if (!name.trim() || pin.trim().length < 4) return;
    onAdd({ name: name.trim(), pin: pin.trim(), role });
    setName(""); setPin(""); setRole("staff"); setOpen(false);
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 10 }}>
        <PageHeader title="Users" subtitle="Manage your team's access" />
        <PrimaryButton onClick={() => setOpen((v) => !v)}>{open ? <X size={15} /> : <Plus size={15} />} {open ? "Cancel" : "New User"}</PrimaryButton>
      </div>

      {open && (
        <Card style={{ marginBottom: 20 }}>
          <form onSubmit={submit} className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 140px 140px auto", gap: 10, alignItems: "end" }}>
            <div><label style={labelStyle}>Name</label><input style={inputStyle} value={name} onChange={(e) => setName(e.target.value)} /></div>
            <div><label style={labelStyle}>PIN (4+ digits)</label><input style={inputStyle} value={pin} onChange={(e) => setPin(e.target.value)} /></div>
            <div><label style={labelStyle}>Role</label><select style={inputStyle} value={role} onChange={(e) => setRole(e.target.value)}><option value="staff">Staff (limited)</option><option value="admin">Admin</option></select></div>
            <PrimaryButton type="submit">Add</PrimaryButton>
          </form>
        </Card>
      )}

      <Card>
        <div style={{ overflowX: "auto" }}><table style={{ minWidth: 560 }}>
          <thead><tr style={{ borderBottom: `1px solid ${PALETTE.line}` }}><Th>Name</Th><Th>Role</Th><Th> </Th></tr></thead>
          <tbody>
            {users.map((u) => (
              <tr key={u.id} className="row-hover" style={{ borderBottom: `1px solid ${PALETTE.line}` }}>
                <Td>{u.name}</Td><Td>{u.role === "admin" ? "Admin" : "Staff"}</Td>
                <Td><button className="pin-btn" onClick={() => onDelete(u.id)} style={{ background: "transparent", color: PALETTE.debit }}><Trash2 size={14} /></button></Td>
              </tr>
            ))}
          </tbody>
        </table></div>
      </Card>
      <p style={{ fontSize: 12, color: PALETTE.inkSoft, marginTop: 12 }}>Note: this PIN system is a simple access control for your 4–5 trusted team members — not banking-grade security. Don't share the link outside your team.</p>
    </div>
  );
}
