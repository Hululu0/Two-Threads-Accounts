import React, { useState, useEffect, useMemo, useCallback } from "react";
import {
  LayoutDashboard, FileText, Receipt, Wallet, ScrollText, Package, ClipboardList,
  Plus, Trash2, X, LogOut, ShieldCheck, User, AlertCircle,
  CheckCircle2, Loader2, PieChart, Menu, Search, Download, Eye, Edit2, TrendingUp
} from "lucide-react";
import { supabase } from "./supabaseClient";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";

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
  { id: "acc-1020", code: "1020", name: "bKash", type: "asset" },
  { id: "acc-1200", code: "1200", name: "Accounts Receivable", type: "asset" },
  { id: "acc-1300", code: "1300", name: "Inventory", type: "asset" },
  { id: "acc-2000", code: "2000", name: "Accounts Payable", type: "liability" },
  { id: "acc-3000", code: "3000", name: "Owner's Equity", type: "equity" },
  { id: "acc-4000", code: "4000", name: "Sales / Service Income", type: "income" },
  { id: "acc-5990", code: "5990", name: "Cost of Goods Sold", type: "expense" },
];

// The business's real expense categories — used to seed Chart of Accounts via the
// "Add Standard Categories" quick-setup button, and for brand-new installs.
const STANDARD_EXPENSE_CATEGORIES = [
  "Transportation", "Printing & Accessories", "Iffat Remuneration", "Raisa Remuneration",
  "Manager Salary", "Moderator 1 (Sadia) Salary", "Moderator 2 (Adiba) Salary", "Tailor",
  "Loan Repay", "Other Stitching Expense", "Purchase", "Office Rent", "Cost Of Investment",
  "Marketing Stock", "China Import", "Customer Refund", "Campaign", "Savings",
];

const PAYMENT_METHODS = ["Cash", "Bank Transfer", "Mobile Banking (bKash/Nagad/Rocket)", "Card", "Cheque", "Other"];
const ALL_TABS = [
  { key: "dashboard", label: "Dashboard" },
  { key: "invoices", label: "Invoices" },
  { key: "bills", label: "Bills" },
  { key: "expenses", label: "Expenses" },
  { key: "income", label: "Income" },
  { key: "inventory", label: "Inventory" },
  { key: "accounts", label: "Chart of Accounts" },
  { key: "journal", label: "Journal" },
  { key: "reports", label: "Reports" },
];
const MEDIA_OPTIONS = ["Facebook", "Instagram", "TikTok", "WhatsApp", "Website", "Walk-in", "Other"];
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

function generateInvoicePDF(invoice) {
  const doc = new jsPDF({ unit: "pt", format: "a4" });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const marginX = 48;
  let y = 50;

  const INK = [30, 26, 27];
  const INK_SOFT = [120, 112, 113];
  const ACCENT = [200, 50, 75];
  const ACCENT_SOFT = [253, 233, 236];
  const LINE = [225, 218, 219];
  const ROW_ALT = [250, 246, 247];

  const paidTotal = (invoice.payments || []).reduce((s, p) => s + (Number(p.amount) || 0), 0);
  const balanceDue = invoice.total - paidTotal;

  // ---- Header: brand (left) + Invoice number (right) ----
  doc.setFont("helvetica", "bold");
  doc.setFontSize(16);
  doc.setTextColor(...INK);
  doc.text("Two Threads", marginX, y);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8.5);
  doc.setTextColor(...INK_SOFT);
  doc.text("Crafted elegance, made yours", marginX, y + 14);

  doc.setFont("helvetica", "bold");
  doc.setFontSize(15);
  doc.setTextColor(...ACCENT);
  doc.text(`Invoice ${invoice.number}`, pageWidth - marginX, y, { align: "right" });

  y += 40;
  doc.setDrawColor(...INK);
  doc.setLineWidth(1.4);
  doc.line(marginX, y, pageWidth - marginX, y);
  doc.setLineWidth(1);
  y += 26;

  // ---- Bill To ----
  const billToTop = y;
  doc.setFont("helvetica", "bold");
  doc.setFontSize(9);
  doc.setTextColor(...INK_SOFT);
  doc.text("BILL TO", marginX, y);
  y += 15;
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.setTextColor(...INK);
  doc.text(invoice.customer || "-", marginX, y);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9.5);
  doc.setTextColor(...INK_SOFT);
  if (invoice.address) { y += 14; doc.text(invoice.address, marginX, y, { maxWidth: 220 }); }
  if (invoice.phone) { y += 14; doc.text(`Phone: ${invoice.phone}`, marginX, y); }

  // ---- Date / Please Pay / Due Date boxes (right side) ----
  const boxW = 96, boxH = 46, boxGap = 6;
  const boxesRight = pageWidth - marginX;
  const boxesLeft = boxesRight - (boxW * 3 + boxGap * 2);
  const boxY = billToTop - 6;

  const drawBox = (x, label, value, filled) => {
    doc.setFillColor(...(filled ? ACCENT : ACCENT_SOFT));
    doc.roundedRect(x, boxY, boxW, boxH, 4, 4, "F");
    doc.setFont("helvetica", "normal");
    doc.setFontSize(7.5);
    doc.setTextColor(...(filled ? [255, 255, 255] : ACCENT));
    doc.text(label, x + boxW / 2, boxY + 15, { align: "center" });
    doc.setFont("helvetica", "bold");
    doc.setFontSize(10.5);
    doc.text(value, x + boxW / 2, boxY + 32, { align: "center" });
  };
  drawBox(boxesLeft, "DATE", invoice.date || "-", false);
  drawBox(boxesLeft + boxW + boxGap, "PLEASE PAY", fmtMoney(balanceDue), true);
  drawBox(boxesLeft + (boxW + boxGap) * 2, "DUE DATE", invoice.dueDate || "-", false);

  y = Math.max(y, boxY + boxH) + 30;

  // ---- Items table ----
  doc.setFont("helvetica", "bold");
  doc.setFontSize(8.5);
  doc.setTextColor(...INK_SOFT);
  doc.text("DESCRIPTION", marginX, y);
  doc.text("QTY", pageWidth - marginX - 190, y, { align: "right" });
  doc.text("RATE", pageWidth - marginX - 100, y, { align: "right" });
  doc.text("AMOUNT", pageWidth - marginX, y, { align: "right" });
  y += 8;
  doc.setDrawColor(...INK);
  doc.line(marginX, y, pageWidth - marginX, y);
  y += 6;

  doc.setFont("helvetica", "normal");
  doc.setFontSize(10);
  (invoice.items || []).forEach((it, idx) => {
    const rowY = y + 17;
    if (idx % 2 === 1) {
      doc.setFillColor(...ROW_ALT);
      doc.rect(marginX, y, pageWidth - marginX * 2, 24, "F");
    }
    doc.setTextColor(...INK);
    doc.text(String(it.desc || ""), marginX, rowY - 3, { maxWidth: pageWidth - marginX * 2 - 220 });
    doc.text(String(it.qty), pageWidth - marginX - 190, rowY - 3, { align: "right" });
    doc.text(fmtMoney(it.rate), pageWidth - marginX - 100, rowY - 3, { align: "right" });
    doc.text(fmtMoney((Number(it.qty) || 0) * (Number(it.rate) || 0)), pageWidth - marginX, rowY - 3, { align: "right" });
    y += 24;
  });
  y += 4;
  doc.setDrawColor(...LINE);
  doc.line(marginX, y, pageWidth - marginX, y);
  y += 26;

  // ---- Note (left) + Summary (right), side by side ----
  const noteTop = y;
  doc.setFont("helvetica", "bold");
  doc.setFontSize(9);
  doc.setTextColor(...INK);
  doc.text("Dear Customer,", marginX, y);
  y += 13;
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8.5);
  doc.setTextColor(...INK_SOFT);
  const noteLines = doc.splitTextToSize(
    "Thank you for shopping with Two Threads. Please check your items on delivery — if anything is missing or damaged, contact us right away so we can help.",
    260
  );
  doc.text(noteLines, marginX, y);

  let sy = noteTop;
  const summaryX = pageWidth - marginX - 190;
  doc.setFont("helvetica", "normal");
  doc.setFontSize(10);
  if (invoice.discount > 0) {
    doc.setTextColor(...INK_SOFT);
    doc.text("Subtotal", summaryX, sy);
    doc.setTextColor(...INK);
    doc.text(fmtMoney(invoice.subtotal != null ? invoice.subtotal : invoice.total), pageWidth - marginX, sy, { align: "right" });
    sy += 17;
    doc.setTextColor(...INK_SOFT);
    doc.text("Discount", summaryX, sy);
    doc.setTextColor(...INK);
    doc.text(`-${fmtMoney(invoice.discount)}`, pageWidth - marginX, sy, { align: "right" });
    sy += 17;
  }
  doc.setTextColor(...INK_SOFT);
  doc.text("Total", summaryX, sy);
  doc.setTextColor(...INK);
  doc.text(fmtMoney(invoice.total), pageWidth - marginX, sy, { align: "right" });
  sy += 17;
  if (paidTotal > 0) {
    doc.setTextColor(...INK_SOFT);
    doc.text("Deposit / Paid", summaryX, sy);
    doc.setTextColor(...INK);
    doc.text(fmtMoney(paidTotal), pageWidth - marginX, sy, { align: "right" });
    sy += 17;
  }
  sy += 4;
  doc.setDrawColor(...LINE);
  doc.line(summaryX - 10, sy - 12, pageWidth - marginX, sy - 12);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(9.5);
  doc.setTextColor(...ACCENT);
  doc.text("TOTAL DUE", summaryX, sy);
  doc.setFontSize(15);
  doc.text(fmtMoney(balanceDue), pageWidth - marginX, sy + 18, { align: "right" });

  y = Math.max(noteTop + 13 + noteLines.length * 11, sy + 30) + 20;

  if ((invoice.payments || []).length > 0) {
    doc.setFont("helvetica", "bold");
    doc.setFontSize(8.5);
    doc.setTextColor(...INK_SOFT);
    doc.text("PAYMENT HISTORY", marginX, y);
    y += 13;
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8.5);
    invoice.payments.forEach((p) => {
      doc.text(`${p.date} — ${fmtMoney(p.amount)} (${p.method || "Payment"})`, marginX, y);
      y += 12;
    });
    y += 10;
  }

  // ---- Footer ----
  const footerY = Math.max(y + 30, pageHeight - 90);
  doc.setDrawColor(...LINE);
  doc.line(marginX, footerY, pageWidth - marginX, footerY);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.setTextColor(...ACCENT);
  doc.text("Thank you for your order!", marginX, footerY + 22);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);
  doc.setTextColor(...INK_SOFT);
  doc.text("Two Threads", marginX, footerY + 36);

  doc.setDrawColor(...INK_SOFT);
  doc.line(pageWidth - marginX - 150, footerY + 32, pageWidth - marginX, footerY + 32);
  doc.text("Customer Signature", pageWidth - marginX - 150, footerY + 44);

  doc.save(`${invoice.number}.pdf`);
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
  incomeEntries: "ledger-income-v2",
};
const SESSION_KEY = "two-threads-session-user-id";

/* ---------------------------------------------------------
   Design tokens — deep wine/maroon gradient, glass cards
--------------------------------------------------------- */

const PALETTE = {
  bg: "#2b070d",
  card: "rgba(255,255,255,0.055)",
  cardSolid: "#4a1119",
  ink: "#F5EDEE",
  inkSoft: "rgba(245,237,238,0.62)",
  accent: "#E24C63",
  accentSoft: "rgba(226,76,99,0.18)",
  credit: "#4ADE95",
  creditSoft: "rgba(74,222,149,0.16)",
  debit: "#FF7A85",
  debitSoft: "rgba(255,122,133,0.16)",
  line: "rgba(255,255,255,0.14)",
  chip: "rgba(255,255,255,0.09)",
  sidebarActive: "rgba(226,76,99,0.22)",
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
      html, body { overflow-x: hidden; width: 100%; }
      html { background: ${PALETTE.bg}; }
      #root { min-height: 100vh; width: 100%; }
      body {
        margin: 0;
        min-height: 100vh;
        background-color: ${PALETTE.bg};
        background-image:
          radial-gradient(1300px circle at 18% -8%, #7d1c29 0%, transparent 55%),
          radial-gradient(1100px circle at 100% 10%, #601420 0%, transparent 50%),
          radial-gradient(1000px circle at 10% 100%, #55111c 0%, transparent 55%);
        background-attachment: fixed;
        color: ${PALETTE.ink};
      }
      input, select { font-family: ${FONT.body}; }
      input::placeholder { color: rgba(245,237,238,0.38); }
      input[type="date"]::-webkit-calendar-picker-indicator { filter: invert(1); opacity: 0.7; }
      select option { background: ${PALETTE.cardSolid}; color: ${PALETTE.ink}; }
      table { border-collapse: collapse; width: 100%; }
      .pin-btn { cursor: pointer; border: none; transition: transform .12s ease, box-shadow .12s ease, opacity .12s ease; }
      .pin-btn:hover { transform: translateY(-1px); opacity: 0.94; }
      .pin-btn:active { transform: translateY(0); }
      .pin-card { transition: box-shadow .15s ease, transform .15s ease, background .15s ease; backdrop-filter: blur(14px); -webkit-backdrop-filter: blur(14px); }
      .pin-card:hover { background: rgba(255,255,255,0.08); box-shadow: 0 14px 34px rgba(0,0,0,0.32); transform: translateY(-2px); }
      .row-hover:hover { background: rgba(255,255,255,0.05); }
      ::-webkit-scrollbar { width: 8px; height: 8px; }
      ::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.18); border-radius: 4px; }

      .mobile-topbar { display: none; }
      .sidebar-overlay { display: none; }

      @media (max-width: 860px) {
        .app-shell { flex-direction: column; }
        .app-sidebar {
          position: fixed; top: 0; left: 0; height: 100vh; z-index: 200;
          transform: translateX(-100%); transition: transform .22s ease;
          box-shadow: 0 0 0 rgba(0,0,0,0);
        }
        .app-sidebar.open { transform: translateX(0); box-shadow: 10px 0 30px rgba(0,0,0,0.4); }
        .mobile-topbar {
          display: flex; align-items: center; justify-content: space-between;
          padding: 14px 16px; background: ${PALETTE.cardSolid}; border-bottom: 1px solid ${PALETTE.line};
          position: sticky; top: 0; z-index: 60;
        }
        .sidebar-overlay.open {
          display: block; position: fixed; inset: 0; background: rgba(0,0,0,0.5); z-index: 150;
        }
        .app-main { padding: 18px 16px !important; width: 100%; max-width: 100vw; box-sizing: border-box; }
        .responsive-grid { grid-template-columns: 1fr !important; }
        .responsive-form-row { grid-template-columns: 1fr !important; }
        table { min-width: 560px; }
      }
    `}</style>
  );
}

const styles = {
  bootScreen: { minHeight: "100vh", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", background: "transparent", color: PALETTE.ink },
  appShell: { display: "flex", minHeight: "100vh", background: "transparent", fontFamily: FONT.body, color: PALETTE.ink },
  main: { flex: 1, padding: "30px 36px", minWidth: 0, position: "relative" },
  toast: {
    position: "fixed", top: 20, right: 24, background: PALETTE.cardSolid, color: PALETTE.ink, padding: "10px 18px",
    borderRadius: 999, borderLeft: "4px solid", boxShadow: "0 8px 24px rgba(0,0,0,0.35)",
    display: "flex", alignItems: "center", gap: 8, fontSize: 14, zIndex: 50,
  },
};

const labelStyle = { display: "block", fontSize: 12, color: PALETTE.inkSoft, marginBottom: 5, marginTop: 14, fontWeight: 600, letterSpacing: 0.2 };
const inputStyle = {
  width: "100%", padding: "10px 13px", borderRadius: 12, border: `1px solid ${PALETTE.line}`,
  fontSize: 14, background: "rgba(255,255,255,0.06)", color: PALETTE.ink, outline: "none",
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
    <div className={`pin-card ${className || ""}`} style={{ background: PALETTE.card, border: `1px solid ${PALETTE.line}`, borderRadius: 18, padding: 20, boxShadow: "0 4px 18px rgba(0,0,0,0.22)", ...style }}>
      {children}
    </div>
  );
}
function PrimaryButton({ children, onClick, type = "button", style, disabled }) {
  return (
    <button type={type} onClick={onClick} disabled={disabled} className="pin-btn" style={{
      display: "flex", alignItems: "center", gap: 6, background: PALETTE.accent, color: "#fff",
      padding: "10px 18px", borderRadius: 999, fontSize: 13.5, fontWeight: 600, fontFamily: FONT.body,
      boxShadow: "0 6px 18px rgba(226,76,99,0.35)", opacity: disabled ? 0.5 : 1, ...style,
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
function Td({ children, align, mono, ...rest }) {
  return <td {...rest} style={{ textAlign: align || "left", padding: "11px 10px", fontSize: 13.5, fontFamily: mono ? FONT.mono : FONT.body, color: PALETTE.ink, ...(rest.style || {}) }}>{children}</td>;
}
function EmptyState({ text }) {
  return <div style={{ padding: "30px 10px", textAlign: "center", color: PALETTE.inkSoft, fontSize: 13.5 }}>{text}</div>;
}
function Badge({ children, tone = "neutral" }) {
  const tones = {
    neutral: { bg: PALETTE.chip, color: PALETTE.inkSoft },
    good: { bg: PALETTE.creditSoft, color: PALETTE.credit },
    bad: { bg: PALETTE.debitSoft, color: PALETTE.debit },
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
  const [incomeEntries, setIncomeEntries] = useState([]);

  const [currentUser, setCurrentUser] = useState(null);
  const [tab, setTab] = useState("dashboard");
  const [toast, setToast] = useState(null);
  const [mobileNavOpen, setMobileNavOpen] = useState(false);

  useEffect(() => {
    (async () => {
      try {
        const [u, a, j, inv, exp, prod, bl, inc] = await Promise.all([
          storageGet(KEYS.users, null),
          storageGet(KEYS.accounts, null),
          storageGet(KEYS.journal, []),
          storageGet(KEYS.invoices, []),
          storageGet(KEYS.expenses, []),
          storageGet(KEYS.products, []),
          storageGet(KEYS.bills, []),
          storageGet(KEYS.incomeEntries, []),
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
        setIncomeEntries(inc || []);

        const savedUserId = localStorage.getItem(SESSION_KEY);
        if (savedUserId) {
          const match = finalUsers.find((u) => u.id === savedUserId);
          if (match) setCurrentUser(match);
          else localStorage.removeItem(SESSION_KEY);
        }
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
  const persistIncomeEntries = async (n) => { setIncomeEntries(n); await storageSet(KEYS.incomeEntries, n); };

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

  useEffect(() => {
    if (!currentUser) return;
    if (tab === "users" && currentUser.role !== "admin") { setTab("dashboard"); return; }
    if (currentUser.role !== "admin" && Array.isArray(currentUser.permissions) && !currentUser.permissions.includes(tab)) {
      setTab("dashboard");
    }
  }, [tab, currentUser]);

  // Keep the logged-in session in sync with the latest user record (role/permission
  // changes take effect immediately); log out automatically if the account was removed.
  useEffect(() => {
    if (!currentUser || booting) return;
    const latest = users.find((u) => u.id === currentUser.id);
    if (!latest) {
      localStorage.removeItem(SESSION_KEY);
      setCurrentUser(null);
    } else if (JSON.stringify(latest) !== JSON.stringify(currentUser)) {
      setCurrentUser(latest);
    }
  }, [users, currentUser, booting]);

  if (booting) {
    return (
      <div style={styles.bootScreen}>
        <GlobalStyles />
        <Loader2 size={28} style={{ animation: "spin 1s linear infinite" }} />
        <p style={{ marginTop: 12, fontFamily: FONT.body, color: PALETTE.ink }}>Loading…</p>
        <style>{`@keyframes spin { from { transform: rotate(0deg);} to { transform: rotate(360deg);} }`}</style>
      </div>
    );
  }
  if (error) {
    return (
      <div style={styles.bootScreen}>
        <GlobalStyles />
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
          localStorage.setItem(SESSION_KEY, admin.id);
          setTab("dashboard");
        }}
        onLogin={(u) => { setCurrentUser(u); localStorage.setItem(SESSION_KEY, u.id); setTab("dashboard"); }}
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
          <div style={{ background: "#fff", borderRadius: 9, padding: "4px 7px", display: "inline-flex" }}>
            <img src="/logo.png" alt="Two Threads" style={{ height: 20, width: "auto", display: "block" }} />
          </div>
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
        onLogout={() => { localStorage.removeItem(SESSION_KEY); setCurrentUser(null); }}
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
            onEdit={async (id, updates) => { await persistProducts(products.map((p) => (p.id === id ? { ...p, ...updates } : p))); showToast("Product updated"); }}
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
            onAdd={async (invoice, cogsJournalEntry, qtyChanges, depositInfo) => {
              if (qtyChanges.length) {
                const nextProducts = products.map((p) => {
                  const chg = qtyChanges.find((c) => c.productId === p.id);
                  if (!chg) return p;
                  return { ...p, qty: (Number(p.qty) || 0) - chg.qty };
                });
                await persistProducts(nextProducts);
              }
              const newInvoice = { ...invoice, id: uid("inv"), createdBy: currentUser.name };
              let journalToSave = journal;
              if (cogsJournalEntry) {
                const je = { ...cogsJournalEntry, id: uid("je"), createdBy: currentUser.name };
                journalToSave = [je, ...journalToSave];
                newInvoice.journalId = je.id;
              }
              if (depositInfo && depositInfo.amount > 0 && depositInfo.accountId && incomeAccount) {
                const payJe = {
                  id: uid("je"), date: depositInfo.date || todayStr(), memo: `Deposit received — Invoice ${invoice.number} (${depositInfo.method})`,
                  lines: [{ accountId: depositInfo.accountId, debit: depositInfo.amount, credit: 0 }, { accountId: incomeAccount.id, debit: 0, credit: depositInfo.amount }],
                  createdBy: currentUser.name, source: "invoice-payment", refId: newInvoice.id,
                };
                journalToSave = [payJe, ...journalToSave];
                newInvoice.payments = [{ id: uid("pay"), date: depositInfo.date || todayStr(), amount: depositInfo.amount, accountId: depositInfo.accountId, method: depositInfo.method }];
                newInvoice.status = depositInfo.amount >= invoice.total ? "paid" : "partial";
              }
              await persistJournal(journalToSave);
              await persistInvoices([newInvoice, ...invoices]);
              showToast("Invoice created");
            }}
            onEdit={async (invoiceId, updatedInvoice, cogsJournalEntry, oldQtyChanges, newQtyChanges) => {
              const oldInvoice = invoices.find((i) => i.id === invoiceId);
              let nextProducts = products;
              if (oldQtyChanges.length) {
                nextProducts = nextProducts.map((p) => {
                  const chg = oldQtyChanges.find((c) => c.productId === p.id);
                  if (!chg) return p;
                  return { ...p, qty: (Number(p.qty) || 0) + chg.qty };
                });
              }
              if (newQtyChanges.length) {
                nextProducts = nextProducts.map((p) => {
                  const chg = newQtyChanges.find((c) => c.productId === p.id);
                  if (!chg) return p;
                  return { ...p, qty: (Number(p.qty) || 0) - chg.qty };
                });
              }
              if (oldQtyChanges.length || newQtyChanges.length) await persistProducts(nextProducts);

              let nextJournal = oldInvoice ? journal.filter((j) => j.id !== oldInvoice.journalId) : journal;
              let newJournalId = null;
              if (cogsJournalEntry) {
                const je = { ...cogsJournalEntry, id: uid("je"), createdBy: currentUser.name };
                nextJournal = [je, ...nextJournal];
                newJournalId = je.id;
              }
              await persistJournal(nextJournal);

              const paidTotal = (oldInvoice?.payments || []).reduce((s, p) => s + (Number(p.amount) || 0), 0);
              const status = paidTotal <= 0 ? "unpaid" : paidTotal >= updatedInvoice.total ? "paid" : "partial";

              await persistInvoices(invoices.map((i) => (i.id === invoiceId ? { ...i, ...updatedInvoice, journalId: newJournalId, status } : i)));
              showToast("Invoice updated");
            }}
            onRecordPayment={async (invoice, amount, accountId, method, date) => {
              const je = {
                id: uid("je"), date: date || todayStr(), memo: `Payment received — Invoice ${invoice.number} (${method})`,
                lines: [
                  { accountId: accountId, debit: amount, credit: 0 },
                  { accountId: incomeAccount ? incomeAccount.id : "", debit: 0, credit: amount },
                ],
                createdBy: currentUser.name, source: "invoice-payment", refId: invoice.id,
              };
              await persistJournal([je, ...journal]);
              const payments = [...(invoice.payments || []), { id: uid("pay"), date: date || todayStr(), amount, accountId, method }];
              const paidTotal = payments.reduce((s, p) => s + (Number(p.amount) || 0), 0);
              const status = paidTotal <= 0 ? "unpaid" : paidTotal >= invoice.total ? "paid" : "partial";
              await persistInvoices(invoices.map((i) => (i.id === invoice.id ? { ...i, payments, status } : i)));
              showToast("Payment recorded");
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
              await persistJournal(journal.filter((j) => j.id !== invoice.journalId && j.refId !== invoice.id));
              showToast("Invoice deleted");
            }}
            onBulkImport={async ({ newInvoices, newJournalEntries }) => {
              await persistJournal([...newJournalEntries, ...journal]);
              await persistInvoices([...newInvoices, ...invoices]);
              showToast(`${newInvoices.length} invoices imported`);
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
            onEdit={async (expenseId, updatedExpense, journalEntry) => {
              const old = expenses.find((e) => e.id === expenseId);
              const je = { ...journalEntry, id: uid("je"), createdBy: currentUser.name };
              const nextJournal = old ? journal.filter((j) => j.id !== old.journalId) : journal;
              await persistJournal([je, ...nextJournal]);
              await persistExpenses(expenses.map((e) => (e.id === expenseId ? { ...e, ...updatedExpense, journalId: je.id } : e)));
              showToast("Expense updated");
            }}
            onDelete={async (expense) => {
              await persistExpenses(expenses.filter((e) => e.id !== expense.id));
              await persistJournal(journal.filter((j) => j.id !== expense.journalId));
              showToast("Expense deleted");
            }}
            onBulkImport={async ({ newExpenses, newJournalEntries, newAccounts }) => {
              let baseAccounts = accounts;
              if (newAccounts.length) {
                baseAccounts = [...accounts, ...newAccounts];
                await persistAccounts(baseAccounts);
              }
              await persistJournal([...newJournalEntries, ...journal]);
              await persistExpenses([...newExpenses, ...expenses]);
              showToast(`${newExpenses.length} expenses imported`);
            }}
          />
        )}

        {tab === "income" && (
          <IncomeTab
            accounts={accounts} incomeEntries={incomeEntries} currentUser={currentUser}
            onAddCategory={async (acc) => { await persistAccounts([...accounts, { ...acc, id: uid("acc") }]); return acc; }}
            onAdd={async (entry, journalEntry) => {
              const je = { ...journalEntry, id: uid("je"), createdBy: currentUser.name };
              await persistJournal([je, ...journal]);
              await persistIncomeEntries([{ ...entry, id: uid("inc"), journalId: je.id, createdBy: currentUser.name }, ...incomeEntries]);
              showToast("Income recorded");
            }}
            onEdit={async (entryId, updatedEntry, journalEntry) => {
              const old = incomeEntries.find((e) => e.id === entryId);
              const je = { ...journalEntry, id: uid("je"), createdBy: currentUser.name };
              const nextJournal = old ? journal.filter((j) => j.id !== old.journalId) : journal;
              await persistJournal([je, ...nextJournal]);
              await persistIncomeEntries(incomeEntries.map((e) => (e.id === entryId ? { ...e, ...updatedEntry, journalId: je.id } : e)));
              showToast("Income updated");
            }}
            onDelete={async (entry) => {
              await persistIncomeEntries(incomeEntries.filter((e) => e.id !== entry.id));
              await persistJournal(journal.filter((j) => j.id !== entry.journalId));
              showToast("Income entry deleted");
            }}
          />
        )}

        {tab === "accounts" && (
          <ChartOfAccounts
            accounts={accounts} balances={balances} currentUser={currentUser}
            onAdd={async (acc) => { await persistAccounts([...accounts, { ...acc, id: uid("acc") }]); showToast("Account added"); }}
            onAddMany={async (accs) => { await persistAccounts([...accounts, ...accs.map((a) => ({ ...a, id: uid("acc") }))]); showToast(`${accs.length} accounts added`); }}
            onEdit={async (id, updates) => { await persistAccounts(accounts.map((a) => (a.id === id ? { ...a, ...updates } : a))); showToast("Account updated"); }}
            onDelete={async (id) => { await persistAccounts(accounts.filter((a) => a.id !== id)); showToast("Account deleted"); }}
            onAdjustBalance={async (account, targetBalance, date, note) => {
              const equityAccount = accounts.find((a) => a.name === "Owner's Equity");
              if (!equityAccount) { showToast("Owner's Equity account not found"); return; }
              const currentBalance = balances[account.id] || 0;
              const delta = targetBalance - currentBalance;
              if (delta === 0) { showToast("Balance already matches"); return; }
              const side = normalSideFor(account.type);
              const lines = side === "debit"
                ? (delta > 0
                  ? [{ accountId: account.id, debit: delta, credit: 0 }, { accountId: equityAccount.id, debit: 0, credit: delta }]
                  : [{ accountId: equityAccount.id, debit: -delta, credit: 0 }, { accountId: account.id, debit: 0, credit: -delta }])
                : (delta > 0
                  ? [{ accountId: equityAccount.id, debit: delta, credit: 0 }, { accountId: account.id, debit: 0, credit: delta }]
                  : [{ accountId: account.id, debit: -delta, credit: 0 }, { accountId: equityAccount.id, debit: 0, credit: -delta }]);
              const je = {
                id: uid("je"), date: date || todayStr(), memo: note || `Balance adjustment — ${account.name}`,
                lines, createdBy: currentUser.name, source: "balance-adjustment",
              };
              await persistJournal([je, ...journal]);
              showToast("Balance updated");
            }}
            onTransfer={async (fromId, toId, amount, date, note) => {
              const je = {
                id: uid("je"), date: date || todayStr(), memo: note || "Account transfer",
                lines: [{ accountId: toId, debit: amount, credit: 0 }, { accountId: fromId, debit: 0, credit: amount }],
                createdBy: currentUser.name, source: "transfer",
              };
              await persistJournal([je, ...journal]);
              showToast("Transfer recorded");
            }}
          />
        )}

        {tab === "journal" && (
          <Journal
            accounts={accounts} journal={journal} currentUser={currentUser}
            onAdd={async (entry) => { await persistJournal([{ ...entry, id: uid("je"), createdBy: currentUser.name }, ...journal]); showToast("Journal entry added"); }}
            onDelete={async (id) => { await persistJournal(journal.filter((j) => j.id !== id)); showToast("Entry deleted"); }}
          />
        )}

        {tab === "reports" && <Reports accounts={accounts} balances={balances} invoices={invoices} expenses={expenses} journal={journal} products={products} />}

        {tab === "users" && currentUser.role === "admin" && (
          <UsersPanel
            users={users}
            onAdd={async (u) => { await persistUsers([...users, { ...u, id: uid("usr") }]); showToast("User added"); }}
            onEdit={async (id, updates) => { await persistUsers(users.map((u) => (u.id === id ? { ...u, ...updates } : u))); showToast("User updated"); }}
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
    <div style={{
      minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: FONT.body,
      background: `${PALETTE.bg}`,
      backgroundImage: "radial-gradient(1300px circle at 18% -8%, #7d1c29 0%, transparent 55%), radial-gradient(1100px circle at 100% 10%, #601420 0%, transparent 50%), radial-gradient(1000px circle at 10% 100%, #55111c 0%, transparent 55%)",
    }}>
      <GlobalStyles />
      <form onSubmit={submit} className="pin-card" style={{
        background: PALETTE.card, border: `1px solid ${PALETTE.line}`, width: 380, maxWidth: "90vw",
        padding: "38px 34px", borderRadius: 24, boxShadow: "0 30px 80px rgba(0,0,0,0.45)",
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 6 }}>
          <div style={{ background: "#fff", borderRadius: 14, padding: "8px 12px", display: "inline-flex" }}>
            <img src="/logo.png" alt="Two Threads" style={{ height: 40, width: "auto", display: "block" }} />
          </div>
        </div>
        <p style={{ color: PALETTE.inkSoft, fontSize: 13, marginTop: 12, marginBottom: 22 }}>
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
  const iconMap = { dashboard: LayoutDashboard, invoices: FileText, bills: ClipboardList, expenses: Wallet, income: TrendingUp, inventory: Package, accounts: ScrollText, journal: Receipt, reports: PieChart };
  const allowed = currentUser.role === "admin" || !Array.isArray(currentUser.permissions)
    ? ALL_TABS.map((t) => t.key)
    : currentUser.permissions;
  const items = ALL_TABS.filter((t) => allowed.includes(t.key)).map((t) => ({ ...t, icon: iconMap[t.key] }));
  if (currentUser.role === "admin") items.push({ key: "users", label: "Users", icon: User });

  return (
    <aside className={`app-sidebar ${mobileOpen ? "open" : ""}`} style={{ width: 236, background: PALETTE.cardSolid, borderRight: `1px solid ${PALETTE.line}`, display: "flex", flexDirection: "column", padding: "26px 14px", flexShrink: 0 }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "0 8px", marginBottom: 26 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ background: "#fff", borderRadius: 10, padding: "5px 8px", display: "inline-flex" }}>
            <img src="/logo.png" alt="Two Threads" style={{ height: 24, width: "auto", display: "block" }} />
          </div>
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
  const cash = accounts.filter((a) => a.type === "asset" && a.name !== "Accounts Receivable" && a.name !== "Inventory").reduce((s, a) => s + (balances[a.id] || 0), 0);
  const receivable = accounts.filter((a) => a.name === "Accounts Receivable").reduce((s, a) => s + (balances[a.id] || 0), 0);
  const income = accounts.filter((a) => a.type === "income").reduce((s, a) => s + (balances[a.id] || 0), 0);
  const expenseTotal = accounts.filter((a) => a.type === "expense").reduce((s, a) => s + (balances[a.id] || 0), 0);
  const inventoryValue = products.reduce((s, p) => s + (Number(p.qty) || 0) * (Number(p.costPrice) || 0), 0);
  const recent = journal.slice(0, 6);
  const lowStock = products.filter((p) => (Number(p.qty) || 0) <= (Number(p.reorderLevel) || 3));

  const cards = [
    { label: "Cash, Bank & bKash", value: cash, color: PALETTE.credit },
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

function Inventory({ products, currentUser, onAdd, onEdit, onDelete, onImport }) {
  const [open, setOpen] = useState(false);
  const [editingProduct, setEditingProduct] = useState(null);
  const [form, setForm] = useState({ name: "", sku: "", category: "", unit: "pcs", qty: "", costPrice: "", salePrice: "", reorderLevel: "3", photo: "" });
  const [importing, setImporting] = useState(false);
  const [importMsg, setImportMsg] = useState("");
  const [search, setSearch] = useState("");
  const [photoBusy, setPhotoBusy] = useState(false);
  const isAdmin = currentUser.role === "admin";

  const filteredProducts = products.filter((p) => {
    const q = search.trim().toLowerCase();
    if (!q) return true;
    return p.name.toLowerCase().includes(q) || (p.sku || "").toLowerCase().includes(q);
  });

  const handlePhoto = (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    setPhotoBusy(true);
    const reader = new FileReader();
    reader.onload = (evt) => {
      const img = new Image();
      img.onload = () => {
        const maxDim = 480;
        let { width, height } = img;
        if (width > height && width > maxDim) { height = Math.round((height * maxDim) / width); width = maxDim; }
        else if (height > maxDim) { width = Math.round((width * maxDim) / height); height = maxDim; }
        const canvas = document.createElement("canvas");
        canvas.width = width; canvas.height = height;
        canvas.getContext("2d").drawImage(img, 0, 0, width, height);
        const dataUrl = canvas.toDataURL("image/jpeg", 0.72);
        setForm((f) => ({ ...f, photo: dataUrl }));
        setPhotoBusy(false);
      };
      img.src = evt.target.result;
    };
    reader.readAsDataURL(file);
  };

  const startEditProduct = (p) => {
    setEditingProduct(p);
    setForm({
      name: p.name || "", sku: p.sku || "", category: p.category || "", unit: p.unit || "pcs", qty: String(p.qty ?? ""),
      costPrice: String(p.costPrice ?? ""), salePrice: String(p.salePrice ?? ""), reorderLevel: String(p.reorderLevel ?? "3"),
      photo: p.photo || "",
    });
    setOpen(true);
  };

  const cancelProductForm = () => {
    setOpen(false);
    setEditingProduct(null);
    setForm({ name: "", sku: "", category: "", unit: "pcs", qty: "", costPrice: "", salePrice: "", reorderLevel: "3", photo: "" });
  };

  const submit = (e) => {
    e.preventDefault();
    if (!form.name.trim()) return;
    const payload = {
      name: form.name.trim(), sku: form.sku.trim(), category: form.category.trim() || "Uncategorized", unit: form.unit.trim() || "pcs",
      qty: Number(form.qty) || 0, costPrice: Number(form.costPrice) || 0, salePrice: Number(form.salePrice) || 0,
      reorderLevel: Number(form.reorderLevel) || 3, photo: form.photo || "",
    };
    if (editingProduct) {
      onEdit(editingProduct.id, payload);
    } else {
      onAdd(payload);
    }
    cancelProductForm();
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
          const category = String(pick(row, ["Product Type", "Category"]) || "Uncategorized").trim();
          const qty = Number(pick(row, ["Quantity", "Qty", "Remaing Stock", "Remaining Stock"])) || 0;
          const costPrice = Number(pick(row, ["Purchase Price", "Cost Price"])) || 0;
          const salePrice = Number(pick(row, ["Selling Price", "Sale Price"])) || 0;
          parsed.push({ name, sku, category, unit: "pcs", qty, costPrice, salePrice, reorderLevel: 3 });
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
              display: "flex", alignItems: "center", gap: 6, background: "rgba(255,255,255,0.08)", color: PALETTE.ink,
              padding: "10px 16px", borderRadius: 999, fontSize: 13.5, fontWeight: 600, border: `1px solid ${PALETTE.line}`, cursor: "pointer",
            }}>
              {importing ? <Loader2 size={15} style={{ animation: "spin 1s linear infinite" }} /> : <FileText size={15} />}
              {importing ? "Importing…" : "Import from Excel"}
              <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} style={{ display: "none" }} />
            </label>
            <PrimaryButton onClick={() => (open ? cancelProductForm() : setOpen(true))}>{open ? <X size={15} /> : <Plus size={15} />} {open ? "Cancel" : "New Product"}</PrimaryButton>
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
          <div style={{ fontFamily: FONT.display, fontSize: 15, fontWeight: 600, marginBottom: 12 }}>{editingProduct ? `Editing ${editingProduct.name}` : "New Product"}</div>
          <form onSubmit={submit} className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 110px", gap: 12 }}>
            <div style={{ gridColumn: "1 / -1", display: "flex", alignItems: "center", gap: 14 }}>
              <div style={{ width: 64, height: 64, borderRadius: 14, overflow: "hidden", background: PALETTE.chip, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                {form.photo ? <img src={form.photo} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : <Package size={22} color={PALETTE.inkSoft} />}
              </div>
              <div>
                <label className="pin-btn" style={{
                  display: "inline-flex", alignItems: "center", gap: 6, background: "rgba(255,255,255,0.08)", color: PALETTE.ink,
                  padding: "8px 14px", borderRadius: 999, fontSize: 12.5, fontWeight: 600, border: `1px solid ${PALETTE.line}`, cursor: "pointer",
                }}>
                  {photoBusy ? "Processing…" : form.photo ? "Change Photo" : "Add Photo (optional)"}
                  <input type="file" accept="image/*" onChange={handlePhoto} style={{ display: "none" }} />
                </label>
                {form.photo && <GhostButton onClick={() => setForm({ ...form, photo: "" })} style={{ marginLeft: 8 }}>Remove</GhostButton>}
              </div>
            </div>
            <div><label style={labelStyle}>Product Name</label><input style={inputStyle} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder="e.g. Blue Cotton Dress" /></div>
            <div><label style={labelStyle}>SKU (optional)</label><input style={inputStyle} value={form.sku} onChange={(e) => setForm({ ...form, sku: e.target.value })} placeholder="SKU-001" /></div>
            <div><label style={labelStyle}>Category</label><input style={inputStyle} value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })} placeholder="e.g. Dress, Saree" /></div>
            <div><label style={labelStyle}>Unit</label><input style={inputStyle} value={form.unit} onChange={(e) => setForm({ ...form, unit: e.target.value })} placeholder="pcs" /></div>
            <div><label style={labelStyle}>Opening Qty</label><input style={inputStyle} type="number" value={form.qty} onChange={(e) => setForm({ ...form, qty: e.target.value })} placeholder="0" /></div>
            <div><label style={labelStyle}>Cost Price (BDT)</label><input style={inputStyle} type="number" value={form.costPrice} onChange={(e) => setForm({ ...form, costPrice: e.target.value })} placeholder="0" /></div>
            <div><label style={labelStyle}>Sale Price (BDT)</label><input style={inputStyle} type="number" value={form.salePrice} onChange={(e) => setForm({ ...form, salePrice: e.target.value })} placeholder="0" /></div>
            <div style={{ gridColumn: "1 / -1", display: "flex", justifyContent: "space-between", alignItems: "end" }}>
              <div style={{ maxWidth: 160 }}><label style={labelStyle}>Reorder Alert Below</label><input style={inputStyle} type="number" value={form.reorderLevel} onChange={(e) => setForm({ ...form, reorderLevel: e.target.value })} /></div>
              <PrimaryButton type="submit">{editingProduct ? "Save Changes" : "Add Product"}</PrimaryButton>
            </div>
          </form>
        </Card>
      )}

      <Card style={{ marginBottom: 16, padding: 12 }}>
        <div style={{ position: "relative" }}>
          <Search size={16} color={PALETTE.inkSoft} style={{ position: "absolute", left: 14, top: "50%", transform: "translateY(-50%)" }} />
          <input
            style={{ ...inputStyle, paddingLeft: 38 }}
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Search products by name or SKU…"
          />
        </div>
      </Card>

      {filteredProducts.length === 0 ? (
        <Card><EmptyState text={products.length === 0 ? "No products yet — add your first product to start tracking stock." : "No products match your search."} /></Card>
      ) : (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))", gap: 16 }}>
          {filteredProducts.map((p) => {
            const low = (Number(p.qty) || 0) <= (Number(p.reorderLevel) || 3);
            const out = (Number(p.qty) || 0) <= 0;
            return (
              <Card key={p.id} style={{ padding: 0, overflow: "hidden" }}>
                {p.photo ? (
                  <div style={{ height: 140, overflow: "hidden" }}>
                    <img src={p.photo} alt={p.name} style={{ width: "100%", height: "100%", objectFit: "cover", display: "block" }} />
                  </div>
                ) : (
                  <div style={{ height: 84, background: `linear-gradient(135deg, ${colorFor(p.name)}, ${colorFor(p.name + "x")})`, display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontSize: 30, fontWeight: 700, fontFamily: FONT.display }}>
                    {p.name.slice(0, 1).toUpperCase()}
                  </div>
                )}
                <div style={{ padding: 16 }}>
                  <div style={{ fontWeight: 600, fontSize: 14.5, marginBottom: 2 }}>{p.name}</div>
                  <div style={{ fontSize: 11.5, color: PALETTE.inkSoft, marginBottom: 10 }}>{p.sku || "No SKU"}{p.category ? ` · ${p.category}` : ""}</div>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                    <Badge tone={out ? "bad" : low ? "neutral" : "good"}>{p.qty} {p.unit} in stock</Badge>
                    {isAdmin && (
                      <div style={{ display: "flex", gap: 2 }}>
                        <button className="pin-btn" onClick={() => startEditProduct(p)} style={{ background: "transparent", color: PALETTE.inkSoft }}><Edit2 size={14} /></button>
                        <button className="pin-btn" onClick={() => onDelete(p.id)} style={{ background: "transparent", color: PALETTE.debit }}><Trash2 size={14} /></button>
                      </div>
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
  const cashAccounts = accounts.filter((a) => a.type === "asset" && a.name !== "Accounts Receivable" && a.name !== "Inventory");

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

function Invoices({ accounts, products, invoices, currentUser, onAdd, onEdit, onRecordPayment, onDelete, onBulkImport }) {
  const [open, setOpen] = useState(false);
  const [editingInvoice, setEditingInvoice] = useState(null);
  const [customer, setCustomer] = useState("");
  const [phone, setPhone] = useState("");
  const [address, setAddress] = useState("");
  const [salesBy, setSalesBy] = useState("");
  const [media, setMedia] = useState("Facebook");
  const [date, setDate] = useState(todayStr());
  const [dueDate, setDueDate] = useState(todayStr());
  const [items, setItems] = useState([{ productId: "", desc: "", qty: 1, rate: "" }]);
  const [discount, setDiscount] = useState("");
  const [deposit, setDeposit] = useState("");
  const [depositAccount, setDepositAccount] = useState("");
  const [depositMethod, setDepositMethod] = useState("Cash");
  const [importing, setImporting] = useState(false);
  const [importMsg, setImportMsg] = useState("");
  const [viewingInvoice, setViewingInvoice] = useState(null);
  const [payingInvoice, setPayingInvoice] = useState(null);

  const incomeAccount = accounts.find((a) => a.type === "income");
  const inventoryAccount = accounts.find((a) => a.name === "Inventory");
  const cogsAccount = accounts.find((a) => a.name === "Cost of Goods Sold");
  const cashAccounts = accounts.filter((a) => a.type === "asset" && a.name !== "Accounts Receivable" && a.name !== "Inventory");

  const subtotal = items.reduce((s, it) => s + (Number(it.qty) || 0) * (Number(it.rate) || 0), 0);
  const total = Math.max(subtotal - (Number(discount) || 0), 0);

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

  const startEditInvoice = (inv) => {
    setEditingInvoice(inv);
    setCustomer(inv.customer || ""); setPhone(inv.phone || ""); setAddress(inv.address || "");
    setSalesBy(inv.salesBy || ""); setMedia(inv.media || "Facebook");
    setDate(inv.date || todayStr()); setDueDate(inv.dueDate || todayStr());
    setItems(inv.items && inv.items.length ? inv.items.map((it) => ({ ...it })) : [{ productId: "", desc: "", qty: 1, rate: "" }]);
    setDiscount(inv.discount ? String(inv.discount) : "");
    setDeposit(""); setDepositAccount("");
    setOpen(true);
    setViewingInvoice(null);
  };

  const cancelForm = () => {
    setOpen(false);
    setEditingInvoice(null);
    setCustomer(""); setPhone(""); setAddress(""); setSalesBy(""); setMedia("Facebook");
    setItems([{ productId: "", desc: "", qty: 1, rate: "" }]);
    setDiscount(""); setDeposit(""); setDepositAccount("");
  };

  const buildCogs = (cleanItems) => {
    const lines = [];
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
    return { lines, qtyChanges, totalCogs };
  };

  const submit = (e) => {
    e.preventDefault();
    if (!customer.trim() || total <= 0) return;
    const cleanItems = items.filter((it) => it.desc.trim() && (Number(it.qty) || 0) > 0);
    const discountAmt = Number(discount) || 0;

    if (editingInvoice) {
      const { lines, qtyChanges: newQtyChanges } = buildCogs(cleanItems);
      const oldQtyChanges = (editingInvoice.items || []).filter((it) => it.productId).map((it) => ({ productId: it.productId, qty: Number(it.qty) || 0 }));
      const cogsJournalEntry = lines.length > 0 ? { date, memo: `COGS — Invoice ${editingInvoice.number}`, lines, source: "invoice-cogs" } : null;
      const updatedInvoice = {
        customer: customer.trim(), phone: phone.trim(), address: address.trim(), salesBy: salesBy.trim(), media,
        date, dueDate, items: cleanItems, subtotal, discount: discountAmt, total,
      };
      onEdit(editingInvoice.id, updatedInvoice, cogsJournalEntry, oldQtyChanges, newQtyChanges);
      cancelForm();
      return;
    }

    const number = `INV-${String(invoices.length + 1).padStart(4, "0")}`;
    const invoice = {
      number, customer: customer.trim(), phone: phone.trim(), address: address.trim(), salesBy: salesBy.trim(), media,
      date, dueDate, items: cleanItems, subtotal, discount: discountAmt, total, payments: [], status: "unpaid",
    };
    const { lines, qtyChanges } = buildCogs(cleanItems);
    const cogsJournalEntry = lines.length > 0 ? { date, memo: `COGS — Invoice ${number}`, lines, source: "invoice-cogs" } : null;

    const depositAmt = Number(deposit) || 0;
    onAdd(invoice, cogsJournalEntry, qtyChanges, depositAmt > 0 ? { amount: depositAmt, accountId: depositAccount, method: depositMethod, date } : null);
    cancelForm();
  };

  const pick = (row, keys) => {
    for (const k of keys) {
      for (const rk of Object.keys(row)) {
        if (rk.trim().toLowerCase() === k.toLowerCase()) return row[rk];
      }
    }
    return undefined;
  };

  const toDateStr = (v) => {
    if (!v) return todayStr();
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    const d = new Date(v);
    if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
    return todayStr();
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
        const wb = XLSX.read(data, { type: "array", cellDates: true });
        const sheetName = wb.SheetNames.find((n) => n.trim().toLowerCase().includes("sales")) || wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

        const cashFallback = accounts.find((a) => a.name === "Cash in Hand") || cashAccounts[0];
        const newInvoices = [];
        const newJournalEntries = [];
        let n = invoices.length;

        rows.forEach((row) => {
          const customerName = String(pick(row, ["Customer Name"]) || "").trim();
          const qty = Number(pick(row, ["Qty", "Quantity"])) || 0;
          if (!customerName || qty <= 0) return;
          const rate = Number(pick(row, ["Unit Price"])) || 0;
          const rowTotal = Number(pick(row, ["Price"])) || qty * rate;
          if (rowTotal <= 0) return;

          n += 1;
          const number = `INV-${String(n).padStart(4, "0")}`;
          const desc = String(pick(row, ["Product", "Dress name"]) || "Item").trim();
          const rowDate = toDateStr(pick(row, ["Date"]));

          const payments = [];
          const advanceAmt = Number(pick(row, ["Advance"])) || 0;
          if (advanceAmt > 0) payments.push({ id: uid("pay"), date: rowDate, amount: advanceAmt, accountId: cashFallback ? cashFallback.id : "", method: "Imported" });
          const paymentAmt = Number(pick(row, ["Payment Amount"])) || 0;
          if (paymentAmt > 0) {
            const payDate = toDateStr(pick(row, ["Payment Date"])) || rowDate;
            payments.push({ id: uid("pay"), date: payDate, amount: paymentAmt, accountId: cashFallback ? cashFallback.id : "", method: "Imported" });
          }
          const paidTotal = payments.reduce((s, p) => s + p.amount, 0);
          const status = paidTotal <= 0 ? "unpaid" : paidTotal >= rowTotal ? "paid" : "partial";

          const invoice = {
            id: uid("inv"), number, customer: customerName,
            phone: String(pick(row, ["Contact number", "Contact Number"]) || "").trim(),
            address: String(pick(row, ["Address"]) || "").trim(),
            salesBy: String(pick(row, ["Sales By"]) || "").trim(),
            date: rowDate, dueDate: rowDate,
            items: [{ productId: null, desc, qty, rate }],
            subtotal: rowTotal, discount: 0, total: rowTotal, payments, status,
            createdBy: currentUser.name, historical: true,
          };

          payments.forEach((p) => {
            if (!p.accountId || !incomeAccount) return;
            newJournalEntries.push({
              id: uid("je"), date: p.date, memo: `Payment received — Invoice ${number} (imported)`,
              lines: [{ accountId: p.accountId, debit: p.amount, credit: 0 }, { accountId: incomeAccount.id, debit: 0, credit: p.amount }],
              createdBy: currentUser.name, source: "invoice-payment", refId: invoice.id,
            });
          });

          newInvoices.push(invoice);
        });

        if (newInvoices.length === 0) {
          setImportMsg('No sales rows found. Make sure the sheet has "Customer Name", "Qty", "Unit Price"/"Price" columns.');
        } else {
          const ok = window.confirm(
            `${newInvoices.length} past sales found. Only amounts actually received (advance/payment) will be added to income — unpaid balances stay on record but won't affect your books. Inventory stock will NOT be changed. Continue?`
          );
          if (ok) onBulkImport({ newInvoices, newJournalEntries });
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
        <PageHeader title="Invoices" subtitle="Create a new invoice — inventory and journal entries update automatically" />
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          {currentUser.role === "admin" && (
            <label className="pin-btn" style={{
              display: "flex", alignItems: "center", gap: 6, background: "rgba(255,255,255,0.08)", color: PALETTE.ink,
              padding: "10px 16px", borderRadius: 999, fontSize: 13.5, fontWeight: 600, border: `1px solid ${PALETTE.line}`, cursor: "pointer",
            }}>
              {importing ? <Loader2 size={15} style={{ animation: "spin 1s linear infinite" }} /> : <FileText size={15} />}
              {importing ? "Importing…" : "Import Past Sales"}
              <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} style={{ display: "none" }} />
            </label>
          )}
          <PrimaryButton onClick={() => (open ? cancelForm() : setOpen(true))}>{open ? <X size={15} /> : <Plus size={15} />} {open ? "Cancel" : "New Invoice"}</PrimaryButton>
        </div>
      </div>

      {importMsg && (
        <Card style={{ marginBottom: 16, borderColor: PALETTE.debit }}>
          <div style={{ color: PALETTE.debit, fontSize: 13, display: "flex", gap: 8, alignItems: "center" }}><AlertCircle size={15} /> {importMsg}</div>
        </Card>
      )}

      {open && (
        <Card style={{ marginBottom: 20 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 10, marginBottom: 16 }}>
            <div>
              <div style={{ fontFamily: FONT.display, fontSize: 17, fontWeight: 600 }}>{editingInvoice ? `Editing ${editingInvoice.number}` : "New Invoice"}</div>
              <div style={{ fontSize: 12, color: PALETTE.inkSoft, marginTop: 2 }}>Customer & Invoice Info — fill details, add items below</div>
            </div>
            {editingInvoice && (
              <div style={{ textAlign: "right" }}>
                <div style={{ fontSize: 10.5, color: PALETTE.inkSoft, textTransform: "uppercase", letterSpacing: 0.4 }}>Remaining Due</div>
                <div style={{ fontSize: 22, fontWeight: 700, fontFamily: FONT.mono, color: PALETTE.accent }}>
                  {fmtMoney(Math.max(editingInvoice.total - (editingInvoice.payments || []).reduce((s, p) => s + (Number(p.amount) || 0), 0), 0))}
                </div>
                {editingInvoice.status !== "paid" && (
                  <GhostButton onClick={() => { setPayingInvoice(editingInvoice); cancelForm(); }} style={{ marginTop: 2 }}>Receive Payment →</GhostButton>
                )}
              </div>
            )}
          </div>

          <form onSubmit={submit}>
            <Card style={{ padding: 16, marginBottom: 16, background: "rgba(255,255,255,0.03)" }}>
              <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                <div><label style={labelStyle}>Customer Name</label><input style={inputStyle} value={customer} onChange={(e) => setCustomer(e.target.value)} placeholder="Customer name" /></div>
                <div><label style={labelStyle}>Phone (optional)</label><input style={inputStyle} value={phone} onChange={(e) => setPhone(e.target.value)} placeholder="01XXXXXXXXX" /></div>
              </div>
              <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: 10 }}>
                <div><label style={labelStyle}>Billing Address (optional)</label><input style={inputStyle} value={address} onChange={(e) => setAddress(e.target.value)} placeholder="Delivery address" /></div>
                <div><label style={labelStyle}>Sales By (optional)</label><input style={inputStyle} value={salesBy} onChange={(e) => setSalesBy(e.target.value)} placeholder="Staff name" /></div>
              </div>
              <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginTop: 10 }}>
                <div><label style={labelStyle}>Media</label><select style={inputStyle} value={media} onChange={(e) => setMedia(e.target.value)}>{MEDIA_OPTIONS.map((m) => <option key={m} value={m}>{m}</option>)}</select></div>
                <div><label style={labelStyle}>Invoice Date</label><input type="date" style={inputStyle} value={date} onChange={(e) => setDate(e.target.value)} /></div>
                <div><label style={labelStyle}>Due Date</label><input type="date" style={inputStyle} value={dueDate} onChange={(e) => setDueDate(e.target.value)} /></div>
              </div>
            </Card>

            <div style={{ fontSize: 11, color: PALETTE.inkSoft, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 8 }}>Products / Services</div>
            <div style={{ border: `1px solid ${PALETTE.line}`, borderRadius: 12, overflow: "hidden", marginBottom: 14 }}>
              <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "160px 1fr 70px 100px 30px", gap: 8, padding: "8px 10px", background: PALETTE.chip, fontSize: 10.5, fontWeight: 700, color: PALETTE.inkSoft, textTransform: "uppercase", letterSpacing: 0.3 }}>
                <span>Product</span><span>Description</span><span>Qty</span><span>Rate</span><span></span>
              </div>
              {items.map((it, i) => {
                const p = products.find((pr) => pr.id === it.productId);
                const short = p && Number(it.qty) > Number(p.qty);
                return (
                  <div key={i} style={{ padding: "8px 10px", borderTop: `1px solid ${PALETTE.line}` }}>
                    <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "160px 1fr 70px 100px 30px", gap: 8 }}>
                      <select style={inputStyle} value={it.productId} onChange={(e) => updateItem(i, "productId", e.target.value)}>
                        <option value="">Other</option>
                        {products.map((pr) => <option key={pr.id} value={pr.id}>{pr.name}</option>)}
                      </select>
                      <input style={inputStyle} placeholder="Description" value={it.desc} onChange={(e) => updateItem(i, "desc", e.target.value)} disabled={!!it.productId} />
                      <input style={inputStyle} type="number" placeholder="Qty" value={it.qty} onChange={(e) => updateItem(i, "qty", e.target.value)} />
                      <input style={inputStyle} type="number" placeholder="Rate" value={it.rate} onChange={(e) => updateItem(i, "rate", e.target.value)} />
                      {items.length > 1 && <button type="button" className="pin-btn" onClick={() => setItems(items.filter((_, idx) => idx !== i))} style={{ background: "transparent", color: PALETTE.debit }}><Trash2 size={14} /></button>}
                    </div>
                    {p && <div style={{ fontSize: 11, color: short ? PALETTE.debit : PALETTE.inkSoft, marginTop: 4 }}>
                      {p.qty} {p.unit} in stock{short ? " — not enough stock, this will oversell" : ""}
                    </div>}
                  </div>
                );
              })}
              <div style={{ padding: "8px 10px", borderTop: `1px solid ${PALETTE.line}` }}>
                <GhostButton onClick={() => setItems([...items, { productId: "", desc: "", qty: 1, rate: "" }])}>+ Add line</GhostButton>
              </div>
            </div>

            <div style={{ display: "flex", justifyContent: "flex-end" }}>
              <Card style={{ padding: 16, width: "100%", maxWidth: 340, background: "rgba(255,255,255,0.03)" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                  <span style={{ fontSize: 13, color: PALETTE.inkSoft }}>Subtotal</span>
                  <span style={{ fontFamily: FONT.mono, fontSize: 13 }}>{fmtMoney(subtotal)}</span>
                </div>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8, gap: 8 }}>
                  <label style={{ fontSize: 13, color: PALETTE.inkSoft }}>Discount</label>
                  <input style={{ ...inputStyle, width: 110, padding: "6px 8px", textAlign: "right" }} type="number" value={discount} onChange={(e) => setDiscount(e.target.value)} placeholder="0" />
                </div>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", borderTop: `1px solid ${PALETTE.line}`, paddingTop: 8, marginBottom: 8 }}>
                  <span style={{ fontSize: 15, fontWeight: 700 }}>Total</span>
                  <span style={{ fontFamily: FONT.mono, fontSize: 16, fontWeight: 700 }}>{fmtMoney(total)}</span>
                </div>

                {!editingInvoice && (
                  <>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8, gap: 8 }}>
                      <label style={{ fontSize: 13, color: PALETTE.inkSoft }}>Deposit</label>
                      <input style={{ ...inputStyle, width: 110, padding: "6px 8px", textAlign: "right" }} type="number" value={deposit} onChange={(e) => setDeposit(e.target.value)} placeholder="0" />
                    </div>
                    {Number(deposit) > 0 && (
                      <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
                        <select style={{ ...inputStyle, padding: "6px 8px", fontSize: 12.5 }} value={depositMethod} onChange={(e) => setDepositMethod(e.target.value)}>{PAYMENT_METHODS.map((m) => <option key={m} value={m}>{m}</option>)}</select>
                        <select style={{ ...inputStyle, padding: "6px 8px", fontSize: 12.5 }} value={depositAccount} onChange={(e) => setDepositAccount(e.target.value)}><option value="">Deposit to…</option>{cashAccounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}</select>
                      </div>
                    )}
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", borderTop: `1px solid ${PALETTE.line}`, paddingTop: 8 }}>
                      <span style={{ fontSize: 14, fontWeight: 700, color: PALETTE.accent }}>Remaining Due</span>
                      <span style={{ fontFamily: FONT.mono, fontSize: 15, fontWeight: 700, color: PALETTE.accent }}>{fmtMoney(Math.max(total - (Number(deposit) || 0), 0))}</span>
                    </div>
                  </>
                )}
              </Card>
            </div>

            <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 16 }}>
              <PrimaryButton type="submit">{editingInvoice ? "Save Changes" : "Create Invoice"}</PrimaryButton>
            </div>
          </form>
        </Card>
      )}

      <Card>
        {invoices.length === 0 ? <EmptyState text="No invoices yet" /> : (
          <div style={{ overflowX: "auto" }}><table style={{ minWidth: 720 }}>
            <thead><tr style={{ borderBottom: `1px solid ${PALETTE.line}` }}><Th>No.</Th><Th>Customer</Th><Th>Phone</Th><Th>Date</Th><Th align="right">Total</Th><Th align="right">Due</Th><Th>Status</Th><Th> </Th></tr></thead>
            <tbody>
              {invoices.map((inv) => {
                const paidTotal = (inv.payments || []).reduce((s, p) => s + (Number(p.amount) || 0), 0);
                const balanceDue = inv.total - paidTotal;
                const statusTone = inv.status === "paid" ? "good" : inv.status === "partial" ? "neutral" : "bad";
                const statusLabel = inv.status === "paid" ? "Paid" : inv.status === "partial" ? "Partial" : "Unpaid";
                return (
                  <tr key={inv.id} className="row-hover" style={{ borderBottom: `1px solid ${PALETTE.line}` }}>
                    <Td mono>
                      <button className="pin-btn" onClick={() => setViewingInvoice(inv)} style={{ background: "transparent", color: PALETTE.accent, fontFamily: FONT.mono, fontSize: 13.5, padding: 0, textDecoration: "underline" }}>
                        {inv.number}
                      </button>
                    </Td>
                    <Td>{inv.customer}{inv.salesBy ? <div style={{ fontSize: 11, color: PALETTE.inkSoft }}>by {inv.salesBy}</div> : null}</Td>
                    <Td style={{ fontSize: 12.5, color: PALETTE.inkSoft }}>{inv.phone || "—"}</Td>
                    <Td>{inv.date}</Td>
                    <Td align="right" mono>{fmtMoney(inv.total)}</Td>
                    <Td align="right" mono>{balanceDue > 0 ? fmtMoney(balanceDue) : "—"}</Td>
                    <Td><Badge tone={statusTone}>{statusLabel}</Badge></Td>
                    <Td>
                      <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                        {inv.status !== "paid" && (
                          <GhostButton onClick={() => setPayingInvoice(inv)}>Receive Payment</GhostButton>
                        )}
                        {currentUser.role === "admin" && (
                          <>
                            <button className="pin-btn" onClick={() => startEditInvoice(inv)} style={{ background: "transparent", color: PALETTE.inkSoft }}><Edit2 size={14} /></button>
                            <button className="pin-btn" onClick={() => onDelete(inv)} style={{ background: "transparent", color: PALETTE.debit }}><Trash2 size={14} /></button>
                          </>
                        )}
                      </div>
                    </Td>
                  </tr>
                );
              })}
            </tbody>
          </table></div>
        )}
      </Card>

      {viewingInvoice && (
        <InvoiceDetailModal
          invoice={viewingInvoice}
          onClose={() => setViewingInvoice(null)}
          onEdit={currentUser.role === "admin" ? () => startEditInvoice(viewingInvoice) : null}
          onReceivePayment={viewingInvoice.status !== "paid" ? () => { setPayingInvoice(viewingInvoice); setViewingInvoice(null); } : null}
        />
      )}
      {payingInvoice && (
        <ReceivePaymentPage
          invoice={payingInvoice}
          cashAccounts={cashAccounts}
          onClose={() => setPayingInvoice(null)}
          onConfirm={(amount, accountId, method, payDate) => {
            onRecordPayment(payingInvoice, amount, accountId, method, payDate);
            setPayingInvoice(null);
          }}
        />
      )}
    </div>
  );
}

function ReceivePaymentPage({ invoice, cashAccounts, onClose, onConfirm }) {
  const paidTotal = (invoice.payments || []).reduce((s, p) => s + (Number(p.amount) || 0), 0);
  const balanceDue = Math.max(invoice.total - paidTotal, 0);
  const [amount, setAmount] = useState(String(balanceDue));
  const [accountId, setAccountId] = useState("");
  const [method, setMethod] = useState("Cash");
  const [payDate, setPayDate] = useState(todayStr());

  const submit = (e) => {
    e.preventDefault();
    const amt = Number(amount);
    if (!(amt > 0) || !accountId) return;
    onConfirm(amt, accountId, method, payDate);
  };

  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(20,20,20,0.5)", zIndex: 320, display: "flex", alignItems: "flex-start", justifyContent: "center", overflowY: "auto", padding: "40px 16px" }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: PALETTE.cardSolid, border: `1px solid ${PALETTE.line}`, borderRadius: 20, maxWidth: 440, width: "100%", padding: 28, boxShadow: "0 30px 80px rgba(0,0,0,0.5)" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 6 }}>
          <div style={{ fontFamily: FONT.display, fontSize: 19, fontWeight: 600 }}>Receive Payment</div>
          <button className="pin-btn" onClick={onClose} style={{ background: PALETTE.chip, padding: 8, borderRadius: 999 }}><X size={16} /></button>
        </div>
        <p style={{ fontSize: 13, color: PALETTE.inkSoft, marginTop: 0, marginBottom: 20 }}>{invoice.number} — {invoice.customer}</p>

        <Card style={{ marginBottom: 18, padding: 16 }}>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, color: PALETTE.inkSoft, marginBottom: 4 }}><span>Total</span><span style={{ fontFamily: FONT.mono, color: PALETTE.ink }}>{fmtMoney(invoice.total)}</span></div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, color: PALETTE.inkSoft, marginBottom: 4 }}><span>Already Paid</span><span style={{ fontFamily: FONT.mono, color: PALETTE.credit }}>{fmtMoney(paidTotal)}</span></div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 15, fontWeight: 700, marginTop: 6 }}><span>Remaining Due</span><span style={{ fontFamily: FONT.mono, color: PALETTE.accent }}>{fmtMoney(balanceDue)}</span></div>
        </Card>

        <form onSubmit={submit}>
          <label style={labelStyle}>Payment Date</label>
          <input type="date" style={inputStyle} value={payDate} onChange={(e) => setPayDate(e.target.value)} />
          <label style={labelStyle}>Amount Received (BDT)</label>
          <input type="number" style={inputStyle} value={amount} onChange={(e) => setAmount(e.target.value)} placeholder="0" />
          <label style={labelStyle}>Method</label>
          <select style={inputStyle} value={method} onChange={(e) => setMethod(e.target.value)}>{PAYMENT_METHODS.map((m) => <option key={m} value={m}>{m}</option>)}</select>
          <label style={labelStyle}>Received Into (Account)</label>
          <select style={inputStyle} value={accountId} onChange={(e) => setAccountId(e.target.value)}>
            <option value="">Select account</option>
            {cashAccounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
          </select>
          <PrimaryButton type="submit" style={{ width: "100%", justifyContent: "center", marginTop: 20, padding: "12px 0" }} disabled={!(Number(amount) > 0) || !accountId}>
            Confirm Payment
          </PrimaryButton>
        </form>
      </div>
    </div>
  );
}

function InvoiceDetailModal({ invoice, onClose, onEdit, onReceivePayment }) {
  const paidTotal = (invoice.payments || []).reduce((s, p) => s + (Number(p.amount) || 0), 0);
  const balanceDue = invoice.total - paidTotal;
  const statusTone = invoice.status === "paid" ? "good" : invoice.status === "partial" ? "neutral" : "bad";
  const statusLabel = invoice.status === "paid" ? "Paid" : invoice.status === "partial" ? "Partial" : "Unpaid";

  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(20,20,20,0.45)", zIndex: 300, display: "flex", alignItems: "flex-start", justifyContent: "center", overflowY: "auto", padding: "40px 16px" }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: PALETTE.cardSolid, border: `1px solid ${PALETTE.line}`, borderRadius: 20, maxWidth: 560, width: "100%", padding: 28, boxShadow: "0 30px 80px rgba(0,0,0,0.5)" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 18 }}>
          <div>
            <div style={{ fontFamily: FONT.display, fontSize: 20, fontWeight: 600 }}>{invoice.number}</div>
            <div style={{ fontSize: 12.5, color: PALETTE.inkSoft, marginTop: 2 }}>{invoice.date}{invoice.dueDate && invoice.dueDate !== invoice.date ? ` · Due ${invoice.dueDate}` : ""}</div>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            {onEdit && <button className="pin-btn" onClick={onEdit} style={{ background: PALETTE.chip, padding: 8, borderRadius: 999 }}><Edit2 size={16} /></button>}
            <button className="pin-btn" onClick={onClose} style={{ background: PALETTE.chip, padding: 8, borderRadius: 999 }}><X size={16} /></button>
          </div>
        </div>

        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 18, flexWrap: "wrap", gap: 10 }}>
          <div>
            <div style={{ fontSize: 11, color: PALETTE.inkSoft, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.3, marginBottom: 4 }}>Bill To</div>
            <div style={{ fontSize: 14, fontWeight: 600 }}>{invoice.customer}</div>
            {invoice.phone && <div style={{ fontSize: 12.5, color: PALETTE.inkSoft }}>{invoice.phone}</div>}
            {invoice.address && <div style={{ fontSize: 12.5, color: PALETTE.inkSoft }}>{invoice.address}</div>}
            {invoice.salesBy && <div style={{ fontSize: 12, color: PALETTE.inkSoft, marginTop: 4 }}>Sales by {invoice.salesBy}</div>}
            {invoice.media && <div style={{ fontSize: 12, color: PALETTE.inkSoft, marginTop: 2 }}>Media: {invoice.media}</div>}
          </div>
          <Badge tone={statusTone}>{statusLabel}</Badge>
        </div>

        <div style={{ border: `1px solid ${PALETTE.line}`, borderRadius: 12, overflow: "hidden", marginBottom: 16 }}>
          <table style={{ width: "100%" }}>
            <thead><tr style={{ background: PALETTE.chip }}><Th>Item</Th><Th align="right">Qty</Th><Th align="right">Rate</Th><Th align="right">Amount</Th></tr></thead>
            <tbody>
              {(invoice.items || []).map((it, idx) => (
                <tr key={idx} style={{ borderTop: `1px solid ${PALETTE.line}` }}>
                  <Td>{it.desc}</Td>
                  <Td align="right" mono>{it.qty}</Td>
                  <Td align="right" mono>{fmtMoney(it.rate)}</Td>
                  <Td align="right" mono>{fmtMoney((Number(it.qty) || 0) * (Number(it.rate) || 0))}</Td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 6, alignItems: "flex-end", marginBottom: 18 }}>
          {invoice.discount > 0 ? (
            <>
              <div style={{ fontSize: 13, color: PALETTE.inkSoft }}>Subtotal: <span style={{ fontFamily: FONT.mono, color: PALETTE.ink }}>{fmtMoney(invoice.subtotal != null ? invoice.subtotal : invoice.total)}</span></div>
              <div style={{ fontSize: 13, color: PALETTE.inkSoft }}>Discount: <span style={{ fontFamily: FONT.mono, color: PALETTE.debit }}>-{fmtMoney(invoice.discount)}</span></div>
            </>
          ) : null}
          <div style={{ fontSize: 14, color: PALETTE.inkSoft }}>Total: <span style={{ fontFamily: FONT.mono, color: PALETTE.ink }}>{fmtMoney(invoice.total)}</span></div>
          {paidTotal > 0 && <div style={{ fontSize: 13, color: PALETTE.inkSoft }}>Deposit / Paid: <span style={{ fontFamily: FONT.mono, color: PALETTE.credit }}>{fmtMoney(paidTotal)}</span></div>}
          <div style={{ fontSize: 16, fontWeight: 700 }}>Remaining Due: <span style={{ fontFamily: FONT.mono }}>{fmtMoney(balanceDue)}</span></div>
        </div>

        {(invoice.payments || []).length > 0 && (
          <div style={{ marginBottom: 20 }}>
            <div style={{ fontSize: 11, color: PALETTE.inkSoft, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.3, marginBottom: 6 }}>Payment History</div>
            {invoice.payments.map((p) => (
              <div key={p.id} style={{ fontSize: 12.5, color: PALETTE.inkSoft, display: "flex", justifyContent: "space-between", padding: "4px 0" }}>
                <span>{p.date} · {p.method}</span>
                <span style={{ fontFamily: FONT.mono }}>{fmtMoney(p.amount)}</span>
              </div>
            ))}
          </div>
        )}

        {onReceivePayment && (
          <PrimaryButton onClick={onReceivePayment} style={{ width: "100%", justifyContent: "center", padding: "12px 0", marginBottom: 10 }}>
            Receive Payment
          </PrimaryButton>
        )}
        <button
          type="button" className="pin-btn" onClick={() => generateInvoicePDF(invoice)}
          style={{
            width: "100%", justifyContent: "center", display: "flex", alignItems: "center", gap: 6,
            background: "rgba(255,255,255,0.08)", color: PALETTE.ink, border: `1px solid ${PALETTE.line}`,
            padding: "12px 0", borderRadius: 999, fontSize: 13.5, fontWeight: 600,
          }}
        >
          <Download size={16} /> Download PDF
        </button>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Expenses
--------------------------------------------------------- */

function Expenses({ accounts, expenses, currentUser, onAdd, onEdit, onDelete, onBulkImport }) {
  const [open, setOpen] = useState(false);
  const [date, setDate] = useState(todayStr());
  const [vendor, setVendor] = useState("");
  const [accountId, setAccountId] = useState("");
  const [paymentAccountId, setPaymentAccountId] = useState("");
  const [amount, setAmount] = useState("");
  const [note, setNote] = useState("");
  const [editingId, setEditingId] = useState(null);
  const [editForm, setEditForm] = useState({});
  const [importing, setImporting] = useState(false);
  const [importMsg, setImportMsg] = useState("");

  const expenseAccounts = accounts.filter((a) => a.type === "expense" && a.name !== "Cost of Goods Sold");
  const cashAccounts = accounts.filter((a) => a.type === "asset" && a.name !== "Accounts Receivable" && a.name !== "Inventory");

  const submit = (e) => {
    e.preventDefault();
    const amt = Number(amount);
    if (!accountId || !paymentAccountId || !(amt > 0)) return;
    const categoryName = expenseAccounts.find((a) => a.id === accountId)?.name || "Expense";
    const vendorLabel = vendor.trim() || categoryName;
    const expense = { date, vendor: vendorLabel, accountId, paymentAccountId, amount: amt, note: note.trim() };
    const journalEntry = { date, memo: `Expense — ${vendorLabel}`, lines: [{ accountId, debit: amt, credit: 0 }, { accountId: paymentAccountId, debit: 0, credit: amt }], source: "expense" };
    onAdd(expense, journalEntry);
    setVendor(""); setAmount(""); setNote(""); setOpen(false);
  };

  const startEdit = (exp) => {
    setEditingId(exp.id);
    setEditForm({ date: exp.date, vendor: exp.vendor, accountId: exp.accountId, paymentAccountId: exp.paymentAccountId, amount: String(exp.amount), note: exp.note || "" });
  };
  const saveEdit = (expenseId) => {
    const amt = Number(editForm.amount);
    if (!editForm.accountId || !editForm.paymentAccountId || !(amt > 0)) return;
    const categoryName = expenseAccounts.find((a) => a.id === editForm.accountId)?.name || "Expense";
    const vendorLabel = editForm.vendor.trim() || categoryName;
    const updatedExpense = { date: editForm.date, vendor: vendorLabel, accountId: editForm.accountId, paymentAccountId: editForm.paymentAccountId, amount: amt, note: editForm.note.trim() };
    const journalEntry = { date: editForm.date, memo: `Expense — ${vendorLabel}`, lines: [{ accountId: editForm.accountId, debit: amt, credit: 0 }, { accountId: editForm.paymentAccountId, debit: 0, credit: amt }], source: "expense" };
    onEdit(expenseId, updatedExpense, journalEntry);
    setEditingId(null);
  };

  const pick = (row, keys) => {
    for (const k of keys) {
      for (const rk of Object.keys(row)) {
        if (rk.trim().toLowerCase() === k.toLowerCase()) return row[rk];
      }
    }
    return undefined;
  };
  const toDateStr = (v) => {
    if (!v) return todayStr();
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    const d = new Date(v);
    if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
    return todayStr();
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
        const wb = XLSX.read(data, { type: "array", cellDates: true });
        const sheetName = wb.SheetNames.find((n) => n.trim().toLowerCase().includes("expense")) || wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

        let workingAccounts = [...accounts];
        const newAccounts = [];
        const findOrCreateExpenseAccount = (name) => {
          const clean = name.trim();
          let acc = workingAccounts.find((a) => a.type === "expense" && a.name.toLowerCase() === clean.toLowerCase());
          if (acc) return acc;
          const existingCodes = workingAccounts.filter((a) => a.type === "expense").map((a) => Number(a.code) || 0);
          const code = String(Math.max(5000, ...existingCodes, 0) + 10);
          acc = { id: uid("acc"), code, name: clean, type: "expense" };
          workingAccounts.push(acc);
          newAccounts.push(acc);
          return acc;
        };
        const defaultCash = accounts.find((a) => a.name === "Cash in Hand") || cashAccounts[0];
        const findPaymentAccount = (name) => {
          if (!name) return defaultCash;
          const clean = String(name).trim().toLowerCase();
          const found = workingAccounts.find((a) => a.type === "asset" && a.name.toLowerCase() === clean);
          return found || defaultCash;
        };

        const newExpenses = [];
        const newJournalEntries = [];

        rows.forEach((row) => {
          const category = String(pick(row, ["Expense Catagory", "Expense Category", "Category"]) || "").trim();
          const amt = Number(pick(row, ["Expense Amount", "Amount"])) || 0;
          if (!category || amt <= 0) return;
          if (category.toLowerCase().includes("opening balace") || category.toLowerCase().includes("opening balance")) return;

          const vendorName = String(pick(row, ["Expense Details", "Vendor", "Details"]) || category).trim();
          const rowDate = toDateStr(pick(row, ["Date"]));
          const methodName = pick(row, ["Method", "Paid From"]);
          const categoryAccount = findOrCreateExpenseAccount(category);
          const paymentAccount = findPaymentAccount(methodName);
          if (!paymentAccount) return;

          const expense = {
            id: uid("exp"), date: rowDate, vendor: vendorName || category,
            accountId: categoryAccount.id, paymentAccountId: paymentAccount.id,
            amount: amt, note: "Imported", createdBy: currentUser.name,
          };
          const je = {
            id: uid("je"), date: rowDate, memo: `Expense — ${expense.vendor} (imported)`,
            lines: [{ accountId: categoryAccount.id, debit: amt, credit: 0 }, { accountId: paymentAccount.id, debit: 0, credit: amt }],
            createdBy: currentUser.name, source: "expense", refId: expense.id,
          };
          expense.journalId = je.id;
          newExpenses.push(expense);
          newJournalEntries.push(je);
        });

        if (newExpenses.length === 0) {
          setImportMsg('No expense rows found. Make sure the sheet has "Expense Catagory" and "Expense Amount" columns.');
        } else {
          const ok = window.confirm(`${newExpenses.length} expenses found${newAccounts.length ? ` (will also create ${newAccounts.length} new categories)` : ""}. Import them?`);
          if (ok) onBulkImport({ newExpenses, newJournalEntries, newAccounts });
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
        <PageHeader title="Expenses" subtitle="Operating costs like rent, utilities, and salaries" />
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          {currentUser.role === "admin" && (
            <label className="pin-btn" style={{
              display: "flex", alignItems: "center", gap: 6, background: "rgba(255,255,255,0.08)", color: PALETTE.ink,
              padding: "10px 16px", borderRadius: 999, fontSize: 13.5, fontWeight: 600, border: `1px solid ${PALETTE.line}`, cursor: "pointer",
            }}>
              {importing ? <Loader2 size={15} style={{ animation: "spin 1s linear infinite" }} /> : <FileText size={15} />}
              {importing ? "Importing…" : "Import from Excel"}
              <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} style={{ display: "none" }} />
            </label>
          )}
          <PrimaryButton onClick={() => setOpen((v) => !v)}>{open ? <X size={15} /> : <Plus size={15} />} {open ? "Cancel" : "New Expense"}</PrimaryButton>
        </div>
      </div>

      {importMsg && (
        <Card style={{ marginBottom: 16, borderColor: PALETTE.debit }}>
          <div style={{ color: PALETTE.debit, fontSize: 13, display: "flex", gap: 8, alignItems: "center" }}><AlertCircle size={15} /> {importMsg}</div>
        </Card>
      )}

      {open && (
        <Card style={{ marginBottom: 20 }}>
          <form onSubmit={submit}>
            <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "140px 1fr 140px", gap: 10 }}>
              <div><label style={labelStyle}>Date</label><input type="date" style={inputStyle} value={date} onChange={(e) => setDate(e.target.value)} /></div>
              <div><label style={labelStyle}>Vendor / Paid To (optional)</label><input style={inputStyle} value={vendor} onChange={(e) => setVendor(e.target.value)} placeholder="e.g. Electricity Bill" /></div>
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
          <div style={{ overflowX: "auto" }}><table style={{ minWidth: 640 }}>
            <thead><tr style={{ borderBottom: `1px solid ${PALETTE.line}` }}><Th>Date</Th><Th>Vendor</Th><Th>Category</Th><Th>Paid From</Th><Th align="right">Amount</Th><Th>By</Th><Th> </Th></tr></thead>
            <tbody>
              {expenses.map((exp) => {
                const acc = accounts.find((a) => a.id === exp.accountId);
                const paidFrom = accounts.find((a) => a.id === exp.paymentAccountId);
                if (editingId === exp.id) {
                  return (
                    <tr key={exp.id} style={{ borderBottom: `1px solid ${PALETTE.line}` }}>
                      <Td colSpan={7} style={{ padding: "12px 10px" }}>
                        <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "130px 1fr 110px", gap: 8 }}>
                          <input style={{ ...inputStyle, padding: "6px 8px" }} type="date" value={editForm.date} onChange={(e) => setEditForm({ ...editForm, date: e.target.value })} />
                          <input style={{ ...inputStyle, padding: "6px 8px" }} value={editForm.vendor} onChange={(e) => setEditForm({ ...editForm, vendor: e.target.value })} placeholder="Vendor" />
                          <input style={{ ...inputStyle, padding: "6px 8px" }} type="number" value={editForm.amount} onChange={(e) => setEditForm({ ...editForm, amount: e.target.value })} placeholder="Amount" />
                        </div>
                        <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginTop: 8 }}>
                          <select style={{ ...inputStyle, padding: "6px 8px" }} value={editForm.accountId} onChange={(e) => setEditForm({ ...editForm, accountId: e.target.value })}>
                            {expenseAccounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
                          </select>
                          <select style={{ ...inputStyle, padding: "6px 8px" }} value={editForm.paymentAccountId} onChange={(e) => setEditForm({ ...editForm, paymentAccountId: e.target.value })}>
                            {cashAccounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
                          </select>
                        </div>
                        <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
                          <button className="pin-btn" onClick={() => saveEdit(exp.id)} style={{ background: PALETTE.credit, color: "#fff", fontSize: 12, padding: "6px 14px", borderRadius: 999 }}>Save</button>
                          <button className="pin-btn" onClick={() => setEditingId(null)} style={{ background: "transparent", color: PALETTE.inkSoft, fontSize: 12 }}>Cancel</button>
                        </div>
                      </Td>
                    </tr>
                  );
                }
                return (
                  <tr key={exp.id} className="row-hover" style={{ borderBottom: `1px solid ${PALETTE.line}` }}>
                    <Td>{exp.date}</Td><Td>{exp.vendor}</Td><Td>{acc?.name || "—"}</Td>
                    <Td style={{ color: PALETTE.inkSoft, fontSize: 12.5 }}>{paidFrom?.name || "—"}</Td>
                    <Td align="right" mono>{fmtMoney(exp.amount)}</Td><Td>{exp.createdBy}</Td>
                    <Td>
                      {currentUser.role === "admin" && (
                        <div style={{ display: "flex", gap: 2 }}>
                          <button className="pin-btn" onClick={() => startEdit(exp)} style={{ background: "transparent", color: PALETTE.inkSoft, padding: 5 }}><Edit2 size={14} /></button>
                          <button className="pin-btn" onClick={() => onDelete(exp)} style={{ background: "transparent", color: PALETTE.debit, padding: 5 }}><Trash2 size={14} /></button>
                        </div>
                      )}
                    </Td>
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
   Income — manual/other income (loans, investor funds, etc.)
--------------------------------------------------------- */

function IncomeTab({ accounts, incomeEntries, currentUser, onAdd, onEdit, onDelete, onAddCategory }) {
  const [open, setOpen] = useState(false);
  const [date, setDate] = useState(todayStr());
  const [source, setSource] = useState("");
  const [accountId, setAccountId] = useState("");
  const [depositAccountId, setDepositAccountId] = useState("");
  const [amount, setAmount] = useState("");
  const [note, setNote] = useState("");
  const [newCategory, setNewCategory] = useState("");
  const [addingCategory, setAddingCategory] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [editForm, setEditForm] = useState({});

  const incomeAccounts = accounts.filter((a) => a.type === "income");
  const cashAccounts = accounts.filter((a) => a.type === "asset" && a.name !== "Accounts Receivable" && a.name !== "Inventory");

  const addCategory = async () => {
    const clean = newCategory.trim();
    if (!clean) return;
    const existingCodes = accounts.filter((a) => a.type === "income").map((a) => Number(a.code) || 0);
    const code = String(Math.max(4000, ...existingCodes, 0) + 10);
    const acc = { code, name: clean, type: "income" };
    await onAddCategory(acc);
    setNewCategory("");
    setAddingCategory(false);
  };

  const submit = (e) => {
    e.preventDefault();
    const amt = Number(amount);
    if (!accountId || !depositAccountId || !(amt > 0)) return;
    const categoryName = incomeAccounts.find((a) => a.id === accountId)?.name || "Income";
    const sourceLabel = source.trim() || categoryName;
    const entry = { date, source: sourceLabel, accountId, depositAccountId, amount: amt, note: note.trim() };
    const journalEntry = { date, memo: `Income — ${sourceLabel}`, lines: [{ accountId: depositAccountId, debit: amt, credit: 0 }, { accountId, debit: 0, credit: amt }], source: "other-income" };
    onAdd(entry, journalEntry);
    setSource(""); setAmount(""); setNote(""); setOpen(false);
  };

  const startEdit = (entry) => {
    setEditingId(entry.id);
    setEditForm({ date: entry.date, source: entry.source, accountId: entry.accountId, depositAccountId: entry.depositAccountId, amount: String(entry.amount), note: entry.note || "" });
  };
  const saveEdit = (entryId) => {
    const amt = Number(editForm.amount);
    if (!editForm.accountId || !editForm.depositAccountId || !(amt > 0)) return;
    const categoryName = incomeAccounts.find((a) => a.id === editForm.accountId)?.name || "Income";
    const sourceLabel = editForm.source.trim() || categoryName;
    const updatedEntry = { date: editForm.date, source: sourceLabel, accountId: editForm.accountId, depositAccountId: editForm.depositAccountId, amount: amt, note: editForm.note.trim() };
    const journalEntry = { date: editForm.date, memo: `Income — ${sourceLabel}`, lines: [{ accountId: editForm.depositAccountId, debit: amt, credit: 0 }, { accountId: editForm.accountId, debit: 0, credit: amt }], source: "other-income" };
    onEdit(entryId, updatedEntry, journalEntry);
    setEditingId(null);
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 10 }}>
        <PageHeader title="Income" subtitle="Record money coming in that isn't a product sale — loans, investor funds, refunds, etc." />
        <PrimaryButton onClick={() => setOpen((v) => !v)}>{open ? <X size={15} /> : <Plus size={15} />} {open ? "Cancel" : "New Income"}</PrimaryButton>
      </div>

      {open && (
        <Card style={{ marginBottom: 20 }}>
          <form onSubmit={submit}>
            <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "140px 1fr 140px", gap: 10 }}>
              <div><label style={labelStyle}>Date</label><input type="date" style={inputStyle} value={date} onChange={(e) => setDate(e.target.value)} /></div>
              <div><label style={labelStyle}>Source / Description (optional)</label><input style={inputStyle} value={source} onChange={(e) => setSource(e.target.value)} placeholder="e.g. Loan from Rahim Bhai" /></div>
              <div><label style={labelStyle}>Amount (BDT)</label><input style={inputStyle} type="number" value={amount} onChange={(e) => setAmount(e.target.value)} placeholder="0" /></div>
            </div>
            <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: 10 }}>
              <div>
                <label style={labelStyle}>Income Category</label>
                <select style={inputStyle} value={accountId} onChange={(e) => setAccountId(e.target.value)}>
                  <option value="">Select</option>
                  {incomeAccounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
                </select>
                {!addingCategory ? (
                  <GhostButton onClick={() => setAddingCategory(true)} style={{ marginTop: 4 }}>+ New category</GhostButton>
                ) : (
                  <div style={{ display: "flex", gap: 6, marginTop: 6 }}>
                    <input style={{ ...inputStyle, padding: "6px 8px", fontSize: 12.5 }} value={newCategory} onChange={(e) => setNewCategory(e.target.value)} placeholder="e.g. Loan, Investment" />
                    <button type="button" className="pin-btn" onClick={addCategory} style={{ background: PALETTE.credit, color: "#fff", fontSize: 12, padding: "6px 12px", borderRadius: 999 }}>Add</button>
                    <button type="button" className="pin-btn" onClick={() => { setAddingCategory(false); setNewCategory(""); }} style={{ background: "transparent", color: PALETTE.inkSoft, fontSize: 12 }}>Cancel</button>
                  </div>
                )}
              </div>
              <div><label style={labelStyle}>Deposit To (Account)</label><select style={inputStyle} value={depositAccountId} onChange={(e) => setDepositAccountId(e.target.value)}><option value="">Select</option>{cashAccounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}</select></div>
            </div>
            <div style={{ marginTop: 10 }}><label style={labelStyle}>Note (optional)</label><input style={inputStyle} value={note} onChange={(e) => setNote(e.target.value)} placeholder="Additional details" /></div>
            <PrimaryButton type="submit" style={{ marginTop: 16 }}>Save Income</PrimaryButton>
          </form>
        </Card>
      )}

      <Card>
        {incomeEntries.length === 0 ? <EmptyState text="No other income recorded yet" /> : (
          <div style={{ overflowX: "auto" }}><table style={{ minWidth: 640 }}>
            <thead><tr style={{ borderBottom: `1px solid ${PALETTE.line}` }}><Th>Date</Th><Th>Source</Th><Th>Category</Th><Th>Deposited To</Th><Th align="right">Amount</Th><Th>By</Th><Th> </Th></tr></thead>
            <tbody>
              {incomeEntries.map((entry) => {
                const acc = accounts.find((a) => a.id === entry.accountId);
                const depositAcc = accounts.find((a) => a.id === entry.depositAccountId);
                if (editingId === entry.id) {
                  return (
                    <tr key={entry.id} style={{ borderBottom: `1px solid ${PALETTE.line}` }}>
                      <Td colSpan={7} style={{ padding: "12px 10px" }}>
                        <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "130px 1fr 110px", gap: 8 }}>
                          <input style={{ ...inputStyle, padding: "6px 8px" }} type="date" value={editForm.date} onChange={(e) => setEditForm({ ...editForm, date: e.target.value })} />
                          <input style={{ ...inputStyle, padding: "6px 8px" }} value={editForm.source} onChange={(e) => setEditForm({ ...editForm, source: e.target.value })} placeholder="Source" />
                          <input style={{ ...inputStyle, padding: "6px 8px" }} type="number" value={editForm.amount} onChange={(e) => setEditForm({ ...editForm, amount: e.target.value })} placeholder="Amount" />
                        </div>
                        <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginTop: 8 }}>
                          <select style={{ ...inputStyle, padding: "6px 8px" }} value={editForm.accountId} onChange={(e) => setEditForm({ ...editForm, accountId: e.target.value })}>
                            {incomeAccounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
                          </select>
                          <select style={{ ...inputStyle, padding: "6px 8px" }} value={editForm.depositAccountId} onChange={(e) => setEditForm({ ...editForm, depositAccountId: e.target.value })}>
                            {cashAccounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
                          </select>
                        </div>
                        <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
                          <button className="pin-btn" onClick={() => saveEdit(entry.id)} style={{ background: PALETTE.credit, color: "#fff", fontSize: 12, padding: "6px 14px", borderRadius: 999 }}>Save</button>
                          <button className="pin-btn" onClick={() => setEditingId(null)} style={{ background: "transparent", color: PALETTE.inkSoft, fontSize: 12 }}>Cancel</button>
                        </div>
                      </Td>
                    </tr>
                  );
                }
                return (
                  <tr key={entry.id} className="row-hover" style={{ borderBottom: `1px solid ${PALETTE.line}` }}>
                    <Td>{entry.date}</Td><Td>{entry.source}</Td><Td>{acc?.name || "—"}</Td>
                    <Td style={{ color: PALETTE.inkSoft, fontSize: 12.5 }}>{depositAcc?.name || "—"}</Td>
                    <Td align="right" mono style={{ color: PALETTE.credit }}>{fmtMoney(entry.amount)}</Td>
                    <Td>{entry.createdBy}</Td>
                    <Td>
                      {currentUser.role === "admin" && (
                        <div style={{ display: "flex", gap: 2 }}>
                          <button className="pin-btn" onClick={() => startEdit(entry)} style={{ background: "transparent", color: PALETTE.inkSoft, padding: 5 }}><Edit2 size={14} /></button>
                          <button className="pin-btn" onClick={() => onDelete(entry)} style={{ background: "transparent", color: PALETTE.debit, padding: 5 }}><Trash2 size={14} /></button>
                        </div>
                      )}
                    </Td>
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

function ChartOfAccounts({ accounts, balances, currentUser, onAdd, onAddMany, onEdit, onDelete, onAdjustBalance, onTransfer }) {
  const [open, setOpen] = useState(false);
  const [form, setForm] = useState({ code: "", name: "", type: "asset" });
  const [editingId, setEditingId] = useState(null);
  const [editForm, setEditForm] = useState({ code: "", name: "" });
  const [balancingId, setBalancingId] = useState(null);
  const [balanceForm, setBalanceForm] = useState({ amount: "", date: todayStr(), note: "" });
  const [transferOpen, setTransferOpen] = useState(false);
  const [transferForm, setTransferForm] = useState({ from: "", to: "", amount: "", date: todayStr(), note: "" });
  const isAdmin = currentUser.role === "admin";
  const grouped = ACCOUNT_TYPES.map((t) => ({ ...t, items: accounts.filter((a) => a.type === t.key) }));

  const missingCategories = STANDARD_EXPENSE_CATEGORIES.filter(
    (name) => !accounts.some((a) => a.name.toLowerCase() === name.toLowerCase())
  );

  const addStandardCategories = () => {
    if (missingCategories.length === 0) return;
    const existingExpenseCodes = accounts.filter((a) => a.type === "expense").map((a) => Number(a.code) || 0);
    let nextCode = Math.max(5000, ...existingExpenseCodes) + 10;
    const toAdd = missingCategories.map((name) => {
      const acc = { code: String(nextCode), name, type: "expense" };
      nextCode += 10;
      return acc;
    });
    onAddMany(toAdd);
  };

  const submit = (e) => {
    e.preventDefault();
    if (!form.code.trim() || !form.name.trim()) return;
    onAdd({ code: form.code.trim(), name: form.name.trim(), type: form.type });
    setForm({ code: "", name: "", type: "asset" });
    setOpen(false);
  };

  const startEdit = (a) => { setEditingId(a.id); setEditForm({ code: a.code, name: a.name }); setBalancingId(null); };
  const saveEdit = (id) => {
    if (!editForm.code.trim() || !editForm.name.trim()) return;
    onEdit(id, { code: editForm.code.trim(), name: editForm.name.trim() });
    setEditingId(null);
  };

  const startBalance = (a) => { setBalancingId(a.id); setBalanceForm({ amount: String(balances[a.id] || 0), date: todayStr(), note: "" }); setEditingId(null); };
  const saveBalance = (a) => {
    const amt = Number(balanceForm.amount);
    if (isNaN(amt)) return;
    onAdjustBalance(a, amt, balanceForm.date, balanceForm.note);
    setBalancingId(null);
  };

  const submitTransfer = (e) => {
    e.preventDefault();
    const amt = Number(transferForm.amount);
    if (!transferForm.from || !transferForm.to || transferForm.from === transferForm.to || !(amt > 0)) return;
    onTransfer(transferForm.from, transferForm.to, amt, transferForm.date, transferForm.note);
    setTransferForm({ from: "", to: "", amount: "", date: todayStr(), note: "" });
    setTransferOpen(false);
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 10 }}>
        <PageHeader title="Chart of Accounts" subtitle="All accounts, balances, opening balances, and transfers" />
        {isAdmin && (
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            {missingCategories.length > 0 && (
              <GhostButton onClick={addStandardCategories}>+ Add {missingCategories.length} Standard Categories</GhostButton>
            )}
            <label className="pin-btn" onClick={() => { setTransferOpen((v) => !v); setOpen(false); }} style={{
              display: "flex", alignItems: "center", gap: 6, background: "rgba(255,255,255,0.08)", color: PALETTE.ink,
              padding: "10px 16px", borderRadius: 999, fontSize: 13.5, fontWeight: 600, border: `1px solid ${PALETTE.line}`, cursor: "pointer",
            }}>
              <Wallet size={15} /> Transfer Funds
            </label>
            <PrimaryButton onClick={() => { setOpen((v) => !v); setTransferOpen(false); }}>{open ? <X size={15} /> : <Plus size={15} />} {open ? "Cancel" : "New Account"}</PrimaryButton>
          </div>
        )}
      </div>

      {transferOpen && (
        <Card style={{ marginBottom: 20 }}>
          <div style={{ fontFamily: FONT.display, fontSize: 15, fontWeight: 600, marginBottom: 4 }}>Transfer Between Accounts</div>
          <p style={{ fontSize: 12.5, color: PALETTE.inkSoft, marginTop: 0, marginBottom: 14 }}>Move money from one account to another — e.g. Cash to Bank, or Bank to bKash.</p>
          <form onSubmit={submitTransfer} className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 140px", gap: 10 }}>
            <div>
              <label style={labelStyle}>From</label>
              <select style={inputStyle} value={transferForm.from} onChange={(e) => setTransferForm({ ...transferForm, from: e.target.value })}>
                <option value="">Select account</option>
                {accounts.map((a) => <option key={a.id} value={a.id}>{a.code} — {a.name}</option>)}
              </select>
            </div>
            <div>
              <label style={labelStyle}>To</label>
              <select style={inputStyle} value={transferForm.to} onChange={(e) => setTransferForm({ ...transferForm, to: e.target.value })}>
                <option value="">Select account</option>
                {accounts.map((a) => <option key={a.id} value={a.id}>{a.code} — {a.name}</option>)}
              </select>
            </div>
            <div><label style={labelStyle}>Amount (BDT)</label><input style={inputStyle} type="number" value={transferForm.amount} onChange={(e) => setTransferForm({ ...transferForm, amount: e.target.value })} placeholder="0" /></div>
            <div><label style={labelStyle}>Date</label><input style={inputStyle} type="date" value={transferForm.date} onChange={(e) => setTransferForm({ ...transferForm, date: e.target.value })} /></div>
            <div style={{ gridColumn: "span 2" }}><label style={labelStyle}>Note (optional)</label><input style={inputStyle} value={transferForm.note} onChange={(e) => setTransferForm({ ...transferForm, note: e.target.value })} placeholder="e.g. Weekly cash deposit" /></div>
            <div style={{ display: "flex", alignItems: "end" }}><PrimaryButton type="submit" style={{ width: "100%", justifyContent: "center" }}>Transfer</PrimaryButton></div>
          </form>
        </Card>
      )}

      {open && (
        <Card style={{ marginBottom: 20 }}>
          <form onSubmit={submit} className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "100px 1fr 160px auto", gap: 10, alignItems: "end" }}>
            <div><label style={labelStyle}>Code</label><input style={inputStyle} value={form.code} onChange={(e) => setForm({ ...form, code: e.target.value })} placeholder="6000" /></div>
            <div><label style={labelStyle}>Name</label><input style={inputStyle} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder="Account name (e.g. a new Income category)" /></div>
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
              <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                {g.items.map((a) => (
                  <div key={a.id} style={{ borderBottom: `1px solid ${PALETTE.line}`, padding: "8px 0" }}>
                    {editingId === a.id ? (
                      <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" }}>
                        <input style={{ ...inputStyle, width: 70, padding: "6px 8px" }} value={editForm.code} onChange={(e) => setEditForm({ ...editForm, code: e.target.value })} />
                        <input style={{ ...inputStyle, flex: 1, minWidth: 100, padding: "6px 8px" }} value={editForm.name} onChange={(e) => setEditForm({ ...editForm, name: e.target.value })} />
                        <button className="pin-btn" onClick={() => saveEdit(a.id)} style={{ background: PALETTE.credit, color: "#fff", fontSize: 12, padding: "5px 10px", borderRadius: 999 }}>Save</button>
                        <button className="pin-btn" onClick={() => setEditingId(null)} style={{ background: "transparent", color: PALETTE.inkSoft, fontSize: 12 }}>Cancel</button>
                      </div>
                    ) : balancingId === a.id ? (
                      <div>
                        <div style={{ fontSize: 12.5, marginBottom: 6, color: PALETTE.inkSoft }}>Set balance for <b style={{ color: PALETTE.ink }}>{a.name}</b></div>
                        <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
                          <input style={{ ...inputStyle, width: 110, padding: "6px 8px" }} type="number" value={balanceForm.amount} onChange={(e) => setBalanceForm({ ...balanceForm, amount: e.target.value })} placeholder="Amount" />
                          <input style={{ ...inputStyle, width: 140, padding: "6px 8px" }} type="date" value={balanceForm.date} onChange={(e) => setBalanceForm({ ...balanceForm, date: e.target.value })} />
                          <input style={{ ...inputStyle, flex: 1, minWidth: 100, padding: "6px 8px" }} value={balanceForm.note} onChange={(e) => setBalanceForm({ ...balanceForm, note: e.target.value })} placeholder="Note (optional)" />
                          <button className="pin-btn" onClick={() => saveBalance(a)} style={{ background: PALETTE.credit, color: "#fff", fontSize: 12, padding: "5px 10px", borderRadius: 999 }}>Save</button>
                          <button className="pin-btn" onClick={() => setBalancingId(null)} style={{ background: "transparent", color: PALETTE.inkSoft, fontSize: 12 }}>Cancel</button>
                        </div>
                      </div>
                    ) : (
                      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <span style={{ fontFamily: FONT.mono, fontSize: 12.5, color: PALETTE.inkSoft, width: 46 }}>{a.code}</span>
                        <span style={{ flex: 1, fontSize: 13.5 }}>{a.name}</span>
                        <span style={{ fontFamily: FONT.mono, fontSize: 13.5, marginRight: 4 }}>{fmtMoney(balances[a.id] || 0)}</span>
                        {isAdmin && (
                          <div style={{ display: "flex", gap: 2 }}>
                            <button className="pin-btn" onClick={() => startBalance(a)} title="Set balance" style={{ background: "transparent", color: PALETTE.inkSoft, padding: 5 }}><Wallet size={13} /></button>
                            <button className="pin-btn" onClick={() => startEdit(a)} title="Rename" style={{ background: "transparent", color: PALETTE.inkSoft, padding: 5 }}><Edit2 size={13} /></button>
                            <button className="pin-btn" onClick={() => onDelete(a.id)} title="Delete" style={{ background: "transparent", color: PALETTE.debit, padding: 5 }}><Trash2 size={13} /></button>
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                ))}
              </div>
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

function Reports({ accounts, balances, invoices, expenses, journal, products }) {
  const [fromDate, setFromDate] = useState("");
  const [toDate, setToDate] = useState("");

  const inRange = (dateStr) => {
    if (!dateStr) return true;
    if (fromDate && dateStr < fromDate) return false;
    if (toDate && dateStr > toDate) return false;
    return true;
  };

  const monthKey = (dateStr) => (dateStr && dateStr.length >= 7 ? dateStr.slice(0, 7) : "Unknown");
  const monthLabel = (key) => {
    if (key === "Unknown") return "Unknown";
    const [y, m] = key.split("-");
    const d = new Date(Number(y), Number(m) - 1, 1);
    return d.toLocaleDateString("en-US", { month: "short", year: "numeric" });
  };

  // Sales by month — from invoice totals (what was invoiced, regardless of payment)
  const salesByMonth = useMemo(() => {
    const map = {};
    invoices.filter((inv) => inRange(inv.date)).forEach((inv) => {
      const k = monthKey(inv.date);
      map[k] = (map[k] || 0) + (Number(inv.total) || 0);
    });
    return Object.entries(map).sort((a, b) => b[0].localeCompare(a[0]));
  }, [invoices, fromDate, toDate]);

  // Income by month — from journal lines that credited an income-type account (cash actually received)
  const incomeByMonth = useMemo(() => {
    const incomeAccountIds = new Set(accounts.filter((a) => a.type === "income").map((a) => a.id));
    const map = {};
    journal.filter((j) => inRange(j.date)).forEach((j) => {
      j.lines.forEach((line) => {
        if (incomeAccountIds.has(line.accountId)) {
          const k = monthKey(j.date);
          map[k] = (map[k] || 0) + (Number(line.credit) || 0) - (Number(line.debit) || 0);
        }
      });
    });
    return Object.entries(map).sort((a, b) => b[0].localeCompare(a[0]));
  }, [journal, accounts, fromDate, toDate]);

  // Expense by month
  const expenseByMonth = useMemo(() => {
    const map = {};
    expenses.filter((e) => inRange(e.date)).forEach((e) => {
      const k = monthKey(e.date);
      map[k] = (map[k] || 0) + (Number(e.amount) || 0);
    });
    return Object.entries(map).sort((a, b) => b[0].localeCompare(a[0]));
  }, [expenses, fromDate, toDate]);

  // Sales by Customer Summary — grouped by "Sales By" team/person, then customer
  const salesByCustomer = useMemo(() => {
    const groups = {};
    invoices.filter((inv) => inRange(inv.date)).forEach((inv) => {
      const groupName = inv.salesBy?.trim() || "Unassigned";
      const customerName = inv.customer?.trim() || "Unknown Customer";
      if (!groups[groupName]) groups[groupName] = { customers: {}, total: 0 };
      groups[groupName].customers[customerName] = (groups[groupName].customers[customerName] || 0) + (Number(inv.total) || 0);
      groups[groupName].total += Number(inv.total) || 0;
    });
    return Object.entries(groups)
      .map(([name, g]) => ({ name, total: g.total, customers: Object.entries(g.customers).sort((a, b) => b[1] - a[1]) }))
      .sort((a, b) => b.total - a.total);
  }, [invoices, fromDate, toDate]);
  const salesByCustomerGrandTotal = salesByCustomer.reduce((s, g) => s + g.total, 0);

  // Sales by Product Summary — grouped by product category
  const salesByProduct = useMemo(() => {
    const productMap = {};
    products.forEach((p) => { productMap[p.id] = p; });
    const groups = {};
    invoices.filter((inv) => inRange(inv.date)).forEach((inv) => {
      (inv.items || []).forEach((it) => {
        const p = it.productId ? productMap[it.productId] : null;
        const category = p ? (p.category || "Uncategorized") : "Other (no product link)";
        const key = p ? p.name : (it.desc || "Item");
        const qty = Number(it.qty) || 0;
        const amount = qty * (Number(it.rate) || 0);
        const cos = p ? qty * (Number(p.costPrice) || 0) : 0;
        if (!groups[category]) groups[category] = { items: {}, qty: 0, amount: 0, cos: 0 };
        if (!groups[category].items[key]) groups[category].items[key] = { qty: 0, amount: 0, cos: 0 };
        groups[category].items[key].qty += qty;
        groups[category].items[key].amount += amount;
        groups[category].items[key].cos += cos;
        groups[category].qty += qty;
        groups[category].amount += amount;
        groups[category].cos += cos;
      });
    });
    return Object.entries(groups)
      .map(([name, g]) => ({
        name, qty: g.qty, amount: g.amount, cos: g.cos, margin: g.amount - g.cos,
        items: Object.entries(g.items).map(([iname, iv]) => ({ name: iname, ...iv, margin: iv.amount - iv.cos })).sort((a, b) => b.amount - a.amount),
      }))
      .sort((a, b) => b.amount - a.amount);
  }, [invoices, products, fromDate, toDate]);
  const salesByProductGrandTotal = salesByProduct.reduce((s, g) => s + g.amount, 0);

  // Profit & Loss — Income (by income account, in range), Cost of Sales (COGS, in range), Expenses (by category, in range)
  const profitLoss = useMemo(() => {
    const incomeAccounts = accounts.filter((a) => a.type === "income");
    const cogsAccount = accounts.find((a) => a.name === "Cost of Goods Sold");
    const incomeTotals = {};
    let cogsTotal = 0;
    journal.filter((j) => inRange(j.date)).forEach((j) => {
      j.lines.forEach((line) => {
        const acc = accounts.find((a) => a.id === line.accountId);
        if (!acc) return;
        const net = (Number(line.credit) || 0) - (Number(line.debit) || 0);
        if (acc.type === "income") incomeTotals[acc.id] = (incomeTotals[acc.id] || 0) + net;
        if (cogsAccount && acc.id === cogsAccount.id) cogsTotal += (Number(line.debit) || 0) - (Number(line.credit) || 0);
      });
    });
    const incomeRows = incomeAccounts.map((a) => ({ name: a.name, amount: incomeTotals[a.id] || 0 })).filter((r) => r.amount !== 0);
    const totalIncome = incomeRows.reduce((s, r) => s + r.amount, 0);
    const discountsGiven = invoices.filter((inv) => inRange(inv.date)).reduce((s, inv) => s + (Number(inv.discount) || 0), 0);
    const grossProfit = totalIncome - cogsTotal;

    const expenseAccounts = accounts.filter((a) => a.type === "expense" && a.name !== "Cost of Goods Sold");
    const expenseTotals = {};
    expenses.filter((e) => inRange(e.date)).forEach((e) => {
      expenseTotals[e.accountId] = (expenseTotals[e.accountId] || 0) + (Number(e.amount) || 0);
    });
    const expenseRows = expenseAccounts.map((a) => ({ name: a.name, amount: expenseTotals[a.id] || 0 })).filter((r) => r.amount !== 0).sort((a, b) => b.amount - a.amount);
    const totalExpenses = expenseRows.reduce((s, r) => s + r.amount, 0);
    const netEarnings = grossProfit - totalExpenses;

    return { incomeRows, totalIncome, discountsGiven, cogsTotal, grossProfit, expenseRows, totalExpenses, netEarnings };
  }, [journal, accounts, expenses, invoices, fromDate, toDate]);

  let totalDebit = 0, totalCredit = 0;
  const rows = accounts.map((a) => {
    const bal = balances[a.id] || 0;
    const side = normalSideFor(a.type);
    const debit = side === "debit" ? Math.max(bal, 0) : Math.max(-bal, 0);
    const credit = side === "credit" ? Math.max(bal, 0) : Math.max(-bal, 0);
    totalDebit += debit; totalCredit += credit;
    return { ...a, debit, credit };
  });

  const MonthlyCard = ({ title, subtitle, data, tone }) => {
    const total = data.reduce((s, [, v]) => s + v, 0);
    return (
      <Card>
        <div style={{ fontFamily: FONT.display, fontSize: 15, fontWeight: 600, marginBottom: 2 }}>{title}</div>
        <div style={{ fontSize: 12, color: PALETTE.inkSoft, marginBottom: 12 }}>{subtitle}</div>
        {data.length === 0 ? <EmptyState text="No data for this period" /> : (
          <>
            <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 12 }}>
              {data.map(([k, v]) => (
                <div key={k} style={{ display: "flex", justifyContent: "space-between", fontSize: 13.5 }}>
                  <span style={{ color: PALETTE.inkSoft }}>{monthLabel(k)}</span>
                  <span style={{ fontFamily: FONT.mono, color: PALETTE.ink }}>{fmtMoney(v)}</span>
                </div>
              ))}
            </div>
            <div style={{ borderTop: `1px solid ${PALETTE.line}`, paddingTop: 10, display: "flex", justifyContent: "space-between", fontSize: 14, fontWeight: 700 }}>
              <span>Total</span>
              <span style={{ fontFamily: FONT.mono, color: tone === "credit" ? PALETTE.credit : tone === "debit" ? PALETTE.debit : PALETTE.ink }}>{fmtMoney(total)}</span>
            </div>
          </>
        )}
      </Card>
    );
  };

  return (
    <div>
      <PageHeader title="Reports" subtitle="Sales, income, and expense reports by month — filter by date to customize" />

      <Card style={{ marginBottom: 20, padding: 16 }}>
        <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr auto", gap: 10, alignItems: "end" }}>
          <div><label style={labelStyle}>From</label><input type="date" style={inputStyle} value={fromDate} onChange={(e) => setFromDate(e.target.value)} /></div>
          <div><label style={labelStyle}>To</label><input type="date" style={inputStyle} value={toDate} onChange={(e) => setToDate(e.target.value)} /></div>
          {(fromDate || toDate) && <GhostButton onClick={() => { setFromDate(""); setToDate(""); }}>Clear filter</GhostButton>}
        </div>
      </Card>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(260px,1fr))", gap: 16, marginBottom: 24 }}>
        <MonthlyCard title="Sales Report" subtitle="Total invoiced amount, by month" data={salesByMonth} tone="ink" />
        <MonthlyCard title="Income Report" subtitle="Actual money received, by month" data={incomeByMonth} tone="credit" />
        <MonthlyCard title="Expense Report" subtitle="Total expenses, by month" data={expenseByMonth} tone="debit" />
      </div>

      <PageHeader title="Sales by Customer Summary" subtitle="Invoiced amount grouped by Sales By, then customer" />
      <Card style={{ marginBottom: 24 }}>
        {salesByCustomer.length === 0 ? <EmptyState text="No sales for this period" /> : (
          <div style={{ overflowX: "auto" }}><table style={{ minWidth: 460 }}>
            <tbody>
              {salesByCustomer.map((g) => (
                <React.Fragment key={g.name}>
                  <tr><Td colSpan={2} style={{ paddingTop: 14, fontWeight: 700, color: PALETTE.accent, fontSize: 12.5, textTransform: "uppercase", letterSpacing: 0.3 }}>{g.name}</Td></tr>
                  {g.customers.map(([cname, amt]) => (
                    <tr key={cname} className="row-hover" style={{ borderBottom: `1px solid ${PALETTE.line}` }}>
                      <Td style={{ paddingLeft: 20 }}>{cname}</Td>
                      <Td align="right" mono>{fmtMoney(amt)}</Td>
                    </tr>
                  ))}
                  <tr style={{ borderBottom: `1px solid ${PALETTE.line}` }}>
                    <Td style={{ fontWeight: 700, fontSize: 12.5 }}>Total for {g.name}</Td>
                    <Td align="right" mono style={{ fontWeight: 700 }}>{fmtMoney(g.total)}</Td>
                  </tr>
                </React.Fragment>
              ))}
              <tr style={{ borderTop: `2px solid ${PALETTE.ink}` }}>
                <Td style={{ fontWeight: 700, fontSize: 14 }}>Grand Total</Td>
                <Td align="right" mono style={{ fontWeight: 700, fontSize: 14 }}>{fmtMoney(salesByCustomerGrandTotal)}</Td>
              </tr>
            </tbody>
          </table></div>
        )}
      </Card>

      <PageHeader title="Sales by Product Summary" subtitle="Quantity, revenue, cost of sales, and margin — grouped by category" />
      <Card style={{ marginBottom: 24 }}>
        {salesByProduct.length === 0 ? <EmptyState text="No sales for this period" /> : (
          <div style={{ overflowX: "auto" }}><table style={{ minWidth: 640 }}>
            <thead><tr style={{ borderBottom: `1px solid ${PALETTE.line}` }}><Th>Product</Th><Th align="right">Qty</Th><Th align="right">Amount</Th><Th align="right">COS</Th><Th align="right">Margin</Th><Th align="right">Margin %</Th></tr></thead>
            <tbody>
              {salesByProduct.map((g) => (
                <React.Fragment key={g.name}>
                  <tr><Td colSpan={6} style={{ paddingTop: 14, fontWeight: 700, color: PALETTE.accent, fontSize: 12.5, textTransform: "uppercase", letterSpacing: 0.3 }}>{g.name}</Td></tr>
                  {g.items.map((it) => (
                    <tr key={it.name} className="row-hover" style={{ borderBottom: `1px solid ${PALETTE.line}` }}>
                      <Td style={{ paddingLeft: 20 }}>{it.name}</Td>
                      <Td align="right" mono>{it.qty}</Td>
                      <Td align="right" mono>{fmtMoney(it.amount)}</Td>
                      <Td align="right" mono>{fmtMoney(it.cos)}</Td>
                      <Td align="right" mono>{fmtMoney(it.margin)}</Td>
                      <Td align="right" mono>{it.amount > 0 ? `${((it.margin / it.amount) * 100).toFixed(1)}%` : "—"}</Td>
                    </tr>
                  ))}
                  <tr style={{ borderBottom: `1px solid ${PALETTE.line}` }}>
                    <Td style={{ fontWeight: 700, fontSize: 12.5 }}>Total for {g.name}</Td>
                    <Td align="right" mono style={{ fontWeight: 700 }}>{g.qty}</Td>
                    <Td align="right" mono style={{ fontWeight: 700 }}>{fmtMoney(g.amount)}</Td>
                    <Td align="right" mono style={{ fontWeight: 700 }}>{fmtMoney(g.cos)}</Td>
                    <Td align="right" mono style={{ fontWeight: 700 }}>{fmtMoney(g.margin)}</Td>
                    <Td align="right" mono style={{ fontWeight: 700 }}>{g.amount > 0 ? `${((g.margin / g.amount) * 100).toFixed(1)}%` : "—"}</Td>
                  </tr>
                </React.Fragment>
              ))}
              <tr style={{ borderTop: `2px solid ${PALETTE.ink}` }}>
                <Td style={{ fontWeight: 700, fontSize: 14 }}>Grand Total</Td>
                <Td></Td>
                <Td align="right" mono style={{ fontWeight: 700, fontSize: 14 }}>{fmtMoney(salesByProductGrandTotal)}</Td>
                <Td></Td><Td></Td><Td></Td>
              </tr>
            </tbody>
          </table></div>
        )}
      </Card>

      <PageHeader title="Profit & Loss" subtitle="Income, cost of sales, and expenses for the selected period" />
      <Card style={{ marginBottom: 24 }}>
        <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 6 }}>Income</div>
        {profitLoss.discountsGiven > 0 && (
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, color: PALETTE.inkSoft, padding: "4px 0" }}>
            <span>Discounts given</span><span style={{ fontFamily: FONT.mono }}>-{fmtMoney(profitLoss.discountsGiven)}</span>
          </div>
        )}
        {profitLoss.incomeRows.map((r) => (
          <div key={r.name} style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "4px 0" }}>
            <span style={{ color: PALETTE.inkSoft }}>{r.name}</span><span style={{ fontFamily: FONT.mono }}>{fmtMoney(r.amount)}</span>
          </div>
        ))}
        <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 700, fontSize: 13.5, borderTop: `1px solid ${PALETTE.line}`, marginTop: 6, paddingTop: 8 }}>
          <span>Total Income</span><span style={{ fontFamily: FONT.mono }}>{fmtMoney(profitLoss.totalIncome)}</span>
        </div>

        <div style={{ fontWeight: 700, fontSize: 13, marginTop: 20, marginBottom: 6 }}>Cost of Sales</div>
        <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 700, fontSize: 13.5, borderTop: `1px solid ${PALETTE.line}`, paddingTop: 8 }}>
          <span>Total for Cost of Sales</span><span style={{ fontFamily: FONT.mono }}>{fmtMoney(profitLoss.cogsTotal)}</span>
        </div>

        <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 700, fontSize: 15, marginTop: 16, color: PALETTE.credit }}>
          <span>Gross Profit</span><span style={{ fontFamily: FONT.mono }}>{fmtMoney(profitLoss.grossProfit)}</span>
        </div>

        <div style={{ fontWeight: 700, fontSize: 13, marginTop: 20, marginBottom: 6 }}>Expenses</div>
        {profitLoss.expenseRows.length === 0 ? <EmptyState text="No expenses for this period" /> : profitLoss.expenseRows.map((r) => (
          <div key={r.name} style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "4px 0" }}>
            <span style={{ color: PALETTE.inkSoft }}>{r.name}</span><span style={{ fontFamily: FONT.mono }}>{fmtMoney(r.amount)}</span>
          </div>
        ))}
        <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 700, fontSize: 13.5, borderTop: `1px solid ${PALETTE.line}`, marginTop: 6, paddingTop: 8 }}>
          <span>Total Expenses</span><span style={{ fontFamily: FONT.mono }}>{fmtMoney(profitLoss.totalExpenses)}</span>
        </div>

        <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 700, fontSize: 17, marginTop: 18, paddingTop: 14, borderTop: `2px solid ${PALETTE.ink}`, color: profitLoss.netEarnings >= 0 ? PALETTE.credit : PALETTE.debit }}>
          <span>Net Earnings</span><span style={{ fontFamily: FONT.mono }}>{fmtMoney(profitLoss.netEarnings)}</span>
        </div>
      </Card>

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

function UsersPanel({ users, onAdd, onEdit, onDelete }) {
  const [open, setOpen] = useState(false);
  const [name, setName] = useState("");
  const [pin, setPin] = useState("");
  const [role, setRole] = useState("staff");
  const [permissions, setPermissions] = useState(ALL_TABS.map((t) => t.key));
  const [editingId, setEditingId] = useState(null);
  const [editRole, setEditRole] = useState("staff");
  const [editPermissions, setEditPermissions] = useState([]);

  const togglePerm = (key, list, setList) => {
    if (key === "dashboard") return; // always included
    setList(list.includes(key) ? list.filter((k) => k !== key) : [...list, key]);
  };

  const submit = (e) => {
    e.preventDefault();
    if (!name.trim() || pin.trim().length < 4) return;
    onAdd({ name: name.trim(), pin: pin.trim(), role, permissions: role === "staff" ? permissions : ALL_TABS.map((t) => t.key) });
    setName(""); setPin(""); setRole("staff"); setPermissions(ALL_TABS.map((t) => t.key)); setOpen(false);
  };

  const startEdit = (u) => {
    setEditingId(u.id);
    setEditRole(u.role);
    setEditPermissions(Array.isArray(u.permissions) ? u.permissions : ALL_TABS.map((t) => t.key));
  };
  const saveEdit = (id) => {
    onEdit(id, { role: editRole, permissions: editRole === "staff" ? editPermissions : ALL_TABS.map((t) => t.key) });
    setEditingId(null);
  };

  const PermGrid = ({ list, setList, disabled }) => (
    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(140px,1fr))", gap: 6, marginTop: 4 }}>
      {ALL_TABS.map((t) => (
        <label key={t.key} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, color: disabled ? PALETTE.inkSoft : PALETTE.ink, opacity: t.key === "dashboard" ? 0.6 : 1 }}>
          <input
            type="checkbox"
            checked={t.key === "dashboard" ? true : list.includes(t.key)}
            disabled={disabled || t.key === "dashboard"}
            onChange={() => togglePerm(t.key, list, setList)}
          />
          {t.label}
        </label>
      ))}
    </div>
  );

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 10 }}>
        <PageHeader title="Users" subtitle="Manage your team's access and which menus each person can see" />
        <PrimaryButton onClick={() => setOpen((v) => !v)}>{open ? <X size={15} /> : <Plus size={15} />} {open ? "Cancel" : "New User"}</PrimaryButton>
      </div>

      {open && (
        <Card style={{ marginBottom: 20 }}>
          <form onSubmit={submit}>
            <div className="responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 140px 140px", gap: 10 }}>
              <div><label style={labelStyle}>Name</label><input style={inputStyle} value={name} onChange={(e) => setName(e.target.value)} /></div>
              <div><label style={labelStyle}>PIN (4+ digits)</label><input style={inputStyle} value={pin} onChange={(e) => setPin(e.target.value)} /></div>
              <div><label style={labelStyle}>Role</label><select style={inputStyle} value={role} onChange={(e) => setRole(e.target.value)}><option value="staff">Staff (limited)</option><option value="admin">Admin</option></select></div>
            </div>
            {role === "staff" && (
              <div style={{ marginTop: 14 }}>
                <label style={labelStyle}>Menu Access</label>
                <p style={{ fontSize: 11.5, color: PALETTE.inkSoft, margin: "2px 0 4px" }}>Choose which tabs this person can see. Dashboard is always visible.</p>
                <PermGrid list={permissions} setList={setPermissions} />
              </div>
            )}
            <PrimaryButton type="submit" style={{ marginTop: 16 }}>Add User</PrimaryButton>
          </form>
        </Card>
      )}

      <Card>
        <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
          {users.map((u) => (
            <div key={u.id} style={{ borderBottom: `1px solid ${PALETTE.line}`, padding: "10px 0" }}>
              {editingId === u.id ? (
                <div>
                  <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
                    <span style={{ fontSize: 13.5, fontWeight: 600 }}>{u.name}</span>
                    <select style={{ ...inputStyle, width: 140, padding: "6px 8px" }} value={editRole} onChange={(e) => setEditRole(e.target.value)}>
                      <option value="staff">Staff (limited)</option>
                      <option value="admin">Admin</option>
                    </select>
                  </div>
                  {editRole === "staff" && <PermGrid list={editPermissions} setList={setEditPermissions} />}
                  <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
                    <button className="pin-btn" onClick={() => saveEdit(u.id)} style={{ background: PALETTE.credit, color: "#fff", fontSize: 12, padding: "6px 14px", borderRadius: 999 }}>Save</button>
                    <button className="pin-btn" onClick={() => setEditingId(null)} style={{ background: "transparent", color: PALETTE.inkSoft, fontSize: 12 }}>Cancel</button>
                  </div>
                </div>
              ) : (
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 13.5, fontWeight: 600 }}>{u.name}</div>
                    <div style={{ fontSize: 11.5, color: PALETTE.inkSoft, marginTop: 2 }}>
                      {u.role === "admin" ? "Admin — full access" : `Staff — ${Array.isArray(u.permissions) ? u.permissions.length : ALL_TABS.length} of ${ALL_TABS.length} menus`}
                    </div>
                  </div>
                  <button className="pin-btn" onClick={() => startEdit(u)} style={{ background: "transparent", color: PALETTE.inkSoft, padding: 6 }}><Edit2 size={14} /></button>
                  <button className="pin-btn" onClick={() => onDelete(u.id)} style={{ background: "transparent", color: PALETTE.debit, padding: 6 }}><Trash2 size={14} /></button>
                </div>
              )}
            </div>
          ))}
        </div>
      </Card>
      <p style={{ fontSize: 12, color: PALETTE.inkSoft, marginTop: 12 }}>Note: this PIN system is a simple access control for your 4–5 trusted team members — not banking-grade security. Don't share the link outside your team.</p>
    </div>
  );
}
