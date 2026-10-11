import { memo, useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, Component, createContext, useContext, Fragment } from "react";
import { createPortal } from "react-dom";

class ViewErrorBoundary extends Component {
  constructor(props) { super(props); this.state = { hasError: false, error: null }; }
  static getDerivedStateFromError(error) { return { hasError: true, error }; }
  render() {
    if (this.state.hasError) return (
      <div style={{ padding: "40px 24px", fontFamily: "monospace", color: "#dc2626", background: "#fff5f5", minHeight: "100vh" }}>
        <strong style={{ fontSize: "1rem" }}>CRASH:</strong>
        <pre style={{ marginTop: 12, fontSize: "0.78rem", whiteSpace: "pre-wrap", wordBreak: "break-all" }}>
          {this.state.error?.message}
          {"\n\n"}
          {this.state.error?.stack}
        </pre>
      </div>
    );
    return this.props.children;
  }
}
import { createRoot } from "react-dom/client";
import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, ImageRun, TabStopType, VerticalAlign, HeightRule, LineRuleType, SectionType,
} from "docx";
import ExcelJS from "exceljs";
import {
  ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  PieChart, Pie, Cell,
} from "recharts";
import "./styles.css";
import LeaveView, { LeaveSignatoriesPanel } from "./LeaveView.jsx";
import psaLogo from "./psa-logo.png";
import { PSA_RP_B64 } from "./psa-rp-logo.js";
import { BAGONG_PILIPINAS_B64 } from "./bagong-pilipinas-logo.js";

const apiUrl = import.meta.env.VITE_API_URL ?? "/api";

const EMPLOYMENT_TYPES = ["Regular", "Contractual"];
const PAY_TYPES = ["Regular", "13th Month", "Mid-Year Bonus", "Year-End Bonus", "Other"];
const PERA_AMOUNT = 2000;

// Global deduction/earning types — fetched once from DB, used across all views
const TypesContext = createContext({ dedTypes: [], earnTypes: [], setDedTypes: () => {}, setEarnTypes: () => {}, signatories: null });

// "DEDUCTIONS" column — computed rows are auto-derived from basic pay; item rows are typed per employee.
const DEDUCTIONS_CATALOG = [
  {
    group: "GSIS",
    computed: { label: "Life and Ret", rate: 0.09, col: "gsis_life_ret" },
    items: [
      { label: "GSIS Consolidated Loan", col: "gsis_consolidated_loan" },
      { label: "CPL",                    col: "gsis_cpl" },
      { label: "GSIS MPL Lite",          col: "gsis_mpl_lite" },
      { label: "GSIS MPL",               col: "gsis_mpl" },
      { label: "REL Loan",               col: "gsis_rel_loan" },
      { label: "Policy Loan Regular",    col: "gsis_policy_loan" },
      { label: "Educational Asst.",      col: "gsis_educational" },
      { label: "Cash Advance",           col: "gsis_cash_advance" },
      { label: "Emergency Loan",         col: "gsis_emergency_loan" },
      { label: "Optional Insurance",      col: "optl_ins" },
      { label: "GFAL",                    col: "gfal" },
      { label: "UOLI Loan",               col: "uoli_loan" },
    ],
  },
  {
    group: "PAG-IBIG",
    computed: { label: "Premium", rate: 0.02, col: "pagibig_premium" },
    items: [
      { label: "MP2",          col: "pagibig_mp2" },
      { label: "MPL/Calamity", col: "pagibig_mpl_calamity" },
    ],
  },
  { group: "LBP Salary Loan", col: "lbp_salary_loan",  items: [], bold: false },
  { group: "Medic",    computed: { label: "Medic",    rate: 0.025, col: "medic" },     col: "medic",           items: [], bold: false },
  { group: "Medic2",   col: "medic2",           items: [], bold: false },
  { group: "W/Hold Tax", col: "withholding_tax", items: [], bold: false },
  { group: "A/P for Last Payment", col: "ap_last_payment", items: [], bold: false },
];

// "LESS DEDUCTIONS" column — all typed per employee, except Union Dues which is computed.
const LESS_DEDUCTIONS_CATALOG = [
  { group: "Union Dues", computed: { label: "Union Dues", rate: 0.005, col: "union_dues" }, col: "union_dues", items: [], bold: false },
  {
    group: "PROVIDENT FUND",
    items: [
      { label: "Premium",                            col: "provident_premium" },
      { label: "General Purpose Loan (GPL)",         col: "provident_gpl" },
      { label: "Emergency Loan (EML)",               col: "provident_eml" },
      { label: "Salary Loan-Motor Vehicle (SLMV)",   col: "provident_slmv" },
    ],
  },
  { group: "CEMCO", col: "cemco", items: [], bold: false },
];

// Builds a DEDUCTIONS_CATALOG-compatible structure from DB deduction types.
// Rules:
//   2+ manual items under same category → group with dropdown picker (e.g. GSIS, PAG-IBIG)
//   computed + exactly 1 manual → two individual sibling rows (e.g. MEDIC + MEDIC 2)
//   single manual, no computed → standalone row
//   computed only → standalone computed row
function buildCatalogFromTypes(types, side) {
  const groups = {};
  const groupOrder = [];
  for (const t of types.filter(t => t.side === side)) {
    const key = t.category || t.display_name || t.name;
    if (!groups[key]) {
      groups[key] = { group: key, _manualItems: [], computed: null, bold: false };
      groupOrder.push(key);
    }
    if (t.computation_mode === 'rate') {
      groups[key].computed = { label: t.display_name || t.name, rate: Number(t.rate), col: t.name, merge_in_excel: !!t.merge_in_excel };
      groups[key].col = t.name;
    } else {
      groups[key]._manualItems.push({ label: t.display_name || t.name, col: t.name, merge_in_excel: !!t.merge_in_excel });
    }
  }
  return groupOrder.flatMap(key => {
    const g = groups[key];
    // Single manual, no computed → standalone row using item label
    if (g._manualItems.length === 1 && !g.computed) {
      return [{ group: g._manualItems[0].label, items: [], computed: null, bold: false, col: g._manualItems[0].col }];
    }
    // Computed only → standalone computed row
    if (g.computed && g._manualItems.length === 0) {
      return [{ group: g.computed.label, items: [], computed: g.computed, bold: false, col: g.computed.col }];
    }
    // Computed + exactly 1 manual → two sibling rows (MEDIC + MEDIC 2)
    if (g.computed && g._manualItems.length === 1) {
      return [
        { group: g.computed.label, items: [], computed: g.computed, bold: false, col: g.computed.col },
        { group: g._manualItems[0].label, items: [], computed: null, bold: false, col: g._manualItems[0].col },
      ];
    }
    // 2+ manual items (with or without computed) → group picker; items carry per-item merge_in_excel
    return [{ group: g.group, items: g._manualItems, computed: g.computed, bold: g.bold, col: g.col }];
  });
}

// Flattens a deduction catalog into a single ordered list of {label, col} leaf columns
// (e.g. for the Excel register export, where every deduction type needs its own column).
function flattenDeductionCatalog(catalog) {
  const cols = [];
  for (const g of catalog) {
    if (g.items.length === 0) {
      cols.push({ label: g.group, col: g.col });
    } else {
      if (g.computed) cols.push({ label: `${g.group} - ${g.computed.label}`, col: g.computed.col });
      for (const item of g.items) cols.push({ label: `${g.group} - ${item.label}`, col: item.col });
    }
  }
  return cols;
}

function sumCatalogTyped(catalog, ded) {
  let sum = 0;
  for (const g of catalog) {
    if (g.items.length === 0) {
      if (g.computed) continue;
      sum += Number(ded[g.col] ?? 0);
    } else {
      for (const item of g.items) sum += Number(ded[item.col] ?? 0);
    }
  }
  return sum;
}

const roundAmt = v => Math.round(v * 100) / 100;

function sumCatalogComputed(catalog, ded, basicPay) {
  return catalog.reduce((s, g) => {
    if (!g.computed) return s;
    const stored = ded[g.computed.col];
    if (stored != null) return s + Number(stored);
    if (!g.computed.rate) return s;
    return s + roundAmt(basicPay * g.computed.rate);
  }, 0);
}

function fmt(n) {
  return Number(n ?? 0).toLocaleString("en-PH", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}


// Plain-data version of what DeductionGroup renders — used by the .docx exporter,
// which builds paragraphs directly rather than JSX. Always lists every catalog item
// (read-only convention), matching the printed payslip rather than the admin editor.
function catalogRows(catalog, ded, basicPay) {
  const rows = [];
  for (const g of catalog) {
    const hasItems = g.items.length > 0;
    if (!hasItems) {
      const value = g.computed ? roundAmt(basicPay * g.computed.rate) : Number(ded[g.col] ?? 0);
      if (value !== 0) rows.push({ label: g.group, value, bold: g.bold !== false });
      continue;
    }
    const subRows = [];
    if (g.computed) {
      const v = roundAmt(basicPay * g.computed.rate);
      if (v !== 0) subRows.push({ label: g.computed.label, value: v, indent: true });
    }
    for (const item of g.items) {
      const v = Number(ded[item.col] ?? 0);
      if (v !== 0) subRows.push({ label: item.label, value: v, indent: true });
    }
    if (subRows.length > 0) {
      rows.push({ label: g.group, header: true });
      rows.push(...subRows);
    }
  }
  return rows;
}

function formatPeriodRange(period) {
  if (!period) return "";
  const start = new Date(period.period_start);
  const end   = new Date(period.period_end);
  const startMonth = start.toLocaleDateString("en-PH", { month: "long" });
  const endMonth   = end.toLocaleDateString("en-PH", { month: "long" });
  if (startMonth === endMonth && start.getFullYear() === end.getFullYear()) {
    return `${startMonth} ${String(start.getDate()).padStart(2, "0")}-${String(end.getDate()).padStart(2, "0")}, ${end.getFullYear()}`;
  }
  return `${startMonth} ${start.getDate()}, ${start.getFullYear()} - ${endMonth} ${end.getDate()}, ${end.getFullYear()}`;
}

// ── Payslip .docx export ────────────────────────────────────────────────────
const DOCX_NO_BORDERS = {
  top:    { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left:   { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right:  { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  insideVertical:   { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};
const DOCX_OUTER_BORDER = {
  top:    { style: BorderStyle.DOUBLE, size: 6, color: "0d2346" },
  bottom: { style: BorderStyle.DOUBLE, size: 6, color: "0d2346" },
  left:   { style: BorderStyle.DOUBLE, size: 6, color: "0d2346" },
  right:  { style: BorderStyle.DOUBLE, size: 6, color: "0d2346" },
  insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  insideVertical:   { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};

function docxRow({ label, value, bold, indent, header }) {
  if (header) {
    return new Paragraph({
      spacing: { before: 55, after: 10, line: 220, lineRule: LineRuleType.EXACT },
      children: [
        new TextRun({ text: label.toUpperCase(), bold: true, size: 16, color: "1a3a6b", allCaps: true }),
      ],
    });
  }
  const numVal = Number(value ?? 0);
  // Non-bold rows show blank when value is 0; bold totals always show the number
  const display = bold ? fmt(numVal) : (numVal === 0 ? "" : fmt(numVal));
  return new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: 3900 }],
    indent: indent ? { left: 160 } : undefined,
    spacing: { after: 6, line: 250, lineRule: LineRuleType.EXACT },
    border: bold ? { top: { style: BorderStyle.SINGLE, size: 5, color: "0d2346" } } : undefined,
    children: [
      new TextRun({ text: label, bold: !!bold, size: 20, color: bold ? "0d2346" : "445566" }),
      new TextRun({ text: `\t${display}`, bold: !!bold, size: 20, color: "0d2346" }),
    ],
  });
}

function docxTwoColumnTable(leftRows, rightRows) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: DOCX_NO_BORDERS,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 57, type: WidthType.PERCENTAGE },
            borders: DOCX_NO_BORDERS,
            margins: { right: 180, top: 40, bottom: 40 },
            verticalAlign: VerticalAlign.TOP,
            children: leftRows.map(docxRow),
          }),
          new TableCell({
            width: { size: 43, type: WidthType.PERCENTAGE },
            borders: DOCX_NO_BORDERS,
            margins: { left: 180, top: 40, bottom: 40 },
            verticalAlign: VerticalAlign.TOP,
            children: rightRows.map(docxRow),
          }),
        ],
      }),
    ],
  });
}

function docxHeaderTable(logoData) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: DOCX_NO_BORDERS,
    rows: [
      new TableRow({
        children: [
          // Logo
          new TableCell({
            width: { size: 12, type: WidthType.PERCENTAGE },
            borders: DOCX_NO_BORDERS,
            verticalAlign: VerticalAlign.CENTER,
            margins: { right: 100 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new ImageRun({ data: logoData, type: "png", transformation: { width: 62, height: 62 } })],
              }),
            ],
          }),
          // Agency name (center)
          new TableCell({
            width: { size: 58, type: WidthType.PERCENTAGE },
            borders: DOCX_NO_BORDERS,
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 16 },
                children: [new TextRun({ text: "Republic of the Philippines", size: 18, italics: true, color: "445566" })],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 16 },
                children: [new TextRun({ text: "PHILIPPINE STATISTICS AUTHORITY", bold: true, size: 26, color: "0d2346" })],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: "Regional Statistical Services Office No. II", size: 19, color: "445566" })],
              }),
            ],
          }),
          // PAY SLIP label
          new TableCell({
            width: { size: 30, type: WidthType.PERCENTAGE },
            borders: { ...DOCX_NO_BORDERS, left: { style: BorderStyle.SINGLE, size: 6, color: "0d2346" } },
            verticalAlign: VerticalAlign.CENTER,
            margins: { left: 200 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 20 },
                children: [new TextRun({ text: "PAY SLIP", bold: true, size: 46, color: "0d2346" })],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

function docxDividerLine() {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 10, color: "0d2346" } },
    children: [],
  });
}

function docxNameDateParagraph(item, period) {
  return new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: 8800 }],
    spacing: { before: 30, after: 30 },
    children: [
      new TextRun({ text: "Employee Name: ", size: 20, color: "445566" }),
      new TextRun({ text: item.employee_name?.toUpperCase() ?? "", bold: true, size: 20, color: "0d2346" }),
      new TextRun({ text: `\tPeriod: `, size: 20, color: "445566" }),
      new TextRun({ text: formatPeriodRange(period), bold: true, size: 20, color: "0d2346" }),
    ],
  });
}

function docxSignaturesTable(signatories) {
  const prepName  = signatories?.prepared_by_name  || "";
  const prepTitle = signatories?.prepared_by_title || "";
  const certName  = signatories?.certified_by_name  || "";
  const certTitle = signatories?.certified_by_title || "";

  const sigCell = (label, name, title) => [
    new Paragraph({ spacing: { before: 120 }, children: [new TextRun({ text: label, size: 19, color: "445566" })] }),
    new Paragraph({ spacing: { before: 120 }, children: [new TextRun({ text: "", size: 20 })] }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 60 },
      children: [new TextRun({ text: name, bold: true, size: 20, color: "0d2346" })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: title, size: 18, color: "445566", italics: true })],
    }),
  ];

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: DOCX_NO_BORDERS,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            borders: DOCX_NO_BORDERS,
            margins: { right: 180 },
            children: sigCell("Prepared by:", prepName, prepTitle),
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            borders: DOCX_NO_BORDERS,
            margins: { left: 180 },
            children: sigCell("Certified by:", certName, certTitle),
          }),
        ],
      }),
    ],
  });
}

function docxSummaryRow(item) {
  const sumCell = (label, value) => new TableCell({
    width: { size: 36, type: WidthType.PERCENTAGE },
    borders: DOCX_NO_BORDERS,
    margins: { top: 50, bottom: 50, left: 0, right: 0 },
    children: [
      new Paragraph({ spacing: { after: 10 }, children: [new TextRun({ text: label, size: 15, color: "445566" })] }),
      new Paragraph({ children: [new TextRun({ text: `₱ ${fmt(Number(value ?? 0))}`, bold: true, size: 19, color: "0d2346" })] }),
    ],
  });
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: { ...DOCX_NO_BORDERS, top: { style: BorderStyle.SINGLE, size: 8, color: "0d2346" } },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 28, type: WidthType.PERCENTAGE },
            borders: DOCX_NO_BORDERS,
            margins: { top: 50, bottom: 50, left: 0, right: 40 },
            children: [
              new Paragraph({ spacing: { after: 10 }, children: [new TextRun({ text: "Take Home Pay", size: 15, color: "445566" })] }),
              new Paragraph({ children: [new TextRun({ text: `₱ ${fmt(Number(item.net_pay ?? 0))}`, bold: true, size: 23, color: "166534" })] }),
            ],
          }),
          sumCell("1st Quincena", item.first_quincena),
          sumCell("2nd Quincena", item.second_quincena),
        ],
      }),
    ],
  });
}

function docxCutLine() {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 60 },
    children: [
      new TextRun({ text: "✂  —  —  —  —  —  —  —  —  —  —  —  CUT HERE  —  —  —  —  —  —  —  —  —  —  —  ✂", size: 16, color: "8899aa", italics: true }),
    ],
  });
}

function buildPayslipCell(item, period, logoData, signatories, lessDedCatalog = LESS_DEDUCTIONS_CATALOG, dedCatalog = DEDUCTIONS_CATALOG, earnTypes = []) {
  const ded      = item.deductions ?? {};
  const basicPay = Number(item.basic_pay) || 0;
  const pera     = Number(item.allowances?.pera ?? PERA_AMOUNT);
  const peraActive = earnTypes.length === 0 || earnTypes.some(t => t.name === 'pera');

  const customEarnRows = earnTypes
    .filter(t => t.name !== 'basic_pay' && t.name !== 'pera')
    .map(t => ({
      label: t.display_name || t.name,
      value: t.computation_mode === 'fixed' && t.amount_fixed != null
        ? Number(t.amount_fixed)
        : Number(item.allowances?.[t.name] ?? 0),
    }));

  const lessDedTotal          = sumCatalogComputed(lessDedCatalog, ded, basicPay) + sumCatalogTyped(lessDedCatalog, ded);
  const deductionsColumnTotal = sumCatalogComputed(dedCatalog, ded, basicPay) + sumCatalogTyped(dedCatalog, ded);
  const totalSalaryEarned     = Number(item.gross_pay) - lessDedTotal;

  const leftRows = [
    { label: "Gross Earnings", header: true },
    { label: "Basic Salary", value: basicPay },
    ...(peraActive ? [{ label: "PERA", value: pera }] : []),
    ...customEarnRows,
    { label: "Total Gross", value: Number(item.gross_pay), bold: true },
    { label: "Less Deductions", header: true },
    ...catalogRows(lessDedCatalog, ded, basicPay),
    { label: "Total Salary/Earned for the Period", value: totalSalaryEarned, bold: true },
  ];
  const rightRows = [
    { label: "Deductions", header: true },
    ...catalogRows(dedCatalog, ded, basicPay),
    { label: "Total Deductions", value: deductionsColumnTotal, bold: true },
  ];

  return new TableCell({
    borders: DOCX_NO_BORDERS,
    margins: { top: 60, bottom: 60, left: 180, right: 180 },
    children: [
      docxHeaderTable(logoData),
      docxNameDateParagraph(item, period),
      docxTwoColumnTable(leftRows, rightRows),
      docxSummaryRow(item),
      docxSignaturesTable(signatories),
    ],
  });
}

// Both payslips of a pair go in one 3-row table so Word never splits them across pages
function buildPageTable(item1, item2, period, logoData, signatories, lessDedCatalog = LESS_DEDUCTIONS_CATALOG, dedCatalog = DEDUCTIONS_CATALOG, earnTypes = []) {
  const rows = [
    new TableRow({
      height: { value: 9440, rule: HeightRule.EXACT },
      cantSplit: true,
      children: [buildPayslipCell(item1, period, logoData, signatories, lessDedCatalog, dedCatalog, earnTypes)],
    }),
  ];
  if (item2) {
    rows.push(
      new TableRow({
        height: { value: 320, rule: HeightRule.EXACT },
        cantSplit: true,
        children: [new TableCell({
          borders: DOCX_NO_BORDERS,
          margins: { top: 0, bottom: 0, left: 0, right: 0 },
          verticalAlign: VerticalAlign.CENTER,
          children: [docxCutLine()],
        })],
      }),
      new TableRow({
        height: { value: 9440, rule: HeightRule.EXACT },
        cantSplit: true,
        children: [buildPayslipCell(item2, period, logoData, signatories, lessDedCatalog, dedCatalog, earnTypes)],
      }),
    );
  }
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: DOCX_NO_BORDERS,
    rows,
  });
}

async function generatePayslipsDocx(items, period, signatories, lessDedCatalog = LESS_DEDUCTIONS_CATALOG, dedCatalog = DEDUCTIONS_CATALOG, earnTypes = []) {
  const logoData = await fetch(psaLogo).then(r => r.arrayBuffer());

  const pageProps = {
    size: { width: 12240, height: 20160 },
    margin: { top: 420, bottom: 420, left: 560, right: 560 },
  };

  const sections = [];
  for (let i = 0; i < items.length; i += 2) {
    sections.push({
      properties: {
        ...(i > 0 ? { type: SectionType.NEXT_PAGE } : {}),
        page: pageProps,
      },
      children: [buildPageTable(items[i], items[i + 1] ?? null, period, logoData, signatories, lessDedCatalog, dedCatalog, earnTypes)],
    });
  }

  const doc = new Document({
    styles: { default: { document: { run: { font: "Arial" } } } },
    sections,
  });
  return Packer.toBlob(doc);
}

// Column layout follows the office's Appendix 33 General Payroll format (DBM/COA). Per office
// request (2026-07-09), Regular and Contractual employees are split into separate sheet tabs —
// same template/layout repeated per employment type — rather than one combined sheet.
// Provident Fund Loans and both PAG-IBIG columns are single combined values with their
// component names as a wrapped note in the header. GSIS Emerg combines the two legacy
// emergency-loan fields into one. Optional Insurance, GFAL, UOLI Loan, and A/P for Last Payment
// are newly tracked fields (added to the schema alongside this export). Educational Asst., Cash
// Advance, and Medic 2 were dropped from the office's hand-edited version but are kept here since
// they're still real tracked data — remove them if unwanted.
//
// Every amount below is read from the same computed fields the Pay Slip screen already shows
// (item.basic_pay, item.gross_pay, item.deductions, etc.) — nothing here recomputes payroll math
// independently, so the export can never disagree with what's on screen.
const ONES = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten",
  "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"];
const TENS = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];

function numberToWords(num) {
  if (num === 0) return "Zero";
  function chunk(n) {
    let str = "";
    if (n >= 100) { str += `${ONES[Math.floor(n / 100)]} Hundred `; n %= 100; }
    if (n >= 20)  { str += `${TENS[Math.floor(n / 10)]} `; n %= 10; }
    if (n > 0)    str += `${ONES[n]} `;
    return str.trim();
  }
  const groups = ["", " Thousand", " Million", " Billion"];
  let words = "", i = 0, n = Math.floor(num);
  while (n > 0) {
    const part = n % 1000;
    if (part > 0) words = `${chunk(part)}${groups[i]} ${words}`;
    n = Math.floor(n / 1000);
    i++;
  }
  return words.trim();
}

function amountInWords(amount) {
  const pesos    = Math.floor(amount);
  const centavos = Math.round((amount - pesos) * 100);
  const pesosWords = pesos === 0 ? "Zero" : numberToWords(pesos);
  return `${pesosWords} Peso${pesos === 1 ? "" : "s"} and ${String(centavos).padStart(2, "0")}/100`;
}

// Journal Entry account codes, confirmed against the office's own Appendix 33 workbook.
const JOURNAL_ENTRY_CODES = [
  { code: "701", label: "Basic Salary",           cols: ["__basic"] },
  { code: "711", label: "PERA",                   cols: ["__pera"] },
  { code: "413", label: "GSIS",                   cols: ["gsis_life_ret", "gsis_consolidated_loan", "gsis_educational", "gsis_cash_advance", "optl_ins", "gfal", "gsis_rel_loan", "uoli_loan", "gsis_policy_loan", "gsis_mpl_lite", "gsis_mpl", "gsis_cpl", "gsis_emergency_loan"] },
  // Confirmed with office (2026-07-08): Union Dues stays folded into 439 alongside Provident Fund.
  { code: "414", label: "PAG-IBIG / HDMF",        cols: ["pagibig_premium", "pagibig_mp2", "pagibig_mpl_calamity"] },
  { code: "415", label: "Medic",                  cols: ["medic", "medic2"] },
  { code: "412", label: "W/Hold Tax",             cols: ["withholding_tax"] },
  { code: "439", label: "Provident Fund / Other", cols: ["provident_premium", "union_dues", "provident_gpl", "provident_eml", "provident_slmv", "cemco", "lbp_salary_loan", "ap_last_payment"] },
  { code: "403", label: "Take-Home / Net Pay",    cols: ["__takehome"] },
];

async function generatePayslipsXlsx(items, period, signatories, lessDedCatalog = LESS_DEDUCTIONS_CATALOG, dedCatalog = DEDUCTIONS_CATALOG, earnTypes = []) {
  // Fixed employee-info columns (never change)
  const peraActive = earnTypes.length === 0 || earnTypes.some(t => t.name === 'pera');
  const customEarnCols = earnTypes
    .filter(t => t.name !== 'basic_pay' && t.name !== 'pera')
    .map(t => ({ label: t.display_name || t.name, col: `__earn_${t.name}` }));
  const fixedPreCols = [
    { label: "No.",                   col: "__no" },
    { label: "Name of Employee",      col: "__name" },
    { label: "Position",              col: "__position" },
    { label: "Employee\nNumber",      col: "__empnum" },
    { label: "Rate Per\nMonth",       col: "__basic" },
    ...(peraActive ? [{ label: "PERA", col: "__pera" }] : []),
    ...customEarnCols,
  ];
  // Known less-deduction columns with their short Excel labels
  const hardcodedLessDedCols = [
    { label: "Prov. Fund\nPrem.",     col: "provident_premium" },
    { label: "Union\nDues",           col: "union_dues" },
    { label: "GPL",   col: "provident_gpl" },
    { label: "EML",   col: "provident_eml" },
    { label: "SLMV",  col: "provident_slmv" },
    { label: "CEMCO",                 col: "cemco" },
  ];
  // Any less-deduction type added through PayrollTypes settings that isn't already hardcoded
  const knownLessColSet = new Set(hardcodedLessDedCols.map(c => c.col));
  const dynamicLessDedCols = flattenDeductionCatalog(lessDedCatalog)
    .filter(s => !knownLessColSet.has(s.col))
    .map(s => ({ label: s.label, col: s.col }));
  // Fixed mid columns (totals computed from less-deductions)
  const fixedMidCols = [
    { label: "Total\nSalary",         col: "__totalSalary" },
    { label: "Earned for\nthe Period",col: "__earned" },
  ];
  // Short Excel labels for known deduction column keys
  const EXCEL_KNOWN_LABELS = {
    gsis_life_ret:          "Life &\nRet",
    gsis_consolidated_loan: "Consol.\nLoan",
    gsis_educational:       "Educ.\nAsst.",
    gsis_cash_advance:      "Cash\nAdvance",
    pagibig_premium:        "REG",
    pagibig_mp2:            "MP2",
    pagibig_mpl_calamity:   "MPL-Cal.",
    optl_ins:               "OPTL\nINS",
    gfal:                   "GFAL",
    gsis_rel_loan:          "REL\nLoan",
    uoli_loan:              "UOLI\nLoan",
    gsis_policy_loan:       "PL Reg",
    gsis_mpl_lite:          "MPL\nLite",
    gsis_mpl:               "MPL",
    gsis_cpl:               "CPL",
    gsis_emergency_loan:    "GSIS\nEmerg.",
    medic:                  "Medic",
    medic2:                 "Medic 2",
    withholding_tax:        "W/Hold\nTax",
    lbp_salary_loan:        "LBP Sal.\nLoan",
    ap_last_payment:        "A/P Last\nPay.",
  };
  // Build deduction column specs from dedCatalog
  const dedColSpecs = [];
  const dynamicHeaderGroups = [];
  for (const g of dedCatalog) {
    const isPickerGroup = g.items && g.items.length >= 2;
    if (isPickerGroup) {
      // Checked items go under a group header (like PAG-IBIG); unchecked are standalone
      const groupedSpecs = [];
      const standaloneSpecs = [];
      if (g.computed) {
        const spec = { label: EXCEL_KNOWN_LABELS[g.computed.col] ?? g.computed.label, col: g.computed.col };
        (g.computed.merge_in_excel ? groupedSpecs : standaloneSpecs).push(spec);
      }
      for (const item of g.items) {
        const spec = { label: EXCEL_KNOWN_LABELS[item.col] ?? item.label, col: item.col };
        (item.merge_in_excel ? groupedSpecs : standaloneSpecs).push(spec);
      }
      // Grouped items first (consecutive in leafCols so the header spans only them)
      for (const s of groupedSpecs) dedColSpecs.push(s);
      for (const s of standaloneSpecs) dedColSpecs.push(s);
      if (groupedSpecs.length >= 2) {
        dynamicHeaderGroups.push({ label: g.group, cols: groupedSpecs.map(s => s.col) });
      }
    } else {
      // Standalone or sibling row
      const col = g.col ?? g.computed?.col;
      if (col) dedColSpecs.push({ label: EXCEL_KNOWN_LABELS[col] ?? g.group, col });
    }
  }
  // Fixed tail columns
  const fixedEndCols = [
    { label: "Take-Home\nPay",        col: "__takehome" },
    { label: "1st\nQuincena",         col: "__q1" },
    { label: "2nd\nQuincena",         col: "__q2" },
    { label: "Signature",             col: "__sig" },
  ];
  const leafCols = [
    ...fixedPreCols,
    ...hardcodedLessDedCols, ...dynamicLessDedCols,
    ...fixedMidCols,
    ...dedColSpecs,
    ...fixedEndCols,
  ];
  const nonNumericCols = new Set(["__no", "__name", "__position", "__empnum", "__sig"]);

  // PAG-IBIG Premium is a "computed" catalog field (2% of basic pay unless typed) — same rate
  // constant the Pay Slip screen and .docx export already read via sumCatalogComputed, reused
  // here so the split-out column matches what the rest of the app shows for the same employee.
  const pagibigGroup = dedCatalog.find(g => g.group === "PAG-IBIG");
  const pagibigPremiumComputed = pagibigGroup?.computed ?? DEDUCTIONS_CATALOG.find(g => g.group === "PAG-IBIG").computed;

  const workbook = new ExcelJS.Workbook();
  let logoImageId = null;
  try {
    const logoBlob = await fetch(psaLogo).then(r => r.blob());
    const psaLogoB64 = await new Promise(resolve => {
      const fr = new FileReader();
      fr.onload = () => resolve(fr.result.split(",")[1]);
      fr.readAsDataURL(logoBlob);
    });
    logoImageId = workbook.addImage({ base64: psaLogoB64, extension: "png" });
  } catch { /* ignore if logo fails */ }
  let bagongImageId = null;
  try {
    bagongImageId = workbook.addImage({ base64: BAGONG_PILIPINAS_B64, extension: "png" });
  } catch { /* ignore */ }
  // Build tab list from actual data so "Permanent" and other types get their own sheet.
  // Standard types come first, then any extra types found in the data.
  const typesInData = [...new Set(items.map(item => item.employment_type || "Regular"))];
  const orderedTypes = [
    ...EMPLOYMENT_TYPES.filter(t => typesInData.includes(t)),
    ...typesInData.filter(t => !EMPLOYMENT_TYPES.includes(t)),
  ];
  for (const empType of orderedTypes) {
    const groupItems = items
      .filter(item => (item.employment_type || "Regular") === empType)
      .sort((a, b) => {
        const divA = (a.office_division || "").toLowerCase();
        const divB = (b.office_division || "").toLowerCase();
        if (divA !== divB) return divA.localeCompare(divB);
        const posA = (a.position || "").toLowerCase();
        const posB = (b.position || "").toLowerCase();
        return posA.localeCompare(posB);
      });
    buildPayrollWorksheet(workbook, empType, groupItems, period, signatories, leafCols, nonNumericCols, pagibigPremiumComputed, logoImageId, bagongImageId, lessDedCatalog, dedCatalog, earnTypes, dynamicHeaderGroups);
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

function buildPayrollWorksheet(workbook, empType, items, period, signatories, leafCols, nonNumericCols, pagibigPremiumComputed, logoImageId, bagongImageId, lessDedCatalog = LESS_DEDUCTIONS_CATALOG, dedCatalog = DEDUCTIONS_CATALOG, earnTypes = [], dynamicHeaderGroups = []) {
  const customEarnTypes = earnTypes.filter(t => t.name !== 'basic_pay' && t.name !== 'pera');
  const rows = items.map((item, idx) => {
    const ded      = item.deductions ?? {};
    const basicPay = Number(item.basic_pay ?? 0);
    const pera     = Number(item.allowances?.pera ?? 0);
    const grossPay = Number(item.gross_pay ?? 0);
    const lessDedTotal      = sumCatalogComputed(lessDedCatalog, ded, basicPay) + sumCatalogTyped(lessDedCatalog, ded);
    const dedColumnTotal    = sumCatalogComputed(dedCatalog, ded, basicPay) + sumCatalogTyped(dedCatalog, ded);
    const earnedForPeriod   = roundAmt(grossPay - lessDedTotal);
    const takeHome          = roundAmt(earnedForPeriod - dedColumnTotal);
    const pagibigPremiumOverride = ded[pagibigPremiumComputed.col];
    const pagibigPremiumAmt = pagibigPremiumOverride != null
      ? Number(pagibigPremiumOverride)
      : roundAmt(basicPay * pagibigPremiumComputed.rate);

    const special = {
      __no: idx + 1,
      __name: (item.employee_name ?? "").toUpperCase(),
      __position: item.position ?? "",
      __empnum: item.employee_number ?? "",
      __basic: basicPay,
      __pera: pera,
      ...Object.fromEntries(customEarnTypes.map(t => [`__earn_${t.name}`, Number(item.allowances?.[t.name] ?? 0)])),
      __gross: grossPay,
      __totalSalary: earnedForPeriod,
      __earned: earnedForPeriod,
      __takehome: takeHome,
      __q1: Number(item.first_quincena ?? 0),
      __q2: Number(item.second_quincena ?? 0),
      __sig: "",
      pagibig_premium: pagibigPremiumAmt,
    };

    return leafCols.map(spec => spec.col in special ? special[spec.col] : Number(ded[spec.col] ?? 0));
  });

  const totals = leafCols.map((spec, i) => {
    if (spec.col === "__no" || spec.col === "__sig" || spec.col === "__position" || spec.col === "__empnum") return "";
    if (spec.col === "__name") return "TOTAL";
    return roundAmt(rows.reduce((sum, r) => sum + r[i], 0));
  });

  const period_label = formatPeriodRange(period);
  const payrollNo = period?.period_end
    ? `${new Date(period.period_end).getFullYear()}-${String(new Date(period.period_end).getMonth() + 1).padStart(2, "0")}`
    : "";

  const ws = workbook.addWorksheet(`${empType} Payroll`);
  const lastCol = leafCols.length;

  const colWidths = { __no: 5, __name: 28, __position: 20, __empnum: 15, __sig: 20, __basic: 16, __pera: 12, provident_premium: 10, union_dues: 10 };
  ws.columns = leafCols.map(spec => ({
    width: colWidths[spec.col] ?? 13,
    style: { font: { name: 'Arial', size: 10 } },
  }));

  // Print setup — too wide for portrait at this column count; scale to fit page width and
  // repeat the header row on every printed page.
  ws.pageSetup.orientation  = "landscape";
  ws.pageSetup.fitToPage    = true;
  ws.pageSetup.fitToWidth   = 1;
  ws.pageSetup.fitToHeight  = 0;
  ws.pageSetup.margins      = { left: 0.3, right: 0.3, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 };
  ws.views = [{ showGridLines: false }];

  const centerMergedRow = (values) => {
    const row = ws.addRow(values);
    ws.mergeCells(row.number, 1, row.number, lastCol);
    ws.getCell(row.number, 1).alignment = { horizontal: "center" };
    return row;
  };

  // ── LETTERHEAD ─────────────────────────────────────────────────────────────
  const hRow1 = ws.addRow([]); hRow1.height = 36;
  const hRow2 = ws.addRow([]); hRow2.height = 36;
  const hRow3 = ws.addRow([]); hRow3.height = 30;

  // psa-logo.png seal (960×930px, ~1:1 ratio) — 120×120px, top-left overlay
  if (logoImageId !== null) {
    ws.addImage(logoImageId, { tl: { col: 0, row: 0 }, ext: { width: 120, height: 120 }, editAs: "oneCell" });
  }
  // Bagong Pilipinas logo — top-right overlay, 130×130px, shifted left so Appendix 33 has room
  if (bagongImageId !== null) {
    ws.addImage(bagongImageId, { tl: { col: lastCol - 5, row: 0 }, ext: { width: 130, height: 130 }, editAs: "oneCell" });
  }

  // PAYROLL title — centered across full sheet width
  ws.mergeCells(hRow1.number, 1, hRow2.number, lastCol);
  ws.getCell(hRow1.number, 1).value = "PAYROLL";
  ws.getCell(hRow1.number, 1).alignment = { horizontal: "center", vertical: "middle" };
  ws.getCell(hRow1.number, 1).font = { bold: true, size: 20 };

  ws.mergeCells(hRow3.number, 1, hRow3.number, lastCol);
  ws.getCell(hRow3.number, 1).value = `For the period: ${period_label}`;
  ws.getCell(hRow3.number, 1).alignment = { horizontal: "center", vertical: "middle" };
  ws.getCell(hRow3.number, 1).font = { bold: true, size: 11 };

  // Appendix 33 + General Payroll subtitle
  const appendixRow = ws.addRow([]);
  ws.getCell(appendixRow.number, lastCol).value = "Appendix 33";
  ws.getCell(appendixRow.number, lastCol).alignment = { horizontal: "right" };
  ws.getCell(appendixRow.number, lastCol).font = { bold: true };
  const gpRow = centerMergedRow([`GENERAL PAYROLL — ${empType.toUpperCase()} EMPLOYEES`]);
  gpRow.getCell(1).font = { bold: true };

  ws.addRow([]);

  const entityRow = ws.addRow(["Entity Name: Philippine Statistics Authority"]);
  entityRow.getCell(1).font = { bold: true };
  ws.getCell(entityRow.number, lastCol - 1).value = `Payroll No.: ${payrollNo}`;
  ws.getCell(entityRow.number, lastCol - 1).font = { bold: true };
  const fundRow = ws.addRow(["Fund Cluster: "]);
  fundRow.getCell(1).font = { bold: true };
  ws.getCell(fundRow.number, lastCol - 1).value = "Sheet ______ of ________ Sheets";
  ws.getCell(fundRow.number, lastCol - 1).font = { bold: true };
  const adaRow = ws.addRow([]);
  ws.getCell(adaRow.number, lastCol - 1).value = "ADA No.: ________________________";
  ws.getCell(adaRow.number, lastCol - 1).font = { bold: true };
  ws.addRow([]);

  centerMergedRow([`We acknowledge receipt of the sums shown opposite our names as full compensation for our services rendered for the period: ${period_label}`]);
  const generatedAt = new Date();
  centerMergedRow([`Date: ${generatedAt.toLocaleDateString("en-PH", { year: "numeric", month: "long", day: "numeric" })} ${generatedAt.toLocaleTimeString("en-PH", { hour: "2-digit", minute: "2-digit" })}`]);
  ws.addRow([]);

  const gridBorder = {
    top: { style: "thin", color: { argb: "FF999999" } },
    bottom: { style: "thin", color: { argb: "FF999999" } },
    left: { style: "thin", color: { argb: "FF999999" } },
    right: { style: "thin", color: { argb: "FF999999" } },
  };

  // 2-row header: group label merged across sub-columns on top row,
  // individual sub-labels on bottom row; standalone columns span both rows.
  const HEADER_GROUPS = [
    { label: "PROVIDENT FUND LOANS", cols: ["provident_gpl", "provident_eml", "provident_slmv"] },
    { label: "PAG-IBIG",             cols: ["pagibig_premium", "pagibig_mp2", "pagibig_mpl_calamity"] },
    ...dynamicHeaderGroups,
  ];
  const groupColSet = new Set(HEADER_GROUPS.flatMap(g => g.cols));

  const hGrpRow = ws.addRow([]);  hGrpRow.height = 28;
  const hLeafRow = ws.addRow([]); hLeafRow.height = 28;

  leafCols.forEach((spec, idx) => {
    const col = idx + 1;
    if (groupColSet.has(spec.col)) {
      hLeafRow.getCell(col).value = spec.label;
    } else {
      hGrpRow.getCell(col).value = spec.label;
      ws.mergeCells(hGrpRow.number, col, hLeafRow.number, col);
    }
  });
  for (const g of HEADER_GROUPS) {
    const si = leafCols.findIndex(s => s.col === g.cols[0]) + 1;
    const ei = leafCols.findIndex(s => s.col === g.cols[g.cols.length - 1]) + 1;
    if (si > 0 && ei > 0) {
      ws.mergeCells(hGrpRow.number, si, hGrpRow.number, ei);
      const c = ws.getCell(hGrpRow.number, si);
      c.value = g.label;
      c.font = { bold: true };
      c.alignment = { vertical: "middle", horizontal: "center" };
      c.border = gridBorder;
    }
  }
  for (let col = 1; col <= lastCol; col++) {
    const spec = leafCols[col - 1];
    const gc = ws.getCell(hGrpRow.number, col);
    gc.font = { bold: true };
    gc.border = gridBorder;
    gc.alignment = { wrapText: true, vertical: "middle", horizontal: "center" };
    if (groupColSet.has(spec.col)) {
      const lc = ws.getCell(hLeafRow.number, col);
      lc.font = { bold: true };
      lc.border = gridBorder;
      lc.alignment = { vertical: "middle", horizontal: "center" };
    }
  }
  ws.pageSetup.printTitlesRow = `${hGrpRow.number}:${hLeafRow.number}`;

  // Helpers for formula injection
  const cl = (n) => { let s = ''; while (n > 0) { const r = (n-1)%26; s = String.fromCharCode(65+r)+s; n = Math.floor((n-1)/26); } return s; };
  const colIdxMap = Object.fromEntries(leafCols.map((spec, i) => [spec.col, i + 1]));
  // cols that are auto-derived (rate × basic pay) — from both catalogs
  const computedRateCols = {};
  for (const cat of [lessDedCatalog, dedCatalog]) {
    for (const g of cat) { if (g.computed) computedRateCols[g.computed.col] = g.computed.rate; }
  }
  const lessDedColKeys = flattenDeductionCatalog(lessDedCatalog).map(s => s.col);
  // dedColKeys: all cols between __earned and __takehome (handles merged group cols too)
  const earnedIdx  = leafCols.findIndex(s => s.col === '__earned');
  const thIdx      = leafCols.findIndex(s => s.col === '__takehome');
  const dedColKeys = (earnedIdx >= 0 && thIdx > earnedIdx)
    ? leafCols.slice(earnedIdx + 1, thIdx).map(s => s.col)
    : flattenDeductionCatalog(dedCatalog).map(s => s.col);

  let firstDataRowNum = null;
  let lastDataRowNum  = null;
  for (let ri = 0; ri < rows.length; ri++) {
    const row = ws.addRow(rows[ri]);
    const rn  = row.number;
    if (firstDataRowNum === null) firstDataRowNum = rn;
    lastDataRowNum = rn;
    row.eachCell((cell, c) => {
      if (!nonNumericCols.has(leafCols[c - 1].col)) cell.numFmt = "#,##0.00";
      cell.border = gridBorder;
    });

    const ded       = items[ri].deductions ?? {};
    const basicRef       = `${cl(colIdxMap['__basic'])}${rn}`;
    const peraRef        = colIdxMap['__pera'] ? `${cl(colIdxMap['__pera'])}${rn}` : null;
    const customEarnRefs = customEarnTypes
      .map(t => colIdxMap[`__earn_${t.name}`])
      .filter(Boolean)
      .map(ci => `${cl(ci)}${rn}`);
    const grossRef       = [basicRef, ...(peraRef ? [peraRef] : []), ...customEarnRefs].join('+');
    const totalSalaryRef = `${cl(colIdxMap['__totalSalary'])}${rn}`;
    const earnedRef      = `${cl(colIdxMap['__earned'])}${rn}`;

    // Rate-based computed deductions: show formula unless user typed a DIFFERENT value
    const itemBasicPay = Number(items[ri].basic_pay ?? 0);
    for (const [colKey, rate] of Object.entries(computedRateCols)) {
      const ci = colIdxMap[colKey]; if (!ci) continue;
      const storedVal   = ded[colKey];
      const computedVal = roundAmt(itemBasicPay * rate);
      // Skip formula only when the stored value genuinely differs (real user override)
      if (storedVal != null && Math.abs(Number(storedVal) - computedVal) > 0.01) continue;
      const cell = ws.getCell(rn, ci);
      cell.value = { formula: `=ROUND(${basicRef}*${rate},2)`, result: Math.round(itemBasicPay * rate * 100) / 100 };
    }

    // Total Salary = Gross − SUM(less-deduction cols) [same formula as Earned]
    if (colIdxMap['__totalSalary']) {
      const refs = lessDedColKeys.filter(k => colIdxMap[k]).map(k => `${cl(colIdxMap[k])}${rn}`);
      const cell = ws.getCell(rn, colIdxMap['__totalSalary']);
      cell.value = { formula: `=${grossRef}-SUM(${refs.join(',')})`, result: Number(cell.value) };
    }

    // Earned for Period = SUM(all right-side deductions + Take-Home) — right-side verification total
    if (colIdxMap['__earned'] && colIdxMap['__takehome']) {
      const firstDedKey = dedColKeys.find(k => colIdxMap[k]);
      if (firstDedKey) {
        const rangeStart = `${cl(colIdxMap[firstDedKey])}${rn}`;
        const rangeEnd   = `${cl(colIdxMap['__takehome'])}${rn}`;
        const cell = ws.getCell(rn, colIdxMap['__earned']);
        cell.value = { formula: `=SUM(${rangeStart}:${rangeEnd})`, result: Number(cell.value) };
      }
    }

    // Take-Home Pay = Total Salary − SUM(deduction cols)
    if (colIdxMap['__takehome']) {
      const refs = dedColKeys.filter(k => colIdxMap[k]).map(k => `${cl(colIdxMap[k])}${rn}`);
      const cell = ws.getCell(rn, colIdxMap['__takehome']);
      cell.value = { formula: `=${totalSalaryRef}-SUM(${refs.join(',')})`, result: Number(cell.value) };
    }

    // Only inject Excel formulas when Q1 has been explicitly stored (HR override).
    // Never fabricate Q1 as net_pay/2 — show 0 if the field is absent.
    if (colIdxMap['__q1'] && colIdxMap['__q2'] && colIdxMap['__takehome']) {
      const storedQ1 = items[ri].first_quincena;
      if (storedQ1 != null) {
        const thRef = `${cl(colIdxMap['__takehome'])}${rn}`;
        const q1Ref = `${cl(colIdxMap['__q1'])}${rn}`;
        const q1Cell = ws.getCell(rn, colIdxMap['__q1']);
        q1Cell.value = { formula: `=ROUND(${thRef}/2,2)`, result: Number(q1Cell.value) };
        const q2Cell = ws.getCell(rn, colIdxMap['__q2']);
        q2Cell.value = { formula: `=${thRef}-${q1Ref}`, result: Number(q2Cell.value) };
      } else {
        ws.getCell(rn, colIdxMap['__q1']).value = 0;
        ws.getCell(rn, colIdxMap['__q2']).value = 0;
      }
    }
  }

  const totalsRow = ws.addRow(totals);
  totalsRow.eachCell((cell, c) => {
    const spec = leafCols[c - 1];
    cell.font = { bold: true };
    if (!nonNumericCols.has(spec.col)) {
      cell.numFmt = "#,##0.00";
      // SUM formula over all data rows for each numeric column
      if (firstDataRowNum) {
        const colL = cl(c);
        cell.value = { formula: `=SUM(${colL}${firstDataRowNum}:${colL}${lastDataRowNum})`, result: Number(cell.value) || 0 };
      }
    }
    cell.border = gridBorder;
  });

  // Group totals row: blank except for merged combined totals for each HEADER_GROUP.
  // All cells get grid borders so the row looks complete; only group cells have values.
  const groupTotalsRow = ws.addRow([]);
  leafCols.forEach((_, idx) => {
    groupTotalsRow.getCell(idx + 1).border = gridBorder;
  });
  for (const g of HEADER_GROUPS) {
    const si = leafCols.findIndex(c => c.col === g.cols[0]) + 1;
    const ei = leafCols.findIndex(c => c.col === g.cols[g.cols.length - 1]) + 1;
    if (si > 0 && ei > 0) {
      const combined = roundAmt(g.cols.reduce((s, col) => {
        const i = leafCols.findIndex(c => c.col === col);
        return s + (i >= 0 ? Number(totals[i]) || 0 : 0);
      }, 0));
      ws.mergeCells(groupTotalsRow.number, si, groupTotalsRow.number, ei);
      const cell = ws.getCell(groupTotalsRow.number, si);
      cell.value = firstDataRowNum
        ? { formula: `=SUM(${cl(si)}${firstDataRowNum}:${cl(ei)}${lastDataRowNum})`, result: combined }
        : combined;
      cell.font = { bold: true };
      cell.numFmt = "#,##0.00";
      cell.alignment = { horizontal: "center" };
      cell.border = gridBorder;
    }
  }

  const colTotal = (colName) => {
    const i = leafCols.findIndex(c => c.col === colName);
    return i === -1 ? 0 : Number(totals[i]) || 0;
  };

  // ── Certification block ───────────────────────────────────────────────────────
  const takeHomeTotal = colTotal("__takehome");

  const prepName       = signatories?.prepared_by_name   || "";
  const prepTitle      = signatories?.prepared_by_title  || "";
  const certName       = signatories?.certified_by_name  || "";
  const certTitle      = signatories?.certified_by_title || "";
  const approvingName  = signatories?.approving_name  || null;
  const approvingTitle = signatories?.approving_title || null;
  const disbursingName  = signatories?.disbursing_name  || null;
  const disbursingTitle = signatories?.disbursing_title || null;

  // Signatory block — cols capped to 24 so all sections are on-screen.
  // A|C: left (1..sigMid) and right (sigMid+1..sigEnd) with surrounding box borders.
  // B|D|E: B (1..sigMid), D (sigMid+1..dEnd), E (dEnd+1..sigEnd), each boxed.
  const sigEnd = Math.min(lastCol, 20);
  const sigMid = Math.floor(sigEnd / 2);   // ≈12
  const dEnd   = sigEnd - 4;               // ≈20
  const dw     = 3;
  const sigLine = "_______________________";
  const T = { style: 'thin' };

  // Write a merged cell with value/font/alignment/border; skip merge when c1===c2
  const sc = (rn, c1, c2, val, opts = {}) => {
    if (c1 > c2 || c1 > sigEnd) return;
    if (c1 < c2) ws.mergeCells(rn, c1, rn, c2);
    const cell = ws.getCell(rn, c1);
    if (val !== undefined) cell.value = val ?? null;
    if (opts.font)   cell.font      = opts.font;
    if (opts.align)  cell.alignment = opts.align;
    if (opts.numFmt) cell.numFmt    = opts.numFmt;
    if (opts.border) cell.border    = opts.border;
    return cell;
  };

  // Build a border object from top/bottom/left/right booleans
  const bdr = (t, b, l, r) => ({
    ...(t ? { top: T } : {}), ...(b ? { bottom: T } : {}),
    ...(l ? { left: T } : {}), ...(r ? { right: T } : {}),
  });

  ws.addRow([]);

  // ── A | C (3 rows: header+amount | name+sig | title+Date) ───────────────────────────────────
  const acR = [];
  for (let i = 0; i < 3; i++) { const r = ws.addRow([]); r.height = [36, 20, 36][i]; acR.push(r.number); }
  const [a1, a2, a3] = acR;

  const aNameEnd = sigMid - dw;   // A name cols; sig = last 2 cols of A section
  const cNameEnd = sigEnd - dw;   // C name cols; sig = last 2 cols of C section

  // Row 1: A text | C text (most of C section) | C amount right-aligned (last 2 cols)
  sc(a1, 1, sigMid, "A. CERTIFIED. Services duly rendered as stated.",
     { align: { wrapText: true, vertical: 'top' }, border: bdr(1,0,1,0) });
  sc(a1, sigMid+1, cNameEnd, `C. APPROVED FOR PAYMENT. ${amountInWords(takeHomeTotal)}.`,
     { align: { wrapText: true, vertical: 'top' }, border: bdr(1,0,1,0) });
  sc(a1, cNameEnd+1, sigEnd, takeHomeTotal,
     { numFmt: "#,##0.00", align: { horizontal: 'right', vertical: 'bottom' }, border: bdr(1,0,0,1) });

  // Row 2: name centered+underline | sig line indented so it doesn't hit the right border
  sc(a2, 1, aNameEnd, prepName,
     { font: { bold: true, underline: true }, align: { horizontal: 'center' }, border: bdr(0,0,1,0) });
  sc(a2, aNameEnd+1, sigMid-1, null, { border: { bottom: T } });          // sig; col sigMid = right pad (no border)
  sc(a2, sigMid+1, cNameEnd, approvingName || "(Approving Official)",
     { font: { bold: true, underline: true }, align: { horizontal: 'center' }, border: bdr(0,0,1,0) });
  sc(a2, cNameEnd+1, sigEnd-1, null, { border: { bottom: T } });          // sig; col sigEnd = right pad (outer right border only)
  ws.getCell(a2, sigEnd).border = { right: T };

  // Row 3: title top-aligned for bottom breathing room | Date centered
  sc(a3, 1, aNameEnd, prepTitle, { align: { horizontal: 'center', vertical: 'top' }, border: bdr(0,1,1,0) });
  sc(a3, aNameEnd+1, sigMid-1, "Date", { align: { horizontal: 'center', vertical: 'top' }, border: bdr(0,1,0,0) }); // same cols as A sig line
  ws.getCell(a3, sigMid).border = { bottom: T };                                                                      // right-pad col: close A box bottom
  sc(a3, sigMid+1, cNameEnd, approvingTitle || "(Designation)", { align: { horizontal: 'center', vertical: 'top' }, border: bdr(0,1,1,0) });
  sc(a3, cNameEnd+1, sigEnd-1, "Date", { align: { horizontal: 'center', vertical: 'top' }, border: bdr(0,1,0,0) }); // same cols as C sig line
  ws.getCell(a3, sigEnd).border = { bottom: T, right: T };                                                            // right-pad col: close C box bottom + outer right

  // ── B | D | E (5 rows: header | 2×blank+E | name+E3 | title+Date+E4) ────────────────────────
  const bdeR = [];
  for (let i = 0; i < 5; i++) { const r = ws.addRow([]); r.height = [22, 12, 12, 18, 26][i]; bdeR.push(r.number); }
  const [b1, b2, b3, b4, b5] = bdeR;

  // Row 1: B text + compact amount box | D blank | E header
  {
    const amtW = 2;
    const amtStart = sigMid - amtW + 1;
    sc(b1, 1, amtStart-1, "B. CERTIFIED. Supporting documents complete and cash available in the amount of",
       { align: { wrapText: true, vertical: 'top' }, border: bdr(1,0,1,0) });
    if (amtStart <= sigMid) ws.mergeCells(b1, amtStart, b1, sigMid);
    const amtCell = ws.getCell(b1, amtStart);
    amtCell.value = takeHomeTotal;
    amtCell.numFmt = "#,##0.00";
    amtCell.alignment = { horizontal: 'center', vertical: 'middle' };
    amtCell.font = { underline: true };
    amtCell.border = { top: T };
    sc(b1, sigMid+1, dEnd, null, { border: bdr(1,0,1,0) });
    sc(b1, dEnd+1, sigEnd, "E.", { border: bdr(1,0,1,1) });
  }

  // Rows 2-3: blank B|D, E fields 1-2
  const eFields = ["ORS/BURS NO.: _______________", "Date: _______________", "JEV No.: _______________", "Date: _______________"];
  [b2, b3].forEach((rn, i) => {
    sc(rn, 1, sigMid, null, { border: bdr(0,0,1,0) });
    sc(rn, sigMid+1, dEnd, null, { border: bdr(0,0,1,0) });
    sc(rn, dEnd+1, sigEnd, eFields[i], { border: bdr(0,0,1,1) });
  });

  // Row 4: name+sig indented (B) | name+sig indented (D) | E field 3
  {
    const bNameEnd = sigMid - dw;
    const dNameEnd = dEnd - dw;
    sc(b4, 1, bNameEnd, certName,
       { font: { bold: true, underline: true }, align: { horizontal: 'center' }, border: bdr(0,0,1,0) });
    sc(b4, bNameEnd+1, sigMid-1, null, { border: { bottom: T } });         // sig; col sigMid = right pad
    sc(b4, sigMid+1, dNameEnd, disbursingName || "(Disbursing Officer)",
       { font: { bold: true, underline: true }, align: { horizontal: 'center' }, border: bdr(0,0,1,0) });
    sc(b4, dNameEnd+1, dEnd-1, null, { border: { bottom: T } });           // sig; col dEnd = right pad
    sc(b4, dEnd+1, sigEnd, eFields[2], { align: { vertical: 'top' }, border: bdr(0,0,1,1) });
  }

  // Row 5: title top-aligned for bottom breathing room | Date centered | closes all boxes
  {
    const bNameEnd = sigMid - dw;
    const dNameEnd = dEnd - dw;
    sc(b5, 1, bNameEnd, certTitle, { align: { horizontal: 'center', vertical: 'top' }, border: bdr(0,1,1,0) });
    sc(b5, bNameEnd+1, sigMid-1, "Date", { align: { horizontal: 'center', vertical: 'top' }, border: bdr(0,1,0,0) }); // same cols as B sig line
    ws.getCell(b5, sigMid).border = { bottom: T };                                                                      // right-pad col: close B box bottom
    sc(b5, sigMid+1, dNameEnd, disbursingTitle || "(Designation)", { align: { horizontal: 'center', vertical: 'top' }, border: bdr(0,1,1,0) });
    sc(b5, dNameEnd+1, dEnd-1, "Date", { align: { horizontal: 'center', vertical: 'top' }, border: bdr(0,1,0,0) });   // same cols as D sig line
    ws.getCell(b5, dEnd).border = { bottom: T };                                                                        // right-pad col: close D box bottom
    sc(b5, dEnd+1, sigEnd, eFields[3], { align: { vertical: 'top' }, border: bdr(0,1,1,1) });
  }

  ws.addRow([]);

  // ── Reconciliation + Journal Entry — side-by-side in the same row band:
  //    cols 1-5: Reconciliation | cols 7-11: JE 1st Quincena | cols 13-17: JE 2nd Quincena
  const gsisCols = JOURNAL_ENTRY_CODES.find(je => je.code === "413").cols;
  const gsisSubTotal = roundAmt(gsisCols.reduce((sum, c) => sum + colTotal(c), 0));

  const divisionTotals = new Map();
  for (const item of items) {
    const div = item.office_division?.trim() || "Unassigned";
    const cur = divisionTotals.get(div) ?? { q1: 0, q2: 0 };
    cur.q1 = roundAmt(cur.q1 + Number(item.first_quincena ?? 0));
    cur.q2 = roundAmt(cur.q2 + Number(item.second_quincena ?? 0));
    divisionTotals.set(div, cur);
  }

  const halfSplit = (total) => { const f = roundAmt(total / 2); return [f, roundAmt(total - f)]; };
  const jeAmountForCode = (je, half) => {
    if (je.code === "403") return half === 1 ? colTotal("__q1") : colTotal("__q2");
    return halfSplit(je.cols.reduce((s, c) => s + colTotal(c), 0))[half - 1];
  };

  // Pre-add enough rows so getCell() can fill them in
  // recon: 1+1+1+1+1+1+1+divs+1 = 8+divs rows; JE: 1+1+codes rows
  const reconRowCount = 8 + divisionTotals.size;
  const jeRowCount    = 2 + JOURNAL_ENTRY_CODES.length;
  const sideRowCount  = Math.max(reconRowCount, jeRowCount);
  const sideStartRow  = ws.lastRow.number + 1;
  for (let i = 0; i < sideRowCount; i++) ws.addRow([]);

  const wc = (rowOff, col, value, opts = {}) => {
    const cell = ws.getCell(sideStartRow + rowOff, col);
    if (value !== null && value !== undefined) cell.value = value;
    if (opts.font)      cell.font      = opts.font;
    if (opts.alignment) cell.alignment = opts.alignment;
    if (opts.numFmt)    cell.numFmt    = opts.numFmt;
  };
  const wm = (rowOff, c1, c2) => ws.mergeCells(sideStartRow + rowOff, c1, sideStartRow + rowOff, c2);

  // RECONCILIATION (cols 1-5)
  let rr = 0;
  wc(rr++, 1, "RECONCILIATION", { font: { bold: true, underline: true } });
  wc(rr++, 1, "GSIS Sub-Total Check", { font: { bold: true } });
  wc(rr, 2, "Sum of GSIS detail columns above (Life & Ret + all GSIS loan/insurance items)", { alignment: { wrapText: true } });
  wm(rr, 2, 3);  wc(rr++, 4, gsisSubTotal, { numFmt: "#,##0.00" });
  wc(rr, 2, "Ties out to Journal Entry Code 413 (GSIS)", { alignment: { wrapText: true } });
  wm(rr, 2, 3);  wc(rr++, 4, gsisSubTotal, { numFmt: "#,##0.00" });
  rr++; // blank
  wc(rr++, 1, "1st & 2nd Quincena Totals by Division", { font: { bold: true } });
  wc(rr, 2, "Division",     { font: { bold: true } });
  wc(rr, 3, "1st Quincena", { font: { bold: true } });
  wc(rr, 4, "2nd Quincena", { font: { bold: true } });
  wc(rr++, 5, "Total",      { font: { bold: true } });
  for (const [div, t] of [...divisionTotals.entries()].sort((a, b) => a[0].localeCompare(b[0]))) {
    wc(rr, 2, div);
    wc(rr, 3, t.q1,                    { numFmt: "#,##0.00" });
    wc(rr, 4, t.q2,                    { numFmt: "#,##0.00" });
    wc(rr++, 5, roundAmt(t.q1 + t.q2), { numFmt: "#,##0.00" });
  }
  wc(rr, 2, "GRAND TOTAL",                                 { font: { bold: true } });
  wc(rr, 3, colTotal("__q1"),                              { font: { bold: true }, numFmt: "#,##0.00" });
  wc(rr, 4, colTotal("__q2"),                              { font: { bold: true }, numFmt: "#,##0.00" });
  wc(rr, 5, roundAmt(colTotal("__q1") + colTotal("__q2")), { font: { bold: true }, numFmt: "#,##0.00" });

  // JOURNAL ENTRY — 1ST QUINCENA (cols 7-11)
  const je1Base = 7;
  let jr = 0;
  wc(jr++, je1Base, "JOURNAL ENTRY — 1ST QUINCENA", { font: { bold: true } });
  wc(jr, je1Base,     "JOUI",         { font: { bold: true } });
  wc(jr, je1Base + 2, "Account Code", { font: { bold: true } });
  wc(jr++, je1Base + 3, "Amount",     { font: { bold: true } });
  for (const je of JOURNAL_ENTRY_CODES) {
    wc(jr, je1Base + 1, je.label);
    wc(jr, je1Base + 2, je.code);
    wc(jr++, je1Base + 3, roundAmt(jeAmountForCode(je, 1)), { numFmt: "#,##0.00" });
  }

  // JOURNAL ENTRY — 2ND QUINCENA (cols 13-17)
  const je2Base = 13;
  let jr2 = 0;
  wc(jr2++, je2Base, "JOURNAL ENTRY — 2ND QUINCENA", { font: { bold: true } });
  wc(jr2, je2Base,     "JOUI",         { font: { bold: true } });
  wc(jr2, je2Base + 2, "Account Code", { font: { bold: true } });
  wc(jr2++, je2Base + 3, "Amount",     { font: { bold: true } });
  for (const je of JOURNAL_ENTRY_CODES) {
    wc(jr2, je2Base + 1, je.label);
    wc(jr2, je2Base + 2, je.code);
    wc(jr2++, je2Base + 3, roundAmt(jeAmountForCode(je, 2)), { numFmt: "#,##0.00" });
  }

  // Widen JE title, description, and amount cols so labels aren't cut off
  ws.getColumn(je1Base).width          = 32;
  ws.getColumn(je1Base + 1).width      = 22;
  ws.getColumn(je1Base + 2).width      = 14;
  ws.getColumn(je1Base + 3).width      = 14;
  ws.getColumn(je2Base).width          = 32;
  ws.getColumn(je2Base + 1).width      = 22;
  ws.getColumn(je2Base + 2).width      = 14;
  ws.getColumn(je2Base + 3).width      = 14;

  // ── Post-pass: Arial font on every cell + white fill (extended for logo rows) ────────────────
  const whiteFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
  ws.eachRow((row, r) => {
    const colMax = r <= 3 ? lastCol + 20 : lastCol;
    for (let c = 1; c <= colMax; c++) {
      const cell = ws.getCell(r, c);
      cell.fill = whiteFill;
      const f = cell.font ?? {};
      if (!f.name) cell.font = { name: 'Arial', ...f };
    }
  });
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

// ── CustomSelect ──────────────────────────────────────────────────────────────
function CustomSelect({ value, onChange, options, placeholder, style }) {
  const [open, setOpen] = useState(false);
  const [dropPos, setDropPos] = useState({ top: 0, left: 0, width: 0 });
  const wrapRef = useRef(null);
  const btnRef = useRef(null);

  useEffect(() => {
    function onDown(e) {
      if (wrapRef.current && !wrapRef.current.contains(e.target)) setOpen(false);
    }
    document.addEventListener("mousedown", onDown);
    return () => document.removeEventListener("mousedown", onDown);
  }, []);

  function handleOpen() {
    if (btnRef.current) {
      const rect = btnRef.current.getBoundingClientRect();
      const spaceBelow = window.innerHeight - rect.bottom;
      const dropHeight = Math.min(options.length * 41 + 2, 262);
      const openUp = spaceBelow < dropHeight && rect.top > dropHeight;
      const minW = 240;
      const w = Math.max(rect.width, minW);
      let left = rect.left;
      if (left + w > window.innerWidth - 8) left = Math.max(8, window.innerWidth - w - 8);
      setDropPos({
        top: openUp ? rect.top - dropHeight - 6 : rect.bottom + 6,
        left,
        width: w,
      });
    }
    setOpen(v => !v);
  }

  const selected = options.find(o => o.value === value);

  return (
    <div ref={wrapRef} style={{ position: "relative", ...style }}>
      <button
        ref={btnRef}
        type="button"
        onClick={handleOpen}
        style={{
          width: "100%", display: "flex", alignItems: "center",
          justifyContent: "space-between", gap: 8,
          padding: "9px 13px",
          border: open ? "1.5px solid #4a90d9" : "1.5px solid var(--border-2)",
          borderRadius: "var(--radius)", background: "var(--surface)",
          color: selected ? "var(--text)" : "#a8bcd4",
          fontSize: "0.95rem", fontWeight: 400, cursor: "pointer",
          textAlign: "left",
          boxShadow: open ? "0 0 0 3px rgba(74,144,217,0.12)" : "none",
          transition: "border-color 0.15s, box-shadow 0.15s",
        }}
      >
        <span style={{ flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
          {selected ? selected.label : placeholder}
        </span>
        <svg width="14" height="14" viewBox="0 0 16 16" style={{ flexShrink: 0, transition: "transform 0.2s", transform: open ? "rotate(180deg)" : "rotate(0deg)" }}>
          <path d="M4 6l4 4 4-4" stroke="#637491" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" fill="none"/>
        </svg>
      </button>
      {open && (
        <div style={{
          position: "fixed", top: dropPos.top, left: dropPos.left, zIndex: 9999,
          width: dropPos.width, maxWidth: 400, background: "var(--surface)",
          border: "1px solid var(--border-2)",
          borderRadius: "var(--radius-lg)", boxShadow: "var(--shadow-lg)",
          overflow: "hidden",
        }}>
          <div style={{ maxHeight: 260, overflowY: "auto", overflowX: "hidden", scrollbarWidth: "thin" }}>
            {options.map((opt, i) => {
              const isSelected = opt.value === value;
              return (
                <button
                  key={opt.value}
                  type="button"
                  onMouseDown={e => e.preventDefault()}
                  onClick={() => { onChange(opt.value); setOpen(false); }}
                  onMouseEnter={e => { e.currentTarget.style.background = "#f5f8ff"; }}
                  onMouseLeave={e => { e.currentTarget.style.background = "transparent"; }}
                  style={{
                    display: "flex", alignItems: "center", justifyContent: "space-between",
                    width: "100%", padding: "10px 14px",
                    background: "transparent",
                    color: "var(--text)",
                    fontWeight: 500,
                    border: 0, borderBottom: i < options.length - 1 ? "1px solid #f0f4fb" : 0,
                    borderRadius: 0, textAlign: "left", cursor: "pointer",
                    fontSize: "0.9rem", transition: "background 0.1s",
                  }}
                >
                  <span>{opt.label}</span>
                  {isSelected && (
                    <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="var(--psa-blue)" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                      <polyline points="2 7 5.5 10.5 12 3.5"/>
                    </svg>
                  )}
                </button>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}

// ── KebabMenu ─────────────────────────────────────────────────────────────────
function KebabMenu({ items }) {
  const [open, setOpen] = useState(false);
  const wrapRef = useRef(null);

  useEffect(() => {
    function onDown(e) {
      if (wrapRef.current && !wrapRef.current.contains(e.target)) setOpen(false);
    }
    document.addEventListener("mousedown", onDown);
    return () => document.removeEventListener("mousedown", onDown);
  }, []);

  return (
    <div ref={wrapRef} style={{ position: "relative", flexShrink: 0 }}>
      <button
        type="button"
        onClick={() => setOpen(v => !v)}
        title="More actions"
        className="ghost-button"
        style={{ width: 32, height: 32, padding: 0, display: "flex", alignItems: "center", justifyContent: "center" }}
      >
        <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
          <circle cx="12" cy="5" r="1.8"/><circle cx="12" cy="12" r="1.8"/><circle cx="12" cy="19" r="1.8"/>
        </svg>
      </button>
      {open && (
        <div style={{
          position: "absolute", top: "calc(100% + 6px)", right: 0, zIndex: 500,
          minWidth: 170, background: "var(--surface)",
          border: "1px solid var(--border-2)", borderRadius: "var(--radius)",
          boxShadow: "var(--shadow-lg)", overflow: "hidden",
        }}>
          {items.map((it, i) => (
            <button
              key={i}
              type="button"
              onClick={() => { setOpen(false); it.onClick(); }}
              style={{
                display: "flex", alignItems: "center", gap: 8, width: "100%",
                padding: "10px 14px", background: "transparent",
                border: 0, borderRadius: 0, borderBottom: i < items.length - 1 ? "1px solid #f0f4fb" : 0,
                textAlign: "left", cursor: "pointer", fontSize: "0.86rem", fontWeight: 600,
                color: it.danger ? "#b02020" : "var(--text)",
              }}
              onMouseEnter={e => e.currentTarget.style.background = it.danger ? "#fff2f2" : "var(--surface-2)"}
              onMouseLeave={e => e.currentTarget.style.background = "transparent"}
            >
              {it.label}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

// ── Sidebar ───────────────────────────────────────────────────────────────────
// badges: { [viewId]: count } shown beside a nav item (e.g. waiting leave pop-ups).
function Sidebar({ collapsed, open, isMobile, view, setView, isAdmin, can, canAny, hasRole, badges = {} }) {
  const navGroups = [
    {
      label: "Main",
      items: [
        {
          id: "dashboard", label: "Dashboard", visible: can("dashboard.view"),
          icon: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/></svg>
        },
        {
          id: "employee-dashboard", label: "Employee Dashboard", visible: !!hasRole,
          icon: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M20 7H4a2 2 0 0 0-2 2v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2z"/><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"/></svg>
        },
        {
          id: "files-201", label: "201 Files", visible: can("files201.view") || isAdmin,
          icon: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="9" y1="13" x2="15" y2="13"/><line x1="9" y1="17" x2="12" y2="17"/></svg>
        },
      ],
    },
    {
      label: "Employee",
      items: [
        {
          id: "employees", label: "Employees", visible: can("employees.view") || can("deductions.view") || can("payslips.view") || isAdmin,
          icon: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>
        },
        {
          id: "leave", label: "Leave", visible: canAny("leave.file", "leave.hr", "leave.crasd", "leave.director", "leave.manage"),
          icon: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/><path d="M9 16l2 2 4-4"/></svg>
        },
      ],
    },
    {
      label: "Payroll",
      items: [
        {
          id: "archive", label: "Archive", visible: canAny("archive.view", "archive.manage"),
          icon: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="21 8 21 21 3 21 3 8"/><rect x="1" y="3" width="22" height="5"/><line x1="10" y1="12" x2="14" y2="12"/></svg>
        },
      ],
    },
    {
      label: "Organization",
      items: [
        {
          id: "salary-rank", label: "Salary Rank", visible: canAny("settings.reference.view", "settings.reference.manage"),
          icon: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>
        },
        {
          id: "positions", label: "Workforce Setup", visible: canAny("settings.reference.view", "settings.reference.manage"),
          icon: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="2" y="7" width="20" height="14" rx="2"/><path d="M16 7V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v2"/><line x1="12" y1="12" x2="12" y2="16"/><line x1="10" y1="14" x2="14" y2="14"/></svg>
        },
      ],
    },
    {
      label: "System",
      items: [
        {
          id: "settings", label: "Settings",
          visible: isAdmin || canAny("settings.payroll_types.view","settings.payroll_types.manage","account.change_password","account.avatar"),
          icon: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>
        },
      ],
    },
  ];

  const visibleGroups = navGroups
    .map(group => ({ ...group, items: group.items.filter(item => item.visible !== false) }))
    .filter(group => group.items.length > 0);

  const [collapsedGroups, setCollapsedGroups] = useState({});
  const toggleGroup = (label) => setCollapsedGroups(prev => ({ ...prev, [label]: !prev[label] }));

  // On mobile: use --open to slide in from fixed position
  // On desktop: use --collapsed to shrink width (existing behaviour)
  const sidebarClass = [
    "nav-sidebar",
    isMobile && open ? "nav-sidebar--open" : "",
    !isMobile && collapsed ? "nav-sidebar--collapsed" : "",
  ].filter(Boolean).join(" ");

  return (
    <nav className={sidebarClass}>
      <div className="brand">
        <img src={psaLogo} alt="PSA" className="brand-logo" />
        <div className="brand-text">
          <span className="brand-title">Philippine Statistics Authority</span>
          <span className="brand-sub">Payroll System</span>
        </div>
      </div>
      <div className="nav-menu">
        {visibleGroups.map(group => {
          const isCollapsed = !!collapsedGroups[group.label];
          return (
            <Fragment key={group.label}>
              <button
                type="button"
                className="nav-section-label nav-section-label--toggle"
                onClick={() => toggleGroup(group.label)}
                aria-expanded={!isCollapsed}
              >
                <span>{group.label}</span>
                <svg className={`nav-section-chevron${isCollapsed ? " nav-section-chevron--collapsed" : ""}`} width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                  <polyline points="6 9 12 15 18 9"/>
                </svg>
              </button>
              {!isCollapsed && group.items.map(item => (
                <button
                  key={item.id}
                  type="button"
                  className={`nav-item${view === item.id ? " nav-item--active" : ""}`}
                  aria-current={view === item.id ? "page" : undefined}
                  title={item.label}
                  onClick={() => setView(item.id)}
                >
                  <span className="nav-item-icon">{item.icon}</span>
                  <span className="nav-item-label">{item.label}</span>
                  {badges[item.id] > 0 && (
                    <span className="nav-item-badge" aria-label={`${badges[item.id]} waiting`}>{badges[item.id] > 9 ? "9+" : badges[item.id]}</span>
                  )}
                </button>
              ))}
            </Fragment>
          );
        })}
      </div>
      <div className="nav-footer">
        Regional Statistical Services Office II
      </div>
    </nav>
  );
}

// ── TopBar ────────────────────────────────────────────────────────────────────
function TopBar({ view, currentUser, hasAvatar, avatarBlobUrl, sidebarCollapsed, isMobile, sidebarOpen, onToggleSidebar, onLogout, onGoSettings }) {
  const labels = { dashboard: "Dashboard", "employee-dashboard": "Employee Dashboard", "files-201": "201 Files", employees: "Employees", leave: "Leave", deductions: "Compensations", adjustments: "Adjustments", payslips: "Pay Slips", archive: "Payroll Archive", "salary-rank": "Salary Rank", positions: "Workforce Setup", settings: "Settings" };
  const avatarUrl = avatarBlobUrl ?? null;
  const pageLabel = labels[view] ?? view;

  return (
    <header className="topbar">
      <div className="topbar-left">
        <button
          type="button"
          onClick={onToggleSidebar}
          className="sidebar-toggle-btn"
          aria-label={isMobile ? "Open navigation menu" : (sidebarCollapsed ? "Expand sidebar" : "Collapse sidebar")}
        >
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round">
            <line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="18" x2="21" y2="18"/>
          </svg>
        </button>
        {/* Desktop breadcrumb — hidden on mobile */}
        <span className="topbar-crumb">
          <span className="topbar-crumb-prefix">PSA</span>
          <span className="topbar-crumb-sep"> / </span>
          <span className="topbar-crumb-prefix">RSSO II</span>
          <span className="topbar-crumb-sep"> / </span>
          <span className="topbar-crumb-current">{pageLabel}</span>
        </span>
      </div>

      {/* Mobile center: page title */}
      {isMobile && (
        <span className="topbar-mobile-title">{pageLabel}</span>
      )}

      <div className="topbar-right">
        <span className="topbar-username" style={{ fontSize: "0.78rem", color: "#888", textTransform: "uppercase" }}>{currentUser}</span>
        <div
          onClick={onGoSettings}
          className="topbar-avatar"
          title="Profile"
          style={{ width: 34, height: 34, borderRadius: "50%", background: avatarUrl ? "transparent" : "#0038a8", overflow: "hidden", cursor: "pointer", flexShrink: 0, position: "relative" }}
        >
          {avatarUrl ? (
            <img src={avatarUrl} alt="Avatar" style={{ position: "absolute", top: 0, left: 0, width: "100%", height: "100%", objectFit: "cover", display: "block" }} />
          ) : (
            <span style={{ position: "absolute", top: 0, left: 0, width: "100%", height: "100%", display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontSize: "0.85rem", fontWeight: 600 }}>
              {currentUser?.[0]?.toUpperCase() ?? "?"}
            </span>
          )}
        </div>
        <button onClick={onLogout} className="ghost-button topbar-signout-btn" style={{ fontSize: "0.78rem" }}>
          {isMobile ? (
            <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/>
              <polyline points="16 17 21 12 16 7"/>
              <line x1="21" y1="12" x2="9" y2="12"/>
            </svg>
          ) : (
            <span className="topbar-signout-text">Sign Out</span>
          )}
        </button>
      </div>
    </header>
  );
}

// ── StatCard ──────────────────────────────────────────────────────────────────
function StatCard({ label, value, color = "var(--psa-blue)", breakdown }) {
  const [open, setOpen] = useState(false);
  return (
    <div
      onClick={() => breakdown && setOpen(o => !o)}
      style={{
        background: "var(--surface)", border: "1px solid var(--border)", borderRadius: "var(--radius-lg)",
        padding: "14px 18px", boxShadow: "var(--shadow-sm)",
        cursor: breakdown ? "pointer" : "default",
        userSelect: "none",
      }}
    >
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
        <div style={{ fontSize: "0.72rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.07em" }}>{label}</div>
        {breakdown && <span style={{ fontSize: "0.65rem", color: "var(--text-muted)" }}>{open ? "▲" : "▼"}</span>}
      </div>
      <div style={{ fontSize: "1.75rem", fontWeight: 800, color, letterSpacing: "-0.03em", lineHeight: 1 }}>{value}</div>
      {open && breakdown && (
        <div style={{ marginTop: 10, borderTop: "1px solid var(--border)", paddingTop: 10, display: "flex", flexDirection: "column", gap: 6 }}>
          {breakdown.map(({ label: bl, value: bv, color: bc }) => (
            <div key={bl} style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <span style={{ fontSize: "0.78rem", color: "var(--text-muted)", fontWeight: 600 }}>{bl}</span>
              <span style={{ fontSize: "1.05rem", fontWeight: 700, color: bc ?? color }}>{bv}</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ── DashboardView ─────────────────────────────────────────────────────────────
// READ-ONLY PAGE — never add POST/PUT/PATCH/DELETE here.
// This page only fetches and displays data. It does not write to the database.
// Deductions are fetched with ?readonly=1 to prevent syncComputedDeductionsToLedger from running.
// Step increments are displayed for awareness only — processing must happen elsewhere.
// See CLAUDE.md "STRICT READ-ONLY PAGES" for the full rule.
function DashboardView({ refreshKey, isSuperAdmin, permissions = [] }) {
  const canFinancial = isSuperAdmin || permissions.includes("dashboard.financial_report");
  const [stats, setStats]         = useState(null);
  const [loading, setLoading]     = useState(true);
  const [salaryPage, setSalaryPage]   = useState(0);
  const [quincenaPage, setQuincenaPage] = useState(0);
  const hasData = useRef(false);

  useEffect(() => {
    if (!hasData.current) setLoading(true);
    fetch(`${apiUrl}/stats`, { credentials: "include" })
      .then(r => r.json())
      .then(s => {
        setStats(s);
        hasData.current = true;
      })
      .finally(() => setLoading(false));
  }, [refreshKey]);


  if (loading) return (
    <div style={{ padding: "80px 24px", textAlign: "center", color: "var(--text-muted)" }}>
      <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ animation: "spin 1s linear infinite", display: "block", margin: "0 auto 12px" }}>
        <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
      </svg>
      <div style={{ fontSize: "0.82rem" }}>Loading…</div>
    </div>
  );

  const deductions = stats?.deductionsByCategory ?? {};
  const topBurden  = stats?.topLoanBurden ?? [];
  const empStatus  = stats?.employeeStatus ?? { active: 0, onLeave: 0, suspended: 0, otherInactive: 0, all: [] };
  const onLeaveList   = empStatus.all.filter(e => e.status === 'On Leave');
  const suspendedList = empStatus.all.filter(e => e.status === 'Suspended');

  return (
    <div style={{ maxWidth: "100%" }}>

      {/* ── STAT CARDS ── */}
      {canFinancial ? (
        <>
          <div style={{ display: "flex", alignItems: "center", gap: 7, margin: "2px 0 14px" }}>
            <span style={{ fontSize: "0.6rem", fontWeight: 700, color: "var(--text-muted)", letterSpacing: "0.12em", textTransform: "uppercase" }}>Philippine Statistics Authority · RSSO II</span>
            <span style={{ width: 3, height: 3, borderRadius: "50%", background: "var(--border)", display: "inline-block", flexShrink: 0 }} />
            <span style={{ fontSize: "0.6rem", color: "var(--text-muted)", letterSpacing: "0.06em" }}>Payroll System Overview</span>
          </div>
          <div className="grid-4-col dash-stat-cards" style={{ gap: 14, marginBottom: 14 }}>
            {[
              { label: "Total Employees", value: stats?.totalEmployees ?? 0, sub: `${stats?.regularEmployees ?? 0} Reg · ${stats?.contractualEmployees ?? 0} Cont`, accent: "#0038a8",
                icon: <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2M9 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8ZM23 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75"/> },
              { label: "Monthly Payroll", value: `₱ ${fmt(stats?.totalNetPaid ?? 0)}`, sub: `₱${fmt(stats?.regularSalary ?? 0)} reg · ₱${fmt(stats?.contractualSalary ?? 0)} cont`, accent: "#ce1126",
                icon: <><rect x="1" y="5" width="22" height="14" rx="2"/><line x1="1" y1="10" x2="23" y2="10"/></> },
              { label: "Avg. Take-Home", value: `₱ ${fmt(stats?.avgNetPay ?? 0)}`, sub: "Per employee, net of deductions", accent: "#15803d",
                icon: <polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/> },
              { label: "Archives", value: stats?.archiveCount ?? 0, sub: "Quincenas saved", accent: "#b45309",
                icon: <><rect x="1" y="3" width="22" height="5" rx="1"/><path d="M3 8v11a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V8"/><line x1="10" y1="13" x2="14" y2="13"/></> },
            ].map(card => (
              <div key={card.label} className="dash-stat-card" style={{ background: "var(--surface)", border: "1px solid var(--border)", borderRadius: "var(--radius-lg)", padding: "18px 20px" }}>
                <div className="dash-stat-icon" style={{
                  width: 34, height: 34, borderRadius: 10, marginBottom: 14,
                  background: `${card.accent}15`, display: "flex", alignItems: "center", justifyContent: "center",
                }}>
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke={card.accent} strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">{card.icon}</svg>
                </div>
                <div style={{ fontSize: "0.6rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: 5 }}>{card.label}</div>
                <div className="dash-stat-value" style={{ fontSize: "1.4rem", fontWeight: 800, color: "var(--text)", letterSpacing: "-0.01em", lineHeight: 1.15, marginBottom: 5 }}>{card.value}</div>
                <div className="dash-stat-sub" style={{ fontSize: "0.68rem", color: "var(--text-muted)" }}>{card.sub}</div>
              </div>
            ))}
          </div>

          {/* ── CHARTS ── */}
          <div className="grid-chart-split" style={{ gap: 14, marginBottom: 14 }}>
            <div style={{ background: "var(--surface)", border: "1px solid var(--border)", borderRadius: "var(--radius-lg)", padding: "18px 22px" }}>
              <div style={{ fontSize: "0.78rem", fontWeight: 700, color: "var(--text)", marginBottom: 2 }}>Employees by Division</div>
              <div style={{ fontSize: "0.68rem", color: "var(--text-muted)", marginBottom: 14 }}>Headcount per division</div>
              {(stats?.employeesByDivision ?? []).length === 0 ? (
                <div style={{ color: "var(--text-muted)", fontSize: "0.83rem", padding: "40px 0", textAlign: "center" }}>No employee data yet.</div>
              ) : (
                <div onTouchEnd={() => document.activeElement?.blur()}>
                <ResponsiveContainer width="100%" height={260}>
                  <BarChart tabIndex={-1} data={stats.employeesByDivision.map(d => ({ ...d, total: (d.Regular ?? 0) + (d.Contractual ?? 0) }))} margin={{ top: 4, right: 8, left: -18, bottom: 4 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="var(--border)" vertical={false} />
                    <XAxis dataKey="division" tick={{ fontSize: 11 }} interval={0} angle={-20} textAnchor="end" height={50} />
                    <YAxis tick={{ fontSize: 11 }} allowDecimals={false} />
                    <Tooltip contentStyle={{ fontSize: "0.75rem", borderRadius: 8 }} formatter={v => [v, "Employees"]} />
                    <Bar dataKey="total" name="Employees" radius={[4, 4, 0, 0]}>
                      {stats.employeesByDivision.map((_, i) => (
                        <Cell key={i} fill={["#0038a8", "#ce1126", "#fcd116", "#15803d", "#b45309", "#6d28d9"][i % 6]} />
                      ))}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
                </div>
              )}
            </div>
            <div style={{ background: "var(--surface)", border: "1px solid var(--border)", borderRadius: "var(--radius-lg)", padding: "18px 22px" }}>
              <div style={{ fontSize: "0.78rem", fontWeight: 700, color: "var(--text)", marginBottom: 2 }}>Employee Status</div>
              <div style={{ fontSize: "0.68rem", color: "var(--text-muted)", marginBottom: 14 }}>Active · On Leave · Suspended</div>
              {empStatus.active + empStatus.onLeave + empStatus.suspended + empStatus.otherInactive === 0 ? (
                <div style={{ color: "var(--text-muted)", fontSize: "0.83rem", padding: "40px 0", textAlign: "center" }}>No employee data yet.</div>
              ) : (() => {
                const statusData = [
                  { name: "Active", value: empStatus.active, color: "#0038a8" },
                  { name: "On Leave", value: empStatus.onLeave, color: "#f59e0b" },
                  { name: "Suspended", value: empStatus.suspended, color: "#ce1126" },
                  { name: "Other Inactive", value: empStatus.otherInactive, color: "#64748b" },
                ].filter(d => d.value > 0);
                return (
                  <div onTouchEnd={() => document.activeElement?.blur()}>
                  <ResponsiveContainer width="100%" height={260}>
                    <PieChart tabIndex={-1}>
                      <Pie data={statusData} dataKey="value" nameKey="name" cx="50%" cy="50%" innerRadius={54} outerRadius={84} paddingAngle={2}>
                        {statusData.map((d, i) => <Cell key={i} fill={d.color} />)}
                      </Pie>
                      <Tooltip formatter={(v, n) => [v + " employees", n]} contentStyle={{ fontSize: "0.75rem", borderRadius: 8 }} />
                      <Legend wrapperStyle={{ fontSize: "0.7rem" }} />
                    </PieChart>
                  </ResponsiveContainer>
                  </div>
                );
              })()}
            </div>
          </div>

        </>
      ) : (
        <div className="page-header">
          <div className="page-header-title">Dashboard</div>
          <div className="page-header-sub">PSA Payroll System overview</div>
        </div>
      )}


    </div>
  );
}

// ── EmployeeStatusView ────────────────────────────────────────────────────────
function LoanTypeDropdown({ value, onChange, loanGroups, optionLabel }) {
  const [open, setOpen] = useState(false);
  const isMobile = useIsMobile();
  const ref = useRef(null);
  useEffect(() => {
    function handleClick(e) { if (ref.current && !ref.current.contains(e.target)) setOpen(false); }
    document.addEventListener("mousedown", handleClick);
    return () => document.removeEventListener("mousedown", handleClick);
  }, []);
  const selected = loanGroups.find(g => g.k === value);
  return (
    <div ref={ref} style={{ flex: 1, position: "relative" }}>
      <button type="button" onClick={() => setOpen(o => !o)}
        style={{ width: "100%", display: "flex", alignItems: "center", justifyContent: "space-between", padding: isMobile ? "11px 14px" : "7px 12px", border: `1.5px solid ${open ? "#4a90d9" : "var(--border)"}`, borderRadius: "var(--radius-sm)", background: "var(--surface)", color: "var(--text)", fontSize: isMobile ? "0.9rem" : "0.82rem", fontWeight: 600, cursor: "pointer", boxShadow: open ? "0 0 0 3px rgba(74,144,217,0.12)" : "var(--shadow-xs)", textAlign: "left", WebkitTapHighlightColor: "transparent" }}>
        <span>{selected ? optionLabel(selected) : "Select loan"}</span>
        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ flexShrink: 0, transform: open ? "rotate(180deg)" : "none", transition: "transform 0.15s", color: "var(--text-muted)" }}>
          <polyline points="6 9 12 15 18 9"/>
        </svg>
      </button>
      {open && (
        <div style={{ position: "absolute", top: "calc(100% + 4px)", left: 0, right: 0, background: "var(--surface)", border: "1.5px solid var(--border-2)", borderRadius: "var(--radius-sm)", boxShadow: "0 8px 24px rgba(0,0,0,0.12)", zIndex: 200, maxHeight: 280, overflowY: "auto" }}>
          {loanGroups.map(g => {
            const isSelected = g.k === value;
            return (
              <button key={g.k} type="button"
                onClick={() => { onChange(g.k); setOpen(false); }}
                style={{ width: "100%", display: "flex", alignItems: "center", justifyContent: "space-between", padding: isMobile ? "12px 16px" : "8px 14px", border: "none", background: isSelected ? "var(--psa-navy)" : "transparent", color: isSelected ? "#fff" : "var(--text)", fontSize: isMobile ? "0.9rem" : "0.82rem", fontWeight: isSelected ? 700 : 400, cursor: "pointer", textAlign: "left", WebkitTapHighlightColor: "transparent" }}
                onMouseEnter={e => { if (!isSelected) e.currentTarget.style.background = "var(--surface-2)"; }}
                onMouseLeave={e => { if (!isSelected) e.currentTarget.style.background = "transparent"; }}>
                <span>{optionLabel(g)}</span>
                {isSelected && (
                  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"/></svg>
                )}
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}

function LedgerAccordion({ loanPayments = [], loans = [], tab = "active", setTab }) {
  const isMobile = useIsMobile();
  const activeLoans = loans.filter(l => l.status === "active");
  const closedLoans = loans.filter(l => l.status !== "active");

  // Dropdown built from the loans list directly — always shows all loans in the tab
  const tabLoans = tab === "active" ? activeLoans : closedLoans;
  const labelCount = {};
  for (const l of tabLoans) { const lbl = l.loan_type_name ?? "Loan"; labelCount[lbl] = (labelCount[lbl] ?? 0) + 1; }
  const loanGroups = tabLoans.map(l => ({ k: `loan:${l.id}`, label: l.loan_type_name ?? "Loan", loanId: l.id }));
  const optionLabel = g => labelCount[g.label] > 1 ? `${g.label} #${g.loanId}` : g.label;

  const [selKey, setSelKey] = useState(() => loanGroups[0]?.k ?? "");
  const [rowPage, setRowPage] = useState(0);
  const [expandedCard, setExpandedCard] = useState(null);

  // Reset to first loan of new tab when tab changes
  const prevTab = useRef(tab);
  useEffect(() => {
    if (prevTab.current !== tab) {
      prevTab.current = tab;
      setSelKey(loanGroups[0]?.k ?? "");
      setRowPage(0);
    }
  });

  // Payment rows for the selected loan from loanPayments
  const selLoanId = selKey.replace("loan:", "");
  const rows = [...loanPayments]
    .filter(r => String(r.loan_id) === String(selLoanId))
    .sort((a, b) => new Date(b.processed_at) - new Date(a.processed_at));

  const PAGE = 5;
  const totalPages = Math.ceil(rows.length / PAGE) || 1;
  const safePage   = Math.min(rowPage, totalPages - 1);
  const pageRows   = rows.slice(safePage * PAGE, (safePage + 1) * PAGE);
  const fmtAmt = n => `₱${Number(n ?? 0).toLocaleString("en-PH", { minimumFractionDigits: 2 })}`;
  const fmtDt  = d => d ? new Date(d).toLocaleString("en-PH", { year:"numeric", month:"short", day:"2-digit", hour:"2-digit", minute:"2-digit", hour12: true }) : "—";

  return (
    <div>
      {/* Toolbar: Active/Closed tabs left, LOAN dropdown right */}
      <div className="loan-toolbar">
        {setTab && loans.length > 0 && (
          <div className="loan-toolbar-tabs">
            <WsTabs small label="Loan status" value={tab} onChange={setTab}
              tabs={[
                { id: "active", label: `Active${activeLoans.length ? ` (${activeLoans.length})` : ""}` },
                { id: "closed", label: `Closed${closedLoans.length ? ` (${closedLoans.length})` : ""}` },
              ]} />
          </div>
        )}
        {loanGroups.length > 1 && (
          <div className="loan-toolbar-select">
            {!isMobile && <label style={{ fontSize: "0.6rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.1em", flexShrink: 0 }}>Loan</label>}
            <LoanTypeDropdown
              value={selKey}
              onChange={k => { setSelKey(k); setRowPage(0); }}
              loanGroups={loanGroups}
              optionLabel={optionLabel}
            />
          </div>
        )}
      </div>
      {/* Selected loan summary */}
      {(() => {
        const sel = tabLoans.find(l => String(l.id) === String(selLoanId));
        if (!sel) return null;
        const fmtP = n => `₱${Number(n ?? 0).toLocaleString("en-PH", { minimumFractionDigits: 2 })}`;
        const total = Number(sel.total_amount ?? 0);
        const remaining = Number(sel.remaining_balance ?? 0);
        const monthly = Number(sel.monthly_amount ?? 0);
        const paid = total - remaining;
        const pct = total > 0 ? Math.max(0, Math.min(100, (paid / total) * 100)) : 0;
        const isActive = sel.status === "active";
        const fmtD = d => d ? new Date(d).toLocaleDateString("en-PH", { year: "numeric", month: "short", day: "numeric" }) : "—";
        if (isMobile) return (
          <div style={{ padding: "14px 16px 16px", borderBottom: "1px solid var(--border)", background: "var(--surface-2)" }}>
            {/* Loan name — only shown when no picker (1 loan in tab) */}
            {loanGroups.length <= 1 && (
              <div style={{ fontSize: "0.92rem", fontWeight: 700, color: "var(--text)", marginBottom: 6 }}>
                {sel.loan_type_name ?? "Loan"}
              </div>
            )}
            {/* Date row */}
            <div style={{ fontSize: "0.72rem", color: "var(--text-muted)", marginBottom: 12 }}>
              Started <strong style={{ color: "var(--text)" }}>{fmtD(sel.start_date)}</strong>
              {!isActive && sel.closed_date && <> · Closed <strong style={{ color: "var(--text)" }}>{fmtD(sel.closed_date)}</strong></>}
            </div>
            {/* 2×2 stats grid */}
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "12px 0", marginBottom: 14 }}>
              {[["Total Loan", fmtP(total), false], ["Monthly", fmtP(monthly), true], ["Amount Paid", fmtP(paid), false], ["Remaining", fmtP(remaining), true]].map(([lbl, val, rightCol]) => (
                <div key={lbl} style={{ paddingLeft: rightCol ? 16 : 0, borderLeft: rightCol ? "1px solid var(--border)" : "none" }}>
                  <div style={{ fontSize: "0.62rem", color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.07em", fontWeight: 700, marginBottom: 3 }}>{lbl}</div>
                  <div style={{ fontSize: "0.95rem", fontWeight: 700, color: lbl === "Remaining" ? "var(--psa-navy)" : "var(--text)", fontVariantNumeric: "tabular-nums" }}>{val}</div>
                </div>
              ))}
            </div>
            {/* Progress bar */}
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <div style={{ flex: 1, height: 6, borderRadius: 999, background: "var(--border)", overflow: "hidden" }}>
                <div style={{ height: "100%", width: `${pct}%`, borderRadius: 999, background: isActive ? "var(--psa-navy)" : "var(--text-muted)", transition: "width 0.4s" }} />
              </div>
              <span style={{ fontSize: "0.72rem", fontWeight: 700, color: "var(--text-muted)", minWidth: 52, textAlign: "right" }}>{Math.round(pct)}% paid</span>
            </div>
          </div>
        );
        return (
          <div style={{ padding: "10px 14px 12px", borderBottom: "1px solid var(--border)", background: "var(--surface-2)" }}>
            <div style={{ fontSize: "0.7rem", color: "var(--text-muted)", marginBottom: 6 }}>
              Started <strong style={{ color: "var(--text)" }}>{fmtD(sel.start_date)}</strong>
              {!isActive && sel.closed_date && <> · Closed <strong style={{ color: "var(--text)" }}>{fmtD(sel.closed_date)}</strong></>}
            </div>
            <div className="loan-summary-stats" style={{ display: "flex", gap: 20, flexWrap: "wrap" }}>
              {[["Total Loan", fmtP(total)], ["Monthly", fmtP(monthly)], ["Remaining", fmtP(remaining)]].map(([lbl, val]) => (
                <div key={lbl}>
                  <div style={{ fontSize: "0.65rem", color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.06em" }}>{lbl}</div>
                  <div style={{ fontSize: "0.85rem", fontWeight: 700, color: "var(--text)", fontVariantNumeric: "tabular-nums" }}>{val}</div>
                </div>
              ))}
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 8 }}>
              <div style={{ flex: 1, height: 3, borderRadius: 999, background: "var(--border)", overflow: "hidden" }}>
                <div style={{ height: "100%", width: `${pct}%`, borderRadius: 999, background: isActive ? "var(--psa-navy)" : "var(--text-muted)", transition: "width 0.4s" }} />
              </div>
              <span style={{ fontSize: "0.65rem", fontWeight: 700, color: "var(--text-muted)", minWidth: 40, textAlign: "right" }}>{Math.round(pct)}% paid</span>
            </div>
          </div>
        );
      })()}
      {/* Payment history table — desktop only */}
      <div className="loan-table-wrap" style={{ display: isMobile ? "none" : undefined }}>
        {rows.length === 0 ? (
          <div className="loan-empty">
            <svg width="34" height="34" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" strokeLinejoin="round">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="8" y1="13" x2="16" y2="13"/><line x1="8" y1="17" x2="13" y2="17"/>
            </svg>
            <div className="loan-empty-title">No payments recorded yet</div>
            <div className="loan-empty-sub">Deductions for this loan will appear here once the first payroll run processes it.</div>
          </div>
        ) : (
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.82rem" }}>
          <thead>
            <tr style={{ background: "var(--surface-2)" }}>
              <th style={{ padding: "7px 10px", textAlign: "left", fontSize: "0.68rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", whiteSpace: "nowrap" }}>Amount</th>
              <th className="loan-col-hide" style={{ padding: "7px 10px", textAlign: "left", fontSize: "0.68rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", whiteSpace: "nowrap" }}>Bal. Before</th>
              <th className="loan-col-hide" style={{ padding: "7px 10px", textAlign: "left", fontSize: "0.68rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", whiteSpace: "nowrap" }}>Bal. After</th>
              <th style={{ padding: "7px 10px", textAlign: "left", fontSize: "0.68rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", whiteSpace: "nowrap" }}>Processed At</th>
            </tr>
          </thead>
          <tbody>
            {pageRows.map((row, i) => (
              <tr key={i} style={{ borderTop: "1px solid var(--border)" }}>
                <td style={{ padding: "8px 10px", fontVariantNumeric: "tabular-nums", fontWeight: 600 }}>{fmtAmt(row.amount_deducted)}</td>
                <td className="loan-col-hide" style={{ padding: "8px 10px", fontVariantNumeric: "tabular-nums", color: "var(--text-muted)" }}>{fmtAmt(row.balance_before)}</td>
                <td className="loan-col-hide" style={{ padding: "8px 10px", fontVariantNumeric: "tabular-nums", color: "var(--text-muted)" }}>{fmtAmt(row.balance_after)}</td>
                <td style={{ padding: "8px 10px", color: "var(--text-muted)" }}>
                  {row.processed_at ? (() => {
                    const d = new Date(row.processed_at);
                    return <>
                      <span style={{ display: "block", whiteSpace: "nowrap", fontSize: "0.82rem" }}>{d.toLocaleDateString("en-PH", { year: "numeric", month: "short", day: "numeric" })}</span>
                      <span style={{ display: "block", whiteSpace: "nowrap", fontSize: "0.72rem", color: "var(--text-muted)" }}>{d.toLocaleTimeString("en-PH", { hour: "2-digit", minute: "2-digit", hour12: true })}</span>
                    </>;
                  })() : "—"}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        )}
      </div>

      {/* Mobile card list — replaces table on small screens */}
      <div className="loan-tx-cards" style={{ display: isMobile ? "flex" : "none", flexDirection: "column" }}>
        {isMobile && rows.length > 0 && (
          <div style={{ padding: "10px 16px 6px", borderBottom: "1px solid var(--border)" }}>
            <span style={{ fontSize: "0.62rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.1em" }}>Payment History</span>
          </div>
        )}
        {rows.length === 0 ? (
          <div className="loan-empty">
            <svg width="34" height="34" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" strokeLinejoin="round">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="8" y1="13" x2="16" y2="13"/><line x1="8" y1="17" x2="13" y2="17"/>
            </svg>
            <div className="loan-empty-title">No payments recorded yet</div>
            <div className="loan-empty-sub">Deductions for this loan will appear here once the first payroll run processes it.</div>
          </div>
        ) : pageRows.map((row, i) => {
          const d = row.processed_at ? new Date(row.processed_at) : null;
          const isOpen = expandedCard === i;
          const hasBal = row.balance_before != null || row.balance_after != null;
          return (
            <div key={i}
              className={`loan-tx-card${isOpen ? " loan-tx-card--open" : ""}`}
              onClick={() => setExpandedCard(isOpen ? null : i)}
              style={{ cursor: hasBal ? "pointer" : "default" }}
            >
              <div className="loan-tx-card__row">
                <span className="loan-tx-card__amount">{fmtAmt(row.amount_deducted)}</span>
                <div className="loan-tx-card__right">
                  {d ? <>
                    <span className="loan-tx-card__date">{d.toLocaleDateString("en-PH", { year: "numeric", month: "short", day: "numeric" })}</span>
                    <span className="loan-tx-card__time">{d.toLocaleTimeString("en-PH", { hour: "2-digit", minute: "2-digit", hour12: true })}</span>
                  </> : <span className="loan-tx-card__date">—</span>}
                  {hasBal && <span className="loan-tx-card__chevron">▾</span>}
                </div>
              </div>
              {isOpen && hasBal && (
                <div className="loan-tx-card__detail">
                  <div className="loan-tx-card__detail-col">
                    <span className="loan-tx-card__detail-label">Balance Before</span>
                    <span className="loan-tx-card__detail-val">{fmtAmt(row.balance_before)}</span>
                  </div>
                  <div className="loan-tx-card__detail-col">
                    <span className="loan-tx-card__detail-label">Balance After</span>
                    <span className="loan-tx-card__detail-val loan-tx-card__detail-val--after">{fmtAmt(row.balance_after)}</span>
                  </div>
                </div>
              )}
            </div>
          );
        })}
      </div>

      {totalPages > 1 && (
        <div style={{ display: "flex", alignItems: "center", justifyContent: "flex-end", gap: 8, padding: "10px 14px", borderTop: "1px solid var(--border)" }}>
          <button onClick={() => setRowPage(p => Math.max(0, p - 1))} disabled={safePage === 0}
            style={{ background: "none", border: "1px solid var(--border)", borderRadius: 6, padding: "5px 14px", fontSize: "0.9rem", cursor: safePage === 0 ? "default" : "pointer", color: safePage === 0 ? "var(--text-muted)" : "var(--psa-navy)", fontWeight: 700 }}>‹</button>
          <span style={{ fontSize: "0.75rem", color: "var(--text-muted)", minWidth: 36, textAlign: "center" }}>{safePage + 1} / {totalPages}</span>
          <button onClick={() => setRowPage(p => Math.min(totalPages - 1, p + 1))} disabled={safePage === totalPages - 1}
            style={{ background: "none", border: "1px solid var(--border)", borderRadius: 6, padding: "5px 14px", fontSize: "0.9rem", cursor: safePage === totalPages - 1 ? "default" : "pointer", color: safePage === totalPages - 1 ? "var(--text-muted)" : "var(--psa-navy)", fontWeight: 700 }}>›</button>
        </div>
      )}
    </div>
  );
}

function MiniDropdown({ value, onChange, options }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  useEffect(() => {
    function handleClick(e) { if (ref.current && !ref.current.contains(e.target)) setOpen(false); }
    document.addEventListener("mousedown", handleClick);
    return () => document.removeEventListener("mousedown", handleClick);
  }, []);
  const selected = options.find(o => String(o.value) === String(value)) ?? options[0];
  return (
    <div ref={ref} style={{ position: "relative" }}>
      <button type="button" onClick={() => setOpen(o => !o)}
        style={{ display: "flex", alignItems: "center", gap: 6, padding: "4px 9px", fontSize: "0.78rem", fontWeight: 500, borderRadius: "var(--radius-sm)", border: `1.5px solid ${open ? "#4a90d9" : "var(--border)"}`, background: "var(--surface)", color: "var(--text)", cursor: "pointer", boxShadow: open ? "0 0 0 3px rgba(74,144,217,0.12)" : "none" }}>
        <span>{selected?.label ?? "—"}</span>
        <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" style={{ flexShrink: 0, transform: open ? "rotate(180deg)" : "none", transition: "transform 0.15s", color: "var(--text-muted)" }}>
          <polyline points="6 9 12 15 18 9"/>
        </svg>
      </button>
      {open && (
        <div style={{ position: "absolute", top: "calc(100% + 4px)", left: 0, background: "var(--surface)", border: "1.5px solid var(--border-2)", borderRadius: "var(--radius-sm)", boxShadow: "0 8px 24px rgba(0,0,0,0.12)", zIndex: 200, maxHeight: 220, overflowY: "auto", minWidth: 80 }}>
          {options.map(o => {
            const isSelected = String(o.value) === String(value);
            return (
              <button key={o.value} type="button"
                onClick={() => { onChange(o.value); setOpen(false); }}
                style={{ width: "100%", display: "flex", alignItems: "center", justifyContent: "space-between", padding: "7px 12px", border: "none", background: isSelected ? "var(--psa-navy)" : "transparent", color: isSelected ? "#fff" : "var(--text)", fontSize: "0.8rem", fontWeight: isSelected ? 700 : 400, cursor: "pointer", textAlign: "left", whiteSpace: "nowrap" }}
                onMouseEnter={e => { if (!isSelected) e.currentTarget.style.background = "var(--surface-2)"; }}
                onMouseLeave={e => { if (!isSelected) e.currentTarget.style.background = "transparent"; }}>
                {o.label}
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}

function PayPeriodsPanel({ payPeriods }) {
  const MON = ["","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const fmtP = n => `₱${Number(n ?? 0).toLocaleString("en-PH", { minimumFractionDigits: 2 })}`;

  const years = [...new Set(payPeriods.map(p => p.year))].sort((a, b) => b - a);
  const monthsFor = yr => [...new Set(payPeriods.filter(p => p.year == yr).map(p => p.month))].sort((a, b) => b - a);
  const qFor = (yr, mo) => [...new Set(payPeriods.filter(p => p.year == yr && p.month == mo).map(p => p.quincena))].sort((a, b) => a - b);

  const [selYear,  setSelYear]  = useState(years[0] ?? "");
  const [selMonth, setSelMonth] = useState(() => monthsFor(years[0])[0] ?? "");
  const [selQ,     setSelQ]     = useState(() => { const qs = qFor(years[0], monthsFor(years[0])[0]); return qs[0] ?? 1; });

  // Resync whenever the underlying data changes (e.g. refetch after switching employees)
  // and the currently selected year is no longer valid for this data set.
  useEffect(() => {
    if (years.length > 0 && !years.includes(selYear)) {
      setSelYear(years[0]);
    }
  }, [payPeriods]);

  useEffect(() => {
    const mo = monthsFor(selYear)[0] ?? "";
    setSelMonth(mo);
    const qs = qFor(selYear, mo);
    setSelQ(qs[0] ?? 1);
  }, [selYear]);
  useEffect(() => {
    const qs = qFor(selYear, selMonth);
    setSelQ(qs[0] ?? 1);
  }, [selMonth]);

  const quincenas = qFor(selYear, selMonth);
  const record = payPeriods.find(p => p.year == selYear && p.month == selMonth && p.quincena == selQ) ?? null;

  const mkSel = (label, val, setVal, options) => (
    <div style={{ display: "flex", alignItems: "center", gap: 5 }}>
      <span style={{ fontSize: "0.58rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.1em" }}>{label}</span>
      <MiniDropdown value={val} onChange={setVal} options={options} />
    </div>
  );

  const sectionHeader = label => (
    <div style={{ padding: "5px 16px 3px", background: "var(--surface-2)", fontSize: "0.58rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.1em" }}>{label}</div>
  );
  const dataRow = (label, value, valueStyle = {}) => (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 16px", borderBottom: "1px solid var(--border)" }}>
      <span style={{ fontSize: "0.82rem", color: "var(--text-muted)" }}>{label}</span>
      <span style={{ fontSize: "0.85rem", fontWeight: 500, fontVariantNumeric: "tabular-nums", ...valueStyle }}>{value}</span>
    </div>
  );

  const snap = record?.snapshot ?? null;
  // snapshot.type_labels stores display names as they were at archive time — use those directly
  const labels = snap?.type_labels ?? {};
  const getLabel = key => labels[key] || key.replace(/_/g, " ").replace(/\b\w/g, c => c.toUpperCase());

  const snapDeds   = snap?.deductions ?? {};
  const snapAllows = snap?.allowances ?? {};
  const dedLines   = Object.entries(snapDeds).filter(([, v]) => Number(v) !== 0);
  const earnLines  = Object.entries(snapAllows).filter(([, v]) => Number(v) !== 0);

  return (
    <div>
      {/* Filter bar */}
      <div style={{ padding: "9px 14px", borderBottom: "1px solid var(--border)", background: "var(--surface-2)", display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
        {mkSel("Year",  selYear,  setSelYear,  years.map(y => ({ value: y, label: String(y) })))}
        {mkSel("Month", selMonth, setSelMonth, monthsFor(selYear).map(m => ({ value: m, label: MON[m] })))}
        <div style={{ marginLeft: "auto", display: "flex", gap: 4 }}>
          {quincenas.map(q => (
            <button key={q} onClick={() => setSelQ(q)} style={{
              padding: "4px 13px", fontSize: "0.74rem", fontWeight: 600, borderRadius: 999, cursor: "pointer",
              color: selQ == q ? "#fff" : "var(--text-muted)",
              background: selQ == q ? "var(--psa-navy)" : "transparent",
              border: `1px solid ${selQ == q ? "var(--psa-navy)" : "var(--border)"}`,
            }}>{q === 1 ? "1st Q" : "2nd Q"}</button>
          ))}
        </div>
      </div>

      {!record ? (
        <div style={{ padding: "28px 16px", textAlign: "center", color: "var(--text-muted)", fontSize: "0.82rem" }}>No record for selected period.</div>
      ) : (
        <div style={{ padding: "16px 18px" }}>
          {/* Period title — plain text, no box */}
          <div style={{ marginBottom: 14 }}>
            <div style={{ fontSize: "1rem", fontWeight: 700, color: "var(--text)" }}>
              {MON[record.month]} {record.year} &mdash; {record.quincena === 1 ? "1st Quincena" : "2nd Quincena"}
            </div>
            {record.saved_at && (
              <div style={{ fontSize: "0.7rem", color: "var(--text-muted)", marginTop: 2 }}>
                Archived {new Date(record.saved_at).toLocaleDateString("en-PH", { year: "numeric", month: "short", day: "numeric" })}
              </div>
            )}
          </div>

          {/* EARNINGS */}
          <div style={{ fontSize: "0.65rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: 6 }}>Earnings</div>
          {snap?.basic_pay != null && (
            <div style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", fontSize: "0.83rem", color: "var(--text)" }}>
              <span>Basic Pay</span><span style={{ fontVariantNumeric: "tabular-nums" }}>{fmtP(snap.basic_pay)}</span>
            </div>
          )}
          {earnLines.map(([k, v]) => (
            <div key={k} style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", fontSize: "0.83rem", color: "var(--text)" }}>
              <span>{getLabel(k)}</span><span style={{ fontVariantNumeric: "tabular-nums" }}>{fmtP(v)}</span>
            </div>
          ))}
          <div style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", marginTop: 2, borderTop: "1px solid var(--border)", fontSize: "0.83rem", fontWeight: 700, color: "var(--text)" }}>
            <span>Gross Pay</span><span style={{ fontVariantNumeric: "tabular-nums" }}>{fmtP(record.gross_pay)}</span>
          </div>

          {/* DEDUCTIONS */}
          <div style={{ fontSize: "0.65rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.1em", marginTop: 16, marginBottom: 6 }}>Deductions</div>
          {dedLines.length > 0
            ? dedLines.map(([k, v]) => (
                <div key={k} style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", fontSize: "0.83rem", color: "var(--text)" }}>
                  <span>{getLabel(k)}</span><span style={{ fontVariantNumeric: "tabular-nums", color: "var(--psa-red)" }}>{fmtP(v)}</span>
                </div>
              ))
            : null
          }
          <div style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", marginTop: 2, borderTop: "1px solid var(--border)", fontSize: "0.83rem", fontWeight: 700, color: "var(--text)" }}>
            <span>Total Deductions</span><span style={{ fontVariantNumeric: "tabular-nums", color: "var(--psa-red)" }}>{fmtP(record.total_deductions)}</span>
          </div>

          {/* NET / TAKE HOME */}
          <div style={{ marginTop: 16, borderTop: "2px solid var(--border)", paddingTop: 12 }}>
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: "0.83rem", color: "var(--text)", marginBottom: 8 }}>
              <span>Net Pay</span><span style={{ fontVariantNumeric: "tabular-nums", fontWeight: 600 }}>{fmtP(record.net_pay)}</span>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline" }}>
              <span style={{ fontSize: "0.88rem", fontWeight: 700, color: "var(--text)" }}>
                Take Home <span style={{ fontSize: "0.72rem", fontWeight: 400, color: "var(--text-muted)" }}>per Quincena</span>
              </span>
              <span style={{ fontSize: "1.1rem", fontWeight: 800, color: "var(--psa-navy)", fontVariantNumeric: "tabular-nums" }}>{fmtP(record.first_quincena != null ? Number(record.first_quincena) : 0)}</span>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// READ-ONLY PAGE — never add POST/PUT/PATCH/DELETE here.
// Shows the linked employee's personal record, deduction ledger history, and position history.
// All data is fetched from GET endpoints only. This is a historical/audit display for the employee.
function EmployeeStatusPanel({ emp, posHistory, suspHistory, showAllSusp, setShowAllSusp }) {
  const activePos = posHistory?.find(p => p.status === "active") ?? posHistory?.[0] ?? null;
  const [expandedRec, setExpandedRec] = useState(null);
  const fmtDT   = d => d ? new Date(d).toLocaleString("en-PH", { year: "numeric", month: "short", day: "numeric", hour: "2-digit", minute: "2-digit" }) : null;
  const fmtDate = d => d ? new Date(d).toLocaleDateString("en-PH", { year: "numeric", month: "short", day: "numeric" }) : null;
  const fmtDuration = (hrs, mins) => {
    const h = Number(hrs) || 0, m = Number(mins) || 0;
    if (h < 1) return m > 0 ? `${m} min` : "< 1 hr";
    if (h < 24) return m > 0 ? `${h} hr ${m} min` : `${h} hr`;
    const days = Math.floor(h / 24), remH = h % 24;
    return remH > 0 ? `${days} day${days !== 1 ? "s" : ""} ${remH} hr` : `${days} day${days !== 1 ? "s" : ""}`;
  };
  const sRow = (label, value, valueStyle = {}) => (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 0", gap: 12 }}>
      <span style={{ fontSize: "0.92rem", fontWeight: 500, color: "var(--text-muted)", flexShrink: 0 }}>{label}</span>
      <span style={{ fontSize: "0.95rem", fontWeight: 700, color: "var(--text)", textAlign: "right", ...valueStyle }}>{value}</span>
    </div>
  );
  const empActive   = emp?.status === "active";
  const currentSusp = suspHistory?.[0] ?? null;
  const isOngoing   = currentSusp && !currentSusp.end_date;
  const pastRecs    = empActive ? (suspHistory ?? []) : (suspHistory?.slice(1) ?? []);
  const SUSP_PAGE   = 5;
  const suspPages   = Math.ceil(pastRecs.length / SUSP_PAGE) || 1;
  const safeSusp    = Math.min(showAllSusp, suspPages - 1);
  const visiblePast = pastRecs.slice(safeSusp * SUSP_PAGE, (safeSusp + 1) * SUSP_PAGE);
  return (
    <div style={{ padding: "18px 20px" }}>
      <div style={{ fontSize: "0.68rem", fontWeight: 800, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: 6, paddingBottom: 8, borderBottom: "1px solid var(--border)" }}>Employee</div>
      {sRow("Full Name", emp?.name ?? "—", { fontWeight: 800 })}
      {sRow("Division", emp?.office_division || "—")}
      {sRow("Position", activePos?.position_title ?? emp?.position ?? "—")}
      {activePos?.start_date && sRow("Appointed", new Date(activePos.start_date).toLocaleDateString("en-PH", { year: "numeric", month: "long", day: "numeric" }))}
      <div style={{ fontSize: "0.68rem", fontWeight: 800, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.1em", marginTop: 20, marginBottom: 6, paddingBottom: 8, borderBottom: "1px solid var(--border)" }}>Employment Status</div>
      {sRow("Status",
        empActive ? "Active" : (emp?.separation_reason ?? "Inactive"),
        { fontWeight: 800, color: empActive ? "#166534" : emp?.separation_reason === "Suspended" ? "#854d0e" : "#1e40af" }
      )}
      {!empActive && currentSusp && (
        <>
          {sRow("Type", currentSusp.reason)}
          {sRow("Start", fmtDT(currentSusp.start_date) ?? "—")}
          {sRow(isOngoing ? "Return Date" : "End", fmtDT(currentSusp.end_date) ?? "Open-ended")}
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 0", gap: 12 }}>
            <span style={{ fontSize: "0.92rem", fontWeight: 500, color: "var(--text-muted)", flexShrink: 0 }}>{isOngoing ? "Duration So Far" : "Duration"}</span>
            <span style={{ fontSize: "0.95rem", fontWeight: 700, color: "var(--psa-navy)", fontVariantNumeric: "tabular-nums" }}>{fmtDuration(currentSusp.total_hours, currentSusp.extra_minutes)}</span>
          </div>
        </>
      )}
      {suspHistory === null ? (
        <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "12px 0", color: "var(--text-muted)", fontSize: "0.83rem" }}>
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ animation: "spin 1s linear infinite", flexShrink: 0 }}><path d="M21 12a9 9 0 1 1-6.219-8.56"/></svg>
          Loading…
        </div>
      ) : empActive && pastRecs.length === 0 ? (
        <div style={{ padding: "9px 0", fontSize: "0.88rem", color: "var(--text-muted)" }}>No leave or suspension records.</div>
      ) : pastRecs.length > 0 ? (
        <>
          <div style={{ fontSize: "0.68rem", fontWeight: 800, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.1em", marginTop: 20, marginBottom: 6, paddingBottom: 8, borderBottom: "1px solid var(--border)" }}>
            {empActive ? "Records" : "Past Records"}{` (${pastRecs.length})`}
          </div>
          {visiblePast.map((rec, i) => {
            const isOpen = expandedRec === i;
            const accentColor = rec.reason === "Suspended" ? "#854d0e" : "#1e40af";
            const endLabel = rec.end_date ? fmtDate(rec.end_date) : "Ongoing";
            return (
              <div key={i}
                onClick={() => setExpandedRec(isOpen ? null : i)}
                style={{ padding: "11px 0", borderTop: i === 0 ? "none" : "1px solid var(--border)", cursor: "pointer", WebkitTapHighlightColor: "transparent" }}
              >
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12 }}>
                  <span style={{ fontSize: "0.95rem", fontWeight: 700, color: accentColor }}>{rec.reason}</span>
                  <span style={{ fontSize: "0.82rem", fontWeight: 500, color: "var(--text-muted)", textAlign: "right" }}>
                    {fmtDate(rec.start_date) ?? "—"}
                  </span>
                </div>
                {isOpen && (
                  <div style={{ marginTop: 10, display: "flex", flexDirection: "column", gap: 4 }}>
                    <div style={{ display: "flex", justifyContent: "space-between" }}>
                      <span style={{ fontSize: "0.82rem", color: "var(--text-muted)" }}>Started</span>
                      <span style={{ fontSize: "0.82rem", fontWeight: 600, color: "var(--text)" }}>{fmtDate(rec.start_date) ?? "—"}</span>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between" }}>
                      <span style={{ fontSize: "0.82rem", color: "var(--text-muted)" }}>{rec.end_date ? "Ended" : "Return Date"}</span>
                      <span style={{ fontSize: "0.82rem", fontWeight: 600, color: "var(--text)" }}>{endLabel}</span>
                    </div>
                    {(rec.total_hours > 0 || rec.extra_minutes > 0) && (
                      <div style={{ display: "flex", justifyContent: "space-between" }}>
                        <span style={{ fontSize: "0.82rem", color: "var(--text-muted)" }}>Duration</span>
                        <span style={{ fontSize: "0.82rem", fontWeight: 700, color: "var(--psa-navy)" }}>{fmtDuration(rec.total_hours, rec.extra_minutes)}</span>
                      </div>
                    )}
                  </div>
                )}
              </div>
            );
          })}
          {suspPages > 1 && (
            <div style={{ display: "flex", alignItems: "center", justifyContent: "flex-end", gap: 8, paddingTop: 8, marginTop: 4, borderTop: "1px solid var(--border)" }}>
              <button onClick={() => setShowAllSusp(p => Math.max(0, p - 1))} disabled={safeSusp === 0}
                style={{ background: "none", border: "1px solid var(--border)", borderRadius: 6, padding: "5px 14px", fontSize: "0.9rem", cursor: safeSusp === 0 ? "default" : "pointer", color: safeSusp === 0 ? "var(--text-muted)" : "var(--psa-navy)", fontWeight: 700 }}>‹</button>
              <span style={{ fontSize: "0.75rem", color: "var(--text-muted)", minWidth: 36, textAlign: "center" }}>{safeSusp + 1} / {suspPages}</span>
              <button onClick={() => setShowAllSusp(p => Math.min(suspPages - 1, p + 1))} disabled={safeSusp === suspPages - 1}
                style={{ background: "none", border: "1px solid var(--border)", borderRadius: 6, padding: "5px 14px", fontSize: "0.9rem", cursor: safeSusp === suspPages - 1 ? "default" : "pointer", color: safeSusp === suspPages - 1 ? "var(--text-muted)" : "var(--psa-navy)", fontWeight: 700 }}>›</button>
            </div>
          )}
        </>
      ) : null}
    </div>
  );
}

const LOANS_PER_PAGE = 2;
function MyLoansPanel({ loans, tab }) {
  const [page, setPage] = useState(0);
  const fmtP = n => `₱${Number(n).toLocaleString("en-PH", { minimumFractionDigits: 2 })}`;
  const activeLoans = loans.filter(l => l.status === "active");
  const closedLoans = loans.filter(l => l.status !== "active");
  const visible = tab === "active" ? activeLoans : closedLoans;
  const totalPages = Math.ceil(visible.length / LOANS_PER_PAGE) || 1;
  const safePage = Math.min(page, totalPages - 1);
  const pageLoans = visible.slice(safePage * LOANS_PER_PAGE, (safePage + 1) * LOANS_PER_PAGE);
  return (
    <div>
      <div style={{ padding: "10px 14px", display: "flex", flexDirection: "column", gap: 10 }}>
        {visible.length === 0 ? (
          <div style={{ padding: "24px 0", textAlign: "center", color: "var(--text-muted)", fontSize: "0.82rem" }}>
            No {tab} loans.
          </div>
        ) : pageLoans.map(loan => {
          const remaining = Number(loan.remaining_balance ?? 0);
          const total = Number(loan.total_amount ?? 0);
          const monthly = Number(loan.monthly_amount ?? 0);
          const paid = total - remaining;
          const pct = total > 0 ? Math.max(0, Math.min(100, (paid / total) * 100)) : 0;
          const isActive = loan.status === "active";
          const fmtD = d => d ? new Date(d).toLocaleDateString("en-PH", { year: "numeric", month: "short", day: "numeric" }) : "—";
          return (
            <div key={loan.id} style={{ borderBottom: "1px solid var(--border)", paddingBottom: 12 }}>
              <div style={{ marginBottom: 6 }}>
                <span style={{ fontSize: "0.72rem", fontWeight: 700, color: "var(--text)", textTransform: "uppercase", letterSpacing: "0.07em" }}>
                  {loan.loan_type_name ?? "Loan"}
                </span>
              </div>
              <div style={{ display: "flex", gap: 16, marginBottom: 8 }}>
                <div style={{ fontSize: "0.72rem", color: "var(--text-muted)" }}>
                  Started <span style={{ color: "var(--text)", fontWeight: 600 }}>{fmtD(loan.start_date)}</span>
                </div>
                {!isActive && loan.closed_date && (
                  <div style={{ fontSize: "0.72rem", color: "var(--text-muted)" }}>
                    Closed <span style={{ color: "var(--text)", fontWeight: 600 }}>{fmtD(loan.closed_date)}</span>
                  </div>
                )}
              </div>
              {[["Total Loan", fmtP(total)], ["Monthly Payment", fmtP(monthly)], ["Remaining Balance", fmtP(remaining)]].map(([lbl, val]) => (
                <div key={lbl} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "3px 0" }}>
                  <span style={{ fontSize: "0.8rem", color: "var(--text-muted)" }}>{lbl}</span>
                  <span style={{ fontSize: "0.8rem", fontWeight: 600, color: "var(--text)", fontVariantNumeric: "tabular-nums" }}>{val}</span>
                </div>
              ))}
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 8 }}>
                <div style={{ flex: 1, height: 3, borderRadius: 999, background: "var(--surface-2)", overflow: "hidden" }}>
                  <div style={{ height: "100%", width: `${pct}%`, borderRadius: 999, background: isActive ? "var(--psa-navy)" : "var(--text-muted)", transition: "width 0.4s" }} />
                </div>
                <span style={{ fontSize: "0.65rem", fontWeight: 700, color: "var(--text-muted)", minWidth: 36, textAlign: "right" }}>{Math.round(pct)}% paid</span>
              </div>
            </div>
          );
        })}
      </div>
      {totalPages > 1 && (
        <div style={{ display: "flex", alignItems: "center", justifyContent: "flex-end", gap: 8, padding: "6px 14px 10px", borderTop: "1px solid var(--border)" }}>
          <button onClick={() => setPage(p => Math.max(0, p - 1))} disabled={safePage === 0}
            style={{ background: "none", border: "1px solid var(--border)", borderRadius: 6, padding: "3px 12px", fontSize: "0.85rem", cursor: safePage === 0 ? "default" : "pointer", color: safePage === 0 ? "var(--text-muted)" : "var(--psa-navy)", fontWeight: 700 }}>‹</button>
          <span style={{ fontSize: "0.72rem", color: "var(--text-muted)", minWidth: 36, textAlign: "center" }}>{safePage + 1} / {totalPages}</span>
          <button onClick={() => setPage(p => Math.min(totalPages - 1, p + 1))} disabled={safePage === totalPages - 1}
            style={{ background: "none", border: "1px solid var(--border)", borderRadius: 6, padding: "3px 12px", fontSize: "0.85rem", cursor: safePage === totalPages - 1 ? "default" : "pointer", color: safePage === totalPages - 1 ? "var(--text-muted)" : "var(--psa-navy)", fontWeight: 700 }}>›</button>
        </div>
      )}
    </div>
  );
}

// See CLAUDE.md "STRICT READ-ONLY PAGES" for the full rule.
function EmployeeStatusView({ refreshKey, hasAvatar, avatarBlobUrl }) {
  const apiUrl = import.meta.env.VITE_API_URL ?? "/api";
  const [emp, setEmp]             = useState(null);
  const [loading, setLoading]     = useState(true);
  const [posHistory, setPosHistory] = useState(null);
  const [myDeds, setMyDeds]       = useState(null);
  const [dedsError, setDedsError] = useState(false);
  const [suspHistory, setSuspHistory] = useState(null);
  const [showAllSusp, setShowAllSusp] = useState(0);
  const [dedPage,     setDedPage]     = useState(0);
  const [historyTab, setHistoryTab] = useState("ledger");
  const [myLoans, setMyLoans] = useState(null);
  const [myLoansTab, setMyLoansTab] = useState("active");
  const [myLoanPayments, setMyLoanPayments] = useState(null);
  const isMobile = useIsMobile();
  const [mobileSection, setMobileSection] = useState("payslip");
  const effectiveHistoryTab = isMobile ? (mobileSection === "status" ? "status" : "ledger") : historyTab;

  useEffect(() => {
    setLoading(true);
    fetch(`${apiUrl}/me/dashboard`, { credentials: "include" })
      .then(r => r.ok ? r.json() : null)
      .then(d => { setEmp(d); setLoading(false); })
      .catch(() => setLoading(false));
  }, [refreshKey]);

  useEffect(() => {
    if (!emp?.id) return;
    setPosHistory(null);
    setMyDeds(null);
    setDedsError(false);
    setSuspHistory(null);
    setMyLoans(null);
    setShowAllSusp(0);
    setDedPage(0);
    fetch(`${apiUrl}/me/suspension`, { credentials: "include" })
      .then(r => r.ok ? r.json() : [])
      .then(setSuspHistory)
      .catch(() => setSuspHistory([]));
    fetch(`${apiUrl}/employees/${emp.id}/position-history`, { credentials: "include" })
      .then(r => r.ok ? r.json() : [])
      .then(setPosHistory)
      .catch(() => setPosHistory([]));
    /* READ-ONLY: ?readonly=1 prevents syncComputedDeductionsToLedger from running */
    fetch(`${apiUrl}/employees/${emp.id}/deductions?readonly=1`, { credentials: "include" })
      .then(r => r.ok ? r.json() : null)
      .then(d => { setMyDeds(d); setDedsError(!d); })
      .catch(() => { setMyDeds(null); setDedsError(true); });
    fetch(`${apiUrl}/employees/${emp.id}/loans`, { credentials: "include" })
      .then(r => r.ok ? r.json() : [])
      .then(setMyLoans)
      .catch(() => setMyLoans([]));
    fetch(`${apiUrl}/employees/${emp.id}/loan-payments`, { credentials: "include" })
      .then(r => r.ok ? r.json() : [])
      .then(setMyLoanPayments)
      .catch(() => setMyLoanPayments([]));
  }, [emp?.id, refreshKey]);

  const initials  = emp?.name ? emp.name.split(" ").filter(Boolean).slice(0, 2).map(w => w[0]).join("").toUpperCase() : "?";
  const avatarUrl = avatarBlobUrl ?? null;

  /* Payslip calculations — computed once, used in stat strip + breakdown */
  const _sides        = myDeds?._sides ?? {};
  const _skip         = new Set(["first_quincena","basic_pay","_sides"]);
  const dedEntries    = myDeds ? Object.entries(myDeds).filter(([k, v]) => !_skip.has(k) && Number(v) > 0 && (_sides[k] ?? "deduction") === "deduction") : [];
  const lessEntries   = myDeds ? Object.entries(myDeds).filter(([k, v]) => !_skip.has(k) && Number(v) > 0 && _sides[k] === "less_deduction") : [];
  const grossEntries  = myDeds ? Object.entries(myDeds).filter(([k, v]) => !_skip.has(k) && Number(v) > 0 && _sides[k] === "gross") : [];
  const totalDed      = dedEntries.reduce((s, [, v]) => s + Number(v), 0);
  const lessDedTotal  = lessEntries.reduce((s, [, v]) => s + Number(v), 0);
  const grossPay      = Number(emp?.basic_monthly_pay ?? 0) + grossEntries.reduce((s, [, v]) => s + Number(v), 0);
  const netPay        = grossPay - totalDed - lessDedTotal;
  const toLabel       = s => s.replace(/_/g, " ").replace(/\b\w/g, c => c.toUpperCase());

  const todayDay      = new Date().getDate();
  const isFirstQ      = todayDay <= 15;
  const q1Amount      = emp?.first_quincena_override != null
    ? Math.round(Number(emp.first_quincena_override) * 100) / 100
    : Math.round(netPay / 2 * 100) / 100;
  const q2Amount      = Math.round((netPay - q1Amount) * 100) / 100;
  const currentQAmt   = isFirstQ ? q1Amount : q2Amount;
  const currentQLabel = isFirstQ ? "1st Quincena" : "2nd Quincena";
  const now           = new Date();
  const monthShort    = now.toLocaleDateString("en-PH", { month: "short" });
  const lastDay       = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
  const periodRange   = isFirstQ ? `${monthShort} 1–15, ${now.getFullYear()}` : `${monthShort} 16–${lastDay}, ${now.getFullYear()}`;

  return (
    <div style={{ maxWidth: "100%" }}>
      {loading ? (
        <div style={{ padding: "80px 24px", textAlign: "center", color: "var(--text-muted)" }}>
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ animation: "spin 1s linear infinite", display: "block", margin: "0 auto 12px" }}>
            <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
          </svg>
          <div style={{ fontSize: "0.82rem" }}>Loading…</div>
        </div>
      ) : !emp ? (
        <div className="panel" style={{ textAlign: "center" }}>
          <div style={{ fontSize: "1rem", fontWeight: 600, color: "var(--text)", marginBottom: 6 }}>No employee record linked</div>
          <div style={{ fontSize: "0.85rem", color: "var(--text-muted)" }}>Ask your administrator to link your account to your employee record.</div>
        </div>
      ) : (
        <div className="emp-dash">

          {/* ── BANNER ── */}
          <div className="emp-banner emp-banner-card">
            {/* Philippine flag strip — matches sidebar exactly */}
            <div style={{ height: 3, background: "linear-gradient(90deg, #0038a8 0% 33.3%, #fcd116 33.3% 66.6%, #ce1126 66.6% 100%)" }} />
            <div className="emp-banner-body">
              <div className="emp-agency">
                <span>Philippine Statistics Authority · RSSO II</span>
                <span className="emp-agency-sep" aria-hidden="true" />
                <span className="emp-agency-sub">Employee Self-Service Portal</span>
              </div>
              {/* Profile row */}
              <div className="emp-banner-profile">
                <div style={{ display: "flex", alignItems: "center", gap: 18 }}>
                  {/* Rounded-rect avatar — modern card style */}
                  <div style={{ width: 72, height: 72, borderRadius: 12, overflow: "hidden", background: "var(--psa-blue)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: "1.6rem", fontWeight: 800, color: "rgba(255,255,255,0.9)", flexShrink: 0, border: "1px solid rgba(255,255,255,0.1)" }}>
                    {avatarUrl ? <img src={avatarUrl} alt="Avatar" style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : initials}
                  </div>
                  <div>
                    <div style={{ fontSize: "1.5rem", fontWeight: 800, color: "#fff", lineHeight: 1.1, marginBottom: 4, letterSpacing: "-0.02em" }}>{emp.name}</div>
                    <div style={{ fontSize: "0.86rem", color: "rgba(255,255,255,0.68)", marginBottom: 12 }}>{emp.position || "—"}</div>
                    <div className="emp-chips">
                      {[
                        emp.employee_number ? ["Employee no.", `#${emp.employee_number}`] : null,
                        emp.employment_type ? ["Type", emp.employment_type] : null,
                        emp.office_division ? ["Division", emp.office_division] : null,
                        emp.salary_rank ? ["Salary grade", emp.salary_rank] : null,
                      ].filter(Boolean).map(([label, value]) => (
                        <span key={label} className="emp-chip" title={label}>{value}</span>
                      ))}
                    </div>
                  </div>
                </div>
                {/* Basic Monthly Pay — white, clean, minimal */}
                <div className="emp-banner-pay">
                  <div style={{ fontSize: "0.62rem", fontWeight: 700, color: "rgba(255,255,255,0.6)", textTransform: "uppercase", letterSpacing: "0.12em", marginBottom: 6 }}>Basic Monthly Pay</div>
                  <div style={{ fontSize: "1.85rem", fontWeight: 800, color: "#fff", letterSpacing: "-0.02em", lineHeight: 1 }}>₱ {fmt(emp.basic_monthly_pay ?? 0)}</div>
                  <div className="emp-period">Pay period · {currentQLabel}, {periodRange}</div>
                </div>
              </div>

            </div>
          </div>

          {/* ── MOBILE SECTION TABS (hidden on desktop) ── */}
          {isMobile && (
            <div className="emp-section-tabs">
              <WsTabs small label="Dashboard sections" value={mobileSection} onChange={setMobileSection}
                tabs={[{ id: "payslip", label: "Payslip" }, { id: "loans", label: "Loans" }, { id: "status", label: "Status" }]} />
            </div>
          )}

          {/* ── PAYSLIP + HISTORY — two separate cards ── */}
          <div className="emp-cols">

            {/* LEFT — Payslip (flat receipt on a plain sheet) */}
            <section className="emp-sheet" aria-label="Current payslip" style={{ display: isMobile && mobileSection !== "payslip" ? "none" : undefined }}>
              {!myDeds && !dedsError && <WsSpinner label="Loading payslip…" />}
              {!myDeds && dedsError && <div className="ws-empty">Couldn't load your payslip. Refresh the page, or contact your administrator if this keeps happening.</div>}
              {myDeds && (
                <div className="emp-payslip">
                  <div className="ps-head">
                    <span className="ps-head-title">Current Payslip</span>
                    <span className="ps-head-period">{currentQLabel} · {periodRange}</span>
                  </div>

                  {/* Earnings */}
                  <div className="ps-label">Earnings</div>
                  <div className="ps-row">
                    <span>Basic Pay</span>
                    <span className="ps-amt">₱{fmt(emp.basic_monthly_pay ?? 0)}</span>
                  </div>
                  {grossEntries.map(([key, val]) => (
                    <div key={key} className="ps-row">
                      <span>{toLabel(key)}</span>
                      <span className="ps-amt">₱{fmt(val)}</span>
                    </div>
                  ))}
                  <div className="ps-row ps-sub">
                    <span>Gross Pay</span>
                    <span className="ps-amt">₱{fmt(grossPay)}</span>
                  </div>

                  {/* Deductions & Contributions */}
                  <div className="ps-label ps-section">Deductions &amp; Contributions</div>
                  {(dedEntries.length > 0 || lessEntries.length > 0) ? (
                    <>
                      {[...lessEntries, ...dedEntries].map(([key, val]) => (
                        <div key={key} className="ps-row">
                          <span>{toLabel(key)}</span>
                          <span className="ps-amt ps-neg">₱{fmt(val)}</span>
                        </div>
                      ))}
                      <div className="ps-row ps-sub">
                        <span>Total Deductions</span>
                        <span className="ps-amt ps-neg">₱{fmt(totalDed + lessDedTotal)}</span>
                      </div>
                    </>
                  ) : (
                    <div className="ps-row" style={{ color: "var(--text-muted)" }}>No deductions on record.</div>
                  )}

                  {/* Net / Take Home */}
                  <div className="ps-total">
                    <div className="ps-row" style={{ marginBottom: 2 }}>
                      <span>Net Pay</span>
                      <span className="ps-amt" style={{ fontWeight: 600 }}>₱{fmt(netPay)}</span>
                    </div>
                    <div className="ps-final">
                      <span className="ps-final-label">Take Home</span>
                      <span className="ps-final-value">₱{fmt(currentQAmt)}</span>
                    </div>
                  </div>
                </div>
              )}
            </section>

            {/* RIGHT — History (loan ledger + status) */}
            <section className="emp-ledger-card" aria-label="History" style={{ display: isMobile && mobileSection === "payslip" ? "none" : undefined }}>
              {/* Desktop inner tab bar — replaced by mobile top tabs on small screens */}
              {!isMobile && (
                <div className="emp-history-head">
                  <span className="ws-kicker">History</span>
                  <WsTabs small label="History" value={historyTab} onChange={setHistoryTab}
                    tabs={[{ id: "ledger", label: "Loan payments" }, { id: "status", label: "Status" }]} />
                </div>
              )}
              <div style={{ flex: 1 }}>
                {effectiveHistoryTab === "ledger" && (
                  <div>
                    {myLoanPayments === null ? (
                      <div style={{ padding: "40px 0", textAlign: "center", color: "var(--text-muted)" }}>
                        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ animation: "spin 1s linear infinite", display: "block", margin: "0 auto 10px" }}><path d="M21 12a9 9 0 1 1-6.219-8.56"/></svg>
                        <div style={{ fontSize: "0.82rem" }}>Loading…</div>
                      </div>
                    ) : (
                      <LedgerAccordion
                        loanPayments={myLoanPayments ?? []}
                        loans={myLoans ?? []}
                        tab={myLoansTab}
                        setTab={setMyLoansTab}
                      />
                    )}
                  </div>
                )}
                {effectiveHistoryTab === "status" && (
                  <EmployeeStatusPanel
                    emp={emp}
                    posHistory={posHistory}
                    suspHistory={suspHistory}
                    showAllSusp={showAllSusp}
                    setShowAllSusp={setShowAllSusp}
                  />
                )}
              </div>
            </section>

          </div>
        </div>
      )}
    </div>
  );
}

// ── 201 File document types ───────────────────────────────────────────────────
const DOCS_201 = [
  { item: "A", name: "Appointment (CS FORM 33)" },
  { item: "B", name: "Oath of Office" },
  { item: "C", name: "Certificate of Assumption to Duty" },
  { item: "D", name: "Position Description Form (PDF)" },
  { item: "E", name: "Personal Data Sheet (PDS)" },
  { item: "F", name: "NOSA / NOSI" },
  { item: "G", name: "Certificate of Eligibility" },
  { item: "H", name: "Transcript of Records / Diploma" },
  { item: "I", name: "Service Record" },
  { item: "J", name: "Certificate of Leave Balance" },
  { item: "K", name: "SALN" },
  { item: "L", name: "Marriage Contract" },
  { item: "M", name: "Medical Certificate" },
  { item: "N", name: "Clearances" },
  { item: "O", name: "Special Orders / Memorandum" },
  { item: "P", name: "Certificate of Training" },
  { item: "Q", name: "Others" },
];

const MIME_BY_EXT = { jpg: "image/jpeg", jpeg: "image/jpeg", png: "image/png", webp: "image/webp", gif: "image/gif", pdf: "application/pdf" };
const mimeFromName = name => MIME_BY_EXT[(name ?? "").split(".").pop()?.toLowerCase()] ?? "";

function Files201Tab({ emp, basePath }) {
  const base = basePath ?? `${apiUrl}/employees/${emp.id}/201-files`;
  const isMobile = useIsMobile();
  const [files, setFiles] = useState({});
  const [loading, setLoading] = useState(true);
  const [uploading, setUploading] = useState({});
  const [uploadMode, setUploadMode] = useState({});   // item -> "add" | "replace"
  const [collapsed, setCollapsed] = useState({});
  const [query, setQuery] = useState("");
  const [preview, setPreview] = useState(null);   // { src, filename, mimeType, docName }
  const [zoomed, setZoomed] = useState(false);
  const [previewError, setPreviewError] = useState("");

  const loadFiles = useCallback(async () => {
    try {
      const res = await fetch(base, { credentials: "include" });
      if (!res.ok) return;
      const rows = await res.json();
      const grouped = {};
      for (const r of rows) {
        if (!grouped[r.docItem]) grouped[r.docItem] = [];
        grouped[r.docItem].push(r);
      }
      setFiles(grouped);
    } finally { setLoading(false); }
  }, [base]);

  useEffect(() => { loadFiles(); }, [loadFiles]);

  const closePreview = useCallback(() => {
    setPreview(null);
    setZoomed(false);
    setPreviewError("");
  }, []);

  // Esc closes the viewer; page scroll is locked while it is open
  useEffect(() => {
    if (!preview) return;
    const onKey = e => { if (e.key === "Escape") closePreview(); };
    window.addEventListener("keydown", onKey);
    const prevOverflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => { window.removeEventListener("keydown", onKey); document.body.style.overflow = prevOverflow; };
  }, [preview, closePreview]);

  async function handleUpload(docItem, file, mode = "add") {
    const existing = files[docItem] ?? [];
    // "Replace" means exactly that — the old files go, the new one takes their place.
    if (mode === "replace" && existing.length) {
      const what = existing.length === 1 ? `"${existing[0].filename}"` : `${existing.length} files`;
      if (!window.confirm(`Replace ${what} with "${file.name}"?`)) return;
    }
    setUploading(u => ({ ...u, [docItem]: true }));
    try {
      const fd = new FormData();
      fd.append("file", file);
      fd.append("docItem", docItem);
      const res = await fetch(base, { method: "POST", body: fd, credentials: "include" });
      if (!res.ok) { const b = await res.json().catch(() => ({})); alert(b.error ?? "Upload failed."); return; }
      // Only remove the old ones once the new file is safely stored.
      if (mode === "replace") {
        for (const old of existing) {
          await fetch(`${base}/${old.id}`, { method: "DELETE", credentials: "include" }).catch(() => {});
        }
      }
      await loadFiles();
    } finally { setUploading(u => { const n = { ...u }; delete n[docItem]; return n; }); }
  }

  // Opens the file in an overlay on this page — never navigates away.
  // The element loads the endpoint directly: the browser streams and decodes it
  // natively and sends the session cookie (same site, port is irrelevant), so
  // there is no multi-megabyte blob held in JS and nothing to revoke.
  function handleView(file, docName) {
    setPreviewError("");
    setZoomed(false);
    setPreview({
      src: `${base}/${file.id}/view`,
      filename: file.filename,
      mimeType: file.mimeType || mimeFromName(file.filename) || "",
      docName,
    });
  }

  async function handleDelete(fileId, filename) {
    if (!window.confirm(`Delete "${filename}"?`)) return;
    await fetch(`${base}/${fileId}`, { method: "DELETE", credentials: "include" });
    await loadFiles();
  }

  const totalUploaded = Object.values(files).reduce((s, arr) => s + arr.length, 0);

  // Search matches the item letter, the document name, or any uploaded file name
  const q = query.trim().toLowerCase();
  const visibleDocs = useMemo(() => {
    if (!q) return DOCS_201;
    return DOCS_201.filter(({ item, name }) =>
      item.toLowerCase() === q ||
      name.toLowerCase().includes(q) ||
      (files[item] ?? []).some(f => (f.filename ?? "").toLowerCase().includes(q))
    );
  }, [q, files]);

  if (loading) return (
    <div style={{ padding: "80px 24px", textAlign: "center", color: "var(--text-muted)" }}>
      <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ animation: "spin 1s linear infinite", display: "block", margin: "0 auto 12px" }}>
        <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
      </svg>
      <div style={{ fontSize: "0.82rem" }}>Loading…</div>
    </div>
  );

  const isImage = !!preview && (preview.mimeType.startsWith("image/") || /\.(jpe?g|png|webp|gif)$/i.test(preview.filename));

  return (
    <div style={{ padding: "12px 0" }}>
      {/* ── Search + progress ─────────────────────────────────────────────── */}
      <div style={{ padding: isMobile ? "0 16px 12px" : "0 16px 14px", display: "flex", flexDirection: isMobile ? "column" : "row", gap: 8, alignItems: isMobile ? "stretch" : "center" }}>
        <div style={{ position: "relative", flex: 1 }}>
          <svg style={{ position: "absolute", left: 12, top: "50%", transform: "translateY(-50%)", pointerEvents: "none" }} width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#a8bcd4" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
          </svg>
          <input
            className="form-input"
            style={{ borderRadius: "var(--radius)", border: "1.5px solid var(--border-2)", padding: "10px 34px 10px 36px", fontSize: "0.9rem", width: "100%", boxSizing: "border-box" }}
            placeholder="Search documents or file names…"
            value={query}
            onChange={e => setQuery(e.target.value)}
          />
          {query && (
            <button type="button" onClick={() => setQuery("")} title="Clear search"
              style={{ position: "absolute", right: 8, top: "50%", transform: "translateY(-50%)", background: "transparent", border: "none", boxShadow: "none",
                color: "var(--text-muted)", cursor: "pointer", fontSize: "0.95rem", lineHeight: 1, padding: "4px 6px" }}>✕</button>
          )}
        </div>
        <span style={{ fontSize: "0.78rem", color: "var(--text-muted)", flexShrink: 0, whiteSpace: "nowrap" }}>
          {q
            ? `${visibleDocs.length} of ${DOCS_201.length} shown`
            : `${totalUploaded} of ${DOCS_201.length} documents uploaded`}
        </span>
      </div>

      {/* ── Document checklist ────────────────────────────────────────────── */}
      {visibleDocs.length === 0 && (
        <div style={{ borderTop: "1px solid var(--border)", padding: "48px 24px", textAlign: "center", color: "var(--text-muted)" }}>
          <svg width="34" height="34" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" style={{ opacity: 0.35, marginBottom: 10 }}>
            <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
          </svg>
          <div style={{ fontSize: "0.86rem", fontWeight: 600, color: "var(--text)" }}>No documents match “{query.trim()}”</div>
          <div style={{ fontSize: "0.78rem", marginTop: 4 }}>Try a document name, a file name, or a single letter such as “A”.</div>
        </div>
      )}
      {visibleDocs.map(({ item, name }) => {
        const docFiles = files[item] ?? [];
        const done = docFiles.length > 0;
        // While searching, matching rows stay open so the file is visible
        const isOpen = done && (q ? true : !collapsed[item]);
        const inputId = `file-201-${emp.id}-${item}`;
        return (
          <div key={item} style={{ borderTop: "1px solid var(--border)" }}>
            {/* Row header — clicking collapses/expands when files exist */}
            <div
              style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 16px",
                cursor: done && !q ? "pointer" : "default", userSelect: "none" }}
              onClick={done && !q ? () => setCollapsed(c => ({ ...c, [item]: !c[item] })) : undefined}>
              <span style={{ width: 22, height: 22, borderRadius: "50%", display: "inline-flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
                background: done ? "#16a34a" : "transparent",
                border: `2px solid ${done ? "#16a34a" : "var(--border)"}` }}>
                {done && (
                  <svg width="12" height="12" viewBox="0 0 16 16" fill="none">
                    <path d="M3 8.5l3.2 3.2L13 5" stroke="#fff" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"/>
                  </svg>
                )}
              </span>
              <span style={{ flex: 1, fontSize: "0.82rem", fontWeight: done ? 600 : 400, color: done ? "var(--text)" : "var(--text-muted)" }}>
                <span style={{ color: "var(--psa-navy)", fontWeight: 700 }}>{item}.</span> {name}
              </span>
              {/* Chevron only when files exist */}
              {done && !q && (
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="var(--text-muted)" strokeWidth="2.2"
                  strokeLinecap="round" strokeLinejoin="round" style={{ flexShrink: 0, transform: isOpen ? "rotate(0deg)" : "rotate(-90deg)", transition: "transform 0.18s" }}>
                  <polyline points="6 9 12 15 18 9"/>
                </svg>
              )}
              <input type="file" id={inputId} style={{ display: "none" }}
                accept=".pdf,.jpg,.jpeg,.png,.webp,application/pdf,image/jpeg,image/png,image/webp"
                onChange={e => { const f = e.target.files?.[0]; if (f) handleUpload(item, f, uploadMode[item] ?? "add"); e.target.value = ""; }} />
              {uploading[item]
                ? <span style={{ fontSize: "0.73rem", color: "var(--text-muted)", flexShrink: 0 }}>Uploading…</span>
                : (() => {
                    const openPicker = (mode) => (e) => {
                      e.stopPropagation();
                      setUploadMode(m => ({ ...m, [item]: mode }));
                      document.getElementById(inputId)?.click();
                    };
                    const pill = (filled) => ({
                      fontSize: "0.73rem", fontWeight: 600, padding: "4px 12px", borderRadius: 999, flexShrink: 0, cursor: "pointer",
                      border: "1.5px solid var(--psa-navy)",
                      background: filled ? "var(--psa-navy)" : "transparent",
                      color: filled ? "#fff" : "var(--psa-navy)",
                    });
                    return done ? (
                      <span style={{ display: "flex", gap: 6, flexShrink: 0 }}>
                        <button type="button" onClick={openPicker("add")} title="Add another file to this document" style={pill(true)}>+ Add</button>
                        <button type="button" onClick={openPicker("replace")} title="Replace the current file" style={pill(false)}>Replace</button>
                      </span>
                    ) : (
                      <button type="button" onClick={openPicker("add")} style={pill(true)}>Upload</button>
                    );
                  })()
              }
            </div>
            {/* File list — only shown when expanded */}
            {isOpen && docFiles.map(f => (
              <div key={f.id} style={{ display: "flex", flexWrap: "wrap", alignItems: "center", gap: 8, marginBottom: 6,
                marginLeft: isMobile ? 16 : 46, marginRight: 16,
                padding: "7px 10px", background: "var(--surface-2)", borderRadius: 8, border: "1px solid var(--border)" }}>
                {/* Name keeps the whole first line on phones so the buttons stay tappable */}
                <div style={{ display: "flex", alignItems: "center", gap: 8, minWidth: 0, flex: isMobile ? "1 1 100%" : 1 }}>
                  <svg width="13" height="13" viewBox="0 0 24 24" fill="var(--psa-navy)" style={{ flexShrink: 0 }}>
                    <path d="M14 2H6c-1.1 0-2 .9-2 2v16c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V8l-6-6zm-1 7V3.5L18.5 9H13z"/>
                  </svg>
                  <span style={{ flex: 1, minWidth: 0, fontSize: "0.78rem", color: "var(--text)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{f.filename}</span>
                </div>
                <span style={{ fontSize: "0.68rem", color: "var(--text-muted)", flexShrink: 0, marginLeft: isMobile ? "auto" : 0 }}>{f.mimeType?.startsWith("image/") ? "Photo" : "PDF"}</span>
                <button type="button" onClick={() => handleView(f, name)}
                  style={{ fontSize: "0.73rem", fontWeight: 600, padding: "4px 12px", borderRadius: 999, flexShrink: 0,
                    border: "1.5px solid var(--psa-navy)", background: "transparent", color: "var(--psa-navy)",
                    cursor: "pointer", whiteSpace: "nowrap" }}>
                  View
                </button>
                <button type="button" onClick={() => handleDelete(f.id, f.filename)}
                  style={{ fontSize: "0.73rem", fontWeight: 600, padding: "4px 12px", borderRadius: 999, flexShrink: 0,
                    border: "1.5px solid var(--text-muted, #888)", background: "transparent", color: "var(--text-muted, #888)",
                    cursor: "pointer", whiteSpace: "nowrap" }}>
                  Delete
                </button>
              </div>
            ))}
          </div>
        );
      })}

      {/* ── In-page viewer ────────────────────────────────────────────────── */}
      {preview && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(13,35,70,0.45)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: isMobile ? 0 : 16, backdropFilter: "blur(2px)" }}
          onClick={e => { if (e.target === e.currentTarget) closePreview(); }}
        >
          <div style={{ background: "var(--surface)", borderRadius: isMobile ? 0 : 16, width: "100%", maxWidth: 920,
            height: isMobile ? "100%" : "88vh", display: "flex", flexDirection: "column", overflow: "hidden",
            boxShadow: "0 20px 60px rgba(0,0,0,0.25)" }}>
            {/* Header */}
            <div style={{ padding: "14px 14px 14px 20px", borderBottom: "1px solid var(--border)", display: "flex", alignItems: "center", gap: 10, flexShrink: 0 }}>
              <svg width="15" height="15" viewBox="0 0 24 24" fill="var(--psa-navy)" style={{ flexShrink: 0 }}>
                <path d="M14 2H6c-1.1 0-2 .9-2 2v16c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V8l-6-6zm-1 7V3.5L18.5 9H13z"/>
              </svg>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontSize: "0.92rem", fontWeight: 800, color: "var(--psa-navy)", letterSpacing: "-0.01em", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{preview.filename}</div>
                <div style={{ fontSize: "0.74rem", color: "var(--text-muted)", marginTop: 2, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{preview.docName}</div>
              </div>
              <a href={preview.src} download={preview.filename}
                style={{ fontSize: "0.73rem", fontWeight: 600, padding: "5px 13px", borderRadius: 999, flexShrink: 0, textDecoration: "none",
                  border: "1.5px solid var(--psa-navy)", background: "transparent", color: "var(--psa-navy)", whiteSpace: "nowrap" }}>
                Download
              </a>
              <button type="button" onClick={closePreview} title="Close (Esc)"
                style={{ padding: "4px 8px", background: "transparent", border: "none", color: "var(--text-muted)", cursor: "pointer", fontSize: "1.2rem", lineHeight: 1, boxShadow: "none", flexShrink: 0 }}
                onMouseEnter={e => { e.currentTarget.style.color = "#991b1b"; }}
                onMouseLeave={e => { e.currentTarget.style.color = "var(--text-muted)"; }}
              >✕</button>
            </div>
            {/* Body */}
            <div style={{ flex: 1, minHeight: 0, overflow: "auto", background: "var(--surface-2)",
              display: "flex", alignItems: zoomed ? "flex-start" : "center", justifyContent: zoomed ? "flex-start" : "center", padding: isImage ? 14 : 0 }}>
              {previewError ? (
                <div style={{ textAlign: "center", color: "var(--text-muted)", padding: 32, margin: "auto" }}>
                  <svg width="38" height="38" viewBox="0 0 24 24" fill="none" stroke="#b45309" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" style={{ marginBottom: 12 }}>
                    <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>
                  </svg>
                  <div style={{ fontSize: "0.88rem", fontWeight: 600, color: "var(--text)", marginBottom: 6 }}>This file could not be displayed</div>
                  <div style={{ fontSize: "0.8rem", lineHeight: 1.5, maxWidth: 380, margin: "0 auto" }}>{previewError}</div>
                  <a href={preview.src} target="_blank" rel="noopener noreferrer"
                    style={{ display: "inline-block", marginTop: 14, fontSize: "0.78rem", fontWeight: 600, color: "var(--psa-navy)" }}>
                    Open it in a new tab
                  </a>
                </div>
              ) : isImage ? (
                <img src={preview.src} alt={preview.filename}
                  onClick={() => setZoomed(z => !z)}
                  onLoad={() => setPreviewError("")}
                  onError={() => setPreviewError("The server refused the request or the file is unreadable.")}
                  title={zoomed ? "Tap to fit" : "Tap to zoom"}
                  style={zoomed
                    ? { maxWidth: "none", width: "auto", borderRadius: 8, display: "block", cursor: "zoom-out" }
                    : { maxWidth: "100%", maxHeight: "100%", objectFit: "contain", borderRadius: 8, display: "block", cursor: "zoom-in", margin: "auto" }} />
              ) : (
                <div style={{ width: "100%", height: "100%", display: "flex", flexDirection: "column" }}>
                  <iframe title={preview.filename} src={preview.src} style={{ flex: 1, width: "100%", border: "none", display: "block" }} />
                  {/* Many phone browsers refuse to render a PDF in an iframe — give them a way out */}
                  {isMobile && (
                    <div style={{ padding: "8px 14px", fontSize: "0.74rem", color: "var(--text-muted)", textAlign: "center", borderTop: "1px solid var(--border)", background: "var(--surface)" }}>
                      PDF not showing?{" "}
                      <a href={preview.src} target="_blank" rel="noopener noreferrer" style={{ color: "var(--psa-navy)", fontWeight: 600 }}>Open it here</a>
                    </div>
                  )}
                </div>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ── Files201View — each user sees ONLY their own 201 files ───────────────────
function Files201View() {
  const [employee, setEmployee] = useState(undefined); // undefined = loading, null = no link
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    fetch(`${apiUrl}/me/dashboard`, { credentials: "include" })
      .then(r => r.ok ? r.json() : null)
      .then(data => { setEmployee(data ?? null); setLoading(false); })
      .catch(() => { setEmployee(null); setLoading(false); });
  }, []);

  if (loading) return (
    <div style={{ padding: "80px 24px", textAlign: "center", color: "var(--text-muted)" }}>
      <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ animation: "spin 1s linear infinite", display: "block", margin: "0 auto 12px" }}>
        <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
      </svg>
      <div style={{ fontSize: "0.82rem" }}>Loading…</div>
    </div>
  );

  if (!employee) return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", height: "100%", minHeight: 400, color: "var(--text-muted)", textAlign: "center", padding: 40 }}>
      <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" strokeLinejoin="round" style={{ opacity: 0.3, marginBottom: 16 }}>
        <circle cx="12" cy="8" r="4"/><path d="M6 20v-2a6 6 0 0 1 12 0v2"/>
      </svg>
      <div style={{ fontSize: "1rem", fontWeight: 600, marginBottom: 6 }}>No employee record linked</div>
      <div style={{ fontSize: "0.84rem" }}>Your account is not linked to an employee record. Contact your administrator.</div>
    </div>
  );

  return (
    <div style={{ maxWidth: 720, margin: "0 auto", padding: "0 0 40px" }}>
      <div style={{ padding: "20px 20px 14px", borderBottom: "1px solid var(--border)" }}>
        <div style={{ fontSize: "1.05rem", fontWeight: 700, color: "var(--text)" }}>{employee.name}</div>
        <div style={{ fontSize: "0.78rem", color: "var(--text-muted)", marginTop: 2 }}>{employee.position} · {employee.office_division || "—"}</div>
      </div>
      <Files201Tab emp={employee} basePath={`${apiUrl}/me/201-files`} />
    </div>
  );
}

// ── EmployeeHubView (Employees + Compensation + Employee Accounts, merged) ────
// ── Shared workspace UI (Employees, Workforce Setup, Salary Rank, Archive) ────
const WS_TAB_ICONS = {
  directory: <><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></>,
  status:    <polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/>,
  positions: <><rect x="2" y="7" width="20" height="14" rx="2"/><path d="M16 7V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v2"/></>,
  reference: <><ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"/><path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"/></>,
  grades:    <><polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 17 12 22 22 17"/><polyline points="2 12 12 17 22 12"/></>,
  archive:   <><polyline points="21 8 21 21 3 21 3 8"/><rect x="1" y="3" width="22" height="5"/><line x1="10" y1="12" x2="14" y2="12"/></>,
  history:   <><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></>,
};

function WsTabs({ tabs, value, onChange, label, small = false }) {
  return (
    <div className={`st-tabs ws-tabs${small ? " st-tabs--sm" : ""}`} role="tablist" aria-label={label}>
      {tabs.map(t => {
        const icon = WS_TAB_ICONS[t.icon ?? t.id];
        return (
          <button key={t.id} type="button" role="tab" aria-selected={value === t.id}
            className={`st-tab${value === t.id ? " is-on" : ""}`} onClick={() => onChange?.(t.id)}>
            {icon && <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">{icon}</svg>}
            {t.label}
            {t.count > 0 && <span className="st-tab-badge">{t.count}</span>}
          </button>
        );
      })}
    </div>
  );
}

function WsSearch({ value, onChange, placeholder }) {
  return (
    <div className="um-search ws-search">
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
      <input type="text" value={value} onChange={e => onChange(e.target.value)} placeholder={placeholder} />
      {value && <button type="button" className="um-clear" onClick={() => onChange("")} aria-label="Clear search">×</button>}
    </div>
  );
}

const WsSpinner = ({ label = "Loading…" }) => (
  <div className="ws-empty">
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ animation: "spin 1s linear infinite" }} aria-hidden="true"><path d="M21 12a9 9 0 1 1-6.219-8.56"/></svg>
    {label}
  </div>
);

function EmployeeHubView({ isAdmin, isSuperAdmin, permissions = [], refreshKey }) {
  const can = (key) => isSuperAdmin || permissions.includes(key);
  const canDeductions = isAdmin || can("deductions.edit");
  const isMobile = useIsMobile();
  const { dedTypes, earnTypes } = useContext(TypesContext);

  const [hubTab, setHubTab] = useState("directory");

  // ── Directory (Employees + Compensation) ──
  const [employees, setEmployees] = useState([]);
  const [query, setQuery]         = useState("");
  const [selectedId, setSelectedId] = useState(null);
  const [openId, setOpenId]       = useState(null); // mobile accordion
  const [editEmp, setEditEmp]     = useState(null);
  const [showAdd, setShowAdd]     = useState(false);
  const [loading, setLoading]     = useState(false);
  const [fetchError, setFetchError] = useState(null);
  const [empTab, setEmpTab] = useState("payslip"); // "payslip" | "finance" | "records"
  const [recordsTab, setRecordsTab] = useState("loans"); // "loans" | "history"
  const payslipEditMode = empTab === "finance";
  const [payslipItem, setPayslipItem]         = useState(null);
  const [payslipLoading, setPayslipLoading]   = useState(false);
  const [embeddedTakeHome, setEmbeddedTakeHome] = useState(null);
  const [selSgSteps, setSelSgSteps]           = useState([]);
  const [exporting, setExporting]         = useState(false);
  const [exportingXlsx, setExportingXlsx] = useState(false);
  const [deactivatingEmp, setDeactivatingEmp]       = useState(null);
  const [separationReason, setSeparationReason]     = useState("On Leave");
  const [suspensionStart, setSuspensionStart]       = useState("");
  const [suspensionEnd, setSuspensionEnd]           = useState("");
  const [deactivating, setDeactivating]             = useState(false);
  const [selAvatar, setSelAvatar]                   = useState(null);
  const [avatarOpen, setAvatarOpen]                 = useState(false);
  const hasData = useRef(false);


  const fetchEmployees = useCallback(async (q = "", silent = false) => {
    if (!silent) setLoading(true);
    setFetchError(null);
    try {
      const limit = q.trim() ? 50 : 5;
      const res = await fetch(`${apiUrl}/employees?q=${encodeURIComponent(q)}&limit=${limit}`, { credentials: "include" });
      if (res.ok) {
        const json = await res.json();
        setEmployees(Array.isArray(json) ? json : (json.data ?? []));
        hasData.current = true;
      } else {
        const errText = await res.text().catch(() => res.status);
        setFetchError(`HTTP ${res.status}: ${errText}`);
      }
    } catch (e) {
      setFetchError(`Network error: ${e.message}`);
    } finally { setLoading(false); }
  }, []);

  useEffect(() => { fetchEmployees("", hasData.current); }, [fetchEmployees, refreshKey]);

  useEffect(() => {
    const t = setTimeout(() => fetchEmployees(query), 300);
    return () => clearTimeout(t);
  }, [query, fetchEmployees]);

  const selectedEmp = employees.find(e => e.id === selectedId) ?? null;

  // Profile photo comes from the employee's linked user account (uploaded in My Account).
  useEffect(() => {
    setSelAvatar(null);
    setAvatarOpen(false);
    if (!selectedId) return;
    let cancelled = false;
    fetch(`${apiUrl}/employees/${selectedId}/profile`, { credentials: "include" })
      .then(r => r.ok ? r.json() : null)
      .then(d => { if (!cancelled) setSelAvatar(d?.employee?.avatarDataUrl ?? null); })
      .catch(() => {});
    return () => { cancelled = true; };
  }, [selectedId]);

  useEffect(() => {
    if (!avatarOpen) return;
    const onKey = (e) => { if (e.key === "Escape") setAvatarOpen(false); };
    document.addEventListener("keydown", onKey);
    return () => document.removeEventListener("keydown", onKey);
  }, [avatarOpen]);

  useEffect(() => {
    setEmpTab("payslip");
    setSelSgSteps([]);
    if (!selectedEmp) { setPayslipItem(null); return; }
    setPayslipLoading(true);
    fetch(`${apiUrl}/employees/live-payslips?q=${encodeURIComponent(selectedEmp.name)}&limit=50`, { credentials: "include" })
      .then(r => r.ok ? r.json() : [])
      .then(list => setPayslipItem(Array.isArray(list) ? (list.find(i => i.employee_id === selectedEmp.id) ?? null) : null))
      .finally(() => setPayslipLoading(false));
    if (selectedEmp.salary_rank) {
      fetch(`${apiUrl}/salary-grades`, { credentials: "include" })
        .then(r => r.ok ? r.json() : [])
        .then(all => {
          const latest = {};
          for (const r of all.filter(r => r.salary_grade === selectedEmp.salary_rank)) {
            if (!latest[r.step] || r.effective_date > latest[r.step].effective_date) latest[r.step] = r;
          }
          const steps = Object.values(latest).sort((a, b) => a.step - b.step);
          setSelSgSteps(steps);
        }).catch(() => {});
    }
  }, [selectedEmp?.id, selectedEmp?.basic_monthly_pay, selectedEmp?.position_id]);

  const period = currentPayPeriod();

  async function handleExportDocx() {
    setExporting(true);
    try {
      const allRes = await fetch(`${apiUrl}/employees/live-payslips?q=${encodeURIComponent(query)}&limit=9999`, { credentials: "include" });
      const allItems = await allRes.json();
      const exportItems = Array.isArray(allItems) && allItems.length ? allItems : [];
      const signatories = await fetch(`${apiUrl}/settings/signatories`, { credentials: "include" }).then(r => r.ok ? r.json() : null);
      const lessDedCatalog = dedTypes.length > 0 ? buildCatalogFromTypes(dedTypes, 'less_deduction') : LESS_DEDUCTIONS_CATALOG;
      const dedCatalog     = dedTypes.length > 0 ? buildCatalogFromTypes(dedTypes, 'deduction')      : DEDUCTIONS_CATALOG;
      const blob = await generatePayslipsDocx(exportItems, period, signatories, lessDedCatalog, dedCatalog, earnTypes);
      const now = new Date();
      downloadBlob(blob, `Payslips - ${now.toLocaleDateString("en-PH", { month: "long", year: "numeric" })}.docx`);
    } catch { alert("Failed to generate payslip document."); }
    finally { setExporting(false); }
  }

  async function handleExportXlsx() {
    setExportingXlsx(true);
    try {
      const allRes = await fetch(`${apiUrl}/employees/live-payslips?q=${encodeURIComponent(query)}&limit=9999`, { credentials: "include" });
      const allItems = await allRes.json();
      const exportItems = Array.isArray(allItems) && allItems.length ? allItems : [];
      const signatories = await fetch(`${apiUrl}/settings/signatories`, { credentials: "include" }).then(r => r.ok ? r.json() : null);
      const lessDedCatalog = dedTypes.length > 0 ? buildCatalogFromTypes(dedTypes, 'less_deduction') : LESS_DEDUCTIONS_CATALOG;
      const dedCatalog     = dedTypes.length > 0 ? buildCatalogFromTypes(dedTypes, 'deduction')      : DEDUCTIONS_CATALOG;
      const blob = await generatePayslipsXlsx(exportItems, period, signatories, lessDedCatalog, dedCatalog, earnTypes);
      const now = new Date();
      downloadBlob(blob, `Payroll Register - ${now.toLocaleDateString("en-PH", { month: "long", year: "numeric" })}.xlsx`);
    } catch { alert("Failed to generate payroll register."); }
    finally { setExportingXlsx(false); }
  }

  const patchEmpInList = useCallback((empId, updates) => {
    setEmployees(prev => prev.map(e => e.id === empId ? { ...e, ...updates } : e));
  }, []);

  const bumpToTop = useCallback((empId) => {
    if (query.trim()) return;
    setEmployees(prev => {
      const idx = prev.findIndex(e => e.id === empId);
      if (idx <= 0) return prev;
      const next = [...prev];
      const [moved] = next.splice(idx, 1);
      return [moved, ...next];
    });
  }, [query]);

  function handleDelete(emp) {
    setSeparationReason("On Leave");
    const now = new Date();
    const local = new Date(now.getTime() - now.getTimezoneOffset() * 60000).toISOString().slice(0, 16);
    setSuspensionStart(local);
    setSuspensionEnd("");
    setDeactivatingEmp(emp);
  }

  async function confirmDeactivate() {
    if (!deactivatingEmp) return;
    const isTemp = ["On Leave", "Suspended"].includes(separationReason);
    if (isTemp && !suspensionStart) { alert("Please set a start date and time."); return; }
    setDeactivating(true);
    try {
      const res = await fetch(`${apiUrl}/employees/${deactivatingEmp.id}`, {
        method: "DELETE", credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          reason: separationReason,
          ...(isTemp && { suspensionStart, suspensionEnd: suspensionEnd || null }),
        }),
      });
      if (res.ok) {
        setDeactivatingEmp(null);
        fetchEmployees(query);
        if (selectedId === deactivatingEmp.id) setSelectedId(null);
        if (openId === deactivatingEmp.id) setOpenId(null);
      }
    } finally { setDeactivating(false); }
  }

  // ── Inactive Employees ──
  const [inactiveEmps, setInactiveEmps]       = useState([]);
  const [inactiveLoading, setInactiveLoading] = useState(true);
  const [activating, setActivating]           = useState(null);
  const [activateConfirm, setActivateConfirm] = useState(null);
  const [permSepTarget, setPermSepTarget]     = useState(null);
  const [permSepReason, setPermSepReason]     = useState("Resigned");
  const [permSepping, setPermSepping]         = useState(false);
  const [permSepConfirm, setPermSepConfirm]   = useState(false);
  const [inactiveQuery, setInactiveQuery]     = useState("");
  const hasInactive = useRef(false);

  const fetchInactive = useCallback(async (q = "", silent = false) => {
    if (!silent) setInactiveLoading(true);
    try {
      const r = await fetch(`${apiUrl}/employees/inactive?q=${encodeURIComponent(q)}&limit=500`, { credentials: "include" });
      if (r.ok) setInactiveEmps(await r.json());
    } finally { setInactiveLoading(false); }
  }, []);

  useEffect(() => { fetchInactive("", hasInactive.current); hasInactive.current = true; }, [fetchInactive, refreshKey]);
  useEffect(() => {
    const t = setTimeout(() => fetchInactive(inactiveQuery), 300);
    return () => clearTimeout(t);
  }, [inactiveQuery, fetchInactive]);

  async function handleActivate(id) {
    setActivating(id);
    setActivateConfirm(null);
    const res = await fetch(`${apiUrl}/employees/${id}/activate`, { method: "PUT", credentials: "include" });
    if (res.ok) { fetchInactive(inactiveQuery, true); fetchSuspLog(); }
    setActivating(null);
  }

  async function handlePermSep() {
    if (!permSepTarget) return;
    setPermSepping(true);
    const res = await fetch(`${apiUrl}/employees/${permSepTarget.key}`, {
      method: "DELETE", credentials: "include",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ reason: permSepReason }),
    });
    if (res.ok) {
      setPermSepTarget(null); setPermSepConfirm(false);
      setSuspSelected(null);
      fetchInactive(inactiveQuery, true);
      fetchSuspLog();
    }
    setPermSepping(false);
  }

  const reasonColors = {
    Resigned:   { bg: "#fef2f2", color: "#991b1b" },
    Retired:    { bg: "#fff7ed", color: "#9a3412" },
    Terminated: { bg: "#fef2f2", color: "#7f1d1d" },
    "On Leave": { bg: "#eff6ff", color: "#1e40af" },
    Suspended:  { bg: "#fefce8", color: "#854d0e" },
  };

  // ── Suspension & Leave Log ──
  const [suspLog, setSuspLog]                 = useState([]);
  const [suspPage, setSuspPage]               = useState(0);
  const [permPage, setPermPage]               = useState(0);
  const [suspLogLoading, setSuspLogLoading]   = useState(false);
  const [suspSearch, setSuspSearch]           = useState("");
  const [suspSelected, setSuspSelected]       = useState(null);
  const [suspDetailPage, setSuspDetailPage]   = useState(0);
  const [logAllEmps, setLogAllEmps]           = useState([]);
  const [suspProfile, setSuspProfile]         = useState(null);
  const [suspProfileLoading, setSuspProfileLoading] = useState(false);
  const [suspAvatarUrl, setSuspAvatarUrl]     = useState(null);
  const [suspPosHistory, setSuspPosHistory]   = useState([]);
  const [suspPosPage, setSuspPosPage]         = useState(0);
  const [suspSgSteps, setSuspSgSteps]         = useState([]);
  const [suspStepPage, setSuspStepPage]       = useState(0);
  const [userLinks, setUserLinks]             = useState([]);
  const [linkSaving, setLinkSaving]           = useState(false);
  const [linkSaved, setLinkSaved]             = useState(false);
  const [linkDropOpen, setLinkDropOpen]       = useState(false);
  const [linkSearch, setLinkSearch]           = useState("");
  const [linkModalSel, setLinkModalSel]       = useState("");
  const [linkConfirm, setLinkConfirm]         = useState(false);

  const hLinkedUser = suspSelected ? userLinks.find(u => u.employee_id === suspSelected.key) ?? null : null;
  const hTakenIds   = new Set(userLinks.filter(u => u.employee_id && u.employee_id !== (suspSelected?.key ?? "")).map(u => u.employee_id));
  const hAvailUsers = userLinks.filter(u => !hTakenIds.has(u.employee_id ?? "") || u.id === hLinkedUser?.id);

  async function handleLinkChange(newUserId) {
    setLinkSaving(true); setLinkSaved(false);
    if (hLinkedUser && hLinkedUser.id !== newUserId) {
      await fetch(`${apiUrl}/admin/users/${hLinkedUser.id}/link-employee`, {
        method: "PATCH", credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ employeeId: null }),
      });
    }
    if (newUserId) {
      await fetch(`${apiUrl}/admin/users/${newUserId}/link-employee`, {
        method: "PATCH", credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ employeeId: suspSelected?.key }),
      });
    }
    const r = await fetch(`${apiUrl}/admin/user-employee-links?limit=200`, { credentials: "include" });
    if (r.ok) setUserLinks(await r.json());
    setLinkSaving(false); setLinkSaved(true);
    setTimeout(() => setLinkSaved(false), 2000);
  }

  const fetchSuspLog = useCallback(async () => {
    setSuspLogLoading(true);
    try {
      const [rLog, rAll, rLinks] = await Promise.all([
        fetch(`${apiUrl}/employees/suspension-log`, { credentials: "include" }),
        fetch(`${apiUrl}/employees?q=&limit=100`, { credentials: "include" }),
        fetch(`${apiUrl}/admin/user-employee-links?limit=200`, { credentials: "include" }),
      ]);
      if (rLog.ok) setSuspLog(await rLog.json());
      if (rAll.ok) { const j = await rAll.json(); setLogAllEmps(j.data ?? j); }
      if (rLinks.ok) setUserLinks(await rLinks.json());
    } finally { setSuspLogLoading(false); }
  }, []);

  useEffect(() => { if (hubTab === "inactive") fetchSuspLog(); }, [hubTab, fetchSuspLog, refreshKey]);


  const tabs = [
    { id: "directory", label: "Directory" },
    ...(isAdmin || can("employees.restore") || can("employees.view") ? [{ id: "inactive", label: "Employee Status", icon: "status", count: inactiveEmps.length }] : []),
  ];

  const deleteButtonStyle = { width: 36, height: 36, padding: 0, display: "flex", alignItems: "center", justifyContent: "center", borderRadius: "var(--radius-sm)", background: "#fee2e2", color: "#991b1b", border: "1.5px solid #fecaca", cursor: "pointer", flexShrink: 0 };

  return (
    <div style={{ maxWidth: "100%" }}>
      <div className="page-header">
        <div className="page-header-title">Employees</div>
        <div className="page-header-sub">Directory, compensation, and account management</div>
      </div>

      <div className="ws-frame-head">
        <WsTabs tabs={tabs} value={hubTab} label="Employee sections"
          onChange={id => { setHubTab(id); setPermPage(0); setSuspSelected(null); }} />
      </div>

      <div className="ws-frame-body">

        {/* ══════════ DIRECTORY TAB (Employees + Compensation) ══════════ */}
        {hubTab === "directory" && (
          <>
            <div className="ws-toolbar">
              <WsSearch value={query} onChange={setQuery} placeholder="Search by name, number, position…" />
              <div className="ws-toolbar-actions">
                {can("payslips.export_pdf") && (
                  <button type="button" className="um-btn um-btn--neutral" title="Download payslips (.docx)" onClick={handleExportDocx} disabled={exporting}>
                    {exporting
                      ? <svg style={{ animation: "spin 0.7s linear infinite" }} width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" aria-hidden="true"><circle cx="12" cy="12" r="10" strokeOpacity="0.25"/><path d="M12 2a10 10 0 0 1 10 10"/></svg>
                      : <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#1d4ed8" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="8" y1="13" x2="16" y2="13"/><line x1="8" y1="17" x2="13" y2="17"/></svg>}
                    {!isMobile && <span>{exporting ? "Generating…" : "Payslips .docx"}</span>}
                  </button>
                )}
                {can("payslips.export_excel") && (
                  <button type="button" className="um-btn um-btn--neutral" title="Download payroll register (.xlsx)" onClick={handleExportXlsx} disabled={exportingXlsx}>
                    {exportingXlsx
                      ? <svg style={{ animation: "spin 0.7s linear infinite" }} width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" aria-hidden="true"><circle cx="12" cy="12" r="10" strokeOpacity="0.25"/><path d="M12 2a10 10 0 0 1 10 10"/></svg>
                      : <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#15803d" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="9" y1="3" x2="9" y2="21"/></svg>}
                    {!isMobile && <span>{exportingXlsx ? "Generating…" : "Register .xlsx"}</span>}
                  </button>
                )}
                {can("employees.create") && (
                  <button type="button" className="btn-primary ws-add" onClick={() => setShowAdd(true)}>
                    {isMobile ? "+ Add" : "+ Add employee"}
                  </button>
                )}
              </div>
            </div>

            {isMobile ? (
              /* ── MOBILE: accordion, tap to expand compensation inline ── */
              <div>
                {loading ? (
                  <div className="empty" style={{ padding: "20px 16px" }}>Searching…</div>
                ) : fetchError ? (
                  <div className="empty" style={{ padding: "20px 16px", color: "#991b1b", fontSize: "0.8rem" }}>{fetchError}</div>
                ) : employees.length === 0 ? (
                  <div className="empty" style={{ padding: "20px 16px" }}>No employees found.</div>
                ) : (
                  <div style={{ display: "flex", flexDirection: "column" }}>
                    {employees.map((emp, empIdx) => {
                      const isOpen = openId === emp.id;
                      const isLast = empIdx === employees.length - 1;
                      return (
                        <div key={emp.id}>
                          <div
                            className={`ws-item ws-item--row${isOpen ? " is-on" : ""}`}
                            style={{ borderBottom: "1px solid var(--border)" }}
                            onClick={() => {
                              const opening = openId !== emp.id;
                              setOpenId(v => v === emp.id ? null : emp.id);
                              if (opening) { bumpToTop(emp.id); setSelectedId(emp.id); fetch(`${apiUrl}/employees/${emp.id}/search-hit`, { method: "POST", credentials: "include" }).catch(() => {}); }
                            }}>
                            <div className="ws-item-main">
                              <div className="ws-item-title">{emp.name}</div>
                              <div className="ws-item-sub">{emp.employee_number} · {emp.position || "No position"}</div>
                            </div>
                            <svg width="13" height="13" viewBox="0 0 16 16" style={{ flexShrink: 0, transition: "transform 0.2s", transform: isOpen ? "rotate(180deg)" : "rotate(0deg)", color: "var(--text-muted)" }}>
                              <path d="M4 6l4 4 4-4" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" fill="none"/>
                            </svg>
                          </div>
                          {isOpen && (
                            <div style={{ overflow: "hidden", background: "var(--surface-2)", borderBottom: isLast ? "none" : "1px solid var(--border)" }}>
                              {/* Action buttons */}
                              <div style={{ display: "flex", justifyContent: "flex-end", gap: 6, padding: "8px 12px", borderBottom: "1px solid var(--border)" }}>
                                {can("employees.edit") && (
                                  <button type="button" title="Edit employee" onClick={() => { bumpToTop(emp.id); setEditEmp(emp); }}
                                    style={{ display: "inline-flex", alignItems: "center", gap: 5, padding: "5px 12px", borderRadius: 999, border: "1.5px solid var(--border)", background: "transparent", color: "var(--text-muted)", fontSize: "0.75rem", fontWeight: 600, cursor: "pointer" }}>
                                    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                    Edit
                                  </button>
                                )}
                                {can("employees.delete") && (
                                  <button type="button" title="Set On Leave / Suspend" onClick={() => handleDelete(emp)}
                                    style={{ display: "inline-flex", alignItems: "center", gap: 5, padding: "5px 12px", borderRadius: 999, border: "1.5px solid #fecaca", background: "#fee2e2", color: "#991b1b", fontSize: "0.75rem", fontWeight: 600, cursor: "pointer" }}>
                                    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10"/><line x1="10" y1="15" x2="10" y2="9"/><line x1="14" y1="15" x2="14" y2="9"/></svg>
                                    On Leave
                                  </button>
                                )}
                              </div>
                              {/* Tab nav */}
                              <div style={{ display: "flex", padding: "8px 12px 0", gap: 4, background: "var(--surface-2)", borderBottom: "1px solid var(--border)" }}>
                                {[{ label: "Pay Slip", key: "payslip" }, { label: "Payroll Summary", key: "finance" }, { label: "Records", key: "records" }].map(tab => (
                                  <button key={tab.key} type="button" onClick={() => setEmpTab(tab.key)}
                                    style={{ fontSize: "0.77rem", fontWeight: 600, padding: "6px 13px", borderRadius: "6px 6px 0 0", border: "none", cursor: "pointer", background: empTab === tab.key ? "var(--psa-navy)" : "transparent", color: empTab === tab.key ? "#fff" : "var(--text-muted)", transition: "all 0.15s" }}>
                                    {tab.label}
                                  </button>
                                ))}
                              </div>
                              {/* Pay Slip tab */}
                              {empTab === "payslip" && (
                                <>
                                  <div style={{ display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: 10, padding: "12px" }}>
                                    <InfoField label="Employee Number" value={emp.employee_number || "Not set"} />
                                    <InfoField label="Basic Monthly Pay" value={`₱ ${fmt(emp.basic_monthly_pay)}`} />
                                    {emp.salary_grade_code && (
                                      <InfoField label="Salary Grade / Step" value={`${emp.salary_grade_code} — Step ${emp.salary_grade_step}`} />
                                    )}
                                    <InfoField label="Employment Type" value={emp.employment_type} />
                                    <InfoField label="Office/Division" value={emp.office_division || "—"} />
                                    <InfoField label="Linked Account" value={emp.linked_username ? `@${emp.linked_username}` : "Not linked"} />
                                  </div>
                                  {payslipLoading ? (
                                    <div style={{ padding: "24px 12px", textAlign: "center", color: "var(--text-muted)", fontSize: "0.85rem" }}>Loading pay slip…</div>
                                  ) : payslipItem ? (
                                    <PaySlipCard hideHeader item={payslipItem} period={period} isOpen={true} onToggle={() => {}} />
                                  ) : (
                                    <div style={{ padding: "24px 12px", textAlign: "center", color: "var(--text-muted)", fontSize: "0.85rem" }}>No pay slip data available yet.</div>
                                  )}
                                </>
                              )}
                              {/* Payroll Summary tab */}
                              {empTab === "finance" && canDeductions && (
                                <EmployeeDeductionCard hideHeader hideLoans emp={emp} isAdmin={canDeductions} onEmpPatch={updates => patchEmpInList(emp.id, updates)} onTakeHomePay={setEmbeddedTakeHome} />
                              )}
                              {empTab === "finance" && !canDeductions && (
                                <div style={{ padding: "24px 12px", textAlign: "center", color: "var(--text-muted)", fontSize: "0.85rem" }}>You don't have permission to view payroll summary.</div>
                              )}
                              {/* Records tab */}
                              {empTab === "records" && (
                                <div style={{ padding: "12px" }}>
                                  <div style={{ display: "flex", gap: 4, marginBottom: 12 }}>
                                    {[{ label: "Loans", key: "loans" }, { label: "Payment History", key: "history" }].map(t => (
                                      <button key={t.key} type="button" onClick={() => setRecordsTab(t.key)}
                                        style={{ fontSize: "0.78rem", fontWeight: 600, padding: "5px 16px", borderRadius: 999, border: recordsTab === t.key ? "none" : "1.5px solid var(--border)", cursor: "pointer", background: recordsTab === t.key ? "var(--psa-navy)" : "transparent", color: recordsTab === t.key ? "#fff" : "var(--text-muted)", transition: "all 0.15s" }}>
                                        {t.label}
                                      </button>
                                    ))}
                                  </div>
                                  {recordsTab === "loans" ? (
                                    <LoansCard emp={emp} isAdmin={canDeductions} />
                                  ) : (
                                    <AdjustmentsCard emp={emp} />
                                  )}
                                </div>
                              )}
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            ) : (
              /* ── DESKTOP: master-detail — click employee, see compensation ── */
              <div className="ws-split">
                {/* LEFT — search results list */}
                <div className="ws-list">
                  {loading ? (
                    <WsSpinner label="Searching…" />
                  ) : fetchError ? (
                    <div className="ws-empty" style={{ color: "#991b1b" }}>{fetchError}</div>
                  ) : employees.length === 0 ? (
                    <div className="ws-empty">No employees found.</div>
                  ) : employees.map(emp => (
                    <div
                      key={emp.id}
                      role="button" tabIndex={0}
                      className={`ws-item${selectedId === emp.id ? " is-on" : ""}`}
                      onKeyDown={e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); e.currentTarget.click(); } }}
                      onClick={() => {
                        setSelectedId(emp.id);
                        bumpToTop(emp.id);
                        fetch(`${apiUrl}/employees/${emp.id}/search-hit`, { method: "POST", credentials: "include" }).catch(() => {});
                      }}
                    >
                      <div className="ws-item-main">
                        <div className="ws-item-title">{emp.name}</div>
                        <div className="ws-item-sub" title={emp.position}>{emp.employee_number} · {emp.position || "No position"}</div>
                      </div>
                    </div>
                  ))}
                </div>

                {/* RIGHT — selected employee's info + compensation */}
                <div className="ws-detail ws-pad">
                  {selectedEmp ? (
                    <>
                      <div className="ws-detail-head" style={{ padding: "0 0 14px", marginBottom: 14 }}>
                        <div className="ws-ident">
                          {selAvatar ? (
                            <button type="button" className="ws-avatar ws-avatar--photo" title="View photo" aria-label={`View photo of ${selectedEmp.name}`} onClick={() => setAvatarOpen(true)}>
                              <img src={selAvatar} alt="" />
                            </button>
                          ) : (
                            <span className="ws-avatar" title="No profile photo yet">{(selectedEmp.name ?? "?").trim().charAt(0).toUpperCase()}</span>
                          )}
                          {avatarOpen && selAvatar && (
                            <div className="ar-backdrop" onClick={() => setAvatarOpen(false)}>
                              <div className="ws-photo" role="dialog" aria-modal="true" aria-label={`Photo of ${selectedEmp.name}`} onClick={e => e.stopPropagation()}>
                                <img src={selAvatar} alt={`Photo of ${selectedEmp.name}`} />
                                <div className="ws-photo-cap">
                                  <strong>{selectedEmp.name}</strong>
                                  <span>#{selectedEmp.employee_number || "—"} · {selectedEmp.position || "No position"}{selectedEmp.office_division ? ` · ${selectedEmp.office_division}` : ""}</span>
                                </div>
                                <button type="button" className="ws-photo-close" aria-label="Close" onClick={() => setAvatarOpen(false)}>✕</button>
                              </div>
                            </div>
                          )}
                          <div style={{ minWidth: 0 }}>
                            <div className="ws-dtitle" style={{ marginTop: 0 }}>{selectedEmp.name}</div>
                            <div className="ws-item-sub">#{selectedEmp.employee_number || "—"} · {selectedEmp.position || "No position"}{selectedEmp.office_division ? ` · ${selectedEmp.office_division}` : ""}</div>
                          </div>
                        </div>
                        <div className="ws-actions">
                          {can("employees.delete") && (
                            <button type="button" className="um-btn um-btn--warn um-btn--sm" title="Set on leave or suspend" onClick={() => handleDelete(selectedEmp)}>
                              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="10"/><line x1="10" y1="15" x2="10" y2="9"/><line x1="14" y1="15" x2="14" y2="9"/></svg>
                              On leave / suspend
                            </button>
                          )}
                          {can("employees.edit") && (
                            <button type="button" className="um-btn um-btn--neutral um-btn--sm" onClick={() => { bumpToTop(selectedEmp.id); setEditEmp(selectedEmp); }}>
                              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                              Edit
                            </button>
                          )}
                        </div>
                      </div>
                      {canDeductions && (
                        <div style={{ marginBottom: 16 }}>
                          <WsTabs small label="Employee details"
                            tabs={[{ id: "payslip", label: "Pay Slip" }, { id: "finance", label: "Payroll Summary" }, { id: "records", label: "Records" }]}
                            value={empTab} onChange={setEmpTab} />
                        </div>
                      )}
                      {empTab === "payslip" && <div className="grid-3-col" style={{ gap: 12, marginBottom: 18 }}>
                        <InfoField label="Employee Number" value={selectedEmp.employee_number || "Not set"} />
                        <InfoField label="Full Name" value={selectedEmp.name || "—"} />
                        <InfoField label="Basic Monthly Pay" value={`₱ ${fmt(selectedEmp.basic_monthly_pay)}`} />
                        {selectedEmp.salary_grade_code && (
                          <InfoField
                            label="Salary Grade / Step"
                            value={`${selectedEmp.salary_grade_code} — Step ${selectedEmp.salary_grade_step}`}
                          />
                        )}
                        <InfoField label="Employment Type" value={selectedEmp.employment_type} />
                        <InfoField label="Office/Division" value={selectedEmp.office_division || "—"} />
                        <InfoField
                          label="Appointed Date"
                          value={(selectedEmp.appointed_date || selectedEmp.start_date) ? new Date(selectedEmp.appointed_date || selectedEmp.start_date).toLocaleDateString("en-PH", { year: "numeric", month: "long", day: "numeric" }) : "—"}
                        />
                        <InfoField
                          label="Linked Account"
                          value={selectedEmp.linked_username ? `@${selectedEmp.linked_username}` : "Not linked"}
                        />
                        {empTab === "payslip" && payslipItem?.net_pay != null && (
                          <div>
                            <div style={{ fontSize: "0.72rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 4 }}>Take Home Pay</div>
                            <div style={{ fontSize: "0.95rem", fontWeight: 800, color: "#166534" }}>₱ {fmt(payslipItem.net_pay)}</div>
                          </div>
                        )}
                      </div>}
                      {empTab === "records" ? (
                        <div>
                          <div style={{ marginBottom: 12 }}>
                            <WsTabs small label="Records"
                              tabs={[{ id: "loans", label: "Loans" }, { id: "history", label: "Payment History" }]}
                              value={recordsTab} onChange={setRecordsTab} />
                          </div>
                          {recordsTab === "loans" ? (
                            <LoansCard emp={selectedEmp} isAdmin={canDeductions} />
                          ) : (
                            <AdjustmentsCard emp={selectedEmp} />
                          )}
                        </div>
                      ) : payslipEditMode ? (
                        <>
                          <div style={{ fontSize: "1.1rem", fontWeight: 800, color: "var(--psa-navy)", marginBottom: 12, letterSpacing: "-0.01em" }}>Payroll Summary</div>
                          <EmployeeDeductionCard hideHeader hideLoans emp={selectedEmp} isAdmin={canDeductions} onEmpPatch={updates => patchEmpInList(selectedEmp.id, updates)} onTakeHomePay={setEmbeddedTakeHome} />
                        </>
                      ) : payslipLoading ? (
                        <div className="panel" style={{ padding: 40, textAlign: "center", color: "var(--text-muted)" }}>Loading pay slip…</div>
                      ) : payslipItem ? (
                        <PaySlipCard hideHeader item={payslipItem} period={period} isOpen={true} onToggle={() => {}} />
                      ) : (
                        <div className="panel" style={{ padding: 40, textAlign: "center", color: "var(--text-muted)" }}>No pay slip data available yet.</div>
                      )}
                    </>
                  ) : (
                    <div className="ws-empty-state">
                      <svg width="44" height="44" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                        <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/>
                      </svg>
                      <strong>No employee selected</strong>
                      <span style={{ fontSize: "0.84rem" }}>Select an employee on the left to see their pay slip, payroll summary and records.</span>
                    </div>
                  )}
                </div>
              </div>
            )}
          </>
        )}

        {/* ══════════ INACTIVE & LEAVE LOG TAB (merged) ══════════ */}
        {hubTab === "inactive" && (isAdmin || can("employees.restore") || can("employees.view")) && (() => {
          const fmtDate = d => d ? new Date(d).toLocaleDateString("en-PH", { month: "short", day: "numeric", year: "numeric" }) : "—";
          const fmtDT   = d => d ? new Date(d).toLocaleString("en-PH", { month: "short", day: "numeric", year: "numeric", hour: "2-digit", minute: "2-digit" }) : null;
          const fmtDuration = (hrs, mins) => {
            const h = Number(hrs) || 0, m = Number(mins) || 0;
            if (h < 1) return m > 0 ? `${m} min` : "< 1 hr";
            if (h < 24) return m > 0 ? `${h} hr ${m} min` : `${h} hr`;
            const days = Math.floor(h / 24), remH = h % 24;
            return remH > 0 ? `${days}d ${remH}hr` : `${days}d`;
          };

          async function selectSuspEmployee(empId, logEntry, toggle) {
            if (toggle) { setSuspSelected(null); setSuspProfile(null); setSuspAvatarUrl(null); setSuspPosHistory([]); setSuspPosPage(0); setSuspDetailPage(0); setSuspSgSteps([]); setSuspStepPage(0); return; }
            setSuspSelected(logEntry);
            setSuspDetailPage(0);
            setSuspStepPage(0);
            setSuspProfile(null);
            setSuspAvatarUrl(null);
            setSuspPosHistory([]);
            setSuspPosPage(0);
            setSuspSgSteps([]);
            setSuspProfileLoading(true);
            try {
              const [pr, posR] = await Promise.all([
                fetch(`${apiUrl}/employees/${empId}/profile`, { credentials: "include" }),
                fetch(`${apiUrl}/employees/${empId}/position-history`, { credentials: "include" }),
              ]);
              if (pr.ok) { const d = await pr.json(); setSuspProfile(d); if (d.employee?.avatarDataUrl) setSuspAvatarUrl(d.employee.avatarDataUrl); }
              if (posR.ok) { const d = await posR.json(); setSuspPosHistory(Array.isArray(d) ? d : []); }
            } finally { setSuspProfileLoading(false); }
          }

          // Build suspension records map keyed by employee_id
          const grouped = {};
          for (const s of suspLog) {
            const key = s.employee_id ?? s.employee_name;
            if (!grouped[key]) grouped[key] = { key, name: s.employee_name, number: s.employee_number, records: [] };
            grouped[key].records.push(s);
          }

          const q = inactiveQuery.trim().toLowerCase();

          // TOP — currently On Leave or Suspended (show ALL, no pagination)
          const activeLeave = inactiveEmps.filter(e =>
            ["On Leave", "Suspended"].includes(e.separation_reason) &&
            (!q || e.name?.toLowerCase().includes(q) || e.employee_number?.toLowerCase().includes(q))
          );

          // BOTTOM — ALL employees (active + inactive combined), paginated 5
          const inactiveIds = new Set(inactiveEmps.map(e => e.id));
          const allEmps = [
            ...inactiveEmps,
            ...logAllEmps.filter(e => !inactiveIds.has(e.id)),
          ].sort((a, b) => a.name.localeCompare(b.name));
          const allFiltered = allEmps.filter(e =>
            !q || e.name?.toLowerCase().includes(q) || e.employee_number?.toLowerCase().includes(q)
          );
          const DIR_PAGE = 5;
          const dirPages   = Math.ceil(allFiltered.length / DIR_PAGE) || 1;
          const safePerm   = Math.min(permPage, dirPages - 1);
          const visibleDir = allFiltered.slice(safePerm * DIR_PAGE, (safePerm + 1) * DIR_PAGE);

          // Map inactive employees by id for Activate button lookup
          const inactiveMap = new Map(inactiveEmps.map(e => [e.id, e]));

          return (
            <div className="grid-chart-split" style={{ minHeight: 400 }}>
              {/* ── Left panel (search list — now right visually) ── */}
              <div style={{ order: 2, borderLeft: "1px solid var(--border)", overflowY: "auto", padding: "20px 28px 20px" }}>

                {/* Search */}
                <div style={{ position: "relative", marginBottom: 16 }}>
                  <svg style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)", pointerEvents: "none" }} width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="var(--text-muted)" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
                  </svg>
                  <input style={{ width: "100%", boxSizing: "border-box", paddingLeft: 32, paddingRight: 12, paddingTop: 8, paddingBottom: 8, fontSize: "0.83rem", border: "1px solid var(--border)", borderRadius: 8, background: "var(--surface-2)", color: "var(--text)", outline: "none" }}
                    placeholder="Search by name or employee number…" value={inactiveQuery} onChange={e => { setInactiveQuery(e.target.value); setPermPage(0); }} />
                </div>

                {suspLogLoading ? (
                  <div style={{ padding: "40px 0", textAlign: "center", color: "var(--text-muted)" }}>
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ animation: "spin 1s linear infinite", display: "block", margin: "0 auto 10px" }}><path d="M21 12a9 9 0 1 1-6.219-8.56"/></svg>
                    <div style={{ fontSize: "0.82rem" }}>Loading…</div>
                  </div>
                ) : (
                  <>
                    {/* TOP — Active Leave & Suspensions, show ALL */}
                    {activeLeave.length > 0 && (
                      <div style={{ marginBottom: 20 }}>
                        <div style={{ fontSize: "0.6rem", fontWeight: 700, color: "#1e40af", textTransform: "uppercase", letterSpacing: "0.12em", marginBottom: 8 }}>
                          Active Leave & Suspensions ({activeLeave.length})
                        </div>
                        {activeLeave.map(emp => {
                          const logEntry = grouped[emp.id] ?? { key: emp.id, name: emp.name, number: emp.employee_number, records: [] };
                          const isSelected = suspSelected?.key === emp.id;
                          const rc = reasonColors[emp.separation_reason] ?? { color: "#64748b" };
                          return (
                            <div key={emp.id} onClick={() => selectSuspEmployee(emp.id, logEntry, isSelected)}
                              style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 10px", borderRadius: 6, cursor: "pointer", background: isSelected ? "var(--surface-2)" : "transparent" }}>
                              <div>
                                <div style={{ fontSize: "0.85rem", fontWeight: 700 }}>{emp.name}</div>
                                <div style={{ fontSize: "0.72rem", color: "var(--text-muted)", marginTop: 2 }}>{emp.employee_number} · {emp.position}{emp.office_division ? ` · ${emp.office_division}` : ""}</div>
                              </div>
                              <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0, marginLeft: 12 }}>
                                <span style={{ fontSize: "0.7rem", fontWeight: 700, color: rc.color }}>{emp.separation_reason}</span>
                                <button type="button" onClick={e => { e.stopPropagation(); setActivateConfirm(emp); }} disabled={activating === emp.id}
                                  style={{ fontSize: "0.7rem", fontWeight: 700, padding: "3px 10px", borderRadius: 999, border: "1px solid #166534", background: "transparent", color: "#166534", cursor: "pointer" }}>
                                  {activating === emp.id ? "…" : "Activate"}
                                </button>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    )}

                    {/* BOTTOM — All employees, paginated 5 */}
                    <div>
                      <div style={{ fontSize: "0.6rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.12em", marginBottom: 8 }}>
                        All Employees ({allFiltered.length})
                      </div>
                      {allFiltered.length === 0 ? (
                        <div style={{ fontSize: "0.83rem", color: "var(--text-muted)", padding: "8px 0" }}>No results for "{inactiveQuery}".</div>
                      ) : (
                        <>
                          {visibleDir.map(emp => {
                            const logEntry = grouped[emp.id] ?? { key: emp.id, name: emp.name, number: emp.employee_number, records: [] };
                            const isSelected = suspSelected?.key === emp.id;
                            const inactiveInfo = inactiveMap.get(emp.id);
                            const rc = inactiveInfo ? (reasonColors[inactiveInfo.separation_reason] ?? { color: "#64748b" }) : null;
                            return (
                              <div key={emp.id} onClick={() => selectSuspEmployee(emp.id, logEntry, isSelected)}
                                style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 10px", borderRadius: 6, cursor: "pointer", background: isSelected ? "var(--surface-2)" : "transparent" }}>
                                <div>
                                  <div style={{ fontSize: "0.85rem", fontWeight: 700 }}>{emp.name}</div>
                                  <div style={{ fontSize: "0.72rem", color: "var(--text-muted)", marginTop: 2 }}>{emp.employee_number}{emp.position ? ` · ${emp.position}` : ""}{emp.office_division ? ` · ${emp.office_division}` : ""}</div>
                                </div>
                                <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0, marginLeft: 12 }}>
                                  {rc && <span style={{ fontSize: "0.7rem", fontWeight: 700, color: rc.color }}>{inactiveInfo.separation_reason}</span>}
                                  {logEntry.records.length > 0 && <span style={{ fontSize: "0.68rem", color: "var(--text-muted)", background: "var(--surface-2)", border: "1px solid var(--border)", borderRadius: 999, padding: "2px 7px" }}>{logEntry.records.length} record{logEntry.records.length !== 1 ? "s" : ""}</span>}
                                </div>
                              </div>
                            );
                          })}
                          {dirPages > 1 && (
                            <div style={{ display: "flex", alignItems: "center", justifyContent: "flex-end", gap: 8, paddingTop: 8, marginTop: 4, borderTop: "1px solid var(--border)" }}>
                              <button onClick={() => setPermPage(p => Math.max(0, p - 1))} disabled={safePerm === 0}
                                style={{ background: "none", border: "1px solid var(--border)", borderRadius: 6, padding: "5px 14px", fontSize: "0.9rem", cursor: safePerm === 0 ? "default" : "pointer", color: safePerm === 0 ? "var(--text-muted)" : "var(--psa-navy)", fontWeight: 700 }}>‹</button>
                              <span style={{ fontSize: "0.75rem", color: "var(--text-muted)", minWidth: 36, textAlign: "center" }}>{safePerm + 1} / {dirPages}</span>
                              <button onClick={() => setPermPage(p => Math.min(dirPages - 1, p + 1))} disabled={safePerm === dirPages - 1}
                                style={{ background: "none", border: "1px solid var(--border)", borderRadius: 6, padding: "5px 14px", fontSize: "0.9rem", cursor: safePerm === dirPages - 1 ? "default" : "pointer", color: safePerm === dirPages - 1 ? "var(--text-muted)" : "var(--psa-navy)", fontWeight: 700 }}>›</button>
                            </div>
                          )}
                        </>
                      )}
                    </div>
                  </>
                )}
              </div>

              {/* ── Right: profile/detail — now left visually ── */}
              <div style={{ order: 1, padding: "20px 28px", overflowY: "auto" }}>
                {suspSelected ? (
                  <>
                    <button onClick={() => { setSuspSelected(null); setSuspProfile(null); setSuspAvatarUrl(null); setSuspPosHistory([]); setSuspPosPage(0); setSuspDetailPage(0); }}
                      style={{ display: "inline-flex", alignItems: "center", gap: 5, marginBottom: 16, padding: "4px 10px 4px 6px", borderRadius: 6, border: "1px solid var(--border)", background: "var(--surface-2)", color: "var(--text-muted)", fontSize: "0.75rem", fontWeight: 600, cursor: "pointer" }}
                      onMouseEnter={e => e.currentTarget.style.background = "var(--border)"}
                      onMouseLeave={e => e.currentTarget.style.background = "var(--surface-2)"}>
                      <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="15 18 9 12 15 6"/></svg>
                      Back
                    </button>

                    {/* ── Employee header with avatar ── */}
                    {(() => {
                      const selEmp = [...(logAllEmps ?? []), ...(inactiveEmps ?? [])].find(e => e.id === suspSelected.key);
                      const isAlreadyPerm = ["Resigned", "Retired", "Terminated"].includes(selEmp?.separation_reason);
                      const statusLabel = selEmp?.separation_reason ?? "Active";
                      const statusColor = statusLabel === "Active" ? "#166534" : statusLabel === "On Leave" ? "#1e40af" : statusLabel === "Suspended" ? "#854d0e" : "#64748b";
                      const initials = (suspSelected.name ?? "?").split(" ").map(w => w[0]).filter(Boolean).slice(0, 2).join("").toUpperCase();
                      return (
                        <div style={{ display: "flex", alignItems: "center", gap: 12, paddingBottom: 14, borderBottom: "1px solid var(--border)", marginBottom: 4 }}>
                          <div style={{ width: 44, height: 44, flexShrink: 0 }}>
                            {suspAvatarUrl
                              ? <img src={suspAvatarUrl} alt="" style={{ width: 44, height: 44, borderRadius: "50%", objectFit: "cover", display: "block" }} />
                              : <div style={{ width: 44, height: 44, borderRadius: "50%", background: "var(--psa-navy)", color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontSize: "0.85rem", fontWeight: 800, letterSpacing: "0.04em" }}>{initials}</div>
                            }
                          </div>
                          <div style={{ minWidth: 0, flex: 1 }}>
                            <div style={{ fontSize: "0.95rem", fontWeight: 700, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{suspSelected.name}</div>
                            <div style={{ fontSize: "0.7rem", color: "var(--text-muted)", marginTop: 1 }}>{suspSelected.number}</div>
                          </div>
                          <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 6, flexShrink: 0 }}>
                            <span style={{ fontSize: "0.7rem", fontWeight: 700, color: statusColor }}>{statusLabel}</span>
                            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                              {isAdmin && (
                                <button
                                  type="button"
                                  disabled={linkSaving}
                                  title="Link user account"
                                  onClick={() => { setLinkModalSel(hLinkedUser?.id ?? ""); setLinkSearch(""); setLinkConfirm(false); setLinkDropOpen(true); }}
                                  style={{ height: 36, padding: "0 10px", display: "flex", alignItems: "center", gap: 6, borderRadius: "var(--radius-sm)", border: "1px solid var(--border)", background: "var(--surface-2)", color: hLinkedUser ? "#166534" : "var(--text-muted)", cursor: "pointer", fontSize: "0.75rem", fontWeight: hLinkedUser ? 700 : 400, maxWidth: 150, outline: "none" }}
                                >
                                  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                                    <path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/>
                                    <path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/>
                                  </svg>
                                  <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                                    {linkSaved ? "✓ Saved" : hLinkedUser ? hLinkedUser.username : "No link"}
                                  </span>
                                </button>
                              )}
                              {!isAlreadyPerm && (
                                <button type="button" title="Permanent Separation" onClick={() => { setPermSepTarget(suspSelected); setPermSepReason("Resigned"); }}
                                  style={deleteButtonStyle}>
                                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                                    <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/>
                                    <circle cx="9" cy="7" r="4"/>
                                    <line x1="17" y1="8" x2="23" y2="14"/>
                                    <line x1="23" y1="8" x2="17" y2="14"/>
                                  </svg>
                                </button>
                              )}
                            </div>
                          </div>
                        </div>
                      );
                    })()}

                    {/* ── Profile details ── */}
                    {suspProfileLoading ? (
                      <div style={{ fontSize: "0.8rem", color: "var(--text-muted)", padding: "12px 0", textAlign: "center" }}>Loading profile…</div>
                    ) : suspProfile?.employee && (() => {
                      const e = suspProfile.employee;
                      const fmtD = d => d ? new Date(d).toLocaleDateString("en-PH", { month: "short", day: "numeric", year: "numeric" }) : "—";
                      const SL = ({ children }) => (
                        <div style={{ fontSize: "0.65rem", fontWeight: 800, color: "var(--text)", textTransform: "uppercase", letterSpacing: "0.12em", marginTop: 20, marginBottom: 8 }}>{children}</div>
                      );
                      const Row = ({ label, children }) => (
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "4px 0" }}>
                          <span style={{ fontSize: "0.78rem", color: "var(--text-muted)", flexShrink: 0 }}>{label}</span>
                          <span style={{ fontSize: "0.8rem", fontWeight: 600, textAlign: "right", marginLeft: 10 }}>{children}</span>
                        </div>
                      );
                      return (
                        <>
                          <SL>Employee</SL>
                          {e.office_division && <Row label="Division">{e.office_division}</Row>}
                          {e.position && <Row label="Position">{e.position}</Row>}
                          {e.employment_type && <Row label="Type">{e.employment_type}</Row>}
                          <Row label="Start Date">{e.start_date ? fmtD(e.start_date) : "—"}</Row>

                          {/* Step Increment History — from employee_promotion table */}
                          {suspSgSteps.length > 0 && (() => {
                            const today = new Date();
                            const allRows = [...suspSgSteps]
                              .filter(r => new Date(r.effective_date) <= today)
                              .sort((a, b) => b.step - a.step);
                            const currentRow = allRows[0] ?? null;
                            // Compute from effective_date + 3 years — current step end_date is NULL (unknown)
                            const expectedNextDate = currentRow?.effective_date
                              ? (() => { const d = new Date(currentRow.effective_date); d.setFullYear(d.getFullYear() + 3); return d; })()
                              : null;
                            const daysUntil = expectedNextDate ? Math.round((expectedNextDate - today) / 86400000) : null;
                            const nextStepNum = currentRow ? currentRow.step + 1 : null;
                            const SPAGE = 3;
                            const sTotal = Math.ceil(allRows.length / SPAGE) || 1;
                            const sSafe  = Math.min(suspStepPage, sTotal - 1);
                            const sItems = allRows.slice(sSafe * SPAGE, (sSafe + 1) * SPAGE);
                            const gradeLabel = suspSgSteps[0]?.salary_grade ?? e.salary_rank;
                            const maxedAll = currentRow?.step === 8;
                            return (
                              <>
                                <div style={{ fontSize: "0.65rem", fontWeight: 800, color: "var(--text)", textTransform: "uppercase", letterSpacing: "0.12em", marginTop: 20, marginBottom: 8 }}>
                                  Step Increment History ({allRows.length})
                                  {gradeLabel && <span style={{ marginLeft: 7, fontWeight: 500, color: "var(--text-muted)", textTransform: "none", letterSpacing: 0 }}>· {gradeLabel}</span>}
                                </div>
                                {expectedNextDate && daysUntil != null && nextStepNum <= 8 && (
                                  <div style={{ borderLeft: "2px solid #94a3b8", paddingLeft: 10, marginBottom: 12 }}>
                                    <div style={{ fontSize: "0.79rem", fontWeight: 700, color: "var(--text)" }}>
                                      Expected Step {nextStepNum}: {fmtD(expectedNextDate)}
                                    </div>
                                    <div style={{ fontSize: "0.69rem", color: "var(--text-muted)", marginTop: 2 }}>
                                      {daysUntil.toLocaleString()} days from today
                                    </div>
                                  </div>
                                )}
                                {sItems.map((row, i) => {
                                  const maxStep   = allRows[0]?.step ?? 0;
                                  const isCurrent = row.step === maxStep;
                                  const labelColor  = isCurrent ? "#1d4ed8" : "#166534";
                                  const statusLabel = isCurrent ? "Current" : "Completed";
                                  return (
                                    <div key={row.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", padding: "8px 0", borderTop: i > 0 ? "1px solid var(--border)" : "none" }}>
                                      <div style={{ flex: 1, minWidth: 0 }}>
                                        <span style={{ fontSize: "0.78rem", fontWeight: 700, color: labelColor }}>Step {row.step}</span>
                                        <div style={{ fontSize: "0.7rem", color: "var(--text-muted)", marginTop: 3 }}>{fmtD(new Date(row.effective_date))}</div>
                                        {isCurrent && row.salary_grade && row.amount && (
                                          <div style={{ fontSize: "0.7rem", color: "var(--text-muted)", marginTop: 2 }}>
                                            {row.salary_grade} · ₱{Number(row.amount).toLocaleString("en-PH", { minimumFractionDigits: 2 })}
                                          </div>
                                        )}
                                      </div>
                                      <div style={{ textAlign: "right", flexShrink: 0, marginLeft: 12 }}>
                                        <div style={{ fontSize: "0.68rem", fontWeight: 700, color: labelColor }}>{statusLabel}</div>
                                      </div>
                                    </div>
                                  );
                                })}
                                {sTotal > 1 && (
                                  <div style={{ display: "flex", alignItems: "center", justifyContent: "flex-end", gap: 8, paddingTop: 8 }}>
                                    <button onClick={() => setSuspStepPage(p => Math.max(0, p - 1))} disabled={sSafe === 0}
                                      style={{ background: "none", border: "1px solid var(--border)", borderRadius: 6, padding: "5px 14px", fontSize: "0.9rem", fontWeight: 700, cursor: sSafe === 0 ? "default" : "pointer", color: sSafe === 0 ? "var(--text-muted)" : "var(--psa-navy)" }}>‹</button>
                                    <span style={{ fontSize: "0.75rem", color: "var(--text-muted)", minWidth: 40, textAlign: "center" }}>{sSafe + 1} / {sTotal}</span>
                                    <button onClick={() => setSuspStepPage(p => Math.min(sTotal - 1, p + 1))} disabled={sSafe === sTotal - 1}
                                      style={{ background: "none", border: "1px solid var(--border)", borderRadius: 6, padding: "5px 14px", fontSize: "0.9rem", fontWeight: 700, cursor: sSafe === sTotal - 1 ? "default" : "pointer", color: sSafe === sTotal - 1 ? "var(--text-muted)" : "var(--psa-navy)" }}>›</button>
                                  </div>
                                )}
                                {maxedAll && <div style={{ fontSize: "0.7rem", color: "var(--text-muted)", paddingTop: 6 }}>Maximum step reached.</div>}
                              </>
                            );
                          })()}
                        </>
                      );
                    })()}

                    <div style={{ fontSize: "0.65rem", fontWeight: 800, color: "var(--text)", textTransform: "uppercase", letterSpacing: "0.12em", marginTop: 20, marginBottom: 10 }}>
                      Records ({suspSelected.records.length})
                    </div>
                    {(() => {
                      const sorted = [...suspSelected.records].sort((a, b) => {
                        if (!a.end_date && b.end_date) return -1;  // no end = active = top
                        if (a.end_date && !b.end_date) return 1;
                        const endA = a.end_date ? new Date(a.end_date) : 0;
                        const endB = b.end_date ? new Date(b.end_date) : 0;
                        return endB - endA;
                      });
                      const RPAGE = 3; const rTotal = Math.ceil(sorted.length / RPAGE) || 1;
                      const rSafe = Math.min(suspDetailPage, rTotal - 1);
                      const rItems = sorted.slice(rSafe * RPAGE, (rSafe + 1) * RPAGE);
                      return (
                        <>
                          {rItems.map((s, i) => {
                            const ongoing = !s.end_date;
                            const reasonColor = s.reason === "Suspended" ? "#854d0e" : "#1e40af";
                            return (
                              <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", padding: "8px 0", borderTop: i > 0 ? "1px solid var(--border)" : "none" }}>
                                <div style={{ flex: 1, minWidth: 0 }}>
                                  <span style={{ fontSize: "0.78rem", fontWeight: 700, color: reasonColor }}>{s.reason}</span>
                                  <div style={{ fontSize: "0.7rem", color: "var(--text-muted)", marginTop: 3 }}>
                                    Start: {fmtDT(s.start_date) ?? "—"}
                                  </div>
                                  <div style={{ fontSize: "0.7rem", marginTop: 2 }}>
                                    {ongoing
                                      ? <span style={{ color: "#dc2626", fontWeight: 700 }}>No return date set</span>
                                      : <span style={{ color: "var(--text-muted)" }}>Returns: <span style={{ color: "var(--text)", fontWeight: 600 }}>{fmtDT(s.end_date)}</span></span>
                                    }
                                  </div>
                                </div>
                                <div style={{ textAlign: "right", flexShrink: 0, marginLeft: 12 }}>
                                  <div style={{ fontSize: "0.78rem", fontWeight: 700, color: "var(--psa-navy)" }}>{fmtDuration(s.total_hours, s.extra_minutes)}</div>
                                  <div style={{ fontSize: "0.68rem", fontWeight: 700, color: ongoing ? "#dc2626" : "#166534", marginTop: 2 }}>{ongoing ? "Ongoing" : "Scheduled"}</div>
                                </div>
                              </div>
                            );
                          })}
                          {rTotal > 1 && (
                            <div style={{ display: "flex", alignItems: "center", justifyContent: "flex-end", gap: 8, paddingTop: 10 }}>
                              <button onClick={() => setSuspDetailPage(p => Math.max(0, p - 1))} disabled={rSafe === 0}
                                style={{ background: "none", border: "1px solid var(--border)", borderRadius: 6, padding: "5px 14px", fontSize: "0.9rem", fontWeight: 700, cursor: rSafe === 0 ? "default" : "pointer", color: rSafe === 0 ? "var(--text-muted)" : "var(--psa-navy)" }}>‹</button>
                              <span style={{ fontSize: "0.75rem", color: "var(--text-muted)", minWidth: 40, textAlign: "center" }}>{rSafe + 1} / {rTotal}</span>
                              <button onClick={() => setSuspDetailPage(p => Math.min(rTotal - 1, p + 1))} disabled={rSafe === rTotal - 1}
                                style={{ background: "none", border: "1px solid var(--border)", borderRadius: 6, padding: "5px 14px", fontSize: "0.9rem", fontWeight: 700, cursor: rSafe === rTotal - 1 ? "default" : "pointer", color: rSafe === rTotal - 1 ? "var(--text-muted)" : "var(--psa-navy)" }}>›</button>
                            </div>
                          )}
                        </>
                      );
                    })()}

                    {/* ── Position History ── */}
                    {suspPosHistory.length > 0 && (() => {
                      const PH = 3;
                      const phTotal = Math.ceil(suspPosHistory.length / PH) || 1;
                      const phSafe  = Math.min(suspPosPage, phTotal - 1);
                      const phItems = suspPosHistory.slice(phSafe * PH, (phSafe + 1) * PH);
                      const fmtD = d => d ? new Date(d).toLocaleDateString("en-PH", { month: "short", day: "numeric", year: "numeric" }) : "—";
                      return (
                        <>
                          <div style={{ fontSize: "0.65rem", fontWeight: 800, color: "var(--text)", textTransform: "uppercase", letterSpacing: "0.12em", marginTop: 20, marginBottom: 8 }}>
                            Position History ({suspPosHistory.length})
                          </div>
                          {phItems.map((ph, i) => {
                            const isCurrent = ph.status === "active";
                            return (
                              <div key={ph.id ?? i} style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", padding: "6px 0", borderBottom: i < phItems.length - 1 ? "1px solid var(--border)" : "none" }}>
                                <div style={{ minWidth: 0, flex: 1 }}>
                                  <div style={{ fontSize: "0.78rem", fontWeight: 600, color: isCurrent ? "var(--psa-navy)" : "var(--text)" }}>
                                    {ph.position_title}
                                    {isCurrent && <span style={{ marginLeft: 6, fontSize: "0.63rem", fontWeight: 700, background: "#dbeafe", color: "#1d4ed8", borderRadius: 4, padding: "1px 5px" }}>Current</span>}
                                  </div>
                                  <div style={{ fontSize: "0.68rem", color: "var(--text-muted)", marginTop: 2 }}>
                                    {ph.movement_type && <span style={{ marginRight: 6, textTransform: "capitalize" }}>{ph.movement_type}</span>}
                                    {fmtD(ph.start_date)}{ph.end_date ? ` → ${fmtD(ph.end_date)}` : ""}
                                  </div>
                                </div>
                              </div>
                            );
                          })}
                          {phTotal > 1 && (
                            <div style={{ display: "flex", alignItems: "center", justifyContent: "flex-end", gap: 8, paddingTop: 8, marginTop: 4 }}>
                              <button onClick={() => setSuspPosPage(p => Math.max(0, p - 1))} disabled={phSafe === 0}
                                style={{ background: "none", border: "1px solid var(--border)", borderRadius: 6, padding: "5px 14px", fontSize: "0.9rem", fontWeight: 700, cursor: phSafe === 0 ? "default" : "pointer", color: phSafe === 0 ? "var(--text-muted)" : "var(--psa-navy)" }}>‹</button>
                              <span style={{ fontSize: "0.75rem", color: "var(--text-muted)", minWidth: 40, textAlign: "center" }}>{phSafe + 1} / {phTotal}</span>
                              <button onClick={() => setSuspPosPage(p => Math.min(phTotal - 1, p + 1))} disabled={phSafe === phTotal - 1}
                                style={{ background: "none", border: "1px solid var(--border)", borderRadius: 6, padding: "5px 14px", fontSize: "0.9rem", fontWeight: 700, cursor: phSafe === phTotal - 1 ? "default" : "pointer", color: phSafe === phTotal - 1 ? "var(--text-muted)" : "var(--psa-navy)" }}>›</button>
                            </div>
                          )}
                        </>
                      );
                    })()}
                  </>
                ) : (
                  <div style={{ color: "var(--text-muted)", fontSize: "0.83rem", padding: "40px 0", textAlign: "center" }}>
                    Select an employee to view their profile and records.
                  </div>
                )}
              </div>
            </div>
          );
        })()}

      {showAdd && (
        <EmployeeFormModal
          title="Add Employee"
          onSave={async (data) => {
            const res = await fetch(`${apiUrl}/employees`, {
              method: "POST", credentials: "include",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify(data),
            });
            if (res.ok) {
              if (data.positionId) fetch(`${apiUrl}/positions/${data.positionId}/search-hit`, { method: "POST", credentials: "include" }).catch(() => {});
              if (data.divisionId) fetch(`${apiUrl}/divisions/${data.divisionId}/search-hit`, { method: "POST", credentials: "include" }).catch(() => {});
              setShowAdd(false); fetchEmployees(query);
            } else { const d = await res.json(); alert(d.error ?? "Failed to add employee."); }
          }}
          onClose={() => setShowAdd(false)}
        />
      )}

      {editEmp && (
        <EmployeeFormModal
          title="Edit & Promote"
          initial={editEmp}
          onSave={async (data) => {
            const res = await fetch(`${apiUrl}/employees/${editEmp.id}`, {
              method: "PATCH", credentials: "include",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify(data),
            });
            if (res.ok) {
              if (data.positionId) fetch(`${apiUrl}/positions/${data.positionId}/search-hit`, { method: "POST", credentials: "include" }).catch(() => {});
              if (data.divisionId) fetch(`${apiUrl}/divisions/${data.divisionId}/search-hit`, { method: "POST", credentials: "include" }).catch(() => {});
              setEditEmp(null); fetchEmployees(query);
            } else { const d = await res.json(); alert(d.error ?? "Failed to update employee."); }
          }}
          onClose={() => setEditEmp(null)}
        />
      )}

      {/* ── Activate confirmation modal ── */}
      {activateConfirm && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center" }}>
          <div style={{ background: "var(--surface)", borderRadius: "var(--radius-lg)", padding: "28px 32px", maxWidth: 420, width: "90%", boxShadow: "0 8px 32px rgba(0,0,0,0.18)" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 14 }}>
              <div style={{ width: 36, height: 36, borderRadius: "50%", background: "#fef9c3", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#ca8a04" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>
              </div>
              <div>
                <div style={{ fontSize: "0.95rem", fontWeight: 700 }}>End Leave Now?</div>
                <div style={{ fontSize: "0.78rem", color: "var(--text-muted)", marginTop: 2 }}>{activateConfirm.name}</div>
              </div>
            </div>
            <p style={{ fontSize: "0.83rem", color: "var(--text-muted)", lineHeight: 1.6, margin: "0 0 20px" }}>
              You are about to manually end this employee's <strong style={{ color: "var(--text)" }}>{activateConfirm.separation_reason}</strong> and restore their active status effective <strong style={{ color: "var(--text)" }}>right now</strong>.
              {activateConfirm.end_date && (
                <span> Their scheduled return date was <strong style={{ color: "var(--text)" }}>{new Date(activateConfirm.end_date).toLocaleString("en-PH", { month: "short", day: "numeric", year: "numeric", hour: "2-digit", minute: "2-digit" })}</strong>.</span>
              )}
            </p>
            <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
              <button type="button" onClick={() => setActivateConfirm(null)} disabled={activating === activateConfirm.id}
                style={{ padding: "8px 18px", borderRadius: "var(--radius-sm)", border: "1px solid var(--border)", background: "var(--surface-2)", color: "var(--text-muted)", fontWeight: 600, fontSize: "0.83rem", cursor: "pointer" }}>
                Cancel
              </button>
              <button type="button" onClick={() => handleActivate(activateConfirm.id)} disabled={activating === activateConfirm.id}
                style={{ padding: "8px 20px", borderRadius: "var(--radius-sm)", border: "none", background: "#166534", color: "#fff", fontWeight: 700, fontSize: "0.83rem", cursor: activating === activateConfirm.id ? "wait" : "pointer", opacity: activating === activateConfirm.id ? 0.7 : 1 }}>
                {activating === activateConfirm.id ? "Activating…" : "Yes, End Leave Now"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ── Permanent Separation modal (Leave Log only) ── */}
      {permSepTarget && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
          <div style={{ background: "var(--surface)", borderRadius: "var(--radius)", boxShadow: "0 8px 32px rgba(0,0,0,0.2)", width: "100%", maxWidth: 420, padding: 28 }}>
            <div style={{ marginBottom: 6 }}>
              <div style={{ fontWeight: 800, fontSize: "1.05rem", color: "var(--psa-navy)", letterSpacing: "-0.01em" }}>Permanent Separation</div>
              <div style={{ fontSize: "0.78rem", color: "var(--text-muted)", marginTop: 2 }}>{permSepTarget.name}</div>
            </div>

            {!permSepConfirm ? (
              /* ── Step 1: pick reason ── */
              <>
                <p style={{ fontSize: "0.82rem", color: "var(--text-muted)", lineHeight: 1.6, margin: "12px 0 16px" }}>
                  This is <strong style={{ color: "var(--text)" }}>permanent and irreversible</strong>. All deductions will be deactivated and the employee account will be closed. Select the reason:
                </p>
                <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 20 }}>
                  {[
                    { val: "Resigned",   label: "Resigned",   note: "Voluntary separation" },
                    { val: "Retired",    label: "Retired",    note: "End of service" },
                    { val: "Terminated", label: "Terminated", note: "Involuntary separation" },
                  ].map(({ val, label, note }) => (
                    <label key={val} style={{ display: "flex", alignItems: "center", gap: 10, cursor: "pointer", padding: "9px 12px", borderRadius: "var(--radius-sm)", border: `1.5px solid ${permSepReason === val ? "#ef4444" : "var(--border)"}`, background: permSepReason === val ? "#fef2f2" : "var(--surface-2)" }}>
                      <input type="radio" name="perm-sep-reason" value={val} checked={permSepReason === val} onChange={() => setPermSepReason(val)} style={{ accentColor: "#ef4444", width: 15, height: 15, flexShrink: 0 }} />
                      <div>
                        <div style={{ fontWeight: 600, fontSize: "0.85rem" }}>{label}</div>
                        <div style={{ fontSize: "0.7rem", color: "var(--text-muted)" }}>{note}</div>
                      </div>
                    </label>
                  ))}
                </div>
                <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
                  <button type="button" className="ghost-button" onClick={() => { setPermSepTarget(null); setPermSepConfirm(false); }} style={{ minWidth: 90 }}>Cancel</button>
                  <button type="button" onClick={() => setPermSepConfirm(true)}
                    style={{ minWidth: 130, padding: "8px 20px", background: "#b91c1c", color: "#fff", fontWeight: 700, fontSize: "0.9rem", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>
                    Continue — {permSepReason}
                  </button>
                </div>
              </>
            ) : (
              /* ── Step 2: final irreversible warning ── */
              <>
                <div style={{ margin: "16px 0", padding: "14px 16px", borderRadius: "var(--radius-sm)", border: "2px solid #dc2626", background: "#fef2f2" }}>
                  <div style={{ fontSize: "0.85rem", fontWeight: 800, color: "#991b1b", marginBottom: 6 }}>⚠ Last Warning — This cannot be undone</div>
                  <div style={{ fontSize: "0.82rem", color: "#7f1d1d", lineHeight: 1.6 }}>
                    You are about to permanently mark <strong>{permSepTarget.name}</strong> as <strong>{permSepReason}</strong>.
                    All payroll deductions will be deactivated and this employee record will be closed.
                    There is no way to reverse this action.
                  </div>
                </div>
                <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
                  <button type="button" className="ghost-button" onClick={() => setPermSepConfirm(false)} disabled={permSepping} style={{ minWidth: 90 }}>Go Back</button>
                  <button type="button" onClick={async () => { await handlePermSep(); setPermSepConfirm(false); }} disabled={permSepping}
                    style={{ minWidth: 130, padding: "8px 20px", background: permSepping ? "var(--text-muted)" : "#b91c1c", color: "#fff", fontWeight: 700, fontSize: "0.9rem", cursor: permSepping ? "wait" : "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>
                    {permSepping ? "Processing…" : `Yes, ${permSepReason} — Finalize`}
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      )}

      {/* ── Link User Account modal ── */}
      {linkDropOpen && suspSelected && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
          <div style={{ background: "var(--surface)", borderRadius: "var(--radius)", boxShadow: "0 8px 32px rgba(0,0,0,0.2)", width: "100%", maxWidth: 420, padding: 28 }}>
            {/* Header */}
            <div style={{ marginBottom: 6 }}>
              <div style={{ fontWeight: 800, fontSize: "1.05rem", color: "var(--psa-navy)", letterSpacing: "-0.01em" }}>Link User Account</div>
              <div style={{ fontSize: "0.78rem", color: "var(--text-muted)", marginTop: 2 }}>{suspSelected.name}</div>
            </div>
            <p style={{ fontSize: "0.82rem", color: "var(--text-muted)", lineHeight: 1.6, margin: "12px 0 14px" }}>
              Select a user account to link to this employee. Linked users can view their payslip through <strong style={{ color: "var(--text)" }}>Employee Dashboard</strong>.
            </p>

            {/* Search — same style as employee search panel */}
            <div style={{ position: "relative", marginBottom: 10 }}>
              <svg style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)", pointerEvents: "none" }} width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="var(--text-muted)" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
              </svg>
              <input
                autoFocus
                type="text"
                placeholder="Search by username…"
                value={linkSearch}
                onChange={e => setLinkSearch(e.target.value)}
                style={{ width: "100%", boxSizing: "border-box", paddingLeft: 32, paddingRight: linkSearch ? 28 : 12, paddingTop: 8, paddingBottom: 8, fontSize: "0.83rem", border: "1px solid var(--border)", borderRadius: 8, background: "var(--surface-2)", color: "var(--text)", outline: "none" }}
              />
              {linkSearch && (
                <button type="button" onClick={() => setLinkSearch("")} style={{ position: "absolute", right: 8, top: "50%", transform: "translateY(-50%)", border: "none", background: "none", cursor: "pointer", padding: 0, color: "var(--text-muted)", fontSize: "1rem", lineHeight: 1 }}>×</button>
              )}
            </div>

            {/* User list — top 5 by default, all results when searching */}
            <div style={{ display: "flex", flexDirection: "column", maxHeight: 260, overflowY: "auto", border: "1px solid var(--border)", borderRadius: 8, overflow: "hidden", marginBottom: 18 }}>
              {(() => {
                const q = linkSearch.toLowerCase();
                const filtered = q
                  ? hAvailUsers.filter(u => u.username.toLowerCase().includes(q))
                  : [
                      ...hAvailUsers.filter(u => u.id === hLinkedUser?.id),
                      ...hAvailUsers.filter(u => u.id !== hLinkedUser?.id),
                    ].slice(0, 5);

                const rows = [];

                if (!q) {
                  rows.push(
                    <div
                      key="__nolink"
                      onClick={() => setLinkModalSel("")}
                      style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 14px", cursor: "pointer", background: linkModalSel === "" ? "#eff6ff" : "transparent", borderBottom: "1px solid var(--border)" }}
                    >
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke={linkModalSel === "" ? "var(--psa-navy)" : "var(--text-muted)"} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                        <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
                      </svg>
                      <div style={{ flex: 1 }}>
                        <div style={{ fontSize: "0.83rem", fontWeight: linkModalSel === "" ? 700 : 500, color: linkModalSel === "" ? "var(--psa-navy)" : "var(--text)" }}>No link</div>
                        <div style={{ fontSize: "0.7rem", color: "var(--text-muted)" }}>Remove existing account link</div>
                      </div>
                      {linkModalSel === "" && <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="var(--psa-navy)" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"/></svg>}
                    </div>
                  );
                }

                filtered.forEach((u, i) => {
                  const isSelected = u.id === linkModalSel;
                  const isCurrent  = u.id === hLinkedUser?.id;
                  rows.push(
                    <div
                      key={u.id}
                      onClick={() => setLinkModalSel(u.id)}
                      style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 14px", cursor: "pointer", background: isSelected ? "#f0fdf4" : "transparent", borderBottom: i < filtered.length - 1 ? "1px solid var(--border)" : "none" }}
                    >
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke={isSelected ? "#166534" : "var(--text-muted)"} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                        <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/>
                      </svg>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontSize: "0.83rem", fontWeight: isSelected ? 700 : 500, color: isSelected ? "#166534" : "var(--text)" }}>{u.username}</div>
                        {u.role && <div style={{ fontSize: "0.7rem", color: "var(--text-muted)", textTransform: "capitalize" }}>{u.role}</div>}
                      </div>
                      {isCurrent && <span style={{ fontSize: "0.65rem", fontWeight: 700, color: "#166534", background: "#dcfce7", borderRadius: 4, padding: "1px 6px", flexShrink: 0 }}>Current</span>}
                      {isSelected && <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#166534" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"/></svg>}
                    </div>
                  );
                });

                if (rows.length === 0) {
                  rows.push(
                    <div key="empty" style={{ padding: "20px", fontSize: "0.78rem", color: "var(--text-muted)", textAlign: "center" }}>
                      No users match "{linkSearch}"
                    </div>
                  );
                }

                return rows;
              })()}
            </div>

            {/* Actions */}
            {!linkConfirm ? (
              /* Step 1 — pick user then proceed */
              <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
                <button type="button" className="ghost-button" onClick={() => { setLinkDropOpen(false); setLinkSearch(""); setLinkConfirm(false); }} style={{ minWidth: 90 }}>
                  Cancel
                </button>
                <button type="button" onClick={() => setLinkConfirm(true)}
                  style={{ minWidth: 130, padding: "8px 20px", background: "var(--psa-navy)", color: "#fff", fontWeight: 700, fontSize: "0.9rem", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>
                  {linkModalSel ? `Continue — ${hAvailUsers.find(u => u.id === linkModalSel)?.username ?? ""}` : "Continue — Remove Link"}
                </button>
              </div>
            ) : (
              /* Step 2 — final warning before saving */
              <>
                <div style={{ margin: "4px 0 16px", padding: "14px 16px", borderRadius: "var(--radius-sm)", border: "2px solid var(--psa-navy)", background: "#eff6ff" }}>
                  <div style={{ fontSize: "0.85rem", fontWeight: 800, color: "var(--psa-navy)", marginBottom: 6 }}>⚠ Confirm Account Link</div>
                  <div style={{ fontSize: "0.82rem", color: "#1e3a5f", lineHeight: 1.6 }}>
                    {linkModalSel
                      ? <>User <strong>{hAvailUsers.find(u => u.id === linkModalSel)?.username}</strong> will be able to view <strong>{suspSelected?.name}</strong>'s payslip and personal payroll data through Employee Dashboard. Only link the correct account.</>
                      : <>The existing account link for <strong>{suspSelected?.name}</strong> will be removed. That user will no longer see their payslip data.</>
                    }
                  </div>
                </div>
                <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
                  <button type="button" className="ghost-button" onClick={() => setLinkConfirm(false)} disabled={linkSaving} style={{ minWidth: 90 }}>
                    Go Back
                  </button>
                  <button type="button" disabled={linkSaving}
                    onClick={async () => { await handleLinkChange(linkModalSel); setLinkDropOpen(false); setLinkSearch(""); setLinkConfirm(false); }}
                    style={{ minWidth: 130, padding: "8px 20px", background: linkSaving ? "var(--text-muted)" : "var(--psa-navy)", color: "#fff", fontWeight: 700, fontSize: "0.9rem", cursor: linkSaving ? "wait" : "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>
                    {linkSaving ? "Saving…" : linkModalSel ? "Yes, Link Account" : "Yes, Remove Link"}
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      )}

      </div>

      {/* ── Deactivation reason modal (temporary only) ── */}
      {deactivatingEmp && (() => {
        const isTemporary = ["On Leave", "Suspended"].includes(separationReason);
        return (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
            <div style={{ background: "var(--surface)", borderRadius: "var(--radius)", boxShadow: "0 8px 32px rgba(0,0,0,0.2)", width: "100%", maxWidth: 420, padding: 28, maxHeight: "90vh", overflowY: "auto" }}>

              {/* Header */}
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 6 }}>
                <div>
                  <div style={{ fontWeight: 800, fontSize: "1.05rem", color: "var(--psa-navy)", letterSpacing: "-0.01em" }}>Set On Leave / Suspend</div>
                  <div style={{ fontSize: "0.78rem", color: "var(--text-muted)", marginTop: 2 }}>{deactivatingEmp.name}</div>
                </div>
              </div>

              <p style={{ fontSize: "0.82rem", color: "var(--text-muted)", lineHeight: 1.6, margin: "12px 0 16px" }}>
                This is <strong style={{ color: "var(--text)" }}>temporary</strong>. All deductions are preserved and the employee can be reactivated at any time. Select the type:
              </p>

              <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 16 }}>
                {[
                  { val: "On Leave",  note: "Temporary absence, deductions preserved" },
                  { val: "Suspended", note: "Disciplinary action, temporary" },
                ].map(({ val, note }) => (
                  <label key={val} style={{ display: "flex", alignItems: "center", gap: 10, cursor: "pointer", padding: "9px 12px", borderRadius: "var(--radius-sm)", border: `1.5px solid ${separationReason === val ? "#f97316" : "var(--border)"}`, background: separationReason === val ? "#fef2f2" : "var(--surface-2)" }}>
                    <input type="radio" name="sep-reason" value={val} checked={separationReason === val} onChange={e => setSeparationReason(e.target.value)} style={{ accentColor: "#d97706", width: 15, height: 15, flexShrink: 0 }} />
                    <div>
                      <div style={{ fontWeight: 600, fontSize: "0.85rem" }}>{val}</div>
                      <div style={{ fontSize: "0.7rem", color: "var(--text-muted)" }}>{note}</div>
                    </div>
                  </label>
                ))}
              </div>

              {/* Date/time inputs */}
              {isTemporary && (
                <div style={{ marginBottom: 20, display: "flex", flexDirection: "column", gap: 10 }}>
                  <div style={{ fontSize: "0.75rem", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.06em", color: "var(--text-muted)" }}>
                    {separationReason === "On Leave" ? "Leave Period" : "Suspension Period"}
                  </div>
                  <div>
                    <div style={{ fontSize: "0.75rem", color: "var(--text-muted)", marginBottom: 4 }}>Start Date & Time <span style={{ color: "#c0392b" }}>*</span></div>
                    <input type="datetime-local" value={suspensionStart} onChange={e => setSuspensionStart(e.target.value)}
                      style={{ width: "100%", border: "1.5px solid var(--border-2)", borderRadius: "var(--radius-sm)", padding: "8px 10px", fontSize: "0.83rem", background: "var(--surface-2)", color: "var(--text)", boxSizing: "border-box" }} />
                  </div>
                  <div>
                    <div style={{ fontSize: "0.75rem", color: "var(--text-muted)", marginBottom: 4 }}>
                      Expected End Date & Time <span style={{ fontSize: "0.7rem", color: "var(--text-muted)" }}>(leave blank if open-ended)</span>
                    </div>
                    <input type="datetime-local" value={suspensionEnd} onChange={e => setSuspensionEnd(e.target.value)} min={suspensionStart}
                      style={{ width: "100%", border: "1.5px solid var(--border-2)", borderRadius: "var(--radius-sm)", padding: "8px 10px", fontSize: "0.83rem", background: "var(--surface-2)", color: "var(--text)", boxSizing: "border-box" }} />
                  </div>
                </div>
              )}

              <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
                <button type="button" className="ghost-button" onClick={() => setDeactivatingEmp(null)} disabled={deactivating} style={{ minWidth: 90 }}>
                  Cancel
                </button>
                <button type="button" onClick={confirmDeactivate} disabled={deactivating}
                  style={{ minWidth: 130, padding: "8px 20px", background: deactivating ? "var(--text-muted)" : "#d97706", color: "#fff", fontWeight: 700, fontSize: "0.9rem", cursor: deactivating ? "wait" : "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}>
                  {deactivating ? "Processing…" : `Confirm — ${separationReason}`}
                </button>
              </div>
            </div>
          </div>
        );
      })()}
    </div>
  );
}

function InfoField({ label, value }) {
  return (
    <div>
      <div style={{ fontSize: "0.72rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 4 }}>{label}</div>
      <div style={{ fontSize: "0.95rem", fontWeight: 600, color: "var(--text)" }}>{value}</div>
    </div>
  );
}

function EmployeeFormModal({ title, initial, onSave, onClose }) {
  const [form, setForm] = useState({
    employeeNumber:      initial?.employee_number ?? "",
    name:                initial?.name ?? "",
    position:            initial?.position ?? "",
    positionId:          initial?.position_id || "",
    divisionId:          initial?.division_id || "",
    employmentType:      initial?.employment_type ?? "Regular",
    basicMonthlyPay:     initial?.basic_monthly_pay ?? "",
    startDate:           "",
    initialStep:         "",
    movementType:        "APPOINTMENT",
  });
  const [saving, setSaving]                 = useState(false);
  const [positions, setPositions]           = useState([]);
  const [divisions, setDivisions]           = useState([]);
  const [lookupsLoading, setLookupsLoading] = useState(true);

  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));

  useEffect(() => {
    Promise.all([
      fetch(`${apiUrl}/positions?limit=500`, { credentials: "include" }).then(r => r.ok ? r.json() : []),
      fetch(`${apiUrl}/divisions`, { credentials: "include" }).then(r => r.ok ? r.json() : []),
    ]).then(([pos, div]) => {
      setPositions(Array.isArray(pos) ? pos : []);
      setDivisions(Array.isArray(div) ? div : []);
    }).finally(() => setLookupsLoading(false));
  }, []);

  function handlePositionChange(id) {
    const found = positions.find(p => p.id === id);
    setForm(f => ({
      ...f,
      positionId:        id,
      position:          found?.name ?? f.position,
      salaryGradeStepId: "",
    }));
  }
  function handleDivisionChange(id) {
    set("divisionId", id);
  }
  const [confirmOpen, setConfirmOpen] = useState(false);

  async function handleSubmit(e) {
    e.preventDefault();
    if (!form.employeeNumber.trim()) { alert("Employee number is required."); return; }
    if (!form.name.trim()) { alert("Full name is required."); return; }
    if (!form.positionId) { alert("Position is required."); return; }
    if (!form.divisionId) { alert("Division is required."); return; }
    if (!initial && !form.startDate) { alert("Start date is required."); return; }
    const selectedPos = positions.find(p => p.id === form.positionId);
    if (selectedPos && !selectedPos.has_steps) {
      alert("This position has no salary grade or steps configured. Go to Positions and assign a salary grade first.");
      return;
    }
    setConfirmOpen(true);
  }

  async function handleConfirmedSave() {
    setConfirmOpen(false);
    setSaving(true);
    const payload = { ...form };
    if (!payload.startDate) delete payload.startDate;
    // Always send initialStep (as null for auto, or the picked number) so backend can clear granted_step
    payload.initialStep = payload.initialStep ? parseInt(payload.initialStep, 10) : null;
    delete payload.incrementDay;
    await onSave(payload);
    setSaving(false);
  }

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(13,35,70,0.45)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: "16px", backdropFilter: "blur(2px)" }}>
      <div style={{ background: "var(--surface)", borderRadius: 16, width: "100%", maxWidth: 560, maxHeight: "92vh", overflowY: "auto", boxShadow: "0 20px 60px rgba(0,0,0,0.25)", display: "flex", flexDirection: "column" }}>

        {/* ── Modal Header ── */}
        <div style={{ padding: "20px 24px 18px", borderBottom: "1px solid var(--border)", display: "flex", alignItems: "center", justifyContent: "space-between", flexShrink: 0 }}>
          <div>
            <div style={{ fontSize: "1.05rem", fontWeight: 800, color: "var(--psa-navy)", letterSpacing: "-0.01em" }}>{title}</div>
            <div style={{ fontSize: "0.78rem", color: "var(--text-muted)", marginTop: 2 }}>Fill in the employee details below</div>
          </div>
          <button
            type="button"
            onClick={onClose}
            aria-label="Close"
            style={{ padding: "4px 8px", background: "transparent", border: "none", color: "var(--text-muted)", cursor: "pointer", flexShrink: 0, fontSize: "1.2rem", fontWeight: 400, lineHeight: 1, boxShadow: "none", transition: "color 0.15s" }}
            onMouseEnter={e => { e.currentTarget.style.color = "#991b1b"; }}
            onMouseLeave={e => { e.currentTarget.style.color = "var(--text-muted)"; }}
          >✕</button>
        </div>

        {/* ── Form Body ── */}
        <form onSubmit={handleSubmit} style={{ padding: "22px 24px", display: "flex", flexDirection: "column", gap: 16, flex: 1 }}>
          {lookupsLoading ? (
            <div style={{ color: "var(--text-muted)", fontSize: "0.88rem", padding: "8px 0" }}>Loading reference data…</div>
          ) : (
            <>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
              <label>
                Employee Number
                <input value={form.employeeNumber} onChange={e => set("employeeNumber", e.target.value)} placeholder="e.g. EMP-0001" required />
              </label>
              <label>
                Full Name
                <input value={form.name} onChange={e => set("name", e.target.value)} placeholder="e.g. Juan Dela Cruz" required />
              </label>
              <label style={{ display: "grid", gap: 6 }}>
                Employment Type
                <CustomSelect value={form.employmentType} onChange={v => set("employmentType", v)} options={EMPLOYMENT_TYPES.map(t => ({ value: t, label: t }))} placeholder="Select type" />
              </label>
              <label style={{ display: "grid", gap: 6 }}>
                Position
                {positions.length > 0 ? (
                  <>
                  <CustomSelect
                    value={form.positionId}
                    onChange={handlePositionChange}
                    options={positions
                      .filter(p => p.has_steps || p.id === form.positionId)
                      .map(p => ({ value: p.id, label: p.name }))}
                    placeholder="Select position"
                  />
                  {!!form.positionId && !positions.find(p => p.id === form.positionId)?.has_steps && (
                    <span style={{ fontSize: "0.75rem", color: "#b91c1c" }}>
                      This position has no salary grade or steps — configure it in Positions first.
                    </span>
                  )}
                  </>
                ) : (
                  <input value={form.position} onChange={e => set("position", e.target.value)} placeholder="e.g. Administrative Officer" required />
                )}
              </label>
              {initial && (
                <label style={{ display: "grid", gap: 6 }}>
                  Movement Type
                  <select value={form.movementType} onChange={e => set("movementType", e.target.value)}>
                    <option value="APPOINTMENT">Appointment</option>
                    <option value="PROMOTION">Promotion</option>
                    <option value="TRANSFER">Transfer</option>
                  </select>
                </label>
              )}
              <label style={{ display: "grid", gap: 6 }}>
                Division
                <CustomSelect
                  value={form.divisionId}
                  onChange={handleDivisionChange}
                  options={divisions.map(d => ({ value: d.id, label: d.name }))}
                  placeholder="Select division"
                />
              </label>
              <label style={{ display: "grid", gap: 6 }}>
                {initial ? "Appointment Date (optional)" : "Appointed Date"}
                <input
                  type="date"
                  value={form.startDate}
                  onChange={e => set("startDate", e.target.value)}
                  required={!initial}
                />
              </label>
              <label style={{ display: "grid", gap: 6 }}>
                Starting Step (optional)
                <select value={form.initialStep} onChange={e => set("initialStep", e.target.value ? Number(e.target.value) : "")}>
                  <option value="">— auto from date —</option>
                  {[1,2,3,4,5,6,7,8].map(s => <option key={s} value={s}>Step {s}</option>)}
                </select>
              </label>
            </div>
            </>
          )}

          {/* ── Footer ── */}
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", paddingTop: 12, marginTop: 4, borderTop: "1px solid var(--border)" }}>
            <button type="button" className="ghost-button" onClick={onClose} style={{ minWidth: 90 }}>Cancel</button>
            <button
              type="submit"
              disabled={saving || lookupsLoading}
              style={{ minWidth: 130, background: (saving || lookupsLoading) ? "var(--text-muted)" : "var(--psa-blue)", display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}
            >
              {saving ? (
                <>
                  <svg style={{ animation: "spin 0.7s linear infinite" }} width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><circle cx="12" cy="12" r="10" strokeOpacity="0.25"/><path d="M12 2a10 10 0 0 1 10 10" /></svg>
                  Saving…
                </>
              ) : "Save Employee"}
            </button>
          </div>
        </form>
      </div>

      {/* ── Confirmation Dialog ── */}
      {confirmOpen && (() => {
        const selectedPos = positions.find(p => p.id === form.positionId);
        const gradeCode = selectedPos?.default_salary_grade;
        return (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", zIndex: 1100, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
            <div style={{ background: "var(--surface)", borderRadius: 14, maxWidth: 420, width: "100%", padding: "28px 28px 22px", boxShadow: "0 24px 64px rgba(0,0,0,0.3)" }}>
              <div style={{ fontSize: "1rem", fontWeight: 800, color: "var(--psa-navy)", marginBottom: 6 }}>Confirm Changes</div>
              <div style={{ fontSize: "0.85rem", color: "var(--text-muted)", marginBottom: 16, lineHeight: 1.6 }}>
                You are about to save the following changes for <strong>{form.name}</strong>:
              </div>
              <div style={{ background: "var(--surface-alt, #f4f6fb)", borderRadius: 8, padding: "12px 14px", fontSize: "0.83rem", lineHeight: 1.8, marginBottom: 20, border: "1px solid var(--border)" }}>
                <div><strong>Employee No.:</strong> {form.employeeNumber}</div>
                <div><strong>Position:</strong> {selectedPos?.name ?? form.position ?? "—"}</div>
                {initial && <div><strong>Movement Type:</strong> {form.movementType}</div>}
                {gradeCode && <div><strong>Salary Grade:</strong> {gradeCode}</div>}
                <div><strong>Division:</strong> {divisions.find(d => d.id === form.divisionId)?.name ?? "—"}</div>
                <div><strong>Employment Type:</strong> {form.employmentType}</div>
              </div>
              <div style={{ fontSize: "0.78rem", color: "#b91c1c", fontWeight: 600, marginBottom: 18 }}>
                This will update the employee record and cannot be undone.
              </div>
              <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
                <button type="button" className="ghost-button" onClick={() => setConfirmOpen(false)} style={{ minWidth: 90 }}>Go Back</button>
                <button type="button" onClick={handleConfirmedSave} style={{ minWidth: 130, background: "#b91c1c" }}>Confirm Save</button>
              </div>
            </div>
          </div>
        );
      })()}
    </div>
  );
}

function useIsMobile() {
  const [isMobile, setIsMobile] = useState(() => window.innerWidth < 768);
  useEffect(() => {
    const handler = () => setIsMobile(window.innerWidth < 768);
    window.addEventListener("resize", handler);
    return () => window.removeEventListener("resize", handler);
  }, []);
  return isMobile;
}

function LoansCard({ emp, isAdmin }) {
  const [loans, setLoans]               = useState([]);
  const [loansLoading, setLoansLoading] = useState(false);

  const fetchLoans = useCallback(async () => {
    setLoansLoading(true);
    try {
      const res = await fetch(`${apiUrl}/employees/${emp.id}/loans`, { credentials: "include" });
      if (res.ok) setLoans(await res.json());
    } finally { setLoansLoading(false); }
  }, [emp.id]);

  useEffect(() => { fetchLoans(); }, [fetchLoans]);

  const handleSimulatePeriod = async () => {
    setLoansLoading(true);
    try {
      const res = await fetch(`${apiUrl}/employees/${emp.id}/loans/simulate-period`, { method: "POST", credentials: "include" });
      if (res.ok) {
        const data = await res.json();
        setLoans(data.loans ?? []);
        if (data.halted?.length > 0) {
          alert(`${data.halted.length} loan(s) halted and closed — remaining balance was below monthly deduction.`);
        } else if (data.deducted?.length > 0) {
          alert(data.deducted.map(d => `✓ ${d.loan_type_name ?? "Loan"}: deducted ₱${Number(d.monthly_amount ?? d.monthly ?? 0).toLocaleString("en-PH", { minimumFractionDigits: 2 })} — remaining ₱${Number(d.newBalance ?? 0).toLocaleString("en-PH", { minimumFractionDigits: 2 })}`).join("\n"));
        } else {
          alert("No deductions were made. Check that the loan has an active monthly amount set.");
        }
      } else { const e = await res.json().catch(() => ({})); alert(e.error || "Simulate failed."); }
    } finally { setLoansLoading(false); }
  };

  const handleCloseLoan = async (loan) => {
    if (!window.confirm(`Close "${loan.loan_type_name}"? Remaining: ₱${Number(loan.remaining_balance).toLocaleString("en-PH", { minimumFractionDigits: 2 })}`)) return;
    const res = await fetch(`${apiUrl}/employees/${emp.id}/loans/${loan.id}/close`, { method: "PATCH", credentials: "include" });
    if (res.ok) await fetchLoans();
    else { const e = await res.json(); alert(e.error || "Failed to close loan."); }
  };

  const fmtLoan = v => "₱" + Number(v).toLocaleString("en-PH", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const fmtDate = d => d ? new Date(d).toLocaleDateString("en-PH", { year: "numeric", month: "short", day: "numeric" }) : "—";

  const [loanSection, setLoanSection] = useState("active");
  const [selType,     setSelType]     = useState("");
  const [loanPage,    setLoanPage]    = useState(0);

  const activeLoans  = loans.filter(l => l.status === "active");
  const closedLoans  = loans.filter(l => l.status !== "active");
  const sectionLoans = loanSection === "active" ? activeLoans : closedLoans;

  // Build type groups (same pattern as LedgerAccordion)
  const typeGroups = useMemo(() => {
    const map = {}, order = [];
    sectionLoans.forEach(loan => {
      const key = loan.loan_type_name || "Unknown";
      if (!map[key]) { map[key] = []; order.push(key); }
      map[key].push(loan);
    });
    return order.map(k => ({ key: k, loans: map[k] }));
  }, [sectionLoans]);

  // Auto-select first type when section changes
  useEffect(() => {
    if (typeGroups.length > 0 && (!selType || !typeGroups.find(g => g.key === selType))) {
      setSelType(typeGroups[0]?.key ?? "");
      setLoanPage(0);
    }
  }, [typeGroups]);

  const switchSection = s => { setLoanSection(s); setSelType(""); setLoanPage(0); };

  const selGroup = typeGroups.find(g => g.key === selType) ?? typeGroups[0] ?? null;
  const rows     = selGroup?.loans ?? [];
  const safePg   = Math.min(loanPage, Math.max(0, rows.length - 1));

  return (
    <div className="panel" style={{ padding: 0, overflow: "hidden" }}>
      {/* Header */}
      <div style={{ padding: "12px 16px", borderBottom: "1px solid var(--border)", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div>
          <div style={{ fontSize: "0.65rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: 1 }}>Loan Overview</div>
          <div style={{ fontWeight: 700, fontSize: "0.92rem", color: "var(--text)" }}>{emp.name}</div>
        </div>
      </div>

      {loansLoading ? (
        <div style={{ padding: "60px 24px", textAlign: "center", color: "var(--text-muted)" }}>
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ animation: "spin 1s linear infinite", display: "block", margin: "0 auto 12px" }}>
            <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
          </svg>
          <div style={{ fontSize: "0.82rem" }}>Loading…</div>
        </div>
      ) : loans.length === 0 ? (
        <div style={{ padding: "32px 16px", textAlign: "center", color: "var(--text-muted)", fontSize: "0.85rem" }}>No loans on record.</div>
      ) : (
        <>
          {/* Section pills — Active / Closed */}
          <div style={{ padding: "8px 14px", borderBottom: "1px solid var(--border)", background: "var(--surface-2)", display: "flex", gap: 4, alignItems: "center", flexWrap: "wrap" }}>
            {[["active", "Active", activeLoans.length], ["closed", "Closed", closedLoans.length]].map(([key, label, count]) => {
              if (count === 0) return null;
              const isSel = loanSection === key;
              return (
                <button key={key} onClick={() => switchSection(key)} style={{
                  padding: "4px 12px", fontSize: "0.74rem", fontWeight: 600, borderRadius: 999, cursor: "pointer", border: "none", outline: "none",
                  color: isSel ? "#fff" : "var(--text-muted)",
                  background: isSel ? "#0038a8" : "transparent",
                  boxShadow: isSel ? "none" : "inset 0 0 0 1px var(--border)",
                }}>{label}</button>
              );
            })}
          </div>

          {/* TYPE select */}
          {typeGroups.length > 1 && (
            <TypeDropdown
              groups={typeGroups.map(g => ({ key: g.key, label: g.key }))}
              selKey={selType}
              onSelect={key => { setSelType(key); setLoanPage(0); }}
            />
          )}
          {typeGroups.length === 1 && (
            <div style={{ padding: "6px 14px", borderBottom: "1px solid var(--border)", display: "flex", alignItems: "center", gap: 8 }}>
              <span style={{ fontSize: "0.6rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.1em" }}>Type</span>
              <span style={{ fontSize: "0.88rem", fontWeight: 500, color: "var(--text)" }}>{selGroup?.key}</span>
            </div>
          )}

          {/* One loan at a time — formal record view */}
          {rows.length === 0 ? (
            <div style={{ padding: "28px 16px", textAlign: "center", color: "var(--text-muted)", fontSize: "0.82rem" }}>No entries</div>
          ) : (() => {
            const loan = rows[safePg];
            const total = Number(loan.total_amount), remaining = Number(loan.remaining_balance), monthly = Number(loan.monthly_amount);
            const paid = total - remaining;
            const pct = total > 0 ? Math.max(0, Math.min(100, (paid / total) * 100)) : 0;
            const monthsLeft = monthly > 0 ? Math.ceil(remaining / monthly) : null;
            const isLoanActive = loan.status === "active";
            const willHalt = isLoanActive && remaining < monthly && monthly > 0;
            const dRow = (label, value, valStyle = {}) => (
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "7px 16px" }}>
                <span style={{ fontSize: "0.82rem", color: "var(--text-muted)" }}>{label}</span>
                <span style={{ fontSize: "0.83rem", fontWeight: 500, fontVariantNumeric: "tabular-nums", ...valStyle }}>{value}</span>
              </div>
            );
            return (
              <div>
                {/* Record header */}
                <div style={{ padding: "12px 16px 4px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <span style={{ fontWeight: 700, fontSize: "0.95rem", color: "var(--text)" }}>Loan #{loan.id}</span>
                    {willHalt && <span style={{ fontSize: "0.68rem", fontWeight: 700, padding: "2px 9px", borderRadius: 999, background: "#fef9c3", color: "#854d0e" }}>⚠ Will halt</span>}
                  </div>
                  {loanSection === "active" && isAdmin && isLoanActive && (
                    <button type="button" onClick={() => handleCloseLoan(loan)}
                      style={{ fontSize: "0.72rem", fontWeight: 700, padding: "3px 10px", borderRadius: 5, background: "#fee2e2", color: "#991b1b", border: "1px solid #fecaca", cursor: "pointer" }}>
                      Close
                    </button>
                  )}
                </div>
                <div style={{ padding: "0 16px 12px", fontSize: "0.7rem", color: "var(--text-muted)", fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.08em" }}>
                  Started &nbsp; {fmtDate(loan.start_date)}
                </div>

                {dRow("Total Loan", fmtLoan(total))}
                {dRow("Monthly Deduction", fmtLoan(monthly))}
                {dRow("Amount Paid", fmtLoan(paid))}
                {dRow("Remaining Balance", fmtLoan(remaining), { color: remaining > 0 ? "#dc2626" : "#16a34a", fontWeight: 700 })}
                {loan.closed_date && dRow("End Date", fmtDate(loan.closed_date))}

                <div style={{ padding: "10px 16px 16px" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
                    <span style={{ fontSize: "0.75rem", color: "var(--text-muted)" }}>Progress</span>
                    <span style={{ fontSize: "0.75rem", color: "var(--text-muted)" }}>
                      {pct >= 100 ? "Paid off" : monthsLeft != null ? `~${monthsLeft} mo. left` : "—"} &nbsp;({pct.toFixed(1)}%)
                    </span>
                  </div>
                  <div style={{ height: 6, background: "var(--border)", borderRadius: 99, overflow: "hidden" }}>
                    <div style={{ height: "100%", width: `${pct}%`, background: pct >= 100 ? "#16a34a" : "var(--psa-navy)", borderRadius: 99, transition: "width 0.4s" }} />
                  </div>
                </div>

                {/* Pagination — one loan per page */}
                {rows.length > 1 && (
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "10px 16px", borderTop: "1px solid var(--border)" }}>
                    <span style={{ fontSize: "0.75rem", color: "var(--text-muted)" }}>{safePg + 1} of {rows.length} loan{rows.length > 1 ? "s" : ""}</span>
                    <div style={{ display: "flex", gap: 6 }}>
                      <button onClick={() => setLoanPage(p => Math.max(0, p - 1))} disabled={safePg === 0}
                        style={{ background: "none", border: "1px solid var(--border)", borderRadius: 6, padding: "5px 14px", fontSize: "0.9rem", cursor: safePg === 0 ? "default" : "pointer", color: safePg === 0 ? "var(--text-muted)" : "var(--psa-navy)", fontWeight: 700 }}>‹</button>
                      <button onClick={() => setLoanPage(p => Math.min(rows.length - 1, p + 1))} disabled={safePg === rows.length - 1}
                        style={{ background: "none", border: "1px solid var(--border)", borderRadius: 6, padding: "5px 14px", fontSize: "0.9rem", cursor: safePg === rows.length - 1 ? "default" : "pointer", color: safePg === rows.length - 1 ? "var(--text-muted)" : "var(--psa-navy)", fontWeight: 700 }}>›</button>
                    </div>
                  </div>
                )}
              </div>
            );
          })()}
        </>
      )}
    </div>
  );
}

function EmployeeDeductionCard({ emp, isAdmin, compact = false, hideHeader = false, hideLoans = false, onEmpPatch, onTakeHomePay }) {
  const isMobile = useIsMobile();
  const { dedTypes, earnTypes } = useContext(TypesContext);
  const LESS_DED = dedTypes.length > 0 ? buildCatalogFromTypes(dedTypes, 'less_deduction') : LESS_DEDUCTIONS_CATALOG;
  const DED      = dedTypes.length > 0 ? buildCatalogFromTypes(dedTypes, 'deduction')      : DEDUCTIONS_CATALOG;
  const peraActive = earnTypes.length === 0 || earnTypes.some(t => t.name === 'pera');
  const customEarnTypes = earnTypes.filter(t => t.name !== 'basic_pay' && t.name !== 'pera');

  const [deductions, setDeductions]         = useState({});
  const [loading, setLoading]               = useState(false);
  const [quincenaOverride, setQuincenaOverride] = useState(emp.first_quincena_override ?? null);

  useEffect(() => { setQuincenaOverride(emp.first_quincena_override ?? null); }, [emp.first_quincena_override]);

  const fetchDeductions = useCallback(async () => {
    setLoading(true);
    try {
      const res = await fetch(`${apiUrl}/employees/${emp.id}/deductions`, { credentials: "include" });
      if (res.ok) {
        const raw = await res.json();
        setDeductions(raw && typeof raw === "object" && !Array.isArray(raw) ? raw : {});
      }
    } finally { setLoading(false); }
  }, [emp.id]);

  useEffect(() => { fetchDeductions(); }, [fetchDeductions]);

  const basicPay = Number(emp.basic_monthly_pay) || 0;
  const pera     = peraActive ? Number(deductions.pera ?? PERA_AMOUNT) : 0;
  const customEarnTotal = customEarnTypes.reduce((sum, t) => {
    const val = t.computation_mode === 'fixed' && t.amount_fixed != null
      ? Number(t.amount_fixed)
      : Number(deductions[t.name] ?? 0);
    return sum + val;
  }, 0);
  const grossPay = basicPay + pera + customEarnTotal;

  const lessDedTotal      = sumCatalogComputed(LESS_DED, deductions, basicPay) + sumCatalogTyped(LESS_DED, deductions);
  const dedColumnTotal    = sumCatalogComputed(DED, deductions, basicPay) + sumCatalogTyped(DED, deductions);
  const totalSalaryEarned = grossPay - lessDedTotal;
  const takeHomePay       = totalSalaryEarned - dedColumnTotal;
  useEffect(() => { if (onTakeHomePay) onTakeHomePay(takeHomePay); }, [takeHomePay, onTakeHomePay]);
  const firstQuincena      = quincenaOverride != null
    ? Number(quincenaOverride)
    : Math.round(takeHomePay / 2 * 100) / 100;
  const secondQuincena     = Math.round((takeHomePay - firstQuincena) * 100) / 100;
  const quincenaOverridden = quincenaOverride != null;

  async function handleSave(col, value) {
    const res = await fetch(`${apiUrl}/employees/${emp.id}/deductions`, {
      method: "PUT", credentials: "include",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ column: col, value }),
    });
    if (res.ok) setDeductions(await res.json());
    else { const d = await res.json(); alert(d.error ?? "Failed to save deduction."); }
  }

  async function handleRemove(col) {
    const res = await fetch(`${apiUrl}/employees/${emp.id}/deductions`, {
      method: "PUT", credentials: "include",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ column: col, value: null }),
    });
    if (res.ok) setDeductions(await res.json());
  }

  async function handleQuincenaSave(which, value) {
    const parsed = parseFloat(value);
    if (isNaN(parsed) || parsed < 0 || parsed > takeHomePay) return;
    const newFirst = which === "first" ? parsed : Math.round((takeHomePay - parsed) * 100) / 100;
    setQuincenaOverride(newFirst);
    const res = await fetch(`${apiUrl}/employees/${emp.id}`, {
      method: "PATCH", credentials: "include",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ firstQuincenaOverride: newFirst }),
    });
    if (res.ok) {
      if (onEmpPatch) onEmpPatch({ first_quincena_override: newFirst });
    } else {
      setQuincenaOverride(emp.first_quincena_override ?? null);
      alert("Failed to save quincena override. Please try again.");
    }
  }

  async function handleQuincenaReset() {
    setQuincenaOverride(null);
    const res = await fetch(`${apiUrl}/employees/${emp.id}`, {
      method: "PATCH", credentials: "include",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ firstQuincenaOverride: null }),
    });
    if (res.ok) {
      if (onEmpPatch) onEmpPatch({ first_quincena_override: null });
    } else {
      setQuincenaOverride(emp.first_quincena_override ?? null);
      alert("Failed to reset quincena override.");
    }
  }

  // ── Loans ──
  const [loans, setLoans]               = useState([]);
  const [loansLoading, setLoansLoading] = useState(false);
  const [showAddLoan, setShowAddLoan]   = useState(false);
  const [loanSaving, setLoanSaving]     = useState(false);
  const [addLoanForm, setAddLoanForm]   = useState({ payrollTypeId: "", totalAmount: "", amountPerDeduction: "", startDate: "" });
  const loanTypes = dedTypes.filter(d => d.is_loan);

  const fetchLoans = useCallback(async () => {
    setLoansLoading(true);
    try {
      const res = await fetch(`${apiUrl}/employees/${emp.id}/loans`, { credentials: "include" });
      if (res.ok) setLoans(await res.json());
    } finally { setLoansLoading(false); }
  }, [emp.id]);

  useEffect(() => { fetchLoans(); }, [fetchLoans]);

  const handleSimulatePeriod = async () => {
    setLoansLoading(true);
    try {
      const res = await fetch(`${apiUrl}/employees/${emp.id}/loans/simulate-period`, {
        method: "POST", credentials: "include",
      });
      if (res.ok) {
        const data = await res.json();
        setLoans(data.loans ?? []);
        await fetchDeductions();
        if (data.halted?.length > 0) {
          alert(`${data.halted.length} loan(s) halted and closed — remaining balance was below monthly deduction.`);
        } else if (data.deducted?.length > 0) {
          const msgs = data.deducted.map(d => {
            const amt = Number(d.monthly_amount ?? d.monthly ?? 0);
            const rem = Number(d.newBalance ?? 0);
            return `✓ ${d.loan_type_name ?? "Loan"}: deducted ₱${amt.toLocaleString("en-PH", { minimumFractionDigits: 2 })} — new remaining ₱${rem.toLocaleString("en-PH", { minimumFractionDigits: 2 })}`;
          });
          alert(msgs.join("\n"));
        } else {
          alert("No deductions were made.\n\nPossible reason: no active loan found, or monthly amount is ₱0.\nCheck Finance Audit → Loans to verify the loan has an active monthly deduction set.");
        }
      } else {
        const e = await res.json().catch(() => ({}));
        alert(e.error || "Simulate failed.");
      }
    } finally { setLoansLoading(false); }
  };

  const handleAddLoan = async () => {
    if (!addLoanForm.payrollTypeId || !addLoanForm.totalAmount || !addLoanForm.amountPerDeduction || !addLoanForm.startDate) return;
    setLoanSaving(true);
    try {
      const res = await fetch(`${apiUrl}/employees/${emp.id}/loans`, {
        method: "POST", credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ payrollTypeId: Number(addLoanForm.payrollTypeId), totalAmount: Number(addLoanForm.totalAmount), amountPerDeduction: Number(addLoanForm.amountPerDeduction), startDate: addLoanForm.startDate }),
      });
      if (res.ok) {
        setShowAddLoan(false);
        setAddLoanForm({ payrollTypeId: "", totalAmount: "", amountPerDeduction: "", startDate: "" });
        await fetchLoans();
        await fetchDeductions();
      } else {
        const e = await res.json();
        alert(e.error || "Failed to add loan.");
      }
    } finally { setLoanSaving(false); }
  };

  const handleCloseLoan = async (loan) => {
    if (!window.confirm(`Close "${loan.loan_type_name}"? Remaining: ₱${Number(loan.remaining_balance).toLocaleString("en-PH", { minimumFractionDigits: 2 })}`)) return;
    const res = await fetch(`${apiUrl}/employees/${emp.id}/loans/${loan.id}/close`, { method: "PATCH", credentials: "include" });
    if (res.ok) { await fetchLoans(); await fetchDeductions(); }
    else { const e = await res.json(); alert(e.error || "Failed to close loan."); }
  };

  const loanCols = new Set(loanTypes.map(t => t.name));
  const loanByCol = {};
  for (const loan of loans) {
    if (loan.status !== "active") continue;
    if (loan.loan_type_col) loanByCol[loan.loan_type_col] = loan;
  }

  const handleAddLoanFromPicker = async (item, totalAmount, monthlyAmount) => {
    const typeMatch = dedTypes.find(t => t.name === item.col);
    if (!typeMatch) return alert("Loan type not found in configuration.");
    const today = new Date().toISOString().split("T")[0];
    try {
      const res = await fetch(`${apiUrl}/employees/${emp.id}/loans`, {
        method: "POST", credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ payrollTypeId: typeMatch.id, totalAmount, amountPerDeduction: monthlyAmount, startDate: today }),
      });
      if (res.ok) { await fetchLoans(); await fetchDeductions(); }
      else { const e = await res.json(); alert(e.error || "Failed to add loan."); }
    } catch { alert("Network error adding loan."); }
  };

  const fmtLoan = v => "₱" + Number(v).toLocaleString("en-PH", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const loanInputStyle = { width: "100%", boxSizing: "border-box", padding: "7px 10px", fontSize: "0.83rem", border: "1.5px solid var(--border)", borderRadius: 6, background: "var(--surface-2)", color: "var(--text)" };
  const loanLabelStyle = { fontSize: "0.72rem", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.06em", color: "var(--text-muted)", marginBottom: 3, display: "block" };

  return (
    <div>
      {!hideHeader && (compact ? (
        /* Mobile accordion: skip name (already in header above), just show take home pay */
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 16px", borderBottom: "1px solid var(--border)", background: "#f0f6ff" }}>
          <span style={{ fontSize: "0.78rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.06em" }}>Take Home Pay</span>
          <span style={{ fontSize: "1.05rem", fontWeight: 800, color: "#166534" }}>₱ {fmt(takeHomePay)}</span>
        </div>
      ) : (
        /* Desktop: full name + position + take home pay header */
        <div className="job-card__header" style={{ borderBottom: "1px solid var(--border)", background: "#f0f6ff" }}>
          <div className="job-card__info">
            <div className="job-card__name">{emp.name}</div>
            <div className="job-card__tags">
              <span className="job-card__tag">{emp.employee_number}</span>
              <span className="job-card__tag-sep">·</span>
              <span className="job-card__tag">{emp.position}</span>
              {emp.office_division && <><span className="job-card__tag-sep">·</span><span className="job-card__tag">{emp.office_division}</span></>}
            </div>
          </div>
          <div style={{ textAlign: "right", flexShrink: 0 }}>
            <div style={{ fontSize: "0.68rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.06em" }}>Take Home Pay</div>
            <div style={{ fontSize: "1.1rem", fontWeight: 800, color: "#166534" }}>₱ {fmt(takeHomePay)}</div>
          </div>
        </div>
      ))}

      {/* Deduction breakdown — always visible */}
      <div style={{ padding: "20px 24px" }}>
        {loading ? (
          <div style={{ padding: "40px 0", textAlign: "center", color: "var(--text-muted)" }}>
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ animation: "spin 1s linear infinite", display: "block", margin: "0 auto 10px" }}>
              <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
            </svg>
            <div style={{ fontSize: "0.82rem" }}>Loading…</div>
          </div>
        ) : (
          <div>
            <div className="breakdown-title">Gross Salary</div>
            <DeductionRow label="Basic Salary" bold amount={basicPay} editable={isAdmin} />
            {peraActive && <DeductionRow label="PERA" bold computed amount={pera} editable={isAdmin} onSave={isAdmin ? v => handleSave("pera", v) : undefined} />}
            {customEarnTypes.map(t => {
              const isFixed = t.computation_mode === 'fixed' && t.amount_fixed != null;
              const override = deductions[t.name];
              const val = override != null ? Number(override) : (isFixed ? Number(t.amount_fixed) : 0);
              return (
                <DeductionRow
                  key={t.name}
                  label={t.display_name || t.name}
                  bold
                  computed={isFixed && override == null}
                  amount={val}
                  editable={isAdmin}
                  onSave={isAdmin ? v => handleSave(t.name, v) : undefined}
                  onRemove={isAdmin && override != null ? () => handleRemove(t.name) : undefined}
                />
              );
            })}
            <DeductionRow label="Total" bold amount={grossPay} editable={isAdmin} />

            <div className="breakdown-title" style={{ marginTop: 18 }}>Deductions</div>
            {DED.map(g => (
              <DeductionGroup key={g.group} group={g} deductions={deductions} isAdmin={isAdmin} basicPay={basicPay} onSave={handleSave} onRemove={handleRemove} loanCols={loanCols} onAddLoan={handleAddLoanFromPicker} loanByCol={loanByCol} onLoanRefresh={async () => { await fetchLoans(); await fetchDeductions(); }} />
            ))}
            <DeductionRow label="Total Deductions" bold amount={dedColumnTotal} editable={false} />

            <div className="breakdown-title" style={{ marginTop: 18 }}>Less Deductions</div>
            {LESS_DED.map(g => (
              <DeductionGroup key={g.group} group={g} deductions={deductions} isAdmin={isAdmin} basicPay={basicPay} onSave={handleSave} onRemove={handleRemove} />
            ))}
            <DeductionRow label="Total Less Deductions" bold amount={lessDedTotal} editable={false} />
            <DeductionRow label="Total Salary/Earned for the Period" bold amount={totalSalaryEarned} editable={false} />
          </div>
        )}
        <div style={{ marginTop: 16, paddingTop: 14, borderTop: "1px solid var(--border)" }}>
          <div style={{ display: "flex", flexDirection: isMobile ? "column" : "row", gap: isMobile ? 0 : 24 }}>
            <div style={{ flex: 1, minWidth: 0 }}>
              <DeductionRow label="1st Quincena" computed amount={firstQuincena}
                editable={isAdmin}
                onSave={isAdmin ? v => handleQuincenaSave("first", v) : undefined}
                onRemove={isAdmin && quincenaOverridden ? handleQuincenaReset : undefined}
              />
            </div>
            <div style={{ flex: 1, minWidth: 0 }}>
              <DeductionRow label="2nd Quincena" computed amount={secondQuincena}
                editable={isAdmin}
                onSave={isAdmin ? v => handleQuincenaSave("second", v) : undefined}
                onRemove={isAdmin && quincenaOverridden ? handleQuincenaReset : undefined}
              />
            </div>
          </div>
        </div>

        {/* ── LOANS SECTION ── */}
        {!hideLoans && <div style={{ marginTop: 24, paddingTop: 18, borderTop: "2px solid var(--border)" }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
            <div className="breakdown-title" style={{ marginTop: 0 }}>Loans</div>
          </div>

          {/* Loans List */}
          {loansLoading ? (
            <div style={{ color: "var(--text-muted)", fontSize: "0.83rem" }}>Loading loans…</div>
          ) : loans.length === 0 ? (
            <div style={{ color: "var(--text-muted)", fontSize: "0.83rem" }}>No active loans.</div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {loans.map(loan => {
                const total = Number(loan.total_amount);
                const remaining = Number(loan.remaining_balance);
                const monthly = Number(loan.monthly_amount);
                const pct = total > 0 ? Math.max(0, Math.min(100, ((total - remaining) / total) * 100)) : 0;
                const monthsLeft = monthly > 0 ? Math.ceil(remaining / monthly) : null;
                const isActive = loan.status === "active";
                return (
                  <div key={loan.id} style={{ border: "1.5px solid var(--border)", borderRadius: 8, padding: "12px 14px", background: isActive ? "var(--surface)" : "var(--surface-2)", opacity: isActive ? 1 : 0.65 }}>
                    <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 12, marginBottom: 8 }}>
                      <div>
                        <span style={{ fontWeight: 700, fontSize: "0.9rem", color: "var(--text)" }}>{loan.loan_type_name}</span>
                        {loan.status !== "active" && (
                          <span style={{ marginLeft: 8, fontSize: "0.7rem", fontWeight: 700, padding: "2px 7px", borderRadius: 4, background: "#f3f4f6", color: "#6b7280", border: "1px solid #e5e7eb" }}>{loan.status}</span>
                        )}
                        {isActive && remaining < monthly && monthly > 0 && (
                          <span style={{ marginLeft: 8, fontSize: "0.7rem", fontWeight: 700, padding: "2px 7px", borderRadius: 4, background: "#fef9c3", color: "#854d0e", border: "1px solid #fde047" }}>⚠️ Will halt</span>
                        )}
                      </div>
                      {isAdmin && isActive && (
                        <button type="button" onClick={() => handleCloseLoan(loan)}
                          style={{ fontSize: "0.72rem", fontWeight: 700, padding: "3px 10px", borderRadius: 5, background: "#fee2e2", color: "#991b1b", border: "1px solid #fecaca", cursor: "pointer", flexShrink: 0 }}>
                          Close
                        </button>
                      )}
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 8, marginBottom: 10 }}>
                      <div>
                        <div style={{ fontSize: "0.68rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 2 }}>Total Loan</div>
                        <div style={{ fontSize: "0.88rem", fontWeight: 700, color: "var(--text)" }}>{fmtLoan(total)}</div>
                      </div>
                      <div>
                        <div style={{ fontSize: "0.68rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 2 }}>Monthly Deduction</div>
                        <div style={{ fontSize: "0.88rem", fontWeight: 700, color: "var(--text)" }}>{fmtLoan(monthly)}</div>
                      </div>
                      <div>
                        <div style={{ fontSize: "0.68rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 2 }}>Remaining</div>
                        <div style={{ fontSize: "0.88rem", fontWeight: 700, color: remaining > 0 ? "#dc2626" : "#16a34a" }}>{fmtLoan(remaining)}</div>
                      </div>
                    </div>
                    {/* Progress bar */}
                    <div style={{ height: 5, background: "var(--border)", borderRadius: 99, overflow: "hidden" }}>
                      <div style={{ height: "100%", width: `${pct}%`, background: pct >= 100 ? "#16a34a" : "var(--psa-navy)", borderRadius: 99, transition: "width 0.4s" }} />
                    </div>
                    {monthsLeft != null && isActive && (
                      <div style={{ fontSize: "0.72rem", color: "var(--text-muted)", marginTop: 5 }}>
                        ~{monthsLeft} month{monthsLeft !== 1 ? "s" : ""} remaining
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </div>}
      </div>
    </div>
  );
}

function LoanDeductionRow({ item, loan, deductions, isAdmin, onSave, onRefresh }) {
  const isMobile = useIsMobile();
  const [editingMonthly, setEditingMonthly] = useState(false);
  const [monthlyInput, setMonthlyInput]     = useState("");
  const [adding, setAdding]                 = useState(false);
  const [addAmount, setAddAmount]           = useState("");

  const monthlyAmt = Number(loan?.monthly_amount ?? deductions[item.col] ?? 0);
  const remaining  = loan ? Number(loan.remaining_balance) : 0;

  async function handleMonthlySave() {
    const newVal = Number(monthlyInput);
    if (!newVal || newVal <= 0) { setEditingMonthly(false); return; }
    if (newVal === monthlyAmt) { setEditingMonthly(false); return; }
    try {
      const res = await fetch(`${apiUrl}/employees/${loan.employee_id}/loans/${loan.id}`, {
        method: "PATCH", credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ monthlyAmount: newVal }),
      });
      if (!res.ok) { const e = await res.json().catch(() => ({})); alert(e.error || "Failed to update monthly amount."); }
      else if (onRefresh) await onRefresh();
    } catch { alert("Network error."); }
    setEditingMonthly(false);
  }

  async function handleAddSubmit() {
    const add = Number(addAmount);
    if (!add || add <= 0) return alert("Enter a valid amount to add.");
    try {
      const res = await fetch(`${apiUrl}/employees/${loan.employee_id}/loans/${loan.id}`, {
        method: "PATCH", credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ addAmount: add }),
      });
      if (!res.ok) { const e = await res.json().catch(() => ({})); return alert(e.error || `Error ${res.status}`); }
      setAdding(false);
      setAddAmount("");
      if (onRefresh) onRefresh();
    } catch { alert("Network error."); }
  }

  const btnStyle = { width: 30, height: 30, padding: 0, border: "none", background: "#1a4f8a", color: "#fff", cursor: "pointer", borderRadius: 6, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, boxShadow: "0 1px 3px rgba(26,79,138,0.25)" };
  const editSvg = (
    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
      <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
    </svg>
  );

  return (
    <div style={{ borderBottom: "1px solid var(--border)" }}>
      {isMobile ? (
        /* Mobile: stacked layout */
        <div style={{ padding: "9px 0 9px 14px", fontSize: "0.93rem" }}>
          {/* Row 1: label + remaining + add button */}
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, marginBottom: 5 }}>
            <span style={{ color: "var(--text-label)", fontWeight: 600 }}>{item.label}</span>
            {loan && (
              <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                <span style={{ fontSize: "0.81rem", color: "var(--text-muted)" }}>
                  Rem: <span style={{ fontWeight: 600, color: "var(--text)" }}>{fmt(remaining)}</span>
                </span>
                {isAdmin && !adding && (
                  <button type="button" onClick={() => { setAdding(true); setAddAmount(""); }} title="Add to loan"
                    style={{ ...btnStyle, width: 22, height: 22 }}>
                    <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                      <line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>
                    </svg>
                  </button>
                )}
              </div>
            )}
          </div>
          {/* Add form */}
          {isAdmin && loan && adding && (
            <div style={{ display: "flex", gap: 6, marginBottom: 5 }}>
              <input autoFocus type="text" placeholder="0.00" value={addAmount} onChange={e => setAddAmount(e.target.value)}
                onKeyDown={e => { if (e.key === "Enter") handleAddSubmit(); if (e.key === "Escape") { setAdding(false); setAddAmount(""); } }}
                style={{ flex: 1, textAlign: "right", padding: "4px 8px", fontSize: "0.88rem", border: "1.5px solid var(--psa-navy)", borderRadius: 4 }} />
              <button type="button" onClick={handleAddSubmit}
                style={{ padding: "4px 10px", height: 28, background: "#1a4f8a", color: "#fff", border: "none", borderRadius: 6, cursor: "pointer", fontSize: "0.82rem", flexShrink: 0 }}>Add</button>
              <button type="button" onClick={() => { setAdding(false); setAddAmount(""); }}
                style={{ padding: "4px 8px", height: 28, background: "transparent", color: "var(--text-muted)", border: "1px solid var(--border)", borderRadius: 6, cursor: "pointer", fontSize: "0.82rem", flexShrink: 0 }}>✕</button>
            </div>
          )}
          {/* Row 2: monthly */}
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
            <span style={{ fontSize: "0.8rem", color: "var(--text-muted)" }}>Monthly deduction</span>
            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
              {editingMonthly ? (
                <input autoFocus type="text" value={monthlyInput}
                  onChange={e => setMonthlyInput(e.target.value)}
                  onBlur={handleMonthlySave}
                  onKeyDown={e => { if (e.key === "Enter") e.target.blur(); if (e.key === "Escape") setEditingMonthly(false); }}
                  style={{ width: 100, textAlign: "right", padding: "4px 8px", fontSize: "0.88rem", border: "1.5px solid var(--psa-navy)", borderRadius: 4 }} />
              ) : (
                <span style={{ fontWeight: 600, color: "var(--text)", whiteSpace: "nowrap" }}>{fmt(monthlyAmt)}</span>
              )}
              {isAdmin && !editingMonthly && (
                <button type="button" onClick={() => { setMonthlyInput(monthlyAmt.toFixed(2)); setEditingMonthly(true); }} title="Edit monthly" style={{ ...btnStyle, width: 26, height: 26 }}>{editSvg}</button>
              )}
            </div>
          </div>
        </div>
      ) : (
        /* Desktop: original horizontal layout */
        <div style={{ display: "flex", alignItems: "center", padding: "9px 0", paddingLeft: 14, fontSize: "0.93rem" }}>
          <span style={{ color: "var(--text-label)", flexShrink: 0 }}>{item.label}</span>
          <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}>
            {loan && <span style={{ fontSize: "0.82rem", color: "var(--text-muted)", whiteSpace: "nowrap" }}>
              Remaining: <span style={{ fontWeight: 600, color: "var(--text)" }}>{fmt(remaining)}</span>
            </span>}
            {isAdmin && loan && !adding && (
              <button type="button" onClick={() => { setAdding(true); setAddAmount(""); }} title="Add to loan"
                style={{ ...btnStyle, width: 24, height: 24 }}>
                <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                  <line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>
                </svg>
              </button>
            )}
            {isAdmin && loan && adding && (
              <>
                <input autoFocus type="text" placeholder="0.00" value={addAmount} onChange={e => setAddAmount(e.target.value)}
                  onKeyDown={e => { if (e.key === "Enter") handleAddSubmit(); if (e.key === "Escape") { setAdding(false); setAddAmount(""); } }}
                  style={{ width: 110, textAlign: "right", padding: "4px 8px", fontSize: "0.88rem", border: "1.5px solid var(--psa-navy)", borderRadius: 4 }} />
                <button type="button" onClick={handleAddSubmit}
                  style={{ padding: "4px 12px", height: 28, background: "#1a4f8a", color: "#fff", border: "none", borderRadius: 6, cursor: "pointer", fontSize: "0.82rem", flexShrink: 0 }}>Add</button>
                <button type="button" onClick={() => { setAdding(false); setAddAmount(""); }}
                  style={{ padding: "4px 8px", height: 28, background: "transparent", color: "var(--text-muted)", border: "1px solid var(--border)", borderRadius: 6, cursor: "pointer", fontSize: "0.82rem", flexShrink: 0 }}>✕</button>
              </>
            )}
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
            <span style={{ fontSize: "0.82rem", color: "var(--text-muted)" }}>Monthly:</span>
            {editingMonthly ? (
              <input autoFocus type="text" value={monthlyInput}
                onChange={e => setMonthlyInput(e.target.value)}
                onBlur={handleMonthlySave}
                onKeyDown={e => { if (e.key === "Enter") e.target.blur(); if (e.key === "Escape") setEditingMonthly(false); }}
                style={{ width: 110, textAlign: "right", padding: "6px 9px", fontSize: "0.93rem" }} />
            ) : (
              <span
                onClick={isAdmin ? () => setEditingMonthly(true) : undefined}
                title={isAdmin ? "Click to edit" : undefined}
                style={{ fontWeight: 600, color: "var(--text)", whiteSpace: "nowrap", minWidth: 80, textAlign: "right", cursor: isAdmin ? "pointer" : "default", padding: "4px 8px", borderRadius: 4 }}
                onMouseEnter={isAdmin ? e => e.currentTarget.style.background = "var(--hover, rgba(0,56,168,0.07))" : undefined}
                onMouseLeave={isAdmin ? e => e.currentTarget.style.background = "transparent" : undefined}
              >{fmt(monthlyAmt)}</span>
            )}
            {isAdmin && (
              <div style={{ width: 30, flexShrink: 0, display: "flex", justifyContent: "center" }}>
                {!editingMonthly && <button type="button" onClick={() => { setMonthlyInput(monthlyAmt.toFixed(2)); setEditingMonthly(true); }} title="Edit monthly" style={btnStyle}>{editSvg}</button>}
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

function DeductionGroup({ group, deductions, isAdmin, basicPay, onSave, onRemove, loanCols, onAddLoan, loanByCol, onLoanRefresh }) {
  const hasItems = group.items.length > 0;

  if (!hasItems) {
    if (group.computed) {
      const displayAmount = Number(deductions[group.computed.col] ?? roundAmt(basicPay * group.computed.rate));
      return (
        <DeductionRow label={group.group} bold computed
          amount={displayAmount} editable={isAdmin}
          onSave={isAdmin && onSave ? v => onSave(group.computed.col, v) : undefined}
        />
      );
    }
    {
      const loanForCol = (loanByCol && group.col && loanByCol[group.col]) || null;
      const isLoanCol = (loanCols && group.col && loanCols.has(group.col)) || !!loanForCol;
      if (isLoanCol) {
        if (!loanForCol) return null;
        return <LoanDeductionRow item={{ label: group.group, col: group.col }} loan={loanForCol} deductions={deductions} isAdmin={isAdmin} onSave={onSave} onRefresh={onLoanRefresh} />;
      }
    }
    const amount = Number(deductions[group.col] ?? 0);
    const isLoanType = loanCols && group.col && loanCols.has(group.col);
    if (amount === 0 && (!isAdmin || isLoanType)) return null;
    return (
      <DeductionRow label={group.group} bold
        amount={amount} editable={isAdmin}
        onSave={onSave ? v => onSave(group.col, v) : undefined}
      />
    );
  }

  const computedAmount = group.computed ? roundAmt(basicPay * group.computed.rate) : 0;

  const useDropdownPicker = isAdmin && group.items.length > 0 && !!onSave;
  const visibleItems = useDropdownPicker
    ? group.items.filter(item => {
        const isLoanItem = loanCols && loanCols.has(item.col);
        if (isLoanItem) return !!(loanByCol && loanByCol[item.col]);
        return deductions[item.col] != null;
      })
    : (isAdmin ? group.items : group.items.filter(item => Number(deductions[item.col] ?? 0) !== 0 || (loanCols && loanCols.has(item.col) && loanByCol && loanByCol[item.col])));

  const showComputed = group.computed && (isAdmin || computedAmount !== 0);
  if (!isAdmin && visibleItems.length === 0 && !showComputed) return null;

  return (
    <div style={{ marginBottom: 6 }}>
      <div style={{ fontWeight: 700, fontSize: "0.85rem", color: "var(--text)", marginTop: 10, marginBottom: 2 }}>{group.group}</div>
      {useDropdownPicker && (
        <AddDeductionRow group={group} addedItems={visibleItems} onAdd={(item, amount) => onSave(item.col, amount)} loanCols={loanCols} onAddLoan={onAddLoan} />
      )}
      {showComputed && (
        <DeductionRow label={group.computed.label} indent computed
          amount={Number(deductions[group.computed.col] ?? computedAmount)} editable={isAdmin}
          onSave={isAdmin && onSave ? v => onSave(group.computed.col, v) : undefined}
        />
      )}
      {visibleItems.map(item => {
        const loan = loanByCol && loanByCol[item.col];
        const isLoanItem = loanCols && loanCols.has(item.col);
        if (isLoanItem && loan && loan.status === "active") {
          return <LoanDeductionRow key={item.col} item={item} loan={loan} deductions={deductions} isAdmin={isAdmin} onSave={onSave} onRefresh={onLoanRefresh} />;
        }
        return (
          <DeductionRow key={item.col} label={item.label} indent
            amount={Number(deductions[item.col] ?? 0)} editable={isAdmin}
            onSave={onSave ? v => onSave(item.col, v) : undefined}
            onRemove={useDropdownPicker && onRemove && !isLoanItem ? () => onRemove(item.col) : undefined}
          />
        );
      })}
    </div>
  );
}

function AddDeductionRow({ group, addedItems, onAdd, loanCols, onAddLoan }) {
  const [selectedItem, setSelectedItem] = useState(null);
  const [amount, setAmount]             = useState("");
  const [totalLoan, setTotalLoan]       = useState("");
  const [monthlyDed, setMonthlyDed]     = useState("");
  const remaining = group.items.filter(i => !addedItems.some(a => a.col === i.col));
  const isLoan = selectedItem && loanCols && loanCols.has(selectedItem.col);

  if (remaining.length === 0) return null;

  function handleAdd() {
    if (!selectedItem) return;
    if (isLoan) {
      if (!totalLoan || !monthlyDed) return;
      if (onAddLoan) onAddLoan(selectedItem, Number(totalLoan), Number(monthlyDed));
      setSelectedItem(null); setTotalLoan(""); setMonthlyDed("");
    } else {
      onAdd(selectedItem, Number(amount) || 0);
      setSelectedItem(null); setAmount("");
    }
  }

  const inSt = { padding: "8px 10px", fontSize: "0.88rem", border: "1.5px solid var(--border)", borderRadius: 6, background: "var(--surface)", color: "var(--text)", flex: "1 1 0", minWidth: 0, boxSizing: "border-box" };

  return (
    <div style={{ marginTop: 10, paddingLeft: 14 }}>
      {/* Row 1: dropdown full-width; Row 2: amount + Add (stacked on mobile) */}
      <div className="add-deduction-row" style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        <CustomSelect
          value={selectedItem?.col ?? ""}
          onChange={v => { setSelectedItem(group.items.find(i => i.col === v)); setAmount(""); setTotalLoan(""); setMonthlyDed(""); }}
          placeholder="+ Select item to add…"
          options={remaining.map(i => ({ value: i.col, label: i.label }))}
        />
        {!isLoan && (
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            <input type="text" value={amount} onChange={e => setAmount(e.target.value)}
              placeholder="0.00" className="add-deduction-amount"
              style={{ flex: 1, padding: "8px 10px", fontSize: "0.92rem" }} />
            <button type="button" onClick={handleAdd} disabled={!selectedItem}
              style={{ padding: "8px 18px", fontSize: "0.88rem", whiteSpace: "nowrap", flexShrink: 0 }}>
              Add
            </button>
          </div>
        )}
      </div>
      {/* Row 2: loan inputs + Add (only when loan selected) */}
      {isLoan && (
        <div style={{ marginTop: 6 }}>
          <div style={{ display: "flex", gap: 8, alignItems: "flex-end", flexWrap: "wrap" }}>
            <div style={{ flex: 1, minWidth: 120 }}>
              <div style={{ fontSize: "0.68rem", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.06em", color: "var(--text-muted)", marginBottom: 4 }}>Total Loan Amount</div>
              <input type="text" placeholder="0.00" value={totalLoan} onChange={e => setTotalLoan(e.target.value)} style={inSt} />
            </div>
            <div style={{ flex: 1, minWidth: 120 }}>
              <div style={{ fontSize: "0.68rem", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.06em", color: "var(--text-muted)", marginBottom: 4 }}>Monthly Deduction</div>
              <input type="text" placeholder="0.00" value={monthlyDed} onChange={e => setMonthlyDed(e.target.value)} style={inSt} />
            </div>
            <button type="button" onClick={handleAdd} disabled={!totalLoan || !monthlyDed}
              style={{ padding: "8px 18px", fontSize: "0.88rem", whiteSpace: "nowrap", flexShrink: 0, alignSelf: "flex-end" }}>
              Add
            </button>
          </div>
          {totalLoan && monthlyDed && Number(monthlyDed) > 0 && (
            <div style={{ fontSize: "0.72rem", color: "var(--text-muted)", marginTop: 4 }}>
              ~{Math.ceil(Number(totalLoan) / Number(monthlyDed))} months to pay off
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function DeductionRow({ label, amount, editable, indent, bold, onSave, onRemove, computed }) {
  const [editing, setEditing] = useState(false);
  const canEdit = editable && !!onSave;
  const showInput = canEdit && editing;

  return (
    <div style={{
      display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8,
      padding: "9px 0", paddingLeft: indent ? 14 : 0,
      fontSize: "0.93rem", borderBottom: "1px solid var(--border)",
      fontWeight: bold ? 700 : 400,
    }}>
      <span style={{ color: bold ? "var(--text)" : "var(--text-label)" }}>{label}</span>
      <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
        {showInput ? (
          <input
            type="number" min="0" step="0.01"
            defaultValue={amount ? Number(amount).toFixed(2) : ""}
            placeholder="0.00"
            ref={el => { if (el) { el.focus(); el.select(); } }}
            onClick={e => e.target.focus()}
            onBlur={e => {
              if (Number(e.target.value || 0) !== Number(amount || 0)) {
                if (window.confirm("Do you want to proceed with the changes?")) {
                  onSave(e.target.value || 0);
                }
              }
              setEditing(false);
            }}
            onKeyDown={e => {
              if (e.key === "Enter") e.target.blur();
              if (e.key === "Escape") setEditing(false);
            }}
            style={{ width: 110, textAlign: "right", padding: "6px 9px", fontSize: "0.93rem" }}
          />
        ) : (
          <span
            onClick={canEdit ? () => setEditing(true) : undefined}
            title={canEdit ? "Click to edit" : undefined}
            style={{ fontWeight: 600, color: "var(--text)", whiteSpace: "nowrap", minWidth: 80, textAlign: "right", cursor: canEdit ? "pointer" : "default", padding: "4px 8px", borderRadius: 4, background: canEdit ? "transparent" : undefined }}
            onMouseEnter={canEdit ? e => e.currentTarget.style.background = "var(--hover, rgba(0,56,168,0.07))" : undefined}
            onMouseLeave={canEdit ? e => e.currentTarget.style.background = "transparent" : undefined}
          >
            {Number(amount) > 0 ? (bold ? `₱ ${fmt(amount)}` : fmt(amount)) : (editable ? "—" : "")}
          </span>
        )}
        {/* Button slot — only reserve space when editable so read-only amounts stay flush right */}
        {editable && <div style={{ width: 30, flexShrink: 0, display: "flex", justifyContent: "center" }}>
          {canEdit && !showInput && !onRemove && (
            <button
              type="button" onClick={() => setEditing(true)} title="Edit"
              style={{ width: 30, height: 30, padding: 0, display: "flex", alignItems: "center", justifyContent: "center", borderRadius: 6, background: "#1a4f8a", color: "#ffffff", border: "none", cursor: "pointer", flexShrink: 0, boxShadow: "0 1px 3px rgba(26,79,138,0.25)" }}
            >
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
                <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
              </svg>
            </button>
          )}
          {onRemove && (
            <button
              type="button" onClick={onRemove} title="Remove"
              style={{ width: 30, height: 30, padding: 0, display: "flex", alignItems: "center", justifyContent: "center", borderRadius: 6, background: "#fee2e2", color: "#991b1b", border: "1px solid #fecaca", cursor: "pointer", flexShrink: 0 }}
            >
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
              </svg>
            </button>
          )}
        </div>}
      </div>
    </div>
  );
}


function currentPayPeriod() {
  const now = new Date();
  const y = now.getFullYear(), m = now.getMonth();
  const d = now.getDate();
  const lastDay = new Date(y, m + 1, 0).getDate();
  return d <= 15
    ? { period_start: new Date(y, m, 1).toISOString(), period_end: new Date(y, m, 15).toISOString() }
    : { period_start: new Date(y, m, 16).toISOString(), period_end: new Date(y, m, lastDay).toISOString() };
}

// Returns true if a catalog group has at least one active entry in the employee's ledger.
function hasLedgerData(g, ded) {
  if (g.computed && ded[g.computed.col] != null) return true;
  if (!g.computed && !g.items.length && g.col && ded[g.col] != null) return true;
  if (g.items.some(i => ded[i.col] != null)) return true;
  return false;
}

const PaySlipCard = memo(function PaySlipCard({ item, period, isOpen, onToggle, hideHeader = false }) {
  const { dedTypes, earnTypes, signatories } = useContext(TypesContext);
  const articleRef = useRef(null);
  const LESS_DED_ALL = dedTypes.length > 0 ? buildCatalogFromTypes(dedTypes, 'less_deduction') : LESS_DEDUCTIONS_CATALOG;
  const DED_ALL      = dedTypes.length > 0 ? buildCatalogFromTypes(dedTypes, 'deduction')      : DEDUCTIONS_CATALOG;
  const peraActive = earnTypes.length === 0 || earnTypes.some(t => t.name === 'pera');

  const customEarnTypes = earnTypes.filter(t => t.name !== 'basic_pay' && t.name !== 'pera');

  const ded      = item.deductions ?? {};
  const basicPay = Number(item.basic_pay) || 0;
  const pera     = Number(item.allowances?.pera ?? PERA_AMOUNT);

  // Only render groups that have actual ledger entries for this employee
  const LESS_DED = LESS_DED_ALL.filter(g => hasLedgerData(g, ded));
  const DED      = DED_ALL.filter(g => hasLedgerData(g, ded));

  const lessDedTotal          = sumCatalogComputed(LESS_DED, ded, basicPay) + sumCatalogTyped(LESS_DED, ded);
  const deductionsColumnTotal = sumCatalogComputed(DED, ded, basicPay) + sumCatalogTyped(DED, ded);
  const totalSalaryEarned     = Number(item.gross_pay) - lessDedTotal;

  function handlePrint() {
    const el = articleRef.current;
    if (!el) return;
    // Clone into a dedicated print portal appended directly to <body>, so the
    // print-only CSS can hide every other element on the page unconditionally —
    // immune to whatever layout/wrapper this card happens to be embedded in.
    const portal = document.createElement('div');
    portal.className = 'payslip-print-portal';
    const clone = el.cloneNode(true);
    clone.classList.add('payslip-to-print');
    portal.appendChild(clone);
    document.body.appendChild(portal);
    document.body.classList.add('payslip-single-print');
    // Inject @page { margin: 0 } to suppress Chrome's URL/datetime headers
    const pageStyle = document.createElement('style');
    pageStyle.id = '__ps_page_override';
    pageStyle.textContent = '@page { margin: 0 !important; size: A4 portrait; }';
    document.head.appendChild(pageStyle);
    window.print();
    document.head.removeChild(pageStyle);
    document.body.classList.remove('payslip-single-print');
    document.body.removeChild(portal);
  }

  return (
    <article ref={articleRef} className="job-card payslip-card">
      {!hideHeader && (
        <div className="job-card__header" onClick={onToggle} style={{ cursor: "pointer" }}>
          <div className="job-card__info">
            <div className="job-card__name">{item.employee_name}</div>
            <div className="job-card__tags">
              <span className="job-card__tag">{item.position}</span>
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 20, flexShrink: 0 }}>
            <div style={{ textAlign: "right" }}>
              <div style={{ fontSize: "0.68rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.06em" }}>Take Home Pay</div>
              <div style={{ fontSize: "1.05rem", fontWeight: 800, color: "#166534" }}>₱ {fmt(item.net_pay)}</div>
            </div>
            <svg width="14" height="14" viewBox="0 0 16 16" style={{ flexShrink: 0, transition: "transform 0.2s", transform: isOpen ? "rotate(180deg)" : "rotate(0deg)" }}>
              <path d="M4 6l4 4 4-4" stroke="#637491" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" fill="none"/>
            </svg>
          </div>
        </div>
      )}

      {isOpen && (
        <div className="payslip-body" style={{ borderTop: "1px solid var(--border)", background: "var(--surface-2)", padding: "20px 24px", borderRadius: "0 0 var(--radius) var(--radius)" }}>
          {/* Print button — screen only */}
          <div className="screen-only" style={{ display: "flex", justifyContent: "flex-end", marginBottom: 14 }}>
            <button
              type="button"
              onClick={handlePrint}
              title="Print this pay slip"
              style={{
                display: "flex", alignItems: "center", gap: 6,
                padding: "6px 14px", borderRadius: 7, border: "1.5px solid var(--border-2)",
                background: "var(--surface)", color: "var(--text)", fontSize: "0.84rem",
                fontWeight: 600, cursor: "pointer",
              }}
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/>
                <rect x="6" y="14" width="12" height="8"/>
              </svg>
              Print
            </button>
          </div>

          <div className="payslip-doc-header">
            <img src={psaLogo} alt="PSA" className="payslip-logo" />
            <div className="payslip-agency">
              <div>Republic of the Philippines</div>
              <div style={{ fontWeight: 800 }}>Philippine Statistics Authority</div>
              <div>Regional Statistical Services Office II</div>
            </div>
            <div className="payslip-title">PAY&nbsp;&nbsp;SLIP</div>
          </div>
          <div className="payslip-meta">
            <span>Name: <b>{item.employee_name?.toUpperCase()}</b></span>
            <span>{formatPeriodRange(period)}</span>
          </div>

          <div className="deduction-two-col" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gap: 32 }}>
            <div>
              <div className="breakdown-title">Gross Earnings</div>
              <BreakdownRow label="Basic Salary" value={basicPay} />
              {peraActive && <BreakdownRow label="PERA" value={pera} />}
              {customEarnTypes.map(t => {
                const val = item.allowances?.[t.name] != null
                  ? Number(item.allowances[t.name])
                  : (t.computation_mode === 'fixed' && t.amount_fixed != null ? Number(t.amount_fixed) : 0);
                if (val === 0) return null;
                return <BreakdownRow key={t.name} label={t.display_name || t.name} value={val} />;
              })}
              <BreakdownRow label="Total Gross" value={item.gross_pay} bold />

              {LESS_DED.length > 0 && (
                <>
                  <div className="breakdown-title" style={{ marginTop: 18 }}>Less Deductions</div>
                  {LESS_DED.map(g => (
                    <DeductionGroup key={g.group} group={g} deductions={ded} isAdmin={false} basicPay={basicPay} />
                  ))}
                </>
              )}
              <BreakdownRow label="Total Salary/Earned for the Period" value={totalSalaryEarned} bold />
            </div>
            <div>
              <div className="breakdown-title">Deductions</div>
              {DED.map(g => (
                <DeductionGroup key={g.group} group={g} deductions={ded} isAdmin={false} basicPay={basicPay} />
              ))}
              <BreakdownRow label="Total Deductions" value={deductionsColumnTotal} bold />
            </div>
          </div>

          {/* Take Home Pay + Quincenas — always at the bottom, shown on screen and print */}
          <div style={{ marginTop: 16, paddingTop: 14, borderTop: "2px solid var(--border)", display: "flex", justifyContent: "space-between", alignItems: "center", gap: 24, flexWrap: "wrap" }}>
            <div style={{ display: "flex", alignItems: "baseline", gap: 10 }}>
              <span style={{ fontWeight: 700, color: "var(--text-label)" }}>Take Home Pay</span>
              <span style={{ fontSize: "1.2rem", fontWeight: 800, color: "#166534" }}>₱ {fmt(item.net_pay)}</span>
            </div>
            <div style={{ display: "flex", gap: 24, flex: 1, maxWidth: 480 }}>
              <div style={{ flex: 1 }}>
                <BreakdownRow label="1st Quincena" value={item.first_quincena} />
              </div>
              <div style={{ flex: 1 }}>
                <BreakdownRow label="2nd Quincena" value={item.second_quincena} />
              </div>
            </div>
          </div>

          <div className="payslip-signatories">
            <div className="payslip-sig">
              <div>Prepared by:</div>
              <div className="payslip-sig-line" />
              {signatories?.prepared_by_name  && <div className="payslip-sig-name">{signatories.prepared_by_name}</div>}
              {signatories?.prepared_by_title && <div className="payslip-sig-title">{signatories.prepared_by_title}</div>}
            </div>
            <div className="payslip-sig">
              <div>Certified by:</div>
              <div className="payslip-sig-line" />
              {signatories?.certified_by_name  && <div className="payslip-sig-name">{signatories.certified_by_name}</div>}
              {signatories?.certified_by_title && <div className="payslip-sig-title">{signatories.certified_by_title}</div>}
            </div>
          </div>
        </div>
      )}
    </article>
  );
});

function BreakdownRow({ label, value, bold }) {
  return (
    <div style={{
      display: "flex", justifyContent: "space-between", gap: 12,
      padding: "7px 0", borderBottom: bold ? "2px solid var(--border)" : "1px solid var(--border)",
      fontWeight: bold ? 700 : 400, fontSize: bold ? "0.92rem" : "0.88rem",
      color: bold ? "var(--text)" : "var(--text-muted)",
    }}>
      <span>{label}</span>
      <span style={{ whiteSpace: "nowrap", flexShrink: 0 }}>{bold ? "₱ " : ""}{fmt(value)}</span>
    </div>
  );
}

// ── adminBtn ──────────────────────────────────────────────────────────────────
function adminBtn(color) {
  const map = {
    green: { bg: "#166534", text: "#ffffff", border: "#166534" },
    red:   { bg: "#c0392b", text: "#ffffff", border: "#c0392b" },
  };
  const c = map[color];
  return {
    padding: "6px 16px", minWidth: 72, height: 32,
    background: c.bg, color: c.text,
    border: `1.5px solid ${c.border}`,
    borderRadius: 6,
    fontWeight: 700, fontSize: "0.82rem",
    cursor: "pointer", transition: "opacity 0.12s",
    display: "inline-flex", alignItems: "center", justifyContent: "center",
  };
}

// ── LookupTablePanel ─────────────────────────────────────────────────────────
function LookupTablePanel({ title, endpoint, showItemNo = false, showSalaryGrade = false, onRefresh, searchable = false, defaultLimit = null, canManage = true, externalQuery = null }) {
  const [items, setItems]               = useState([]);
  const [loading, setLoading]           = useState(true);
  const [newName, setNewName]           = useState("");
  const [newItemNo, setNewItemNo]       = useState("");
  const [newSalaryGrade, setNewSalaryGrade] = useState("");
  const [adding, setAdding]             = useState(false);
  const [showAddModal, setShowAddModal] = useState(false);
  const [editId, setEditId]             = useState(null);
  const [editName, setEditName]         = useState("");
  const [editItemNo, setEditItemNo]     = useState("");
  const [editSalaryGrade, setEditSalaryGrade] = useState("");
  const [q, setQ]                       = useState("");
  const [showAll, setShowAll]           = useState(false);

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const res = await fetch(`${apiUrl}/${endpoint}`, { credentials: "include" });
      if (res.ok) setItems(await res.json());
    } finally { setLoading(false); }
  }, [endpoint]);

  useEffect(() => { load(); }, [load]);

  async function handleAdd(e) {
    e.preventDefault();
    if (!newName.trim()) return;
    setAdding(true);
    try {
      const body = { name: newName };
      if (showItemNo) body.itemNo = newItemNo;
      if (showSalaryGrade) body.salaryGrade = newSalaryGrade ? parseInt(newSalaryGrade, 10) : null;
      const res = await fetch(`${apiUrl}/${endpoint}`, {
        method: "POST", credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });
      if (res.ok) {
        setNewName(""); setNewItemNo(""); setNewSalaryGrade("");
        setShowAddModal(false);
        await load();
        if (onRefresh) onRefresh();
      } else {
        const d = await res.json();
        alert(d.error ?? "Failed to add.");
      }
    } finally { setAdding(false); }
  }

  function startEdit(item) {
    setEditId(item.id);
    setEditName(item.name);
    setEditItemNo(item.item_no ?? "");
    setEditSalaryGrade(item.salary_grade ? String(item.salary_grade) : "");
  }

  async function handleSaveEdit(id) {
    if (!editName.trim()) return;
    const body = showItemNo ? { name: editName, itemNo: editItemNo } : { name: editName };
    if (showSalaryGrade) body.salaryGrade = editSalaryGrade ? parseInt(editSalaryGrade, 10) : null;
    const res = await fetch(`${apiUrl}/${endpoint}/${id}`, {
      method: "PATCH", credentials: "include",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (res.ok) {
      setEditId(null);
      await load();
      if (onRefresh) onRefresh();
    } else {
      const d = await res.json();
      alert(d.error ?? "Failed to save.");
    }
  }

  async function handleDelete(item) {
    if (!window.confirm(`Delete "${item.name}"?`)) return;
    const res = await fetch(`${apiUrl}/${endpoint}/${item.id}`, { method: "DELETE", credentials: "include" });
    if (res.ok) {
      await load();
      if (onRefresh) onRefresh();
    } else {
      const d = await res.json();
      alert(d.error ?? "Cannot delete — it may be in use by employees.");
    }
  }

  const activeQ = externalQuery !== null ? externalQuery : q;
  const filtered = activeQ.trim()
    ? items.filter(i => i.name.toLowerCase().includes(activeQ.trim().toLowerCase()))
    : items;
  const limit = defaultLimit && !showAll && !activeQ.trim() ? defaultLimit : filtered.length;
  const visible = filtered.slice(0, limit);

  const singular = title.replace(/s$/, "");
  const showSearch = externalQuery === null && (searchable || items.length > 6);
  const editKeys = (id) => (e) => { if (e.key === "Enter") handleSaveEdit(id); if (e.key === "Escape") setEditId(null); };

  return (
    <div>
      <div className="ws-toolbar">
        <div className="ws-title">{title}<span className="ws-count">{items.length}</span></div>
        {showSearch
          ? <WsSearch value={q} onChange={v => { setQ(v); setShowAll(false); }} placeholder={`Search ${title.toLowerCase()}…`} />
          : <span style={{ flex: 1 }} />}
        {canManage && (
          <button type="button" className="btn-primary ws-add" onClick={() => { setNewName(""); setNewSalaryGrade(""); setShowAddModal(true); }}>
            + Add {singular.toLowerCase()}
          </button>
        )}
      </div>

      <div className="ws-pad">
        {loading ? (
          <WsSpinner />
        ) : items.length === 0 ? (
          <div className="ws-hint">No {title.toLowerCase()} yet.{canManage ? ` Use “Add ${singular.toLowerCase()}” to create one.` : ""}</div>
        ) : (
          <>
            <div className="lk-list">
              {visible.map(item => (
                <div key={item.id} className={`lk-row${editId === item.id ? " is-editing" : ""}`}>
                  <span className="lk-icon" aria-hidden="true">
                    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="4" y="2" width="16" height="20" rx="2"/><line x1="9" y1="6" x2="9.01" y2="6"/><line x1="15" y1="6" x2="15.01" y2="6"/><line x1="9" y1="10" x2="9.01" y2="10"/><line x1="15" y1="10" x2="15.01" y2="10"/><line x1="9" y1="14" x2="9.01" y2="14"/><line x1="15" y1="14" x2="15.01" y2="14"/><path d="M10 22v-4h4v4"/></svg>
                  </span>
                  {editId === item.id ? (
                    <>
                      <input value={editName} onChange={e => setEditName(e.target.value)} autoFocus onKeyDown={editKeys(item.id)} aria-label={`${singular} name`} />
                      {showSalaryGrade && (
                        <input type="number" min="1" max="33" value={editSalaryGrade} onChange={e => setEditSalaryGrade(e.target.value)} placeholder="SG" onKeyDown={editKeys(item.id)} style={{ flex: "0 0 80px" }} aria-label="Salary grade" />
                      )}
                      <span className="lk-actions">
                        <button type="button" className="btn-primary ws-btn-sm" onClick={() => handleSaveEdit(item.id)} disabled={!editName.trim()}>Save</button>
                        <button type="button" className="um-btn um-btn--neutral um-btn--sm" onClick={() => setEditId(null)}>Cancel</button>
                      </span>
                    </>
                  ) : (
                    <>
                      <span className="lk-name">{item.name}</span>
                      {showSalaryGrade && item.salary_grade && <span className="ws-pill">SG {item.salary_grade}</span>}
                      {canManage && (
                        <span className="lk-actions">
                          <button type="button" className="um-btn um-btn--neutral um-btn--sm" onClick={() => startEdit(item)}>Edit</button>
                          <button type="button" className="ws-icon-btn ws-icon-btn--danger ws-icon-btn--sm" onClick={() => handleDelete(item)} title={`Delete ${singular.toLowerCase()}`} aria-label={`Delete ${item.name}`}>
                            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/></svg>
                          </button>
                        </span>
                      )}
                    </>
                  )}
                </div>
              ))}
            </div>
            {activeQ.trim() && filtered.length === 0 && (
              <div className="ws-empty">No {title.toLowerCase()} match “{activeQ.trim()}”.</div>
            )}
            {!activeQ.trim() && defaultLimit && filtered.length > defaultLimit && (
              <button type="button" className="um-btn um-btn--neutral pt-more" onClick={() => setShowAll(v => !v)}>
                {showAll ? "Show less" : `Show all ${filtered.length}`}
              </button>
            )}
            {endpoint === "divisions" && (
              <div className="ws-hint" style={{ marginTop: 14 }}>
                <span>An employee's division appears as the office on their leave application, e.g. “RSSO II - {items[0]?.name ?? "CRS"}”, and decides which chief receives it (set per chief in Settings → Leave Signatories).</span>
              </div>
            )}
          </>
        )}
      </div>

      {/* Add Modal */}
      {showAddModal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(13,35,70,0.45)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 16, backdropFilter: "blur(2px)" }}
          onClick={e => { if (e.target === e.currentTarget) setShowAddModal(false); }}
        >
          <div style={{ background: "var(--surface)", borderRadius: 16, width: "100%", maxWidth: 420, boxShadow: "0 20px 60px rgba(0,0,0,0.25)" }}>
            {/* Header */}
            <div style={{ padding: "20px 24px 18px", borderBottom: "1px solid var(--border)", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
              <div>
                <div style={{ fontSize: "1.05rem", fontWeight: 800, color: "var(--psa-navy)", letterSpacing: "-0.01em" }}>Add {title.replace(/s$/, "")}</div>
                <div style={{ fontSize: "0.78rem", color: "var(--text-muted)", marginTop: 2 }}>Fill in the details below</div>
              </div>
              <button type="button" onClick={() => setShowAddModal(false)} style={{ padding: "4px 8px", background: "transparent", border: "none", color: "var(--text-muted)", cursor: "pointer", fontSize: "1.2rem", lineHeight: 1, boxShadow: "none" }}
                onMouseEnter={e => { e.currentTarget.style.color = "#991b1b"; }}
                onMouseLeave={e => { e.currentTarget.style.color = "var(--text-muted)"; }}
              >✕</button>
            </div>
            {/* Form */}
            <form onSubmit={handleAdd} style={{ padding: "22px 24px", display: "flex", flexDirection: "column", gap: 14 }}>
              <label style={{ display: "flex", flexDirection: "column", gap: 6, fontSize: "0.88rem", fontWeight: 600, color: "var(--psa-navy)" }}>
                {title.replace(/s$/, "")} Name
                <input
                  className="form-input"
                  value={newName}
                  onChange={e => setNewName(e.target.value)}
                  placeholder={`e.g. ${title === "Positions" ? "Administrative Officer" : title === "Divisions" ? "Regional Operations" : "Contractual"}`}
                  autoFocus
                  required
                />
              </label>
              {showSalaryGrade && (
                <label style={{ display: "flex", flexDirection: "column", gap: 6, fontSize: "0.88rem", fontWeight: 600, color: "var(--psa-navy)" }}>
                  Salary Grade
                  <input
                    className="form-input"
                    type="number" min="1" max="33"
                    value={newSalaryGrade}
                    onChange={e => setNewSalaryGrade(e.target.value)}
                    placeholder="1–33"
                  />
                </label>
              )}
              <div style={{ borderTop: "1px solid var(--border)", paddingTop: 16, display: "flex", justifyContent: "flex-end", gap: 10 }}>
                <button type="button" onClick={() => setShowAddModal(false)} style={{ padding: "10px 20px", fontSize: "0.9rem", background: "var(--surface)", border: "1.5px solid var(--border-2)", borderRadius: "var(--radius)", cursor: "pointer", fontWeight: 600, color: "var(--text)" }}>
                  Cancel
                </button>
                <button type="submit" disabled={adding || !newName.trim()} style={{ padding: "10px 20px", fontSize: "0.9rem", fontWeight: 700, borderRadius: "var(--radius)" }}>
                  {adding ? "Saving…" : "Save"}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}

// ── EmployeeSalaryStandingView ────────────────────────────────────────────────
function EmployeeSalaryStandingView() {
  const [q, setQ]               = useState("");
  const [employees, setEmployees] = useState([]);
  const [loading, setLoading]   = useState(true);
  const [editId, setEditId]     = useState(null);
  const [editStep, setEditStep] = useState(1);
  const [saving, setSaving]     = useState(false);

  const load = useCallback(async (search = "") => {
    setLoading(true);
    try {
      const res = await fetch(`${apiUrl}/employees/salary-standing?q=${encodeURIComponent(search)}`, { credentials: "include" });
      if (res.ok) setEmployees(await res.json());
    } finally { setLoading(false); }
  }, []);

  useEffect(() => { load(""); }, [load]);

  function handleSearch(e) { e.preventDefault(); load(q); }

  function startEdit(emp) { setEditId(emp.id); setEditStep(emp.step ?? 1); }

  async function saveStep(id) {
    setSaving(true);
    try {
      const res = await fetch(`${apiUrl}/employees/${id}/step`, {
        method: "PATCH", credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ step: editStep }),
      });
      if (res.ok) { setEditId(null); load(q); }
      else { const d = await res.json(); alert(d.error ?? "Failed to update step."); }
    } finally { setSaving(false); }
  }

  const thStyle = { padding: "11px 14px", textAlign: "left", fontWeight: 700, fontSize: "0.72rem", color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.06em", whiteSpace: "nowrap" };
  const tdStyle = { padding: "12px 14px", verticalAlign: "middle" };

  return (
    <>
      <div className="page-header">
        <div className="page-header-title">Salary Standing</div>
        <div className="page-header-sub">Employee salary grades and step increments — updating a step auto-recalculates Basic Monthly Pay</div>
      </div>

      <div style={{ maxWidth: 940, margin: "0 auto" }}>
        <form onSubmit={handleSearch} style={{ display: "flex", gap: 8, marginBottom: 16 }}>
          <div style={{ position: "relative", flex: 1 }}>
            <svg style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)", color: "var(--text-muted)", pointerEvents: "none" }} width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
            </svg>
            <input
              value={q}
              onChange={e => setQ(e.target.value)}
              placeholder="Search by name or employee number…"
              style={{ width: "100%", boxSizing: "border-box", padding: "10px 12px 10px 34px", border: "1.5px solid var(--border-2)", borderRadius: "var(--radius)", fontSize: "0.9rem", background: "var(--surface)" }}
            />
          </div>
          <button type="submit" style={{ padding: "10px 20px", fontSize: "0.9rem", whiteSpace: "nowrap" }}>Search</button>
        </form>

        {loading ? (
          <div style={{ padding: "40px 0", textAlign: "center", color: "var(--text-muted)" }}>
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ animation: "spin 1s linear infinite", display: "block", margin: "0 auto 10px" }}>
              <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
            </svg>
            <div style={{ fontSize: "0.82rem" }}>Loading…</div>
          </div>
        ) : employees.length === 0 ? (
          <div style={{ color: "var(--text-muted)", fontSize: "0.88rem", padding: "20px 0", textAlign: "center" }}>No employees found.</div>
        ) : (
          <div className="panel" style={{ padding: 0, overflow: "hidden" }}>
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.88rem" }}>
                <thead>
                  <tr style={{ background: "var(--surface-2)", borderBottom: "2px solid var(--border)" }}>
                    <th style={thStyle}>Emp. No.</th>
                    <th style={thStyle}>Name</th>
                    <th style={thStyle}>Position</th>
                    <th style={{ ...thStyle, textAlign: "center" }}>SG</th>
                    <th style={{ ...thStyle, textAlign: "center" }}>Step</th>
                    <th style={{ ...thStyle, textAlign: "right" }}>Monthly Pay</th>
                    <th style={thStyle}></th>
                  </tr>
                </thead>
                <tbody>
                  {employees.map((emp, i) => {
                    const isEditing = editId === emp.id;
                    const hasSG = !!emp.salary_grade;
                    const payMismatch = hasSG && emp.rate_amount != null &&
                      Math.abs(Number(emp.basic_monthly_pay) - Number(emp.rate_amount)) > 0.01;
                    return (
                      <tr key={emp.id} style={{ borderBottom: "1px solid var(--border)", background: i % 2 === 0 ? "var(--surface)" : "var(--surface-2)" }}>
                        <td style={{ ...tdStyle, color: "var(--text-muted)", fontSize: "0.8rem", fontFamily: "monospace" }}>{emp.employee_number ?? "—"}</td>
                        <td style={{ ...tdStyle, fontWeight: 700 }}>{emp.name}</td>
                        <td style={{ ...tdStyle, color: "var(--text-muted)", fontSize: "0.83rem", maxWidth: 200 }}>{emp.position || "—"}</td>
                        <td style={{ ...tdStyle, textAlign: "center" }}>
                          {hasSG ? (
                            <span style={{ fontSize: "0.75rem", background: "#eff6ff", color: "#1d4ed8", border: "1px solid #bfdbfe", borderRadius: 4, padding: "3px 8px", fontWeight: 700 }}>
                              SG {emp.salary_grade}
                            </span>
                          ) : <span style={{ color: "var(--text-muted)" }}>—</span>}
                        </td>
                        <td style={{ ...tdStyle, textAlign: "center" }}>
                          {isEditing ? (
                            <select
                              value={editStep}
                              onChange={e => setEditStep(parseInt(e.target.value, 10))}
                              style={{ padding: "5px 8px", borderRadius: 4, border: "1.5px solid #4a90d9", fontSize: "0.88rem", background: "var(--surface)", cursor: "pointer" }}
                            >
                              {[1,2,3,4,5,6,7,8].map(s => <option key={s} value={s}>Step {s}</option>)}
                            </select>
                          ) : (
                            <span style={{ fontWeight: 600, color: hasSG ? "var(--text)" : "var(--text-muted)" }}>
                              Step {emp.step ?? 1}
                            </span>
                          )}
                        </td>
                        <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, fontVariantNumeric: "tabular-nums" }}>
                          ₱{Number(emp.basic_monthly_pay).toLocaleString("en-PH", { minimumFractionDigits: 2 })}
                          {payMismatch && (
                            <div style={{ fontSize: "0.72rem", color: "#d97706", fontWeight: 400, marginTop: 2 }}>
                              table: ₱{Number(emp.rate_amount).toLocaleString("en-PH", { minimumFractionDigits: 2 })}
                            </div>
                          )}
                        </td>
                        <td style={{ ...tdStyle, whiteSpace: "nowrap" }}>
                          {isEditing ? (
                            <div style={{ display: "flex", gap: 6 }}>
                              <button
                                onClick={() => saveStep(emp.id)} disabled={saving}
                                style={{ padding: "5px 14px", fontSize: "0.82rem", background: "var(--psa-blue)", color: "#fff", border: "none", borderRadius: 4, cursor: "pointer", fontWeight: 600 }}
                              >{saving ? "…" : "Save"}</button>
                              <button
                                onClick={() => setEditId(null)}
                                style={{ padding: "5px 10px", fontSize: "0.82rem", background: "var(--surface)", border: "1.5px solid var(--border-2)", borderRadius: 4, cursor: "pointer" }}
                              >Cancel</button>
                            </div>
                          ) : (
                            <button
                              onClick={() => startEdit(emp)}
                              style={{ padding: "5px 12px", fontSize: "0.82rem", background: "var(--surface)", border: "1.5px solid var(--border-2)", borderRadius: "var(--radius-sm)", cursor: "pointer", fontWeight: 600, color: "var(--text)" }}
                              onMouseEnter={e => e.currentTarget.style.background = "var(--surface-2)"}
                              onMouseLeave={e => e.currentTarget.style.background = "var(--surface)"}
                            >Edit Step</button>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </>
  );
}


// ── PayrollTypesPanel ─────────────────────────────────────────────────────────
function PayrollTypesPanel({ canManage = true }) {
  const { dedTypes, earnTypes, setDedTypes, setEarnTypes } = useContext(TypesContext);
  const [activeTab, setActiveTab]       = useState("deduction");
  const [modalOpen, setModalOpen]       = useState(false);
  const [editItem, setEditItem]         = useState(null);
  const [saving, setSaving]             = useState(false);
  const [showAll, setShowAll]           = useState(false);
  const [showInactive, setShowInactive] = useState(false);
  const [allDedTypes, setAllDedTypes]   = useState(null);
  const [allEarnTypes, setAllEarnTypes] = useState(null);

  async function loadAllTypes() {
    const [dRes, eRes] = await Promise.all([
      fetch(`${apiUrl}/settings/deduction-types?all=1`, { credentials: "include" }),
      fetch(`${apiUrl}/settings/earning-types?all=1`,   { credentials: "include" }),
    ]);
    if (dRes.ok) setAllDedTypes(await dRes.json());
    if (eRes.ok) setAllEarnTypes(await eRes.json());
  }

  function toggleInactive() {
    const next = !showInactive;
    setShowInactive(next);
    if (next && allDedTypes === null) loadAllTypes();
  }

  const emptyForm = {
    display_name: "", name: "",
    catMode: "standalone",   // "standalone" | "categorized"
    category: "", catSelectVal: "__new__", newCatInput: "",
    side: "deduction", computation_mode: "manual",
    rate: "", amount_fixed: "", sort_order: "0", is_active: true,
    merge_in_excel: false, is_loan: false,
  };
  const [form, setForm] = useState(emptyForm);

  const slugify = str => str.toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "");

  // groupCats = categories with 2+ manual items (GSIS, PAG-IBIG, PROVIDENT — real picker groups)
  // standaloneCats = categories with 0-1 manual items (single rows or sibling pairs like MEDIC)
  const catManualCount = {};
  for (const t of dedTypes) {
    if (t.side === form.side && t.category && t.computation_mode !== 'rate') {
      catManualCount[t.category] = (catManualCount[t.category] || 0) + 1;
    }
  }
  // Also include computed-only categories (rate items with no manual siblings) in standaloneCats
  for (const t of dedTypes) {
    if (t.side === form.side && t.category && t.computation_mode === 'rate') {
      if (catManualCount[t.category] === undefined) catManualCount[t.category] = 0;
    }
  }
  const groupCats     = Object.keys(catManualCount).filter(c => catManualCount[c] >= 2).sort();
  const standaloneCats = Object.keys(catManualCount).filter(c => catManualCount[c] <= 1).sort();

  function openAdd() {
    setEditItem(null);
    const typesList = activeTab === "income" ? earnTypes : dedTypes;
    const maxSort = typesList.length > 0 ? Math.max(...typesList.map(t => t.sort_order ?? 0)) : -10;
    setForm({ ...emptyForm, sort_order: String(maxSort + 10) });
    setModalOpen(true);
  }

  function openEdit(item) {
    markRecent(item.id);
    setEditItem(item);
    const itemSide = item.side || "deduction";
    const manCount = {};
    for (const t of dedTypes) {
      if (t.side === itemSide && t.category && t.computation_mode !== 'rate') {
        manCount[t.category] = (manCount[t.category] || 0) + 1;
      }
    }
    const existingCat = item.category || "";
    const isPickerGroup = (manCount[existingCat] ?? 0) >= 2;
    const catMode = isPickerGroup ? "categorized" : "standalone";
    setForm({
      display_name:     item.display_name || item.name,
      name:             item.name,
      catMode,
      category:         existingCat,
      catSelectVal:     existingCat || "__new__",
      newCatInput:      "",
      side:             itemSide,
      computation_mode: item.computation_mode || "manual",
      rate:             item.rate != null ? String(Math.round(Number(item.rate) * 10000) / 100) : "",
      amount_fixed:     item.amount_fixed != null ? String(item.amount_fixed) : "",
      sort_order:       String(item.sort_order ?? 0),
      is_active:        !!item.is_active,
      merge_in_excel:   !!item.merge_in_excel,
      is_loan:          !!item.is_loan,
    });
    setModalOpen(true);
  }

  function closeModal() { setModalOpen(false); setEditItem(null); setForm(emptyForm); }

  async function handleSubmit(e) {
    e.preventDefault();
    if (!form.display_name.trim()) return;
    setSaving(true);
    try {
      const isIncome = activeTab === "income";
      const endpoint = isIncome ? "earning-types" : "deduction-types";
      const payload  = {
        display_name:     form.display_name.trim(),
        computation_mode: form.computation_mode,
        rate:             form.computation_mode === "rate"  ? (Number(form.rate) / 100 || null) : null,
        amount_fixed:     form.computation_mode === "fixed" ? (Number(form.amount_fixed) || null) : null,
        sort_order:       parseInt(form.sort_order, 10) || 0,
        is_active:        form.is_active,
      };
      if (!isIncome) {
        // Standalone new → uppercase the typed name; categorized → use as-is
        const rawCat = form.category.trim();
        payload.category = rawCat
          ? (form.catMode === "standalone" && form.catSelectVal === "__new__" ? rawCat.toUpperCase() : rawCat)
          : null;
        payload.side = form.side || "deduction";
        payload.merge_in_excel = form.catMode === "categorized" ? !!form.merge_in_excel : false;
        payload.is_loan = !!form.is_loan;
      }
      let res;
      if (editItem) {
        res = await fetch(`${apiUrl}/settings/${endpoint}/${editItem.id}`, {
          method: "PATCH", credentials: "include",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
        });
      } else {
        const namePart = form.name.trim() || slugify(form.display_name);
        if (!namePart) { alert("Could not generate a slug from the display name."); setSaving(false); return; }
        payload.name = namePart;
        res = await fetch(`${apiUrl}/settings/${endpoint}`, {
          method: "POST", credentials: "include",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
        });
      }
      if (!res.ok) { const d = await res.json(); alert(d.error ?? "Failed to save."); return; }
      const refreshRes = await fetch(`${apiUrl}/settings/${endpoint}`, { credentials: "include" });
      if (refreshRes.ok) {
        if (isIncome) setEarnTypes(await refreshRes.json());
        else setDedTypes(await refreshRes.json());
      }
      if (allDedTypes !== null) await loadAllTypes();
      closeModal();
    } finally { setSaving(false); }
  }

  async function handleDelete(item) {
    const isIncome = activeTab === "income";
    const endpoint = isIncome ? "earning-types" : "deduction-types";
    const label    = item.display_name || item.name;
    if (!window.confirm(`Remove "${label}"?\n\nIf employees have data for this type it will be deactivated instead of deleted.`)) return;
    const res = await fetch(`${apiUrl}/settings/${endpoint}/${item.id}`, { method: "DELETE", credentials: "include" });
    if (!res.ok) { const d = await res.json(); alert(d.error ?? "Failed to delete."); return; }
    const result = await res.json();
    if (result.deactivated) alert(`"${label}" has existing data and was deactivated (hidden from the editor).`);
    const refreshRes = await fetch(`${apiUrl}/settings/${endpoint}`, { credentials: "include" });
    if (refreshRes.ok) {
      if (isIncome) setEarnTypes(await refreshRes.json());
      else setDedTypes(await refreshRes.json());
    }
    if (allDedTypes !== null) await loadAllTypes();
  }

  const activeOnlyList = activeTab === "income" ? earnTypes : dedTypes;
  const fullList       = activeTab === "income" ? (allEarnTypes ?? earnTypes) : (allDedTypes ?? dedTypes);
  const currentList    = showInactive ? fullList : activeOnlyList;
  const LIMIT = 8;

  // Recent-first ordering: IDs stored in localStorage, most-recently-opened first
  const [recentIds, setRecentIds] = useState(() => {
    try { return JSON.parse(localStorage.getItem(`psa_recent_${activeTab}`) || "[]"); }
    catch { return []; }
  });
  useEffect(() => {
    try { setRecentIds(JSON.parse(localStorage.getItem(`psa_recent_${activeTab}`) || "[]")); }
    catch { setRecentIds([]); }
  }, [activeTab]);

  function markRecent(id) {
    const updated = [id, ...recentIds.filter(x => x !== id)].slice(0, 40);
    setRecentIds(updated);
    localStorage.setItem(`psa_recent_${activeTab}`, JSON.stringify(updated));
  }

  // Sort by recency, then by sort_order for the rest
  const recentRank = Object.fromEntries(recentIds.map((id, i) => [id, i]));
  const sortedList = [...currentList].sort((a, b) => {
    const ra = recentRank[a.id] ?? Infinity;
    const rb = recentRank[b.id] ?? Infinity;
    if (ra !== rb) return ra - rb;
    return (a.sort_order ?? 0) - (b.sort_order ?? 0);
  });

  const [searchQuery, setSearchQuery] = useState("");
  const [typeFilter, setTypeFilter]   = useState("all");
  const sq = searchQuery.trim().toLowerCase();
  const filteredList = sq
    ? sortedList.filter(t =>
        (t.display_name || t.name).toLowerCase().includes(sq) ||
        (t.category || "").toLowerCase().includes(sq)
      )
    : sortedList;

  const isSearching = !!sq;
  const visible = (isSearching || showAll) ? filteredList : filteredList.slice(0, LIMIT);

  const sideLabel = s => s === "less_deduction" ? "Less Ded." : "Deduction";
  const sideColor = s => s === "less_deduction"
    ? { bg: "#ecfdf5", color: "#065f46", border: "#6ee7b7" }
    : { bg: "#fef3c7", color: "#92400e", border: "#fcd34d" };
  const modeTag = item => {
    if (item.computation_mode === "rate")  return `${(Number(item.rate) * 100).toFixed(2)}% of basic`;
    if (item.computation_mode === "fixed") return `Fixed ₱${Number(item.amount_fixed ?? 0).toLocaleString("en-PH", { minimumFractionDigits: 2 })}`;
    return null;
  };

  const inputStyle = { width: "100%", padding: "10px 12px", fontSize: "0.9rem", borderRadius: 6, border: "1.5px solid var(--border-2)", background: "var(--surface)", color: "var(--text)", outline: "none", boxSizing: "border-box" };
  const labelStyle = { display: "flex", flexDirection: "column", gap: 6, fontSize: "0.83rem", fontWeight: 600, color: "var(--text-label)", textTransform: "uppercase", letterSpacing: "0.04em" };
  const fieldHeadStyle = { fontSize: "0.8rem", fontWeight: 700, color: "var(--text-label)", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 10 };
  const radioRowStyle  = { display: "flex", gap: 24 };
  const radioLabelStyle = { display: "flex", alignItems: "center", gap: 8, cursor: "pointer", fontSize: "0.92rem", fontWeight: 500, color: "var(--text)" };
  const radioInputStyle = { width: 16, height: 16, accentColor: "var(--psa-navy)", cursor: "pointer", flexShrink: 0 };

  return (
    <>
      {/* Modal */}
      {modalOpen && (
        <div
          onClick={e => { if (e.target === e.currentTarget) closeModal(); }}
          style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}
        >
          <div style={{ background: "var(--surface)", borderRadius: 10, width: "100%", maxWidth: 460, boxShadow: "0 20px 60px rgba(0,0,0,0.25)", overflow: "hidden" }}>

            {/* Modal header */}
            <div style={{ padding: "18px 24px", borderBottom: "1px solid var(--border)", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
              <div style={{ fontWeight: 700, fontSize: "1rem", color: "var(--text)" }}>
                {editItem ? `Edit — ${editItem.display_name || editItem.name}` : `Add ${activeTab === "income" ? "Income" : "Deduction"} Type`}
              </div>
              <button type="button" onClick={closeModal} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--text-muted)", padding: 4, display: "flex" }}>
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                  <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
                </svg>
              </button>
            </div>

            {/* Modal body */}
            <form onSubmit={handleSubmit}>
              <div style={{ padding: "20px 24px", display: "flex", flexDirection: "column", gap: 18 }}>

                {/* 1. Side — only shown for standalone items; categorized items are always deduction side */}
                {activeTab === "deduction" && form.catMode === "standalone" && (
                  <div>
                    <div style={fieldHeadStyle}>Side</div>
                    <div style={radioRowStyle}>
                      {[["less_deduction", "Less Deduction"], ["deduction", "Deduction"]].map(([val, lbl]) => (
                        <label key={val} style={radioLabelStyle}>
                          <input
                            type="radio" name="modal-side" value={val}
                            checked={form.side === val}
                            onChange={() => setForm(f => ({ ...f, side: val, catMode: "standalone", category: "", catSelectVal: "__new__", newCatInput: "" }))}
                            style={radioInputStyle}
                          />
                          {lbl}
                        </label>
                      ))}
                    </div>
                  </div>
                )}

                {/* 2. Category — Standalone | Categorized (deduction tab only) */}
                {activeTab === "deduction" && (
                  <div>
                    <div style={fieldHeadStyle}>Category</div>

                    {/* Mode radio */}
                    <div style={{ ...radioRowStyle, marginBottom: 10 }}>
                      {[["standalone", "Standalone"], ["categorized", "Categorized"]].map(([val, lbl]) => (
                        <label key={val} style={radioLabelStyle}>
                          <input
                            type="radio" name="modal-catmode" value={val}
                            checked={form.catMode === val}
                            onChange={() => {
                              const allCats = [...new Set([...groupCats, ...standaloneCats.filter(c => catManualCount[c] >= 1)])].sort();
                              const firstCat = allCats[0] || "";
                              setForm(f => ({
                                ...f, catMode: val,
                                catSelectVal: val === "categorized" ? (firstCat || "__new__") : "__new__",
                                category:     val === "categorized" ? firstCat : "",
                                newCatInput:  "",
                              }));
                            }}
                            style={radioInputStyle}
                          />
                          {lbl}
                        </label>
                      ))}
                    </div>

                    {/* Standalone: existing single-item cats + Add New */}
                    {form.catMode === "standalone" && (
                      <>
                        <select
                          value={form.catSelectVal}
                          onChange={e => {
                            const v = e.target.value;
                            setForm(f => ({ ...f, catSelectVal: v, category: v === "__new__" ? f.newCatInput.toUpperCase() : v, newCatInput: v === "__new__" ? f.newCatInput : "" }));
                          }}
                          style={{ ...inputStyle, cursor: "pointer" }}
                        >
                          <option value="__new__">+ Add New Standalone</option>
                          {standaloneCats.map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                        {form.catSelectVal === "__new__" && (
                          <input
                            type="text" autoFocus
                            placeholder="e.g. CEMCO, Children Fund"
                            value={form.newCatInput}
                            onChange={e => { const v = e.target.value; setForm(f => ({ ...f, newCatInput: v, category: v.toUpperCase() })); }}
                            style={{ ...inputStyle, marginTop: 8 }}
                          />
                        )}
                      </>
                    )}

                    {/* Categorized: existing groups + New Category */}
                    {form.catMode === "categorized" && (
                      <>
                        <select
                          value={form.catSelectVal}
                          onChange={e => {
                            const v = e.target.value;
                            setForm(f => ({ ...f, catSelectVal: v, category: v === "__new__" ? f.newCatInput : v, newCatInput: v === "__new__" ? f.newCatInput : "" }));
                          }}
                          style={{ ...inputStyle, cursor: "pointer" }}
                        >
                          {[...new Set([...groupCats, ...standaloneCats.filter(c => catManualCount[c] >= 1)])].sort().map(c => <option key={c} value={c}>{c}</option>)}
                          <option value="__new__">+ New Category</option>
                        </select>
                        {form.catSelectVal === "__new__" && (
                          <input
                            type="text" autoFocus
                            placeholder="e.g. GSIS, PAG-IBIG"
                            value={form.newCatInput}
                            onChange={e => { const v = e.target.value; setForm(f => ({ ...f, newCatInput: v, category: v })); }}
                            style={{ ...inputStyle, marginTop: 8 }}
                          />
                        )}
                      </>
                    )}
                  </div>
                )}

                {/* 3. Name */}
                <label style={labelStyle}>
                  Name
                  <input
                    style={inputStyle} value={form.display_name} required
                    placeholder={activeTab === "income" ? "e.g. RATA, Honorarium" : "e.g. Children Fund"}
                    onChange={e => {
                      const v = e.target.value;
                      setForm(f => ({ ...f, display_name: v, ...(editItem ? {} : { name: slugify(v) }) }));
                    }}
                  />
                </label>

                {/* 4. Type — radio buttons */}
                <div>
                  <div style={fieldHeadStyle}>Type</div>
                  <div style={radioRowStyle}>
                    {[["manual", "Manual"], ["rate", "Rate"], ["fixed", "Fixed"]].map(([val, lbl]) => (
                      <label key={val} style={radioLabelStyle}>
                        <input
                          type="radio" name="modal-computation" value={val}
                          checked={form.computation_mode === val}
                          onChange={() => setForm(f => ({ ...f, computation_mode: val }))}
                          style={radioInputStyle}
                        />
                        {lbl}
                      </label>
                    ))}
                  </div>
                </div>

                {/* Rate field */}
                {form.computation_mode === "rate" && (
                  <label style={labelStyle}>
                    Rate (%)
                    <input type="number" step="0.01" min="0" max="100" style={inputStyle}
                      value={form.rate} placeholder="e.g. 9 for 9%"
                      onChange={e => setForm(f => ({ ...f, rate: e.target.value }))}
                    />
                  </label>
                )}
                {/* Fixed amount field */}
                {form.computation_mode === "fixed" && (
                  <label style={labelStyle}>
                    Fixed Amount (₱)
                    <input type="number" step="0.01" min="0" style={inputStyle}
                      value={form.amount_fixed} placeholder="2000.00"
                      onChange={e => setForm(f => ({ ...f, amount_fixed: e.target.value }))}
                    />
                  </label>
                )}

                {/* Sort Order — edit only */}
                {editItem && (
                  <label style={labelStyle}>
                    Sort Order
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <input type="number" min="0" step="1" style={{ ...inputStyle, width: 90 }}
                        value={form.sort_order}
                        onChange={e => setForm(f => ({ ...f, sort_order: e.target.value }))}
                      />
                      <span style={{ fontSize: "0.78rem", color: "var(--text-muted)", fontWeight: 400, textTransform: "none", letterSpacing: 0 }}>
                        Lower = appears first
                      </span>
                    </div>
                  </label>
                )}

                {/* Is Loan — deduction types only */}
                {activeTab === "deduction" && (
                  <label style={{ display: "flex", alignItems: "center", gap: 10, cursor: "pointer" }}>
                    <input type="checkbox" checked={form.is_loan}
                      onChange={e => setForm(f => ({ ...f, is_loan: e.target.checked }))}
                      style={{ width: 16, height: 16, cursor: "pointer", accentColor: "var(--psa-navy)", flexShrink: 0 }}
                    />
                    <span style={{ fontSize: "0.88rem", fontWeight: 600, color: "var(--text)" }}>Is Loan</span>
                    <span style={{ fontSize: "0.78rem", color: "var(--text-muted)" }}>Enables loan tracking in the Loans tab</span>
                  </label>
                )}

                {/* Active toggle (edit only) */}
                {editItem && (
                  <label style={{ display: "flex", alignItems: "center", gap: 10, cursor: "pointer" }}>
                    <input type="checkbox" checked={form.is_active}
                      onChange={e => setForm(f => ({ ...f, is_active: e.target.checked }))}
                      style={{ width: 16, height: 16, cursor: "pointer", accentColor: "var(--psa-navy)", flexShrink: 0 }}
                    />
                    <span style={{ fontSize: "0.88rem", fontWeight: 600, color: "var(--text)" }}>Active</span>
                    <span style={{ fontSize: "0.78rem", color: "var(--text-muted)" }}>visible in the deduction editor</span>
                  </label>
                )}

                {/* Merge in Excel — only for categorized deduction items, shown below Active */}
                {activeTab === "deduction" && form.catMode === "categorized" && (
                  <label style={{ display: "flex", alignItems: "center", gap: 10, cursor: "pointer" }}>
                    <input
                      type="checkbox"
                      checked={form.merge_in_excel}
                      onChange={e => setForm(f => ({ ...f, merge_in_excel: e.target.checked }))}
                      style={{ width: 16, height: 16, cursor: "pointer", accentColor: "var(--psa-navy)", flexShrink: 0 }}
                    />
                    <span style={{ fontSize: "0.88rem", fontWeight: 600, color: "var(--text)", whiteSpace: "nowrap" }}>Merge in Excel</span>
                    <span style={{ fontSize: "0.78rem", color: "var(--text-muted)" }}>Group with other checked items under a shared header in Excel</span>
                  </label>
                )}
              </div>

              {/* Modal footer */}
              <div style={{ padding: "14px 24px", borderTop: "1px solid var(--border)", display: "flex", justifyContent: "flex-end", gap: 10 }}>
                <button type="button" onClick={closeModal}
                  style={{ padding: "9px 20px", fontSize: "0.88rem", fontWeight: 600, borderRadius: 6, border: "1.5px solid var(--border-2)", background: "var(--surface)", color: "var(--text)", cursor: "pointer" }}
                >Cancel</button>
                <button type="submit" disabled={saving} className="btn-primary"
                  style={{ padding: "9px 20px", fontSize: "0.88rem", opacity: saving ? 0.6 : 1, cursor: saving ? "not-allowed" : "pointer" }}
                >{saving ? "Saving…" : editItem ? "Save Changes" : `Add ${activeTab === "income" ? "Income" : "Deduction"} Type`}</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {(() => {
        const isManual = (t) => !t.computation_mode || t.computation_mode === "manual";
        const KINDS = [
          ["all",      "All",          () => true],
          ["computed", "Computed",     t => !isManual(t)],
          ["manual",   "Per employee", isManual],
          ["loan",     "Loans",        t => !!t.is_loan],
        ];
        const kindCounts = Object.fromEntries(KINDS.map(([k, , f]) => [k, filteredList.filter(f).length]));
        const kinds = KINDS.filter(([k]) => k === "all" || kindCounts[k] > 0);
        const inKind = KINDS.find(([k]) => k === typeFilter)?.[2] ?? (() => true);
        const list = filteredList.filter(inKind);
        const expanded = isSearching || showAll || typeFilter !== "all";
        const rows = expanded ? list : list.slice(0, LIMIT);
        const computeLabel = (t) => modeTag(t) ?? "Entered per employee";
        return (
          <>
            <div className="pt-top">
              <div className="pt-tabs" role="tablist" aria-label="Payroll type">
                {[["income", "Income", earnTypes.length], ["deduction", "Deductions", dedTypes.length]].map(([id, label, n]) => (
                  <button key={id} type="button" role="tab" aria-selected={activeTab === id}
                    className={`pt-tab${activeTab === id ? " is-on" : ""}`}
                    onClick={() => { setActiveTab(id); setShowAll(false); setSearchQuery(""); setTypeFilter("all"); }}>
                    {label}<span className="pt-tab-n">{n}</span>
                  </button>
                ))}
              </div>
              {canManage && (
                <button type="button" className="rm-create pt-add" onClick={openAdd}>
                  <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  Add {activeTab === "income" ? "income" : "deduction"} type
                </button>
              )}
            </div>

            <div className="pt-tools">
              <div className="um-search pt-search">
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
                <input type="text" placeholder="Search by name or category…" value={searchQuery}
                  onChange={e => { setSearchQuery(e.target.value); setShowAll(false); }} />
                {searchQuery && <button type="button" className="um-clear" onClick={() => setSearchQuery("")} aria-label="Clear search">×</button>}
              </div>
              <div className="um-filters pt-filters">
                {kinds.map(([k, label]) => (
                  <button key={k} type="button" className={`um-filter${typeFilter === k ? " is-on" : ""}`}
                    onClick={() => { setTypeFilter(k); setShowAll(false); }}>
                    {label}<span className="um-filter-n">{kindCounts[k]}</span>
                  </button>
                ))}
              </div>
              <label className="pt-inactive">
                <input type="checkbox" checked={showInactive} onChange={toggleInactive} />
                Show deactivated
              </label>
            </div>

            {currentList.length === 0 ? (
              <div className="um-empty">No types yet.{canManage ? " Use “Add” to create one." : ""}</div>
            ) : list.length === 0 ? (
              <div className="um-empty">No types match{sq ? ` “${searchQuery.trim()}”` : " this filter"}.</div>
            ) : (
              <div className="pt-list">
                {rows.map(item => (
                  <div key={item.id} className={`pt-row${item.is_active ? "" : " is-inactive"}`}>
                    <span className={`pt-mode-dot pt-mode--${isManual(item) ? "manual" : item.computation_mode}`} aria-hidden="true" />
                    <div className="pt-name-cell">
                      <span className="pt-name">{item.display_name || item.name}</span>
                      {item.category && <span className="um-sub">{item.category}</span>}
                    </div>
                    <span className="pt-compute">{computeLabel(item)}</span>
                    <span className="pt-tags">
                      {!!item.is_loan && <span className="pt-tag pt-tag--loan">Loan</span>}
                      {item.side === "less_deduction" && <span className="pt-tag pt-tag--less">Less deduction</span>}
                      {!item.is_active && <span className="pt-tag pt-tag--off">Inactive</span>}
                    </span>
                    {canManage && (
                      <span className="pt-actions">
                        <button type="button" className="um-btn um-btn--neutral um-btn--sm" onClick={() => openEdit(item)}>Edit</button>
                        {item.name !== "basic_pay" ? (
                          <button type="button" className="pt-del" onClick={() => handleDelete(item)} title="Remove" aria-label={`Remove ${item.display_name || item.name}`}>
                            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                              <polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/>
                              <path d="M10 11v6"/><path d="M14 11v6"/><path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/>
                            </svg>
                          </button>
                        ) : <span className="pt-del-spacer" />}
                      </span>
                    )}
                  </div>
                ))}
              </div>
            )}

            {!expanded && list.length > LIMIT && (
              <button type="button" className="um-btn um-btn--neutral pt-more" onClick={() => setShowAll(true)}>
                Show {list.length - LIMIT} more
              </button>
            )}
            {showAll && !isSearching && typeFilter === "all" && list.length > LIMIT && (
              <button type="button" className="um-btn um-btn--neutral pt-more" onClick={() => setShowAll(false)}>Show less</button>
            )}
          </>
        );
      })()}

    </>
  );
}



// ── PositionView ──────────────────────────────────────────────────────────────

// Grade codes are stored inconsistently ("SG7", "SG 9"); show them uniformly as "SG 7".
const sgDisplay = (code) => code ? `SG ${String(code).replace(/^sg\s*/i, "").trim()}` : "";

const SgCaret = ({ open }) => (
  <svg className="sgd-caret" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true" style={{ transform: open ? "rotate(180deg)" : "none" }}>
    <polyline points="6 9 12 15 18 9"/>
  </svg>
);
const SgCheck = () => (
  <svg className="sgd-check" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><polyline points="20 6 9 17 4 12"/></svg>
);

// Menu is portaled to <body> with fixed positioning so card overflow can't clip it;
// it opens upward when there isn't room below.
function useAnchoredMenu(open, setOpen, maxH = 280) {
  const anchorRef = useRef(null);
  const menuRef = useRef(null);
  const [pos, setPos] = useState(null);
  useLayoutEffect(() => {
    if (!open) { setPos(null); return; }
    function place(e) {
      if (e && e.type === "scroll" && menuRef.current?.contains(e.target)) return;
      const r = anchorRef.current?.getBoundingClientRect();
      if (!r) return;
      const gap = 6, edge = 12;
      const below = window.innerHeight - r.bottom - gap - edge;
      const above = r.top - gap - edge;
      const up = below < Math.min(maxH, 220) && above > below;
      setPos({
        left: r.left, width: r.width,
        top: up ? undefined : r.bottom + gap,
        bottom: up ? window.innerHeight - r.top + gap : undefined,
        maxHeight: Math.max(140, Math.min(maxH, up ? above : below)),
      });
    }
    place();
    window.addEventListener("resize", place);
    window.addEventListener("scroll", place, true);
    return () => { window.removeEventListener("resize", place); window.removeEventListener("scroll", place, true); };
  }, [open, maxH]);
  useEffect(() => {
    if (!open) return;
    function handleClick(e) {
      if (anchorRef.current?.contains(e.target) || menuRef.current?.contains(e.target)) return;
      setOpen(false);
    }
    function handleKey(e) { if (e.key === "Escape") setOpen(false); }
    document.addEventListener("mousedown", handleClick);
    document.addEventListener("keydown", handleKey);
    return () => { document.removeEventListener("mousedown", handleClick); document.removeEventListener("keydown", handleKey); };
  }, [open, setOpen]);
  // Bring the selected option into view when the menu opens.
  const placed = !!pos;
  useEffect(() => {
    const menu = menuRef.current;
    const on = menu?.querySelector(".is-on");
    if (on) menu.scrollTop = on.offsetTop - menu.clientHeight / 2 + on.offsetHeight / 2;
  }, [placed]);
  const menuStyle = pos ? { position: "fixed", left: pos.left, width: pos.width, top: pos.top, bottom: pos.bottom, maxHeight: pos.maxHeight } : null;
  return { anchorRef, menuRef, menuStyle };
}

function SalaryGradeDropdown({ value, onChange, sgOptions, sgSteps = [] }) {
  const [open, setOpen] = useState(false);
  const { anchorRef, menuRef, menuStyle } = useAnchoredMenu(open, setOpen);
  const rangeOf = (code) => {
    const amts = sgSteps.filter(s => s.salary_grade === code).map(s => Number(s.amount)).filter(n => n > 0);
    if (!amts.length) return "No steps yet";
    const peso = n => `₱${n.toLocaleString("en-PH", { maximumFractionDigits: 0 })}`;
    return `${peso(Math.min(...amts))} – ${peso(Math.max(...amts))}`;
  };
  return (
    <div ref={anchorRef} className="sgd">
      <button type="button" className={`sgd-trigger${open ? " is-open" : ""}`} onClick={() => setOpen(o => !o)} aria-haspopup="listbox" aria-expanded={open}>
        <span className="sgd-main">
          <span className={value ? "sgd-value" : "sgd-placeholder"}>{value ? sgDisplay(value) : "No salary grade"}</span>
          {value && <span className="sgd-range">{rangeOf(value)}</span>}
        </span>
        <SgCaret open={open} />
      </button>
      {open && menuStyle && createPortal(
        <div ref={menuRef} className="sgd-menu" style={menuStyle} role="listbox">
          {[{ id: "__none__", grade_code: "" }, ...sgOptions].map(sg => {
            const code = sg.grade_code || "";
            const on = code === (value || "");
            return (
              <button key={sg.id} type="button" role="option" aria-selected={on}
                className={`sgd-opt${on ? " is-on" : ""}`} onClick={() => { onChange(code); setOpen(false); }}>
                <span className="sgd-opt-name">{code ? sgDisplay(code) : "No salary grade"}</span>
                {code && <span className="sgd-opt-meta">{rangeOf(code)}</span>}
                {on && <SgCheck />}
              </button>
            );
          })}
        </div>,
        document.body
      )}
    </div>
  );
}

function PositionFormFields({ f, setF, sgOptions, sgSteps, existingNames = [], ignoreName = null }) {
  const fmtAmt = n => n == null ? "—" : `₱${Number(n).toLocaleString("en-PH", { minimumFractionDigits: 2 })}`;
  const steps = f.defaultSalaryGrade
    ? sgSteps.filter(s => s.salary_grade === f.defaultSalaryGrade).sort((a, b) => a.step - b.step)
    : [];
  const typed = f.name.trim().toLowerCase();
  const isDuplicate = !!typed && typed !== (ignoreName ?? "").trim().toLowerCase()
    && existingNames.some(n => (n ?? "").trim().toLowerCase() === typed);
  return (
    <div className="pf">
      <div className="pf-field">
        <label className="ar-label" htmlFor="pf-title">Position title <span className="pf-req">*</span></label>
        <input id="pf-title" value={f.name} onChange={e => setF(x => ({ ...x, name: e.target.value }))} placeholder="e.g. Statistician I" autoComplete="off" />
        {isDuplicate
          ? <span className="pf-warn">A position with this title already exists.</span>
          : <span className="pf-help">Use the title exactly as it appears on the plantilla.</span>}
      </div>
      <div className="pf-field">
        <span className="ar-label">Default salary grade</span>
        <SalaryGradeDropdown value={f.defaultSalaryGrade} onChange={v => setF(x => ({ ...x, defaultSalaryGrade: v }))} sgOptions={sgOptions} sgSteps={sgSteps} />
        <span className="pf-help">Optional. Links this position to a salary grade so its monthly steps show here.</span>
        {steps.length > 0 && (
          <div className="pf-steps">
            <div className="ws-kicker">{sgDisplay(f.defaultSalaryGrade)} · monthly salary by step</div>
            <div className="pf-steps-grid">
              {steps.map(s => (
                <div key={s.id} className="pf-step">
                  <span>Step {s.step}</span>
                  <strong>{fmtAmt(s.amount)}</strong>
                </div>
              ))}
            </div>
          </div>
        )}
        {f.defaultSalaryGrade && steps.length === 0 && (
          <div className="ws-hint">No salary steps are set up for {f.defaultSalaryGrade} yet. Add them in Salary Rank.</div>
        )}
      </div>
    </div>
  );
}

function PositionStepsTable({ grade, steps, fmt, fmtDate }) {
  if (!grade) return <div className="ws-hint">No default salary grade yet. Edit the position to assign one.</div>;
  if (steps.length === 0) return <div className="ws-hint">No salary steps are set up for {grade}. Add them in Salary Rank.</div>;
  return (
    <>
      <div className="ws-kicker" style={{ marginBottom: 8 }}>{sgDisplay(grade)} salary steps</div>
      <div style={{ overflowX: "auto", border: "1px solid var(--border)", borderRadius: 12 }}>
        <table className="ws-table">
          <thead>
            <tr><th>Step</th><th>Monthly amount</th><th>Effective</th></tr>
          </thead>
          <tbody>
            {steps.map(s => (
              <tr key={s.id ?? s.step}>
                <td style={{ fontWeight: 700, color: "var(--psa-navy)" }}>Step {s.step}</td>
                <td style={{ fontWeight: 600, color: "var(--psa-navy)" }}>{fmt(s.amount)}</td>
                <td style={{ color: "var(--text-muted)", fontSize: "0.8rem" }}>{s.effective_date ? fmtDate(s.effective_date) : "—"}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </>
  );
}

function PositionView({ noHeader = false, noPadding = false, listMaxHeight = null }) {
  const [positions, setPositions]       = useState([]);
  const [sgList, setSgList]             = useState([]);   // [{id, grade_code}]
  const [sgSteps, setSgSteps]           = useState([]);   // full steps [{salary_grade, step, amount}]
  const [selected, setSelected]         = useState(null);
  const [query, setQuery]               = useState("");
  const [loading, setLoading]           = useState(false);
  const [saving, setSaving]             = useState(false);
  const [adding, setAdding]             = useState(false);
  const [form, setForm]                 = useState({ name: "", defaultSalaryGrade: "" });
  const [editForm, setEditForm]         = useState(null);
  const isMobile = useIsMobile();
  const fmt = n => n == null ? "—" : `₱${Number(n).toLocaleString("en-PH", { minimumFractionDigits: 2 })}`;
  const fmtDate = d => { try { return new Date(d).toLocaleDateString("en-PH", { year: "numeric", month: "short", day: "numeric" }); } catch { return d; } };
  const hasPos = useRef(false);

  const fetchPositions = useCallback(async (q = "", silent = false) => {
    if (!silent) setLoading(true);
    try {
      const limit = q.trim() ? 50 : 5;
      const r = await fetch(`${apiUrl}/positions?q=${encodeURIComponent(q)}&limit=${limit}`, { credentials: "include" });
      if (r.ok) setPositions(await r.json());
    } finally { if (!silent) setLoading(false); }
  }, [apiUrl]);

  useEffect(() => {
    const loadMeta = async () => {
      const [sgListRes, sgStepsRes] = await Promise.all([
        fetch(`${apiUrl}/salary-grades/list`, { credentials: "include" }),
        fetch(`${apiUrl}/salary-grades`, { credentials: "include" }),
      ]);
      if (sgListRes.ok)  { const d = await sgListRes.json();  setSgList(Array.isArray(d) ? d : []); }
      if (sgStepsRes.ok) { const d = await sgStepsRes.json(); setSgSteps(Array.isArray(d) ? d : []); }
    };
    loadMeta();
  }, [apiUrl]);

  // The list only holds the 5 most recently used positions, so overview numbers come from the full list.
  const [posTotals, setPosTotals] = useState(null);
  const loadPosTotals = useCallback(async () => {
    const r = await fetch(`${apiUrl}/positions?limit=500`, { credentials: "include" });
    if (r.ok) {
      const all = await r.json();
      setPosTotals({ total: all.length, filled: all.filter(p => p.employee_count > 0).length, names: all.map(p => p.name) });
    }
  }, [apiUrl]);
  useEffect(() => { loadPosTotals(); }, [loadPosTotals]);

  const isFirstQuery = useRef(true);
  useEffect(() => { fetchPositions("", hasPos.current); hasPos.current = true; }, [fetchPositions]);

  useEffect(() => {
    if (isFirstQuery.current) { isFirstQuery.current = false; return; }
    const t = setTimeout(() => fetchPositions(query), 300);
    return () => clearTimeout(t);
  }, [query, fetchPositions]);

  const bumpPosToTop = (posId) => {
    if (query.trim()) return;
    setPositions(prev => {
      const idx = prev.findIndex(p => p.id === posId);
      if (idx <= 0) return prev;
      const updated = [...prev];
      const [pos] = updated.splice(idx, 1);
      updated.unshift(pos);
      return updated;
    });
  };

  const selectedPos = positions.find(p => p.id === selected) ?? null;

  const sgOptions = sgList.sort((a, b) =>
    parseFloat(a.grade_code.replace(/\D/g, "")) - parseFloat(b.grade_code.replace(/\D/g, ""))
  );

  const stepsForGrade = (gradeCode) =>
    sgSteps.filter(s => s.salary_grade === gradeCode).sort((a, b) => a.step - b.step);

  const postBody = (f) => ({
    name: f.name.trim(),
    defaultSalaryGrade: f.defaultSalaryGrade || null,
  });

  const createPosition = async () => {
    if (!form.name.trim()) return;
    setSaving(true);
    try {
      const res = await fetch(`${apiUrl}/positions`, {
        method: "POST", credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(postBody(form)),
      });
      if (res.ok) {
        const created = await res.json();
        setSelected(created.id);
        setAdding(false);
        setForm({ name: "", defaultSalaryGrade: "" });
        fetchPositions(query, true);
        loadPosTotals();
      } else {
        const err = await res.json();
        alert(err.error ?? "Failed to create position.");
      }
    } finally { setSaving(false); }
  };

  const saveEdit = async () => {
    if (!editForm?.name?.trim()) return;
    setSaving(true);
    try {
      const res = await fetch(`${apiUrl}/positions/${selected}`, {
        method: "PATCH", credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(postBody(editForm)),
      });
      if (res.ok) {
        const updated = await res.json();
        setPositions(prev => prev.map(p => p.id === selected ? { ...p, ...updated, name: updated.name ?? editForm.name } : p));
        setEditForm(null);
      }
    } finally { setSaving(false); }
  };

  const deletePosition = async (id) => {
    if (!window.confirm("Delete this position? Employees assigned here will lose their position link.")) return;
    await fetch(`${apiUrl}/positions/${id}`, { method: "DELETE", credentials: "include" });
    if (selected === id) setSelected(null);
    fetchPositions(query, true);
    loadPosTotals();
  };

  const totalPositions = posTotals?.total ?? positions.length;
  const filledCount = posTotals?.filled ?? positions.filter(p => p.employee_count > 0).length;

  const PositionSearchRow = (inFrame) => (
    <div className={inFrame ? "ws-toolbar" : "ws-toolbar ws-toolbar--bare"}>
      <WsSearch value={query} onChange={setQuery} placeholder="Search positions…" />
      <button type="button" className="btn-primary ws-add" onClick={() => { setAdding(true); setSelected(null); }}>+ Add position</button>
    </div>
  );

  const PositionCards = (
    <div>
      {loading ? (
        <WsSpinner />
      ) : positions.length === 0 ? (
        <div className="ws-empty">No positions found.</div>
      ) : positions.map(p => {
        const isSelected = selected === p.id;
        const isFilled = p.employee_count > 0;
        return (
          <div key={p.id} role="button" tabIndex={0}
            className={`ws-item${isSelected ? " is-on" : ""}`}
            onKeyDown={e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); e.currentTarget.click(); } }}
            onClick={() => {
              const opening = selected !== p.id;
              setSelected(p.id);
              setEditForm(null);
              if (opening) {
                bumpPosToTop(p.id);
                fetch(`${apiUrl}/positions/${p.id}/search-hit`, { method: "POST", credentials: "include" }).catch(() => {});
              }
            }}
          >
            <div className="ws-item-main">
              <div className="ws-item-title">{p.name}</div>
              <div className="ws-item-sub">{sgDisplay(p.default_salary_grade) || <em>No grade set</em>}</div>
            </div>
            <span className={`ws-pill${isFilled ? " ws-pill--ok" : ""}`}>
              {isFilled ? `${p.employee_count} filled` : "Vacant"}
            </span>
          </div>
        );
      })}
      {!loading && !query.trim() && totalPositions > positions.length && (
        <div className="ws-list-note">Showing your {positions.length} most recently used of {totalPositions} positions. Search to find the rest.</div>
      )}
    </div>
  );

  const ListPanel = isMobile ? (
    <>
      {PositionSearchRow(true)}
      <div style={{ padding: 8 }}>{PositionCards}</div>
    </>
  ) : (
    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
      <div className="panel" style={{ marginBottom: 0 }}>
        {PositionSearchRow(false)}
      </div>
      {PositionCards}
    </div>
  );

  const DetailPanel = adding ? (
    <div className={noPadding ? undefined : "panel"}>
      <div className="ws-detail-head">
        <div>
          <div className="ws-kicker">New position</div>
          <div className="ws-dtitle">{form.name.trim() || "Untitled position"}</div>
          <div className="ws-item-sub">{sgDisplay(form.defaultSalaryGrade) || "No salary grade yet"}</div>
        </div>
      </div>
      <div className="ws-pad">
        <PositionFormFields f={form} setF={setForm} sgOptions={sgOptions} sgSteps={sgSteps} existingNames={posTotals?.names ?? positions.map(p => p.name)} />
        <div className="ws-form-actions">
          <button type="button" className="um-btn um-btn--neutral" onClick={() => setAdding(false)} disabled={saving}>Cancel</button>
          <button type="button" className="btn-primary ws-btn-sm" onClick={createPosition} disabled={saving || !form.name.trim()}>{saving ? "Saving…" : "Save position"}</button>
        </div>
      </div>
    </div>
  ) : selectedPos ? (
    <div style={{ height: "100%", overflow: "auto" }}>
      <div className="ws-detail-head">
        <div>
          <div className="ws-kicker">Position</div>
          <div className="ws-dtitle">{selectedPos.name}</div>
          <div className="ws-item-sub">
            {sgDisplay(selectedPos.default_salary_grade) || "No grade set"} · {selectedPos.employee_count > 0 ? `${selectedPos.employee_count} employee${selectedPos.employee_count === 1 ? "" : "s"}` : "Vacant"}
          </div>
        </div>
        <div className="ws-actions">
          <button type="button" className="um-btn um-btn--danger" onClick={() => deletePosition(selectedPos.id)}>
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/></svg>
            Delete
          </button>
          <button type="button" className="um-btn um-btn--neutral"
            onClick={() => setEditForm({ name: selectedPos.name, defaultSalaryGrade: selectedPos.default_salary_grade ?? "" })}>
            Edit position
          </button>
        </div>
      </div>
      <div style={{ padding: isMobile ? "16px" : "20px" }}>
        {editForm ? (
          <>
            <PositionFormFields f={editForm} setF={setEditForm} sgOptions={sgOptions} sgSteps={sgSteps}
              existingNames={posTotals?.names ?? positions.map(p => p.name)} ignoreName={selectedPos.name} />
            <div className="ws-form-actions">
              <button type="button" className="um-btn um-btn--neutral" onClick={() => setEditForm(null)} disabled={saving}>Cancel</button>
              <button type="button" className="btn-primary ws-btn-sm" onClick={saveEdit} disabled={saving || !editForm.name.trim()}>{saving ? "Saving…" : "Save changes"}</button>
            </div>
          </>
        ) : (
          <PositionStepsTable
            grade={selectedPos.default_salary_grade}
            steps={stepsForGrade(selectedPos.default_salary_grade)}
            fmt={fmt}
            fmtDate={fmtDate}
          />
        )}
      </div>
    </div>
  ) : (
    <div className="ws-overview">
      <div className="ws-kicker">Positions overview</div>
      <div className="ws-stats">
        {[
          ["Positions", totalPositions, ""],
          ["Filled", filledCount, " ws-stat-value--ok"],
          ["Vacant", totalPositions - filledCount, " ws-stat-value--muted"],
        ].map(([lbl, val, tone]) => (
          <div key={lbl} className="ws-stat">
            <div className="ws-stat-label">{lbl}</div>
            <div className={`ws-stat-value${tone}`}>{val}</div>
          </div>
        ))}
      </div>
      <div className="ws-hint">
        <span>Select a position on the left to see its salary steps, or add a new one.</span>
        <button type="button" className="btn-primary ws-btn-sm" onClick={() => setAdding(true)}>+ Add position</button>
      </div>
    </div>
  );

  return (
    <>
      {!noHeader && (
        <div className="page-header">
          <div className="page-header-title">Positions</div>
          <div className="page-header-sub">Manage plantilla positions and their default salary grades</div>
        </div>
      )}
      {isMobile && adding && (
        <div className="ar-backdrop" role="dialog" aria-modal="true" aria-label="Add position">
          <div className="ar-panel">
            <div className="ar-head">
              <div>
                <div className="ar-title">Add position</div>
                <div className="um-sub">Enter the plantilla title and, optionally, its salary grade.</div>
              </div>
              <button type="button" className="ar-close" onClick={() => setAdding(false)} disabled={saving} aria-label="Close">✕</button>
            </div>
            <div className="ar-body" style={{ paddingTop: 16 }}>
              <PositionFormFields f={form} setF={setForm} sgOptions={sgOptions} sgSteps={sgSteps} existingNames={posTotals?.names ?? positions.map(p => p.name)} />
            </div>
            <div className="ar-foot">
              <div className="ar-foot-actions">
                <button type="button" className="um-btn um-btn--neutral" onClick={() => setAdding(false)} disabled={saving}>Cancel</button>
                <button type="button" className="btn-primary ws-btn-sm" onClick={createPosition} disabled={saving || !form.name.trim()}>{saving ? "Saving…" : "Save position"}</button>
              </div>
            </div>
          </div>
        </div>
      )}
      {isMobile ? (
        <div style={{ padding: noPadding ? "0" : "0 0 24px" }}>
          {selected ? (
            <>
              <div style={{ padding: "10px 16px", borderBottom: "1px solid var(--border)" }}>
                <button
                  style={{ display: "inline-flex", alignItems: "center", gap: 5, padding: "5px 12px", borderRadius: 7, border: "1.5px solid var(--border-2)", background: "transparent", color: "var(--text-muted)", fontSize: "0.8rem", fontWeight: 600, cursor: "pointer" }}
                  onClick={() => { setSelected(null); setAdding(false); }}>
                  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="15 18 9 12 15 6"/></svg>
                  Back
                </button>
              </div>
              {DetailPanel}
            </>
          ) : ListPanel}
        </div>
      ) : noPadding ? (
        <>
          {PositionSearchRow(true)}
          <div className="ws-split">
            <div className="ws-list">{PositionCards}</div>
            <div className="ws-detail">{DetailPanel}</div>
          </div>
        </>
      ) : (
        <div style={{ display: "flex", justifyContent: "center", padding: "0 24px 24px" }}>
          <div style={{ display: "flex", gap: 20, alignItems: "start" }}>
            <div style={{ width: 320, flexShrink: 0 }}>{ListPanel}</div>
            <div style={{ width: 500, flexShrink: 0 }}>{DetailPanel}</div>
          </div>
        </div>
      )}
    </>
  );
}

// ── WorkforceSetupView ────────────────────────────────────────────────────────

function WorkforceSetupView({ isAdmin = false, permissions = [] }) {
  const canManageRef = isAdmin || permissions.includes("settings.reference.manage");
  const [wfTab, setWfTab] = useState("positions");

  const tabs = [
    { id: "positions", label: "Positions" },
    { id: "reference", label: "Divisions" },
  ];

  return (
    <>
      <div className="page-header">
        <div className="page-header-title">Workforce Setup</div>
        <div className="page-header-sub">Manage positions and office divisions</div>
      </div>

      <div className="ws-frame-head">
        <WsTabs tabs={tabs} value={wfTab} onChange={setWfTab} label="Workforce setup sections" />
      </div>

      {/* Divisions are a short list, so that tab sizes to its rows instead of the tall frame. */}
      <div className={`ws-frame-body${wfTab === "reference" ? " ws-frame-body--fit" : ""}`}>
        {wfTab === "positions" && (
          <PositionView noHeader noPadding />
        )}
        {wfTab === "reference" && (
          <LookupTablePanel title="Divisions" endpoint="divisions" showItemNo={false} defaultLimit={null} canManage={canManageRef} />
        )}
      </div>
    </>
  );
}

function RankGradeDropdown({ value, onChange, existingGradeNumbers }) {
  const [open, setOpen] = useState(false);
  const { anchorRef, menuRef, menuStyle } = useAnchoredMenu(open, setOpen);
  const options = Array.from({ length: 33 }, (_, i) => i + 1);
  return (
    <div ref={anchorRef} className="sgd">
      <button type="button" className={`sgd-trigger${open ? " is-open" : ""}`} onClick={() => setOpen(o => !o)} aria-haspopup="listbox" aria-expanded={open}>
        <span className="sgd-main">
          <span className={value ? "sgd-value" : "sgd-placeholder"}>{value ? `SG ${value}` : "Select a salary grade"}</span>
        </span>
        <SgCaret open={open} />
      </button>
      {open && menuStyle && createPortal(
        <div ref={menuRef} className="sgd-menu" style={menuStyle} role="listbox">
          {options.map(sg => {
            const taken = existingGradeNumbers.has(String(sg));
            const on = String(sg) === String(value || "");
            return (
              <button key={sg} type="button" role="option" aria-selected={on} disabled={taken}
                className={`sgd-opt${on ? " is-on" : ""}`}
                onClick={() => { if (taken) return; onChange(String(sg)); setOpen(false); }}>
                <span className="sgd-opt-name">SG {sg}</span>
                {taken && <span className="sgd-opt-meta">Already added</span>}
                {on && <SgCheck />}
              </button>
            );
          })}
        </div>,
        document.body
      )}
    </div>
  );
}

// ── SalaryRankView ────────────────────────────────────────────────────────────

function SalaryRankView() {
  const [groups, setGroups]           = useState([]);
  const [selectedRank, setSelectedRank] = useState(null);
  const [query, setQuery]             = useState("");
  const [loading, setLoading]         = useState(false);
  const [addingRank, setAddingRank]   = useState(false);
  const [newRank, setNewRank]         = useState("");
  const [newStepCount, setNewStepCount] = useState(8);
  const [newAmounts, setNewAmounts]   = useState({});
  const [editId, setEditId]           = useState(null);
  const [editVal, setEditVal]         = useState("");
  const [saving, setSaving]           = useState(false);
  const isMobile = useIsMobile();
  const hasRanks = useRef(false);

  const fmt = (n) => n == null ? "—" : `₱${Number(n).toLocaleString("en-PH", { minimumFractionDigits: 2 })}`;
  const fmtDate = (d) => { try { return new Date(d).toLocaleDateString("en-PH", { year: "numeric", month: "short", day: "numeric" }); } catch { return d; } };

  const fetchRanks = useCallback(async (q = "", silent = false) => {
    if (!silent) setLoading(true);
    try {
      const limit = q.trim() ? 50 : 100;
      const r = await fetch(`${apiUrl}/salary-grades/ranked-list?q=${encodeURIComponent(q)}&limit=${limit}`, { credentials: "include" });
      if (r.ok) { const data = await r.json(); setGroups(data); hasRanks.current = true; }
    } finally { if (!silent) setLoading(false); }
  }, []);

  const isFirstRankQuery = useRef(true);
  useEffect(() => { fetchRanks(); }, [fetchRanks]);

  // Full list of grade codes already in use — needed to block duplicates in the
  // Add Rank dropdown, since `groups` above only holds a top-5/search-limited slice.
  // Every step of every grade, so the overview reflects all grades, not just the loaded slice.
  const [existingGradeCodes, setExistingGradeCodes] = useState([]);
  const [allSteps, setAllSteps] = useState([]);
  const fetchExistingGrades = useCallback(async () => {
    try {
      const [r, s] = await Promise.all([
        fetch(`${apiUrl}/salary-grades/list`, { credentials: "include" }),
        fetch(`${apiUrl}/salary-grades`, { credentials: "include" }),
      ]);
      if (r.ok) setExistingGradeCodes(await r.json());
      if (s.ok) { const d = await s.json(); setAllSteps(Array.isArray(d) ? d : []); }
    } catch {}
  }, []);
  useEffect(() => { fetchExistingGrades(); }, [fetchExistingGrades]);
  const existingGradeNumbers = useMemo(
    () => new Set(existingGradeCodes.map(g => (g.grade_code ?? "").replace(/\D/g, ""))),
    [existingGradeCodes]
  );

  // Debounced search
  useEffect(() => {
    if (isFirstRankQuery.current) { isFirstRankQuery.current = false; return; }
    const t = setTimeout(() => fetchRanks(query), 300);
    return () => clearTimeout(t);
  }, [query, fetchRanks]);

  const bumpRankToTop = (rank) => {
    if (query.trim()) return;
    setGroups(prev => {
      const idx = prev.findIndex(g => g.rank === rank);
      if (idx <= 0) return prev;
      const next = [...prev];
      const [item] = next.splice(idx, 1);
      next.unshift(item);
      return next;
    });
  };

  // Same recency rule as Workforce positions: opening a grade records it, so the
  // server lists recently used grades first next time.
  const selectRank = (rank) => {
    if (selectedRank !== rank) {
      fetch(`${apiUrl}/salary-grades/rank/${encodeURIComponent(rank)}/search-hit`, { method: "POST", credentials: "include" }).catch(() => {});
    }
    setSelectedRank(rank);
    bumpRankToTop(rank);
  };

  const selectedGroup = groups.find(g => g.rank === selectedRank) ?? null;

  const saveAmount = async (id, val) => {
    setSaving(true);
    try {
      const res = await fetch(`${apiUrl}/salary-grades/${id}`, {
        method: "PATCH", credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ amount: parseFloat(val) || 0 }),
      });
      if (res.ok) {
        const updated = await res.json();
        setGroups(prev => prev.map(g => ({ ...g, steps: g.steps.map(s => s.id === id ? { ...s, ...updated } : s) })));
        fetchExistingGrades();
      }
    } finally { setSaving(false); setEditId(null); }
  };

  const deleteStep = async (id) => {
    if (!window.confirm("Delete this step?")) return;
    const r = await fetch(`${apiUrl}/salary-grades/${id}`, { method: "DELETE", credentials: "include" });
    if (!r.ok) { const e = await r.json().catch(() => ({})); alert(e.error ?? "Failed to delete step."); return; }
    setGroups(prev => prev.map(g => ({ ...g, steps: g.steps.filter(s => s.id !== id) })).filter(g => g.steps.length > 0));
    fetchExistingGrades();
  };

  const deleteRank = async (rank) => {
    if (!window.confirm(`Delete entire rank "${rank}" and all its steps?`)) return;
    const r = await fetch(`${apiUrl}/salary-grades/rank/${encodeURIComponent(rank)}`, { method: "DELETE", credentials: "include" });
    if (!r.ok) { const e = await r.json().catch(() => ({})); alert(e.error ?? "Failed to delete rank."); return; }
    setGroups(prev => prev.filter(g => g.rank !== rank));
    if (selectedRank === rank) setSelectedRank(null);
    fetchExistingGrades();
  };

  const createRank = async () => {
    if (!newRank.trim()) return;
    if (existingGradeNumbers.has(newRank.trim())) { alert("SG " + newRank.trim() + " already exists."); return; }
    const stepCount = Number(newStepCount) || 1;
    const hasAmount = Array.from({ length: stepCount }, (_, i) => i + 1).some(s => parseFloat(newAmounts[s]) > 0);
    if (!hasAmount) { alert("Please enter at least one step amount before creating the rank."); return; }
    const gradeCode = `SG ${newRank.trim()}`;
    setSaving(true);
    try {
      for (let s = 1; s <= stepCount; s++) {
        await fetch(`${apiUrl}/salary-grades`, {
          method: "POST", credentials: "include",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ salary_grade: gradeCode, step: s, amount: parseFloat(newAmounts[s]) || 0 }),
        });
      }
      await fetchRanks(query);
      await fetchExistingGrades();
      setSelectedRank(gradeCode);
      setAddingRank(false);
      setNewRank(""); setNewStepCount(8); setNewAmounts({});
    } finally { setSaving(false); }
  };

  // The list keeps the server's order (most recently used first), like Workforce positions.
  const totalGrades = existingGradeCodes.length || groups.length;

  const RankList = (
    <div>
      {loading ? (
        <WsSpinner />
      ) : groups.length === 0 ? (
        <div className="ws-empty">No salary grades found.</div>
      ) : groups.map(g => {
        const gradeLabel = "SG " + g.rank.replace(/^sg\s*/i, "").trim();
        const minAmt = g.steps[0]?.amount;
        const maxAmt = g.steps[g.steps.length - 1]?.amount;
        const effYear = g.steps[0]?.effective_date ? new Date(g.steps[0].effective_date).getUTCFullYear() : null;
        const isSelected = selectedRank === g.rank;
        return (
          <div key={g.rank} role="button" tabIndex={0}
            className={`ws-item${isSelected ? " is-on" : ""}`}
            onKeyDown={e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); selectRank(g.rank); } }}
            onClick={() => selectRank(g.rank)}
          >
            <div className="ws-item-main">
              <div className="ws-item-title">{gradeLabel}</div>
              <div className="ws-item-sub">{g.steps.length} steps · {fmt(minAmt)} – {fmt(maxAmt)}</div>
            </div>
            {effYear && <span className="ws-pill" title="Effective year">{effYear}</span>}
          </div>
        );
      })}
      {!loading && !query.trim() && totalGrades > groups.length && (
        <div className="ws-list-note">Showing your {groups.length} most recently used of {totalGrades} grades. Search to find the rest.</div>
      )}
    </div>
  );

  const amounts    = allSteps.map(s => Number(s.amount)).filter(n => n > 0);
  const lowestPay  = amounts.length ? Math.min(...amounts) : null;
  const highestPay = amounts.length ? Math.max(...amounts) : null;
  const sgNums     = existingGradeCodes.map(g => parseFloat((g.grade_code ?? "").replace(/\D/g, ""))).filter(Boolean);
  const sgRangeLabel = sgNums.length ? `SG${Math.min(...sgNums)} – SG${Math.max(...sgNums)}` : "—";
  const stepCounts = Object.values(allSteps.reduce((acc, s) => ({ ...acc, [s.salary_grade]: (acc[s.salary_grade] ?? 0) + 1 }), {}));
  const stepsPerGrade = stepCounts.length ? Math.max(...stepCounts) : null;

  const Detail = selectedGroup ? (
    <div style={{ height: "100%", overflow: "auto" }}>
      <div className="ws-detail-head">
        <div>
          <div className="ws-kicker">Salary grade</div>
          <div className="ws-dtitle">
            {/^sg/i.test(selectedGroup.rank) ? selectedGroup.rank.toUpperCase() : `SG ${selectedGroup.rank}`}
          </div>
          <div className="ws-item-sub">
            {selectedGroup.steps.length} steps · effective {selectedGroup.steps[0]?.effective_date ? fmtDate(selectedGroup.steps[0].effective_date) : "—"}
          </div>
        </div>
        <div className="ws-actions">
          <button type="button" className="um-btn um-btn--danger" onClick={() => deleteRank(selectedGroup.rank)}>
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/></svg>
            Delete grade
          </button>
        </div>
      </div>
      <div style={{ overflowX: "auto" }}>
        <table className="ws-table">
          <thead>
            <tr>
              <th>Step</th>
              <th>Monthly amount</th>
              <th />
            </tr>
          </thead>
          <tbody>
            {selectedGroup.steps.map(step => (
              <tr key={step.id}>
                <td style={{ fontWeight: 700, color: "var(--psa-navy)", whiteSpace: "nowrap" }}>Step {step.step}</td>
                <td>
                  {editId === step.id ? (
                    <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" }}>
                      <input type="number" autoFocus value={editVal}
                        onChange={e => setEditVal(e.target.value)}
                        onKeyDown={e => { if (e.key === "Enter") saveAmount(step.id, editVal); if (e.key === "Escape") setEditId(null); }}
                        style={{ padding: "6px 10px", fontSize: "0.86rem", width: isMobile ? 120 : 150, borderRadius: 8 }} />
                      <button type="button" className="btn-primary ws-btn-sm" onClick={() => saveAmount(step.id, editVal)} disabled={saving}>Save</button>
                      <button type="button" className="um-btn um-btn--neutral um-btn--sm" onClick={() => setEditId(null)}>Cancel</button>
                    </div>
                  ) : (
                    <span style={{ fontWeight: 600, color: "var(--psa-navy)", cursor: "pointer" }} title="Click to edit" onClick={() => { setEditId(step.id); setEditVal(step.amount ?? ""); }}>{fmt(step.amount)}</span>
                  )}
                </td>
                <td style={{ textAlign: "right" }}>
                  {editId !== step.id && (
                    <div style={{ display: "inline-flex", gap: 6 }}>
                      <button type="button" className="um-btn um-btn--neutral um-btn--sm" onClick={() => { setEditId(step.id); setEditVal(step.amount ?? ""); }}>Edit</button>
                      <button type="button" className="ws-icon-btn ws-icon-btn--danger ws-icon-btn--sm" title="Delete step" aria-label={`Delete step ${step.step}`} onClick={() => deleteStep(step.id)}>
                        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/></svg>
                      </button>
                    </div>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  ) : (
    <div className="ws-overview">
      <div className="ws-kicker">Salary grades overview</div>
      <div className="ws-stats">
        <div className="ws-stat">
          <div className="ws-stat-label">Grades</div>
          <div className="ws-stat-value">{totalGrades}</div>
        </div>
        <div className="ws-stat">
          <div className="ws-stat-label">Grade range</div>
          <div className="ws-stat-value ws-stat-value--sm">{sgRangeLabel}</div>
        </div>
        <div className="ws-stat">
          <div className="ws-stat-label">Steps per grade</div>
          <div className="ws-stat-value ws-stat-value--ok">{stepsPerGrade ?? "—"}</div>
        </div>
        <div className="ws-stat ws-stat--wide">
          <div className="ws-stat-label">Monthly pay range</div>
          <div className="ws-stat-value ws-stat-value--sm">{lowestPay != null ? `${fmt(lowestPay)} – ${fmt(highestPay)}` : "—"}</div>
        </div>
      </div>
      <div className="ws-hint">
        <span>Select a grade on the left to view or edit its step amounts.</span>
        <button type="button" className="btn-primary ws-btn-sm" onClick={() => { fetchExistingGrades(); setAddingRank(true); }}>+ Add grade</button>
      </div>
    </div>
  );

  return (
    <>
      <div className="page-header">
        <div className="page-header-title">Salary Rank</div>
        <div className="page-header-sub">Manage salary grades, steps, and monthly amounts</div>
      </div>

      {/* Add Rank modal */}
      {addingRank && (() => {
        const count = Math.min(20, Math.max(1, Number(newStepCount) || 1));
        const stepNums = Array.from({ length: count }, (_, i) => i + 1);
        const vals = stepNums.map(s => parseFloat(newAmounts[s]));
        const entered = vals.filter(v => v > 0);
        const allFilled = entered.length === count;
        // Step amounts normally rise; flag a drop as a likely typo (not blocking).
        const dropAt = vals.findIndex((v, i) => i > 0 && v > 0 && vals[i - 1] > 0 && v < vals[i - 1]);
        const setCount = n => setNewStepCount(Math.min(20, Math.max(1, n)));
        const footNote = !newRank.trim() ? "Pick a salary grade to continue."
          : !allFilled ? `Enter an amount for every step (${count - entered.length} left).`
          : `Ready to add SG ${newRank} with ${count} steps.`;
        return (
          <div className="ar-backdrop" role="dialog" aria-modal="true" aria-label="Add salary grade">
            <div className="ar-panel">
              <div className="ar-head">
                <div>
                  <div className="ar-title">Add salary grade</div>
                  <div className="um-sub">Choose the grade, set its number of steps, then enter each monthly amount.</div>
                </div>
                <button type="button" className="ar-close" onClick={() => setAddingRank(false)} disabled={saving} aria-label="Close">✕</button>
              </div>

              <div className="ar-top">
                <div className="ar-field">
                  <div className="ar-label">Salary grade</div>
                  <RankGradeDropdown value={newRank} onChange={setNewRank} existingGradeNumbers={existingGradeNumbers} />
                </div>
                <div className="ar-field">
                  <div className="ar-label">Steps</div>
                  <div className="ar-stepper">
                    <button type="button" onClick={() => setCount(count - 1)} disabled={count <= 1} aria-label="Fewer steps">−</button>
                    <span aria-live="polite">{count}</span>
                    <button type="button" onClick={() => setCount(count + 1)} disabled={count >= 20} aria-label="More steps">+</button>
                  </div>
                </div>
              </div>

              <div className="ar-body">
                <div className="ar-label-row">
                  <span className="ar-label">Monthly amount per step</span>
                  <span className={`ws-pill${allFilled ? " ws-pill--ok" : ""}`}>{entered.length} of {count} filled</span>
                </div>
                <div className="ar-steps">
                  {stepNums.map(s => (
                    <label key={s} className="ar-step">
                      <span className="ar-step-n">Step {s}</span>
                      <span className="ar-money">
                        <span aria-hidden="true">₱</span>
                        <input type="number" inputMode="decimal" min="0" step="0.01" placeholder="0.00"
                          value={newAmounts[s] ?? ""} onChange={e => setNewAmounts(p => ({ ...p, [s]: e.target.value }))}
                          aria-label={`Step ${s} monthly amount`} />
                      </span>
                    </label>
                  ))}
                </div>
                {dropAt > 0 && (
                  <div className="ar-warn">Step {dropAt + 1} is lower than step {dropAt}. Double-check the amounts.</div>
                )}
                {entered.length > 0 && (
                  <div className="ar-range">Range: <strong>{fmt(Math.min(...entered))} – {fmt(Math.max(...entered))}</strong></div>
                )}
              </div>

              <div className="ar-foot">
                <span className="um-sub">{footNote}</span>
                <div className="ar-foot-actions">
                  <button type="button" className="um-btn um-btn--neutral" onClick={() => setAddingRank(false)} disabled={saving}>Cancel</button>
                  <button type="button" className="btn-primary ws-btn-sm" onClick={createRank} disabled={saving || !newRank.trim() || !allFilled}>
                    {saving ? "Adding…" : "Add grade"}
                  </button>
                </div>
              </div>
            </div>
          </div>
        );
      })()}

      <div className="ws-frame-body ws-frame-body--solo">
        {(!isMobile || !selectedRank) && (
          <div className="ws-toolbar">
            <WsSearch value={query} onChange={setQuery} placeholder="Search by salary grade…" />
            <button type="button" className="btn-primary ws-add" onClick={() => { fetchExistingGrades(); setAddingRank(true); }}>+ Add grade</button>
          </div>
        )}

        {isMobile ? (
          <div style={{ padding: 16 }}>
            {!selectedRank ? RankList : (
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                <button className="ghost-button" style={{ alignSelf: "flex-start", display: "inline-flex", alignItems: "center", gap: 6, padding: "6px 14px", fontSize: "0.82rem", fontWeight: 600, borderRadius: "var(--radius)", border: "1.5px solid var(--border-2)" }} onClick={() => setSelectedRank(null)}>
                  <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="15 18 9 12 15 6"/></svg>
                  Back
                </button>
                {Detail}
              </div>
            )}
          </div>
        ) : (
          <div className="ws-split">
            <div className="ws-list">{RankList}</div>
            <div className="ws-detail">{Detail}</div>
          </div>
        )}
      </div>
    </>
  );
}

// ── TypeDropdown — minimal custom select used inside AdjustmentsCard ──────────

function TypeDropdown({ groups, selKey, onSelect }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  useEffect(() => {
    if (!open) return;
    const handler = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [open]);
  const selLabel = groups.find(g => g.key === selKey)?.label ?? groups[0]?.label ?? "";
  return (
    <div style={{ padding: "6px 14px", borderBottom: "1px solid var(--border)", display: "flex", alignItems: "center", gap: 8 }}>
      <span style={{ fontSize: "0.6rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.1em", flexShrink: 0 }}>Type</span>
      <div ref={ref} style={{ position: "relative" }}>
        <button
          onClick={() => setOpen(v => !v)}
          style={{
            display: "flex", alignItems: "center", gap: 6,
            padding: "6px 14px", fontSize: "0.88rem", fontWeight: 500,
            borderRadius: 6, border: "1px solid var(--border)",
            background: "var(--surface)", color: "var(--text)",
            cursor: "pointer", outline: "none", whiteSpace: "nowrap",
          }}
        >
          {selLabel}
          <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"
            style={{ color: "var(--text-muted)", flexShrink: 0, transition: "transform 0.15s", transform: open ? "rotate(180deg)" : "none" }}>
            <polyline points="6 9 12 15 18 9"/>
          </svg>
        </button>
        {open && (
          <div style={{
            position: "absolute", top: "calc(100% + 4px)", left: 0, zIndex: 50,
            background: "var(--surface)", border: "1px solid var(--border)",
            borderRadius: 8, boxShadow: "0 4px 16px rgba(0,0,0,0.10)",
            minWidth: 190, overflow: "hidden",
          }}>
            {groups.map(g => {
              const active = g.key === selKey;
              return (
                <button key={g.key}
                  onClick={() => { onSelect(g.key); setOpen(false); }}
                  style={{
                    display: "block", width: "100%", textAlign: "left",
                    padding: "8px 14px", fontSize: "0.82rem",
                    fontWeight: active ? 600 : 400,
                    color: active ? "var(--psa-navy)" : "var(--text)",
                    background: active ? "var(--surface-2)" : "transparent",
                    border: "none", cursor: "pointer",
                  }}
                  onMouseEnter={e => { if (!active) e.currentTarget.style.background = "var(--surface-2)"; }}
                  onMouseLeave={e => { if (!active) e.currentTarget.style.background = "transparent"; }}
                >{g.label}</button>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}

// ── AdjustmentsCard ───────────────────────────────────────────────────────────

function AdjustmentsCard({ emp }) {
  const { dedTypes, earnTypes } = useContext(TypesContext);
  const [deds, setDeds]       = useState({});
  const [logs, setLogs]       = useState([]);
  const [loading, setLoading] = useState(false);
  const [section, setSection] = useState("Earnings");
  const [selKey, setSelKey]   = useState("");
  const [rowPage, setRowPage] = useState(0);

  useEffect(() => {
    if (!emp?.id || !emp?.name) return;
    setLoading(true);
    Promise.all([
      fetch(`${apiUrl}/employees/${emp.id}/deductions`, { credentials: "include" })
        .then(r => r.ok ? r.json() : {}).catch(() => ({})),
      fetch(`${apiUrl}/audit-log?entity=employee_deduction&q=${encodeURIComponent(emp.name)}&limit=50`, { credentials: "include" })
        .then(r => r.ok ? r.json() : []).catch(() => []),
    ]).then(([dedsData, logsData]) => {
      setDeds(dedsData && typeof dedsData === "object" ? dedsData : {});
      setLogs(logsData);
      setSelKey("");
      setRowPage(0);
    }).finally(() => setLoading(false));
  }, [emp?.id]);

  const fmt       = (n) => n == null ? "—" : `₱${Number(n).toLocaleString("en-PH", { minimumFractionDigits: 2 })}`;
  const fmtWhen   = (d) => { if (!d) return "—"; const dt = new Date(d); return isNaN(dt) ? String(d) : dt.toLocaleString("en-PH", { year: "numeric", month: "short", day: "2-digit", hour: "2-digit", minute: "2-digit" }); };
  const safeParse = (v) => { if (v == null) return null; if (typeof v === "object") return v; try { return JSON.parse(v); } catch { return null; } };

  const LOAN_ACTIONS = new Set(["create_loan", "extend_loan", "loan_deduction"]);
  const SECTION_COLOR = { Loans: "#0038a8", Earnings: "#166534", Deductions: "#374151" };

  const sideByName = useMemo(() => {
    const map = {};
    // Keys stored lowercase — looked up with .toLowerCase() to survive display_name renames
    (earnTypes ?? []).forEach(t => {
      if (t.display_name) map[t.display_name.toLowerCase()] = 'gross';
      if (t.name)         map[t.name.toLowerCase()]         = 'gross';
    });
    (dedTypes ?? []).forEach(t => {
      const side = t.side ?? 'deduction';
      if (t.display_name) map[t.display_name.toLowerCase()] = side;
      if (t.name)         map[t.name.toLowerCase()]         = side;
    });
    return map;
  }, [dedTypes, earnTypes]);

  const getLabel = (log) => {
    const o = safeParse(log.old_value), n = safeParse(log.new_value);
    return n?.payroll_type || o?.payroll_type || n?.loan_type || o?.loan_type || "—";
  };

  const getDesc = (log) => {
    const o = safeParse(log.old_value), n = safeParse(log.new_value);
    switch (log.action) {
      case "create_loan":      return `${fmt(n?.amount)} / ${fmt(n?.total_amount)}`;
      case "extend_loan":      return `+${fmt(n?.added_amount)} / ${fmt(n?.new_total)}`;
      case "loan_deduction":   return `${fmt(n?.deducted)} / ${fmt(n?.remaining_balance)}`;
      case "create_deduction": return fmt(n?.amount);
      case "update_deduction": return fmt(n?.amount);
      case "remove_deduction": return fmt(o?.amount);
      default: return "—";
    }
  };

  // Build groups keyed by section
  const allGroups = useMemo(() => {
    const map = {}, order = [];
    logs.forEach(log => {
      const label   = getLabel(log);
      const isLoan  = LOAN_ACTIONS.has(log.action);
      const sec     = isLoan ? "Loans" : sideByName[label?.toLowerCase()] === "gross" ? "Earnings" : "Deductions";
      const key     = `${sec}__${label}`;
      if (!map[key]) { map[key] = { key, label, section: sec, entries: [] }; order.push(key); }
      map[key].entries.push(log);
    });
    // Supplement Earnings and Deductions with any type that has a ledger value but no audit_log entry
    const makeEntry = (name, label, sec, amount) => ({
      id: `synth_${name}`, action: "create_deduction", changed_by: "system", changed_at: null,
      new_value: JSON.stringify({ employee_name: emp?.name, payroll_type: label, amount }),
      old_value: null, _synthetic: true,
    });
    (earnTypes ?? []).forEach(t => {
      const label = t.display_name || t.name;
      const key   = `Earnings__${label}`;
      if (map[key]) return;
      const amount = deds[t.name] != null
        ? Number(deds[t.name])
        : (t.computation_mode === 'fixed' && t.amount_fixed != null ? Number(t.amount_fixed) : null);
      if (amount == null || amount === 0) return;
      map[key] = { key, label, section: "Earnings", entries: [makeEntry(t.name, label, "Earnings", amount)] };
      order.push(key);
    });
    (dedTypes ?? []).forEach(t => {
      const label = t.display_name || t.name;
      const sec   = t.side === 'less_deduction' ? "Deductions" : "Deductions";
      const key   = `Deductions__${label}`;
      if (map[key]) return;
      const amount = deds[t.name] != null ? Number(deds[t.name]) : null;
      if (amount == null || amount === 0) return;
      map[key] = { key, label, section: "Deductions", entries: [makeEntry(t.name, label, sec, amount)] };
      order.push(key);
    });
    return order.map(k => map[k]);
  }, [logs, sideByName, earnTypes, deds, emp?.name]);

  const sectionGroups = useMemo(() => {
    const sg = { Loans: [], Earnings: [], Deductions: [] };
    allGroups.forEach(g => { if (sg[g.section]) sg[g.section].push(g); });
    return sg;
  }, [allGroups]);

  // Auto-select first available section and type on load (Loans removed — shown in Loans tab)
  useEffect(() => {
    for (const s of ["Earnings", "Deductions"]) {
      if (sectionGroups[s]?.length > 0) {
        setSection(s);
        setSelKey(sectionGroups[s][0].key);
        setRowPage(0);
        return;
      }
    }
  }, [allGroups]);

  // When user switches section, auto-select its first type
  const handleSection = (s) => {
    setSection(s);
    setSelKey(sectionGroups[s]?.[0]?.key ?? "");
    setRowPage(0);
  };

  const curGroups  = sectionGroups[section] ?? [];
  const selGroup   = curGroups.find(g => g.key === selKey) ?? curGroups[0] ?? null;
  const rows       = selGroup?.entries ?? [];
  const PAGE       = 5;
  const totalPages = Math.ceil(rows.length / PAGE) || 1;
  const safePg     = Math.min(rowPage, totalPages - 1);
  const pageRows   = rows.slice(safePg * PAGE, (safePg + 1) * PAGE);
  const sectionCol = SECTION_COLOR[section] ?? "#374151";

  if (loading) return (
    <div className="panel" style={{ padding: "80px 24px", textAlign: "center", color: "var(--text-muted)" }}>
      <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ animation: "spin 1s linear infinite", display: "block", margin: "0 auto 12px" }}>
        <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
      </svg>
      <div style={{ fontSize: "0.82rem" }}>Loading…</div>
    </div>
  );

  return (
    <div className="panel" style={{ padding: 0, overflow: "hidden" }}>
      {/* Header */}
      <div style={{ padding: "14px 20px", borderBottom: "1px solid var(--border)" }}>
        <div style={{ fontSize: "0.62rem", fontWeight: 800, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: 2 }}>Payment History</div>
        <div style={{ fontWeight: 800, fontSize: "1rem" }}>{emp.name}</div>
      </div>

      {allGroups.length === 0 ? (
        <div style={{ padding: "40px 18px", textAlign: "center", color: "var(--text-muted)" }}>
          <div style={{ fontWeight: 600, fontSize: "0.95rem", marginBottom: 4 }}>No changes logged yet</div>
          <div style={{ fontSize: "0.82rem" }}>Deduction edits for this employee will appear here.</div>
        </div>
      ) : (
        <>
          {/* Section pills — Earnings / Deductions only (Loans moved to dedicated Loans tab) */}
          <div style={{ padding: "8px 14px", borderBottom: "1px solid var(--border)", background: "var(--surface-2)", display: "flex", gap: 4, alignItems: "center", flexWrap: "wrap" }}>
            {["Earnings", "Deductions"].map(s => {
              const count = sectionGroups[s]?.length ?? 0;
              if (count === 0) return null;
              const active = section === s;
              return (
                <button key={s} onClick={() => handleSection(s)} style={{
                  padding: "4px 12px", fontSize: "0.74rem", fontWeight: 600, borderRadius: 999, cursor: "pointer", border: "none", outline: "none",
                  color: active ? "#fff" : "var(--text-muted)",
                  background: active ? (SECTION_COLOR[s] ?? "var(--psa-navy)") : "transparent",
                  boxShadow: active ? "none" : "inset 0 0 0 1px var(--border)",
                }}>{s}</button>
              );
            })}
          </div>

          {/* Type dropdown — only if current section has >1 type */}
          {curGroups.length > 1 && (
            <TypeDropdown
              groups={curGroups}
              selKey={selKey}
              onSelect={(key) => { setSelKey(key); setRowPage(0); }}
            />
          )}

          {/* Single type label when no dropdown needed */}
          {curGroups.length === 1 && selGroup && (
            <div style={{ padding: "9px 14px", borderBottom: "1px solid var(--border)", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
              <span style={{ fontSize: "0.85rem", fontWeight: 600, color: "var(--text)" }}>{selGroup.label}</span>
              <span style={{ fontSize: "0.68rem", fontWeight: 700, padding: "2px 9px", borderRadius: 999, background: sectionCol + "18", color: sectionCol }}>
                {rows.length} {rows.length === 1 ? "entry" : "entries"}
              </span>
            </div>
          )}

          {/* Table */}
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.82rem" }}>
              <thead>
                <tr style={{ background: "var(--surface-2)" }}>
                  {["Amount", "Date", "By"].map(h => (
                    <th key={h} style={{ padding: "7px 14px", textAlign: "left", fontSize: "0.68rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", whiteSpace: "nowrap" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {pageRows.length === 0 ? (
                  <tr><td colSpan={3} style={{ padding: "16px 14px", textAlign: "center", color: "var(--text-muted)", fontSize: "0.82rem" }}>No entries</td></tr>
                ) : pageRows.map((log) => (
                  <tr key={log.id} style={{ borderTop: "1px solid var(--border)" }}>
                    <td style={{ padding: "9px 14px", color: "var(--text)" }}>{getDesc(log)}</td>
                    <td style={{ padding: "9px 14px", whiteSpace: "nowrap", color: "var(--text-muted)", fontSize: "0.78rem" }}>{log._synthetic ? "Current" : fmtWhen(log.changed_at)}</td>
                    <td style={{ padding: "9px 14px", color: "var(--text-muted)", fontSize: "0.78rem", whiteSpace: "nowrap" }}>{log.changed_by}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {totalPages > 1 && (
            <div style={{ display: "flex", alignItems: "center", justifyContent: "flex-end", gap: 8, padding: "10px 14px", borderTop: "1px solid var(--border)" }}>
              <button onClick={() => setRowPage(p => Math.max(0, p - 1))} disabled={safePg === 0}
                style={{ background: "none", border: "1px solid var(--border)", borderRadius: 6, padding: "5px 14px", fontSize: "0.9rem", cursor: safePg === 0 ? "default" : "pointer", color: safePg === 0 ? "var(--text-muted)" : "var(--psa-navy)", fontWeight: 700 }}>‹</button>
              <span style={{ fontSize: "0.75rem", color: "var(--text-muted)", minWidth: 36, textAlign: "center" }}>{safePg + 1} / {totalPages}</span>
              <button onClick={() => setRowPage(p => Math.min(totalPages - 1, p + 1))} disabled={safePg === totalPages - 1}
                style={{ background: "none", border: "1px solid var(--border)", borderRadius: 6, padding: "5px 14px", fontSize: "0.9rem", cursor: safePg === totalPages - 1 ? "default" : "pointer", color: safePg === totalPages - 1 ? "var(--text-muted)" : "var(--psa-navy)", fontWeight: 700 }}>›</button>
            </div>
          )}
        </>
      )}
    </div>
  );
}

// ── AdjustmentsView ───────────────────────────────────────────────────────────

function AdjustmentsView({ isAdmin, refreshKey }) {
  const [employees, setEmployees] = useState([]);
  const [query, setQuery]         = useState("");
  const [selectedId, setSelectedId] = useState(null);
  const [openId, setOpenId]         = useState(null);
  const [loading, setLoading]       = useState(false);
  const isMobile = useIsMobile();
  const hasData = useRef(false);

  const fetchEmployees = useCallback(async (q = "", silent = false) => {
    if (!silent) setLoading(true);
    try {
      const res = await fetch(`${apiUrl}/employees?q=${encodeURIComponent(q)}&limit=5`, { credentials: "include" });
      if (res.ok) {
        const json = await res.json();
        setEmployees(Array.isArray(json) ? json : (json.data ?? []));
        hasData.current = true;
      }
    } finally { setLoading(false); }
  }, []);

  useEffect(() => { fetchEmployees("", hasData.current); }, [fetchEmployees, refreshKey]);

  useEffect(() => {
    const t = setTimeout(() => fetchEmployees(query), 300);
    return () => clearTimeout(t);
  }, [query, fetchEmployees]);

  const selectedEmp = employees.find(e => e.id === selectedId) ?? null;

  const bumpToTop = useCallback((empId) => {
    if (query.trim()) return;
    setEmployees(prev => {
      const idx = prev.findIndex(e => e.id === empId);
      if (idx <= 0) return prev;
      const next = [...prev];
      const [moved] = next.splice(idx, 1);
      return [moved, ...next];
    });
  }, [query]);

  const searchPanel = (
    <div className="panel" style={{ marginBottom: 10 }}>
      <div style={{ position: "relative" }}>
        <svg style={{ position: "absolute", left: 12, top: "50%", transform: "translateY(-50%)", pointerEvents: "none" }} width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#a8bcd4" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
          <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
        </svg>
        <input
          className="form-input"
          style={{ borderRadius: "var(--radius)", border: "1.5px solid var(--border-2)", padding: "10px 14px 10px 36px", fontSize: "0.9rem", width: "100%" }}
          placeholder="Search employees…"
          value={query}
          onChange={e => setQuery(e.target.value)}
        />
      </div>
    </div>
  );

  return (
    <>
      <div className="page-header">
        <div className="page-header-title">Adjustments</div>
        <div className="page-header-sub">History of changes made to each employee's deductions</div>
      </div>

      {isMobile ? (
        <div>
          {searchPanel}
          {loading ? (
            <div className="empty" style={{ padding: "20px 0" }}>Searching…</div>
          ) : employees.length === 0 ? (
            <div className="empty" style={{ padding: "20px 0" }}>No employees found.</div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {employees.map(emp => {
                const isOpen = openId === emp.id;
                return (
                  <div key={emp.id}>
                    <div
                      className="job-card"
                      style={{
                        cursor: "pointer",
                        background: isOpen ? "#e8f0fa" : undefined,
                        outline: isOpen ? "2px solid #1a4f8a" : "none",
                        outlineOffset: -2,
                        borderRadius: isOpen ? "var(--radius) var(--radius) 0 0" : undefined,
                      }}
                      onClick={() => {
                        const opening = openId !== emp.id;
                        setOpenId(v => v === emp.id ? null : emp.id);
                        if (opening) { bumpToTop(emp.id); fetch(`${apiUrl}/employees/${emp.id}/search-hit`, { method: "POST", credentials: "include" }).catch(() => {}); }
                      }}
                    >
                      <div style={{ padding: "14px 16px" }}>
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                          <div className="job-card__name">{emp.name}</div>
                          <svg width="13" height="13" viewBox="0 0 16 16" style={{ flexShrink: 0, transition: "transform 0.2s", transform: isOpen ? "rotate(180deg)" : "rotate(0deg)" }}>
                            <path d="M4 6l4 4 4-4" stroke="#637491" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" fill="none"/>
                          </svg>
                        </div>
                        <div style={{ marginTop: 3, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
                          <span style={{ fontSize: "0.78rem", color: "var(--text-muted)", flexShrink: 0 }}>{emp.employee_number}</span>
                          <span className="job-card__tag" title={emp.position} style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontSize: "0.74rem", minWidth: 0 }}>{emp.position}</span>
                        </div>
                      </div>
                    </div>
                    {isOpen && (
                      <div style={{ borderTop: "none", borderRadius: "0 0 var(--radius) var(--radius)", overflow: "hidden" }}>
                        <AdjustmentsCard emp={emp} />
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </div>
      ) : (
        <div className="deduction-master-detail">

          <div className="deduction-master-detail__list">
            {searchPanel}
            {loading ? (
              <div className="empty" style={{ padding: "20px 0" }}>Searching…</div>
            ) : employees.length === 0 ? (
              <div className="empty" style={{ padding: "20px 0" }}>No employees found.</div>
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                {employees.map(emp => (
                  <div
                    key={emp.id}
                    className={`job-card${selectedId === emp.id ? " job-card--selected" : ""}`}
                    style={{ cursor: "pointer", transition: "background 0.15s" }}
                    onClick={() => {
                      setSelectedId(emp.id);
                      bumpToTop(emp.id);
                      fetch(`${apiUrl}/employees/${emp.id}/search-hit`, { method: "POST", credentials: "include" }).catch(() => {});
                    }}
                  >
                    <div style={{ padding: "12px 16px" }}>
                      <div className="job-card__name" style={{ fontSize: "0.92rem" }}>{emp.name}</div>
                      <div style={{ marginTop: 3, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
                        <span style={{ fontSize: "0.78rem", color: "var(--text-muted)", flexShrink: 0 }}>{emp.employee_number}</span>
                        <span className="job-card__tag" title={emp.position} style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontSize: "0.74rem", minWidth: 0 }}>{emp.position}</span>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>

          <div className="deduction-master-detail__detail">
            {selectedEmp ? (
              <AdjustmentsCard emp={selectedEmp} />
            ) : (
              <div className="panel" style={{ height: "100%", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", textAlign: "center", color: "var(--text-muted)", padding: "40px" }}>
                <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" strokeLinejoin="round" style={{ opacity: 0.3, marginBottom: 16 }}>
                  <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/>
                </svg>
                <div style={{ fontSize: "1rem", fontWeight: 600, marginBottom: 6 }}>No employee selected</div>
                <div style={{ fontSize: "0.84rem" }}>Select an employee from the list to view their deductions.</div>
              </div>
            )}
          </div>

        </div>
      )}
    </>
  );
}

// ── ArchiveView ───────────────────────────────────────────────────────────────
const MONTH_NAMES = ["January","February","March","April","May","June","July","August","September","October","November","December"];

function ArchiveView({ isSuperAdmin, permissions = [] }) {
  const isMobile = useIsMobile();
  const canAny = (...keys) => isSuperAdmin || keys.some(k => permissions.includes(k));
  const can    = (key)     => isSuperAdmin || permissions.includes(key);
  const canView      = canAny("archive.view", "archive.manage");
  const canManage    = can("archive.manage");
  const canExport    = canAny("archive.export", "archive.manage");
  const canPayslip   = canAny("archive.payslip", "archive.manage");
  const canBreakdown = canAny("archive.breakdown", "archive.manage");
  const canAnnual    = canAny("archive.annual", "archive.manage");

  const { dedTypes, earnTypes } = useContext(TypesContext);

  const curYear = new Date().getFullYear();
  const [selectedYear, setSelectedYear]   = useState(curYear);
  const [archives, setArchives]           = useState([]);
  const [saving, setSaving]               = useState(null);
  const [exporting, setExporting]         = useState(null);
  const [exportingQSummary, setExportingQSummary]   = useState(false);
  const [selectedMonth, setSelectedMonth] = useState(new Date().getMonth() + 1);
  const [selectedQ, setSelectedQ]         = useState(null);
  const [slotSnapshot, setSlotSnapshot]   = useState(null);  // full snapshot {items, dedTypes, earnTypes}
  const [loadingItems, setLoadingItems]   = useState(false);
  const [showYearSummary, setShowYearSummary] = useState(false);
  const [yearSnapshots, setYearSnapshots]     = useState([]);
  const [loadingYear, setLoadingYear]         = useState(false);

  // ── Employee History tab ──
  const [archiveTab, setArchiveTab]           = useState("archive");
  const [histQuery, setHistQuery]             = useState("");
  const [histEmployees, setHistEmployees]     = useState([]);
  const [histEmpLoading, setHistEmpLoading]   = useState(false);
  const [histSelectedEmp, setHistSelectedEmp] = useState(null);
  const [histRecords, setHistRecords]         = useState([]);
  const [histLoading, setHistLoading]         = useState(false);
  const [histExpandedRow, setHistExpandedRow] = useState(null);
  const [histRecordsPage, setHistRecordsPage] = useState(0);
  const [histPage, setHistPage]               = useState(0);

  const loadArchives = useCallback(async () => {
    const res = await fetch(`${apiUrl}/archives`, { credentials: "include" });
    if (res.ok) setArchives(await res.json());
  }, []);

  const searchHistEmployees = useCallback(async (q) => {
    setHistEmpLoading(true);
    try {
      const limit = q.trim() ? 50 : 5;
      const res = await fetch(`${apiUrl}/employees/search-all?q=${encodeURIComponent(q)}&limit=${limit}`, { credentials: "include" });
      if (res.ok) setHistEmployees(await res.json());
    } finally { setHistEmpLoading(false); }
  }, [apiUrl]);

  const bumpHistToTop = (empId) => {
    if (histQuery.trim()) return;
    setHistEmployees(prev => {
      const idx = prev.findIndex(e => e.id === empId);
      if (idx <= 0) return prev;
      const updated = [...prev];
      const [emp] = updated.splice(idx, 1);
      updated.unshift(emp);
      return updated;
    });
  };

  useEffect(() => {
    if (archiveTab !== "history") return;
    setHistPage(0);
    const t = setTimeout(() => searchHistEmployees(histQuery), 300);
    return () => clearTimeout(t);
  }, [histQuery, archiveTab, searchHistEmployees]);

  useEffect(() => {
    if (archiveTab === "history" && histEmployees.length === 0 && !histEmpLoading) {
      searchHistEmployees("");
    }
  }, [archiveTab]); // eslint-disable-line

  async function selectHistEmployee(emp) {
    bumpHistToTop(emp.id);
    fetch(`${apiUrl}/employees/${emp.id}/search-hit`, { method: "POST", credentials: "include" }).catch(() => {});
    setHistSelectedEmp(emp);
    setHistRecords([]);
    setHistExpandedRow(null);
    setHistRecordsPage(0);
    setHistLoading(true);
    try {
      const res = await fetch(`${apiUrl}/employees/${emp.id}/payroll-history`, { credentials: "include" });
      if (res.ok) setHistRecords(await res.json());
    } finally { setHistLoading(false); }
  }

  useEffect(() => { if (canView) loadArchives(); }, [canView, loadArchives]);

  useEffect(() => {
    const saved = archives.filter(a => Number(a.year) === selectedYear).map(a => Number(a.month));
    if (saved.length > 0) setSelectedMonth(Math.min(...saved));
    else setSelectedMonth(new Date().getMonth() + 1);
    setSelectedQ(null);
    setSlotSnapshot(null);
    setShowYearSummary(false);
    setYearSnapshots([]);
  }, [archives, selectedYear]);

  useEffect(() => {
    if (selectedMonth === null) return;
    setSelectedQ(null);
    setSlotSnapshot(null);
    setShowYearSummary(false);
  }, [selectedMonth]);

  async function selectQuincena(q) {
    setSelectedQ(q);
    setSlotSnapshot(null);
    const slot = getSlot(selectedMonth, q);
    if (!slot) return;
    setLoadingItems(true);
    try {
      const res = await fetch(`${apiUrl}/archives/${slot.year}/${slot.month}/${slot.quincena}`, { credentials: "include" });
      if (res.ok) {
        const arc = await res.json();
        setSlotSnapshot(arc.snapshot ?? null);
      }
    } finally { setLoadingItems(false); }
  }

  async function loadYearSummary() {
    setShowYearSummary(true);
    setSelectedMonth(null);
    setSelectedQ(null);
    setSlotSnapshot(null);
    setYearSnapshots([]);
    setLoadingYear(true);
    try {
      const slots = archives.filter(a => Number(a.year) === selectedYear);
      const results = await Promise.all(slots.map(async slot => {
        const res = await fetch(`${apiUrl}/archives/${slot.year}/${slot.month}/${slot.quincena}`, { credentials: "include" });
        if (!res.ok) return null;
        const arc = await res.json();
        return { month: Number(slot.month), quincena: Number(slot.quincena), snapshot: arc.snapshot };
      }));
      setYearSnapshots(results.filter(Boolean));
    } finally { setLoadingYear(false); }
  }

  const getSlot = (month, quincena) =>
    archives.find(a => Number(a.year) === selectedYear && Number(a.month) === month && Number(a.quincena) === quincena) ?? null;

  async function handleSave(month, quincena) {
    setSaving({ month, quincena });
    try {
      const [itemsRes, sigRes] = await Promise.all([
        fetch(`${apiUrl}/employees/live-payslips?limit=9999`, { credentials: "include" }),
        fetch(`${apiUrl}/settings/signatories`, { credentials: "include" }),
      ]);
      const items      = itemsRes.ok ? await itemsRes.json() : [];
      const signatories = sigRes.ok ? await sigRes.json() : null;

      const lastDay = new Date(selectedYear, month, 0).getDate();
      const period = {
        period_start: new Date(selectedYear, month - 1, quincena === 1 ? 1  : 16).toISOString(),
        period_end:   new Date(selectedYear, month - 1, quincena === 1 ? 15 : lastDay).toISOString(),
      };

      const res = await fetch(`${apiUrl}/archives/${selectedYear}/${month}/${quincena}`, {
        method: "PUT", credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ items, period, signatories, dedTypes, earnTypes }),
      });
      if (!res.ok) { const d = await res.json(); alert(d.error ?? "Failed to save."); return; }
      await loadArchives();
    } finally { setSaving(null); }
  }

  async function handleExport(slot) {
    setExporting(slot.id);
    try {
      const res = await fetch(`${apiUrl}/archives/${slot.year}/${slot.month}/${slot.quincena}`, { credentials: "include" });
      if (!res.ok) { alert("Failed to load archive."); return; }
      const arc = await res.json();
      const { items, period, signatories, dedTypes: aDed, earnTypes: aEarn } = arc.snapshot;
      const lessDed = aDed?.length > 0 ? buildCatalogFromTypes(aDed, 'less_deduction') : LESS_DEDUCTIONS_CATALOG;
      const ded     = aDed?.length > 0 ? buildCatalogFromTypes(aDed, 'deduction')      : DEDUCTIONS_CATALOG;
      const blob = await generatePayslipsXlsx(items, period, signatories, lessDed, ded, aEarn ?? []);
      const mn = MONTH_NAMES[Number(slot.month) - 1];
      downloadBlob(blob, `Payroll Archive - ${mn} ${slot.year} Q${slot.quincena}.xlsx`);
    } catch { alert("Failed to generate Excel."); }
    finally { setExporting(null); }
  }

  async function handleExportQuincenaSummary() {
    if (!slotSnapshot) return;
    setExportingQSummary(true);
    try {
      const items      = slotSnapshot.items ?? [];
      const sDedTypes  = slotSnapshot.dedTypes ?? [];
      const sEarnTypes = slotSnapshot.earnTypes ?? [];
      const sum = fn => items.reduce((s, i) => s + Number(fn(i) || 0), 0);

      const totalBasic        = sum(i => i.basic_pay);
      const totalPera         = sum(i => i.allowances?.pera ?? 0);
      const totalGross        = sum(i => i.gross_pay);
      const takeHomeThisQ     = selectedQ === 1 ? sum(i => i.first_quincena) : sum(i => i.second_quincena);
      const peraActive        = sEarnTypes.length === 0 || sEarnTypes.some(t => t.name === 'pera');
      const customEarns       = sEarnTypes.filter(t => t.name !== 'pera' && t.name !== 'basic_pay');
      const lessDedList       = sDedTypes.filter(t => t.side === 'less_deduction');
      const mainDedList       = sDedTypes.filter(t => t.side === 'deduction');
      const dedTotals = {};
      for (const t of sDedTypes) dedTotals[t.name] = sum(i => i.deductions?.[t.name] ?? 0);
      const totalLessDed      = lessDedList.reduce((s, t) => s + (dedTotals[t.name] ?? 0), 0);
      const totalSalaryEarned = totalGross - totalLessDed;
      const totalMainDed      = mainDedList.reduce((s, t) => s + (dedTotals[t.name] ?? 0), 0);

      const slot      = getSlot(selectedMonth, selectedQ);
      const mn        = MONTH_NAMES[selectedMonth - 1];
      const lastDay   = new Date(selectedYear, selectedMonth, 0).getDate();
      const periodStr = selectedQ === 1 ? `${mn} 1–15, ${selectedYear}` : `${mn} 16–${lastDay}, ${selectedYear}`;
      const qLabel    = selectedQ === 1 ? '1st Quincena' : '2nd Quincena';

      const wb     = new ExcelJS.Workbook();
      const navy   = { argb: 'FF103669' };
      const green  = { argb: 'FF166534' };
      const numFmt = '#,##0.00';

      const ws = wb.addWorksheet('Payroll Summary');
      ws.columns = [{ width: 38 }, { width: 16 }];

      const titleRow = ws.addRow([`${qLabel} — Payroll Summary`]);
      titleRow.getCell(1).font = { bold: true, size: 14, color: navy };
      ws.addRow([`Period: ${periodStr}  ·  ${items.length} Employees`]).getCell(1).font = { italic: true, color: { argb: 'FF888888' } };
      if (slot) ws.addRow([`Saved: ${new Date(slot.updated_at).toLocaleDateString('en-PH', { month: 'short', day: 'numeric', year: 'numeric' })} by ${slot.saved_by}`]).getCell(1).font = { italic: true, color: { argb: 'FF888888' } };
      ws.addRow([]);

      const addSection = label => {
        const r = ws.addRow([label]);
        r.getCell(1).font = { bold: true, size: 9, color: navy };
        r.height = 18;
      };
      const addItem = (label, val, opts = {}) => {
        const r = ws.addRow([label, Number(val || 0)]);
        if (opts.bold) {
          r.getCell(1).font = { bold: true, ...(opts.color ? { color: opts.color } : {}) };
          r.getCell(2).font = { bold: true, ...(opts.color ? { color: opts.color } : {}) };
        }
        if (opts.border) {
          r.getCell(1).border = { top: { style: 'thin' } };
          r.getCell(2).border = { top: { style: 'thin' } };
        }
        r.getCell(2).numFmt = numFmt;
        r.getCell(2).alignment = { horizontal: 'right' };
      };

      addSection('GROSS EARNINGS');
      addItem('Basic Salary', totalBasic);
      if (peraActive) addItem('PERA', totalPera);
      for (const t of customEarns) addItem(t.display_name || t.name, sum(i => i.allowances?.[t.name] ?? 0));
      addItem('Total Gross', totalGross, { bold: true, border: true });
      ws.addRow([]);

      addSection('LESS DEDUCTIONS');
      for (const t of lessDedList) addItem(t.display_name || t.name, dedTotals[t.name] ?? 0);
      addItem('Total Less Deductions', totalLessDed, { bold: true, border: true });
      ws.addRow([]);

      addSection('SALARY EARNED FOR THE PERIOD');
      addItem('Total Salary / Earned', totalSalaryEarned, { bold: true });
      ws.addRow([]);

      addSection('DEDUCTIONS');
      for (const t of mainDedList) addItem(t.display_name || t.name, dedTotals[t.name] ?? 0);
      addItem('Total Deductions', totalMainDed, { bold: true, border: true });
      ws.addRow([]);

      const thRow = ws.addRow(['TOTAL TAKE HOME PAY', Number(takeHomeThisQ)]);
      thRow.getCell(1).font = { bold: true, size: 12, color: green };
      thRow.getCell(2).font = { bold: true, size: 12, color: green };
      thRow.getCell(2).numFmt = numFmt;
      thRow.getCell(2).alignment = { horizontal: 'right' };

      const buffer = await wb.xlsx.writeBuffer();
      downloadBlob(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `Payroll Summary - ${mn} ${selectedYear} Q${selectedQ}.xlsx`);
    } catch { alert('Failed to generate Payroll Summary Excel.'); }
    finally { setExportingQSummary(false); }
  }

  async function handleDelete(slot) {
    const mn = MONTH_NAMES[Number(slot.month) - 1];
    if (!window.confirm(`Delete archive for ${mn} ${slot.year} — ${slot.quincena === 1 || slot.quincena === "1" ? "1st" : "2nd"} Quincena?\n\nThis cannot be undone.`)) return;
    await fetch(`${apiUrl}/archives/${slot.id}`, { method: "DELETE", credentials: "include" });
    await loadArchives();
  }

  const savedYears = [...new Set(archives.map(a => Number(a.year)))];
  const years = [...new Set([curYear, ...savedYears])].sort((a, b) => b - a);

  if (!canView) return (
    <div className="empty" style={{ padding: "60px 0" }}>You don't have permission to view the archive.</div>
  );

  const yearSaved = archives.filter(a => Number(a.year) === selectedYear).length;
  const mn = selectedMonth ? MONTH_NAMES[selectedMonth - 1] : "";

  return (
    <>
      <div className="page-header">
        <div className="page-header-title">Payroll Archive</div>
        <div className="page-header-sub">Browse, save, and export payroll records by period</div>
      </div>

      <WsTabs label="Archive sections" value={archiveTab} onChange={setArchiveTab}
        tabs={[{ id: "archive", label: isMobile ? "Archive" : "Payroll Archive" }, { id: "history", label: isMobile ? "History" : "Employee History" }]} />

      {/* ── Employee History tab ── */}
      {archiveTab === "history" && (() => {
        const empListPanel = (
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            <div className="panel search-form" style={{ marginBottom: 0 }}>
              <div className="search-combo-wrapper">
                <svg style={{ position: "absolute", left: 12, top: "50%", transform: "translateY(-50%)", pointerEvents: "none" }} width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#a8bcd4" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
                </svg>
                <input value={histQuery} onChange={e => setHistQuery(e.target.value)} placeholder="Search employees…"
                  className="form-input"
                  style={{ borderRadius: "var(--radius)", border: "1.5px solid var(--border-2)", padding: "10px 14px 10px 36px", fontSize: "0.95rem", width: "100%" }} />
              </div>
            </div>
            {histEmpLoading && <div style={{ padding: "14px", fontSize: "0.8rem", color: "var(--text-muted)", textAlign: "center" }}>Searching…</div>}
            {!histEmpLoading && histEmployees.length === 0 && (
              <div style={{ padding: "14px", fontSize: "0.8rem", color: "var(--text-muted)", textAlign: "center" }}>No employees found</div>
            )}
            {!histEmpLoading && histEmployees.length > 0 && (() => {
              const HIST_PAGE_SIZE = 12;
              const totalPages = Math.ceil(histEmployees.length / HIST_PAGE_SIZE);
              const pageEmps = histEmployees.slice(histPage * HIST_PAGE_SIZE, (histPage + 1) * HIST_PAGE_SIZE);
              return (
                <>
                  <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                    {pageEmps.map(emp => {
                      const isSelected = histSelectedEmp?.id === emp.id;
                      return (
                        <div key={emp.id}
                          className={`job-card${isSelected ? " job-card--selected" : ""}`}
                          style={{ cursor: "pointer", transition: "background 0.15s" }}
                          onClick={() => selectHistEmployee(emp)}>
                          <div style={{ padding: "12px 16px" }}>
                            <div className="job-card__name" style={{ fontSize: "0.92rem" }}>
                              {emp.name}
                              {emp.status === "inactive" && (
                                <span style={{ marginLeft: 6, background: isSelected ? "rgba(255,255,255,0.2)" : "#fee2e2", color: isSelected ? "#fff" : "#991b1b", borderRadius: 4, padding: "1px 5px", fontSize: "0.63rem", fontWeight: 700, verticalAlign: "middle" }}>Inactive</span>
                              )}
                            </div>
                            <div style={{ marginTop: 3, display: "flex", alignItems: "center", gap: 6 }}>
                              <span style={{ fontSize: "0.78rem", color: "var(--text-muted)", flexShrink: 0 }}>{emp.employee_number}</span>
                              <span className="job-card__tag" style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontSize: "0.74rem", minWidth: 0 }}>{emp.position || "—"}</span>
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                  {totalPages > 1 && (
                    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "8px 4px 2px", marginTop: 2 }}>
                      <button onClick={() => setHistPage(p => Math.max(0, p - 1))} disabled={histPage === 0}
                        style={{ padding: "4px 12px", fontSize: "0.78rem", borderRadius: 6, border: "1px solid var(--border)", background: "var(--surface)", color: histPage === 0 ? "var(--text-muted)" : "var(--text)", cursor: histPage === 0 ? "default" : "pointer" }}>← Prev</button>
                      <span style={{ fontSize: "0.75rem", color: "var(--text-muted)" }}>{histPage + 1} / {totalPages}</span>
                      <button onClick={() => setHistPage(p => Math.min(totalPages - 1, p + 1))} disabled={histPage === totalPages - 1}
                        style={{ padding: "4px 12px", fontSize: "0.78rem", borderRadius: 6, border: "1px solid var(--border)", background: "var(--surface)", color: histPage === totalPages - 1 ? "var(--text-muted)" : "var(--text)", cursor: histPage === totalPages - 1 ? "default" : "pointer" }}>Next →</button>
                    </div>
                  )}
                </>
              );
            })()}
          </div>
        );

        const HIST_PAGE_SIZE = 10;
        const histPageCount = Math.ceil(histRecords.length / HIST_PAGE_SIZE);
        const pagedHistRecords = histRecords.slice(histRecordsPage * HIST_PAGE_SIZE, (histRecordsPage + 1) * HIST_PAGE_SIZE);

        const historyRecordsPanel = histSelectedEmp && (
          <>
            {/* Back button on mobile */}
            {isMobile && (
              <button type="button" onClick={() => setHistSelectedEmp(null)}
                style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 12, padding: "6px 12px", borderRadius: 999, border: "1.5px solid var(--border)", background: "transparent", color: "var(--text-muted)", fontSize: "0.8rem", fontWeight: 600, cursor: "pointer" }}>
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="15 18 9 12 15 6"/></svg>
                All Employees
              </button>
            )}
            {histLoading && <div style={{ padding: "60px 0", textAlign: "center", color: "var(--text-muted)", fontSize: "0.85rem" }}>Loading history…</div>}
            {!histLoading && histRecords.length === 0 && (
              <div style={{ padding: "60px 0", textAlign: "center", color: "var(--text-muted)", fontSize: "0.85rem" }}>
                No payslip history found for <b>{histSelectedEmp.name}</b>.<br/>
                <span style={{ fontSize: "0.78rem" }}>History is recorded when an archive snapshot is saved.</span>
              </div>
            )}
            {!histLoading && histRecords.length > 0 && (
              <>
                <div style={{ marginBottom: 10, display: "flex", alignItems: "baseline", gap: 8, flexWrap: "wrap" }}>
                  <span style={{ fontWeight: 700, fontSize: "0.95rem", color: "var(--psa-navy)" }}>{histSelectedEmp.name}</span>
                  <span style={{ fontSize: "0.78rem", color: "var(--text-muted)" }}>{histRecords.length} record{histRecords.length !== 1 ? "s" : ""}</span>
                  {histSelectedEmp.status === "inactive" && (
                    <span style={{ background: "#fee2e2", color: "#991b1b", borderRadius: 4, padding: "1px 6px", fontSize: "0.68rem", fontWeight: 700 }}>Inactive Employee</span>
                  )}
                </div>

                {isMobile ? (
                  /* Mobile: tap-to-expand cards */
                  <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                    {pagedHistRecords.map((rec, idx) => {
                      const rowKey  = `${rec.year}-${rec.month}-${rec.quincena}`;
                      const isOpen  = histExpandedRow === rowKey;
                      const mName   = MONTH_NAMES[rec.month - 1];
                      const qLabel  = rec.quincena === 1 ? "1st Quincena" : "2nd Quincena";
                      const snap    = rec.snapshot;
                      const takeHome = rec.quincena === 1 ? snap?.first_quincena : snap?.second_quincena;
                      const deds    = snap?.deductions ?? {};
                      const allocs  = snap?.allowances ?? {};
                      const labels  = snap?.type_labels ?? {};
                      const toLabel = k => labels[k] || k.replace(/_/g, " ").replace(/\b\w/g, c => c.toUpperCase());
                      const dedKeys = Object.keys(deds).filter(k => Number(deds[k]) > 0);
                      const earnKeys = Object.keys(allocs).filter(k => k !== "pera" && k !== "basic_pay");
                      return (
                        <div key={rowKey} style={{ border: "1px solid var(--border)", borderRadius: "var(--radius)", background: "var(--surface)", overflow: "hidden" }}>
                          <div style={{ padding: "11px 14px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, cursor: snap ? "pointer" : "default", background: isOpen ? "#e8f0fa" : "var(--surface)" }}
                            onClick={() => snap && setHistExpandedRow(isOpen ? null : rowKey)}>
                            <div>
                              <div style={{ fontWeight: 700, fontSize: "0.88rem", color: "var(--psa-navy)" }}>{mName} {rec.year} <span style={{ fontWeight: 400, color: "var(--text-muted)", fontSize: "0.77rem" }}>{qLabel}</span></div>
                              <div style={{ marginTop: 3, display: "flex", gap: 10, flexWrap: "wrap" }}>
                                <span style={{ fontSize: "0.77rem", color: "var(--text-muted)" }}>Gross: <b style={{ color: "var(--text)" }}>₱ {fmt(rec.gross_pay)}</b></span>
                                <span style={{ fontSize: "0.77rem", color: "#991b1b" }}>Ded: <b>₱ {fmt(rec.total_deductions)}</b></span>
                                <span style={{ fontSize: "0.77rem", color: "#166534", fontWeight: 700 }}>Take Home: ₱ {fmt(takeHome ?? 0)}</span>
                              </div>
                            </div>
                            {snap && <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"
                              style={{ flexShrink: 0, transform: isOpen ? "rotate(180deg)" : "none", transition: "transform 0.15s" }}>
                              <polyline points="6 9 12 15 18 9"/>
                            </svg>}
                          </div>
                          {isOpen && snap && (
                            <div style={{ padding: "12px 14px", borderTop: "1px solid var(--border)", background: "var(--surface-2)", display: "flex", flexDirection: "column", gap: 14 }}>
                              <div>
                                <div style={{ fontSize: "0.67rem", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.07em", color: "var(--text-muted)", marginBottom: 6 }}>Earnings</div>
                                <BreakdownRow label="Basic Pay" value={snap.basic_pay ?? 0} />
                                {allocs.pera != null && allocs.pera > 0 && <BreakdownRow label={toLabel("pera")} value={allocs.pera} />}
                                {earnKeys.map(k => <BreakdownRow key={k} label={toLabel(k)} value={allocs[k]} />)}
                                <BreakdownRow label="Gross Pay" value={snap.gross_pay ?? 0} bold />
                              </div>
                              <div>
                                <div style={{ fontSize: "0.67rem", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.07em", color: "var(--text-muted)", marginBottom: 6 }}>Deductions</div>
                                {dedKeys.length === 0 && <div style={{ fontSize: "0.78rem", color: "var(--text-muted)" }}>None</div>}
                                {dedKeys.map(k => <BreakdownRow key={k} label={toLabel(k)} value={deds[k]} />)}
                                <BreakdownRow label="Total Deductions" value={rec.total_deductions} bold />
                              </div>
                              <div>
                                <div style={{ fontSize: "0.67rem", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.07em", color: "var(--text-muted)", marginBottom: 6 }}>Take Home</div>
                                <BreakdownRow label="Net Pay" value={snap.net_pay ?? 0} />
                                <BreakdownRow label={rec.quincena === 1 ? "1st Quincena" : "2nd Quincena"} value={takeHome ?? 0} bold />
                              </div>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                ) : (
                  /* Desktop: full table */
                  <div style={{ border: "1px solid var(--border)", borderRadius: "var(--radius)", overflow: "hidden", background: "var(--surface)" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.81rem" }}>
                      <thead>
                        <tr style={{ background: "var(--surface-2)", borderBottom: "2px solid var(--border)" }}>
                          <th style={{ padding: "10px 14px", textAlign: "left", fontWeight: 700, fontSize: "0.67rem", textTransform: "uppercase", letterSpacing: "0.05em", color: "var(--text-muted)" }}>Period</th>
                          <th style={{ padding: "10px 12px", textAlign: "right", fontWeight: 700, fontSize: "0.67rem", textTransform: "uppercase", letterSpacing: "0.05em", color: "var(--text-muted)" }}>Gross Pay</th>
                          <th style={{ padding: "10px 12px", textAlign: "right", fontWeight: 700, fontSize: "0.67rem", textTransform: "uppercase", letterSpacing: "0.05em", color: "var(--text-muted)" }}>Deductions</th>
                          <th style={{ padding: "10px 12px", textAlign: "right", fontWeight: 700, fontSize: "0.67rem", textTransform: "uppercase", letterSpacing: "0.05em", color: "var(--text-muted)" }}>Net Pay</th>
                          <th style={{ padding: "10px 12px", textAlign: "right", fontWeight: 700, fontSize: "0.67rem", textTransform: "uppercase", letterSpacing: "0.05em", color: "#166534" }}>Take Home</th>
                          <th style={{ padding: "10px 12px", width: 30 }}></th>
                        </tr>
                      </thead>
                      <tbody>
                        {pagedHistRecords.map((rec, idx) => {
                          const rowKey  = `${rec.year}-${rec.month}-${rec.quincena}`;
                          const isOpen  = histExpandedRow === rowKey;
                          const mName   = MONTH_NAMES[rec.month - 1];
                          const qLabel  = rec.quincena === 1 ? "1st Q" : "2nd Q";
                          const snap    = rec.snapshot;
                          const takeHome = rec.quincena === 1 ? snap?.first_quincena : snap?.second_quincena;
                          const rowBg   = idx % 2 === 0 ? "var(--surface)" : "var(--surface-2)";
                          return (
                            <Fragment key={rowKey}>
                              <tr style={{ borderBottom: "1px solid var(--border)", background: rowBg, cursor: snap ? "pointer" : "default" }}
                                onClick={() => snap && setHistExpandedRow(isOpen ? null : rowKey)}>
                                <td style={{ padding: "9px 14px", fontWeight: 600, color: "var(--psa-navy)" }}>
                                  {mName} {rec.year} <span style={{ fontWeight: 400, color: "var(--text-muted)", fontSize: "0.75rem" }}>{qLabel}</span>
                                </td>
                                <td style={{ padding: "9px 12px", textAlign: "right", fontVariantNumeric: "tabular-nums" }}>₱ {fmt(rec.gross_pay)}</td>
                                <td style={{ padding: "9px 12px", textAlign: "right", fontVariantNumeric: "tabular-nums", color: "#991b1b" }}>₱ {fmt(rec.total_deductions)}</td>
                                <td style={{ padding: "9px 12px", textAlign: "right", fontVariantNumeric: "tabular-nums" }}>₱ {fmt(rec.net_pay)}</td>
                                <td style={{ padding: "9px 12px", textAlign: "right", fontVariantNumeric: "tabular-nums", fontWeight: 700, color: "#166534" }}>₱ {fmt(takeHome ?? 0)}</td>
                                <td style={{ padding: "9px 10px", textAlign: "center", color: "var(--text-muted)" }}>
                                  {snap && <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"
                                    style={{ transform: isOpen ? "rotate(180deg)" : "none", transition: "transform 0.15s" }}>
                                    <polyline points="6 9 12 15 18 9"/>
                                  </svg>}
                                </td>
                              </tr>
                              {isOpen && snap && (() => {
                                const deds    = snap.deductions ?? {};
                                const allocs  = snap.allowances ?? {};
                                const labels  = snap.type_labels ?? {};
                                const toLabel = k => labels[k] || k.replace(/_/g, " ").replace(/\b\w/g, c => c.toUpperCase());
                                const earnKeys = Object.keys(allocs).filter(k => k !== "pera" && k !== "basic_pay");
                                const dedKeys  = Object.keys(deds).filter(k => Number(deds[k]) > 0);
                                return (
                                  <tr style={{ background: "var(--surface-2)" }}>
                                    <td colSpan={6} style={{ padding: "14px 20px", borderBottom: "1px solid var(--border)" }}>
                                      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 20 }}>
                                        <div>
                                          <div style={{ fontSize: "0.67rem", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.07em", color: "var(--text-muted)", marginBottom: 6 }}>Earnings</div>
                                          <BreakdownRow label="Basic Pay" value={snap.basic_pay ?? 0} />
                                          {allocs.pera != null && allocs.pera > 0 && <BreakdownRow label={toLabel("pera")} value={allocs.pera} />}
                                          {earnKeys.map(k => <BreakdownRow key={k} label={toLabel(k)} value={allocs[k]} />)}
                                          <BreakdownRow label="Gross Pay" value={snap.gross_pay ?? 0} bold />
                                        </div>
                                        <div>
                                          <div style={{ fontSize: "0.67rem", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.07em", color: "var(--text-muted)", marginBottom: 6 }}>Deductions</div>
                                          {dedKeys.length === 0 && <div style={{ fontSize: "0.78rem", color: "var(--text-muted)" }}>None</div>}
                                          {dedKeys.map(k => <BreakdownRow key={k} label={toLabel(k)} value={deds[k]} />)}
                                          <BreakdownRow label="Total Deductions" value={rec.total_deductions} bold />
                                        </div>
                                        <div>
                                          <div style={{ fontSize: "0.67rem", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.07em", color: "var(--text-muted)", marginBottom: 6 }}>Take Home</div>
                                          <BreakdownRow label="Net Pay" value={snap.net_pay ?? 0} />
                                          <BreakdownRow label={rec.quincena === 1 ? "1st Quincena" : "2nd Quincena"} value={takeHome ?? 0} bold />
                                        </div>
                                      </div>
                                    </td>
                                  </tr>
                                );
                              })()}
                            </Fragment>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}
                {histPageCount > 1 && (
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 10, marginTop: 14, flexWrap: "wrap" }}>
                    <button
                      disabled={histRecordsPage === 0}
                      onClick={() => { setHistRecordsPage(p => p - 1); setHistExpandedRow(null); }}
                      style={{ padding: "6px 18px", borderRadius: 6, border: "1px solid var(--border)", background: histRecordsPage === 0 ? "var(--surface-2)" : "var(--psa-navy)", color: histRecordsPage === 0 ? "var(--text-muted)" : "#fff", fontWeight: 600, fontSize: "0.82rem", cursor: histRecordsPage === 0 ? "not-allowed" : "pointer", opacity: histRecordsPage === 0 ? 0.5 : 1 }}>
                      ← Prev
                    </button>
                    <span style={{ fontSize: "0.8rem", color: "var(--text-muted)", fontVariantNumeric: "tabular-nums" }}>
                      Page {histRecordsPage + 1} of {histPageCount}
                    </span>
                    <button
                      disabled={histRecordsPage >= histPageCount - 1}
                      onClick={() => { setHistRecordsPage(p => p + 1); setHistExpandedRow(null); }}
                      style={{ padding: "6px 18px", borderRadius: 6, border: "1px solid var(--border)", background: histRecordsPage >= histPageCount - 1 ? "var(--surface-2)" : "var(--psa-navy)", color: histRecordsPage >= histPageCount - 1 ? "var(--text-muted)" : "#fff", fontWeight: 600, fontSize: "0.82rem", cursor: histRecordsPage >= histPageCount - 1 ? "not-allowed" : "pointer", opacity: histRecordsPage >= histPageCount - 1 ? 0.5 : 1 }}>
                      Next →
                    </button>
                  </div>
                )}
              </>
            )}
          </>
        );

        return isMobile ? (
          /* Mobile: single-panel — list or detail */
          <div>
            {!histSelectedEmp ? empListPanel : historyRecordsPanel}
          </div>
        ) : (
          /* Desktop: side-by-side */
          <div style={{ display: "flex", gap: 20, alignItems: "flex-start" }}>
            <div style={{ width: 280, flexShrink: 0 }}>{empListPanel}</div>
            <div style={{ flex: 1, minWidth: 0 }}>
              {!histSelectedEmp
                ? <div style={{ padding: "60px 0", textAlign: "center", color: "var(--text-muted)", fontSize: "0.85rem" }}>Select an employee to view their payslip history</div>
                : historyRecordsPanel}
            </div>
          </div>
        );
      })()}

      {archiveTab === "archive" && <div className="archive-layout">

        {/* ── LEFT: month sidebar ── */}
        <div className="archive-sidebar">
          {isMobile ? (() => {
            const cnt = archives.filter(a => Number(a.year) === selectedYear).length;
            return (
              <>
                {/* Mobile year navigator — bigger, tappable */}
                <div style={{ display: "flex", alignItems: "center", padding: "10px 12px", borderBottom: "1px solid var(--border)", background: "var(--surface-2)", gap: 8 }}>
                  <button type="button" onClick={() => setSelectedYear(y => y - 1)}
                    style={{ width: 38, height: 38, borderRadius: 8, border: "1.5px solid var(--border)", background: "var(--surface)", color: "var(--text)", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, fontSize: "1.3rem", lineHeight: 1 }}>‹</button>
                  <div style={{ flex: 1, textAlign: "center" }}>
                    <div style={{ fontWeight: 800, fontSize: "1.25rem", color: "var(--psa-navy)", lineHeight: 1 }}>{selectedYear}</div>
                    <div style={{ fontSize: "0.69rem", color: "var(--text-muted)", marginTop: 2 }}>
                      {cnt > 0 ? `${cnt} of 24 periods saved` : "No archives yet"}
                    </div>
                  </div>
                  <button type="button" onClick={() => setSelectedYear(y => y + 1)} disabled={selectedYear >= curYear}
                    style={{ width: 38, height: 38, borderRadius: 8, border: "1.5px solid var(--border)", background: selectedYear >= curYear ? "var(--surface-2)" : "var(--surface)", color: selectedYear >= curYear ? "var(--border)" : "var(--text)", cursor: selectedYear >= curYear ? "default" : "pointer", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, fontSize: "1.3rem", lineHeight: 1 }}>›</button>
                </div>

                {/* Mobile month grid — 4 columns, all 12 months visible */}
                <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 6, padding: "10px 12px" }}>
                  {MONTH_NAMES.map((monthName, idx) => {
                    const month      = idx + 1;
                    const s1         = getSlot(month, 1);
                    const s2         = getSlot(month, 2);
                    const savedCount = (s1 ? 1 : 0) + (s2 ? 1 : 0);
                    const isSelected = selectedMonth === month && !showYearSummary;
                    return (
                      <button key={month} type="button"
                        onClick={() => { setShowYearSummary(false); setSelectedMonth(month); }}
                        style={{
                          padding: "9px 4px 7px",
                          background: isSelected ? "var(--psa-navy)" : savedCount > 0 ? "var(--surface-2)" : "transparent",
                          color: isSelected ? "#fff" : "var(--text)",
                          border: `1.5px solid ${isSelected ? "var(--psa-navy)" : savedCount > 0 ? "var(--border)" : "transparent"}`,
                          borderRadius: 8,
                          cursor: "pointer",
                          fontSize: "0.79rem",
                          fontWeight: isSelected ? 700 : 500,
                          display: "flex",
                          flexDirection: "column",
                          alignItems: "center",
                          gap: 4,
                          transition: "background 0.12s",
                        }}>
                        <span>{monthName.slice(0, 3)}</span>
                        {savedCount > 0 ? (
                          <span style={{
                            width: 6, height: 6, borderRadius: 999, display: "block", flexShrink: 0,
                            background: isSelected ? "rgba(255,255,255,0.8)" : savedCount === 2 ? "#16a34a" : "#f59e0b",
                          }} />
                        ) : (
                          <span style={{ width: 6, height: 6, display: "block", flexShrink: 0 }} />
                        )}
                      </button>
                    );
                  })}
                </div>

                {/* Annual Summary */}
                {canAnnual && (
                  <div style={{ padding: "0 12px 10px" }}>
                    <button type="button" onClick={loadYearSummary}
                      style={{
                        width: "100%", padding: "8px 14px", borderRadius: 8, cursor: "pointer",
                        background: showYearSummary ? "var(--psa-navy)" : "transparent",
                        color: showYearSummary ? "#fff" : "var(--text-muted)",
                        border: `1.5px solid ${showYearSummary ? "var(--psa-navy)" : "var(--border)"}`,
                        fontSize: "0.82rem", fontWeight: 600,
                        display: "flex", alignItems: "center", justifyContent: "space-between",
                      }}>
                      <span>Annual Summary</span>
                      {yearSaved > 0 && (
                        <span style={{
                          fontSize: "0.63rem", fontWeight: 700, borderRadius: 8, padding: "1px 7px",
                          background: showYearSummary ? "rgba(255,255,255,0.2)" : "#dbeafe",
                          color: showYearSummary ? "#fff" : "var(--psa-blue)",
                        }}>{yearSaved}/24</span>
                      )}
                    </button>
                  </div>
                )}
              </>
            );
          })() : (() => {
            const cnt = archives.filter(a => Number(a.year) === selectedYear).length;
            return (
              <>
                {/* Desktop year navigator */}
                <div className="ws-year-nav">
                  <button type="button" className="ws-icon-btn ws-icon-btn--sm" onClick={() => setSelectedYear(y => y - 1)} aria-label="Previous year">
                    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><polyline points="15 18 9 12 15 6"/></svg>
                  </button>
                  <div className="ws-year-nav-label">
                    {selectedYear}
                    {cnt > 0 && <span className="ws-pill" style={{ marginLeft: 8, verticalAlign: "middle" }}>{cnt}/24</span>}
                  </div>
                  <button type="button" className="ws-icon-btn ws-icon-btn--sm" onClick={() => setSelectedYear(y => y + 1)} disabled={selectedYear >= curYear} aria-label="Next year">
                    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><polyline points="9 18 15 12 9 6"/></svg>
                  </button>
                </div>

                {/* Desktop month list */}
                <div className="archive-month-list">
                  {MONTH_NAMES.map((monthName, idx) => {
                    const month      = idx + 1;
                    const s1         = getSlot(month, 1);
                    const s2         = getSlot(month, 2);
                    const savedCount = (s1 ? 1 : 0) + (s2 ? 1 : 0);
                    const isSelected = selectedMonth === month && !showYearSummary;
                    return (
                      <div key={month} className="archive-month-item">
                        <button type="button" onClick={() => { setShowYearSummary(false); setSelectedMonth(month); }}
                          className={"archive-month-btn" + (isSelected ? " archive-month-btn--active" : "")}>
                          <span>{monthName}</span>
                          {savedCount > 0 && (
                            <span style={{
                              fontSize: "0.65rem", fontWeight: 700,
                              color: isSelected ? "var(--psa-navy)" : savedCount === 2 ? "#166534" : "#92400e",
                              background: isSelected ? "#fff" : savedCount === 2 ? "#dcfce7" : "#fef3c7",
                              borderRadius: 8, padding: "1px 7px"
                            }}>{savedCount}/2</span>
                          )}
                        </button>
                      </div>
                    );
                  })}
                  {canAnnual && <div className="archive-annual-wrap">
                    <button type="button" onClick={loadYearSummary}
                      className={"archive-month-btn archive-annual-btn" + (showYearSummary ? " archive-month-btn--active" : "")}
                      style={{ boxShadow: showYearSummary ? "var(--shadow-xs)" : "none" }}>
                      <span>Annual Summary</span>
                      {yearSaved > 0 && (
                        <span style={{ fontSize: "0.65rem", fontWeight: 700, color: "var(--psa-blue)", background: "#dbeafe", borderRadius: 8, padding: "1px 6px" }}>{yearSaved}/24</span>
                      )}
                    </button>
                  </div>}
                </div>
              </>
            );
          })()}
        </div>

        {/* ── RIGHT: detail panel ── */}
        <div className="archive-panel">
          {selectedMonth && (
            <>
              {/* Month header */}
              <div className="ws-month-head">
                <div>
                  <span className="ws-month-title">{mn}</span>
                  <span className="ws-month-year">{selectedYear}</span>
                </div>
                <span className="ws-progress" title={`${yearSaved} of 24 pay periods in ${selectedYear} are archived`}>
                  <span className="ws-progress-bar"><span style={{ width: `${(yearSaved / 24) * 100}%` }} /></span>
                  {yearSaved} of 24 periods archived in {selectedYear}
                </span>
              </div>

              {/* Quincena cards */}
              <div className="archive-quincena-grid">
                {[1, 2].map(q => {
                  const slot     = getSlot(selectedMonth, q);
                  const isActive = selectedQ === q;
                  const lastDay  = new Date(selectedYear, selectedMonth, 0).getDate();
                  const qLabel   = q === 1 ? "1st Quincena" : "2nd Quincena";
                  const qRange   = q === 1 ? `${mn} 1–15` : `${mn} 16–${lastDay}`;
                  const isSaving = saving?.month === selectedMonth && saving?.quincena === q;
                  const isExp    = exporting === slot?.id;

                  const actionButtons = slot ? (
                    <>
                      {canExport && (
                        <button type="button" className="btn-primary ws-btn-sm" onClick={() => handleExport(slot)} disabled={isExp} style={{ flex: isMobile ? 1 : "none" }}>
                          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>
                          </svg>
                          {isExp ? "Exporting…" : "Export"}
                        </button>
                      )}
                      {canManage && <>
                        <button type="button" className="ws-icon-btn" onClick={() => handleSave(selectedMonth, q)} disabled={isSaving} title="Re-save snapshot from current payroll" aria-label="Re-save snapshot">
                          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><polyline points="1 4 1 10 7 10"/><path d="M3.51 15a9 9 0 1 0 .49-4.71"/></svg>
                        </button>
                        <button type="button" className="ws-icon-btn ws-icon-btn--danger" onClick={() => handleDelete(slot)} title="Delete snapshot" aria-label="Delete snapshot">
                          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M9 6V4h6v2"/></svg>
                        </button>
                      </>}
                    </>
                  ) : canManage ? (
                    <button type="button" className="btn-primary ws-btn-sm" onClick={() => handleSave(selectedMonth, q)} disabled={isSaving} style={{ flex: isMobile ? 1 : "none" }}>
                      {isSaving ? "Saving…" : "Save snapshot"}
                    </button>
                  ) : null;

                  return (
                    <article key={q}
                      className={`ws-qcard${isActive ? " is-active" : ""}${slot ? "" : " is-empty"}${slot && canPayslip ? " is-clickable" : ""}`}
                      onClick={() => slot && canPayslip && selectQuincena(q)}>
                      <div className="ws-qcard-top">
                        <div>
                          <div className="ws-qcard-title">{qLabel}</div>
                          <div className="ws-item-sub">{qRange}, {selectedYear}</div>
                        </div>
                        <span className={`ws-pill${slot ? " ws-pill--ok" : ""}`}>
                          {slot && <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><polyline points="20 6 9 17 4 12"/></svg>}
                          {slot ? "Saved" : "Not saved"}
                        </span>
                      </div>
                      <div className="ws-qcard-meta">
                        {slot ? (
                          <>
                            <span><strong>{slot.row_count}</strong> employees</span>
                            <span>Saved {new Date(slot.updated_at).toLocaleDateString("en-PH", { month: "short", day: "numeric", year: "numeric" })} by {slot.saved_by}</span>
                          </>
                        ) : (
                          <span>No snapshot saved for this period yet.</span>
                        )}
                      </div>
                      {actionButtons && (
                        <div className="ws-qcard-actions" onClick={e => e.stopPropagation()}>
                          {actionButtons}
                        </div>
                      )}
                    </article>
                  );
                })}
              </div>

              {/* Payroll Summary Receipt */}
              {selectedQ && (
                <article className="job-card payslip-card">
                  {loadingItems ? (
                    <div style={{ padding: 40, textAlign: "center", color: "var(--text-muted)", fontSize: "0.85rem" }}>Loading summary…</div>
                  ) : !getSlot(selectedMonth, selectedQ) ? (
                    <div style={{ padding: 40, textAlign: "center", color: "var(--text-muted)", fontSize: "0.85rem" }}>Save a snapshot first to view the payroll summary.</div>
                  ) : slotSnapshot ? (() => {
                    const items     = slotSnapshot.items ?? [];
                    const sDedTypes = slotSnapshot.dedTypes ?? [];
                    const sEarnTypes= slotSnapshot.earnTypes ?? [];
                    const sum = (fn) => items.reduce((s, i) => s + Number(fn(i) || 0), 0);

                    const totalBasic    = sum(i => i.basic_pay);
                    const totalPera     = sum(i => i.allowances?.pera ?? 0);
                    const totalGross    = sum(i => i.gross_pay);
                    const totalNetPay   = sum(i => i.net_pay);
                    const totalFirstQ   = sum(i => i.first_quincena);
                    const totalSecondQ  = sum(i => i.second_quincena);
                    const takeHomeThisQ = selectedQ === 1 ? totalFirstQ : totalSecondQ;

                    const peraActive    = sEarnTypes.length === 0 || sEarnTypes.some(t => t.name === 'pera');
                    const customEarns   = sEarnTypes.filter(t => t.name !== 'pera' && t.name !== 'basic_pay');
                    const lessDedList   = sDedTypes.filter(t => t.side === 'less_deduction');
                    const mainDedList   = sDedTypes.filter(t => t.side === 'deduction');

                    const dedTotals = {};
                    for (const t of sDedTypes) {
                      dedTotals[t.name] = sum(i => i.deductions?.[t.name] ?? 0);
                    }
                    const totalLessDed      = lessDedList.reduce((s, t) => s + (dedTotals[t.name] ?? 0), 0);
                    const totalSalaryEarned = totalGross - totalLessDed;
                    const totalMainDed      = mainDedList.reduce((s, t) => s + (dedTotals[t.name] ?? 0), 0);

                    const slot = getSlot(selectedMonth, selectedQ);
                    const lastDay = new Date(selectedYear, selectedMonth, 0).getDate();
                    const periodStr = selectedQ === 1
                      ? `${mn} 1–15, ${selectedYear}`
                      : `${mn} 16–${lastDay}, ${selectedYear}`;
                    const qLabel = selectedQ === 1 ? "1st Quincena" : "2nd Quincena";

                    return (
                      <>
                        {/* Card header — same structure as PaySlipCard */}
                        <div className="job-card__header">
                          <div className="job-card__info">
                            <div className="job-card__name">{qLabel} — Payroll Summary</div>
                            <div className="job-card__tags">
                              <span className="job-card__tag">{periodStr}</span>
                              <span className="job-card__tag-sep">·</span>
                              <span className="job-card__tag">{items.length} employees</span>
                              <span className="job-card__tag-sep">·</span>
                              <span className="job-card__tag">saved {new Date(slot.updated_at).toLocaleDateString("en-PH", { month: "short", day: "numeric", year: "numeric" })} by {slot.saved_by}</span>
                            </div>
                          </div>
                          <div style={{ display: "flex", alignItems: "center", gap: 10, flexShrink: 0 }}>
                            {canExport && (
                              <button type="button" title={exportingQSummary ? "Generating…" : "Export Payroll Summary (.xlsx)"}
                                onClick={handleExportQuincenaSummary} disabled={exportingQSummary}
                                style={{ width: 32, height: 32, borderRadius: 7, border: "none", cursor: exportingQSummary ? "wait" : "pointer", background: exportingQSummary ? "var(--text-muted)" : "#1b7a43", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                                {exportingQSummary
                                  ? <svg style={{ animation: "spin 0.7s linear infinite" }} width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2.5"><circle cx="12" cy="12" r="10" strokeOpacity="0.25"/><path d="M12 2a10 10 0 0 1 10 10"/></svg>
                                  : <img alt="" width="14" height="14" src={`data:image/svg+xml,${encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="9" y1="3" x2="9" y2="21"/></svg>')}`} />}
                              </button>
                            )}
                            <div style={{ textAlign: "right" }}>
                              <div style={{ fontSize: "0.68rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.06em" }}>Total Take Home Pay</div>
                              <div style={{ fontSize: "1.05rem", fontWeight: 800, color: "#166534" }}>₱ {fmt(takeHomeThisQ)}</div>
                            </div>
                          </div>
                        </div>

                        {/* Payslip body — same classes as PaySlipCard */}
                        <div className="payslip-body" style={{ borderTop: "1px solid var(--border)", background: "var(--surface-2)", padding: "20px 24px", borderRadius: "0 0 var(--radius) var(--radius)" }}>
                          <div className="payslip-doc-header">
                            <img src={psaLogo} alt="PSA" className="payslip-logo" />
                            <div className="payslip-agency">
                              <div>Republic of the Philippines</div>
                              <div style={{ fontWeight: 800 }}>Philippine Statistics Authority</div>
                              <div>Regional Statistical Services Office II</div>
                            </div>
                            <div className="payslip-title">PAYROLL&nbsp;&nbsp;SUMMARY</div>
                          </div>
                          <div className="payslip-meta">
                            <span>Period: <b>{periodStr}</b></span>
                            <span>{items.length} Employees</span>
                          </div>

                          <div className="deduction-two-col" style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : "repeat(auto-fit, minmax(280px, 1fr))", gap: isMobile ? 16 : 32 }}>
                            {/* LEFT: Earnings + Less Deductions */}
                            <div>
                              <div className="breakdown-title">Gross Earnings</div>
                              <BreakdownRow label="Basic Salary" value={totalBasic} />
                              {peraActive && totalPera > 0 && <BreakdownRow label="PERA" value={totalPera} />}
                              {customEarns.filter(t => !isMobile || sum(i => i.allowances?.[t.name] ?? 0) > 0).map(t => (
                                <BreakdownRow key={t.name} label={t.display_name ?? t.name} value={sum(i => i.allowances?.[t.name] ?? 0)} />
                              ))}
                              <BreakdownRow label="Total Gross" value={totalGross} bold />

                              {lessDedList.length > 0 && <>
                                <div className="breakdown-title" style={{ marginTop: 18 }}>Less Deductions</div>
                                {lessDedList.filter(t => !isMobile || (dedTotals[t.name] ?? 0) > 0).map(t => (
                                  <BreakdownRow key={t.name} label={t.display_name ?? t.name} value={dedTotals[t.name] ?? 0} />
                                ))}
                                <BreakdownRow label="Total Less Deductions" value={totalLessDed} bold />
                              </>}

                              <div className="breakdown-title" style={{ marginTop: 18 }}>Salary Earned for the Period</div>
                              <BreakdownRow label="Total Salary / Earned" value={totalSalaryEarned} bold />
                            </div>

                            {/* RIGHT: Deductions */}
                            <div>
                              <div className="breakdown-title">Deductions</div>
                              {mainDedList.length === 0 && <div style={{ fontSize: "0.78rem", color: "var(--text-muted)" }}>No deduction types recorded.</div>}
                              {mainDedList.filter(t => !isMobile || (dedTotals[t.name] ?? 0) > 0).map(t => (
                                <BreakdownRow key={t.name} label={t.display_name ?? t.name} value={dedTotals[t.name] ?? 0} />
                              ))}
                              {isMobile && mainDedList.filter(t => (dedTotals[t.name] ?? 0) === 0).length > 0 && (
                                <div style={{ fontSize: "0.75rem", color: "var(--text-muted)", padding: "4px 0", fontStyle: "italic" }}>
                                  {mainDedList.filter(t => (dedTotals[t.name] ?? 0) === 0).length} deduction types with ₱0 hidden
                                </div>
                              )}
                              <BreakdownRow label="Total Deductions" value={totalMainDed} bold />
                            </div>
                          </div>

                          {/* Take Home bottom */}
                          <div className="screen-only" style={{ marginTop: 16, paddingTop: 14, borderTop: "2px solid var(--border)", display: "flex", justifyContent: "space-between", alignItems: "center", gap: 24, flexWrap: "wrap" }}>
                            <div style={{ display: "flex", alignItems: "baseline", gap: 10 }}>
                              <span style={{ fontWeight: 700, color: "var(--text-label)" }}>Total Take Home Pay</span>
                              <span style={{ fontSize: "1.2rem", fontWeight: 800, color: "#166534" }}>₱ {fmt(takeHomeThisQ)}</span>
                            </div>
                            <div style={{ border: "1px solid #16a34a", borderRadius: 8, padding: "8px 16px", background: "#f0fdf4" }}>
                              <div style={{ fontSize: "0.69rem", color: "#166534", fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.05em" }}>{qLabel}</div>
                              <div style={{ fontSize: "1rem", fontWeight: 800, color: "#166534", marginTop: 2 }}>₱ {fmt(takeHomeThisQ)}</div>
                            </div>
                          </div>
                        </div>
                      </>
                    );
                  })() : (
                    <div style={{ padding: "60px 24px", textAlign: "center", color: "var(--text-muted)" }}>
                      <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ animation: "spin 1s linear infinite", display: "block", margin: "0 auto 12px" }}><path d="M21 12a9 9 0 1 1-6.219-8.56"/></svg>
                      <div style={{ fontSize: "0.82rem" }}>Loading…</div>
                    </div>
                  )}
                </article>
              )}

              {/* Employee breakdown table */}
              {canBreakdown && selectedQ && slotSnapshot && slotSnapshot.items?.length > 0 && (() => {
                const items     = slotSnapshot.items ?? [];
                const sDedTypes = slotSnapshot.dedTypes ?? [];
                const takeHomeCol = selectedQ === 1 ? "first_quincena" : "second_quincena";

                const lessDedTypes = sDedTypes.filter(t => t.side === 'less_deduction');
                const mainDedTypes  = sDedTypes.filter(t => t.side === 'deduction');
                const TH = ({ children, green }) => (
                  <th style={{ padding: "10px 12px", textAlign: "right", fontWeight: 700, fontSize: "0.67rem", textTransform: "uppercase", letterSpacing: "0.05em", color: green ? "#166534" : "var(--text-muted)", whiteSpace: "nowrap", borderRight: "1px solid var(--border)", background: "var(--surface-2)" }}>{children}</th>
                );
                const TD = ({ children, bold, green, muted }) => (
                  <td style={{ padding: "8px 12px", textAlign: "right", fontVariantNumeric: "tabular-nums", fontSize: "0.8rem", fontWeight: bold ? 700 : 400, color: green ? "#166534" : muted ? "var(--text-muted)" : "var(--text)", borderRight: "1px solid var(--border)", whiteSpace: "nowrap" }}>{children}</td>
                );

                const totalTakeHome = items.reduce((s, i) => s + Number(i[takeHomeCol] ?? 0), 0);
                const totalBasicAll = items.reduce((s, i) => s + Number(i.basic_pay || 0), 0);
                const totalGrossAll = items.reduce((s, i) => s + Number(i.gross_pay || 0), 0);

                return (
                  <div style={{ marginTop: 16 }}>
                    <div style={{ fontSize: "0.7rem", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.08em", color: "var(--text-muted)", marginBottom: 8 }}>
                      Employee Breakdown — {items.length} records
                    </div>

                    {isMobile ? (
                      /* ── Mobile: card list ── */
                      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                        {items.map((item, idx) => {
                          const totalLessDed = lessDedTypes.reduce((s, t) => s + Number(item.deductions?.[t.name] ?? 0), 0);
                          const salaryEarned = Number(item.gross_pay) - totalLessDed;
                          const totalMainDed = mainDedTypes.reduce((s, t) => s + Number(item.deductions?.[t.name] ?? 0), 0);
                          const takeHome     = Number(item[takeHomeCol] ?? 0);
                          return (
                            <div key={item.id ?? idx} style={{ border: "1px solid var(--border)", borderRadius: "var(--radius)", background: "var(--surface)", overflow: "hidden" }}>
                              <div style={{ padding: "10px 14px", borderBottom: "1px solid var(--border)", background: "var(--surface-2)" }}>
                                <div style={{ fontWeight: 700, fontSize: "0.88rem", color: "var(--psa-navy)" }}>{item.employee_name}</div>
                                <div style={{ fontSize: "0.74rem", color: "var(--text-muted)", marginTop: 2 }}>{item.position}</div>
                              </div>
                              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 0 }}>
                                {[
                                  { label: "Basic Pay", value: fmt(item.basic_pay) },
                                  { label: "Gross Pay", value: fmt(item.gross_pay) },
                                  { label: "Total Deductions", value: fmt(totalMainDed) },
                                  { label: "Take Home", value: fmt(takeHome), green: true },
                                ].map((row, i) => (
                                  <div key={i} style={{ padding: "8px 14px", borderBottom: i < 2 ? "1px solid var(--border)" : "none", borderRight: i % 2 === 0 ? "1px solid var(--border)" : "none" }}>
                                    <div style={{ fontSize: "0.67rem", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em", color: "var(--text-muted)", marginBottom: 2 }}>{row.label}</div>
                                    <div style={{ fontSize: "0.88rem", fontWeight: row.green ? 800 : 600, color: row.green ? "#166534" : "var(--text)", fontVariantNumeric: "tabular-nums" }}>₱ {row.value}</div>
                                  </div>
                                ))}
                              </div>
                            </div>
                          );
                        })}
                        {/* Mobile totals bar */}
                        <div style={{ border: "1.5px solid #16a34a", borderRadius: "var(--radius)", background: "#f0fdf4", padding: "12px 14px" }}>
                          <div style={{ fontSize: "0.68rem", fontWeight: 700, color: "#166534", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 8 }}>Payroll Totals</div>
                          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                            {[
                              { label: "Basic Pay", value: fmt(totalBasicAll) },
                              { label: "Gross Pay", value: fmt(totalGrossAll) },
                              { label: "Total Deductions", value: fmt(items.reduce((s, i) => s + mainDedTypes.reduce((d, t) => d + Number(i.deductions?.[t.name] ?? 0), 0), 0)) },
                              { label: "Total Take Home", value: fmt(totalTakeHome), green: true },
                            ].map((r, i) => (
                              <div key={i}>
                                <div style={{ fontSize: "0.66rem", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em", color: r.green ? "#166534" : "var(--text-muted)", marginBottom: 1 }}>{r.label}</div>
                                <div style={{ fontSize: r.green ? "1rem" : "0.88rem", fontWeight: r.green ? 800 : 600, color: r.green ? "#166534" : "var(--text)", fontVariantNumeric: "tabular-nums" }}>₱ {r.value}</div>
                              </div>
                            ))}
                          </div>
                        </div>
                      </div>
                    ) : (
                      /* ── Desktop: full horizontal-scroll table ── */
                      <div style={{ border: "1px solid var(--border)", borderRadius: "var(--radius)", overflowX: "auto", background: "var(--surface)" }}>
                        <table style={{ width: "max-content", minWidth: "100%", borderCollapse: "collapse", fontSize: "0.8rem" }}>
                          <thead>
                            <tr style={{ background: "var(--surface-2)", borderBottom: "2px solid var(--border)" }}>
                              <th style={{ padding: "10px 14px", textAlign: "left", fontWeight: 700, fontSize: "0.67rem", textTransform: "uppercase", letterSpacing: "0.05em", color: "var(--text-muted)", whiteSpace: "nowrap", minWidth: 200, position: "sticky", left: 0, background: "var(--surface-2)", boxShadow: "3px 0 6px -2px rgba(0,0,0,0.10)", zIndex: 1 }}>Employee</th>
                              <TH>Basic Pay</TH>
                              <TH>Gross Pay</TH>
                              {lessDedTypes.map(t => <TH key={t.name}>{t.display_name ?? t.name}</TH>)}
                              <TH>Salary Earned</TH>
                              {mainDedTypes.map(t => <TH key={t.name}>{t.display_name ?? t.name}</TH>)}
                              <TH>Total Ded.</TH>
                              <TH green>Take Home</TH>
                            </tr>
                          </thead>
                          <tbody>
                            {items.map((item, idx) => {
                              const totalLessDed = lessDedTypes.reduce((s, t) => s + Number(item.deductions?.[t.name] ?? 0), 0);
                              const salaryEarned = Number(item.gross_pay) - totalLessDed;
                              const totalMainDed = mainDedTypes.reduce((s, t) => s + Number(item.deductions?.[t.name] ?? 0), 0);
                              const takeHome     = Number(item[takeHomeCol] ?? 0);
                              const rowBg        = idx % 2 === 0 ? "var(--surface)" : "var(--surface-2)";
                              return (
                                <tr key={item.id ?? idx} style={{ borderBottom: "1px solid var(--border)", background: rowBg }}>
                                  <td style={{ padding: "9px 14px", background: rowBg, position: "sticky", left: 0, boxShadow: "3px 0 6px -2px rgba(0,0,0,0.08)", zIndex: 1, minWidth: 200 }}>
                                    <div style={{ fontWeight: 600, fontSize: "0.82rem", color: "var(--psa-navy)", whiteSpace: "nowrap" }}>{item.employee_name}</div>
                                    <div style={{ fontSize: "0.71rem", color: "var(--text-muted)", marginTop: 1, whiteSpace: "nowrap" }}>{item.position}</div>
                                  </td>
                                  <TD>{fmt(item.basic_pay)}</TD>
                                  <TD>{fmt(item.gross_pay)}</TD>
                                  {lessDedTypes.map(t => <TD key={t.name} muted>{fmt(item.deductions?.[t.name] ?? 0)}</TD>)}
                                  <TD bold>{fmt(salaryEarned)}</TD>
                                  {mainDedTypes.map(t => <TD key={t.name} muted>{fmt(item.deductions?.[t.name] ?? 0)}</TD>)}
                                  <TD muted>{fmt(totalMainDed)}</TD>
                                  <TD bold green>{fmt(takeHome)}</TD>
                                </tr>
                              );
                            })}
                          </tbody>
                          <tfoot>
                            <tr style={{ borderTop: "2px solid var(--border)", background: "#f0fdf4" }}>
                              <td style={{ padding: "10px 14px", fontWeight: 800, fontSize: "0.77rem", color: "var(--psa-navy)", position: "sticky", left: 0, background: "#f0fdf4", boxShadow: "3px 0 6px -2px rgba(0,0,0,0.08)", letterSpacing: "0.04em" }}>TOTAL</td>
                              <TD bold>{fmt(totalBasicAll)}</TD>
                              <TD bold>{fmt(totalGrossAll)}</TD>
                              {lessDedTypes.map(t => <TD key={t.name} bold>{fmt(items.reduce((s, i) => s + Number(i.deductions?.[t.name] ?? 0), 0))}</TD>)}
                              <TD bold>{fmt(items.reduce((s, i) => s + Number(i.gross_pay || 0) - lessDedTypes.reduce((d, t) => d + Number(i.deductions?.[t.name] ?? 0), 0), 0))}</TD>
                              {mainDedTypes.map(t => <TD key={t.name} bold>{fmt(items.reduce((s, i) => s + Number(i.deductions?.[t.name] ?? 0), 0))}</TD>)}
                              <TD bold>{fmt(items.reduce((s, i) => s + mainDedTypes.reduce((d, t) => d + Number(i.deductions?.[t.name] ?? 0), 0), 0))}</TD>
                              <TD bold green>{fmt(totalTakeHome)}</TD>
                            </tr>
                          </tfoot>
                        </table>
                      </div>
                    )}
                  </div>
                );
              })()}
            </>
          )}

          {/* ── Annual Summary panel ── */}
          {showYearSummary && canAnnual && (() => {
            if (loadingYear) return (
              <div style={{ padding: 60, textAlign: "center", color: "var(--text-muted)", fontSize: "0.85rem" }}>Loading annual summary…</div>
            );
            if (yearSnapshots.length === 0) return (
              <div style={{ padding: 60, textAlign: "center", color: "var(--text-muted)", fontSize: "0.85rem" }}>No saved archives for {selectedYear}.</div>
            );

            // Year totals for the breakdown table's TOTAL row.
            const allEntries = yearSnapshots.flatMap(({ quincena, snapshot }) =>
              (snapshot.items ?? []).map(item => ({ ...item, _q: quincena }))
            );
            const sum = fn => allEntries.reduce((s, i) => s + Number(fn(i) ?? 0), 0);
            const totalGross    = sum(i => i.gross_pay);
            const totalTakeHome = sum(i => Number(i._q) === 1 ? i.first_quincena : i.second_quincena);
            const savedSlotCount = yearSnapshots.length;

            return (
              <>
                <div style={{ marginBottom: 14, display: "flex", alignItems: "center", gap: 8 }}>
                  <span style={{ fontWeight: 800, fontSize: "1.1rem", color: "var(--text)" }}>Annual Summary</span>
                  <span style={{ fontWeight: 400, fontSize: "0.88rem", color: "var(--text-muted)" }}>{selectedYear}</span>
                  <span style={{ marginLeft: "auto", fontSize: "0.74rem", color: "var(--text-muted)" }}>{savedSlotCount} of 24 slots archived</span>
                </div>

                {/* Quincena breakdown table */}
                <div>
                  <div style={{ fontSize: "0.7rem", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.08em", color: "var(--text-muted)", marginBottom: 8 }}>
                    Quincena Breakdown — {savedSlotCount} saved periods
                  </div>
                  <div style={{ border: "1px solid var(--border)", borderRadius: "var(--radius)", overflowX: "auto", background: "var(--surface)" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.8rem" }}>
                      <thead>
                        <tr style={{ background: "var(--surface-2)", borderBottom: "2px solid var(--border)" }}>
                          <th style={{ padding: "10px 14px", textAlign: "left", fontWeight: 700, fontSize: "0.67rem", textTransform: "uppercase", letterSpacing: "0.05em", color: "var(--text-muted)" }}>Period</th>
                          <th style={{ padding: "10px 12px", textAlign: "right", fontWeight: 700, fontSize: "0.67rem", textTransform: "uppercase", letterSpacing: "0.05em", color: "var(--text-muted)" }}>Employees</th>
                          <th style={{ padding: "10px 12px", textAlign: "right", fontWeight: 700, fontSize: "0.67rem", textTransform: "uppercase", letterSpacing: "0.05em", color: "var(--text-muted)" }}>Gross Pay</th>
                          <th style={{ padding: "10px 12px", textAlign: "right", fontWeight: 700, fontSize: "0.67rem", textTransform: "uppercase", letterSpacing: "0.05em", color: "#166534" }}>Take Home</th>
                        </tr>
                      </thead>
                      <tbody>
                        {yearSnapshots.sort((a, b) => a.month - b.month || a.quincena - b.quincena).map(({ month, quincena, snapshot }, idx) => {
                          const items = snapshot.items ?? [];
                          const gross = items.reduce((s, i) => s + Number(i.gross_pay || 0), 0);
                          const takeHome = items.reduce((s, i) => s + Number(quincena === 1 ? i.first_quincena : i.second_quincena || 0), 0);
                          const rowBg = idx % 2 === 0 ? "var(--surface)" : "var(--surface-2)";
                          return (
                            <tr key={`${month}-${quincena}`} style={{ borderBottom: "1px solid var(--border)", background: rowBg }}>
                              <td style={{ padding: "9px 14px", fontWeight: 600, color: "var(--psa-navy)", fontSize: "0.82rem" }}>
                                {MONTH_NAMES[month - 1]} — {quincena === 1 ? "1st Quincena" : "2nd Quincena"}
                              </td>
                              <td style={{ padding: "9px 12px", textAlign: "right", color: "var(--text-muted)" }}>{items.length}</td>
                              <td style={{ padding: "9px 12px", textAlign: "right", fontVariantNumeric: "tabular-nums" }}>{fmt(gross)}</td>
                              <td style={{ padding: "9px 12px", textAlign: "right", fontWeight: 700, color: "#166534", fontVariantNumeric: "tabular-nums" }}>{fmt(takeHome)}</td>
                            </tr>
                          );
                        })}
                      </tbody>
                      <tfoot>
                        <tr style={{ borderTop: "2px solid var(--border)", background: "#f0fdf4" }}>
                          <td style={{ padding: "10px 14px", fontWeight: 800, fontSize: "0.77rem", color: "var(--psa-navy)", letterSpacing: "0.04em" }}>TOTAL</td>
                          <td style={{ padding: "10px 12px" }} />
                          <td style={{ padding: "10px 12px", textAlign: "right", fontWeight: 700, fontVariantNumeric: "tabular-nums" }}>{fmt(totalGross)}</td>
                          <td style={{ padding: "10px 12px", textAlign: "right", fontWeight: 800, color: "#166534", fontVariantNumeric: "tabular-nums" }}>{fmt(totalTakeHome)}</td>
                        </tr>
                      </tfoot>
                    </table>
                  </div>
                </div>
              </>
            );
          })()}
        </div>

      </div>}

    </>
  );
}

// ── SettingsView ──────────────────────────────────────────────────────────────
// Pages shown in the role permission editor, and used to summarise what a role covers.
// `keys: null` marks a page every signed-in user gets automatically.
const ROLE_PERMISSION_SECTIONS = [
  { label: "Main", pages: [
    { label: "Dashboard",          keys: ["dashboard.view", "dashboard.financial_report"], hint: "Overview stats and the financial report" },
    { label: "Employee Dashboard", keys: null,                                              hint: "Every signed-in user's own payslip and history" },
    { label: "201 Files",          keys: ["files201.view"],                                 hint: "Their own 201 documents (self-service)" },
  ]},
  { label: "Employee", pages: [
    { label: "Employees",     keys: ["employees.view", "employees.create", "employees.edit", "employees.delete"], hint: "The employee directory and records" },
    { label: "Compensations", keys: ["deductions.view", "deductions.edit"],                                       hint: "Employee deductions and loans" },
  ]},
  { label: "Leave", pages: [
    { label: "File own leave",                keys: ["leave.file"],     hint: "Submit their own CS Form No. 6" },
    { label: "HR certification (7.A)",        keys: ["leave.hr"],       hint: "Certify leave credits and forward to CRASD" },
    { label: "Chief recommendation (7.B)",    keys: ["leave.crasd"],    hint: "Recommend approval or disapproval to the Director" },
    { label: "Director approval",             keys: ["leave.director"], hint: "Give the final approval or disapproval" },
    { label: "Manage all leave applications", keys: ["leave.manage"],   hint: "See every application and correct Section 7" },
  ]},
  { label: "Payroll", pages: [
    { label: "Pay Slips",   keys: ["payslips.view", "payslips.export_pdf", "payslips.export_excel"], hint: "Payslips and the payroll register" },
    { label: "Archive",     keys: ["archive.view", "archive.manage"],                                hint: "Saved payroll periods" },
    { label: "Adjustments", keys: ["deductions.view", "deductions.edit", "deductions.bulk_edit"],    hint: "Deduction changes across many employees" },
  ]},
  { label: "Organization", pages: [
    // Both pages are gated by the same Reference Data permission, so they switch together.
    { label: "Salary Rank",     keys: ["settings.reference.view", "settings.reference.manage"], hint: "Salary grades and steps (shares access with Workforce Setup)" },
    { label: "Workforce Setup", keys: ["settings.reference.view", "settings.reference.manage"], hint: "Positions and divisions (shares access with Salary Rank)" },
  ]},
  { label: "System", pages: [
    { label: "Change Password",       keys: ["account.change_password", "account.avatar"],               hint: "Their own password and profile picture" },
    { label: "Payroll Configuration", keys: ["settings.payroll_types.view", "settings.payroll_types.manage"], hint: "Payroll types and deduction settings" },
  ]},
];

// Short labels for the individual permissions inside a page.
const PERMISSION_PART_LABELS = {
  "dashboard.view": "View", "dashboard.financial_report": "Financial report",
  "employees.view": "View", "employees.create": "Add", "employees.edit": "Edit", "employees.delete": "Deactivate",
  "deductions.view": "View", "deductions.edit": "Edit", "deductions.bulk_edit": "Bulk edit",
  "payslips.view": "View", "payslips.export_pdf": "Export DOCX/PDF", "payslips.export_excel": "Export Excel",
  "archive.view": "View", "archive.manage": "Save / delete",
  "settings.reference.view": "View", "settings.reference.manage": "Manage",
  "account.change_password": "Password", "account.avatar": "Profile picture",
  "settings.payroll_types.view": "View", "settings.payroll_types.manage": "Manage",
};

function RoleManagerPanel({ refreshKey, currentUser, isSuperAdmin }) {
  const [roles, setRoles]               = useState([]);
  const [catalog, setCatalog]           = useState({});
  const [loading, setLoading]           = useState(true);
  const [showCreateModal, setShowCreateModal] = useState(false);
  const [editPermsRole, setEditPermsRole]     = useState(null);
  const [createForm, setCreateForm]     = useState({ name: "" });
  const [createError, setCreateError]   = useState("");
  const [createSaving, setCreateSaving] = useState(false);
  const [permsSaving, setPermsSaving]   = useState(false);
  const isSuperadminRole = editPermsRole?.name?.toLowerCase() === "superadmin";
  const [selectedPerms, setSelectedPerms] = useState(new Set());
  const loadData = useCallback(async () => {
    setLoading(true);
    try {
      const [rolesRes, catalogRes] = await Promise.all([
        fetch(`${apiUrl}/admin/roles`, { credentials: "include" }),
        fetch(`${apiUrl}/admin/permissions`, { credentials: "include" }),
      ]);
      if (rolesRes.ok) setRoles(await rolesRes.json());
      if (catalogRes.ok) setCatalog(await catalogRes.json());
    } finally { setLoading(false); }
  }, []);

  useEffect(() => { loadData(); }, [loadData, refreshKey]);

  async function handleCreateRole(e) {
    e.preventDefault();
    setCreateError("");
    const trimmedName = createForm.name.trim();
    if (!trimmedName) { setCreateError("Role name is required."); return; }
    const dupe = roles.find(r => r.name.toLowerCase() === trimmedName.toLowerCase());
    if (dupe) { setCreateError(`A role named "${dupe.name}" already exists.`); return; }
    if (createSaving) return;
    setCreateSaving(true);
    try {
      const res = await fetch(`${apiUrl}/admin/roles`, {
        method: "POST", credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name: trimmedName, permissionKeys: [] }),
      });
      const d = await res.json();
      if (!res.ok) { setCreateError(d.error ?? "Failed to create role."); return; }
      setShowCreateModal(false);
      setCreateForm({ name: "" });
      await loadData();
    } finally { setCreateSaving(false); }
  }

  async function handleDeleteRole(role) {
    if (!window.confirm(`Delete role "${role.name}"?\n\nThis cannot be undone. Roles with assigned users cannot be deleted.`)) return;
    const res = await fetch(`${apiUrl}/admin/roles/${role.id}`, { method: "DELETE", credentials: "include" });
    const d = await res.json();
    if (!res.ok) { alert(d.error ?? "Failed to delete role."); return; }
    await loadData();
  }

  function openPermEditor(role) {
    setEditPermsRole(role);
    setSelectedPerms(new Set(Array.isArray(role.permission_keys) ? role.permission_keys : []));
  }

  async function handleSavePermissions() {
    if (!editPermsRole) return;
    setPermsSaving(true);
    try {
      const res = await fetch(`${apiUrl}/admin/roles/${editPermsRole.id}/permissions`, {
        method: "PUT", credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ permissionKeys: [...selectedPerms] }),
      });
      const d = await res.json();
      if (!res.ok) { alert(d.error ?? "Failed to save permissions."); return; }
      setEditPermsRole(null);
      await loadData();
    } finally { setPermsSaving(false); }
  }

  const togglePerm = (key) => setSelectedPerms(prev => {
    const next = new Set(prev);
    if (next.has(key)) next.delete(key); else next.add(key);
    return next;
  });

  const toggleCategory = (keys) => {
    const allSelected = keys.every(k => selectedPerms.has(k));
    setSelectedPerms(prev => {
      const next = new Set(prev);
      if (allSelected) keys.forEach(k => next.delete(k)); else keys.forEach(k => next.add(k));
      return next;
    });
  };

  if (loading) return (
    <div style={{ display: "flex", alignItems: "center", gap: 8, color: "var(--text-muted)", fontSize: "0.88rem", padding: "8px 0" }}>
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ animation: "spin 1s linear infinite", flexShrink: 0 }}>
        <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
      </svg>
      Loading roles…
    </div>
  );

  const areasFor = (keys) => [...new Set(ROLE_PERMISSION_SECTIONS.flatMap(s => s.pages)
    .filter(p => p.keys && p.keys.some(k => keys.includes(k)))
    .map(p => p.label))];
  const inUse = roles.filter(r => (r.user_count ?? 0) > 0).length;
  const ShieldIcon = () => (
    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>
  );

  return (
    <div>
      <div className="rm-head">
        <div className="um-sub">
          {roles.length === 0 ? "No roles created yet" : `${roles.length} role${roles.length !== 1 ? "s" : ""} · ${inUse} in use`}
        </div>
        <button type="button" className="rm-create"
          onClick={() => { setCreateForm({ name: "" }); setCreateError(""); setShowCreateModal(true); }}>
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
          Create role
        </button>
      </div>

      {roles.length === 0 ? (
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 10, padding: "40px 16px", background: "var(--surface-2)", border: "1.5px dashed var(--border)", borderRadius: "var(--radius)", textAlign: "center" }}>
          <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="var(--text-muted)" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" style={{ opacity: 0.4 }}>
            <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>
          </svg>
          <div style={{ fontSize: "0.88rem", fontWeight: 600, color: "var(--text-muted)" }}>No roles yet</div>
          <div style={{ fontSize: "0.8rem", color: "var(--text-muted)", opacity: 0.7 }}>Create a role to start assigning permissions to users.</div>
        </div>
      ) : (
        <div className="rm-list">
          {[...roles].sort((a, b) => (b.name.toLowerCase() === "superadmin") - (a.name.toLowerCase() === "superadmin")).map(role => {
            const keys = Array.isArray(role.permission_keys) ? role.permission_keys : [];
            const userCount = role.user_count ?? 0;
            const isSuper = role.name.toLowerCase() === "superadmin";
            const areas = areasFor(keys);
            const shown = areas.slice(0, 5);
            return (
              <div key={role.id} className="rm-item">
                <span className={`rm-badge${isSuper ? " is-super" : ""}`}>{role.name.slice(0, 2).toUpperCase()}</span>
                <div className="rm-main">
                  <div className="rm-title">
                    <span className="rm-name">{role.name}</span>
                    {isSuper && <span className="rm-full">Full access</span>}
                  </div>
                  <div className="rm-meta">
                    {!isSuper && <span>{keys.length} permission{keys.length !== 1 ? "s" : ""}</span>}
                    {userCount > 0
                      ? <span className="rm-users">{userCount} user{userCount !== 1 ? "s" : ""}</span>
                      : <span>Not assigned to anyone</span>}
                  </div>
                  <div className="rm-areas">
                    {isSuper ? (
                      <span className="rm-area rm-area--all">Every page and action</span>
                    ) : keys.length === 0 ? (
                      <span className="um-flag">No access yet: users with this role can't open any page</span>
                    ) : (
                      <>
                        {shown.map(a => <span key={a} className="rm-area">{a}</span>)}
                        {areas.length > shown.length && <span className="rm-area rm-area--more">+{areas.length - shown.length} more</span>}
                      </>
                    )}
                  </div>
                </div>
                <div className="rm-actions">
                  <button type="button" className="um-btn um-btn--neutral" onClick={() => openPermEditor(role)}>
                    <ShieldIcon /> {isSuper ? "View access" : "Edit access"}
                  </button>
                  {!isSuper && (
                    <button type="button" className="rm-del" onClick={() => handleDeleteRole(role)} title="Delete role" aria-label={`Delete role ${role.name}`}>
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                        <polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/>
                        <path d="M10 11v6"/><path d="M14 11v6"/><path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/>
                      </svg>
                    </button>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      )}

      {/* Create Role Modal */}
      {showCreateModal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.48)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
          <div style={{ background: "var(--surface)", borderRadius: "var(--radius-lg)", padding: "28px 28px 24px", width: "100%", maxWidth: 440, boxShadow: "var(--shadow-lg)" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 20 }}>
              <div style={{ width: 36, height: 36, borderRadius: "var(--radius-sm)", background: "#eff6ff", border: "1px solid #bfdbfe", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#1d4ed8" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>
                </svg>
              </div>
              <h3 style={{ margin: 0, color: "var(--text)", fontSize: "1rem", fontWeight: 700 }}>Create New Role</h3>
            </div>
            <form onSubmit={handleCreateRole} style={{ display: "flex", flexDirection: "column", gap: 14 }}>
              <label style={{ fontSize: "0.88rem", fontWeight: 600, color: "var(--text)" }}>
                Role Name <span style={{ color: "#dc2626" }}>*</span>
                <input
                  type="text"
                  value={createForm.name}
                  onChange={e => setCreateForm(f => ({ ...f, name: e.target.value }))}
                  placeholder="e.g. Payroll Clerk"
                  maxLength={100}
                  required
                  autoFocus
                  style={{ display: "block", width: "100%", marginTop: 5, padding: "8px 12px", borderRadius: "var(--radius-sm)", border: "1.5px solid var(--border-2)", fontSize: "0.9rem", boxSizing: "border-box", color: "var(--text)", background: "var(--surface)" }}
                />
              </label>
              {createError && (
                <div style={{ display: "flex", alignItems: "center", gap: 6, color: "#dc2626", fontSize: "0.85rem", background: "#fef2f2", border: "1px solid #fecaca", borderRadius: "var(--radius-sm)", padding: "8px 12px" }}>
                  <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
                  {createError}
                </div>
              )}
              <div className="form-actions" style={{ marginTop: 4 }}>
                <button type="button" className="ghost-button" onClick={() => setShowCreateModal(false)}>Cancel</button>
                <button type="submit" disabled={createSaving}>
                  {createSaving ? "Creating…" : "Create Role"}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Edit Permissions Side Panel */}
      {editPermsRole && (() => {
        const allKeys = [...new Set(ROLE_PERMISSION_SECTIONS.flatMap(s => s.pages.flatMap(p => p.keys ?? [])))];
        const has = (k) => isSuperadminRole || selectedPerms.has(k);
        const onCount = allKeys.filter(has).length;
        const initial = new Set(Array.isArray(editPermsRole.permission_keys) ? editPermsRole.permission_keys : []);
        const dirty = !isSuperadminRole && (initial.size !== selectedPerms.size || [...selectedPerms].some(k => !initial.has(k)));
        const setKeys = (keys, on) => setSelectedPerms(prev => {
          const next = new Set(prev);
          keys.forEach(k => on ? next.add(k) : next.delete(k));
          return next;
        });
        const stateOf = (keys) => {
          const n = keys.filter(has).length;
          return n === 0 ? "off" : n === keys.length ? "on" : "some";
        };
        const Switch = ({ state, onToggle, label }) => (
          <button type="button" role="switch" aria-label={label}
            aria-checked={state === "on" ? "true" : state === "some" ? "mixed" : "false"}
            className={`pe-switch is-${state}`} disabled={isSuperadminRole} onClick={onToggle}>
            <span className="pe-knob" />
          </button>
        );
        return (
          <div className="pe-backdrop">
            <aside className="pe-panel" role="dialog" aria-modal="true" aria-label={`Permissions for ${editPermsRole.name}`}>
              <header className="pe-head">
                <div className="pe-title">
                  <span className={`rm-badge${isSuperadminRole ? " is-super" : ""}`}>{editPermsRole.name.slice(0, 2).toUpperCase()}</span>
                  <div>
                    <div className="pe-kicker">Role permissions</div>
                    <h3 className="pe-name">
                      {editPermsRole.name}
                      {isSuperadminRole && <span className="rm-full">Full access</span>}
                    </h3>
                  </div>
                </div>
                <div>
                  <div className="pe-progress-row">
                    <span><strong>{onCount}</strong> of {allKeys.length} permissions on</span>
                    {!isSuperadminRole && (
                      <span className="pe-bulk">
                        <button type="button" onClick={() => setKeys(allKeys, true)}>All on</button>
                        <button type="button" onClick={() => setKeys(allKeys, false)}>All off</button>
                      </span>
                    )}
                  </div>
                  <div className="pe-bar"><span style={{ width: `${allKeys.length ? (onCount / allKeys.length) * 100 : 0}%` }} /></div>
                </div>
                {isSuperadminRole && (
                  <div className="pe-note">superadmin always has every permission. This role can't be edited.</div>
                )}
              </header>

              <div className="pe-body">
                {ROLE_PERMISSION_SECTIONS.map(section => {
                  const sectionKeys = [...new Set(section.pages.flatMap(p => p.keys ?? []))];
                  const sectionState = stateOf(sectionKeys);
                  const pages = section.pages.filter(p => p.keys);
                  const pagesOn = pages.filter(p => stateOf(p.keys) === "on").length;
                  return (
                    <section key={section.label} className="pe-section">
                      <div className="pe-section-head">
                        <span className="pe-section-name">{section.label}</span>
                        <span className="pe-section-count">{pagesOn}/{pages.length}</span>
                        <Switch state={sectionState} label={`Everything in ${section.label}`}
                          onToggle={() => setKeys(sectionKeys, sectionState !== "on")} />
                      </div>
                      {section.pages.map(page => {
                        if (!page.keys) return (
                          <div key={page.label} className="pe-page is-auto">
                            <div className="pe-page-text">
                              <span className="pe-page-name">{page.label}</span>
                              <span className="pe-hint">{page.hint}</span>
                            </div>
                            <span className="pe-auto">Always on</span>
                          </div>
                        );
                        const pageState = stateOf(page.keys);
                        return (
                          <div key={page.label} className="pe-page">
                            <div className="pe-page-main">
                              <div className="pe-page-text">
                                <span className="pe-page-name">{page.label}</span>
                                {page.hint && <span className="pe-hint">{page.hint}</span>}
                              </div>
                              <Switch state={pageState} label={page.label} onToggle={() => setKeys(page.keys, pageState !== "on")} />
                            </div>
                            {page.keys.length > 1 && (
                              <div className="pe-parts">
                                {page.keys.map(k => (
                                  <button key={k} type="button" className={`pe-part${has(k) ? " is-on" : ""}`}
                                    disabled={isSuperadminRole} aria-pressed={has(k)} onClick={() => setKeys([k], !has(k))}>
                                    {has(k) && (
                                      <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><polyline points="20 6 9 17 4 12"/></svg>
                                    )}
                                    {PERMISSION_PART_LABELS[k] ?? k}
                                  </button>
                                ))}
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </section>
                  );
                })}
              </div>

              <footer className="pe-foot">
                {dirty && <span className="pe-dirty">Unsaved changes</span>}
                <button type="button" className="um-btn um-btn--neutral" onClick={() => setEditPermsRole(null)}>
                  {isSuperadminRole ? "Close" : "Cancel"}
                </button>
                {!isSuperadminRole && (
                  <button type="button" className="pe-save" onClick={handleSavePermissions} disabled={permsSaving || !dirty}>
                    {permsSaving ? "Saving…" : "Save permissions"}
                  </button>
                )}
              </footer>
            </aside>
          </div>
        );
      })()}
    </div>
  );
}

// ── UserEmployeeLinkPanel ──────────────────────────────────────────────────────
function UserEmployeeLinkPanel({ apiUrl }) {
  const [links, setLinks]         = useState([]);
  const [employees, setEmployees] = useState([]);
  const [loading, setLoading]     = useState(true);
  const [saving, setSaving]       = useState(null);
  const [saved, setSaved]         = useState(null);
  const [openId, setOpenId]       = useState(null);
  const [linkQuery, setLinkQuery] = useState("");
  const hasLinks = useRef(false);

  const fetchLinks = useCallback(async (q = "", silent = false) => {
    if (!silent) setLoading(true);
    try {
      const limit = q.trim() ? 50 : 5;
      const [uRes, eRes] = await Promise.all([
        fetch(`${apiUrl}/admin/user-employee-links?q=${encodeURIComponent(q)}&limit=${limit}`, { credentials: "include" }),
        employees.length === 0 ? fetch(`${apiUrl}/employees?limit=200`, { credentials: "include" }) : Promise.resolve(null),
      ]);
      if (uRes.ok) setLinks(await uRes.json());
      if (eRes?.ok) { const d = await eRes.json(); setEmployees(Array.isArray(d) ? d : (d.data ?? [])); }
    } finally { setLoading(false); }
  }, [apiUrl, employees.length]);

  useEffect(() => { fetchLinks("", hasLinks.current); hasLinks.current = true; }, [fetchLinks]);

  useEffect(() => {
    const t = setTimeout(() => fetchLinks(linkQuery), 300);
    return () => clearTimeout(t);
  }, [linkQuery, fetchLinks]);

  async function handleLink(userId, employeeId) {
    setSaving(userId);
    await fetch(`${apiUrl}/admin/users/${userId}/link-employee`, {
      method: "PATCH", credentials: "include",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ employeeId: employeeId || null }),
    });
    setSaving(null);
    setSaved(userId);
    setTimeout(() => setSaved(null), 2000);
    fetchLinks(linkQuery, true);
  }

  const takenEmployeeIds = new Set(links.filter(u => u.employee_id).map(u => u.employee_id));

  const bumpLink = (userId) => {
    if (linkQuery.trim()) return;
    setLinks(prev => {
      const idx = prev.findIndex(u => u.id === userId);
      if (idx <= 0) return prev;
      const updated = [...prev];
      const [user] = updated.splice(idx, 1);
      updated.unshift(user);
      return updated;
    });
  };

  return (
    <div style={{ display: "grid", gridTemplateColumns: "3fr 2fr", minHeight: 400 }}>
      {/* Left — search + list */}
      <div style={{ padding: "20px 28px", borderRight: "1px solid var(--border)" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <span style={{ fontSize: "0.6rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.12em" }}>System Users</span>
        </div>
        <div style={{ position: "relative", marginBottom: 16 }}>
          <svg style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)", pointerEvents: "none" }} width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="var(--text-muted)" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
          </svg>
          <input style={{ width: "100%", boxSizing: "border-box", paddingLeft: 32, paddingRight: 12, paddingTop: 8, paddingBottom: 8, fontSize: "0.83rem", border: "1px solid var(--border)", borderRadius: 8, background: "var(--surface-2)", color: "var(--text)", outline: "none" }}
            placeholder="Search users…" value={linkQuery}
            onChange={e => { setLinkQuery(e.target.value); setOpenId(null); }} />
        </div>
        {loading ? (
          <div style={{ padding: "40px 0", textAlign: "center", color: "var(--text-muted)" }}>
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ animation: "spin 1s linear infinite", display: "block", margin: "0 auto 10px" }}><path d="M21 12a9 9 0 1 1-6.219-8.56"/></svg>
            <div style={{ fontSize: "0.82rem" }}>Loading…</div>
          </div>
        ) : links.length === 0 ? (
          <div style={{ textAlign: "center", padding: "40px 0", color: "var(--text-muted)", fontSize: "0.83rem" }}>
            {linkQuery.trim() ? `No results for "${linkQuery}".` : "No users found."}
          </div>
        ) : (
          <LinkList links={links} employees={employees} takenEmployeeIds={takenEmployeeIds} linkQuery={linkQuery} openId={openId} setOpenId={setOpenId} saving={saving} saved={saved} handleLink={handleLink} bumpLink={bumpLink} />
        )}
      </div>
      {/* Right — summary */}
      <div style={{ padding: "20px 28px" }}>
        <div style={{ fontSize: "0.6rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.12em", marginBottom: 14 }}>Link Status</div>
        {[
          { label: "Linked", count: links.filter(u => u.employee_id).length, color: "#166534" },
          { label: "Unlinked", count: links.filter(u => !u.employee_id).length, color: "#991b1b" },
        ].map(r => (
          <div key={r.label} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0" }}>
            <span style={{ fontSize: "0.83rem", fontWeight: 600 }}>{r.label}</span>
            <span style={{ fontSize: "0.85rem", fontWeight: 800, color: r.count > 0 ? r.color : "var(--text-muted)" }}>{r.count}</span>
          </div>
        ))}
        <div style={{ marginTop: 16, paddingTop: 14, borderTop: "1px solid var(--border)" }}>
          <div style={{ display: "flex", justifyContent: "space-between" }}>
            <span style={{ fontSize: "0.83rem", fontWeight: 700 }}>Total Users</span>
            <span style={{ fontSize: "0.9rem", fontWeight: 800, color: "var(--psa-navy)" }}>{links.length}</span>
          </div>
        </div>
      </div>
    </div>
  );
}

function LinkList({ links, employees, takenEmployeeIds, linkQuery, openId, setOpenId, saving, saved, handleLink, bumpLink }) {
  return (
    <div>
      {links.map(u => {
        const isLinked  = !!u.employee_id;
        const isOpen    = openId === u.id;
        const available = employees.filter(e => !takenEmployeeIds.has(e.id) || e.id === u.employee_id);
        return (
          <div key={u.id}>
            <div
              onClick={() => { const opening = openId !== u.id; setOpenId(isOpen ? null : u.id); if (opening) bumpLink(u.id); }}
              style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 8px", cursor: "pointer", borderRadius: 6, transition: "background 0.12s" }}
              onMouseEnter={e => e.currentTarget.style.background = "var(--surface-2)"}
              onMouseLeave={e => e.currentTarget.style.background = "transparent"}
            >
              <div>
                <div style={{ fontSize: "0.85rem", fontWeight: 700 }}>{u.username}</div>
                <div style={{ fontSize: "0.72rem", color: "var(--text-muted)", marginTop: 2 }}>
                  {isLinked
                    ? <><span style={{ color: "#166534", fontWeight: 600 }}>{u.employee_name}</span>{u.position ? ` · ${u.position}` : ""}</>
                    : <span style={{ fontStyle: "italic" }}>No employee linked</span>
                  }
                </div>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 10, flexShrink: 0, marginLeft: 12 }}>
                {saved === u.id && <span style={{ fontSize: "0.72rem", color: "#166534", fontWeight: 700 }}>✓ Saved</span>}
                <span style={{ fontSize: "0.72rem", fontWeight: 700, color: isLinked ? "#166534" : "#991b1b" }}>
                  {isLinked ? "Linked" : "Unlinked"}
                </span>
              </div>
            </div>
            {isOpen && (
              <div style={{ padding: "10px 8px 14px", marginBottom: 4 }}>
                <div style={{ fontSize: "0.6rem", fontWeight: 700, color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.12em", marginBottom: 8 }}>Assign Employee</div>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <select
                    value={u.employee_id ?? ""}
                    disabled={saving === u.id}
                    onChange={e => handleLink(u.id, e.target.value)}
                    style={{ flex: 1, maxWidth: 300, fontSize: "0.83rem", fontWeight: 500, padding: "8px 12px", borderRadius: 8, border: "1px solid var(--border)", background: "var(--surface)", color: "var(--text)", cursor: "pointer", outline: "none", appearance: "auto" }}
                  >
                    <option value="">— no link —</option>
                    {available.map(e => (
                      <option key={e.id} value={e.id}>{e.name}{e.employee_number ? ` · ${e.employee_number}` : ""}</option>
                    ))}
                  </select>
                  {saving === u.id
                    ? <span style={{ fontSize: "0.75rem", color: "var(--text-muted)" }}>Saving…</span>
                    : saved === u.id
                      ? <span style={{ fontSize: "0.75rem", fontWeight: 700, color: "#166534" }}>✓ Saved</span>
                      : null
                  }
                </div>
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

function SettingsView({ currentUser, isAdmin, isSuperAdmin, permissions = [], hasAvatar, avatarBlobUrl, setAvatarBlobUrl, setCurrentUser, setHasAvatar, refreshKey }) {
  const can    = (key) => isSuperAdmin || permissions.includes(key);
  const canAny = (...keys) => isSuperAdmin || keys.some(k => permissions.includes(k));
  const isMobile = useIsMobile();
  const settingsTabs = [
    { id: "account", label: "My Account", visible: can("account.avatar") || can("account.change_password") },
    { id: "users",   label: "User Management", visible: isSuperAdmin },
    { id: "payroll", label: "Payroll Configuration", visible: isSuperAdmin || canAny("settings.payroll_types.view", "settings.payroll_types.manage") },
    { id: "leave",   label: "Leave Signatories", visible: can("leave.manage") },
  ].filter(t => t.visible);
  const [activeTab, setActiveTab] = useState(settingsTabs[0]?.id ?? "account");
  const [pendingUsers, setPendingUsers]         = useState([]);
  const [allUsers, setAllUsers]                 = useState([]);
  const [pendingLoading, setPendingLoading]     = useState(false);
  const [allUsersLoading, setAllUsersLoading]   = useState(false);
  const [pwForm, setPwForm]                 = useState({ current: "", next: "", confirm: "" });
  const [pwError, setPwError]               = useState("");
  const [pwSuccess, setPwSuccess]           = useState("");
  const [showPw, setShowPw]                 = useState(false);
  const [avatarLoading, setAvatarLoading]   = useState(false);
  const [avatarUrl, setAvatarUrl]           = useState(avatarBlobUrl ?? null);
  const [sigForm, setSigForm]               = useState({ preparedByName: "", preparedByTitle: "", certifiedByName: "", certifiedByTitle: "", approvingName: "", approvingTitle: "", disbursingName: "", disbursingTitle: "" });
  const [sigLoading, setSigLoading]         = useState(false);
  const [sigSaving, setSigSaving]           = useState(false);
  const [sigSuccess, setSigSuccess]         = useState("");
  const [resetPwModal, setResetPwModal]       = useState(null);
  const [resetPwForm, setResetPwForm]         = useState({ newPassword: "", confirmPassword: "" });
  const [resetPwError, setResetPwError]       = useState("");
  const [resetPwLoading, setResetPwLoading]   = useState(false);
  const [userSearch, setUserSearch]           = useState("");
  const [userPage, setUserPage]               = useState(0);
  const [userFilter, setUserFilter]           = useState("all");
  const [openUserId, setOpenUserId]           = useState(null);
  const [recentUserIds, setRecentUserIds]     = useState(() => { try { return JSON.parse(localStorage.getItem("psa_recent_users") || "[]"); } catch { return []; } });
  const [rolePickerOpen, setRolePickerOpen]   = useState(false);
  const [settingsRoles, setSettingsRoles]     = useState([]);
  const [roleDropdowns, setRoleDropdowns]     = useState({});
  const [linkEmpSearch, setLinkEmpSearch]     = useState("");
  const [linkEmpResults, setLinkEmpResults]   = useState([]);
  const [linkEmpLoading, setLinkEmpLoading]   = useState(false);
  const fileRef = useRef(null);

  useEffect(() => { if (avatarBlobUrl) setAvatarUrl(avatarBlobUrl); }, [avatarBlobUrl]);

  const loadSignatories = useCallback(async () => {
    if (!isAdmin) return;
    setSigLoading(true);
    try {
      const res = await fetch(`${apiUrl}/settings/signatories`, { credentials: "include" });
      if (res.ok) {
        const d = await res.json();
        setSigForm({
          preparedByName:  d?.prepared_by_name  ?? "",
          preparedByTitle: d?.prepared_by_title ?? "",
          certifiedByName:  d?.certified_by_name  ?? "",
          certifiedByTitle: d?.certified_by_title ?? "",
          approvingName:  d?.approving_name  ?? "",
          approvingTitle: d?.approving_title ?? "",
          disbursingName:  d?.disbursing_name  ?? "",
          disbursingTitle: d?.disbursing_title ?? "",
        });
      }
    } finally { setSigLoading(false); }
  }, [isAdmin]);

  useEffect(() => { loadSignatories(); }, [loadSignatories]);

  async function handleSaveSignatories(e) {
    e.preventDefault();
    setSigSuccess("");
    setSigSaving(true);
    const res = await fetch(`${apiUrl}/settings/signatories`, {
      method: "PUT", credentials: "include",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(sigForm),
    });
    setSigSaving(false);
    if (res.ok) setSigSuccess("Signatories updated.");
    else { const d = await res.json(); alert(d.error ?? "Failed to save signatories."); }
  }

  const loadPendingUsers = useCallback(async () => {
    if (!isAdmin && !can("admin.users.approve")) return;
    setPendingLoading(true);
    try {
      const res = await fetch(`${apiUrl}/admin/pending-users`, { credentials: "include" });
      if (res.ok) setPendingUsers(await res.json());
    } finally { setPendingLoading(false); }
  }, [isAdmin, permissions]);

  const loadAllUsers = useCallback(async () => {
    if (!isAdmin && !can("admin.users.view")) return;
    setAllUsersLoading(true);
    try {
      const [usersRes, rolesRes] = await Promise.all([
        fetch(`${apiUrl}/admin/users`, { credentials: "include" }),
        fetch(`${apiUrl}/admin/roles`, { credentials: "include" }),
      ]);
      if (usersRes.ok) setAllUsers(await usersRes.json());
      if (rolesRes.ok) setSettingsRoles(await rolesRes.json());
    } finally { setAllUsersLoading(false); }
  }, [isAdmin, permissions]);

  useEffect(() => { loadPendingUsers(); loadAllUsers(); }, [loadPendingUsers, loadAllUsers, refreshKey]);

  async function handleApprove(id) {
    await fetch(`${apiUrl}/admin/approve/${id}`, { method: "POST", credentials: "include" });
    loadPendingUsers(); loadAllUsers();
  }
  async function handleReject(id) {
    await fetch(`${apiUrl}/admin/reject/${id}`, { method: "POST", credentials: "include" });
    loadPendingUsers(); loadAllUsers();
  }
  async function handleBlock(id, username) {
    if (!window.confirm(`Block "${username}"?\n\nThis will prevent them from logging in.`)) return;
    await fetch(`${apiUrl}/admin/block/${id}`, { method: "POST", credentials: "include" });
    loadAllUsers();
  }
  async function handleUnblock(id) {
    await fetch(`${apiUrl}/admin/unblock/${id}`, { method: "POST", credentials: "include" });
    loadAllUsers();
  }
  async function handleDeleteUser(id, username) {
    if (!window.confirm(`Delete user "${username}"?\n\nThis action cannot be undone.`)) return;
    await fetch(`${apiUrl}/admin/users/${id}`, { method: "DELETE", credentials: "include" });
    loadAllUsers();
  }

  async function handleLinkEmployee(userId, employeeId) {
    const res = await fetch(`${apiUrl}/admin/users/${userId}/link-employee`, {
      method: "PATCH", credentials: "include",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ employeeId: employeeId || null }),
    });
    if (res.ok) { setLinkEmpSearch(""); setLinkEmpResults([]); loadAllUsers(); }
    else { const d = await res.json(); alert(d.error ?? "Failed to link employee."); }
  }

  async function searchLinkEmp(q) {
    setLinkEmpSearch(q);
    if (!q.trim()) { setLinkEmpResults([]); return; }
    setLinkEmpLoading(true);
    try {
      const r = await fetch(`${apiUrl}/employees?q=${encodeURIComponent(q)}&limit=8`, { credentials: "include" });
      if (r.ok) { const d = await r.json(); setLinkEmpResults(d.data ?? []); }
    } finally { setLinkEmpLoading(false); }
  }

  async function handleAssignRoleSettings(userId, roleId) {
    setRoleDropdowns(prev => ({ ...prev, [userId]: roleId }));
    const res = await fetch(`${apiUrl}/admin/users/${userId}/role`, {
      method: "PATCH", credentials: "include",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ roleId: roleId || null }),
    });
    if (!res.ok) {
      const d = await res.json();
      alert(d.error ?? "Failed to assign role.");
      loadAllUsers();
    }
  }

  async function handleAdminResetPassword(e) {
    e.preventDefault();
    setResetPwError("");
    if (resetPwForm.newPassword !== resetPwForm.confirmPassword) { setResetPwError("Passwords do not match."); return; }
    setResetPwLoading(true);
    try {
      const res = await fetch(`${apiUrl}/admin/users/${resetPwModal.id}/password`, {
        method: "PATCH", credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ newPassword: resetPwForm.newPassword }),
      });
      const d = await res.json();
      if (res.ok) {
        setResetPwModal(null);
        setResetPwForm({ newPassword: "", confirmPassword: "" });
        alert(`Password for "${resetPwModal.username}" has been reset successfully.`);
      } else {
        setResetPwError(d.error ?? "Failed to reset password.");
      }
    } finally { setResetPwLoading(false); }
  }

  async function handlePasswordChange(e) {
    e.preventDefault();
    setPwError(""); setPwSuccess("");
    if (pwForm.next !== pwForm.confirm) { setPwError("New passwords do not match."); return; }
    const res = await fetch(`${apiUrl}/auth/password`, {
      method: "PATCH", credentials: "include",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ currentPassword: pwForm.current, newPassword: pwForm.next }),
    });
    const d = await res.json();
    if (res.ok) { setPwSuccess("Password changed successfully."); setPwForm({ current: "", next: "", confirm: "" }); }
    else setPwError(d.error ?? "Failed to change password.");
  }

  async function handleAvatarUpload(file) {
    if (!file) return;
    setAvatarLoading(true);
    const fd = new FormData();
    fd.append("avatar", file);
    const res = await fetch(`${apiUrl}/auth/avatar`, { method: "POST", credentials: "include", body: fd });
    setAvatarLoading(false);
    if (res.ok) {
      setHasAvatar(true);
      // Fetch fresh /auth/me to get updated avatarDataUrl
      fetch(`${apiUrl}/auth/me`, { credentials: "include", cache: "no-store" })
        .then(r => r.ok ? r.json() : null)
        .then(data => {
          if (data?.avatarDataUrl) {
            setAvatarBlobUrl(data.avatarDataUrl);
            setAvatarUrl(data.avatarDataUrl);
          }
        })
        .catch(() => {});
      alert("Avatar updated.");
    } else { const d = await res.json(); alert(d.error ?? "Upload failed."); }
  }

  return (
    <>
      <div className="page-header">
        <div className="page-header-title">Settings</div>
        <div className="page-header-sub">
          {{
            account: "Your profile photo and password",
            users:   "Accounts, approvals, roles and permissions",
            payroll: "Payslip signatories and income / deduction types",
            leave:   "Names printed on the leave form (CS Form No. 6)",
          }[activeTab] ?? "Manage your account settings"}
        </div>
      </div>

      {settingsTabs.length > 1 && (
        <div className="st-tabs" role="tablist" aria-label="Settings sections">
          {settingsTabs.map(t => (
            <button key={t.id} type="button" role="tab" aria-selected={activeTab === t.id}
              className={`st-tab${activeTab === t.id ? " is-on" : ""}`}
              onClick={() => setActiveTab(t.id)}>
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                {t.id === "account" && <><circle cx="12" cy="8" r="4"/><path d="M4 21v-1a6 6 0 0 1 6-6h4a6 6 0 0 1 6 6v1"/></>}
                {t.id === "users" && <><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></>}
                {t.id === "payroll" && <><rect x="2" y="5" width="20" height="14" rx="2"/><line x1="2" y1="10" x2="22" y2="10"/><line x1="6" y1="15" x2="10" y2="15"/></>}
                {t.id === "leave" && <><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/><path d="m9 16 2 2 4-4"/></>}
              </svg>
              {isMobile ? { account: "Account", users: "Users", payroll: "Payroll", leave: "Leave" }[t.id] : t.label}
              {t.id === "users" && pendingUsers.length > 0 && (
                <span className="st-tab-badge" title={`${pendingUsers.length} waiting for approval`}>{pendingUsers.length}</span>
              )}
            </button>
          ))}
        </div>
      )}

      {activeTab === "account" && (can("account.avatar") || can("account.change_password")) && (() => {
        const pw = pwForm.next;
        const rules = [
          ["8+ characters",        pw.length >= 8],
          ["Uppercase letter",     /[A-Z]/.test(pw)],
          ["Number",               /[0-9]/.test(pw)],
          ["Symbol",               /[^A-Za-z0-9]/.test(pw)],
          ["Different from current", pw.length > 0 && pw !== pwForm.current],
          ["Passwords match",      pw.length > 0 && pw === pwForm.confirm],
        ];
        const editPw = (field) => (e) => { setPwSuccess(""); setPwForm(f => ({ ...f, [field]: e.target.value })); };
        const pwType = showPw ? "text" : "password";
        const canPhoto = can("account.avatar");
        const initial = <span className="acct-initial">{currentUser?.[0]?.toUpperCase() ?? "?"}</span>;
        return (
          <div className="acct">
            <div className="panel acct-profile">
              {canPhoto ? (
                <button type="button" className="acct-avatar" onClick={() => fileRef.current?.click()} disabled={avatarLoading} title="Change photo" aria-label="Change profile photo">
                  {avatarUrl ? <img src={avatarUrl} alt="" /> : initial}
                  <span className="acct-cam" aria-hidden="true">
                    <svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor"><path d="M12 15.2A3.2 3.2 0 1 1 12 8.8a3.2 3.2 0 0 1 0 6.4zM9 3L7.17 5H4a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2h-3.17L15 3H9z"/></svg>
                  </span>
                </button>
              ) : (
                <span className="acct-avatar is-static">{avatarUrl ? <img src={avatarUrl} alt="" /> : initial}</span>
              )}
              <input ref={fileRef} type="file" accept="image/jpeg,image/png,image/webp,image/gif" style={{ display: "none" }}
                onChange={e => { const f = e.target.files?.[0]; if (f) handleAvatarUpload(f); e.target.value = ""; }} />

              <div className="acct-id">
                <div className="acct-name">
                  {currentUser}
                  {isSuperAdmin && <span className="rm-full">Superadmin</span>}
                </div>
                {canPhoto && (
                  <>
                    <div className="um-sub">Profile photo: JPEG, PNG, WebP or GIF, up to 2 MB</div>
                    <div className="acct-id-actions">
                      <button type="button" className="um-btn um-btn--neutral um-btn--sm" onClick={() => fileRef.current?.click()} disabled={avatarLoading}>
                        {avatarLoading ? "Uploading…" : avatarUrl ? "Change photo" : "Upload photo"}
                      </button>
                    </div>
                  </>
                )}
              </div>

              {!isMobile && (
                <a className="acct-desktop" href="https://github.com/JohnCasAsh/psa-payroll-system/releases/latest/download/PSAPayrollSetup.exe" download>
                  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>
                  </svg>
                  <span className="acct-desktop-text">
                    <span className="acct-desktop-title">Windows desktop app</span>
                    <span className="um-sub">Download the installer (.exe)</span>
                  </span>
                </a>
              )}
            </div>

            {can("account.change_password") && (
              <div className="panel acct-pw">
                <div className="acct-pw-intro">
                  <h3 className="panel-title"><span className="panel-title-dot" />Change password</h3>
                  <p className="um-sub">Use 8+ characters with an uppercase letter, a number, and a symbol.</p>
                </div>

                <form onSubmit={handlePasswordChange} className="acct-pw-form">
                  <label className="acct-field acct-full">
                    Current password
                    <input type={pwType} value={pwForm.current} onChange={editPw("current")} placeholder="Enter current password" autoComplete="current-password" required />
                  </label>
                  <label className="acct-field">
                    New password
                    <input type={pwType} value={pwForm.next} onChange={editPw("next")} placeholder="New password" autoComplete="new-password" required />
                  </label>
                  <label className="acct-field">
                    Confirm new password
                    <input type={pwType} value={pwForm.confirm} onChange={editPw("confirm")} placeholder="Repeat new password" autoComplete="new-password" required />
                  </label>
                  {pw.length > 0 && (
                    <ul className="acct-rules acct-full" aria-label="Password requirements">
                      {rules.map(([label, ok]) => (
                        <li key={label} className={ok ? "is-ok" : ""}>
                          {ok
                            ? <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><polyline points="20 6 9 17 4 12"/></svg>
                            : <span className="acct-rule-dot" aria-hidden="true" />}
                          {label}
                        </li>
                      ))}
                    </ul>
                  )}
                  {pwError   && <div className="error acct-full">{pwError}</div>}
                  {pwSuccess && <div className="acct-success acct-full">{pwSuccess}</div>}
                  <div className="acct-pw-actions acct-full">
                    <label className="acct-show">
                      <input type="checkbox" checked={showPw} onChange={e => setShowPw(e.target.checked)} />
                      Show passwords
                    </label>
                    <button type="submit" className="acct-submit">Update password</button>
                  </div>
                </form>
              </div>
            )}
          </div>
        );
      })()}

      {activeTab === "payroll" && isAdmin && (
        <div className="panel">
          <div className="sig-head">
            <div>
              <h3 className="panel-title"><span className="panel-title-dot" />Payslip Signatories</h3>
              <div className="um-sub">These names print on the signature lines of every payslip.</div>
            </div>
            {sigSuccess && (
              <div className="sig-saved">
                <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><polyline points="20 6 9 17 4 12"/></svg>
                Saved
              </div>
            )}
          </div>
          {sigLoading ? (
            <div className="um-empty">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ animation: "spin 1s linear infinite" }}><path d="M21 12a9 9 0 1 1-6.219-8.56"/></svg>
              Loading…
            </div>
          ) : (
            <form onSubmit={handleSaveSignatories}>
              <div className="sig-grid">
                {[
                  { label: "Prepared By", sub: "Certification of correctness", nameKey: "preparedByName", titleKey: "preparedByTitle", namePh: "e.g. Maria L. Santos", titlePh: "e.g. Administrative Officer IV", required: true, section: "A" },
                  { label: "Certified By", sub: "Certified correct", nameKey: "certifiedByName", titleKey: "certifiedByTitle", namePh: "e.g. Ana P. Garcia", titlePh: "e.g. Accountant III", required: true, section: "B" },
                  { label: "Approving Official", sub: "Approved for payment", nameKey: "approvingName", titleKey: "approvingTitle", namePh: "e.g. Juan D. Dela Cruz", titlePh: "e.g. Regional Director", section: "C" },
                  { label: "Disbursing Officer", sub: "Section D", nameKey: "disbursingName", titleKey: "disbursingTitle", namePh: "e.g. Maria S. Santos", titlePh: "e.g. Accountant II", section: "D" },
                ].map(sig => {
                  const name = (sigForm[sig.nameKey] ?? "").trim();
                  const title = (sigForm[sig.titleKey] ?? "").trim();
                  return (
                    <div key={sig.nameKey} className="sig-card">
                      <div className="sig-card-head">
                        <span className="sig-letter">{sig.section}</span>
                        <div className="sig-card-title">
                          <div className="sig-role">
                            {sig.label}
                            {sig.required && <span className="sig-req">Required</span>}
                          </div>
                          <div className="um-sub">{sig.sub}</div>
                        </div>
                      </div>
                      <div className="sig-fields">
                        <label className="sig-field">
                          Full name
                          <input value={sigForm[sig.nameKey]} onChange={e => setSigForm(f => ({ ...f, [sig.nameKey]: e.target.value }))} placeholder={sig.namePh} required={!!sig.required} />
                        </label>
                        <label className="sig-field">
                          Title / designation
                          <input value={sigForm[sig.titleKey]} onChange={e => setSigForm(f => ({ ...f, [sig.titleKey]: e.target.value }))} placeholder={sig.titlePh} required={!!sig.required} />
                        </label>
                      </div>
                      <div className="sig-preview" aria-hidden="true">
                        <span className="sig-preview-kicker">On the payslip</span>
                        <span className={`sig-preview-name${name ? "" : " is-empty"}`}>{name || "Name"}</span>
                        <span className="sig-preview-line" />
                        <span className={`sig-preview-title${title ? "" : " is-empty"}`}>{title || "Title / designation"}</span>
                      </div>
                    </div>
                  );
                })}
              </div>
              <div className="sig-foot">
                <span className="um-sub">Archived payrolls keep the signatories they were saved with.</span>
                <button type="submit" className="sig-save" disabled={sigSaving}>{sigSaving ? "Saving…" : "Save signatories"}</button>
              </div>
            </form>
          )}
        </div>
      )}

      {activeTab === "users" && isSuperAdmin && (() => {
        const q = userSearch.trim().toLowerCase();
        const roleNameOf = (u) => settingsRoles.find(r => r.id === u.role_id)?.name ?? null;
        const recentRank = Object.fromEntries(recentUserIds.map((id, i) => [id, i]));
        const sortedUsers = [...allUsers].sort((a, b) => {
          const ra = recentRank[a.id] ?? Infinity;
          const rb = recentRank[b.id] ?? Infinity;
          return ra !== rb ? ra - rb : a.username.localeCompare(b.username);
        });
        const FILTERS = [
          ["all",      "All",        () => true],
          ["norole",   "No role",    u => !u.role_id],
          ["unlinked", "Not linked", u => !u.employee_id],
          ["blocked",  "Blocked",    u => u.status === "blocked"],
        ];
        const counts = Object.fromEntries(FILTERS.map(([k, , f]) => [k, allUsers.filter(f).length]));
        const inFilter = FILTERS.find(([k]) => k === userFilter)?.[2] ?? (() => true);
        const PAGE_SIZE = 5;
        const allFiltered = sortedUsers.filter(inFilter).filter(u => !q
          || u.username.toLowerCase().includes(q)
          || (roleNameOf(u) ?? "").toLowerCase().includes(q)
          || (u.employee_name ?? "").toLowerCase().includes(q));
        const totalPages = Math.ceil(allFiltered.length / PAGE_SIZE);
        const safePage = Math.min(userPage ?? 0, Math.max(0, totalPages - 1));
        const filtered = allFiltered.slice(safePage * PAGE_SIZE, safePage * PAGE_SIZE + PAGE_SIZE);

        const STATUS = {
          active:   ["Active",   "um-chip--ok"],
          pending:  ["Pending",  "um-chip--warn"],
          blocked:  ["Blocked",  "um-chip--bad"],
          rejected: ["Rejected", "um-chip--bad"],
        };
        const AVATAR_COLORS = ["#0d2346", "#1a3a6b", "#0f766e", "#7c2d12", "#5b21b6", "#9d174d", "#1e40af"];
        const avatarColor = (u) => u.status === "blocked" ? "#94a3b8"
          : AVATAR_COLORS[[...u.username].reduce((h, c) => (h * 31 + c.charCodeAt(0)) >>> 0, 7) % AVATAR_COLORS.length];
        const fmtDate = (d) => d ? new Date(d).toLocaleDateString("en-PH", { month: "short", day: "numeric", year: "numeric" }) : "—";
        const fmtDateTime = (d) => d ? new Date(d).toLocaleString("en-PH", { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" }) : "—";
        const Chevron = ({ className }) => (
          <svg className={className} width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M9 18l6-6-6-6"/></svg>
        );
        const Caret = () => (
          <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true" style={{ color: "var(--text-muted)", flexShrink: 0 }}><polyline points="6 9 12 15 18 9"/></svg>
        );

        return (
          <>
            {pendingUsers.length > 0 && (
              <div className="um-pending">
                <div className="um-pending-head">
                  <span className="um-pending-dot" aria-hidden="true" />
                  <strong>{pendingUsers.length} account{pendingUsers.length > 1 ? "s" : ""} waiting for approval</strong>
                  <span>New sign-ups can't sign in until approved.</span>
                </div>
                <div className="um-pending-list">
                  {pendingUsers.map(u => (
                    <div key={u.id} className="um-pending-row">
                      <span className="um-avatar" style={{ background: avatarColor(u) }}>{u.username[0].toUpperCase()}</span>
                      <div className="um-acct-text">
                        <span className="um-name">{u.username}</span>
                        <span className="um-sub">Registered {fmtDateTime(u.created_at)}</span>
                      </div>
                      <div className="um-pending-actions">
                        <button type="button" className="um-btn um-btn--ghost-danger" onClick={() => handleReject(u.id)}>Reject</button>
                        <button type="button" className="um-btn um-btn--ok" onClick={() => handleApprove(u.id)}>Approve</button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            <div className="panel" style={{ marginBottom: 16 }}>
              <div className="um-head">
                <div>
                  <h3 className="panel-title"><span className="panel-title-dot" />User Accounts</h3>
                  <div className="um-sub um-head-sub">
                    {allUsers.length} account{allUsers.length === 1 ? "" : "s"} · {allUsers.length - counts.unlinked} linked to an employee record
                  </div>
                </div>
                <div className="um-search">
                  <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
                  <input type="text" placeholder="Search username, role, or employee…" value={userSearch}
                    onChange={e => { setUserSearch(e.target.value); setUserPage(0); }} />
                  {userSearch && <button type="button" className="um-clear" onClick={() => setUserSearch("")} aria-label="Clear search">×</button>}
                </div>
              </div>

              <div className="um-filters" role="tablist" aria-label="Filter accounts">
                {FILTERS.map(([k, label]) => (
                  <button key={k} type="button" role="tab" aria-selected={userFilter === k}
                    className={`um-filter${userFilter === k ? " is-on" : ""}${(k === "norole" || k === "unlinked") && counts[k] > 0 ? " has-warn" : ""}`}
                    onClick={() => { setUserFilter(k); setUserPage(0); }}>
                    {label}<span className="um-filter-n">{counts[k]}</span>
                  </button>
                ))}
              </div>

              {allUsersLoading ? (
                <div className="um-empty">
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ animation: "spin 1s linear infinite" }}><path d="M21 12a9 9 0 1 1-6.219-8.56"/></svg>
                  Loading accounts…
                </div>
              ) : (
                <>
                  <div className="um-cols" aria-hidden="true">
                    <span>Account</span><span>Role</span><span>Employee record</span><span>Status</span><span />
                  </div>
                  <div className="um-list">
                    {filtered.length === 0 ? (
                      <div className="um-empty">{q ? `No accounts match “${userSearch.trim()}”.` : "No accounts in this view."}</div>
                    ) : filtered.map(u => {
                      const roleName = roleNameOf(u);
                      const isOpen = openUserId === u.id;
                      const currentRoleId = roleDropdowns[u.id] !== undefined ? roleDropdowns[u.id] : (u.role_id ?? "");
                      const currentRoleName = settingsRoles.find(r => r.id === currentRoleId)?.name ?? null;
                      const allOptions = [{ id: "", name: "No Role" }, ...settingsRoles];
                      const [statusLabel, statusClass] = STATUS[u.status] ?? [u.status, "um-chip--warn"];
                      const setupNote = !u.role_id && !u.employee_id ? "Finish setup: assign a role and link this login to an employee record."
                        : !u.role_id ? "No role yet: this account can sign in but can't open any page."
                        : !u.employee_id ? "Not linked to an employee: it can't see payslips or 201 files, or file leave."
                        : null;
                      return (
                        <div key={u.id} className={`um-item${isOpen ? " is-open" : ""}${u.status === "blocked" ? " is-blocked" : ""}`}>
                          <button type="button" className="um-row" aria-expanded={isOpen}
                            onClick={() => { setOpenUserId(isOpen ? null : u.id); setRolePickerOpen(false); setLinkEmpResults([]); setLinkEmpSearch(""); }}>
                            <span className="um-acct um-c-acct">
                              <span className="um-avatar" style={{ background: avatarColor(u) }}>{u.username[0].toUpperCase()}</span>
                              <span className="um-acct-text">
                                <span className="um-name">{u.username}</span>
                                <span className="um-sub">Joined {fmtDate(u.created_at)}</span>
                              </span>
                            </span>
                            <span className="um-meta">
                              <span className="um-cell">{roleName ? <span className="um-role">{roleName}</span> : <span className="um-flag">No role</span>}</span>
                              <span className="um-cell">
                                {u.employee_id
                                  ? <span className="um-emp"><span className="um-emp-name">{u.employee_name ?? "—"}</span><span className="um-sub">#{u.employee_number}</span></span>
                                  : <span className="um-flag um-flag--soft">Not linked</span>}
                              </span>
                            </span>
                            <span className="um-cell um-c-status"><span className={`um-chip ${statusClass}`}>{statusLabel}</span></span>
                            <Chevron className="um-chev" />
                          </button>

                          {isOpen && (
                            <div className="um-detail">
                              {setupNote && (
                                <div className="um-setup">
                                  <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true" style={{ flexShrink: 0, marginTop: 1 }}><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
                                  <span>{setupNote}</span>
                                </div>
                              )}

                              <div className="um-fields">
                                <div className="um-field">
                                  <div className="um-label">Role</div>
                                  <div className="um-help">Controls which pages and actions this account can use.</div>
                                  <div className="um-picker">
                                    <button type="button" className="um-select" onClick={() => setRolePickerOpen(v => !v)} aria-expanded={rolePickerOpen}>
                                      <span>{currentRoleName ?? "No Role"}</span><Caret />
                                    </button>
                                    {rolePickerOpen && (
                                      <>
                                        <div style={{ position: "fixed", inset: 0, zIndex: 98 }} onClick={() => setRolePickerOpen(false)} />
                                        <div className="um-menu">
                                          {allOptions.map(r => (
                                            <button key={r.id} type="button" className={`um-opt${r.id === currentRoleId ? " is-on" : ""}`}
                                              onClick={() => { handleAssignRoleSettings(u.id, r.id); setRolePickerOpen(false); }}>
                                              {r.name}
                                            </button>
                                          ))}
                                        </div>
                                      </>
                                    )}
                                  </div>
                                </div>

                                <div className="um-field">
                                  <div className="um-label">Employee record</div>
                                  <div className="um-help">Links this login to a payroll record for payslips, 201 files, and leave.</div>
                                  {u.employee_id ? (
                                    <div className="um-linked">
                                      <span className="um-emp">
                                        <span className="um-emp-name">{u.employee_name ?? "—"}</span>
                                        <span className="um-sub">Employee #{u.employee_number}</span>
                                      </span>
                                      <button type="button" className="um-btn um-btn--neutral um-btn--sm" onClick={() => handleLinkEmployee(u.id, null)}>Unlink</button>
                                    </div>
                                  ) : (
                                    <div className="um-picker">
                                      <button type="button" className="um-select"
                                        onClick={async () => {
                                          if (linkEmpResults.length > 0) { setLinkEmpResults([]); return; }
                                          setLinkEmpLoading(true);
                                          try {
                                            const r = await fetch(`${apiUrl}/employees?limit=100`, { credentials: "include" });
                                            if (r.ok) { const d = await r.json(); setLinkEmpResults(d.data ?? []); }
                                          } finally { setLinkEmpLoading(false); }
                                        }}>
                                        <span>{linkEmpLoading ? "Loading…" : "Select employee"}</span><Caret />
                                      </button>
                                      {linkEmpResults.length > 0 && (
                                        <div className="um-menu">
                                          <div className="um-menu-search">
                                            <input autoFocus type="text" placeholder="Search name or number…" value={linkEmpSearch}
                                              onChange={e => setLinkEmpSearch(e.target.value)} />
                                          </div>
                                          <div className="um-menu-list">
                                            {linkEmpResults
                                              .filter(emp => !linkEmpSearch.trim() || emp.name?.toLowerCase().includes(linkEmpSearch.toLowerCase()) || emp.employee_number?.includes(linkEmpSearch.trim()))
                                              .map(emp => (
                                                <button key={emp.id} type="button" className="um-opt"
                                                  onClick={() => { handleLinkEmployee(u.id, emp.id); setLinkEmpResults([]); setLinkEmpSearch(""); }}>
                                                  <span className="um-emp-name" style={{ flex: 1 }}>{emp.name}</span>
                                                  <span className="um-sub">#{emp.employee_number}</span>
                                                </button>
                                              ))}
                                          </div>
                                        </div>
                                      )}
                                    </div>
                                  )}
                                </div>
                              </div>

                              <div className="um-actions">
                                <button type="button" className="um-btn um-btn--danger" onClick={() => handleDeleteUser(u.id, u.username)}>
                                  <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/></svg>
                                  Delete account
                                </button>
                                <div className="um-actions-right">
                                  <button type="button" className="um-btn um-btn--neutral"
                                    onClick={() => { setResetPwModal({ id: u.id, username: u.username }); setResetPwForm({ newPassword: "", confirmPassword: "" }); setResetPwError(""); }}>
                                    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d="M21 2l-2 2m-7.61 7.61a5.5 5.5 0 1 1-7.778 7.778 5.5 5.5 0 0 1 7.777-7.777zm0 0L15.5 7.5m0 0l3 3L22 7l-3-3m-3.5 3.5L19 4"/></svg>
                                    Reset password
                                  </button>
                                  {u.status === "active" && (
                                    <button type="button" className="um-btn um-btn--warn" onClick={() => handleBlock(u.id, u.username)}>
                                      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>
                                      Block
                                    </button>
                                  )}
                                  {u.status === "blocked" && (
                                    <button type="button" className="um-btn um-btn--ok" onClick={() => handleUnblock(u.id)}>
                                      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 9.9-1"/></svg>
                                      Unblock
                                    </button>
                                  )}
                                </div>
                              </div>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                  {totalPages > 1 && (
                    <div className="um-pager">
                      <button type="button" className="um-btn um-btn--neutral um-btn--sm" onClick={() => setUserPage(p => Math.max(0, p - 1))} disabled={safePage === 0}>← Prev</button>
                      <span>Page {safePage + 1} of {totalPages}</span>
                      <button type="button" className="um-btn um-btn--neutral um-btn--sm" onClick={() => setUserPage(p => Math.min(totalPages - 1, p + 1))} disabled={safePage >= totalPages - 1}>Next →</button>
                    </div>
                  )}
                </>
              )}
            </div>

            {/* Role Manager — full width below */}
            <div className="panel">
              <div className="section-heading" style={{ marginBottom: 16 }}>
                <h3 className="panel-title"><span className="panel-title-dot" />Role Manager</h3>
              </div>
              <RoleManagerPanel refreshKey={refreshKey} currentUser={currentUser} isSuperAdmin={isSuperAdmin} />
            </div>
          </>
        );
      })()}

      {activeTab === "payroll" && (isAdmin || canAny("settings.payroll_types.view", "settings.payroll_types.manage")) && (
        <div className="panel" style={{ marginTop: 24 }}>
          <PayrollTypesPanel canManage={isAdmin || can("settings.payroll_types.manage")} />
        </div>
      )}

      {activeTab === "leave" && can("leave.manage") && <LeaveSignatoriesPanel />}

      {resetPwModal && (
        <div
          style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}
          onClick={e => { if (e.target === e.currentTarget) setResetPwModal(null); }}
        >
          <div style={{ background: "var(--surface)", borderRadius: "var(--radius-lg)", padding: "28px 32px", width: "100%", maxWidth: 400, boxShadow: "0 20px 60px rgba(0,0,0,0.3)" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 20 }}>
              <div style={{ width: 38, height: 38, borderRadius: "50%", background: "#eff6ff", border: "2px solid #bfdbfe", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="#1d4ed8" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M21 2l-2 2m-7.61 7.61a5.5 5.5 0 1 1-7.778 7.778 5.5 5.5 0 0 1 7.777-7.777zm0 0L15.5 7.5m0 0l3 3L22 7l-3-3m-3.5 3.5L19 4"/>
                </svg>
              </div>
              <div>
                <div style={{ fontWeight: 700, fontSize: "1rem", color: "var(--text)" }}>Reset Password</div>
                <div style={{ fontSize: "0.8rem", color: "var(--text-muted)" }}>for <strong>{resetPwModal.username}</strong></div>
              </div>
            </div>
            <form onSubmit={handleAdminResetPassword} style={{ display: "flex", flexDirection: "column", gap: 14 }}>
              <label style={{ fontSize: "0.88rem", fontWeight: 500, color: "var(--text)" }}>New Password
                <input
                  type="password"
                  value={resetPwForm.newPassword}
                  onChange={e => setResetPwForm(f => ({ ...f, newPassword: e.target.value }))}
                  placeholder="8+ chars, uppercase, number, symbol"
                  autoComplete="new-password"
                  required
                  style={{ marginTop: 4, display: "block", width: "100%", boxSizing: "border-box" }}
                />
              </label>
              <label style={{ fontSize: "0.88rem", fontWeight: 500, color: "var(--text)" }}>Confirm Password
                <input
                  type="password"
                  value={resetPwForm.confirmPassword}
                  onChange={e => setResetPwForm(f => ({ ...f, confirmPassword: e.target.value }))}
                  placeholder="Repeat new password"
                  autoComplete="new-password"
                  required
                  style={{ marginTop: 4, display: "block", width: "100%", boxSizing: "border-box" }}
                />
              </label>
              {resetPwError && <div className="error">{resetPwError}</div>}
              <div className="form-actions" style={{ marginTop: 4 }}>
                <button type="button" onClick={() => setResetPwModal(null)}>Cancel</button>
                <button type="submit" disabled={resetPwLoading}>{resetPwLoading ? "Saving…" : "Reset Password"}</button>
              </div>
            </form>
          </div>
        </div>
      )}

    </>
  );
}

// ── App Root ──────────────────────────────────────────────────────────────────
function App() {
  const [currentUser, setCurrentUser]     = useState(null);
  const [isAdmin, setIsAdmin]             = useState(false);
  const [isSuperAdmin, setIsSuperAdmin]   = useState(false);
  const [permissions, setPermissions]     = useState([]);
  const [userRole, setUserRole]           = useState(null);
  const [hasAvatar, setHasAvatar]         = useState(false);
  const [avatarBlobUrl, setAvatarBlobUrl] = useState(null);
  const [view, setView]                   = useState(() => window.location.hash.slice(1) || "dashboard");
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [sidebarOpen, setSidebarOpen]     = useState(false);
  const [isMobile, setIsMobile]           = useState(() => window.innerWidth < 768);
  const [ready, setReady]                 = useState(false);
  const [authError, setAuthError]         = useState(false);
  const [authKey, setAuthKey]             = useState(0);
  const [showLoader, setShowLoader]       = useState(false);
  const [refreshKey, setRefreshKey]       = useState(0);
  const [dedTypes, setDedTypes]           = useState([]);
  const [earnTypes, setEarnTypes]         = useState([]);
  const [signatories, setSignatories]     = useState(null);
  const [leaveNotes, setLeaveNotes]       = useState([]);   // unread leave pop-ups
  const [notifKey, setNotifKey]           = useState(0);    // bumped by the "leave-notify" SSE event
  const [leaveFocus, setLeaveFocus]       = useState(null); // leave id to open when going to Leave

  // Permission helpers — memoised functions to avoid recreating on every render
  const can    = useCallback((key) => isSuperAdmin || permissions.includes(key), [isSuperAdmin, permissions]);
  const canAny = useCallback((...keys) => isSuperAdmin || keys.some(k => permissions.includes(k)), [isSuperAdmin, permissions]);

  // Keep URL hash in sync so refresh restores the current view
  useEffect(() => { window.location.hash = view; }, [view]);

  // Track screen size and reset sidebar state on breakpoint change
  useEffect(() => {
    function handleResize() {
      const mobile = window.innerWidth < 768;
      setIsMobile(mobile);
      if (!mobile) setSidebarOpen(false);
    }
    window.addEventListener("resize", handleResize);
    return () => window.removeEventListener("resize", handleResize);
  }, []);

  // Only show the loading screen after 500ms — invisible on fast connections
  useEffect(() => {
    if (ready) { setShowLoader(false); return; }
    setShowLoader(false);
    const t = setTimeout(() => setShowLoader(true), 500);
    return () => clearTimeout(t);
  }, [authKey, ready]);

  useEffect(() => {
    let attempts = 0;
    let cancelled = false;
    setAuthError(false);
    function tryAuth() {
      if (cancelled) return;
      attempts++;
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), 8000);
      fetch(`${apiUrl}/auth/me`, { credentials: "include", signal: controller.signal })
        .then(r => {
          clearTimeout(timer);
          if (r.status === 401 || r.status === 403) { window.location.replace("/login.html"); return null; }
          if (!r.ok) throw new Error("server_error");
          return r.json();
        })
        .then(data => {
          if (!data || cancelled) return;
          setCurrentUser(data.username);
          setHasAvatar(!!data.hasAvatar);
          if (data.avatarDataUrl) setAvatarBlobUrl(data.avatarDataUrl);
          setIsAdmin(!!data.isAdmin);
          setIsSuperAdmin(!!data.isSuperAdmin);
          const perms = Array.isArray(data.permissions) ? data.permissions : [];
          setPermissions(perms);
          setUserRole(data.role ?? null);
          if (!data.isSuperAdmin && perms.length === 0 && data.role) setView("employee-dashboard");
          setReady(true);
        })
        .catch(() => {
          clearTimeout(timer);
          if (cancelled) return;
          if (attempts < 4) setTimeout(tryAuth, 3000);
          else setAuthError(true);
        });
    }
    tryAuth();
    return () => { cancelled = true; };
  }, [authKey]);

  // Fetch deduction/earning types and signatories on login and whenever SSE/polling fires a refresh
  useEffect(() => {
    if (!ready) return;
    fetch(`${apiUrl}/settings/deduction-types`, { credentials: "include" })
      .then(r => r.ok ? r.json() : []).then(setDedTypes).catch(() => {});
    fetch(`${apiUrl}/settings/earning-types`, { credentials: "include" })
      .then(r => r.ok ? r.json() : []).then(setEarnTypes).catch(() => {});
    fetch(`${apiUrl}/settings/signatories`, { credentials: "include" })
      .then(r => r.ok ? r.json() : null).then(setSignatories).catch(() => {});
  }, [ready, refreshKey]);

  // Leave pop-ups: read-only fetch on login, on every refresh, and on each
  // "leave-notify" event. Marking one read only happens when the user clicks.
  useEffect(() => {
    if (!ready) return;
    fetch(`${apiUrl}/leave-notifications`, { credentials: "include" })
      .then(r => r.ok ? r.json() : []).then(d => setLeaveNotes(Array.isArray(d) ? d : [])).catch(() => {});
  }, [ready, refreshKey, notifKey]);

  const dismissLeaveNote = useCallback((id) => {
    setLeaveNotes(ns => ns.filter(n => n.id !== id));
    fetch(`${apiUrl}/leave-notifications/${id}/read`, { method: "POST", credentials: "include" }).catch(() => {});
  }, []);
  const dismissAllLeaveNotes = useCallback(() => {
    setLeaveNotes([]);
    fetch(`${apiUrl}/leave-notifications/read-all`, { method: "POST", credentials: "include" }).catch(() => {});
  }, []);
  const openLeaveNote = useCallback((n) => {
    dismissLeaveNote(n.id);
    setLeaveFocus(n.leave_id);
    setView("leave");
    setSidebarOpen(false);
  }, [dismissLeaveNote]);

  useEffect(() => {
    if (!ready) return;
    let es, retryTimeout, pollInterval, debounceTimer;
    function triggerRefresh() {
      clearTimeout(debounceTimer);
      debounceTimer = setTimeout(() => setRefreshKey(k => k + 1), 500);
    }
    function stopPolling() { clearInterval(pollInterval); pollInterval = null; }
    function connect() {
      es = new EventSource(`${apiUrl}/events`, { withCredentials: true });
      es.addEventListener("connected", () => stopPolling());
      es.addEventListener("refresh", triggerRefresh);
      es.addEventListener("leave-notify", () => setNotifKey(k => k + 1));
      es.onerror = () => {
        es.close();
        if (!pollInterval) pollInterval = setInterval(triggerRefresh, 30_000);
        retryTimeout = setTimeout(connect, 30_000);
      };
    }
    connect();
    return () => {
      es?.close();
      clearTimeout(retryTimeout);
      clearTimeout(debounceTimer);
      stopPolling();
    };
  }, [ready]);

  async function handleLogout() {
    await fetch(`${apiUrl}/auth/logout`, { method: "POST", credentials: "include" }).catch(() => {});
    window.location.replace("/login.html");
  }

  function handleNavSelect(id) {
    setView(id);
    // Close drawer on mobile after navigation
    if (isMobile) setSidebarOpen(false);
  }

  if (!ready) {
    if (!authError && !showLoader) return null;
    return (
      <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center",
        height: "100vh", background: "#f0ede8", gap: 16 }}>
        <div style={{ width: 72, height: 72, backgroundImage: "url('/PSA.png')", backgroundSize: "contain",
          backgroundRepeat: "no-repeat", backgroundPosition: "center", opacity: 0.75 }} />
        {authError ? (
          <>
            <div style={{ fontSize: "0.95rem", color: "#b02020", fontWeight: 600 }}>Cannot reach server</div>
            <div style={{ fontSize: "0.82rem", color: "#888", maxWidth: 260, textAlign: "center", lineHeight: 1.6 }}>
              The server is temporarily unavailable. Please wait and try again.
            </div>
            <button onClick={() => setAuthKey(k => k + 1)}
              style={{ marginTop: 8, padding: "10px 28px", background: "#1a3a5c", color: "#fff",
                border: "none", borderRadius: 6, cursor: "pointer", fontSize: "0.9rem", fontWeight: 600 }}>
              Try Again
            </button>
            <a href="/login.html" style={{ fontSize: "0.82rem", color: "#888", marginTop: 4 }}>Back to login</a>
          </>
        ) : (
          <>
            <div style={{ fontSize: "0.85rem", color: "#aaa", letterSpacing: "0.05em" }}>Connecting…</div>
            <div style={{ width: 160, height: 2, background: "#e0ddd5", borderRadius: 99, overflow: "hidden" }}>
              <div style={{ height: "100%", background: "#1a3a5c", borderRadius: 99,
                animation: "psa-bar 1.4s ease-in-out infinite" }} />
            </div>
            <style>{`@keyframes psa-bar{0%{width:0%;margin-left:0}50%{width:70%;margin-left:15%}100%{width:0%;margin-left:100%}}`}</style>
          </>
        )}
      </div>
    );
  }

  return (
    <TypesContext.Provider value={{ dedTypes, earnTypes, setDedTypes, setEarnTypes, signatories }}>
    <div className={`app-shell${!isMobile && sidebarCollapsed ? " sidebar-collapsed" : ""}`}>
      {/* Mobile overlay backdrop — clicking it closes the sidebar */}
      {isMobile && sidebarOpen && (
        <div className="sidebar-overlay" onClick={() => setSidebarOpen(false)} />
      )}
      <Sidebar
        collapsed={!isMobile && sidebarCollapsed}
        open={isMobile ? sidebarOpen : true}
        isMobile={isMobile}
        view={view}
        setView={handleNavSelect}
        isAdmin={isAdmin}
        can={can}
        canAny={canAny}
        hasRole={isSuperAdmin || !!userRole}
        badges={{ leave: leaveNotes.length }}
      />
      <div className="workspace">
        <TopBar
          view={view}
          currentUser={currentUser}
          hasAvatar={hasAvatar}
          avatarBlobUrl={avatarBlobUrl}
          sidebarCollapsed={sidebarCollapsed}
          isMobile={isMobile}
          sidebarOpen={sidebarOpen}
          onToggleSidebar={() => {
            if (isMobile) {
              setSidebarOpen(v => !v);
            } else {
              setSidebarCollapsed(v => !v);
            }
          }}
          onLogout={handleLogout}
          onGoSettings={() => { setView("settings"); if (isMobile) setSidebarOpen(false); }}
        />
        <div className="page-content">
          {/* No role assigned — show message to non-super-admin users with no role at all */}
          {!isSuperAdmin && !userRole && permissions.length === 0 ? (
            <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", minHeight: 320, textAlign: "center", padding: "60px 24px" }}>
              <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" style={{ marginBottom: 20 }}>
                <circle cx="12" cy="12" r="9"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
              </svg>
              <div style={{ fontSize: "1.1rem", fontWeight: 700, color: "var(--text)", marginBottom: 8 }}>No role assigned yet.</div>
              <div style={{ fontSize: "0.92rem", color: "var(--text-muted)", maxWidth: 380 }}>
                Your account is active but has not been assigned a role. Please contact your system administrator to have a role assigned before you can access the system.
              </div>
            </div>
          ) : (
            <>
              {view === "dashboard"  && <ViewErrorBoundary><DashboardView refreshKey={refreshKey} isSuperAdmin={isSuperAdmin} permissions={permissions} /></ViewErrorBoundary>}
              {view === "files-201"  && <ViewErrorBoundary><Files201View /></ViewErrorBoundary>}
              {view === "employees"  && <ViewErrorBoundary><EmployeeHubView isAdmin={isAdmin} isSuperAdmin={isSuperAdmin} permissions={permissions} refreshKey={refreshKey} /></ViewErrorBoundary>}
              {view === "employee-dashboard" && <ViewErrorBoundary><EmployeeStatusView refreshKey={refreshKey} hasAvatar={hasAvatar} avatarBlobUrl={avatarBlobUrl} /></ViewErrorBoundary>}
              {view === "adjustments" && <ViewErrorBoundary><AdjustmentsView refreshKey={refreshKey} /></ViewErrorBoundary>}
              {view === "salary-rank" && <ViewErrorBoundary><SalaryRankView /></ViewErrorBoundary>}
              {view === "positions"   && <ViewErrorBoundary><WorkforceSetupView isAdmin={isAdmin} permissions={permissions} /></ViewErrorBoundary>}
              {view === "leave"      && <ViewErrorBoundary><LeaveView currentUser={currentUser} isSuperAdmin={isSuperAdmin} permissions={permissions}
                refreshSignal={notifKey} focusId={leaveFocus} onFocusDone={() => setLeaveFocus(null)} /></ViewErrorBoundary>}
              {view === "archive"    && <ViewErrorBoundary><ArchiveView isSuperAdmin={isSuperAdmin} permissions={permissions} /></ViewErrorBoundary>}
              {view === "settings"   && <ViewErrorBoundary><SettingsView currentUser={currentUser} isAdmin={isAdmin} isSuperAdmin={isSuperAdmin} permissions={permissions} hasAvatar={hasAvatar} avatarBlobUrl={avatarBlobUrl} setAvatarBlobUrl={setAvatarBlobUrl} setCurrentUser={setCurrentUser} setHasAvatar={setHasAvatar} refreshKey={refreshKey} /></ViewErrorBoundary>}
            </>
          )}
        </div>
      </div>
      <LeaveToasts notes={leaveNotes} onOpen={openLeaveNote} onDismiss={dismissLeaveNote} onDismissAll={dismissAllLeaveNotes} />
    </div>
    </TypesContext.Provider>
  );
}

// Leave pop-ups (top-right on every page): "waiting for your action" for the
// next officer, progress and the final result for the applicant.
const LEAVE_NOTE_META = {
  action:      { title: "Needs your action", tone: "is-action" },
  info:        { title: "Leave update",      tone: "is-info" },
  approved:    { title: "Leave approved",    tone: "is-good" },
  disapproved: { title: "Leave disapproved", tone: "is-bad" },
};

function LeaveToasts({ notes, onOpen, onDismiss, onDismissAll }) {
  if (!notes.length) return null;
  return (
    <div className="lt-stack" role="region" aria-label="Leave notifications" aria-live="polite">
      {notes.slice(0, 3).map(n => {
        const meta = LEAVE_NOTE_META[n.kind] ?? LEAVE_NOTE_META.info;
        return (
          <div key={n.id} className={`lt-toast ${meta.tone}`}>
            <div className="lt-body">
              <div className="lt-title">{meta.title}</div>
              <div className="lt-msg">{n.message}</div>
              <div className="lt-actions">
                <button type="button" className="lt-open" onClick={() => onOpen(n)}>{n.kind === "action" ? "Review" : "Open"}</button>
                <button type="button" className="um-btn um-btn--neutral um-btn--sm" onClick={() => onDismiss(n.id)}>Dismiss</button>
              </div>
            </div>
            <button type="button" className="lt-x" aria-label="Dismiss" onClick={() => onDismiss(n.id)}>✕</button>
          </div>
        );
      })}
      {notes.length > 3 && (
        <div className="lt-more">
          <span>+{notes.length - 3} more</span>
          <button type="button" className="um-btn um-btn--neutral um-btn--sm" onClick={onDismissAll}>Dismiss all</button>
        </div>
      )}
    </div>
  );
}

createRoot(document.getElementById("root")).render(
  <ViewErrorBoundary>
    <App />
  </ViewErrorBoundary>
);
