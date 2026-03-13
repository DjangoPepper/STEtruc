import { useState, useCallback, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import type { PointageData } from "./types";
import { useWindowWidth } from "../tabs/useWindowWidth";

type CellValue = string | number | boolean | null;
type RawData = CellValue[][];

// ─── Palette navy STEtruc ──────────────────────────────────────────────────
const T = {
  bg:         "#0D1B2E",
  bgCard:     "#1A2535",
  bgDark:     "#0A1628",
  border:     "#1E3A5F",
  border2:    "#2D3F55",
  accent:     "#38BDF8",
  accentDim:  "#38BDF855",
  success:    "#10B981",
  warning:    "#F59E0B",
  error:      "#EF4444",
  text:       "#F0F9FF",
  textMuted:  "#94A3B8",
  textDim:    "#64748B",
  repeat:     "#7C3AED",
  repeatBg:   "#1A1030",
  rowHover:   "#1E2D3E",
  selRowBg:   "#1E3A5F",
  selRowTxt:  "#38BDF8",
} as const;

interface ExcelCleanerProps {
  dark?: boolean;
  onDarkToggle?: () => void;
  onSendToPointage: (data: PointageData) => void;
}

interface ParsedData {
  headers: string[];
  rows: CellValue[][];
  headerRowIndex: number;
}

function detectHeaders(data: RawData): ParsedData {
  if (data.length === 0) return { headers: [], rows: [], headerRowIndex: 0 };
  let headerRowIndex = 0;
  for (let i = 0; i < Math.min(5, data.length); i++) {
    const row = data[i];
    const stringCount = row.filter((c) => typeof c === "string" && c.trim() !== "").length;
    const filledCount = row.filter((c) => c !== null && c !== undefined && c !== "").length;
    if (filledCount > 0 && stringCount / filledCount > 0.6) { headerRowIndex = i; break; }
  }
  const headers = data[headerRowIndex].map((h, i) =>
    h !== null && h !== undefined && h !== "" ? String(h) : `Colonne ${i + 1}`
  );
  return { headers, rows: data.slice(headerRowIndex + 1), headerRowIndex };
}

function computeRepetitiveValues(rows: RawData, colIndex: number, threshold = 0.35): Set<string> {
  const counts: Record<string, number> = {};
  let total = 0;
  for (const row of rows) {
    const val = row[colIndex];
    if (val !== null && val !== undefined && val !== "") {
      const key = String(val); counts[key] = (counts[key] || 0) + 1; total++;
    }
  }
  const result = new Set<string>();
  if (total === 0) return result;
  for (const [val, count] of Object.entries(counts))
    if (count / total >= threshold && count > 1) result.add(val);
  return result;
}

// ─── Btn ──────────────────────────────────────────────────────────────────
function Btn({
  children, onClick, variant = "default", small, title, style: extraStyle,
}: {
  children: React.ReactNode;
  onClick?: (e?: React.MouseEvent<HTMLButtonElement>) => void;
  variant?: "default" | "accent" | "success" | "danger" | "ghost";
  small?: boolean;
  title?: string;
  style?: React.CSSProperties;
}) {
  const [hov, setHov] = useState(false);
  const base: React.CSSProperties = {
    fontFamily: "'IBM Plex Mono', monospace",
    fontSize: small ? 10 : 12,
    fontWeight: 600,
    letterSpacing: "0.08em",
    textTransform: "uppercase",
    padding: small ? "5px 10px" : "9px 16px",
    border: "1px solid",
    borderRadius: 4,
    cursor: "pointer",
    transition: "all 0.15s",
    whiteSpace: "nowrap",
    ...extraStyle,
  };
  const variants: Record<string, React.CSSProperties> = {
    default: {
      background: hov ? T.bgCard : T.bgDark,
      borderColor: hov ? T.border2 : T.border,
      color: hov ? T.text : T.textMuted,
    },
    accent: {
      background: hov ? "#0ea5e9" : T.accent,
      borderColor: T.accent,
      color: "#0D1B2E",
    },
    success: {
      background: hov ? "#059669" : "#0f2a20",
      borderColor: T.success,
      color: T.success,
    },
    danger: {
      background: hov ? "#1f0a0a" : "transparent",
      borderColor: hov ? T.error : "#7f1d1d",
      color: hov ? T.error : "#fca5a5",
    },
    ghost: {
      background: "transparent",
      borderColor: "transparent",
      color: hov ? T.text : T.textMuted,
    },
  };
  return (
    <button
      style={{ ...base, ...variants[variant] }}
      onClick={(e) => {
        e.stopPropagation();
        onClick?.(e);
      }}
      onMouseEnter={() => setHov(true)}
      onMouseLeave={() => setHov(false)}
      title={title}
    >
      {children}
    </button>
  );
}

// ─── AddRowModal ──────────────────────────────────────────────────────────
function AddRowModal({
  headers, hiddenCols, onClose, onAdd,
}: {
  headers: string[];
  hiddenCols: Set<number>;
  onClose: () => void;
  onAdd: (row: CellValue[]) => void;
}) {
  const [values, setValues] = useState<string[]>(headers.map(() => ""));

  const handleAdd = () => {
    onAdd(values.map((v) => (v.trim() === "" ? null : v.trim())));
    onClose();
  };

  const visibleCols = headers.map((h, i) => ({ h, i })).filter(({ i }) => !hiddenCols.has(i));

  return (
    <div
      style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.75)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100 }}
      onClick={onClose}
    >
      <div
        style={{ background: T.bgCard, border: `1px solid ${T.border2}`, borderRadius: 8, padding: 28, minWidth: "min(360px, calc(100vw - 32px))", maxWidth: 600, maxHeight: "80vh", overflowY: "auto", boxShadow: "0 24px 60px rgba(0,0,0,0.6)" }}
        onClick={(e) => e.stopPropagation()}
      >
        {/* Title */}
        <div style={{ fontSize: 9, letterSpacing: "0.2em", textTransform: "uppercase", color: T.textDim, marginBottom: 4 }}>IEC</div>
        <div style={{ fontSize: 16, fontWeight: 700, color: T.text, marginBottom: 20 }}>
          + Ajouter une ligne
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {visibleCols.map(({ h, i }) => (
            <div key={i} style={{ display: "flex", flexDirection: "column", gap: 4 }}>
              <label style={{ fontSize: 9, letterSpacing: "0.12em", textTransform: "uppercase", color: T.textDim, fontFamily: "'IBM Plex Mono', monospace" }}>{h}</label>
              <input
                value={values[i]}
                onChange={(e) => { const next = [...values]; next[i] = e.target.value; setValues(next); }}
                onKeyDown={(e) => { if (e.key === "Enter") handleAdd(); if (e.key === "Escape") onClose(); }}
                style={{
                  fontFamily: "'IBM Plex Mono', monospace", fontSize: 13,
                  padding: "9px 12px",
                  background: T.bgDark,
                  border: `1px solid ${T.border}`,
                  borderRadius: 4,
                  color: T.text,
                  outline: "none",
                  transition: "border-color 0.15s",
                }}
                onFocus={(e) => { e.target.style.borderColor = T.accent; }}
                onBlur={(e) => { e.target.style.borderColor = T.border; }}
                placeholder={`Valeur pour ${h}…`}
              />
            </div>
          ))}
        </div>

        <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 24 }}>
          <Btn onClick={onClose} variant="ghost">Annuler</Btn>
          <Btn onClick={handleAdd} variant="success">Ajouter</Btn>
        </div>
      </div>
    </div>
  );
}

// ─── ExcelCleaner ─────────────────────────────────────────────────────────
export default function ExcelCleaner({ onSendToPointage }: ExcelCleanerProps) {
  const vw       = useWindowWidth();
  const isMobile = vw < 640;

  const [fileName, setFileName]       = useState<string | null>(null);
  const [workbook, setWorkbook]       = useState<XLSX.WorkBook | null>(null);
  const [sheetNames, setSheetNames]   = useState<string[]>([]);
  const [activeSheet, setActiveSheet] = useState<string | null>(null);
  const [parsed, setParsed]           = useState<ParsedData | null>(null);
  const [hiddenCols, setHiddenCols]   = useState<Set<number>>(new Set());
  const [hiddenRows, setHiddenRows]   = useState<Set<number>>(new Set());
  const [editingHeader, setEditingHeader] = useState<number | null>(null);
  const [headers, setHeaders]         = useState<string[]>([]);
  const [addedRows, setAddedRows]     = useState<CellValue[][]>([]);
  const [selectMode, setSelectMode]   = useState<"none" | "col" | "row">("none");
  const [selectedItems, setSelectedItems] = useState<Set<number>>(new Set());
  const [dimRepeated, setDimRepeated] = useState(true);
  const [exportFileName, setExportFileName] = useState("données_nettoyées");
  const [showAddRow, setShowAddRow]   = useState(false);
  const [hiddenSheets, setHiddenSheets] = useState<Set<string>>(new Set());
  const [sheetSelectMode, setSheetSelectMode] = useState<"none" | "delete" | "keep">("none");
  const [selectedSheets, setSelectedSheets] = useState<Set<string>>(new Set());
  const fileInputRef = useRef<HTMLInputElement>(null);

  const allRows = useMemo(() => {
    if (!parsed) return [];
    return [...addedRows, ...parsed.rows];
  }, [parsed, addedRows]);

  const loadSheet = useCallback((wb: XLSX.WorkBook, sheetName: string) => {
    const sheet = wb.Sheets[sheetName];
    const raw: RawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null }) as RawData;
    const result = detectHeaders(raw);
    setParsed(result);
    setHeaders(result.headers);
    setHiddenCols(new Set()); setHiddenRows(new Set());
    setSelectMode("none"); setSelectedItems(new Set());
    setEditingHeader(null); setAddedRows([]);
  }, []);

  const handleFile = useCallback((file: File) => {
    setFileName(file.name);
    const baseName = file.name.replace(/\.[^/.]+$/, "");
    setExportFileName(`${baseName}_cleaned`);
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer);
      const wb = XLSX.read(data, { type: "array" });
      setWorkbook(wb); setSheetNames(wb.SheetNames);
      setActiveSheet(wb.SheetNames[0]);
      setHiddenSheets(new Set());
      setSheetSelectMode("none"); setSelectedSheets(new Set());
      loadSheet(wb, wb.SheetNames[0]);
    };
    reader.readAsArrayBuffer(file);
  }, [loadSheet]);

  const handleSheetChange = (name: string) => {
    if (!workbook) return;
    setActiveSheet(name); loadSheet(workbook, name);
  };

  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  }, [handleFile]);

  const toggleSelectItem = (idx: number) => {
    setSelectedItems((prev) => {
      const next = new Set(prev);
      next.has(idx) ? next.delete(idx) : next.add(idx); return next;
    });
  };

  const repetitiveByCol = useMemo(() => {
    if (!parsed || !dimRepeated) return new Map<number, Set<string>>();
    const map = new Map<number, Set<string>>();
    headers.forEach((_, ci) => map.set(ci, computeRepetitiveValues(allRows, ci)));
    return map;
  }, [parsed, headers, dimRepeated, allRows]);

  const applyColAction = () => {
    if (selectMode !== "col") return;
    setHiddenCols((prev) => new Set([...prev, ...selectedItems]));
    setSelectedItems(new Set()); setSelectMode("none");
  };

  const applyRowDeletion = () => {
    if (selectMode !== "row") return;
    setHiddenRows((prev) => new Set([...prev, ...selectedItems]));
    setSelectedItems(new Set()); setSelectMode("none");
  };

  const exportClean = () => {
    if (!workbook) return;
    const wb2 = XLSX.utils.book_new();
    const visibleSheets = sheetNames.filter((n) => !hiddenSheets.has(n));
    visibleSheets.forEach((sheetName) => {
      if (sheetName === activeSheet && parsed) {
        const visibleHeaders = headers.filter((_, i) => !hiddenCols.has(i));
        const visibleRows = allRows
          .filter((_, i) => !hiddenRows.has(i))
          .map((row) => {
            const padded = [...row];
            while (padded.length < headers.length) padded.push(null);
            return padded.filter((_, ci) => !hiddenCols.has(ci));
          });
        const ws = XLSX.utils.aoa_to_sheet([visibleHeaders, ...visibleRows]);
        XLSX.utils.book_append_sheet(wb2, ws, sheetName);
      } else {
        XLSX.utils.book_append_sheet(wb2, workbook.Sheets[sheetName], sheetName);
      }
    });
    if (wb2.SheetNames.length === 0) return;
    XLSX.writeFile(wb2, `${exportFileName.trim() || "données_nettoyées"}.xlsx`);
  };

  const applySheetAction = () => {
    if (sheetSelectMode === "delete") {
      const next = new Set(hiddenSheets);
      selectedSheets.forEach((n) => next.add(n));
      setHiddenSheets(next);
      if (selectedSheets.has(activeSheet ?? "")) {
        const first = sheetNames.find((n) => !next.has(n));
        if (first && workbook) { setActiveSheet(first); loadSheet(workbook, first); }
      }
    } else {
      const next = new Set<string>(sheetNames.filter((n) => !selectedSheets.has(n)));
      setHiddenSheets(next);
      if (next.has(activeSheet ?? "")) {
        const first = sheetNames.find((n) => !next.has(n));
        if (first && workbook) { setActiveSheet(first); loadSheet(workbook, first); }
      }
    }
    setSelectedSheets(new Set()); setSheetSelectMode("none");
  };

  const visibleColCount = parsed ? headers.filter((_, i) => !hiddenCols.has(i)).length : 0;
  const visibleRowCount = allRows.filter((_, i) => !hiddenRows.has(i)).length;

  // ─── CSS ────────────────────────────────────────────────────────────────
  const css = `
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@300;400;500;600&family=Space+Grotesk:wght@300;400;500;600;700&display=swap');
    * { box-sizing: border-box; margin: 0; padding: 0; }

    ::-webkit-scrollbar { width: 4px; height: 4px; }
    ::-webkit-scrollbar-track { background: transparent; }
    ::-webkit-scrollbar-thumb { background: ${T.border2}; border-radius: 2px; }

    .iec-drop {
      border: 1.5px dashed ${T.border2}; border-radius: 8px; padding: 60px 40px;
      text-align: center; cursor: pointer; transition: all 0.2s; background: ${T.bgCard};
    }
    .iec-drop:hover { border-color: ${T.accent}; background: ${T.bgDark}; }

    .iec-sheet-tabs { display: flex; border-bottom: 1px solid ${T.border}; overflow-x: auto; margin-bottom: 0; }
    .iec-tab {
      font-family: 'IBM Plex Mono', monospace; font-size: 11px; padding: 9px 18px;
      cursor: pointer; border: 1px solid transparent; border-bottom: none;
      border-radius: 4px 4px 0 0; color: ${T.textMuted}; background: transparent;
      transition: all 0.15s; white-space: nowrap; letter-spacing: 0.05em;
      position: relative; top: 1px;
    }
    .iec-tab:hover  { color: ${T.text}; background: ${T.bgCard}; border-color: ${T.border}; border-bottom-color: transparent; }
    .iec-tab.active { color: ${T.accent}; background: ${T.bg}; border-color: ${T.border2}; border-bottom-color: ${T.bg}; }
    .iec-tab.hidden { opacity: 0.3; text-decoration: line-through; }
    .iec-tab.sel-del  { color: #fca5a5 !important; background: #2a0e0e !important; border-color: ${T.error} !important; border-bottom-color: transparent !important; }
    .iec-tab.sel-keep { color: ${T.success} !important; background: #0f2a20 !important; border-color: ${T.success} !important; border-bottom-color: transparent !important; }

    .iec-table { border-collapse: collapse; width: 100%; font-size: 12px; }
    .iec-table thead tr th {
      background: ${T.bgDark}; color: ${T.textDim}; font-weight: 600;
      font-size: 10px; letter-spacing: 0.12em; text-transform: uppercase;
      padding: 10px 14px; text-align: left; border-bottom: 1px solid ${T.border};
      white-space: nowrap; position: sticky; top: 0; z-index: 2;
      font-family: 'IBM Plex Mono', monospace;
    }
    .iec-table th.col-sel { cursor: pointer; }
    .iec-table th.col-sel:hover { background: ${T.selRowBg}; color: ${T.accent}; }
    .iec-table th.col-selected { background: #1f0a0a !important; color: #fca5a5 !important; border-bottom-color: ${T.error} !important; }

    .iec-table td {
      padding: 8px 14px; border-bottom: 1px solid ${T.border};
      color: ${T.textMuted}; white-space: nowrap;
      max-width: 220px; overflow: hidden; text-overflow: ellipsis;
      font-family: 'IBM Plex Mono', monospace; font-size: 12px;
    }
    .iec-table tr:hover td { background: ${T.rowHover}; }
    .iec-table tr.row-selected td { background: ${T.selRowBg} !important; color: ${T.selRowTxt}; }
    .iec-table tr.added-row td { background: #0a1f10 !important; }
    .iec-table tr.added-row:hover td { background: #0f2a18 !important; }

    .td-rownum {
      color: ${T.textDim} !important; font-size: 11px !important;
      background: ${T.bgDark} !important;
      border-right: 1px solid ${T.border} !important;
      min-width: 46px; width: 46px; text-align: center !important;
      user-select: none; font-variant-numeric: tabular-nums;
    }
    .th-rownum {
      background: ${T.bgDark} !important;
      border-right: 1px solid ${T.border} !important;
      width: 46px; min-width: 46px; text-align: center !important;
      color: ${T.textDim} !important;
    }
    .iec-table tr:hover .td-rownum { background: ${T.bgCard} !important; }
    .iec-table tr.row-selected .td-rownum { background: #122a40 !important; color: ${T.accent} !important; }
    .td-rownum.row-sel { cursor: pointer; }
    .td-rownum.row-sel:hover { color: ${T.accent} !important; background: ${T.selRowBg} !important; }

    .cell-repeated { color: ${T.repeat} !important; font-style: italic; background: ${T.repeatBg}; }
    .iec-table tr:hover .cell-repeated { color: #a78bfa !important; background: #1f1540; }

    .header-label { display: flex; align-items: center; gap: 5px; cursor: pointer; }
    .header-label:hover .edit-pencil { opacity: 1; }
    .edit-pencil { opacity: 0; font-size: 9px; color: ${T.accent}; transition: opacity 0.15s; flex-shrink: 0; }

    .header-input {
      background: transparent; border: none; border-bottom: 1px solid ${T.accent};
      color: ${T.accent}; font-family: 'IBM Plex Mono', monospace;
      font-size: 10px; letter-spacing: 0.08em; text-transform: uppercase;
      width: 100%; min-width: 50px; outline: none; padding: 2px 0;
    }

    .info-bar {
      margin-bottom: 10px; padding: 9px 14px;
      background: ${T.selRowBg}; border: 1px solid ${T.border2};
      border-radius: 4px; font-size: 11px; color: ${T.accent};
      font-family: 'IBM Plex Mono', monospace; letter-spacing: 0.04em;
    }
    .info-bar.del { background: #1f0a0a; border-color: #7f1d1d; color: #fca5a5; }

    .badge-hidden {
      display: inline-block; font-size: 9px; padding: 2px 6px; border-radius: 2px;
      background: #2a0e0e; color: #fca5a5; margin-left: 4px;
      font-family: 'IBM Plex Mono', monospace;
    }

    .dim-chip {
      display: flex; align-items: center; gap: 6px;
      font-size: 10px; letter-spacing: 0.06em; text-transform: uppercase;
      color: ${T.textMuted}; cursor: pointer; padding: 7px 12px;
      border: 1px solid ${T.border}; border-radius: 4px; background: ${T.bgDark};
      transition: all 0.15s; user-select: none;
      font-family: 'IBM Plex Mono', monospace;
    }
    .dim-chip:hover { border-color: ${T.border2}; color: ${T.text}; }
    .dim-chip.on { color: #a78bfa; border-color: #3d2a6a; background: ${T.repeatBg}; }
    .dim-dot { width: 7px; height: 7px; border-radius: 50%; background: currentColor; flex-shrink: 0; }

    .export-wrap { display: flex; align-items: center; border: 1px solid ${T.border2}; border-radius: 4px; overflow: hidden; }
    .export-input {
      font-family: 'IBM Plex Mono', monospace; font-size: 11px; letter-spacing: 0.05em;
      background: ${T.bgCard}; color: ${T.accent};
      border: none; outline: none; padding: 8px 10px; width: 180px;
    }
    .export-ext {
      font-family: 'IBM Plex Mono', monospace; font-size: 11px;
      color: ${T.textDim}; background: ${T.bgCard};
      padding: 8px 6px 8px 0; pointer-events: none;
    }
    .export-btn {
      font-family: 'IBM Plex Mono', monospace; font-size: 11px; font-weight: 600;
      letter-spacing: 0.08em; text-transform: uppercase; padding: 8px 14px;
      border: none; border-left: 1px solid ${T.border2};
      cursor: pointer; background: ${T.bgDark}; color: ${T.success};
      transition: all 0.15s;
    }
    .export-btn:hover { background: #0f2a20; color: #6ee7b7; }

    @media (max-width: 639px) {
      .iec-tab { font-size: 14px !important; padding: 12px 16px !important; }
      .export-input { font-size: 15px !important; padding: 13px 10px !important; width: 130px !important; }
      .export-ext { font-size: 15px !important; padding: 13px 4px 13px 0 !important; }
      .export-btn { font-size: 15px !important; padding: 13px 14px !important; letter-spacing: 0 !important; text-transform: none !important; }
      .dim-chip { font-size: 13px !important; padding: 11px 14px !important; }
      .info-bar { font-size: 14px !important; padding: 12px 16px !important; }
      .iec-table { font-size: 14px !important; }
      .iec-table thead tr th { font-size: 12px !important; padding: 12px 14px !important; }
      .iec-table td { padding: 11px 14px !important; }
    }
  `;

  // ─── render ───────────────────────────────────────────────────────────────
  return (
    <div style={{
      fontFamily: "'Space Grotesk', sans-serif",
      minHeight: "100vh",
      background: T.bg,
      color: T.text,
    }}>
      <style>{css}</style>

      {showAddRow && parsed && (
        <AddRowModal
          headers={headers}
          hiddenCols={hiddenCols}
          onClose={() => setShowAddRow(false)}
          onAdd={(row) => {
            const padded = headers.map((_, i) => row[i] ?? null);
            setAddedRows((prev) => [padded, ...prev]);
          }}
        />
      )}

      {/* Page header */}
      <div style={{
        padding: "16px 16px 12px",
        background: T.bgDark,
        borderBottom: `1px solid ${T.border}`,
      }}>
        <div style={{ color: T.textDim, fontSize: 9, letterSpacing: "0.2em", textTransform: "uppercase", marginBottom: 4 }}>
          IEC
        </div>
        <h1 style={{ color: T.text, fontSize: 20, fontWeight: 700, fontFamily: "'Space Grotesk', sans-serif" }}>
          Excel <span style={{ color: T.accent }}>Cleaner</span>
        </h1>
        <div style={{ color: T.textMuted, fontSize: 11, marginTop: 2 }}>
          Nettoyez et exportez vos fichiers Excel
        </div>
      </div>

      {/* Stats bar (when file loaded) */}
      {parsed && (
        <div style={{
          padding: "10px 16px",
          background: T.bgDark,
          borderBottom: `1px solid ${T.border}`,
          display: "flex",
          alignItems: "center",
          gap: 16,
          flexWrap: "wrap",
        }}>
          <span style={{ fontSize: 11, color: T.textMuted, fontFamily: "'IBM Plex Mono', monospace" }}>
            📄 <span style={{ color: T.textMuted }}>{fileName}</span>
            {activeSheet && sheetNames.length > 1 && <span style={{ color: T.textDim, marginLeft: 6 }}>· {activeSheet}</span>}
          </span>
          <div style={{ flex: 1 }} />
          <span style={{ fontSize: 11, fontFamily: "'IBM Plex Mono', monospace", color: T.textMuted }}>
            <span style={{ color: T.accent, fontWeight: 700 }}>{visibleColCount}</span> col ·{" "}
            <span style={{ color: T.accent, fontWeight: 700 }}>{visibleRowCount}</span> lignes
            {sheetNames.length > 1 && (
              <span style={{ marginLeft: 10, color: T.textMuted }}>
                · <span style={{ color: T.accent, fontWeight: 700 }}>{sheetNames.length - hiddenSheets.size}</span>/{sheetNames.length} onglets
              </span>
            )}
            {(hiddenCols.size > 0 || hiddenRows.size > 0) && (
              <span style={{ marginLeft: 10, color: T.error }}>({hiddenCols.size + hiddenRows.size} masqués)</span>
            )}
            {addedRows.length > 0 && (
              <span style={{ marginLeft: 10, color: T.success }}>+{addedRows.length} ajoutée(s)</span>
            )}
          </span>
          {/* Actions */}
          <Btn
            variant="success"
            small
            title="Envoyer les données nettoyées vers l'onglet Pointage"
            onClick={() => {
              if (!parsed) return;
              const visHdrs = headers.filter((_, i) => !hiddenCols.has(i));
              const visRows = allRows
                .filter((_, i) => !hiddenRows.has(i))
                .map((row) => {
                  const padded = [...row];
                  while (padded.length < headers.length) padded.push(null);
                  return padded.filter((_, ci) => !hiddenCols.has(ci));
                });
              onSendToPointage({ headers: visHdrs, rows: visRows, fileName: fileName ?? "fichier" });
            }}
          >
            ↗ Pointage
          </Btn>
          <div className="export-wrap">
            <input
              className="export-input"
              value={exportFileName}
              onChange={(e) => setExportFileName(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter") exportClean(); }}
              spellCheck={false}
            />
            <span className="export-ext">.xlsx</span>
            <button className="export-btn" onClick={exportClean}>↓ Exporter</button>
          </div>
          <Btn
            variant="danger"
            small
            onClick={() => {
              setParsed(null); setFileName(null); setWorkbook(null);
              setSheetNames([]); setActiveSheet(null); setAddedRows([]);
              setHiddenSheets(new Set()); setSheetSelectMode("none"); setSelectedSheets(new Set());
            }}
          >
            ✕ Reset
          </Btn>
        </div>
      )}

      <div style={{ padding: isMobile ? "12px 10px" : "20px 16px", maxWidth: 960, margin: "0 auto" }}>

        {/* ── Drop zone ── */}
        {!parsed ? (
          <div
            className="iec-drop"
            onDrop={onDrop}
            onDragOver={(e) => e.preventDefault()}
            onClick={() => fileInputRef.current?.click()}
          >
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              style={{ display: "none" }}
              onChange={(e) => { const f = e.target.files?.[0]; if (f) handleFile(f); }}
            />
            <div style={{ fontSize: 36, marginBottom: 16, opacity: 0.15, color: T.accent }}>⊞</div>
            <div style={{ fontSize: 14, color: T.textMuted, marginBottom: 8, fontFamily: "'Space Grotesk', sans-serif", fontWeight: 500 }}>
              Glissez un fichier Excel ou cliquez pour sélectionner
            </div>
            <div style={{ fontSize: 10, color: T.textDim, letterSpacing: "0.12em", textTransform: "uppercase", fontFamily: "'IBM Plex Mono', monospace" }}>
              .xlsx · .xls · .csv
            </div>
          </div>
        ) : (
          <div>
            {/* ── Onglets (multi-sheet) ── */}
            {sheetNames.length > 1 && (
              <div style={{
                background: T.bgCard,
                border: `1px solid ${T.border}`,
                borderRadius: 8,
                padding: "12px 14px",
                marginBottom: 16,
              }}>
                {/* Section label */}
                <div style={{ fontSize: 9, letterSpacing: "0.15em", textTransform: "uppercase", color: T.textDim, marginBottom: 8, fontFamily: "'IBM Plex Mono', monospace" }}>
                  Onglets · {sheetNames.length - hiddenSheets.size}/{sheetNames.length} visibles
                </div>
                {/* Controls */}
                <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 8, flexWrap: "wrap" }}>
                  {sheetSelectMode === "none" ? (
                    <>
                      <Btn
                        small
                        variant="danger"
                        onClick={() => {
                          setSheetSelectMode("delete");
                          setSelectedSheets(new Set());
                        }}
                      >
                        ✕ Supprimer onglets
                      </Btn>
                      <Btn
                        small
                        variant="success"
                        onClick={() => {
                          setSheetSelectMode("keep");
                          setSelectedSheets(new Set());
                        }}
                      >
                        ✓ Conserver onglets
                      </Btn>
                      {hiddenSheets.size > 0 && (
                        <Btn small onClick={() => setHiddenSheets(new Set())}>Restaurer onglets</Btn>
                      )}
                    </>
                  ) : (
                    <>
                      <span style={{ fontSize: 11, color: sheetSelectMode === "delete" ? "#fca5a5" : T.success, letterSpacing: "0.06em", fontFamily: "'IBM Plex Mono', monospace" }}>
                        {sheetSelectMode === "delete" ? "→ Cliquez les onglets à supprimer" : "→ Cliquez les onglets à conserver"}
                      </span>
                      {selectedSheets.size > 0 && (
                        <Btn
                          small
                          variant={sheetSelectMode === "keep" ? "success" : "danger"}
                          onClick={applySheetAction}
                        >
                          {sheetSelectMode === "delete" ? `Supprimer ${selectedSheets.size}` : `Garder ${selectedSheets.size}`}
                        </Btn>
                      )}
                      <Btn
                        small
                        onClick={() => {
                          setSheetSelectMode("none");
                          setSelectedSheets(new Set());
                        }}
                      >
                        Annuler
                      </Btn>
                    </>
                  )}
                </div>
                {/* Tabs */}
                <div className="iec-sheet-tabs">
                  {sheetNames.map((name) => {
                    const isHidden = hiddenSheets.has(name);
                    const isActive = activeSheet === name;
                    const isSel = selectedSheets.has(name);
                    const inSel = sheetSelectMode !== "none";
                    let cls = "iec-tab";
                    if (isHidden) cls += " hidden";
                    if (isActive && !inSel) cls += " active";
                    if (isSel && sheetSelectMode === "delete") cls += " sel-del";
                    if (isSel && sheetSelectMode === "keep") cls += " sel-keep";
                    return (
                      <button
                        key={name}
                        className={cls}
                        onClick={() => {
                          if (inSel) {
                            setSelectedSheets((prev) => {
                              const next = new Set(prev);
                              next.has(name) ? next.delete(name) : next.add(name);
                              return next;
                            });
                          } else if (!isHidden) {
                            handleSheetChange(name);
                          }
                        }}
                        title={isHidden ? `${name} (masqué)` : name}
                      >
                        {name}
                        {isHidden && <span style={{ marginLeft: 4, fontSize: 8 }}>✕</span>}
                      </button>
                    );
                  })}
                </div>
              </div>
            )}

            {/* ── Toolbar ── */}
            <div style={{
              background: T.bgCard,
              border: `1px solid ${T.border}`,
              borderRadius: 8,
              padding: "10px 14px",
              marginBottom: 12,
              display: "flex",
              gap: 8,
              alignItems: "center",
              flexWrap: "wrap",
            }}>
              {/* dim chip */}
              <div className={`dim-chip ${dimRepeated ? "on" : ""}`} onClick={() => setDimRepeated((v) => !v)}>
                <span className="dim-dot" />Griser répétitions
              </div>

              <div style={{ flex: 1 }} />

              <Btn small onClick={() => setShowAddRow(true)}>+ Ligne</Btn>
              <Btn
                small
                variant={selectMode === "col" ? "accent" : "default"}
                onClick={() => { setSelectMode(selectMode === "col" ? "none" : "col"); setSelectedItems(new Set()); }}
              >
                {selectMode === "col" ? "✓ " : ""}Colonnes
              </Btn>
              <Btn
                small
                variant={selectMode === "row" ? "accent" : "default"}
                onClick={() => { setSelectMode(selectMode === "row" ? "none" : "row"); setSelectedItems(new Set()); }}
              >
                {selectMode === "row" ? "✓ " : ""}Lignes
              </Btn>
              {selectedItems.size > 0 && (
                <Btn
                  small
                  variant="danger"
                  onClick={selectMode === "col" ? applyColAction : applyRowDeletion}
                >
                  {`Supprimer ${selectedItems.size} ${selectMode === "col" ? "col." : "ligne(s)"}`}
                </Btn>
              )}
              {(hiddenCols.size > 0 || hiddenRows.size > 0) && (
                <Btn small onClick={() => { setHiddenCols(new Set()); setHiddenRows(new Set()); }}>Restaurer tout</Btn>
              )}
            </div>

            {/* Context hints */}
            {selectMode === "col" && (
              <div className="info-bar del">→ Cliquez sur les en-têtes de colonnes à supprimer, puis confirmez</div>
            )}
            {selectMode === "row" && (
              <div className="info-bar del">→ Cliquez sur les numéros de lignes à supprimer, puis confirmez</div>
            )}
            {selectMode === "none" && (
              <div style={{ marginBottom: 8, fontSize: 10, color: T.textDim, letterSpacing: "0.04em", fontFamily: "'IBM Plex Mono', monospace" }}>
                Cliquez sur un en-tête pour le renommer · Activez Colonnes ou Lignes pour sélectionner
              </div>
            )}

            {/* Table */}
            <div style={{
              overflow: "auto",
              border: `1px solid ${T.border}`,
              borderRadius: 8,
              maxHeight: "calc(100vh - 360px)",
              background: T.bgCard,
            }}>
              <table className="iec-table">
                <thead>
                  <tr>
                    <th className="th-rownum">#</th>
                    {headers.map((h, ci) => {
                      if (hiddenCols.has(ci)) return null;
                      const isSel = selectMode === "col" && selectedItems.has(ci);
                      const isEditing = editingHeader === ci;
                      const cls = [
                        selectMode === "col" ? "col-sel" : "",
                        isSel ? "col-selected" : "",
                      ].filter(Boolean).join(" ");
                      return (
                        <th
                          key={ci}
                          className={cls}
                          onClick={() => {
                            if (selectMode === "col") { toggleSelectItem(ci); return; }
                            if (selectMode === "none") setEditingHeader(ci);
                          }}
                        >
                          {isEditing ? (
                            <input
                              className="header-input"
                              value={headers[ci]}
                              autoFocus
                              onChange={(e) => { const next = [...headers]; next[ci] = e.target.value; setHeaders(next); }}
                              onBlur={() => setEditingHeader(null)}
                              onKeyDown={(e) => { if (e.key === "Enter" || e.key === "Escape") setEditingHeader(null); }}
                              onClick={(e) => e.stopPropagation()}
                            />
                          ) : (
                            <span className="header-label" title="Cliquer pour renommer">
                              {h}
                              {selectMode === "none" && <span className="edit-pencil">✎</span>}
                            </span>
                          )}
                        </th>
                      );
                    })}
                  </tr>
                </thead>
                <tbody>
                  {allRows.map((row, ri) => {
                    if (hiddenRows.has(ri)) return null;
                    const isRowSel = selectMode === "row" && selectedItems.has(ri);
                    const isAdded = ri < addedRows.length;
                    return (
                      <tr key={ri} className={[isRowSel ? "row-selected" : "", isAdded ? "added-row" : ""].filter(Boolean).join(" ")}>
                        <td
                          className={`td-rownum ${selectMode === "row" ? "row-sel" : ""}`}
                          onClick={() => selectMode === "row" && toggleSelectItem(ri)}
                          title={isAdded ? "Ligne ajoutée manuellement" : undefined}
                        >
                          {isAdded
                            ? <span style={{ color: T.success, fontSize: 9 }}>+{ri + 1}</span>
                            : ri + 1}
                        </td>
                        {headers.map((_, ci) => {
                          if (hiddenCols.has(ci)) return null;
                          const cell = row[ci] ?? null;
                          const strVal = cell !== null ? String(cell) : null;
                          const isRepeated = !isAdded && dimRepeated && strVal !== null && (repetitiveByCol.get(ci)?.has(strVal) ?? false);
                          return (
                            <td key={ci} title={strVal ?? ""} className={isRepeated ? "cell-repeated" : ""}>
                              {strVal ?? <span style={{ color: T.textDim }}>—</span>}
                            </td>
                          );
                        })}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            {/* Footer */}
            <div style={{ marginTop: 10, display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 8 }}>
              {(hiddenCols.size > 0 || hiddenRows.size > 0) && (
                <div style={{ fontSize: 11, color: T.textMuted, display: "flex", gap: 16, flexWrap: "wrap", fontFamily: "'IBM Plex Mono', monospace" }}>
                  {hiddenCols.size > 0 && (
                    <span>
                      Colonnes masquées :
                      {[...hiddenCols].map((i) => (
                        <span key={i} className="badge-hidden">{headers[i] || `col ${i + 1}`}</span>
                      ))}
                    </span>
                  )}
                  {hiddenRows.size > 0 && (
                    <span><span className="badge-hidden">{hiddenRows.size} ligne(s) masquée(s)</span></span>
                  )}
                </div>
              )}
              {dimRepeated && (
                <div style={{ fontSize: 10, color: "#a78bfa", letterSpacing: "0.04em", fontFamily: "'IBM Plex Mono', monospace" }}>
                  ◆ Valeurs grisées = ≥35 % des lignes de leur colonne
                </div>
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
