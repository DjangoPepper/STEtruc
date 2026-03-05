// ============================================================
// STEtruc — Nettoyeur Excel Portatif
// Visuel  : coil-deploy  (navy dark, mobile-first, bottom-nav)
// Fonctions: STEpi        (Excel import, multi-sheets, clean, export)
// Stack    : React 18 + TypeScript + Vite → GitHub Pages
// ============================================================

import {
  useState, useCallback, useRef, useMemo, useEffect,
  useContext, createContext,
} from "react";
import * as XLSX from "xlsx";

// ─────────────────────────────────────────────
// 1. TYPES
// ─────────────────────────────────────────────

type CellValue = string | number | boolean | null;
type RawData = CellValue[][];
type Tab = "import" | "tableau" | "rapport" | "export";
type SelectMode = "none" | "col" | "row" | "dest";
type SheetSelectMode = "none" | "delete" | "keep";

type Destination = { name: string; color: string; excludeFromReport?: boolean };

interface ParsedData {
  headers: string[];
  rows: RawData;
  headerRowIndex: number;
}

// État de nettoyage par onglet
interface SheetState {
  parsed: ParsedData;
  headers: string[];
  addedRows: RawData;
  hiddenCols: Set<number>;
  hiddenRows: Set<number>;
}

interface AppState {
  // navigation
  activeTab: Tab;
  setActiveTab: (t: Tab) => void;
  // toast
  toast: { msg: string; type: "success" | "error" | "info" } | null;
  showToast: (msg: string, type?: "success" | "error" | "info") => void;
  // file
  fileName: string | null;
  setFileName: (n: string | null) => void;
  workbook: XLSX.WorkBook | null;
  setWorkbook: (w: XLSX.WorkBook | null) => void;
  sheetNames: string[];
  setSheetNames: (s: string[]) => void;
  activeSheet: string | null;
  setActiveSheet: (s: string | null) => void;
  // parsed data
  parsed: ParsedData | null;
  setParsed: (p: ParsedData | null) => void;
  headers: string[];
  setHeaders: (h: string[]) => void;
  addedRows: RawData;
  setAddedRows: React.Dispatch<React.SetStateAction<RawData>>;
  // hidden sets
  hiddenCols: Set<number>;
  setHiddenCols: React.Dispatch<React.SetStateAction<Set<number>>>;
  hiddenRows: Set<number>;
  setHiddenRows: React.Dispatch<React.SetStateAction<Set<number>>>;
  hiddenSheets: Set<string>;
  setHiddenSheets: React.Dispatch<React.SetStateAction<Set<string>>>;
  // selection
  selectMode: SelectMode;
  setSelectMode: React.Dispatch<React.SetStateAction<SelectMode>>;
  selectedItems: Set<number>;
  setSelectedItems: React.Dispatch<React.SetStateAction<Set<number>>>;
  editingHeader: number | null;
  setEditingHeader: React.Dispatch<React.SetStateAction<number | null>>;
  // sheet select
  sheetSelectMode: SheetSelectMode;
  setSheetSelectMode: React.Dispatch<React.SetStateAction<SheetSelectMode>>;
  selectedSheets: Set<string>;
  setSelectedSheets: React.Dispatch<React.SetStateAction<Set<string>>>;
  // options
  dimRepeated: boolean;
  setDimRepeated: React.Dispatch<React.SetStateAction<boolean>>;
  exportFileName: string;
  setExportFileName: React.Dispatch<React.SetStateAction<string>>;
  // split/grouping formats keyed by header name
  splitFormats: Record<string, string>;
  setSplitFormats: React.Dispatch<React.SetStateAction<Record<string, string>>>;
  autoRefFmt: boolean;
  setAutoRefFmt: React.Dispatch<React.SetStateAction<boolean>>;
  poidsUnit: "t" | "kg";
  setPoidsUnit: React.Dispatch<React.SetStateAction<"t" | "kg">>;
  // mapping (colonnes REF/Rang/Poids/Destination) + extras
  mapping: { rang: string; reference: string; poids: string; dch: string };
  setMapping: React.Dispatch<React.SetStateAction<{ rang: string; reference: string; poids: string; dch: string }>>;
  extras: { col: string; label: string }[];
  setExtras: React.Dispatch<React.SetStateAction<{ col: string; label: string }[]>>;
  // pointage tracking
  pointedRows: Set<number>;
  setPointedRows: React.Dispatch<React.SetStateAction<Set<number>>>;
  rowOverrides: Map<number, Record<number, CellValue>>;
  setRowOverrides: React.Dispatch<React.SetStateAction<Map<number, Record<number, CellValue>>>>;
  // destination assignment (mouvements)
  destinations: Destination[];
  setDestinations: React.Dispatch<React.SetStateAction<Destination[]>>;
  selectedDest: string;
  setSelectedDest: React.Dispatch<React.SetStateAction<string>>;
  rowDestinations: Map<number, string>;
  setRowDestinations: React.Dispatch<React.SetStateAction<Map<number, string>>>;
  reassignedRows: Map<number, { from: string; to: string }[]>;
  setReassignedRows: React.Dispatch<React.SetStateAction<Map<number, { from: string; to: string }[]>>>;
  // winwin modal
  winwinModalOpen: boolean;
  setWinwinModalOpen: React.Dispatch<React.SetStateAction<boolean>>;
  // helpers
  loadSheet: (wb: XLSX.WorkBook, sheet: string) => void;
  handleFile: (file: File) => void;
  allRows: RawData;
  repetitiveByCol: Map<number, Set<string>>;
  sheetStates: React.MutableRefObject<Map<string, SheetState>>;
}

// ─────────────────────────────────────────────
// 2. UTILS
// ─────────────────────────────────────────────

function detectHeaders(data: RawData): ParsedData {
  if (data.length === 0) return { headers: [], rows: [], headerRowIndex: 0 };
  let headerRowIndex = 0;
  for (let i = 0; i < Math.min(5, data.length); i++) {
    const row = data[i];
    const stringCount = row.filter((c) => typeof c === "string" && String(c).trim() !== "").length;
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
      const key = String(val);
      counts[key] = (counts[key] ?? 0) + 1;
      total++;
    }
  }
  const result = new Set<string>();
  if (total === 0) return result;
  for (const [val, count] of Object.entries(counts))
    if (count / total >= threshold && count > 1) result.add(val);
  return result;
}


// ─────────────────────────────────────────────
// 3. CONTEXT
// ─────────────────────────────────────────────

const AppCtx = createContext<AppState | null>(null);
const useApp = () => {
  const ctx = useContext(AppCtx);
  if (!ctx) throw new Error("useApp outside provider");
  return ctx;
};

function AppProvider({ children }: { children: React.ReactNode }) {
  const [activeTab, setActiveTab] = useState<Tab>("import");
  const [toast, setToast] = useState<AppState["toast"]>(null);

  const [fileName, setFileName] = useState<string | null>(null);
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [activeSheet, setActiveSheet] = useState<string | null>(null);

  const [parsed, setParsed] = useState<ParsedData | null>(null);
  const [headers, setHeaders] = useState<string[]>([]);
  const [addedRows, setAddedRows] = useState<RawData>([]);

  const [hiddenCols, setHiddenCols] = useState<Set<number>>(new Set());
  const [hiddenRows, setHiddenRows] = useState<Set<number>>(new Set());
  const [hiddenSheets, setHiddenSheets] = useState<Set<string>>(new Set());

  const [selectMode, setSelectMode] = useState<SelectMode>("none");
  const [selectedItems, setSelectedItems] = useState<Set<number>>(new Set());
  const [editingHeader, setEditingHeader] = useState<number | null>(null);

  const [sheetSelectMode, setSheetSelectMode] = useState<SheetSelectMode>("none");
  const [selectedSheets, setSelectedSheets] = useState<Set<string>>(new Set());

  const [dimRepeated, setDimRepeated] = useState(true);
  const [exportFileName, setExportFileName] = useState("données_nettoyées");

  // ── Helpers localStorage ──────────────────────────────────────────────────
  function lsGet<T>(key: string, fallback: T): T {
    try { const v = localStorage.getItem(key); return v !== null ? (JSON.parse(v) as T) : fallback; } catch { return fallback; }
  }
  function lsSet(key: string, val: unknown) { try { localStorage.setItem(key, JSON.stringify(val)); } catch {} }

  const [splitFormats, setSplitFormats] = useState<Record<string, string>>(() => lsGet("ste_splitFormats", {}));
  const [autoRefFmt, setAutoRefFmt] = useState<boolean>(() => lsGet("ste_autoRefFmt", false));
  const [poidsUnit, setPoidsUnit] = useState<"t" | "kg">(() => lsGet("ste_poidsUnit", "t"));
  const [mapping, setMapping] = useState<{ rang: string; reference: string; poids: string; dch: string }>(() => lsGet("ste_mapping", { rang: "", reference: "", poids: "", dch: "" }));
  const [extras, setExtras] = useState<{ col: string; label: string }[]>(() => lsGet("ste_extras", []));
  const [pointedRows, setPointedRows] = useState<Set<number>>(new Set());
  const [rowOverrides, setRowOverrides] = useState<Map<number, Record<number, CellValue>>>(new Map());
  const defaultDestinations: Destination[] = [
    { name: "H1", color: "#00c87a" },
    { name: "H2", color: "#f447d1" },
    { name: "H3", color: "#3cbefc" },
    { name: "H4", color: "#ff9b2c" },
    { name: "HNE", color: "#94A3B8", excludeFromReport: true },
  ];
  const [destinations, setDestinations] = useState<Destination[]>(() => lsGet("ste_destinations", defaultDestinations));
  const [selectedDest, setSelectedDest] = useState<string>("");
  const [rowDestinations, setRowDestinations] = useState<Map<number, string>>(new Map());
  const [reassignedRows, setReassignedRows] = useState<Map<number, { from: string; to: string }[]>>(new Map());
  const [winwinModalOpen, setWinwinModalOpen] = useState(false);

  // Stockage de l'état de chaque onglet (persist entre changements d'onglet)
  const sheetStates = useRef<Map<string, SheetState>>(new Map());

  const showToast = useCallback((msg: string, type: "success" | "error" | "info" = "info") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3000);
  }, []);

  // ── Persistance localStorage des préférences utilisateur ──────────────────
  useEffect(() => { lsSet("ste_splitFormats", splitFormats); }, [splitFormats]);
  useEffect(() => { lsSet("ste_autoRefFmt", autoRefFmt); }, [autoRefFmt]);
  useEffect(() => { lsSet("ste_poidsUnit", poidsUnit); }, [poidsUnit]);
  useEffect(() => { lsSet("ste_mapping", mapping); }, [mapping]);
  useEffect(() => { lsSet("ste_extras", extras); }, [extras]);
  useEffect(() => { lsSet("ste_destinations", destinations); }, [destinations]);

  // loadSheet : sauvegarde l'onglet courant AVANT de charger le nouveau
  const loadSheet = useCallback((wb: XLSX.WorkBook, sheetName: string) => {
    // On lit les valeurs depuis les refs pour éviter les closures périmées
    // Elles sont mises à jour via useEffect ci-dessous
    const curSheet = activeSheetRef.current;
    if (curSheet && curSheet !== sheetName && parsedRef.current) {
      sheetStates.current.set(curSheet, {
        parsed:     parsedRef.current,
        headers:    headersRef.current,
        addedRows:  addedRowsRef.current,
        hiddenCols: hiddenColsRef.current,
        hiddenRows: hiddenRowsRef.current,
      });
    }

    // Restaurer l'état sauvegardé ou charger depuis le workbook
    const saved = sheetStates.current.get(sheetName);
    if (saved) {
      setParsed(saved.parsed);
      setHeaders(saved.headers);
      setAddedRows(saved.addedRows);
      setHiddenCols(saved.hiddenCols);
      setHiddenRows(saved.hiddenRows);
    } else {
      const sheet = wb.Sheets[sheetName];
      const raw: RawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null }) as RawData;
      const result = detectHeaders(raw);
      setParsed(result);
      setHeaders(result.headers);
      setHiddenCols(new Set());
      setHiddenRows(new Set());
      setAddedRows([]);
    }
    setSelectMode("none");
    setSelectedItems(new Set());
    setEditingHeader(null);
  }, []); // pas de deps : on utilise des refs pour lire l'état courant

  const handleFile = useCallback((file: File) => {
    sheetStates.current.clear(); // reset total
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer);
      const wb = XLSX.read(data, { type: "array" });
      setWorkbook(wb);
      setSheetNames(wb.SheetNames);
      // Auto-select "winwin" sheet if it exists
      const winwinSheet = wb.SheetNames.find(s => s.toLowerCase() === "winwin");
      const defaultSheet = winwinSheet ?? wb.SheetNames[0];
      if (winwinSheet) { setWinwinModalOpen(true); setAutoRefFmt(true); }
      setActiveSheet(defaultSheet);
      setHiddenSheets(winwinSheet ? new Set(wb.SheetNames.filter(s => s !== winwinSheet)) : new Set());
      setSheetSelectMode("none");
      setSelectedSheets(new Set());
      setPointedRows(new Set());
      setRowOverrides(new Map());
      setRowDestinations(new Map());
      setReassignedRows(new Map());
      setSplitFormats({});
      loadSheet(wb, defaultSheet);
    };
    reader.readAsArrayBuffer(file);
  }, [loadSheet]);

  // Refs miroirs pour accès synchrone depuis loadSheet (pas de closure périmée)
  const activeSheetRef  = useRef(activeSheet);
  const parsedRef       = useRef(parsed);
  const headersRef      = useRef(headers);
  const addedRowsRef    = useRef(addedRows);
  const hiddenColsRef   = useRef(hiddenCols);
  const hiddenRowsRef   = useRef(hiddenRows);

  // Mise à jour des refs à chaque render
  activeSheetRef.current  = activeSheet;
  parsedRef.current       = parsed;
  headersRef.current      = headers;
  addedRowsRef.current    = addedRows;
  hiddenColsRef.current   = hiddenCols;
  hiddenRowsRef.current   = hiddenRows;

  const allRows = useMemo(() => {
    if (!parsed) return [];
    return [...addedRows, ...parsed.rows];
  }, [parsed, addedRows]);

  const repetitiveByCol = useMemo(() => {
    if (!parsed || !dimRepeated) return new Map<number, Set<string>>();
    const map = new Map<number, Set<string>>();
    headers.forEach((_, ci) => map.set(ci, computeRepetitiveValues(allRows, ci)));
    return map;
  }, [parsed, headers, dimRepeated, allRows]);

  return (
    <AppCtx.Provider value={{
      activeTab, setActiveTab,
      toast, showToast,
      fileName, setFileName,
      workbook, setWorkbook,
      sheetNames, setSheetNames,
      activeSheet, setActiveSheet,
      parsed, setParsed,
      headers, setHeaders,
      addedRows, setAddedRows,
      hiddenCols, setHiddenCols,
      hiddenRows, setHiddenRows,
      hiddenSheets, setHiddenSheets,
      selectMode, setSelectMode,
      selectedItems, setSelectedItems,
      editingHeader, setEditingHeader,
      sheetSelectMode, setSheetSelectMode,
      selectedSheets, setSelectedSheets,
      dimRepeated, setDimRepeated,
      exportFileName, setExportFileName,
      splitFormats, setSplitFormats,
      autoRefFmt, setAutoRefFmt,
      poidsUnit, setPoidsUnit,
      mapping, setMapping,
      extras, setExtras,
      pointedRows, setPointedRows,
      rowOverrides, setRowOverrides,
      destinations, setDestinations,
      selectedDest, setSelectedDest,
      rowDestinations, setRowDestinations,
      reassignedRows, setReassignedRows,
      winwinModalOpen, setWinwinModalOpen,
      loadSheet, handleFile,
      allRows, repetitiveByCol,
      sheetStates,
    }}>
      {children}
    </AppCtx.Provider>
  );
}

// ─────────────────────────────────────────────
// 4. COMPOSANTS UI (style coil-deploy)
// ─────────────────────────────────────────────

// Palette navy (coil-deploy)
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

// Toast
function Toast() {
  const { toast } = useApp();
  if (!toast) return null;
  const bg = toast.type === "success" ? T.success : toast.type === "error" ? T.error : T.accent;
  return (
    <div style={{
      position: "fixed", top: 16, left: "50%", transform: "translateX(-50%)",
      background: bg, color: "#0F172A", padding: "10px 20px", borderRadius: 10,
      fontWeight: 700, fontSize: 14, zIndex: 9999,
      boxShadow: "0 4px 20px rgba(0,0,0,0.4)",
      maxWidth: "90vw", textAlign: "center", letterSpacing: 0.3,
      fontFamily: "'Share Tech Mono', monospace",
    }}>
      {toast.msg}
    </div>
  );
}

// Bottom Navigation
function BottomNav() {
  const { activeTab, setActiveTab, parsed } = useApp();
  const tabs: { id: Tab; icon: string; label: string }[] = [
    { id: "import",  icon: "⬇️",  label: "Import"  },
    { id: "tableau", icon: "📊",  label: "Pointage" },
    { id: "rapport", icon: "📋",  label: "Rapport"  },
    { id: "export",  icon: "⬆️",  label: "Export"  },
  ];
  return (
    <nav style={{
      position: "fixed", bottom: 0, left: 0, right: 0, height: 64,
      background: "#0F172A", borderTop: `2px solid ${T.border}`,
      display: "flex", alignItems: "stretch", zIndex: 100,
      maxWidth: 540, margin: "0 auto",
    }}>
      {tabs.map((t) => {
        const isActive = activeTab === t.id;
        const disabled = (t.id === "tableau" || t.id === "rapport" || t.id === "export") && !parsed;
        return (
          <button
            key={t.id}
            onClick={() => !disabled && setActiveTab(t.id)}
            style={{
              flex: 1, display: "flex", flexDirection: "column", alignItems: "center",
              justifyContent: "center", background: "none", border: "none",
              color: disabled ? T.textDim : isActive ? T.accent : T.textDim,
              fontSize: 11, fontWeight: 700, cursor: disabled ? "not-allowed" : "pointer",
              gap: 2, letterSpacing: 0.5,
              borderTop: isActive ? `3px solid ${T.accent}` : "3px solid transparent",
              transition: "all 0.15s",
              fontFamily: "'Share Tech Mono', monospace",
            }}
          >
            <span style={{ fontSize: 20 }}>{t.icon}</span>
            {t.label}
          </button>
        );
      })}
    </nav>
  );
}

// Header de page
function PageHeader({ title, subtitle }: { title: string; subtitle?: string }) {
  return (
    <div style={{
      padding: "16px 16px 12px",
      background: T.bgDark,
      borderBottom: `1px solid ${T.border}`,
    }}>
      <div style={{ color: T.textDim, fontSize: 9, letterSpacing: "0.2em", textTransform: "uppercase", marginBottom: 4 }}>
        STEtruc
      </div>
      <h1 style={{ color: T.text, fontSize: 20, fontWeight: 700, letterSpacing: "-0.01em" }}>
        {title}
      </h1>
      {subtitle && (
        <div style={{ color: T.textMuted, fontSize: 11, marginTop: 2 }}>{subtitle}</div>
      )}
    </div>
  );
}

// Bouton générique (style coil-deploy)
function Btn({
  children, onClick, color = T.accent, textColor = "#0F172A",
  small, disabled, fullWidth, danger,
}: {
  children: React.ReactNode;
  onClick?: () => void;
  color?: string;
  textColor?: string;
  small?: boolean;
  disabled?: boolean;
  fullWidth?: boolean;
  danger?: boolean;
}) {
  const bg = danger ? T.error : disabled ? T.border2 : color;
  const txt = danger ? "#fff" : disabled ? T.textDim : textColor;
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      style={{
        background: bg, color: txt,
        border: "none", borderRadius: 8,
        padding: small ? "7px 12px" : "11px 18px",
        fontWeight: 700, fontSize: small ? 11 : 13,
        cursor: disabled ? "not-allowed" : "pointer",
        width: fullWidth ? "100%" : undefined,
        letterSpacing: 0.3,
        transition: "all 0.15s",
        boxShadow: disabled ? "none" : `0 2px 8px ${bg}44`,
        fontFamily: "'Share Tech Mono', monospace",
        whiteSpace: "nowrap",
      }}
    >
      {children}
    </button>
  );
}


// Stats Bar compacte
function StatsBar() {
  const { parsed, addedRows, hiddenCols, hiddenRows, headers } = useApp();
  if (!parsed) return null;
  const visibleCols = headers.filter((_, i) => !hiddenCols.has(i)).length;
  const allRows = [...addedRows, ...parsed.rows];
  const visibleRows = allRows.filter((_, i) => !hiddenRows.has(i)).length;

  return (
    <div style={{
      display: "flex", alignItems: "center", gap: 0,
      background: T.bgDark, borderBottom: `1px solid ${T.border}`,
      padding: "3px 12px",
    }}>
      {[
        { label: "COL", value: visibleCols, color: T.accent },
        { label: "LIG", value: visibleRows, color: T.success },
      ].map((s, i) => (
        <span key={s.label} style={{ fontSize: 10, color: T.textDim, display: "flex", alignItems: "center", gap: 3 }}>
          {i > 0 && <span style={{ color: T.border2, margin: "0 6px" }}>·</span>}
          <span style={{ color: s.color, fontWeight: 900 }}>{s.value}</span>
          <span style={{ letterSpacing: "0.06em" }}> {s.label}</span>
        </span>
      ))}
    </div>
  );
}

// Infos fichier+stats en ligne (utilisé dans la barre unique de TablePage)
function PointageInfos() {
  const { fileName } = useApp();
  return (
    <>
      {fileName && (
        <>
          <span style={{ color: "#7C3AED99", fontSize: 12 }}>·</span>
          <span style={{ color: T.textMuted, fontSize: 11, fontFamily: "'Share Tech Mono', monospace", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: 100 }}>
            📄 {fileName}
          </span>
        </>
      )}
    </>
  );
}

// Section title
function SectionTitle({ icon, text, count, right }: {
  icon?: string; text: string; count?: number;
  right?: React.ReactNode;
}) {
  return (
    <div style={{
      display: "flex", alignItems: "center", justifyContent: "space-between",
      padding: "12px 16px 8px", borderBottom: `1px solid ${T.border}`,
    }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
        {icon && <span style={{ fontSize: 16 }}>{icon}</span>}
        <span style={{ color: T.text, fontWeight: 800, fontSize: 14 }}>{text}</span>
        {count !== undefined && (
          <span style={{
            background: T.border, color: T.accent, borderRadius: 99,
            fontSize: 11, fontWeight: 700, padding: "1px 8px", border: `1px solid ${T.accentDim}`,
          }}>{count}</span>
        )}
      </div>
      {right}
    </div>
  );
}

// Empty state
function EmptyState({ icon, text, sub }: { icon: string; text: string; sub?: string }) {
  return (
    <div style={{ textAlign: "center", padding: "48px 20px" }}>
      <div style={{ fontSize: 52, marginBottom: 12 }}>{icon}</div>
      <div style={{ color: T.textMuted, fontSize: 15, fontWeight: 700, marginBottom: 6 }}>{text}</div>
      {sub && <div style={{ color: T.textDim, fontSize: 12 }}>{sub}</div>}
    </div>
  );
}

// Modal ajout de ligne (adapté STEpi → style coil-deploy)
function AddRowModal({ onClose }: { onClose: () => void }) {
  const { headers, hiddenCols, setAddedRows, showToast } = useApp();
  const [values, setValues] = useState<string[]>(headers.map(() => ""));

  const visibleCols = headers.map((h, i) => ({ h, i })).filter(({ i }) => !hiddenCols.has(i));

  const handleAdd = () => {
    const row = headers.map((_, i) => {
      const v = values[i];
      return v.trim() === "" ? null : v.trim();
    });
    setAddedRows((prev) => [row, ...prev]);
    showToast("Ligne ajoutée", "success");
    onClose();
  };

  return (
    <div
      style={{
        position: "fixed", inset: 0, background: "rgba(0,0,0,0.7)",
        display: "flex", alignItems: "flex-end", zIndex: 200,
      }}
      onClick={onClose}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          width: "100%", maxWidth: 540, margin: "0 auto",
          background: T.bgCard, borderRadius: "18px 18px 0 0",
          padding: 20, paddingBottom: 90, maxHeight: "80vh", overflowY: "auto",
          border: `1px solid ${T.border}`,
        }}
      >
        <div style={{ width: 36, height: 4, background: T.border2, borderRadius: 99, margin: "0 auto 16px" }} />
        <div style={{ color: T.accent, fontSize: 12, letterSpacing: "0.12em", textTransform: "uppercase", marginBottom: 16, fontWeight: 700 }}>
          + Ajouter une ligne
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {visibleCols.map(({ h, i }) => (
            <div key={i}>
              <div style={{ color: T.textDim, fontSize: 10, marginBottom: 4, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.5 }}>
                {h}
              </div>
              <input
                value={values[i]}
                onChange={(e) => {
                  const next = [...values];
                  next[i] = e.target.value;
                  setValues(next);
                }}
                onKeyDown={(e) => {
                  if (e.key === "Enter") handleAdd();
                  if (e.key === "Escape") onClose();
                }}
                placeholder={`Valeur pour ${h}…`}
                style={{
                  width: "100%", background: T.bgDark, border: `1px solid ${T.border2}`,
                  borderRadius: 8, color: T.text, fontSize: 14, padding: "10px 12px",
                  outline: "none", fontFamily: "'Share Tech Mono', monospace",
                }}
                onFocus={(e) => { e.target.style.borderColor = T.accent; }}
                onBlur={(e) => { e.target.style.borderColor = T.border2; }}
              />
            </div>
          ))}
        </div>
        <div style={{ display: "flex", gap: 10, marginTop: 20 }}>
          <Btn onClick={onClose} color={T.border} textColor={T.textMuted} fullWidth>Annuler</Btn>
          <Btn onClick={handleAdd} color={T.success} textColor="#0F172A" fullWidth>✅ Ajouter</Btn>
        </div>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────
// 5. PAGE IMPORT
// ─────────────────────────────────────────────

// Helper: apply visual grouping pattern "4 4 1" to a string value
function applyGrouping(value: string, pattern: string): string {
  if (!pattern || !pattern.trim()) return value;
  const sizes = pattern.trim().split(/\s+/).map(Number).filter((n) => n > 0);
  if (sizes.length === 0) return value;
  const str = String(value);
  let result = "";
  let pos = 0;
  for (const size of sizes) {
    if (pos >= str.length) break;
    if (result) result += " ";
    result += str.slice(pos, pos + size);
    pos += size;
  }
  if (pos < str.length) result += (result ? " " : "") + str.slice(pos);
  return result;
}

// Helper: thousands separator (right-to-left groups of 3, preserves decimals)
function thsep(val: string): string {
  const [int, dec] = val.split(".");
  const separated = int.replace(/\B(?=(\d{3})+(?!\d))/g, " ");
  return dec !== undefined ? separated + "." + dec : separated;
}

// Helper: auto-format reference (groupement intelligent selon longueur/contenu)
function autoFormatRef(val: string, globalFmt: string): string {
  const str = val.trim();
  if (!str) return str;
  if (/^\d+$/.test(str)) {
    if (str.length === 8) return applyGrouping(str, "3 2 3");
    if (str.length === 7) return applyGrouping(str, "3 2 2");
    if (str.length === 6) return applyGrouping(str, "3 3");
    return globalFmt ? applyGrouping(str, globalFmt) : str;
  }
  const letters = [...str.matchAll(/[A-Za-z]/g)];
  if (letters.length === 1) {
    const idx = letters[0].index!;
    if (idx > 0 && idx < str.length - 1) {
      return str.slice(0, idx) + " " + str[idx].toUpperCase() + " " + str.slice(idx + 1);
    }
  }
  return globalFmt ? applyGrouping(str, globalFmt) : str;
}

function ImportPage() {
  const {
    handleFile, parsed, setParsed, fileName,
    sheetNames, activeSheet, setActiveSheet, workbook, loadSheet,
    showToast, setActiveTab,
    headers, setHeaders,
    hiddenCols, setHiddenCols,
    hiddenSheets, setHiddenSheets,
    splitFormats, setSplitFormats,
    mapping, setMapping,
    extras, setExtras,
    autoRefFmt, setAutoRefFmt,
    poidsUnit, setPoidsUnit,
    winwinModalOpen, setWinwinModalOpen,
  } = useApp();

  const fileInputRef = useRef<HTMLInputElement>(null);
  const [step, setStep] = useState<1 | 2 | 3>(1);
  const [editableRows, setEditableRows] = useState<Record<string, string>[]>([]);
  const [editingHdr, setEditingHdr] = useState<number | null>(null);
  // split formats: now from global context
  const [splitEditingField, setSplitEditingField] = useState<string | null>(null);
  const [splitInputValue, setSplitInputValue] = useState("");
  const [openOnglets,   setOpenOnglets]   = useState(false);
  const [openAtypiques, setOpenAtypiques] = useState(false);
  const [openHeaders,   setOpenHeaders]   = useState(false);
  const [openDonnees,   setOpenDonnees]   = useState(false);
  const [openMapping,   setOpenMapping]   = useState(false);
  const [openApercu,    setOpenApercu]    = useState(false);
  const suppressCollapseRef = useRef(false);

  // Auto-close winwin modal after 3s
  useEffect(() => {
    if (!winwinModalOpen) return;
    const t = setTimeout(() => setWinwinModalOpen(false), 3000);
    return () => clearTimeout(t);
  }, [winwinModalOpen]);

  // Re-initialize editableRows when the parsed data changes (new file or new sheet).
  // Headers intentionally excluded to avoid re-init when user renames columns.
  // eslint-disable-next-line react-hooks/exhaustive-deps
  useEffect(() => {
    if (!parsed || headers.length === 0) return;
    const rows5 = parsed.rows.slice(0, 5).map((row) => {
      const obj: Record<string, string> = {};
      headers.forEach((h, i) => {
        obj[h] = row[i] !== null && row[i] !== undefined ? String(row[i]) : "";
      });
      return obj;
    });
    setEditableRows(rows5);
    setSplitFormats({});
    // Auto-detect column mapping from header names
    setMapping({
      rang:      headers.find((k) => /rang|row|line|ligne/i.test(k)) ?? "",
      reference: headers.find((k) => /ref|coil|serial|num|id|bobine/i.test(k)) ?? headers[0] ?? "",
      poids:     headers.find((k) => /poids|weight|kg|tonne|masse/i.test(k)) ?? "",
      dch:       headers.find((k) => /dch|d[eé]chargement|dest(ination)?/i.test(k)) ?? "",
    });
    setExtras([]);
    setStep((s) => (s === 1 ? 2 : s));
    if (!suppressCollapseRef.current) {
      setOpenOnglets(false);
      setOpenAtypiques(false);
      setOpenHeaders(false);
      setOpenDonnees(false);
      setOpenMapping(false);
      setOpenApercu(false);
    }
    suppressCollapseRef.current = false;
  }, [parsed, activeSheet]); // intentionally excludes `headers`

  const onFile = useCallback((file: File) => { handleFile(file); }, [handleFile]);

  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) onFile(file);
  }, [onFile]);

  // Rename a column: update context headers + remap editableRows keys
  const renameColumn = useCallback((idx: number, newName: string) => {
    const trimmed = newName.trim();
    if (!trimmed || trimmed === headers[idx]) { setEditingHdr(null); return; }
    const oldName = headers[idx];
    setHeaders(headers.map((h, i) => (i === idx ? trimmed : h)));
    setEditableRows((prev) =>
      prev.map((row) => {
        const next: Record<string, string> = {};
        Object.entries(row).forEach(([k, v]) => { next[k === oldName ? trimmed : k] = v; });
        return next;
      })
    );
    setSplitFormats((prev) => {
      if (!prev[oldName]) return prev;
      const next = { ...prev };
      next[trimmed] = prev[oldName];
      delete next[oldName];
      return next;
    });
    setMapping((m) => {
      const updated = { ...m };
      (Object.keys(updated) as (keyof typeof updated)[]).forEach((k) => {
        if (updated[k] === oldName) updated[k] = trimmed;
      });
      return updated;
    });
    setExtras((exs) => exs.map((e) => e.col === oldName ? { ...e, col: trimmed } : e));
    setEditingHdr(null);
  }, [headers, setHeaders]);

  // Hide (mask) a column
  const deleteColumn = useCallback((idx: number) => {
    setHiddenCols((prev) => new Set([...prev, idx]));
  }, [setHiddenCols]);

  // Promote current header row to a data row, reset headers to Col1, Col2…
  const promoteHeaderToRow = useCallback(() => {
    if (!parsed) return;
    const newHeaders = headers.map((_, i) => `Col${i + 1}`);
    const headerAsRow: CellValue[] = headers.map((h) => h);
    setParsed({ ...parsed, headers: newHeaders, rows: [headerAsRow, ...parsed.rows], headerRowIndex: 0 });
    setHeaders(newHeaders);
    setSplitFormats({});
    setEditingHdr(null);
    // editableRows will be re-initialized by the effect when `parsed` changes
  }, [parsed, headers, setParsed, setHeaders]);

  const visibleCols = useMemo(
    () => headers.map((h, i) => ({ h, i })).filter(({ i }) => !hiddenCols.has(i)),
    [headers, hiddenCols]
  );

  // Anomaly detection
  const anomalyInfo = useMemo(() => {
    if (editableRows.length === 0) return { isAnomalous: (_: Record<string, string>) => false };
    const nonEmptyCounts = editableRows.map((row) =>
      visibleCols.filter(({ h }) => (row[h] ?? "").trim() !== "").length
    );
    const sorted = [...nonEmptyCounts].sort((a, b) => a - b);
    const median = sorted[Math.floor(sorted.length / 2)];
    const numericCols = visibleCols.filter(({ h }) => {
      const vals = editableRows.map((r) => r[h]).filter((v) => v && v.trim() !== "");
      if (vals.length === 0) return false;
      return vals.filter((v) => !isNaN(parseFloat(v)) && isFinite(Number(v))).length / vals.length > 0.5;
    });
    const isAnomalous = (row: Record<string, string>) => {
      const filled = visibleCols.filter(({ h }) => (row[h] ?? "").trim() !== "").length;
      if (filled < median - 1) return true;
      return numericCols.some(({ h }) => {
        const v = row[h] ?? "";
        return v.trim() !== "" && (isNaN(parseFloat(v)) || !isFinite(Number(v)));
      });
    };
    return { isAnomalous };
  }, [editableRows, visibleCols]);

  const STEP_LABELS = ["Fichier", "Colonnes", "Confirme"] as const;

  return (
    <div style={{ flex: 1, overflowY: "auto", paddingBottom: 80, background: T.bg }}>
      {winwinModalOpen && (
        <div
          onClick={() => setWinwinModalOpen(false)}
          style={{
            position: "fixed", inset: 0, background: "#000000bb",
            zIndex: 500, display: "flex", alignItems: "center", justifyContent: "center",
          }}
        >
          <div style={{ borderRadius: 18, overflow: "hidden", boxShadow: "0 24px 80px #000000cc", maxWidth: 320, width: "90%" }}>
            <img
              src="/STEtruc/P1060411.JPG"
              alt="winwin"
              style={{ width: "100%", display: "block" }}
            />
          </div>
        </div>
      )}
      <PageHeader title="Import" subtitle="Chargez et configurez votre fichier Excel" />

      {/* ── Step indicator ── */}
      <div style={{ display: "flex", background: T.bgDark, borderBottom: `1px solid ${T.border}` }}>
        {STEP_LABELS.map((l, i) => {
          const reachable = i + 1 < step;
          return (
            <div
              key={l}
              onClick={() => { if (reachable) setStep((i + 1) as 1 | 2 | 3); }}
              style={{
                flex: 1, textAlign: "center", padding: "10px 4px",
                borderBottom: `3px solid ${step > i ? T.success : step === i + 1 ? T.accent : "transparent"}`,
                cursor: reachable ? "pointer" : "default",
                opacity: reachable ? 1 : undefined,
                transition: "opacity 0.15s",
              }}
            >
              <div style={{
                color: step > i ? T.success : step === i + 1 ? T.accent : T.textDim,
                fontWeight: 800, fontSize: 12,
              }}>
                <span style={{
                  display: "inline-block", width: 20, height: 20, borderRadius: "50%",
                  background: step > i ? T.success : step === i + 1 ? T.accent : T.border,
                  color: "#0F172A", lineHeight: "20px", fontSize: 11, marginRight: 4,
                  fontWeight: 900, textAlign: "center",
                }}>{i + 1}</span>
                {l}
              </div>
            </div>
          );
        })}
      </div>

      <div style={{ padding: 16 }}>

        {/* ── STEP 1 : File upload ── */}
        {step === 1 && (
          <div>
            <div
              onDrop={onDrop}
              onDragOver={(e) => e.preventDefault()}
              onClick={() => fileInputRef.current?.click()}
              style={{
                border: `2px dashed ${T.accentDim}`, borderRadius: 14,
                padding: "48px 24px", textAlign: "center", cursor: "pointer",
                background: T.bgDark, marginBottom: 16,
              }}
            >
              <div style={{ fontSize: 48, marginBottom: 8 }}>📊</div>
              <div style={{ color: T.accent, fontWeight: 800, fontSize: 16, marginBottom: 4 }}>
                Charger fichier Excel
              </div>
              <div style={{ color: T.textDim, fontSize: 12 }}>
                .xlsx / .xls — colonnes et onglets vides auto-épurés
              </div>
            </div>
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls"
              style={{ display: "none" }}
              onChange={(e) => e.target.files?.[0] && onFile(e.target.files[0])}
            />
            {parsed && (
              <Btn onClick={() => setStep(2)} color={T.border2} textColor={T.textMuted} fullWidth>
                ← Reprendre le fichier actif
              </Btn>
            )}
          </div>
        )}

        {/* ── STEP 2 : Columns editor + data preview ── */}
        {step === 2 && parsed && (
          <div>
            {/* Sheet selector */}
            {sheetNames.length > 1 && (
              <div style={{ marginBottom: 14, background: T.bgCard, borderRadius: 12, padding: 14, border: `1px solid ${T.accentDim}` }}>
                <div
                  onClick={() => setOpenOnglets((o) => !o)}
                  style={{ color: T.accent, fontWeight: 700, fontSize: 12, marginBottom: openOnglets ? 8 : 0, textTransform: "uppercase", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span>🗂 Onglets ({sheetNames.length - hiddenSheets.size}/{sheetNames.length})</span>
                  <span style={{ opacity: 0.5, fontSize: 10 }}>{openOnglets ? "▲" : "▼"}</span>
                </div>
                {openOnglets && (
                  <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
                    {sheetNames.map((name) => {
                      const isActive = activeSheet === name;
                      const isExcluded = hiddenSheets.has(name);
                      return (
                        <div key={name} style={{ display: "flex", alignItems: "center", gap: 6 }}>
                          <button
                            onClick={() => { if (workbook && !isExcluded) { setActiveSheet(name); loadSheet(workbook, name); } }}
                            style={{
                              flex: 1, textAlign: "left",
                              background: isActive ? T.accent : isExcluded ? T.bgDark : T.bgDark,
                              color: isActive ? "#0F172A" : isExcluded ? T.textDim : T.textMuted,
                              border: `1px solid ${isActive ? T.accent : T.border2}`,
                              borderRadius: 7, padding: "6px 12px", fontSize: 12, fontWeight: 700,
                              cursor: isExcluded ? "default" : "pointer",
                              fontFamily: "'Share Tech Mono', monospace",
                              opacity: isExcluded ? 0.45 : 1,
                              textDecoration: isExcluded ? "line-through" : "none",
                            }}
                          >{name}</button>
                          <button
                            disabled={isActive}
                            onClick={() => {
                              if (isActive) return;
                              setHiddenSheets((prev) => {
                                const next = new Set(prev);
                                next.has(name) ? next.delete(name) : next.add(name);
                                return next;
                              });
                            }}
                            title={isActive ? "Onglet actif — impossible à exclure" : isExcluded ? "Inclure cet onglet" : "Exclure cet onglet"}
                            style={{
                              flexShrink: 0,
                              background: isExcluded ? `${T.success}22` : `${T.error}22`,
                              border: `1px solid ${isExcluded ? T.success : T.error}55`,
                              borderRadius: 6, padding: "5px 9px",
                              color: isExcluded ? T.success : T.error,
                              fontSize: 10, fontWeight: 800,
                              cursor: isActive ? "not-allowed" : "pointer",
                              opacity: isActive ? 0.3 : 1,
                              fontFamily: "'Share Tech Mono', monospace",
                              whiteSpace: "nowrap",
                            }}
                          >{isExcluded ? "✓ inclure" : "✕ exclure"}</button>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            )}

            {/* ── Lignes atypiques ── */}
            {(() => {
              const atypical = editableRows
                .map((row, ri) => ({ row, ri }))
                .filter(({ row }) => anomalyInfo.isAnomalous(row));
              if (atypical.length === 0) return (
                <div style={{ marginBottom: 14, background: "#111A14", borderRadius: 12, border: `1px solid ${T.success}44`, overflow: "hidden" }}>
                  <div style={{ padding: "10px 14px", background: "#0B1410", display: "flex", alignItems: "center", gap: 8 }}>
                    <span style={{ color: T.success, fontSize: 14 }}>✅</span>
                    <span style={{ color: T.success, fontWeight: 700, fontSize: 12, textTransform: "uppercase" }}>Aucune ligne atypique détectée</span>
                  </div>
                </div>
              );
              return (
                <div style={{ marginBottom: 14, background: "#1E1433", borderRadius: 12, border: `1px solid ${T.warning}55`, overflow: "hidden" }}>
                  <div style={{ padding: "10px 14px", background: "#140D24", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                    <span style={{ color: T.warning, fontWeight: 800, fontSize: 12, textTransform: "uppercase" }}>
                      ⚠ Lignes atypiques ({atypical.length})
                    </span>
                    <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                      <button
                        onClick={() => setEditableRows((prev) => prev.filter((row) => !anomalyInfo.isAnomalous(row)))}
                        style={{
                          background: `${T.error}22`, border: `1px solid ${T.error}55`,
                          borderRadius: 6, color: T.error, fontSize: 10, fontWeight: 700,
                          cursor: "pointer", padding: "3px 8px", whiteSpace: "nowrap",
                        }}
                      >✕ Supprimer tout</button>
                      <span
                        onClick={() => setOpenAtypiques((o) => !o)}
                        style={{ opacity: 0.5, fontSize: 10, cursor: "pointer", color: T.warning, padding: "0 2px" }}
                      >{openAtypiques ? "▲" : "▼"}</span>
                    </div>
                  </div>
                  {openAtypiques && <div style={{ maxHeight: 220, overflowY: "auto" }}>
                    {atypical.map(({ row, ri }) => (
                      <div key={ri} style={{
                        display: "flex", alignItems: "center", gap: 8,
                        padding: "7px 14px", borderBottom: `1px solid ${T.border}22`,
                        background: "#5B21B611",
                      }}>
                        <span style={{ color: T.warning, fontSize: 12, flexShrink: 0 }}>⚠</span>
                        <div style={{ flex: 1, display: "flex", gap: 8, flexWrap: "wrap", overflow: "hidden" }}>
                          {visibleCols.slice(0, 6).map(({ h }) => {
                            const v = row[h] ?? "";
                            return v.trim() !== "" ? (
                              <span key={h} style={{ color: T.textMuted, fontSize: 11, fontFamily: "monospace", whiteSpace: "nowrap" }}>
                                <span style={{ color: T.textDim, fontSize: 9 }}>{h}: </span>{v}
                              </span>
                            ) : null;
                          })}
                          {visibleCols.length > 6 && (
                            <span style={{ color: T.textDim, fontSize: 10 }}>+{visibleCols.length - 6} col.</span>
                          )}
                        </div>
                        <button
                          onClick={() => setEditableRows((prev) => prev.filter((_, idx) => idx !== ri))}
                          style={{ background: "none", border: "none", color: T.error, cursor: "pointer", fontSize: 16, padding: 0, flexShrink: 0 }}
                          title="Supprimer cette ligne"
                        >✕</button>
                      </div>
                    ))}
                  </div>}
                </div>
              );
            })()}

            {/* Headers editor */}
            <div style={{ marginBottom: 14, background: "#2A1020", borderRadius: 12, border: `1px solid ${T.border2}`, overflow: "hidden" }}>
              <div style={{
                padding: "10px 14px", background: "#1C0A16",
                display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8,
              }}>
                <span style={{ color: T.textMuted, fontWeight: 800, fontSize: 12, textTransform: "uppercase" }}>
                  📋 Headers · {visibleCols.length}/{headers.length} colonnes
                </span>
                <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" }}>
                  <button
                    onClick={promoteHeaderToRow}
                    title="Convertir les headers en ligne de données et créer des noms génériques Col1, Col2…"
                    style={{
                      background: `${T.warning}22`, border: `1px solid ${T.warning}55`,
                      borderRadius: 6, color: T.warning, fontSize: 10, fontWeight: 700,
                      cursor: "pointer", padding: "3px 8px", whiteSpace: "nowrap",
                    }}
                  >↩ Header → ligne</button>
                  <button
                    onClick={() => {
                      const newIdx = headers.length;
                      const newName = `Col${newIdx + 1}`;
                      suppressCollapseRef.current = true;
                      setOpenHeaders(true);
                      setHeaders([...headers, newName]);
                      setEditableRows(prev => prev.map(row => ({ ...row, [newName]: "" })));
                      if (parsed) setParsed({ ...parsed, headers: [...parsed.headers, newName], rows: parsed.rows.map(r => [...r, null]) });
                      setEditingHdr(newIdx);
                    }}
                    title="Ajouter une colonne vide en dernière position"
                    style={{
                      background: `${T.accent}22`, border: `1px solid ${T.accent}55`,
                      borderRadius: 6, color: T.accent, fontSize: 10, fontWeight: 700,
                      cursor: "pointer", padding: "3px 8px", whiteSpace: "nowrap",
                    }}
                  >+ Col.</button>
                  {hiddenCols.size > 0 && (
                    <button
                      onClick={() => setHiddenCols(new Set())}
                      style={{
                        background: `${T.success}22`, border: `1px solid ${T.success}44`,
                        borderRadius: 6, color: T.success, fontSize: 10, fontWeight: 700,
                        cursor: "pointer", padding: "3px 8px",
                      }}
                    >↺ Restaurer ({hiddenCols.size})</button>
                  )}
                  <span
                    onClick={() => { setOpenHeaders((o) => !o); setOpenDonnees((o) => !o); }}
                    style={{ opacity: 0.5, fontSize: 10, cursor: "pointer", color: T.textMuted, padding: "0 2px" }}
                  >{openHeaders ? "▲" : "▼"}</span>
                </div>
              </div>
              {openHeaders && <div style={{ padding: "10px 14px", display: "flex", gap: 6, flexWrap: "wrap" }}>
                {headers.map((h, i) => {
                  if (hiddenCols.has(i)) return null;
                  return (
                    <span key={i} style={{
                      background: T.bgDark, color: T.text, borderRadius: 6,
                      fontSize: 11, padding: "4px 6px 4px 10px",
                      fontFamily: "'Share Tech Mono', monospace",
                      border: `1px solid ${editingHdr === i ? T.accent : T.border2}`,
                      display: "flex", alignItems: "center", gap: 5,
                    }}>
                      <span style={{ color: T.textDim, fontSize: 9, fontWeight: 700 }}>{i + 1}</span>
                      {editingHdr === i ? (
                        <input
                          autoFocus
                          defaultValue={h}
                          onBlur={(e) => renameColumn(i, e.target.value)}
                          onKeyDown={(e) => {
                            if (e.key === "Enter") renameColumn(i, (e.target as HTMLInputElement).value);
                            if (e.key === "Escape") setEditingHdr(null);
                          }}
                          style={{
                            background: "transparent", border: "none", outline: "none",
                            color: T.accent, fontFamily: "'Share Tech Mono', monospace",
                            fontSize: 11, width: Math.max(60, h.length * 8),
                          }}
                        />
                      ) : (
                        <span
                          onDoubleClick={() => setEditingHdr(i)}
                          title="Double-cliquer pour renommer"
                          style={{ cursor: "text" }}
                        >{h}</span>
                      )}
                      <button
                        onClick={() => deleteColumn(i)}
                        title={`Masquer la colonne "${h}"`}
                        style={{
                          background: "none", border: "none", color: T.error,
                          cursor: "pointer", fontSize: 12, padding: "0 0 0 4px",
                          lineHeight: 1, display: "flex", alignItems: "center",
                        }}
                      >✕</button>
                    </span>
                  );
                })}
              </div>}
            </div>

            {/* Editable data preview */}
            {editableRows.length > 0 && (
              <div style={{ marginBottom: 14, background: T.bgDark, borderRadius: 12, border: `1px solid ${T.border2}`, overflow: "hidden" }}>
                <div
                  onClick={() => setOpenDonnees((o) => !o)}
                  style={{ padding: "10px 14px", background: T.bgCard, display: "flex", alignItems: "center", gap: 8, cursor: "pointer" }}
                >
                  <span style={{ color: T.accent, fontWeight: 800, fontSize: 12, textTransform: "uppercase", flex: 1 }}>
                    ✏️ Données — 5 premières lignes ({editableRows.length} total)
                  </span>
                  <span style={{ color: T.textDim, fontSize: 10 }}>Modifiables avant import</span>
                  <span style={{ color: T.textDim, fontSize: 10, opacity: 0.6, marginLeft: 4 }}>{openDonnees ? "▲" : "▼"}</span>
                </div>
                {openDonnees && <div style={{ overflowX: "auto", maxHeight: 260, overflowY: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                    <thead>
                      <tr style={{ position: "sticky", top: 0, background: T.bgDark, zIndex: 1 }}>
                        {visibleCols.map(({ h, i }) => (
                          <th key={i} style={{
                            color: T.textDim, padding: "5px 6px", textAlign: "left",
                            borderBottom: `1px solid ${T.border}`, whiteSpace: "nowrap", fontWeight: 700,
                          }}>
                            {h}
                          </th>
                        ))}
                        <th style={{ width: 40, borderBottom: `1px solid ${T.border}` }} />
                      </tr>
                    </thead>
                    <tbody>
                      {editableRows.slice(0, 5).map((row, ri) => {
                        const anomalous = anomalyInfo.isAnomalous(row);
                        return (
                          <tr key={ri} style={{
                            borderBottom: `1px solid ${T.border}22`,
                            background: anomalous ? "#7C150822" : "transparent",
                            outline: anomalous ? `1px solid ${T.error}33` : "none",
                          }}>
                            {visibleCols.map(({ h }) => (
                              <td key={h} style={{ padding: "3px 4px" }}>
                                <input
                                  value={row[h] ?? ""}
                                  onChange={(e) => setEditableRows((prev) =>
                                    prev.map((r, idx) => idx === ri ? { ...r, [h]: e.target.value } : r)
                                  )}
                                  style={{
                                    background: T.bgCard, border: `1px solid ${T.border2}55`,
                                    borderRadius: 4, color: T.text, fontSize: 11,
                                    padding: "3px 6px", width: "100%", minWidth: 60,
                                    outline: "none", boxSizing: "border-box",
                                    fontFamily: "'Share Tech Mono', monospace",
                                  }}
                                />
                              </td>
                            ))}
                            <td style={{ padding: "3px 4px", textAlign: "center", whiteSpace: "nowrap" }}>
                              {anomalous && (
                                <span title="Ligne atypique" style={{ color: T.warning, fontSize: 12, marginRight: 2 }}>⚠</span>
                              )}
                              <button
                                onClick={() => setEditableRows((prev) => prev.filter((_, idx) => idx !== ri))}
                                style={{ background: "none", border: "none", color: T.error, cursor: "pointer", fontSize: 14, padding: 0 }}
                                title="Supprimer cette ligne"
                              >✕</button>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>}
              </div>
            )}
            {visibleCols.length > 0 && (
              <div style={{ marginBottom: 14, background: T.bgCard, borderRadius: 12, border: `1px solid ${T.border2}`, overflow: "hidden" }}>
                {/* Header bar */}
                <div
                  onClick={() => { setOpenMapping((o) => !o); setOpenApercu((o) => !o); }}
                  style={{ padding: "8px 14px", background: T.bgDark, borderBottom: openMapping ? `1px solid ${T.border2}` : "none", display: "flex", alignItems: "center", gap: 8, cursor: "pointer" }}
                >
                  <span style={{ color: T.textMuted, fontWeight: 800, fontSize: 12, textTransform: "uppercase", flex: 1 }}>🗂 Mapping colonnes</span>
                  <span style={{ color: T.textDim, fontSize: 10, opacity: 0.6, marginLeft: 4 }}>{openMapping ? "▲" : "▼"}</span>
                </div>
                {openMapping && <div style={{ padding: "10px 14px" }}>
                <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
                  {([
                    { field: "rang"      as const, label: "📍 Rang" },
                    { field: "reference" as const, label: "🏷 Référence *" },
                    { field: "poids"     as const, label: "⚖️ Poids (t)" },
                    { field: "dch"       as const, label: "🏗️ Destination" },
                  ] as { field: keyof typeof mapping; label: string }[]).map(({ field, label }) => (
                    <div key={field} style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ color: T.textMuted, fontSize: 11, marginBottom: 4, fontWeight: 600 }}>{label}</div>
                      <select
                        value={mapping[field]}
                        onChange={(e) => setMapping((m) => ({ ...m, [field]: e.target.value }))}
                        style={{
                          width: "100%", background: T.bgCard,
                          border: `1px solid ${mapping[field] ? T.accent : T.border2}`,
                          borderRadius: 8, color: T.text, fontSize: 12, padding: "7px 6px",
                          outline: "none", fontFamily: "'Share Tech Mono', monospace",
                          boxSizing: "border-box",
                        }}
                      >
                        <option value="">— — —</option>
                        {visibleCols.map(({ h }) => <option key={h} value={h}>{h}</option>)}
                      </select>
                    </div>
                  ))}
                </div>
                {/* Extra columns */}
                <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                  {extras.map((ex, idx) => (
                    <div key={idx} style={{ display: "flex", gap: 6, alignItems: "flex-end" }}>
                      <div style={{ flex: "0 0 110px", minWidth: 0 }}>
                        {idx === 0 && <div style={{ color: T.textMuted, fontSize: 11, marginBottom: 4, fontWeight: 600 }}>Nom affiché</div>}
                        <input
                          value={ex.label}
                          onChange={(e) => setExtras((exs) => exs.map((x, i) => i === idx ? { ...x, label: e.target.value } : x))}
                          placeholder="ex: PREPA"
                          style={{
                            width: "100%", background: T.bgCard, border: `1px solid ${T.border2}`,
                            borderRadius: 8, color: T.accent, fontSize: 12, padding: "7px 8px",
                            outline: "none", fontFamily: "'Share Tech Mono', monospace", boxSizing: "border-box",
                          }}
                        />
                      </div>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        {idx === 0 && <div style={{ color: T.textMuted, fontSize: 11, marginBottom: 4, fontWeight: 600 }}>+ Colonne supplémentaire</div>}
                        <select
                          value={ex.col}
                          onChange={(e) => {
                            const newCol = e.target.value;
                            setExtras((exs) => exs.map((x, i) => {
                              if (i !== idx) return x;
                              // Auto-remplir le label si vide, "EXTRA", ou si label == ancienne colonne
                              const isDefault = !x.label.trim() || x.label === "EXTRA" || x.label === x.col;
                              const autoLabel = (isDefault && newCol) ? newCol : x.label;
                              return { ...x, col: newCol, label: autoLabel };
                            }));
                          }}
                          style={{
                            width: "100%", background: T.bgCard,
                            border: `1px solid ${ex.col ? T.success : T.border2}`,
                            borderRadius: 8, color: T.text, fontSize: 12, padding: "7px 6px",
                            outline: "none", fontFamily: "'Share Tech Mono', monospace",
                            boxSizing: "border-box",
                          }}
                        >
                          <option value="">— Aucune —</option>
                          {visibleCols.map(({ h }) => <option key={h} value={h}>{h}</option>)}
                        </select>
                      </div>
                      <button
                        onClick={() => setExtras((exs) => exs.filter((_, i) => i !== idx))}
                        title="Supprimer cette colonne"
                        style={{
                          flexShrink: 0, background: `${T.error}22`,
                          border: `1px solid ${T.error}55`, borderRadius: 7,
                          color: T.error, fontSize: 13, cursor: "pointer",
                          padding: "6px 9px", lineHeight: 1,
                        }}
                      >🗑</button>
                    </div>
                  ))}
                </div>
                </div>}
              </div>
            )}
            {editableRows.length > 0 && (
              <div style={{ background: "#0C2D3A", borderRadius: 10, marginBottom: 14, overflow: "hidden", border: `1px solid #1A5F7A55` }}>
                <div
                  onClick={() => setOpenApercu((o) => !o)}
                  style={{ padding: "8px 12px", background: "#103848", borderBottom: openApercu ? `1px solid #1A5F7A55` : "none", display: "flex", alignItems: "center", gap: 8, cursor: "pointer" }}
                >
                  <span style={{ color: T.textDim, fontSize: 11, textTransform: "uppercase", fontWeight: 700, flex: 1 }}>Aperçu — 5 premières lignes</span>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <span style={{ color: T.textDim, fontSize: 10 }}>Cliquer sur un header pour grouper</span>
                    <button
                      onClick={(e) => { e.stopPropagation(); setAutoRefFmt((v) => !v); }}
                      style={{
                        background: autoRefFmt ? `${T.accent}22` : T.bgDark,
                        border: `1px solid ${autoRefFmt ? T.accent : T.border2}`, borderRadius: 6,
                        color: autoRefFmt ? T.accent : T.textDim, fontSize: 10, fontWeight: 700,
                        cursor: "pointer", padding: "2px 7px", fontFamily: "'Share Tech Mono', monospace",
                      }}
                      title="Formatage automatique des références (selon longueur/alphanum)"
                    >🔢 Réf. auto</button>
                    <button
                      onClick={(e) => { e.stopPropagation(); setPoidsUnit((u) => u === "t" ? "kg" : "t"); }}
                      style={{
                        background: poidsUnit === "kg" ? `${T.warning}33` : T.bgDark,
                        border: `1px solid ${T.warning}66`, borderRadius: 6,
                        color: T.warning, fontSize: 10, fontWeight: 700,
                        cursor: "pointer", padding: "2px 7px", fontFamily: "'Share Tech Mono', monospace",
                      }}
                      title="Basculer tonnes / kilos"
                    >{poidsUnit === "t" ? "⚖️ t ⇄ kg" : "⚖️ kg ⇄ t"}</button>
                  </div>
                  <span style={{ color: T.textDim, fontSize: 10, opacity: 0.6, marginLeft: 4 }}>{openApercu ? "▲" : "▼"}</span>
                </div>
                {openApercu && <div style={{ padding: 10, overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                    <thead>
                      <tr>
                        {([
                          { key: "rang",      label: "📍 Rang",      color: T.textMuted },
                          { key: "reference", label: "🏷 Référence", color: T.text },
                          { key: "poids",     label: `⚖️ Poids (${poidsUnit})`, color: T.warning },
                          { key: "dch",       label: "🏗️ Destination",  color: T.accent },
                          ...extras.map((e, i) => ({ key: `extra_${i}`, label: e.label || "EXTRA", color: T.success })),
                        ] as { key: string; label: string; color: string }[]).map(({ key, label }) => (
                          <th key={key} style={{ padding: "4px 6px", textAlign: "left", borderBottom: `1px solid ${T.border}` }}>
                            {splitEditingField === key ? (
                              <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
                                <input
                                  autoFocus
                                  value={splitInputValue}
                                  onChange={(e) => setSplitInputValue(e.target.value)}
                                  placeholder="ex: 4 4 1"
                                  onKeyDown={(e) => {
                                    if (e.key === "Enter") { setSplitFormats((p) => ({ ...p, [key]: splitInputValue })); setSplitEditingField(null); }
                                    if (e.key === "Escape") setSplitEditingField(null);
                                  }}
                                  onBlur={() => { setSplitFormats((p) => ({ ...p, [key]: splitInputValue })); setSplitEditingField(null); }}
                                  style={{
                                    background: T.bgCard, border: `1px solid ${T.accent}`, borderRadius: 4,
                                    color: T.accent, fontSize: 10, padding: "2px 6px", width: 70, outline: "none",
                                    fontFamily: "'Share Tech Mono', monospace",
                                  }}
                                />
                                <button
                                  onClick={() => { setSplitFormats((p) => { const n = { ...p }; delete n[key]; return n; }); setSplitEditingField(null); }}
                                  style={{ background: "none", border: "none", color: T.error, cursor: "pointer", fontSize: 11, padding: 0 }}
                                >✕</button>
                              </div>
                            ) : (
                              <span
                                onClick={() => { setSplitEditingField(key); setSplitInputValue(splitFormats[key] || ""); }}
                                title="Cliquer pour définir un groupement visuel (ex: 4 4 1)"
                                style={{ color: splitFormats[key] ? T.success : key.startsWith("extra") ? T.success : T.accent, cursor: "pointer", fontWeight: 700, display: "inline-flex", alignItems: "center", gap: 4 }}
                              >
                                {label}
                                {splitFormats[key] && <span style={{ fontSize: 9, color: `${T.success}88`, fontFamily: "monospace" }}>[{splitFormats[key]}]</span>}
                                <span style={{ fontSize: 9, opacity: 0.4 }}>✎</span>
                              </span>
                            )}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {editableRows.slice(0, 5).map((row, ri) => (
                        <tr key={ri} style={{ borderBottom: `1px solid ${T.border}22` }}>
                          <td style={{ color: T.textMuted, padding: "4px 6px", fontFamily: "monospace" }}>
                            {mapping.rang ? applyGrouping(row[mapping.rang] ?? "", splitFormats["rang"] || "") || "—" : "—"}
                          </td>
                          <td style={{ color: T.text, padding: "4px 6px", fontFamily: "monospace", fontSize: 11 }}>
                            {mapping.reference ? (() => {
                              const raw = String(row[mapping.reference] ?? "").slice(0, 28);
                              return (autoRefFmt ? autoFormatRef(raw, splitFormats["reference"] || "") : applyGrouping(raw, splitFormats["reference"] || "")) || "—";
                            })() : "—"}
                          </td>
                          <td style={{ color: T.warning, padding: "4px 6px", fontFamily: "monospace" }}>
                            {mapping.poids ? (() => {
                              const raw = parseFloat(row[mapping.poids] ?? "") || 0;
                              const val = poidsUnit === "kg" ? (raw * 1000).toFixed(0) : parseFloat(raw.toFixed(3)).toString();
                              const poidsFmt = splitFormats["poids"] ?? "";
                              return (poidsFmt ? applyGrouping(val, poidsFmt) : thsep(val)) + " " + (poidsUnit === "kg" ? "kg" : "t");
                            })() : "—"}
                          </td>
                          <td style={{ color: T.accent, padding: "4px 6px", fontFamily: "monospace", fontSize: 11 }}>
                            {mapping.dch ? applyGrouping(String(row[mapping.dch] ?? "").slice(0, 24), splitFormats["dch"] || "") || "—" : "—"}
                          </td>
                          {extras.map((e, i) => (
                            <td key={i} style={{ color: T.success, padding: "4px 6px", fontFamily: "monospace", fontSize: 11 }}>
                              {e.col ? applyGrouping(String(row[e.col] ?? "").slice(0, 24), splitFormats[`extra_${i}`] || "") || <span style={{ color: T.textDim }}>—</span> : <span style={{ color: T.textDim, fontSize: 10 }}>— — —</span>}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>}
              </div>
            )}

            <div style={{ display: "flex", gap: 10, marginTop: 4 }}>
              <Btn onClick={() => setStep(1)} color={T.border2} textColor={T.textMuted} fullWidth>← Retour</Btn>
              <Btn onClick={() => {
                if (!mapping.dch) {
                  const destName = "Dest";
                  setHeaders([...headers, destName]);
                  setEditableRows(prev => prev.map(row => ({ ...row, [destName]: "" })));
                  setMapping(m => ({ ...m, dch: destName }));
                }
                setStep(3);
              }} color={T.accent} textColor="#0F172A" fullWidth>Suivant →</Btn>
            </div>
          </div>
        )}

        {/* ── STEP 3 : Confirmation ── */}
        {step === 3 && parsed && (() => {
          const visibleCount = headers.filter((_, i) => !hiddenCols.has(i)).length;
          const totalRows = parsed.rows.length;
          return (
            <div>
              <div style={{ background: T.bgCard, borderRadius: 12, padding: 16, marginBottom: 14, border: `1px solid ${T.success}55` }}>
                <div style={{ color: T.success, fontWeight: 800, fontSize: 16, marginBottom: 10 }}>
                  ✅ Prêt à analyser
                </div>
                {([
                  ["Fichier", fileName ?? "—"],
                  ["Onglet actif", activeSheet ?? "—"],
                  ["Colonnes visibles", String(visibleCount)],
                  ["Lignes totales", String(totalRows)],
                  ["📍 Colonne Rang", mapping.rang || "— non mappé"],
                  ["🏷 Colonne Référence", mapping.reference || "— non mappé"],
                  ["⚖️ Colonne Poids", mapping.poids || "— non mappé"],
                  ["🏗️ Colonne Destination", mapping.dch || "— non mappé"],
                  ["⚖️ Unité poids", poidsUnit === "kg" ? "Kilogrammes (kg)" : "Tonnes (t)"],
                  ["🔢 Réf. auto-groupée", autoRefFmt ? "Activé" : "Désactivé"],
                  ...extras.filter((e) => e.label.trim()).map((e) => [`🔖 ${e.label}`, e.col ? e.col : "+ colonne vide"]),
                ] as [string, string][]).map(([k, v]) => (
                  <div key={k} style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", borderBottom: `1px solid ${T.border2}33` }}>
                    <span style={{ color: T.textMuted, fontSize: 13 }}>{k}</span>
                    <span style={{ color: T.text, fontWeight: 700, fontSize: 13, fontFamily: "monospace" }}>{v}</span>
                  </div>
                ))}
                {Object.keys(splitFormats).length > 0 && (
                  <div style={{ marginTop: 10 }}>
                    <div style={{ color: T.textDim, fontSize: 10, textTransform: "uppercase", letterSpacing: "0.1em", fontWeight: 700, marginBottom: 6 }}>Groupements par colonne</div>
                    {Object.entries(splitFormats).map(([col, fmt]) => (
                      <div key={col} style={{ display: "flex", justifyContent: "space-between", padding: "4px 0", borderBottom: `1px solid ${T.border2}22` }}>
                        <span style={{ color: T.textMuted, fontSize: 12 }}>{col}</span>
                        <span style={{ color: T.accent, fontWeight: 700, fontSize: 12, fontFamily: "monospace" }}>{fmt}</span>
                      </div>
                    ))}
                  </div>
                )}
              </div>
              <div style={{ display: "flex", gap: 10 }}>
                <Btn onClick={() => setStep(2)} color={T.border2} textColor={T.textMuted} fullWidth>← Retour</Btn>
                <Btn
                  onClick={() => {
                    if (!parsed) return;
                    // Reconstruire parsed.rows depuis editableRows (délétions appliquées) + lignes au-delà de 200
                    const tail = parsed.rows.slice(200);
                    const currentHeaders = headers;
                    let newRows: RawData = [
                      ...editableRows.map((rowObj) =>
                        currentHeaders.map((h) => {
                          const v = rowObj[h] ?? "";
                          return v === "" ? null : v;
                        })
                      ),
                      ...tail,
                    ];
                    const newHeaders = [...currentHeaders];
                    // Apply extras: rename existing cols or add new empty cols
                    extras.forEach((ex) => {
                      const label = ex.label.trim() || "EXTRA";
                      if (ex.col) {
                        const idx = newHeaders.indexOf(ex.col);
                        if (idx >= 0) newHeaders[idx] = label;
                      } else {
                        newHeaders.push(label);
                        newRows = newRows.map((r) => [...r, null]);
                      }
                    });
                    setHeaders(newHeaders);
                    setParsed({ ...parsed, headers: newHeaders, rows: newRows });
                    showToast("✅ Fichier prêt", "success");
                    setActiveTab("tableau");
                  }}
                  color={T.success} textColor="#0F172A" fullWidth
                >📊 Voir le pointage →</Btn>
              </div>
            </div>
          );
        })()}

      </div>
    </div>
  );
}

// ─────────────────────────────────────────────
// 6. PAGE POINTAGE
// ─────────────────────────────────────────────

// Barre de navigation rapide entre onglets (usage dans TablePage)
function TablePage() {
  const {
    parsed, headers, setHeaders,
    hiddenCols, setHiddenCols: _setHiddenCols, hiddenRows, setHiddenRows: _setHiddenRows,
    selectMode, setSelectMode, selectedItems, setSelectedItems,
    editingHeader, setEditingHeader,
    dimRepeated, setDimRepeated,
    allRows, repetitiveByCol, addedRows,
    setActiveTab, showToast,
    splitFormats, setSplitFormats,
    mapping, extras,
    pointedRows, setPointedRows,
    rowOverrides, setRowOverrides,
    autoRefFmt, setAutoRefFmt,
    poidsUnit, setPoidsUnit,
    destinations, setDestinations,
    selectedDest, setSelectedDest,
    rowDestinations, setRowDestinations,
    reassignedRows: _reassignedRows, setReassignedRows,
  } = useApp();

  const [showAddRow,  setShowAddRow]  = useState(false);
  const [colFilters,  setColFilters]  = useState<Record<number, string>>({});
  const [sortCol,     setSortCol]     = useState<number | null>(null);
  const [sortDir,     setSortDir]     = useState<"asc" | "desc">("asc");
  const [pageIdx,     setPageIdx]     = useState(0);
  const [pageSize,    setPageSize]    = useState(20);
  const [modalRow,    setModalRow]    = useState<{ row: CellValue[]; rowNum: number; ri: number } | null>(null);
  const [openToolbar, setOpenToolbar] = useState(false);
  const [openMouvements, setOpenMouvements] = useState(false);
  const [newDestName, setNewDestName] = useState("");
  const [confirmClearDest, setConfirmClearDest] = useState(false);
  const touchStartPos = useRef<{ x: number; y: number } | null>(null);
  const touchScrolled = useRef(false);
  const [confirmReassign, setConfirmReassign] = useState<{ ri: number; from: string; to: string } | null>(null);
  const [confirmUnpoint, setConfirmUnpoint] = useState<{ ri: number; ref: string } | null>(null);
  const [doubleVerif, setDoubleVerif] = useState(true);
  const [doubleVerifModal, setDoubleVerifModal] = useState<{ ri: number; ref: string; correct: number; choices: number[]; error: boolean; pendingDest?: string } | null>(null);
  const [refGroupModal, setRefGroupModal] = useState(false);
  const [refGroupInput, setRefGroupInput] = useState("");
  const longPressRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  if (!parsed) {
    return (
      <div style={{ flex: 1, overflowY: "auto", paddingBottom: 80, background: T.bg }}>
        <PageHeader title="Pointage" />
        <EmptyState icon="📁" text="Aucun fichier chargé" sub="Importez un fichier Excel d'abord" />
        <div style={{ padding: 16 }}>
          <Btn onClick={() => setActiveTab("import")} color={T.accent} textColor="#0F172A" fullWidth>
            ⬇️ Aller à l'import
          </Btn>
        </div>
      </div>
    );
  }

  const toggleSelectItem = (idx: number) => {
    setSelectedItems((prev) => {
      const next = new Set(prev);
      next.has(idx) ? next.delete(idx) : next.add(idx);
      return next;
    });
  };

  const applyColAction = () => {
    if (selectMode !== "col") return;
    _setHiddenCols((prev) => new Set([...prev, ...selectedItems]));
    setSelectedItems(new Set());
    setSelectMode("none");
    showToast(`${selectedItems.size} colonne(s) masquée(s)`, "info");
  };

  const applyRowDeletion = () => {
    if (selectMode !== "row") return;
    _setHiddenRows((prev) => new Set([...prev, ...selectedItems]));
    setSelectedItems(new Set());
    setSelectMode("none");
    showToast(`${selectedItems.size} ligne(s) masquée(s)`, "info");
  };

  const visibleColCount = headers.filter((_, i) => !hiddenCols.has(i)).length;
  void visibleColCount; // kept for reference

  // ── Column ordering: prepa last, destination column renamed to DEST ──
  const prepaIdx = headers.findIndex((h) => /^prepa$/i.test(h.trim()));
  const destIdx  = headers.findIndex((h) => /^dest(ination)?$/i.test(h.trim()));
  const visibleCols: number[] = headers
    .map((_, i) => i)
    .filter((i) => !hiddenCols.has(i) && i !== prepaIdx);
  if (prepaIdx >= 0 && !hiddenCols.has(prepaIdx)) visibleCols.push(prepaIdx);

  // Labels affichés dans les en-têtes : priorité aux noms définis dans l'Import
  const mappingLabels = new Map<number, string>();
  if (mapping.rang)      { const i = headers.indexOf(mapping.rang);      if (i >= 0) mappingLabels.set(i, "📍 Rang"); }
  if (mapping.reference) { const i = headers.indexOf(mapping.reference); if (i >= 0) mappingLabels.set(i, "🏷 REF"); }
  if (mapping.poids)     { const i = headers.indexOf(mapping.poids);     if (i >= 0) mappingLabels.set(i, `⚖️ Poids(${poidsUnit})`); }
  if (mapping.dch)       { const i = headers.indexOf(mapping.dch);       if (i >= 0) mappingLabels.set(i, "🏗️ DEST"); }
  extras.forEach((ex) => {
    if (ex.col && ex.label.trim()) {
      const i = headers.indexOf(ex.col);
      if (i >= 0) mappingLabels.set(i, ex.label.trim());
    }
  });
  const colLabel = (i: number): string =>
    mappingLabels.get(i) ?? (i === destIdx ? "DEST" : headers[i]);

  // ── Filter / sort / paginate ──
  const baseRows = allRows
    .map((row, ri) => ({ row, ri }))
    .filter(({ ri }) => !hiddenRows.has(ri));

  const filteredRows = baseRows.filter(({ row }) =>
    visibleCols.every((ci) => {
      const f = (colFilters[ci] ?? "").trim().toLowerCase();
      if (!f) return true;
      const cell = row[ci];
      return cell !== null && String(cell).toLowerCase().includes(f);
    })
  );

  const sortedRows = sortCol !== null
    ? [...filteredRows].sort((a, b) => {
        const va = a.row[sortCol] ?? "";
        const vb = b.row[sortCol] ?? "";
        const na = Number(va), nb = Number(vb);
        const numCmp = !isNaN(na) && !isNaN(nb) ? na - nb : 0;
        const cmp = numCmp !== 0 ? numCmp : String(va).localeCompare(String(vb));
        return sortDir === "asc" ? cmp : -cmp;
      })
    : filteredRows;

  const PAGE_SIZES = [10, 20, 40, 100, 200];
  const totalPages = Math.max(1, Math.ceil(sortedRows.length / pageSize));
  const safePage   = Math.min(pageIdx, totalPages - 1);
  const pageRows   = sortedRows.slice(safePage * pageSize, (safePage + 1) * pageSize);

  // Indices des colonnes référence et poids
  const refColIdx   = mapping.reference ? headers.indexOf(mapping.reference)  : -1;
  const poidsColIdx = mapping.poids     ? headers.indexOf(mapping.poids)      : -1;
  const dchColIdx   = mapping.dch       ? headers.indexOf(mapping.dch)        : -1;
  const rangColIdx  = mapping.rang      ? headers.indexOf(mapping.rang)       : -1;

  // Stats par destination
  const destStats = useMemo(() => {
    const stats = new Map<string, { count: number; weight: number }>();
    for (const [ri2, dest] of rowDestinations) {
      if (!stats.has(dest)) stats.set(dest, { count: 0, weight: 0 });
      const s = stats.get(dest)!;
      s.count++;
      if (poidsColIdx >= 0) {
        const rw = allRows[ri2];
        if (rw) {
          const rawW = parseFloat(String(rw[poidsColIdx] ?? "")) || 0;
          s.weight += poidsUnit === "kg" ? rawW * 1000 : rawW;
        }
      }
    }
    return stats;
  }, [rowDestinations, allRows, poidsColIdx, poidsUnit]);

  // Largeur des colonnes — compression proportionnelle pour éviter le scroll
  const numColW = Math.max(28, String(allRows.length).length * 8 + 14);

  const colWidths = useMemo(() => {
    const CH_DATA = 9.5;   // px/char police 18px
    const CH_HDR  = 8;    // px/char police 14px uppercase
    const MIN_W   = 38;
    const MAX_W   = 200;
    const PAD     = 20;    // padding interne
    const NUM_W   = numColW;    // colonne # — dynamique selon nb de lignes
    const VIEWPORT = 380; // largeur cible mobile
    const sample = sortedRows.slice(0, 80);
    const ideals = new Map<number, number>();
    visibleCols.forEach((ci) => {
      const hLen = colLabel(ci).replace(/[^\x00-\x7F]/g, "  ").length; // emoji = 2
      let maxData = 0;
      for (const { row } of sample) {
        const v = row[ci];
        if (v !== null && v !== undefined) maxData = Math.max(maxData, String(v).length);
      }
      const ideal = Math.max(hLen * CH_HDR, maxData * CH_DATA) + PAD;
      ideals.set(ci, Math.min(MAX_W, Math.max(MIN_W, ideal)));
    });
    const totalIdeal = NUM_W + [...ideals.values()].reduce((a, b) => a + b, 0);
    const widths = new Map<number, number>();
    if (totalIdeal <= VIEWPORT) {
      // Tout tient : on étire uniformément pour remplir
      const extra = (VIEWPORT - totalIdeal) / visibleCols.length;
      visibleCols.forEach((ci) => widths.set(ci, (ideals.get(ci) ?? MIN_W) + extra));
    } else {
      // Compression proportionnelle avec min garanti
      const available = VIEWPORT - NUM_W - visibleCols.length * MIN_W;
      const idealTotal = [...ideals.values()].reduce((a, b) => a + b, 0) - visibleCols.length * MIN_W;
      const ratio = idealTotal > 0 ? Math.max(0, available) / idealTotal : 0;
      visibleCols.forEach((ci) => {
        const ideal = ideals.get(ci) ?? MIN_W;
        widths.set(ci, Math.round(MIN_W + (ideal - MIN_W) * ratio));
      });
    }
    return widths;
  }, [sortedRows, visibleCols, headers, mapping, extras, poidsUnit, numColW]); // eslint-disable-line react-hooks/exhaustive-deps

  // Détection des lignes atypiques (visuel uniquement)
  const anomalyRowSet = useMemo(() => {
    if (allRows.length === 0) return new Set<number>();
    const visibleRIs = allRows.map((_, ri) => ri).filter((ri) => !hiddenRows.has(ri));
    if (visibleRIs.length === 0) return new Set<number>();
    const counts = visibleRIs.map((ri) =>
      allRows[ri].filter((v) => v !== null && v !== undefined && String(v).trim() !== "").length
    );
    const sorted = [...counts].sort((a, b) => a - b);
    const median = sorted[Math.floor(sorted.length / 2)];
    const anomalous = new Set<number>();
    visibleRIs.forEach((ri, idx) => { if (counts[idx] < median - 1) anomalous.add(ri); });
    return anomalous;
  }, [allRows, hiddenRows]);

  // Résoudre le format de groupement : par nom de colonne réel OU par clé logique
  const getColFmt = (ci: number): string => {
    const byName = splitFormats[headers[ci]];
    if (byName) return byName;
    if (ci === refColIdx)   return splitFormats["reference"] ?? "";
    if (ci === poidsColIdx) return splitFormats["poids"]     ?? "";
    if (ci === rangColIdx)  return splitFormats["rang"]      ?? "";
    if (ci === dchColIdx)   return splitFormats["dch"]       ?? "";
    const exIdx = extras.findIndex((ex) => ex.col && headers.indexOf(ex.col) === ci);
    if (exIdx >= 0)         return splitFormats[`extra_${exIdx}`] ?? "";
    return "";
  };

  const handleSortCol = (ci: number) => {
    if (selectMode === "col") { toggleSelectItem(ci); return; }
    if (selectMode !== "none") return;
    if (sortCol !== ci) { setSortCol(ci); setSortDir("asc"); }
    else if (sortDir === "asc") { setSortDir("desc"); }
    else { setSortCol(null); }
    setPageIdx(0);
  };

  const css = `
    .pt-table { border-collapse: collapse; width: 100%; table-layout: fixed; font-size: 18px; }
    .pt-table thead th {
      background: #131E2E; color: ${T.textMuted};
      font-size: 14px; letter-spacing: 0.07em; text-transform: uppercase;
      padding: 0; text-align: left;
      border-bottom: 2px solid ${T.border};
      border-right: 1px solid #7C3AED88;
      white-space: normal; word-break: break-word;
      position: sticky; top: 0; z-index: 2;
      font-weight: 700; vertical-align: top;
    }
    .pt-table thead th .th-inner {
      padding: 6px 8px 3px; display: flex; align-items: center; gap: 3px;
      cursor: pointer; user-select: none; transition: color 0.12s;
    }
    .pt-table thead th .th-inner:hover { color: ${T.accent}; }
    .pt-table thead th.th-sorted { background: #0F1F30; }
    .pt-table thead th.th-sorted .th-inner { color: ${T.accent}; }
    .pt-table thead th .th-filter { padding: 2px 6px 5px; }
    .pt-table thead th .th-filter input {
      width: 100%; background: ${T.bg}; border: 1px solid ${T.border}55;
      border-radius: 4px; color: ${T.textMuted}; font-size: 10px;
      padding: 2px 5px; outline: none; font-family: 'Share Tech Mono', monospace;
      box-sizing: border-box;
    }
    .pt-table thead th .th-filter input:focus { border-color: ${T.accent}55; }
    .pt-table th.th-num-h { background: #0E1826; border-right: 1px solid ${T.border};
      width: ${numColW}px; min-width: ${numColW}px; text-align: center; position: sticky; top: 0;
      z-index: 3; vertical-align: top; padding: 0; }
    .pt-table td.td-num {
      background: #0E1826; border-right: 1px solid #7C3AED88;
      width: ${numColW}px; min-width: ${numColW}px; text-align: center;
      color: ${T.textDim}; font-size: 10px; padding: 7px 4px;
      white-space: nowrap; user-select: none;
    }
    .pt-table td {
      padding: 9px 10px; border-bottom: 1px solid ${T.border}22;
      border-right: 1px solid #7C3AED44;
      color: #C8D8E8; white-space: nowrap;
      overflow: hidden; text-overflow: ellipsis;
    }
    .pt-table tr:hover td, .pt-table tr:hover .td-num { background: ${T.rowHover}; }
    .pt-table tr.row-sel td { background: ${T.selRowBg} !important; color: ${T.selRowTxt}; }
    .pt-table tr.added-row td { background: #0A1F10 !important; }
    .pt-table tr.click-row { cursor: pointer; }
    .th-col-sel .th-inner { cursor: pointer; }
    .th-col-sel .th-inner:hover { color: ${T.error} !important; }
    .th-col-del { background: #2A0E0E !important; }
    .th-col-del .th-inner { color: ${T.error} !important; }
    .cell-rep { color: ${T.repeat} !important; font-style: italic; background: ${T.repeatBg}; }
    .header-input {
      background: transparent; border: none; border-bottom: 1px solid ${T.accent};
      color: ${T.accent}; font-family: 'Share Tech Mono', monospace;
      font-size: 10px; letter-spacing: 0.08em; text-transform: uppercase;
      width: 100%; min-width: 50px; outline: none; padding: 2px 0;
    }
    .ste-btn {
      font-family: 'Share Tech Mono', monospace; font-size: 11px; font-weight: 700;
      letter-spacing: 0.05em; padding: 7px 12px;
      border: 1px solid ${T.border2}; border-radius: 7px;
      cursor: pointer; background: ${T.bgCard}; color: ${T.textMuted};
      transition: all 0.15s; white-space: nowrap;
    }
    .ste-btn:hover { border-color: ${T.accent}; color: ${T.accent}; }
    .ste-btn.active { background: ${T.border}; border-color: ${T.accent}; color: ${T.accent}; }
    .ste-btn.danger { border-color: ${T.error}55; color: ${T.error}; }
    .ste-btn.danger:hover { background: #2A0E0E; }
    .dim-chip {
      display: flex; align-items: center; gap: 5px;
      font-size: 10px; letter-spacing: 0.06em; text-transform: uppercase;
      color: ${T.textDim}; cursor: pointer; padding: 7px 10px;
      border: 1px solid ${T.border2}; border-radius: 7px; background: ${T.bgCard};
      transition: all 0.15s; user-select: none;
    }
    .dim-chip:hover { border-color: ${T.accent}; }
    .dim-chip.on { color: #A78BFA; border-color: #3D2A6A; background: #160F2A; }
    .dim-dot { width: 7px; height: 7px; border-radius: 50%; background: currentColor; flex-shrink: 0; }
    .page-btn {
      background: #160F2A; border: 1px solid #3D2A6A; border-radius: 5px;
      color: #A78BFA; font-family: 'Share Tech Mono', monospace; font-size: 11px;
      padding: 4px 9px; cursor: pointer; transition: all 0.12s;
    }
    .page-btn:hover:not(:disabled) { border-color: #A78BFA; color: #D4BBFF; background: #1E1238; }
    .page-btn:disabled { opacity: 0.25; cursor: not-allowed; }
    .page-btn.cur { background: #7C3AED33; border-color: #A78BFA; color: #D4BBFF; font-weight: 700; }
    .pt-modal-overlay {
      position: fixed; inset: 0; background: #00000088; z-index: 200;
      display: flex; align-items: center; justify-content: center; padding: 20px;
    }
    .pt-modal {
      background: ${T.bgCard}; border: 1px solid ${T.border2}; border-radius: 14px;
      max-width: 500px; width: 100%; max-height: 80dvh; overflow-y: auto;
      box-shadow: 0 20px 60px #00000099;
    }
    .pt-modal-hdr {
      padding: 14px 16px; background: ${T.bgDark}; border-bottom: 1px solid ${T.border};
      border-radius: 14px 14px 0 0; display: flex; justify-content: space-between; align-items: center;
      position: sticky; top: 0;
    }
    .pt-modal-row {
      display: flex; gap: 10px; padding: 9px 16px; border-bottom: 1px solid ${T.border}22; align-items: flex-start;
    }
    .pt-modal-row:last-child { border-bottom: none; }
  `;

  return (
    <div style={{ flex: 1, display: "flex", flexDirection: "column", background: T.bg, paddingBottom: 64 }}>
      <style>{css}</style>
      {/* ── Barre unique : info + outils + retract ── */}
      <div style={{
        padding: "5px 12px",
        background: T.bgDark,
        borderBottom: `1px solid #7C3AED66`,
        display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap",
      }}>
        {/* Titre */}
        <span style={{ color: T.text, fontSize: 15, fontWeight: 800, letterSpacing: "-0.01em", whiteSpace: "nowrap" }}>Pointage</span>
        {/* Infos fichier/stats */}
        <PointageInfos />
        {/* Séparateur */}
        <span style={{ color: "#7C3AED99", fontSize: 12, margin: "0 2px" }}>|</span>
        {/* Boutons Outils + Mouvement */}
        <button
          className={`ste-btn${openToolbar ? " active" : ""}`}
          onClick={() => setOpenToolbar((o) => !o)}
          style={{ marginLeft: "auto", flexShrink: 0 }}
        >{openToolbar ? "✓ " : ""}Outils</button>
        <button
          className={`ste-btn${openMouvements ? " active" : ""}`}
          onClick={() => setOpenMouvements((o) => !o)}
          style={{ flexShrink: 0 }}
        >{openMouvements ? "✓ " : ""}Mouvement</button>
      </div>

      {/* Cadre Outils */}
      {openToolbar && (
        <div style={{
          margin: "8px 12px 0",
          padding: "10px 14px",
          background: T.bgCard,
          border: `1px solid #7C3AED66`,
          borderRadius: 10,
          display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center",
        }}>
          <div className={`dim-chip${dimRepeated ? " on" : ""}`} onClick={() => setDimRepeated((v) => !v)}>
            <span className="dim-dot" />Rép.
          </div>
          <div
            className={`dim-chip${autoRefFmt ? " on" : ""}`}
            onClick={() => { if (autoRefFmt) { setRefGroupInput(splitFormats["reference"] ?? ""); setRefGroupModal(true); } setAutoRefFmt((v) => !v); }}
            title="Formatage automatique des références"
          >🔢 Réf.</div>
          <button
            className="ste-btn"
            onClick={() => setPoidsUnit((u) => u === "t" ? "kg" : "t")}
            style={{ color: poidsUnit === "kg" ? T.warning : undefined, borderColor: poidsUnit === "kg" ? `${T.warning}66` : undefined }}
            title="Basculer tonnes / kilos"
          >⚖️ {poidsUnit}</button>
          <button
            className={`ste-btn${selectMode === "col" ? " active" : ""}`}
            onClick={() => { setSelectMode(selectMode === "col" ? "none" : "col"); setSelectedItems(new Set()); }}
          >{selectMode === "col" ? "✓ " : ""}Col.</button>
          <button
            className={`ste-btn${selectMode === "row" ? " active" : ""}`}
            onClick={() => { setSelectMode(selectMode === "row" ? "none" : "row"); setSelectedItems(new Set()); }}
          >{selectMode === "row" ? "✓ " : ""}Lig.</button>
          {selectedItems.size > 0 && (
            <button className="ste-btn danger" onClick={selectMode === "col" ? applyColAction : applyRowDeletion}>
              {selectMode === "col" ? `Masquer ${selectedItems.size} col.` : `Masquer ${selectedItems.size} lig.`}
            </button>
          )}
          {(hiddenCols.size > 0 || hiddenRows.size > 0) && (
            <button className="ste-btn" onClick={() => { _setHiddenCols(new Set()); _setHiddenRows(new Set()); }}>↺</button>
          )}
          {sortCol !== null && (
            <button className="ste-btn" onClick={() => { setSortCol(null); setPageIdx(0); }}>✕ Tri</button>
          )}
          {Object.values(colFilters).some((v) => v.trim()) && (
            <button className="ste-btn" onClick={() => { setColFilters({}); setPageIdx(0); }}>✕ Filtres</button>
          )}
          <button className="ste-btn" style={{ marginLeft: "auto", fontSize: 11 }} onClick={() => setOpenToolbar(false)}>✕</button>
        </div>
      )}

      {/* Cadre Mouvements */}
      {openMouvements && (
        <div style={{
          margin: "8px 12px 0",
          padding: "12px 14px",
          background: T.bgCard,
          border: `1px solid #7C3AED66`,
          borderRadius: 10,
        }}>
          {/* En-tête */}
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
            <span style={{ fontSize: 13, fontWeight: 700, color: "#A78BFA", letterSpacing: "0.04em" }}>Mouvements</span>
            <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
              <div
                onClick={() => setDoubleVerif(v => !v)}
                style={{
                  display: "flex", alignItems: "center", gap: 5, cursor: "pointer",
                  background: doubleVerif ? "#34D39922" : T.bgDark,
                  border: `1px solid ${doubleVerif ? "#34D399" : T.border2}`,
                  borderRadius: 20, padding: "3px 10px 3px 6px", userSelect: "none",
                }}
              >
                <div style={{
                  width: 28, height: 16, borderRadius: 8, position: "relative",
                  background: doubleVerif ? "#34D399" : T.border2, transition: "background 0.2s",
                }}>
                  <div style={{
                    position: "absolute", top: 2, left: doubleVerif ? 14 : 2,
                    width: 12, height: 12, borderRadius: "50%",
                    background: "#fff", transition: "left 0.2s",
                  }} />
                </div>
                <span style={{ fontSize: 10, color: doubleVerif ? "#34D399" : T.textDim, fontWeight: 600 }}>Double vérif</span>
              </div>
              <button className="ste-btn" onClick={() => setOpenMouvements(false)} style={{ fontSize: 11 }}>✕</button>
            </div>
          </div>

          {/* Pills destinations */}
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 10 }}>
            {destinations.map(d => (
              <button
                key={d.name}
                className="ste-btn"
                onClick={() => {
                  if (selectedDest === d.name && selectMode === "dest") {
                    setSelectedDest("");
                    setSelectMode("none");
                  } else {
                    setSelectedDest(d.name);
                    setSelectMode("dest");
                  }
                }}
                style={{
                  background: selectedDest === d.name && selectMode === "dest" ? d.color : `${d.color}22`,
                  borderColor: d.color,
                  color: selectedDest === d.name && selectMode === "dest" ? "#0F172A" : d.color,
                  fontWeight: selectedDest === d.name ? 700 : 400,
                  minWidth: 36,
                  outline: selectedDest === d.name && selectMode === "dest" ? `2px solid ${d.color}` : undefined,
                }}
                title={`Cliquer pour affecter les lignes à "${d.name}"`}
              >
                {d.name}
                {destStats.get(d.name) && (
                  <span style={{ marginLeft: 4, fontSize: 9, opacity: 0.85 }}>
                    ({destStats.get(d.name)!.count})
                  </span>
                )}
              </button>
            ))}
          </div>

          {/* Ajoute destination */}
          <div style={{ display: "flex", gap: 6, marginBottom: 10 }}>
            <input
              value={newDestName}
              onChange={e => setNewDestName(e.target.value)}
              onKeyDown={e => {
                if (e.key === "Enter" && newDestName.trim()) {
                  const name = newDestName.trim();
                  if (!destinations.some(d => d.name.toLowerCase() === name.toLowerCase())) {
                    const colors = ["#00c87a","#f447d1","#3cbefc","#ff9b2c","#a78bfa","#f87171","#34d399","#fbbf24"];
                    const color = colors[destinations.length % colors.length];
                    setDestinations(prev => [...prev, { name, color }]);
                  }
                  setNewDestName("");
                }
              }}
              placeholder="Nom destination"
              style={{
                flex: 1, background: T.bg, border: `1px solid ${T.border}`, borderRadius: 6,
                color: T.text, fontFamily: "'Share Tech Mono', monospace", fontSize: 12,
                padding: "4px 8px", outline: "none",
              }}
            />
            <button
              className="ste-btn"
              style={{ fontWeight: 700 }}
              onClick={() => {
                const name = newDestName.trim();
                if (!name) return;
                if (!destinations.some(d => d.name.toLowerCase() === name.toLowerCase())) {
                  const colors = ["#00c87a","#f447d1","#3cbefc","#ff9b2c","#a78bfa","#f87171","#34d399","#fbbf24"];
                  const color = colors[destinations.length % colors.length];
                  setDestinations(prev => [...prev, { name, color }]);
                }
                setNewDestName("");
              }}
            >+</button>
            {destinations.length > 0 && (
              <button
                className="ste-btn danger"
                title="Supprimer la destination sélectionnée"
                onClick={() => {
                  if (!selectedDest) return;
                  setDestinations(prev => prev.filter(d => d.name !== selectedDest));
                  setRowDestinations(prev => {
                    const next = new Map(prev);
                    for (const [k, v] of next) { if (v === selectedDest) next.delete(k); }
                    return next;
                  });
                  setSelectedDest("");
                }}
                disabled={!selectedDest}
              >✕ Dest.</button>
            )}
          </div>

          {/* Stats */}
          {destStats.size > 0 && (
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
              <thead>
                <tr>
                  <th style={{ textAlign: "left", color: T.textDim, padding: "2px 6px 4px 0", fontWeight: 600 }}>Dest.</th>
                  <th style={{ textAlign: "right", color: T.textDim, padding: "2px 6px 4px", fontWeight: 600 }}>Q</th>
                  {poidsColIdx >= 0 && <th style={{ textAlign: "right", color: T.textDim, padding: "2px 0 4px 6px", fontWeight: 600 }}>Poids ({poidsUnit})</th>}
                </tr>
              </thead>
              <tbody>
                {destinations.filter(d => destStats.has(d.name)).map(d => {
                  const s = destStats.get(d.name)!;
                  const poidsFmt = splitFormats["poids"] ?? "";
                  const fmtWeight = (w: number) => {
                    const val = poidsUnit === "kg" ? (w * 1000).toFixed(0) : parseFloat(w.toFixed(3)).toString();
                    return poidsFmt ? applyGrouping(val, poidsFmt) : thsep(val);
                  };
                  return (
                    <tr key={d.name}>
                      <td style={{ padding: "2px 6px 2px 0" }}>
                        <span style={{ background: `${d.color}33`, borderLeft: `3px solid ${d.color}`, padding: "1px 6px", borderRadius: 3, color: d.color, fontWeight: 600 }}>{d.name}</span>
                      </td>
                      <td style={{ textAlign: "right", padding: "2px 6px", color: T.text }}>{s.count}</td>
                      {poidsColIdx >= 0 && (
                        <td style={{ textAlign: "right", padding: "2px 0 2px 6px", color: T.warning }}>
                          {fmtWeight(s.weight)}
                        </td>
                      )}
                    </tr>
                  );
                })}
                {destStats.size > 1 && (() => {
                  const poidsFmt = splitFormats["poids"] ?? "";
                  const totalW = [...destStats.values()].reduce((a, b) => a + b.weight, 0);
                  const fmtTotal = () => {
                    const val = poidsUnit === "kg" ? (totalW * 1000).toFixed(0) : parseFloat(totalW.toFixed(3)).toString();
                    return poidsFmt ? applyGrouping(val, poidsFmt) : thsep(val);
                  };
                  return (
                  <tr style={{ borderTop: `1px solid ${T.border}` }}>
                    <td style={{ padding: "2px 6px 2px 0", color: T.textDim, fontStyle: "italic" }}>Total</td>
                    <td style={{ textAlign: "right", padding: "2px 6px", color: T.accent, fontWeight: 700 }}>
                      {[...destStats.values()].reduce((a, b) => a + b.count, 0)}
                    </td>
                    {poidsColIdx >= 0 && (
                      <td style={{ textAlign: "right", padding: "2px 0 2px 6px", color: T.accent, fontWeight: 700 }}>
                        {fmtTotal()}
                      </td>
                    )}
                  </tr>
                  );
                })()}
              </tbody>
            </table>
          )}

          {/* Effacer toutes les affectations */}
          {rowDestinations.size > 0 && (
            <button
              className="ste-btn danger"
              style={{ marginTop: 10, fontSize: 11 }}
              onClick={() => setConfirmClearDest(true)}
            >✖ Effacer toutes les affectations</button>
          )}
          {confirmClearDest && (
            <div
              onClick={() => setConfirmClearDest(false)}
              style={{
                position: "fixed", inset: 0, background: "#000000aa",
                zIndex: 500, display: "flex", alignItems: "center", justifyContent: "center",
              }}
            >
              <div
                onClick={(e) => e.stopPropagation()}
                style={{
                  background: T.bgCard, border: `1px solid ${T.error}66`,
                  borderRadius: 14, padding: "24px 20px", maxWidth: 320, width: "90%",
                  boxShadow: "0 20px 60px #00000099",
                }}
              >
                <div style={{ color: T.error, fontWeight: 800, fontSize: 15, marginBottom: 10 }}>⚠️ Confirmer la suppression</div>
                <div style={{ color: T.textMuted, fontSize: 13, marginBottom: 20 }}>
                  Toutes les affectations ({rowDestinations.size} ligne{rowDestinations.size > 1 ? "s" : ""}) seront effacées. Cette action est irréversible.
                </div>
                <div style={{ display: "flex", gap: 10 }}>
                  <button
                    className="ste-btn"
                    style={{ flex: 1 }}
                    onClick={() => setConfirmClearDest(false)}
                  >Annuler</button>
                  <button
                    className="ste-btn danger"
                    style={{ flex: 1 }}
                    onClick={() => { setRowDestinations(new Map()); setReassignedRows(new Map()); setConfirmClearDest(false); }}
                  >✖ Effacer</button>
                </div>
              </div>
            </div>
          )}
        </div>
      )}

      {/* Status hints */}
      <div style={{ padding: "6px 12px 0", display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
        {selectMode === "col" && (
          <div style={{ flex: "1 1 100%", marginBottom: 4, padding: "6px 10px", background: "#1F0A0A", border: `1px solid ${T.error}55`, borderRadius: 6, fontSize: 11, color: T.error }}>
            → Cliquez les en-têtes à masquer, puis confirmez
          </div>
        )}
        {selectMode === "row" && (
          <div style={{ flex: "1 1 100%", marginBottom: 4, padding: "6px 10px", background: "#1F0A0A", border: `1px solid ${T.error}55`, borderRadius: 6, fontSize: 11, color: T.error }}>
            → Cliquez les numéros de lignes à masquer, puis confirmez
          </div>
        )}
        {addedRows.length > 0 && (
          <span style={{ color: T.success, fontSize: 10 }}>+ {addedRows.length} ajoutée(s)</span>
        )}
        {pointedRows.size > 0 && (
          <span style={{ color: T.success, fontSize: 10, fontWeight: 700 }}>✅ {pointedRows.size} pointée(s)</span>
        )}
      </div>

      {/* Table */}
      <div style={{
        flex: 1, overflowX: "auto", overflowY: "auto",
        margin: "8px 12px 0", maxHeight: "calc(100dvh - 340px)",
        border: `1px solid ${T.border}`, borderRadius: 8, background: T.bgCard,
        WebkitOverflowScrolling: "touch",
      }}>
        <table className="pt-table">
          <thead>
            <tr>
              <th className="th-num-h">
                <div style={{ padding: "8px 4px 4px", textAlign: "center", color: T.textDim, fontSize: 9 }}>#</div>
                <div className="th-filter" />
              </th>
              {visibleCols.map((ci) => {
                const isSortedCol = sortCol === ci;
                const isSel = selectMode === "col" && selectedItems.has(ci);
                let cls = selectMode === "col" ? "th-col-sel" : "";
                if (isSel) cls += " th-col-del";
                if (isSortedCol) cls += " th-sorted";
                return (
                  <th key={ci} className={cls} style={{ width: colWidths.get(ci) ?? 60 }}>
                    <div className="th-inner" onClick={() => handleSortCol(ci)}>
                      {editingHeader === ci ? (
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
                        <span
                          onDoubleClick={() => selectMode === "none" && setEditingHeader(ci)}
                          title="Double-cliquer pour renommer"
                          style={{ flex: 1, ...(ci === poidsColIdx ? { whiteSpace: "nowrap", fontSize: 12 } : (ci === destIdx || ci === dchColIdx) ? { fontSize: 12 } : (ci === refColIdx || ci === rangColIdx) ? { fontSize: 13 } : {}) }}
                        >
                          {colLabel(ci)}
                        </span>
                      )}
                      {isSortedCol && <span style={{ fontSize: 9, flexShrink: 0 }}>{sortDir === "asc" ? "▲" : "▼"}</span>}
                      {!isSortedCol && <span style={{ fontSize: 8, flexShrink: 0, opacity: 0.2 }}>⇅</span>}
                    </div>
                    <div className="th-filter">
                      <input
                        value={colFilters[ci] ?? ""}
                        onChange={(e) => { setColFilters((p) => ({ ...p, [ci]: e.target.value })); setPageIdx(0); }}
                        placeholder="…"
                        onClick={(e) => e.stopPropagation()}
                      />
                    </div>
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {pageRows.map(({ row, ri }) => {
              const isRowSel = selectMode === "row" && selectedItems.has(ri);
              const isAdded  = ri < addedRows.length;
              const isPointed = pointedRows.has(ri);
              const rowDest = rowDestinations.get(ri);
              const destColor = rowDest ? (destinations.find(d => d.name === rowDest)?.color ?? null) : null;
              const rowStyle: React.CSSProperties = isPointed
                ? { background: "#0A1F10", outline: `1px solid ${T.success}44`, ...(destColor ? { borderLeft: `4px solid ${destColor}` } : {}) }
                : anomalyRowSet.has(ri)
                  ? { background: "#1E0A0A", outline: `1px solid ${T.error}33`, ...(destColor ? { borderLeft: `4px solid ${destColor}` } : {}) }
                  : destColor
                    ? { borderLeft: `4px solid ${destColor}`, background: `${destColor}33` }
                    : {};
              return (
                <tr
                  key={ri}
                  className={`${isRowSel ? "row-sel" : ""} ${isAdded ? "added-row" : ""} ${selectMode === "none" ? "click-row" : ""}`}
                  style={rowStyle}
                  onTouchStart={(e) => { touchStartPos.current = { x: e.touches[0].clientX, y: e.touches[0].clientY }; touchScrolled.current = false; }}
                  onTouchMove={(e) => { if (touchStartPos.current) { const dx = Math.abs(e.touches[0].clientX - touchStartPos.current.x); const dy = Math.abs(e.touches[0].clientY - touchStartPos.current.y); if (dx > 8 || dy > 8) touchScrolled.current = true; } }}
                  onClick={() => {
                    if (touchScrolled.current) return;
                    if (selectMode === "row") { toggleSelectItem(ri); return; }
                    if (selectMode === "dest") {
                      if (!selectedDest) { showToast("⚠️ Sélectionnez une destination d'abord", "error"); return; }
                        const existing = rowDestinations.get(ri);
                        if (existing && existing !== selectedDest) {
                          setConfirmReassign({ ri, from: existing, to: selectedDest });
                        } else {
                          if (doubleVerif) {
                            const poidsIdx2 = mapping.poids ? headers.indexOf(mapping.poids) : -1;
                            const refIdx2   = mapping.reference ? headers.indexOf(mapping.reference) : -1;
                            const getCell2  = (ci: number) => rowOverrides.get(ri)?.[ci] !== undefined ? String(rowOverrides.get(ri)![ci] ?? "") : String(row[ci] ?? "");
                            const refVal2   = refIdx2 >= 0 ? getCell2(refIdx2) : "—";
                            const poidsVal2 = poidsIdx2 >= 0 ? getCell2(poidsIdx2) : null;
                            if (poidsIdx2 >= 0 && poidsVal2) {
                              const realW = parseFloat(String(poidsVal2).replace(",", "."));
                              if (!isNaN(realW) && realW > 0) {
                                const genFake = (): number => {
                                  const pct = 0.02 + Math.random() * 0.07;
                                  const sign = Math.random() > 0.5 ? 1 : -1;
                                  return parseFloat((realW * (1 + sign * pct)).toFixed(realW < 10 ? 3 : 1));
                                };
                                let f1 = genFake(), f2 = genFake(), guard = 0;
                                while (guard++ < 20 && Math.abs(f1 - realW) < realW * 0.005) f1 = genFake();
                                guard = 0;
                                while (guard++ < 20 && (Math.abs(f2 - realW) < realW * 0.005 || Math.abs(f2 - f1) < realW * 0.005)) f2 = genFake();
                                const choices = [realW, f1, f2].sort(() => Math.random() - 0.5);
                                setDoubleVerifModal({ ri, ref: refVal2, correct: realW, choices, error: false, pendingDest: selectedDest });
                                return;
                              }
                            }
                          }
                          setRowDestinations((prev) => {
                            const next = new Map(prev);
                            if (next.get(ri) === selectedDest) next.delete(ri);
                            else next.set(ri, selectedDest);
                            return next;
                          });
                        }
                      return;
                    }
                    if (selectMode === "none" && !openMouvements) {
                      // detail modal disabled
                    }
                  }}
                >
                  <td className="td-num" title={isPointed ? "Pointé ✅" : isAdded ? "Ligne ajoutée" : undefined}>
                    {isPointed
                      ? <span style={{ color: T.success, fontSize: 11 }}>✅</span>
                      : isAdded
                        ? <span style={{ color: T.success, fontSize: 9 }}>+{ri + 1}</span>
                        : ri + 1}
                  </td>
                  {visibleCols.map((ci) => {
                    const rawCell = rowOverrides.get(ri)?.[ci] !== undefined ? rowOverrides.get(ri)![ci] : (row[ci] ?? null);
                    const strVal = rawCell !== null ? String(rawCell) : null;
                    const fmt    = getColFmt(ci);
                    let display: string | null;
                    if (strVal === null) {
                      display = null;
                    } else if (ci === poidsColIdx) {
                      const raw = parseFloat(strVal) || 0;
                      const val = poidsUnit === "kg" ? (raw * 1000).toFixed(0) : parseFloat(raw.toFixed(3)).toString();
                      display = (fmt ? applyGrouping(val, fmt) : thsep(val)) + " " + (poidsUnit === "kg" ? "kg" : "t");
                    } else if (ci === refColIdx && autoRefFmt) {
                      display = autoFormatRef(strVal, fmt);
                    } else {
                      display = fmt ? applyGrouping(strVal, fmt) : strVal;
                    }
                    const isRep  = !isAdded && dimRepeated && strVal !== null && (repetitiveByCol.get(ci)?.has(strVal) ?? false);
                    const isDestCol = ci === destIdx;
                    const destCellVal = isDestCol && rowDest ? rowDest : null;
                    const isRefCol = ci === refColIdx;
                    return (
                      <td
                        key={ci}
                        title={destCellVal ?? strVal ?? ""}
                        className={isRep ? "cell-rep" : ""}
                        onPointerDown={isRefCol ? (e) => { e.stopPropagation(); longPressRef.current = setTimeout(() => { setRefGroupInput(splitFormats["reference"] ?? ""); setRefGroupModal(true); }, 500); } : undefined}
                        onPointerUp={isRefCol ? () => { if (longPressRef.current) { clearTimeout(longPressRef.current); longPressRef.current = null; } } : undefined}
                        onPointerLeave={isRefCol ? () => { if (longPressRef.current) { clearTimeout(longPressRef.current); longPressRef.current = null; } } : undefined}
                      >
                        {isDestCol && destCellVal
                          ? <span style={{ color: destColor ?? T.accent, fontWeight: 700, fontSize: 11 }}>{destCellVal}</span>
                          : display ?? <span style={{ color: T.textDim }}>—</span>}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
            {pageRows.length === 0 && (
              <tr>
                <td colSpan={visibleCols.length + 1} style={{ textAlign: "center", padding: "32px", color: T.textDim }}>
                  Aucune ligne correspondante
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {/* Pagination */}
      <div style={{ margin: "8px 12px 4px", display: "flex", gap: 5, flexWrap: "wrap", alignItems: "center" }}>
        <button className="page-btn" onClick={() => setPageIdx(0)} disabled={safePage === 0}>«</button>
        <button className="page-btn" onClick={() => setPageIdx((p) => Math.max(0, p - 1))} disabled={safePage === 0}>‹</button>
        {Array.from({ length: totalPages }, (_, i) => i)
          .filter((i) => Math.abs(i - safePage) <= 2 || i === 0 || i === totalPages - 1)
          .reduce<(number | "…")[]>((acc, i, idx, arr) => {
            if (idx > 0 && i - (arr[idx - 1] as number) > 1) acc.push("…");
            acc.push(i);
            return acc;
          }, [])
          .map((item, idx) =>
            item === "…"
              ? <span key={`e${idx}`} style={{ color: "#7C3AED66", fontSize: 11 }}>…</span>
              : <button key={item} className={`page-btn${item === safePage ? " cur" : ""}`} onClick={() => setPageIdx(item as number)}>{(item as number) + 1}</button>
          )}
        <button className="page-btn" onClick={() => setPageIdx((p) => Math.min(totalPages - 1, p + 1))} disabled={safePage >= totalPages - 1}>›</button>
        <button className="page-btn" onClick={() => setPageIdx(totalPages - 1)} disabled={safePage >= totalPages - 1}>»</button>
        <span style={{ color: "#7C3AED88", fontSize: 10, margin: "0 4px" }}>|</span>
        <select
          value={pageSize}
          onChange={(e) => { setPageSize(Number(e.target.value)); setPageIdx(0); }}
          style={{ background: "#160F2A", border: "1px solid #3D2A6A", borderRadius: 5, color: "#A78BFA", fontFamily: "'Share Tech Mono', monospace", fontSize: 11, padding: "3px 6px", cursor: "pointer" }}
        >
          {PAGE_SIZES.map((s) => <option key={s} value={s}>{s} / page</option>)}
        </select>
        <span style={{ color: "#7C3AED99", fontSize: 10 }}>
          {safePage * pageSize + 1}–{Math.min((safePage + 1) * pageSize, sortedRows.length)} / {sortedRows.length}
        </span>
      </div>

      {/* Add row modal */}
      {showAddRow && <AddRowModal onClose={() => setShowAddRow(false)} />}

      {confirmUnpoint && (
        <div
          onClick={() => setConfirmUnpoint(null)}
          style={{ position: "fixed", inset: 0, background: "#000000aa", zIndex: 500, display: "flex", alignItems: "center", justifyContent: "center" }}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            style={{ background: T.bgCard, border: `1px solid ${T.error}66`, borderRadius: 14, padding: "22px 20px", maxWidth: 300, width: "90%", boxShadow: "0 20px 60px #00000099" }}
          >
            <div style={{ color: T.error, fontWeight: 800, fontSize: 14, marginBottom: 10 }}>❌ Dépointer la ligne ?</div>
            <div style={{ color: T.textMuted, fontSize: 12, marginBottom: 18 }}>
              Retirer le pointage de <strong style={{ color: T.accent, fontFamily: "monospace" }}>{confirmUnpoint.ref}</strong> ?<br />
              <span style={{ color: T.textDim, fontSize: 11 }}>Cette action est réversible.</span>
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button className="ste-btn" style={{ flex: 1 }} onClick={() => setConfirmUnpoint(null)}>Annuler</button>
              <button className="ste-btn danger" style={{ flex: 1 }} onClick={() => {
                setPointedRows(prev => { const next = new Set(prev); next.delete(confirmUnpoint!.ri); return next; });
                setConfirmUnpoint(null);
                setModalRow(null);
              }}>Dépointer</button>
            </div>
          </div>
        </div>
      )}

      {confirmReassign && (
        <div
          onClick={() => setConfirmReassign(null)}
          style={{ position: "fixed", inset: 0, background: "#000000aa", zIndex: 500, display: "flex", alignItems: "center", justifyContent: "center" }}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            style={{ background: T.bgCard, border: `1px solid ${T.warning}66`, borderRadius: 14, padding: "22px 20px", maxWidth: 300, width: "90%", boxShadow: "0 20px 60px #00000099" }}
          >
            <div style={{ color: T.warning, fontWeight: 800, fontSize: 14, marginBottom: 10 }}>⚠️ Ligne déjà affectée</div>
            <div style={{ color: T.textMuted, fontSize: 12, marginBottom: 18 }}>
              Cette ligne est déjà affectée à <strong style={{ color: (destinations.find(d => d.name === confirmReassign.from)?.color ?? T.accent) }}>{confirmReassign.from}</strong>.<br />
              Réaffecter à <strong style={{ color: (destinations.find(d => d.name === confirmReassign.to)?.color ?? T.accent) }}>{confirmReassign.to}</strong> ?
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button className="ste-btn" style={{ flex: 1 }} onClick={() => setConfirmReassign(null)}>Annuler</button>
              <button className="ste-btn danger" style={{ flex: 1 }} onClick={() => {
                setRowDestinations((prev) => { const next = new Map(prev); next.set(confirmReassign.ri, confirmReassign.to); return next; });
                setReassignedRows((prev: Map<number, { from: string; to: string }[]>) => {
                  const next = new Map(prev);
                  const history: { from: string; to: string }[] = next.get(confirmReassign.ri) ?? [];
                  next.set(confirmReassign.ri, [...history, { from: confirmReassign.from, to: confirmReassign.to }]);
                  return next;
                });
                setConfirmReassign(null);
              }}>Réaffecter</button>
            </div>
          </div>
        </div>
      )}

      {/* Row detail / affectation modal */}
      {modalRow && (() => {
        const { row, rowNum, ri } = modalRow;
        const refIdx   = mapping.reference ? headers.indexOf(mapping.reference) : -1;
        const rangIdx  = mapping.rang      ? headers.indexOf(mapping.rang)      : -1;
        const poidsIdx = mapping.poids     ? headers.indexOf(mapping.poids)     : -1;
        const getCell  = (ci: number) => rowOverrides.get(ri)?.[ci] !== undefined ? String(rowOverrides.get(ri)![ci] ?? "") : String(row[ci] ?? "");
        const refVal   = refIdx  >= 0 ? getCell(refIdx)  : "—";
        const rangVal  = rangIdx >= 0 ? getCell(rangIdx) : null;
        const poidsVal = poidsIdx >= 0 ? getCell(poidsIdx) : null;
        const isPointed = pointedRows.has(ri);
        const togglePointed = () => setPointedRows((prev) => {
          const next = new Set(prev); next.has(ri) ? next.delete(ri) : next.add(ri); return next;
        });
        const isEmptyRow = row.every((cell, ci) => {
          const v = rowOverrides.get(ri)?.[ci] !== undefined ? String(rowOverrides.get(ri)![ci] ?? "") : String(cell ?? "");
          return v.trim() === "";
        });
        const handlePointerClick = () => {
          if (isEmptyRow) return;
          if (isPointed) {
            setConfirmUnpoint({ ri, ref: refVal });
            return;
          }
          togglePointed();
        };
        const setOverride = (ci: number, val: string) => setRowOverrides((prev) => {
          const next = new Map(prev);
          const r = { ...(next.get(ri) ?? {}) }; r[ci] = val; next.set(ri, r); return next;
        });
        // extra cols (col vide = nouvelle colonne ajoutée à la fin)
        const extraCols = extras.map((ex) => ({ ...ex, idx: headers.indexOf(ex.col) }));
        return (
          <div className="pt-modal-overlay" onClick={() => setModalRow(null)}>
            <div className="pt-modal" onClick={(e) => e.stopPropagation()}>
              {/* Header */}
              <div className="pt-modal-hdr">
                <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                  <span style={{ color: T.textDim, fontSize: 9, letterSpacing: "0.12em", textTransform: "uppercase" }}>Ligne #{rowNum}</span>
                  <span style={{ color: T.accent, fontWeight: 900, fontSize: 15, fontFamily: "'Share Tech Mono', monospace", letterSpacing: "0.06em" }}>🏷 {refVal}</span>
                </div>
                <button onClick={() => setModalRow(null)} style={{ background: "none", border: "none", color: T.textDim, fontSize: 22, cursor: "pointer", lineHeight: 1, padding: 0 }}>×</button>
              </div>

              {/* Pointer toggle */}
              <div style={{ padding: "10px 16px", borderBottom: `1px solid ${T.border}22`, display: "flex", alignItems: "center", gap: 10 }}>
                <button
                  onClick={handlePointerClick}
                  style={{
                    padding: "8px 18px", borderRadius: 8, fontWeight: 800, fontSize: 13,
                    fontFamily: "'Share Tech Mono', monospace", transition: "all 0.15s",
                    cursor: isEmptyRow ? "not-allowed" : "pointer",
                    opacity: isEmptyRow ? 0.35 : 1,
                    background: isPointed ? `${T.success}22` : T.bgDark,
                    border: `2px solid ${isPointed ? T.success : T.border2}`,
                    color: isPointed ? T.success : T.textMuted,
                    boxShadow: isPointed ? `0 0 12px ${T.success}44` : "none",
                  }}
                >{isEmptyRow ? "⊘ Ligne vide" : isPointed ? "✅ Pointé" : "○ Pointer"}</button>
                {isPointed && <span style={{ color: T.success, fontSize: 11 }}>✔ Confirmé dans l'état</span>}
                {isEmptyRow && <span style={{ color: T.textDim, fontSize: 11 }}>Impossible de pointer une ligne vide</span>}
              </div>

              {/* Champs clés */}
              {(rangVal !== null || poidsVal !== null) && (
                <div style={{ display: "flex", gap: 0, borderBottom: `1px solid ${T.border}22` }}>
                  {rangVal !== null && (
                    <div style={{ flex: 1, padding: "8px 16px", borderRight: `1px solid ${T.border}22` }}>
                      <div style={{ color: T.textDim, fontSize: 9, textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: 3 }}>📍 Rang</div>
                      <div style={{ color: T.textMuted, fontFamily: "monospace", fontSize: 13, fontWeight: 700 }}>{rangVal || "—"}</div>
                    </div>
                  )}
                  {poidsVal !== null && (
                    <div style={{ flex: 1, padding: "8px 16px" }}>
                      <div style={{ color: T.textDim, fontSize: 9, textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: 3 }}>⚖️ Poids</div>
                      <div style={{ color: T.warning, fontFamily: "monospace", fontSize: 13, fontWeight: 700 }}>{poidsVal || "—"}</div>
                    </div>
                  )}
                </div>
              )}

              {/* Affectation extras */}
              {extraCols.length > 0 && (
                <div style={{ padding: "10px 16px", borderBottom: `1px solid ${T.border}22` }}>
                  <div style={{ color: "#A78BFA", fontSize: 9, textTransform: "uppercase", letterSpacing: "0.12em", fontWeight: 700, marginBottom: 8 }}>Affectation</div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 7 }}>
                    {extraCols.map((ex) => {
                      const ci = ex.idx >= 0 ? ex.idx : -1;
                      return (
                        <div key={ex.label} style={{ display: "flex", alignItems: "center", gap: 8 }}>
                          <span style={{ minWidth: 80, color: "#A78BFA", fontSize: 11, fontWeight: 700, fontFamily: "'Share Tech Mono', monospace" }}>{ex.label}</span>
                          <input
                            value={ci >= 0 ? (rowOverrides.get(ri)?.[ci] !== undefined ? String(rowOverrides.get(ri)![ci] ?? "") : String(row[ci] ?? "")) : ""}
                            onChange={(e) => ci >= 0 && setOverride(ci, e.target.value)}
                            placeholder={ci < 0 ? "Colonne non mappée" : "—"}
                            disabled={ci < 0}
                            style={{
                              flex: 1, background: ci < 0 ? T.bgDark : T.bgCard,
                              border: `1px solid ${ci < 0 ? T.border : "#7C3AED88"}`,
                              borderRadius: 7, color: ci < 0 ? T.textDim : T.text,
                              fontSize: 12, padding: "6px 10px", outline: "none",
                              fontFamily: "'Share Tech Mono', monospace",
                            }}
                          />
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}

              {/* Toutes les autres colonnes */}
              <div style={{ maxHeight: 260, overflowY: "auto" }}>
                {visibleCols.map((ci) => {
                  const val    = row[ci];
                  const strVal = val !== null && val !== undefined ? String(val) : null;
                  const fmt    = getColFmt(ci);
                  const display = strVal !== null ? (fmt ? applyGrouping(strVal, fmt) : strVal) : null;
                  return (
                    <div key={ci} className="pt-modal-row">
                      <span style={{ minWidth: 120, color: T.textDim, fontSize: 11, textTransform: "uppercase", letterSpacing: "0.06em", flexShrink: 0, paddingTop: 2 }}>
                        {colLabel(ci)}
                      </span>
                      <span style={{ color: display ? T.text : T.textDim, fontSize: 13, fontFamily: "monospace", wordBreak: "break-all" }}>
                        {display ?? "—"}
                      </span>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        );
      })()}

      {/* Double Vérif modal */}
      {doubleVerifModal && (() => {
        const { ref, correct, choices, error } = doubleVerifModal;
        const poidsFmt = splitFormats["poids"] ?? "";
        const unitLabel = poidsUnit === "kg" ? "kg" : "t";
        const fmtChoice = (w: number) => {
          const val = poidsUnit === "kg" ? (w * 1000).toFixed(0) : parseFloat(w.toFixed(3)).toString();
          return (poidsFmt ? applyGrouping(val, poidsFmt) : thsep(val)) + " " + unitLabel;
        };
        const handleChoice = (w: number) => {
          if (error) return;
          const isCorrect = Math.abs(w - correct) < 1e-9;
          if (isCorrect) {
            if (doubleVerifModal!.pendingDest) {
              setRowDestinations(prev => { const next = new Map(prev); next.set(doubleVerifModal!.ri, doubleVerifModal!.pendingDest!); return next; });
            } else {
              setPointedRows(prev => { const next = new Set(prev); next.add(doubleVerifModal!.ri); return next; });
            }
            setDoubleVerifModal(null);
          } else {
            setDoubleVerifModal(prev => prev ? { ...prev, error: true } : null);
            setTimeout(() => setDoubleVerifModal(null), 5000);
          }
        };
        return (
          <div className="pt-modal-overlay" style={{ zIndex: 1100 }} onClick={() => !error && setDoubleVerifModal(null)}>
            <div className="pt-modal" onClick={e => e.stopPropagation()} style={{ maxWidth: 320 }}>
              <div className="pt-modal-hdr">
                <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                  <span style={{ color: T.textDim, fontSize: 9, letterSpacing: "0.12em", textTransform: "uppercase" }}>Double vérification</span>
                  <span style={{ color: T.accent, fontWeight: 900, fontSize: 15, fontFamily: "'Share Tech Mono', monospace" }}>🏷 {ref}</span>
                </div>
                {!error && <button onClick={() => setDoubleVerifModal(null)} style={{ background: "none", border: "none", color: T.textDim, fontSize: 22, cursor: "pointer", lineHeight: 1, padding: 0 }}>×</button>}
              </div>
              {error ? (
                <div style={{ padding: "24px 16px", textAlign: "center" }}>
                  <div style={{ fontSize: 36, marginBottom: 12 }}>❌</div>
                  <div style={{ color: T.error, fontWeight: 800, fontSize: 14, marginBottom: 6 }}>Poids incorrect !</div>
                  <div style={{ color: T.textDim, fontSize: 11 }}>Fermeture automatique dans 5 secondes…</div>
                </div>
              ) : (
                <div style={{ padding: "16px" }}>
                  <div style={{ color: T.textDim, fontSize: 11, marginBottom: 14, textAlign: "center" }}>
                    Sélectionnez le <strong style={{ color: T.text }}>poids correct</strong> pour valider le pointage
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                    {choices.map((w, i) => (
                      <button key={i} onClick={() => handleChoice(w)} style={{
                        padding: "12px 16px", borderRadius: 10, cursor: "pointer",
                        background: T.bgDark, border: `2px solid ${T.border2}`,
                        color: T.warning, fontFamily: "'Share Tech Mono', monospace",
                        fontSize: 16, fontWeight: 700, textAlign: "center",
                        transition: "all 0.1s",
                      }}
                        onMouseEnter={e => { (e.currentTarget as HTMLButtonElement).style.borderColor = T.warning; (e.currentTarget as HTMLButtonElement).style.background = `${T.warning}11`; }}
                        onMouseLeave={e => { (e.currentTarget as HTMLButtonElement).style.borderColor = T.border2; (e.currentTarget as HTMLButtonElement).style.background = T.bgDark; }}
                      >{fmtChoice(w)}</button>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </div>
        );
      })()}

      {refGroupModal && (
        <div style={{ position: "fixed", inset: 0, background: "#000a", zIndex: 2000, display: "flex", flexDirection: "column", justifyContent: "flex-end" }} onClick={() => setRefGroupModal(false)}>
          <div style={{ background: T.bgCard, borderRadius: "18px 18px 0 0", padding: "20px 16px 32px", boxShadow: "0 -4px 40px #0008" }} onClick={e => e.stopPropagation()}>
            <div style={{ color: T.text, fontWeight: 700, fontSize: 14, marginBottom: 4 }}>🏷 Format de groupement — REF</div>
            <div style={{ color: T.textDim, fontSize: 11, marginBottom: 12 }}>Séparateur visuel (ex: "3 2 3" → ABC DE FGH)</div>
            <input
              style={{ width: "100%", background: T.bgDark, border: `1px solid ${T.border2}`, borderRadius: 8, padding: "10px 12px", color: T.text, fontSize: 13, fontFamily: "monospace", boxSizing: "border-box" }}
              value={refGroupInput}
              onChange={e => setRefGroupInput(e.target.value)}
              placeholder="ex: 3 2 3"
              autoFocus
              onKeyDown={e => {
                if (e.key === "Enter") { setSplitFormats(p => ({ ...p, reference: refGroupInput.trim() })); setRefGroupModal(false); }
                if (e.key === "Escape") setRefGroupModal(false);
              }}
            />
            <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
              <button className="ste-btn" style={{ flex: 1 }} onClick={() => { setSplitFormats(p => { const n = { ...p }; delete n["reference"]; return n; }); setRefGroupModal(false); }}>🗑 Supprimer</button>
              <button className="ste-btn" style={{ flex: 2, background: T.accent, color: "#0F172A", fontWeight: 700 }} onClick={() => { setSplitFormats(p => ({ ...p, reference: refGroupInput.trim() })); setRefGroupModal(false); }}>✔ Appliquer</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────
// 7. PAGE EXPORT
// ─────────────────────────────────────────────

// Sauvegarde l'onglet actif dans sheetStates avant export / résumé
function useFlushActiveSheet() {
  const { activeSheet, parsed, headers, addedRows, hiddenCols, hiddenRows, sheetStates } = useApp();
  return useCallback(() => {
    if (activeSheet && parsed) {
      sheetStates.current.set(activeSheet, { parsed, headers, addedRows, hiddenCols, hiddenRows });
    }
  }, [activeSheet, parsed, headers, addedRows, hiddenCols, hiddenRows, sheetStates]);
}

// ─────────────────────────────────────────────
// 7. PAGE RAPPORT
// ─────────────────────────────────────────────

function RapportPage() {
  const {
    parsed, headers, allRows, hiddenRows,
    mapping, extras, splitFormats,
    pointedRows,
    destinations, rowDestinations, reassignedRows,
    poidsUnit, autoRefFmt,
  } = useApp();

  const [openResume, setOpenResume] = useState(false);
  const [openMouvements, setOpenMouvements] = useState(true);
  const [openChargement, setOpenChargement] = useState(false);
  const [openDechargement, setOpenDechargement] = useState(false);
  const [openTally, setOpenTally] = useState(false);
  const [openPointees, setOpenPointees] = useState(true);
  const [openReaff, setOpenReaff] = useState(true);

  function rLs<T>(key: string, fb: T): T { try { const v = localStorage.getItem(key); return v !== null ? JSON.parse(v) as T : fb; } catch { return fb; } }
  function wLs(key: string, v: unknown) { try { localStorage.setItem(key, JSON.stringify(v)); } catch {} }

  const [tallyPrev, setTallyPrevRaw] = useState<Record<string, { qty: string; weight: string }>>(() => rLs("ste_tallyPrev", {}));
  const setTallyPrev: typeof setTallyPrevRaw = (v) => { setTallyPrevRaw(prev => { const next = typeof v === "function" ? v(prev) : v; wLs("ste_tallyPrev", next); return next; }); };

  const [chargementMaxi, setChargementMaxiRaw] = useState<Record<string, { qty: string; weight: string }>>(() => rLs("ste_chargementMaxi", {}));
  const setChargementMaxi: typeof setChargementMaxiRaw = (v) => { setChargementMaxiRaw(prev => { const next = typeof v === "function" ? v(prev) : v; wLs("ste_chargementMaxi", next); return next; }); };

  const [dechargementMaxi, setDechargementMaxiRaw] = useState<Record<string, { qty: string; weight: string }>>(() => rLs("ste_dechargementMaxi", {}));
  const setDechargementMaxi: typeof setDechargementMaxiRaw = (v) => { setDechargementMaxiRaw(prev => { const next = typeof v === "function" ? v(prev) : v; wLs("ste_dechargementMaxi", next); return next; }); };

  if (!parsed) {
    return (
      <div style={{ flex: 1, overflowY: "auto", paddingBottom: 80, background: T.bg }}>
        <PageHeader title="Rapport" />
        <EmptyState icon="📋" text="Aucun fichier chargé" sub="Importez un fichier Excel d'abord" />
      </div>
    );
  }

  const poidsIdx    = mapping.poids ? headers.indexOf(mapping.poids) : -1;
  const refIdx      = mapping.reference ? headers.indexOf(mapping.reference) : -1;
  const rangIdx     = mapping.rang ? headers.indexOf(mapping.rang) : -1;
  const poidsFmt    = splitFormats["poids"] ?? "";
  const unitLabel   = poidsUnit === "kg" ? "kg" : "t";

  const fmtWeight = (w: number): string => {
    const val = poidsUnit === "kg" ? (w * 1000).toFixed(0) : parseFloat(w.toFixed(3)).toString();
    return (poidsFmt ? applyGrouping(val, poidsFmt) : thsep(val)) + " " + unitLabel;
  };

  const visibleRows = allRows.map((row, ri) => ({ row, ri })).filter(({ ri }) => !hiddenRows.has(ri));

  // Stats par destination
  const excludedFromReport = new Set(destinations.filter(d => d.excludeFromReport).map(d => d.name));
  const destStatsMap = new Map<string, { count: number; weight: number }>();
  for (const { row, ri } of visibleRows) {
    const dest = rowDestinations.get(ri);
    if (!dest || excludedFromReport.has(dest)) continue;
    const w = poidsIdx >= 0 ? (parseFloat(String(row[poidsIdx] ?? "")) || 0) : 0;
    const s = destStatsMap.get(dest) ?? { count: 0, weight: 0 };
    destStatsMap.set(dest, { count: s.count + 1, weight: s.weight + w });
  }

  const totalPointed = pointedRows.size;
  const totalAffected = [...rowDestinations.entries()].filter(([, d]) => !excludedFromReport.has(d)).length;
  const totalRows = visibleRows.length;
  const totalWeight = poidsIdx >= 0
    ? visibleRows.reduce((acc, { row }) => acc + (parseFloat(String(row[poidsIdx] ?? "")) || 0), 0)
    : 0;
  const affectedWeight = poidsIdx >= 0
    ? [...rowDestinations.entries()].reduce((acc, [ri, d]) => {
        if (excludedFromReport.has(d)) return acc;
        const row = allRows[ri];
        return acc + (row ? (parseFloat(String(row[poidsIdx] ?? "")) || 0) : 0);
      }, 0)
    : 0;

  const extraCols = extras.filter(e => e.col && e.label.trim());

  const SectionHeader = ({ label, open, toggle, color }: { label: string; open: boolean; toggle: () => void; color: string }) => (
    <div
      onClick={toggle}
      style={{ display: "flex", alignItems: "center", cursor: "pointer", userSelect: "none",
        color, fontWeight: 800, fontSize: 11, textTransform: "uppercase", letterSpacing: "0.08em",
        marginBottom: open ? 10 : 0 }}
    >
      <span style={{ flex: 1 }}>{label}</span>
      <span style={{ fontSize: 12, opacity: 0.55, transition: "transform 0.2s", display: "inline-block", transform: open ? "rotate(90deg)" : "rotate(0deg)" }}>▶</span>
    </div>
  );

  return (
    <div style={{ flex: 1, overflowY: "auto", paddingBottom: 80, background: T.bg }}>
      <div style={{ position: "relative" }}>
        <PageHeader title="Rapport" subtitle="Synthèse pointage & mouvements" />
        <button
          onClick={() => {
            setTallyPrev({});
            setChargementMaxi({});
            setDechargementMaxi({});
            ["ste_tallyPrev", "ste_chargementMaxi", "ste_dechargementMaxi"].forEach(k => { try { localStorage.removeItem(k); } catch {} });
          }}
          title="Effacer toutes les valeurs saisies manuellement"
          style={{ position: "absolute", top: 14, right: 12, background: `${T.error}22`, border: `1px solid ${T.error}55`, borderRadius: 8, color: T.error, fontSize: 16, padding: "4px 10px", cursor: "pointer", lineHeight: 1 }}
        >🗑</button>
      </div>

      <div style={{ padding: "12px 12px 0" }}>

        {/* Tableau par destination - Mouvements */}
        {destStatsMap.size > 0 && (
          <div style={{ background: T.bgCard, borderRadius: 12, border: `1px solid #7C3AED66`, padding: "12px 14px", marginBottom: 12 }}>
            <SectionHeader label="🏗️ Mouvements" open={openMouvements} toggle={() => setOpenMouvements(o => !o)} color="#A78BFA" />
            {openMouvements && <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
              <thead>
                <tr style={{ borderBottom: `1px solid ${T.border}` }}>
                  <th style={{ textAlign: "left", color: T.textDim, padding: "3px 8px 6px 0", fontWeight: 600 }}>Dest.</th>
                  <th style={{ textAlign: "right", color: T.textDim, padding: "3px 6px 6px", fontWeight: 600 }}>Qté</th>
                  {poidsIdx >= 0 && <th style={{ textAlign: "right", color: T.textDim, padding: "3px 0 6px 6px", fontWeight: 600 }}>Poids ({unitLabel})</th>}
                </tr>
              </thead>
              <tbody>
                {destinations.filter(d => destStatsMap.has(d.name)).map(d => {
                  const s = destStatsMap.get(d.name)!;
                  return (
                    <tr key={d.name} style={{ borderBottom: `1px solid ${T.border2}33` }}>
                      <td style={{ padding: "5px 8px 5px 0" }}>
                        <span style={{ background: `${d.color}33`, borderLeft: `3px solid ${d.color}`, padding: "2px 8px", borderRadius: 4, color: d.color, fontWeight: 700 }}>{d.name}</span>
                      </td>
                      <td style={{ textAlign: "right", padding: "5px 6px", color: T.text, fontFamily: "monospace" }}>{s.count}</td>
                      {poidsIdx >= 0 && <td style={{ textAlign: "right", padding: "5px 0 5px 6px", color: T.warning, fontFamily: "monospace" }}>{fmtWeight(s.weight)}</td>}
                    </tr>
                  );
                })}
                {destStatsMap.size > 1 && (
                  <tr style={{ borderTop: `1px solid ${T.border}` }}>
                    <td style={{ padding: "5px 8px 5px 0", color: T.textDim, fontStyle: "italic" }}>Total affecté</td>
                    <td style={{ textAlign: "right", padding: "5px 6px", color: T.accent, fontWeight: 700, fontFamily: "monospace" }}>
                      {[...destStatsMap.values()].reduce((a, b) => a + b.count, 0)}
                    </td>
                    {poidsIdx >= 0 && (
                      <td style={{ textAlign: "right", padding: "5px 0 5px 6px", color: T.accent, fontWeight: 700, fontFamily: "monospace" }}>
                        {fmtWeight([...destStatsMap.values()].reduce((a, b) => a + b.weight, 0))}
                      </td>
                    )}
                  </tr>
                )}
              </tbody>
            </table>}
          </div>
        )}

        {/* Tally */}
        <div style={{ background: T.bgCard, borderRadius: 12, border: `1px solid #05966955`, padding: "12px 14px", marginBottom: 12 }}>
          <SectionHeader label="📋 Tally" open={openTally} toggle={() => setOpenTally(o => !o)} color="#34D399" />
          {openTally && (() => {
            const tallyDests = destinations.filter(d => !d.excludeFromReport);
            const setPrev = (name: string, field: "qty" | "weight", val: string) =>
              setTallyPrev(p => ({ ...p, [name]: { qty: p[name]?.qty ?? "", weight: p[name]?.weight ?? "", [field]: val } }));
            const thStyle = (align: "left" | "right" | "center" = "right"): React.CSSProperties => ({
              textAlign: align, color: T.textDim, padding: "3px 5px 5px", fontWeight: 600, fontSize: 10, whiteSpace: "nowrap"
            });
            const tdStyle = (color: string = T.text): React.CSSProperties => ({
              textAlign: "right", padding: "4px 5px", fontFamily: "monospace", fontSize: 10, color
            });
            let sumTodayQty = 0, sumTodayW = 0, sumPrevQty = 0, sumPrevW = 0;
            for (const d of tallyDests) {
              const s = destStatsMap.get(d.name);
              sumTodayQty += s?.count ?? 0;
              sumTodayW   += s?.weight ?? 0;
              sumPrevQty  += parseFloat(tallyPrev[d.name]?.qty ?? "0") || 0;
              sumPrevW    += parseFloat(String(tallyPrev[d.name]?.weight ?? "0").replace(",", ".")) || 0;
            }
            return (
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                  <thead>
                    <tr style={{ borderBottom: `1px solid ${T.border}44` }}>
                      <th style={{ ...thStyle("left"), padding: "3px 6px 0 0" }} rowSpan={2}>Dest.</th>
                      <th colSpan={2} style={{ ...thStyle(), color: "#60A5FA", borderBottom: `1px solid #60A5FA44`, paddingBottom: 2 }}>TODAY</th>
                      <th colSpan={2} style={{ ...thStyle(), color: "#FB923C", borderBottom: `1px solid #FB923C44`, paddingBottom: 2 }}>PREV.</th>
                      <th colSpan={2} style={{ ...thStyle(), color: "#34D399", borderBottom: `1px solid #34D39944`, paddingBottom: 2 }}>TOTAL</th>
                    </tr>
                    <tr style={{ borderBottom: `1px solid ${T.border}` }}>
                      <th style={thStyle()}>Qté</th><th style={thStyle()}>Poids</th>
                      <th style={{ ...thStyle(), color: "#FB923C" }}>Qté</th><th style={{ ...thStyle(), color: "#FB923C" }}>Poids</th>
                      <th style={{ ...thStyle(), color: "#34D399" }}>Qté</th><th style={{ ...thStyle(), color: "#34D399" }}>Poids</th>
                    </tr>
                  </thead>
                  <tbody>
                    {tallyDests.map(d => {
                      const s = destStatsMap.get(d.name);
                      const tQty = s?.count ?? 0;
                      const tW   = s?.weight ?? 0;
                      const pQty = parseFloat(tallyPrev[d.name]?.qty ?? "0") || 0;
                      const pW   = parseFloat(String(tallyPrev[d.name]?.weight ?? "0").replace(",", ".")) || 0;
                      const totQty = tQty + pQty;
                      const totW   = tW + pW;
                      const inputStyle: React.CSSProperties = {
                        width: 58, textAlign: "right", background: T.bgDark,
                        border: `1px solid ${T.border2}`, borderRadius: 4,
                        color: "#FB923C", fontFamily: "monospace", fontSize: 10,
                        padding: "2px 4px"
                      };
                      return (
                        <tr key={d.name} style={{ borderBottom: `1px solid ${T.border2}22` }}>
                          <td style={{ padding: "4px 6px 4px 0" }}>
                            <span style={{ background: `${d.color}22`, borderLeft: `3px solid ${d.color}`, padding: "1px 6px", borderRadius: 3, color: d.color, fontWeight: 700, fontSize: 10 }}>{d.name}</span>
                          </td>
                          <td style={tdStyle("#60A5FA")}>{tQty || "—"}</td>
                          <td style={tdStyle("#60A5FA")}>{tW > 0 ? fmtWeight(tW) : "—"}</td>
                          <td style={{ padding: "3px 5px", textAlign: "right" }}>
                            <input type="number" min={0} style={inputStyle}
                              value={tallyPrev[d.name]?.qty ?? ""}
                              onChange={e => setPrev(d.name, "qty", e.target.value)}
                              placeholder="0" />
                          </td>
                          <td style={{ padding: "3px 5px", textAlign: "right" }}>
                            <input type="text" inputMode="decimal" style={{ ...inputStyle, width: 70 }}
                              value={(() => { const raw = tallyPrev[d.name]?.weight ?? ""; const n = parseFloat(raw.replace(",", ".")); return isNaN(n) ? raw : thsep(String(n)); })()}
                              onChange={e => setPrev(d.name, "weight", e.target.value.replace(/\s/g, ""))}
                              placeholder="0" />
                          </td>
                          <td style={tdStyle("#34D399")}>{totQty || "—"}</td>
                          <td style={tdStyle("#34D399")}>{totW > 0 ? fmtWeight(totW) : "—"}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                  <tfoot>
                    <tr style={{ borderTop: `1px solid ${T.border}` }}>
                      <td style={{ padding: "5px 6px 4px 0", color: T.textDim, fontSize: 10, fontStyle: "italic" }}>Total</td>
                      <td style={tdStyle("#60A5FA")}><strong>{sumTodayQty || "—"}</strong></td>
                      <td style={tdStyle("#60A5FA")}><strong>{sumTodayW > 0 ? fmtWeight(sumTodayW) : "—"}</strong></td>
                      <td style={tdStyle("#FB923C")}><strong>{sumPrevQty || "—"}</strong></td>
                      <td style={tdStyle("#FB923C")}><strong>{sumPrevW > 0 ? fmtWeight(sumPrevW) : "—"}</strong></td>
                      <td style={tdStyle("#34D399")}><strong>{(sumTodayQty + sumPrevQty) || "—"}</strong></td>
                      <td style={tdStyle("#34D399")}><strong>{(sumTodayW + sumPrevW) > 0 ? fmtWeight(sumTodayW + sumPrevW) : "—"}</strong></td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            );
          })()}
        </div>

        {/* Chargement */}
        <div style={{ background: T.bgCard, borderRadius: 12, border: `1px solid #2563EB55`, padding: "12px 14px", marginBottom: 12 }}>
          <SectionHeader label="🚢 Chargement" open={openChargement} toggle={() => setOpenChargement(o => !o)} color="#60A5FA" />
          {openChargement && (() => {
            const dests = destinations.filter(d => !d.excludeFromReport);
            const setMaxi = (name: string, field: "qty" | "weight", val: string) =>
              setChargementMaxi(p => ({ ...p, [name]: { qty: p[name]?.qty ?? "", weight: p[name]?.weight ?? "", [field]: val } }));
            const thS = (align: "left" | "right" | "center" = "right"): React.CSSProperties => ({
              textAlign: align, color: T.textDim, padding: "3px 5px 5px", fontWeight: 600, fontSize: 10, whiteSpace: "nowrap"
            });
            const tdS = (color: string = T.text): React.CSSProperties => ({
              textAlign: "right", padding: "4px 5px", fontFamily: "monospace", fontSize: 10, color
            });
            const inputS: React.CSSProperties = {
              width: 62, textAlign: "right", background: T.bgDark,
              border: `1px solid ${T.border2}`, borderRadius: 4,
              color: "#FB923C", fontFamily: "monospace", fontSize: 10, padding: "2px 4px"
            };
            let sumTotQty = 0, sumTotW = 0, sumMaxiQty = 0, sumMaxiW = 0;
            for (const d of dests) {
              const s = destStatsMap.get(d.name);
              const pQty = parseFloat(tallyPrev[d.name]?.qty ?? "0") || 0;
              const pW   = parseFloat(String(tallyPrev[d.name]?.weight ?? "0").replace(",", ".")) || 0;
              sumTotQty  += (s?.count ?? 0) + pQty;
              sumTotW    += (s?.weight ?? 0) + pW;
              sumMaxiQty += parseFloat(chargementMaxi[d.name]?.qty ?? "0") || 0;
              sumMaxiW   += parseFloat(String(chargementMaxi[d.name]?.weight ?? "0").replace(",", ".")) || 0;
            }
            const sumEstW   = sumMaxiW - sumTotW;
            const sumMoyW   = sumTotQty > 0 ? sumTotW / sumTotQty : NaN;
            const sumEstQty = !isNaN(sumMoyW) && sumMoyW > 0 ? Math.floor(sumEstW / sumMoyW) : NaN;
            return (
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                  <thead>
                    <tr style={{ borderBottom: `1px solid ${T.border}44` }}>
                      <th style={{ ...thS("left"), padding: "3px 6px 0 0" }} rowSpan={2}>Destination</th>
                      <th colSpan={2} style={{ ...thS(), color: "#60A5FA", borderBottom: `1px solid #60A5FA44`, paddingBottom: 2 }}>TOTAL</th>
                      <th colSpan={2} style={{ ...thS(), color: "#FB923C", opacity: 0.6, borderBottom: `1px solid #FB923C44`, paddingBottom: 2 }}>MAXI</th>
                      <th colSpan={2} style={{ ...thS(), color: "#34D399", borderBottom: `1px solid #34D39944`, paddingBottom: 2 }}>ESTIM.</th>
                      <th style={{ ...thS(), color: "#A78BFA", fontSize: 9 }} rowSpan={2}>Moy.</th>
                    </tr>
                    <tr style={{ borderBottom: `1px solid ${T.border}` }}>
                      <th style={thS()}>Qté</th><th style={thS()}>Poids</th>
                      <th style={{ ...thS(), color: "#FB923C", opacity: 0.6 }}>Qté</th><th style={{ ...thS(), color: "#FB923C", opacity: 0.6 }}>Poids</th>
                      <th style={{ ...thS(), color: "#34D399" }}>Qté</th><th style={{ ...thS(), color: "#34D399" }}>Poids</th>
                    </tr>
                  </thead>
                  <tbody>
                    {dests.map(d => {
                      const s = destStatsMap.get(d.name);
                      const pQty = parseFloat(tallyPrev[d.name]?.qty ?? "0") || 0;
                      const pW   = parseFloat(String(tallyPrev[d.name]?.weight ?? "0").replace(",", ".")) || 0;
                      const tQty = (s?.count ?? 0) + pQty;
                      const tW   = (s?.weight ?? 0) + pW;
                      const mQty = parseFloat(chargementMaxi[d.name]?.qty ?? "0") || 0;
                      const mW   = parseFloat(String(chargementMaxi[d.name]?.weight ?? "0").replace(",", ".")) || 0;
                      const eW   = mW - tW;
                      const moyW = tQty > 0 ? tW / tQty : NaN;
                      const effectiveMoyW = !isNaN(moyW) && moyW > 0 ? moyW : sumMoyW;
                      const eQty = !isNaN(effectiveMoyW) && effectiveMoyW > 0 ? Math.floor(eW / effectiveMoyW) : NaN;
                      return (
                        <tr key={d.name} style={{ borderBottom: `1px solid ${T.border2}22` }}>
                          <td style={{ padding: "4px 6px 4px 0" }}>
                            <span style={{ background: `${d.color}22`, borderLeft: `3px solid ${d.color}`, padding: "1px 6px", borderRadius: 3, color: d.color, fontWeight: 700, fontSize: 10 }}>{d.name}</span>
                          </td>
                          <td style={tdS("#60A5FA")}>{tQty || "—"}</td>
                          <td style={tdS("#60A5FA")}>{tW > 0 ? fmtWeight(tW) : "—"}</td>
                          <td style={{ padding: "3px 5px", textAlign: "right", color: T.textDim, opacity: 0.35, fontFamily: "monospace", fontSize: 11 }}>{mQty > 0 ? thsep(String(mQty)) : "—"}</td>
                          <td style={{ padding: "3px 5px", textAlign: "right" }}>
                            <input type="text" inputMode="decimal" style={{ ...inputS, width: 72 }}
                              value={(() => { const raw = chargementMaxi[d.name]?.weight ?? ""; const n = parseFloat(raw.replace(",", ".")); return isNaN(n) ? raw : thsep(String(n)); })()}
                              onChange={e => setMaxi(d.name, "weight", e.target.value.replace(/\s/g, ""))}
                              placeholder="0" />
                          </td>
                          <td style={tdS(eQty < 0 ? T.error : "#34D399")}>{mW > 0 ? (eQty < 0 ? `(${Math.abs(eQty)})` : isNaN(eQty) ? "—" : eQty || "—") : "—"}</td>
                          <td style={tdS(eW < 0 ? T.error : "#34D399")}>{mW > 0 ? (eW < 0 ? `(${fmtWeight(Math.abs(eW))})` : eW > 0 ? fmtWeight(eW) : "—") : "—"}</td>
                          <td style={tdS("#A78BFA")}>{!isNaN(moyW) && moyW > 0 ? fmtWeight(moyW) : "—"}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                  <tfoot>
                    <tr style={{ borderTop: `1px solid ${T.border}` }}>
                      <td style={{ padding: "5px 6px 4px 0", color: T.textDim, fontSize: 10, fontStyle: "italic" }}>Total</td>
                      <td style={tdS("#60A5FA")}><strong>{sumTotQty || "—"}</strong></td>
                      <td style={tdS("#60A5FA")}><strong>{sumTotW > 0 ? fmtWeight(sumTotW) : "—"}</strong></td>
                      <td style={tdS("#FB923C44")}><strong style={{ opacity: 0.35 }}>{sumMaxiQty > 0 ? thsep(String(sumMaxiQty)) : "—"}</strong></td>
                      <td style={tdS("#FB923C")}><strong>{sumMaxiW > 0 ? fmtWeight(sumMaxiW) : "—"}</strong></td>
                      <td style={tdS(sumEstQty < 0 ? T.error : "#34D399")}><strong>{sumMaxiW > 0 ? (isNaN(sumEstQty) ? "—" : sumEstQty < 0 ? `(${Math.abs(sumEstQty)})` : sumEstQty || "—") : "—"}</strong></td>
                      <td style={tdS(sumEstW < 0 ? T.error : "#34D399")}><strong>{sumMaxiW > 0 ? (sumEstW < 0 ? `(${fmtWeight(Math.abs(sumEstW))})` : sumEstW > 0 ? fmtWeight(sumEstW) : "—") : "—"}</strong></td>
                      <td style={tdS("#A78BFA")}><strong>{sumTotQty > 0 && sumTotW > 0 ? fmtWeight(sumTotW / sumTotQty) : "—"}</strong></td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            );
          })()}
        </div>

        {/* Déchargement */}
        <div style={{ background: T.bgCard, borderRadius: 12, border: `1px solid #D9770655`, padding: "12px 14px", marginBottom: 12 }}>
          <SectionHeader label="⚓ Déchargement" open={openDechargement} toggle={() => setOpenDechargement(o => !o)} color="#FB923C" />
          {openDechargement && (() => {
            const dests = destinations.filter(d => !d.excludeFromReport);
            const setMaxi = (name: string, field: "qty" | "weight", val: string) =>
              setDechargementMaxi(p => ({ ...p, [name]: { qty: p[name]?.qty ?? "", weight: p[name]?.weight ?? "", [field]: val } }));
            const thS = (align: "left" | "right" | "center" = "right"): React.CSSProperties => ({
              textAlign: align, color: T.textDim, padding: "3px 5px 5px", fontWeight: 600, fontSize: 10, whiteSpace: "nowrap"
            });
            const tdS = (color: string = T.text): React.CSSProperties => ({
              textAlign: "right", padding: "4px 5px", fontFamily: "monospace", fontSize: 10, color
            });
            const inputS: React.CSSProperties = {
              width: 62, textAlign: "right", background: T.bgDark,
              border: `1px solid ${T.border2}`, borderRadius: 4,
              color: "#FB923C", fontFamily: "monospace", fontSize: 10, padding: "2px 4px"
            };
            let sumTotQty = 0, sumTotW = 0, sumMaxiQty = 0, sumMaxiW = 0;
            for (const d of dests) {
              const s = destStatsMap.get(d.name);
              const pQty = parseFloat(tallyPrev[d.name]?.qty ?? "0") || 0;
              const pW   = parseFloat(String(tallyPrev[d.name]?.weight ?? "0").replace(",", ".")) || 0;
              sumTotQty  += (s?.count ?? 0) + pQty;
              sumTotW    += (s?.weight ?? 0) + pW;
              sumMaxiQty += parseFloat(dechargementMaxi[d.name]?.qty ?? "0") || 0;
              sumMaxiW   += parseFloat(String(dechargementMaxi[d.name]?.weight ?? "0").replace(",", ".")) || 0;
            }
            const sumEstW   = sumTotW - sumMaxiW;
            const sumEstQty = sumTotQty - sumMaxiQty;
            return (
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                  <thead>
                    <tr style={{ borderBottom: `1px solid ${T.border}44` }}>
                      <th style={{ ...thS("left"), padding: "3px 6px 0 0" }} rowSpan={2}>Destination</th>
                      <th colSpan={2} style={{ ...thS(), color: "#60A5FA", borderBottom: `1px solid #60A5FA44`, paddingBottom: 2 }}>DÉCHARGÉ</th>
                      <th colSpan={2} style={{ ...thS(), color: "#FB923C", borderBottom: `1px solid #FB923C44`, paddingBottom: 2 }}>TOTAL</th>
                      <th colSpan={2} style={{ ...thS(), color: "#34D399", borderBottom: `1px solid #34D39944`, paddingBottom: 2 }}>RESTANT</th>
                      <th style={{ ...thS(), color: "#A78BFA", fontSize: 9 }} rowSpan={2}>Moy.</th>
                    </tr>
                    <tr style={{ borderBottom: `1px solid ${T.border}` }}>
                      <th style={thS()}>Qté</th><th style={thS()}>Poids</th>
                      <th style={{ ...thS(), color: "#FB923C" }}>Qté</th><th style={{ ...thS(), color: "#FB923C" }}>Poids</th>
                      <th style={{ ...thS(), color: "#34D399" }}>Qté</th><th style={{ ...thS(), color: "#34D399" }}>Poids</th>
                    </tr>
                  </thead>
                  <tbody>
                    {dests.map(d => {
                      const s = destStatsMap.get(d.name);
                      const pQty = parseFloat(tallyPrev[d.name]?.qty ?? "0") || 0;
                      const pW   = parseFloat(String(tallyPrev[d.name]?.weight ?? "0").replace(",", ".")) || 0;
                      const tQty = (s?.count ?? 0) + pQty;
                      const tW   = (s?.weight ?? 0) + pW;
                      const mQty = parseFloat(dechargementMaxi[d.name]?.qty ?? "0") || 0;
                      const mW   = parseFloat(String(dechargementMaxi[d.name]?.weight ?? "0").replace(",", ".")) || 0;
                      const eW   = tW - mW;
                      const eQty = tQty - mQty;
                      const moyW = tQty > 0 ? tW / tQty : NaN;
                      return (
                        <tr key={d.name} style={{ borderBottom: `1px solid ${T.border2}22` }}>
                          <td style={{ padding: "4px 6px 4px 0" }}>
                            <span style={{ background: `${d.color}22`, borderLeft: `3px solid ${d.color}`, padding: "1px 6px", borderRadius: 3, color: d.color, fontWeight: 700, fontSize: 10 }}>{d.name}</span>
                          </td>
                          <td style={tdS("#60A5FA")}>{tQty || "—"}</td>
                          <td style={tdS("#60A5FA")}>{tW > 0 ? fmtWeight(tW) : "—"}</td>
                          <td style={{ padding: "3px 5px", textAlign: "right" }}>
                            <input type="number" min={0} style={inputS}
                              value={dechargementMaxi[d.name]?.qty ?? ""}
                              onChange={e => setMaxi(d.name, "qty", e.target.value)}
                              placeholder="0" />
                          </td>
                          <td style={{ padding: "3px 5px", textAlign: "right" }}>
                            <input type="text" inputMode="decimal" style={{ ...inputS, width: 72 }}
                              value={(() => { const raw = dechargementMaxi[d.name]?.weight ?? ""; const n = parseFloat(raw.replace(",", ".")); return isNaN(n) ? raw : thsep(String(n)); })()}
                              onChange={e => setMaxi(d.name, "weight", e.target.value.replace(/\s/g, ""))}
                              placeholder="0" />
                          </td>
                          <td style={tdS(eQty < 0 ? "#34D399" : T.error)}>{mQty > 0 || tQty > 0 ? (eQty < 0 ? `(${Math.abs(eQty)})` : eQty || "—") : "—"}</td>
                          <td style={tdS(eW < 0 ? "#34D399" : T.error)}>{mW > 0 || tW > 0 ? (eW < 0 ? `(${fmtWeight(Math.abs(eW))})` : eW > 0 ? fmtWeight(eW) : "—") : "—"}</td>
                          <td style={tdS("#A78BFA")}>{!isNaN(moyW) && moyW > 0 ? fmtWeight(moyW) : "—"}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                  <tfoot>
                    <tr style={{ borderTop: `1px solid ${T.border}` }}>
                      <td style={{ padding: "5px 6px 4px 0", color: T.textDim, fontSize: 10, fontStyle: "italic" }}>Total</td>
                      <td style={tdS("#60A5FA")}><strong>{sumTotQty || "—"}</strong></td>
                      <td style={tdS("#60A5FA")}><strong>{sumTotW > 0 ? fmtWeight(sumTotW) : "—"}</strong></td>
                      <td style={tdS("#FB923C")}><strong>{sumMaxiQty || "—"}</strong></td>
                      <td style={tdS("#FB923C")}><strong>{sumMaxiW > 0 ? fmtWeight(sumMaxiW) : "—"}</strong></td>
                      <td style={tdS(sumEstQty < 0 ? "#34D399" : T.error)}><strong>{sumEstQty < 0 ? `(${Math.abs(sumEstQty)})` : sumEstQty || "—"}</strong></td>
                      <td style={tdS(sumEstW < 0 ? "#34D399" : T.error)}><strong>{sumEstW < 0 ? `(${fmtWeight(Math.abs(sumEstW))})` : sumEstW > 0 ? fmtWeight(sumEstW) : "—"}</strong></td>
                      <td style={tdS("#A78BFA")}><strong>{sumTotQty > 0 && sumTotW > 0 ? fmtWeight(sumTotW / sumTotQty) : "—"}</strong></td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            );
          })()}
        </div>

        {/* Résumé général */}
        <div style={{ background: T.bgCard, borderRadius: 12, border: `1px solid ${T.border2}`, padding: "12px 14px", marginBottom: 12 }}>
          <SectionHeader label="📊 Résumé" open={openResume} toggle={() => setOpenResume(o => !o)} color={T.textMuted} />
          {openResume && <div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "8px 12px" }}>
            {[
              { label: "Lignes totales",   val: String(totalRows),       color: T.text },
              { label: "Pointées",          val: String(totalPointed),    color: T.success },
              { label: "Affectées (dest.)", val: String(totalAffected),   color: "#A78BFA" },
              ...(reassignedRows.size > 0 ? [
                { label: "Réaffectations",  val: String(reassignedRows.size), color: T.error },
              ] : []),
              ...(poidsIdx >= 0 ? [
                { label: `Poids total`,     val: fmtWeight(totalWeight),  color: T.warning },
                { label: `Poids affecté`,   val: fmtWeight(affectedWeight), color: "#A78BFA" },
                { label: `Poids moyen`,     val: totalRows > 0 && totalWeight > 0 ? (poidsUnit === "kg" ? fmtWeight(totalWeight / totalRows) : thsep(Math.round(totalWeight / totalRows).toString()) + " t") : "—", color: "#F9A8D4" },
              ] : []),
            ].map(({ label, val, color }) => (
              <div key={label} style={{ background: T.bgDark, borderRadius: 8, padding: "8px 10px" }}>
                <div style={{ color: T.textDim, fontSize: 10, marginBottom: 2 }}>{label}</div>
                <div style={{ color, fontWeight: 700, fontSize: 14, fontFamily: "'Share Tech Mono', monospace" }}>{val}</div>
              </div>
            ))}
          </div>
          {(() => {
            const hneRows = [...rowDestinations.entries()].filter(([, d]) => excludedFromReport.has(d));
            if (hneRows.length === 0) return null;
            const hneColor = "#94A3B8";
            return (
              <div style={{ marginTop: 12 }}>
                <div style={{ color: hneColor, fontWeight: 700, fontSize: 11, marginBottom: 6 }}>
                  🚫 Hors rapport — {hneRows.length} ligne{hneRows.length > 1 ? "s" : ""}
                </div>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                  <thead>
                    <tr style={{ borderBottom: `1px solid ${T.border}` }}>
                      <th style={{ textAlign: "left", color: T.textDim, padding: "3px 6px 6px 0", fontWeight: 600 }}>#</th>
                      {rangIdx >= 0 && <th style={{ textAlign: "left", color: T.textDim, padding: "3px 6px 6px", fontWeight: 600 }}>📍 Rang</th>}
                      {refIdx >= 0 && <th style={{ textAlign: "left", color: T.textDim, padding: "3px 6px 6px", fontWeight: 600 }}>🏷 REF</th>}
                      {poidsIdx >= 0 && <th style={{ textAlign: "right", color: T.textDim, padding: "3px 0 6px 6px", fontWeight: 600 }}>Poids</th>}
                      <th style={{ textAlign: "right", color: T.textDim, padding: "3px 0 6px 6px", fontWeight: 600 }}>Dest.</th>
                    </tr>
                  </thead>
                  <tbody>
                    {hneRows.map(([ri, dest]) => {
                      const row = allRows[ri];
                      const refRaw = row && refIdx >= 0 ? String(row[refIdx] ?? "").slice(0, 24) : "";
                      const refDisplay = refIdx >= 0 ? (autoRefFmt ? autoFormatRef(refRaw, splitFormats["reference"] || "") : applyGrouping(refRaw, splitFormats["reference"] || "")) : "";
                      const destColor = destinations.find(d => d.name === dest)?.color ?? hneColor;
                      return (
                        <tr key={ri} style={{ borderBottom: `1px solid ${T.border2}22` }}>
                          <td style={{ padding: "3px 6px 3px 0", color: T.textDim, fontFamily: "monospace" }}>{ri + 1}</td>
                          {rangIdx >= 0 && <td style={{ padding: "3px 6px 3px 0", color: T.textMuted, fontFamily: "monospace" }}>{applyGrouping(String(row?.[rangIdx] ?? ""), splitFormats["rang"] || "") || "—"}</td>}
                          {refIdx >= 0 && <td style={{ padding: "3px 6px", color: T.text, fontFamily: "monospace" }}>{refDisplay || "—"}</td>}
                          {poidsIdx >= 0 && <td style={{ textAlign: "right", padding: "3px 0 3px 6px", color: T.warning, fontFamily: "monospace" }}>{fmtWeight(parseFloat(String(row?.[poidsIdx] ?? "").replace(",", ".")) || 0)}</td>}
                          <td style={{ textAlign: "right", padding: "3px 0 3px 6px" }}>
                            <span style={{ color: destColor, fontWeight: 700, fontSize: 10 }}>{dest}</span>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            );
          })()}
          </div>}
        </div>

        {/* Lignes pointées */}
        {pointedRows.size > 0 && (
          <div style={{ background: T.bgCard, borderRadius: 12, border: `1px solid ${T.success}44`, padding: "12px 14px", marginBottom: 12 }}>
            <SectionHeader label={`✅ Lignes pointées (${pointedRows.size})`} open={openPointees} toggle={() => setOpenPointees(o => !o)} color={T.success} />
            {openPointees && <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
              <thead>
                <tr style={{ borderBottom: `1px solid ${T.border}` }}>
                  {rangIdx >= 0 && <th style={{ textAlign: "left", color: T.textDim, padding: "3px 6px 6px 0", fontWeight: 600 }}>📍 Rang</th>}
                  {refIdx  >= 0 && <th style={{ textAlign: "left", color: T.textDim, padding: "3px 6px 6px", fontWeight: 600 }}>🏷 Référence</th>}
                  {poidsIdx >= 0 && <th style={{ textAlign: "right", color: T.textDim, padding: "3px 0 6px 6px", fontWeight: 600 }}>Poids</th>}
                  {extraCols.map(e => <th key={e.col} style={{ textAlign: "left", color: T.textDim, padding: "3px 6px 6px", fontWeight: 600 }}>{e.label}</th>)}
                  <th style={{ textAlign: "right", color: T.textDim, padding: "3px 0 6px 6px", fontWeight: 600 }}>Dest.</th>
                </tr>
              </thead>
              <tbody>
                {visibleRows.filter(({ ri }) => pointedRows.has(ri)).map(({ row, ri }) => {
                  const dest = rowDestinations.get(ri);
                  const destColor = dest ? (destinations.find(d => d.name === dest)?.color ?? T.accent) : null;
                  const refRaw = refIdx >= 0 ? String(row[refIdx] ?? "").slice(0, 24) : "";
                  const refDisplay = refIdx >= 0 ? (autoRefFmt ? autoFormatRef(refRaw, splitFormats["reference"] || "") : applyGrouping(refRaw, splitFormats["reference"] || "")) : "";
                  return (
                    <tr key={ri} style={{ borderBottom: `1px solid ${T.border2}22` }}>
                      {rangIdx >= 0 && <td style={{ padding: "3px 6px 3px 0", color: T.textMuted, fontFamily: "monospace" }}>{applyGrouping(String(row[rangIdx] ?? ""), splitFormats["rang"] || "") || "—"}</td>}
                      {refIdx  >= 0 && <td style={{ padding: "3px 6px", color: T.text, fontFamily: "monospace" }}>{refDisplay || "—"}</td>}
                      {poidsIdx >= 0 && <td style={{ textAlign: "right", padding: "3px 0 3px 6px", color: T.warning, fontFamily: "monospace" }}>{fmtWeight(parseFloat(String(row[poidsIdx] ?? "")) || 0)}</td>}
                      {extraCols.map(e => {
                        const ci = headers.indexOf(e.col);
                        return <td key={e.col} style={{ padding: "3px 6px", color: T.success, fontFamily: "monospace" }}>{ci >= 0 ? String(row[ci] ?? "") || "—" : "—"}</td>;
                      })}
                      <td style={{ textAlign: "right", padding: "3px 0 3px 6px" }}>
                        {dest ? <span style={{ color: destColor ?? T.accent, fontWeight: 700, fontSize: 10 }}>{dest}</span> : <span style={{ color: T.textDim, fontSize: 10 }}>—</span>}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>}
          </div>
        )}

        {/* Réaffectations — doublons / erreurs potentiels */}
        {reassignedRows.size > 0 && (
          <div style={{ background: T.bgCard, borderRadius: 12, border: `1px solid ${T.error}55`, padding: "12px 14px", marginBottom: 12 }}>
            <SectionHeader label={`⚠️ Réaffectations (${reassignedRows.size} ligne${reassignedRows.size > 1 ? "s" : ""})`} open={openReaff} toggle={() => setOpenReaff(o => !o)} color={T.error} />
            {openReaff && <><div style={{ color: T.textDim, fontSize: 10, marginBottom: 8 }}>Ces lignes ont changé de destination — vérifier doublons ou erreurs de saisie.</div>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
              <thead>
                <tr style={{ borderBottom: `1px solid ${T.border}` }}>
                  <th style={{ textAlign: "left", color: T.textDim, padding: "3px 6px 6px 0", fontWeight: 600 }}>#</th>
                  {refIdx  >= 0 && <th style={{ textAlign: "left", color: T.textDim, padding: "3px 6px 6px", fontWeight: 600 }}>Référence</th>}
                  {poidsIdx >= 0 && <th style={{ textAlign: "right", color: T.textDim, padding: "3px 6px 6px", fontWeight: 600 }}>Poids</th>}
                  <th style={{ textAlign: "left", color: T.textDim, padding: "3px 6px 6px", fontWeight: 600 }}>Ancien</th>
                  <th style={{ textAlign: "left", color: T.textDim, padding: "3px 0 6px 6px", fontWeight: 600 }}>Actuel</th>
                </tr>
              </thead>
              <tbody>
                {[...reassignedRows.entries()].map(([ri, history]) => {
                  const row = allRows[ri];
                  const refRaw = row && refIdx >= 0 ? String(row[refIdx] ?? "").slice(0, 20) : "";
                  const refDisplay = autoRefFmt ? autoFormatRef(refRaw, splitFormats["reference"] || "") : applyGrouping(refRaw, splitFormats["reference"] || "");
                  const poidsRaw = row && poidsIdx >= 0 ? parseFloat(String(row[poidsIdx] ?? "").replace(",", ".")) : NaN;
                  const poidsDisplay = !isNaN(poidsRaw) ? fmtWeight(poidsRaw) : "—";
                  const current = rowDestinations.get(ri) ?? "—";
                  const currentColor = destinations.find(d => d.name === current)?.color ?? T.textDim;
                  return (
                    <tr key={ri} style={{ borderBottom: `1px solid ${T.border2}22` }}>
                      <td style={{ padding: "4px 6px 4px 0", color: T.textDim, fontFamily: "monospace" }}>{ri + 1}</td>
                      {refIdx >= 0 && <td style={{ padding: "4px 6px", color: T.text, fontFamily: "monospace", fontSize: 10 }}>{refDisplay || "—"}</td>}
                      {poidsIdx >= 0 && <td style={{ padding: "4px 6px", color: T.text, fontFamily: "monospace", fontSize: 10, textAlign: "right" }}>{poidsDisplay}</td>}
                      <td style={{ padding: "4px 6px" }}>
                        <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                          {history.map((h, i) => {
                            const fromColor = destinations.find(d => d.name === h.from)?.color ?? T.textDim;
                            const toColor   = destinations.find(d => d.name === h.to)?.color   ?? T.textDim;
                            return (
                              <span key={i} style={{ fontSize: 10, color: T.textDim }}>
                                <span style={{ color: fromColor, fontWeight: 600 }}>{h.from}</span>
                                <span style={{ opacity: 0.5 }}> → </span>
                                <span style={{ color: toColor, fontWeight: 600 }}>{h.to}</span>
                              </span>
                            );
                          })}
                        </div>
                      </td>
                      <td style={{ padding: "4px 0 4px 6px" }}>
                        <span style={{ color: currentColor, fontWeight: 700, background: `${currentColor}22`, padding: "1px 6px", borderRadius: 4 }}>{current}</span>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table></> }
          </div>
        )}

        {destStatsMap.size === 0 && pointedRows.size === 0 && reassignedRows.size === 0 && (
          <EmptyState icon="📋" text="Aucune donnée de rapport" sub="Pointez des lignes ou affectez des destinations dans l'onglet Pointage" />
        )}

      </div>
    </div>
  );
}

function ExportPage() {
  const {
    parsed, workbook, sheetNames, hiddenSheets,
    exportFileName, setExportFileName,
    showToast, setActiveTab,
    setParsed, setFileName, setWorkbook, setSheetNames, setActiveSheet,
    setAddedRows, setHiddenSheets, setSheetSelectMode, setSelectedSheets,
    sheetStates, setPointedRows, setRowOverrides,
  } = useApp();

  const flushActive = useFlushActiveSheet();

  const exportClean = () => {
    if (!workbook || !parsed) return;
    // S'assurer que l'onglet actif est bien sauvegardé avant export
    flushActive();

    const wb2 = XLSX.utils.book_new();
    const visibleSheets = sheetNames.filter((n) => !hiddenSheets.has(n));

    visibleSheets.forEach((sheetName) => {
      // Chercher l'état nettoyé dans sheetStates, sinon utiliser le workbook brut
      const saved = sheetStates.current.get(sheetName);
      if (saved) {
        const { headers: sh, addedRows: sa, hiddenCols: sc, hiddenRows: sr, parsed: sp } = saved;
        const visibleHeaders = sh.filter((_, i) => !sc.has(i));
        const allSheetRows = [...sa, ...sp.rows];
        const visibleRows = allSheetRows
          .filter((_, i) => !sr.has(i))
          .map((row) => {
            const padded = [...row];
            while (padded.length < sh.length) padded.push(null);
            return padded.filter((_, ci) => !sc.has(ci));
          });
        const ws = XLSX.utils.aoa_to_sheet([visibleHeaders, ...visibleRows]);
        XLSX.utils.book_append_sheet(wb2, ws, sheetName);
      } else {
        // Onglet jamais visité → export brut
        const ws = workbook.Sheets[sheetName];
        XLSX.utils.book_append_sheet(wb2, ws, sheetName);
      }
    });

    if (wb2.SheetNames.length === 0) { showToast("Rien à exporter", "error"); return; }
    XLSX.writeFile(wb2, `${exportFileName.trim() || "données_nettoyées"}.xlsx`);
    showToast("Fichier exporté !", "success");
  };

  const handleReset = () => {
    sheetStates.current.clear();
    setParsed(null);
    setFileName(null);
    setWorkbook(null);
    setSheetNames([]);
    setActiveSheet(null);
    setAddedRows([]);
    setHiddenSheets(new Set());
    setSheetSelectMode("none");
    setSelectedSheets(new Set());
    setPointedRows(new Set());
    setRowOverrides(new Map());
    setActiveTab("import");
    showToast("Réinitialisé", "info");
  };

  if (!parsed) {
    return (
      <div style={{ flex: 1, overflowY: "auto", paddingBottom: 80, background: T.bg }}>
        <PageHeader title="Export" />
        <EmptyState icon="📭" text="Aucune donnée à exporter" sub="Importez et nettoyez un fichier d'abord" />
        <div style={{ padding: 16 }}>
          <Btn onClick={() => setActiveTab("import")} color={T.accent} textColor="#0F172A" fullWidth>
            ⬇️ Aller à l'import
          </Btn>
        </div>
      </div>
    );
  }

  // Flush active sheet before computing summary stats (ref mutation, safe in render)
  flushActive();

  const visibleSheets = sheetNames.filter((n) => !hiddenSheets.has(n));

  // Totaux multi-onglets (tous onglets sauvegardés + bruts)
  let totalColsAll = 0, visibleColsAll = 0;
  let totalRowsAll = 0, visibleRowsAll = 0;
  let totalAddedAll = 0;
  visibleSheets.forEach((sn) => {
    const s = sheetStates.current.get(sn);
    if (s) {
      const allR = [...s.addedRows, ...s.parsed.rows];
      totalColsAll   += s.headers.length;
      visibleColsAll += s.headers.filter((_, i) => !s.hiddenCols.has(i)).length;
      totalRowsAll   += allR.length;
      visibleRowsAll += allR.filter((_, i) => !s.hiddenRows.has(i)).length;
      totalAddedAll  += s.addedRows.length;
    } else if (workbook?.Sheets[sn]) {
      const raw = XLSX.utils.sheet_to_json(workbook.Sheets[sn], { header: 1, defval: null }) as RawData;
      totalColsAll   += raw[0]?.length ?? 0;
      visibleColsAll += raw[0]?.length ?? 0;
      totalRowsAll   += Math.max(0, raw.length - 1);
      visibleRowsAll += Math.max(0, raw.length - 1);
    }
  });

  return (
    <div style={{ flex: 1, overflowY: "auto", paddingBottom: 80, background: T.bg }}>
      <PageHeader title="Export" subtitle="Téléchargez le fichier nettoyé" />
      <StatsBar />

      <div style={{ padding: 16 }}>
        {/* Summary card */}
        <div style={{
          background: T.bgCard, border: `1px solid ${T.border}`,
          borderRadius: 12, overflow: "hidden", marginBottom: 16,
        }}>
          <SectionTitle icon="📋" text="Résumé" />
          <div style={{ padding: "12px 16px" }}>
            {([
              { label: "Colonnes visibles", value: `${visibleColsAll} / ${totalColsAll}`,             color: T.accent },
              { label: "Lignes visibles",   value: `${visibleRowsAll} / ${totalRowsAll}`,             color: T.success },
              { label: "Onglets exportés",  value: `${visibleSheets.length} / ${sheetNames.length}`, color: T.warning },
              { label: "Lignes ajoutées",   value: `+ ${totalAddedAll}`,                             color: T.success },
            ] as { label: string; value: string; color: string }[]).map((r) => (
              <div key={r.label} style={{
                display: "flex", justifyContent: "space-between",
                padding: "8px 0", borderBottom: `1px solid ${T.border}22`,
              }}>
                <span style={{ color: T.textMuted, fontSize: 13 }}>{r.label}</span>
                <span style={{ color: r.color, fontWeight: 700, fontSize: 13, fontFamily: "monospace" }}>
                  {r.value}
                </span>
              </div>
            ))}
          </div>
        </div>

        {/* Modifications */}
        <div style={{ background: T.bgCard, border: `1px solid ${T.border}`, borderRadius: 12, marginBottom: 16, overflow: "hidden" }}>
          <SectionTitle icon="✏️" text="Modifications actives" />
          <div style={{ padding: "10px 16px", display: "flex", flexDirection: "column", gap: 6 }}>
            {/* Détail par onglet */}
            {visibleSheets.map((sn) => {
              const s = sheetStates.current.get(sn);
              const modifs: string[] = [];
              if (s) {
                if (s.hiddenCols.size > 0) modifs.push(`${s.hiddenCols.size} col. masquée(s)`);
                if (s.hiddenRows.size > 0) modifs.push(`${s.hiddenRows.size} ligne(s) masquée(s)`);
                if (s.addedRows.length > 0) modifs.push(`+${s.addedRows.length} ajoutée(s)`);
              }
              return (
                <div key={sn} style={{ display: "flex", alignItems: "flex-start", gap: 8, padding: "4px 0", borderBottom: `1px solid ${T.border}22` }}>
                  <span style={{ color: T.accent, fontSize: 11, minWidth: 12 }}>🗂</span>
                  <div>
                    <span style={{ color: T.text, fontSize: 12, fontWeight: 700 }}>{sn}</span>
                    {modifs.length > 0
                      ? <span style={{ color: T.warning, fontSize: 11, marginLeft: 8 }}>{modifs.join(" · ")}</span>
                      : <span style={{ color: T.textDim, fontSize: 11, marginLeft: 8 }}>{s ? "nettoyé" : "brut (non visité)"}</span>
                    }
                  </div>
                </div>
              );
            })}
            {hiddenSheets.size > 0 && (
              <div style={{ display: "flex", gap: 8, color: T.error, fontSize: 12, paddingTop: 4 }}>
                <span>🚫</span>
                {hiddenSheets.size} onglet(s) exclu(s)
              </div>
            )}
          </div>
        </div>

        {/* Export card */}
        <div style={{ background: T.bgCard, border: `1px solid ${T.border}`, borderRadius: 12, marginBottom: 16, overflow: "hidden" }}>
          <SectionTitle icon="⬆️" text="Télécharger" />
          <div style={{ padding: 16 }}>
            <div style={{ color: T.textMuted, fontSize: 11, marginBottom: 6, fontWeight: 600 }}>
              Nom du fichier
            </div>
            <div style={{ display: "flex", alignItems: "center", marginBottom: 16, border: `1px solid ${T.accentDim}`, borderRadius: 8, overflow: "hidden" }}>
              <input
                value={exportFileName}
                onChange={(e) => setExportFileName(e.target.value)}
                onKeyDown={(e) => { if (e.key === "Enter") exportClean(); }}
                spellCheck={false}
                style={{
                  flex: 1, background: T.bgDark, border: "none", outline: "none",
                  color: T.success, fontSize: 13, padding: "10px 12px",
                  fontFamily: "'Share Tech Mono', monospace", fontWeight: 700,
                }}
              />
              <span style={{
                background: T.bgDark, color: T.textDim,
                fontSize: 12, padding: "10px 10px 10px 0",
                fontFamily: "'Share Tech Mono', monospace",
                pointerEvents: "none",
              }}>.xlsx</span>
            </div>
            <Btn onClick={exportClean} color={T.success} textColor="#0F172A" fullWidth>
              ⬇️ Télécharger le fichier nettoyé
            </Btn>
          </div>
        </div>

        {/* Danger zone */}
        <div style={{ background: "#1A0A0A", border: `1px solid ${T.error}33`, borderRadius: 12, overflow: "hidden" }}>
          <SectionTitle icon="⚠️" text="Zone dangereuse" />
          <div style={{ padding: 16 }}>
            <div style={{ color: T.textMuted, fontSize: 12, marginBottom: 12 }}>
              Efface toutes les données chargées et les modifications.
            </div>
            <Btn onClick={handleReset} danger fullWidth>
              ✕ Réinitialiser
            </Btn>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────
// 8. APP ROOT
// ─────────────────────────────────────────────

function AppInner() {
  const { activeTab } = useApp();

  const pages: Record<Tab, React.ReactNode> = {
    import:  <ImportPage />,
    tableau: <TablePage />,
    rapport: <RapportPage />,
    export:  <ExportPage />,
  };

  return (
    <div style={{
      display: "flex", flexDirection: "column",
      height: "100dvh", maxWidth: 540, margin: "0 auto",
      background: T.bg, color: T.text,
      fontFamily: "'Share Tech Mono', 'IBM Plex Mono', 'Courier New', monospace",
      position: "relative", overflow: "hidden",
    }}>
      <Toast />
      <div style={{ flex: 1, display: "flex", flexDirection: "column", overflowY: "auto" }}>
        {pages[activeTab]}
      </div>
      <BottomNav />
    </div>
  );
}

export default function App() {
  return (
    <AppProvider>
      <AppInner />
    </AppProvider>
  );
}
