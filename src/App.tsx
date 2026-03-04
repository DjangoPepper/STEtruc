// ============================================================
// STEtruc — Nettoyeur Excel Portatif
// Visuel  : coil-deploy  (navy dark, mobile-first, bottom-nav)
// Fonctions: STEpi        (Excel import, multi-sheets, clean, export)
// Stack    : React 18 + TypeScript + Vite → GitHub Pages
// ============================================================

import {
  useState, useCallback, useRef, useMemo,
  useContext, createContext,
} from "react";
import * as XLSX from "xlsx";

// ─────────────────────────────────────────────
// 1. TYPES
// ─────────────────────────────────────────────

type CellValue = string | number | boolean | null;
type RawData = CellValue[][];
type Tab = "import" | "tableau" | "export";
type SelectMode = "none" | "col" | "row";
type SheetSelectMode = "none" | "delete" | "keep";

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

  // Stockage de l'état de chaque onglet (persist entre changements d'onglet)
  const sheetStates = useRef<Map<string, SheetState>>(new Map());

  const showToast = useCallback((msg: string, type: "success" | "error" | "info" = "info") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3000);
  }, []);

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
      setActiveSheet(wb.SheetNames[0]);
      setHiddenSheets(new Set());
      setSheetSelectMode("none");
      setSelectedSheets(new Set());
      loadSheet(wb, wb.SheetNames[0]);
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
    { id: "tableau", icon: "📊",  label: "Tableau" },
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
        const disabled = (t.id === "tableau" || t.id === "export") && !parsed;
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
  const { parsed, addedRows, hiddenCols, hiddenRows, hiddenSheets, sheetNames, headers } = useApp();
  if (!parsed) return null;
  const visibleCols = headers.filter((_, i) => !hiddenCols.has(i)).length;
  const allRows = [...addedRows, ...parsed.rows];
  const visibleRows = allRows.filter((_, i) => !hiddenRows.has(i)).length;
  const visibleShts = sheetNames.filter((n) => !hiddenSheets.has(n)).length;

  return (
    <div style={{
      display: "grid", gridTemplateColumns: "repeat(3, 1fr)",
      background: T.bgDark, borderBottom: `1px solid ${T.border}`,
    }}>
      {[
        { label: "Colonnes", value: visibleCols, color: T.accent },
        { label: "Lignes",   value: visibleRows, color: T.success },
        { label: "Onglets",  value: visibleShts, color: T.warning },
      ].map((s) => (
        <div key={s.label} style={{
          padding: "8px 4px", textAlign: "center",
          borderRight: `1px solid ${T.border}`,
        }}>
          <div style={{ color: s.color, fontSize: 17, fontWeight: 900, fontFamily: "monospace" }}>
            {s.value}
          </div>
          <div style={{ color: T.textDim, fontSize: 9, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.5 }}>
            {s.label}
          </div>
        </div>
      ))}
    </div>
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

function SheetTabsSection() {
  const {
    sheetNames, hiddenSheets, setHiddenSheets, activeSheet, setActiveSheet,
    workbook, loadSheet,
    sheetSelectMode, setSheetSelectMode, selectedSheets, setSelectedSheets,
    showToast,
  } = useApp();

  if (sheetNames.length <= 1) return null;

  const applySheetAction = () => {
    if (sheetSelectMode === "delete") {
      const next = new Set(hiddenSheets);
      selectedSheets.forEach((n) => next.add(n));
      setHiddenSheets(next);
      if (selectedSheets.has(activeSheet ?? "")) {
        const first = sheetNames.find((n) => !next.has(n));
        if (first && workbook) { setActiveSheet(first); loadSheet(workbook, first); }
      }
      showToast(`${selectedSheets.size} onglet(s) masqué(s)`, "info");
    } else {
      const next = new Set<string>(sheetNames.filter((n) => !selectedSheets.has(n)));
      setHiddenSheets(next);
      if (next.has(activeSheet ?? "")) {
        const first = sheetNames.find((n) => !next.has(n));
        if (first && workbook) { setActiveSheet(first); loadSheet(workbook, first); }
      }
      showToast(`Onglets conservés : ${selectedSheets.size}`, "success");
    }
    setSelectedSheets(new Set());
    setSheetSelectMode("none");
  };

  return (
    <div style={{ margin: "14px 16px 0", background: T.bgCard, borderRadius: 12, border: `1px solid ${T.border}`, overflow: "hidden" }}>
      <div style={{ padding: "10px 14px", background: T.bgDark, display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", borderBottom: `1px solid ${T.border}` }}>
        <span style={{ color: T.accent, fontWeight: 700, fontSize: 12, flex: 1 }}>
          🗂 Onglets ({sheetNames.length - hiddenSheets.size}/{sheetNames.length} visibles)
        </span>
        {sheetSelectMode === "none" ? (
          <>
            <Btn small onClick={() => { setSheetSelectMode("delete"); setSelectedSheets(new Set()); }}
              color={T.border} textColor={T.textMuted}>✕ Masquer</Btn>
            <Btn small onClick={() => { setSheetSelectMode("keep"); setSelectedSheets(new Set()); }}
              color={T.border} textColor={T.textMuted}>✓ Garder</Btn>
            {hiddenSheets.size > 0 && (
              <Btn small onClick={() => setHiddenSheets(new Set())} color={T.success} textColor="#0F172A">
                Restaurer
              </Btn>
            )}
          </>
        ) : (
          <>
            <span style={{ color: sheetSelectMode === "delete" ? T.error : T.success, fontSize: 10 }}>
              {sheetSelectMode === "delete" ? "→ Cliquez onglets à masquer" : "→ Cliquez onglets à garder"}
            </span>
            {selectedSheets.size > 0 && (
              <Btn small onClick={applySheetAction}
                color={sheetSelectMode === "keep" ? T.success : T.error}
                textColor="#0F172A">
                {sheetSelectMode === "delete" ? `Masquer ${selectedSheets.size}` : `Garder ${selectedSheets.size}`}
              </Btn>
            )}
            <Btn small onClick={() => { setSheetSelectMode("none"); setSelectedSheets(new Set()); }}
              color={T.border} textColor={T.textMuted}>Annuler</Btn>
          </>
        )}
      </div>
      {/* Tabs */}
      <div style={{ display: "flex", overflowX: "auto", padding: "8px 10px", gap: 6, flexWrap: "wrap" }}>
        {sheetNames.map((name) => {
          const isHidden = hiddenSheets.has(name);
          const isActive = activeSheet === name;
          const isSel = selectedSheets.has(name);
          const inSelectMode = sheetSelectMode !== "none";
          const borderColor = isSel
            ? (sheetSelectMode === "delete" ? T.error : T.success)
            : isActive ? T.accent : T.border2;
          return (
            <button
              key={name}
              onClick={() => {
                if (inSelectMode) {
                  setSelectedSheets((prev) => {
                    const next = new Set(prev);
                    next.has(name) ? next.delete(name) : next.add(name);
                    return next;
                  });
                } else if (!isHidden && workbook) {
                  setActiveSheet(name);
                  loadSheet(workbook, name);
                }
              }}
              style={{
                background: isActive && !inSelectMode ? T.border : T.bgDark,
                border: `1px solid ${borderColor}`,
                borderRadius: 7, padding: "6px 12px", cursor: "pointer",
                color: isHidden ? T.textDim : isActive && !inSelectMode ? T.accent : T.textMuted,
                fontSize: 12, fontWeight: 700,
                fontFamily: "'Share Tech Mono', monospace",
                opacity: isHidden ? 0.4 : 1,
                textDecoration: isHidden ? "line-through" : "none",
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
  );
}

function ImportPage() {
  const { handleFile, parsed, fileName, sheetNames, showToast, setActiveTab } = useApp();
  const fileInputRef = useRef<HTMLInputElement>(null);

  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) {
      handleFile(file);
      showToast("Fichier chargé !", "success");
      setTimeout(() => setActiveTab("tableau"), 400);
    }
  }, [handleFile, showToast, setActiveTab]);

  const onFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      handleFile(file);
      showToast("Fichier chargé !", "success");
      setTimeout(() => setActiveTab("tableau"), 400);
    }
  };

  return (
    <div style={{ flex: 1, overflowY: "auto", paddingBottom: 80, background: T.bg }}>
      <PageHeader title="Import" subtitle="Chargez un fichier Excel ou CSV" />

      {/* Drop zone */}
      <div style={{ padding: 16 }}>
        <div
          onDrop={onDrop}
          onDragOver={(e) => e.preventDefault()}
          onClick={() => fileInputRef.current?.click()}
          style={{
            border: `1.5px dashed ${T.border}`,
            borderRadius: 14, padding: "48px 24px",
            textAlign: "center", cursor: "pointer",
            background: T.bgCard, transition: "all 0.2s",
          }}
          onMouseEnter={(e) => { (e.currentTarget as HTMLDivElement).style.borderColor = T.accent; }}
          onMouseLeave={(e) => { (e.currentTarget as HTMLDivElement).style.borderColor = T.border; }}
        >
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls,.csv"
            style={{ display: "none" }}
            onChange={onFileChange}
          />
          <div style={{ fontSize: 40, marginBottom: 14, opacity: 0.3 }}>⊞</div>
          <div style={{ fontSize: 14, color: T.textMuted, marginBottom: 8 }}>
            Glissez un fichier ou cliquez pour sélectionner
          </div>
          <div style={{ fontSize: 10, color: T.textDim, letterSpacing: "0.1em", textTransform: "uppercase" }}>
            .xlsx · .xls · .csv
          </div>
        </div>
      </div>

      {/* Fichier actif */}
      {parsed && (
        <div style={{ padding: "0 16px" }}>
          <div style={{
            background: T.bgCard, border: `1px solid ${T.border}`,
            borderLeft: `4px solid ${T.success}`,
            borderRadius: 10, padding: "12px 14px", marginBottom: 14,
          }}>
            <div style={{ color: T.success, fontSize: 11, fontWeight: 700, marginBottom: 4 }}>
              ✅ Fichier chargé
            </div>
            <div style={{ color: T.text, fontSize: 13, fontFamily: "monospace", fontWeight: 700 }}>
              📄 {fileName}
            </div>
            {sheetNames.length > 1 && (
              <div style={{ color: T.textMuted, fontSize: 11, marginTop: 4 }}>
                {sheetNames.length} onglets détectés
              </div>
            )}
          </div>

          {/* Sélecteur d'onglets */}
          <SheetTabsSection />

          {/* CTA vers tableau */}
          <div style={{ marginTop: 20 }}>
            <Btn onClick={() => setActiveTab("tableau")} color={T.accent} textColor="#0F172A" fullWidth>
              📊 Voir le tableau →
            </Btn>
          </div>
        </div>
      )}

      {!parsed && (
        <div style={{ padding: "0 16px" }}>
          <div style={{ background: T.bgCard, border: `1px solid ${T.border}`, borderRadius: 12, padding: 16, marginTop: 8 }}>
            <div style={{ color: T.accent, fontSize: 11, fontWeight: 700, marginBottom: 10, textTransform: "uppercase" }}>
              ⚡ Fonctionnalités
            </div>
            {[
              "Import Excel (.xlsx / .xls / .csv)",
              "Détection automatique des en-têtes",
              "Gestion multi-onglets",
              "Masquage colonnes / lignes",
              "Renommage des en-têtes",
              "Ajout de nouvelles lignes",
              "Grisage des valeurs répétitives",
              "Export du fichier nettoyé",
            ].map((f) => (
              <div key={f} style={{
                display: "flex", alignItems: "center", gap: 8,
                padding: "6px 0", borderBottom: `1px solid ${T.border}22`,
                color: T.textMuted, fontSize: 12,
              }}>
                <span style={{ color: T.success, fontSize: 10 }}>◆</span>
                {f}
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────
// 6. PAGE TABLEAU
// ─────────────────────────────────────────────

// Barre de navigation rapide entre onglets (usage dans TablePage)
function SheetSwitcher() {
  const {
    sheetNames, hiddenSheets, activeSheet, setActiveSheet,
    workbook, loadSheet,
  } = useApp();

  const visibleSheets = sheetNames.filter((n) => !hiddenSheets.has(n));
  if (visibleSheets.length <= 1) return null;

  return (
    <div style={{
      display: "flex", overflowX: "auto", gap: 4,
      padding: "8px 12px",
      background: T.bgDark,
      borderBottom: `1px solid ${T.border}`,
    }}>
      {visibleSheets.map((name) => {
        const isActive = activeSheet === name;
        return (
          <button
            key={name}
            onClick={() => {
              if (!isActive && workbook) {
                setActiveSheet(name);
                loadSheet(workbook, name);
              }
            }}
            style={{
              flexShrink: 0,
              background: isActive ? T.border : T.bgCard,
              border: `1px solid ${isActive ? T.accent : T.border2}`,
              borderRadius: 7, padding: "5px 12px",
              cursor: isActive ? "default" : "pointer",
              color: isActive ? T.accent : T.textMuted,
              fontSize: 11, fontWeight: 700,
              fontFamily: "'Share Tech Mono', monospace",
              transition: "all 0.15s",
            }}
            onMouseEnter={(e) => {
              if (!isActive) (e.currentTarget as HTMLButtonElement).style.borderColor = T.accent;
            }}
            onMouseLeave={(e) => {
              if (!isActive) (e.currentTarget as HTMLButtonElement).style.borderColor = T.border2;
            }}
          >
            {name}
          </button>
        );
      })}
      <span style={{
        marginLeft: "auto", flexShrink: 0,
        color: T.textDim, fontSize: 10, alignSelf: "center",
        letterSpacing: "0.05em",
      }}>
        {visibleSheets.indexOf(activeSheet ?? "") + 1}/{visibleSheets.length}
      </span>
    </div>
  );
}

function TablePage() {
  const {
    parsed, headers, setHeaders,
    hiddenCols, setHiddenCols, hiddenRows, setHiddenRows,
    selectMode, setSelectMode, selectedItems, setSelectedItems,
    editingHeader, setEditingHeader,
    dimRepeated, setDimRepeated,
    fileName, sheetNames,
    allRows, repetitiveByCol, addedRows,
    showToast, setActiveTab,
  } = useApp();

  const [showAddRow, setShowAddRow] = useState(false);

  if (!parsed) {
    return (
      <div style={{ flex: 1, overflowY: "auto", paddingBottom: 80, background: T.bg }}>
        <PageHeader title="Tableau" />
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
    setHiddenCols((prev) => new Set([...prev, ...selectedItems]));
    setSelectedItems(new Set());
    setSelectMode("none");
    showToast(`${selectedItems.size} colonne(s) masquée(s)`, "info");
  };

  const applyRowDeletion = () => {
    if (selectMode !== "row") return;
    setHiddenRows((prev) => new Set([...prev, ...selectedItems]));
    setSelectedItems(new Set());
    setSelectMode("none");
    showToast(`${selectedItems.size} ligne(s) masquée(s)`, "info");
  };

  const visibleColCount = headers.filter((_, i) => !hiddenCols.has(i)).length;
  const visibleRowCount = allRows.filter((_, i) => !hiddenRows.has(i)).length;

  // CSS injected via <style>
  const css = `
    .ste-table { border-collapse: collapse; width: 100%; font-size: 12px; }
    .ste-table thead tr th {
      background: #131E2E; color: ${T.textMuted};
      font-size: 10px; letter-spacing: 0.1em; text-transform: uppercase;
      padding: 9px 12px; text-align: left;
      border-bottom: 1px solid ${T.border};
      white-space: nowrap; position: sticky; top: 0; z-index: 2;
      font-weight: 600;
    }
    .ste-table td {
      padding: 7px 12px; border-bottom: 1px solid ${T.border}22;
      color: #C8D8E8; white-space: nowrap;
      max-width: 200px; overflow: hidden; text-overflow: ellipsis;
    }
    .ste-table tr:hover td { background: ${T.rowHover}; }
    .ste-table tr.row-sel td { background: ${T.selRowBg} !important; color: ${T.selRowTxt}; }
    .ste-table tr.added-row td { background: #0A1F10 !important; }
    .ste-table tr.added-row:hover td { background: #0F2A18 !important; }

    .th-num { background: #0E1826 !important; border-right: 1px solid ${T.border} !important; width: 38px; min-width: 38px; text-align: center !important; }
    .td-num { color: ${T.textDim} !important; font-size: 11px !important; background: #0E1826 !important; border-right: 1px solid ${T.border} !important; min-width: 38px; width: 38px; text-align: center !important; user-select: none; }
    .td-num.row-sel-click { cursor: pointer; }
    .td-num.row-sel-click:hover { color: ${T.accent} !important; background: #1E3A5F !important; }

    .th-col-sel { cursor: pointer; }
    .th-col-sel:hover { background: #1A2D3D !important; color: ${T.accent} !important; }
    .th-col-del { background: #2A0E0E !important; color: ${T.error} !important; border-bottom-color: ${T.error} !important; }

    .cell-rep { color: ${T.repeat} !important; font-style: italic; background: ${T.repeatBg}; }
    .header-input {
      background: transparent; border: none; border-bottom: 1px solid ${T.accent};
      color: ${T.accent}; font-family: 'Share Tech Mono', monospace;
      font-size: 10px; letter-spacing: 0.08em; text-transform: uppercase;
      width: 100%; min-width: 50px; outline: none; padding: 2px 0;
    }
    .header-txt { display: flex; align-items: center; gap: 4px; cursor: pointer; }
    .header-txt:hover .pencil { opacity: 1; }
    .pencil { opacity: 0; font-size: 9px; color: ${T.accent}; transition: opacity 0.15s; }

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
      fontFamily: 'Share Tech Mono', monospace;
    }
    .dim-chip:hover { border-color: ${T.accent}; }
    .dim-chip.on { color: #A78BFA; border-color: #3D2A6A; background: #160F2A; }
    .dim-dot { width: 7px; height: 7px; border-radius: 50%; background: currentColor; flex-shrink: 0; }

    .info-bar {
      margin-bottom: 10px; padding: 7px 12px;
      background: #0F2A1A; border: 1px solid #1A3A28;
      border-radius: 7px; font-size: 11px; color: ${T.success};
    }
    .info-bar.del { background: #1F0A0A; border-color: #3A1515; color: ${T.error}; }
  `;

  return (
    <div style={{ flex: 1, display: "flex", flexDirection: "column", background: T.bg, paddingBottom: 64 }}>
      <style>{css}</style>

      <PageHeader
        title="Tableau"
        subtitle={fileName ? `📄 ${fileName}` : undefined}
      />
      <StatsBar />

      {/* Sélecteur d'onglets inline (si multi-sheets) */}
      {sheetNames.length > 1 && (
        <SheetSwitcher />
      )}

      {/* Toolbar */}
      <div style={{
        padding: "10px 12px", background: T.bgDark,
        borderBottom: `1px solid ${T.border}`,
        display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center",
      }}>
        {/* Dim chip */}
        <div className={`dim-chip${dimRepeated ? " on" : ""}`} onClick={() => setDimRepeated((v) => !v)}>
          <span className="dim-dot" />Répétitions
        </div>

        {/* Add row */}
        <button className="ste-btn" onClick={() => setShowAddRow(true)}>+ Ligne</button>

        {/* Col select */}
        <button
          className={`ste-btn${selectMode === "col" ? " active" : ""}`}
          onClick={() => { setSelectMode(selectMode === "col" ? "none" : "col"); setSelectedItems(new Set()); }}
        >
          {selectMode === "col" ? "✓ " : ""}Colonnes
        </button>

        {/* Row select */}
        <button
          className={`ste-btn${selectMode === "row" ? " active" : ""}`}
          onClick={() => { setSelectMode(selectMode === "row" ? "none" : "row"); setSelectedItems(new Set()); }}
        >
          {selectMode === "row" ? "✓ " : ""}Lignes
        </button>

        {/* Apply / confirm */}
        {selectedItems.size > 0 && (
          <button className="ste-btn danger" onClick={selectMode === "col" ? applyColAction : applyRowDeletion}>
            {selectMode === "col"
              ? `Masquer ${selectedItems.size} col.`
              : `Masquer ${selectedItems.size} ligne(s)`}
          </button>
        )}

        {/* Restore */}
        {(hiddenCols.size > 0 || hiddenRows.size > 0) && (
          <button className="ste-btn" onClick={() => { setHiddenCols(new Set()); setHiddenRows(new Set()); }}>
            Restaurer tout
          </button>
        )}
      </div>

      {/* Hints */}
      <div style={{ padding: "8px 12px 0" }}>
        {selectMode === "col" && (
          <div className="info-bar del">→ Cliquez les en-têtes à masquer, puis confirmez</div>
        )}
        {selectMode === "row" && (
          <div className="info-bar del">→ Cliquez les numéros de lignes à masquer, puis confirmez</div>
        )}
        {selectMode === "none" && (hiddenCols.size > 0 || hiddenRows.size > 0) && (
          <div className="info-bar">
            {hiddenCols.size > 0 && <span>{hiddenCols.size} col. masquée(s) · </span>}
            {hiddenRows.size > 0 && <span>{hiddenRows.size} ligne(s) masquée(s)</span>}
          </div>
        )}
        {addedRows.length > 0 && selectMode === "none" && (
          <div style={{ marginBottom: 8, fontSize: 10, color: T.success }}>
            + {addedRows.length} ligne(s) ajoutée(s) manuellement
          </div>
        )}
        {dimRepeated && visibleRowCount > 0 && (
          <div style={{ marginBottom: 8, fontSize: 10, color: T.repeat, letterSpacing: "0.04em" }}>
            ◆ Valeurs grisées = ≥35% des lignes de leur colonne
          </div>
        )}
      </div>

      {/* Table */}
      <div style={{
        flex: 1, overflowX: "auto", overflowY: "auto",
        margin: "0 12px 12px", maxHeight: "calc(100dvh - 320px)",
        border: `1px solid ${T.border}`, borderRadius: 8, background: T.bgCard,
      }}>
        <table className="ste-table">
          <thead>
            <tr>
              <th className="th-num">#</th>
              {headers.map((h, ci) => {
                if (hiddenCols.has(ci)) return null;
                const isSel = selectMode === "col" && selectedItems.has(ci);
                const isEditing = editingHeader === ci;
                let cls = selectMode === "col" ? "th-col-sel" : "";
                if (isSel) cls += " th-col-del";
                return (
                  <th key={ci} className={cls}
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
                        onChange={(e) => {
                          const next = [...headers];
                          next[ci] = e.target.value;
                          setHeaders(next);
                        }}
                        onBlur={() => setEditingHeader(null)}
                        onKeyDown={(e) => { if (e.key === "Enter" || e.key === "Escape") setEditingHeader(null); }}
                        onClick={(e) => e.stopPropagation()}
                      />
                    ) : (
                      <span className="header-txt" title="Cliquer pour renommer">
                        {h}
                        {selectMode === "none" && <span className="pencil">✎</span>}
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
                <tr key={ri} className={`${isRowSel ? "row-sel" : ""} ${isAdded ? "added-row" : ""}`}>
                  <td
                    className={`td-num${selectMode === "row" ? " row-sel-click" : ""}`}
                    onClick={() => selectMode === "row" && toggleSelectItem(ri)}
                    title={isAdded ? "Ligne ajoutée manuellement" : undefined}
                  >
                    {isAdded
                      ? <span style={{ color: T.success, fontSize: 9 }}>+{ri + 1}</span>
                      : ri + 1
                    }
                  </td>
                  {headers.map((_, ci) => {
                    if (hiddenCols.has(ci)) return null;
                    const cell = row[ci] ?? null;
                    const strVal = cell !== null ? String(cell) : null;
                    const isRep = !isAdded && dimRepeated && strVal !== null && (repetitiveByCol.get(ci)?.has(strVal) ?? false);
                    return (
                      <td key={ci} title={strVal ?? ""} className={isRep ? "cell-rep" : ""}>
                        {strVal ?? <span style={{ color: T.textDim }}>—</span>}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
            {allRows.every((_, i) => hiddenRows.has(i)) && (
              <tr>
                <td colSpan={visibleColCount + 1} style={{ textAlign: "center", padding: "32px", color: T.textDim }}>
                  Toutes les lignes sont masquées
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {/* Add row modal */}
      {showAddRow && <AddRowModal onClose={() => setShowAddRow(false)} />}
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

function ExportPage() {
  const {
    parsed, workbook, sheetNames, hiddenSheets,
    exportFileName, setExportFileName,
    showToast, setActiveTab,
    setParsed, setFileName, setWorkbook, setSheetNames, setActiveSheet,
    setAddedRows, setHiddenSheets, setSheetSelectMode, setSelectedSheets,
    sheetStates,
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
