// ============================================================
// STEtruc — AppContext + AppProvider
// ============================================================

import {
  useState, useCallback, useRef, useMemo, useEffect,
  useContext, createContext,
} from "react";
import * as XLSX from "xlsx";
import {
  Tab, SelectMode, SheetSelectMode, Destination, ParsedData,
  SheetState, RawData, CellValue,
  detectHeaders, computeRepetitiveValues,
} from "./types";

export interface AppState {
  activeTab: Tab;
  setActiveTab: (t: Tab) => void;
  toast: { msg: string; type: "success" | "error" | "info" } | null;
  showToast: (msg: string, type?: "success" | "error" | "info") => void;
  fileName: string | null;
  setFileName: (n: string | null) => void;
  workbook: XLSX.WorkBook | null;
  setWorkbook: (w: XLSX.WorkBook | null) => void;
  sheetNames: string[];
  setSheetNames: (s: string[]) => void;
  activeSheet: string | null;
  setActiveSheet: (s: string | null) => void;
  parsed: ParsedData | null;
  setParsed: (p: ParsedData | null) => void;
  headers: string[];
  setHeaders: (h: string[]) => void;
  addedRows: RawData;
  setAddedRows: React.Dispatch<React.SetStateAction<RawData>>;
  hiddenCols: Set<number>;
  setHiddenCols: React.Dispatch<React.SetStateAction<Set<number>>>;
  hiddenRows: Set<number>;
  setHiddenRows: React.Dispatch<React.SetStateAction<Set<number>>>;
  hiddenSheets: Set<string>;
  setHiddenSheets: React.Dispatch<React.SetStateAction<Set<string>>>;
  selectMode: SelectMode;
  setSelectMode: React.Dispatch<React.SetStateAction<SelectMode>>;
  selectedItems: Set<number>;
  setSelectedItems: React.Dispatch<React.SetStateAction<Set<number>>>;
  editingHeader: number | null;
  setEditingHeader: React.Dispatch<React.SetStateAction<number | null>>;
  sheetSelectMode: SheetSelectMode;
  setSheetSelectMode: React.Dispatch<React.SetStateAction<SheetSelectMode>>;
  selectedSheets: Set<string>;
  setSelectedSheets: React.Dispatch<React.SetStateAction<Set<string>>>;
  dimRepeated: boolean;
  setDimRepeated: React.Dispatch<React.SetStateAction<boolean>>;
  exportFileName: string;
  setExportFileName: React.Dispatch<React.SetStateAction<string>>;
  splitFormats: Record<string, string>;
  setSplitFormats: React.Dispatch<React.SetStateAction<Record<string, string>>>;
  autoRefFmt: boolean;
  setAutoRefFmt: React.Dispatch<React.SetStateAction<boolean>>;
  poidsUnit: "t" | "kg";
  setPoidsUnit: React.Dispatch<React.SetStateAction<"t" | "kg">>;
  mapping: { rang: string; reference: string; poids: string; dch: string };
  setMapping: React.Dispatch<React.SetStateAction<{ rang: string; reference: string; poids: string; dch: string }>>;
  extras: { col: string; label: string }[];
  setExtras: React.Dispatch<React.SetStateAction<{ col: string; label: string }[]>>;
  pointedRows: Set<number>;
  setPointedRows: React.Dispatch<React.SetStateAction<Set<number>>>;
  rowOverrides: Map<number, Record<number, CellValue>>;
  setRowOverrides: React.Dispatch<React.SetStateAction<Map<number, Record<number, CellValue>>>>;
  destinations: Destination[];
  setDestinations: React.Dispatch<React.SetStateAction<Destination[]>>;
  selectedDest: string;
  setSelectedDest: React.Dispatch<React.SetStateAction<string>>;
  rowDestinations: Map<number, string>;
  setRowDestinations: React.Dispatch<React.SetStateAction<Map<number, string>>>;
  reassignedRows: Map<number, { from: string; to: string }[]>;
  setReassignedRows: React.Dispatch<React.SetStateAction<Map<number, { from: string; to: string }[]>>>;
  winwinModalOpen: boolean;
  setWinwinModalOpen: React.Dispatch<React.SetStateAction<boolean>>;
  tallyPrev: Record<string, { qty: string; weight: string }>;
  setTallyPrev: (v: Record<string, { qty: string; weight: string }> | ((p: Record<string, { qty: string; weight: string }>) => Record<string, { qty: string; weight: string }>)) => void;
  chargementMaxi: Record<string, { qty: string; weight: string }>;
  setChargementMaxi: (v: Record<string, { qty: string; weight: string }> | ((p: Record<string, { qty: string; weight: string }>) => Record<string, { qty: string; weight: string }>)) => void;
  dechargementMaxi: Record<string, { qty: string; weight: string }>;
  setDechargementMaxi: (v: Record<string, { qty: string; weight: string }> | ((p: Record<string, { qty: string; weight: string }>) => Record<string, { qty: string; weight: string }>)) => void;
  loadSheet: (wb: XLSX.WorkBook, sheet: string) => void;
  handleFile: (file: File) => void;
  allRows: RawData;
  repetitiveByCol: Map<number, Set<string>>;
  sheetStates: React.MutableRefObject<Map<string, SheetState>>;
}

const AppCtx = createContext<AppState | null>(null);

export const useApp = () => {
  const ctx = useContext(AppCtx);
  if (!ctx) throw new Error("useApp outside provider");
  return ctx;
};

function lsGet<T>(key: string, fallback: T): T {
  try { const v = localStorage.getItem(key); return v !== null ? (JSON.parse(v) as T) : fallback; } catch { return fallback; }
}
function lsSet(key: string, val: unknown) { try { localStorage.setItem(key, JSON.stringify(val)); } catch {} }

export function AppProvider({ children }: { children: React.ReactNode }) {
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
    { name: "STOCK", color: "#64748B", excludeFromReport: true },
  ];
  const [destinations, setDestinations] = useState<Destination[]>(() => {
    const saved = lsGet<Destination[]>("ste_destinations", defaultDestinations);
    if (!saved.some(d => d.name === "STOCK")) {
      const hneIdx = saved.findIndex(d => d.name === "HNE");
      const next = [...saved];
      next.splice(hneIdx >= 0 ? hneIdx + 1 : next.length, 0, { name: "STOCK", color: "#64748B", excludeFromReport: true });
      return next;
    }
    return saved;
  });

  const [tallyPrevRaw, setTallyPrevRaw] = useState<Record<string, { qty: string; weight: string }>>(() => lsGet("ste_tallyPrev", {}));
  const setTallyPrev = (v: Record<string, { qty: string; weight: string }> | ((p: Record<string, { qty: string; weight: string }>) => Record<string, { qty: string; weight: string }>)) => {
    setTallyPrevRaw(prev => { const next = typeof v === "function" ? v(prev) : v; lsSet("ste_tallyPrev", next); return next; });
  };
  const [chargementMaxiRaw, setChargementMaxiRaw] = useState<Record<string, { qty: string; weight: string }>>(() => lsGet("ste_chargementMaxi", {}));
  const setChargementMaxi = (v: Record<string, { qty: string; weight: string }> | ((p: Record<string, { qty: string; weight: string }>) => Record<string, { qty: string; weight: string }>)) => {
    setChargementMaxiRaw(prev => { const next = typeof v === "function" ? v(prev) : v; lsSet("ste_chargementMaxi", next); return next; });
  };
  const [dechargementMaxiRaw, setDechargementMaxiRaw] = useState<Record<string, { qty: string; weight: string }>>(() => lsGet("ste_dechargementMaxi", {}));
  const setDechargementMaxi = (v: Record<string, { qty: string; weight: string }> | ((p: Record<string, { qty: string; weight: string }>) => Record<string, { qty: string; weight: string }>)) => {
    setDechargementMaxiRaw(prev => { const next = typeof v === "function" ? v(prev) : v; lsSet("ste_dechargementMaxi", next); return next; });
  };

  const [selectedDest, setSelectedDest] = useState<string>("");
  const [rowDestinations, setRowDestinations] = useState<Map<number, string>>(new Map());
  const [reassignedRows, setReassignedRows] = useState<Map<number, { from: string; to: string }[]>>(new Map());
  const [winwinModalOpen, setWinwinModalOpen] = useState(false);

  const sheetStates = useRef<Map<string, SheetState>>(new Map());

  const showToast = useCallback((msg: string, type: "success" | "error" | "info" = "info") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3000);
  }, []);

  useEffect(() => { lsSet("ste_splitFormats", splitFormats); }, [splitFormats]);
  useEffect(() => { lsSet("ste_autoRefFmt", autoRefFmt); }, [autoRefFmt]);
  useEffect(() => { lsSet("ste_poidsUnit", poidsUnit); }, [poidsUnit]);
  useEffect(() => { lsSet("ste_mapping", mapping); }, [mapping]);
  useEffect(() => { lsSet("ste_extras", extras); }, [extras]);
  useEffect(() => { lsSet("ste_destinations", destinations); }, [destinations]);

  // Refs miroirs pour accès synchrone
  const activeSheetRef = useRef(activeSheet);
  const parsedRef      = useRef(parsed);
  const headersRef     = useRef(headers);
  const addedRowsRef   = useRef(addedRows);
  const hiddenColsRef  = useRef(hiddenCols);
  const hiddenRowsRef  = useRef(hiddenRows);
  activeSheetRef.current = activeSheet;
  parsedRef.current      = parsed;
  headersRef.current     = headers;
  addedRowsRef.current   = addedRows;
  hiddenColsRef.current  = hiddenCols;
  hiddenRowsRef.current  = hiddenRows;

  const loadSheet = useCallback((wb: XLSX.WorkBook, sheetName: string) => {
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
  }, []);

  const handleFile = useCallback((file: File) => {
    sheetStates.current.clear();
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer);
      const wb = XLSX.read(data, { type: "array" });
      setWorkbook(wb);
      setSheetNames(wb.SheetNames);
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
      ["ste_tallyPrev", "ste_chargementMaxi", "ste_dechargementMaxi"].forEach(k => { try { localStorage.removeItem(k); } catch {} });
      setTallyPrev({});
      setChargementMaxi({});
      setDechargementMaxi({});
      loadSheet(wb, defaultSheet);
    };
    reader.readAsArrayBuffer(file);
  }, [loadSheet]);

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
      tallyPrev: tallyPrevRaw, setTallyPrev,
      chargementMaxi: chargementMaxiRaw, setChargementMaxi,
      dechargementMaxi: dechargementMaxiRaw, setDechargementMaxi,
      loadSheet, handleFile,
      allRows, repetitiveByCol,
      sheetStates,
    }}>
      {children}
    </AppCtx.Provider>
  );
}
