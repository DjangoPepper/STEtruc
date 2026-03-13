// ============================================================
// STEtruc — Types, Utils, Palette
// ============================================================

export type CellValue = string | number | boolean | null;
export type RawData = CellValue[][];
export type Tab = "import" | "iec" | "tableau" | "rapport" | "export";
export type SelectMode = "none" | "col" | "row" | "dest";
export type SheetSelectMode = "none" | "delete" | "keep";

export type Destination = { name: string; color: string; excludeFromReport?: boolean };

export interface ParsedData {
  headers: string[];
  rows: RawData;
  headerRowIndex: number;
}

export interface SheetState {
  parsed: ParsedData;
  headers: string[];
  addedRows: RawData;
  hiddenCols: Set<number>;
  hiddenRows: Set<number>;
}

// ─── Palette navy (coil-deploy) ───────────────────────────────
export const T = {
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

// ─── Utils ────────────────────────────────────────────────────

export function detectHeaders(data: RawData): ParsedData {
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

export function computeRepetitiveValues(rows: RawData, colIndex: number, threshold = 0.35): Set<string> {
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

export function applyGrouping(value: string, pattern: string): string {
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

export function thsep(val: string): string {
  const [int, dec] = val.split(".");
  const separated = int.replace(/\B(?=(\d{3})+(?!\d))/g, " ");
  return dec !== undefined ? separated + "." + dec : separated;
}

export function autoFormatRef(val: string, globalFmt: string): string {
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
