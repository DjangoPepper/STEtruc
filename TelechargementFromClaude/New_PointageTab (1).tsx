// ============================================================
// STEtruc — Onglet Pointage
// ============================================================

import { useState, useRef, useMemo, useEffect } from "react";
import { T, CellValue, applyGrouping, thsep, autoFormatRef } from "../types";
import { useApp } from "../AppContext";
import { Btn, EmptyState, PointageInfos, AddRowModal } from "../components";

export default function PointageTab() {
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
  const [quickConfirmModal, setQuickConfirmModal] = useState<{ ri: number; ref: string; rang: string; poids: string; dest: string; destColor: string } | null>(null);
  const [quickConfirmCountdown, setQuickConfirmCountdown] = useState(5);
  const [noDestWarning, setNoDestWarning] = useState(false);
  const [refGroupModal, setRefGroupModal] = useState(false);
  const [refGroupInput, setRefGroupInput] = useState("");
  const longPressRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  useEffect(() => {
    if (!quickConfirmModal) return;
    setQuickConfirmCountdown(5);
    const interval = setInterval(() => {
      setQuickConfirmCountdown(prev => {
        if (prev <= 1) { clearInterval(interval); setQuickConfirmModal(null); return 0; }
        return prev - 1;
      });
    }, 1000);
    return () => clearInterval(interval);
  }, [quickConfirmModal?.ri]); // eslint-disable-line react-hooks/exhaustive-deps

  if (!parsed) {
    return (
      <div style={{ flex: 1, overflowY: "auto", paddingBottom: 80, background: T.bg }}>
        <div style={{ padding: "16px 16px 12px", background: T.bgDark, borderBottom: `1px solid ${T.border}` }}>
          <div style={{ color: T.textDim, fontSize: 9, letterSpacing: "0.2em", textTransform: "uppercase", marginBottom: 4 }}>STEtruc</div>
          <h1 style={{ color: T.text, fontSize: 20, fontWeight: 700 }}>Pointage</h1>
        </div>
        <EmptyState icon="📁" text="Aucun fichier chargé" sub="Importez un fichier Excel d'abord" />
        <div style={{ padding: 16 }}>
          <Btn onClick={() => setActiveTab("import")} color={T.accent} textColor="#0F172A" fullWidth>⬇️ Aller à l'import</Btn>
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
    setSelectedItems(new Set()); setSelectMode("none");
    showToast(`${selectedItems.size} colonne(s) masquée(s)`, "info");
  };

  const applyRowDeletion = () => {
    if (selectMode !== "row") return;
    _setHiddenRows((prev) => new Set([...prev, ...selectedItems]));
    setSelectedItems(new Set()); setSelectMode("none");
    showToast(`${selectedItems.size} ligne(s) masquée(s)`, "info");
  };

  const prepaIdx = headers.findIndex((h) => /^prepa$/i.test(h.trim()));
  const destIdx  = headers.findIndex((h) => /^dest(ination)?$/i.test(h.trim()));
  const visibleCols: number[] = headers.map((_, i) => i).filter((i) => !hiddenCols.has(i) && i !== prepaIdx);
  if (prepaIdx >= 0 && !hiddenCols.has(prepaIdx)) visibleCols.push(prepaIdx);

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
  const colLabel = (i: number): string => mappingLabels.get(i) ?? (i === destIdx ? "DEST" : headers[i]);

  const baseRows = allRows.map((row, ri) => ({ row, ri })).filter(({ ri }) => !hiddenRows.has(ri));
  const filteredRows = baseRows.filter(({ row, ri }) =>
    visibleCols.every((ci) => {
      const f = (colFilters[ci] ?? "").trim().toLowerCase();
      if (!f) return true;
      // Pour la colonne Dest, la valeur réelle vient de rowDestinations
      const cell = ci === destIdx ? (rowDestinations.get(ri) ?? row[ci]) : row[ci];
      return cell !== null && String(cell).toLowerCase().includes(f);
    })
  );
  const sortedRows = sortCol !== null
    ? [...filteredRows].sort((a, b) => {
        // Pour la colonne Dest, comparer les valeurs de rowDestinations
        const va = sortCol === destIdx
          ? (rowDestinations.get(a.ri) ?? a.row[sortCol] ?? "")
          : (a.row[sortCol] ?? "");
        const vb = sortCol === destIdx
          ? (rowDestinations.get(b.ri) ?? b.row[sortCol] ?? "")
          : (b.row[sortCol] ?? "");
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

  const refColIdx   = mapping.reference ? headers.indexOf(mapping.reference) : -1;
  const poidsColIdx = mapping.poids     ? headers.indexOf(mapping.poids)     : -1;
  const dchColIdx   = mapping.dch       ? headers.indexOf(mapping.dch)       : -1;
  const rangColIdx  = mapping.rang      ? headers.indexOf(mapping.rang)      : -1;

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

  const numColW = Math.max(28, String(allRows.length).length * 8 + 14);

  const colWidths = useMemo(() => {
    const CH_DATA = 9.5; const CH_HDR = 8; const MIN_W = 38; const MAX_W = 200;
    const PAD = 20; const NUM_W = numColW; const VIEWPORT = 380;
    const sample = sortedRows.slice(0, 80);
    const ideals = new Map<number, number>();
    visibleCols.forEach((ci) => {
      const hLen = colLabel(ci).replace(/[^\x00-\x7F]/g, "  ").length;
      let maxData = 0;
      for (const { row } of sample) { const v = row[ci]; if (v !== null && v !== undefined) maxData = Math.max(maxData, String(v).length); }
      ideals.set(ci, Math.min(MAX_W, Math.max(MIN_W, Math.max(hLen * CH_HDR, maxData * CH_DATA) + PAD)));
    });
    const totalIdeal = NUM_W + [...ideals.values()].reduce((a, b) => a + b, 0);
    const widths = new Map<number, number>();
    if (totalIdeal <= VIEWPORT) {
      const extra = (VIEWPORT - totalIdeal) / visibleCols.length;
      visibleCols.forEach((ci) => widths.set(ci, (ideals.get(ci) ?? MIN_W) + extra));
    } else {
      const available = VIEWPORT - NUM_W - visibleCols.length * MIN_W;
      const idealTotal = [...ideals.values()].reduce((a, b) => a + b, 0) - visibleCols.length * MIN_W;
      const ratio = idealTotal > 0 ? Math.max(0, available) / idealTotal : 0;
      visibleCols.forEach((ci) => { const ideal = ideals.get(ci) ?? MIN_W; widths.set(ci, Math.round(MIN_W + (ideal - MIN_W) * ratio)); });
    }
    return widths;
  }, [sortedRows, visibleCols, headers, mapping, extras, poidsUnit, numColW]); // eslint-disable-line react-hooks/exhaustive-deps

  const tableFontSize = useMemo(() => {
    const PAD = 20; const CH_RATIO = 0.60;
    let maxFs = 14;
    visibleCols.forEach(ci => {
      const w = (colWidths.get(ci) ?? 38) - PAD;
      if (w <= 0) return;
      let maxLen = 0;
      for (const { row: r } of sortedRows.slice(0, 100)) { const v = r[ci]; if (v !== null && v !== undefined) maxLen = Math.max(maxLen, String(v).length); }
      if (maxLen > 0) { const fs = Math.floor(w / (maxLen * CH_RATIO)); maxFs = Math.min(maxFs, fs); }
    });
    return Math.max(9, maxFs);
  }, [colWidths, visibleCols, sortedRows]);

  const anomalyRowSet = useMemo(() => {
    if (allRows.length === 0) return new Set<number>();
    const visibleRIs = allRows.map((_, ri) => ri).filter((ri) => !hiddenRows.has(ri));
    if (visibleRIs.length === 0) return new Set<number>();
    const counts = visibleRIs.map((ri) => allRows[ri].filter((v) => v !== null && v !== undefined && String(v).trim() !== "").length);
    const sorted = [...counts].sort((a, b) => a - b);
    const median = sorted[Math.floor(sorted.length / 2)];
    const anomalous = new Set<number>();
    visibleRIs.forEach((ri, idx) => { if (counts[idx] < median - 1) anomalous.add(ri); });
    return anomalous;
  }, [allRows, hiddenRows]);

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
    if (openToolbar && selectMode === "none") { setSelectMode("col"); setSelectedItems(new Set([ci])); return; }
    if (selectMode !== "none") return;
    if (sortCol !== ci) { setSortCol(ci); setSortDir("asc"); }
    else if (sortDir === "asc") { setSortDir("desc"); }
    else { setSortCol(null); }
    setPageIdx(0);
  };

  const css = `
    .pt-table { border-collapse: collapse; width: 100%; table-layout: fixed; font-size: ${tableFontSize}px; }
    .pt-table thead th { background: #131E2E; color: ${T.textMuted}; font-size: 14px; letter-spacing: 0.07em; text-transform: uppercase; padding: 0; text-align: left; border-bottom: 2px solid ${T.border}; border-right: 1px solid #7C3AED88; white-space: normal; word-break: break-word; position: sticky; top: 0; z-index: 2; font-weight: 700; vertical-align: top; }
    .pt-table thead th .th-inner { padding: 6px 8px 3px; display: flex; align-items: center; gap: 3px; cursor: pointer; user-select: none; transition: color 0.12s; }
    .pt-table thead th .th-inner:hover { color: ${T.accent}; }
    .pt-table thead th.th-sorted { background: #0F1F30; }
    .pt-table thead th.th-sorted .th-inner { color: ${T.accent}; }
    .pt-table thead th .th-filter { padding: 2px 6px 5px; }
    .pt-table thead th .th-filter input { width: 100%; background: ${T.bg}; border: 1px solid ${T.border}55; border-radius: 4px; color: ${T.textMuted}; font-size: 10px; padding: 2px 5px; outline: none; font-family: 'Share Tech Mono', monospace; box-sizing: border-box; }
    .pt-table thead th .th-filter input:focus { border-color: ${T.accent}55; }
    .pt-table th.th-num-h { background: #0E1826; border-right: 1px solid ${T.border}; width: ${numColW}px; min-width: ${numColW}px; text-align: center; position: sticky; top: 0; z-index: 3; vertical-align: top; padding: 0; }
    .pt-table td.td-num { background: #0E1826; border-right: 1px solid #7C3AED88; width: ${numColW}px; min-width: ${numColW}px; text-align: center; color: ${T.textDim}; font-size: 10px; padding: 7px 4px; white-space: nowrap; user-select: none; }
    .pt-table td { padding: 9px 10px; border-bottom: 1px solid ${T.border}22; border-right: 1px solid #7C3AED44; color: #C8D8E8; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .pt-table tr:hover td, .pt-table tr:hover .td-num { background: ${T.rowHover}; }
    .pt-table tr.row-sel td { background: ${T.selRowBg} !important; color: ${T.selRowTxt}; }
    .pt-table tr.added-row td { background: #0A1F10 !important; }
    .pt-table tr.click-row { cursor: pointer; }
    .th-col-sel .th-inner { cursor: pointer; }
    .th-col-sel .th-inner:hover { color: ${T.error} !important; }
    .th-col-del { background: #2A0E0E !important; }
    .th-col-del .th-inner { color: ${T.error} !important; }
    .cell-rep { color: ${T.repeat} !important; font-style: italic; background: ${T.repeatBg}; }
    .header-input { background: transparent; border: none; border-bottom: 1px solid ${T.accent}; color: ${T.accent}; font-family: 'Share Tech Mono', monospace; font-size: 10px; letter-spacing: 0.08em; text-transform: uppercase; width: 100%; min-width: 50px; outline: none; padding: 2px 0; }
    .ste-btn { font-family: 'Share Tech Mono', monospace; font-size: 11px; font-weight: 700; letter-spacing: 0.05em; padding: 7px 12px; border: 1px solid ${T.border2}; border-radius: 7px; cursor: pointer; background: ${T.bgCard}; color: ${T.textMuted}; transition: all 0.15s; white-space: nowrap; }
    .ste-btn:hover { border-color: ${T.accent}; color: ${T.accent}; }
    .ste-btn.active { background: ${T.border}; border-color: ${T.accent}; color: ${T.accent}; }
    .ste-btn.danger { border-color: ${T.error}55; color: ${T.error}; }
    .ste-btn.danger:hover { background: #2A0E0E; }
    .dim-chip { display: flex; align-items: center; gap: 5px; font-size: 10px; letter-spacing: 0.06em; text-transform: uppercase; color: ${T.textDim}; cursor: pointer; padding: 7px 10px; border: 1px solid ${T.border2}; border-radius: 7px; background: ${T.bgCard}; transition: all 0.15s; user-select: none; }
    .dim-chip:hover { border-color: ${T.accent}; }
    .dim-chip.on { color: #A78BFA; border-color: #3D2A6A; background: #160F2A; }
    .dim-dot { width: 7px; height: 7px; border-radius: 50%; background: currentColor; flex-shrink: 0; }
    .page-btn { background: #160F2A; border: 1px solid #3D2A6A; border-radius: 5px; color: #A78BFA; font-family: 'Share Tech Mono', monospace; font-size: 11px; padding: 4px 9px; cursor: pointer; transition: all 0.12s; }
    .page-btn:hover:not(:disabled) { border-color: #A78BFA; color: #D4BBFF; background: #1E1238; }
    .page-btn:disabled { opacity: 0.25; cursor: not-allowed; }
    .page-btn.cur { background: #7C3AED33; border-color: #A78BFA; color: #D4BBFF; font-weight: 700; }
    .pt-modal-overlay { position: fixed; inset: 0; background: #00000088; z-index: 200; display: flex; align-items: center; justify-content: center; padding: 20px; }
    .pt-modal { background: ${T.bgCard}; border: 1px solid ${T.border2}; border-radius: 14px; max-width: 500px; width: 100%; max-height: 80dvh; overflow-y: auto; box-shadow: 0 20px 60px #00000099; }
    .pt-modal-hdr { padding: 14px 16px; background: ${T.bgDark}; border-bottom: 1px solid ${T.border}; border-radius: 14px 14px 0 0; display: flex; justify-content: space-between; align-items: center; position: sticky; top: 0; }
    .pt-modal-row { display: flex; gap: 10px; padding: 9px 16px; border-bottom: 1px solid ${T.border}22; align-items: flex-start; }
    .pt-modal-row:last-child { border-bottom: none; }
  `;

  return (
    <div style={{ flex: 1, display: "flex", flexDirection: "column", background: T.bg, overflow: "hidden" }}>
      <style>{css}</style>

      {/* Barre supérieure */}
      <div style={{ padding: "5px 12px", background: T.bgDark, borderBottom: `1px solid #7C3AED66`, display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
        <span style={{ color: T.text, fontSize: 15, fontWeight: 800, letterSpacing: "-0.01em", whiteSpace: "nowrap" }}>Pointage</span>
        <PointageInfos />
        <span style={{ color: "#7C3AED99", fontSize: 12, margin: "0 2px" }}>|</span>
        <div style={{ marginLeft: "auto", display: "flex", gap: 8, flexShrink: 0 }}>
          <button className={`ste-btn${openToolbar ? " active" : ""}`} onClick={() => { setOpenToolbar((o) => !o); setOpenMouvements(false); }}>
            {openToolbar ? "✓ " : ""}Outils
          </button>
          <button
            className={`ste-btn${openMouvements ? " active" : ""}`}
            onClick={() => { setOpenMouvements((o) => !o); setOpenToolbar(false); }}
            style={(selectedDest && selectMode === "dest" ? (() => { const dc = destinations.find(d => d.name === selectedDest)?.color; return dc ? { background: `${dc}33`, borderColor: dc, color: dc } : {}; })() : {})}
          >
            {openMouvements ? "✓ " : ""}Mouvement
          </button>
        </div>
      </div>

      {/* Cadre Outils */}
      {openToolbar && (
        <div style={{ margin: "8px 12px 0", padding: "10px 14px", background: T.bgCard, border: `1px solid #7C3AED66`, borderRadius: 10, display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center", maxHeight: "35dvh", overflowY: "auto", flexShrink: 0 }}>
          <div className={`dim-chip${dimRepeated ? " on" : ""}`} onClick={() => setDimRepeated((v) => !v)}><span className="dim-dot" />Rép.</div>
          <div className={`dim-chip${autoRefFmt ? " on" : ""}`} onClick={() => { if (autoRefFmt) { setRefGroupInput(splitFormats["reference"] ?? ""); setRefGroupModal(true); } setAutoRefFmt((v) => !v); }} title="Formatage automatique des références">🔢 Réf.</div>
          <button className="ste-btn" onClick={() => setPoidsUnit((u) => u === "t" ? "kg" : "t")} style={{ color: poidsUnit === "kg" ? T.warning : undefined, borderColor: poidsUnit === "kg" ? `${T.warning}66` : undefined }}>⚖️ {poidsUnit}</button>
          <button className={`ste-btn${selectMode === "col" ? " active" : ""}`} onClick={() => { setSelectMode(selectMode === "col" ? "none" : "col"); setSelectedItems(new Set()); }}>{selectMode === "col" ? "✓ " : ""}Col.</button>
          <button className={`ste-btn${selectMode === "row" ? " active" : ""}`} onClick={() => { setSelectMode(selectMode === "row" ? "none" : "row"); setSelectedItems(new Set()); }}>{selectMode === "row" ? "✓ " : ""}Lig.</button>
          {selectedItems.size > 0 && <button className="ste-btn danger" onClick={selectMode === "col" ? applyColAction : applyRowDeletion}>{selectMode === "col" ? `Masquer ${selectedItems.size} col.` : `Masquer ${selectedItems.size} lig.`}</button>}
          {(hiddenCols.size > 0 || hiddenRows.size > 0) && <button className="ste-btn" onClick={() => { _setHiddenCols(new Set()); _setHiddenRows(new Set()); }}>↺</button>}
          {sortCol !== null && <button className="ste-btn" onClick={() => { setSortCol(null); setPageIdx(0); }}>✕ Tri</button>}
          {Object.values(colFilters).some((v) => v.trim()) && <button className="ste-btn" onClick={() => { setColFilters({}); setPageIdx(0); }}>✕ Filtres</button>}
          <button className="ste-btn" style={{ marginLeft: "auto", fontSize: 11 }} onClick={() => setOpenToolbar(false)}>✕</button>
        </div>
      )}

      {/* Cadre Mouvements */}
      {openMouvements && (
        <div style={{ margin: "8px 12px 0", padding: "12px 14px", background: T.bgCard, border: `1px solid #7C3AED66`, borderRadius: 10, maxHeight: "45dvh", overflowY: "auto", flexShrink: 0 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
            <span style={{ fontSize: 13, fontWeight: 700, color: "#A78BFA", letterSpacing: "0.04em" }}>Mouvements</span>
            <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
              <div
                onClick={() => setDoubleVerif(v => !v)}
                style={{ display: "flex", alignItems: "center", gap: 5, cursor: "pointer", background: doubleVerif ? "#34D39922" : T.bgDark, border: `1px solid ${doubleVerif ? "#34D399" : T.border2}`, borderRadius: 20, padding: "3px 10px 3px 6px", userSelect: "none" }}
              >
                <div style={{ width: 28, height: 16, borderRadius: 8, position: "relative", background: doubleVerif ? "#34D399" : T.border2, transition: "background 0.2s" }}>
                  <div style={{ position: "absolute", top: 2, left: doubleVerif ? 14 : 2, width: 12, height: 12, borderRadius: "50%", background: "#fff", transition: "left 0.2s" }} />
                </div>
                <span style={{ fontSize: 10, color: doubleVerif ? "#34D399" : T.textDim, fontWeight: 600 }}>Double vérif</span>
              </div>
              <button className="ste-btn" onClick={() => setOpenMouvements(false)} style={{ fontSize: 11 }}>✕</button>
            </div>
          </div>

          {/* Pills destinations */}
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 10 }}>
            {destinations.map(d => (
              <button key={d.name} className="ste-btn"
                onClick={() => {
                  if (selectedDest === d.name && selectMode === "dest") { setSelectedDest(""); setSelectMode("none"); }
                  else { setSelectedDest(d.name); setSelectMode("dest"); }
                }}
                style={{ background: selectedDest === d.name && selectMode === "dest" ? d.color : `${d.color}22`, borderColor: d.color, color: selectedDest === d.name && selectMode === "dest" ? "#0F172A" : d.color, fontWeight: selectedDest === d.name ? 700 : 400, minWidth: 36, outline: selectedDest === d.name && selectMode === "dest" ? `2px solid ${d.color}` : undefined }}
              >
                {d.name}
                {destStats.get(d.name) && <span style={{ marginLeft: 4, fontSize: 9, opacity: 0.85 }}>({destStats.get(d.name)!.count})</span>}
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
                    setDestinations(prev => [...prev, { name, color: colors[destinations.length % colors.length] }]);
                  }
                  setNewDestName("");
                }
              }}
              placeholder="Nom destination"
              style={{ flex: 1, background: T.bg, border: `1px solid ${T.border}`, borderRadius: 6, color: T.text, fontFamily: "'Share Tech Mono', monospace", fontSize: 12, padding: "4px 8px", outline: "none" }}
            />
            <button className="ste-btn" style={{ fontWeight: 700 }} onClick={() => {
              const name = newDestName.trim();
              if (!name) return;
              if (!destinations.some(d => d.name.toLowerCase() === name.toLowerCase())) {
                const colors = ["#00c87a","#f447d1","#3cbefc","#ff9b2c","#a78bfa","#f87171","#34d399","#fbbf24"];
                setDestinations(prev => [...prev, { name, color: colors[destinations.length % colors.length] }]);
              }
              setNewDestName("");
            }}>+</button>
            {destinations.length > 0 && (
              <button className="ste-btn danger" title="Supprimer la destination sélectionnée" onClick={() => {
                if (!selectedDest) return;
                setDestinations(prev => prev.filter(d => d.name !== selectedDest));
                setRowDestinations(prev => { const next = new Map(prev); for (const [k, v] of next) { if (v === selectedDest) next.delete(k); } return next; });
                setSelectedDest("");
              }} disabled={!selectedDest}>✕ Dest.</button>
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
                  const fmtW = (w: number) => { const val = poidsUnit === "kg" ? (w * 1000).toFixed(0) : parseFloat(w.toFixed(3)).toString(); return poidsFmt ? applyGrouping(val, poidsFmt) : thsep(val); };
                  return (
                    <tr key={d.name}>
                      <td style={{ padding: "2px 6px 2px 0" }}><span style={{ background: `${d.color}33`, borderLeft: `3px solid ${d.color}`, padding: "1px 6px", borderRadius: 3, color: d.color, fontWeight: 600 }}>{d.name}</span></td>
                      <td style={{ textAlign: "right", padding: "2px 6px", color: T.text }}>{s.count}</td>
                      {poidsColIdx >= 0 && <td style={{ textAlign: "right", padding: "2px 0 2px 6px", color: T.warning }}>{fmtW(s.weight)}</td>}
                    </tr>
                  );
                })}
                {destStats.size > 1 && (() => {
                  const poidsFmt = splitFormats["poids"] ?? "";
                  const totalW = [...destStats.values()].reduce((a, b) => a + b.weight, 0);
                  const fmtTotal = () => { const val = poidsUnit === "kg" ? (totalW * 1000).toFixed(0) : parseFloat(totalW.toFixed(3)).toString(); return poidsFmt ? applyGrouping(val, poidsFmt) : thsep(val); };
                  return (
                    <tr style={{ borderTop: `1px solid ${T.border}` }}>
                      <td style={{ padding: "2px 6px 2px 0", color: T.textDim, fontStyle: "italic" }}>Total</td>
                      <td style={{ textAlign: "right", padding: "2px 6px", color: T.accent, fontWeight: 700 }}>{[...destStats.values()].reduce((a, b) => a + b.count, 0)}</td>
                      {poidsColIdx >= 0 && <td style={{ textAlign: "right", padding: "2px 0 2px 6px", color: T.accent, fontWeight: 700 }}>{fmtTotal()}</td>}
                    </tr>
                  );
                })()}
              </tbody>
            </table>
          )}

          {rowDestinations.size > 0 && (
            <button className="ste-btn danger" style={{ marginTop: 10, fontSize: 11 }} onClick={() => setConfirmClearDest(true)}>✖ Effacer toutes les affectations</button>
          )}
          {confirmClearDest && (
            <div onClick={() => setConfirmClearDest(false)} style={{ position: "fixed", inset: 0, background: "#000000aa", zIndex: 500, display: "flex", alignItems: "center", justifyContent: "center" }}>
              <div onClick={(e) => e.stopPropagation()} style={{ background: T.bgCard, border: `1px solid ${T.error}66`, borderRadius: 14, padding: "24px 20px", maxWidth: 320, width: "90%", boxShadow: "0 20px 60px #00000099" }}>
                <div style={{ color: T.error, fontWeight: 800, fontSize: 15, marginBottom: 10 }}>⚠️ Confirmer la suppression</div>
                <div style={{ color: T.textMuted, fontSize: 13, marginBottom: 20 }}>Toutes les affectations ({rowDestinations.size} ligne{rowDestinations.size > 1 ? "s" : ""}) seront effacées. Cette action est irréversible.</div>
                <div style={{ display: "flex", gap: 10 }}>
                  <button className="ste-btn" style={{ flex: 1 }} onClick={() => setConfirmClearDest(false)}>Annuler</button>
                  <button className="ste-btn danger" style={{ flex: 1 }} onClick={() => { setRowDestinations(new Map()); setReassignedRows(new Map()); setConfirmClearDest(false); }}>✖ Effacer</button>
                </div>
              </div>
            </div>
          )}
        </div>
      )}

      {/* Status hints */}
      <div style={{ padding: "6px 12px 0", display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
        {selectMode === "col" && <div style={{ flex: "1 1 100%", marginBottom: 4, padding: "6px 10px", background: "#1F0A0A", border: `1px solid ${T.error}55`, borderRadius: 6, fontSize: 11, color: T.error }}>→ Cliquez les en-têtes à masquer, puis confirmez</div>}
        {selectMode === "row" && <div style={{ flex: "1 1 100%", marginBottom: 4, padding: "6px 10px", background: "#1F0A0A", border: `1px solid ${T.error}55`, borderRadius: 6, fontSize: 11, color: T.error }}>→ Cliquez les numéros de lignes à masquer, puis confirmez</div>}
        {addedRows.length > 0 && <span style={{ color: T.success, fontSize: 10 }}>+ {addedRows.length} ajoutée(s)</span>}
        {pointedRows.size > 0 && <span style={{ color: T.success, fontSize: 10, fontWeight: 700 }}>✅ {pointedRows.size} pointée(s)</span>}
      </div>

      {/* Table */}
      <div style={{ flex: 1, overflowX: "auto", overflowY: "auto", margin: "8px 12px 0", border: `1px solid ${T.border}`, borderRadius: 8, background: T.bgCard, WebkitOverflowScrolling: "touch", minHeight: 0 }}>
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
                        <input className="header-input" value={headers[ci]} autoFocus
                          onChange={(e) => { const next = [...headers]; next[ci] = e.target.value; setHeaders(next); }}
                          onBlur={() => setEditingHeader(null)}
                          onKeyDown={(e) => { if (e.key === "Enter" || e.key === "Escape") setEditingHeader(null); }}
                          onClick={(e) => e.stopPropagation()}
                        />
                      ) : (
                        <span onDoubleClick={() => selectMode === "none" && setEditingHeader(ci)} style={{ flex: 1 }}>{colLabel(ci)}</span>
                      )}
                      {isSortedCol && <span style={{ fontSize: 9, flexShrink: 0 }}>{sortDir === "asc" ? "▲" : "▼"}</span>}
                      {!isSortedCol && <span style={{ fontSize: 8, flexShrink: 0, opacity: 0.2 }}>⇅</span>}
                    </div>
                    <div className="th-filter">
                      <input value={colFilters[ci] ?? ""} onChange={(e) => { setColFilters((p) => ({ ...p, [ci]: e.target.value })); setPageIdx(0); }} placeholder="…" onClick={(e) => e.stopPropagation()} />
                    </div>
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {pageRows.map(({ row, ri }) => {
              const isRowSel  = selectMode === "row" && selectedItems.has(ri);
              const isAdded   = ri < addedRows.length;
              const isPointed = pointedRows.has(ri);
              const rowDest   = rowDestinations.get(ri);
              const destColor = rowDest ? (destinations.find(d => d.name === rowDest)?.color ?? null) : null;
              const rowStyle: React.CSSProperties = isPointed
                ? { background: "#0A1F10", outline: `1px solid ${T.success}44`, ...(destColor ? { borderLeft: `4px solid ${destColor}` } : {}) }
                : anomalyRowSet.has(ri)
                  ? { background: "#1E0A0A", outline: `1px solid ${T.error}33`, ...(destColor ? { borderLeft: `4px solid ${destColor}` } : {}) }
                  : destColor ? { borderLeft: `4px solid ${destColor}`, background: `${destColor}33` } : {};
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
                      if (!selectedDest) { setNoDestWarning(true); return; }
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
                              const genFake = (): number => { const pct = 0.02 + Math.random() * 0.07; const sign = Math.random() > 0.5 ? 1 : -1; return parseFloat((realW * (1 + sign * pct)).toFixed(realW < 10 ? 3 : 1)); };
                              let f1 = genFake(), f2 = genFake(), guard = 0;
                              while (guard++ < 20 && Math.abs(f1 - realW) < realW * 0.005) f1 = genFake();
                              guard = 0;
                              while (guard++ < 20 && (Math.abs(f2 - realW) < realW * 0.005 || Math.abs(f2 - f1) < realW * 0.005)) f2 = genFake();
                              const choices = [realW, f1, f2].sort(() => Math.random() - 0.5);
                              setDoubleVerifModal({ ri, ref: refVal2, correct: realW, choices, error: false, pendingDest: selectedDest });
                              return;
                            }
                          }
                        } else {
                          const refIdx2   = mapping.reference ? headers.indexOf(mapping.reference) : -1;
                          const rangIdx2  = mapping.rang      ? headers.indexOf(mapping.rang)      : -1;
                          const poidsIdx2 = mapping.poids     ? headers.indexOf(mapping.poids)     : -1;
                          const getCell2  = (ci: number) => rowOverrides.get(ri)?.[ci] !== undefined ? String(rowOverrides.get(ri)![ci] ?? "") : String(row[ci] ?? "");
                          const dc = destinations.find(d => d.name === selectedDest)?.color ?? T.accent;
                          setQuickConfirmModal({ ri, ref: refIdx2 >= 0 ? getCell2(refIdx2) : "—", rang: rangIdx2 >= 0 ? getCell2(rangIdx2) : "—", poids: poidsIdx2 >= 0 ? getCell2(poidsIdx2) : "—", dest: selectedDest, destColor: dc });
                          return;
                        }
                        setRowDestinations((prev) => { const next = new Map(prev); if (next.get(ri) === selectedDest) next.delete(ri); else next.set(ri, selectedDest); return next; });
                      }
                      return;
                    }
                  }}
                >
                  <td className="td-num" onClick={openToolbar ? (e) => { e.stopPropagation(); if (selectMode !== "row") { setSelectMode("row"); setSelectedItems(new Set([ri])); } else { toggleSelectItem(ri); } } : undefined} style={openToolbar ? { cursor: "pointer" } : undefined}>
                    {isPointed ? <span style={{ color: T.success, fontSize: 11 }}>✅</span> : isAdded ? <span style={{ color: T.success, fontSize: 9 }}>+{ri + 1}</span> : ri + 1}
                  </td>
                  {visibleCols.map((ci) => {
                    const rawCell = rowOverrides.get(ri)?.[ci] !== undefined ? rowOverrides.get(ri)![ci] : (row[ci] ?? null);
                    const strVal = rawCell !== null ? String(rawCell) : null;
                    const fmt    = getColFmt(ci);
                    let display: string | null;
                    if (strVal === null) { display = null; }
                    else if (ci === poidsColIdx) { const raw = parseFloat(strVal) || 0; const val = poidsUnit === "kg" ? (raw * 1000).toFixed(0) : parseFloat(raw.toFixed(3)).toString(); display = (fmt ? applyGrouping(val, fmt) : thsep(val)) + " " + (poidsUnit === "kg" ? "kg" : "t"); }
                    else if (ci === refColIdx && autoRefFmt) { display = autoFormatRef(strVal, fmt); }
                    else { display = fmt ? applyGrouping(strVal, fmt) : strVal; }
                    const isRep  = !isAdded && dimRepeated && strVal !== null && (repetitiveByCol.get(ci)?.has(strVal) ?? false);
                    const isDestCol = ci === destIdx;
                    const destCellVal = isDestCol && rowDest ? rowDest : null;
                    const isRefCol = ci === refColIdx;
                    return (
                      <td key={ci} title={destCellVal ?? strVal ?? ""} className={isRep ? "cell-rep" : ""}
                        onPointerDown={isRefCol ? (e) => { e.stopPropagation(); longPressRef.current = setTimeout(() => { setRefGroupInput(splitFormats["reference"] ?? ""); setRefGroupModal(true); }, 500); } : undefined}
                        onPointerUp={isRefCol ? () => { if (longPressRef.current) { clearTimeout(longPressRef.current); longPressRef.current = null; } } : undefined}
                        onPointerLeave={isRefCol ? () => { if (longPressRef.current) { clearTimeout(longPressRef.current); longPressRef.current = null; } } : undefined}
                      >
                        {isDestCol ? destCellVal ? <span style={{ color: destColor ?? T.accent, fontWeight: 700, fontSize: 11 }}>{destCellVal}</span> : <span style={{ color: T.textDim, fontSize: 11, opacity: 0.5 }}>STOCK</span>
                          : display ?? <span style={{ color: T.textDim }}>—</span>}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
            {pageRows.length === 0 && (
              <tr><td colSpan={visibleCols.length + 1} style={{ textAlign: "center", padding: "32px", color: T.textDim }}>Aucune ligne correspondante</td></tr>
            )}
          </tbody>
        </table>
      </div>

      {/* Pagination */}
      <div style={{ margin: "8px 12px 4px", display: "flex", gap: 5, flexWrap: "wrap", alignItems: "center", flexShrink: 0 }}>
        <button className="page-btn" onClick={() => setPageIdx(0)} disabled={safePage === 0}>«</button>
        <button className="page-btn" onClick={() => setPageIdx((p) => Math.max(0, p - 1))} disabled={safePage === 0}>‹</button>
        {Array.from({ length: totalPages }, (_, i) => i).filter((i) => Math.abs(i - safePage) <= 2 || i === 0 || i === totalPages - 1).reduce<(number | "…")[]>((acc, i, idx, arr) => { if (idx > 0 && i - (arr[idx - 1] as number) > 1) acc.push("…"); acc.push(i); return acc; }, []).map((item, idx) =>
          item === "…" ? <span key={`e${idx}`} style={{ color: "#7C3AED66", fontSize: 11 }}>…</span>
            : <button key={item} className={`page-btn${item === safePage ? " cur" : ""}`} onClick={() => setPageIdx(item as number)}>{(item as number) + 1}</button>
        )}
        <button className="page-btn" onClick={() => setPageIdx((p) => Math.min(totalPages - 1, p + 1))} disabled={safePage >= totalPages - 1}>›</button>
        <button className="page-btn" onClick={() => setPageIdx(totalPages - 1)} disabled={safePage >= totalPages - 1}>»</button>
        <span style={{ color: "#7C3AED88", fontSize: 10, margin: "0 4px" }}>|</span>
        <select value={pageSize} onChange={(e) => { setPageSize(Number(e.target.value)); setPageIdx(0); }} style={{ background: "#160F2A", border: "1px solid #3D2A6A", borderRadius: 5, color: "#A78BFA", fontFamily: "'Share Tech Mono', monospace", fontSize: 11, padding: "3px 6px", cursor: "pointer" }}>
          {PAGE_SIZES.map((s) => <option key={s} value={s}>{s} / page</option>)}
        </select>
        <span style={{ color: "#7C3AED99", fontSize: 10 }}>{safePage * pageSize + 1}–{Math.min((safePage + 1) * pageSize, sortedRows.length)} / {sortedRows.length}</span>
      </div>

      {showAddRow && <AddRowModal onClose={() => setShowAddRow(false)} />}

      {/* Modals */}
      {confirmUnpoint && (
        <div onClick={() => setConfirmUnpoint(null)} style={{ position: "fixed", inset: 0, background: "#000000aa", zIndex: 500, display: "flex", alignItems: "center", justifyContent: "center" }}>
          <div onClick={(e) => e.stopPropagation()} style={{ background: T.bgCard, border: `1px solid ${T.error}66`, borderRadius: 14, padding: "22px 20px", maxWidth: 300, width: "90%", boxShadow: "0 20px 60px #00000099" }}>
            <div style={{ color: T.error, fontWeight: 800, fontSize: 14, marginBottom: 10 }}>❌ Dépointer la ligne ?</div>
            <div style={{ color: T.textMuted, fontSize: 12, marginBottom: 18 }}>Retirer le pointage de <strong style={{ color: T.accent, fontFamily: "monospace" }}>{confirmUnpoint.ref}</strong> ?</div>
            <div style={{ display: "flex", gap: 10 }}>
              <button className="ste-btn" style={{ flex: 1 }} onClick={() => setConfirmUnpoint(null)}>Annuler</button>
              <button className="ste-btn danger" style={{ flex: 1 }} onClick={() => { setPointedRows(prev => { const next = new Set(prev); next.delete(confirmUnpoint!.ri); return next; }); setConfirmUnpoint(null); setModalRow(null); }}>Dépointer</button>
            </div>
          </div>
        </div>
      )}

      {confirmReassign && (
        <div onClick={() => setConfirmReassign(null)} style={{ position: "fixed", inset: 0, background: "#000000aa", zIndex: 500, display: "flex", alignItems: "center", justifyContent: "center" }}>
          <div onClick={(e) => e.stopPropagation()} style={{ background: T.bgCard, border: `1px solid ${T.warning}66`, borderRadius: 14, padding: "22px 20px", maxWidth: 300, width: "90%", boxShadow: "0 20px 60px #00000099" }}>
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
                  const next = new Map(prev); const history = next.get(confirmReassign.ri) ?? [];
                  next.set(confirmReassign.ri, [...history, { from: confirmReassign.from, to: confirmReassign.to }]); return next;
                });
                setConfirmReassign(null);
              }}>Réaffecter</button>
            </div>
          </div>
        </div>
      )}

      {/* Row detail modal */}
      {modalRow && (() => {
        const { row, rowNum, ri } = modalRow;
        const refIdx   = mapping.reference ? headers.indexOf(mapping.reference) : -1;
        const getCell  = (ci: number) => rowOverrides.get(ri)?.[ci] !== undefined ? String(rowOverrides.get(ri)![ci] ?? "") : String(row[ci] ?? "");
        const refVal   = refIdx >= 0 ? getCell(refIdx) : "—";
        return (
          <div className="pt-modal-overlay" onClick={() => setModalRow(null)}>
            <div className="pt-modal" onClick={(e) => e.stopPropagation()}>
              <div className="pt-modal-hdr">
                <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                  <span style={{ color: T.textDim, fontSize: 9, letterSpacing: "0.12em", textTransform: "uppercase" }}>Ligne #{rowNum}</span>
                  <span style={{ color: T.accent, fontWeight: 900, fontSize: 15, fontFamily: "'Share Tech Mono', monospace" }}>🏷 {refVal}</span>
                </div>
                <button onClick={() => setModalRow(null)} style={{ background: "none", border: "none", color: T.textDim, fontSize: 22, cursor: "pointer" }}>×</button>
              </div>
              <div style={{ maxHeight: 260, overflowY: "auto" }}>
                {visibleCols.map((ci) => {
                  const val = row[ci]; const strVal = val !== null && val !== undefined ? String(val) : null;
                  const fmt = getColFmt(ci); const display = strVal !== null ? (fmt ? applyGrouping(strVal, fmt) : strVal) : null;
                  return (
                    <div key={ci} className="pt-modal-row">
                      <span style={{ minWidth: 120, color: T.textDim, fontSize: 11, textTransform: "uppercase", letterSpacing: "0.06em", flexShrink: 0, paddingTop: 2 }}>{colLabel(ci)}</span>
                      <span style={{ color: display ? T.text : T.textDim, fontSize: 13, fontFamily: "monospace", wordBreak: "break-all" }}>{display ?? "—"}</span>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        );
      })()}

      {/* Quick Confirm modal */}
      {quickConfirmModal && (() => {
        const { ref, rang, poids, dest, destColor } = quickConfirmModal;
        const pct = (quickConfirmCountdown / 5) * 100;
        const doConfirm = () => { setRowDestinations(prev => { const next = new Map(prev); next.set(quickConfirmModal.ri, dest); return next; }); setQuickConfirmModal(null); };
        const doCancel = () => setQuickConfirmModal(null);
        return (
          <div className="pt-modal-overlay" style={{ zIndex: 1100 }} onClick={doCancel}>
            <div className="pt-modal" onClick={e => e.stopPropagation()} style={{ maxWidth: 320 }}>
              <div className="pt-modal-hdr">
                <div><span style={{ color: T.textDim, fontSize: 9, letterSpacing: "0.12em", textTransform: "uppercase" }}>Vérification</span><br /><span style={{ color: T.accent, fontWeight: 900, fontSize: 15, fontFamily: "'Share Tech Mono', monospace" }}>🏷 {ref}</span></div>
                <button onClick={doCancel} style={{ background: "none", border: "none", color: T.textDim, fontSize: 22, cursor: "pointer" }}>×</button>
              </div>
              <div style={{ padding: "14px 16px" }}>
                {[{ label: "Rang", value: rang }, { label: "Référence", value: ref }, { label: "Poids", value: poids }].map(({ label, value }) => (
                  <div key={label} style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", borderBottom: `1px solid ${T.border}33` }}>
                    <span style={{ color: T.textDim, fontSize: 12 }}>{label}</span>
                    <span style={{ color: T.text, fontWeight: 700, fontSize: 13, fontFamily: "monospace" }}>{value}</span>
                  </div>
                ))}
                <div style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", marginBottom: 16, borderBottom: `1px solid ${T.border}33` }}>
                  <span style={{ color: T.textDim, fontSize: 12 }}>Destination</span>
                  <span style={{ color: destColor, fontWeight: 800, fontSize: 13, fontFamily: "monospace" }}>{dest}</span>
                </div>
                <div style={{ marginBottom: 14 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                    <span style={{ color: T.textDim, fontSize: 10 }}>Annulation automatique dans…</span>
                    <span style={{ color: T.warning, fontWeight: 800, fontSize: 14, fontFamily: "monospace" }}>{quickConfirmCountdown}s</span>
                  </div>
                  <div style={{ height: 6, borderRadius: 99, background: T.border, overflow: "hidden" }}>
                    <div style={{ height: "100%", borderRadius: 99, background: T.warning, width: `${pct}%`, transition: "width 1s linear" }} />
                  </div>
                </div>
                <div style={{ display: "flex", gap: 10 }}>
                  <button className="ste-btn" style={{ flex: 1 }} onClick={doCancel}>✕ Annuler</button>
                  <button className="ste-btn" style={{ flex: 1, background: `${destColor}33`, borderColor: destColor, color: destColor, fontWeight: 800 }} onClick={doConfirm}>✓ Confirmer</button>
                </div>
              </div>
            </div>
          </div>
        );
      })()}

      {/* Double Vérif modal */}
      {doubleVerifModal && (() => {
        const { ref, correct, choices, error } = doubleVerifModal;
        const poidsFmt = splitFormats["poids"] ?? ""; const unitLabel = poidsUnit === "kg" ? "kg" : "t";
        const fmtChoice = (w: number) => { const val = poidsUnit === "kg" ? (w * 1000).toFixed(0) : parseFloat(w.toFixed(3)).toString(); return (poidsFmt ? applyGrouping(val, poidsFmt) : thsep(val)) + " " + unitLabel; };
        const handleChoice = (w: number) => {
          if (error) return;
          const isCorrect = Math.abs(w - correct) < 1e-9;
          if (isCorrect) {
            if (doubleVerifModal!.pendingDest) { setRowDestinations(prev => { const next = new Map(prev); next.set(doubleVerifModal!.ri, doubleVerifModal!.pendingDest!); return next; }); }
            else { setPointedRows(prev => { const next = new Set(prev); next.add(doubleVerifModal!.ri); return next; }); }
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
                <div><span style={{ color: T.textDim, fontSize: 9, letterSpacing: "0.12em", textTransform: "uppercase" }}>Double vérification</span><br /><span style={{ color: T.accent, fontWeight: 900, fontSize: 15, fontFamily: "'Share Tech Mono', monospace" }}>🏷 {ref}</span></div>
                {!error && <button onClick={() => setDoubleVerifModal(null)} style={{ background: "none", border: "none", color: T.textDim, fontSize: 22, cursor: "pointer" }}>×</button>}
              </div>
              {error ? (
                <div style={{ padding: "24px 16px", textAlign: "center" }}>
                  <div style={{ fontSize: 36, marginBottom: 12 }}>❌</div>
                  <div style={{ color: T.error, fontWeight: 800, fontSize: 14, marginBottom: 6 }}>Poids incorrect !</div>
                  <div style={{ color: T.textDim, fontSize: 11 }}>Fermeture automatique dans 5 secondes…</div>
                </div>
              ) : (
                <div style={{ padding: "16px" }}>
                  <div style={{ color: T.textDim, fontSize: 11, marginBottom: 14, textAlign: "center" }}>Sélectionnez le <strong style={{ color: T.text }}>poids correct</strong> pour valider le pointage</div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                    {choices.map((w, i) => (
                      <button key={i} onClick={() => handleChoice(w)} style={{ padding: "12px 16px", borderRadius: 10, cursor: "pointer", background: T.bgDark, border: `2px solid ${T.border2}`, color: T.warning, fontFamily: "'Share Tech Mono', monospace", fontSize: 16, fontWeight: 700, textAlign: "center", transition: "all 0.1s" }}
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

      {noDestWarning && (
        <div style={{ position: "fixed", inset: 0, background: "#000a", zIndex: 2000, display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }} onClick={() => setNoDestWarning(false)}>
          <div style={{ background: T.bgCard, borderRadius: 16, padding: "24px 20px", maxWidth: 320, width: "100%", textAlign: "center" }} onClick={e => e.stopPropagation()}>
            <div style={{ fontSize: 36, marginBottom: 10 }}>⚠️</div>
            <div style={{ color: T.text, fontWeight: 700, fontSize: 15, marginBottom: 8 }}>Aucune destination sélectionnée</div>
            <div style={{ color: T.textDim, fontSize: 12, marginBottom: 20 }}>Sélectionnez une destination dans le panneau Mouvements avant d'affecter des lignes.</div>
            <button className="ste-btn" style={{ width: "100%", background: T.accent, color: "#0F172A", fontWeight: 700 }} onClick={() => setNoDestWarning(false)}>OK</button>
          </div>
        </div>
      )}

      {refGroupModal && (
        <div style={{ position: "fixed", inset: 0, background: "#000a", zIndex: 2000, display: "flex", flexDirection: "column", justifyContent: "flex-end" }} onClick={() => setRefGroupModal(false)}>
          <div style={{ background: T.bgCard, borderRadius: "18px 18px 0 0", padding: "20px 16px 32px" }} onClick={e => e.stopPropagation()}>
            <div style={{ color: T.text, fontWeight: 700, fontSize: 14, marginBottom: 4 }}>🏷 Format de groupement — REF</div>
            <div style={{ color: T.textDim, fontSize: 11, marginBottom: 12 }}>Séparateur visuel (ex: "3 2 3" → ABC DE FGH)</div>
            <input
              style={{ width: "100%", background: T.bgDark, border: `1px solid ${T.border2}`, borderRadius: 8, padding: "10px 12px", color: T.text, fontSize: 13, fontFamily: "monospace", boxSizing: "border-box" }}
              value={refGroupInput} onChange={e => setRefGroupInput(e.target.value)} placeholder="ex: 3 2 3" autoFocus
              onKeyDown={e => { if (e.key === "Enter") { setSplitFormats(p => ({ ...p, reference: refGroupInput.trim() })); setRefGroupModal(false); } if (e.key === "Escape") setRefGroupModal(false); }}
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
