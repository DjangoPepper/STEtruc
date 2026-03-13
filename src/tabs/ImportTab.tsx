// ============================================================
// STEtruc — Onglet Import
// ============================================================

import { useState, useCallback, useRef, useMemo, useEffect } from "react";
// import { T, RawData, CellValue, ParsedData, applyGrouping, thsep, autoFormatRef } from "../types";
import { T, RawData, CellValue, applyGrouping, thsep, autoFormatRef } from "../types";
import { useApp } from "../AppContext";
import { Btn } from "../components";

export default function ImportPage() {
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
  const [splitEditingField, setSplitEditingField] = useState<string | null>(null);
  const [splitInputValue, setSplitInputValue] = useState("");
  const [openOnglets,   setOpenOnglets]   = useState(false);
  const [openAtypiques, setOpenAtypiques] = useState(false);
  const [openHeaders,   setOpenHeaders]   = useState(false);
  const [openDonnees,   setOpenDonnees]   = useState(false);
  const [openMapping,   setOpenMapping]   = useState(false);
  const [openApercu,    setOpenApercu]    = useState(false);
  const suppressCollapseRef = useRef(false);

  useEffect(() => {
    if (!winwinModalOpen) return;
    const t = setTimeout(() => setWinwinModalOpen(false), 3000);
    return () => clearTimeout(t);
  }, [winwinModalOpen]);

  // eslint-disable-next-line react-hooks/exhaustive-deps
  useEffect(() => {
    if (!parsed || headers.length === 0) return;
    const allRows = parsed.rows.map((row) => {
      const obj: Record<string, string> = {};
      headers.forEach((h, i) => { obj[h] = row[i] !== null && row[i] !== undefined ? String(row[i]) : ""; });
      return obj;
    });
    setEditableRows(allRows);
    setSplitFormats({});
    setMapping({
      rang:      headers.find((k) => /rang|n°|row|id|line|ligne/i.test(k)) ?? "",
      reference: headers.find((k) => /ref|coil|serial|coils|brames|bobine/i.test(k)) ?? headers[0] ?? "",
      poids:     headers.find((k) => /poids|weight|kg|ton(ne)|masse|bruto/i.test(k)) ?? "",
      dch:       headers.find((k) => /dch|d[eé]chargement|dest(ination)?/i.test(k)) ?? "",
    });
    setExtras([]);
    setStep((s) => (s === 1 ? 2 : s));
    if (!suppressCollapseRef.current) {
      setOpenOnglets(false); setOpenAtypiques(false); setOpenHeaders(false);
      setOpenDonnees(false); setOpenMapping(false); setOpenApercu(false);
    }
    suppressCollapseRef.current = false;
  }, [parsed, activeSheet]);

  const onFile = useCallback((file: File) => { handleFile(file); }, [handleFile]);
  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) onFile(file);
  }, [onFile]);

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
      const next = { ...prev }; next[trimmed] = prev[oldName]; delete next[oldName]; return next;
    });
    setMapping((m) => {
      const updated = { ...m };
      (Object.keys(updated) as (keyof typeof updated)[]).forEach((k) => { if (updated[k] === oldName) updated[k] = trimmed; });
      return updated;
    });
    setExtras((exs) => exs.map((e) => e.col === oldName ? { ...e, col: trimmed } : e));
    setEditingHdr(null);
  }, [headers, setHeaders]);

  const deleteColumn = useCallback((idx: number) => {
    setHiddenCols((prev) => new Set([...prev, idx]));
  }, [setHiddenCols]);

  const promoteHeaderToRow = useCallback(() => {
    if (!parsed) return;
    const newHeaders = headers.map((_, i) => `Col${i + 1}`);
    const headerAsRow: CellValue[] = headers.map((h) => h);
    setParsed({ ...parsed, headers: newHeaders, rows: [headerAsRow, ...parsed.rows], headerRowIndex: 0 });
    setHeaders(newHeaders);
    setSplitFormats({});
    setEditingHdr(null);
  }, [parsed, headers, setParsed, setHeaders]);

  const visibleCols = useMemo(
    () => headers.map((h, i) => ({ h, i })).filter(({ i }) => !hiddenCols.has(i)),
    [headers, hiddenCols]
  );

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
          style={{ position: "fixed", inset: 0, background: "#000000bb", zIndex: 500, display: "flex", alignItems: "center", justifyContent: "center" }}
        >
          <div style={{ borderRadius: 18, overflow: "hidden", boxShadow: "0 24px 80px #000000cc", maxWidth: 320, width: "90%" }}>
            <img src="/STEtruc/P1060411.JPG" alt="winwin" style={{ width: "100%", display: "block" }} />
          </div>
        </div>
      )}

      {/* Header */}
      <div style={{ padding: "16px 16px 12px", background: T.bgDark, borderBottom: `1px solid ${T.border}` }}>
        <div style={{ color: T.textDim, fontSize: 9, letterSpacing: "0.2em", textTransform: "uppercase", marginBottom: 4 }}>STEtruc</div>
        <h1 style={{ color: T.text, fontSize: 20, fontWeight: 700 }}>Import</h1>
        <div style={{ color: T.textMuted, fontSize: 11, marginTop: 2 }}>Chargez et configurez votre fichier Excel</div>
      </div>

      {/* Step indicator */}
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
              }}
            >
              <div style={{ color: step > i ? T.success : step === i + 1 ? T.accent : T.textDim, fontWeight: 800, fontSize: 12 }}>
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

        {/* ── STEP 1 ── */}
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
              <div style={{ fontSize: 96, marginBottom: 8 }}>📊</div>
              <div style={{ color: T.accent, fontWeight: 800, fontSize: 16, marginBottom: 4 }}>Charger fichier Excel</div>
              <div style={{ color: T.textDim, fontSize: 12 }}>.xlsx / .xls — colonnes et onglets vides auto-épurés</div>
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

        {/* ── STEP 2 ── */}
        {step === 2 && parsed && (
          <div>
            {/* Sheet selector */}
            {sheetNames.length > 1 && (
              <div style={{ marginBottom: 14, background: T.bgCard, borderRadius: 12, padding: 14, border: `1px solid ${T.accentDim}` }}>
                <div
                  onClick={() => setOpenOnglets((o) => !o)}
                  style={{ color: T.accent, fontWeight: 700, fontSize: 12, marginBottom: openOnglets ? 8 : 0, textTransform: "uppercase", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center" }}
                >
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
                              background: isActive ? T.accent : T.bgDark,
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

            {/* Lignes atypiques */}
            {(() => {
              const atypical = editableRows.map((row, ri) => ({ row, ri })).filter(({ row }) => anomalyInfo.isAnomalous(row));
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
                    <span style={{ color: T.warning, fontWeight: 800, fontSize: 12, textTransform: "uppercase" }}>⚠ Lignes atypiques ({atypical.length})</span>
                    <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                      <button
                        onClick={() => setEditableRows((prev) => prev.filter((row) => !anomalyInfo.isAnomalous(row)))}
                        style={{ background: `${T.error}22`, border: `1px solid ${T.error}55`, borderRadius: 6, color: T.error, fontSize: 10, fontWeight: 700, cursor: "pointer", padding: "3px 8px", whiteSpace: "nowrap" }}
                      >✕ Supprimer tout</button>
                      <span onClick={() => setOpenAtypiques((o) => !o)} style={{ opacity: 0.5, fontSize: 10, cursor: "pointer", color: T.warning, padding: "0 2px" }}>{openAtypiques ? "▲" : "▼"}</span>
                    </div>
                  </div>
                  {openAtypiques && <div style={{ maxHeight: 220, overflowY: "auto" }}>
                    {atypical.map(({ row, ri }) => (
                      <div key={ri} style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 14px", borderBottom: `1px solid ${T.border}22`, background: "#5B21B611" }}>
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
                          {visibleCols.length > 6 && <span style={{ color: T.textDim, fontSize: 10 }}>+{visibleCols.length - 6} col.</span>}
                        </div>
                        <button
                          onClick={() => setEditableRows((prev) => prev.filter((_, idx) => idx !== ri))}
                          style={{ background: "none", border: "none", color: T.error, cursor: "pointer", fontSize: 16, padding: 0, flexShrink: 0 }}
                        >✕</button>
                      </div>
                    ))}
                  </div>}
                </div>
              );
            })()}

            {/* Headers editor */}
            <div style={{ marginBottom: 14, background: "#2A1020", borderRadius: 12, border: `1px solid ${T.border2}`, overflow: "hidden" }}>
              <div style={{ padding: "10px 14px", background: "#1C0A16", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                <span style={{ color: T.textMuted, fontWeight: 800, fontSize: 12, textTransform: "uppercase" }}>
                  📋 Headers · {visibleCols.length}/{headers.length} colonnes
                </span>
                <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" }}>
                  <button
                    onClick={promoteHeaderToRow}
                    style={{ background: `${T.warning}22`, border: `1px solid ${T.warning}55`, borderRadius: 6, color: T.warning, fontSize: 10, fontWeight: 700, cursor: "pointer", padding: "3px 8px", whiteSpace: "nowrap" }}
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
                    style={{ background: `${T.accent}22`, border: `1px solid ${T.accent}55`, borderRadius: 6, color: T.accent, fontSize: 10, fontWeight: 700, cursor: "pointer", padding: "3px 8px", whiteSpace: "nowrap" }}
                  >+ Col.</button>
                  {hiddenCols.size > 0 && (
                    <button
                      onClick={() => setHiddenCols(new Set())}
                      style={{ background: `${T.success}22`, border: `1px solid ${T.success}44`, borderRadius: 6, color: T.success, fontSize: 10, fontWeight: 700, cursor: "pointer", padding: "3px 8px" }}
                    >↺ Restaurer ({hiddenCols.size})</button>
                  )}
                  <span
                    onClick={() => { setOpenHeaders((o) => !o); setOpenDonnees((o) => !o); }}
                    style={{ opacity: 0.5, fontSize: 10, cursor: "pointer", color: T.textMuted, padding: "0 2px" }}
                  >{openHeaders ? "▲" : "▼"}</span>
                </div>
              </div>
              {openHeaders && (
                <div style={{ padding: "10px 14px", display: "flex", gap: 6, flexWrap: "wrap" }}>
                  {headers.map((h, i) => {
                    if (hiddenCols.has(i)) return null;
                    return (
                      <span key={i} style={{
                        background: T.bgDark, color: T.text, borderRadius: 6, fontSize: 11,
                        padding: "4px 6px 4px 10px", fontFamily: "'Share Tech Mono', monospace",
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
                            style={{ background: "transparent", border: "none", outline: "none", color: T.accent, fontFamily: "'Share Tech Mono', monospace", fontSize: 11, width: Math.max(60, h.length * 8) }}
                          />
                        ) : (
                          <span onDoubleClick={() => setEditingHdr(i)} style={{ cursor: "text" }}>{h}</span>
                        )}
                      </span>
                    );
                  })}
                </div>
              )}
            </div>

            {/* Editable data preview */}
            {editableRows.length > 0 && (
              <div style={{ marginBottom: 14, background: T.bgDark, borderRadius: 12, border: `1px solid ${T.border2}`, overflow: "hidden" }}>
                <div
                  onClick={() => setOpenDonnees((o) => !o)}
                  style={{ padding: "10px 14px", background: T.bgCard, display: "flex", alignItems: "center", gap: 8, cursor: "pointer" }}
                >
                  <span style={{ color: T.accent, fontWeight: 800, fontSize: 12, textTransform: "uppercase", flex: 1 }}>
                    ✏️ Données ({editableRows.length} lignes)
                  </span>
                  <span style={{ color: T.textDim, fontSize: 10 }}>Modifiables avant import</span>
                  <span style={{ color: T.textDim, fontSize: 10, opacity: 0.6, marginLeft: 4 }}>{openDonnees ? "▲" : "▼"}</span>
                </div>
                {openDonnees && (
                  <div style={{ overflowX: "auto", maxHeight: 260, overflowY: "auto" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                      <thead>
                        <tr style={{ position: "sticky", top: 0, background: T.bgDark, zIndex: 1 }}>
                          {visibleCols.map(({ h, i }) => (
                            <th key={i}
                              onClick={openHeaders ? () => deleteColumn(i) : undefined}
                              style={{ color: openHeaders ? T.error : T.textDim, padding: "5px 6px", textAlign: "left", borderBottom: `1px solid ${T.border}`, whiteSpace: "nowrap", fontWeight: 700, cursor: openHeaders ? "pointer" : "default", background: openHeaders ? `${T.error}11` : undefined }}
                            >
                              {openHeaders && <span style={{ fontSize: 9, marginRight: 3, opacity: 0.6 }}>✕</span>}
                              {h}
                            </th>
                          ))}
                          <th style={{ width: 40, borderBottom: `1px solid ${T.border}` }} />
                        </tr>
                      </thead>
                      <tbody>
                        {editableRows.map((row, ri) => {
                          const anomalous = anomalyInfo.isAnomalous(row);
                          return (
                            <tr key={ri} style={{ borderBottom: `1px solid ${T.border}22`, background: anomalous ? "#7C150822" : "transparent", outline: anomalous ? `1px solid ${T.error}33` : "none" }}>
                              {visibleCols.map(({ h }) => (
                                <td key={h} style={{ padding: "3px 4px" }}>
                                  <input
                                    value={row[h] ?? ""}
                                    onChange={(e) => setEditableRows((prev) => prev.map((r, idx) => idx === ri ? { ...r, [h]: e.target.value } : r))}
                                    style={{ background: T.bgCard, border: `1px solid ${T.border2}55`, borderRadius: 4, color: T.text, fontSize: 11, padding: "3px 6px", width: "100%", minWidth: 60, outline: "none", boxSizing: "border-box", fontFamily: "'Share Tech Mono', monospace" }}
                                  />
                                </td>
                              ))}
                              <td style={{ padding: "3px 4px", textAlign: "center", whiteSpace: "nowrap" }}>
                                {anomalous && <span title="Ligne atypique" style={{ color: T.warning, fontSize: 12, marginRight: 2 }}>⚠</span>}
                                <button onClick={() => setEditableRows((prev) => prev.filter((_, idx) => idx !== ri))} style={{ background: "none", border: "none", color: T.error, cursor: "pointer", fontSize: 14, padding: 0 }}>✕</button>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            )}

            {/* Mapping colonnes */}
            {visibleCols.length > 0 && (
              <div style={{ marginBottom: 14, background: T.bgCard, borderRadius: 12, border: `1px solid ${T.border2}`, overflow: "hidden" }}>
                <div
                  onClick={() => { setOpenMapping((o) => !o); setOpenApercu((o) => !o); }}
                  style={{ padding: "8px 14px", background: T.bgDark, borderBottom: openMapping ? `1px solid ${T.border2}` : "none", display: "flex", alignItems: "center", gap: 8, cursor: "pointer" }}
                >
                  <span style={{ color: T.textMuted, fontWeight: 800, fontSize: 12, textTransform: "uppercase", flex: 1 }}>🗂 Mapping colonnes</span>
                  <span style={{ color: T.textDim, fontSize: 10, opacity: 0.6, marginLeft: 4 }}>{openMapping ? "▲" : "▼"}</span>
                </div>
                {openMapping && (
                  <div style={{ padding: "10px 14px" }}>
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
                            style={{ width: "100%", background: T.bgCard, border: `1px solid ${mapping[field] ? T.accent : T.border2}`, borderRadius: 8, color: T.text, fontSize: 12, padding: "7px 6px", outline: "none", fontFamily: "'Share Tech Mono', monospace", boxSizing: "border-box" }}
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
                              style={{ width: "100%", background: T.bgCard, border: `1px solid ${T.border2}`, borderRadius: 8, color: T.accent, fontSize: 12, padding: "7px 8px", outline: "none", fontFamily: "'Share Tech Mono', monospace", boxSizing: "border-box" }}
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
                                  const isDefault = !x.label.trim() || x.label === "EXTRA" || x.label === x.col;
                                  const autoLabel = (isDefault && newCol) ? newCol : x.label;
                                  return { ...x, col: newCol, label: autoLabel };
                                }));
                              }}
                              style={{ width: "100%", background: T.bgCard, border: `1px solid ${ex.col ? T.success : T.border2}`, borderRadius: 8, color: T.text, fontSize: 12, padding: "7px 6px", outline: "none", fontFamily: "'Share Tech Mono', monospace", boxSizing: "border-box" }}
                            >
                              <option value="">— Aucune —</option>
                              {visibleCols.map(({ h }) => <option key={h} value={h}>{h}</option>)}
                            </select>
                          </div>
                          <button
                            onClick={() => setExtras((exs) => exs.filter((_, i) => i !== idx))}
                            style={{ flexShrink: 0, background: `${T.error}22`, border: `1px solid ${T.error}55`, borderRadius: 7, color: T.error, fontSize: 13, cursor: "pointer", padding: "6px 9px", lineHeight: 1 }}
                          >🗑</button>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* Aperçu */}
            {editableRows.length > 0 && (
              <div style={{ background: "#0C2D3A", borderRadius: 10, marginBottom: 14, overflow: "hidden", border: `1px solid #1A5F7A55` }}>
                <div
                  onClick={() => setOpenApercu((o) => !o)}
                  style={{ padding: "8px 12px", background: "#103848", borderBottom: openApercu ? `1px solid #1A5F7A55` : "none", display: "flex", alignItems: "center", gap: 8, cursor: "pointer" }}
                >
                  <span style={{ color: T.textDim, fontSize: 11, textTransform: "uppercase", fontWeight: 700, flex: 1 }}>Aperçu — {editableRows.length} lignes</span>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <span style={{ color: T.textDim, fontSize: 10 }}>Cliquer sur un header pour grouper</span>
                    <button
                      onClick={(e) => { e.stopPropagation(); setAutoRefFmt((v) => !v); }}
                      style={{ background: autoRefFmt ? `${T.accent}22` : T.bgDark, border: `1px solid ${autoRefFmt ? T.accent : T.border2}`, borderRadius: 6, color: autoRefFmt ? T.accent : T.textDim, fontSize: 10, fontWeight: 700, cursor: "pointer", padding: "2px 7px", fontFamily: "'Share Tech Mono', monospace" }}
                    >🔢 Réf. auto</button>
                    <button
                      onClick={(e) => { e.stopPropagation(); setPoidsUnit((u) => u === "t" ? "kg" : "t"); }}
                      style={{ background: poidsUnit === "kg" ? `${T.warning}33` : T.bgDark, border: `1px solid ${T.warning}66`, borderRadius: 6, color: T.warning, fontSize: 10, fontWeight: 700, cursor: "pointer", padding: "2px 7px", fontFamily: "'Share Tech Mono', monospace" }}
                    >{poidsUnit === "t" ? "⚖️ t ⇄ kg" : "⚖️ kg ⇄ t"}</button>
                  </div>
                  <span style={{ color: T.textDim, fontSize: 10, opacity: 0.6, marginLeft: 4 }}>{openApercu ? "▲" : "▼"}</span>
                </div>
                {openApercu && (
                  <div style={{ padding: 10, overflowX: "auto" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                      <thead>
                        <tr>
                          {([
                            { key: "rang",      label: "📍 Rang"      },
                            { key: "reference", label: "🏷 Référence" },
                            { key: "poids",     label: `⚖️ Poids (${poidsUnit})` },
                            { key: "dch",       label: "🏗️ Destination"  },
                            ...extras.map((e, i) => ({ key: `extra_${i}`, label: e.label || "EXTRA" })),
                          ] as { key: string; label: string }[]).map(({ key, label }) => (
                            <th key={key} style={{ padding: "4px 6px", textAlign: "left", borderBottom: `1px solid ${T.border}` }}>
                              {splitEditingField === key ? (
                                <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
                                  <input
                                    autoFocus value={splitInputValue}
                                    onChange={(e) => setSplitInputValue(e.target.value)}
                                    placeholder="ex: 4 4 1"
                                    onKeyDown={(e) => {
                                      if (e.key === "Enter") { setSplitFormats((p) => ({ ...p, [key]: splitInputValue })); setSplitEditingField(null); }
                                      if (e.key === "Escape") setSplitEditingField(null);
                                    }}
                                    onBlur={() => { setSplitFormats((p) => ({ ...p, [key]: splitInputValue })); setSplitEditingField(null); }}
                                    style={{ background: T.bgCard, border: `1px solid ${T.accent}`, borderRadius: 4, color: T.accent, fontSize: 10, padding: "2px 6px", width: 70, outline: "none", fontFamily: "'Share Tech Mono', monospace" }}
                                  />
                                  <button
                                    onClick={() => { setSplitFormats((p) => { const n = { ...p }; delete n[key]; return n; }); setSplitEditingField(null); }}
                                    style={{ background: "none", border: "none", color: T.error, cursor: "pointer", fontSize: 11, padding: 0 }}
                                  >✕</button>
                                </div>
                              ) : (
                                <span
                                  onClick={() => { setSplitEditingField(key); setSplitInputValue(splitFormats[key] || ""); }}
                                  style={{ color: splitFormats[key] ? T.success : T.accent, cursor: "pointer", fontWeight: 700, display: "inline-flex", alignItems: "center", gap: 4 }}
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
                        {editableRows.map((row, ri) => (
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
                  </div>
                )}
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

        {/* ── STEP 3 ── */}
        {step === 3 && parsed && (() => {
          const visibleCount = headers.filter((_, i) => !hiddenCols.has(i)).length;
          const totalRows = parsed.rows.length;
          return (
            <div>
              <div style={{ background: T.bgCard, borderRadius: 12, padding: 16, marginBottom: 14, border: `1px solid ${T.success}55` }}>
                <div style={{ color: T.success, fontWeight: 800, fontSize: 16, marginBottom: 10 }}>✅ Prêt à analyser</div>
                {([
                  ["Fichier", fileName ?? "—"],
                  ["Onglet actif", activeSheet ?? "—"],
                  ["Colonnes importables", String(visibleCount)],
                  ["Lignes totales", String(totalRows)],
                  ["📍 Colonne Rang", mapping.rang || "— non mappé"],
                  ["🏷 Colonne Référence", mapping.reference || "— non mappé"],
                  ["⚖️ Colonne Poids", mapping.poids || "— non mappé"],
                  ["🏗️ Colonne Destination", mapping.dch || "— non mappé"],
                  ["⚖️ Unité poids", poidsUnit === "kg" ? "Kilogrammes (kg)" : "Tonnes (t)"],
                  ["🔢 Réf. auto-groupée", autoRefFmt ? "Activé" : "Désactivé"],
                  ...extras.filter((e) => e.label.trim()).map((e) => [
                    `🔖 ${e.label}`,
                    e.col ? e.col : "+ colonne vide"
                  ]),
                  ["Dernier déploiement", new Date().toLocaleString("fr-FR")],
                  ["Dernier git push", "Voir la console ou l'historique git"],
                ] as [string, string][]).map(([k, v]) => (
                  <div key={k} style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", borderBottom: `1px solid ${T.border2}33` }}>
                    <span style={{ color: T.textMuted, fontSize: 13 }}>{k}</span>
                    <span style={{ color: T.text, fontWeight: 700, fontSize: 13, fontFamily: "monospace" }}>{v}</span>
                  </div>
                ))}
              </div>
              <div style={{ display: "flex", gap: 10 }}>
                <Btn onClick={() => setStep(2)} color={T.border2} textColor={T.textMuted} fullWidth>← Retour</Btn>
                <Btn
                  onClick={() => {
                    if (!parsed) return;
                    const tail = parsed.rows.slice(editableRows.length);
                    const currentHeaders = headers;
                    let newRows: RawData = [
                      ...editableRows.map((rowObj) => currentHeaders.map((h) => { const v = rowObj[h] ?? ""; return v === "" ? null : v; })),
                      ...tail,
                    ];
                    const newHeaders = [...currentHeaders];
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
