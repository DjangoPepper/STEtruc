// ============================================================
// STEtruc — Onglet Export
// ============================================================

import { useCallback } from "react";
import * as XLSX from "xlsx";
import { T, RawData } from "../types";
import { useApp } from "../AppContext";
import { Btn, EmptyState, StatsBar, SectionTitle } from "../components";

// Hook pour flush l'onglet actif dans sheetStates avant export
export function useFlushActiveSheet() {
  const { activeSheet, parsed, headers, addedRows, hiddenCols, hiddenRows, sheetStates } = useApp();
  return useCallback(() => {
    if (activeSheet && parsed) {
      sheetStates.current.set(activeSheet, { parsed, headers, addedRows, hiddenCols, hiddenRows });
    }
  }, [activeSheet, parsed, headers, addedRows, hiddenCols, hiddenRows, sheetStates]);
}

export default function ExportTab() {
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
    flushActive();
    const wb2 = XLSX.utils.book_new();
    const visibleSheets = sheetNames.filter((n) => !hiddenSheets.has(n));
    visibleSheets.forEach((sheetName) => {
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
    setParsed(null); setFileName(null); setWorkbook(null); setSheetNames([]);
    setActiveSheet(null); setAddedRows([]); setHiddenSheets(new Set());
    setSheetSelectMode("none"); setSelectedSheets(new Set());
    setPointedRows(new Set()); setRowOverrides(new Map());
    setActiveTab("import");
    showToast("Réinitialisé", "info");
  };

  if (!parsed) {
    return (
      <div style={{ flex: 1, overflowY: "auto", paddingBottom: 80, background: T.bg }}>
        <div style={{ padding: "16px 16px 12px", background: T.bgDark, borderBottom: `1px solid ${T.border}` }}>
          <div style={{ color: T.textDim, fontSize: 9, letterSpacing: "0.2em", textTransform: "uppercase", marginBottom: 4 }}>STEtruc</div>
          <h1 style={{ color: T.text, fontSize: 20, fontWeight: 700 }}>Export</h1>
        </div>
        <EmptyState icon="📭" text="Aucune donnée à exporter" sub="Importez et nettoyez un fichier d'abord" />
        <div style={{ padding: 16 }}>
          <Btn onClick={() => setActiveTab("import")} color={T.accent} textColor="#0F172A" fullWidth>⬇️ Aller à l'import</Btn>
        </div>
      </div>
    );
  }

  flushActive();

  const visibleSheets = sheetNames.filter((n) => !hiddenSheets.has(n));
  let totalColsAll = 0, visibleColsAll = 0, totalRowsAll = 0, visibleRowsAll = 0, totalAddedAll = 0;
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
      <div style={{ padding: "16px 16px 12px", background: T.bgDark, borderBottom: `1px solid ${T.border}` }}>
        <div style={{ color: T.textDim, fontSize: 9, letterSpacing: "0.2em", textTransform: "uppercase", marginBottom: 4 }}>STEtruc</div>
        <h1 style={{ color: T.text, fontSize: 20, fontWeight: 700 }}>Export</h1>
        <div style={{ color: T.textMuted, fontSize: 11, marginTop: 2 }}>Téléchargez le fichier nettoyé</div>
      </div>
      <StatsBar />

      <div style={{ padding: 16 }}>
        {/* Summary */}
        <div style={{ background: T.bgCard, border: `1px solid ${T.border}`, borderRadius: 12, overflow: "hidden", marginBottom: 16 }}>
          <SectionTitle icon="📋" text="Résumé" />
          <div style={{ padding: "12px 16px" }}>
            {([
              { label: "Colonnes visibles", value: `${visibleColsAll} / ${totalColsAll}`,             color: T.accent },
              { label: "Lignes visibles",   value: `${visibleRowsAll} / ${totalRowsAll}`,             color: T.success },
              { label: "Onglets exportés",  value: `${visibleSheets.length} / ${sheetNames.length}`, color: T.warning },
              { label: "Lignes ajoutées",   value: `+ ${totalAddedAll}`,                             color: T.success },
            ] as { label: string; value: string; color: string }[]).map((r) => (
              <div key={r.label} style={{ display: "flex", justifyContent: "space-between", padding: "8px 0", borderBottom: `1px solid ${T.border}22` }}>
                <span style={{ color: T.textMuted, fontSize: 13 }}>{r.label}</span>
                <span style={{ color: r.color, fontWeight: 700, fontSize: 13, fontFamily: "monospace" }}>{r.value}</span>
              </div>
            ))}
          </div>
        </div>

        {/* Modifications */}
        <div style={{ background: T.bgCard, border: `1px solid ${T.border}`, borderRadius: 12, marginBottom: 16, overflow: "hidden" }}>
          <SectionTitle icon="✏️" text="Modifications actives" />
          <div style={{ padding: "10px 16px", display: "flex", flexDirection: "column", gap: 6 }}>
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
                      : <span style={{ color: T.textDim, fontSize: 11, marginLeft: 8 }}>{s ? "nettoyé" : "brut (non visité)"}</span>}
                  </div>
                </div>
              );
            })}
            {hiddenSheets.size > 0 && <div style={{ display: "flex", gap: 8, color: T.error, fontSize: 12, paddingTop: 4 }}><span>🚫</span>{hiddenSheets.size} onglet(s) exclu(s)</div>}
          </div>
        </div>

        {/* Export */}
        <div style={{ background: T.bgCard, border: `1px solid ${T.border}`, borderRadius: 12, marginBottom: 16, overflow: "hidden" }}>
          <SectionTitle icon="⬆️" text="Télécharger" />
          <div style={{ padding: 16 }}>
            <div style={{ color: T.textMuted, fontSize: 11, marginBottom: 6, fontWeight: 600 }}>Nom du fichier</div>
            <div style={{ display: "flex", alignItems: "center", marginBottom: 16, border: `1px solid ${T.accentDim}`, borderRadius: 8, overflow: "hidden" }}>
              <input
                value={exportFileName}
                onChange={(e) => setExportFileName(e.target.value)}
                onKeyDown={(e) => { if (e.key === "Enter") exportClean(); }}
                spellCheck={false}
                style={{ flex: 1, background: T.bgDark, border: "none", outline: "none", color: T.success, fontSize: 13, padding: "10px 12px", fontFamily: "'Share Tech Mono', monospace", fontWeight: 700 }}
              />
              <span style={{ background: T.bgDark, color: T.textDim, fontSize: 12, padding: "10px 10px 10px 0", fontFamily: "'Share Tech Mono', monospace", pointerEvents: "none" }}>.xlsx</span>
            </div>
            <Btn onClick={exportClean} color={T.success} textColor="#0F172A" fullWidth>⬇️ Télécharger le fichier nettoyé</Btn>
          </div>
        </div>

        {/* Danger zone */}
        <div style={{ background: "#1A0A0A", border: `1px solid ${T.error}33`, borderRadius: 12, overflow: "hidden" }}>
          <SectionTitle icon="⚠️" text="Zone dangereuse" />
          <div style={{ padding: 16 }}>
            <div style={{ color: T.textMuted, fontSize: 12, marginBottom: 12 }}>Efface toutes les données chargées et les modifications.</div>
            <Btn onClick={handleReset} danger fullWidth>✕ Réinitialiser</Btn>
          </div>
        </div>
      </div>
    </div>
  );
}
