// ============================================================
// STEtruc — Onglet Rapport
// ============================================================

import { useState } from "react";
import { T, applyGrouping, thsep, autoFormatRef } from "../types";
import { useApp } from "../AppContext";
import { EmptyState } from "../components";

export default function RapportTab() {
  const {
    parsed, headers, allRows, hiddenRows,
    mapping, extras, splitFormats,
    pointedRows,
    destinations, rowDestinations, reassignedRows,
    poidsUnit, autoRefFmt,
    tallyPrev, setTallyPrev,
    chargementMaxi, setChargementMaxi,
    dechargementMaxi, setDechargementMaxi,
  } = useApp();

  const [openResume,       setOpenResume]       = useState(false);
  const [openMouvements,   setOpenMouvements]   = useState(true);
  const [openChargement,   setOpenChargement]   = useState(false);
  const [openDechargement, setOpenDechargement] = useState(false);
  const [openTally,        setOpenTally]        = useState(false);
  const [openPointees,     setOpenPointees]     = useState(true);
  const [openReaff,        setOpenReaff]        = useState(true);

  if (!parsed) {
    return (
      <div style={{ flex: 1, overflowY: "auto", paddingBottom: 80, background: T.bg }}>
        <div style={{ padding: "16px 16px 12px", background: T.bgDark, borderBottom: `1px solid ${T.border}` }}>
          <div style={{ color: T.textDim, fontSize: 9, letterSpacing: "0.2em", textTransform: "uppercase", marginBottom: 4 }}>STEtruc</div>
          <h1 style={{ color: T.text, fontSize: 20, fontWeight: 700 }}>Rapport</h1>
        </div>
        <EmptyState icon="📋" text="Aucun fichier chargé" sub="Importez un fichier Excel d'abord" />
      </div>
    );
  }

  const poidsIdx  = mapping.poids     ? headers.indexOf(mapping.poids)     : -1;
  const refIdx    = mapping.reference ? headers.indexOf(mapping.reference) : -1;
  const rangIdx   = mapping.rang      ? headers.indexOf(mapping.rang)      : -1;
  const poidsFmt  = splitFormats["poids"] ?? "";
  const unitLabel = poidsUnit === "kg" ? "kg" : "t";

  const fmtWeight = (w: number): string => {
    const val = poidsUnit === "kg" ? (w * 1000).toFixed(0) : parseFloat(w.toFixed(3)).toString();
    return (poidsFmt ? applyGrouping(val, poidsFmt) : thsep(val)) + " " + unitLabel;
  };

  const visibleRows = allRows.map((row, ri) => ({ row, ri })).filter(({ ri }) => !hiddenRows.has(ri));
  const excludedFromReport = new Set(destinations.filter(d => d.excludeFromReport).map(d => d.name));

  const destStatsMap = new Map<string, { count: number; weight: number }>();
  for (const { row, ri } of visibleRows) {
    const dest = rowDestinations.get(ri);
    if (!dest || excludedFromReport.has(dest)) continue;
    const w = poidsIdx >= 0 ? (parseFloat(String(row[poidsIdx] ?? "")) || 0) : 0;
    const s = destStatsMap.get(dest) ?? { count: 0, weight: 0 };
    destStatsMap.set(dest, { count: s.count + 1, weight: s.weight + w });
  }

  void pointedRows;
  const totalAffected = [...rowDestinations.entries()].filter(([, d]) => !excludedFromReport.has(d)).length;
  const totalRows     = visibleRows.length;
  const totalWeight   = poidsIdx >= 0 ? visibleRows.reduce((acc, { row }) => acc + (parseFloat(String(row[poidsIdx] ?? "")) || 0), 0) : 0;
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
      style={{ display: "flex", alignItems: "center", cursor: "pointer", userSelect: "none", color, fontWeight: 800, fontSize: 11, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: open ? 10 : 0 }}
    >
      <span style={{ flex: 1 }}>{label}</span>
      <span style={{ fontSize: 12, opacity: 0.55, transition: "transform 0.2s", display: "inline-block", transform: open ? "rotate(90deg)" : "rotate(0deg)" }}>▶</span>
    </div>
  );

  return (
    <div style={{ flex: 1, overflowY: "auto", paddingBottom: 80, background: T.bg }}>
      <div style={{ position: "relative" }}>
        <div style={{ padding: "16px 16px 12px", background: T.bgDark, borderBottom: `1px solid ${T.border}` }}>
          <div style={{ color: T.textDim, fontSize: 9, letterSpacing: "0.2em", textTransform: "uppercase", marginBottom: 4 }}>STEtruc</div>
          <h1 style={{ color: T.text, fontSize: 20, fontWeight: 700 }}>Rapport</h1>
          <div style={{ color: T.textMuted, fontSize: 11, marginTop: 2 }}>Synthèse pointage & mouvements</div>
        </div>
        <button
          onClick={() => {
            setTallyPrev({}); setChargementMaxi({}); setDechargementMaxi({});
            ["ste_tallyPrev","ste_chargementMaxi","ste_dechargementMaxi"].forEach(k => { try { localStorage.removeItem(k); } catch {} });
          }}
          title="Effacer toutes les valeurs saisies manuellement"
          style={{ position: "absolute", top: 14, right: 12, background: `${T.error}22`, border: `1px solid ${T.error}55`, borderRadius: 8, color: T.error, fontSize: 16, padding: "4px 10px", cursor: "pointer", lineHeight: 1 }}
        >🗑</button>
      </div>

      <div style={{ padding: "12px 12px 0" }}>

        {/* Mouvements */}
        {destStatsMap.size > 0 && (
          <div style={{ background: T.bgCard, borderRadius: 12, border: `1px solid #7C3AED66`, padding: "12px 14px", marginBottom: 12 }}>
            <SectionHeader label="🏗️ Mouvements" open={openMouvements} toggle={() => setOpenMouvements(o => !o)} color="#A78BFA" />
            {openMouvements && (
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
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
                        <td style={{ padding: "5px 8px 5px 0" }}><span style={{ background: `${d.color}33`, borderLeft: `3px solid ${d.color}`, padding: "2px 8px", borderRadius: 4, color: d.color, fontWeight: 700 }}>{d.name}</span></td>
                        <td style={{ textAlign: "right", padding: "5px 6px", color: T.text, fontFamily: "monospace" }}>{s.count}</td>
                        {poidsIdx >= 0 && <td style={{ textAlign: "right", padding: "5px 0 5px 6px", color: T.warning, fontFamily: "monospace" }}>{fmtWeight(s.weight)}</td>}
                      </tr>
                    );
                  })}
                  {destStatsMap.size > 1 && (
                    <tr style={{ borderTop: `1px solid ${T.border}` }}>
                      <td style={{ padding: "5px 8px 5px 0", color: T.textDim, fontStyle: "italic" }}>Total affecté</td>
                      <td style={{ textAlign: "right", padding: "5px 6px", color: T.accent, fontWeight: 700, fontFamily: "monospace" }}>{[...destStatsMap.values()].reduce((a, b) => a + b.count, 0)}</td>
                      {poidsIdx >= 0 && <td style={{ textAlign: "right", padding: "5px 0 5px 6px", color: T.accent, fontWeight: 700, fontFamily: "monospace" }}>{fmtWeight([...destStatsMap.values()].reduce((a, b) => a + b.weight, 0))}</td>}
                    </tr>
                  )}
                </tbody>
              </table>
            )}
          </div>
        )}

        {/* Tally */}
        <div style={{ background: T.bgCard, borderRadius: 12, border: `1px solid #05966955`, padding: "12px 14px", marginBottom: 12 }}>
          <SectionHeader label="📋 Tally" open={openTally} toggle={() => setOpenTally(o => !o)} color="#34D399" />
          {openTally && (() => {
            const tallyDests = destinations.filter(d => !d.excludeFromReport);
            const setPrev = (name: string, field: "qty" | "weight", val: string) =>
              setTallyPrev(p => ({ ...p, [name]: { qty: p[name]?.qty ?? "", weight: p[name]?.weight ?? "", [field]: val } }));
            const thStyle = (align: "left" | "right" | "center" = "right"): React.CSSProperties => ({ textAlign: align, color: T.textDim, padding: "3px 5px 5px", fontWeight: 600, fontSize: 10, whiteSpace: "nowrap" });
            const tdStyle = (color: string = T.text): React.CSSProperties => ({ textAlign: "right", padding: "4px 5px", fontFamily: "monospace", fontSize: 10, color });
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
                      const tQty = s?.count ?? 0; const tW = s?.weight ?? 0;
                      const pQty = parseFloat(tallyPrev[d.name]?.qty ?? "0") || 0;
                      const pW   = parseFloat(String(tallyPrev[d.name]?.weight ?? "0").replace(",", ".")) || 0;
                      const totQty = tQty + pQty; const totW = tW + pW;
                      const inputStyle: React.CSSProperties = { width: 58, textAlign: "right", background: T.bgDark, border: `1px solid ${T.border2}`, borderRadius: 4, color: "#FB923C", fontFamily: "monospace", fontSize: 10, padding: "2px 4px" };
                      return (
                        <tr key={d.name} style={{ borderBottom: `1px solid ${T.border2}22` }}>
                          <td style={{ padding: "4px 6px 4px 0" }}><span style={{ background: `${d.color}22`, borderLeft: `3px solid ${d.color}`, padding: "1px 6px", borderRadius: 3, color: d.color, fontWeight: 700, fontSize: 10 }}>{d.name}</span></td>
                          <td style={tdStyle("#60A5FA")}>{tQty || "—"}</td>
                          <td style={tdStyle("#60A5FA")}>{tW > 0 ? fmtWeight(tW) : "—"}</td>
                          <td style={{ padding: "3px 5px", textAlign: "right" }}><input type="number" min={0} style={inputStyle} value={tallyPrev[d.name]?.qty ?? ""} onChange={e => setPrev(d.name, "qty", e.target.value)} placeholder="0" /></td>
                          <td style={{ padding: "3px 5px", textAlign: "right" }}>
                            <input type="text" inputMode="decimal" style={{ ...inputStyle, width: 70 }}
                              value={(() => { const raw = tallyPrev[d.name]?.weight ?? ""; const n = parseFloat(raw.replace(",", ".")); return isNaN(n) ? raw : thsep(String(n)); })()}
                              onChange={e => setPrev(d.name, "weight", e.target.value.replace(/\s/g, ""))} placeholder="0" />
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
            const thS = (align: "left" | "right" | "center" = "right"): React.CSSProperties => ({ textAlign: align, color: T.textDim, padding: "3px 5px 5px", fontWeight: 600, fontSize: 10, whiteSpace: "nowrap" });
            const tdS = (color: string = T.text): React.CSSProperties => ({ textAlign: "right", padding: "4px 5px", fontFamily: "monospace", fontSize: 10, color });
            const inputS: React.CSSProperties = { width: 62, textAlign: "right", background: T.bgDark, border: `1px solid ${T.border2}`, borderRadius: 4, color: "#FB923C", fontFamily: "monospace", fontSize: 10, padding: "2px 4px" };
            let sumTotQty = 0, sumTotW = 0, sumMaxiQty = 0, sumMaxiW = 0;
            for (const d of dests) {
              const s = destStatsMap.get(d.name);
              const pQty = parseFloat(tallyPrev[d.name]?.qty ?? "0") || 0;
              const pW   = parseFloat(String(tallyPrev[d.name]?.weight ?? "0").replace(",", ".")) || 0;
              sumTotQty  += (s?.count ?? 0) + pQty; sumTotW += (s?.weight ?? 0) + pW;
              sumMaxiQty += parseFloat(chargementMaxi[d.name]?.qty ?? "0") || 0;
              sumMaxiW   += parseFloat(String(chargementMaxi[d.name]?.weight ?? "0").replace(",", ".")) || 0;
            }
            const sumEstW = sumMaxiW - sumTotW;
            const sumMoyW = sumTotQty > 0 ? sumTotW / sumTotQty : NaN;
            const sumEstQty = !isNaN(sumMoyW) && sumMoyW > 0 ? Math.floor(sumEstW / sumMoyW) : NaN;
            return (
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                  <thead>
                    <tr style={{ borderBottom: `1px solid ${T.border}44` }}>
                      <th style={{ ...thS("left"), padding: "3px 6px 0 0" }} rowSpan={2}>Dest.</th>
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
                      const tQty = (s?.count ?? 0) + pQty; const tW = (s?.weight ?? 0) + pW;
                      const mQty = parseFloat(chargementMaxi[d.name]?.qty ?? "0") || 0;
                      const mW   = parseFloat(String(chargementMaxi[d.name]?.weight ?? "0").replace(",", ".")) || 0;
                      const eW = mW - tW; const moyW = tQty > 0 ? tW / tQty : NaN;
                      const effectiveMoyW = !isNaN(moyW) && moyW > 0 ? moyW : sumMoyW;
                      const eQty = !isNaN(effectiveMoyW) && effectiveMoyW > 0 ? Math.floor(eW / effectiveMoyW) : NaN;
                      return (
                        <tr key={d.name} style={{ borderBottom: `1px solid ${T.border2}22` }}>
                          <td style={{ padding: "4px 6px 4px 0" }}><span style={{ background: `${d.color}22`, borderLeft: `3px solid ${d.color}`, padding: "1px 6px", borderRadius: 3, color: d.color, fontWeight: 700, fontSize: 10 }}>{d.name}</span></td>
                          <td style={tdS("#60A5FA")}>{tQty || "—"}</td>
                          <td style={tdS("#60A5FA")}>{tW > 0 ? fmtWeight(tW) : "—"}</td>
                          <td style={{ padding: "3px 5px", textAlign: "right", color: T.textDim, opacity: 0.35, fontFamily: "monospace", fontSize: 11 }}>{mQty > 0 ? thsep(String(mQty)) : "—"}</td>
                          <td style={{ padding: "3px 5px", textAlign: "right" }}>
                            <input type="text" inputMode="decimal" style={{ ...inputS, width: 72 }}
                              value={(() => { const raw = chargementMaxi[d.name]?.weight ?? ""; const n = parseFloat(raw.replace(",", ".")); return isNaN(n) ? raw : thsep(String(n)); })()}
                              onChange={e => setMaxi(d.name, "weight", e.target.value.replace(/\s/g, ""))} placeholder="0" />
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
            const thS = (align: "left" | "right" | "center" = "right"): React.CSSProperties => ({ textAlign: align, color: T.textDim, padding: "3px 5px 5px", fontWeight: 600, fontSize: 10, whiteSpace: "nowrap" });
            const tdS = (color: string = T.text): React.CSSProperties => ({ textAlign: "right", padding: "4px 5px", fontFamily: "monospace", fontSize: 10, color });
            const inputS: React.CSSProperties = { width: 62, textAlign: "right", background: T.bgDark, border: `1px solid ${T.border2}`, borderRadius: 4, color: "#FB923C", fontFamily: "monospace", fontSize: 10, padding: "2px 4px" };
            let sumTotQty = 0, sumTotW = 0, sumMaxiQty = 0, sumMaxiW = 0;
            for (const d of dests) {
              const s = destStatsMap.get(d.name);
              const pQty = parseFloat(tallyPrev[d.name]?.qty ?? "0") || 0;
              const pW   = parseFloat(String(tallyPrev[d.name]?.weight ?? "0").replace(",", ".")) || 0;
              sumTotQty += (s?.count ?? 0) + pQty; sumTotW += (s?.weight ?? 0) + pW;
              sumMaxiQty += parseFloat(dechargementMaxi[d.name]?.qty ?? "0") || 0;
              sumMaxiW   += parseFloat(String(dechargementMaxi[d.name]?.weight ?? "0").replace(",", ".")) || 0;
            }
            const sumEstW = sumTotW - sumMaxiW; const sumEstQty = sumTotQty - sumMaxiQty;
            return (
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                  <thead>
                    <tr style={{ borderBottom: `1px solid ${T.border}44` }}>
                      <th style={{ ...thS("left"), padding: "3px 6px 0 0" }} rowSpan={2}>Dest.</th>
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
                      const tQty = (s?.count ?? 0) + pQty; const tW = (s?.weight ?? 0) + pW;
                      const mQty = parseFloat(dechargementMaxi[d.name]?.qty ?? "0") || 0;
                      const mW   = parseFloat(String(dechargementMaxi[d.name]?.weight ?? "0").replace(",", ".")) || 0;
                      const eW = tW - mW; const eQty = tQty - mQty; const moyW = tQty > 0 ? tW / tQty : NaN;
                      return (
                        <tr key={d.name} style={{ borderBottom: `1px solid ${T.border2}22` }}>
                          <td style={{ padding: "4px 6px 4px 0" }}><span style={{ background: `${d.color}22`, borderLeft: `3px solid ${d.color}`, padding: "1px 6px", borderRadius: 3, color: d.color, fontWeight: 700, fontSize: 10 }}>{d.name}</span></td>
                          <td style={tdS("#60A5FA")}>{tQty || "—"}</td>
                          <td style={tdS("#60A5FA")}>{tW > 0 ? fmtWeight(tW) : "—"}</td>
                          <td style={{ padding: "3px 5px", textAlign: "right" }}><input type="number" min={0} style={inputS} value={dechargementMaxi[d.name]?.qty ?? ""} onChange={e => setMaxi(d.name, "qty", e.target.value)} placeholder="0" /></td>
                          <td style={{ padding: "3px 5px", textAlign: "right" }}>
                            <input type="text" inputMode="decimal" style={{ ...inputS, width: 72 }}
                              value={(() => { const raw = dechargementMaxi[d.name]?.weight ?? ""; const n = parseFloat(raw.replace(",", ".")); return isNaN(n) ? raw : thsep(String(n)); })()}
                              onChange={e => setMaxi(d.name, "weight", e.target.value.replace(/\s/g, ""))} placeholder="0" />
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
          {openResume && (
            <div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "8px 12px" }}>
                {[
                  { label: "Lignes totales",    val: String(totalRows),       color: T.text },
                  { label: "Affectées (dest.)",  val: String(totalAffected),   color: "#A78BFA" },
                  ...(reassignedRows.size > 0 ? [{ label: "Réaffectations", val: String(reassignedRows.size), color: T.error }] : []),
                  ...(poidsIdx >= 0 ? [
                    { label: `Poids total`,    val: fmtWeight(totalWeight),    color: T.warning },
                    { label: `Poids affecté`,  val: fmtWeight(affectedWeight), color: "#A78BFA" },
                    { label: `Poids moyen`,    val: totalRows > 0 && totalWeight > 0 ? (poidsUnit === "kg" ? fmtWeight(totalWeight / totalRows) : thsep(Math.round(totalWeight / totalRows).toString()) + " t") : "—", color: "#F9A8D4" },
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
                    <div style={{ color: hneColor, fontWeight: 700, fontSize: 11, marginBottom: 6 }}>🚫 Hors rapport — {hneRows.length} ligne{hneRows.length > 1 ? "s" : ""}</div>
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
                              <td style={{ textAlign: "right", padding: "3px 0 3px 6px" }}><span style={{ color: destColor, fontWeight: 700, fontSize: 10 }}>{dest}</span></td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                );
              })()}
            </div>
          )}
        </div>

        {/* Lignes pointées */}
        {pointedRows.size > 0 && (
          <div style={{ background: T.bgCard, borderRadius: 12, border: `1px solid ${T.success}44`, padding: "12px 14px", marginBottom: 12 }}>
            <SectionHeader label={`✅ Lignes pointées (${pointedRows.size})`} open={openPointees} toggle={() => setOpenPointees(o => !o)} color={T.success} />
            {openPointees && (
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
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
                        {extraCols.map(e => { const ci = headers.indexOf(e.col); return <td key={e.col} style={{ padding: "3px 6px", color: T.success, fontFamily: "monospace" }}>{ci >= 0 ? String(row[ci] ?? "") || "—" : "—"}</td>; })}
                        <td style={{ textAlign: "right", padding: "3px 0 3px 6px" }}>
                          {dest ? <span style={{ color: destColor ?? T.accent, fontWeight: 700, fontSize: 10 }}>{dest}</span> : <span style={{ color: T.textDim, fontSize: 10 }}>—</span>}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            )}
          </div>
        )}

        {/* Réaffectations */}
        {reassignedRows.size > 0 && (
          <div style={{ background: T.bgCard, borderRadius: 12, border: `1px solid ${T.error}55`, padding: "12px 14px", marginBottom: 12 }}>
            <SectionHeader label={`⚠️ Réaffectations (${reassignedRows.size} ligne${reassignedRows.size > 1 ? "s" : ""})`} open={openReaff} toggle={() => setOpenReaff(o => !o)} color={T.error} />
            {openReaff && (
              <>
                <div style={{ color: T.textDim, fontSize: 10, marginBottom: 8 }}>Ces lignes ont changé de destination — vérifier doublons ou erreurs de saisie.</div>
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
                      const current = rowDestinations.get(ri) ?? "—";
                      const currentColor = destinations.find(d => d.name === current)?.color ?? T.textDim;
                      return (
                        <tr key={ri} style={{ borderBottom: `1px solid ${T.border2}22` }}>
                          <td style={{ padding: "4px 6px 4px 0", color: T.textDim, fontFamily: "monospace" }}>{ri + 1}</td>
                          {refIdx >= 0 && <td style={{ padding: "4px 6px", color: T.text, fontFamily: "monospace", fontSize: 10 }}>{refDisplay || "—"}</td>}
                          {poidsIdx >= 0 && <td style={{ padding: "4px 6px", color: T.text, fontFamily: "monospace", fontSize: 10, textAlign: "right" }}>{!isNaN(poidsRaw) ? fmtWeight(poidsRaw) : "—"}</td>}
                          <td style={{ padding: "4px 6px" }}>
                            <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                              {history.map((h, i) => {
                                const fromColor = destinations.find(d => d.name === h.from)?.color ?? T.textDim;
                                const toColor   = destinations.find(d => d.name === h.to)?.color   ?? T.textDim;
                                return <span key={i} style={{ fontSize: 10, color: T.textDim }}><span style={{ color: fromColor, fontWeight: 600 }}>{h.from}</span><span style={{ opacity: 0.5 }}> → </span><span style={{ color: toColor, fontWeight: 600 }}>{h.to}</span></span>;
                              })}
                            </div>
                          </td>
                          <td style={{ padding: "4px 0 4px 6px" }}><span style={{ color: currentColor, fontWeight: 700, background: `${currentColor}22`, padding: "1px 6px", borderRadius: 4 }}>{current}</span></td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </>
            )}
          </div>
        )}

        {destStatsMap.size === 0 && pointedRows.size === 0 && reassignedRows.size === 0 && (
          <EmptyState icon="📋" text="Aucune donnée de rapport" sub="Pointez des lignes ou affectez des destinations dans l'onglet Pointage" />
        )}
      </div>
    </div>
  );
}
