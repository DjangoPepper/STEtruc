// ============================================================
// STEtruc — Composants UI partagés
// ============================================================

import { useState } from "react";
import { T } from "./types";
import { useApp } from "./AppContext";

// ─── Toast ────────────────────────────────────────────────────
export function Toast() {
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

// ─── Bottom Navigation ────────────────────────────────────────
export function BottomNav() {
  const { activeTab, setActiveTab, parsed } = useApp();
  const [darkMode, setDarkMode] = useState(true);
  type Tab = "import" | "iec" | "tableau" | "rapport" | "export";
  const tabs: { id: Tab; icon: string; label: string }[] = [
    { id: "import",  icon: "⬇️",   label: "Import"   },
    { id: "tableau", icon: "📊",   label: "Pointage" },
    { id: "rapport", icon: "📋",   label: "Rapport"  },
    { id: "export",  icon: "⬆️",   label: "Export"   },
  ];
  if (activeTab === "tableau") return null;
  return (
    <nav style={{
      position: "fixed", bottom: 0, left: 0, right: 0, height: 54,
      background: darkMode ? "#0F172A" : "#F8FAFC", borderTop: `2px solid ${darkMode ? T.border : '#CBD5E1'}`,
      display: "flex", alignItems: "stretch", justifyContent: "space-between", zIndex: 100,
      maxWidth: 540, margin: "0 auto",
      gap: 1,
    }}>
      {tabs.map((t, idx) => {
        const isActive = activeTab === t.id;
        const disabled = (t.id === "tableau" || t.id === "rapport" || t.id === "export") && !parsed;
        return (
          <button
            key={t.id}
            onClick={() => !disabled && setActiveTab(t.id)}
            style={{
              flex: 1, display: "flex", flexDirection: "column", alignItems: "center",
              justifyContent: "center", background: "none", border: "none",
              color: disabled ? (darkMode ? T.textDim : '#64748B') : isActive ? (darkMode ? T.accent : '#0F172A') : (darkMode ? T.textDim : '#64748B'),
              fontSize: 11, fontWeight: 700, cursor: disabled ? "not-allowed" : "pointer",
              gap: 0.5, // Réduit le gap entre icônes
              letterSpacing: 0.5,
              borderTop: isActive ? `3px solid ${darkMode ? T.accent : '#0F172A'}` : "3px solid transparent",
              transition: "all 0.15s",
              fontFamily: "'Share Tech Mono', monospace",
              minHeight: 44,
              padding: '0 2px',
            }}
          >
            <span style={{ fontSize: 20 }}>{t.icon}</span>
            {t.label}
          </button>
        );
      })}
      {/* Bouton sombre/clair à droite de Export */}
      <button
        onClick={() => setDarkMode((d) => !d)}
        style={{
          marginLeft: 6,
          background: darkMode ? T.bgCard : '#F1F5F9',
          color: darkMode ? T.accent : '#0F172A',
          border: `1px solid ${darkMode ? T.accent : '#CBD5E1'}`,
          borderRadius: 8,
          fontSize: 13,
          fontWeight: 700,
          padding: '4px 10px',
          cursor: 'pointer',
          minWidth: 36,
          alignSelf: 'center',
        }}
        title={darkMode ? 'Mode sombre' : 'Mode clair'}
      >
        {darkMode ? '🌙' : '☀️'}
      </button>
    </nav>
  );
}

// ─── Page Header ──────────────────────────────────────────────
export function PageHeader({ title, subtitle }: { title: string; subtitle?: string }) {
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

// ─── Btn ──────────────────────────────────────────────────────
export function Btn({
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

// ─── Stats Bar ────────────────────────────────────────────────
export function StatsBar() {
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

// ─── PointageInfos ────────────────────────────────────────────
export function PointageInfos() {
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

// ─── Section Title ────────────────────────────────────────────
export function SectionTitle({ icon, text, count, right }: {
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

// ─── Empty State ──────────────────────────────────────────────
export function EmptyState({ icon, text, sub }: { icon: string; text: string; sub?: string }) {
  return (
    <div style={{ textAlign: "center", padding: "48px 20px" }}>
      <div style={{ fontSize: 52, marginBottom: 12 }}>{icon}</div>
      <div style={{ color: T.textMuted, fontSize: 15, fontWeight: 700, marginBottom: 6 }}>{text}</div>
      {sub && <div style={{ color: T.textDim, fontSize: 12 }}>{sub}</div>}
    </div>
  );
}

// ─── Add Row Modal ────────────────────────────────────────────
export function AddRowModal({ onClose }: { onClose: () => void }) {
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
      style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.7)", display: "flex", alignItems: "flex-end", zIndex: 200 }}
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
              <div style={{ color: T.textDim, fontSize: 10, marginBottom: 4, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.5 }}>{h}</div>
              <input
                value={values[i]}
                onChange={(e) => { const next = [...values]; next[i] = e.target.value; setValues(next); }}
                onKeyDown={(e) => { if (e.key === "Enter") handleAdd(); if (e.key === "Escape") onClose(); }}
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
