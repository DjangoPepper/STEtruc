// ============================================================
// STEtruc — App root (refactorisé multi-fichiers)
// ============================================================

import { AppProvider, useApp } from "./AppContext";
import { Toast, BottomNav } from "./components";
import { T } from "./types";
import { Tab } from "./types";

// Onglets
import ImportTab  from "./tabs/ImportTab";
import PointageTab from "./tabs/PointageTab";
import RapportTab  from "./tabs/RapportTab";
import ExportTab   from "./tabs/ExportTab";
import ExcelCleaner from "./tabs/ExcelCleanerTab";



import { useApp as useAppContext } from "./AppContext";

function IecPage() {
  // Utilise le contexte global pour mettre à jour les données de pointage
  const { setParsed, setHeaders, setFileName } = useAppContext();
  const handleSendToPointage = (data: { headers: string[]; rows: any[][]; fileName: string }) => {
    setParsed({ headers: data.headers, rows: data.rows, headerRowIndex: 0 });
    setHeaders(data.headers);
    setFileName(data.fileName);
  };
  return <ExcelCleaner dark={false} onDarkToggle={() => {}} onSendToPointage={handleSendToPointage} />;
}


import React, { useEffect, useState } from "react";

function AppInner() {
  const { activeTab, setActiveTab } = useApp();
  const [darkMode, setDarkMode] = useState(true);

  useEffect(() => {
    const handler = (e: Event) => {
      const custom = e as CustomEvent;
      if (custom.detail && custom.detail.tab) {
        setActiveTab(custom.detail.tab);
      }
    };
    window.addEventListener("STEtruc_setActiveTab", handler);
    return () => window.removeEventListener("STEtruc_setActiveTab", handler);
  }, [setActiveTab]);

  const pages: Record<Tab, React.ReactNode> = {
    import:  <ImportTab />, 
    iec:     <IecPage />, 
    tableau: <PointageTab />, 
    rapport: <RapportTab />, 
    export:  <ExportTab />, 
  };

  return (
    <div style={{
      display: "flex", flexDirection: "column",
      height: "100dvh", maxWidth: 540, margin: "0 auto",
      background: darkMode ? T.bg : '#F8FAFC', color: darkMode ? T.text : '#0F172A',
      fontFamily: "'Share Tech Mono', 'IBM Plex Mono', 'Courier New', monospace",
      position: "relative", overflow: "hidden",
      transition: 'background 0.2s, color 0.2s',
    }}>
      <Toast />
      <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>
        {/* Propagation du mode sombre/clair */}
        {pages[activeTab]}
      </div>
      <BottomNav darkMode={darkMode} setDarkMode={setDarkMode} />
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
