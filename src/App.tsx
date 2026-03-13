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


function IecPage() {
  return <ExcelCleaner dark={false} onDarkToggle={() => {}} onSendToPointage={() => {}} />;
}


import { useEffect } from "react";

function AppInner() {
  const { activeTab, setActiveTab } = useApp();

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
      background: T.bg, color: T.text,
      fontFamily: "'Share Tech Mono', 'IBM Plex Mono', 'Courier New', monospace",
      position: "relative", overflow: "hidden",
    }}>
      <Toast />
      <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>
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
