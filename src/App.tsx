import React, { useState } from 'react';
import { Header } from './components/Header';
import { OcKeyModal } from './components/OcKeyModal';
import { UnifiedSearchAndDriveExporter } from './components/UnifiedSearchAndDriveExporter';

export default function App() {
  const [ocKey, setOcKey] = useState('ceiai_law_test');
  const [isOcKeyModalOpen, setIsOcKeyModalOpen] = useState(false);

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans antialiased flex flex-col selection:bg-emerald-600 selection:text-white">
      {/* Top Header */}
      <Header
        ocKey={ocKey}
        onOpenOcKeyModal={() => setIsOcKeyModalOpen(true)}
      />

      {/* Main Content Area: Dedicated Revision History & Excel Exporter */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6 space-y-6">
        <UnifiedSearchAndDriveExporter
          ocKey={ocKey}
          onOpenOcKeyModal={() => setIsOcKeyModalOpen(true)}
        />
      </main>

      {/* Footer */}
      <footer className="border-t border-slate-200 py-6 mt-12 bg-white text-slate-500 text-xs">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 flex flex-col sm:flex-row items-center justify-between gap-4">
          <p>© 2026 대한민국 관세법령 & 행정규칙 개정연혁 엑셀 다운로드 자동화 시스템</p>
          <div className="flex items-center space-x-4">
            <a
              href="https://open.law.go.kr/LSO/usr/usrOcInfoMod.do"
              target="_blank"
              rel="noopener noreferrer"
              className="hover:text-emerald-700 transition-colors font-medium"
            >
              국가법령 Open API 센터
            </a>
            <span>•</span>
            <button
              onClick={() => setIsOcKeyModalOpen(true)}
              className="hover:text-emerald-700 transition-colors font-mono font-bold text-slate-700 cursor-pointer"
            >
              API Key ({ocKey})
            </button>
          </div>
        </div>
      </footer>

      {/* OC Key Modal */}
      <OcKeyModal
        isOpen={isOcKeyModalOpen}
        onClose={() => setIsOcKeyModalOpen(false)}
        currentOcKey={ocKey}
        onSaveOcKey={setOcKey}
      />
    </div>
  );
}
