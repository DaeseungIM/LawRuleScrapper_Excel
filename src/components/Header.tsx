import React from 'react';
import { Key, FileSpreadsheet, Download, ShieldCheck } from 'lucide-react';

interface HeaderProps {
  ocKey: string;
  onOpenOcKeyModal: () => void;
}

export const Header: React.FC<HeaderProps> = ({
  ocKey,
  onOpenOcKeyModal,
}) => {
  return (
    <header className="bg-white text-slate-900 border-b border-slate-200 sticky top-0 z-30 shadow-xs">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
        {/* Brand & Logo */}
        <div className="flex items-center space-x-3">
          <div className="bg-emerald-600 p-2 rounded-xl text-white shadow-sm flex items-center justify-center">
            <FileSpreadsheet className="w-5 h-5" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h1 className="text-base sm:text-lg font-black tracking-tight text-slate-900">
                관세법령 & 행정규칙 개정연혁 엑셀 다운로드
              </h1>
              <span className="hidden sm:inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-bold bg-emerald-50 text-emerald-700 border border-emerald-200/80">
                <Download className="w-3 h-3" />
                XLSX / ZIP 지원
              </span>
            </div>
            <p className="text-xs text-slate-500 font-medium">
              국가법령정보 오픈API (법령 & 행정규칙) 개정연혁 실시간 조회 및 엑셀(XLSX) 내보내기
            </p>
          </div>
        </div>

        {/* Right Action Bar */}
        <div className="flex items-center space-x-2 sm:space-x-3">
          <div className="hidden lg:flex items-center gap-1.5 text-xs text-slate-500 font-medium mr-2">
            <ShieldCheck className="w-4 h-4 text-emerald-600" />
            <span>국가법령정보포털 공식 연계</span>
          </div>

          {/* OC Key Config Button */}
          <button
            onClick={onOpenOcKeyModal}
            className="flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-semibold bg-slate-50 hover:bg-slate-100 text-slate-700 border border-slate-200 transition-colors shadow-2xs"
            title="국가법령 Open API 키 설정"
          >
            <Key className="w-3.5 h-3.5 text-amber-500" />
            <span className="hidden md:inline">API 인증키:</span>
            <span className="font-mono text-indigo-600 font-bold">{ocKey}</span>
          </button>
        </div>
      </div>
    </header>
  );
};
