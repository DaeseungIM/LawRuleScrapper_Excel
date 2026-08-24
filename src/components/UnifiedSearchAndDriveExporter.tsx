import React, { useState, useEffect, useMemo, useRef } from 'react';
import JSZip from 'jszip';
import {
  Search,
  BookOpen,
  FileSpreadsheet,
  CheckCircle2,
  AlertCircle,
  Loader2,
  ChevronRight,
  Sparkles,
  RotateCcw,
  CheckSquare,
  Square,
  Filter,
  Download,
  FolderArchive,
  RefreshCw,
  RotateCw,
  HelpCircle,
  Clock,
  ArrowUpDown,
  X,
  ArrowLeft,
  Check,
  FileCheck,
  ListOrdered,
  ListCheck,
  Layers,
  FileText,
  SlidersHorizontal
} from 'lucide-react';
import {
  SearchTargetType,
  LawSubType,
  SearchMatchMode,
  SearchWorkflowStep,
  UnifiedSearchItem,
  UnifiedRevisionItem,
} from '../types';
import { safeFetchJson } from '../lib/apiHelper';

export function formatElapsedDuration(seconds?: number): string {
  if (!seconds || seconds <= 0) return '0초';
  const hrs = Math.floor(seconds / 3600);
  const mins = Math.floor((seconds % 3600) / 60);
  const secs = seconds % 60;
  if (hrs > 0) {
    return `${hrs}시간 ${mins}분 ${secs}초`;
  }
  if (mins > 0) {
    return `${mins}분 ${secs}초`;
  }
  return `${secs}초`;
}

interface UnifiedSearchAndDriveExporterProps {
  ocKey: string;
  onOpenOcKeyModal: () => void;
}

export const UnifiedSearchAndDriveExporter: React.FC<UnifiedSearchAndDriveExporterProps> = ({
  ocKey,
  onOpenOcKeyModal,
}) => {
  // 1. 유형 선택 필터 상태
  const [targetType, setTargetType] = useState<SearchTargetType>('law');
  const [selectedSubTypes, setSelectedSubTypes] = useState<Record<LawSubType, boolean>>({
    law: true,
    decree: true,
    rule: true,
  });

  // 2. 키워드 입력 및 조회 옵션 상태
  const [searchKeyword, setSearchKeyword] = useState<string>('관세법');
  const [searchError, setSearchError] = useState<string | null>(null);
  const [matchMode, setMatchMode] = useState<SearchMatchMode>('exact');
  const [isSearching, setIsSearching] = useState<boolean>(false);
  const [hasSearched, setHasSearched] = useState<boolean>(false);

  // 2단계 검색 프로세스 상태 (1단계: 법령 선택, 2단계: 개정목록 조회)
  const [searchStep, setSearchStep] = useState<SearchWorkflowStep>('step1_select_law');
  const [candidateLaws, setCandidateLaws] = useState<UnifiedSearchItem[]>([]);
  const [selectedCandidateLaws, setSelectedCandidateLaws] = useState<Record<string, boolean>>({});
  const [isLoadingRevisions, setIsLoadingRevisions] = useState<boolean>(false);
  const [lastSearchMeta, setLastSearchMeta] = useState<{
    keyword: string;
    matchMode: SearchMatchMode;
    targetType: SearchTargetType;
    count: number;
    matchedAliasNote?: string;
  } | null>(null);

  // 3. 개정 목록 라인 단위 시각화 상태
  const [revisions, setRevisions] = useState<UnifiedRevisionItem[]>([]);
  const [selectedRevisions, setSelectedRevisions] = useState<Record<string, boolean>>({});
  const [filterQuery, setFilterQuery] = useState<string>('');
  const [ruleTypeFilter, setRuleTypeFilter] = useState<string>('all');
  const [sortField, setSortField] = useState<'hierarchy' | 'name' | 'enforcementDate' | 'promulgationDate' | 'promulgationNo'>('hierarchy');
  const [sortOrder, setSortOrder] = useState<'asc' | 'desc'>('desc');

  // 4. 폴더명 및 다운로드 상태
  const [customFolderName, setCustomFolderName] = useState<string>('');
  const [isCustomFolderEdited, setIsCustomFolderEdited] = useState<boolean>(false);

  // 5. 엑셀(.xlsx / .zip) 다운로드 상태
  const [isDownloadingExcel, setIsDownloadingExcel] = useState<boolean>(false);
  const [excelElapsedSeconds, setExcelElapsedSeconds] = useState<number>(0);
  const excelTimerRef = useRef<NodeJS.Timeout | null>(null);
  const [excelDownloadStatus, setExcelDownloadStatus] = useState<{
    message: string;
    total: number;
    current?: number;
    percent?: number;
  } | null>(null);
  const [downloadSuccessNotice, setDownloadSuccessNotice] = useState<{
    filename: string;
    itemCount: number;
    folderName?: string;
    elapsedSeconds?: number;
  } | null>(null);

  // Clean up timers on unmount
  useEffect(() => {
    return () => {
      if (excelTimerRef.current) clearInterval(excelTimerRef.current);
    };
  }, []);

  // Helper to normalize law title spacing
  function normalizeLawTitleSpacing(title: string): string {
    if (!title) return title;
    return title
      .replace(/([^\s]+)시행령(?:\b|$)/g, '$1 시행령')
      .replace(/([^\s]+)시행규칙(?:\b|$)/g, '$1 시행규칙')
      .replace(/\s{2,}/g, ' ')
      .trim();
  }

  // Helper to compute folder name
  const computeDefaultFolderName = (
    keyword: string,
    target: SearchTargetType,
    subTypes: Record<LawSubType, boolean>
  ): string => {
    const rawKeyword = normalizeLawTitleSpacing(keyword.trim());
    if (!rawKeyword) {
      return '';
    }

    const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');

    if (target === 'admrul') {
      return `[${rawKeyword.replace(/[\/\\:*?"<>|]/g, '_')}_개정목록_${today}]`;
    }

    let cleanBase = rawKeyword;
    if (cleanBase.endsWith(' 시행령') || (cleanBase.endsWith('시행령') && cleanBase.length > 3)) {
      cleanBase = cleanBase.replace(/\s*시행령$/, '');
    } else if (cleanBase.endsWith(' 시행규칙') || (cleanBase.endsWith('시행규칙') && cleanBase.length > 4)) {
      cleanBase = cleanBase.replace(/\s*시행규칙$/, '');
    } else if (cleanBase.endsWith(' 법률') || cleanBase.endsWith('법률')) {
      cleanBase = cleanBase.replace(/\s*법률$/, '');
    }
    cleanBase = normalizeLawTitleSpacing(cleanBase);

    const isAllSelected = subTypes.law && subTypes.decree && subTypes.rule;
    const isOnlyDecree = !subTypes.law && subTypes.decree && !subTypes.rule;
    const isOnlyRule = !subTypes.law && !subTypes.decree && subTypes.rule;
    const isDecreeAndRule = !subTypes.law && subTypes.decree && !subTypes.rule;
    const isOnlyLaw = subTypes.law && !subTypes.decree && !subTypes.rule;
    const isLawAndDecree = subTypes.law && subTypes.decree && !subTypes.rule;
    const isLawAndRule = subTypes.law && !subTypes.decree && subTypes.rule;

    let qualifiedName = cleanBase;
    if (isAllSelected) {
      qualifiedName = cleanBase;
    } else if (isOnlyDecree) {
      qualifiedName = `${cleanBase} 시행령`;
    } else if (isOnlyRule) {
      qualifiedName = `${cleanBase} 시행규칙`;
    } else if (isDecreeAndRule) {
      qualifiedName = `${cleanBase} 시행령·시행규칙`;
    } else if (isOnlyLaw) {
      qualifiedName = `${cleanBase} 법률`;
    } else if (isLawAndDecree) {
      qualifiedName = `${cleanBase} 법·시행령`;
    } else if (isLawAndRule) {
      qualifiedName = `${cleanBase} 법·시행규칙`;
    }

    return `[${qualifiedName.replace(/[\/\\:*?"<>|]/g, '_')}_개정목록_${today}]`;
  };

  // Auto-generate default folder name based on search & subType checkboxes if not manually edited
  useEffect(() => {
    if (!isCustomFolderEdited) {
      if (searchKeyword.trim()) {
        const generated = computeDefaultFolderName(searchKeyword, targetType, selectedSubTypes);
        setCustomFolderName(generated);
      } else {
        setCustomFolderName('');
      }
    }
  }, [searchKeyword, targetType, selectedSubTypes, isCustomFolderEdited]);

  // Reset folder name to auto computed name
  const handleResetFolderName = () => {
    const defaultKeyword = searchKeyword.trim() || (revisions[0]?.name) || (targetType === 'admrul' ? '행정규칙' : '관세법');
    const generated = computeDefaultFolderName(defaultKeyword, targetType, selectedSubTypes);
    setCustomFolderName(generated);
    setIsCustomFolderEdited(false);
  };

  // SubType checkbox toggle
  const handleToggleSubType = (subType: LawSubType) => {
    setSelectedSubTypes((prev) => {
      const next = { ...prev, [subType]: !prev[subType] };
      const hasAny = Object.values(next).some(Boolean);
      if (!hasAny) return prev;
      return next;
    });
  };

  // Build comma-separated subTypes string
  const getSubTypesParam = (): string => {
    const active: string[] = [];
    if (selectedSubTypes.law) active.push('law');
    if (selectedSubTypes.decree) active.push('decree');
    if (selectedSubTypes.rule) active.push('rule');
    return active.join(',') || 'law,decree,rule';
  };

  // Helper to parse numeric date for sorting
  const parseDateNum = (dateStr: any, fallbackDateStr?: any): number => {
    if (!dateStr) {
      if (fallbackDateStr) return parseDateNum(fallbackDateStr);
      return 0;
    }
    const str = String(dateStr).trim();
    if (str.includes('9999') || str.includes('미정') || str.toLowerCase().includes('unknown')) {
      if (fallbackDateStr) return parseDateNum(fallbackDateStr);
      return 0;
    }
    const parts = str.match(/\d+/g);
    if (parts && parts.length >= 3) {
      const yyyy = parts[0].padStart(4, '20');
      const mm = parts[1].padStart(2, '0');
      const dd = parts[2].padStart(2, '0');
      return parseInt(`${yyyy}${mm}${dd}`, 10) || 0;
    }
    const clean = str.replace(/[^0-9]/g, '');
    if (clean.length >= 8) return parseInt(clean.slice(0, 8), 10) || 0;
    if (clean.length >= 4) return parseInt(clean.padEnd(8, '0'), 10) || 0;
    if (fallbackDateStr) return parseDateNum(fallbackDateStr);
    return 0;
  };

  // Hierarchy rank: 1: 법 -> 2: 시행령 -> 3: 시행규칙 -> 4: 행정규칙
  const getHierarchyRank = (item: UnifiedRevisionItem): number => {
    if (!item) return 99;
    if (item.targetType === 'admrul') return 4;
    const subType = (item.subType || '').toLowerCase();
    const name = (item.name || '').trim();
    const ruleType = (item.ruleType || (item as any).lawType || '').trim();

    if (
      subType === 'law' ||
      ruleType === '법률' ||
      (!name.includes('시행령') && !name.includes('시행규칙') && !ruleType.includes('대통령령') && !ruleType.includes('부령'))
    ) {
      if (!name.includes('시행령') && !name.includes('시행규칙')) return 1;
    }
    if (subType === 'decree' || name.includes('시행령') || ruleType.includes('대통령령')) {
      return 2;
    }
    if (subType === 'rule' || name.includes('시행규칙') || ruleType.includes('부령') || ruleType.includes('총리령')) {
      return 3;
    }
    return 1;
  };

  // Master revision sorting function
  const sortRevisionsByHierarchyAndDate = (list: UnifiedRevisionItem[]): UnifiedRevisionItem[] => {
    if (!Array.isArray(list)) return [];
    return [...list].sort((a, b) => {
      const rankA = getHierarchyRank(a);
      const rankB = getHierarchyRank(b);
      if (rankA !== rankB) {
        return rankA - rankB; // Ascending: 법(1) -> 시행령(2) -> 시행규칙(3) -> 행정규칙(4)
      }

      const promA = parseDateNum(a.promulgationDate);
      const promB = parseDateNum(b.promulgationDate);

      const dateA = parseDateNum(a.enforcementDate, a.promulgationDate);
      const dateB = parseDateNum(b.enforcementDate, b.promulgationDate);
      if (dateB !== dateA) return dateB - dateA;

      if (promB !== promA) return promB - promA;

      const noA = parseInt(String(a.promulgationNo || a.seq || a.id || a.lawMst || '').replace(/[^0-9]/g, ''), 10) || 0;
      const noB = parseInt(String(b.promulgationNo || b.seq || b.id || b.lawMst || '').replace(/[^0-9]/g, ''), 10) || 0;
      return noB - noA;
    });
  };

  // Unique key generator for revision items
  const getRevKey = (item: UnifiedRevisionItem) => {
    const safeId = item.id || item.seq || item.lawMst || item.lawId || '';
    const safeName = (item.name || '').replace(/\s+/g, '');
    const safeEnf = (item.enforcementDate || '').replace(/\D/g, '');
    const safePromD = (item.promulgationDate || '').replace(/\D/g, '');
    const safePromN = (item.promulgationNo || '').replace(/\s+/g, '');
    return `${item.targetType || 'law'}_${safeId}_${safeName}_${safeEnf}_${safePromD}_${safePromN}`;
  };

  // Trigger browser file download safely to user's PC
  const triggerBrowserDownload = (blob: Blob, filename: string, count = 1, elapsedSeconds?: number) => {
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();

    const finalDuration = elapsedSeconds !== undefined && elapsedSeconds > 0 ? elapsedSeconds : (excelElapsedSeconds || 1);

    setDownloadSuccessNotice({
      filename,
      itemCount: count,
      elapsedSeconds: finalDuration,
    });
    setTimeout(() => {
      setDownloadSuccessNotice((prev) => (prev?.filename === filename ? null : prev));
    }, 12000);

    setTimeout(() => {
      if (document.body.contains(a)) {
        document.body.removeChild(a);
      }
      window.URL.revokeObjectURL(url);
    }, 60000);
  };

  // Candidate key generator for Step 1
  const getCandidateKey = (law: UnifiedSearchItem): string => {
    return law.id || `${law.name}_${law.subType || law.ruleType || ''}_${law.promulgationDate}_${law.promulgationNo}`;
  };

  // Step 1: Candidate selection toggles
  const handleToggleCandidateLaw = (key: string) => {
    setSelectedCandidateLaws((prev) => ({
      ...prev,
      [key]: !prev[key],
    }));
  };

  const handleSelectAllCandidates = (select: boolean) => {
    const next: Record<string, boolean> = {};
    if (select) {
      candidateLaws.forEach((l) => {
        next[getCandidateKey(l)] = true;
      });
    }
    setSelectedCandidateLaws(next);
  };

  const handleSelectSubTypeCandidates = (subType: LawSubType) => {
    const next: Record<string, boolean> = { ...selectedCandidateLaws };
    candidateLaws.forEach((l) => {
      if (l.subType === subType) {
        next[getCandidateKey(l)] = true;
      }
    });
    setSelectedCandidateLaws(next);
  };

  // Step 1: Execute Law Search
  const handleSearch = async (overrideKeyword?: string) => {
    const q = (overrideKeyword !== undefined ? overrideKeyword : searchKeyword).trim();
    if (!q) {
      setSearchError('법령명 또는 행정규칙명을 입력하세요.');
      return;
    }
    setSearchError(null);
    setHasSearched(true);
    setIsSearching(true);
    setCandidateLaws([]);
    setSelectedCandidateLaws({});
    setRevisions([]);
    setSelectedRevisions({});
    setSearchStep('step1_select_law');

    try {
      const subTypesParam = getSubTypesParam();
      const endpoint = `/api/unified/search?ocKey=${encodeURIComponent(
        ocKey
      )}&targetType=${targetType}&query=${encodeURIComponent(
        q
      )}&subTypes=${encodeURIComponent(subTypesParam)}&matchMode=${matchMode}`;

      const data = await safeFetchJson<any>(endpoint);

      if (data.success) {
        const rawList: UnifiedSearchItem[] = data.results || [];
        setCandidateLaws(rawList);
        setLastSearchMeta({
          keyword: q,
          matchMode,
          targetType,
          count: rawList.length,
          matchedAliasNote: data.matchedAliasNote,
        });

        // Default: select all candidate laws found
        const initialSelected: Record<string, boolean> = {};
        rawList.forEach((law) => {
          initialSelected[getCandidateKey(law)] = true;
        });
        setSelectedCandidateLaws(initialSelected);
      } else {
        setSearchError(data.error || '법령 검색에 실패했습니다.');
      }
    } catch (err: any) {
      console.error('Search candidate laws error:', err);
      setSearchError(err.message || '검색 중 오류가 발생했습니다.');
    } finally {
      setIsSearching(false);
    }
  };

  // Step 2: Fetch and display revision history for selected laws
  const handleProceedToStep2 = async () => {
    const selectedLaws = candidateLaws.filter((law) => !!selectedCandidateLaws[getCandidateKey(law)]);

    if (selectedLaws.length === 0) {
      alert('개정 연혁을 조회할 법령을 최소 1개 이상 체크해 주세요.');
      return;
    }

    setIsLoadingRevisions(true);
    setSearchStep('step2_view_revisions');
    setRevisions([]);
    setSelectedRevisions({});

    try {
      const selectedNames = Array.from(new Set(selectedLaws.map((l) => l.name.trim()).filter(Boolean)));
      const subTypesParam = getSubTypesParam();
      const namesParam = selectedNames.join(',');
      const endpoint = `/api/unified/revisions?ocKey=${encodeURIComponent(
        ocKey
      )}&targetType=${targetType}&names=${encodeURIComponent(
        namesParam
      )}&subTypes=${encodeURIComponent(subTypesParam)}&matchMode=exact`;

      const data = await safeFetchJson<any>(endpoint);

      if (data.success) {
        const rawList: UnifiedRevisionItem[] = data.revisions || [];
        const revList = sortRevisionsByHierarchyAndDate(rawList);
        setRevisions(revList);

        // Default: select all revisions
        const initialSelected: Record<string, boolean> = {};
        revList.forEach((r) => {
          initialSelected[getRevKey(r)] = true;
        });
        setSelectedRevisions(initialSelected);
      } else {
        alert(`개정 연혁 조회 실패: ${data.error || '알 수 없는 오류'}`);
      }
    } catch (err: any) {
      console.error('Fetch revisions error:', err);
      alert(`개정 연혁 조회 중 오류: ${err.message || '네트워크 오류'}`);
    } finally {
      setIsLoadingRevisions(false);
    }
  };

  const handleBackToStep1 = () => {
    setSearchStep('step1_select_law');
  };

  // Direct Excel Preset / Selected / All Download handler (packaged progressively via JSZip)
  const handleDownloadExcelZip = async (mode: 'preset' | 'selected' | 'all', presetCount?: number) => {
    let targetList: UnifiedRevisionItem[] = [];
    const cleanLawName = (revisions[0]?.name || searchKeyword || (targetType === 'admrul' ? '행정규칙' : '관세법')).trim();

    if (mode === 'preset') {
      const count = presetCount || 5;
      targetList = filteredRevisions.slice(0, count);
    } else if (mode === 'selected') {
      targetList = filteredRevisions.filter((r) => selectedRevisions[getRevKey(r)]);
    } else {
      targetList = filteredRevisions;
    }

    if (targetList.length === 0) {
      alert('다운로드할 개정연혁 항목이 없습니다.');
      return;
    }

    const sortedTargetList = sortRevisionsByHierarchyAndDate(targetList);
    const finalFolderName = (customFolderName.trim() || computeDefaultFolderName(searchKeyword || cleanLawName, targetType, selectedSubTypes)).replace(/[\/\\:*?"<>|]/g, '_');

    const startTs = Date.now();
    setExcelElapsedSeconds(0);
    if (excelTimerRef.current) clearInterval(excelTimerRef.current);
    excelTimerRef.current = setInterval(() => {
      setExcelElapsedSeconds(Math.floor((Date.now() - startTs) / 1000));
    }, 1000);

    setIsDownloadingExcel(true);
    setExcelDownloadStatus({
      message: `총 ${sortedTargetList.length}건의 엑셀 파일 패키징을 준비 중입니다...`,
      total: sortedTargetList.length,
      current: 0,
      percent: 0,
    });

    try {
      const zip = new JSZip();
      const zipFolder = zip.folder(finalFolderName) || zip;
      const BATCH_SIZE = 12;

      for (let i = 0; i < sortedTargetList.length; i += BATCH_SIZE) {
        const chunk = sortedTargetList.slice(i, i + BATCH_SIZE);
        const currentProcessed = Math.min(i + chunk.length, sortedTargetList.length);
        const percent = Math.round((currentProcessed / sortedTargetList.length) * 100);

        setExcelDownloadStatus({
          message: `엑셀 파일 생성 및 압축 중... (${currentProcessed}/${sortedTargetList.length}건 완료)`,
          total: sortedTargetList.length,
          current: currentProcessed,
          percent,
        });

        // Fetch this batch of workbooks as base64
        const batchRes = await safeFetchJson<any>('/api/unified/export-excel-batch', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            ocKey,
            targetType,
            selectedItem: { name: cleanLawName },
            revisions: chunk,
            cleanLawName,
            targetFolderName: finalFolderName,
            includeMasterSummary: i === 0,
            allRevisions: sortedTargetList,
            startIndex: i,
          }),
        });

        if (!batchRes.success) {
          throw new Error(batchRes.error || `엑셀 배치 생성 실패 (항목 ${i + 1}~${currentProcessed})`);
        }

        if (Array.isArray(batchRes.files)) {
          for (const fileItem of batchRes.files) {
            if (fileItem.filename && fileItem.base64) {
              zipFolder.file(fileItem.filename, fileItem.base64, { base64: true });
            }
          }
        }

        if (i + BATCH_SIZE < sortedTargetList.length) {
          await new Promise((r) => setTimeout(r, 80));
        }
      }

      setExcelDownloadStatus({
        message: `ZIP 압축 파일 생성 완료! PC 다운로드 시작 중...`,
        total: sortedTargetList.length,
        current: sortedTargetList.length,
        percent: 100,
      });

      const zipBlob = await zip.generateAsync({
        type: 'blob',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 },
      });

      if (excelTimerRef.current) clearInterval(excelTimerRef.current);
      const finalExcelElapsed = Math.max(1, Math.floor((Date.now() - startTs) / 1000));
      setExcelElapsedSeconds(finalExcelElapsed);

      const zipFileName = `${finalFolderName}_(${sortedTargetList.length}건)_엑셀모음.zip`;
      triggerBrowserDownload(zipBlob, zipFileName, sortedTargetList.length, finalExcelElapsed);
    } catch (err: any) {
      if (excelTimerRef.current) clearInterval(excelTimerRef.current);
      console.error('Excel export error:', err);
      alert(`엑셀 다운로드 오류: ${err.message || '다운로드 중 오류가 발생했습니다.'}`);
    } finally {
      if (excelTimerRef.current) clearInterval(excelTimerRef.current);
      setIsDownloadingExcel(false);
      setExcelDownloadStatus(null);
    }
  };

  // Direct Single Revision Excel Download (.xlsx)
  const handleDownloadSingleRevisionExcel = async (rev: UnifiedRevisionItem, e: React.MouseEvent) => {
    e.stopPropagation();
    const cleanLawName = (rev.name || searchKeyword || (targetType === 'admrul' ? '행정규칙' : '관세법')).trim();

    const startTs = Date.now();
    setExcelElapsedSeconds(0);
    if (excelTimerRef.current) clearInterval(excelTimerRef.current);
    excelTimerRef.current = setInterval(() => {
      setExcelElapsedSeconds(Math.floor((Date.now() - startTs) / 1000));
    }, 1000);

    setIsDownloadingExcel(true);
    setExcelDownloadStatus({
      message: `[${rev.name}] ${rev.promulgationNo || '개정본'} 엑셀 파일 생성 중...`,
      total: 1,
      current: 0,
      percent: 50,
    });

    try {
      const response = await fetch('/api/unified/export-single-excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          ocKey,
          targetType: rev.targetType || targetType,
          revision: rev,
          cleanLawName,
        }),
      });

      if (!response.ok) {
        const errJson = await response.json().catch(() => ({}));
        throw new Error(errJson.error || '엑셀 파일 생성 중 오류가 발생했습니다.');
      }

      const blob = await response.blob();
      const enfFormatted = (rev.enforcementDate && rev.enforcementDate !== '시행미정' && !rev.enforcementDate.includes('9999'))
        ? `[시행 ${rev.enforcementDate.trim()}]`
        : '[시행미정]';
      const promNoFormatted = (rev.promulgationNo || '개정본').trim();
      const promDateFormatted = (rev.promulgationDate || '').trim();
      const revTypeFormatted = (rev.revisionType || '일부개정').trim();
      const promPartFormatted = promDateFormatted
        ? `[${promNoFormatted}, ${promDateFormatted}, ${revTypeFormatted}]`
        : `[${promNoFormatted}, ${revTypeFormatted}]`;
      const filename = `${(rev.name || cleanLawName).trim()} ${enfFormatted} ${promPartFormatted}.xlsx`.replace(/[\/\\:*?"<>|]/g, '_');

      if (excelTimerRef.current) clearInterval(excelTimerRef.current);
      const finalExcelElapsed = Math.max(1, Math.floor((Date.now() - startTs) / 1000));
      setExcelElapsedSeconds(finalExcelElapsed);

      triggerBrowserDownload(blob, filename, 1, finalExcelElapsed);
    } catch (err: any) {
      if (excelTimerRef.current) clearInterval(excelTimerRef.current);
      console.error('Single revision excel export error:', err);
      alert(`엑셀 다운로드 오류: ${err.message}`);
    } finally {
      if (excelTimerRef.current) clearInterval(excelTimerRef.current);
      setIsDownloadingExcel(false);
      setExcelDownloadStatus(null);
    }
  };

  // Standalone Buchik (.xlsx) Direct Download handler
  const handleDownloadBuchikExcel = async () => {
    if (!filteredRevisions || filteredRevisions.length === 0) {
      alert('다운로드할 개정연혁 데이터가 없습니다.');
      return;
    }
    const cleanLawName = (revisions[0]?.name || searchKeyword || (targetType === 'admrul' ? '행정규칙' : '관세법')).trim();

    const startTs = Date.now();
    setExcelElapsedSeconds(0);
    if (excelTimerRef.current) clearInterval(excelTimerRef.current);
    excelTimerRef.current = setInterval(() => {
      setExcelElapsedSeconds(Math.floor((Date.now() - startTs) / 1000));
    }, 1000);

    setIsDownloadingExcel(true);
    setExcelDownloadStatus({
      message: `'${cleanLawName}' 부칙 데이터를 추출하여 엑셀 파일 생성 중입니다...`,
      total: 1,
      current: 0,
      percent: 50,
    });

    try {
      const response = await fetch('/api/unified/export-buchik-excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          ocKey,
          targetType,
          cleanLawName,
          latestRevision: filteredRevisions[0],
        }),
      });

      if (!response.ok) {
        const errJson = await response.json().catch(() => ({}));
        throw new Error(errJson.error || '부칙 엑셀 파일 생성 중 오류가 발생했습니다.');
      }

      const blob = await response.blob();
      const filename = `000_[${cleanLawName}]_부칙.xlsx`.replace(/[\/\\:*?"<>|]/g, '_');

      if (excelTimerRef.current) clearInterval(excelTimerRef.current);
      const finalExcelElapsed = Math.max(1, Math.floor((Date.now() - startTs) / 1000));
      setExcelElapsedSeconds(finalExcelElapsed);

      triggerBrowserDownload(blob, filename, 1, finalExcelElapsed);
    } catch (err: any) {
      if (excelTimerRef.current) clearInterval(excelTimerRef.current);
      console.error('Buchik excel export error:', err);
      alert(`부칙 엑셀 다운로드 오류: ${err.message}`);
    } finally {
      if (excelTimerRef.current) clearInterval(excelTimerRef.current);
      setIsDownloadingExcel(false);
      setExcelDownloadStatus(null);
    }
  };

  // Toggle single revision checkbox
  const handleToggleRevision = (key: string) => {
    setSelectedRevisions((prev) => ({
      ...prev,
      [key]: !prev[key],
    }));
  };

  // Select all / Deselect all
  const handleSelectAll = (checked: boolean) => {
    const next: Record<string, boolean> = {};
    filteredRevisions.forEach((r) => {
      next[getRevKey(r)] = checked;
    });
    setSelectedRevisions(next);
  };

  // Quick preset selector
  const handleQuickSelectCount = (count: number) => {
    const next: Record<string, boolean> = {};
    filteredRevisions.slice(0, count).forEach((r) => {
      next[getRevKey(r)] = true;
    });
    setSelectedRevisions(next);
  };

  // Select by Subtype in revisions
  const handleSelectSubTypesRevisions = (subType: LawSubType) => {
    const next: Record<string, boolean> = { ...selectedRevisions };
    filteredRevisions.forEach((r) => {
      if (r.subType === subType) {
        next[getRevKey(r)] = true;
      }
    });
    setSelectedRevisions(next);
  };

  // Handler to toggle column sorting
  const handleSortToggle = (field: 'hierarchy' | 'name' | 'enforcementDate' | 'promulgationDate' | 'promulgationNo') => {
    if (sortField === field) {
      setSortOrder((prev) => (prev === 'asc' ? 'desc' : 'asc'));
    } else {
      setSortField(field);
      setSortOrder(field === 'name' ? 'asc' : 'desc');
    }
  };

  // Filtered & sorted revisions
  const filteredRevisions = useMemo(() => {
    const filtered = revisions.filter((rev) => {
      if (ruleTypeFilter !== 'all') {
        if (ruleTypeFilter === 'law' && rev.subType !== 'law') return false;
        if (ruleTypeFilter === 'decree' && rev.subType !== 'decree') return false;
        if (ruleTypeFilter === 'rule' && rev.subType !== 'rule') return false;
      }
      if (filterQuery.trim()) {
        const q = filterQuery.toLowerCase();
        const matchesName = rev.name.toLowerCase().includes(q);
        const matchesDate = (rev.enforcementDate || '').includes(q) || (rev.promulgationDate || '').includes(q);
        const matchesNo = (rev.promulgationNo || '').includes(q);
        const matchesType = (rev.revisionType || '').toLowerCase().includes(q);
        return matchesName || matchesDate || matchesNo || matchesType;
      }
      return true;
    });

    if (sortField === 'hierarchy') {
      const sorted = sortRevisionsByHierarchyAndDate(filtered);
      return sortOrder === 'asc' ? sorted.reverse() : sorted;
    }

    return [...filtered].sort((a, b) => {
      if (sortField === 'name') {
        const cmp = (a.name || '').localeCompare(b.name || '', 'ko-KR');
        if (cmp !== 0) return sortOrder === 'asc' ? cmp : -cmp;
        const dateA = parseDateNum(a.enforcementDate);
        const dateB = parseDateNum(b.enforcementDate);
        return dateB - dateA;
      }

      if (sortField === 'enforcementDate') {
        const dateA = parseDateNum(a.enforcementDate, a.promulgationDate);
        const dateB = parseDateNum(b.enforcementDate, b.promulgationDate);
        if (dateA !== dateB) return sortOrder === 'asc' ? dateA - dateB : dateB - dateA;
        const promA = parseDateNum(a.promulgationDate);
        const promB = parseDateNum(b.promulgationDate);
        return promB - promA;
      }

      if (sortField === 'promulgationDate') {
        const promA = parseDateNum(a.promulgationDate);
        const promB = parseDateNum(b.promulgationDate);
        if (promA !== promB) return sortOrder === 'asc' ? promA - promB : promB - promA;
      }

      if (sortField === 'promulgationNo') {
        const noA = parseInt(String(a.promulgationNo || a.seq || a.id || a.lawMst || '').replace(/[^0-9]/g, ''), 10) || 0;
        const noB = parseInt(String(b.promulgationNo || b.seq || b.id || b.lawMst || '').replace(/[^0-9]/g, ''), 10) || 0;
        if (noA !== noB) return sortOrder === 'asc' ? noA - noB : noB - noA;
      }

      return 0;
    });
  }, [revisions, ruleTypeFilter, filterQuery, sortField, sortOrder]);

  // Counts for Step 1
  const selectedCandidateCount = Object.values(selectedCandidateLaws).filter(Boolean).length;
  const isAllCandidatesSelected = candidateLaws.length > 0 && selectedCandidateCount === candidateLaws.length;
  const candidateLawCount = candidateLaws.filter((l) => l.subType === 'law' || l.ruleType === '법률').length;
  const candidateDecreeCount = candidateLaws.filter((l) => l.subType === 'decree' || l.ruleType?.includes('대통령령') || l.name.includes('시행령')).length;
  const candidateRuleCount = candidateLaws.filter((l) => l.subType === 'rule' || l.ruleType?.includes('부령') || l.name.includes('시행규칙')).length;

  // Counts for Step 2
  const selectedCount = Object.values(selectedRevisions).filter(Boolean).length;
  const isAllSelected = filteredRevisions.length > 0 && selectedCount >= filteredRevisions.length;
  const lawCount = revisions.filter((r) => r.subType === 'law' || r.ruleType === '법률').length;
  const decreeCount = revisions.filter((r) => r.subType === 'decree' || r.ruleType?.includes('대통령령') || r.name.includes('시행령')).length;
  const ruleCount = revisions.filter((r) => r.subType === 'rule' || r.ruleType?.includes('부령') || r.name.includes('시행규칙')).length;
  const admrulCount = revisions.filter((r) => r.targetType === 'admrul').length;

  return (
    <div className="space-y-6">
      {/* 1. Search Control Card */}
      <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs space-y-4">
        {/* Target Switcher & Search Bar */}
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          {/* Target Type Selector */}
          <div className="flex items-center space-x-1.5 bg-slate-100 p-1 rounded-xl border border-slate-200 shrink-0">
            <button
              id="target-type-law-btn"
              onClick={() => {
                setTargetType('law');
                setSearchStep('step1_select_law');
                setCandidateLaws([]);
                setRevisions([]);
              }}
              className={`px-4 py-2 rounded-lg text-xs font-bold transition-all flex items-center space-x-1.5 cursor-pointer ${
                targetType === 'law'
                  ? 'bg-emerald-600 text-white shadow-xs'
                  : 'text-slate-600 hover:text-slate-900 hover:bg-white/60'
              }`}
            >
              <BookOpen className="w-3.5 h-3.5" />
              <span>법령 (법·령·규칙)</span>
            </button>
            <button
              id="target-type-admrul-btn"
              onClick={() => {
                setTargetType('admrul');
                setSearchStep('step1_select_law');
                setCandidateLaws([]);
                setRevisions([]);
              }}
              className={`px-4 py-2 rounded-lg text-xs font-bold transition-all flex items-center space-x-1.5 cursor-pointer ${
                targetType === 'admrul'
                  ? 'bg-emerald-600 text-white shadow-xs'
                  : 'text-slate-600 hover:text-slate-900 hover:bg-white/60'
              }`}
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              <span>행정규칙 (고시·훈령·예규)</span>
            </button>
          </div>

          {/* SubType Checkboxes (for Laws) */}
          {targetType === 'law' && (
            <div className="flex items-center space-x-3 text-xs font-semibold text-slate-700 bg-slate-50 px-3.5 py-1.5 rounded-xl border border-slate-200">
              <span className="text-slate-400 font-bold text-[11px]">수집 범위:</span>
              <label className="flex items-center space-x-1.5 cursor-pointer hover:text-emerald-700">
                <input
                  type="checkbox"
                  checked={selectedSubTypes.law}
                  onChange={() => handleToggleSubType('law')}
                  className="rounded text-emerald-600 focus:ring-emerald-500 w-4 h-4 cursor-pointer"
                />
                <span>법률</span>
              </label>
              <label className="flex items-center space-x-1.5 cursor-pointer hover:text-emerald-700">
                <input
                  type="checkbox"
                  checked={selectedSubTypes.decree}
                  onChange={() => handleToggleSubType('decree')}
                  className="rounded text-emerald-600 focus:ring-emerald-500 w-4 h-4 cursor-pointer"
                />
                <span>대통령령(시행령)</span>
              </label>
              <label className="flex items-center space-x-1.5 cursor-pointer hover:text-emerald-700">
                <input
                  type="checkbox"
                  checked={selectedSubTypes.rule}
                  onChange={() => handleToggleSubType('rule')}
                  className="rounded text-emerald-600 focus:ring-emerald-500 w-4 h-4 cursor-pointer"
                />
                <span>부령(시행규칙)</span>
              </label>
            </div>
          )}

          {/* Search Mode Toggle */}
          <div className="flex items-center space-x-1.5 text-xs text-slate-600 bg-slate-50 px-3 py-1.5 rounded-xl border border-slate-200 shrink-0">
            <span className="text-slate-400 text-[11px]">검색 매칭:</span>
            <button
              onClick={() => setMatchMode('exact')}
              className={`px-2.5 py-1 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                matchMode === 'exact'
                  ? 'bg-white text-emerald-800 shadow-2xs border border-slate-200'
                  : 'text-slate-500 hover:text-slate-800'
              }`}
              title="검색어와 정확히 일치하는 법령/행정규칙을 검색합니다. (약칭/구명칭 자동 매핑)"
            >
              정확히 일치 (권장)
            </button>
            <button
              onClick={() => setMatchMode('contains')}
              className={`px-2.5 py-1 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                matchMode === 'contains'
                  ? 'bg-white text-emerald-800 shadow-2xs border border-slate-200'
                  : 'text-slate-500 hover:text-slate-800'
              }`}
              title="검색어가 포함된 모든 법령 및 행정규칙을 폭넓게 검색합니다."
            >
              포함 (확장)
            </button>
          </div>
        </div>

        {/* Input Bar & Actions */}
        <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2">
          <div className="relative flex-1">
            <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
            <input
              id="main-search-input"
              type="text"
              value={searchKeyword}
              onChange={(e) => setSearchKeyword(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') handleSearch();
              }}
              placeholder={
                targetType === 'law'
                  ? '법령명을 입력하세요 (예: 관세법, 외국환거래법, 환특법, 자유무역협정관세법 등)'
                  : '행정규칙명을 입력하세요 (예: 관세평가 운영에 관한 고시, 수출통관 사무처리에 관한 고시 등)'
              }
              className="w-full pl-10 pr-4 py-2.5 bg-slate-50 border border-slate-300 rounded-xl text-sm font-semibold focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-600 transition-all placeholder:text-slate-400"
            />
          </div>

          <button
            id="main-search-btn"
            onClick={() => handleSearch()}
            disabled={isSearching}
            className="px-6 py-2.5 bg-emerald-600 hover:bg-emerald-700 active:bg-emerald-800 text-white rounded-xl text-sm font-bold flex items-center justify-center space-x-2 shadow-sm transition-all disabled:opacity-50 shrink-0 cursor-pointer"
          >
            {isSearching ? (
              <>
                <Loader2 className="w-4 h-4 animate-spin" />
                <span>조회 중...</span>
              </>
            ) : (
              <>
                <Search className="w-4 h-4" />
                <span>법령 검색</span>
              </>
            )}
          </button>
        </div>

        {/* Quick Keyword Presets */}
        <div className="flex flex-wrap items-center gap-1.5 pt-1 text-xs text-slate-500">
          <span className="font-semibold text-slate-400 mr-1 flex items-center space-x-1">
            <Sparkles className="w-3.5 h-3.5 text-amber-500" />
            <span>추천 검색어:</span>
          </span>
          {targetType === 'law' ? (
            <>
              {['관세법', '외국환거래법', '수출용원재료관세환급특례법', '자유무역협정관세특례법', '대외무역법'].map((kw) => (
                <button
                  key={kw}
                  onClick={() => {
                    setSearchKeyword(kw);
                    handleSearch(kw);
                  }}
                  className="px-2.5 py-1 bg-slate-100 hover:bg-emerald-50 hover:text-emerald-800 hover:border-emerald-200 border border-slate-200 rounded-lg text-slate-700 transition-all cursor-pointer"
                >
                  {kw}
                </button>
              ))}
            </>
          ) : (
            <>
              {[
                '관세평가 운영에 관한 고시',
                '수출통관 사무처리에 관한 고시',
                '수입통관 사무처리에 관한 고시',
                '보세판매장 운영에 관한 고시',
                '외국환거래규정',
              ].map((kw) => (
                <button
                  key={kw}
                  onClick={() => {
                    setSearchKeyword(kw);
                    handleSearch(kw);
                  }}
                  className="px-2.5 py-1 bg-slate-100 hover:bg-emerald-50 hover:text-emerald-800 hover:border-emerald-200 border border-slate-200 rounded-lg text-slate-700 transition-all cursor-pointer"
                >
                  {kw}
                </button>
              ))}
            </>
          )}
        </div>

        {/* Error message */}
        {searchError && (
          <div className="p-3 bg-rose-50 text-rose-700 rounded-xl text-xs flex items-center justify-between border border-rose-200">
            <div className="flex items-center space-x-2">
              <AlertCircle className="w-4 h-4 shrink-0" />
              <span>{searchError}</span>
            </div>
            <button
              onClick={() => setSearchError(null)}
              className="text-rose-500 hover:text-rose-700 p-1 cursor-pointer"
            >
              <X className="w-3.5 h-3.5" />
            </button>
          </div>
        )}
      </div>

      {/* Live Excel Download Progress Bar */}
      {isDownloadingExcel && excelDownloadStatus && (
        <div className="p-4 bg-emerald-50 border border-emerald-300 rounded-2xl animate-fade-in shadow-sm space-y-2.5">
          <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2 text-xs font-bold text-emerald-950">
            <div className="flex items-center space-x-2">
              <Loader2 className="w-4 h-4 animate-spin text-emerald-700 shrink-0" />
              <span className="leading-snug">{excelDownloadStatus.message}</span>
            </div>
            <div className="flex items-center space-x-2.5 shrink-0">
              <span className="bg-emerald-200/90 text-emerald-950 px-2.5 py-1 rounded-full text-[11px] font-mono flex items-center space-x-1 border border-emerald-300">
                <Clock className="w-3 h-3 text-emerald-800" />
                <span>경과시간:</span>
                <strong>{formatElapsedDuration(excelElapsedSeconds)}</strong>
              </span>
              <span className="bg-emerald-600 text-white px-2.5 py-1 rounded-full text-[11px]">
                {excelDownloadStatus.percent !== undefined ? `${excelDownloadStatus.percent}%` : '패키징 중...'}
              </span>
            </div>
          </div>
          <div className="w-full bg-emerald-200/70 rounded-full h-2.5 overflow-hidden">
            <div
              className="bg-emerald-600 h-2.5 rounded-full transition-all duration-300"
              style={{ width: `${excelDownloadStatus.percent || 10}%` }}
            ></div>
          </div>
        </div>
      )}

      {/* Download Success Notice Banner */}
      {downloadSuccessNotice && (
        <div className="p-4 bg-emerald-50 border border-emerald-300 rounded-2xl animate-fade-in shadow-xs flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
          <div className="flex items-start sm:items-center space-x-3">
            <div className="p-2 rounded-xl bg-emerald-100 text-emerald-800 shrink-0 border border-emerald-300">
              <CheckCircle2 className="w-5 h-5" />
            </div>
            <div>
              <div className="flex flex-wrap items-center gap-2">
                <h4 className="text-sm font-extrabold text-emerald-950">
                  🎉 엑셀 파일 다운로드가 시작되었습니다!
                </h4>
                <span className="px-2 py-0.5 rounded-full text-[11px] font-bold bg-emerald-600 text-white">
                  {downloadSuccessNotice.itemCount}건 완료
                </span>
                <span className="px-2 py-0.5 rounded-full text-[11px] font-mono font-bold bg-white text-emerald-900 border border-emerald-300">
                  ⏱️ {formatElapsedDuration(downloadSuccessNotice.elapsedSeconds)}
                </span>
              </div>
              <p className="text-xs text-emerald-900 mt-1 font-medium">
                파일명: <strong className="font-mono bg-white px-1.5 py-0.5 rounded border border-emerald-200 text-emerald-950 font-bold">{downloadSuccessNotice.filename}</strong>
                <span className="ml-2 text-emerald-700">(내 PC [다운로드] 폴더에 저장됩니다)</span>
              </p>
            </div>
          </div>
          <button
            onClick={() => setDownloadSuccessNotice(null)}
            className="p-1.5 text-emerald-600 hover:text-emerald-900 hover:bg-emerald-100 rounded-lg transition-colors shrink-0 self-end sm:self-center cursor-pointer"
            title="알림 닫기"
          >
            <X className="w-4 h-4" />
          </button>
        </div>
      )}

      {/* Step 1: Candidate Laws Selection View */}
      {searchStep === 'step1_select_law' && (
        <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-xs">
          {/* Header */}
          <div className="p-4 bg-slate-50/90 border-b border-slate-200 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
            <div className="flex items-center space-x-2.5">
              <div className="w-6 h-6 rounded-full bg-emerald-600 text-white text-xs font-black flex items-center justify-center">
                1
              </div>
              <h3 className="text-sm font-extrabold text-slate-900">
                1단계: 대상 법령/행정규칙 선택
              </h3>
              {candidateLaws.length > 0 && (
                <span className="px-2.5 py-0.5 rounded-full text-xs font-bold bg-emerald-100 text-emerald-800 border border-emerald-200">
                  총 {candidateLaws.length}건 검색됨
                </span>
              )}
            </div>

            {candidateLaws.length > 0 && (
              <div className="flex flex-wrap items-center gap-2">
                <button
                  id="select-all-candidates-btn"
                  onClick={() => handleSelectAllCandidates(!isAllCandidatesSelected)}
                  className="px-2.5 py-1 text-xs font-semibold bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 rounded-lg transition-colors cursor-pointer"
                >
                  {isAllCandidatesSelected ? '전체 해제' : '전체 선택'}
                </button>
                {targetType === 'law' && (
                  <>
                    <button
                      onClick={() => handleSelectSubTypeCandidates('law')}
                      className="px-2.5 py-1 text-xs font-semibold bg-blue-50 border border-blue-200 text-blue-800 hover:bg-blue-100 rounded-lg transition-colors cursor-pointer"
                    >
                      법률만
                    </button>
                    <button
                      onClick={() => handleSelectSubTypeCandidates('decree')}
                      className="px-2.5 py-1 text-xs font-semibold bg-amber-50 border border-amber-200 text-amber-800 hover:bg-amber-100 rounded-lg transition-colors cursor-pointer"
                    >
                      시행령만
                    </button>
                    <button
                      onClick={() => handleSelectSubTypeCandidates('rule')}
                      className="px-2.5 py-1 text-xs font-semibold bg-emerald-50 border border-emerald-200 text-emerald-800 hover:bg-emerald-100 rounded-lg transition-colors cursor-pointer"
                    >
                      시행규칙만
                    </button>
                  </>
                )}

                <button
                  id="proceed-to-step2-btn"
                  onClick={handleProceedToStep2}
                  disabled={isLoadingRevisions || selectedCandidateCount === 0}
                  className="px-4 py-1.5 bg-emerald-600 hover:bg-emerald-700 active:bg-emerald-800 text-white rounded-lg text-xs font-bold flex items-center space-x-1.5 shadow-xs transition-all disabled:opacity-50 cursor-pointer"
                >
                  {isLoadingRevisions ? (
                    <>
                      <Loader2 className="w-3.5 h-3.5 animate-spin" />
                      <span>개정연혁 수집 중...</span>
                    </>
                  ) : (
                    <>
                      <span>개정연혁 조회 ({selectedCandidateCount}건)</span>
                      <ChevronRight className="w-4 h-4" />
                    </>
                  )}
                </button>
              </div>
            )}
          </div>

          {/* Alias / Predecessor Note */}
          {lastSearchMeta?.matchedAliasNote && (
            <div className="px-4 py-2.5 bg-amber-50 border-b border-amber-200 text-xs text-amber-900 flex items-center space-x-2">
              <Sparkles className="w-4 h-4 text-amber-600 shrink-0" />
              <span>{lastSearchMeta.matchedAliasNote}</span>
            </div>
          )}

          {/* Table / Empty State */}
          <div className="overflow-x-auto min-h-[220px]">
            {!hasSearched ? (
              <div className="py-16 px-6 flex flex-col items-center justify-center space-y-3 text-slate-400 text-center">
                <div className="w-12 h-12 rounded-2xl bg-emerald-50 flex items-center justify-center text-emerald-600 mb-1 border border-emerald-100">
                  <Search className="w-6 h-6" />
                </div>
                <p className="text-sm font-bold text-slate-700">
                  상단 검색창에 법령명(예: 관세법)을 입력하고 [법령 검색]을 누르세요.
                </p>
                <p className="text-xs text-slate-500 max-w-md">
                  법령 및 시행령, 시행규칙 또는 행정규칙(고시) 목록이 검색되며, 원하는 대상을 선택하여 개정 연혁을 전수 수집하고 엑셀로 내려받을 수 있습니다.
                </p>
              </div>
            ) : isSearching ? (
              <div className="py-16 flex flex-col items-center justify-center space-y-3 text-slate-400">
                <Loader2 className="w-8 h-8 animate-spin text-emerald-600" />
                <p className="text-sm font-semibold text-slate-600">
                  국가법령정보포털에서 법령 목록을 검색 중입니다...
                </p>
              </div>
            ) : candidateLaws.length === 0 ? (
              <div className="py-16 flex flex-col items-center justify-center space-y-2 text-slate-400">
                <BookOpen className="w-8 h-8 text-slate-300" />
                <p className="text-sm font-semibold text-slate-600">
                  검색 결과와 일치하는 법령 또는 행정규칙이 없습니다.
                </p>
                <p className="text-xs text-slate-400">
                  키워드를 확인하시거나 [포함 (확장)] 검색 모드로 전환해 보세요.
                </p>
              </div>
            ) : (
              <table className="w-full text-left text-xs border-collapse">
                <thead className="bg-slate-100/80 sticky top-0 z-10 border-b border-slate-200 text-slate-600 font-bold">
                  <tr>
                    <th className="p-3 w-10 text-center">선택</th>
                    <th className="p-3 w-28">구분</th>
                    <th className="p-3 min-w-[220px]">법령/행정규칙명</th>
                    <th className="p-3 w-28">시행일자</th>
                    <th className="p-3 w-28">공포(발령)일자</th>
                    <th className="p-3 min-w-[140px]">공포번호</th>
                    <th className="p-3 w-28">소관부처</th>
                    <th className="p-3 w-24">법령ID</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200 bg-white">
                  {candidateLaws.map((law) => {
                    const key = getCandidateKey(law);
                    const isChecked = !!selectedCandidateLaws[key];

                    const isDecree = law.subType === 'decree' || law.name.includes('시행령') || law.ruleType?.includes('대통령령');
                    const isRule = law.subType === 'rule' || law.name.includes('시행규칙') || law.ruleType?.includes('부령');
                    const isAdmrul = targetType === 'admrul' || law.targetType === 'admrul';

                    let badgeClass = 'bg-blue-100 text-blue-800 border-blue-200';
                    let badgeText = law.ruleType || '법률';
                    if (isDecree) {
                      badgeClass = 'bg-amber-100 text-amber-800 border-amber-200';
                      badgeText = law.ruleType || '대통령령(시행령)';
                    } else if (isRule) {
                      badgeClass = 'bg-emerald-100 text-emerald-800 border-emerald-200';
                      badgeText = law.ruleType || '부령(시행규칙)';
                    } else if (isAdmrul) {
                      badgeClass = 'bg-purple-100 text-purple-800 border-purple-200';
                      badgeText = law.ruleType || '고시';
                    }

                    return (
                      <tr
                        key={key}
                        onClick={() => handleToggleCandidateLaw(key)}
                        className={`cursor-pointer transition-colors ${
                          isChecked ? 'bg-emerald-50/40 hover:bg-emerald-50/70' : 'hover:bg-slate-50'
                        }`}
                      >
                        <td className="p-3 text-center" onClick={(e) => e.stopPropagation()}>
                          <input
                            type="checkbox"
                            checked={isChecked}
                            onChange={() => handleToggleCandidateLaw(key)}
                            className="rounded text-emerald-600 focus:ring-emerald-500 w-4 h-4 cursor-pointer"
                          />
                        </td>
                        <td className="p-3">
                          <span className={`px-2 py-0.5 rounded-full text-[11px] font-bold border inline-block ${badgeClass}`}>
                            {badgeText}
                          </span>
                        </td>
                        <td className="p-3 font-bold text-slate-900">
                          <div className="flex items-center space-x-1.5">
                            <span>{law.name}</span>
                            {law.isPredecessor && (
                              <span className="px-1.5 py-0.2 rounded text-[10px] font-bold bg-amber-100 text-amber-800 border border-amber-300">
                                변경전 명칭
                              </span>
                            )}
                          </div>
                        </td>
                        <td className="p-3 font-semibold text-emerald-700">
                          {law.enforcementDate || '-'}
                        </td>
                        <td className="p-3 text-slate-700">
                          {law.promulgationDate || '-'}
                        </td>
                        <td className="p-3 text-slate-700">
                          {law.promulgationNo || '-'}
                        </td>
                        <td className="p-3 text-slate-600">
                          {law.department || '기획재정부'}
                        </td>
                        <td className="p-3 text-slate-400 font-mono text-[11px]">
                          {law.id || '-'}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            )}
          </div>

          {/* Bottom Step 1 Action Bar */}
          {candidateLaws.length > 0 && (
            <div className="p-4 bg-slate-50 border-t border-slate-200 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
              <div className="text-xs text-slate-600">
                선택된 <strong>{selectedCandidateCount}개</strong>의 법령에 대한 전체 개정 연혁을 수집하여 2단계 목록으로 표시합니다.
              </div>
              <button
                id="bottom-proceed-to-step2-btn"
                onClick={handleProceedToStep2}
                disabled={isLoadingRevisions || selectedCandidateCount === 0}
                className="px-5 py-2 bg-emerald-600 hover:bg-emerald-700 active:bg-emerald-800 text-white rounded-xl text-xs font-bold flex items-center justify-center space-x-2 shadow-sm transition-all disabled:opacity-50 cursor-pointer"
              >
                {isLoadingRevisions ? (
                  <>
                    <Loader2 className="w-4 h-4 animate-spin" />
                    <span>개정 연혁 전수 수집 중...</span>
                  </>
                ) : (
                  <>
                    <span>선택 법령 개정연혁 조회 및 엑셀 다운로드 (2단계 이동)</span>
                    <ChevronRight className="w-4 h-4" />
                  </>
                )}
              </button>
            </div>
          )}
        </div>
      )}

      {/* Step 2: Revisions Table & Direct Excel Download View */}
      {searchStep === 'step2_view_revisions' && (
        <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-xs space-y-0">
          {/* Header Bar */}
          <div className="p-4 bg-slate-50/90 border-b border-slate-200 flex flex-col lg:flex-row lg:items-center lg:justify-between gap-3">
            <div className="flex items-center space-x-3">
              <button
                id="back-to-step1-btn"
                onClick={handleBackToStep1}
                className="p-1.5 bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 rounded-lg text-xs font-bold flex items-center space-x-1 transition-colors cursor-pointer"
                title="1단계 대상 법령 목록으로 돌아가기"
              >
                <ArrowLeft className="w-4 h-4" />
                <span>법령 다시 선택</span>
              </button>

              <div className="h-4 w-[1px] bg-slate-300 hidden sm:block"></div>

              <div>
                <div className="flex items-center space-x-2">
                  <h3 className="text-sm font-extrabold text-slate-900">
                    2단계: 개정연혁 목록 및 엑셀 다운로드
                  </h3>
                  <span className="px-2.5 py-0.5 rounded-full text-xs font-bold bg-emerald-100 text-emerald-800 border border-emerald-300">
                    총 {revisions.length}건 수집 완료
                  </span>
                </div>
                <p className="text-xs text-slate-500 mt-0.5">
                  수집된 개정본을 선택하여 일괄 ZIP 압축 다운로드하거나, 행별로 즉시 개별 .xlsx 파일로 다운로드할 수 있습니다.
                </p>
              </div>
            </div>

            {/* Excel Download Action Buttons */}
            <div className="flex flex-wrap items-center gap-2">
              {/* Standalone Buchik (.xlsx) Download Button */}
              <button
                id="download-buchik-excel-btn"
                onClick={handleDownloadBuchikExcel}
                disabled={isDownloadingExcel || filteredRevisions.length === 0}
                className="px-3.5 py-2 bg-blue-600 hover:bg-blue-700 active:bg-blue-800 text-white rounded-xl text-xs font-bold flex items-center space-x-1.5 shadow-sm transition-all disabled:opacity-50 cursor-pointer"
                title="최신 개정본 기준 부칙(附則) 조항 전용 엑셀 파일 (.xlsx) 즉시 다운로드"
              >
                <FileText className="w-3.5 h-3.5" />
                <span>부칙 단독 엑셀</span>
              </button>

              {/* Selected Excel ZIP Download Button */}
              <button
                id="download-selected-excel-btn"
                onClick={() => handleDownloadExcelZip('selected')}
                disabled={isDownloadingExcel || selectedCount === 0}
                className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 active:bg-emerald-800 text-white rounded-xl text-xs font-bold flex items-center space-x-1.5 shadow-sm transition-all disabled:opacity-50 cursor-pointer"
                title="선택된 개정본 파일들을 ZIP 압축하여 다운로드합니다."
              >
                <Download className="w-3.5 h-3.5" />
                <span>선택 엑셀 다운로드 ({selectedCount}건)</span>
              </button>

              {/* All Excel ZIP Download Button */}
              <button
                id="download-all-excel-btn"
                onClick={() => handleDownloadExcelZip('all')}
                disabled={isDownloadingExcel || filteredRevisions.length === 0}
                className="px-3.5 py-2 bg-slate-800 hover:bg-slate-900 active:bg-black text-white rounded-xl text-xs font-bold flex items-center space-x-1.5 shadow-sm transition-all disabled:opacity-50 cursor-pointer"
                title="조회된 전체 개정본 목록을 ZIP 압축하여 일괄 다운로드합니다."
              >
                <FolderArchive className="w-3.5 h-3.5 text-emerald-400" />
                <span>전체 엑셀 다운로드 ({filteredRevisions.length}건)</span>
              </button>
            </div>
          </div>

          {/* Filter, Quick Select, & Sorting Bar */}
          <div className="p-3.5 bg-slate-50 border-b border-slate-200 flex flex-col md:flex-row md:items-center justify-between gap-3 text-xs">
            {/* Quick Selection Shortcuts */}
            <div className="flex flex-wrap items-center gap-1.5">
              <span className="font-bold text-slate-500 text-[11px] mr-1 flex items-center space-x-1">
                <ListCheck className="w-3.5 h-3.5" />
                <span>빠른 선택:</span>
              </span>
              <button
                onClick={() => handleQuickSelectCount(5)}
                className="px-2 py-1 bg-white hover:bg-slate-100 border border-slate-300 rounded text-slate-700 font-semibold transition-colors cursor-pointer"
              >
                최근 5건
              </button>
              <button
                onClick={() => handleQuickSelectCount(10)}
                className="px-2 py-1 bg-white hover:bg-slate-100 border border-slate-300 rounded text-slate-700 font-semibold transition-colors cursor-pointer"
              >
                최근 10건
              </button>
              <button
                onClick={() => handleQuickSelectCount(20)}
                className="px-2 py-1 bg-white hover:bg-slate-100 border border-slate-300 rounded text-slate-700 font-semibold transition-colors cursor-pointer"
              >
                최근 20건
              </button>
              <button
                onClick={() => handleSelectAll(true)}
                className="px-2 py-1 bg-white hover:bg-slate-100 border border-slate-300 rounded text-slate-700 font-semibold transition-colors cursor-pointer"
              >
                전체 선택
              </button>
              <button
                onClick={() => handleSelectAll(false)}
                className="px-2 py-1 bg-white hover:bg-slate-100 border border-slate-300 rounded text-slate-700 font-semibold transition-colors cursor-pointer"
              >
                전체 해제
              </button>

              {targetType === 'law' && (
                <>
                  <span className="text-slate-300 mx-1">|</span>
                  <button
                    onClick={() => handleSelectSubTypesRevisions('law')}
                    className="px-2 py-1 bg-blue-50 hover:bg-blue-100 border border-blue-200 text-blue-800 rounded font-semibold transition-colors cursor-pointer"
                  >
                    법률 ({lawCount})
                  </button>
                  <button
                    onClick={() => handleSelectSubTypesRevisions('decree')}
                    className="px-2 py-1 bg-amber-50 hover:bg-amber-100 border border-amber-200 text-amber-800 rounded font-semibold transition-colors cursor-pointer"
                  >
                    시행령 ({decreeCount})
                  </button>
                  <button
                    onClick={() => handleSelectSubTypesRevisions('rule')}
                    className="px-2 py-1 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 text-emerald-800 rounded font-semibold transition-colors cursor-pointer"
                  >
                    시행규칙 ({ruleCount})
                  </button>
                </>
              )}
            </div>

            {/* Search Filter & SubType Filter */}
            <div className="flex items-center space-x-2 shrink-0">
              <select
                value={ruleTypeFilter}
                onChange={(e) => setRuleTypeFilter(e.target.value)}
                className="px-2.5 py-1 bg-white border border-slate-300 rounded-lg text-xs font-semibold focus:outline-none focus:ring-1 focus:ring-emerald-500"
              >
                <option value="all">전체 구분 보기</option>
                <option value="law">법률만</option>
                <option value="decree">시행령만</option>
                <option value="rule">시행규칙만</option>
              </select>

              <div className="relative">
                <input
                  type="text"
                  value={filterQuery}
                  onChange={(e) => setFilterQuery(e.target.value)}
                  placeholder="결과 내 검색 (번호, 일자 등)..."
                  className="pl-2.5 pr-6 py-1 bg-white border border-slate-300 rounded-lg text-xs font-medium focus:outline-none focus:ring-1 focus:ring-emerald-500 w-44"
                />
                {filterQuery && (
                  <button
                    onClick={() => setFilterQuery('')}
                    className="absolute right-1.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                  >
                    <X className="w-3 h-3" />
                  </button>
                )}
              </div>
            </div>
          </div>

          {/* Revision Table */}
          <div className="overflow-x-auto min-h-[300px]">
            {filteredRevisions.length === 0 ? (
              <div className="py-20 flex flex-col items-center justify-center space-y-2 text-slate-400">
                <BookOpen className="w-8 h-8 text-slate-300" />
                <p className="text-sm font-semibold text-slate-600">
                  조회 조건과 일치하는 개정 연혁 목록이 없습니다.
                </p>
                <p className="text-xs text-slate-400">
                  검색 키워드를 변경하거나 매칭 옵션을 확인해 보세요.
                </p>
              </div>
            ) : (
              <table className="w-full text-left text-xs border-collapse">
                <thead className="bg-slate-100/80 sticky top-0 z-10 border-b border-slate-200 text-slate-600 font-bold">
                  <tr>
                    <th className="p-3 w-10 text-center">
                      <input
                        type="checkbox"
                        checked={isAllSelected}
                        onChange={(e) => handleSelectAll(e.target.checked)}
                        className="rounded text-emerald-600 focus:ring-emerald-500 w-4 h-4 cursor-pointer"
                        title="전체 선택 / 전체 해제"
                      />
                    </th>
                    <th
                      className="p-3 w-28 cursor-pointer hover:bg-slate-200/70 transition-colors select-none"
                      onClick={() => handleSortToggle('hierarchy')}
                      title="법령구분(위계) 정렬"
                    >
                      <div className="flex items-center space-x-1">
                        <span>법령구분</span>
                        {sortField === 'hierarchy' && (
                          <span className="text-emerald-700 font-bold">{sortOrder === 'asc' ? '↑' : '↓'}</span>
                        )}
                      </div>
                    </th>
                    <th
                      className="p-3 min-w-[200px] cursor-pointer hover:bg-slate-200/70 transition-colors select-none"
                      onClick={() => handleSortToggle('name')}
                      title="법령/행정규칙명 가나다순 정렬"
                    >
                      <div className="flex items-center space-x-1">
                        <span>법령/행정규칙명</span>
                        {sortField === 'name' ? (
                          <span className="text-emerald-700 font-bold">{sortOrder === 'asc' ? '▲' : '▼'}</span>
                        ) : (
                          <ArrowUpDown className="w-3 h-3 text-slate-400 opacity-60" />
                        )}
                      </div>
                    </th>
                    <th className="p-3 w-24">제·개정구분</th>
                    <th
                      className="p-3 w-28 cursor-pointer hover:bg-slate-200/70 transition-colors select-none"
                      onClick={() => handleSortToggle('enforcementDate')}
                      title="시행일자 순 정렬"
                    >
                      <div className="flex items-center space-x-1">
                        <span>시행일자</span>
                        {sortField === 'enforcementDate' ? (
                          <span className="text-emerald-700 font-bold">{sortOrder === 'asc' ? '▲' : '▼'}</span>
                        ) : (
                          <ArrowUpDown className="w-3 h-3 text-slate-400 opacity-60" />
                        )}
                      </div>
                    </th>
                    <th
                      className="p-3 w-28 cursor-pointer hover:bg-slate-200/70 transition-colors select-none"
                      onClick={() => handleSortToggle('promulgationDate')}
                      title="개정(공포)일자 순 정렬"
                    >
                      <div className="flex items-center space-x-1">
                        <span>개정일자 (공포·발령)</span>
                        {sortField === 'promulgationDate' && (
                          <span className="text-emerald-700 font-bold">{sortOrder === 'asc' ? '▲' : '▼'}</span>
                        )}
                      </div>
                    </th>
                    <th
                      className="p-3 min-w-[140px] cursor-pointer hover:bg-slate-200/70 transition-colors select-none"
                      onClick={() => handleSortToggle('promulgationNo')}
                      title="공포번호 순 정렬"
                    >
                      <div className="flex items-center space-x-1">
                        <span>공포(발령)번호</span>
                        {sortField === 'promulgationNo' && (
                          <span className="text-emerald-700 font-bold">{sortOrder === 'asc' ? '▲' : '▼'}</span>
                        )}
                      </div>
                    </th>
                    <th className="p-3 w-28">소관부처</th>
                    <th className="p-3 w-24">일련번호(ID)</th>
                    <th className="p-3 w-28 text-center">엑셀 다운</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200 bg-white">
                  {filteredRevisions.map((item, index) => {
                    const revKey = getRevKey(item);
                    const isChecked = !!selectedRevisions[revKey];

                    const isDecree = item.subType === 'decree' || item.name.includes('시행령') || item.ruleType?.includes('대통령령');
                    const isRule = item.subType === 'rule' || item.name.includes('시행규칙') || item.ruleType?.includes('부령');
                    const isAdmrul = item.targetType === 'admrul';

                    let badgeClass = 'bg-blue-100 text-blue-800 border-blue-200';
                    let badgeText = item.ruleType || '법률';
                    if (isDecree) {
                      badgeClass = 'bg-amber-100 text-amber-800 border-amber-200';
                      badgeText = item.ruleType || '대통령령(시행령)';
                    } else if (isRule) {
                      badgeClass = 'bg-emerald-100 text-emerald-800 border-emerald-200';
                      badgeText = item.ruleType || '부령(시행규칙)';
                    } else if (isAdmrul) {
                      badgeClass = 'bg-purple-100 text-purple-800 border-purple-200';
                      badgeText = item.ruleType || '고시';
                    }

                    const revType = item.revisionType || '일부개정';
                    let revClass = 'bg-slate-100 text-slate-700';
                    if (revType.includes('제정')) revClass = 'bg-emerald-100 text-emerald-800 font-bold';
                    else if (revType.includes('전부개정')) revClass = 'bg-indigo-100 text-indigo-800 font-bold';
                    else if (revType.includes('타법개정')) revClass = 'bg-violet-100 text-violet-800';

                    return (
                      <tr
                        key={revKey}
                        onClick={() => handleToggleRevision(revKey)}
                        className={`cursor-pointer transition-colors ${
                          isChecked ? 'bg-emerald-50/40 hover:bg-emerald-50/70' : 'hover:bg-slate-50'
                        }`}
                      >
                        {/* Checkbox */}
                        <td className="p-3 text-center" onClick={(e) => e.stopPropagation()}>
                          <input
                            type="checkbox"
                            checked={isChecked}
                            onChange={() => handleToggleRevision(revKey)}
                            className="rounded text-emerald-600 focus:ring-emerald-500 w-4 h-4 cursor-pointer"
                          />
                        </td>

                        {/* 법령구분 */}
                        <td className="p-3">
                          <span className={`px-2 py-0.5 rounded-full text-[11px] font-bold border inline-block ${badgeClass}`}>
                            {badgeText}
                          </span>
                        </td>

                        {/* 법령명 */}
                        <td className="p-3 font-bold text-slate-900">
                          <div className="flex items-center space-x-1.5 flex-wrap gap-y-1">
                            <span>{item.name}</span>
                            {item.isPredecessor && (
                              <span
                                className="px-1.5 py-0.2 rounded text-[10px] font-bold bg-amber-100 text-amber-800 border border-amber-300"
                                title={item.predecessorNote || '변경 전 구 명칭 연혁'}
                              >
                                변경전 명칭
                              </span>
                            )}
                            {index === 0 && !item.isPredecessor && (
                              <span className="px-1.5 py-0.2 rounded text-[10px] font-bold bg-rose-100 text-rose-700">
                                최신
                              </span>
                            )}
                          </div>
                        </td>

                        {/* 제개정구분 */}
                        <td className="p-3">
                          <span className={`px-2 py-0.5 rounded text-[11px] ${revClass}`}>
                            {revType}
                          </span>
                        </td>

                        {/* 시행일자 */}
                        <td className="p-3 font-semibold">
                          {item.enforcementDate && item.enforcementDate !== '시행미정' && !item.enforcementDate.includes('9999') ? (
                            <span className="text-emerald-700">{item.enforcementDate}</span>
                          ) : (
                            <span className="px-2 py-0.5 rounded text-[11px] font-bold bg-amber-100 text-amber-900 border border-amber-300">
                              시행미정
                            </span>
                          )}
                        </td>

                        {/* 개정일자 (공포일자) */}
                        <td className="p-3 font-semibold text-slate-700">
                          {item.promulgationDate || '-'}
                        </td>

                        {/* 공포번호 */}
                        <td className="p-3 text-slate-700 font-medium">
                          {item.promulgationNo || '-'}
                        </td>

                        {/* 소관부처 */}
                        <td className="p-3 text-slate-600">
                          {item.department || '기획재정부'}
                        </td>

                        {/* 일련번호 */}
                        <td className="p-3 text-slate-400 font-mono text-[11px]">
                          {item.id || item.seq || '-'}
                        </td>

                        {/* 단일 엑셀 다운로드 */}
                        <td className="p-3 text-center" onClick={(e) => e.stopPropagation()}>
                          <button
                            onClick={(e) => handleDownloadSingleRevisionExcel(item, e)}
                            disabled={isDownloadingExcel}
                            className="px-2.5 py-1 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-300 rounded-lg text-[11px] font-bold flex items-center space-x-1 transition-colors shadow-2xs mx-auto disabled:opacity-50 cursor-pointer"
                            title="해당 개정본 단일 엑셀 파일 (.xlsx) 즉시 다운로드"
                          >
                            <Download className="w-3 h-3 text-emerald-600" />
                            <span>엑셀 (.xlsx)</span>
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            )}
          </div>

          {/* Footer Summary & Custom Folder Name Configuration */}
          <div className="p-4 bg-slate-50 border-t border-slate-200 flex flex-col md:flex-row md:items-center md:justify-between gap-4">
            <div className="flex flex-wrap items-center gap-2.5 text-xs">
              <span className="font-bold text-slate-800 flex items-center space-x-1.5 shrink-0">
                <FolderArchive className="w-4 h-4 text-emerald-700" />
                <span>압축(ZIP) 폴더명 설정:</span>
              </span>
              <input
                type="text"
                value={customFolderName}
                onChange={(e) => {
                  setCustomFolderName(e.target.value);
                  setIsCustomFolderEdited(true);
                }}
                placeholder={computeDefaultFolderName(searchKeyword || (revisions[0]?.name ?? ''), targetType, selectedSubTypes)}
                className="px-3 py-1.5 text-xs bg-white border border-slate-300 rounded-lg text-slate-900 font-mono font-bold focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-600 min-w-[280px] shadow-2xs"
              />
              {isCustomFolderEdited ? (
                <button
                  type="button"
                  onClick={handleResetFolderName}
                  className="px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-emerald-50 text-emerald-800 hover:bg-emerald-100 border border-emerald-200 transition-colors flex items-center space-x-1 cursor-pointer"
                  title="기본 폴더명으로 초기화"
                >
                  <RotateCw className="w-3.5 h-3.5 text-emerald-700" />
                  <span>기본명 리셋</span>
                </button>
              ) : (
                <span className="text-[11px] text-slate-500 font-medium">
                  {targetType === 'law' ? '• 법/시행령/시행규칙 선택에 맞춰 자동 부여됩니다.' : '• 행정규칙 명칭이 자동 반영됩니다.'}
                </span>
              )}
            </div>

            <div className="flex items-center space-x-3 text-xs text-slate-600 shrink-0 font-medium">
              <span>선택된 개정본 <strong>{selectedCount}개</strong> (전체 {filteredRevisions.length}개)</span>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
