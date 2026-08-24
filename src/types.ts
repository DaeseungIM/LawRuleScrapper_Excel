export interface LawRevisionItem {
  lawId: string;
  lawMst: string;
  lawName: string;
  promulgationDate: string;
  promulgationNo: string;
  enforcementDate: string;
  revisionType: string; // e.g. 일부개정, 타법개정, 제정
  department: string; // 소관부처 (e.g. 기획재정부)
  lawType: string; // 법령종류 (e.g. 법률)
}

export interface LawInfo {
  lawId: string;
  lawMst: string;
  lawName: string;
  promulgationDate: string;
  promulgationNo: string;
  enforcementDate: string;
  revisionType: string; // e.g. 일부개정, 타법개정, 제정
  department: string; // 소관부처 (e.g. 기획재정부)
  lawType: string; // 법령종류 (e.g. 법률)
  articleCount?: number;
}

export interface LawArticle {
  articleNo: string; // e.g., "제1조"
  articleTitle: string; // e.g., "목적"
  articleContent: string; // Main text
  chapterName?: string; // e.g., "제1장 총칙"
  sectionName?: string; // e.g., "제1절 통칙"
  subsectionName?: string; // e.g., "제1관"
  effectiveDate?: string;
  isDeleted?: boolean;
}

export interface CustomsActData {
  info: LawInfo;
  articles: LawArticle[];
  fetchedAt: string;
}

export interface ExportConfig {
  targetType: 'new' | 'existing';
  spreadsheetIdOrUrl?: string;
  sheetName?: string;
  includeOverview: boolean;
  autoFormat: boolean;
  exportAll140?: boolean;
  exportMode?: 'selected' | 'separate_files_140' | 'single_file_140';
}

export interface ProcessStep {
  id: string;
  title: string;
  status: 'idle' | 'running' | 'success' | 'error';
  message?: string;
}

export interface UserProfile {
  displayName: string | null;
  email: string | null;
  photoURL: string | null;
}

export interface DecisionItem {
  id: string;
  year?: string;
  caseNo: string;
  title: string;
  decisionDate: string;
  department: string;
  targetType: string;
  relLaw?: string;
  itemDesc?: string;
  summary?: string;
  content?: string;
  category?: string;
}

export interface YearlyDecisionStat {
  year: string;
  committeeCount: number; // 위원회결정사항 (04)
  councilCount: number;   // 협의회결정사항 (03)
  caseCount: number;      // 품목분류사례 (01)
  totalCount: number;     // 합계
}

export interface HskItem {
  hskCode: string;
  pureCode: string;
  nameKo: string;
  nameEn: string;
  generalRate: string;
  agreementRate: string;
  unit1: string;
  unit2: string;
  remarks: string;
}

export interface HsExplanatoryItem {
  category: string;
  sectionChapter: string;
  hsHeading: string;
  titleKo: string;
  titleEn: string;
  scopeContent: string;
  guideline: string;
}

export interface HsOpinionItem {
  category: string;
  opinionNo: string;
  subheading: string;
  itemName: string;
  opinionText: string;
  rationale: string;
  remarks: string;
}

export type SearchTargetType = 'law' | 'admrul';
export type LawSubType = 'law' | 'decree' | 'rule'; // 법률, 시행령, 시행규칙
export type SearchMatchMode = 'exact' | 'contains'; // 정확히 일치, 포함
export type SearchWorkflowStep = 'step1_select_law' | 'step2_view_revisions'; // 1단계 (법령 선택), 2단계 (개정목록 조회)

export interface UnifiedSearchItem {
  id: string;
  seq: string;
  name: string;
  targetType: SearchTargetType;
  subType?: LawSubType;
  department: string;
  promulgationDate: string;
  promulgationNo: string;
  enforcementDate: string;
  revisionType: string;
  ruleType?: string; // 법률, 대통령령, 기획재정부령, 훈령, 예규, 고시 등
  currentYn?: string;
  isPredecessor?: boolean;
  predecessorNote?: string;
  matchedAliasNote?: string;
}

export interface UnifiedRevisionItem {
  id: string;
  seq: string;
  name: string;
  lawName?: string;
  lawMst?: string;
  lawId?: string;
  targetType: SearchTargetType;
  subType?: LawSubType;
  promulgationDate: string;
  promulgationNo: string;
  enforcementDate: string;
  revisionType: string;
  department: string;
  ruleType: string;
  buchikText?: string;
  checked?: boolean;
  isPredecessor?: boolean;
  predecessorNote?: string;
  matchedAliasNote?: string;
}

export interface DriveFolderInfo {
  id: string;
  name: string;
  url: string;
  created: boolean;
  isExisting: boolean;
}

export interface DrivePermissionOption {
  type: 'private' | 'anyone' | 'user';
  role: 'reader' | 'writer';
  email?: string;
}

export interface SaveProgressState {
  isSaving: boolean;
  currentStep: number;
  totalSteps: number;
  percentage: number;
  message: string;
  folderInfo?: DriveFolderInfo | null;
  savedSheets?: Array<{
    title?: string;
    name?: string;
    spreadsheetId?: string;
    url?: string;
    isExisting?: boolean;
    isNumbered?: boolean;
    duplicateNumber?: number;
    promulgationNo?: string;
    enforcementDate?: string;
  }>;
  error?: string | null;
  elapsedSeconds?: number;
}

export interface BuchikArticle {
  buchikCategory: string; // e.g. "부칙 <법률 제7849호, 2006. 2. 21.>"
  relatedLaw: string;     // e.g. "(제주특별자치도 설치 및 국제자유도시 조성을 위한 특별법)"
  articleNo: string;      // e.g. "제1조" or "-"
  articleTitle: string;   // e.g. "(시행일)" or "-"
  articleContent: string; // e.g. "제1조 (시행일) 이 법은 2006년 7월 1일부터 시행한다."
}

export interface DriveFolderItem {
  id: string;
  name: string;
  url?: string;
  createdTime?: string;
  modifiedTime?: string;
  isShared?: boolean;
}

export interface DriveSaveConfig {
  permission: DrivePermissionOption;
  customFolderName?: string;
  targetFolderId?: string;
  targetFolderName?: string;
  duplicateHandlingMode?: 'numbering' | 'skip';
}




