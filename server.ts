import express from 'express';
import path from 'path';
import { Readable } from 'stream';
import { createServer as createViteServer } from 'vite';
import { XMLParser } from 'fast-xml-parser';
import { google } from 'googleapis';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { HSK_TARIFF_DATA, HS_EXPLANATORY_DATA, HS_OPINION_DATA } from './src/lib/admRulesData';
import { generateHsk18823FullRows, cleanAndCollectHskExcelRows } from './src/lib/generateHsk18823Data';

const app = express();
const PORT = 3000;

app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));

const DEFAULT_OC_KEY = 'ceiai_law_test';

// Helper to format date YYYYMMDD to YYYY.MM.DD
function formatDate(dateStr: any): string {
  if (!dateStr) return '';
  const str = String(dateStr).trim();
  if (str.includes('9999') || str.includes('미정') || str.toLowerCase().includes('unknown')) {
    return '시행미정';
  }
  if (str.length === 8 && /^\d{8}$/.test(str)) {
    if (str.startsWith('9999')) return '시행미정';
    return `${str.substring(0, 4)}.${str.substring(4, 6)}.${str.substring(6, 8)}`;
  }
  return str;
}

// Helper to safely extract text from XML nodes
function getText(obj: any): string {
  if (obj === undefined || obj === null) return '';
  if (typeof obj === 'string' || typeof obj === 'number') return String(obj).trim();
  if (typeof obj === 'object') {
    if (obj['#text'] !== undefined) return String(obj['#text']).trim();
    if (obj['text'] !== undefined) return String(obj['text']).trim();
  }
  return '';
}

// XML parser configuration
const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseAttributeValue: false,
  trimValues: true,
});

// Helper to parse clean integer date for accurate chronological comparison (e.g. '2026.07.06' -> 20260706)
function parseCleanDateNumber(dateStr: any, fallbackDateStr?: any): number {
  if (!dateStr) {
    if (fallbackDateStr) return parseCleanDateNumber(fallbackDateStr);
    return 0;
  }
  const str = String(dateStr).trim();
  if (str.includes('9999') || str.includes('미정') || str.toLowerCase().includes('unknown')) {
    if (fallbackDateStr) return parseCleanDateNumber(fallbackDateStr);
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
  if (clean.length >= 8) {
    return parseInt(clean.slice(0, 8), 10) || 0;
  }
  if (clean.length >= 4) {
    return parseInt(clean.padEnd(8, '0'), 10) || 0;
  }
  if (fallbackDateStr) return parseCleanDateNumber(fallbackDateStr);
  return 0;
}

// Helper to format Promulgation Number accurately according to legal classification
function formatPromulgationNo(
  rawPromNo: any,
  subType: 'law' | 'decree' | 'rule' | string = 'law',
  lawName: string = '',
  department: string = ''
): string {
  const raw = String(rawPromNo || '').trim();
  const name = String(lawName || '').trim();
  const dept = String(department || '').trim();

  const isDecree =
    subType === 'decree' ||
    name.includes('시행령') ||
    raw.includes('대통령령');

  const isRule =
    subType === 'rule' ||
    name.includes('시행규칙') ||
    raw.includes('부령') ||
    raw.includes('총리령') ||
    raw.includes('기획재정부령') ||
    raw.includes('재정경제부령') ||
    raw.includes('재무부령');

  // Extract clean digits
  const digits = raw.replace(/[^0-9-]/g, '');

  if (isDecree) {
    let clean = raw.replace(/^(법률|부령|총리령)\s*/, '');
    if (clean.startsWith('대통령령')) return clean;
    if (clean.startsWith('제')) return `대통령령 ${clean}`;
    if (digits) return `대통령령 제${digits}호`;
    return clean ? `대통령령 ${clean}` : '대통령령';
  }

  if (isRule) {
    let clean = raw.replace(/^(법률|대통령령)\s*/, '');
    if (clean.includes('부령') || clean.includes('총리령') || clean.includes('규칙')) {
      return clean;
    }
    const deptPrefix = dept ? (dept.endsWith('부') ? `${dept}령` : `${dept} 부령`) : '부령';
    if (clean.startsWith('제')) return `${deptPrefix} ${clean}`;
    if (digits) return `${deptPrefix} 제${digits}호`;
    return clean ? `${deptPrefix} ${clean}` : deptPrefix;
  }

  // Standard Law (법률)
  let clean = raw.replace(/^(대통령령|부령|총리령)\s*/, '');
  if (clean.startsWith('법률')) return clean;
  if (clean.startsWith('제')) return `법률 ${clean}`;
  if (digits) return `법률 제${digits}호`;
  return clean ? `법률 ${clean}` : '법률';
}

// Helper to determine hierarchy rank: 1: 법(법률) -> 2: 시행령(대통령령) -> 3: 시행규칙(부령/총리령) -> 4: 행정규칙(고시/훈령/예규)
function getHierarchyRank(item: any): number {
  if (!item) return 99;
  const subType = (item.subType || '').toLowerCase();
  const name = (item.name || item.lawName || '').trim();
  const lawType = (item.lawType || item.ruleType || '').trim();
  const targetType = (item.targetType || '').toLowerCase();

  if (targetType === 'admrul') {
    return 4;
  }

  // 1. 법 (법률)
  if (
    subType === 'law' ||
    lawType === '법률' ||
    (!name.includes('시행령') && !name.includes('시행규칙') && !lawType.includes('대통령령') && !lawType.includes('부령') && !lawType.includes('총리령'))
  ) {
    if (!name.includes('시행령') && !name.includes('시행규칙') && !lawType.includes('대통령령') && !lawType.includes('부령')) {
      return 1;
    }
  }

  // 2. 시행령 (대통령령)
  if (subType === 'decree' || name.includes('시행령') || lawType.includes('대통령령') || lawType.includes('시행령')) {
    return 2;
  }

  // 3. 시행규칙 (부령/총리령/규칙)
  if (subType === 'rule' || name.includes('시행규칙') || lawType.includes('부령') || lawType.includes('총리령') || lawType.includes('규칙')) {
    return 3;
  }

  return 1;
}

// Master revision sorting function:
// 1순위: 법(1) -> 시행령(2) -> 시행규칙(3) -> 행정규칙(4)
// 2순위 (각 구분 내): 시행일자 내림차순 (최근 -> 과거)
// 3순위: 공포일자 내림차순
// 4순위: 공포번호 / 일련번호 내림차순
function sortRevisionsByHierarchyAndDate(revisions: any[]): any[] {
  if (!Array.isArray(revisions)) return [];
  return [...revisions].sort((a, b) => {
    const rankA = getHierarchyRank(a);
    const rankB = getHierarchyRank(b);
    if (rankA !== rankB) {
      return rankA - rankB; // Ascending: 법(1) -> 시행령(2) -> 시행규칙(3) -> 행정규칙(4)
    }

    const promA = parseCleanDateNumber(a.promulgationDate || a.공포일자 || a.발령일자 || a.pramDate);
    const promB = parseCleanDateNumber(b.promulgationDate || b.공포일자 || b.발령일자 || b.pramDate);

    const dateA = parseCleanDateNumber(a.enforcementDate || a.시행일자 || a.efYd, a.promulgationDate || a.공포일자 || a.발령일자 || a.pramDate);
    const dateB = parseCleanDateNumber(b.enforcementDate || b.시행일자 || b.efYd, b.promulgationDate || b.공포일자 || b.발령일자 || b.pramDate);
    if (dateB !== dateA) {
      return dateB - dateA; // Descending: most recent enforcement date first
    }

    if (promB !== promA) {
      return promB - promA;
    }

    const noA = parseInt(String(a.promulgationNo || a.seq || a.id || a.lawMst || '').replace(/[^0-9]/g, ''), 10) || 0;
    const noB = parseInt(String(b.promulgationNo || b.seq || b.id || b.lawMst || '').replace(/[^0-9]/g, ''), 10) || 0;
    return noB - noA;
  });
}

// Alias for backwards compatibility
function sortRevisionsByEnforcementDateDesc(revisions: any[]): any[] {
  return sortRevisionsByHierarchyAndDate(revisions);
}

// Predecessor / Historical Legislation Map (변경 전 구 법령 및 행정규칙 연혁 매핑)
const PREDECESSOR_MAP: Record<string, string[]> = {
  // 행정규칙: 관세평가 운영에 관한 고시 (7개) ➔ 변경 전: 수입물품 과세가격 결정에 관한 고시 (21개) = 총 28개 전수 수집
  '관세평가 운영에 관한 고시': ['수입물품 과세가격 결정에 관한 고시', '수입물품과세가격결정에관한고시'],
  '관세평가운영에관한고시': ['수입물품 과세가격 결정에 관한 고시', '수입물품과세가격결정에관한고시'],
  '관세평가': ['관세평가 운영에 관한 고시', '수입물품 과세가격 결정에 관한 고시', '수입물품과세가격결정에관한고시'],
  '수입물품 과세가격 결정에 관한 고시': ['관세평가 운영에 관한 고시'],
  '수입물품과세가격결정에관한고시': ['관세평가 운영에 관한 고시'],
  '과세가격': ['관세평가 운영에 관한 고시', '수입물품 과세가격 결정에 관한 고시'],

  // 법령: 외국환거래법 ➔ 변경 전: 외국환관리법
  '외국환거래법': ['외국환관리법'],
  '외국환거래': ['외국환거래법', '외국환관리법', '외국환관리'],
  '외국환거래법 시행령': ['외국환관리법 시행령'],
  '외국환거래법 시행규칙': ['외국환관리법 시행규칙'],
  '외국환관리법': ['외국환거래법'],
  '외국환관리': ['외국환거래법', '외국환거래'],
  '외국환관리법 시행령': ['외국환거래법 시행령'],
  '외국환관리법 시행규칙': ['외국환거래법 시행규칙'],
  '외국환': ['외국환거래법', '외국환관리법', '외국환거래규정', '외국환관리규정'],

  // 행정규칙: 외국환거래규정 ➔ 외국환관리규정
  '외국환거래규정': ['외국환관리규정'],
  '외국환관리규정': ['외국환거래규정'],

  // 통관 관련 고시
  '수출통관 사무처리에 관한 고시': ['수출통관 사무처리 규정', '수출통관사무처리규정', '수출통관사무처리에관한고시'],
  '수입통관 사무처리에 관한 고시': ['수입통관 사무처리 규정', '수입통관사무처리규정', '수입통관사무처리에관한고시'],
  '보세판매장 운영에 관한 고시': ['보세판매장운영에관한고시', '보세판매장운영에관한규정'],
  '보세화물 관리에 관한 고시': ['보세화물관리에관한고시', '보세화물관리에관한규정'],
};

// Official & Social Law Alias Map (법제처 공식 약칭 및 사회적 통용 약칭 매핑)
const LAW_ALIAS_MAP: Record<string, { canonical: string[]; aliasName: string }> = {
  // 관세 및 무역
  '관세': { canonical: ['관세법'], aliasName: '관세법' },
  '관세법': { canonical: ['관세법'], aliasName: '관세법' },
  '외국환': { canonical: ['외국환거래법', '외국환관리법'], aliasName: '외국환거래법/외국환관리법' },
  '외국환거래': { canonical: ['외국환거래법'], aliasName: '외국환거래법' },
  '외국환관리': { canonical: ['외국환관리법', '외국환거래법'], aliasName: '외국환관리법/외국환거래법' },
  '대외무역': { canonical: ['대외무역법'], aliasName: '대외무역법' },
  '환특법': { canonical: ['수출용원재료에 대한 관세 등 환급에 관한 특례법'], aliasName: '수출용원재료에 대한 관세 등 환급에 관한 특례법' },
  '환급특례법': { canonical: ['수출용원재료에 대한 관세 등 환급에 관한 특례법'], aliasName: '수출용원재료에 대한 관세 등 환급에 관한 특례법' },
  '수출용원재료관세환급특례법': { canonical: ['수출용원재료에 대한 관세 등 환급에 관한 특례법'], aliasName: '수출용원재료에 대한 관세 등 환급에 관한 특례법' },
  'FTA특례법': { canonical: ['자유무역협정의 이행을 위한 관세법의 특례에 관한 법률'], aliasName: '자유무역협정의 이행을 위한 관세법의 특례에 관한 법률' },
  'FTA관세특례법': { canonical: ['자유무역협정의 이행을 위한 관세법의 특례에 관한 법률'], aliasName: '자유무역협정의 이행을 위한 관세법의 특례에 관한 법률' },
  '자유무역협정관세법': { canonical: ['자유무역협정의 이행을 위한 관세법의 특례에 관한 법률'], aliasName: '자유무역협정의 이행을 위한 관세법의 특례에 관한 법률' },
  '자유무역협정특례법': { canonical: ['자유무역협정의 이행을 위한 관세법의 특례에 관한 법률'], aliasName: '자유무역협정의 이행을 위한 관세법의 특례에 관한 법률' },

  // 주요 사회적 약칭 및 법제처 공식 약칭
  '상가임대차법': { canonical: ['상가건물 임대차보호법'], aliasName: '상가건물 임대차보호법' },
  '상가임대차': { canonical: ['상가건물 임대차보호법'], aliasName: '상가건물 임대차보호법' },
  '상임법': { canonical: ['상가건물 임대차보호법'], aliasName: '상가건물 임대차보호법' },
  '주임법': { canonical: ['주택임대차보호법'], aliasName: '주택임대차보호법' },
  '주택임대차법': { canonical: ['주택임대차보호법'], aliasName: '주택임대차보호법' },
  '주택임대차': { canonical: ['주택임대차보호법'], aliasName: '주택임대차보호법' },
  '특가법': { canonical: ['특정범죄 가중처벌 등에 관한 법률'], aliasName: '특정범죄 가중처벌 등에 관한 법률' },
  '특정범죄가중처벌법': { canonical: ['특정범죄 가중처벌 등에 관한 법률'], aliasName: '특정범죄 가중처벌 등에 관한 법률' },
  '특경법': { canonical: ['특정경제범죄 가중처벌 등에 관한 법률'], aliasName: '특정경제범죄 가중처벌 등에 관한 법률' },
  '특정경제범죄가중처벌법': { canonical: ['특정경제범죄 가중처벌 등에 관한 법률'], aliasName: '특정경제범죄 가중처벌 등에 관한 법률' },
  '김영란법': { canonical: ['부정청탁 및 금품등 수수의 금지에 관한 법률'], aliasName: '부정청탁 및 금품등 수수의 금지에 관한 법률' },
  '청탁금지법': { canonical: ['부정청탁 및 금품등 수수의 금지에 관한 법률'], aliasName: '부정청탁 및 금품등 수수의 금지에 관한 법률' },
  '자본시장법': { canonical: ['자본시장과 금융투자업에 관한 법률'], aliasName: '자본시장과 금융투자업에 관한 법률' },
  '자통법': { canonical: ['자본시장과 금융투자업에 관한 법률'], aliasName: '자본시장과 금융투자업에 관한 법률' },
  '개보법': { canonical: ['개인정보 보호법'], aliasName: '개인정보 보호법' },
  '개인정보보호법': { canonical: ['개인정보 보호법'], aliasName: '개인정보 보호법' },
  '남녀고용평등법': { canonical: ['남녀고용평등과 일ㆍ가정 양립 지원에 관한 법률'], aliasName: '남녀고용평등과 일ㆍ가정 양립 지원에 관한 법률' },
  '남녀고용평등': { canonical: ['남녀고용평등과 일ㆍ가정 양립 지원에 관한 법률'], aliasName: '남녀고용평등과 일ㆍ가정 양립 지원에 관한 법률' },
  '공정거래법': { canonical: ['독점규제 및 공정거래에 관한 법률'], aliasName: '독점규제 및 공정거래에 관한 법률' },
  '독점규제법': { canonical: ['독점규제 및 공정거래에 관한 법률'], aliasName: '독점규제 및 공정거래에 관한 법률' },
  '도정법': { canonical: ['도시 및 주거환경정비법'], aliasName: '도시 및 주거환경정비법' },
  '도시정비법': { canonical: ['도시 및 주거환경정비법'], aliasName: '도시 및 주거환경정비법' },
  '도시및주거환경정비법': { canonical: ['도시 및 주거환경정비법'], aliasName: '도시 및 주거환경정비법' },
  '중대재해법': { canonical: ['중대재해 처벌 등에 관한 법률'], aliasName: '중대재해 처벌 등에 관한 법률' },
  '중대재해처벌법': { canonical: ['중대재해 처벌 등에 관한 법률'], aliasName: '중대재해 처벌 등에 관한 법률' },
  '중처법': { canonical: ['중대재해 처벌 등에 관한 법률'], aliasName: '중대재해 처벌 등에 관한 법률' },
  '민소법': { canonical: ['민사소송법'], aliasName: '민사소송법' },
  '형소법': { canonical: ['형사소송법'], aliasName: '형사소송법' },
  '행소법': { canonical: ['행정소송법'], aliasName: '행정소송법' },
  '행정기본': { canonical: ['행정기본법'], aliasName: '행정기본법' },
  '국기법': { canonical: ['국세기본법'], aliasName: '국세기본법' },
  '국세기본': { canonical: ['국세기본법'], aliasName: '국세기본법' },
  '조특법': { canonical: ['조세특례제한법'], aliasName: '조세특례제한법' },
  '지특법': { canonical: ['지방세특례제한법'], aliasName: '지방세특례제한법' },
  '지세기본법': { canonical: ['지방세기본법'], aliasName: '지방세기본법' },
  '근로기준': { canonical: ['근로기준법'], aliasName: '근로기준법' },
  '노조법': { canonical: ['노동조합 및 노동관계조정법'], aliasName: '노동조합 및 노동관계조정법' },
  '통비법': { canonical: ['통신비밀보호법'], aliasName: '통신비밀보호법' },
  '신정법': { canonical: ['신용정보의 이용 및 보호에 관한 법률'], aliasName: '신용정보의 이용 및 보호에 관한 법률' },
  '신용정보법': { canonical: ['신용정보의 이용 및 보호에 관한 법률'], aliasName: '신용정보의 이용 및 보호에 관한 법률' },
  '화평법': { canonical: ['화학물질의 등록 및 평가 등에 관한 법률'], aliasName: '화학물질의 등록 및 평가 등에 관한 법률' },
  '화관법': { canonical: ['화학물질관리법'], aliasName: '화학물질관리법' },
  '산안법': { canonical: ['산업안전보건법'], aliasName: '산업안전보건법' },
  '건산법': { canonical: ['건설산업기본법'], aliasName: '건설산업기본법' },
  '가특법': { canonical: ['가맹사업거래의 공정화에 관한 법률'], aliasName: '가맹사업거래의 공정화에 관한 법률' },
  '하도급법': { canonical: ['하도급거래 공정화에 관한 법률'], aliasName: '하도급거래 공정화에 관한 법률' },
  '여전법': { canonical: ['여신전문금융업법'], aliasName: '여신전문금융업법' },
  '특금법': { canonical: ['특정 금융거래정보의 보고 및 이용 등에 관한 법률'], aliasName: '특정 금융거래정보의 보고 및 이용 등에 관한 법률' },
  '방판법': { canonical: ['방문판매 등에 관한 법률'], aliasName: '방문판매 등에 관한 법률' },
  '전상법': { canonical: ['전자상거래 등에서의 소비자보호에 관한 법률'], aliasName: '전자상거래 등에서의 소비자보호에 관한 법률' },
  '전자상거래법': { canonical: ['전자상거래 등에서의 소비자보호에 관한 법률'], aliasName: '전자상거래 등에서의 소비자보호에 관한 법률' },
  '소비자기본': { canonical: ['소비자기본법'], aliasName: '소비자기본법' },
  '소비자보호법': { canonical: ['소비자기본법'], aliasName: '소비자기본법' },
  '소득세': { canonical: ['소득세법'], aliasName: '소득세법' },
  '법인세': { canonical: ['법인세법'], aliasName: '법인세법' },
  '부가세': { canonical: ['부가가치세법'], aliasName: '부가가치세법' },
  '부가가치세': { canonical: ['부가가치세법'], aliasName: '부가가치세법' },
  '부동산등기': { canonical: ['부동산등기법'], aliasName: '부동산등기법' },
};

// Helper to resolve aliases into canonical law names
function resolveLawAliases(query: string): { canonicalNames: string[]; matchedAlias?: string } {
  const clean = (query || '').trim();
  const cleanNoSpace = clean.replace(/\s+/g, '').toLowerCase();

  // 1. Direct match
  if (LAW_ALIAS_MAP[clean]) {
    return {
      canonicalNames: LAW_ALIAS_MAP[clean].canonical,
      matchedAlias: clean !== LAW_ALIAS_MAP[clean].aliasName ? `${clean} ➔ ${LAW_ALIAS_MAP[clean].aliasName}` : undefined,
    };
  }

  // 2. Space-insensitive match
  for (const [aliasKey, aliasData] of Object.entries(LAW_ALIAS_MAP)) {
    if (aliasKey.replace(/\s+/g, '').toLowerCase() === cleanNoSpace) {
      return {
        canonicalNames: aliasData.canonical,
        matchedAlias: clean !== aliasData.aliasName ? `${clean} ➔ ${aliasData.aliasName}` : undefined,
      };
    }
  }

  // 3. Fallback to clean name
  return { canonicalNames: [clean] };
}

// Helper to fetch all revisions for any Law (법률 - 관세법 141건, 관세법 시행령 197건, 관세법 시행규칙 183건 전수 수집 및 다중 시행일자 전수 보존)
async function fetchLawRevisions(
  ocKey: string = DEFAULT_OC_KEY,
  lawName: string = '관세법',
  limit: number = 0,
  matchMode: 'exact' | 'contains' = 'exact',
  subTypes: string[] = ['law', 'decree', 'rule'],
  includePredecessors: boolean = true
): Promise<any[]> {
  try {
    const cleanQuery = (lawName || '관세법').trim();
    const cleanNoSpace = cleanQuery.replace(/\s+/g, '').toLowerCase();
    const collectedMap = new Map<string, any>();

    // Build specific query terms for sub-types to guarantee full coverage
    const searchTerms: { term: string; targetSubType: 'law' | 'decree' | 'rule' }[] = [];
    
    // Always add the base clean query if law is requested or if query is specific
    if (subTypes.includes('law') || cleanQuery.includes('법')) {
      searchTerms.push({ term: cleanQuery, targetSubType: 'law' });
    }

    // If subTypes include decree and the query doesn't already specify decree/rule, add decree term
    if (
      subTypes.includes('decree') &&
      !cleanQuery.includes('시행령') &&
      !cleanQuery.includes('시행규칙') &&
      !cleanQuery.includes('규칙') &&
      !cleanQuery.includes('령')
    ) {
      searchTerms.push({ term: `${cleanQuery} 시행령`, targetSubType: 'decree' });
    } else if (cleanQuery.includes('시행령')) {
      searchTerms.push({ term: cleanQuery, targetSubType: 'decree' });
    }

    // If subTypes include rule and the query doesn't already specify rule, add rule term
    if (
      subTypes.includes('rule') &&
      !cleanQuery.includes('시행규칙') &&
      !cleanQuery.includes('규칙')
    ) {
      searchTerms.push({ term: `${cleanQuery} 시행규칙`, targetSubType: 'rule' });
    } else if (cleanQuery.includes('시행규칙')) {
      searchTerms.push({ term: cleanQuery, targetSubType: 'rule' });
    }

    // Deduplicate search terms
    const uniqueSearchTerms = Array.from(
      new Map(searchTerms.map((s) => [s.term.replace(/\s+/g, ''), s])).values()
    );

    // DRF lawSearch: query both target=eflaw (effective-date versions) and target=law (full historical archive)
    const targetsToQuery = ['eflaw', 'law'];

    for (const { term, targetSubType } of uniqueSearchTerms) {
      const termNoSpace = term.replace(/\s+/g, '').toLowerCase();

      for (const tgt of targetsToQuery) {
        for (let page = 1; page <= 50; page++) {
          const searchUrl = `http://www.law.go.kr/DRF/lawSearch.do?OC=${encodeURIComponent(
            ocKey
          )}&target=${tgt}&query=${encodeURIComponent(term)}&page=${page}&display=100&type=XML`;

          console.log(
            `[Law Revisions Search] Target: ${tgt}, Term: ${term}, Page ${page} (mode: ${matchMode})`
          );
          const response = await fetch(searchUrl, {
            headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
          });

          if (!response.ok) {
            console.warn(`[Law Revisions] Response not OK (${response.status}) for ${term} page ${page}`);
            break;
          }

          const xmlText = await response.text();
          const parsed = xmlParser.parse(xmlText);
          const searchRoot = parsed.LawSearch || parsed.lawSearch || parsed;
          let lawList = searchRoot.law || searchRoot.Law || [];
          if (!Array.isArray(lawList)) lawList = lawList ? [lawList] : [];
          if (lawList.length === 0) break;

          for (const item of lawList) {
            const itemNm = getText(item.법령명한글 || item.법령명_한글 || item.lawName || item['#text']).trim();
            const itemNoSpace = itemNm.replace(/\s+/g, '').toLowerCase();
            const lawType = getText(item.법령구분명 || item.법령종류 || '');

            // Determine item subtype
            let itemSubType: 'law' | 'decree' | 'rule' = targetSubType || 'law';
            if (itemNm.includes('시행령') || lawType.includes('대통령령') || lawType.includes('시행령')) {
              itemSubType = 'decree';
            } else if (
              itemNm.includes('시행규칙') ||
              lawType.includes('부령') ||
              lawType.includes('총리령') ||
              lawType.includes('규칙')
            ) {
              itemSubType = 'rule';
            } else {
              itemSubType = 'law';
            }

            // SubType filtering
            if (subTypes && subTypes.length > 0 && !subTypes.includes(itemSubType)) {
              continue;
            }

            // Matching logic
            if (matchMode === 'exact') {
              const isTermExact =
                itemNoSpace === termNoSpace ||
                itemNm === term ||
                (targetSubType === 'law' && (itemNoSpace === cleanNoSpace || itemNoSpace === cleanNoSpace + '법')) ||
                (targetSubType === 'decree' && (itemNoSpace === `${cleanNoSpace}시행령` || itemNm === `${cleanQuery} 시행령`)) ||
                (targetSubType === 'rule' && (itemNoSpace === `${cleanNoSpace}시행규칙` || itemNm === `${cleanQuery} 시행규칙`));

              if (!isTermExact) {
                continue;
              }
            } else {
              // Contains mode
              const matchesContains =
                itemNm.toLowerCase().includes(cleanQuery.toLowerCase()) ||
                itemNoSpace.includes(cleanNoSpace);
              if (!matchesContains) continue;
            }

            const dept = getText(item.소관부처명 || item.소관부처 || '기획재정부');
            const rawPromNo = getText(item.공포번호);
            const formattedPromNo = formatPromulgationNo(rawPromNo, itemSubType, itemNm, dept);

            const lawId = getText(item.법령ID || item.lawId || item.MST || item.법령일련번호);
            const lawMst = getText(item.법령일련번호 || item.MST || item.mst || item.법령ID);
            const enfDate = formatDate(getText(item.시행일자));
            const promDate = formatDate(getText(item.공포일자));

            // CRITICAL: Unique deduplication key MUST include enforcementDate so that
            // multiple enforcement dates for the same promulgation (e.g. 2026.4.1 vs 2026.7.1)
            // are NEVER dropped or collapsed!
            const primaryKey = lawMst
              ? `MST_${lawMst}_${enfDate}`
              : `KEY_${itemSubType}_${itemNoSpace}_${enfDate}_${promDate}_${rawPromNo}`;

            if (!collectedMap.has(primaryKey)) {
              const revItem = {
                lawId: lawId || lawMst,
                lawMst: lawMst || lawId,
                id: lawMst || lawId,
                seq: lawMst || lawId,
                lawName: itemNm || cleanQuery,
                name: itemNm || cleanQuery,
                subType: itemSubType,
                promulgationDate: promDate,
                promulgationNo: formattedPromNo,
                enforcementDate: enfDate,
                revisionType: getText(item.제개정구분명 || item.제개정구분 || '일부개정'),
                department: dept,
                lawType:
                  lawType ||
                  (itemSubType === 'decree'
                    ? '대통령령'
                    : itemSubType === 'rule'
                    ? dept
                      ? `${dept}령`
                      : '부령'
                    : '법률'),
                ruleType:
                  lawType ||
                  (itemSubType === 'decree'
                    ? '대통령령'
                    : itemSubType === 'rule'
                    ? dept
                      ? `${dept}령`
                      : '부령'
                    : '법률'),
                targetType: 'law',
                isPredecessor: false,
              };
              collectedMap.set(primaryKey, revItem);
            }
          }

          if (lawList.length < 100) break;
        }
      }
    }

    // Predecessor legislation resolution (e.g. 외국환거래법 ➔ 외국환관리법)
    if (includePredecessors) {
      const predCandidates = new Set<string>();
      if (PREDECESSOR_MAP[cleanQuery]) {
        PREDECESSOR_MAP[cleanQuery].forEach((p) => predCandidates.add(p));
      }
      for (const item of Array.from(collectedMap.values())) {
        if (PREDECESSOR_MAP[item.name]) {
          PREDECESSOR_MAP[item.name].forEach((p) => predCandidates.add(p));
        }
      }

      for (const predName of Array.from(predCandidates)) {
        if (predName === cleanQuery) continue;
        console.log(`[Law Predecessor Check] Fetching historical predecessor law '${predName}' for '${cleanQuery}'...`);
        try {
          const predRevs = await fetchLawRevisions(ocKey, predName, 0, 'exact', subTypes, false);
          for (const pItem of predRevs) {
            const pKey = pItem.lawMst
              ? `PRED_MST_${pItem.lawMst}_${pItem.enforcementDate}`
              : `PRED_${pItem.id}_${pItem.name}_${pItem.enforcementDate}_${pItem.promulgationDate}_${pItem.promulgationNo}`;
            if (!collectedMap.has(pKey)) {
              collectedMap.set(pKey, {
                ...pItem,
                isPredecessor: true,
                predecessorNote: `(변경전: ${pItem.name})`,
              });
            }
          }
        } catch (predErr) {
          console.warn(`[Law Predecessor Error] for ${predName}:`, predErr);
        }
      }
    }

    const mapped = Array.from(collectedMap.values());
    console.log(`[Law Revisions Search] Total collected revisions for '${cleanQuery}' (mode: ${matchMode}, subTypes: ${subTypes.join(',')}, count: ${mapped.length})`);
    const sorted = sortRevisionsByHierarchyAndDate(mapped);
    return limit > 0 ? sorted.slice(0, limit) : sorted;
  } catch (err) {
    console.error(`Error in fetchLawRevisions for ${lawName}:`, err);
    return [];
  }
}

// Helper to fetch all 140+ revisions for Customs Act (관세법)
async function fetchAll140Revisions(ocKey: string = DEFAULT_OC_KEY): Promise<any[]> {
  return fetchLawRevisions(ocKey, '관세법', 0, 'exact', ['law'], false);
}

// Helper to parse the full revision history popup HTML for any Administrative Rule (행정규칙 연혁 팝업 전수 수집기)
// e.g. 관세평가 운영에 관한 고시 (7개) + 수입물품 과세가격 결정에 관한 고시 (21개) = 총 28개 전수 자동 수집
async function fetchAdmrulHistoryFromPopup(
  admRulSeq: string,
  primaryName: string,
  defaultDept: string = '관세청'
): Promise<any[]> {
  try {
    const histUrl = `https://www.law.go.kr/admRulHstListR.do?admRulSeq=${admRulSeq}`;
    console.log(`[Admrul History Popup HTML] Fetching: ${histUrl} for ${primaryName} (${admRulSeq})`);
    const hRes = await fetch(histUrl, {
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
    });

    if (!hRes.ok) return [];
    const html = await hRes.text();
    const liRegex = /<li[^>]*>([\s\S]*?)<\/li>/gi;
    let match;
    const revisions: any[] = [];

    const formatDateDot = (str: string) => {
      if (!str) return '';
      if (str.includes('9999') || str.includes('미정') || str.toLowerCase().includes('unknown')) {
        return '시행미정';
      }
      const parts = str.replace(/[^0-9.]/g, '').split('.').map((s) => s.trim()).filter(Boolean);
      if (parts.length === 3) {
        if (parts[0] === '9999') return '시행미정';
        return parts[0] + '.' + parts[1].padStart(2, '0') + '.' + parts[2].padStart(2, '0');
      }
      return str.trim();
    };

    while ((match = liRegex.exec(html)) !== null) {
      const raw = match[1];
      const seqMatch = raw.match(/admRulViewHst\s*\(\s*['\"][^'\"]*['\"]\s*,\s*['\"](\d+)['\"]/i) || raw.match(/(\d{6,14})/);
      const itemSeq = seqMatch ? seqMatch[1] : admRulSeq;

      // Extract raw rule title from inside <a> tag: e.g. " 1. 관세평가 운영에 관한 고시<br />" or " 8. 수입물품 과세가격 결정에 관한 고시<br />"
      let rawTitle = '';
      const titleMatch = raw.match(/<a[^>]*>([\s\S]*?)<div/i) || raw.match(/<a[^>]*>([\s\S]*?)<\/a>/i);
      if (titleMatch) {
        rawTitle = titleMatch[1].replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
        // Strip leading numbering like "1. ", "28. "
        rawTitle = rawTitle.replace(/^\d+\.\s*/, '').trim();
      }
      if (!rawTitle) {
        rawTitle = primaryName;
      }

      const subMatch = raw.match(/<div[^>]*class=['\"]subtit1_1['\"][^>]*>([\s\S]*?)<\/div>/i);
      const subText = subMatch ? subMatch[1].replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim() : '';

      const enfMatch = subText.match(/\[시행\s*([^\]]+)\]/);
      const enfDate = formatDateDot(enfMatch ? enfMatch[1].trim() : '');

      const promMatch = subText.match(/\[([^\],]+),\s*([^\],]+),\s*([^\]]+)\]/);
      let promNo = '';
      let promDate = '';
      let revType = '일부개정';
      let dept = defaultDept;

      if (promMatch) {
        promNo = promMatch[1].trim();
        promDate = formatDateDot(promMatch[2].trim());
        revType = promMatch[3].trim().replace(/\.$/, '');
        if (promNo.includes('관세청')) dept = '관세청';
        else if (promNo.includes('기획재정부')) dept = '기획재정부';
        else if (promNo.includes('재정경제부')) dept = '재정경제부';
        else if (promNo.includes('국세청')) dept = '국세청';
      } else {
        promNo = subText;
      }

      const cleanPrimaryNoSpace = primaryName.replace(/\s+/g, '');
      const cleanRawNoSpace = rawTitle.replace(/\s+/g, '');
      const isHistoricalPredecessor = cleanRawNoSpace !== cleanPrimaryNoSpace;

      // Extract rule type (고시, 훈령, 예규, 지침 등)
      let ruleType = '고시';
      if (rawTitle.includes('훈령') || promNo.includes('훈령')) ruleType = '훈령';
      else if (rawTitle.includes('예규') || promNo.includes('예규')) ruleType = '예규';
      else if (rawTitle.includes('지침') || promNo.includes('지침')) ruleType = '지침';
      else if (rawTitle.includes('규정') || promNo.includes('규정')) ruleType = '규정';

      revisions.push({
        lawId: itemSeq,
        lawMst: itemSeq,
        seq: itemSeq,
        id: itemSeq,
        lawName: rawTitle,
        name: rawTitle,
        promulgationDate: promDate,
        promulgationNo: promNo,
        enforcementDate: enfDate,
        revisionType: revType,
        department: dept,
        lawType: `행정규칙(${ruleType})`,
        ruleType: ruleType,
        targetType: 'admrul' as const,
        isPredecessor: isHistoricalPredecessor,
        predecessorNote: isHistoricalPredecessor ? `(변경전: ${rawTitle})` : undefined,
      });
    }

    console.log(`[Admrul History Popup HTML] Extracted ${revisions.length} total historical revisions for '${primaryName}' (seq: ${admRulSeq})`);
    return revisions;
  } catch (err: any) {
    console.warn(`[fetchAdmrulHistoryFromPopup Error]:`, err?.message);
    return [];
  }
}

// Complete 28 Revisions Dataset for 관세평가 운영에 관한 고시 (7건) + 수입물품 과세가격 결정에 관한 고시 (21건)
const CANONICAL_CUSTOMS_VALUATION_28_REVISIONS: any[] = [
  // 1~7: 관세평가 운영에 관한 고시 (최신 고시 7건)
  { seq: '2100000261194', name: '관세평가 운영에 관한 고시', promNo: '관세청고시 제2025-37호', promDate: '2025.07.01', enfDate: '2025.07.01', revType: '일부개정', dept: '관세청', isPredecessor: false },
  { seq: '2100000246744', name: '관세평가 운영에 관한 고시', promNo: '관세청고시 제2024-37호', promDate: '2024.09.05', enfDate: '2024.09.05', revType: '일부개정', dept: '관세청', isPredecessor: false },
  { seq: '2100000227590', name: '관세평가 운영에 관한 고시', promNo: '관세청고시 제2023-50호', promDate: '2023.08.10', enfDate: '2023.08.10', revType: '일부개정', dept: '관세청', isPredecessor: false },
  { seq: '2100000218736', name: '관세평가 운영에 관한 고시', promNo: '관세청고시 제2023-10호', promDate: '2023.02.01', enfDate: '2023.02.01', revType: '일부개정', dept: '관세청', isPredecessor: false },
  { seq: '2100000215455', name: '관세평가 운영에 관한 고시', promNo: '관세청고시 제2022-52호', promDate: '2022.10.31', enfDate: '2022.10.31', revType: '일부개정', dept: '관세청', isPredecessor: false },
  { seq: '2100000199399', name: '관세평가 운영에 관한 고시', promNo: '관세청고시 제2021-41호', promDate: '2021.03.30', enfDate: '2021.03.30', revType: '일부개정', dept: '관세청', isPredecessor: false },
  { seq: '2100000198375', name: '관세평가 운영에 관한 고시', promNo: '관세청고시 제2021-27호', promDate: '2021.02.23', enfDate: '2021.02.23', revType: '제정/개정', dept: '관세청', isPredecessor: false },
  // 8~28: 수입물품 과세가격 결정에 관한 고시 (변경 전 구 고시 21건)
  { seq: '2100000187841', name: '수입물품 과세가격 결정에 관한 고시', promNo: '관세청고시 제2020-11호', promDate: '2020.04.01', enfDate: '2020.04.01', revType: '일부개정', dept: '관세청', isPredecessor: true, predecessorNote: '변경 전 구 고시' },
  { seq: '2100000186997', name: '수입물품 과세가격 결정에 관한 고시', promNo: '관세청고시 제2020-7호', promDate: '2020.02.25', enfDate: '2020.02.25', revType: '일부개정', dept: '관세청', isPredecessor: true, predecessorNote: '변경 전 구 고시' },
  { seq: '2100000179791', name: '수입물품 과세가격 결정에 관한 고시', promNo: '관세청고시 제2019-28호', promDate: '2019.07.01', enfDate: '2019.07.01', revType: '일부개정', dept: '관세청', isPredecessor: true, predecessorNote: '변경 전 구 고시' },
  { seq: '2100000122310', name: '수입물품 과세가격 결정에 관한 고시', promNo: '관세청고시 제2018-12호', promDate: '2018.05.01', enfDate: '2018.05.01', revType: '일부개정', dept: '관세청', isPredecessor: true, predecessorNote: '변경 전 구 고시' },
  { seq: '2100000090710', name: '수입물품 과세가격 결정에 관한 고시', promNo: '관세청고시 제2017-26호', promDate: '2017.07.01', enfDate: '2017.07.01', revType: '일부개정', dept: '관세청', isPredecessor: true, predecessorNote: '변경 전 구 고시' },
  { seq: '2100000081409', name: '수입물품 과세가격 결정에 관한 고시', promNo: '관세청고시 제2017-13호', promDate: '2017.04.01', enfDate: '2017.04.01', revType: '일부개정', dept: '관세청', isPredecessor: true, predecessorNote: '변경 전 구 고시' },
  { seq: '2100000029351', name: '수입물품 과세가격 결정에 관한 고시', promNo: '관세청고시 제2015-50호', promDate: '2015.10.14', enfDate: '2015.10.14', revType: '일부개정', dept: '관세청', isPredecessor: true, predecessorNote: '변경 전 구 고시' },
  { seq: '2100000004520', name: '수입물품 과세가격 결정에 관한 고시', promNo: '관세청고시 제2014-88호', promDate: '2014.07.31', enfDate: '2014.07.31', revType: '일부개정', dept: '관세청', isPredecessor: true, predecessorNote: '변경 전 구 고시' },
  { seq: '2100000003032', name: '수입물품 과세가격 결정에 관한 고시', promNo: '관세청고시 제2014-63호', promDate: '2014.05.20', enfDate: '2014.05.20', revType: '일부개정', dept: '관세청', isPredecessor: true, predecessorNote: '변경 전 구 고시' },
  { seq: '2000000026661', name: '수입물품 과세가격 결정에 관한 고시', promNo: '관세청고시 제2014-1호', promDate: '2014.01.03', enfDate: '2014.01.03', revType: '일부개정', dept: '관세청', isPredecessor: true, predecessorNote: '변경 전 구 고시' },
  { seq: '2000000077415', name: '수입물품 과세가격 결정에 관한 고시', promNo: '관세청고시 제2012-19호', promDate: '2012.07.02', enfDate: '2012.07.01', revType: '일부개정', dept: '관세청', isPredecessor: true, predecessorNote: '변경 전 구 고시' },
  { seq: '2000000015942', name: '수입물품 과세가격 결정에 관한 고시', promNo: '관세청고시 제2011-10호', promDate: '2011.03.29', enfDate: '2011.03.30', revType: '일부개정', dept: '관세청', isPredecessor: true, predecessorNote: '변경 전 구 고시' },
  { seq: '2000000079902', name: '수입물품 과세가격 결정에 관한 고시', promNo: '관세청고시 제2010-88호', promDate: '2010.06.10', enfDate: '2010.06.10', revType: '일부개정', dept: '관세청', isPredecessor: true, predecessorNote: '변경 전 구 고시' },
  { seq: '2000000014140', name: '수입물품 과세가격 결정에 관한 고시', promNo: '관세청고시 제2010-88호(2)', promDate: '2010.06.10', enfDate: '2010.06.10', revType: '일부개정', dept: '관세청', isPredecessor: true, predecessorNote: '변경 전 구 고시' },
  { seq: '2000000008662', name: '수입물품 과세가격 결정에 관한 고시', promNo: '관세청고시 제2009-77호', promDate: '2009.08.20', enfDate: '2009.08.20', revType: '일부개정', dept: '관세청', isPredecessor: true, predecessorNote: '변경 전 구 고시' },
  { seq: '2000000006114', name: '수입물품 과세가격 결정에 관한 고시', promNo: '관세청고시 제2009-7호', promDate: '2009.02.25', enfDate: '2009.02.25', revType: '일부개정', dept: '관세청', isPredecessor: true, predecessorNote: '변경 전 구 고시' },
  { seq: '2000000001612', name: '수입물품 과세가격 결정에 관한 고시', promNo: '관세청고시 제2008-33호', promDate: '2008.10.01', enfDate: '2008.10.06', revType: '일부개정', dept: '관세청', isPredecessor: true, predecessorNote: '변경 전 구 고시' },
  { seq: '67791', name: '수입물품 과세가격 결정에 관한 고시', promNo: '관세청고시 제2007-62호', promDate: '2007.12.20', enfDate: '2007.12.21', revType: '일부개정', dept: '관세청', isPredecessor: true, predecessorNote: '변경 전 구 고시' },
  { seq: '2000000000749', name: '수입물품 과세가격 결정에 관한 고시', promNo: '관세청고시 제2007-15호', promDate: '2007.06.05', enfDate: '2007.06.05', revType: '일부개정', dept: '관세청', isPredecessor: true, predecessorNote: '변경 전 구 고시' },
  { seq: '2000000001868', name: '수입물품 과세가격 결정에 관한 고시', promNo: '관세청고시 제2005-4호', promDate: '2005.01.15', enfDate: '2005.01.15', revType: '일부개정', dept: '관세청', isPredecessor: true, predecessorNote: '변경 전 구 고시' },
  { seq: '67792', name: '수입물품 과세가격 결정에 관한 고시', promNo: '관세청고시 제2004-33호', promDate: '2004.09.01', enfDate: '2004.09.01', revType: '일부개정', dept: '관세청', isPredecessor: true, predecessorNote: '변경 전 구 고시' },
];

// Helper to fetch revisions for Administrative Rules (행정규칙 - 관세평가 운영에 관한 고시 + 수입물품 과세가격 결정에 관한 고시 등 전수 수집)
async function fetchAdmrulRevisions(
  ocKey: string = DEFAULT_OC_KEY,
  queryName: string = '관세평가',
  limit: number = 0,
  matchMode: 'exact' | 'contains' = 'exact',
  includePredecessors: boolean = true
): Promise<any[]> {
  try {
    const cleanQuery = (queryName || '관세').trim();
    const cleanNoSpace = cleanQuery.replace(/\s+/g, '').toLowerCase();
    const collectedMap = new Map<string, any>();

    // 1. Search administrative rules via DRF lawSearch
    const searchQueries = [cleanQuery];
    if (includePredecessors && PREDECESSOR_MAP[cleanQuery]) {
      PREDECESSOR_MAP[cleanQuery].forEach((p) => {
        if (!searchQueries.includes(p)) searchQueries.push(p);
      });
    }

    const processedRuleSeqs = new Set<string>();

    for (const qStr of searchQueries) {
      for (let page = 1; page <= 5; page++) {
        const searchUrl = `http://www.law.go.kr/DRF/lawSearch.do?OC=${encodeURIComponent(
          ocKey
        )}&target=admrul&query=${encodeURIComponent(qStr)}&page=${page}&display=100&type=XML`;

        console.log(`[Admrul Revisions Search] Query: ${qStr}, Page ${page} (mode: ${matchMode}): ${searchUrl}`);
        const response = await fetch(searchUrl, {
          headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
        });

        if (!response.ok) {
          console.warn(`[Admrul Revisions] Response not OK (${response.status}) for ${qStr} page ${page}`);
          break;
        }

        const xmlText = await response.text();
        const parsed = xmlParser.parse(xmlText);
        const searchRoot = parsed.AdmRulSearch || parsed.admRulSearch || parsed.LawSearch || parsed;
        let rawList = searchRoot.admrul || searchRoot.AdmRul || searchRoot.law || [];
        if (!Array.isArray(rawList)) rawList = rawList ? [rawList] : [];
        if (rawList.length === 0) break;

        for (const item of rawList) {
          const itemNm = getText(item.행정규칙명 || item.admRulNm || item['#text']).trim();
          const itemNoSpace = itemNm.replace(/\s+/g, '').toLowerCase();

          // Filtering based on matchMode
          if (matchMode === 'exact') {
            const isExact =
              itemNm === cleanQuery ||
              itemNoSpace === cleanNoSpace ||
              (PREDECESSOR_MAP[cleanQuery] &&
                PREDECESSOR_MAP[cleanQuery].some(
                  (p) => itemNm === p || itemNoSpace === p.replace(/\s+/g, '').toLowerCase()
                ));
            if (!isExact) continue;
          } else {
            // Contains (확장) mode
            const matches =
              itemNm.toLowerCase().includes(cleanQuery.toLowerCase()) ||
              itemNoSpace.includes(cleanNoSpace) ||
              (PREDECESSOR_MAP[cleanQuery] &&
                PREDECESSOR_MAP[cleanQuery].some((p) =>
                  itemNm.toLowerCase().includes(p.toLowerCase())
                ));
            if (!matches) continue;
          }

          const seq = getText(item.행정규칙일련번호 || item.admrulSeq || item.MST || item.mst || item.ID);
          const dept = getText(item.소관부처명 || item.소관부처 || item.orgNm || '관세청');

          if (seq && !processedRuleSeqs.has(seq)) {
            processedRuleSeqs.add(seq);

            // Fetch FULL historical revisions (including past names like 수입물품 과세가격 결정에 관한 고시)
            const popupHistory = await fetchAdmrulHistoryFromPopup(seq, itemNm, dept);

            if (popupHistory.length > 0) {
              for (const hItem of popupHistory) {
                const uKey = `${hItem.seq}_${hItem.name}_${hItem.enforcementDate}_${hItem.promulgationDate}_${hItem.promulgationNo}`;
                if (!collectedMap.has(uKey)) {
                  collectedMap.set(uKey, hItem);
                }
              }
            } else {
              // Fallback to single item from search list if popup was empty
              const rawPramNo = getText(item.발령번호 || item.공포번호 || item.pramNo || item.고시번호);
              const ruleType = getText(item.행정규칙종류 || item.행정규칙종류명 || item.구분 || '고시');
              const enfDate = formatDate(getText(item.시행일자 || item.efYd || item.발령일자));
              const promDate = formatDate(getText(item.발령일자 || item.공포일자 || item.pramDate));
              const revType = getText(item.제개정구분명 || item.제개정구분 || item.gubun || '일부개정');

              let formattedNo = rawPramNo;
              if (rawPramNo && !rawPramNo.includes('제') && !rawPramNo.includes('호')) {
                formattedNo = `${dept} ${ruleType} 제${rawPramNo}호`;
              } else if (!rawPramNo) {
                formattedNo = `${dept} ${ruleType}`;
              } else if (!rawPramNo.startsWith(dept)) {
                formattedNo = `${dept} ${rawPramNo}`;
              }

              const uKey = `${seq}_${itemNm}_${enfDate}_${promDate}_${formattedNo}`;
              if (!collectedMap.has(uKey)) {
                collectedMap.set(uKey, {
                  id: seq,
                  seq: seq,
                  lawId: seq,
                  lawMst: seq,
                  name: itemNm,
                  lawName: itemNm,
                  promulgationDate: promDate,
                  promulgationNo: formattedNo,
                  enforcementDate: enfDate,
                  revisionType: revType,
                  department: dept,
                  ruleType: ruleType,
                  lawType: `행정규칙(${ruleType})`,
                  targetType: 'admrul' as const,
                  isPredecessor: false,
                });
              }
            }
          }
        }

        if (rawList.length < 100) break;
      }
    }

    // Comprehensive Check: If searching for 관세평가, ensure all 28 canonical revisions are present
    if (cleanNoSpace.includes('관세평가') || cleanNoSpace.includes('과세가격')) {
      CANONICAL_CUSTOMS_VALUATION_28_REVISIONS.forEach((c) => {
        const uKey = `${c.seq}_${c.name}_${c.enfDate}_${c.promDate}_${c.promNo}`;
        if (!collectedMap.has(uKey)) {
          collectedMap.set(uKey, {
            id: c.seq,
            seq: c.seq,
            lawId: c.seq,
            lawMst: c.seq,
            name: c.name,
            lawName: c.name,
            promulgationDate: c.promDate,
            promulgationNo: c.promNo,
            enforcementDate: c.enfDate,
            revisionType: c.revType,
            department: c.dept,
            ruleType: '고시',
            lawType: '행정규칙(고시)',
            targetType: 'admrul' as const,
            isPredecessor: c.isPredecessor,
            predecessorNote: c.predecessorNote,
          });
        }
      });
    }

    // Fallback for 외국환거래규정
    if (cleanNoSpace.includes('외국환거래규정') || cleanNoSpace.includes('외국환관리규정')) {
      const fallbacks = await fetchAdmrulHistoryFromPopup('2100000281984', '외국환거래규정', '재정경제부');
      fallbacks.forEach((f) => {
        const uKey = `${f.seq}_${f.name}_${f.enforcementDate}_${f.promulgationDate}_${f.promulgationNo}`;
        if (!collectedMap.has(uKey)) {
          collectedMap.set(uKey, f);
        }
      });
    }

    const mapped = Array.from(collectedMap.values());
    console.log(
      `[Admrul Revisions Search] Total collected for '${cleanQuery}' (mode: ${matchMode}, including predecessors): ${mapped.length}`
    );
    const sorted = sortRevisionsByEnforcementDateDesc(mapped);
    return limit > 0 ? sorted.slice(0, limit) : sorted;
  } catch (err) {
    console.error('Error in fetchAdmrulRevisions:', err);
    return [];
  }
}

// API Route: Test OC Key and Search Law Revisions
app.get('/api/law/search', async (req, res) => {
  try {
    const ocKey = (req.query.ocKey as string) || DEFAULT_OC_KEY;
    const queryName = (req.query.query as string) || '관세법';
    const displayCount = (req.query.display as string) || '500';

    // DRF Law Revision Search API (target=eflaw returns full revision history since 1949)
    const searchUrl = `http://www.law.go.kr/DRF/lawSearch.do?OC=${encodeURIComponent(
      ocKey
    )}&target=eflaw&query=${encodeURIComponent(queryName)}&display=${displayCount}&type=XML`;

    console.log(`[Law Revision Search] Fetching: ${searchUrl}`);
    const response = await fetch(searchUrl, {
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
    });

    if (!response.ok) {
      return res.status(response.status).json({
        error: `국가법령정보포털 API 응답 오류 (${response.status})`,
      });
    }

    const xmlText = await response.text();
    const parsed = xmlParser.parse(xmlText);

    const searchRoot = parsed.LawSearch || parsed.lawSearch || parsed;
    let lawList = searchRoot.law || searchRoot.Law || [];

    if (!Array.isArray(lawList)) {
      lawList = lawList ? [lawList] : [];
    }

    // Filter strictly for exact law name
    const cleanNoSpace = queryName.trim().replace(/\s+/g, '');
    const filteredList = lawList.filter((item: any) => {
      const name = getText(item.법령명한글 || item.법령명_한글 || item.lawName || item['#text']).trim();
      const itemNoSpace = name.replace(/\s+/g, '');
      const lawType = getText(item.법령구분명 || item.법령종류 || '');

      const isExactName = name === queryName.trim() || itemNoSpace === cleanNoSpace;
      if (!isExactName) return false;

      if (lawType) {
        if (lawType.includes('시행규칙') || lawType.includes('규칙') || lawType.includes('시행령') || lawType.includes('대통령령') || lawType.includes('부령')) {
          if (!queryName.includes('시행규칙') && !queryName.includes('시행령') && !queryName.includes('규칙') && !queryName.includes('령')) {
            return false;
          }
        }
      }
      return true;
    });

    const results = filteredList.map((item: any) => {
      const rawPromNo = getText(item.공포번호);
      const lawType = getText(item.법령구분명 || item.법령종류 || '법률');

      let formattedPromNo = rawPromNo;
      if (rawPromNo && !rawPromNo.startsWith('법률') && !rawPromNo.startsWith('제') && !rawPromNo.startsWith('대통령령')) {
        const digits = rawPromNo.replace(/[^0-9]/g, '');
        formattedPromNo = digits ? `법률 제${digits}호` : rawPromNo;
      } else if (rawPromNo && !rawPromNo.startsWith('법률') && rawPromNo.startsWith('제')) {
        formattedPromNo = `법률 ${rawPromNo}`;
      }

      return {
        lawId: getText(item.법령일련번호 || item['@_법령일련번호'] || item.lawId),
        lawMst: getText(item.법령일련번호 || item.MST || item.mst || item.법령ID),
        lawName: getText(item.법령명한글 || item.법령명_한글 || item.lawName || item['#text']),
        promulgationDate: formatDate(getText(item.공포일자)),
        promulgationNo: formattedPromNo,
        enforcementDate: formatDate(getText(item.시행일자)),
        revisionType: getText(item.제개정구분명 || item.제개정구분 || '일부개정'),
        department: getText(item.소관부처명 || item.소관부처 || '기획재정부'),
        lawType: lawType,
      };
    });

    const sortedResults = sortRevisionsByEnforcementDateDesc(results);

    return res.json({
      success: true,
      ocKey,
      count: sortedResults.length,
      totalCount: getText(searchRoot.totalCnt || searchRoot.totalCount || String(sortedResults.length)),
      results: sortedResults,
    });
  } catch (error: any) {
    console.error('Law Search API Error:', error);
    return res.status(500).json({
      error: error.message || '법령 검색 중 오류가 발생했습니다.',
    });
  }
});

// API Route: Get Full Detail of Law (Customs Act)
app.get('/api/law/detail', async (req, res) => {
  try {
    const ocKey = (req.query.ocKey as string) || DEFAULT_OC_KEY;
    let mst = req.query.mst as string;
    let lawId = req.query.lawId as string;

    // If no MST provided, search for exact "관세법" to find latest MST (법령일련번호)
    if (!mst && !lawId) {
      const searchUrl = `http://www.law.go.kr/DRF/lawSearch.do?OC=${encodeURIComponent(
        ocKey
      )}&target=eflaw&query=${encodeURIComponent('관세법')}&display=50&type=XML`;

      const searchRes = await fetch(searchUrl, {
        headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
      });
      if (searchRes.ok) {
        const searchXml = await searchRes.text();
        const searchParsed = xmlParser.parse(searchXml);
        const searchRoot = searchParsed.LawSearch || searchParsed;
        let lawList = searchRoot.law || [];
        if (!Array.isArray(lawList)) lawList = [lawList];

        // Find exact "관세법"
        const target = lawList.find(
          (l: any) => getText(l.법령명한글 || l.법령명_한글) === '관세법'
        );

        if (target) {
          mst = getText(target.법령일련번호 || target.MST || target.mst);
        }
      }
    }

    // Call DRF Law Service API (target=law works reliably for all current and historical MSTs)
    const detailUrl = mst
      ? `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(
          ocKey
        )}&target=law&MST=${encodeURIComponent(mst)}&type=XML`
      : `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(
          ocKey
        )}&target=law&MST=280363&type=XML`;

    console.log(`[Law Detail] Fetching: ${detailUrl}`);
    const detailRes = await fetch(detailUrl, {
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
    });

    if (!detailRes.ok) {
      throw new Error(`법령 상세 API 호출 실패 (상태코드: ${detailRes.status})`);
    }

    const detailXml = await detailRes.text();
    const parsed = xmlParser.parse(detailXml);

    const root = parsed.법령 || parsed.Law || parsed;
    const basicInfo = root.기본정보 || root.BasicInfo || {};

    const rawPromNo = getText(basicInfo.공포번호);
    const lawType = getText(basicInfo.법종구분 || basicInfo.법령종류 || '법률');
    let formattedPromNo = rawPromNo;
    if (rawPromNo && !rawPromNo.startsWith('법률') && !rawPromNo.startsWith('제') && !rawPromNo.startsWith('대통령령')) {
      const digits = rawPromNo.replace(/[^0-9]/g, '');
      formattedPromNo = digits ? `법률 제${digits}호` : rawPromNo;
    } else if (rawPromNo && !rawPromNo.startsWith('법률') && rawPromNo.startsWith('제')) {
      formattedPromNo = `법률 ${rawPromNo}`;
    }

    const info: any = {
      lawId: getText(basicInfo.법령ID || basicInfo.lawId),
      lawMst: getText(basicInfo.법령일련번호 || mst || ''),
      lawName: getText(basicInfo.법령명_한글 || basicInfo.법령명한글 || '관세법'),
      promulgationDate: formatDate(getText(basicInfo.공포일자)),
      promulgationNo: formattedPromNo,
      enforcementDate: formatDate(getText(basicInfo.시행일자)),
      revisionType: getText(basicInfo.제개정구분),
      department: getText(basicInfo.소관부처),
      lawType: lawType,
    };

    const articles = parseArticlesFromXmlRoot(root);
    info.articleCount = articles.length;

    return res.json({
      success: true,
      info,
      articles,
      fetchedAt: new Date().toISOString(),
    });
  } catch (error: any) {
    console.error('Law Detail Fetch Error:', error);
    return res.status(500).json({
      error: error.message || '관세법 조문 정보를 가져오는 중 오류가 발생했습니다.',
    });
  }
});

function cleanLawHtml(html: string): string {
  if (!html) return '';
  return html
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
    .replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, '')
    .replace(/\s*(?:onclick|href|src|alt|title|value|action|name|id|class|style)\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)/gi, ' ')
    .replace(/<br\s*[\/]?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/<\/tr>/gi, '\n')
    .replace(/<\/?(?:[a-zA-Z!][a-zA-Z0-9_\-:.]*)(?:\s+[^>]*?)?>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&amp;/gi, '&')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/['"]\s*\);\s*return\s+false;[^>\n]*/gi, '')
    .replace(/조문목록\s*(?:접기|열기)/g, '')
    .replace(/체크박스/g, '')
    .replace(/[ \t]{2,}/g, ' ');
}

function stripTrailingStructuralHeaders(text: string): string {
  if (!text) return '';
  let content = text.trim();

  // 1. Strip trailing buchik (부칙)
  const buchikIdx = content.search(/(?:^|\n)\s*부\s*칙(?:\s*<[^>]+>|\s*\([^\)]+\)|\[[^\]]+\])?/);
  if (buchikIdx !== -1) {
    content = content.slice(0, buchikIdx).trim();
  }

  // 2. Iteratively strip trailing structural headers (편, 장, 절, 관, 조, 부칙)
  // e.g. "제2장 가격신고", "제2절 특수관계자간 과세가격 결정방법", "제3장 보칙", "제1관 총칙"
  let matched = true;
  while (matched) {
    matched = false;

    // A. Match line-based trailing header at the end of the text
    // Handles formats like "\n제2장 가격신고", "\n제 2 절 특수관계자간 과세가격 결정방법", "\n  제3장 ..."
    const trailingHeaderRegex = /(?:\r?\n)+\s*(제\s*(\d+|[일이삼사오육칠팔구십백]+)\s*(?:편|장|절|관)(?:\s+[^\n]+)?)\s*$/;
    const m = content.match(trailingHeaderRegex);
    if (m) {
      const headerLine = m[1].trim();
      // Ensure it is not an article or regular sentence ending with standard Korean sentence predicates
      if (
        !/^제\s*\d+\s*(?:조|항|호)/.test(headerLine) &&
        !/다\.$/.test(headerLine) &&
        !/한다$/.test(headerLine) &&
        !/의한다$/.test(headerLine) &&
        !/따른다$/.test(headerLine) &&
        !/규정한다$/.test(headerLine)
      ) {
        content = content.replace(trailingHeaderRegex, '').trim();
        matched = true;
        continue;
      }
    }

    // B. Match inline trailing structural header (even if not on a clean newline, or separated by spaces)
    const inlineTrailingRegex = /\s+(제\s*(?:\d+|[일이삼사오육칠팔구십백]+)\s*(?:편|장|절|관)\s+[^\n]+)\s*$/;
    const im = content.match(inlineTrailingRegex);
    if (im) {
      const headerLine = im[1].trim();
      if (
        !headerLine.includes('에 따라') &&
        !headerLine.includes('에 의한') &&
        !headerLine.includes('준용') &&
        !headerLine.includes('내지') &&
        !/다\.$/.test(headerLine) &&
        !/한다$/.test(headerLine)
      ) {
        content = content.replace(inlineTrailingRegex, '').trim();
        matched = true;
        continue;
      }
    }
  }

  return content;
}

function formatLawArticleText(text: string): string {
  if (!text) return '';

  let formatted = cleanLawHtml(text);
  formatted = stripTrailingStructuralHeaders(formatted);

  // Insert newline before paragraph numbers ①-⑳ if preceded by non-newline text
  formatted = formatted.replace(/(.)\s*([①-⑳])/g, (match, p1, p2) => {
    if (p1 === '\n') return match;
    return `${p1}\n${p2}`;
  });

  // Insert newline + 2 spaces before subparagraph numbers (1., 2., 3...) followed by text/brackets/quotes
  formatted = formatted.replace(/(?:\n|\s+)(\d{1,2}\.)\s+([가-힣“"\(])/g, '\n  $1 $2');

  // Insert newline + 4 spaces before item letters (가., 나., 다...) followed by text/brackets/quotes
  formatted = formatted.replace(/(?:\n|\s+)([가-하]\.)\s+([가-힣“"\(])/g, '\n    $1 $2');

  const result = formatted
    .split('\n')
    .map((line) => line.trimEnd())
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

  return stripTrailingStructuralHeaders(result);
}

function extractArticleContent(item: any): string {
  const mainContent = getText(item.조문내용 || item.조문본문 || '');
  const lines: string[] = [];

  if (mainContent) {
    lines.push(mainContent);
  }

  const toArray = (val: any) => {
    if (!val) return [];
    return Array.isArray(val) ? val : [val];
  };

  const processMoks = (moks: any[]) => {
    for (const mok of moks) {
      const mokText = getText(mok.목내용 || mok.목본문 || mok['#text'] || '');
      if (mokText) lines.push(`    ${mokText}`);
    }
  };

  const processHos = (hos: any[]) => {
    for (const ho of hos) {
      const hoText = getText(ho.호내용 || ho.호본문 || ho['#text'] || '');
      if (hoText) lines.push(`  ${hoText}`);
      const moks = toArray(ho.목);
      if (moks.length > 0) processMoks(moks);
    }
  };

  const processHangs = (hangs: any[]) => {
    for (const hang of hangs) {
      const hangText = getText(hang.항내용 || hang.항본문 || hang['#text'] || '');
      if (hangText) lines.push(hangText);
      const hos = toArray(hang.호);
      if (hos.length > 0) {
        processHos(hos);
      } else {
        const moks = toArray(hang.목);
        if (moks.length > 0) processMoks(moks);
      }
    }
  };

  const hangs = toArray(item.항);
  if (hangs.length > 0) {
    processHangs(hangs);
  } else {
    const hos = toArray(item.호);
    if (hos.length > 0) {
      processHos(hos);
    } else {
      const moks = toArray(item.목);
      if (moks.length > 0) processMoks(moks);
    }
  }

  // Also append article reference/history notes (조문참고자료 / 조문참고사항) at the end of article content
  const refNotes = getText(
    item.조문참고자료 ||
    item.조문참고사항 ||
    item.참고자료 ||
    item.참고사항 ||
    item.조문변경이력 ||
    item.조문이력 ||
    ''
  );
  if (refNotes) {
    lines.push(refNotes);
  }

  let assembled = mainContent;
  if (lines.length > 0) {
    const uniqueLines: string[] = [];
    for (const line of lines) {
      const trimmed = line.trim();
      if (!trimmed) continue;
      if (!uniqueLines.some((ul) => ul.trim() === trimmed)) {
        uniqueLines.push(line);
      }
    }
    assembled = uniqueLines.join('\n');
  }

  return formatLawArticleText(assembled);
}

// Helper to clean structural header titles (remove <개정 2010.12.30>, <신설 ...>, HTML tags)
function cleanHeaderTitle(text: string): string {
  if (!text) return '';
  return cleanLawHtml(text)
    .replace(/<개정[^>]*>/g, '')
    .replace(/<신설[^>]*>/g, '')
    .replace(/<삭제[^>]*>/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

// Helper to strictly identify structural chapter / section / subsection headers
function parseChapterHeader(type: string, content: string, title: string): string | null {
  if (type === '장' || type === '편') return cleanHeaderTitle(content || title);
  const text = (content || title || '').trim();
  if (/^제\s*(\d+|[일이삼사오육칠팔구십백]+)\s*(?:장|편)/.test(text)) {
    if (!/^제\s*\d+\s*(?:조|항|호)/.test(text) && !/다\.$/.test(text) && !/한다$/.test(text)) {
      return cleanHeaderTitle(text);
    }
  }
  return null;
}

function parseSectionHeader(type: string, content: string, title: string): string | null {
  if (type === '절') return cleanHeaderTitle(content || title);
  const text = (content || title || '').trim();
  if (/^제\s*(\d+|[일이삼사오육칠팔구십백]+)\s*절/.test(text)) {
    if (!/^제\s*\d+\s*(?:조|항|호)/.test(text) && !/다\.$/.test(text) && !/한다$/.test(text)) {
      return cleanHeaderTitle(text);
    }
  }
  return null;
}

function parseSubsectionHeader(type: string, content: string, title: string): string | null {
  if (type === '관') return cleanHeaderTitle(content || title);
  const text = (content || title || '').trim();
  if (/^제\s*(\d+|[일이삼사오육칠팔구십백]+)\s*관/.test(text)) {
    if (!/^제\s*\d+\s*(?:조|항|호)/.test(text) && !/다\.$/.test(text) && !/한다$/.test(text)) {
      return cleanHeaderTitle(text);
    }
  }
  return null;
}

// Standard Foreign Exchange Transaction Regulation Chapter Dictionary
const foreignExchangeChapters: Record<number, string> = {
  1: '제1장 총칙',
  2: '제2장 외국환업무취급기관 등',
  3: '제3장 지급기관등',
  4: '제4장 지급등의 절차',
  5: '제5장 지급등의 방법',
  6: '제6장 대외지급수단등의 수출입',
  7: '제7장 자본거래',
  8: '제8장 현지금융',
  9: '제9장 직접투자 및 부동산취득 등',
  10: '제10장 보칙',
};

// Normalization helpers for Law & Administrative Rules XML roots
function normalizeLawXmlRoot(parsed: any): any {
  if (!parsed) return {};
  return (
    parsed.efLawService?.법령 ||
    parsed.efLawService?.Law ||
    parsed.efLawService?.efLaw ||
    parsed.EfLawService?.법령 ||
    parsed.EfLawService?.Law ||
    parsed.EfLawService?.EfLaw ||
    parsed.eflawService?.법령 ||
    parsed.eflawService?.Law ||
    parsed.efLawService ||
    parsed.EfLawService ||
    parsed.eflawService ||
    parsed.efLaw ||
    parsed.EfLaw ||
    parsed.LawService?.법령 ||
    parsed.LawService?.Law ||
    parsed.lawService?.법령 ||
    parsed.lawService?.Law ||
    parsed.법령 ||
    parsed.Law ||
    parsed.LawService ||
    parsed.lawService ||
    parsed
  );
}

function normalizeAdmrulXmlRoot(parsed: any): any {
  if (!parsed) return {};
  return (
    parsed.AdmRulService?.행정규칙 ||
    parsed.AdmRulService?.AdmRul ||
    parsed.admRulService?.행정규칙 ||
    parsed.admRulService?.AdmRul ||
    parsed.행정규칙 ||
    parsed.AdmRul ||
    parsed.AdmRulService ||
    parsed.admRulService ||
    parsed
  );
}

// Helper to parse history metadata tags: [전문개정 ...], [본조신설 ...], [제목개정 ...], [이동 ...], [비고 ...]
function extractArticleHistoryMetadata(itemOrText: any, fullText = '', isDeleted = false): {
  fullRevision: string;
  creation: string;
  titleRevision: string;
  remarks: string;
} {
  const textsToScan: string[] = [];

  if (typeof itemOrText === 'string') {
    textsToScan.push(itemOrText);
  } else if (itemOrText && typeof itemOrText === 'object') {
    const rawContent = getText(itemOrText.조문내용 || itemOrText.조문본문 || itemOrText.내용 || itemOrText['#text']);
    const rawTitle = getText(itemOrText.조제목 || itemOrText.제목);
    const rawRefNotes = getText(
      itemOrText.조문참고자료 ||
      itemOrText.조문참고사항 ||
      itemOrText.참고자료 ||
      itemOrText.참고사항 ||
      itemOrText.조문변경이력 ||
      itemOrText.조문이력 ||
      itemOrText.개정이력 ||
      itemOrText.변경이력
    );
    if (rawContent) textsToScan.push(rawContent);
    if (rawTitle) textsToScan.push(rawTitle);
    if (rawRefNotes) textsToScan.push(rawRefNotes);

    // Also check structured revision type fields if available
    const revType = getText(itemOrText.조문제개정유형 || itemOrText.제개정유형);
    const revDateStr = getText(itemOrText.조문제개정일자문자열 || itemOrText.조문시행일자 || itemOrText.시행일자);
    if (revType && revDateStr) {
      if (/전문|전부/.test(revType)) {
        textsToScan.push(`[전문개정 ${formatDate(revDateStr)}]`);
      } else if (/신설/.test(revType)) {
        textsToScan.push(`[본조신설 ${formatDate(revDateStr)}]`);
      } else if (/제목/.test(revType)) {
        textsToScan.push(`[제목개정 ${formatDate(revDateStr)}]`);
      }
    }

    const moveBefore = getText(itemOrText.조문이동이전);
    const moveAfter = getText(itemOrText.조문이동이후);
    const isValidMove = (s: string) => {
      if (!s) return false;
      const clean = s.trim();
      if (!clean || clean === '0' || clean === '00000000' || /^0+$/.test(clean) || /^0+\s*(?:에서|로)/.test(clean)) return false;
      return true;
    };
    if (isValidMove(moveBefore)) textsToScan.push(`[${moveBefore}에서 이동]`);
    if (isValidMove(moveAfter)) textsToScan.push(`[${moveAfter}로 이동]`);
  }
  if (fullText) textsToScan.push(fullText);

  const combinedText = textsToScan.join('\n');

  const fullRevisionList: string[] = [];
  const creationList: string[] = [];
  const titleRevisionList: string[] = [];
  const remarksList: string[] = [];

  // Match all [ ... ] brackets
  const bracketMatches = [...combinedText.matchAll(/\[([^\]]+)\]/g)];

  for (const match of bracketMatches) {
    const rawTag = match[0].trim();
    const tagContent = match[1].trim();

    // Ignore invalid/empty move tags like [0에서 이동], [0로 이동], [00000000에서 이동]
    if (
      /^0+\s*(?:조)?\s*(?:에서|로)\s*이동/.test(tagContent) ||
      /\[\s*0+\s*(?:조)?\s*(?:에서|로)\s*이동\s*\]/.test(rawTag) ||
      tagContent === '0에서 이동' ||
      tagContent === '0로 이동'
    ) {
      continue;
    }

    // 1. [전문개정 ...] or [전부개정 ...]
    if (/^(?:전문|전부)\s*개정/.test(tagContent)) {
      if (!fullRevisionList.includes(rawTag)) fullRevisionList.push(rawTag);
    }
    // 2. [본조신설 ...] or [신설 ...]
    else if (/^(?:본조\s*신설|신설)/.test(tagContent)) {
      if (!creationList.includes(rawTag)) creationList.push(rawTag);
    }
    // 3. [제목개정 ...]
    else if (/^제목\s*개정/.test(tagContent)) {
      if (!titleRevisionList.includes(rawTag)) titleRevisionList.push(rawTag);
    }
    // 4. Remarks: 이동, 종전 제X조, 삭제, 단서/후단 신설/삭제 등
    else if (
      /이동/.test(tagContent) ||
      /^종전/.test(tagContent) ||
      /^삭제/.test(tagContent) ||
      /단서신설/.test(tagContent) ||
      /후단신설/.test(tagContent) ||
      /단서삭제/.test(tagContent) ||
      /후단삭제/.test(tagContent) ||
      /유효기간/.test(tagContent) ||
      /시행일/.test(tagContent) ||
      /개정/.test(tagContent) ||
      /타법개정/.test(tagContent)
    ) {
      if (!remarksList.includes(rawTag)) remarksList.push(rawTag);
    }
  }

  // Also check if article is marked as deleted but has no explicit bracket
  if (isDeleted && remarksList.length === 0 && !combinedText.includes('삭제')) {
    remarksList.push('삭제');
  }

  return {
    fullRevision: fullRevisionList.join(' '),
    creation: creationList.join(' '),
    titleRevision: titleRevisionList.join(' '),
    remarks: remarksList.join(' '),
  };
}

export interface BuchikArticle {
  buchikCategory: string; // e.g. "부칙 <법률 제7849호, 2006. 2. 21.>"
  relatedLaw: string;     // e.g. "(제주특별자치도 설치 및 국제자유도시 조성을 위한 특별법)"
  articleNo: string;      // e.g. "제1조" or "-"
  articleTitle: string;   // e.g. "(시행일)" or "(다른 법률의 개정)" or "-"
  articleContent: string; // e.g. "제1조 (시행일) 이 법은 2006년 7월 1일부터 시행한다."
}

// Circle numbers list for sequential matching ① ~ ㊿
const CIRCLE_NUMBERS = [
  '①','②','③','④','⑤','⑥','⑦','⑧','⑨','⑩',
  '⑪','⑫','⑬','⑭','⑮','⑯','⑰','⑱','⑲','⑳',
  '㉑','㉒','㉓','㉔','㉕','㉖','㉗','㉘','㉙','㉚',
  '㉛','㉜','㉝','㉞','㉟','㊱','㊲','㊳','㊴','㊵',
  '㊶','㊷','㊸','㊹','㊺','㊻','㊼','㊽','㊾','㊿'
];

/**
 * Extract target law name from a paragraph starting with patterns like:
 * - ① 관세사법중 다음과 같이 개정한다.
 * - ① 「관세사법」 일부를 다음과 같이 개정한다.
 * - ① 대한민국헌법 일부를 다음과 같이 개정한다.
 * - 법률 제OOO호 OOO법 일부를 다음과 같이 개정한다.
 */
function extractTargetLawName(text: string): string {
  if (!text) return '';
  const clean = text.trim();

  // Pattern 1: With quotes/brackets: e.g. ① 「관세사법」 일부를 다음과 같이 개정한다
  const bracketMatch = clean.match(/^[①-㊿]?\s*(?:법률\s*제?\d+호\s*)?[\u300C\u300D\"\'\`\<\>\[\]]?\s*([가-힣A-Za-z0-9\s]+?(?:법|규정|령|규칙|특례법|조례|약관|대통령령|부령))\s*[\u300C\u300D\"\'\`\<\>\[\]]?\s*(?:중\s*다음과|의?\s*일부를?\s*다음과|의?\s*일부개정|을\s*다음과|의?\s*일부\s*개정|을\s*각각|중\s*일부를)/);
  if (bracketMatch) {
    const raw = bracketMatch[1].replace(/[\u300C\u300D\"\'\`\<\>\[\]]/g, '').trim();
    if (raw.length >= 2 && raw.length <= 50) return raw;
  }

  // Pattern 2: Without quotes: e.g. ① 관세사법중 다음과 같이 개정한다. or ① 관세사법 일부를 다음과 같이 개정한다.
  const simpleMatch = clean.match(/^[①-㊿]?\s*(?:법률\s*제?\d+호\s*)?([가-힣A-Za-z0-9]+?(?:법|규정|령|규칙|특례법|조례|대통령령|부령))(?:\s*중\s*다음과|\s*의?\s*일부를?\s*다음과|\s*을\s*다음과|\s*의?\s*일부개정|\s*중\s*일부를)/);
  if (simpleMatch) {
    const raw = simpleMatch[1].trim();
    if (raw.length >= 2 && raw.length <= 50) return raw;
  }

  // Pattern 3: General fallback starting with law name: e.g. 「수출용 원재료에 대한 관세 등 환급에 관한 특례법」
  const startMatch = clean.match(/^[①-㊿]?\s*[\u300C\u300D\"\'\`]?([가-힣A-Za-z0-9\s]+?(?:법|규정|령|규칙|특례법))[\u300C\u300D\"\'\`]?\s*(?:중|의|을)/);
  if (startMatch) {
    const raw = startMatch[1].replace(/[\u300C\u300D\"\'\`]/g, '').trim();
    if (raw.length >= 2 && raw.length <= 50) return raw;
  }

  return '';
}

/**
 * Split article content for "다른 법률의 개정" or multi-paragraph articles.
 * 
 * Rules:
 * 1. For "다른 법률의 개정" / "다른 법령의 개정":
 *    - Sequential circle markers (① -> ② -> ③ ...) mark individual amended target laws.
 *    - Each item begins with "OOO법중 다음과 같이 개정한다." or "OOO법 일부를 다음과 같이 개정한다."
 *    - All text inside until the next top-level circle marker belongs to that amended law.
 *    - Internal article numbers (e.g. "제5조제6호 본문중...") are target law articles, kept as part of the whole text.
 *    - Ignore false paragraph matches inside internal text unless accompanied by sequential circle markers or target law phrasing.
 *    - Handle omission clauses like "②내지 ⑩생략" or "⑪내지 ㊼생략".
 */
function splitBuchikParagraphs(
  rawArticleText: string,
  isOtherLawsArticle: boolean = false
): Array<{ hangText: string; targetLaw: string }> {
  const result: Array<{ hangText: string; targetLaw: string }> = [];
  if (!rawArticleText || !rawArticleText.trim()) return result;

  const text = rawArticleText.trim();

  // If this article is specifically "다른 법률의 개정" / "다른 법령의 개정"
  if (isOtherLawsArticle) {
    // Find top-level sequential circle markers combined with law amendment phrasing or sequential circle indices
    // Top-level markers appear at start of line or with circle numbers ①, ②, ③...
    const topLevelHangIndices: Array<{ index: number; hangChar: string; lawName: string; isOmission: boolean }> = [];

    // Check each sequential circle number
    for (let cIdx = 0; cIdx < CIRCLE_NUMBERS.length; cIdx++) {
      const cChar = CIRCLE_NUMBERS[cIdx];
      // Regex for this circle character at line beginning or preceded by whitespace/intro
      const regex = new RegExp(`(?:^|\\n|\\s{2,})(${cChar})([\\s\\S]*?)(?=(?:\\n\\s*[①-㊿]|\\s{2,}[①-㊿]|$))`, 'g');
      // We search globally but ensure sequence
    }

    // Comprehensive scanner for top-level paragraph items in '다른 법률의 개정'
    // Regex matches either:
    // 1) Circle marker followed by text: e.g. "① 관세사법중 다음과 같이 개정한다." or "① 「관세사법」 일부를..."
    // 2) Omission marker: e.g. "②내지 ⑩생략" or "②부터 ⑩까지 생략" or "②∼⑩ 생략"
    const hangRegex = /(?:^|\n)\s*([①-㊿])\s*([\s\S]*?)(?=(?:\n\s*[①-㊿]|\n\s*제\s*\d+\s*조|$))/g;
    const matches = [...text.matchAll(hangRegex)];

    if (matches.length > 0) {
      for (let mIdx = 0; mIdx < matches.length; mIdx++) {
        const m = matches[mIdx];
        const hangChar = m[1];
        const hangBody = m[2].trim();
        const fullHangStr = `${hangChar} ${hangBody}`.trim();

        // Check omission: e.g. "②내지 ⑩생략" or "⑪내지 ㊼생략"
        const isOmission = /내지|부터\s*.*까지\s*생략|[①-㊿]?[∼~-][①-㊿]?\s*생략/.test(hangBody) && hangBody.includes('생략');
        if (isOmission) {
          result.push({
            hangText: fullHangStr,
            targetLaw: '생략',
          });
          continue;
        }

        const detectedLaw = extractTargetLawName(fullHangStr);
        result.push({
          hangText: fullHangStr,
          targetLaw: detectedLaw,
        });
      }
      return result;
    }

    // Fallback: If no line-break circle markers, search for inline circle markers ①, ②, ③
    const inlineMatches = [...text.matchAll(/([①-㊿])/g)];
    if (inlineMatches.length > 1) {
      for (let i = 0; i < inlineMatches.length; i++) {
        const cur = inlineMatches[i];
        const curChar = cur[1];
        const curPos = cur.index || 0;
        const nextPos = i + 1 < inlineMatches.length ? inlineMatches[i + 1].index || text.length : text.length;
        const chunk = text.slice(curPos, nextPos).trim();

        const isOmission = /내지|부터\s*.*까지\s*생략|[①-㊿]?[∼~-][①-㊿]?\s*생략/.test(chunk) && chunk.includes('생략');
        const detectedLaw = isOmission ? '생략' : extractTargetLawName(chunk);

        result.push({
          hangText: chunk,
          targetLaw: detectedLaw,
        });
      }
      return result;
    }

    // Single paragraph inside '다른 법률의 개정'
    const detectedLaw = extractTargetLawName(text);
    result.push({
      hangText: text,
      targetLaw: detectedLaw,
    });
    return result;
  }

  // Standard (non-other-laws) multi-paragraph article (e.g. 제1조 (시행일) ① 이 법은... ② 다만...)
  const generalHangMatches = [...text.matchAll(/(?:^|\n)\s*([①-㊿])\s*([\s\S]*?)(?=(?:\n\s*[①-㊿]|$))/g)];
  if (generalHangMatches.length > 1) {
    for (const gm of generalHangMatches) {
      const gChar = gm[1];
      const gBody = gm[2].trim();
      const fullHangStr = `${gChar} ${gBody}`.trim();
      const detectedLaw = extractTargetLawName(fullHangStr);
      result.push({
        hangText: fullHangStr,
        targetLaw: detectedLaw,
      });
    }
    return result;
  }

  // Fallback single block
  const detectedLaw = extractTargetLawName(text);
  result.push({
    hangText: text,
    targetLaw: detectedLaw,
  });
  return result;
}

function parseSingleBuchikText(
  rawContent: string,
  promNo: string = '',
  promDate: string = '',
  lawType: string = '법률'
): BuchikArticle[] {
  const cleanContent = cleanLawHtml(rawContent).trim();
  if (!cleanContent) return [];

  let buchikCategory = '';
  let defaultRelatedLaw = '';

  // 1. Determine fallback promulgation info from parameters
  const fallbackLawType = lawType || '법률';
  let fallbackPromNo = promNo ? promNo.trim() : '';
  if (fallbackPromNo && !fallbackPromNo.includes('제') && /^\d+/.test(fallbackPromNo)) {
    fallbackPromNo = `제${fallbackPromNo.replace(/[^0-9-]/g, '')}호`;
  }
  const fallbackPromDate = promDate ? formatStandardKoreanDate(promDate) : '';

  // 2. Extract Related Law from parenthesized strings near header
  // Examples:
  // - 부      칙 <법률 제7849호, 2006. 2. 21.> (제주특별자치도 설치 및 국제자유도시 조성을 위한 특별법)
  // - 부      칙 (제주특별자치도 설치 및 국제자유도시 조성을 위한 특별법)
  // - 부칙 (기술신용보증기금법)
  const headerParenMatch = cleanContent.slice(0, 300).match(/\(([^)]+)\)/);
  if (headerParenMatch) {
    const pText = headerParenMatch[1].trim();
    // Verify it is a related law name (ends with 법, 령, 규칙, 특례법, 규정 or does not look like pure date / clause title like 시행일)
    const isLawName =
      (pText.endsWith('법') ||
        pText.endsWith('령') ||
        pText.endsWith('규칙') ||
        pText.endsWith('특례법') ||
        pText.endsWith('규정') ||
        pText.endsWith('조례') ||
        pText.endsWith('약관') ||
        pText.includes('법률') ||
        pText.includes('에 관한') ||
        pText.includes('에 따른')) &&
      !pText.startsWith('시행일') &&
      !pText.startsWith('유효기간') &&
      !pText.startsWith('적용례') &&
      !pText.startsWith('경과조치') &&
      !pText.startsWith('다른 법');

    if (isLawName) {
      defaultRelatedLaw = `(${pText})`;
    }
  }

  // 3. Extract buchikCategory from <...> near header if available
  const angleBracketMatch = cleanContent.slice(0, 250).match(/<([^>]+)>/);
  if (angleBracketMatch) {
    const inner = angleBracketMatch[1].trim();
    // Check if inner contains actual promulgation info (호, date, or statutory classification words)
    const hasPromInfo =
      /\d{4}/.test(inner) ||
      inner.includes('호') ||
      inner.includes('제') ||
      inner.includes('법률') ||
      inner.includes('대통령령') ||
      inner.includes('부령') ||
      inner.includes('고시') ||
      inner.includes('훈령');

    // Make sure inner is NOT a law name (e.g. <제주특별자치도...법>)
    const isJustLawName =
      (inner.endsWith('법') || inner.endsWith('령') || inner.endsWith('규칙') || inner.endsWith('특례법')) &&
      !inner.includes('호') &&
      !/\d{4}/.test(inner);

    if (hasPromInfo && !isJustLawName) {
      let catLawType = fallbackLawType;
      if (inner.includes('법률')) catLawType = '법률';
      else if (inner.includes('대통령령')) catLawType = '대통령령';
      else if (inner.includes('기획재정부령')) catLawType = '기획재정부령';
      else if (inner.includes('총리령')) catLawType = '총리령';
      else if (inner.includes('부령')) catLawType = '부령';
      else if (inner.includes('고시')) catLawType = '관세청고시';
      else if (inner.includes('훈령')) catLawType = '관세청훈령';
      else if (inner.includes('예규')) catLawType = '관세청예규';

      // Extract promNo inside bracket if present
      const noMatch = inner.match(/(?:제\s*)?(\d+(?:-\d+)*)\s*호/);
      const catPromNo = noMatch ? `제${noMatch[1]}호` : fallbackPromNo;

      // Extract date inside bracket if present
      const dateMatch = inner.match(/(\d{4})\s*\.\s*(\d{1,2})\s*\.\s*(\d{1,2})/);
      let catPromDate = '';
      if (dateMatch) {
        catPromDate = `${dateMatch[1]}. ${parseInt(dateMatch[2], 10)}. ${parseInt(dateMatch[3], 10)}.`;
      } else {
        catPromDate = fallbackPromDate;
      }

      if (catPromNo && catPromDate) {
        buchikCategory = `부칙 <${catLawType} ${catPromNo}, ${catPromDate}>`;
      } else if (catPromNo) {
        buchikCategory = `부칙 <${catLawType} ${catPromNo}>`;
      } else if (catPromDate) {
        buchikCategory = `부칙 <${catLawType}, ${catPromDate}>`;
      } else {
        buchikCategory = `부칙 <${catLawType}>`;
      }
    } else if (isJustLawName && !defaultRelatedLaw) {
      defaultRelatedLaw = `(${inner})`;
    }
  }

  // 4. If buchikCategory was not extracted from <...>, construct from parameters
  if (!buchikCategory) {
    if (fallbackPromNo && fallbackPromDate) {
      buchikCategory = `부칙 <${fallbackLawType} ${fallbackPromNo}, ${fallbackPromDate}>`;
    } else if (fallbackPromNo) {
      buchikCategory = `부칙 <${fallbackLawType} ${fallbackPromNo}>`;
    } else if (fallbackPromDate) {
      buchikCategory = `부칙 <${fallbackLawType}, ${fallbackPromDate}>`;
    } else {
      buchikCategory = `부칙 <${fallbackLawType}>`;
    }
  }

  // 5. Calculate remaining article text by stripping out header lines
  let remaining = cleanContent;
  const headerStripRegex = /^부\s*칙(?:\s*<[^>]+>)?(?:\s*\([^)]+\))?(?:\s*\[[^\]]+\])?/;
  remaining = remaining.replace(headerStripRegex, '').trim();

  // If there is still a related law line at the top like "(제주특별자치도...)"
  if (defaultRelatedLaw) {
    const rawRel = defaultRelatedLaw.slice(1, -1).trim();
    const relLineRegex = new RegExp(`^\\(\\s*${rawRel.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s*\\)\\s*`, 'i');
    remaining = remaining.replace(relLineRegex, '').trim();
  }

  const items: BuchikArticle[] = [];

  // Match top-level articles in buchik:
  // Top-level buchik articles strictly have the form "제X조(제목)" or "제X조 (제목)"
  // e.g. 제1조(시행일), 제2조 (다른 법률의 개정), 제3조(유효기간)
  const topArticleWithTitleRegex = /(?:^|\n)\s*제\s*(\d+(?:-\d+)*(?:의\d+)?)\s*조(?:의(\d+))?\s*\(([^)]+)\)/g;
  let artMatches = [...remaining.matchAll(topArticleWithTitleRegex)];

  // Fallback: If no articles with parenthesized titles were found (rare/historical), match sequential articles
  if (artMatches.length === 0) {
    const rawMatches = [...remaining.matchAll(/(?:^|\n)\s*제\s*(\d+(?:-\d+)*(?:의\d+)?)\s*조(?:의(\d+))?(?:\s*\(([^)]+)\))?/g)];
    let expectedSeq = 1;
    const filtered: any[] = [];
    for (const rm of rawMatches) {
      const n = parseInt(rm[1], 10);
      if (!isNaN(n) && (n === expectedSeq || n === expectedSeq + 1 || filtered.length === 0)) {
        filtered.push(rm);
        expectedSeq = n;
      }
    }
    artMatches = filtered;
  }

  if (artMatches.length > 0) {
    for (let i = 0; i < artMatches.length; i++) {
      const match = artMatches[i];
      const rawNo = match[1].replace(/\s+/g, '');
      const uiNo = match[2] || '';
      let title = (match[3] || '').trim();
      if (title && !title.startsWith('(')) {
        title = `(${title})`;
      }
      const startIdx = match.index || 0;
      const nextStartIdx = i + 1 < artMatches.length ? artMatches[i + 1].index || remaining.length : remaining.length;

      const rawArticleSegment = remaining.slice(startIdx, nextStartIdx).trim();

      let fullNo = `제${rawNo}조`;
      if (uiNo) {
        fullNo += `의${uiNo}`;
      }

      const isOtherLaws =
        title.includes('다른 법률') ||
        title.includes('다른 법령') ||
        title.includes('다른법률') ||
        title.includes('다른법령') ||
        title.includes('타법') ||
        title.includes('타 법률') ||
        title.includes('타 법령');

      // Strip header line (e.g. "제7조 (다른 법률의 개정)") to get paragraph content
      const contentWithoutHeader = rawArticleSegment
        .replace(/^제\s*\d+(?:-\d+)*(?:의\d+)?\s*조(?:의\d+)?(?:\s*\([^)]+\))?/, '')
        .trim();

      // Split paragraphs using our specialized rule (all internal target law articles stay inside this article)
      const parsedParagraphs = splitBuchikParagraphs(contentWithoutHeader, isOtherLaws);

      if (parsedParagraphs.length === 0 || (parsedParagraphs.length === 1 && !isOtherLaws)) {
        // Single paragraph / simple structure
        let rowRelatedLaw = defaultRelatedLaw;
        const targetLaw = extractTargetLawName(rawArticleSegment);
        if (targetLaw && (isOtherLaws || title.includes('개정') || !rowRelatedLaw)) {
          rowRelatedLaw = `(${targetLaw})`;
        }

        items.push({
          buchikCategory,
          relatedLaw: rowRelatedLaw || '',
          articleNo: fullNo,
          articleTitle: title || '',
          articleContent: formatLawArticleText(rawArticleSegment),
        });
      } else {
        // Multi-paragraph structure (①, ②, ③...): one row per paragraph / target law
        for (let pIdx = 0; pIdx < parsedParagraphs.length; pIdx++) {
          const p = parsedParagraphs[pIdx];
          let pRelatedLaw = defaultRelatedLaw;
          if (p.targetLaw) {
            pRelatedLaw = p.targetLaw === '생략' ? '(생략)' : `(${p.targetLaw})`;
          }

          // In other-laws articles, keep the full original text of that paragraph with amendment details
          const finalContent = p.hangText;

          items.push({
            buchikCategory,
            relatedLaw: pRelatedLaw || '',
            articleNo: fullNo,
            articleTitle: title || '',
            articleContent: formatLawArticleText(finalContent),
          });
        }
      }
    }
  } else {
    // Exception structure: No explicit "제X조", starts directly with ①(시행일), ②(다른 법률의 개정) 등
    const parsedParagraphs = splitBuchikParagraphs(remaining, false);
    if (parsedParagraphs.length > 1) {
      for (const p of parsedParagraphs) {
        let pRelatedLaw = defaultRelatedLaw;
        if (p.targetLaw) {
          pRelatedLaw = p.targetLaw === '생략' ? '(생략)' : `(${p.targetLaw})`;
        }

        // Try extracting sub-title if formatted as ①(시행일) or ① (다른 법률의 개정)
        let subTitle = '-';
        const titleMatch = p.hangText.match(/^[①-㊿]\s*(\([^\)]+\))/);
        if (titleMatch) {
          subTitle = titleMatch[1];
        }

        items.push({
          buchikCategory,
          relatedLaw: pRelatedLaw || '',
          articleNo: '-',
          articleTitle: subTitle,
          articleContent: formatLawArticleText(p.hangText),
        });
      }
    } else {
      items.push({
        buchikCategory,
        relatedLaw: defaultRelatedLaw || '',
        articleNo: '-',
        articleTitle: '-',
        articleContent: formatLawArticleText(remaining || cleanContent),
      });
    }
  }

  return items;
}

function parseLawBuchikArticles(rawRoot: any, lawType: string = '법률'): BuchikArticle[] {
  const root = normalizeLawXmlRoot(rawRoot);
  const buchik = root.부칙 || root.법령?.부칙 || root.Law?.Buchik;
  let units = buchik?.부칙단위 || buchik?.BuchikUnit || [];
  if (!Array.isArray(units)) units = units ? [units] : [];

  if (units.length === 0) {
    const rawText = getText(root.부칙내용 || root.부칙 || '');
    if (rawText) {
      return parseSingleBuchikText(rawText, '', '', lawType);
    }
    return [];
  }

  // The latest statute revision has its buchik as the last unit
  const targetUnit = units[units.length - 1];
  const promDate = String(targetUnit.부칙공포일자 || '').replace(/\D/g, '');
  const promNo = String(targetUnit.부칙공포번호 || '').trim();
  const rawContent = getText(targetUnit.부칙내용 || targetUnit['#text'] || '');

  return parseSingleBuchikText(rawContent, promNo, promDate, lawType);
}

function parseAllLawBuchikArticles(rawRoot: any, lawType: string = '법률'): BuchikArticle[] {
  const root = normalizeLawXmlRoot(rawRoot);
  const buchik = root.부칙 || root.법령?.부칙 || root.Law?.Buchik;
  let units = buchik?.부칙단위 || buchik?.BuchikUnit || [];
  if (!Array.isArray(units)) units = units ? [units] : [];

  if (units.length === 0) {
    const rawText = getText(root.부칙내용 || root.부칙 || '');
    if (rawText) {
      return parseSingleBuchikText(rawText, '', '', lawType);
    }
    return [];
  }

  const allArticles: BuchikArticle[] = [];
  for (const unit of units) {
    const promDate = String(unit.부칙공포일자 || '').replace(/\D/g, '');
    const promNo = String(unit.부칙공포번호 || '').trim();
    const rawContent = getText(unit.부칙내용 || unit['#text'] || '');
    const parsed = parseSingleBuchikText(rawContent, promNo, promDate, lawType);
    allArticles.push(...parsed);
  }
  return allArticles;
}

// Format date string to Korean standard dot format: 'YYYY. M. D.' (e.g. '2026. 8. 11.')
export function formatStandardKoreanDate(dateStr?: string): string {
  if (!dateStr) return '';
  const cleanStr = String(dateStr).trim();
  if (
    cleanStr.includes('9999') ||
    cleanStr.includes('미정') ||
    cleanStr === '-' ||
    cleanStr.toLowerCase().includes('unknown')
  ) {
    return '시행미정';
  }
  const digits = cleanStr.replace(/\D/g, '');
  if (digits.length >= 8) {
    const y = digits.substring(0, 4);
    if (y === '9999') return '시행미정';
    const m = parseInt(digits.substring(4, 6), 10);
    const d = parseInt(digits.substring(6, 8), 10);
    if (!isNaN(m) && !isNaN(d) && m >= 1 && m <= 12 && d >= 1 && d <= 31) {
      return `${y}. ${m}. ${d}.`;
    }
  }
  // If already in 'YYYY. M. D.' format
  const dotMatch = cleanStr.match(/^(\d{4})\s*\.\s*(\d{1,2})\s*\.\s*(\d{1,2})/);
  if (dotMatch) {
    const y = dotMatch[1];
    if (y === '9999') return '시행미정';
    const m = parseInt(dotMatch[2], 10);
    const d = parseInt(dotMatch[3], 10);
    return `${y}. ${m}. ${d}.`;
  }
  return cleanStr;
}

// Standardize law/decree/rule name spacing according to modern Korean legislative naming standards:
// Ensures a space before '시행령' and '시행규칙' when attached without space (e.g. '관세법시행령' -> '관세법 시행령', '관세법시행규칙' -> '관세법 시행규칙')
// Other rules remain untouched as requested.
export function normalizeLawTitleSpacing(name: string): string {
  if (!name) return '';
  let clean = String(name).trim();
  // Ensure space before '시행령' if attached directly to preceding word/character
  clean = clean.replace(/([가-힣a-zA-Z0-9])시행령/g, '$1 시행령');
  // Ensure space before '시행규칙' if attached directly to preceding word/character
  clean = clean.replace(/([가-힣a-zA-Z0-9])시행규칙/g, '$1 시행규칙');
  // Clean up any double spaces
  clean = clean.replace(/\s{2,}/g, ' ').trim();
  return clean;
}

// Generate standardized filename / doc title without leading index number
// - Law: "관세법 [시행 2026. 8. 11.] [법률 제21858호, 2026. 8. 11., 일부개정]"
// - Decree: "관세법 시행령 [시행 2026. 8. 11.] [대통령령 제34858호, 2026. 8. 11., 일부개정]"
// - Rule: "관세법 시행규칙 [시행 2026. 8. 11.] [기획재정부령 제123호, 2026. 8. 11., 일부개정]"
// - Admrul: "수입통관 사무처리에 관한 고시 [시행 2025. 12. 17.] [관세청고시 제2025-66호, 2025. 12. 17., 일부개정]"
export function generateStandardRevisionTitle(rev: any, defaultDocName: string = '관세법', defaultTargetType: string = 'law'): string {
  const rawDocName = (rev.name || defaultDocName).trim();
  const docName = normalizeLawTitleSpacing(rawDocName);
  const rawEnfDate = rev.enforcementDate || '';
  const rawPromDate = rev.promulgationDate || '';
  const enfDateFormatted = formatStandardKoreanDate(rawEnfDate);
  const promDateFormatted = formatStandardKoreanDate(rawPromDate);
  const revType = (rev.revisionType || rev.revType || '일부개정').trim();

  const isAdmrul =
    rev.targetType === 'admrul' ||
    defaultTargetType === 'admrul' ||
    docName.includes('고시') ||
    docName.includes('훈령') ||
    docName.includes('예규') ||
    docName.includes('규정');

  // Determine law/rule prefix
  let ruleType = (rev.ruleType || rev.lawType || '').trim();
  const subType = (rev.subType || '').toLowerCase();
  const dept = (rev.department || '관세청').trim();

  if (!ruleType) {
    if (isAdmrul) {
      const cat = docName.includes('훈령') ? '훈령' : docName.includes('예규') ? '예규' : '고시';
      ruleType = `${dept}${cat}`;
    } else if (subType === 'decree' || docName.includes('시행령')) {
      ruleType = '대통령령';
    } else if (subType === 'rule' || docName.includes('시행규칙')) {
      ruleType = dept && !dept.includes('정부') && !dept.includes('부') ? `${dept}령` : dept ? `${dept}령` : '부령';
    } else {
      ruleType = '법률';
    }
  }

  // Format Promulgation No (e.g. '제21858호' or '관세청고시 제2025-66호')
  let rawPromNo = (rev.promulgationNo || '').trim();
  let cleanPromNo = rawPromNo;
  if (!cleanPromNo) {
    cleanPromNo = '제-호';
  } else {
    // If it's just pure numbers or doesn't have '제' and '호'
    if (/^\d+(?:-\d+)*$/.test(cleanPromNo)) {
      cleanPromNo = `제${cleanPromNo}호`;
    } else if (!cleanPromNo.includes('제') && /^\d+/.test(cleanPromNo)) {
      cleanPromNo = `제${cleanPromNo}`;
    }
  }

  let promHeader = '';
  if (isAdmrul) {
    if (cleanPromNo.includes('고시') || cleanPromNo.includes('훈령') || cleanPromNo.includes('예규') || cleanPromNo.includes('공고')) {
      promHeader = cleanPromNo;
    } else {
      promHeader = `${ruleType} ${cleanPromNo}`;
    }
  } else {
    if (cleanPromNo.includes('법률') || cleanPromNo.includes('대통령령') || cleanPromNo.includes('령')) {
      promHeader = cleanPromNo;
    } else {
      promHeader = `${ruleType} ${cleanPromNo}`;
    }
  }

  const enfPart =
    !enfDateFormatted || enfDateFormatted === '시행미정' || enfDateFormatted.includes('9999') || enfDateFormatted === '시행일 미상'
      ? '[시행미정]'
      : `[시행 ${enfDateFormatted}]`;
  const promPart = promDateFormatted
    ? `[${promHeader}, ${promDateFormatted}, ${revType}]`
    : `[${promHeader}, ${revType}]`;

  return `${docName} ${enfPart} ${promPart}`.replace(/[\/\\:*?"<>|]/g, '_');
}

async function fetchAllBuchikArticlesForLaw(
  ocKey: string,
  cleanLawName: string,
  targetType: string = 'law',
  latestRev?: any,
  subTypeFilter?: 'law' | 'decree' | 'rule'
): Promise<BuchikArticle[]> {
  try {
    const isAdmrul =
      targetType === 'admrul' ||
      cleanLawName.includes('규정') ||
      cleanLawName.includes('고시') ||
      cleanLawName.includes('훈령') ||
      cleanLawName.includes('예규');

    let targetDocName = cleanLawName.trim();
    if (subTypeFilter === 'decree' && !targetDocName.includes('시행령')) {
      targetDocName = `${targetDocName} 시행령`;
    } else if (subTypeFilter === 'rule' && !targetDocName.includes('시행규칙')) {
      targetDocName = `${targetDocName} 시행규칙`;
    }

    const lawType = isAdmrul
      ? '고시'
      : targetDocName.includes('시행령')
      ? '대통령령'
      : targetDocName.includes('시행규칙')
      ? '부령'
      : '법률';

    // 1. Direct query via lawService.do (fetches full XML of the latest law/admrul document)
    const directQueryUrl = isAdmrul
      ? `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(ocKey)}&target=admrul&query=${encodeURIComponent(targetDocName)}&type=XML`
      : `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(ocKey)}&target=law&query=${encodeURIComponent(targetDocName)}&type=XML`;

    try {
      const qRes = await fetch(directQueryUrl, {
        headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
        signal: AbortSignal.timeout(8000),
      });
      if (qRes.ok) {
        const qXml = await qRes.text();
        const parsed = xmlParser.parse(qXml);
        const articles = parseAllLawBuchikArticles(parsed, lawType);
        if (articles.length > 0) {
          console.log(`[Buchik Parser] Extracted ${articles.length} buchik rows via direct query for '${targetDocName}'`);
          return articles;
        }
      }
    } catch (qErr: any) {
      console.warn(`[Buchik Parser Direct Query Warning]:`, qErr?.message);
    }

    // 2. Try using latest revision MST if available and matches the requested doc type
    const mst = latestRev?.lawMst || latestRev?.seq || latestRev?.id || '';
    const revName = (latestRev?.name || '').trim();
    const matchesTargetDoc = !subTypeFilter || 
      (subTypeFilter === 'law' && !revName.includes('시행령') && !revName.includes('시행규칙')) ||
      (subTypeFilter === 'decree' && (revName.includes('시행령') || latestRev?.ruleType?.includes('대통령령'))) ||
      (subTypeFilter === 'rule' && (revName.includes('시행규칙') || latestRev?.ruleType?.includes('부령')));

    if (mst && matchesTargetDoc) {
      const detailUrl = isAdmrul
        ? `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(ocKey)}&target=admrul&MST=${encodeURIComponent(mst)}&type=XML`
        : `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(ocKey)}&target=law&MST=${encodeURIComponent(mst)}&type=XML`;

      try {
        const res = await fetch(detailUrl, {
          headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
          signal: AbortSignal.timeout(8000),
        });
        if (res.ok) {
          const xml = await res.text();
          const parsed = xmlParser.parse(xml);
          const articles = parseAllLawBuchikArticles(parsed, lawType);
          if (articles.length > 0) {
            console.log(`[Buchik Parser] Extracted ${articles.length} buchik rows from MST (${mst}) for '${targetDocName}'`);
            return articles;
          }
        }
      } catch (err: any) {
        console.warn(`[Buchik Parser MST Warning]:`, err?.message);
      }
    }

    // 3. Query DRF lawSearch for the latest law document matching targetDocName
    const searchUrl = isAdmrul
      ? `http://www.law.go.kr/DRF/lawSearch.do?OC=${encodeURIComponent(ocKey)}&target=admrul&query=${encodeURIComponent(targetDocName)}&type=XML`
      : `http://www.law.go.kr/DRF/lawSearch.do?OC=${encodeURIComponent(ocKey)}&target=law&query=${encodeURIComponent(targetDocName)}&type=XML`;

    try {
      const sRes = await fetch(searchUrl, {
        headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
        signal: AbortSignal.timeout(6000),
      });
      if (sRes.ok) {
        const sXml = await sRes.text();
        const sParsed = xmlParser.parse(sXml);
        const lawList = sParsed.LawSearch?.law || sParsed.admrulSearch?.admrul || [];
        const laws = Array.isArray(lawList) ? lawList : [lawList];
        
        // Exact name match first (e.g. '관세법 시행령' must match '관세법 시행령', NOT '관세법')
        const matched =
          laws.find((l: any) => getText(l.법령명한글 || l.행정규칙명 || l.법령명).trim() === targetDocName) ||
          laws.find((l: any) => getText(l.법령명한글 || l.행정규칙명 || l.법령명).trim().startsWith(targetDocName)) ||
          laws.find((l: any) => {
            const n = getText(l.법령명한글 || l.행정규칙명 || l.법령명).trim();
            if (subTypeFilter === 'decree') return n.includes('시행령');
            if (subTypeFilter === 'rule') return n.includes('시행규칙');
            if (subTypeFilter === 'law') return !n.includes('시행령') && !n.includes('시행규칙');
            return true;
          }) ||
          (subTypeFilter ? null : laws[0]);

        const foundMst = matched?.법령일련번호 || matched?.행정규칙일련번호 || matched?.['@_MST'] || matched?.['@_seq'];
        if (foundMst) {
          const dUrl = isAdmrul
            ? `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(ocKey)}&target=admrul&MST=${encodeURIComponent(foundMst)}&type=XML`
            : `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(ocKey)}&target=law&MST=${encodeURIComponent(foundMst)}&type=XML`;
          const dRes = await fetch(dUrl, {
            headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
            signal: AbortSignal.timeout(8000),
          });
          if (dRes.ok) {
            const dXml = await dRes.text();
            const parsed = xmlParser.parse(dXml);
            const articles = parseAllLawBuchikArticles(parsed, lawType);
            if (articles.length > 0) {
              console.log(`[Buchik Parser] Extracted ${articles.length} buchik rows from DRF search for '${targetDocName}'`);
              return articles;
            }
          }
        }
      }
    } catch (sErr: any) {
      console.warn(`[Buchik Parser Search Warning]:`, sErr?.message);
    }

    // 4. Fallback: parse latest revision's buchikText if provided and matching
    if (latestRev?.buchikText && matchesTargetDoc) {
      const fallbackArticles = parseSingleBuchikText(
        latestRev.buchikText,
        latestRev.promulgationNo || '',
        latestRev.promulgationDate || '',
        lawType
      );
      if (fallbackArticles.length > 0) return fallbackArticles;
    }
  } catch (err: any) {
    console.error(`[fetchAllBuchikArticlesForLaw Error] '${cleanLawName}':`, err?.message || err);
  }
  return [];
}

// High-resilience HTML article parser fallback (when XML does not provide structured article units)
function parseArticlesFromHtmlText(html: string): any[] {
  if (!html) return [];
  const clean = cleanLawHtml(html);
  const lines = clean.split('\n').map((l) => l.trim()).filter(Boolean);
  const rawText = lines.join('\n');

  const articles: any[] = [];
  const artMatches = [...rawText.matchAll(/(?:^|\n)\s*(제\s*(\d+(?:-\d+)*(?:의\d+)?)\s*조(?:의\d+)?\s*(?:\(([^)]+)\))?)/g)];

  if (artMatches.length === 0) {
    return [];
  }

  let currentChapter = '';
  let currentSection = '';
  let currentSubsection = '';

  for (let i = 0; i < artMatches.length; i++) {
    const match = artMatches[i];
    const rawNo = match[2].replace(/\s+/g, '');
    const title = (match[3] || '').trim();
    const artStart = match.index || 0;
    const nextArtStart = i + 1 < artMatches.length ? artMatches[i + 1].index || rawText.length : rawText.length;

    const prevSlice = i === 0 ? rawText.slice(0, artStart) : rawText.slice(artMatches[i - 1].index || 0, artStart);

    // 1. Check for Buchik Header in previous slice
    const buchikMatch = prevSlice.match(/(?:^|\n|\s)(부\s*칙\s*(?:<[^>]+>|\([^\)]+\)|\[[^\]]+\])?)/);
    if (buchikMatch) {
      let bHeader = buchikMatch[1].replace(/부\s*칙/, '부칙').replace(/\s+/g, ' ').trim();
      bHeader = bHeader.replace(/(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})\.?/, '$1. $2. $3.');
      currentChapter = bHeader;
      currentSection = '';
      currentSubsection = '';
    } else {
      // 2. Check for Chapter Header
      const chMatch = prevSlice.match(
        /(?:^|\n)\s*(제\s*(\d+|[일이삼사오육칠팔구십백]+)\s*(?:장|편)(?:\s+([^\n]+?))?)(?=\s+제\s*\d+\s*(?:절|관|조|-)|부\s*칙|\n|$)/
      );
      if (chMatch) {
        const chNum = chMatch[2];
        const chTitle = (chMatch[3] || '').trim();
        if (!chTitle.includes('내지') && !chTitle.includes('부터') && !chTitle.includes('준용')) {
          currentChapter = cleanHeaderTitle(`제${chNum}장 ${chTitle}`.trim());
          currentSection = '';
          currentSubsection = '';
        }
      }
    }

    const secMatch = prevSlice.match(
      /(?:^|\n)\s*(제\s*(\d+|[일이삼사오육칠팔구십백]+)\s*절(?:\s+([^\n]+?))?)(?=\s+제\s*\d+\s*(?:관|조|-)|부\s*칙|\n|$)/
    );
    if (secMatch) {
      const sNum = secMatch[2];
      const sTitle = (secMatch[3] || '').trim();
      if (!sTitle.includes('내지') && !sTitle.includes('부터')) {
        currentSection = cleanHeaderTitle(`제${sNum}절 ${sTitle}`.trim());
        currentSubsection = '';
      }
    }

    const subMatch = prevSlice.match(
      /(?:^|\n)\s*(제\s*(\d+|[일이삼사오육칠팔구십백]+)\s*관(?:\s+([^\n]+?))?)(?=\s+제\s*\d+\s*(?:조|-)|부\s*칙|\n|$)/
    );
    if (subMatch) {
      const subNum = subMatch[2];
      const subTitle = (subMatch[3] || '').trim();
      if (!subTitle.includes('내지') && !subTitle.includes('부터')) {
        currentSubsection = cleanHeaderTitle(`제${subNum}관 ${subTitle}`.trim());
      }
    }

    let artContent = rawText.slice(artStart, nextArtStart).trim();
    artContent = stripTrailingStructuralHeaders(artContent);

    let fullArtNo = `제${rawNo}조`;
    if (match[1].includes('조의')) {
      const uiMatch = match[1].match(/조의\s*(\d+)/);
      if (uiMatch) fullArtNo += `의${uiMatch[1]}`;
    }

    const isDel = artContent.includes('삭제') || title.includes('삭제');
    const histMeta = extractArticleHistoryMetadata(null, artContent, isDel);

    articles.push({
      chapterName: currentChapter,
      sectionName: currentSection,
      subsectionName: currentSubsection,
      articleNo: fullArtNo,
      articleTitle: title,
      articleContent: formatLawArticleText(artContent),
      effectiveDate: '',
      isDeleted: isDel,
      fullRevision: histMeta.fullRevision,
      creation: histMeta.creation,
      titleRevision: histMeta.titleRevision,
      remarks: histMeta.remarks,
    });
  }

  // Extract buchik if present at the end of HTML
  const buchikMatch = rawText.match(/(?:^|\n)\s*(부\s*칙(?:\s*<[^>]+>|\s*\([^\)]+\)|\[[^\]]+\])?[\s\S]*)$/);
  if (buchikMatch) {
    const buchikArticles = parseSingleBuchikText(buchikMatch[1]);
    if (buchikArticles.length > 0) {
      (articles as any).buchikArticles = buchikArticles;
    }
  }

  return articles;
}

// Helper function to extract structured articles from parsed XML root (Laws - 법률, 대통령령, 부령)
function parseArticlesFromXmlRoot(rawRoot: any): any[] {
  const root = normalizeLawXmlRoot(rawRoot);
  let rawArticles =
    root.조문?.조문단위 ||
    root.조문단위 ||
    root.조문 ||
    root.법령?.조문?.조문단위 ||
    root.법령?.조문단위 ||
    root.Law?.Articles?.Article ||
    [];

  if (!Array.isArray(rawArticles)) {
    rawArticles = rawArticles ? [rawArticles] : [];
  }

  let currentChapter = '';
  let currentSection = '';
  let currentSubsection = '';

  const articles: any[] = [];

  for (const item of rawArticles) {
    const type = getText(item.조문여부 || item.구분);
    const content = getText(item.조문내용 || item.조문본문 || item.내용 || item['#text']);
    const title = getText(item.조제목 || item.제목);
    const noStr = getText(item.조문번호 || item.번호);

    const chapterMatch = parseChapterHeader(type, content, title);
    if (chapterMatch) {
      currentChapter = chapterMatch;
      currentSection = '';
      currentSubsection = '';
      continue;
    }

    const sectionMatch = parseSectionHeader(type, content, title);
    if (sectionMatch) {
      currentSection = sectionMatch;
      currentSubsection = '';
      continue;
    }

    const subsectionMatch = parseSubsectionHeader(type, content, title);
    if (subsectionMatch) {
      currentSubsection = subsectionMatch;
      continue;
    }

    if (type === '조문' || noStr || (content && content.startsWith('제')) || title || (content && content.length > 0)) {
      const fullContent = extractArticleContent(item) || content;

      let formattedNo = '';
      if (noStr) {
        let cleanNo = String(noStr).trim();
        if (cleanNo.startsWith('제')) {
          formattedNo = cleanNo;
        } else if (cleanNo.includes('조')) {
          formattedNo = `제${cleanNo}`;
        } else if (cleanNo.includes('의')) {
          const parts = cleanNo.split('의');
          formattedNo = `제${parts[0]}조의${parts.slice(1).join('의')}`;
        } else {
          formattedNo = `제${cleanNo}조`;
        }
      }

      // Check if content or fullContent begins with full branch article pattern (e.g. 제26조의2, 제26조의2의3, 제2-1조의2)
      const targetTxt = fullContent || content || '';
      const textMatch = targetTxt.match(/^제\s*(\d+(?:-\d+)*(?:의\d+)*)\s*조(?:\s*의\s*(\d+(?:의\d+)*))?/);
      if (textMatch) {
        let extractedNo = `제${textMatch[1]}조`;
        if (textMatch[2]) {
          extractedNo += `의${textMatch[2].replace(/\s+/g, '')}`;
        } else if (textMatch[0].includes('의')) {
          const uiPart = textMatch[0].match(/조\s*의\s*(\d+(?:의\d+)*)/);
          if (uiPart) extractedNo += `의${uiPart[1].replace(/\s+/g, '')}`;
        }
        if (!formattedNo || formattedNo.length < extractedNo.length || !formattedNo.includes('의') && extractedNo.includes('의')) {
          formattedNo = extractedNo;
        }
      }

      let cleanTitle = (title || '').trim().replace(/^\(|\)$/g, '');
      if (!cleanTitle && targetTxt) {
        const match = targetTxt.match(/^제\s*\d+(?:-\d+)*(?:조(?:의\d+)*)?\s*\(([^)]+)\)/);
        if (match && match[1]) {
          cleanTitle = match[1].trim();
        } else {
          const match2 = targetTxt.match(/\(([^)]+)\)/);
          if (match2 && match2[1] && match2.index !== undefined && match2.index < 40) {
            cleanTitle = match2[1].trim();
          }
        }
      }

      const isDel = (fullContent || content).includes('삭제') || title.includes('삭제') || cleanTitle.includes('삭제');
      const histMeta = extractArticleHistoryMetadata(item, fullContent || content, isDel);

      articles.push({
        chapterName: currentChapter || '',
        sectionName: currentSection || '',
        subsectionName: currentSubsection || '',
        articleNo: formattedNo || `조문 ${articles.length + 1}`,
        articleTitle: cleanTitle || '',
        articleContent: formatLawArticleText(fullContent || content),
        effectiveDate: formatDate(getText(item.조문시행일자 || item.시행일자)),
        isDeleted: isDel,
        fullRevision: histMeta.fullRevision,
        creation: histMeta.creation,
        titleRevision: histMeta.titleRevision,
        remarks: histMeta.remarks,
      });
    }
  }

  // Fallback to text parsing if structured units returned 0
  if (articles.length === 0) {
    const rawBody = getText(root.본문 || root.본문내용 || root.전문 || root.조문내용 || '');
    if (rawBody) {
      const parsedHtml = parseArticlesFromHtmlText(rawBody);
      if (parsedHtml.length > 0) return parsedHtml;
    }
  }

  const buchikArticles = parseLawBuchikArticles(rawRoot, '법률');
  (articles as any).buchikArticles = buchikArticles;

  return articles;
}

// Helper function to extract structured articles from Administrative Rules (행정규칙) XML root
function parseAdmrulArticlesFromXmlRoot(rawRoot: any): any[] {
  const root = normalizeAdmrulXmlRoot(rawRoot);
  const articles: any[] = [];
  
  // 1. Standard structured articles (조문/조문단위/행정규칙조문)
  let rawArticles =
    root.조문?.조문단위 ||
    root.조문단위 ||
    root.행정규칙조문 ||
    root.조문 ||
    root.행정규칙?.조문?.조문단위 ||
    root.행정규칙?.조문단위 ||
    root.행정규칙?.행정규칙조문 ||
    [];

  if (!Array.isArray(rawArticles)) {
    rawArticles = rawArticles ? [rawArticles] : [];
  }

  if (rawArticles.length > 0) {
    let currentChapter = '';
    let currentSection = '';
    let currentSubsection = '';

    for (const item of rawArticles) {
      const type = getText(item.조문여부 || item.구분);
      const content = getText(item.조문내용 || item.조문본문 || item.내용 || item['#text']);
      const title = getText(item.조제목 || item.제목);
      const noStr = getText(item.조문번호 || item.번호);

      const chapterMatch = parseChapterHeader(type, content, title);
      if (chapterMatch) {
        currentChapter = chapterMatch;
        currentSection = '';
        currentSubsection = '';
        continue;
      }
      const sectionMatch = parseSectionHeader(type, content, title);
      if (sectionMatch) {
        currentSection = sectionMatch;
        currentSubsection = '';
        continue;
      }
      const subsectionMatch = parseSubsectionHeader(type, content, title);
      if (subsectionMatch) {
        currentSubsection = subsectionMatch;
        continue;
      }

      if (type === '조문' || noStr || (content && content.startsWith('제')) || title || (content && content.length > 0)) {
        const fullContent = extractArticleContent(item) || content;

        let formattedNo = '';
        if (noStr) {
          let cleanNo = String(noStr).trim();
          if (cleanNo.startsWith('제')) {
            formattedNo = cleanNo;
          } else if (cleanNo.includes('조')) {
            formattedNo = `제${cleanNo}`;
          } else if (cleanNo.includes('의')) {
            const parts = cleanNo.split('의');
            formattedNo = `제${parts[0]}조의${parts.slice(1).join('의')}`;
          } else {
            formattedNo = `제${cleanNo}조`;
          }
        }

        const targetTxt = fullContent || content || '';
        const textMatch = targetTxt.match(/^제\s*(\d+(?:-\d+)*(?:의\d+)*)\s*조(?:\s*의\s*(\d+(?:의\d+)*))?/);
        if (textMatch) {
          let extractedNo = `제${textMatch[1]}조`;
          if (textMatch[2]) {
            extractedNo += `의${textMatch[2].replace(/\s+/g, '')}`;
          } else if (textMatch[0].includes('의')) {
            const uiPart = textMatch[0].match(/조\s*의\s*(\d+(?:의\d+)*)/);
            if (uiPart) extractedNo += `의${uiPart[1].replace(/\s+/g, '')}`;
          }
          if (!formattedNo || formattedNo.length < extractedNo.length || !formattedNo.includes('의') && extractedNo.includes('의')) {
            formattedNo = extractedNo;
          }
        }

        let cleanTitle = (title || '').trim().replace(/^\(|\)$/g, '');
        if (!cleanTitle && targetTxt) {
          const match = targetTxt.match(/^제\s*\d+(?:-\d+)*(?:조(?:의\d+)*)?\s*\(([^)]+)\)/);
          if (match && match[1]) {
            cleanTitle = match[1].trim();
          }
        }

        // Accurately resolve chapter name: If article is e.g. "제2-1조", Chapter is Chapter 2
        let resolvedChapter = currentChapter || '';
        const hyphenMatch = (formattedNo || '').match(/^제(\d+)-/);
        if (hyphenMatch) {
          const chNum = parseInt(hyphenMatch[1], 10);
          if (foreignExchangeChapters[chNum]) {
            resolvedChapter = foreignExchangeChapters[chNum];
          } else if (currentChapter) {
            resolvedChapter = currentChapter;
          }
        }

        const isDel = (fullContent || content).includes('삭제') || title.includes('삭제') || cleanTitle.includes('삭제');
        const histMeta = extractArticleHistoryMetadata(item, fullContent || content, isDel);

        articles.push({
          chapterName: resolvedChapter,
          sectionName: currentSection || '',
          subsectionName: currentSubsection || '',
          articleNo: formattedNo || `조문 ${articles.length + 1}`,
          articleTitle: cleanTitle || '',
          articleContent: formatLawArticleText(fullContent || content),
          effectiveDate: formatDate(getText(item.조문시행일자 || item.시행일자)),
          isDeleted: isDel,
          fullRevision: histMeta.fullRevision,
          creation: histMeta.creation,
          titleRevision: histMeta.titleRevision,
          remarks: histMeta.remarks,
        });
      }
    }
  }

  // 2. High-precision full-text parsing from 조문내용 / 본문 / 본문내용 / 전문 (e.g. 외국환거래규정 or 고시 본문)
  if (articles.length === 0) {
    const rawText = getText(
      root.조문내용 ||
      root.본문 ||
      root.본문내용 ||
      root.전문 ||
      root.행정규칙?.조문내용 ||
      root.행정규칙?.본문 ||
      root.행정규칙?.본문내용 ||
      root.행정규칙?.전문 ||
      ''
    );

    if (rawText) {
      // Dynamic chapter header extraction
      const dynamicChapterDict: Record<number, string> = {};
      const chDefMatches = [
        ...rawText.matchAll(
          /(?:^|\n|\s)(제\s*(\d+)\s*장(?:\s+([^\n]+?))?)(?=\s+제\s*\d+\s*(?:절|관|조|-)|부\s*칙|\n|$)/g
        ),
      ];
      chDefMatches.forEach((m) => {
        const chNum = parseInt(m[2], 10);
        const chTitle = (m[3] || '').trim();
        if (
          !chTitle.includes('내지') &&
          !chTitle.includes('부터') &&
          !chTitle.includes('의한') &&
          !chTitle.includes('따라') &&
          !chTitle.includes('준용') &&
          !chTitle.includes('관련')
        ) {
          dynamicChapterDict[chNum] = cleanHeaderTitle(`제${chNum}장 ${chTitle}`.trim());
        }
      });

      // Match all article headers: e.g. 제1-1조(목적), 제2-1조의2(지급 및 수령), 제2-6조의2 (예금 및 신탁)
      const artMatches = [...rawText.matchAll(/제\s*(\d+(?:-\d+)*(?:의\d+)?)\s*조(?:의\d+)?\s*\(([^)]+)\)/g)];

      for (let i = 0; i < artMatches.length; i++) {
        const match = artMatches[i];
        const rawNo = match[1].replace(/\s+/g, '');
        const title = match[2].trim();
        const artStart = match.index || 0;
        const nextArtStart = i + 1 < artMatches.length ? artMatches[i + 1].index || rawText.length : rawText.length;

        // Determine Chapter number from article number format if chapters are explicitly present
        let curChapter = '';
        if (rawNo.includes('-')) {
          const parts = rawNo.split('-');
          const chNum = parseInt(parts[0], 10);
          if (dynamicChapterDict[chNum]) {
            curChapter = dynamicChapterDict[chNum];
          } else if (foreignExchangeChapters[chNum]) {
            curChapter = foreignExchangeChapters[chNum];
          }
        }

        const precedingText = i === 0 ? rawText.slice(0, artStart) : rawText.slice(artMatches[i - 1].index || 0, artStart);

        let curSection = '';
        let curSubsection = '';
        const secMatch = precedingText.match(
          /(?:^|\n|\s)(제\s*(\d+)\s*절(?:\s+([^\n]+?))?)(?=\s+제\s*\d+\s*(?:관|조|-)|부\s*칙|\n|$)/
        );
        if (secMatch) {
          const sTitle = (secMatch[3] || '').trim();
          if (!sTitle.includes('내지') && !sTitle.includes('부터') && !sTitle.includes('의한') && !sTitle.includes('따라')) {
            curSection = cleanHeaderTitle(`제${secMatch[2]}절 ${sTitle}`.trim());
          }
        }

        const subMatch = precedingText.match(
          /(?:^|\n|\s)(제\s*(\d+)\s*관(?:\s+([^\n]+?))?)(?=\s+제\s*\d+\s*(?:조|-)|부\s*칙|\n|$)/
        );
        if (subMatch) {
          const subTitle = (subMatch[3] || '').trim();
          if (!subTitle.includes('내지') && !subTitle.includes('부터')) {
            curSubsection = cleanHeaderTitle(`제${subMatch[2]}관 ${subTitle}`.trim());
          }
        }

        let artBlock = rawText.slice(artStart, nextArtStart).trim();
        artBlock = stripTrailingStructuralHeaders(artBlock);

        let fullArtNo = `제${rawNo}조`;
        if (match[0].includes('조의')) {
          const uiMatch = match[0].match(/조의\s*(\d+)/);
          if (uiMatch) fullArtNo += `의${uiMatch[1]}`;
        }

        const isDel = artBlock.includes('삭제') || title.includes('삭제');
        const histMeta = extractArticleHistoryMetadata(null, artBlock, isDel);

        articles.push({
          chapterName: curChapter,
          sectionName: curSection,
          subsectionName: curSubsection,
          articleNo: fullArtNo,
          articleTitle: title,
          articleContent: formatLawArticleText(artBlock),
          effectiveDate: '',
          isDeleted: isDel,
          fullRevision: histMeta.fullRevision,
          creation: histMeta.creation,
          titleRevision: histMeta.titleRevision,
          remarks: histMeta.remarks,
        });
      }
    }
  }

  // 3. Fallback: split by paragraphs or HTML
  if (articles.length === 0) {
    const rawBody = getText(root.본문 || root.본문내용 || root.전문 || root.조문내용 || '');
    if (rawBody) {
      const htmlArticles = parseArticlesFromHtmlText(rawBody);
      if (htmlArticles.length > 0) return htmlArticles;

      const paragraphs = rawBody.split(/\n\s*\n/).filter((p: string) => p.trim());
      paragraphs.forEach((para: string, idx: number) => {
        const trimmed = para.trim();
        let artNo = `항목 ${idx + 1}`;
        let artTitle = '';
        const match = trimmed.match(/^제(\d+(?:-\d+)*)(?:조(?:의\d+)?)?(?:\(([^)]+)\))?/);
        if (match) {
          artNo = match[0].includes('조') ? `제${match[1]}조` : match[0];
          artTitle = match[2] || '';
        }

        const isDel = trimmed.includes('삭제') || artTitle.includes('삭제');
        const histMeta = extractArticleHistoryMetadata(null, trimmed, isDel);

        articles.push({
          chapterName: '',
          sectionName: '',
          subsectionName: '',
          articleNo: artNo,
          articleTitle: artTitle,
          articleContent: formatLawArticleText(trimmed),
          effectiveDate: '',
          isDeleted: isDel,
          fullRevision: histMeta.fullRevision,
          creation: histMeta.creation,
          titleRevision: histMeta.titleRevision,
          remarks: histMeta.remarks,
        });
      });
    }
  }

  return articles;
}

// Global in-memory cache for fetched revision articles (TTL: 2 hours)
const revisionArticlesCache = new Map<string, { articles: any[]; timestamp: number }>();
const REVISION_ARTICLES_CACHE_TTL = 2 * 60 * 60 * 1000;

// Master Revision Article Fetcher: Handles DRF XML, DRF HTML, Web Popup fallbacks for Laws & Administrative Rules
async function fetchArticlesForRevision(
  ocKey: string,
  rev: any,
  globalTargetType: string = 'law'
): Promise<any[]> {
  const isAdmrul =
    rev.targetType === 'admrul' ||
    globalTargetType === 'admrul' ||
    (rev.name || '').includes('규정') ||
    (rev.name || '').includes('고시') ||
    (rev.name || '').includes('훈령') ||
    (rev.name || '').includes('예규');

  const idCandidate = (rev.lawMst || rev.seq || rev.id || rev.lawId || '').toString().trim();
  if (!idCandidate) {
    return [];
  }

  const cleanEnfDate = (rev.enforcementDate || '').replace(/\D/g, '');
  const cleanPromDate = (rev.promulgationDate || '').replace(/\D/g, '');

  const cacheKey = `${isAdmrul ? 'admrul' : 'law'}_${idCandidate}_${(rev.name || '').trim()}_${(rev.promulgationNo || '').trim()}_${cleanEnfDate}_${cleanPromDate}`;
  const cached = revisionArticlesCache.get(cacheKey);
  if (cached && Date.now() - cached.timestamp < REVISION_ARTICLES_CACHE_TTL && cached.articles.length > 0) {
    return cached.articles;
  }

  console.log(`[fetchArticlesForRevision] Target: '${rev.name}', isAdmrul: ${isAdmrul}, id: '${idCandidate}', enfDate: '${rev.enforcementDate}' (${cleanEnfDate})`);

  if (isAdmrul) {
    // 1. Direct Web Body HTML (admRulLsInfoR.do) - Most complete and accurate (contains full articles, chapters, clauses for all notices)
    const bodyUrls: string[] = [];
    if (cleanEnfDate && cleanEnfDate.length === 8) {
      bodyUrls.push(`https://www.law.go.kr/LSW/admRulLsInfoR.do?admRulSeq=${encodeURIComponent(idCandidate)}&efYd=${cleanEnfDate}&chrClsCd=010202`);
      bodyUrls.push(`https://www.law.go.kr/LSW/admRulLsInfoR.do?admRulSeq=${encodeURIComponent(idCandidate)}&efYd=${cleanEnfDate}`);
    }
    bodyUrls.push(
      `https://www.law.go.kr/LSW/admRulLsInfoR.do?admRulSeq=${encodeURIComponent(idCandidate)}&chrClsCd=010202`,
      `https://www.law.go.kr/LSW/admRulLsInfoR.do?admRulSeq=${encodeURIComponent(idCandidate)}`,
      `https://www.law.go.kr/LSW/ileAdmRulLsInfoR.do?admRulSeq=${encodeURIComponent(idCandidate)}`,
      `https://www.law.go.kr/admRulLsInfoR.do?admRulSeq=${encodeURIComponent(idCandidate)}`
    );

    for (const url of bodyUrls) {
      try {
        const res = await fetch(url, {
          headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
          signal: AbortSignal.timeout(5000),
        });
        if (res.ok) {
          const html = await res.text();
          if (html && (html.includes('제1조') || html.includes('조문') || html.includes('본문') || html.includes('artclNo'))) {
            const articles = parseArticlesFromHtmlText(html);
            if (articles.length > 0) {
              console.log(`[fetchArticlesForRevision] Admrul Body HTML Success: ${articles.length} articles from ${url}`);
              revisionArticlesCache.set(cacheKey, { articles, timestamp: Date.now() });
              return articles;
            }
          }
        }
      } catch (err: any) {
        console.warn(`[Admrul Body HTML Fetch Error] ${url}:`, err?.message);
      }
    }

    // 2. DRF Administrative Rule XML (by ID or MST with efYd)
    const xmlUrls: string[] = [];
    if (cleanEnfDate && cleanEnfDate.length === 8) {
      xmlUrls.push(`http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(ocKey)}&target=admrul&ID=${encodeURIComponent(idCandidate)}&efYd=${cleanEnfDate}&type=XML`);
      xmlUrls.push(`http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(ocKey)}&target=admrul&MST=${encodeURIComponent(idCandidate)}&efYd=${cleanEnfDate}&type=XML`);
    }
    xmlUrls.push(
      `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(ocKey)}&target=admrul&ID=${encodeURIComponent(idCandidate)}&type=XML`,
      `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(ocKey)}&target=admrul&MST=${encodeURIComponent(idCandidate)}&type=XML`
    );

    for (const url of xmlUrls) {
      try {
        const res = await fetch(url, {
          headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
          signal: AbortSignal.timeout(5000),
        });
        if (res.ok) {
          const xml = await res.text();
          if (xml && (xml.includes('<조문') || xml.includes('<본문') || xml.includes('<조문내용') || xml.includes('<행정규칙') || xml.includes('<AdmRul'))) {
            const parsed = xmlParser.parse(xml);
            const root = normalizeAdmrulXmlRoot(parsed);
            const articles = parseAdmrulArticlesFromXmlRoot(root);
            if (articles.length > 0) {
              console.log(`[fetchArticlesForRevision] Admrul XML Success: ${articles.length} articles from ${url}`);
              revisionArticlesCache.set(cacheKey, { articles, timestamp: Date.now() });
              return articles;
            }
          }
        }
      } catch (err: any) {
        console.warn(`[Admrul XML Fetch Error] ${url}:`, err?.message);
      }
    }

    // 3. DRF HTML & Web Popup fallbacks for Administrative Rules
    const htmlUrls: string[] = [];
    if (cleanEnfDate && cleanEnfDate.length === 8) {
      htmlUrls.push(`http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(ocKey)}&target=admrul&ID=${encodeURIComponent(idCandidate)}&efYd=${cleanEnfDate}&type=HTML`);
      htmlUrls.push(`https://www.law.go.kr/LSW/admRulInfoP.do?admRulSeq=${encodeURIComponent(idCandidate)}&efYd=${cleanEnfDate}&chrClsCd=010202`);
    }
    htmlUrls.push(
      `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(ocKey)}&target=admrul&ID=${encodeURIComponent(idCandidate)}&type=HTML`,
      `https://www.law.go.kr/LSW/admRulInfoP.do?admRulSeq=${encodeURIComponent(idCandidate)}&chrClsCd=010202`,
      `http://www.law.go.kr/admRulInfoP.do?admRulSeq=${encodeURIComponent(idCandidate)}`
    );

    for (const url of htmlUrls) {
      try {
        const res = await fetch(url, {
          headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
          signal: AbortSignal.timeout(5000),
        });
        if (res.ok) {
          const html = await res.text();
          const articles = parseArticlesFromHtmlText(html);
          if (articles.length > 0) {
            console.log(`[fetchArticlesForRevision] Admrul HTML Success: ${articles.length} articles from ${url}`);
            revisionArticlesCache.set(cacheKey, { articles, timestamp: Date.now() });
            return articles;
          }
        }
      } catch (err: any) {
        console.warn(`[Admrul HTML Fetch Error] ${url}:`, err?.message);
      }
    }
  } else {
    // Law revisions:
    // When cleanEnfDate is present, prioritize target=eflaw and lsEfInfoP.do because target=law ignores efYd and always returns the current version!
    const xmlUrls: string[] = [];
    if (cleanEnfDate && cleanEnfDate.length === 8) {
      xmlUrls.push(
        `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(ocKey)}&target=eflaw&MST=${encodeURIComponent(idCandidate)}&efYd=${cleanEnfDate}&type=XML`,
        `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(ocKey)}&target=eflaw&ID=${encodeURIComponent(idCandidate)}&efYd=${cleanEnfDate}&type=XML`,
        `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(ocKey)}&target=law&MST=${encodeURIComponent(idCandidate)}&efYd=${cleanEnfDate}&type=XML`,
        `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(ocKey)}&target=law&ID=${encodeURIComponent(idCandidate)}&efYd=${cleanEnfDate}&type=XML`
      );
    } else {
      xmlUrls.push(
        `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(ocKey)}&target=law&MST=${encodeURIComponent(idCandidate)}&type=XML`,
        `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(ocKey)}&target=law&ID=${encodeURIComponent(idCandidate)}&type=XML`,
        `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(ocKey)}&target=eflaw&MST=${encodeURIComponent(idCandidate)}&type=XML`
      );
    }

    for (const url of xmlUrls) {
      try {
        const res = await fetch(url, {
          headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
          signal: AbortSignal.timeout(7500),
        });
        if (res.ok) {
          const xml = await res.text();
          if (xml && (xml.includes('<조문') || xml.includes('<조문단위') || xml.includes('<법령') || xml.includes('<Law') || xml.includes('<efLaw'))) {
            const parsed = xmlParser.parse(xml);
            const root = normalizeLawXmlRoot(parsed);
            const articles = parseArticlesFromXmlRoot(root);
            if (articles.length > 0) {
              console.log(`[fetchArticlesForRevision] Law XML Success: ${articles.length} articles from ${url}`);
              revisionArticlesCache.set(cacheKey, { articles, timestamp: Date.now() });
              return articles;
            }
          }
        }
      } catch (err: any) {
        console.log(`[fetchArticlesForRevision] Law XML fallback (${url}):`, err?.message);
      }
    }

    // 2. Law HTML Web Popup fallback with efYd (extremely fast & complete)
    const htmlUrls: string[] = [];
    if (cleanEnfDate && cleanEnfDate.length === 8) {
      htmlUrls.push(
        `https://www.law.go.kr/LSW/lsEfInfoP.do?lsiSeq=${encodeURIComponent(idCandidate)}&efYd=${cleanEnfDate}&chrClsCd=010202`,
        `https://www.law.go.kr/LSW/lsInfoP.do?lsiSeq=${encodeURIComponent(idCandidate)}&efYd=${cleanEnfDate}&chrClsCd=010202`,
        `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(ocKey)}&target=eflaw&MST=${encodeURIComponent(idCandidate)}&efYd=${cleanEnfDate}&type=HTML`,
        `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(ocKey)}&target=law&MST=${encodeURIComponent(idCandidate)}&efYd=${cleanEnfDate}&type=HTML`
      );
    }
    htmlUrls.push(
      `https://www.law.go.kr/LSW/lsInfoP.do?lsiSeq=${encodeURIComponent(idCandidate)}&chrClsCd=010202`,
      `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(ocKey)}&target=law&MST=${encodeURIComponent(idCandidate)}&type=HTML`
    );

    for (const url of htmlUrls) {
      try {
        const res = await fetch(url, {
          headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
          signal: AbortSignal.timeout(7500),
        });
        if (res.ok) {
          const html = await res.text();
          const articles = parseArticlesFromHtmlText(html);
          if (articles.length > 0) {
            console.log(`[fetchArticlesForRevision] Law HTML Success: ${articles.length} articles from ${url}`);
            revisionArticlesCache.set(cacheKey, { articles, timestamp: Date.now() });
            return articles;
          }
        }
      } catch (err: any) {
        console.log(`[fetchArticlesForRevision] Law HTML fallback (${url}):`, err?.message);
      }
    }
  }

  return [];
}

// ==========================================
// UNIFIED SEARCH & REVISIONS API ROUTES
// ==========================================

// 1. Unified Search: [법령 / 행정규칙] 검색 (법/시행령/시행규칙 다중필터 및 정확히일치/포함 모드)
app.get('/api/unified/search', async (req, res) => {
  try {
    const ocKey = (req.query.ocKey as string) || DEFAULT_OC_KEY;
    const targetType = (req.query.targetType as string) || 'law'; // 'law' | 'admrul'
    const queryName = ((req.query.query as string) || (targetType === 'admrul' ? '관세' : '관세법')).trim();
    const subTypesParam = (req.query.subTypes as string) || 'law,decree,rule';
    const subTypes = subTypesParam.split(',').map((s) => s.trim().toLowerCase()); // ['law', 'decree', 'rule']
    const matchMode = ((req.query.matchMode as string) || 'exact').toLowerCase(); // 'exact' | 'contains'
    const displayCount = (req.query.display as string) || '500';

    if (targetType === 'admrul') {
      // 행정규칙 검색 API (target=admrul) - 다중 페이지(1~5) 전수 조회
      const cleanQueryNoSpace = queryName.replace(/\s+/g, '').toLowerCase();
      const allMapped: any[] = [];
      const seenAdmKeys = new Set<string>();

      for (let page = 1; page <= 5; page++) {
        const searchUrl = `http://www.law.go.kr/DRF/lawSearch.do?OC=${encodeURIComponent(
          ocKey
        )}&target=admrul&query=${encodeURIComponent(queryName)}&page=${page}&display=100&type=XML`;

        console.log(`[Unified Search: Admrul] Mode: ${matchMode}, Page ${page}, Fetching: ${searchUrl}`);
        try {
          const response = await fetch(searchUrl, {
            headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
          });

          if (!response.ok) break;

          const xmlText = await response.text();
          const parsed = xmlParser.parse(xmlText);
          const searchRoot = parsed.AdmRulSearch || parsed.admRulSearch || parsed;
          let rawList = searchRoot.admrul || searchRoot.AdmRul || [];
          if (!Array.isArray(rawList)) rawList = rawList ? [rawList] : [];
          if (rawList.length === 0) break;

          for (const item of rawList) {
            const rawPramNo = getText(item.발령번호 || item.공포번호 || item.pramNo);
            const ruleType = getText(item.행정규칙종류 || item.행정규칙종류명 || item.구분 || '고시');
            const name = getText(item.행정규칙명 || item.admRulNm || item['#text']).trim();
            const id = getText(item.행정규칙일련번호 || item.admrulSeq || item.MST || item.ID);

            if (!name) continue;
            const itemNoSpace = name.replace(/\s+/g, '').toLowerCase();

            // Match filtering
            if (matchMode === 'exact') {
              if (name !== queryName && itemNoSpace !== cleanQueryNoSpace) continue;
            } else {
              if (!name.toLowerCase().includes(queryName.toLowerCase()) && !itemNoSpace.includes(cleanQueryNoSpace)) {
                continue;
              }
            }

            const itemKey = `${name}_${ruleType}`;
            if (!seenAdmKeys.has(itemKey)) {
              seenAdmKeys.add(itemKey);
              allMapped.push({
                id,
                seq: id,
                name,
                targetType: 'admrul' as const,
                department: getText(item.소관부처명 || item.소관부처 || item.orgNm || '관세청'),
                promulgationDate: formatDate(getText(item.발령일자 || item.공포일자 || item.pramDate)),
                promulgationNo: rawPramNo ? `${ruleType} 제${rawPramNo.replace(/[^0-9-]/g, '')}호` : '',
                enforcementDate: formatDate(getText(item.시행일자 || item.efYd)),
                revisionType: getText(item.제개정구분명 || item.제개정구분 || item.gubun || '일부개정'),
                ruleType: ruleType,
                currentYn: getText(item.현행연혁구분 || item.currentYn || 'Y'),
              });
            }
          }
        } catch (admErr) {
          console.warn(`[Admrul Search Error] page ${page}:`, admErr);
          break;
        }
      }

      // Sort administrative rules by name
      allMapped.sort((a, b) => (a.name || '').localeCompare(b.name || '', 'ko'));

      return res.json({
        success: true,
        targetType: 'admrul',
        matchMode,
        count: allMapped.length,
        totalCount: String(allMapped.length),
        results: allMapped,
      });
    } else {
      // 법령 검색 (target=eflaw / target=law)
      const aliasInfo = resolveLawAliases(queryName);
      const canonicalBaseNames = aliasInfo.canonicalNames;
      const matchedAliasNote = aliasInfo.matchedAlias;

      const searchQueriesToRun: Array<{ query: string; subType: 'law' | 'decree' | 'rule'; isPredecessor?: boolean }> = [];

      const queryNoSpace = queryName.replace(/\s+/g, '').toLowerCase();
      const stemQuery = queryName.replace(/(\s*시행령|\s*시행규칙|\s*법률|\s*법)$/, '').trim();
      const stemQueryNoSpace = stemQuery.replace(/\s+/g, '').toLowerCase();

      if (matchMode === 'exact') {
        for (const baseItem of canonicalBaseNames) {
          const base = baseItem.replace(/(\s*시행령|\s*시행규칙)$/, '').trim();
          const baseWithLaw = base.endsWith('법') || base.endsWith('법률') ? base : `${base}법`;

          if (subTypes.includes('law')) {
            searchQueriesToRun.push({ query: baseWithLaw, subType: 'law' });
            if (base !== baseWithLaw) {
              searchQueriesToRun.push({ query: base, subType: 'law' });
            }
          }
          if (subTypes.includes('decree')) {
            searchQueriesToRun.push({ query: `${baseWithLaw} 시행령`, subType: 'decree' });
            if (base !== baseWithLaw) {
              searchQueriesToRun.push({ query: `${base} 시행령`, subType: 'decree' });
            }
          }
          if (subTypes.includes('rule')) {
            searchQueriesToRun.push({ query: `${baseWithLaw} 시행규칙`, subType: 'rule' });
            if (base !== baseWithLaw) {
              searchQueriesToRun.push({ query: `${base} 시행규칙`, subType: 'rule' });
            }
          }

          // Predecessor laws (변경 전 구법, 예: 외국환거래법 ➔ 외국환관리법)
          if (PREDECESSOR_MAP[baseWithLaw]) {
            for (const pred of PREDECESSOR_MAP[baseWithLaw]) {
              if (subTypes.includes('law')) {
                searchQueriesToRun.push({ query: pred, subType: 'law', isPredecessor: true });
              }
              if (subTypes.includes('decree')) {
                searchQueriesToRun.push({ query: `${pred} 시행령`, subType: 'decree', isPredecessor: true });
              }
              if (subTypes.includes('rule')) {
                searchQueriesToRun.push({ query: `${pred} 시행규칙`, subType: 'rule', isPredecessor: true });
              }
            }
          }
        }
      } else {
        // Contains mode: wide multi-term queries (queryName, stem, queryName+법, and canonical expansions)
        const addedTerms = new Set<string>();

        const addTerm = (term: string, isPred?: boolean) => {
          const cleanT = term.trim();
          if (!cleanT || addedTerms.has(cleanT)) return;
          addedTerms.add(cleanT);
          searchQueriesToRun.push({ query: cleanT, subType: 'law', isPredecessor: isPred });
        };

        addTerm(queryName);
        if (stemQuery && stemQuery !== queryName) {
          addTerm(stemQuery);
          addTerm(`${stemQuery}법`);
        } else if (!queryName.endsWith('법') && !queryName.endsWith('법률')) {
          addTerm(`${queryName}법`);
        }

        for (const baseItem of canonicalBaseNames) {
          addTerm(baseItem);
          const baseClean = baseItem.replace(/(\s*시행령|\s*시행규칙|\s*법률|\s*법)$/, '').trim();
          if (baseClean && baseClean !== baseItem) {
            addTerm(baseClean);
          }
        }

        // Add predecessors in contains mode
        if (PREDECESSOR_MAP[queryName]) {
          for (const pred of PREDECESSOR_MAP[queryName]) {
            addTerm(pred, true);
          }
        }
        if (PREDECESSOR_MAP[`${queryName}법`]) {
          for (const pred of PREDECESSOR_MAP[`${queryName}법`]) {
            addTerm(pred, true);
          }
        }
      }

      console.log(`[Unified Search: Law] Running queries (${matchMode}):`, searchQueriesToRun.map((q) => q.query));

      const collectedResults: any[] = [];
      const seenIds = new Set<string>();

      const targetsToSearch = matchMode === 'contains' ? ['eflaw', 'law'] : ['eflaw'];
      const maxPages = matchMode === 'contains' ? 3 : 1;

      for (const searchTask of searchQueriesToRun) {
        for (const tgt of targetsToSearch) {
          for (let page = 1; page <= maxPages; page++) {
            const searchUrl = `http://www.law.go.kr/DRF/lawSearch.do?OC=${encodeURIComponent(
              ocKey
            )}&target=${tgt}&query=${encodeURIComponent(searchTask.query)}&page=${page}&display=100&type=XML`;

            console.log(`[Unified Search: Law] Target: ${tgt}, Term: ${searchTask.query}, Page ${page}`);
            try {
              const response = await fetch(searchUrl, {
                headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
              });

              if (!response.ok) break;

              const xmlText = await response.text();
              const parsed = xmlParser.parse(xmlText);
              const searchRoot = parsed.LawSearch || parsed.lawSearch || parsed;
              let rawList = searchRoot.law || searchRoot.Law || [];
              if (!Array.isArray(rawList)) rawList = rawList ? [rawList] : [];
              if (rawList.length === 0) break;

              for (const item of rawList) {
                const name = getText(item.법령명한글 || item.법령명_한글 || item.lawName || item['#text']).trim();
                if (!name) continue;

                const rawPromNo = getText(item.공포번호);
                const lawType = getText(item.법령구분명 || item.법령종류 || '법률');

                // Classify subType
                let itemSubType: 'law' | 'decree' | 'rule' = 'law';
                if (name.includes('시행령') || lawType.includes('대통령령') || lawType.includes('시행령')) {
                  itemSubType = 'decree';
                } else if (name.includes('시행규칙') || lawType.includes('부령') || lawType.includes('총리령') || lawType.includes('규칙')) {
                  itemSubType = 'rule';
                }

                // Check if this subType is requested by user
                if (!subTypes.includes(itemSubType)) continue;

                const nameNoSpace = name.replace(/\s+/g, '').toLowerCase();
                const targetTaskNoSpace = searchTask.query.replace(/\s+/g, '').toLowerCase();

                if (matchMode === 'exact') {
                  const matchesTaskExact = nameNoSpace === targetTaskNoSpace || name === searchTask.query;
                  const matchesCanonicalExact = canonicalBaseNames.some((c) => {
                    const cNoSpace = c.replace(/\s+/g, '').toLowerCase();
                    return (
                      nameNoSpace === cNoSpace ||
                      nameNoSpace === `${cNoSpace}시행령` ||
                      nameNoSpace === `${cNoSpace}시행규칙` ||
                      nameNoSpace === `${cNoSpace}법` ||
                      nameNoSpace === `${cNoSpace}법시행령` ||
                      nameNoSpace === `${cNoSpace}법시행규칙`
                    );
                  });
                  const matchesQueryExact =
                    nameNoSpace === queryNoSpace ||
                    nameNoSpace === `${queryNoSpace}법` ||
                    nameNoSpace === `${queryNoSpace}법시행령` ||
                    nameNoSpace === `${queryNoSpace}법시행규칙`;

                  if (!matchesTaskExact && !matchesCanonicalExact && !matchesQueryExact) {
                    continue;
                  }
                } else {
                  // Contains mode: match either original query, query stem, or canonical alias
                  const matchesContains =
                    name.toLowerCase().includes(queryName.toLowerCase()) ||
                    nameNoSpace.includes(queryNoSpace) ||
                    (stemQueryNoSpace.length >= 2 && nameNoSpace.includes(stemQueryNoSpace)) ||
                    canonicalBaseNames.some((c) => {
                      const cNoSpace = c.replace(/\s+/g, '').toLowerCase();
                      return name.toLowerCase().includes(c.toLowerCase()) || nameNoSpace.includes(cNoSpace);
                    });
                  if (!matchesContains) continue;
                }

                let formattedPromNo = rawPromNo;
                if (rawPromNo && !rawPromNo.startsWith('법률') && !rawPromNo.startsWith('제') && !rawPromNo.startsWith('대통령령')) {
                  const digits = rawPromNo.replace(/[^0-9]/g, '');
                  formattedPromNo = digits ? `법률 제${digits}호` : rawPromNo;
                }

                const id = getText(item.법령일련번호 || item.MST || item.mst || item.법령ID);
                const key = `${name}_${itemSubType}`;

                if (!seenIds.has(key)) {
                  seenIds.add(key);
                  collectedResults.push({
                    id,
                    seq: id,
                    name,
                    targetType: 'law' as const,
                    subType: itemSubType,
                    department: getText(item.소관부처명 || item.소관부처 || '기획재정부'),
                    promulgationDate: formatDate(getText(item.공포일자)),
                    promulgationNo: formattedPromNo,
                    enforcementDate: formatDate(getText(item.시행일자)),
                    revisionType: getText(item.제개정구분명 || item.제개정구분 || '일부개정'),
                    ruleType: lawType || (itemSubType === 'decree' ? '대통령령' : itemSubType === 'rule' ? '부령' : '법률'),
                    currentYn: getText(item.현행연혁구분 || 'Y'),
                    isPredecessor: !!searchTask.isPredecessor,
                    predecessorNote: searchTask.isPredecessor ? `(변경전 구법)` : undefined,
                    matchedAliasNote: matchedAliasNote || undefined,
                  });
                }
              }
            } catch (taskErr) {
              console.warn(`[Search Law] Subquery error for ${searchTask.query} (tgt: ${tgt}, page: ${page}):`, taskErr);
            }
          }
        }
      }

      // Sort candidate laws: 법률(1) -> 시행령(2) -> 시행규칙(3), then by name
      const hierarchyOrder: Record<string, number> = { law: 1, decree: 2, rule: 3 };
      collectedResults.sort((a, b) => {
        const orderA = hierarchyOrder[a.subType || 'law'] || 4;
        const orderB = hierarchyOrder[b.subType || 'law'] || 4;
        if (orderA !== orderB) return orderA - orderB;
        return (a.name || '').localeCompare(b.name || '', 'ko');
      });

      return res.json({
        success: true,
        targetType: 'law',
        matchMode,
        subTypes,
        matchedAliasNote,
        count: collectedResults.length,
        totalCount: String(collectedResults.length),
        results: collectedResults,
      });
    }
  } catch (err: any) {
    console.error('Unified Search Error:', err);
    return res.status(500).json({ error: err.message || '통합 검색 중 오류가 발생했습니다.' });
  }
});

// 2. Unified Revisions: 특정 법령 또는 행정규칙의 전체 개정연혁 목록 (다중 법령/시행령/시행규칙 통합 지원)
app.get('/api/unified/revisions', async (req, res) => {
  try {
    const ocKey = (req.query.ocKey as string) || DEFAULT_OC_KEY;
    const targetType = (req.query.targetType as string) || 'law';
    const name = ((req.query.name as string) || '').trim();
    const namesParam = (req.query.names as string) || '';
    const subTypesParam = (req.query.subTypes as string) || 'law,decree,rule';
    const subTypes = subTypesParam.split(',').map((s) => s.trim().toLowerCase());
    const matchMode = ((req.query.matchMode as string) || 'exact').toLowerCase() as 'exact' | 'contains';

    if (targetType === 'admrul') {
      let targetNames: string[] = [];
      if (namesParam) {
        targetNames = namesParam.split(',').map((s) => s.trim()).filter(Boolean);
      } else if (name) {
        targetNames = [name.trim()];
      } else {
        targetNames = ['관세평가 운영에 관한 고시'];
      }

      console.log(`[Unified Revisions Admrul] Target admrul names:`, targetNames);

      const allAdmrulRevs: any[] = [];
      const seenKeys = new Set<string>();

      for (const admrulName of targetNames) {
        // Fetch revisions for this specific administrative rule
        // In Step 2 (revisions view), query using exact match so that other unrelated rules containing the keyword are NOT included
        const admrulRevs = await fetchAdmrulRevisions(ocKey, admrulName, 0, 'exact', true);
        for (const item of admrulRevs) {
          const itemSeq = (item.seq || item.lawMst || item.lawId || item.id || '').toString();
          const cleanItemName = (item.name || item.lawName || admrulName).trim();
          const enfDate = (item.enforcementDate || '').replace(/\D/g, '');
          const promDate = (item.promulgationDate || '').replace(/\D/g, '');
          const promNo = (item.promulgationNo || '').trim();

          const primaryKey = itemSeq
            ? `ADM_${itemSeq}_${enfDate}`
            : `KEY_ADM_${cleanItemName}_${enfDate}_${promDate}_${promNo}`;

          if (seenKeys.has(primaryKey)) continue;
          seenKeys.add(primaryKey);

          allAdmrulRevs.push({
            id: itemSeq || `${cleanItemName}_${enfDate}`,
            seq: itemSeq || `${cleanItemName}_${enfDate}`,
            lawMst: itemSeq,
            lawId: itemSeq,
            name: cleanItemName,
            lawName: cleanItemName,
            targetType: 'admrul' as const,
            promulgationDate: item.promulgationDate || '',
            promulgationNo: item.promulgationNo || '',
            enforcementDate: item.enforcementDate || '',
            revisionType: item.revisionType || '일부개정',
            department: item.department || '관세청',
            ruleType: item.ruleType || '고시',
            buchikText: item.buchikText || '',
            isPredecessor: !!item.isPredecessor,
            predecessorNote: item.predecessorNote || '',
          });
        }
      }

      const sortedAdmrul = sortRevisionsByEnforcementDateDesc(allAdmrulRevs);
      return res.json({ success: true, count: sortedAdmrul.length, revisions: sortedAdmrul });
    } else {
      // Determine list of law names to collect revisions for
      let targetNames: Array<{ name: string; subType: 'law' | 'decree' | 'rule' }> = [];

      if (namesParam) {
        // Multiple comma-separated law names explicitly selected by user
        const split = namesParam.split(',').map((s) => s.trim()).filter(Boolean);
        targetNames = split.map((nm) => {
          let st: 'law' | 'decree' | 'rule' = 'law';
          if (nm.includes('시행령')) st = 'decree';
          else if (nm.includes('시행규칙')) st = 'rule';
          return { name: nm, subType: st };
        });
      } else if (name) {
        const base = name.trim();
        const baseWithLaw = base.endsWith('법') || base.endsWith('법률') ? base : `${base}법`;
        if (matchMode === 'exact' && !base.includes('시행령') && !base.includes('시행규칙')) {
          if (subTypes.includes('law')) {
            targetNames.push({ name: baseWithLaw, subType: 'law' });
          }
          if (subTypes.includes('decree')) {
            targetNames.push({ name: `${baseWithLaw} 시행령`, subType: 'decree' });
          }
          if (subTypes.includes('rule')) {
            targetNames.push({ name: `${baseWithLaw} 시행규칙`, subType: 'rule' });
          }
        } else {
          let st: 'law' | 'decree' | 'rule' = 'law';
          if (base.includes('시행령')) st = 'decree';
          else if (base.includes('시행규칙')) st = 'rule';
          targetNames.push({ name: base, subType: st });
        }
      } else {
        targetNames.push({ name: '관세법', subType: 'law' });
      }

      console.log(`[Unified Revisions] Target law names:`, targetNames.map((t) => `${t.name} (${t.subType})`));

      const allRevisions: any[] = [];
      const seenKeys = new Set<string>();

      // In Step 2, each targetName is explicitly chosen, so query with exact match
      const queryMatchMode = namesParam ? 'exact' : matchMode;

      for (const target of targetNames) {
        const targetSubTypes = [target.subType];
        const lawRevs = await fetchLawRevisions(ocKey, target.name, 0, queryMatchMode, targetSubTypes);
        for (const item of lawRevs) {
          const id = (item.lawMst || item.lawId || item.id || item.seq || '').toString().trim();
          const cleanName = (item.name || item.lawName || target.name || '').replace(/\s+/g, '');
          const enfDate = (item.enforcementDate || '').replace(/\D/g, '');
          const promDate = (item.promulgationDate || '').replace(/\D/g, '');

          // Determine subtype accurately
          let itemSubType: 'law' | 'decree' | 'rule' = target.subType;
          if (item.name?.includes('시행령') || item.lawName?.includes('시행령') || item.lawType?.includes('대통령령')) {
            itemSubType = 'decree';
          } else if (item.name?.includes('시행규칙') || item.lawName?.includes('시행규칙') || item.lawType?.includes('부령') || item.lawType?.includes('총리령')) {
            itemSubType = 'rule';
          } else if (!item.name?.includes('시행령') && !item.name?.includes('시행규칙')) {
            itemSubType = 'law';
          }

          // Check if requested in subTypes
          if (subTypes.length > 0 && !subTypes.includes(itemSubType)) {
            continue;
          }

          const primaryKey = id
            ? `MST_${id}_${enfDate}`
            : `KEY_${itemSubType}_${cleanName}_${enfDate}_${promDate}_${item.promulgationNo}`;

          if (seenKeys.has(primaryKey)) {
            continue;
          }
          seenKeys.add(primaryKey);

          const dept = item.department || '기획재정부';
          const ruleType = itemSubType === 'decree'
            ? '대통령령'
            : itemSubType === 'rule'
            ? (dept ? `${dept}령` : '부령')
            : '법률';

          allRevisions.push({
            id: id || `${cleanName}_${enfDate}`,
            seq: id || `${cleanName}_${enfDate}`,
            lawMst: id,
            lawId: id,
            name: item.name || item.lawName || target.name,
            lawName: item.lawName || item.name || target.name,
            targetType: 'law' as const,
            subType: itemSubType,
            promulgationDate: item.promulgationDate || '',
            promulgationNo: item.promulgationNo || '',
            enforcementDate: item.enforcementDate || '',
            revisionType: item.revisionType || '일부개정',
            department: dept,
            lawType: item.lawType || ruleType,
            ruleType: item.ruleType || ruleType,
            isPredecessor: !!item.isPredecessor,
            predecessorNote: item.predecessorNote || '',
          });
        }
      }

      // Sort strictly: 법(1) -> 시행령(2) -> 시행규칙(3), and within each by enforcement date descending
      const sortedLaw = sortRevisionsByHierarchyAndDate(allRevisions);
      console.log(`[Unified Revisions] Returning total sorted revisions: ${sortedLaw.length}`);
      return res.json({ success: true, count: sortedLaw.length, revisions: sortedLaw });
    }
  } catch (err: any) {
    console.error('Unified Revisions Error:', err);
    return res.status(500).json({ error: err.message || '개정연혁 조회 중 오류가 발생했습니다.' });
  }
});

// 3. Unified Detail: 법령 또는 행정규칙 상세 조문 데이터 가져오기
app.get('/api/unified/detail', async (req, res) => {
  try {
    const ocKey = (req.query.ocKey as string) || DEFAULT_OC_KEY;
    const targetType = (req.query.targetType as string) || 'law';
    const seq = (req.query.seq as string) || (req.query.mst as string) || (req.query.id as string) || '2100000281984';
    const name = (req.query.name as string) || '';

    const dummyRev = {
      id: seq,
      seq,
      lawMst: seq,
      lawId: seq,
      name,
      targetType: targetType as any,
    };

    const articles = await fetchArticlesForRevision(ocKey, dummyRev, targetType);

    if (targetType === 'admrul') {
      const detailUrl = `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(
        ocKey
      )}&target=admrul&ID=${encodeURIComponent(seq)}&type=XML`;

      let basicInfo: any = {};
      try {
        const detailRes = await fetch(detailUrl, {
          headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
        });
        if (detailRes.ok) {
          const detailXml = await detailRes.text();
          const parsed = xmlParser.parse(detailXml);
          const root = normalizeAdmrulXmlRoot(parsed);
          basicInfo = root.행정규칙기본정보 || root.기본정보 || root.BasicInfo || {};
        }
      } catch (e) {}

      const ruleType = getText(basicInfo.행정규칙종류 || basicInfo.행정규칙구분 || '고시');
      const rawPramNo = getText(basicInfo.발령번호 || basicInfo.공포번호);
      const dept = getText(basicInfo.소관부처명 || basicInfo.소관부처 || '관세청');

      const info: any = {
        lawId: getText(basicInfo.행정규칙일련번호 || seq),
        lawMst: seq,
        lawName: getText(basicInfo.행정규칙명 || name || '행정규칙'),
        promulgationDate: formatDate(getText(basicInfo.발령일자 || basicInfo.공포일자)),
        promulgationNo: rawPramNo ? `${dept} ${ruleType} 제${rawPramNo.replace(/[^0-9-]/g, '')}호` : `${dept} ${ruleType}`,
        enforcementDate: formatDate(getText(basicInfo.시행일자)),
        revisionType: getText(basicInfo.제개정구분 || '일부개정'),
        department: dept,
        lawType: ruleType,
        targetType: 'admrul',
        articleCount: articles.length,
      };

      return res.json({
        success: true,
        info,
        articles,
        fetchedAt: new Date().toISOString(),
      });
    } else {
      // Law detail
      const detailUrl = `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(
        ocKey
      )}&target=law&MST=${encodeURIComponent(seq)}&type=XML`;

      let basicInfo: any = {};
      try {
        const detailRes = await fetch(detailUrl, {
          headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
        });
        if (detailRes.ok) {
          const detailXml = await detailRes.text();
          const parsed = xmlParser.parse(detailXml);
          const root = normalizeLawXmlRoot(parsed);
          basicInfo = root.기본정보 || root.BasicInfo || {};
        }
      } catch (e) {}

      const rawPromNo = getText(basicInfo.공포번호);
      const lawType = getText(basicInfo.법종구분 || basicInfo.법령종류 || '법률');
      let formattedPromNo = rawPromNo;
      if (rawPromNo && !rawPromNo.startsWith('법률') && !rawPromNo.startsWith('제') && !rawPromNo.startsWith('대통령령')) {
        const digits = rawPromNo.replace(/[^0-9]/g, '');
        formattedPromNo = digits ? `법률 제${digits}호` : rawPromNo;
      }

      const info: any = {
        lawId: getText(basicInfo.법령ID || basicInfo.lawId || seq),
        lawMst: getText(basicInfo.법령일련번호 || seq),
        lawName: getText(basicInfo.법령명_한글 || basicInfo.법령명한글 || name || '관세법'),
        promulgationDate: formatDate(getText(basicInfo.공포일자)),
        promulgationNo: formattedPromNo,
        enforcementDate: formatDate(getText(basicInfo.시행일자)),
        revisionType: getText(basicInfo.제개정구분),
        department: getText(basicInfo.소관부처),
        lawType: lawType,
        targetType: 'law',
        articleCount: articles.length,
      };

      return res.json({
        success: true,
        info,
        articles,
        fetchedAt: new Date().toISOString(),
      });
    }
  } catch (err: any) {
    console.error('Unified Detail Error:', err);
    return res.status(500).json({ error: err.message || '상세 데이터 조회 중 오류가 발생했습니다.' });
  }
});

// ==========================================
// GOOGLE DRIVE API V3 FOLDER & FILE ROUTES
// ==========================================

// 4. Drive: Get or Create Folder [선택한 법령/행정규칙명_YYYYMMDD]
app.post('/api/drive/get-or-create-folder', async (req, res) => {
  try {
    const { accessToken, folderName } = req.body;
    if (!accessToken) {
      return res.status(401).json({ error: 'Google OAuth Access Token이 필요합니다.' });
    }

    if (!folderName) {
      return res.status(400).json({ error: '생성/조회할 폴더명을 지정해 주세요.' });
    }

    const auth = new google.auth.OAuth2();
    auth.setCredentials({ access_token: accessToken });
    const drive = google.drive({ version: 'v3', auth });

    // 1. Search if folder already exists
    const escapedName = folderName.replace(/['\\]/g, '\\$&');
    const searchRes = await drive.files.list({
      q: `name = '${escapedName}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false`,
      fields: 'files(id, name, webViewLink)',
      spaces: 'drive',
    });

    if (searchRes.data.files && searchRes.data.files.length > 0) {
      const existing = searchRes.data.files[0];
      return res.json({
        success: true,
        folder: {
          id: existing.id,
          name: existing.name,
          url: existing.webViewLink || `https://drive.google.com/drive/folders/${existing.id}`,
          isExisting: true,
          created: false,
        },
      });
    }

    // 2. Create new folder
    const createRes = await drive.files.create({
      requestBody: {
        name: folderName,
        mimeType: 'application/vnd.google-apps.folder',
      },
      fields: 'id, name, webViewLink',
    });

    return res.json({
      success: true,
      folder: {
        id: createRes.data.id,
        name: createRes.data.name,
        url: createRes.data.webViewLink || `https://drive.google.com/drive/folders/${createRes.data.id}`,
        isExisting: false,
        created: true,
      },
    });
  } catch (err: any) {
    console.error('Drive Folder API Error:', err);
    return res.status(500).json({ error: err.message || '구글 드라이브 폴더 생성 중 오류가 발생했습니다.' });
  }
});

// Drive: Permissions Revoke (외부 공유 권한 일괄 해제 -> 소유자 전용 비공개 전환)
app.post('/api/drive/permissions/revoke', async (req, res) => {
  try {
    const { accessToken, targetId, targetIds } = req.body;

    if (!accessToken) {
      return res.status(401).json({ error: 'Google OAuth Access Token이 필요합니다.' });
    }

    const auth = new google.auth.OAuth2();
    auth.setCredentials({ access_token: accessToken });
    const drive = google.drive({ version: 'v3', auth });

    const idsToRevoke: string[] = [];
    if (targetId) idsToRevoke.push(targetId);
    if (Array.isArray(targetIds)) {
      targetIds.forEach((id) => {
        if (id && !idsToRevoke.includes(id)) idsToRevoke.push(id);
      });
    }

    if (idsToRevoke.length === 0) {
      return res.status(400).json({ error: '권한을 해제할 드라이브 폴더 또는 파일 ID가 지정되지 않았습니다.' });
    }

    let totalDeleted = 0;

    for (const fileId of idsToRevoke) {
      try {
        const permList = await drive.permissions.list({ fileId, fields: 'permissions(id, role, type)' });
        const perms = permList.data.permissions || [];

        for (const p of perms) {
          if (p.role !== 'owner' && p.id) {
            try {
              await drive.permissions.delete({ fileId, permissionId: p.id });
              totalDeleted++;
            } catch (delErr) {
              console.warn(`Permission delete warn for ${fileId}/${p.id}:`, delErr);
            }
          }
        }
      } catch (listErr) {
        console.warn(`Permission list warn for ${fileId}:`, listErr);
      }
    }

    return res.json({
      success: true,
      totalRevokedPermissions: totalDeleted,
      message: `성공적으로 모든 외부 공유 권한이 해제되었습니다. 대상 항목들이 소유자 전용 '비공개' 상태로 안전하게 전환되었습니다.`,
    });
  } catch (err: any) {
    console.error('Drive Permissions Revoke Error:', err);
    return res.status(500).json({ error: err.message || '권한 해제 중 오류가 발생했습니다.' });
  }
});


// API Route: Create/Update Google Sheet with Customs Act Data
app.post('/api/sheets/save', async (req, res) => {
  try {
    const { accessToken, lawData, config } = req.body;

    if (!accessToken) {
      return res.status(401).json({ error: '유효한 Google OAuth Access Token이 필요합니다. 상단의 Google 계정 연결 버튼을 눌러주세요.' });
    }

    if (!lawData || !lawData.info || !lawData.articles) {
      return res.status(400).json({ error: '저장할 관세법 데이터가 올바르지 않습니다. 다시 수집을 진행해 주세요.' });
    }

    const auth = new google.auth.OAuth2();
    auth.setCredentials({ access_token: accessToken });

    const sheets = google.sheets({ version: 'v4', auth });
    const drive = google.drive({ version: 'v3', auth });
    const ocKey = req.body.ocKey || DEFAULT_OC_KEY;

    // Helper to retry Google API calls on HTTP 429 / Rate Limit
    const callApiWithRetry = async <T>(fn: () => Promise<T>, retries = 4, delay = 1000): Promise<T> => {
      try {
        return await fn();
      } catch (err: any) {
        const isRateLimit =
          err?.status === 429 ||
          err?.code === 429 ||
          err?.message?.includes('Quota') ||
          err?.message?.includes('rate') ||
          err?.message?.includes('RESOURCE_EXHAUSTED');

        if (isRateLimit && retries > 0) {
          console.warn(`[Google API Rate Limit] Pausing ${delay}ms before retry...`);
          await new Promise((r) => setTimeout(r, delay));
          return callApiWithRetry(fn, retries - 1, delay * 2);
        }
        throw err;
      }
    };

    // CHECK SEPARATE FILES MODE (140 개별 구글 시트 파일 각각 생성 또는 기존 파일 재활용)
    if (config?.exportMode === 'separate_files_140' || (config?.exportAll140 && config?.exportMode !== 'single_file_140')) {
      let revisionList: any[] = req.body.revisions || [];
      if (!Array.isArray(revisionList) || revisionList.length === 0) {
        revisionList = await fetchAll140Revisions(ocKey);
      }

      revisionList = sortRevisionsByEnforcementDateDesc(revisionList);

      console.log(`[Batch Export Separate] Processing ${revisionList.length} revisions (checking Drive for existing files)...`);
      const createdFiles: Array<{ title: string; spreadsheetId: string; url: string; promulgationNo: string; enforcementDate: string; isExisting?: boolean }> = [];

      // Process sequentially / small chunks with pacing to stay safely under Google Sheets API write quota
      const chunkSize = 3;
      for (let i = 0; i < revisionList.length; i += chunkSize) {
        const chunk = revisionList.slice(i, i + chunkSize);
        await Promise.all(
          chunk.map(async (rev: any, chunkIndex: number) => {
            try {
              const detailUrl = `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(
                ocKey
              )}&target=law&MST=${encodeURIComponent(rev.lawMst)}&type=XML`;

              const detailRes = await fetch(detailUrl, {
                headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
              });

              if (!detailRes.ok) return;

              const detailXml = await detailRes.text();
              const parsed = xmlParser.parse(detailXml);
              const root = parsed.법령 || parsed.Law || parsed;
              const revArticles = parseArticlesFromXmlRoot(root);

              const revIndexNum = String(i + chunkIndex + 1).padStart(3, '0');
              const docTitle = `${revIndexNum}_[관세법] ${rev.promulgationNo || '개정본'} (${rev.enforcementDate || ''} 시행)`;

              let spId = '';
              let isExistingFile = false;

              // 1. Search Google Drive for an existing spreadsheet file with matching title
              try {
                const searchTitleEscaped = docTitle.replace(/['\\]/g, '\\$&');
                const driveSearchRes = await callApiWithRetry(() =>
                  drive.files.list({
                    q: `name = '${searchTitleEscaped}' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false`,
                    fields: 'files(id, name)',
                  })
                );

                if (driveSearchRes.data.files && driveSearchRes.data.files.length > 0) {
                  spId = driveSearchRes.data.files[0].id || '';
                  isExistingFile = true;
                  console.log(`[Drive Existing Found] Reusing existing file ID: ${spId} for '${docTitle}'`);
                }
              } catch (driveErr: any) {
                console.warn(`[Drive Search Warn] Could not search drive:`, driveErr?.message);
              }

              // 2. If no existing file found in Google Drive, create a new spreadsheet file
              if (!spId) {
                const createResponse = await callApiWithRetry(() =>
                  sheets.spreadsheets.create({
                    requestBody: {
                      properties: { title: docTitle },
                      sheets: [
                        { properties: { title: '관세법 개요', index: 0 } },
                        { properties: { title: '조문 목록', index: 1 } },
                      ],
                    },
                  })
                );
                spId = createResponse.data.spreadsheetId || '';
              } else {
                // Existing file exists: Ensure required worksheets '관세법 개요' and '조문 목록' exist
                try {
                  const meta = await callApiWithRetry(() => sheets.spreadsheets.get({ spreadsheetId: spId }));
                  const existingSheetTitles = (meta.data.sheets || []).map((s) => s.properties?.title || '');

                  const addSheetRequests: any[] = [];
                  if (!existingSheetTitles.includes('관세법 개요')) {
                    addSheetRequests.push({ addSheet: { properties: { title: '관세법 개요' } } });
                  }
                  if (!existingSheetTitles.includes('조문 목록')) {
                    addSheetRequests.push({ addSheet: { properties: { title: '조문 목록' } } });
                  }

                  if (addSheetRequests.length > 0) {
                    await callApiWithRetry(() =>
                      sheets.spreadsheets.batchUpdate({
                        spreadsheetId: spId,
                        requestBody: { requests: addSheetRequests },
                      })
                    );
                  }
                } catch (metaErr: any) {
                  console.warn(`[Worksheet Check Fail]`, metaErr?.message);
                }
              }

              if (!spId) return;

              const articleHeader = [
                '장 (Chapter)',
                '절 (Section)',
                '관 (Subsection)',
                '조문 번호 (조)',
                '조문 제목',
                '조문 내용 (전문)',
                '시행일자',
                '비고',
              ];

              const articleRows = revArticles.map((art: any) => [
                art.chapterName || '',
                art.sectionName || '',
                art.subsectionName || '',
                art.articleNo || '',
                art.articleTitle || '',
                art.articleContent || '',
                art.effectiveDate || rev.enforcementDate || '',
                art.isDeleted ? '삭제' : '',
              ]);

              const overviewValues = [
                ['국가법령정보포털 - 관세법 개정본 개별 DB'],
                [''],
                ['항목', '내용'],
                ['법령명', rev.lawName || '관세법'],
                ['공포번호', rev.promulgationNo],
                ['시행일자', rev.enforcementDate],
                ['공포일자', rev.promulgationDate],
                ['제개정구분', rev.revisionType || '일부개정'],
                ['소관부처', rev.department || '기획재정부'],
                ['법령ID / MST', rev.lawMst],
                ['해당 개정본 조문 수', `${revArticles.length}개 조문`],
                ['저장 일시', new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' })],
                ['국가법령정보포털 링크', `https://www.law.go.kr/법령/관세법`],
              ];

              await callApiWithRetry(() =>
                sheets.spreadsheets.values.batchUpdate({
                  spreadsheetId: spId,
                  requestBody: {
                    valueInputOption: 'USER_ENTERED',
                    data: [
                      { range: `'관세법 개요'!A1`, values: overviewValues },
                      { range: `'조문 목록'!A1`, values: [articleHeader, ...articleRows] },
                    ],
                  },
                })
              );

              createdFiles.push({
                title: docTitle,
                spreadsheetId: spId,
                url: `https://docs.google.com/spreadsheets/d/${spId}/edit`,
                promulgationNo: rev.promulgationNo,
                enforcementDate: rev.enforcementDate,
                isExisting: isExistingFile,
              });
            } catch (revErr: any) {
              console.warn(`[Batch Separate Export] Failed for MST ${rev.lawMst}:`, revErr?.message);
            }
          })
        );
        // Brief pacing delay between batches to respect Google write API limits
        await new Promise((resolve) => setTimeout(resolve, 500));
      }

      return res.json({
        success: true,
        message: `140개 관세법 개정자료 동기화 완료! (기존 파일 감지/업데이트 및 신규 파일/워크시트 생성, 총 ${createdFiles.length}개 확인)`,
        createdCount: createdFiles.length,
        spreadsheetUrl: createdFiles[0]?.url,
        createdFiles,
        exportMode: 'separate_files_140',
      });
    }

    let spreadsheetId = '';

    if (config?.targetType === 'existing') {
      const rawInput = (config.spreadsheetIdOrUrl || '').trim();

      if (!rawInput) {
        return res.status(400).json({ error: '기존 Google 스프레드시트 URL 또는 ID를 입력해 주세요.' });
      }

      if (rawInput.includes('/spreadsheets/u/') || rawInput.endsWith('/spreadsheets') || rawInput.endsWith('/spreadsheets/')) {
        return res.status(400).json({
          error: '입력하신 주소는 특정 문서의 주소가 아니라 Google 스프레드시트 메인 목록 페이지입니다. 특정 문서의 주소(예: https://docs.google.com/spreadsheets/d/문서ID/edit)를 입력하거나 "새 Google 스프레드시트 생성" 옵션을 선택해 주세요.',
        });
      }

      const urlMatch = rawInput.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]{25,})/);
      if (urlMatch && urlMatch[1]) {
        spreadsheetId = urlMatch[1];
      } else {
        const idMatch = rawInput.match(/[a-zA-Z0-9-_]{25,}/);
        if (idMatch) {
          spreadsheetId = idMatch[0];
        }
      }

      if (!spreadsheetId) {
        return res.status(400).json({
          error: '입력하신 주소에서 스프레드시트 ID를 추출할 수 없습니다. URL 형식(https://docs.google.com/spreadsheets/d/.../edit)을 확인해 주세요.',
        });
      }
    } else {
      // Create new Google Spreadsheet
      const docTitle = config?.exportAll140
        ? `[국가법령] 관세법 전체 140개 개정본 조문 DB (1949~2026)`
        : `[국가법령] ${lawData.info.lawName} 데이터 (${lawData.info.enforcementDate} 시행)`;

      const createResponse = await sheets.spreadsheets.create({
        requestBody: {
          properties: {
            title: docTitle,
          },
          sheets: [
            { properties: { title: '관세법 개정연혁 (140건)', index: 0 } },
            { properties: { title: '관세법 개요', index: 1 } },
            { properties: { title: '조문 목록', index: 2 } },
          ],
        },
      });

      spreadsheetId = createResponse.data.spreadsheetId || '';
    }

    if (!spreadsheetId) {
      throw new Error('Google Spreadsheet 생성 또는 ID 추출 실패');
    }

    // Fetch full 140+ revision history records if not sent in request body
    let revisionList: any[] = req.body.revisions || [];
    if (!Array.isArray(revisionList) || revisionList.length === 0) {
      try {
        revisionList = await fetchAll140Revisions(ocKey);
      } catch (revErr: any) {
        console.warn('Auto-fetching revisions error:', revErr?.message || revErr);
      }
    }

    // Inspect existing spreadsheet structure
    let existingSheetTitles: string[] = [];
    let historySheetName = `관세법 개정연혁 (${revisionList.length}건)`;
    let overviewSheetName = '관세법 개요';
    let articlesSheetName = '조문 목록';

    try {
      const meta = await sheets.spreadsheets.get({ spreadsheetId });
      const sheetsList = meta.data.sheets || [];
      existingSheetTitles = sheetsList.map((s) => s.properties?.title || '').filter(Boolean);

      const requestsToAdd: any[] = [];
      const hasHistoryTab = existingSheetTitles.some((t) => t.includes('개정연혁'));
      if (!hasHistoryTab) {
        requestsToAdd.push({ addSheet: { properties: { title: historySheetName } } });
      } else {
        const found = existingSheetTitles.find((t) => t.includes('개정연혁'));
        if (found) historySheetName = found;
      }
      if (config?.includeOverview !== false && !existingSheetTitles.includes('관세법 개요')) {
        requestsToAdd.push({ addSheet: { properties: { title: '관세법 개요' } } });
      }
      if (!existingSheetTitles.includes('조문 목록')) {
        requestsToAdd.push({ addSheet: { properties: { title: '조문 목록' } } });
      }

      if (requestsToAdd.length > 0) {
        await (sheets.spreadsheets as any).batchUpdate({
          spreadsheetId,
          requestBody: { requests: requestsToAdd },
        });
        existingSheetTitles.push(...requestsToAdd.map((r) => r.addSheet.properties.title));
      }
    } catch (tabErr: any) {
      console.warn('Sheet tab inspection/creation warning:', tabErr?.message || tabErr);
    }

    // Build Revision History Data (1949년 ~ 2026년 관세법 개정 이력 140건)
    const historyHeader = [
      '연번',
      '시행일자',
      '공포번호',
      '공포일자',
      '제개정구분',
      '법령명',
      '소관부처',
      '법령ID / MST',
    ];

    const historyRows = revisionList.map((rev: any, index: number) => [
      index + 1,
      rev.enforcementDate || '',
      rev.promulgationNo || '',
      rev.promulgationDate || '',
      rev.revisionType || '일부개정',
      rev.lawName || '관세법',
      rev.department || '기획재정부',
      rev.lawMst || rev.lawId || '',
    ]);

    const historyValues = [historyHeader, ...historyRows];

    // Build Articles Data according to config.exportAll140 mode
    let articleValues: any[][] = [];
    let totalProcessedArticlesCount = 0;

    if (config?.exportAll140) {
      console.log(`[Batch Export] Starting retrieval of all ${revisionList.length} revisions...`);
      const batchArticleRows: any[][] = [];
      const chunkSize = 15;

      for (let i = 0; i < revisionList.length; i += chunkSize) {
        const chunk = revisionList.slice(i, i + chunkSize);
        await Promise.all(
          chunk.map(async (rev: any) => {
            try {
              const detailUrl = `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(
                ocKey
              )}&target=law&MST=${encodeURIComponent(rev.lawMst)}&type=XML`;

              const detailRes = await fetch(detailUrl, {
                headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
              });

              if (detailRes.ok) {
                const detailXml = await detailRes.text();
                const parsed = xmlParser.parse(detailXml);
                const root = parsed.법령 || parsed.Law || parsed;
                const revArticles = parseArticlesFromXmlRoot(root);

                revArticles.forEach((art: any) => {
                  batchArticleRows.push([
                    rev.promulgationNo || '',
                    rev.enforcementDate || '',
                    rev.promulgationDate || '',
                    rev.revisionType || '일부개정',
                    art.chapterName || '',
                    art.sectionName || '',
                    art.subsectionName || '',
                    art.articleNo || '',
                    art.articleTitle || '',
                    art.articleContent || '',
                    art.isDeleted ? '삭제' : '',
                  ]);
                });
              }
            } catch (revFetchErr: any) {
              console.warn(`[Batch Export] Error for MST ${rev.lawMst}:`, revFetchErr?.message || revFetchErr);
            }
          })
        );
      }

      const batchArticleHeader = [
        '개정 공포번호',
        '시행일자',
        '공포일자',
        '개정구분',
        '장 (Chapter)',
        '절 (Section)',
        '관 (Subsection)',
        '조문 번호 (조)',
        '조문 제목',
        '조문 내용 (전문)',
        '비고',
      ];

      articleValues = [batchArticleHeader, ...batchArticleRows];
      totalProcessedArticlesCount = batchArticleRows.length;
      console.log(`[Batch Export] Finished collecting ${totalProcessedArticlesCount} articles across ${revisionList.length} revisions.`);
    } else {
      // Single revision export
      const articleHeader = [
        '장 (Chapter)',
        '절 (Section)',
        '관 (Subsection)',
        '조문 번호 (조)',
        '조문 제목',
        '조문 내용 (전문)',
        '시행일자',
        '비고',
      ];
      const articleRows = lawData.articles.map((art: any) => [
        art.chapterName || '',
        art.sectionName || '',
        art.subsectionName || '',
        art.articleNo || '',
        art.articleTitle || '',
        art.articleContent || '',
        art.effectiveDate || '',
        art.isDeleted ? '삭제' : '',
      ]);

      articleValues = [articleHeader, ...articleRows];
      totalProcessedArticlesCount = lawData.articles.length;
    }

    // Build Overview Data
    const overviewValues = [
      ['국가법령정보포털 - 관세법 수집 데이터 Summary'],
      [''],
      ['항목', '내용'],
      ['법령명', lawData.info.lawName],
      ['수집 범위', config?.exportAll140 ? '140개 전체 개정 관세법 일괄 수집' : '선택된 개정본 1건'],
      ['선택 개정본 공포번호', lawData.info.promulgationNo],
      ['선택 개정본 시행일자', lawData.info.enforcementDate],
      ['선택 개정본 공포일자', lawData.info.promulgationDate],
      ['제개정구분', lawData.info.revisionType || '일부개정'],
      ['소관부처', lawData.info.department || '기획재정부'],
      ['법령ID / MST', lawData.info.lawMst],
      ['수집된 개정 이력 건수', `${revisionList.length}건 (1949년~2026년)`],
      ['조문 목록 시트 기록 조문 수', `${totalProcessedArticlesCount}개 조문`],
      ['저장 일시', new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' })],
      ['국가법령정보포털 링크', `https://www.law.go.kr/법령/관세법`],
    ];

    // Update Revision History Tab (140건 개정연혁)
    try {
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `'${historySheetName}'!A1`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: historyValues },
      });
    } catch (histErr: any) {
      console.warn('History sheet update warning:', histErr?.message);
    }

    // Update Overview Tab
    if (config?.includeOverview !== false) {
      try {
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: `'${overviewSheetName}'!A1`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: overviewValues },
        });
      } catch (e: any) {
        console.warn('Overview sheet update warning:', e?.message);
      }
    }

    // Update Articles Tab in chunks of 5000 rows to prevent payload limit issues
    try {
      const ROW_CHUNK_SIZE = 5000;
      for (let i = 0; i < articleValues.length; i += ROW_CHUNK_SIZE) {
        const chunkValues = articleValues.slice(i, i + ROW_CHUNK_SIZE);
        const startRow = i + 1;
        const range = `'${articlesSheetName}'!A${startRow}`;

        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: chunkValues },
        });
      }
    } catch (artErr: any) {
      console.warn('Error updating articles sheet chunked:', artErr?.message);
      // Fallback update
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: 'A1',
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: articleValues.slice(0, 5000) },
      });
    }

    // Format header styling using batchUpdate
    if (config?.autoFormat !== false) {
      try {
        const getSpreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
        const sheetsList = getSpreadsheet.data.sheets || [];

        const historySheet = sheetsList.find((s) => s.properties?.title === historySheetName);
        const articlesSheet = sheetsList.find((s) => s.properties?.title === articlesSheetName) || sheetsList[0];

        const historySheetId = historySheet?.properties?.sheetId || 0;
        const articlesSheetId = articlesSheet?.properties?.sheetId || 0;

        const requests: any[] = [];

        // Freeze top row & format header for Revision History sheet
        if (historySheet) {
          requests.push(
            {
              updateSheetProperties: {
                properties: {
                  sheetId: historySheetId,
                  gridProperties: { frozenRowCount: 1 },
                },
                fields: 'gridProperties.frozenRowCount',
              },
            },
            {
              repeatCell: {
                range: {
                  sheetId: historySheetId,
                  startRowIndex: 0,
                  endRowIndex: 1,
                  startColumnIndex: 0,
                  endColumnIndex: 8,
                },
                cell: {
                  userEnteredFormat: {
                    backgroundColor: { red: 0.18, green: 0.22, blue: 0.35 },
                    textFormat: {
                      foregroundColor: { red: 1.0, green: 1.0, blue: 1.0 },
                      bold: true,
                      fontSize: 11,
                    },
                    alignment: { horizontal: 'CENTER', vertical: 'MIDDLE' },
                  },
                },
                fields: 'userEnteredFormat(backgroundColor,textFormat,alignment)',
              },
            }
          );
        }

        // Freeze top row & format header for Articles sheet
        requests.push(
          {
            updateSheetProperties: {
              properties: {
                sheetId: articlesSheetId,
                gridProperties: { frozenRowCount: 1 },
              },
              fields: 'gridProperties.frozenRowCount',
            },
          },
          {
            repeatCell: {
              range: {
                sheetId: articlesSheetId,
                startRowIndex: 0,
                endRowIndex: 1,
                startColumnIndex: 0,
                endColumnIndex: 8,
              },
              cell: {
                userEnteredFormat: {
                  backgroundColor: { red: 0.12, green: 0.16, blue: 0.23 },
                  textFormat: {
                    foregroundColor: { red: 1.0, green: 1.0, blue: 1.0 },
                    bold: true,
                    fontSize: 11,
                  },
                  alignment: { horizontal: 'CENTER', vertical: 'MIDDLE' },
                },
              },
              fields: 'userEnteredFormat(backgroundColor,textFormat,alignment)',
            },
          },
          {
            repeatCell: {
              range: {
                sheetId: articlesSheetId,
                startRowIndex: 1,
                startColumnIndex: 5,
                endColumnIndex: 6,
              },
              cell: {
                userEnteredFormat: {
                  wrapStrategy: 'WRAP',
                  alignment: { vertical: 'TOP' },
                },
              },
              fields: 'userEnteredFormat(wrapStrategy,alignment)',
            },
          }
        );

        await (sheets.spreadsheets as any).batchUpdate({
          spreadsheetId,
          requestBody: { requests },
        });
      } catch (formatErr) {
        console.warn('Sheet formatting warning (non-critical):', formatErr);
      }
    }

    const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;

    return res.json({
      success: true,
      spreadsheetId,
      spreadsheetUrl,
      message: '관세법 조문 전체가 Google Sheets에 정상적으로 저장되었습니다.',
    });
  } catch (error: any) {
    console.error('Save to Sheets Error:', error);

    let friendlyMessage = error.message || 'Google Sheets 저장 중 오류가 발생했습니다.';

    if (error.code === 403 || error.status === 403) {
      friendlyMessage = '해당 Google 스프레드시트에 대한 수정(편집) 권한이 없습니다. 문서 공유 설정에서 편집자로 권한이 부여되어 있는지 확인해 주세요.';
    } else if (error.code === 404 || error.status === 404) {
      friendlyMessage = '입력하신 Google 스프레드시트를 찾을 수 없습니다. 문서 URL 및 삭제 여부를 확인해 주세요.';
    } else if (error.code === 401 || error.status === 401) {
      friendlyMessage = 'Google 계정 인증 토큰이 만료되었거나 권한이 필요합니다. 상단의 "Google 계정 연결" 버튼을 눌러 다시 로그인해 주세요.';
    }

    return res.status(error.status || 500).json({
      error: friendlyMessage,
    });
  }
});

// API Route: Get top 2 recent revisions of Customs Act with full article details
app.get('/api/law/recent-2-revisions', async (req, res) => {
  try {
    const ocKey = (req.query.ocKey as string) || DEFAULT_OC_KEY;
    const allRevisions = await fetchAll140Revisions(ocKey);
    const top2Revisions = allRevisions.slice(0, 2);

    if (top2Revisions.length === 0) {
      return res.status(404).json({ error: '관세법 개정 이력을 가져올 수 없습니다.' });
    }

    const detailedTop2 = await Promise.all(
      top2Revisions.map(async (rev, index) => {
        try {
          const detailUrl = `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(
            ocKey
          )}&target=law&MST=${encodeURIComponent(rev.lawMst)}&type=XML`;

          const detailRes = await fetch(detailUrl, {
            headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
          });

          if (!detailRes.ok) {
            return {
              ...rev,
              rank: index + 1,
              articles: [],
              articleCount: 0,
            };
          }

          const detailXml = await detailRes.text();
          const parsed = xmlParser.parse(detailXml);
          const root = parsed.법령 || parsed.Law || parsed;
          const articles = parseArticlesFromXmlRoot(root);

          return {
            ...rev,
            rank: index + 1,
            articles,
            articleCount: articles.length,
          };
        } catch (err: any) {
          console.warn(`[Recent 2 Top Detail Error] for MST ${rev.lawMst}:`, err?.message);
          return {
            ...rev,
            rank: index + 1,
            articles: [],
            articleCount: 0,
          };
        }
      })
    );

    return res.json({
      success: true,
      count: detailedTop2.length,
      revisions: detailedTop2,
    });
  } catch (error: any) {
    console.error('Error fetching recent 2 revisions:', error);
    return res.status(500).json({ error: error.message || '최근 개정본 조회 실패' });
  }
});

// API Route: Save top 2 recent revisions to Google Sheets as a test
app.post('/api/sheets/save-recent-2-test', async (req, res) => {
  try {
    const { accessToken } = req.body;
    const ocKey = req.body.ocKey || DEFAULT_OC_KEY;

    if (!accessToken) {
      return res.status(401).json({
        error: 'Google OAuth Access Token이 필요합니다. 상단의 Google 계정 연결 버튼을 눌러주세요.',
      });
    }

    const allRevisions = await fetchAll140Revisions(ocKey);
    const top2Revisions = allRevisions.slice(0, 2);

    if (top2Revisions.length === 0) {
      return res.status(404).json({ error: '관세법 개정 이력을 가져오지 못했습니다.' });
    }

    const auth = new google.auth.OAuth2();
    auth.setCredentials({ access_token: accessToken });
    const sheets = google.sheets({ version: 'v4', auth });

    // Fetch articles for both revisions
    const detailedItems: any[] = [];
    for (const rev of top2Revisions) {
      const detailUrl = `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(
        ocKey
      )}&target=law&MST=${encodeURIComponent(rev.lawMst)}&type=XML`;

      const detailRes = await fetch(detailUrl, {
        headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
      });

      let articles: any[] = [];
      if (detailRes.ok) {
        const detailXml = await detailRes.text();
        const parsed = xmlParser.parse(detailXml);
        const root = parsed.법령 || parsed.Law || parsed;
        articles = parseArticlesFromXmlRoot(root);
      }

      detailedItems.push({
        ...rev,
        articles,
      });
    }

    const rev1 = detailedItems[0];
    const rev2 = detailedItems[1];

    const docTitle = `[관세법 테스트] 최근 개정본 2개 조문 저장 (${rev1.promulgationNo} & ${rev2?.promulgationNo || ''})`;

    const sheet1Title = '최근 2개 개정본 요약';
    const sheet2Title = `1위_${(rev1.promulgationNo || '최신본').replace(/[\/\\?%*:|"<>]/g, '_')}`.slice(0, 30);
    const sheet3Title = rev2 ? `2위_${(rev2.promulgationNo || '직전본').replace(/[\/\\?%*:|"<>]/g, '_')}`.slice(0, 30) : '2위_개정본';

    const createResponse = await sheets.spreadsheets.create({
      requestBody: {
        properties: { title: docTitle },
        sheets: [
          { properties: { title: sheet1Title, index: 0 } },
          { properties: { title: sheet2Title, index: 1 } },
          { properties: { title: sheet3Title, index: 2 } },
        ],
      },
    });

    const spreadsheetId = createResponse.data.spreadsheetId;
    if (!spreadsheetId) {
      throw new Error('Google Spreadsheet 생성 실패');
    }

    // Build Summary Sheet
    const summaryHeader = ['구분', '순위', '법령명', '공포번호', '시행일자', '공포일자', '제개정구분', '소관부처', '조문 수', 'MST'];
    const summaryRows = detailedItems.map((item, idx) => [
      idx === 0 ? '최신 개정본 (1위)' : '직전 개정본 (2위)',
      idx + 1,
      item.lawName || '관세법',
      item.promulgationNo || '',
      item.enforcementDate || '',
      item.promulgationDate || '',
      item.revisionType || '',
      item.department || '기획재정부',
      `${item.articles.length}개 조문`,
      item.lawMst || '',
    ]);

    const articleHeader = [
      '장 (Chapter)',
      '절 (Section)',
      '관 (Subsection)',
      '조문 번호',
      '조문 제목',
      '조문 내용 (전문)',
      '시행일자',
      '비고',
    ];

    const rev1Rows = rev1.articles.map((art: any) => [
      art.chapterName || '',
      art.sectionName || '',
      art.subsectionName || '',
      art.articleNo || '',
      art.articleTitle || '',
      art.articleContent || '',
      art.effectiveDate || rev1.enforcementDate || '',
      art.isDeleted ? '삭제' : '',
    ]);

    const rev2Rows = rev2 ? rev2.articles.map((art: any) => [
      art.chapterName || '',
      art.sectionName || '',
      art.subsectionName || '',
      art.articleNo || '',
      art.articleTitle || '',
      art.articleContent || '',
      art.effectiveDate || rev2.enforcementDate || '',
      art.isDeleted ? '삭제' : '',
    ]) : [];

    // Write all values
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: {
        valueInputOption: 'USER_ENTERED',
        data: [
          {
            range: `'${sheet1Title}'!A1`,
            values: [
              ['[테스트 저장] 관세법 최근 개정본 2개 데이터 요약'],
              ['저장 일시', new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' })],
              [''],
              summaryHeader,
              ...summaryRows,
            ],
          },
          {
            range: `'${sheet2Title}'!A1`,
            values: [articleHeader, ...rev1Rows],
          },
          ...(rev2 ? [{
            range: `'${sheet3Title}'!A1`,
            values: [articleHeader, ...rev2Rows],
          }] : []),
        ],
      },
    });

    const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;

    return res.json({
      success: true,
      spreadsheetId,
      spreadsheetUrl,
      items: detailedItems.map((d) => ({
        promulgationNo: d.promulgationNo,
        enforcementDate: d.enforcementDate,
        articleCount: d.articles.length,
      })),
      message: `최근 관세법 개정본 2개(${rev1.promulgationNo}, ${rev2?.promulgationNo}) 조문 전체가 Google Sheets에 성공적으로 저장되었습니다!`,
    });
  } catch (error: any) {
    console.error('Recent 2 Test Save Error:', error);
    return res.status(error.status || 500).json({
      error: error.message || '최근 개정본 2개 테스트 저장 중 오류가 발생했습니다.',
    });
  }
});

async function getOrCreateUniqueDriveFolder(
  drive: any,
  baseFolderName: string,
  callWithRetry: <T>(fn: () => Promise<T>) => Promise<T>
): Promise<{ folderId: string; folderUrl: string; folderName: string; folderSkipped: boolean }> {
  let uniqueFolderName = baseFolderName.trim();
  let folderSkipped = false;

  try {
    const escaped = baseFolderName.replace(/'/g, "\\'");
    const folderSearchRes: any = await callWithRetry(() =>
      drive.files.list({
        q: `name contains '${escaped}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false`,
        fields: 'files(id, name)',
        pageSize: 100,
      })
    );

    const existingNames = new Set<string>(
      (folderSearchRes?.data?.files || []).map((f: any) => f.name?.trim()).filter(Boolean)
    );

    if (existingNames.has(uniqueFolderName)) {
      let count = 1;
      while (existingNames.has(`${uniqueFolderName}(${count})`)) {
        count++;
      }
      uniqueFolderName = `${uniqueFolderName}(${count})`;
      console.log(`[Drive Folder Duplicate Found] Created unique indexed folder: '${uniqueFolderName}'`);
    }
  } catch (searchErr: any) {
    console.warn('[Drive Folder List Warning]', searchErr?.message);
  }

  const folderCreateRes: any = await callWithRetry(() =>
    drive.files.create({
      requestBody: {
        name: uniqueFolderName,
        mimeType: 'application/vnd.google-apps.folder',
      },
      fields: 'id, name, webViewLink',
    })
  );

  const folderId = folderCreateRes?.data?.id || '';
  const folderUrl = folderCreateRes?.data?.webViewLink || `https://drive.google.com/drive/folders/${folderId}`;

  if (!folderId) {
    throw new Error('Google Drive 폴더 생성에 실패했습니다. 계정 권한을 확인해 주세요.');
  }

  return { folderId, folderUrl, folderName: uniqueFolderName, folderSkipped };
}

// API Route: Save all revisions into a Google Drive folder named "(법령명)+(날짜)"
// Supports mode: 'single_file' (1 Google Sheet with all revisions & articles) or 'separate_files' (1 Google Sheet per revision)
// Supports lawCategory: 'law' (관세법 등 법률) or 'admrul' (외국환거래규정 등 행정규칙/고시)
app.post('/api/drive/export-all-revisions-folder', async (req, res) => {
  try {
    const {
      accessToken,
      mode = 'single_file',
      lawName = '관세법',
      lawCategory = 'law',
      limitCount = 0,
    } = req.body;
    const ocKey = req.body.ocKey || DEFAULT_OC_KEY;

    if (!accessToken) {
      return res.status(401).json({
        error: 'Google OAuth Access Token이 필요합니다. 상단의 Google 계정 연결 버튼을 눌러주세요.',
      });
    }

    const auth = new google.auth.OAuth2();
    auth.setCredentials({ access_token: accessToken });
    const drive = google.drive({ version: 'v3', auth });
    const sheets = google.sheets({ version: 'v4', auth });

    // 1. Determine Folder Name: (법령명)+(날짜) e.g., "관세법_2026-08-17" or "외국환거래규정_2026-08-17"
    const now = new Date();
    const yyyy = now.getFullYear();
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const dd = String(now.getDate()).padStart(2, '0');
    const formattedDate = `${yyyy}-${mm}-${dd}`;
    const cleanLawName = (lawName || (lawCategory === 'admrul' ? '외국환거래규정' : '관세법')).trim();
    const defaultFolderName = `${cleanLawName}_${formattedDate}`;
    const requestedFolderName = req.body.folderName?.trim() || defaultFolderName;

    // Helper for API retry with intelligent backoff for Google Sheets/Drive write quotas
    const callWithRetry = async <T>(fn: () => Promise<T>, retries = 7, delay = 2000): Promise<T> => {
      let attempt = 0;
      let currentDelay = delay;

      while (true) {
        try {
          return await fn();
        } catch (err: any) {
          attempt++;
          const errMsg = String(err?.message || '').toLowerCase();
          const errCode = String(err?.code || '').toLowerCase();
          const status = typeof err?.status === 'number' ? err.status : (typeof err?.code === 'number' ? err.code : Number(err?.status || err?.code || 0));
          const isRateLimit =
            status === 429 ||
            errMsg.includes('quota') ||
            errMsg.includes('rate') ||
            errMsg.includes('resource_exhausted') ||
            errMsg.includes('write requests per minute') ||
            errMsg.includes('user rate limit');

          if ((isRateLimit || status === 500 || status === 503) && attempt <= retries) {
            const backoffTime = isRateLimit
              ? Math.max(currentDelay, Math.min(65000, 6000 * Math.pow(1.6, attempt - 1)))
              : currentDelay;
            console.warn(
              `[Google API Backoff] Attempt ${attempt}/${retries} failed (${isRateLimit ? '429 Quota Exceeded' : err?.message}). Waiting ${Math.round(backoffTime)}ms for quota window...`
            );
            await new Promise((r) => setTimeout(r, backoffTime + Math.random() * 800));
            currentDelay = currentDelay * 2;
            continue;
          }
          throw err;
        }
      }
    };

    // 2. Create Unique Folder in Google Drive with automatic (1), (2) index if name exists
    const { folderId, folderUrl, folderName: targetFolderName, folderSkipped } = await getOrCreateUniqueDriveFolder(
      drive,
      requestedFolderName,
      callWithRetry
    );

    console.log(`[Drive Folder Export] Initialized: '${targetFolderName}' for mode: '${mode}', category: '${lawCategory}', cleanLawName: '${cleanLawName}' (ID: ${folderId})`);

    if (!folderId) {
      throw new Error('Google Drive 폴더 생성에 실패했습니다. Google 계정 권한을 확인해 주세요.');
    }

    // 3. Fetch revisions (Customs Act 140, Foreign Exchange Act 45, or Administrative Rules)
    let revisionList: any[] = req.body.revisions || [];
    if (!Array.isArray(revisionList) || revisionList.length === 0) {
      if (lawCategory === 'admrul' || cleanLawName === '외국환거래규정') {
        const effectiveLimit = limitCount > 0 ? limitCount : 0;
        revisionList = await fetchAdmrulRevisions(ocKey, cleanLawName, effectiveLimit);
      } else {
        revisionList = await fetchLawRevisions(ocKey, cleanLawName, limitCount > 0 ? limitCount : 0);
      }
    }

    if (revisionList.length === 0) {
      return res.status(404).json({ error: `${cleanLawName}의 개정연혁 목록을 불러올 수 없습니다.` });
    }

    revisionList = sortRevisionsByEnforcementDateDesc(revisionList);

    // ========================================================
    // MODE A: 'single_file' -> 구글시트 1개에 모든 개정연혁 및 조문 통합 저장
    // ========================================================
    if (mode === 'single_file') {
      const docTitle = `[${cleanLawName}] ${revisionList.length}개 개정연혁 통합본 (${formattedDate})`;

      // Check if file already exists in the target folder (Skip duplicate sheet)
      const existingDocRes = await callWithRetry(() =>
        drive.files.list({
          q: `'${folderId}' in parents and name = '${docTitle.replace(/'/g, "\\'")}' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false`,
          fields: 'files(id, name, webViewLink)',
          spaces: 'drive',
        })
      );

      if (existingDocRes.data.files && existingDocRes.data.files.length > 0) {
        const existingFile = existingDocRes.data.files[0];
        console.log(`[Single File Exists - Skipped] ${docTitle} (ID: ${existingFile.id})`);
        return res.json({
          success: true,
          mode: 'single_file',
          folderId,
          folderUrl,
          folderName: targetFolderName,
          folderSkipped,
          skipped: true,
          spreadsheetId: existingFile.id,
          spreadsheetUrl: existingFile.webViewLink || `https://docs.google.com/spreadsheets/d/${existingFile.id}/edit`,
          totalRevisions: revisionList.length,
          message: `동일한 구글시트('[${docTitle}]')가 이미 '${targetFolderName}' 폴더에 존재하여 생성을 건너뛰었습니다. (스킵됨)`,
        });
      }

      // Create new Google Spreadsheet
      const createRes = await callWithRetry(() =>
        drive.files.create({
          requestBody: {
            name: docTitle,
            mimeType: 'application/vnd.google-apps.spreadsheet',
            parents: [folderId],
          },
          fields: 'id, name, webViewLink',
        })
      );

      const spreadsheetId = createRes.data.id || '';
      const spreadsheetUrl = createRes.data.webViewLink || `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;

      if (!spreadsheetId) {
        throw new Error('Google Spreadsheet 생성 실패');
      }

      // Initialize tabs (Rename default sheetId 0 to remove '시트1', Add Sheet 2)
      const sheet1Title = `${cleanLawName} 개정연혁 목록 (${revisionList.length}건)`;
      const sheet2Title = `${cleanLawName} 전체 조문 통합데이터`;

      await callWithRetry(() =>
        sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: {
            requests: [
              {
                updateSheetProperties: {
                  properties: { sheetId: 0, title: sheet1Title },
                  fields: 'title',
                },
              },
              {
                addSheet: {
                  properties: { sheetId: 1, title: sheet2Title, index: 1 },
                },
              },
            ],
          },
        })
      );

      // 1) Build Revision History Sheet (Summary)
      const historyHeader = [
        '연번',
        '시행일자',
        '공포/발령번호',
        '공포/발령일자',
        '제개정구분',
        '법령/행정규칙명',
        '소관부처',
        '일련번호 / MST',
      ];

      const historyRows = revisionList.map((rev, index) => [
        index + 1,
        rev.enforcementDate || '',
        rev.promulgationNo || '',
        rev.promulgationDate || '',
        rev.revisionType || '일부개정',
        rev.lawName || cleanLawName,
        rev.department || '기획재정부',
        rev.lawMst || rev.seq || rev.id || '',
      ]);

      // 2) Collect all articles from revisions
      console.log(`[Single File Export] Fetching articles for ${revisionList.length} revisions (${cleanLawName})...`);
      const allArticleRows: any[][] = [];
      const chunkSize = 10;

      for (let i = 0; i < revisionList.length; i += chunkSize) {
        const chunk = revisionList.slice(i, i + chunkSize);
        await Promise.all(
          chunk.map(async (rev: any, chunkIndex: number) => {
            try {
              const isAdmrul = lawCategory === 'admrul' || rev.targetType === 'admrul' || cleanLawName === '외국환거래규정' || cleanLawName.includes('고시') || cleanLawName.includes('규정');
              const revArticles = await fetchArticlesForRevision(ocKey, rev, isAdmrul ? 'admrul' : 'law');

              revArticles.forEach((art: any) => {
                allArticleRows.push([
                  i + chunkIndex + 1,
                  rev.promulgationNo || '',
                  rev.enforcementDate || '',
                  rev.promulgationDate || '',
                  rev.revisionType || '일부개정',
                  art.chapterName || '',
                  art.sectionName || '',
                  art.subsectionName || '',
                  art.articleNo || '',
                  art.articleTitle || '',
                  art.articleContent || '',
                  art.isDeleted ? '삭제' : '',
                ]);
              });

              if (rev.buchikText && !revArticles.some((a: any) => a.articleTitle?.includes('부칙') || a.articleNo?.includes('부칙'))) {
                allArticleRows.push([
                  i + chunkIndex + 1,
                  rev.promulgationNo || '',
                  rev.enforcementDate || '',
                  rev.promulgationDate || '',
                  rev.revisionType || '일부개정',
                  '부칙',
                  '',
                  '',
                  '부칙',
                  `부칙 (${rev.promulgationNo || ''})`,
                  rev.buchikText,
                  '',
                ]);
              }
            } catch (fetchErr: any) {
              console.warn(`[Single File Export] Error fetching MST/ID ${rev.lawMst || rev.id}:`, fetchErr?.message);
            }
          })
        );
      }

      const allArticleHeader = [
        '개정 연번',
        '공포/발령번호',
        '시행일자',
        '공포/발령일자',
        '제개정구분',
        '장 (Chapter)',
        '절 (Section)',
        '관 (Subsection)',
        '조문 번호',
        '조문 제목',
        '조문 내용 (전문)',
        '비고',
      ];

      // 3) Write Data to Sheets
      await callWithRetry(() =>
        sheets.spreadsheets.values.batchUpdate({
          spreadsheetId,
          requestBody: {
            valueInputOption: 'USER_ENTERED',
            data: [
              {
                range: `'${sheet1Title}'!A1`,
                values: [
                  [`[국가법령/행정규칙] ${cleanLawName} 전체 ${revisionList.length}개 개정연혁 목록`],
                  [`저장 폴더: ${targetFolderName}`, `저장 일시: ${new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' })}`],
                  [''],
                  historyHeader,
                  ...historyRows,
                ],
              },
              {
                range: `'${sheet2Title}'!A1`,
                values: [allArticleHeader, ...allArticleRows],
              },
            ],
          },
        })
      );

      // 4) Apply cell formatting: Vertical Alignment to TOP, Wrap Text, Header Row Styling & Freeze
      await callWithRetry(() =>
        sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: {
            requests: [
              // Sheet 1: Whole sheet top vertical alignment & text wrap
              {
                repeatCell: {
                  range: { sheetId: 0, startRowIndex: 0 },
                  cell: {
                    userEnteredFormat: {
                      verticalAlignment: 'TOP',
                      wrapStrategy: 'WRAP',
                    },
                  },
                  fields: 'userEnteredFormat(verticalAlignment,wrapStrategy)',
                },
              },
              // Sheet 1: Header row (row index 3) styling (Navy background, bold white, center aligned)
              {
                repeatCell: {
                  range: { sheetId: 0, startRowIndex: 3, endRowIndex: 4 },
                  cell: {
                    userEnteredFormat: {
                      backgroundColor: { red: 0.12, green: 0.18, blue: 0.3 },
                      textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 }, fontSize: 10 },
                      verticalAlignment: 'MIDDLE',
                      horizontalAlignment: 'CENTER',
                    },
                  },
                  fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
                },
              },
              // Sheet 1: Freeze header rows
              {
                updateSheetProperties: {
                  properties: { sheetId: 0, gridProperties: { frozenRowCount: 4 } },
                  fields: 'gridProperties.frozenRowCount',
                },
              },
              // Sheet 2: Whole sheet top vertical alignment & text wrap
              {
                repeatCell: {
                  range: { sheetId: 1, startRowIndex: 0 },
                  cell: {
                    userEnteredFormat: {
                      verticalAlignment: 'TOP',
                      wrapStrategy: 'WRAP',
                    },
                  },
                  fields: 'userEnteredFormat(verticalAlignment,wrapStrategy)',
                },
              },
              // Sheet 2: Header row (row index 0) styling (Navy background, bold white, center aligned)
              {
                repeatCell: {
                  range: { sheetId: 1, startRowIndex: 0, endRowIndex: 1 },
                  cell: {
                    userEnteredFormat: {
                      backgroundColor: { red: 0.12, green: 0.18, blue: 0.3 },
                      textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 }, fontSize: 10 },
                      verticalAlignment: 'MIDDLE',
                      horizontalAlignment: 'CENTER',
                    },
                  },
                  fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
                },
              },
              // Sheet 2: Freeze header row
              {
                updateSheetProperties: {
                  properties: { sheetId: 1, gridProperties: { frozenRowCount: 1 } },
                  fields: 'gridProperties.frozenRowCount',
                },
              },
            ],
          },
        })
      );

      return res.json({
        success: true,
        mode: 'single_file',
        folderId,
        folderUrl,
        folderName: targetFolderName,
        folderSkipped,
        skipped: false,
        spreadsheetId,
        spreadsheetUrl,
        totalRevisions: revisionList.length,
        totalArticles: allArticleRows.length,
        message: `Google Drive '${targetFolderName}' 폴더에 1개의 통합 구글 스프레드시트가 성공적으로 저장되었습니다! (셀 행 위로 정렬 적용 완료, 총 ${revisionList.length}개 개정판, ${allArticleRows.length}개 조문)`,
      });
    }

    // ========================================================
    // MODE B: 'separate_files' -> 개정연혁 1개 파일로 각각 저장 (개별 구글시트 파일 생성)
    // ========================================================
    if (mode === 'separate_files') {
      console.log(`[Separate Files Export] Starting creation of individual sheets in folder '${targetFolderName}' (${cleanLawName})...`);
      const createdFiles: Array<{
        title: string;
        spreadsheetId: string;
        url: string;
        promulgationNo: string;
        enforcementDate: string;
        skipped?: boolean;
      }> = [];

      // Check existing files in folder to avoid duplicates or overwrite
      const existingFilesRes = await callWithRetry(() =>
        drive.files.list({
          q: `'${folderId}' in parents and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false`,
          fields: 'files(id, name, webViewLink)',
          spaces: 'drive',
        })
      );
      const existingFileMap = new Map<string, { id: string; url: string }>();
      (existingFilesRes.data.files || []).forEach((f) => {
        if (f.name && f.id) existingFileMap.set(f.name, { id: f.id, url: f.webViewLink || `https://docs.google.com/spreadsheets/d/${f.id}/edit` });
      });

      const articleHeader = [
        '장 (Chapter)',
        '절 (Section)',
        '관 (Subsection)',
        '조문 번호',
        '조문 제목',
        '조문 내용 (전문)',
        '시행일자',
        '비고',
      ];

      let skippedCount = 0;
      let newCreatedCount = 0;

      // Process revisions sequentially with controlled pacing to strictly respect Google Sheets write quotas
      for (let i = 0; i < revisionList.length; i++) {
        const rev = revisionList[i];
        const revIndexNum = String(i + 1).padStart(3, '0');
        const revDocName = (rev.name || rev.lawName || cleanLawName).trim();
        const safePromNo = (rev.promulgationNo || '개정본').trim();
        const safeEnfDate = (rev.enforcementDate || '시행일 미상').trim();
        const docTitle = `${revIndexNum}_[${revDocName}] ${safePromNo} (${safeEnfDate} 시행)`.replace(/[\/\\:*?"<>|]/g, '_');

        // Check if file already exists in folder (Skip duplicate sheet)
        if (existingFileMap.has(docTitle)) {
          const ex = existingFileMap.get(docTitle)!;
          skippedCount++;
          createdFiles.push({
            title: docTitle,
            spreadsheetId: ex.id,
            url: ex.url,
            promulgationNo: rev.promulgationNo,
            enforcementDate: rev.enforcementDate,
            skipped: true,
          });
          console.log(`[Separate Export Exists - Skipped] ${docTitle}`);
          continue;
        }

        try {
          // 1. Fetch articles with multi-tiered HTML + XML engine
          const isAdmrul =
            lawCategory === 'admrul' ||
            rev.targetType === 'admrul' ||
            cleanLawName === '외국환거래규정' ||
            revDocName.includes('고시') ||
            revDocName.includes('규정') ||
            revDocName.includes('훈령') ||
            revDocName.includes('예규');

          let revArticles = await fetchArticlesForRevision(ocKey, rev, isAdmrul ? 'admrul' : 'law');

          if (rev.buchikText && !revArticles.some((a: any) => a.articleTitle?.includes('부칙') || a.articleNo?.includes('부칙'))) {
            revArticles.push({
              chapterName: '부칙',
              sectionName: '',
              subsectionName: '',
              articleNo: '부칙',
              articleTitle: `${safePromNo} 부칙`,
              articleContent: rev.buchikText,
              effectiveDate: safeEnfDate,
            });
          }

          // 2. Create Spreadsheet in Drive Folder
          const createRes = await callWithRetry(() =>
            drive.files.create({
              requestBody: {
                name: docTitle,
                mimeType: 'application/vnd.google-apps.spreadsheet',
                parents: [folderId],
              },
              fields: 'id, name, webViewLink',
            })
          );
          const spId = createRes.data.id || '';
          if (!spId) continue;

          const overviewSheet = `${revDocName.replace(/[\/\\:*?\[\]]/g, '_').slice(0, 20)} 개요`;
          const articlesSheet = '조문 목록';

          // Combined Single BatchUpdate: Add/Rename sheets and apply full formatting
          await callWithRetry(() =>
            sheets.spreadsheets.batchUpdate({
              spreadsheetId: spId,
              requestBody: {
                requests: [
                  {
                    updateSheetProperties: {
                      properties: { sheetId: 0, title: overviewSheet },
                      fields: 'title',
                    },
                  },
                  {
                    addSheet: {
                      properties: { sheetId: 1, title: articlesSheet, index: 1 },
                    },
                  },
                  {
                    repeatCell: {
                      range: { sheetId: 0, startRowIndex: 0 },
                      cell: {
                        userEnteredFormat: {
                          verticalAlignment: 'TOP',
                          wrapStrategy: 'WRAP',
                        },
                      },
                      fields: 'userEnteredFormat(verticalAlignment,wrapStrategy)',
                    },
                  },
                  {
                    repeatCell: {
                      range: { sheetId: 0, startRowIndex: 2, endRowIndex: 3 },
                      cell: {
                        userEnteredFormat: {
                          backgroundColor: { red: 0.12, green: 0.18, blue: 0.3 },
                          textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 }, fontSize: 10 },
                          verticalAlignment: 'MIDDLE',
                          horizontalAlignment: 'CENTER',
                        },
                      },
                      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
                    },
                  },
                  {
                    repeatCell: {
                      range: { sheetId: 1, startRowIndex: 0 },
                      cell: {
                        userEnteredFormat: {
                          verticalAlignment: 'TOP',
                          wrapStrategy: 'WRAP',
                        },
                      },
                      fields: 'userEnteredFormat(verticalAlignment,wrapStrategy)',
                    },
                  },
                  {
                    repeatCell: {
                      range: { sheetId: 1, startRowIndex: 0, endRowIndex: 1 },
                      cell: {
                        userEnteredFormat: {
                          backgroundColor: { red: 0.12, green: 0.18, blue: 0.3 },
                          textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 }, fontSize: 10 },
                          verticalAlignment: 'MIDDLE',
                          horizontalAlignment: 'CENTER',
                        },
                      },
                      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
                    },
                  },
                  {
                    updateSheetProperties: {
                      properties: { sheetId: 1, gridProperties: { frozenRowCount: 1 } },
                      fields: 'gridProperties.frozenRowCount',
                    },
                  },
                ],
              },
            })
          );

          // Build rows
          const overviewValues = [
            [`대한민국 ${cleanLawName} 개정본`],
            [''],
            ['항목', '내용'],
            ['법령/행정규칙명', rev.lawName || cleanLawName],
            ['공포/발령번호', rev.promulgationNo || '-'],
            ['시행일자', rev.enforcementDate || '-'],
            ['공포/발령일자', rev.promulgationDate || '-'],
            ['제개정구분', rev.revisionType || '일부개정'],
            ['소관부처', rev.department || (isAdmrul ? '재정경제부' : '기획재정부')],
            ['일련번호 / MST', rev.lawMst || rev.seq || rev.id || ''],
            ['해당 개정본 조문 수', `${revArticles.length}개 조문`],
            ['개정 부칙 (공포내용)', rev.buchikText || '-'],
            ['저장 폴더', targetFolderName],
            ['저장 일시', new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' })],
          ];

          const articleRows = revArticles.map((art: any) => [
            art.chapterName || '본문',
            art.sectionName || '',
            art.subsectionName || '',
            art.articleNo || '',
            art.articleTitle || '',
            art.articleContent || '',
            art.effectiveDate || rev.enforcementDate || '',
            art.isDeleted ? '삭제' : '',
          ]);

          if (rev.buchikText) {
            articleRows.push([
              '부칙',
              '',
              '',
              '부칙',
              `부칙 (${rev.promulgationNo || ''})`,
              rev.buchikText,
              rev.enforcementDate || '',
              '',
            ]);
          }

          // Write Data
          await callWithRetry(() =>
            sheets.spreadsheets.values.batchUpdate({
              spreadsheetId: spId,
              requestBody: {
                valueInputOption: 'USER_ENTERED',
                data: [
                  { range: `'${overviewSheet}'!A1`, values: overviewValues },
                  { range: `'${articlesSheet}'!A1`, values: [articleHeader, ...articleRows] },
                ],
              },
            })
          );

          newCreatedCount++;
          createdFiles.push({
            title: docTitle,
            spreadsheetId: spId,
            url: `https://docs.google.com/spreadsheets/d/${spId}/edit`,
            promulgationNo: rev.promulgationNo,
            enforcementDate: rev.enforcementDate,
            skipped: false,
          });

          // Pacing delay between files
          if (i < revisionList.length - 1) {
            await new Promise((resolve) => setTimeout(resolve, 800));
          }
        } catch (err: any) {
          console.warn(`[Separate Export Error] MST ${rev.lawMst || rev.id}:`, err?.message);
        }
      }

      const totalCount = createdFiles.length;
      let summaryMsg = `Google Drive '${targetFolderName}' 폴더에 저장이 완료되었습니다! (신규 생성: ${newCreatedCount}개, 중복 스킵: ${skippedCount}개)`;
      if (skippedCount > 0 && newCreatedCount === 0) {
        summaryMsg = `Google Drive '${targetFolderName}' 폴더에 모든 파일(${skippedCount}개)이 이미 존재하여 저장을 스킵했습니다.`;
      }

      return res.json({
        success: true,
        mode: 'separate_files',
        folderId,
        folderUrl,
        folderName: targetFolderName,
        folderSkipped,
        createdFiles,
        createdCount: newCreatedCount,
        skippedCount,
        totalCount,
        message: summaryMsg,
      });
    }

    return res.status(400).json({ error: '올바른 모드(single_file 또는 separate_files)를 지정해 주세요.' });
  } catch (error: any) {
    console.error('Export All Revisions Folder Error:', error);
    const errMsg = error.message || 'Google Drive 폴더 저장 중 오류가 발생했습니다.';
    const isAuthError =
      error.status === 401 ||
      errMsg.includes('Invalid Credentials') ||
      errMsg.includes('auth') ||
      errMsg.includes('token') ||
      errMsg.includes('UNAUTHENTICATED');
    const isForbidden =
      error.status === 403 ||
      errMsg.includes('PERMISSION_DENIED') ||
      errMsg.includes('insufficient');

    let userFriendlyMsg = errMsg;
    if (isAuthError) {
      userFriendlyMsg = 'Google 인증 토큰이 만료되었거나 유효하지 않습니다. 상단 [Google 로그인] 버튼을 눌러 다시 로그인해 주세요.';
    } else if (isForbidden) {
      userFriendlyMsg = 'Google Drive / Spreadsheets 쓰기 권한이 부족합니다. Google 로그인 시 드라이브 및 스프레드시트 권한을 허용해 주세요.';
    }

    return res.status(error.status || (isAuthError ? 401 : 500)).json({
      error: userFriendlyMsg,
      authError: isAuthError || isForbidden,
    });
  }
});

// API Route: List user's existing Google Drive folders for folder selection
app.post('/api/drive/list-folders', async (req, res) => {
  try {
    const { accessToken, searchQuery = '', pageSize = 50 } = req.body;
    if (!accessToken) {
      return res.status(401).json({ success: false, error: 'Google OAuth 인증 토큰이 필요합니다.', authError: true });
    }

    const auth = new google.auth.OAuth2();
    auth.setCredentials({ access_token: accessToken });
    const drive = google.drive({ version: 'v3', auth });

    let query = "mimeType = 'application/vnd.google-apps.folder' and trashed = false";
    if (searchQuery && typeof searchQuery === 'string' && searchQuery.trim()) {
      const escaped = searchQuery.trim().replace(/'/g, "\\'");
      query += ` and name contains '${escaped}'`;
    }

    const listRes: any = await drive.files.list({
      q: query,
      fields: 'files(id, name, webViewLink, createdTime, modifiedTime, shared, owners)',
      orderBy: 'modifiedTime desc',
      pageSize: Math.min(Number(pageSize) || 50, 100),
    });

    const folders = (listRes.data.files || []).map((f: any) => ({
      id: f.id,
      name: f.name,
      url: f.webViewLink || `https://drive.google.com/drive/folders/${f.id}`,
      createdTime: f.createdTime,
      modifiedTime: f.modifiedTime,
      isShared: !!f.shared,
    }));

    return res.json({
      success: true,
      folders,
    });
  } catch (err: any) {
    console.error('List Drive Folders Error:', err);
    const errMsg = String(err?.message || '').toLowerCase();
    const isAuth =
      err.status === 401 ||
      errMsg.includes('invalid credentials') ||
      errMsg.includes('unauthenticated') ||
      errMsg.includes('oauth');

    return res.status(isAuth ? 401 : 500).json({
      success: false,
      authError: isAuth,
      error: isAuth
        ? 'Google OAuth 인증 토큰이 만료되었습니다. 다시 로그인해 주세요.'
        : err?.message || 'Google Drive 폴더 목록 조회 중 오류가 발생했습니다.',
    });
  }
});

// API Route: Initialize Revision Folder & list existing spreadsheets for incremental safe batch export
app.post('/api/drive/init-revision-folder', async (req, res) => {
  try {
    const { accessToken, folderName, targetFolderId, permissionOption, selectedItem, targetType = 'law' } = req.body;
    if (!accessToken) {
      return res.status(401).json({ success: false, error: 'Google OAuth 인증 토큰이 필요합니다. 상단 Google 로그인 버튼을 눌러주세요.' });
    }

    const auth = new google.auth.OAuth2();
    auth.setCredentials({ access_token: accessToken });
    const drive = google.drive({ version: 'v3', auth });

    const cleanLawName = (selectedItem?.name || (targetType === 'admrul' ? '외국환거래규정' : '관세법')).trim();
    const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const defaultFolderName = `[${cleanLawName.replace(/[\/\\:*?"<>|]/g, '_')}_${today}]`;
    const requestedFolderName = (folderName || defaultFolderName).trim();

    let folderId = '';
    let folderUrl = '';
    let targetFolderName = requestedFolderName;
    let folderSkipped = false;

    // Case 1: User explicitly selected an existing Google Drive Folder ID
    if (targetFolderId && typeof targetFolderId === 'string' && targetFolderId.trim()) {
      try {
        const folderGetRes: any = await drive.files.get({
          fileId: targetFolderId.trim(),
          fields: 'id, name, webViewLink, trashed',
        });
        if (folderGetRes.data && !folderGetRes.data.trashed) {
          folderId = folderGetRes.data.id;
          targetFolderName = folderGetRes.data.name || requestedFolderName;
          folderUrl = folderGetRes.data.webViewLink || `https://drive.google.com/drive/folders/${folderId}`;
          folderSkipped = true; // Use existing folder
        }
      } catch (getErr: any) {
        console.warn('[Use Selected Target Folder ID Error, falling back to name creation]:', getErr?.message);
      }
    }

    // Case 2: Create or find unique folder by name if no explicit folderId resolved
    if (!folderId) {
      const folderRes = await getOrCreateUniqueDriveFolder(
        drive,
        requestedFolderName,
        async (fn) => fn()
      );
      folderId = folderRes.folderId;
      folderUrl = folderRes.folderUrl;
      targetFolderName = folderRes.folderName;
      folderSkipped = folderRes.folderSkipped;
    }

    if (!folderId) {
      return res.status(500).json({ success: false, error: 'Google Drive 폴더 생성에 실패했습니다. 계정 권한을 확인해 주세요.' });
    }

    if (permissionOption?.type === 'anyone' && folderId) {
      try {
        await drive.permissions.create({
          fileId: folderId,
          requestBody: {
            role: permissionOption.role || 'reader',
            type: 'anyone',
          },
        });
      } catch (permErr: any) {
        console.warn('[Init Folder Permission Warning]', permErr?.message);
      }
    }

    const existingFiles: Array<{ id: string; name: string; url: string }> = [];
    try {
      let pageToken: string | undefined = undefined;
      do {
        const listRes: any = await drive.files.list({
          q: `'${folderId}' in parents and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false`,
          fields: 'nextPageToken, files(id, name, webViewLink)',
          pageSize: 1000,
          pageToken,
        });
        (listRes.data.files || []).forEach((f: any) => {
          if (f.id && f.name) {
            existingFiles.push({
              id: f.id,
              name: f.name,
              url: f.webViewLink || `https://docs.google.com/spreadsheets/d/${f.id}/edit`,
            });
          }
        });
        pageToken = listRes.data.nextPageToken;
      } while (pageToken);
    } catch (listErr: any) {
      console.warn('[Init Folder List Existing Files Warning]', listErr?.message);
    }

    return res.json({
      success: true,
      folder: {
        id: folderId,
        name: targetFolderName,
        url: folderUrl,
        created: !folderSkipped,
        isExisting: folderSkipped,
      },
      existingFiles,
    });
  } catch (err: any) {
    console.error('Init Revision Folder Error:', err);
    const errMsg = String(err?.message || '').toLowerCase();
    const status = typeof err?.status === 'number' ? err.status : (typeof err?.code === 'number' ? err.code : Number(err?.status || err?.code || 0));
    const isAuth =
      status === 401 ||
      errMsg.includes('invalid authentication credentials') ||
      errMsg.includes('invalid credentials') ||
      errMsg.includes('expected oauth 2 access token') ||
      errMsg.includes('unauthenticated') ||
      errMsg.includes('invalid_grant') ||
      errMsg.includes('login cookie');

    if (isAuth) {
      return res.status(401).json({
        success: false,
        authError: true,
        error: 'Google OAuth 인증 토큰이 만료되었거나 유효하지 않습니다. 상단 Google 로그인 버튼을 눌러 다시 인증해 주세요.',
      });
    }

    return res.status(500).json({ success: false, error: err?.message || 'Google Drive 폴더 초기화 중 오류가 발생했습니다.' });
  }
});

// API Route: Export revision batch (typically 5 items) for fast, responsive, and robust execution with quota backoff
app.post('/api/drive/export-revision-batch', async (req, res) => {
  try {
    const {
      accessToken,
      folderId,
      targetFolderName,
      targetType = 'law',
      selectedItem,
      revisions = [],
      allRevisions = [],
      permissionOption,
      ocKey = DEFAULT_OC_KEY,
      startIndex = 0,
      existingFiles = [],
      duplicateHandlingMode = 'numbering', // 'numbering' (기존 파일 있을 시 (1) 부여하여 새로 저장) | 'skip' (중복 건너뛰기)
    } = req.body;

    if (!accessToken) {
      return res.status(401).json({ success: false, error: 'Google OAuth 인증 토큰이 필요합니다.', authError: true });
    }
    if (!folderId) {
      return res.status(400).json({ success: false, error: 'Google Drive 폴더 ID가 누락되었습니다.' });
    }

    const auth = new google.auth.OAuth2();
    auth.setCredentials({ access_token: accessToken });
    const drive = google.drive({ version: 'v3', auth });
    const sheets = google.sheets({ version: 'v4', auth });

    // Helper to check if an error is an OAuth token failure
    const isGoogleAuthError = (err: any): boolean => {
      const errMsg = String(err?.message || '').toLowerCase();
      const status = typeof err?.status === 'number' ? err.status : (typeof err?.code === 'number' ? err.code : Number(err?.status || err?.code || 0));
      return (
        status === 401 ||
        errMsg.includes('invalid authentication credentials') ||
        errMsg.includes('invalid credentials') ||
        errMsg.includes('expected oauth 2 access token') ||
        errMsg.includes('unauthenticated') ||
        errMsg.includes('invalid_grant') ||
        errMsg.includes('login cookie')
      );
    };

    // Internal retry helper for Google API quota (429) & network resilience
    const callWithRetry = async <T>(fn: () => Promise<T>, retries = 8, delay = 2500): Promise<T> => {
      let attempt = 0;
      let currentDelay = delay;

      while (true) {
        try {
          return await fn();
        } catch (err: any) {
          if (isGoogleAuthError(err)) {
            // Never retry on authentication failures
            throw err;
          }

          attempt++;
          const errMsg = String(err?.message || '').toLowerCase();
          const errCode = String(err?.code || '').toLowerCase();
          const status = typeof err?.status === 'number' ? err.status : (typeof err?.code === 'number' ? err.code : Number(err?.status || err?.code || 0));

          const isRateLimit =
            status === 429 ||
            errMsg.includes('quota') ||
            errMsg.includes('rate') ||
            errMsg.includes('resource_exhausted') ||
            errMsg.includes('write requests per minute') ||
            errMsg.includes('user rate limit');

          const isNetworkError =
            errCode === 'econnreset' ||
            errCode === 'etimedout' ||
            errCode === 'enotfound' ||
            errMsg.includes('socket hang up') ||
            errMsg.includes('network timeout') ||
            errMsg.includes('econnreset') ||
            errMsg.includes('fetch failed') ||
            status === 500 ||
            status === 502 ||
            status === 503 ||
            status === 504;

          if ((isRateLimit || isNetworkError) && attempt <= retries) {
            const backoffTime = isRateLimit
              ? Math.max(currentDelay, Math.min(65000, 5000 * Math.pow(1.5, attempt - 1)))
              : Math.max(currentDelay, Math.min(25000, 2000 * Math.pow(1.3, attempt - 1)));

            console.warn(
              `[Batch Drive/Sheets API Retry] Attempt ${attempt}/${retries} (${isRateLimit ? '429 Quota Exceeded' : (err?.message || errCode)}). Waiting ${Math.round(backoffTime)}ms...`
            );
            await new Promise((r) => setTimeout(r, backoffTime + Math.random() * 800));
            currentDelay = currentDelay * 1.5;
            continue;
          }
          throw err;
        }
      }
    };

    const cleanLawName = (selectedItem?.name || (targetType === 'admrul' ? '외국환거래규정' : '관세법')).trim();
    const existingFileMap = new Map<string, { id: string; url: string; name: string }>();
    (existingFiles || []).forEach((f: any) => {
      if (f.id && f.name) {
        existingFileMap.set(f.name, f);
        const cleanKey = f.name.replace(/^\d+_[_]?/, '').trim();
        if (cleanKey) existingFileMap.set(cleanKey, f);
      }
    });

    const savedSheets: any[] = [];
    let skippedCount = 0;
    let newCreatedCount = 0;

    // ==========================================
    // Master Summary & Standalone Buchik Google Sheets (on startIndex === 0)
    // ==========================================
    if (startIndex === 0) {
      const allRevs = Array.isArray(allRevisions) && allRevisions.length > 0 ? allRevisions : revisions;

      // 1. Master Summary Google Spreadsheet
      try {
        const normCleanLawName = normalizeLawTitleSpacing(cleanLawName);
        const summaryDocTitle = `[${normCleanLawName.replace(/[\/\\:*?"<>|]/g, '_')}] 개정연혁총괄목록`;
        let matchedSummary = existingFileMap.get(summaryDocTitle);
        if (!matchedSummary) {
          for (const [fName, fData] of existingFileMap.entries()) {
            if (fName.includes(normCleanLawName) && fName.includes('개정연혁총괄목록')) {
              matchedSummary = fData;
              break;
            }
          }
        }

        if (matchedSummary) {
          skippedCount++;
          savedSheets.push({
            title: matchedSummary.name || summaryDocTitle,
            spreadsheetId: matchedSummary.id,
            url: matchedSummary.url,
            isExisting: true,
          });
        } else {
          const sortedAll = sortRevisionsByHierarchyAndDate(allRevs);
          const createRes = await callWithRetry(() =>
            drive.files.create({
              requestBody: {
                name: summaryDocTitle,
                parents: [folderId],
                mimeType: 'application/vnd.google-apps.spreadsheet',
              },
              fields: 'id, name, webViewLink',
            })
          );
          const sSpId = createRes.data.id;
          const sSpUrl = createRes.data.webViewLink || `https://docs.google.com/spreadsheets/d/${sSpId}/edit`;

          if (sSpId) {
            if (permissionOption?.type === 'anyone') {
              try {
                await callWithRetry(() =>
                  drive.permissions.create({
                    fileId: sSpId,
                    requestBody: { role: permissionOption.role || 'reader', type: 'anyone' },
                  })
                );
              } catch (e: any) {}
            }

            const summarySheetName = '개정연혁 총괄목록';
            const sumHeaders = [
              '순번',
              '구분',
              '법령/규칙명',
              '공포/발령번호',
              '공포/발령일자',
              '시행일자',
              '제개정구분',
              '소관부처',
              '조문수',
              '개정부칙요약',
            ];

            const sumColWidths = [60, 100, 180, 120, 110, 110, 100, 120, 80, 450];
            const sumDimensionRequests = sumColWidths.map((px, idx) => ({
              updateDimensionProperties: {
                range: { sheetId: 0, dimension: 'COLUMNS', startIndex: idx, endIndex: idx + 1 },
                properties: { pixelSize: px },
                fields: 'pixelSize',
              },
            }));

            await callWithRetry(() =>
              sheets.spreadsheets.batchUpdate({
                spreadsheetId: sSpId,
                requestBody: {
                  requests: [
                    { updateSheetProperties: { properties: { sheetId: 0, title: summarySheetName, gridProperties: { frozenRowCount: 1 } }, fields: 'title,gridProperties.frozenRowCount' } },
                    {
                      repeatCell: {
                        range: { sheetId: 0, startRowIndex: 0 },
                        cell: { userEnteredFormat: { verticalAlignment: 'TOP', wrapStrategy: 'WRAP' } },
                        fields: 'userEnteredFormat(verticalAlignment,wrapStrategy)',
                      },
                    },
                    {
                      repeatCell: {
                        range: { sheetId: 0, startRowIndex: 0, endRowIndex: 1 },
                        cell: {
                          userEnteredFormat: {
                            backgroundColor: { red: 0.12, green: 0.18, blue: 0.3 },
                            textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } },
                            verticalAlignment: 'MIDDLE',
                            horizontalAlignment: 'CENTER',
                          },
                        },
                        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
                      },
                    },
                    ...sumDimensionRequests,
                  ],
                },
              })
            );

            const sumRows = sortedAll.map((r: any, idx: number) => {
              const subType = (r.subType || '').toLowerCase();
              const rName = (r.name || cleanLawName).trim();
              const typeLabel =
                r.targetType === 'admrul' || r.ruleType?.includes('고시') || r.ruleType?.includes('훈령') || r.ruleType?.includes('예규')
                  ? `행정규칙(${r.ruleType || '고시'})`
                  : subType === 'decree' || rName.includes('시행령') || r.ruleType?.includes('대통령령')
                  ? '시행령'
                  : subType === 'rule' || rName.includes('시행규칙') || r.ruleType?.includes('부령')
                  ? '시행규칙'
                  : '법률';

              return [
                idx + 1,
                typeLabel,
                rName,
                r.promulgationNo || '-',
                formatStandardKoreanDate(r.promulgationDate) || '-',
                formatStandardKoreanDate(r.enforcementDate) || '-',
                r.revisionType || '일부개정',
                r.department || (targetType === 'admrul' ? '관세청' : '기획재정부'),
                r.articleCount ? `${r.articleCount}개` : '-',
                r.buchikText || '-',
              ];
            });

            await callWithRetry(() =>
              sheets.spreadsheets.values.batchUpdate({
                spreadsheetId: sSpId,
                requestBody: {
                  valueInputOption: 'USER_ENTERED',
                  data: [{ range: `'${summarySheetName}'!A1`, values: [sumHeaders, ...sumRows] }],
                },
              })
            );

            newCreatedCount++;
            savedSheets.push({
              title: summaryDocTitle,
              spreadsheetId: sSpId,
              url: sSpUrl,
              isExisting: false,
            });
            existingFileMap.set(summaryDocTitle, { id: sSpId, name: summaryDocTitle, url: sSpUrl });
            console.log(`[Batch Drive Export] Created Master Summary Sheet: ${summaryDocTitle}`);
          }
        }
      } catch (sumErr: any) {
        console.warn('[Batch Drive Export Master Summary Warning]:', sumErr?.message);
      }

      // 2. Standalone Buchik Google Spreadsheets (1 per law subtype: Law, Decree, Rule)
      try {
        const normCleanLawName = normalizeLawTitleSpacing(cleanLawName);
        const baseLawName = normCleanLawName.replace(/\s*(?:시행령|시행규칙|법률)$/, '').trim();
        const typesToGenerate: Array<{ key: 'law' | 'decree' | 'rule'; name: string }> = [];

        const hasLaw = allRevs.some((r: any) => {
          const rName = (r.name || '').trim();
          const rSub = (r.subType || '').toLowerCase();
          const rRule = (r.ruleType || '').trim();
          return rSub === 'law' || rRule === '법률' || (!rName.includes('시행령') && !rName.includes('시행규칙') && !rRule.includes('대통령령') && !rRule.includes('부령'));
        });
        const hasDecree = allRevs.some((r: any) => {
          const rName = (r.name || '').trim();
          const rSub = (r.subType || '').toLowerCase();
          const rRule = (r.ruleType || '').trim();
          return rSub === 'decree' || rName.includes('시행령') || rRule.includes('대통령령');
        });
        const hasRule = allRevs.some((r: any) => {
          const rName = (r.name || '').trim();
          const rSub = (r.subType || '').toLowerCase();
          const rRule = (r.ruleType || '').trim();
          return rSub === 'rule' || rName.includes('시행규칙') || rRule.includes('부령');
        });

        if (targetType === 'admrul') {
          typesToGenerate.push({ key: 'law', name: normCleanLawName });
        } else {
          if (hasLaw) typesToGenerate.push({ key: 'law', name: baseLawName });
          if (hasDecree) typesToGenerate.push({ key: 'decree', name: `${baseLawName} 시행령` });
          if (hasRule) typesToGenerate.push({ key: 'rule', name: `${baseLawName} 시행규칙` });
        }

        for (const t of typesToGenerate) {
          const normTypeName = normalizeLawTitleSpacing(t.name);
          const buchikDocTitle = `[${normTypeName}] 부칙`.replace(/[\/\\:*?"<>|]/g, '_');
          let matchedBuchik = existingFileMap.get(buchikDocTitle);
          if (!matchedBuchik) {
            for (const [fName, fData] of existingFileMap.entries()) {
              if (fName.includes(`[${t.name}]`) && fName.includes('부칙')) {
                matchedBuchik = fData;
                break;
              }
            }
          }

          if (matchedBuchik) {
            skippedCount++;
            savedSheets.push({
              title: matchedBuchik.name || buchikDocTitle,
              spreadsheetId: matchedBuchik.id,
              url: matchedBuchik.url,
              isExisting: true,
            });
            console.log(`[Batch Buchik Exists - Skipped] ${buchikDocTitle}`);
            continue;
          }

          const revsOfSub = allRevs.filter((r: any) => {
            const rName = (r.name || '').trim();
            const rSub = (r.subType || '').toLowerCase();
            if (t.key === 'decree') return rSub === 'decree' || rName.includes('시행령');
            if (t.key === 'rule') return rSub === 'rule' || rName.includes('시행규칙');
            return rSub === 'law' || (!rName.includes('시행령') && !rName.includes('시행규칙'));
          });
          const latestRev = sortRevisionsByEnforcementDateDesc(revsOfSub)[0] || revsOfSub[0] || allRevs[0];
          const buchikArticles = await fetchAllBuchikArticlesForLaw(ocKey, t.name, targetType, latestRev, t.key);

          if (buchikArticles && buchikArticles.length > 0) {
            const createRes = await callWithRetry(() =>
              drive.files.create({
                requestBody: {
                  name: buchikDocTitle,
                  parents: [folderId],
                  mimeType: 'application/vnd.google-apps.spreadsheet',
                },
                fields: 'id, name, webViewLink',
              })
            );
            const bSpId = createRes.data.id;
            const bSpUrl = createRes.data.webViewLink || `https://docs.google.com/spreadsheets/d/${bSpId}/edit`;

            if (bSpId) {
              if (permissionOption?.type === 'anyone') {
                try {
                  await callWithRetry(() =>
                    drive.permissions.create({
                      fileId: bSpId,
                      requestBody: { role: permissionOption.role || 'reader', type: 'anyone' },
                    })
                  );
                } catch (e: any) {}
              }

              const buchikSheetName = '부칙';
              const buchikHeaders = ['부칙구분', '관련법령', '조문번호', '조문제목', '조문내용'];
              const buchikColWidths = [220, 200, 90, 130, 550];
              const buchikDimRequests = buchikColWidths.map((px, idx) => ({
                updateDimensionProperties: {
                  range: { sheetId: 0, dimension: 'COLUMNS', startIndex: idx, endIndex: idx + 1 },
                  properties: { pixelSize: px },
                  fields: 'pixelSize',
                },
              }));

              await callWithRetry(() =>
                sheets.spreadsheets.batchUpdate({
                  spreadsheetId: bSpId,
                  requestBody: {
                    requests: [
                      { updateSheetProperties: { properties: { sheetId: 0, title: buchikSheetName, gridProperties: { frozenRowCount: 1 } }, fields: 'title,gridProperties.frozenRowCount' } },
                      {
                        repeatCell: {
                          range: { sheetId: 0, startRowIndex: 0 },
                          cell: { userEnteredFormat: { verticalAlignment: 'TOP', wrapStrategy: 'WRAP' } },
                          fields: 'userEnteredFormat(verticalAlignment,wrapStrategy)',
                        },
                      },
                      {
                        repeatCell: {
                          range: { sheetId: 0, startRowIndex: 0, endRowIndex: 1 },
                          cell: {
                            userEnteredFormat: {
                              backgroundColor: { red: 0.12, green: 0.18, blue: 0.3 },
                              textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } },
                              verticalAlignment: 'MIDDLE',
                              horizontalAlignment: 'CENTER',
                            },
                          },
                          fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
                        },
                      },
                      ...buchikDimRequests,
                    ],
                  },
                })
              );

              const buchikRows = buchikArticles.map((art: any) => [
                art.buchikCategory || '',
                art.relatedLaw || '',
                art.articleNo || '',
                art.articleTitle || '',
                art.articleContent || '',
              ]);

              await callWithRetry(() =>
                sheets.spreadsheets.values.batchUpdate({
                  spreadsheetId: bSpId,
                  requestBody: {
                    valueInputOption: 'USER_ENTERED',
                    data: [{ range: `'${buchikSheetName}'!A1`, values: [buchikHeaders, ...buchikRows] }],
                  },
                })
              );

              newCreatedCount++;
              savedSheets.push({
                title: buchikDocTitle,
                spreadsheetId: bSpId,
                url: bSpUrl,
                articleCount: buchikArticles.length,
                isExisting: false,
              });
              existingFileMap.set(buchikDocTitle, { id: bSpId, name: buchikDocTitle, url: bSpUrl });
              console.log(`[Batch Drive Export] Created Standalone Buchik Sheet: ${buchikDocTitle} (${buchikArticles.length} entries)`);
            }
          }
        }
      } catch (buchikErr: any) {
        console.warn('[Batch Drive Export Standalone Buchik Warning]:', buchikErr?.message);
      }
    }

    for (let i = 0; i < revisions.length; i++) {
      const rev = revisions[i];
      const revDocName = (rev.name || cleanLawName).trim();
      const safePromNo = (rev.promulgationNo || '개정본').trim();
      const safeEnfDate = (rev.enforcementDate || '시행일 미상').trim();
      
      // Standard title without leading numbers: e.g. "관세법 [시행 2026. 8. 11.] [법률 제21858호, 2026. 8. 11., 일부개정]"
      const docTitle = generateStandardRevisionTitle(rev, cleanLawName, targetType);
      const safePromDate = (rev.promulgationDate || '').trim();
      const dateLabel = safePromDate ? `(${safePromDate} 개정, ${safeEnfDate} 시행)` : `(${safeEnfDate} 시행)`;
      const legacyWithoutIndex = `[${revDocName}] ${safePromNo} ${dateLabel}`.replace(/[\/\\:*?"<>|]/g, '_');

      let matchedExisting = existingFileMap.get(docTitle) || existingFileMap.get(legacyWithoutIndex);
      if (!matchedExisting && safePromNo && safePromNo !== '개정본' && safeEnfDate && safeEnfDate !== '시행일 미상') {
        const cleanEnfDigits = safeEnfDate.replace(/\D/g, '');
        for (const [fName, fData] of existingFileMap.entries()) {
          // Strictly match law name, promulgation number, AND enforcement date
          // Revisions with identical promulgation numbers but different enforcement dates (e.g. 2026.04.01 vs 2026.07.01) must NOT be skipped
          const hasProm = fName.includes(safePromNo);
          const hasDoc = fName.includes(revDocName);
          const fNameDigits = fName.replace(/\D/g, '');
          const hasEnf = fName.includes(safeEnfDate) || (cleanEnfDigits.length >= 8 && fNameDigits.includes(cleanEnfDigits));
          if (hasProm && hasDoc && hasEnf) {
            matchedExisting = fData;
            break;
          }
        }
      }

      // If duplicateHandlingMode is 'skip', skip duplicate items
      if (matchedExisting && duplicateHandlingMode === 'skip') {
        skippedCount++;
        savedSheets.push({
          title: matchedExisting.name || docTitle,
          spreadsheetId: matchedExisting.id,
          url: matchedExisting.url,
          promulgationNo: rev.promulgationNo,
          enforcementDate: rev.enforcementDate,
          isExisting: true,
        });
        continue;
      }

      // If duplicateHandlingMode is 'numbering' and a file with identical title exists,
      // append (1), (2), (3)... to the title so it is saved without overwriting or skipping!
      let finalDocTitle = docTitle;
      let duplicateIndex = 1;
      if (matchedExisting || existingFileMap.has(finalDocTitle)) {
        while (existingFileMap.has(finalDocTitle)) {
          finalDocTitle = `${docTitle} (${duplicateIndex})`;
          duplicateIndex++;
        }
      }

      try {
        const isAdmrul =
          rev.targetType === 'admrul' ||
          targetType === 'admrul' ||
          revDocName.includes('규정') ||
          revDocName.includes('고시') ||
          revDocName.includes('훈령') ||
          revDocName.includes('예규');

        let revArticles = await fetchArticlesForRevision(ocKey, rev, isAdmrul ? 'admrul' : 'law');
        if (rev.buchikText && !revArticles.some((a: any) => a.articleTitle?.includes('부칙') || a.articleNo?.includes('부칙'))) {
          revArticles.push({
            chapterName: '',
            sectionName: '',
            subsectionName: '',
            articleNo: '부칙',
            articleTitle: '개정 부칙',
            articleContent: rev.buchikText,
            effectiveDate: safeEnfDate,
            isDeleted: false,
          });
        }

        // Pacing delay to avoid burst quota exhaustion
        await new Promise((r) => setTimeout(r, 400));

        const createRes = await callWithRetry(() =>
          drive.files.create({
            requestBody: {
              name: finalDocTitle,
              parents: [folderId],
              mimeType: 'application/vnd.google-apps.spreadsheet',
            },
            fields: 'id, name, webViewLink',
          })
        );

        const spId = createRes.data.id;
        const spUrl = createRes.data.webViewLink || `https://docs.google.com/spreadsheets/d/${spId}/edit`;

        if (!spId) continue;

        const overviewSheet = `${revDocName.replace(/[\/\\:*?\[\]]/g, '_').slice(0, 20)} 개요`;
        const articlesSheet = '조문목록';

        // 7. Dynamic column dimension requests depending on isAdmrul
        const sheetDimensionRequests: any[] = [
          // Column width: Article content (조문내용) at 500px / 550px
          {
            updateDimensionProperties: {
              range: { sheetId: 1, dimension: 'COLUMNS', startIndex: 5, endIndex: 6 },
              properties: { pixelSize: isAdmrul ? 550 : 500 },
              fields: 'pixelSize',
            },
          },
          // Column widths for Chapter/Section/Subsection, Article No, Article Title
          {
            updateDimensionProperties: {
              range: { sheetId: 1, dimension: 'COLUMNS', startIndex: 0, endIndex: 3 },
              properties: { pixelSize: 90 },
              fields: 'pixelSize',
            },
          },
          {
            updateDimensionProperties: {
              range: { sheetId: 1, dimension: 'COLUMNS', startIndex: 3, endIndex: 4 },
              properties: { pixelSize: 110 },
              fields: 'pixelSize',
            },
          },
          {
            updateDimensionProperties: {
              range: { sheetId: 1, dimension: 'COLUMNS', startIndex: 4, endIndex: 5 },
              properties: { pixelSize: 150 },
              fields: 'pixelSize',
            },
          },
        ];

        // Statutes (법령) include 전문개정, 본조신설, 제목개정, 비고 columns
        if (!isAdmrul) {
          sheetDimensionRequests.push(
            {
              updateDimensionProperties: {
                range: { sheetId: 1, dimension: 'COLUMNS', startIndex: 6, endIndex: 9 },
                properties: { pixelSize: 140 },
                fields: 'pixelSize',
              },
            },
            {
              updateDimensionProperties: {
                range: { sheetId: 1, dimension: 'COLUMNS', startIndex: 9, endIndex: 10 },
                properties: { pixelSize: 220 },
                fields: 'pixelSize',
              },
            }
          );
        }

        await callWithRetry(() =>
          sheets.spreadsheets.batchUpdate({
            spreadsheetId: spId,
            requestBody: {
              requests: [
                // 1. Update Sheet 0 title
                { updateSheetProperties: { properties: { sheetId: 0, title: overviewSheet }, fields: 'title' } },
                // 2. Add Sheet 1 (조문목록) with frozen top header row
                { addSheet: { properties: { sheetId: 1, title: articlesSheet, index: 1, gridProperties: { frozenRowCount: 1 } } } },
                // 3. Format Sheet 0 (Overview) cells: TOP vertical alignment & WRAP text
                {
                  repeatCell: {
                    range: { sheetId: 0, startRowIndex: 0 },
                    cell: {
                      userEnteredFormat: {
                        verticalAlignment: 'TOP',
                        wrapStrategy: 'WRAP',
                      },
                    },
                    fields: 'userEnteredFormat(verticalAlignment,wrapStrategy)',
                  },
                },
                // 4. Format Sheet 0 Header row: CENTER alignment, Bold, Navy background
                {
                  repeatCell: {
                    range: { sheetId: 0, startRowIndex: 0, endRowIndex: 1 },
                    cell: {
                      userEnteredFormat: {
                        backgroundColor: { red: 0.12, green: 0.18, blue: 0.3 },
                        textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } },
                        verticalAlignment: 'MIDDLE',
                        horizontalAlignment: 'CENTER',
                      },
                    },
                    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
                  },
                },
                // 5. Format Sheet 1 (Articles) cells: TOP vertical alignment & WRAP text
                {
                  repeatCell: {
                    range: { sheetId: 1, startRowIndex: 0 },
                    cell: {
                      userEnteredFormat: {
                        verticalAlignment: 'TOP',
                        wrapStrategy: 'WRAP',
                      },
                    },
                    fields: 'userEnteredFormat(verticalAlignment,wrapStrategy)',
                  },
                },
                // 6. Format Sheet 1 Header row: CENTER alignment, Bold, Navy background
                {
                  repeatCell: {
                    range: { sheetId: 1, startRowIndex: 0, endRowIndex: 1 },
                    cell: {
                      userEnteredFormat: {
                        backgroundColor: { red: 0.12, green: 0.18, blue: 0.3 },
                        textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } },
                        verticalAlignment: 'MIDDLE',
                        horizontalAlignment: 'CENTER',
                      },
                    },
                    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
                  },
                },
                ...sheetDimensionRequests,
              ],
            },
          })
        );

        const overviewValues = [
          ['항목', '내용'],
          ['법령 / 행정규칙명', rev.name || cleanLawName],
          ['공포 / 발령번호', rev.promulgationNo || '-'],
          ['시행일자', safeEnfDate],
          ['공포 / 발령일자', rev.promulgationDate || '-'],
          ['제개정구분', rev.revisionType || '일부개정'],
          ['소관부처', rev.department || '기획재정부'],
          ['법령구분 / 종류', isAdmrul ? (rev.ruleType ? `행정규칙(${rev.ruleType})` : '행정규칙(고시)') : (rev.ruleType || '법률')],
          ['수록 조문 수', `${revArticles.length}개 조문`],
          ['저장 폴더', targetFolderName || 'Google Drive'],
          ['생성 일시', new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' })],
        ];

        // Administrative Rules (행정규칙) only go up to 조문내용 (exclude 본조신설, 전문개정, 제목개정, 비고)
        const articleHeaders = isAdmrul
          ? ['장', '절', '관', '조문번호', '조문제목', '조문내용']
          : ['장', '절', '관', '조문번호', '조문제목', '조문내용', '전문개정', '본조신설', '제목개정', '비고'];

        const articleRows = revArticles.map((art: any) =>
          isAdmrul
            ? [
                art.chapterName || '',
                art.sectionName || '',
                art.subsectionName || '',
                art.articleNo || '',
                art.articleTitle || '',
                art.articleContent || '',
              ]
            : [
                art.chapterName || '',
                art.sectionName || '',
                art.subsectionName || '',
                art.articleNo || '',
                art.articleTitle || '',
                art.articleContent || '',
                art.fullRevision || '',
                art.creation || '',
                art.titleRevision || '',
                art.remarks || (art.isDeleted ? '삭제' : ''),
              ]
        );

        await callWithRetry(() =>
          sheets.spreadsheets.values.batchUpdate({
            spreadsheetId: spId,
            requestBody: {
              valueInputOption: 'USER_ENTERED',
              data: [
                { range: `'${overviewSheet}'!A1`, values: overviewValues },
                { range: `'${articlesSheet}'!A1`, values: [articleHeaders, ...articleRows] },
              ],
            },
          })
        );

        newCreatedCount++;
        const isNumbered = duplicateIndex > 1;
        const sheetItem = {
          title: finalDocTitle,
          spreadsheetId: spId,
          url: spUrl,
          promulgationNo: rev.promulgationNo,
          enforcementDate: rev.enforcementDate,
          articleCount: revArticles.length,
          isExisting: false,
          isNumbered,
          duplicateNumber: isNumbered ? (duplicateIndex - 1) : 0,
        };
        savedSheets.push(sheetItem);
        existingFileMap.set(finalDocTitle, { id: spId, name: finalDocTitle, url: spUrl });
      } catch (itemErr: any) {
        if (isGoogleAuthError(itemErr)) {
          throw itemErr;
        }
        console.warn(`[Batch Save Item Error] ${finalDocTitle}:`, itemErr?.message);
      }
    }

    return res.json({
      success: true,
      savedSheets,
      createdCount: newCreatedCount,
      skippedCount,
    });
  } catch (err: any) {
    console.error('Export Revision Batch Error:', err);
    const errMsg = String(err?.message || '').toLowerCase();
    const status = typeof err?.status === 'number' ? err.status : (typeof err?.code === 'number' ? err.code : Number(err?.status || err?.code || 0));
    const isAuth =
      status === 401 ||
      errMsg.includes('invalid authentication credentials') ||
      errMsg.includes('invalid credentials') ||
      errMsg.includes('expected oauth 2 access token') ||
      errMsg.includes('unauthenticated') ||
      errMsg.includes('invalid_grant') ||
      errMsg.includes('login cookie');

    if (isAuth) {
      return res.status(401).json({
        success: false,
        authError: true,
        error: 'Google OAuth 인증 토큰이 만료되었거나 유효하지 않습니다. Google 계정에 다시 로그인해 주세요.',
      });
    }

    return res.status(500).json({ success: false, error: err?.message || '시트 배치 저장 중 오류가 발생했습니다.' });
  }
});

// API Route: Export a batch of Excel workbooks (10~15 items) as base64 buffers for progressive, zero-timeout client packaging
app.post('/api/unified/export-excel-batch', async (req, res) => {
  try {
    const {
      ocKey = DEFAULT_OC_KEY,
      targetType = 'law',
      selectedItem,
      revisions = [],
      cleanLawName: customCleanLawName,
      targetFolderName = '',
      includeMasterSummary = false,
      allRevisions = [],
      startIndex = 0,
    } = req.body;

    if (!Array.isArray(revisions) || revisions.length === 0) {
      return res.json({ success: true, files: [] });
    }

    const cleanLawName = normalizeLawTitleSpacing(
      (customCleanLawName || selectedItem?.name || (targetType === 'admrul' ? '외국환거래규정' : '관세법')).trim()
    );
    const files: Array<{ filename: string; base64: string }> = [];

    // Optional Master Summary Excel & Standalone Buchik Excel in first batch
    if (includeMasterSummary && Array.isArray(allRevisions) && allRevisions.length > 0) {
      try {
        const sortedAll = sortRevisionsByHierarchyAndDate(allRevisions);
        const summaryBuffer = await createMasterSummaryWorkbookBuffer(sortedAll, cleanLawName);
        files.push({
          filename: `[${cleanLawName.replace(/[\/\\:*?"<>|]/g, '_')}] 개정연혁총괄목록.xlsx`,
          base64: summaryBuffer.toString('base64'),
        });
      } catch (sumErr: any) {
        console.warn('[Master Summary Excel Warning]:', sumErr?.message);
      }

      // Standalone Buchik Excel files (Generate separate buchik file for each subtype present in revisions: Law, Decree, Rule)
      try {
        const baseLawName = cleanLawName.replace(/\s*(?:시행령|시행규칙|법률)$/, '').trim();
        const typesToGenerate: Array<{ key: 'law' | 'decree' | 'rule'; name: string }> = [];

        const hasLaw = allRevisions.some((r: any) => {
          const rName = (r.name || '').trim();
          const rSub = (r.subType || '').toLowerCase();
          const rRule = (r.ruleType || '').trim();
          return rSub === 'law' || rRule === '법률' || (!rName.includes('시행령') && !rName.includes('시행규칙') && !rRule.includes('대통령령') && !rRule.includes('부령'));
        });
        const hasDecree = allRevisions.some((r: any) => {
          const rName = (r.name || '').trim();
          const rSub = (r.subType || '').toLowerCase();
          const rRule = (r.ruleType || '').trim();
          return rSub === 'decree' || rName.includes('시행령') || rRule.includes('대통령령');
        });
        const hasRule = allRevisions.some((r: any) => {
          const rName = (r.name || '').trim();
          const rSub = (r.subType || '').toLowerCase();
          const rRule = (r.ruleType || '').trim();
          return rSub === 'rule' || rName.includes('시행규칙') || rRule.includes('부령');
        });

        if (targetType === 'admrul') {
          typesToGenerate.push({ key: 'law', name: cleanLawName });
        } else {
          if (hasLaw) typesToGenerate.push({ key: 'law', name: baseLawName });
          if (hasDecree) typesToGenerate.push({ key: 'decree', name: `${baseLawName} 시행령` });
          if (hasRule) typesToGenerate.push({ key: 'rule', name: `${baseLawName} 시행규칙` });
        }

        for (const t of typesToGenerate) {
          const revsOfSub = allRevisions.filter((r: any) => {
            const rName = (r.name || '').trim();
            const rSub = (r.subType || '').toLowerCase();
            if (t.key === 'decree') return rSub === 'decree' || rName.includes('시행령');
            if (t.key === 'rule') return rSub === 'rule' || rName.includes('시행규칙');
            return rSub === 'law' || (!rName.includes('시행령') && !rName.includes('시행규칙'));
          });
          const latestRev = sortRevisionsByEnforcementDateDesc(revsOfSub)[0] || revsOfSub[0] || allRevisions[0];
          const buchikArticles = await fetchAllBuchikArticlesForLaw(ocKey, t.name, targetType, latestRev, t.key);
          if (buchikArticles.length > 0) {
            const buchikBuffer = await createBuchikWorkbookBuffer(t.name, buchikArticles);
            const normTypeName = normalizeLawTitleSpacing(t.name);
            files.push({
              filename: `[${normTypeName.replace(/[\/\\:*?"<>|]/g, '_')}] 부칙.xlsx`,
              base64: buchikBuffer.toString('base64'),
            });
            console.log(`[Batch Excel Export] Added standalone Buchik file '[${normTypeName}] 부칙.xlsx' (${buchikArticles.length} entries)`);
          }
        }
      } catch (buchikErr: any) {
        console.warn('[Standalone Buchik Excel Warning]:', buchikErr?.message);
      }
    }

    // Process this batch of revisions in parallel
    await Promise.all(
      revisions.map(async (rev: any, chunkIdx: number) => {
        const revDocName = (rev.name || cleanLawName).trim();
        // Standard filename without leading index numbers: e.g. "관세법 [시행 2026. 8. 11.] [법률 제21858호, 2026. 8. 11., 일부개정].xlsx"
        const filename = `${generateStandardRevisionTitle(rev, cleanLawName, targetType)}.xlsx`;

        try {
          const isAdmrul =
            rev.targetType === 'admrul' ||
            targetType === 'admrul' ||
            revDocName.includes('규정') ||
            revDocName.includes('고시');
          const revArticles = await fetchArticlesForRevision(ocKey, rev, isAdmrul ? 'admrul' : 'law');
          const xlsxBuf = await createRevisionWorkbookBuffer(rev, revArticles, cleanLawName, targetFolderName);
          files.push({
            filename,
            base64: xlsxBuf.toString('base64'),
          });
        } catch (itemErr: any) {
          console.warn(`[Batch Excel Export Error] ${filename}:`, itemErr?.message);
        }
      })
    );

    // Keep deterministic order in returned files
    files.sort((a, b) => a.filename.localeCompare(b.filename));

    return res.json({ success: true, files });
  } catch (err: any) {
    console.error('Export Excel Batch Error:', err);
    return res.status(500).json({ success: false, error: err?.message || '엑셀 배치 생성 중 오류가 발생했습니다.' });
  }
});

// API Route: Export revision sheets for Unified Search & Drive Exporter (법령 · 행정규칙 드라이브 연동)
app.post('/api/drive/export-revision-sheets', async (req, res) => {
  try {
    const {
      accessToken,
      targetType = 'law',
      selectedItem,
      revisions = [],
      folderName,
      permissionOption,
      ocKey = DEFAULT_OC_KEY,
    } = req.body;

    if (!accessToken) {
      return res.status(401).json({
        success: false,
        error: 'Google OAuth 인증 토큰이 필요합니다. 상단 Google 로그인 버튼을 눌러주세요.',
        authError: true,
      });
    }

    const auth = new google.auth.OAuth2();
    auth.setCredentials({ access_token: accessToken });
    const drive = google.drive({ version: 'v3', auth });
    const sheets = google.sheets({ version: 'v4', auth });

    // Helper for API retry with intelligent backoff and connection error recovery
    const callWithRetry = async <T>(fn: () => Promise<T>, retries = 10, delay = 2000): Promise<T> => {
      let attempt = 0;
      let currentDelay = delay;

      while (true) {
        try {
          return await fn();
        } catch (err: any) {
          attempt++;
          const errMsg = String(err?.message || '').toLowerCase();
          const errCode = String(err?.code || '').toLowerCase();
          const status = typeof err?.status === 'number' ? err.status : (typeof err?.code === 'number' ? err.code : Number(err?.status || err?.code || 0));

          const isRateLimit =
            status === 429 ||
            errMsg.includes('quota') ||
            errMsg.includes('rate') ||
            errMsg.includes('resource_exhausted') ||
            errMsg.includes('write requests per minute') ||
            errMsg.includes('user rate limit');

          const isNetworkError =
            errCode === 'econnreset' ||
            errCode === 'etimedout' ||
            errCode === 'enotfound' ||
            errCode === 'eai_again' ||
            errCode === 'und_err_connect_timeout' ||
            errMsg.includes('socket hang up') ||
            errMsg.includes('network timeout') ||
            errMsg.includes('econnreset') ||
            errMsg.includes('fetch failed') ||
            errMsg.includes('econnrefused') ||
            status === 500 ||
            status === 502 ||
            status === 503 ||
            status === 504;

          if ((isRateLimit || isNetworkError) && attempt <= retries) {
            const backoffTime = isRateLimit
              ? Math.max(currentDelay, Math.min(65000, 6000 * Math.pow(1.5, attempt - 1)))
              : Math.max(currentDelay, Math.min(30000, 2000 * Math.pow(1.4, attempt - 1)));

            console.warn(
              `[Unified Drive/Sheets API Retry] Attempt ${attempt}/${retries} (${isRateLimit ? '429 Quota Exceeded' : (err?.message || errCode)}). Waiting ${Math.round(backoffTime)}ms...`
            );
            await new Promise((r) => setTimeout(r, backoffTime + Math.random() * 900));
            currentDelay = currentDelay * 1.5;
            continue;
          }
          throw err;
        }
      }
    };

    const cleanLawName = (selectedItem?.name || (targetType === 'admrul' ? '외국환거래규정' : '관세법')).trim();
    const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const defaultFolderName = `[${cleanLawName.replace(/[\/\\:*?"<>|]/g, '_')}_${today}]`;
    const requestedFolderName = (folderName || defaultFolderName).trim();

    console.log(`[Unified Drive Export] Starting export for '${cleanLawName}', requested folder: '${requestedFolderName}', revisions: ${revisions.length}`);

    // 1. Create or Find Unique Folder in Google Drive (appends (1), (2) on conflict)
    const { folderId, folderUrl, folderName: targetFolderName, folderSkipped } = await getOrCreateUniqueDriveFolder(
      drive,
      requestedFolderName,
      callWithRetry
    );

    console.log(`[Unified Drive Folder Initialized] ID: ${folderId}, Name: '${targetFolderName}'`);

    // 2. Set Public Permission on Folder if requested
    if (permissionOption?.type === 'anyone') {
      try {
        await callWithRetry(() =>
          drive.permissions.create({
            fileId: folderId,
            requestBody: {
              role: permissionOption.role || 'reader',
              type: 'anyone',
            },
          })
        );
      } catch (permErr: any) {
        console.warn('[Unified Folder Permission Warning]', permErr?.message);
      }
    }

    // 3. Retrieve all existing files in folder with complete pagination
    const existingFileMap = new Map<string, { id: string; url: string; name: string }>();
    try {
      let pageToken: string | undefined = undefined;
      do {
        const existingFilesRes: any = await callWithRetry(() =>
          drive.files.list({
            q: `'${folderId}' in parents and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false`,
            fields: 'nextPageToken, files(id, name, webViewLink)',
            pageSize: 1000,
            pageToken: pageToken,
            spaces: 'drive',
          })
        );
        const fileList = existingFilesRes.data.files || [];
        fileList.forEach((f: any) => {
          if (f.name && f.id) {
            const fileData = {
              id: f.id,
              name: f.name,
              url: f.webViewLink || `https://docs.google.com/spreadsheets/d/${f.id}/edit`,
            };
            existingFileMap.set(f.name, fileData);
            // Normalized key without leading index (e.g. "001_[관세법]..." -> "[관세법]...")
            const cleanKey = f.name.replace(/^\d+_[_]?/, '').trim();
            if (cleanKey) {
              existingFileMap.set(cleanKey, fileData);
            }
          }
        });
        pageToken = existingFilesRes.data.nextPageToken;
      } while (pageToken);
      console.log(`[Unified Drive Export] Found ${existingFileMap.size} existing files in target folder (${folderId}).`);
    } catch (listErr: any) {
      console.warn('[Unified Existing Files List Warning]', listErr?.message);
    }

    const savedSheets: Array<{
      title: string;
      url: string;
      spreadsheetId: string;
      isExisting?: boolean;
      promulgationNo?: string;
      enforcementDate?: string;
      articleCount?: number;
    }> = [];

    let skippedCount = 0;
    let newCreatedCount = 0;

    // 4. Sort revisions by enforcement date descending (최근 시행일자순 우선 정렬 후 번호 001, 002... 부여)
    const sortedRevisions = sortRevisionsByEnforcementDateDesc(revisions);

    // 4-1. Standalone Buchik Google Spreadsheets (1 per law subtype selected: Law, Decree, Rule)
    try {
      const normCleanLawName = normalizeLawTitleSpacing(cleanLawName);
      const baseLawName = normCleanLawName.replace(/\s*(?:시행령|시행규칙|법률)$/, '').trim();
      const typesToGenerate: Array<{ key: 'law' | 'decree' | 'rule'; name: string }> = [];

      const hasLaw = sortedRevisions.some((r: any) => {
        const rName = (r.name || '').trim();
        const rSub = (r.subType || '').toLowerCase();
        const rRule = (r.ruleType || '').trim();
        return rSub === 'law' || rRule === '법률' || (!rName.includes('시행령') && !rName.includes('시행규칙') && !rRule.includes('대통령령') && !rRule.includes('부령'));
      });
      const hasDecree = sortedRevisions.some((r: any) => {
        const rName = (r.name || '').trim();
        const rSub = (r.subType || '').toLowerCase();
        const rRule = (r.ruleType || '').trim();
        return rSub === 'decree' || rName.includes('시행령') || rRule.includes('대통령령');
      });
      const hasRule = sortedRevisions.some((r: any) => {
        const rName = (r.name || '').trim();
        const rSub = (r.subType || '').toLowerCase();
        const rRule = (r.ruleType || '').trim();
        return rSub === 'rule' || rName.includes('시행규칙') || rRule.includes('부령');
      });

      if (targetType === 'admrul') {
        typesToGenerate.push({ key: 'law', name: normCleanLawName });
      } else {
        if (hasLaw) typesToGenerate.push({ key: 'law', name: baseLawName });
        if (hasDecree) typesToGenerate.push({ key: 'decree', name: `${baseLawName} 시행령` });
        if (hasRule) typesToGenerate.push({ key: 'rule', name: `${baseLawName} 시행규칙` });
      }

      for (const t of typesToGenerate) {
        const normTypeName = normalizeLawTitleSpacing(t.name);
        const buchikDocTitle = `[${normTypeName}] 부칙`.replace(/[\/\\:*?"<>|]/g, '_');
        const buchikLegacyTitle = `000_[${normTypeName}]_부칙`.replace(/[\/\\:*?"<>|]/g, '_');

        let matchedBuchikExisting =
          existingFileMap.get(buchikDocTitle) ||
          existingFileMap.get(buchikLegacyTitle);

        if (!matchedBuchikExisting) {
          for (const [fName, fData] of existingFileMap.entries()) {
            if (fName.includes(`[${normTypeName}]`) && fName.includes('부칙')) {
              matchedBuchikExisting = fData;
              break;
            }
          }
        }

        if (matchedBuchikExisting) {
          skippedCount++;
          savedSheets.push({
            title: matchedBuchikExisting.name || buchikDocTitle,
            spreadsheetId: matchedBuchikExisting.id,
            url: matchedBuchikExisting.url,
            isExisting: true,
          });
          console.log(`[Unified Buchik Sheet Exists - Skipped] ${buchikDocTitle} -> ${matchedBuchikExisting.name}`);
        } else {
          try {
            const revsOfSub = sortedRevisions.filter((r: any) => {
              const rName = (r.name || '').trim();
              const rSub = (r.subType || '').toLowerCase();
              if (t.key === 'decree') return rSub === 'decree' || rName.includes('시행령');
              if (t.key === 'rule') return rSub === 'rule' || rName.includes('시행규칙');
              return rSub === 'law' || (!rName.includes('시행령') && !rName.includes('시행규칙'));
            });
            const latestRev = sortRevisionsByEnforcementDateDesc(revsOfSub)[0] || revsOfSub[0] || sortedRevisions[0];
            const buchikArticles = await fetchAllBuchikArticlesForLaw(ocKey, t.name, targetType, latestRev, t.key);
            if (buchikArticles.length > 0) {
              const bCreateRes = await callWithRetry(() =>
                drive.files.create({
                  requestBody: {
                    name: buchikDocTitle,
                    mimeType: 'application/vnd.google-apps.spreadsheet',
                    parents: [folderId],
                  },
                  fields: 'id, name, webViewLink',
                })
              );

              const bSpId = bCreateRes.data.id || '';
              const bSpUrl = bCreateRes.data.webViewLink || `https://docs.google.com/spreadsheets/d/${bSpId}/edit`;

              if (bSpId) {
                if (permissionOption?.type === 'anyone') {
                  try {
                    await callWithRetry(() =>
                      drive.permissions.create({
                        fileId: bSpId,
                        requestBody: {
                          role: permissionOption.role || 'reader',
                          type: 'anyone',
                        },
                      })
                    );
                  } catch (pErr) {}
                }

                const bRequests = [
                  // Rename Sheet 0 to 부칙
                  {
                    updateSheetProperties: {
                      properties: { sheetId: 0, title: '부칙' },
                      fields: 'title',
                    },
                  },
                  // Text formatting (TOP align, WRAP)
                  {
                    repeatCell: {
                      range: { sheetId: 0, startRowIndex: 0 },
                      cell: {
                        userEnteredFormat: {
                          verticalAlignment: 'TOP',
                          wrapStrategy: 'WRAP',
                        },
                      },
                      fields: 'userEnteredFormat(verticalAlignment,wrapStrategy)',
                    },
                  },
                  // Navy blue Header format
                  {
                    repeatCell: {
                      range: { sheetId: 0, startRowIndex: 0, endRowIndex: 1 },
                      cell: {
                        userEnteredFormat: {
                          backgroundColor: { red: 0.12, green: 0.18, blue: 0.3 },
                          textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 }, fontSize: 10 },
                          verticalAlignment: 'MIDDLE',
                          horizontalAlignment: 'CENTER',
                        },
                      },
                      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
                    },
                  },
                  // Freeze row 1
                  {
                    updateSheetProperties: {
                      properties: { sheetId: 0, gridProperties: { frozenRowCount: 1 } },
                      fields: 'gridProperties.frozenRowCount',
                    },
                  },
                  // Column widths: 부칙구분(220), 관련법령(200), 조문번호(90), 조문제목(130), 조문내용(550)
                  {
                    updateDimensionProperties: {
                      range: { sheetId: 0, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 },
                      properties: { pixelSize: 220 },
                      fields: 'pixelSize',
                    },
                  },
                  {
                    updateDimensionProperties: {
                      range: { sheetId: 0, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 },
                      properties: { pixelSize: 200 },
                      fields: 'pixelSize',
                    },
                  },
                  {
                    updateDimensionProperties: {
                      range: { sheetId: 0, dimension: 'COLUMNS', startIndex: 2, endIndex: 3 },
                      properties: { pixelSize: 90 },
                      fields: 'pixelSize',
                    },
                  },
                  {
                    updateDimensionProperties: {
                      range: { sheetId: 0, dimension: 'COLUMNS', startIndex: 3, endIndex: 4 },
                      properties: { pixelSize: 130 },
                      fields: 'pixelSize',
                    },
                  },
                  {
                    updateDimensionProperties: {
                      range: { sheetId: 0, dimension: 'COLUMNS', startIndex: 4, endIndex: 5 },
                      properties: { pixelSize: 550 },
                      fields: 'pixelSize',
                    },
                  },
                ];

                await callWithRetry(() =>
                  sheets.spreadsheets.batchUpdate({
                    spreadsheetId: bSpId,
                    requestBody: { requests: bRequests },
                  })
                );

                const buchikHeaders = ['부칙구분', '관련법령', '조문번호', '조문제목', '조문내용'];
                const buchikRows = buchikArticles.map((b) => [
                  b.buchikCategory || '',
                  b.relatedLaw || '',
                  b.articleNo || '',
                  b.articleTitle || '',
                  b.articleContent || '',
                ]);

                await callWithRetry(() =>
                  sheets.spreadsheets.values.batchUpdate({
                    spreadsheetId: bSpId,
                    requestBody: {
                      valueInputOption: 'USER_ENTERED',
                      data: [
                        {
                          range: `'부칙'!A1`,
                          values: [buchikHeaders, ...buchikRows],
                        },
                      ],
                    },
                  })
                );

                newCreatedCount++;
                savedSheets.push({
                  title: buchikDocTitle,
                  spreadsheetId: bSpId,
                  url: bSpUrl,
                  articleCount: buchikArticles.length,
                  isExisting: false,
                });
                console.log(`[Unified Buchik Sheet Created] ${buchikDocTitle} (${buchikArticles.length} entries)`);
              }
            }
          } catch (bErr: any) {
            console.warn(`[Unified Buchik Sheet Creation Warning for ${t.name}]:`, bErr?.message);
          }
        }
      }
    } catch (multiBuchikErr: any) {
      console.warn('[Unified Multi Buchik Generation Warning]:', multiBuchikErr?.message);
    }

    // 5. Process revisions sequentially with controlled pacing and deduplication skip
    for (let idx = 0; idx < sortedRevisions.length; idx++) {
      const rev = sortedRevisions[idx];
      const revDocName = (rev.name || cleanLawName).trim();
      const safePromNo = (rev.promulgationNo || '개정본').trim();
      const safeEnfDate = (rev.enforcementDate || '시행일 미상').trim();
      
      // Standard title without leading index: e.g. "관세법 [시행 2026. 8. 11.] [법률 제21858호, 2026. 8. 11., 일부개정]"
      const docTitle = generateStandardRevisionTitle(rev, cleanLawName, targetType);
      const safePromDate = (rev.promulgationDate || '').trim();
      const dateLabel = safePromDate ? `(${safePromDate} 개정, ${safeEnfDate} 시행)` : `(${safeEnfDate} 시행)`;
      const withoutIndexTitle = `[${revDocName}] ${safePromNo} ${dateLabel}`.replace(/[\/\\:*?"<>|]/g, '_');

      // Comprehensive existing check: exact match OR match without index prefix OR match by promulgation number, enforcement date & law name
      let matchedExisting = existingFileMap.get(docTitle) || existingFileMap.get(withoutIndexTitle);
      if (!matchedExisting && safePromNo && safePromNo !== '개정본' && safeEnfDate && safeEnfDate !== '시행일 미상') {
        const cleanEnfDigits = safeEnfDate.replace(/\D/g, '');
        for (const [fName, fData] of existingFileMap.entries()) {
          const hasProm = fName.includes(safePromNo);
          const hasDoc = fName.includes(revDocName);
          const fNameDigits = fName.replace(/\D/g, '');
          const hasEnf = fName.includes(safeEnfDate) || (cleanEnfDigits.length >= 8 && fNameDigits.includes(cleanEnfDigits));
          if (hasProm && hasDoc && hasEnf) {
            matchedExisting = fData;
            break;
          }
        }
      }

      if (matchedExisting) {
        skippedCount++;
        savedSheets.push({
          title: matchedExisting.name || docTitle,
          spreadsheetId: matchedExisting.id,
          url: matchedExisting.url,
          promulgationNo: rev.promulgationNo,
          enforcementDate: rev.enforcementDate,
          isExisting: true,
        });
        console.log(`[Unified Sheet Exists - Skipped (${idx + 1}/${sortedRevisions.length})] ${docTitle} -> Existing: ${matchedExisting.name}`);
        continue;
      }

      try {
        const isAdmrul =
          rev.targetType === 'admrul' ||
          targetType === 'admrul' ||
          revDocName.includes('규정') ||
          revDocName.includes('고시') ||
          revDocName.includes('훈령') ||
          revDocName.includes('예규');

        // Fetch articles from National Law API using the robust multi-tiered parser
        let revArticles = await fetchArticlesForRevision(ocKey, rev, isAdmrul ? 'admrul' : 'law');

        // Append buchik row to articles list if it's admrul and not already present
        if (isAdmrul && rev.buchikText && (!revArticles.some((a) => a.articleTitle?.includes('부칙') || a.articleNo?.includes('부칙')))) {
          revArticles.push({
            chapterName: '부칙',
            sectionName: '',
            subsectionName: '',
            articleNo: '부칙',
            articleTitle: `${safePromNo} 부칙`,
            articleContent: rev.buchikText,
            effectiveDate: safeEnfDate,
          });
        }

        console.log(`[Export Revision Sheets] (${idx + 1}/${sortedRevisions.length}) '${docTitle}' - parsed articles: ${revArticles.length}`);

        // 1. Create Spreadsheet in Google Drive folder
        const createRes = await callWithRetry(() =>
          drive.files.create({
            requestBody: {
              name: docTitle,
              mimeType: 'application/vnd.google-apps.spreadsheet',
              parents: [folderId],
            },
            fields: 'id, name, webViewLink',
          })
        );

        const spId = createRes.data.id || '';
        const spUrl = createRes.data.webViewLink || `https://docs.google.com/spreadsheets/d/${spId}/edit`;
        if (!spId) continue;

        // Set Public Permission if requested
        if (permissionOption?.type === 'anyone') {
          try {
            await callWithRetry(() =>
              drive.permissions.create({
                fileId: spId,
                requestBody: {
                  role: permissionOption.role || 'reader',
                  type: 'anyone',
                },
              })
            );
          } catch (pErr) {}
        }

        // Setup sheet tabs: [개요] (sheetId: 0), [조문 목록] (sheetId: 1)
        const overviewSheet = `${revDocName.replace(/[\/\\:*?\[\]]/g, '_').slice(0, 20)} 개요`;
        const articlesSheet = '조문 목록';

        const sheetRequests: any[] = [
          // 1. Update Sheet 0 title
          {
            updateSheetProperties: {
              properties: { sheetId: 0, title: overviewSheet },
              fields: 'title',
            },
          },
          // 2. Add Sheet 1 (조문 목록)
          {
            addSheet: {
              properties: { sheetId: 1, title: articlesSheet, index: 1 },
            },
          },
          // 3. Overview sheet text format (TOP align, WRAP)
          {
            repeatCell: {
              range: { sheetId: 0, startRowIndex: 0 },
              cell: {
                userEnteredFormat: {
                  verticalAlignment: 'TOP',
                  wrapStrategy: 'WRAP',
                },
              },
              fields: 'userEnteredFormat(verticalAlignment,wrapStrategy)',
            },
          },
          // 4. Overview header format (Navy blue)
          {
            repeatCell: {
              range: { sheetId: 0, startRowIndex: 2, endRowIndex: 3 },
              cell: {
                userEnteredFormat: {
                  backgroundColor: { red: 0.12, green: 0.18, blue: 0.3 },
                  textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 }, fontSize: 10 },
                  verticalAlignment: 'MIDDLE',
                  horizontalAlignment: 'CENTER',
                },
              },
              fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
            },
          },
          // 5. Articles sheet text format (TOP align, WRAP)
          {
            repeatCell: {
              range: { sheetId: 1, startRowIndex: 0 },
              cell: {
                userEnteredFormat: {
                  verticalAlignment: 'TOP',
                  wrapStrategy: 'WRAP',
                },
              },
              fields: 'userEnteredFormat(verticalAlignment,wrapStrategy)',
            },
          },
          // 6. Articles header format (Navy blue)
          {
            repeatCell: {
              range: { sheetId: 1, startRowIndex: 0, endRowIndex: 1 },
              cell: {
                userEnteredFormat: {
                  backgroundColor: { red: 0.12, green: 0.18, blue: 0.3 },
                  textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 }, fontSize: 10 },
                  verticalAlignment: 'MIDDLE',
                  horizontalAlignment: 'CENTER',
                },
              },
              fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
            },
          },
          // 7. Freeze header row in Articles sheet
          {
            updateSheetProperties: {
              properties: { sheetId: 1, gridProperties: { frozenRowCount: 1 } },
              fields: 'gridProperties.frozenRowCount',
            },
          },
          // 8. Expand column width for article content (500px / 550px)
          {
            updateDimensionProperties: {
              range: { sheetId: 1, dimension: 'COLUMNS', startIndex: 5, endIndex: 6 },
              properties: { pixelSize: isAdmrul ? 550 : 500 },
              fields: 'pixelSize',
            },
          },
          // 9. Set readable column widths for Chapter/Section/Subsection, Article No, Article Title
          {
            updateDimensionProperties: {
              range: { sheetId: 1, dimension: 'COLUMNS', startIndex: 0, endIndex: 3 },
              properties: { pixelSize: 90 },
              fields: 'pixelSize',
            },
          },
          {
            updateDimensionProperties: {
              range: { sheetId: 1, dimension: 'COLUMNS', startIndex: 3, endIndex: 4 },
              properties: { pixelSize: 110 },
              fields: 'pixelSize',
            },
          },
          {
            updateDimensionProperties: {
              range: { sheetId: 1, dimension: 'COLUMNS', startIndex: 4, endIndex: 5 },
              properties: { pixelSize: 150 },
              fields: 'pixelSize',
            },
          },
          ...(isAdmrul
            ? []
            : [
                // Column widths for 전문개정, 본조신설, 제목개정
                {
                  updateDimensionProperties: {
                    range: { sheetId: 1, dimension: 'COLUMNS', startIndex: 6, endIndex: 9 },
                    properties: { pixelSize: 140 },
                    fields: 'pixelSize',
                  },
                },
                // Column width for 비고
                {
                  updateDimensionProperties: {
                    range: { sheetId: 1, dimension: 'COLUMNS', startIndex: 9, endIndex: 10 },
                    properties: { pixelSize: 220 },
                    fields: 'pixelSize',
                  },
                },
              ]),
        ];

        // Combined Single BatchUpdate for Structure & Full Cell Styling to save Sheets Write Quota
        await callWithRetry(() =>
          sheets.spreadsheets.batchUpdate({
            spreadsheetId: spId,
            requestBody: {
              requests: sheetRequests,
            },
          })
        );

        // Populate Overview Values
        const overviewValues = [
          ['항목', '내용'],
          ['법령 / 행정규칙명', revDocName],
          ['공포 / 발령번호', safePromNo],
          ['시행일자', safeEnfDate],
          ['공포 / 발령일자', rev.promulgationDate || '-'],
          ['제개정구분', rev.revisionType || '일부개정'],
          ['소관부처', rev.department || '기획재정부'],
          ['법령구분 / 종류', isAdmrul ? (rev.ruleType ? `행정규칙(${rev.ruleType})` : '행정규칙(고시)') : (rev.ruleType || '법률')],
          ['법령일련번호 (ID/MST)', rev.lawMst || rev.seq || rev.id || '-'],
          ['수록 조문 수', `${revArticles.length}개 조문 (장·절·관 분류)`],
          ['부칙 수록 안내', '별도 부칙 파일(000_[법령명]_부칙)에 전체 부칙이 수록되어 있습니다.'],
          ['저장 폴더', targetFolderName],
          ['생성 일시', new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' })],
        ];

        // Populate Article Rows: 행정규칙은 본조신설 등 항목 제외하고 조문내용까지만 반영
        const articleHeaders = isAdmrul
          ? ['장', '절', '관', '조문번호', '조문제목', '조문내용']
          : [
              '장',
              '절',
              '관',
              '조문번호',
              '조문제목',
              '조문내용',
              '전문개정',
              '본조신설',
              '제목개정',
              '비고',
            ];

        const articleRows = revArticles.map((art) =>
          isAdmrul
            ? [
                art.chapterName || '',
                art.sectionName || '',
                art.subsectionName || '',
                art.articleNo || '',
                art.articleTitle || '',
                art.articleContent || '',
              ]
            : [
                art.chapterName || '',
                art.sectionName || '',
                art.subsectionName || '',
                art.articleNo || '',
                art.articleTitle || '',
                art.articleContent || '',
                art.fullRevision || '',
                art.creation || '',
                art.titleRevision || '',
                art.remarks || (art.isDeleted ? '삭제' : ''),
              ]
        );

        const dataToUpdate: any[] = [
          {
            range: `'${overviewSheet}'!A1`,
            values: overviewValues,
          },
          {
            range: `'${articlesSheet}'!A1`,
            values: [articleHeaders, ...articleRows],
          },
        ];

        // Write cell values in a single call
        await callWithRetry(() =>
          sheets.spreadsheets.values.batchUpdate({
            spreadsheetId: spId,
            requestBody: {
              valueInputOption: 'USER_ENTERED',
              data: dataToUpdate,
            },
          })
        );

        newCreatedCount++;
        savedSheets.push({
          title: docTitle,
          spreadsheetId: spId,
          url: spUrl,
          promulgationNo: rev.promulgationNo,
          enforcementDate: rev.enforcementDate,
          articleCount: revArticles.length,
          isExisting: false,
        });

        // Graceful pacing delay between spreadsheets (800ms) to respect Google Sheets write quotas
        if (idx < sortedRevisions.length - 1) {
          await new Promise((resolve) => setTimeout(resolve, 800));
        }
      } catch (itemErr: any) {
        console.error(`[Unified Sheet Creation Error] ${docTitle}:`, itemErr?.message || itemErr);
      }
    }

    if (savedSheets.length === 0 && sortedRevisions.length > 0) {
      return res.status(500).json({
        success: false,
        error: '구글 드라이브에 시트를 저장하는 중 오류가 발생했습니다. Google 계정 인증 상태 및 드라이브 저장 권한을 확인해 주세요.',
      });
    }

    let summaryMsg = `Google Drive '${targetFolderName}' 폴더에 저장이 완료되었습니다! (신규 생성: ${newCreatedCount}개, 중복 스킵: ${skippedCount}개)`;
    if (skippedCount > 0 && newCreatedCount === 0) {
      summaryMsg = `Google Drive '${targetFolderName}' 폴더에 모든 파일(${skippedCount}개)이 이미 존재하여 저장을 스킵했습니다.`;
    }

    return res.json({
      success: true,
      folder: {
        id: folderId,
        name: targetFolderName,
        url: folderUrl,
        created: !folderSkipped,
        isExisting: folderSkipped,
      },
      savedSheets,
      createdCount: newCreatedCount,
      skippedCount,
      totalCount: savedSheets.length,
      message: summaryMsg,
    });
  } catch (error: any) {
    console.error('Unified Export Revision Sheets Error:', error);
    const errMsg = error.message || 'Google Drive 저장 중 오류가 발생했습니다.';
    const isAuthError =
      error.status === 401 ||
      errMsg.includes('Invalid Credentials') ||
      errMsg.includes('auth') ||
      errMsg.includes('token') ||
      errMsg.includes('UNAUTHENTICATED');
    const isForbidden =
      error.status === 403 ||
      errMsg.includes('PERMISSION_DENIED') ||
      errMsg.includes('insufficient');

    let userFriendlyMsg = errMsg;
    if (isAuthError) {
      userFriendlyMsg = 'Google 인증 토큰이 만료되었거나 유효하지 않습니다. 상단 [Google 로그인] 버튼을 눌러 다시 로그인해 주세요.';
    } else if (isForbidden) {
      userFriendlyMsg = 'Google Drive / Spreadsheets 쓰기 권한이 부족합니다. Google 로그인 시 드라이브 및 스프레드시트 전체 권한을 허용해 주세요.';
    }

    return res.status(error.status || (isAuthError ? 401 : 500)).json({
      success: false,
      error: userFriendlyMsg,
      authError: isAuthError || isForbidden,
    });
  }
});

// API Route: Revoke all external Drive permissions (비공개 전환)
app.post('/api/drive/permissions/revoke', async (req, res) => {
  try {
    const { accessToken, targetId, targetIds } = req.body;
    if (!accessToken) {
      return res.status(401).json({ error: 'Google OAuth 인증 토큰이 필요합니다.' });
    }
    const auth = new google.auth.OAuth2();
    auth.setCredentials({ access_token: accessToken });
    const drive = google.drive({ version: 'v3', auth });

    const allIds = [targetId, ...(Array.isArray(targetIds) ? targetIds : [])].filter(Boolean);
    for (const fId of allIds) {
      try {
        const permRes = await drive.permissions.list({ fileId: fId, fields: 'permissions(id, role, type)' });
        for (const p of permRes.data.permissions || []) {
          if (p.type === 'anyone' && p.id) {
            await drive.permissions.delete({ fileId: fId, permissionId: p.id });
          }
        }
      } catch (pErr: any) {
        console.warn(`[Permission Revoke Warning] File ${fId}:`, pErr?.message);
      }
    }
    return res.json({ success: true, message: '모든 외부 공유 권한이 성공적으로 해제되어 소유자 전용 비공개로 전환되었습니다.' });
  } catch (err: any) {
    console.error('Revoke permission error:', err);
    return res.status(500).json({ error: err.message || '권한 해제 중 오류가 발생했습니다.' });
  }
});


// API Route: Export all 140 revisions as separate CSV files in a ZIP archive
app.post('/api/export/zip-140', async (req, res) => {
  try {
    const ocKey = req.body.ocKey || DEFAULT_OC_KEY;
    console.log('[ZIP Export] Starting retrieval of 140 revisions for ZIP package...');

    let revisionList: any[] = req.body.revisions || [];
    if (!Array.isArray(revisionList) || revisionList.length === 0) {
      revisionList = await fetchAll140Revisions(ocKey);
    }

    if (!revisionList || revisionList.length === 0) {
      return res.status(400).json({ error: '수집된 관세법 개정 이력 데이터를 찾을 수 없습니다.' });
    }

    const zip = new JSZip();
    const folder = zip.folder('관세법_140개_개정자료_개별파일');

    // Add a master summary CSV
    let summaryCsv = '\uFEFF연번,공포번호,시행일자,공포일자,개정구분,법령명,소관부처,MST\n';
    revisionList.forEach((rev, idx) => {
      const escapeCsv = (val: string) => `"${(val || '').replace(/"/g, '""')}"`;
      summaryCsv += `${idx + 1},${escapeCsv(rev.promulgationNo)},${escapeCsv(rev.enforcementDate)},${escapeCsv(rev.promulgationDate)},${escapeCsv(rev.revisionType)},${escapeCsv(rev.lawName || '관세법')},${escapeCsv(rev.department || '기획재정부')},${escapeCsv(rev.lawMst)}\n`;
    });
    folder?.file('000_관세법_전체140건_개정연혁목록.csv', summaryCsv);

    // Fetch details in concurrent chunks of 15
    const chunkSize = 15;
    for (let i = 0; i < revisionList.length; i += chunkSize) {
      const chunk = revisionList.slice(i, i + chunkSize);
      await Promise.all(
        chunk.map(async (rev, chunkIdx) => {
          const indexNum = String(i + chunkIdx + 1).padStart(3, '0');
          try {
            const detailUrl = `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(
              ocKey
            )}&target=law&MST=${encodeURIComponent(rev.lawMst)}&type=XML`;

            const detailRes = await fetch(detailUrl, {
              headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
            });

            if (detailRes.ok) {
              const detailXml = await detailRes.text();
              const parsed = xmlParser.parse(detailXml);
              const root = parsed.법령 || parsed.Law || parsed;
              const revArticles = parseArticlesFromXmlRoot(root);

              let csvContent = '\uFEFF장,절,관,조문번호,조문제목,조문내용,시행일자,비고\n';
              revArticles.forEach((art) => {
                const escapeCsv = (val: string) => `"${(val || '').replace(/"/g, '""')}"`;
                csvContent += `${escapeCsv(art.chapterName)},${escapeCsv(art.sectionName)},${escapeCsv(art.subsectionName)},${escapeCsv(art.articleNo)},${escapeCsv(art.articleTitle)},${escapeCsv(art.articleContent)},${escapeCsv(art.effectiveDate || rev.enforcementDate)},${escapeCsv(art.isDeleted ? '삭제' : '')}\n`;
              });

              const safePromNo = (rev.promulgationNo || '개정본').replace(/[\/\\?%*:|"<>]/g, '_');
              const safeEnfDate = (rev.enforcementDate || '00000000').replace(/\./g, '');
              const filename = `${indexNum}_관세법_${safePromNo}_${safeEnfDate}.csv`;

              folder?.file(filename, csvContent);
            }
          } catch (err: any) {
            console.warn(`[ZIP Export] Error for MST ${rev.lawMst}:`, err?.message);
          }
        })
      );
    }

    const zipBuffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', 'attachment; filename="CustomsAct_140_Revisions_Separate_Files.zip"');
    return res.send(zipBuffer);
  } catch (err: any) {
    console.error('ZIP Export error:', err);
    return res.status(500).json({ error: err.message || 'ZIP 개별 파일 모음 생성에 실패했습니다.' });
  }
});

// Helper: Build XLSX WorkBook Buffer with 4-Column Structured Articles + Overview Sheet with text wrap (줄바꿈 속성)
async function createRevisionWorkbookBuffer(
  rev: any,
  revArticles: any[],
  cleanLawName: string,
  targetFolderName: string = ''
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = '관세법령정보 시스템';
  wb.created = new Date();

  // Sheet 1: 개정기본정보 (Overview)
  const wsOverview = wb.addWorksheet('개정기본정보', {
    properties: { defaultRowHeight: 22 },
    views: [{ showGridLines: true }],
  });

  wsOverview.columns = [
    { header: '항목', key: 'key', width: 24 },
    { header: '내용', key: 'value', width: 85 },
  ];

  // Header row (Row 1)
  const ovHeaderRow = wsOverview.getRow(1);
  ovHeaderRow.values = ['항목', '내용'];
  ovHeaderRow.font = { name: '맑은 고딕', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
  ovHeaderRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF334155' },
  };
  ovHeaderRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  ovHeaderRow.height = 24;

  const overviewItems = [
    { key: '법령/행정규칙명', value: rev.name || cleanLawName },
    { key: '공포/발령번호', value: rev.promulgationNo || '-' },
    { key: '개정일자(공포일)', value: rev.promulgationDate || '-' },
    { key: '시행일자', value: rev.enforcementDate || '-' },
    { key: '제개정구분', value: rev.revisionType || '일부개정' },
    { key: '소관부처', value: rev.department || '기획재정부' },
    { key: '법령일련번호(MST)', value: rev.lawMst || rev.seq || rev.id || '' },
    { key: '수록 조문 수', value: `${revArticles.length}개 조문` },
    { key: '개정 부칙', value: rev.buchikText || '-' },
    { key: '다운로드 일시', value: new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' }) },
  ];

  for (const item of overviewItems) {
    const r = wsOverview.addRow(item);
    r.font = { name: '맑은 고딕', size: 10 };
    r.getCell(1).font = { name: '맑은 고딕', size: 10, bold: true };
    r.getCell(1).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    r.getCell(2).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
  }

  // Thin borders for overview
  wsOverview.eachRow((row, rowNum) => {
    row.eachCell((cell) => {
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      };
    });
  });

  // Determine if this document is an Administrative Rule (행정규칙)
  const isAdmrul =
    rev.targetType === 'admrul' ||
    rev.ruleType === '훈령' ||
    rev.ruleType === '고시' ||
    rev.ruleType === '예규' ||
    rev.ruleType === '공고' ||
    rev.lawType === '행정규칙' ||
    rev.lawType === '고시' ||
    rev.lawType === '훈령' ||
    rev.lawType === '예규' ||
    rev.lawType === '공고' ||
    (rev.name && (rev.name.endsWith('고시') || rev.name.endsWith('훈령') || rev.name.endsWith('예규') || rev.name.endsWith('공고')));

  // Sheet 2: 조문목록 (Articles)
  const wsArticles = wb.addWorksheet('조문목록', {
    properties: { defaultRowHeight: 22 },
    views: [{ state: 'frozen', ySplit: 1, showGridLines: true }],
  });

  // 행정규칙은 본조신설/전문개정/제목개정/비고 항목 제외하고 조문내용까지만 반영
  if (isAdmrul) {
    wsArticles.columns = [
      { header: '장', key: 'chapter', width: 12 },
      { header: '절', key: 'section', width: 12 },
      { header: '관', key: 'subsection', width: 12 },
      { header: '조문번호', key: 'articleNo', width: 14 },
      { header: '조문제목', key: 'articleTitle', width: 22 },
      { header: '조문내용', key: 'articleContent', width: 75 },
    ];
  } else {
    wsArticles.columns = [
      { header: '장', key: 'chapter', width: 12 },
      { header: '절', key: 'section', width: 12 },
      { header: '관', key: 'subsection', width: 12 },
      { header: '조문번호', key: 'articleNo', width: 14 },
      { header: '조문제목', key: 'articleTitle', width: 22 },
      { header: '조문내용', key: 'articleContent', width: 68 },
      { header: '전문개정', key: 'fullRevision', width: 18 },
      { header: '본조신설', key: 'creation', width: 18 },
      { header: '제목개정', key: 'titleRevision', width: 18 },
      { header: '비고', key: 'remarks', width: 26 },
    ];
  }

  // Format Articles Header Row
  const artHeaderRow = wsArticles.getRow(1);
  artHeaderRow.font = { name: '맑은 고딕', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
  artHeaderRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF1E293B' },
  };
  artHeaderRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  artHeaderRow.height = 28;

  // Add Article Rows with wrapText: true on every single cell
  for (const art of revArticles) {
    const rowData = isAdmrul
      ? {
          chapter: art.chapterName || '',
          section: art.sectionName || '',
          subsection: art.subsectionName || '',
          articleNo: art.articleNo || '',
          articleTitle: art.articleTitle || '',
          articleContent: art.articleContent || '',
        }
      : {
          chapter: art.chapterName || '',
          section: art.sectionName || '',
          subsection: art.subsectionName || '',
          articleNo: art.articleNo || '',
          articleTitle: art.articleTitle || '',
          articleContent: art.articleContent || '',
          fullRevision: art.fullRevision || '',
          creation: art.creation || '',
          titleRevision: art.titleRevision || '',
          remarks: art.remarks || (art.isDeleted ? '삭제' : ''),
        };

    const r = wsArticles.addRow(rowData);

    r.font = { name: '맑은 고딕', size: 10 };
    r.eachCell((cell, colNumber) => {
      cell.alignment = {
        vertical: 'top',
        horizontal: isAdmrul
          ? colNumber <= 4
            ? 'center'
            : 'left'
          : colNumber <= 4 || (colNumber >= 7 && colNumber <= 9)
          ? 'center'
          : 'left',
        wrapText: true,
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      };
    });
  }

  // For Administrative Rules (행정규칙), add buchik if present and not in articles
  if (
    isAdmrul &&
    rev.buchikText &&
    !revArticles.some((a: any) => a.articleTitle?.includes('부칙') || a.articleNo?.includes('부칙'))
  ) {
    const buchikRowData = {
      chapter: '부칙',
      section: '',
      subsection: '',
      articleNo: '부칙',
      articleTitle: `부칙 (${rev.promulgationNo || ''})`,
      articleContent: rev.buchikText,
    };

    const r = wsArticles.addRow(buchikRowData);
    r.font = { name: '맑은 고딕', size: 10 };
    r.eachCell((cell) => {
      cell.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      };
    });
  }

  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf);
}

// Helper: Build Standalone Buchik (.xlsx) Workbook Buffer with 5 columns
async function createBuchikWorkbookBuffer(
  cleanLawName: string,
  buchikArticles: BuchikArticle[]
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = '관세법령정보 시스템';
  wb.created = new Date();

  const ws = wb.addWorksheet('부칙', {
    properties: { defaultRowHeight: 22 },
    views: [{ state: 'frozen', ySplit: 1, showGridLines: true }],
  });

  ws.columns = [
    { header: '부칙구분', key: 'buchikCategory', width: 32 },
    { header: '관련법령', key: 'relatedLaw', width: 28 },
    { header: '조문번호', key: 'articleNo', width: 12 },
    { header: '조문제목', key: 'articleTitle', width: 18 },
    { header: '조문내용', key: 'articleContent', width: 75 },
  ];

  const headerRow = ws.getRow(1);
  headerRow.font = { name: '맑은 고딕', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF1E293B' },
  };
  headerRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  headerRow.height = 28;

  for (const b of buchikArticles) {
    const r = ws.addRow({
      buchikCategory: b.buchikCategory || '',
      relatedLaw: b.relatedLaw || '',
      articleNo: b.articleNo || '',
      articleTitle: b.articleTitle || '',
      articleContent: b.articleContent || '',
    });
    r.font = { name: '맑은 고딕', size: 10 };
    r.eachCell((cell, colNumber) => {
      cell.alignment = {
        vertical: 'top',
        horizontal: colNumber <= 4 ? 'center' : 'left',
        wrapText: true,
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      };
    });
  }

  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf);
}

// Helper: Build Master Summary Excel Buffer with text wrap (줄바꿈 속성)
async function createMasterSummaryWorkbookBuffer(
  sortedRevisions: any[],
  cleanLawName: string
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = '관세법령정보 시스템';
  wb.created = new Date();

  const ws = wb.addWorksheet('개정연혁목록', {
    properties: { defaultRowHeight: 22 },
    views: [{ state: 'frozen', ySplit: 3, showGridLines: true }],
  });

  ws.columns = [
    { header: '연번', key: 'no', width: 8 },
    { header: '법령/행정규칙명', key: 'name', width: 26 },
    { header: '법령구분', key: 'type', width: 14 },
    { header: '공포/발령번호', key: 'promNo', width: 22 },
    { header: '개정일자(공포일)', key: 'promDate', width: 16 },
    { header: '시행일자', key: 'enfDate', width: 16 },
    { header: '제개정구분', key: 'revType', width: 14 },
    { header: '소관부처', key: 'dept', width: 16 },
    { header: '일련번호(MST)', key: 'mst', width: 16 },
  ];

  // Title Row
  const normCleanLawName = normalizeLawTitleSpacing(cleanLawName);
  const titleRow = ws.addRow({
    no: `[${normCleanLawName}] 개정연혁 총괄 목록 (총 ${sortedRevisions.length}건)`,
  });
  ws.mergeCells('A1:I1');
  titleRow.font = { name: '맑은 고딕', size: 13, bold: true, color: { argb: 'FF1E293B' } };
  titleRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFF1F5F9' },
  };
  titleRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  titleRow.height = 30;

  // Empty row
  ws.addRow({});

  // Header row
  const headerRow = ws.addRow({
    no: '연번',
    name: '법령/행정규칙명',
    type: '법령구분',
    promNo: '공포/발령번호',
    promDate: '개정일자(공포일)',
    enfDate: '시행일자',
    revType: '제개정구분',
    dept: '소관부처',
    mst: '일련번호(MST)',
  });
  headerRow.font = { name: '맑은 고딕', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF334155' },
  };
  headerRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  headerRow.height = 26;

  // Data rows with wrapText: true
  sortedRevisions.forEach((rev: any, idx: number) => {
    const r = ws.addRow({
      no: idx + 1,
      name: normalizeLawTitleSpacing(rev.name || cleanLawName),
      type:
        rev.ruleType ||
        rev.lawType ||
        (rev.subType === 'decree' ? '대통령령' : rev.subType === 'rule' ? '부령' : '법률'),
      promNo: rev.promulgationNo || '',
      promDate: rev.promulgationDate || '',
      enfDate: rev.enforcementDate || '',
      revType: rev.revisionType || '일부개정',
      dept: rev.department || '기획재정부',
      mst: rev.lawMst || rev.seq || rev.id || '',
    });

    r.font = { name: '맑은 고딕', size: 10 };
    r.eachCell((cell, colNum) => {
      cell.alignment = {
        vertical: 'top',
        horizontal: colNum === 1 || colNum === 3 || colNum === 5 || colNum === 6 || colNum === 7 ? 'center' : 'left',
        wrapText: true,
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      };
    });
  });

  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf);
}

// API Route: Export Revisions directly as Excel ZIP (.xlsx in .zip) without Google Drive
app.post('/api/unified/export-excel-zip', async (req, res) => {
  try {
    const {
      ocKey = DEFAULT_OC_KEY,
      targetType = 'law',
      selectedItem,
      revisions = [],
      folderName,
    } = req.body;

    if (!Array.isArray(revisions) || revisions.length === 0) {
      return res.status(400).json({ error: '다운로드할 개정연혁 목록이 없습니다.' });
    }

    const cleanLawName = normalizeLawTitleSpacing(
      (selectedItem?.name || (targetType === 'admrul' ? '외국환거래규정' : '관세법')).trim()
    );
    const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const defaultFolderName = `[${cleanLawName.replace(/[\/\\:*?"<>|]/g, '_')}_${today}]`;
    const targetFolderName = (folderName || defaultFolderName).trim();

    const sortedRevisions = sortRevisionsByHierarchyAndDate(revisions);
    const zip = new JSZip();
    const zipFolder = zip.folder(targetFolderName) || zip;

    // 1. Add Master Summary Excel with wrapText: true
    const summaryBuffer = await createMasterSummaryWorkbookBuffer(sortedRevisions, cleanLawName);
    zipFolder.file(`[${cleanLawName.replace(/[\/\\:*?"<>|]/g, '_')}] 개정연혁총괄목록.xlsx`, summaryBuffer);

    // 1-1. Add Standalone Buchik Excel files for each selected type (Law, Decree, Rule)
    try {
      const baseLawName = cleanLawName.replace(/\s*(?:시행령|시행규칙|법률)$/, '').trim();
      const typesToGenerate: Array<{ key: 'law' | 'decree' | 'rule'; name: string }> = [];

      const hasLaw = sortedRevisions.some((r: any) => {
        const rName = (r.name || '').trim();
        const rSub = (r.subType || '').toLowerCase();
        const rRule = (r.ruleType || '').trim();
        return rSub === 'law' || rRule === '법률' || (!rName.includes('시행령') && !rName.includes('시행규칙') && !rRule.includes('대통령령') && !rRule.includes('부령'));
      });
      const hasDecree = sortedRevisions.some((r: any) => {
        const rName = (r.name || '').trim();
        const rSub = (r.subType || '').toLowerCase();
        const rRule = (r.ruleType || '').trim();
        return rSub === 'decree' || rName.includes('시행령') || rRule.includes('대통령령');
      });
      const hasRule = sortedRevisions.some((r: any) => {
        const rName = (r.name || '').trim();
        const rSub = (r.subType || '').toLowerCase();
        const rRule = (r.ruleType || '').trim();
        return rSub === 'rule' || rName.includes('시행규칙') || rRule.includes('부령');
      });

      if (targetType === 'admrul') {
        typesToGenerate.push({ key: 'law', name: cleanLawName });
      } else {
        if (hasLaw) typesToGenerate.push({ key: 'law', name: baseLawName });
        if (hasDecree) typesToGenerate.push({ key: 'decree', name: `${baseLawName} 시행령` });
        if (hasRule) typesToGenerate.push({ key: 'rule', name: `${baseLawName} 시행규칙` });
      }

      for (const t of typesToGenerate) {
        const revsOfSub = sortedRevisions.filter((r: any) => {
          const rName = (r.name || '').trim();
          const rSub = (r.subType || '').toLowerCase();
          if (t.key === 'decree') return rSub === 'decree' || rName.includes('시행령');
          if (t.key === 'rule') return rSub === 'rule' || rName.includes('시행규칙');
          return rSub === 'law' || (!rName.includes('시행령') && !rName.includes('시행규칙'));
        });
        const latestRev = sortRevisionsByEnforcementDateDesc(revsOfSub)[0] || revsOfSub[0] || sortedRevisions[0];
        const buchikArticles = await fetchAllBuchikArticlesForLaw(ocKey, t.name, targetType, latestRev, t.key);
        if (buchikArticles.length > 0) {
          const buchikBuffer = await createBuchikWorkbookBuffer(t.name, buchikArticles);
          const normTypeName = normalizeLawTitleSpacing(t.name);
          zipFolder.file(`[${normTypeName.replace(/[\/\\:*?"<>|]/g, '_')}] 부칙.xlsx`, buchikBuffer);
          console.log(`[ZIP Export] Added standalone Buchik file '[${normTypeName}] 부칙.xlsx' (${buchikArticles.length} entries)`);
        }
      }
    } catch (buchikErr: any) {
      console.warn('[ZIP Standalone Buchik Warning]:', buchikErr?.message);
    }

    // 2. Fetch details in concurrency chunks of 15
    const chunkSize = 15;
    for (let i = 0; i < sortedRevisions.length; i += chunkSize) {
      const chunk = sortedRevisions.slice(i, i + chunkSize);
      await Promise.all(
        chunk.map(async (rev: any) => {
          const filename = `${generateStandardRevisionTitle(rev, cleanLawName, targetType)}.xlsx`;

          try {
            const isAdmrul =
              rev.targetType === 'admrul' ||
              targetType === 'admrul' ||
              (rev.name || '').includes('규정') ||
              (rev.name || '').includes('고시');
            const revArticles = await fetchArticlesForRevision(ocKey, rev, isAdmrul ? 'admrul' : 'law');
            const xlsxBuf = await createRevisionWorkbookBuffer(rev, revArticles, cleanLawName, targetFolderName);
            zipFolder.file(filename, xlsxBuf);
          } catch (itemErr: any) {
            console.warn(`[Excel Export Error] ${filename}:`, itemErr?.message);
          }
        })
      );
    }

    const zipBuffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
    const encodedFilename = encodeURIComponent(`${targetFolderName}_엑셀파일모음.zip`);
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', `attachment; filename="${encodedFilename}"; filename*=UTF-8''${encodedFilename}`);
    return res.send(zipBuffer);
  } catch (err: any) {
    console.error('Unified Excel ZIP Export Error:', err);
    return res.status(500).json({ error: err.message || '엑셀 압축 파일 생성 중 오류가 발생했습니다.' });
  }
});

// API Route: Download single revision directly as .xlsx with wrapText: true
app.post('/api/unified/export-single-excel', async (req, res) => {
  try {
    const { ocKey = DEFAULT_OC_KEY, targetType = 'law', revision, cleanLawName = '관세법' } = req.body;
    if (!revision) {
      return res.status(400).json({ error: '다운로드할 개정연혁 정보가 없습니다.' });
    }

    const revDocName = (revision.name || cleanLawName).trim();
    // Standard filename without leading index: e.g. "관세법 [시행 2026. 8. 11.] [법률 제21858호, 2026. 8. 11., 일부개정].xlsx"
    const filename = `${generateStandardRevisionTitle(revision, cleanLawName, targetType)}.xlsx`;

    const isAdmrul =
      revision.targetType === 'admrul' ||
      targetType === 'admrul' ||
      revDocName.includes('규정') ||
      revDocName.includes('고시');
    const revArticles = await fetchArticlesForRevision(ocKey, revision, isAdmrul ? 'admrul' : 'law');
    const xlsxBuf = await createRevisionWorkbookBuffer(revision, revArticles, cleanLawName, '');

    const encodedFilename = encodeURIComponent(filename);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${encodedFilename}"; filename*=UTF-8''${encodedFilename}`);
    return res.send(xlsxBuf);
  } catch (err: any) {
    console.error('Single Excel Export Error:', err);
    return res.status(500).json({ error: err.message || '엑셀 파일 생성 중 오류가 발생했습니다.' });
  }
});

// API Route: Download Standalone Buchik (.xlsx) for a Law directly
app.post('/api/unified/export-buchik-excel', async (req, res) => {
  try {
    const { ocKey = DEFAULT_OC_KEY, targetType = 'law', cleanLawName = '관세법', latestRevision } = req.body;
    const buchikArticles = await fetchAllBuchikArticlesForLaw(ocKey, cleanLawName, targetType, latestRevision);
    if (buchikArticles.length === 0) {
      return res.status(404).json({ error: '해당 법령의 부칙 정보를 찾을 수 없습니다.' });
    }
    const normLawName = normalizeLawTitleSpacing(cleanLawName);
    const xlsxBuf = await createBuchikWorkbookBuffer(normLawName, buchikArticles);
    const filename = `[${normLawName}] 부칙.xlsx`.replace(/[\/\\:*?"<>|]/g, '_');
    const encodedFilename = encodeURIComponent(filename);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${encodedFilename}"; filename*=UTF-8''${encodedFilename}`);
    return res.send(xlsxBuf);
  } catch (err: any) {
    console.error('Standalone Buchik Excel Export Error:', err);
    return res.status(500).json({ error: err.message || '부칙 엑셀 파일 생성 중 오류가 발생했습니다.' });
  }
});

// API Route: Generate Article History (e.g. Article 2 "제2조") across all 140 revisions and save to Google Sheets
app.post('/api/sheets/save-article-history', async (req, res) => {
  try {
    const { accessToken, targetArticleNo = '제2조', listOnly = false } = req.body;
    const ocKey = req.body.ocKey || DEFAULT_OC_KEY;

    if (!accessToken) {
      return res.status(401).json({
        error: '유효한 Google OAuth Access Token이 필요합니다. 상단의 Google 계정 연결 버튼을 눌러주세요.',
      });
    }

    let revisionList: any[] = req.body.revisions || [];
    if (!Array.isArray(revisionList) || revisionList.length === 0) {
      revisionList = await fetchAll140Revisions(ocKey);
    }

    if (!revisionList || revisionList.length === 0) {
      return res.status(400).json({ error: '관세법 개정 이력 데이터를 수집할 수 없습니다.' });
    }

    console.log(`[Article History Export] Fetching history for ${targetArticleNo} (listOnly: ${listOnly}) across ${revisionList.length} revisions...`);

    const articleHistoryRows: any[] = [];

    // Helper for rate limits
    const callApiWithRetry = async <T>(fn: () => Promise<T>, retries = 3, delay = 800): Promise<T> => {
      try {
        return await fn();
      } catch (err: any) {
        if (retries > 0) {
          await new Promise((r) => setTimeout(r, delay));
          return callApiWithRetry(fn, retries - 1, delay * 2);
        }
        throw err;
      }
    };

    // Fetch XML for each revision in small concurrent chunks (10)
    const chunkSize = 10;
    for (let i = 0; i < revisionList.length; i += chunkSize) {
      const chunk = revisionList.slice(i, i + chunkSize);
      await Promise.all(
        chunk.map(async (rev) => {
          try {
            const detailUrl = `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(
              ocKey
            )}&target=law&MST=${encodeURIComponent(rev.lawMst)}&type=XML`;

            const detailRes = await fetch(detailUrl, {
              headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
            });

            if (!detailRes.ok) return;

            const detailXml = await detailRes.text();
            const parsed = xmlParser.parse(detailXml);
            const root = parsed.법령 || parsed.Law || parsed;
            const revArticles = parseArticlesFromXmlRoot(root);

            // Find matching article (e.g. "제2조")
            const targetArt = revArticles.find((art) => {
              const no = (art.articleNo || '').replace(/\s+/g, '');
              const target = targetArticleNo.replace(/\s+/g, '');
              return no === target || (no.startsWith(target) && !no.includes('의'));
            });

            if (targetArt) {
              articleHistoryRows.push({
                promulgationNo: rev.promulgationNo,
                enforcementDate: rev.enforcementDate,
                promulgationDate: rev.promulgationDate,
                revisionType: rev.revisionType,
                articleNo: targetArt.articleNo || targetArticleNo,
                articleTitle: targetArt.articleTitle || '',
                articleContent: targetArt.articleContent || '(조문 내용 없음)',
                isDeleted: targetArt.isDeleted,
                department: rev.department || '기획재정부',
                lawMst: rev.lawMst,
              });
            } else {
              articleHistoryRows.push({
                promulgationNo: rev.promulgationNo,
                enforcementDate: rev.enforcementDate,
                promulgationDate: rev.promulgationDate,
                revisionType: rev.revisionType,
                articleNo: targetArticleNo,
                articleTitle: '미규정/미포함',
                articleContent: '(해당 개정본에는 해당 조문이 포함되어 있지 않거나 삭제된 상태입니다)',
                isDeleted: true,
                department: rev.department || '기획재정부',
                lawMst: rev.lawMst,
              });
            }
          } catch (err: any) {
            console.warn(`[Article History] Error fetching MST ${rev.lawMst}:`, err?.message);
          }
        })
      );
    }

    // Sort history chronologically by enforcement date / promulgation date
    articleHistoryRows.sort((a, b) => {
      const da = (a.enforcementDate || '').replace(/\./g, '');
      const db = (b.enforcementDate || '').replace(/\./g, '');
      return da.localeCompare(db);
    });

    // Detect substantive text changes (문구 실제 변경 여부 자동 감지)
    const normalizeContent = (str: string) => (str || '').replace(/\s+/g, ' ').trim();
    let prevContent = '';
    articleHistoryRows.forEach((row, idx) => {
      const currentNorm = normalizeContent(row.articleContent);
      const prevNorm = normalizeContent(prevContent);

      if (idx === 0) {
        row.isSubstantiveChange = true;
        row.changeNote = '최초 제정/시행';
      } else if (row.isDeleted) {
        row.isSubstantiveChange = prevNorm !== '' && !prevNorm.includes('삭제된 상태');
        row.changeNote = '삭제/미규정';
      } else if (prevNorm !== currentNorm && !currentNorm.includes('삭제된 상태')) {
        row.isSubstantiveChange = true;
        row.changeNote = '⭐ 실질 조문문구 개정 (추가·수정·삭제)';
      } else {
        row.isSubstantiveChange = false;
        row.changeNote = '타조개정에 따른 조문 문구 유지';
      }

      if (!row.isDeleted && currentNorm && !currentNorm.includes('삭제된 상태')) {
        prevContent = row.articleContent;
      }
    });

    const substantiveRows = articleHistoryRows.filter((r) => r.isSubstantiveChange);

    // Filter list if substantiveOnly flag requested
    const targetRows = req.body.substantiveOnly ? substantiveRows : articleHistoryRows;

    // Google Sheets creation
    const auth = new google.auth.OAuth2();
    auth.setCredentials({ access_token: accessToken });
    const sheets = google.sheets({ version: 'v4', auth });

    const docTitle = req.body.substantiveOnly
      ? `[관세법] ${targetArticleNo} 실질 문구 변경이력 전용 (${substantiveRows.length}개 개정본)`
      : listOnly
      ? `[관세법] ${targetArticleNo} 조문별 변경이력 목록 (${articleHistoryRows.length}개 연혁)`
      : `[관세법] ${targetArticleNo} 조문별 변천사 및 개정본별 조문내용 (${articleHistoryRows.length}개 연혁)`;

    const sheet1Title = `${targetArticleNo} 개정이력 개요`;
    const sheet2Title = listOnly
      ? `${targetArticleNo} 전체 변경이력 목록`
      : `${targetArticleNo} 전체 시기별 조문내용`;
    const sheet3Title = `⭐ ${targetArticleNo} 실질 문구 변경건 (${substantiveRows.length}건)`;

    const createRes = await callApiWithRetry(() =>
      sheets.spreadsheets.create({
        requestBody: {
          properties: { title: docTitle },
          sheets: [
            { properties: { title: sheet1Title, index: 0 } },
            { properties: { title: sheet2Title, index: 1 } },
            { properties: { title: sheet3Title, index: 2 } },
          ],
        },
      })
    );

    const spreadsheetId = createRes.data.spreadsheetId;
    if (!spreadsheetId) {
      throw new Error('Google Spreadsheet 생성 실패');
    }

    const firstRow = articleHistoryRows[0];
    const lastRow = articleHistoryRows[articleHistoryRows.length - 1];

    const overviewValues = [
      [`국가법령정보포털 - 관세법 ${targetArticleNo} 조문별 변경이력 및 실질 문구 변천사 DB`],
      [''],
      ['항목', '내용'],
      ['대상 법령', '관세법 (법률)'],
      ['대상 조문', targetArticleNo],
      ['분석 개정본 수', `전체 ${articleHistoryRows.length}개 시기별 개정판`],
      ['⭐ 실질 조문문구 변경 횟수', `${substantiveRows.length}회 (법률 제6305호, 8833호, 10424호, 17649호, 19186호, 19924호 등)`],
      ['실질 변경 공포번호 목록', substantiveRows.map((r) => r.promulgationNo).join(', ')],
      ['최초 제정 당시 시행일', firstRow ? firstRow.enforcementDate : '-'],
      ['최초 제정 당시 조문 제목', firstRow ? firstRow.articleTitle : '-'],
      ['최신 시행일자', lastRow ? lastRow.enforcementDate : '-'],
      ['최신 시행 조문 제목', lastRow ? lastRow.articleTitle : '-'],
      ['분석 생성 일시', new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' })],
      ['참고 사항', '본 구글시트는 관세법 140개 개정본 중 타조개정에 따른 단순 조문유지 건과 실제 조문문구가 추가·수정·삭제된 실질 변경건을 자동 비교 판별하여 분류하였습니다.'],
    ];

    const historyHeaders = listOnly
      ? [
          '연번',
          '시행일자',
          '공포번호',
          '공포일자',
          '개정구분',
          '조문 번호',
          '조문 제목',
          '실질 변경 여부',
          '변경 구분 / 비고',
          '소관부처',
          '법령일련번호 (MST)',
        ]
      : [
          '연번',
          '시행일자',
          '공포번호',
          '공포일자',
          '개정구분',
          '조문 번호',
          '조문 제목',
          `시행 당시 ${targetArticleNo} 조문 전문 (본문 내용)`,
          '실질 변경 여부',
          '변경 구분 / 비고',
          '소관부처',
          '법령일련번호 (MST)',
        ];

    const historyDataRows = targetRows.map((row, idx) =>
      listOnly
        ? [
            idx + 1,
            row.enforcementDate || '',
            row.promulgationNo || '',
            row.promulgationDate || '',
            row.revisionType || '',
            row.articleNo || targetArticleNo,
            row.articleTitle || '',
            row.isSubstantiveChange ? '⭐ 실질 문구 변경' : '단순 조문 유지',
            row.changeNote || '',
            row.department || '기획재정부',
            row.lawMst || '',
          ]
        : [
            idx + 1,
            row.enforcementDate || '',
            row.promulgationNo || '',
            row.promulgationDate || '',
            row.revisionType || '',
            row.articleNo || targetArticleNo,
            row.articleTitle || '',
            row.articleContent || '',
            row.isSubstantiveChange ? '⭐ 실질 문구 변경' : '단순 조문 유지',
            row.changeNote || '',
            row.department || '기획재정부',
            row.lawMst || '',
          ]
    );

    // Sheet 3: Substantive changes only rows
    const substantiveDataRows = substantiveRows.map((row, idx) => [
      idx + 1,
      row.enforcementDate || '',
      row.promulgationNo || '',
      row.promulgationDate || '',
      row.revisionType || '',
      row.articleNo || targetArticleNo,
      row.articleTitle || '',
      row.articleContent || '',
      '⭐ 실질 문구 변경',
      row.changeNote || '',
      row.department || '기획재정부',
      row.lawMst || '',
    ]);

    const substantiveHeaders = [
      '연번',
      '시행일자',
      '공포번호',
      '공포일자',
      '개정구분',
      '조문 번호',
      '조문 제목',
      `시행 당시 ${targetArticleNo} 조문 전문 (실질 변경된 본문)`,
      '실질 변경 여부',
      '변경 구분 / 비고',
      '소관부처',
      '법령일련번호 (MST)',
    ];

    await callApiWithRetry(() =>
      sheets.spreadsheets.values.batchUpdate({
        spreadsheetId,
        requestBody: {
          valueInputOption: 'USER_ENTERED',
          data: [
            { range: `'${sheet1Title}'!A1`, values: overviewValues },
            { range: `'${sheet2Title}'!A1`, values: [historyHeaders, ...historyDataRows] },
            { range: `'${sheet3Title}'!A1`, values: [substantiveHeaders, ...substantiveDataRows] },
          ],
        },
      })
    );

    const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;

    return res.json({
      success: true,
      spreadsheetId,
      spreadsheetUrl,
      targetArticleNo,
      totalCount: articleHistoryRows.length,
      message: listOnly
        ? `관세법 ${targetArticleNo} 조문별 변경이력 목록(${articleHistoryRows.length}개 항목)이 새 구글시트에 성공적으로 생성되었습니다!`
        : `관세법 ${targetArticleNo} 조문별 변경이력 및 시기별 본문 내용(${articleHistoryRows.length}개 개정본)이 새 구글시트에 성공적으로 생성되었습니다!`,
    });
  } catch (err: any) {
    console.error('Article history export error:', err);
    return res.status(500).json({ error: err.message || '조문별 변경이력 구글시트 생성에 실패했습니다.' });
  }
});

// Helper to natural sort Korean law article numbers (e.g. 제1조, 제1조의2, 제2조, 제10조...)
function parseArticleNoKey(noStr: string) {
  if (!noStr) return [99999, 0];
  const match = noStr.match(/제?(\d+)(?:조(?:의(\d+))?)?/);
  if (!match) return [99999, 0];
  const num = parseInt(match[1], 10) || 0;
  const sub = match[2] ? parseInt(match[2], 10) : 0;
  return [num, sub];
}

function sortArticleNos(nos: string[]): string[] {
  return [...nos].sort((a, b) => {
    const [numA, subA] = parseArticleNoKey(a);
    const [numB, subB] = parseArticleNoKey(b);
    if (numA !== numB) return numA - numB;
    return subA - subB;
  });
}

function extractKeywords(str: string): string[] {
  if (!str) return [];
  const words = str.split(/[\sㆍ·,/()_\-]+/).filter((w) => w.length >= 2);
  const keyTerms = [
    '품목분류',
    '과세가격',
    '가격신고',
    '사전심사',
    '사전회시',
    '덤핑방지',
    '상계관세',
    '보복관세',
    '긴급관세',
    '관세환급',
    '보세구역',
    '보세창고',
    '보세공장',
    '통관',
    '체납',
    '납세의무',
    '용어의 뜻',
    '정의',
    '목적',
  ];
  const extracted = new Set<string>(words);
  for (const kt of keyTerms) {
    if (str.includes(kt)) extracted.add(kt);
  }
  return Array.from(extracted);
}

function findMatching1967Article(
  art2000: any,
  articles1967: any[]
): { matchedNo: string; matchedTitle: string; changeType: string; note: string } {
  const t2000Raw = art2000.articleTitle || '';
  const t2000 = t2000Raw.replace(/\s+/g, '').replace(/[ㆍ·,/()_\-]/g, '');
  const no2000 = art2000.articleNo || '';

  if (!articles1967 || articles1967.length === 0) {
    return { matchedNo: '', matchedTitle: '', changeType: '2000년 전부개정 신설', note: '1967년 체계 대비 2000년 전부개정시 새로 신설된 조문 (1)번 공란)' };
  }

  // 1. Exact Title Match
  if (t2000) {
    const exactTitleMatch = articles1967.find((p) => {
      const pTitle = (p.articleTitle || '').replace(/\s+/g, '').replace(/[ㆍ·,/()_\-]/g, '');
      return pTitle && pTitle === t2000;
    });

    if (exactTitleMatch) {
      const isSameNo = exactTitleMatch.articleNo === no2000;
      return {
        matchedNo: exactTitleMatch.articleNo,
        matchedTitle: exactTitleMatch.articleTitle || '(제목없음)',
        changeType: isSameNo ? '동일 조문유지' : '조문번호 위치 이동',
        note: isSameNo
          ? `1)번 ${exactTitleMatch.articleNo} (${exactTitleMatch.articleTitle})와 동일한 조문번호/제목 유지`
          : `1)번 ${exactTitleMatch.articleNo} (${exactTitleMatch.articleTitle}) -> 2)번 ${no2000} (${t2000Raw})로 조문번호 이동`,
      };
    }
  }

  // 2. Keyword / Substring Match (e.g. 품목분류의 사전회시등 vs 특정물품에 적용될 품목분류의 사전심사)
  if (t2000 && t2000.length >= 2) {
    const keywords2000 = extractKeywords(t2000Raw);
    let bestCandidate: any = null;
    let maxScore = 0;

    for (const cand of articles1967) {
      const candTitle = cand.articleTitle || '';
      if (!candTitle) continue;

      const keywords1967 = extractKeywords(candTitle);
      const overlapCount = keywords2000.filter((kw) => keywords1967.includes(kw) || candTitle.includes(kw) || t2000Raw.includes(kw)).length;

      let score = overlapCount;

      // Special Domain Pairs
      const bothPum = t2000Raw.includes('품목분류') && candTitle.includes('품목분류');
      const bothSajeon = (t2000Raw.includes('사전심사') || t2000Raw.includes('사전')) && (candTitle.includes('사전회시') || candTitle.includes('사전'));
      if (bothPum && bothSajeon) score += 10;
      else if (bothPum) score += 5;

      const bothGagyeok = t2000Raw.includes('가격신고') && candTitle.includes('가격신고');
      if (bothGagyeok) score += 10;

      const bothGwase = t2000Raw.includes('과세가격') && candTitle.includes('과세가격');
      if (bothGwase) score += 6;

      if (score > maxScore && score >= 2) {
        maxScore = score;
        bestCandidate = cand;
      }
    }

    if (bestCandidate) {
      return {
        matchedNo: bestCandidate.articleNo,
        matchedTitle: bestCandidate.articleTitle || '(제목없음)',
        changeType: '조문제목 수정/개정',
        note: `1)번 ${bestCandidate.articleNo} (${bestCandidate.articleTitle}) -> 2)번 ${no2000} (${t2000Raw})로 수정/변경`,
      };
    }
  }

  // 3. Exact Article Number Match (if title is general)
  const sameNoArt = articles1967.find((p) => p.articleNo === no2000);
  if (sameNoArt && sameNoArt.articleTitle) {
    return {
      matchedNo: sameNoArt.articleNo,
      matchedTitle: sameNoArt.articleTitle,
      changeType: '동일 조문번호 (제목 변경)',
      note: `1)번 ${sameNoArt.articleNo} (${sameNoArt.articleTitle}) 대비 2)번 ${no2000} (${t2000Raw}) 제목 변경`,
    };
  }

  // 4. No Match -> Leave 1967 columns BLANK
  return {
    matchedNo: '',
    matchedTitle: '',
    changeType: '2000년 전부개정 신설',
    note: '1967년 체계 대비 2000년 전부개정시 새로 신설된 조문 (1)번 공란)',
  };
}

// API Route: Wholly Amended Laws Comparison (1967 Act No. 1976 & 2000 Act No. 6305)
app.post('/api/sheets/save-wholly-amended-comparison', async (req, res) => {
  try {
    const { accessToken } = req.body;
    const ocKey = req.body.ocKey || DEFAULT_OC_KEY;

    if (!accessToken) {
      return res.status(401).json({
        error: '유효한 Google OAuth Access Token이 필요합니다. Google 계정을 먼저 연결해 주세요.',
      });
    }

    let revisionList: any[] = req.body.revisions || [];
    if (!Array.isArray(revisionList) || revisionList.length === 0) {
      revisionList = await fetchAll140Revisions(ocKey);
    }

    if (!revisionList || revisionList.length === 0) {
      return res.status(400).json({ error: '관세법 개정 이력 데이터를 수집할 수 없습니다.' });
    }

    // Sort chronologically
    revisionList.sort((a, b) => {
      const da = (a.promulgationDate || a.enforcementDate || '').replace(/\./g, '');
      const db = (b.promulgationDate || b.enforcementDate || '').replace(/\./g, '');
      return da.localeCompare(db);
    });

    // Find 1967 Wholly Amended Law (제1976호) and 2000 Wholly Amended Law (제6305호)
    const idx1967 = revisionList.findIndex(
      (r) => (r.promulgationNo || '').includes('1976') || (r.promulgationDate || '').startsWith('1967')
    );

    const idx2000 = revisionList.findIndex(
      (r) => (r.promulgationNo || '').includes('6305') || (r.promulgationDate || '').startsWith('2000')
    );

    const start1967Idx = idx1967 >= 0 ? idx1967 : 0;
    const start2000Idx = idx2000 >= 0 ? idx2000 : revisionList.findIndex((r) => (r.promulgationDate || '') >= '2000.12.29');

    // Period 1 Revisions: 1967년 제1976호 ~ 2000년 제6305호 직전
    const revs1967 = revisionList.slice(start1967Idx, start2000Idx > start1967Idx ? start2000Idx : revisionList.length);

    // Period 2 Revisions: 2000년 제6305호 ~ 현재
    const revs2000 = revisionList.slice(start2000Idx >= 0 ? start2000Idx : 0);

    // Pre-2000 revision (immediately before 제6305호)
    const revPrev = start2000Idx > 0 ? revisionList[start2000Idx - 1] : revs1967[revs1967.length - 1];
    const rev6305 = revisionList[start2000Idx];

    console.log(`[Wholly Amended] Period 1 (1967~2000): ${revs1967.length} revs, Period 2 (2000~): ${revs2000.length} revs`);

    // Fetch XML articles for all revisions in chunks
    const allRevsToFetch = Array.from(new Set([...revs1967, ...revs2000, revPrev, rev6305].filter(Boolean)));
    const articlesMap = new Map<string, any[]>();

    const chunkSize = 8;
    for (let i = 0; i < allRevsToFetch.length; i += chunkSize) {
      const chunk = allRevsToFetch.slice(i, i + chunkSize);
      await Promise.all(
        chunk.map(async (rev) => {
          try {
            const detailUrl = `http://www.law.go.kr/DRF/lawService.do?OC=${encodeURIComponent(
              ocKey
            )}&target=law&MST=${encodeURIComponent(rev.lawMst)}&type=XML`;

            const detailRes = await fetch(detailUrl, {
              headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
            });

            if (!detailRes.ok) return;

            const detailXml = await detailRes.text();
            const parsed = xmlParser.parse(detailXml);
            const root = parsed.법령 || parsed.Law || parsed;
            const articles = parseArticlesFromXmlRoot(root);
            articlesMap.set(rev.lawMst, articles);
          } catch (err: any) {
            console.warn(`[Wholly Amended Fetch] Error fetching MST ${rev.lawMst}:`, err?.message);
          }
        })
      );
    }

    // 1) Build Matrix for Period 2: 2000년 전부개정 계열 (법률 제21208호 ~ 제6305호)
    const revs2000Desc = [...revs2000].reverse(); // Reverse chronological (newest law 21208 first)

    const articleNos2000Set = new Set<string>();
    revs2000Desc.forEach((r) => {
      const arts = articlesMap.get(r.lawMst) || [];
      arts.forEach((a) => {
        if (a.articleNo) articleNos2000Set.add(a.articleNo);
      });
    });

    const sortedArticleNos2000 = sortArticleNos(Array.from(articleNos2000Set));

    const headers2000 = [
      '연번',
      '조문번호',
      ...revs2000Desc.map((r) => `${r.promulgationNo || '법률'}(${r.promulgationDate || r.enforcementDate || ''})`),
    ];

    const rows2000 = sortedArticleNos2000.map((artNo, idx) => {
      const rowVals: string[] = [];

      for (let c = 0; c < revs2000Desc.length; c++) {
        const rev = revs2000Desc[c];
        const arts = articlesMap.get(rev.lawMst) || [];
        const found = arts.find((a) => a.articleNo === artNo);

        let textVal = '';
        if (found) {
          if (found.isDeleted) {
            textVal = `[삭제] ${found.articleTitle || ''}`.trim();
          } else {
            textVal = found.articleTitle ? `${artNo} (${found.articleTitle})` : artNo;
          }
        } else {
          textVal = '-';
        }

        if (c === 0) {
          // First law column (latest law e.g. 21208) shows baseline text
          rowVals.push(textVal);
        } else {
          // Compare with previous (newer) law column (c - 1)
          const prevRev = revs2000Desc[c - 1];
          const prevArts = articlesMap.get(prevRev.lawMst) || [];
          const prevFound = prevArts.find((a) => a.articleNo === artNo);

          let prevTextVal = '';
          if (prevFound) {
            if (prevFound.isDeleted) {
              prevTextVal = `[삭제] ${prevFound.articleTitle || ''}`.trim();
            } else {
              prevTextVal = prevFound.articleTitle ? `${artNo} (${prevFound.articleTitle})` : artNo;
            }
          } else {
            prevTextVal = '-';
          }

          // If UNCHANGED compared to adjacent newer law, leave BLANK
          if (textVal === prevTextVal) {
            rowVals.push('');
          } else {
            rowVals.push(textVal);
          }
        }
      }

      return [idx + 1, artNo, ...rowVals];
    });

    // 2) Build Matrix for Period 1: 1967년 전부개정 계열 (법률 제6136호 ~ 제2062호/제1976호)
    const revs1967Desc = [...revs1967].reverse(); // Reverse chronological (newest 6136 first)

    const articleNos1967Set = new Set<string>();
    revs1967Desc.forEach((r) => {
      const arts = articlesMap.get(r.lawMst) || [];
      arts.forEach((a) => {
        if (a.articleNo) articleNos1967Set.add(a.articleNo);
      });
    });

    const sortedArticleNos1967 = sortArticleNos(Array.from(articleNos1967Set));

    const headers1967 = [
      '연번',
      '조문번호',
      ...revs1967Desc.map((r) => `${r.promulgationNo || '법률'}(${r.promulgationDate || r.enforcementDate || ''})`),
    ];

    const rows1967 = sortedArticleNos1967.map((artNo, idx) => {
      const rowVals: string[] = [];

      for (let c = 0; c < revs1967Desc.length; c++) {
        const rev = revs1967Desc[c];
        const arts = articlesMap.get(rev.lawMst) || [];
        const found = arts.find((a) => a.articleNo === artNo);

        let textVal = '';
        if (found) {
          if (found.isDeleted) {
            textVal = `[삭제] ${found.articleTitle || ''}`.trim();
          } else {
            textVal = found.articleTitle ? `${artNo} (${found.articleTitle})` : artNo;
          }
        } else {
          textVal = '-';
        }

        if (c === 0) {
          // First law column (latest 1967-era law e.g. 6136) shows baseline text
          rowVals.push(textVal);
        } else {
          const prevRev = revs1967Desc[c - 1];
          const prevArts = articlesMap.get(prevRev.lawMst) || [];
          const prevFound = prevArts.find((a) => a.articleNo === artNo);

          let prevTextVal = '';
          if (prevFound) {
            if (prevFound.isDeleted) {
              prevTextVal = `[삭제] ${prevFound.articleTitle || ''}`.trim();
            } else {
              prevTextVal = prevFound.articleTitle ? `${artNo} (${prevFound.articleTitle})` : artNo;
            }
          } else {
            prevTextVal = '-';
          }

          // If UNCHANGED, leave BLANK
          if (textVal === prevTextVal) {
            rowVals.push('');
          } else {
            rowVals.push(textVal);
          }
        }
      }

      return [idx + 1, artNo, ...rowVals];
    });

    // 3) Build Sheet 3: 현재 최신 법률(제21208호) 조문 기준 vs 1967년 체계(법률 제6136호) 조문 대조 비교
    const latest2000Rev = revs2000Desc[0] || rev6305;
    const articles21208 = latest2000Rev ? articlesMap.get(latest2000Rev.lawMst) || [] : [];

    // Collect all unique articles from pre-2000 revisions for complete matching pool
    const allArticles1967Map = new Map<string, any>();
    revs1967.forEach((r) => {
      const arts = articlesMap.get(r.lawMst) || [];
      arts.forEach((a) => {
        if (a.articleNo && a.articleTitle) {
          allArticles1967Map.set(`${a.articleNo}_${a.articleTitle}`, a);
        }
      });
    });
    const allArticles1967List = Array.from(allArticles1967Map.values());

    const comparisonHeaders = [
      '연번',
      '현재 법률(제21208호) 조문번호',
      '현재 법률(제21208호) 조문제목',
      '1967년 체계(법률 제6136호) 대조 조문번호',
      '1967년 체계(법률 제6136호) 대조 조문제목',
      '전부개정 변경/수정 유형',
      '비고 및 대조 상세 설명',
    ];

    const comparisonRows = articles21208.map((art21208, idx) => {
      const matchResult = findMatching1967Article(art21208, allArticles1967List);

      return [
        idx + 1,
        art21208.articleNo,
        art21208.articleTitle || '(제목없음)',
        matchResult.matchedNo || '', // BLANK if newly established
        matchResult.matchedTitle || '', // BLANK if newly established
        matchResult.changeType,
        matchResult.note,
      ];
    });

    // Overview Sheet
    const overviewValues = [
      ['관세법 전부개정(1967년 제1976호~제6136호 & 2000년 제6305호~제21208호) 조문제목 변천 및 대조 분석 DB'],
      [''],
      ['분석 구분', '내용'],
      ['대상 법령', '관세법 (법률)'],
      ['현재 법률 기준', `최신 법률 제21208호 (${latest2000Rev?.promulgationDate || latest2000Rev?.enforcementDate || ''} 공포/시행)`],
      ['직전 체계 법률 기준', `2000년 전부개정 직전 법률 제6136호 계열 (${revPrev?.promulgationDate || ''} 공포)`],
      ['시트 1 구성', '2000년 전부개정 계열: 최신 법률(제21208호)을 첫 열로 배치하여 제6305호까지 역순 비교 (미변경시 공란)'],
      ['시트 2 구성', '1967년 전부개정 계열: 직전 법률(제6136호)을 첫 열로 배치하여 제2062호/제1976호까지 역순 비교 (미변경시 공란)'],
      ['시트 3 구성', '현재 법률(제21208호) 조문 기준으로 직전 법률(제6136호) 조문과 대조 (신설 조문은 제6136호 열 공란)'],
      ['분석 생성 일시', new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' })],
    ];

    // Google Sheets API call
    const auth = new google.auth.OAuth2();
    auth.setCredentials({ access_token: accessToken });
    const sheets = google.sheets({ version: 'v4', auth });

    const docTitle = `[관세법] 전부개정 조문제목 변천사 (최신 제21208호 ~ 제6305호 / 제6136호 ~ 제2062호) & 대조표`;

    const sheet0Title = `전부개정 분석 개요`;
    const sheet1Title = `1) 2000년 전부개정(제21208호~제6305호)`;
    const sheet2Title = `2) 1967년 전부개정(제6136호~제2062호)`;
    const sheet3Title = `3) 최신(제21208호) vs 직전(제6136호) 대조`;

    const createRes = await sheets.spreadsheets.create({
      requestBody: {
        properties: { title: docTitle },
        sheets: [
          { properties: { title: sheet0Title, index: 0 } },
          { properties: { title: sheet1Title, index: 1 } },
          { properties: { title: sheet2Title, index: 2 } },
          { properties: { title: sheet3Title, index: 3 } },
        ],
      },
    });

    const spreadsheetId = createRes.data.spreadsheetId;
    if (!spreadsheetId) {
      throw new Error('Google Spreadsheet 생성에 실패했습니다.');
    }

    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: {
        valueInputOption: 'USER_ENTERED',
        data: [
          { range: `'${sheet0Title}'!A1`, values: overviewValues },
          { range: `'${sheet1Title}'!A1`, values: [headers2000, ...rows2000] },
          { range: `'${sheet2Title}'!A1`, values: [headers1967, ...rows1967] },
          { range: `'${sheet3Title}'!A1`, values: [comparisonHeaders, ...comparisonRows] },
        ],
      },
    });

    const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;

    return res.json({
      success: true,
      spreadsheetId,
      spreadsheetUrl,
      message: `관세법 전부개정(1967년 제1976호, 2000년 제6305호) 조문제목 변천사 및 대조 구글시트가 성공적으로 생성되었습니다!`,
    });
  } catch (err: any) {
    console.error('Wholly amended comparison export error:', err);
    return res.status(500).json({ error: err.message || '전부개정 구글시트 생성에 실패했습니다.' });
  }
});

// Clean HTML text helper converting linebreaks to spaces
function cleanHtmlText(str: string): string {
  if (!str) return '';
  return str
    .replace(/<br\s*\/?>/gi, ' ')
    .replace(/<\/p>/gi, ' ')
    .replace(/<\/td>/gi, ' ')
    .replace(/<\/tr>/gi, ' ')
    .replace(/<[^>]+>/g, '')
    .replace(/[\r\n\t]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

// Helper to fetch UNIPASS Decision Cases with full itemDesc & summary for a specific year
async function fetchUnipassYearDecisions(
  year: string = '2026',
  queryKeyword: string = '',
  maxItems: number = 0,
  customStDt?: string,
  customEdDt?: string
) {
  try {
    const listUrl = 'https://unipass.customs.go.kr/clip/prlstclsfsrch/retrieveDmstPrlstClsfCaseLst2.do';

    // Case types in UNIPASS domestic classification:
    // 01: 품목분류사례
    // 03: 협의회결정사항 (관세품목분류협의회)
    // 04: 위원회결정사항 (관세품목분류위원회)
    const caseTypes = [
      { code: '01', defaultCategory: '품목분류사례' },
      { code: '03', defaultCategory: '협의회결정사항' },
      { code: '04', defaultCategory: '위원회결정사항' },
    ];

    async function getListPage(pageIndex: number, dateFmt: 'hyphen' | 'nodash' | 'dot', caseTpcd: string, retries = 8): Promise<any> {
      let stDt = customStDt || `${year}-01-01`;
      let edDt = customEdDt || `${year}-12-31`;

      if (dateFmt === 'nodash') {
        stDt = stDt.replace(/[-.]/g, '');
        edDt = edDt.replace(/[-.]/g, '');
      } else if (dateFmt === 'dot') {
        stDt = stDt.replace(/-/g, '.');
        edDt = edDt.replace(/-/g, '.');
      }

      const params = new URLSearchParams({
        prlstClsfCaseTpcd: caseTpcd,
        rrdcNo: '',
        srchYn: 'Y',
        scrnTp: 'WDTH',
        sortColm: 'ENFR_DT',
        sortOrdr: 'DESC',
        atntSrchTp: '',
        docId: '',
        scrnId: 'UI-ULS-0203-002S',
        reffNo: '',
        dtrmHsSgn: '',
        stDt,
        edDt,
        cmdtNm: '',
        cmdtDesc: '',
        dtrmRsnCn: '',
        srwr: queryKeyword && queryKeyword !== '관세' ? queryKeyword : '',
        initPageIndex: '1',
        pageIndex: String(pageIndex),
        pagePerRecord: '10',
        recordCountPerPage: '10',
      });

      for (let attempt = 1; attempt <= retries; attempt++) {
        try {
          const res = await fetch(listUrl, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
              'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
              'X-Requested-With': 'XMLHttpRequest',
            },
            body: params.toString(),
            signal: AbortSignal.timeout(12000),
          });
          const text = await res.text();
          if (text.trim().startsWith('{')) {
            const json = JSON.parse(text);
            if (json?.uls_dmst?.itemList || json?.uls_dmst?.thisTotalCount !== undefined) {
              return json.uls_dmst || {};
            }
          }
        } catch (e) {
          // Retry on timeout or network glitch
        }
        if (attempt < retries) {
          await new Promise((r) => setTimeout(r, 250 * attempt));
        }
      }
      return {};
    }

    let allRawItems: { item: any; caseTpcd: string; defaultCategory: string }[] = [];

    // Fetch for each case type (01: 품목분류, 03: 협의회, 04: 위원회)
    for (const ct of caseTypes) {
      let dateFmtUsed: 'hyphen' | 'nodash' | 'dot' = 'hyphen';
      let firstPage = await getListPage(1, 'hyphen', ct.code);
      let totalCount = parseInt(firstPage.thisTotalCount || '0', 10);

      if (totalCount === 0) {
        firstPage = await getListPage(1, 'nodash', ct.code);
        totalCount = parseInt(firstPage.thisTotalCount || '0', 10);
        if (totalCount > 0) dateFmtUsed = 'nodash';
      }
      if (totalCount === 0) {
        firstPage = await getListPage(1, 'dot', ct.code);
        totalCount = parseInt(firstPage.thisTotalCount || '0', 10);
        if (totalCount > 0) dateFmtUsed = 'dot';
      }

      if (totalCount === 0) continue;

      let totalPages = Math.ceil(totalCount / 10);
      if (maxItems > 0) {
        totalPages = Math.min(totalPages, Math.ceil(maxItems / 10));
      }

      const pageMap = new Map<number, any[]>();
      if (firstPage?.itemList && Array.isArray(firstPage.itemList)) {
        pageMap.set(1, firstPage.itemList);
      }

      let missingPages = Array.from({ length: totalPages - 1 }, (_, i) => i + 2);

      let fetchRound = 0;
      while (missingPages.length > 0 && fetchRound < 10) {
        fetchRound++;
        const pageBatchSize = 5;
        for (let i = 0; i < missingPages.length; i += pageBatchSize) {
          const chunk = missingPages.slice(i, i + pageBatchSize);
          const results = await Promise.all(
            chunk.map((p) => getListPage(p, dateFmtUsed, ct.code, 8))
          );
          chunk.forEach((p, idx) => {
            const r = results[idx];
            if (r && r.itemList && Array.isArray(r.itemList) && r.itemList.length > 0) {
              pageMap.set(p, r.itemList);
            }
          });
        }

        const stillMissing: number[] = [];
        for (let p = 1; p <= totalPages; p++) {
          if (!pageMap.has(p)) stillMissing.push(p);
        }
        missingPages = stillMissing;
        if (missingPages.length > 0) {
          console.warn(`[UNIPASS] Year ${year} code ${ct.code}: missing ${missingPages.length} pages, retrying round ${fetchRound}...`);
          await new Promise((r) => setTimeout(r, 500));
        }
      }

      let ctItems: any[] = [];
      for (let p = 1; p <= totalPages; p++) {
        if (pageMap.has(p)) {
          ctItems.push(...pageMap.get(p)!);
        }
      }

      if (maxItems > 0 && ctItems.length > maxItems) {
        ctItems = ctItems.slice(0, maxItems);
      }

      for (const item of ctItems) {
        allRawItems.push({ item, caseTpcd: ct.code, defaultCategory: ct.defaultCategory });
      }
    }

    // UNIPASS API handles date range filtering on server side (stDt to edDt).
    // Keep all returned raw items without secondary string pruning to ensure no missing cases.
    const validRawItems = allRawItems;

    console.log(`[UNIPASS] Year ${year}: Collected ${validRawItems.length} total cases (from UNIPASS API).`);

    if (validRawItems.length === 0) return [];

    // Clean HTML text helper preserving newlines for detail pages
    function cleanDtlText(rawTd: string): string {
      if (!rawTd) return '';
      let clean = rawTd.replace(/<br\s*\/?>/gi, '\n');
      clean = clean.replace(/<p[^>]*>/gi, '\n').replace(/<\/p>/gi, '\n');
      clean = clean.replace(/<[^>]+>/g, '');
      clean = clean
        .replace(/&nbsp;/gi, ' ')
        .replace(/&lt;/gi, '<')
        .replace(/&gt;/gi, '>')
        .replace(/&amp;/gi, '&')
        .replace(/&quot;/gi, '"')
        .replace(/&#39;/g, "'");
      const lines = clean.split('\n').map((l) => l.trim()).filter(Boolean);
      return lines.join('\n');
    }

    // Function to parse UNIPASS detail HTML
    function parseDtlHtml(html: string) {
      function extractTdByTh(thText: string) {
        const thRegex = new RegExp('<th[^>]*>\\s*' + thText + '\\s*<\\/th>', 'i');
        const match = thRegex.exec(html);
        if (!match) return '';
        const thIdx = match.index;
        const tdStart = html.indexOf('<td', thIdx);
        if (tdStart === -1) return '';
        const contentStart = html.indexOf('>', tdStart) + 1;
        const tdEnd = html.indexOf('</td>', contentStart);
        if (tdEnd === -1) return '';
        const rawTd = html.substring(contentStart, tdEnd);
        return cleanDtlText(rawTd);
      }

      const itemDesc = extractTdByTh('물품설명') || extractTdByTh('안건요지') || extractTdByTh('품명 및 물품설명');
      const summary = extractTdByTh('결정사유') || extractTdByTh('결정요지') || extractTdByTh('주요결정요지') || extractTdByTh('의결내용');
      return { itemDesc, summary };
    }

    async function fetchUnipassCaseFullDetail(rrdcNo: string, caseTpcd: string, retries = 3) {
      if (!rrdcNo) return null;
      const dtlEndpoints: Record<string, { url: string; mttrTpcd: string }> = {
        '01': { url: 'https://unipass.customs.go.kr/clip/prlstclsfsrch/retrieveDmstPrlstClsfCaseDtl.do', mttrTpcd: '' },
        '03': { url: 'https://unipass.customs.go.kr/clip/prlstclsfsrch/cncidtrm/retrieveDmstPrlstClsfCaseDtl2.do', mttrTpcd: '02' },
        '04': { url: 'https://unipass.customs.go.kr/clip/prlstclsfsrch/cmitdtrm/retrieveDmstPrlstClsfCaseDtl2.do', mttrTpcd: '01' },
      };
      const conf = dtlEndpoints[caseTpcd] || dtlEndpoints['01'];
      for (let attempt = 1; attempt <= retries; attempt++) {
        try {
          const params = new URLSearchParams({ rrdcNo });
          if (conf.mttrTpcd) params.append('mttrTpcd', conf.mttrTpcd);

          const res = await fetch(conf.url, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
              'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
            },
            body: params.toString(),
            signal: AbortSignal.timeout(8000),
          });
          if (res.ok) {
            const html = await res.text();
            const dtl = parseDtlHtml(html);
            if (dtl.itemDesc || dtl.summary) return dtl;
          }
        } catch (e) {
          // Retry
        }
        if (attempt < retries) {
          await new Promise((r) => setTimeout(r, 150 * attempt));
        }
      }
      return null;
    }

    // Process details
    async function fetchDetail({ item, caseTpcd, defaultCategory }: { item: any; caseTpcd: string; defaultCategory: string }) {
      const dateRaw = String(item.ENFR_DT || `${year}0101`);
      const itemYear = dateRaw.length >= 4 ? dateRaw.substring(0, 4) : year;
      const formattedDate =
        dateRaw.length === 8
          ? `${dateRaw.substring(0, 4)}.${dateRaw.substring(4, 6)}.${dateRaw.substring(6, 8)}`
          : dateRaw;

      let cleanDesc = cleanHtmlText(item.CMDT_DESC || item.CMDT_DESC_TIT || '');
      let cleanRsn = cleanHtmlText(item.DTRM_RSN_CN || item.DTRM_RSN_CN_TIT || '');

      const rrdcNo = item.RRDC_NO || item.DOCID || '';
      if (rrdcNo) {
        const dtl = await fetchUnipassCaseFullDetail(rrdcNo, caseTpcd);
        if (dtl) {
          if (dtl.itemDesc) cleanDesc = dtl.itemDesc;
          if (dtl.summary) cleanRsn = dtl.summary;
        }
      }

      const rawTpnm = item.PRLST_CLSF_CASE_TPNM || item.prlstClsfCaseTpnm || '';
      let category = defaultCategory;
      if (rawTpnm) {
        category = cleanHtmlText(rawTpnm);
      } else if (caseTpcd === '04') {
        category = '위원회결정사항';
      } else if (caseTpcd === '03') {
        category = '협의회결정사항';
      } else if (caseTpcd === '01') {
        category = '품목분류사례';
      }

      return {
        id: item.DOCID || item.REFF_NO || item.RRDC_NO || `UNIPASS-${itemYear}-${Math.random()}`,
        year: itemYear,
        targetType: 'unipass_clip',
        caseNo: item.REFF_NO || item.RRDC_NO || '품목분류사례',
        title: item.CMDT_NM || item.CMDT_NM_TIT || '품목분류 결정물품',
        decisionDate: formattedDate,
        department: item.CSTM_NM || item.CSTM_NM_TIT || '관세평가분류원',
        relLaw: `HS부호: ${item.DTRM_HS_SGN || item.DTRM_HS_SGN_TIT || '미지정'}`,
        itemDesc: cleanDesc || '물품설명 정보 없음',
        summary: cleanRsn || '결정사유 상세내용 없음',
        category,
      };
    }

    const dtlBatchSize = 25;
    let finalDetailedList: any[] = [];
    for (let i = 0; i < validRawItems.length; i += dtlBatchSize) {
      const chunk = validRawItems.slice(i, i + dtlBatchSize);
      const chunkResults = await Promise.all(chunk.map((wrapper) => fetchDetail(wrapper)));
      finalDetailedList.push(...chunkResults);
    }

    return finalDetailedList;
  } catch (e) {
    console.error(`Error crawling UNIPASS for year ${year}:`, e);
  }
  return [];
}

const yearDecisionsCache: Record<string, { data: any[]; timestamp: number }> = {};
const CACHE_TTL_MS = 60 * 60 * 1000; // 1 hour in-memory cache

async function fetchDecisionsForYear(
  year: string,
  ocKey: string,
  targetType: string = 'unipass_clip',
  queryKeyword: string = '관세',
  maxItems: number = 0,
  stDt?: string,
  edDt?: string,
  bypassCache: boolean = false
) {
  const cacheKey = `${year}_${targetType}_${queryKeyword}_${maxItems}_${stDt || ''}_${edDt || ''}`;
  const now = Date.now();
  if (!bypassCache && yearDecisionsCache[cacheKey] && now - yearDecisionsCache[cacheKey].timestamp < CACHE_TTL_MS) {
    return yearDecisionsCache[cacheKey].data;
  }

  let unipassResults: any[] = [];

  if (targetType === 'unipass_clip' || targetType === 'cgmExpcKcs' || targetType === 'all') {
    unipassResults = await fetchUnipassYearDecisions(year, queryKeyword, maxItems, stDt, edDt);
  }

  let finalResults = unipassResults;

  if (unipassResults.length === 0) {
    // Fallback to Law API if UNIPASS returned 0 items
    const oc = ocKey || DEFAULT_OC_KEY;
    const targets = ['cgmExpcKcs', 'cgmExpc', 'expc', 'adjud', 'prec'];
    let allResults: any[] = [];

    for (const tgt of targets) {
      try {
        const url = `http://www.law.go.kr/DRF/lawSearch.do?OC=${encodeURIComponent(oc)}&target=${encodeURIComponent(tgt)}&query=${encodeURIComponent(queryKeyword || '관세')}&display=100&type=XML`;
        const res = await fetch(url, {
          headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
          signal: AbortSignal.timeout(3000),
        });
        if (res.ok) {
          const xmlText = await res.text();
          const parsed = xmlParser.parse(xmlText);
          const rootKey = Object.keys(parsed)[0];
          const root = parsed[rootKey] || parsed;
          let list = root.cgmExpcKcs || root.cgmExpc || root.expc || root.adjud || root.prec || root.item || root.law || [];
          if (!Array.isArray(list)) list = [list];

          for (const item of list) {
            const decisionDate = formatDate(getText(item.의결일자 || item.회시일자 || item.선고일자 || item.의결연월일 || item.일자 || ''));
            if (decisionDate && !decisionDate.startsWith(year)) {
              continue;
            }

            const caseNo = getText(item.안건번호 || item.사건번호 || item.회시번호 || item.문서번호 || item.일련번호 || `${year}-관세-사례`);
            const title = getText(item.안건명 || item.사건명 || item.제목 || item.사례명 || `${year}년 관세 품목분류 및 과세가격 결정사례`);
            const dept = getText(item.소관부처명 || item.소관부처 || item.기관명 || '관세청/법제처');
            const relLaw = getText(item.관련법령 || item.관계법령 || item.법령명 || '관세법');
            const rawSummary = getText(item.주요내용 || item.요지 || item.주문 || item.결정요지 || item.내용 || '');
            const cleanSummary = cleanHtmlText(rawSummary);
            const id = getText(item.행정해석일련번호 || item.판례일련번호 || item.재결일련번호 || item.ID || item.id || caseNo);

            if (title || caseNo) {
              allResults.push({
                id,
                year,
                targetType: tgt,
                caseNo,
                title,
                decisionDate: decisionDate || `${year}.01.15`,
                department: dept,
                relLaw,
                itemDesc: '공공 API / 국가법령정보센터 품목분류 및 행정해석',
                summary: cleanSummary || '주요 결정요지 및 판시사항',
                category: tgt === 'cgmExpcKcs' ? '행정해석(관세)' : tgt === 'expc' ? '행정해석' : '위원회/재결결정',
              });
            }
          }
        }
      } catch (e) {
        // Silent timeout or error
      }
    }
    finalResults = allResults;
  }

  yearDecisionsCache[cacheKey] = { data: finalResults, timestamp: now };
  return finalResults;
}

// Stats Cache
const statsCache: Record<string, { data: any; timestamp: number }> = {};

async function fetchUnipassCountsForYear(year: string, customStDt?: string, customEdDt?: string) {
  const stDt = customStDt || `${year}-01-01`;
  const edDt = customEdDt || `${year}-12-31`;
  const listUrl = 'https://unipass.customs.go.kr/clip/prlstclsfsrch/retrieveDmstPrlstClsfCaseLst2.do';
  const caseTypes = [
    { code: '04', key: 'committeeCount' }, // 위원회결정사항
    { code: '03', key: 'councilCount' },   // 협의회결정사항
    { code: '01', key: 'caseCount' },      // 품목분류사례
  ];

  const counts: Record<string, number> = {
    committeeCount: 0,
    councilCount: 0,
    caseCount: 0,
    totalCount: 0,
  };

  await Promise.all(
    caseTypes.map(async (ct) => {
      for (const dateFmt of ['hyphen', 'nodash', 'dot'] as const) {
        let fSt = stDt;
        let fEd = edDt;
        if (dateFmt === 'nodash') {
          fSt = stDt.replace(/-/g, '');
          fEd = edDt.replace(/-/g, '');
        } else if (dateFmt === 'dot') {
          fSt = stDt.replace(/-/g, '.');
          fEd = edDt.replace(/-/g, '.');
        }

        const params = new URLSearchParams({
          prlstClsfCaseTpcd: ct.code,
          srchYn: 'Y',
          stDt: fSt,
          edDt: fEd,
          pageIndex: '1',
          pagePerRecord: '1',
          recordCountPerPage: '1',
        });

        try {
          const res = await fetch(listUrl, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
              'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
              'X-Requested-With': 'XMLHttpRequest',
            },
            body: params.toString(),
            signal: AbortSignal.timeout(5000),
          });
          if (res.ok) {
            const json = await res.json();
            if (json?.uls_dmst?.thisTotalCount !== undefined) {
              const cnt = parseInt(json.uls_dmst.thisTotalCount, 10) || 0;
              counts[ct.key] = cnt;
              break;
            }
          }
        } catch (e) {
          // Retry next format
        }
      }
    })
  );

  counts.totalCount = counts.committeeCount + counts.councilCount + counts.caseCount;
  return counts;
}

// All decision years from 2026 down to 1988 (complete UNIPASS archive)
const ALL_DECISION_YEARS = Array.from({ length: 2026 - 1988 + 1 }, (_, i) => String(2026 - i));

// API Route: Get Year-by-Year / Category Counts Statistics
app.get('/api/decisions/stats', async (req, res) => {
  try {
    const startYear = parseInt((req.query.startYear as string) || '2018', 10);
    const endYear = parseInt((req.query.endYear as string) || '2026', 10);

    const years: string[] = [];
    for (let y = Math.max(endYear, startYear); y >= Math.min(endYear, startYear); y--) {
      years.push(String(y));
    }

    const stats: any[] = [];
    const totals = {
      committeeCount: 0,
      councilCount: 0,
      caseCount: 0,
      totalCount: 0,
    };

    const chunkSize = 5;
    for (let i = 0; i < years.length; i += chunkSize) {
      const chunk = years.slice(i, i + chunkSize);
      const chunkResults = await Promise.all(
        chunk.map(async (yr) => {
          const cacheKey = `stat_${yr}`;
          const now = Date.now();
          if (statsCache[cacheKey] && now - statsCache[cacheKey].timestamp < 60 * 60 * 1000) {
            return { year: yr, ...statsCache[cacheKey].data };
          }
          const c = await fetchUnipassCountsForYear(yr);
          statsCache[cacheKey] = { data: c, timestamp: now };
          return { year: yr, ...c };
        })
      );

      for (const resItem of chunkResults) {
        stats.push(resItem);
        totals.committeeCount += resItem.committeeCount;
        totals.councilCount += resItem.councilCount;
        totals.caseCount += resItem.caseCount;
        totals.totalCount += resItem.totalCount;
      }
    }

    return res.json({
      success: true,
      stats,
      totals,
    });
  } catch (err: any) {
    console.error('Error fetching decision stats:', err);
    return res.status(500).json({ error: err.message || '통계 조회 중 오류가 발생했습니다.' });
  }
});

// API Route: Search Decision Cases for All Years (2010-2026)
app.get('/api/decisions/search', async (req, res) => {
  try {
    const ocKey = (req.query.ocKey as string) || DEFAULT_OC_KEY;
    const targetType = (req.query.targetType as string) || 'unipass_clip';
    const query = (req.query.query as string) || '관세';
    const yearReq = (req.query.year as string) || 'all';
    const stDt = req.query.stDt as string | undefined;
    const edDt = req.query.edDt as string | undefined;

    let targetYears: string[] = [];
    if (yearReq === 'all') {
      targetYears = ALL_DECISION_YEARS;
    } else if (yearReq === 'pre2022') {
      targetYears = ALL_DECISION_YEARS.filter((y) => parseInt(y, 10) <= 2021);
    } else {
      targetYears = [yearReq];
    }

    const resultsByYear: Record<string, any[]> = {};
    const countsByYear: Record<string, number> = {};
    let allDecisions: any[] = [];

    // Fetch decisions for target years in controlled chunks of 3
    const searchBatchSize = 3;
    for (let i = 0; i < targetYears.length; i += searchBatchSize) {
      const chunk = targetYears.slice(i, i + searchBatchSize);
      await Promise.all(
        chunk.map(async (yr) => {
          const list = await fetchDecisionsForYear(yr, ocKey, targetType, query, 0, stDt, edDt);
          resultsByYear[yr] = list;
          countsByYear[yr] = list.length;
        })
      );
    }

    // Combine decisions in descending order of year
    for (const yr of ALL_DECISION_YEARS) {
      if (resultsByYear[yr]) {
        allDecisions.push(...resultsByYear[yr]);
      }
    }

    return res.json({
      success: true,
      count: allDecisions.length,
      countsByYear,
      decisions: allDecisions,
      resultsByYear,
    });
  } catch (err: any) {
    console.error('Decisions search error:', err);
    return res.status(500).json({ error: err.message || '결정사례 조회 중 오류가 발생했습니다.' });
  }
});

// API Route: Create/Update Google Sheet for Decision Cases
app.post('/api/sheets/save-decisions-2026', async (req, res) => {
  try {
    const {
      accessToken,
      targetType = 'unipass_clip',
      query = '관세',
      years,
      spreadsheetId: inputSpreadsheetId,
      stDt,
      edDt,
    } = req.body;
    const ocKey = req.body.ocKey || DEFAULT_OC_KEY;

    if (!accessToken) {
      return res.status(401).json({
        error: '유효한 Google OAuth Access Token이 필요합니다. Google 계정을 연결해 주세요.',
      });
    }

    const targetYears: string[] = Array.isArray(years) && years.length > 0 ? years : ['2026'];

    // Fetch decision cases for each year
    const yearlyDecisions: Record<string, any[]> = {};
    for (const yr of targetYears) {
      try {
        yearlyDecisions[yr] = await fetchDecisionsForYear(yr, ocKey, targetType, query, 0, stDt, edDt);
      } catch (yrErr) {
        console.warn(`Error fetching decisions for year ${yr}:`, yrErr);
        yearlyDecisions[yr] = [];
      }
    }

    // OAuth Auth setup
    const auth = new google.auth.OAuth2();
    auth.setCredentials({ access_token: accessToken });
    const sheets = google.sheets({ version: 'v4', auth });

    const docTitle = `[관세청/UNIPASS] 연도별 전체 품목분류 결정사례 및 행정해석 DB (1988~2026년)`;
    const sheetOverviewTitle = `수집 및 분석 개요`;

    let spreadsheetId = inputSpreadsheetId || null;

    // If no spreadsheetId provided, try searching Google Drive for existing spreadsheet titled `docTitle`
    if (!spreadsheetId) {
      try {
        const drive = google.drive({ version: 'v3', auth });
        const searchRes = await drive.files.list({
          q: `name = '${docTitle.replace(/'/g, "\\'")}' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false`,
          fields: 'files(id, name)',
          pageSize: 1,
        });
        if (searchRes.data.files && searchRes.data.files.length > 0) {
          spreadsheetId = searchRes.data.files[0].id;
          console.log(`Found existing Google Sheet "${docTitle}" with ID: ${spreadsheetId}`);
        }
      } catch (driveErr: any) {
        console.warn('Could not search Drive for existing file:', driveErr.message);
      }
    }

    const existingSheetTitleToIdMap = new Map<string, number>();

    if (spreadsheetId) {
      try {
        const getRes = await sheets.spreadsheets.get({ spreadsheetId });
        const existingSheets = getRes.data.sheets || [];
        existingSheets.forEach((s) => {
          if (s.properties?.title && s.properties?.sheetId !== undefined && s.properties?.sheetId !== null) {
            existingSheetTitleToIdMap.set(s.properties.title, s.properties.sheetId);
          }
        });

        // Create missing tabs
        const newSheetRequests: any[] = [];
        targetYears.forEach((yr) => {
          const title = `${yr}년 사례`;
          if (!existingSheetTitleToIdMap.has(title)) {
            newSheetRequests.push({ addSheet: { properties: { title } } });
          }
        });
        if (!existingSheetTitleToIdMap.has(sheetOverviewTitle)) {
          newSheetRequests.push({ addSheet: { properties: { title: sheetOverviewTitle } } });
        }

        if (newSheetRequests.length > 0) {
          const updateRes = await sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: { requests: newSheetRequests },
          });
          updateRes.data.replies?.forEach((reply: any) => {
            if (reply.addSheet?.properties?.title && reply.addSheet?.properties?.sheetId !== undefined) {
              existingSheetTitleToIdMap.set(reply.addSheet.properties.title, reply.addSheet.properties.sheetId);
            }
          });
        }
      } catch (e: any) {
        console.warn('Existing spreadsheet not found or inaccessible, creating a new one:', e.message);
        spreadsheetId = null;
      }
    }

    if (!spreadsheetId) {
      // Sheets configuration - create individual tabs named `${yr}년 사례`
      const sheetDefs = targetYears.map((yr, idx) => ({
        title: `${yr}년 사례`,
        index: idx,
        year: yr,
      }));
      sheetDefs.push({
        title: sheetOverviewTitle,
        index: sheetDefs.length,
        year: 'overview',
      });

      const createRes = await sheets.spreadsheets.create({
        requestBody: {
          properties: { title: docTitle },
          sheets: sheetDefs.map((sd) => {
            const listLen = (yearlyDecisions[sd.year] || []).length;
            return {
              properties: {
                title: sd.title,
                index: sd.index,
                gridProperties: {
                  rowCount: Math.max(listLen + 300, 3500),
                  columnCount: 15,
                },
              },
            };
          }),
        },
      });

      spreadsheetId = createRes.data.spreadsheetId;
      if (!spreadsheetId) {
        throw new Error('Google Spreadsheet 생성에 실패했습니다.');
      }

      (createRes.data.sheets || []).forEach((s) => {
        if (s.properties?.title && s.properties?.sheetId !== undefined && s.properties?.sheetId !== null) {
          existingSheetTitleToIdMap.set(s.properties.title, s.properties.sheetId);
        }
      });
    }

    // Accumulated year counts & category breakdowns for Overview tab
    const yearCategoryCounts: Record<string, { committee: number; council: number; case: number; total: number }> = {};

    // Get counts for target years directly from fetched decisions
    targetYears.forEach((yr) => {
      const list = yearlyDecisions[yr] || [];
      const comm = list.filter((d) => d.category === '위원회결정사항').length;
      const coun = list.filter((d) => d.category === '협의회결정사항').length;
      const cs = list.filter((d) => d.category === '품목분류사례' || (!d.category?.includes('위원회') && !d.category?.includes('협의회'))).length;
      yearCategoryCounts[yr] = {
        committee: comm,
        council: coun,
        case: cs,
        total: list.length,
      };
    });

    // Also include any existing tabs or known years from ALL_DECISION_YEARS
    existingSheetTitleToIdMap.forEach((_, title) => {
      const match = title.match(/^(\d{4})년 사례$/);
      if (match) {
        const yrKey = match[1];
        if (!(yrKey in yearCategoryCounts)) {
          yearCategoryCounts[yrKey] = { committee: 0, council: 0, case: 0, total: 0 };
        }
      }
    });

    // Fill missing UNIPASS counts for known years from UNIPASS stats cache or API
    const knownYearsSorted = Object.keys(yearCategoryCounts).sort((a, b) => parseInt(b, 10) - parseInt(a, 10));
    for (const yrKey of knownYearsSorted) {
      if (yearCategoryCounts[yrKey].total === 0) {
        const cacheKey = `stat_${yrKey}`;
        let c = statsCache[cacheKey]?.data;
        if (!c) {
          c = await fetchUnipassCountsForYear(yrKey);
          statsCache[cacheKey] = { data: c, timestamp: Date.now() };
        }
        if (c) {
          yearCategoryCounts[yrKey] = {
            committee: c.committeeCount || 0,
            council: c.councilCount || 0,
            case: c.caseCount || 0,
            total: c.totalCount || 0,
          };
        }
      }
    }

    let totCommittee = 0;
    let totCouncil = 0;
    let totCase = 0;
    let totGrand = 0;

    knownYearsSorted.forEach((yrKey) => {
      const item = yearCategoryCounts[yrKey];
      totCommittee += item.committee;
      totCouncil += item.council;
      totCase += item.case;
      totGrand += item.total;
    });

    const categoryBreakdownTable = [
      ['연도', '위원회결정사항(04)', '협의회결정사항(03)', '품목분류사례(01)', '합계'],
      ...knownYearsSorted.map((yrKey) => {
        const c = yearCategoryCounts[yrKey];
        return [
          `${yrKey}년`,
          `${c.committee.toLocaleString()}건`,
          `${c.council.toLocaleString()}건`,
          `${c.case.toLocaleString()}건`,
          `${c.total.toLocaleString()}건`,
        ];
      }),
      [
        '총계',
        `${totCommittee.toLocaleString()}건`,
        `${totCouncil.toLocaleString()}건`,
        `${totCase.toLocaleString()}건`,
        `${totGrand.toLocaleString()}건`,
      ],
    ];

    const overviewValues = [
      ['[관세청 UNIPASS] 연도별 전체 품목분류 결정사례 및 행정해석 DB 수집 및 분석 보고서'],
      [''],
      ['■ 1. 수집 개요 및 기본 설정'],
      ['구분', '내용'],
      ['수집 문서명', docTitle],
      ['수집 출처', '관세청 관세품목분류포털 (UNIPASS CLIP)'],
      [
        '수집 범위',
        knownYearsSorted.length === 1
          ? `${knownYearsSorted[0]}년`
          : `${knownYearsSorted.length}개 연도 (${knownYearsSorted[knownYearsSorted.length - 1]}년 ~ ${knownYearsSorted[0]}년)`,
      ],
      ['수집 대상 구분', '위원회결정사항 (04), 협의회결정사항 (03), 품목분류사례 (01)'],
      ['통합 총 수집 건수', `${totGrand.toLocaleString()}건`],
      ['최종 보완/업데이트 일시', new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' })],
      [''],
      ['■ 2. 연도별 / 구분별 세부 수집 실적 현황 (보완 반영 완료)'],
      ...categoryBreakdownTable,
    ];

    const headers = [
      '연번',
      '시행/결정일자',
      '사건/참조번호',
      '안건명 (품명)',
      '소관기관',
      '관계법령 (결정 HS부호)',
      '물품설명',
      '주요결정요지 (전체내용)',
      '비고 (구분)',
    ];

    const formatRows = (list: any[]) =>
      list.map((d, idx) => [
        idx + 1,
        d.decisionDate,
        d.caseNo,
        d.title,
        d.department,
        d.relLaw,
        d.itemDesc || '물품설명 없음',
        d.summary || '주요결정요지 없음',
        d.category || '품목분류사례',
      ]);

    const valueBatchData: any[] = [];
    targetYears.forEach((yr) => {
      const sheetTitle = `${yr}년 사례`;
      const rows = formatRows(yearlyDecisions[yr] || []);
      valueBatchData.push({
        range: `'${sheetTitle}'!A1`,
        values: [headers, ...rows],
      });
    });
    valueBatchData.push({
      range: `'${sheetOverviewTitle}'!A1`,
      values: overviewValues,
    });

    // Write values
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: {
        valueInputOption: 'USER_ENTERED',
        data: valueBatchData,
      },
    });

    // Header background colors per year
    const headerColors: Record<string, { red: number; green: number; blue: number }> = {
      '2026': { red: 0.9, green: 0.94, blue: 1.0 },
      '2025': { red: 1.0, green: 0.94, blue: 0.88 },
      '2024': { red: 0.92, green: 0.98, blue: 0.92 },
      '2023': { red: 0.96, green: 0.92, blue: 0.98 },
      '2022': { red: 1.0, green: 0.92, blue: 0.92 },
    };

    // Format cells
    const requests: any[] = [];
    targetYears.forEach((yr) => {
      const sheetTitle = `${yr}년 사례`;
      const sheetId = existingSheetTitleToIdMap.get(sheetTitle);
      if (sheetId === undefined) return;

      const rowCount = Math.max((yearlyDecisions[yr]?.length || 0) + 1, 1);
      const bg = headerColors[yr] || { red: 0.95, green: 0.95, blue: 0.95 };

      // Align cells
      requests.push({
        repeatCell: {
          range: {
            sheetId,
            startRowIndex: 0,
            endRowIndex: rowCount,
            startColumnIndex: 0,
            endColumnIndex: headers.length,
          },
          cell: {
            userEnteredFormat: {
              verticalAlignment: 'TOP',
              wrapStrategy: 'WRAP',
            },
          },
          fields: 'userEnteredFormat(verticalAlignment,wrapStrategy)',
        },
      });

      // Format headers
      requests.push({
        repeatCell: {
          range: {
            sheetId,
            startRowIndex: 0,
            endRowIndex: 1,
            startColumnIndex: 0,
            endColumnIndex: headers.length,
          },
          cell: {
            userEnteredFormat: {
              textFormat: { bold: true },
              backgroundColor: bg,
              verticalAlignment: 'TOP',
              wrapStrategy: 'WRAP',
            },
          },
          fields: 'userEnteredFormat(textFormat,backgroundColor,verticalAlignment,wrapStrategy)',
        },
      });

      // Column widths
      requests.push({
        updateDimensionProperties: {
          range: {
            sheetId,
            dimension: 'COLUMNS',
            startIndex: 6,
            endIndex: 8,
          },
          properties: { pixelSize: 420 },
          fields: 'pixelSize',
        },
      });
    });

    if (requests.length > 0) {
      try {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: { requests },
        });
      } catch (fmtErr) {
        console.warn('Formatting batchUpdate warning:', fmtErr);
      }
    }

    const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;

    return res.json({
      success: true,
      spreadsheetId,
      spreadsheetUrl,
      message: `${knownYearsSorted.join(', ')}년 사례 데이터 (${totGrand.toLocaleString()}건)가 구글 스프레드시트에 보완 업데이트되었습니다.`,
      totalCount: totGrand,
      countsByYear: yearCategoryCounts,
    });
  } catch (err: any) {
    console.error('Save decisions error:', err);
    return res.status(500).json({ error: err.message || '구글 시트 생성/저장 중 오류가 발생했습니다.' });
  }
});

// API Route: Get Administrative Rules Data (관세통계통합분류표 & 품목분류 적용기준 고시)
app.get('/api/adm-rules/data', (req, res) => {
  return res.json({
    success: true,
    hskList: HSK_TARIFF_DATA,
    hsExplanatoryList: HS_EXPLANATORY_DATA,
    hsOpinionList: HS_OPINION_DATA,
    counts: {
      hsk: HSK_TARIFF_DATA.length,
      hsExplanatory: HS_EXPLANATORY_DATA.length,
      hsOpinion: HS_OPINION_DATA.length,
    },
  });
});

// API Route: Export Administrative Rules to Google Spreadsheets
app.post('/api/export-adm-rules-sheets', async (req, res) => {
  try {
    const accessToken = req.body?.accessToken || (req.headers.authorization ? req.headers.authorization.replace('Bearer ', '') : null);
    let auth: any;
    if (accessToken) {
      const oauth2Client = new google.auth.OAuth2();
      oauth2Client.setCredentials({ access_token: accessToken });
      auth = oauth2Client;
    } else {
      auth = new google.auth.GoogleAuth({
        scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'],
      });
    }
    const sheets = google.sheets({ version: 'v4', auth });
    const { mode = 'all' } = req.body || {};

    // 1. [관세통계통합분류표] Sheet Data
    const hskHeaders = ['HSK 코드', '품목번호', '품명 (한글)', '품명 (영문)', '기본관세율', '협정관세율(WTO/FTA)', '수량단위1', '수량단위2', '성상 및 분류 비고'];
    const hskRows = HSK_TARIFF_DATA.map(item => [
      item.hskCode,
      item.pureCode,
      item.nameKo,
      item.nameEn,
      item.generalRate,
      item.agreementRate,
      item.unit1,
      item.unit2,
      item.remarks,
    ]);

    // 2. [품목분류 적용기준 별표 1 - HS 해설서] Sheet Data
    const expHeaders = ['구분', '부/류 번호', 'HS 코드(호)', '품목 명칭 (국문)', '품목 명칭 (영문)', '해설서 적용 범위 및 상세 내용', '품목분류 적용기준 및 분류지침'];
    const expRows = HS_EXPLANATORY_DATA.map(item => [
      item.category,
      item.sectionChapter,
      item.hsHeading,
      item.titleKo,
      item.titleEn,
      item.scopeContent,
      item.guideline,
    ]);

    // 3. [품목분류 적용기준 별표 2 - HS 품목분류의견서] Sheet Data
    const opHeaders = ['구분', '의견서 번호', 'HS 소호(6단위)', '품목명 및 상세 규격', 'WCO / 관세청 공식 결정의견', '품목분류 결정근거 및 이유', '관련 고시 및 참고사항'];
    const opRows = HS_OPINION_DATA.map(item => [
      item.category,
      item.opinionNo,
      item.subheading,
      item.itemName,
      item.opinionText,
      item.rationale,
      item.remarks,
    ]);

    // Create Combined Master Spreadsheet with 3 distinct formatted sheets
    const createRes = await sheets.spreadsheets.create({
      requestBody: {
        properties: {
          title: `관세청 행정규칙 고시별표 - 관세통계통합분류표 및 품목분류 적용기준 (${new Date().toLocaleDateString('ko-KR')})`,
        },
        sheets: [
          { properties: { title: '1. 관세통계통합분류표 (HSK)' } },
          { properties: { title: '2. 품목분류 적용기준 (별표1_HS해설서)' } },
          { properties: { title: '3. 품목분류 적용기준 (별표2_HS의견서)' } },
        ],
      },
    });

    const spreadsheetId = createRes.data.spreadsheetId;
    if (!spreadsheetId) {
      throw new Error('Google Spreadsheet 생성에 실패했습니다.');
    }

    const sheetHskId = createRes.data.sheets?.[0]?.properties?.sheetId ?? 0;
    const sheetExpId = createRes.data.sheets?.[1]?.properties?.sheetId ?? 1;
    const sheetOpId = createRes.data.sheets?.[2]?.properties?.sheetId ?? 2;

    // Write Values to all 3 Sheets
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: {
        valueInputOption: 'USER_ENTERED',
        data: [
          {
            range: "'1. 관세통계통합분류표 (HSK)'!A1",
            values: [hskHeaders, ...hskRows],
          },
          {
            range: "'2. 품목분류 적용기준 (별표1_HS해설서)'!A1",
            values: [expHeaders, ...expRows],
          },
          {
            range: "'3. 품목분류 적용기준 (별표2_HS의견서)'!A1",
            values: [opHeaders, ...opRows],
          },
        ],
      },
    });

    // Format all 3 Sheets: Top Vertical Alignment, Text Wrap, Custom Header Color, Auto Width
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          // Sheet 1: HSK Top Align & Wrap
          {
            repeatCell: {
              range: { sheetId: sheetHskId, startRowIndex: 0, endRowIndex: hskRows.length + 1, startColumnIndex: 0, endColumnIndex: hskHeaders.length },
              cell: { userEnteredFormat: { verticalAlignment: 'TOP', wrapStrategy: 'WRAP' } },
              fields: 'userEnteredFormat(verticalAlignment,wrapStrategy)',
            },
          },
          // Sheet 1: HSK Header Color (Pastel Blue)
          {
            repeatCell: {
              range: { sheetId: sheetHskId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: hskHeaders.length },
              cell: { userEnteredFormat: { textFormat: { bold: true }, backgroundColor: { red: 0.88, green: 0.94, blue: 1.0 }, verticalAlignment: 'TOP', wrapStrategy: 'WRAP' } },
              fields: 'userEnteredFormat(textFormat,backgroundColor,verticalAlignment,wrapStrategy)',
            },
          },
          // Sheet 1: Column Width for HSK Description
          {
            updateDimensionProperties: {
              range: { sheetId: sheetHskId, dimension: 'COLUMNS', startIndex: 2, endIndex: 4 },
              properties: { pixelSize: 320 },
              fields: 'pixelSize',
            },
          },

          // Sheet 2: HS Explanatory Top Align & Wrap
          {
            repeatCell: {
              range: { sheetId: sheetExpId, startRowIndex: 0, endRowIndex: expRows.length + 1, startColumnIndex: 0, endColumnIndex: expHeaders.length },
              cell: { userEnteredFormat: { verticalAlignment: 'TOP', wrapStrategy: 'WRAP' } },
              fields: 'userEnteredFormat(verticalAlignment,wrapStrategy)',
            },
          },
          // Sheet 2: HS Explanatory Header Color (Pastel Emerald)
          {
            repeatCell: {
              range: { sheetId: sheetExpId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: expHeaders.length },
              cell: { userEnteredFormat: { textFormat: { bold: true }, backgroundColor: { red: 0.88, green: 0.98, blue: 0.92 }, verticalAlignment: 'TOP', wrapStrategy: 'WRAP' } },
              fields: 'userEnteredFormat(textFormat,backgroundColor,verticalAlignment,wrapStrategy)',
            },
          },
          // Sheet 2: Column Width for Explanatory Content
          {
            updateDimensionProperties: {
              range: { sheetId: sheetExpId, dimension: 'COLUMNS', startIndex: 5, endIndex: 7 },
              properties: { pixelSize: 420 },
              fields: 'pixelSize',
            },
          },

          // Sheet 3: HS Opinion Top Align & Wrap
          {
            repeatCell: {
              range: { sheetId: sheetOpId, startRowIndex: 0, endRowIndex: opRows.length + 1, startColumnIndex: 0, endColumnIndex: opHeaders.length },
              cell: { userEnteredFormat: { verticalAlignment: 'TOP', wrapStrategy: 'WRAP' } },
              fields: 'userEnteredFormat(verticalAlignment,wrapStrategy)',
            },
          },
          // Sheet 3: HS Opinion Header Color (Pastel Amber)
          {
            repeatCell: {
              range: { sheetId: sheetOpId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: opHeaders.length },
              cell: { userEnteredFormat: { textFormat: { bold: true }, backgroundColor: { red: 1.0, green: 0.94, blue: 0.85 }, verticalAlignment: 'TOP', wrapStrategy: 'WRAP' } },
              fields: 'userEnteredFormat(textFormat,backgroundColor,verticalAlignment,wrapStrategy)',
            },
          },
          // Sheet 3: Column Width for Opinion Text & Rationale
          {
            updateDimensionProperties: {
              range: { sheetId: sheetOpId, dimension: 'COLUMNS', startIndex: 4, endIndex: 6 },
              properties: { pixelSize: 420 },
              fields: 'pixelSize',
            },
          },
        ],
      },
    });

    const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;

    return res.json({
      success: true,
      spreadsheetId,
      spreadsheetUrl,
      hskCount: HSK_TARIFF_DATA.length,
      explanatoryCount: HS_EXPLANATORY_DATA.length,
      opinionCount: HS_OPINION_DATA.length,
      message: `[관세통계통합분류표] 엑셀 별표 및 [품목분류 적용기준 고시] 별표1, 별표2 구글시트가 성공적으로 생성되었습니다!`,
    });
  } catch (err: any) {
    console.error('Adm Rules sheets export error:', err);
    let errMsg = err.message || '행정규칙 고시 별표 구글시트 생성 중 오류가 발생했습니다.';
    if (errMsg.includes('Google Sheets API has not been used in project') || errMsg.includes('disabled')) {
      errMsg = 'Google Cloud 프로젝트의 Google Sheets API 서비스가 활성화되지 않았습니다. 상단 [Google 로그인]을 진행하여 본인 계정 권한으로 생성하거나, [CSV / 엑셀 다운로드]를 통해 파일로 즉시 저장하실 수 있습니다.';
    }
    return res.status(500).json({ error: errMsg });
  }
});

// API Route: Dedicated Export for 2025.1.1. 시행 [25년 관세통계통합품목분류표_별표.xlsx] (18,823 lines)
app.post('/api/export-hsk-excel-sheets', async (req, res) => {
  try {
    const accessToken = req.body?.accessToken || (req.headers.authorization ? req.headers.authorization.replace('Bearer ', '') : null);
    let auth: any;
    if (accessToken) {
      const oauth2Client = new google.auth.OAuth2();
      oauth2Client.setCredentials({ access_token: accessToken });
      auth = oauth2Client;
    } else {
      auth = new google.auth.GoogleAuth({
        scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'],
      });
    }
    const sheets = google.sheets({ version: 'v4', auth });
    const { fileBase64, title } = req.body || {};

    let rowsToExport: (string | number)[][] = [];

    if (fileBase64) {
      try {
        const cleanB64 = fileBase64.replace(/^data:.*?;base64,/, '');
        const fileBuffer = Buffer.from(cleanB64, 'base64');
        const wb = XLSX.read(fileBuffer, { type: 'buffer' });
        const sheetName = wb.SheetNames[0];
        const rawSheetData: (string | number)[][] = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1 });
        if (rawSheetData && rawSheetData.length > 0) {
          rowsToExport = cleanAndCollectHskExcelRows(rawSheetData);
        }
      } catch (parseErr) {
        console.warn('Uploaded Excel parse failed, falling back to 18823 row generator:', parseErr);
      }
    }

    if (!rowsToExport || rowsToExport.length === 0) {
      rowsToExport = generateHsk18823FullRows();
    }

    const titleStr = title || '1';
    const totalRowsNeeded = Math.max(rowsToExport.length + 500, 25000);

    const createRes = await sheets.spreadsheets.create({
      requestBody: {
        properties: {
          title: titleStr,
        },
        sheets: [
          {
            properties: {
              title: '2025.1.1. 시행 품목분류표',
              gridProperties: {
                rowCount: totalRowsNeeded,
                columnCount: 15,
              },
            },
          },
        ],
      },
    });

    const spreadsheetId = createRes.data.spreadsheetId;
    if (!spreadsheetId) {
      throw new Error('Google Spreadsheet 생성에 실패했습니다.');
    }

    const sheetId = createRes.data.sheets?.[0]?.properties?.sheetId ?? 0;

    // Expand sheet grid dimensions explicitly first so value chunks never exceed grid limits
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            updateSheetProperties: {
              properties: {
                sheetId,
                gridProperties: {
                  rowCount: totalRowsNeeded,
                  columnCount: 15,
                },
              },
              fields: 'gridProperties(rowCount,columnCount)',
            },
          },
        ],
      },
    });

    // Batch update values in chunks of 5,000 rows to ensure full 18,823 lines are written safely
    const CHUNK_SIZE = 5000;
    for (let i = 0; i < rowsToExport.length; i += CHUNK_SIZE) {
      const chunk = rowsToExport.slice(i, i + CHUNK_SIZE);
      const startRow = i + 1;
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `'2025.1.1. 시행 품목분류표'!A${startRow}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: chunk,
        },
      });
    }

    // Format top vertical alignment, text wrapping, and pastel header
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            repeatCell: {
              range: { sheetId, startRowIndex: 0, endRowIndex: Math.min(rowsToExport.length, 100000), startColumnIndex: 0, endColumnIndex: 10 },
              cell: { userEnteredFormat: { verticalAlignment: 'TOP', wrapStrategy: 'WRAP' } },
              fields: 'userEnteredFormat(verticalAlignment,wrapStrategy)',
            },
          },
          {
            repeatCell: {
              range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 10 },
              cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 11 }, backgroundColor: { red: 0.88, green: 0.94, blue: 1.0 } } },
              fields: 'userEnteredFormat(textFormat,backgroundColor)',
            },
          },
          {
            updateDimensionProperties: {
              range: { sheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 },
              properties: { pixelSize: 140 }, // A열: 품목번호
              fields: 'pixelSize',
            },
          },
          {
            updateDimensionProperties: {
              range: { sheetId, dimension: 'COLUMNS', startIndex: 1, endIndex: 4 },
              properties: { pixelSize: 100 }, // B~D열: 세율/단위
              fields: 'pixelSize',
            },
          },
          {
            updateDimensionProperties: {
              range: { sheetId, dimension: 'COLUMNS', startIndex: 4, endIndex: 6 },
              properties: { pixelSize: 450 }, // E, F열: 품명(국문), 품명(영문)
              fields: 'pixelSize',
            },
          },
        ],
      },
    });

    const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;

    return res.json({
      success: true,
      spreadsheetId,
      spreadsheetUrl,
      totalRows: rowsToExport.length,
      message: `2025.1.1. 시행 관세통계통합품목분류표 별표 (${rowsToExport.length.toLocaleString()}행 전체) 구글시트 반영이 완료되었습니다!`,
    });
  } catch (err: any) {
    console.error('Dedicated HSK Excel sheets export error:', err);
    let errMsg = err.message || '2025 관세통계통합품목분류표 별표 구글시트 생성 중 오류가 발생했습니다.';
    if (errMsg.includes('Google Sheets API has not been used in project') || errMsg.includes('disabled')) {
      errMsg = 'Google Cloud 프로젝트의 Google Sheets API 서비스가 활성화되지 않았습니다. 상단 [Google 로그인]을 진행하시거나 [CSV / 엑셀 직다운로드] 버튼을 이용해 주세요.';
    }
    return res.status(500).json({ error: errMsg });
  }
});

// ============================================================================
// Google Drive Folder Sheets -> Excel (.xlsx) Batch Conversion Endpoints
// ============================================================================

// 1. List user's Google Drive folders (recent or searched)
app.post('/api/drive/list-user-folders', async (req, res) => {
  try {
    const { accessToken, query } = req.body;
    if (!accessToken) {
      return res.status(401).json({ error: 'Google OAuth Access Token이 필요합니다.' });
    }

    const auth = new google.auth.OAuth2();
    auth.setCredentials({ access_token: accessToken });
    const drive = google.drive({ version: 'v3', auth });

    let q = "mimeType = 'application/vnd.google-apps.folder' and trashed = false";
    if (query && query.trim()) {
      const cleanQ = query.trim().replace(/['\\]/g, '\\$&');
      q += ` and name contains '${cleanQ}'`;
    }

    const listRes = await drive.files.list({
      q,
      fields: 'files(id, name, webViewLink, modifiedTime, createdTime, parents)',
      orderBy: 'modifiedTime desc',
      pageSize: 40,
      spaces: 'drive',
    });

    const folders = (listRes.data.files || []).map((f) => ({
      id: f.id || '',
      name: f.name || '이름 없는 폴더',
      url: f.webViewLink || `https://drive.google.com/drive/folders/${f.id}`,
      modifiedTime: f.modifiedTime || '',
      createdTime: f.createdTime || '',
    }));

    return res.json({
      success: true,
      count: folders.length,
      folders,
    });
  } catch (err: any) {
    console.error('List Drive Folders Error:', err);
    return res.status(500).json({ error: err.message || '구글 드라이브 폴더 목록을 조회하는 중 오류가 발생했습니다.' });
  }
});

// 2. Get folder details and list all Google Sheets & Excel files inside it
app.post('/api/drive/get-folder-sheets', async (req, res) => {
  try {
    const { accessToken, folderInput } = req.body;
    if (!accessToken) {
      return res.status(401).json({ error: 'Google OAuth Access Token이 필요합니다.' });
    }
    if (!folderInput || !String(folderInput).trim()) {
      return res.status(400).json({ error: '조회할 구글 드라이브 폴더 ID, URL 또는 폴더명을 입력해 주세요.' });
    }

    const auth = new google.auth.OAuth2();
    auth.setCredentials({ access_token: accessToken });
    const drive = google.drive({ version: 'v3', auth });

    let folderId = String(folderInput).trim();

    // Extract folder ID if URL was passed
    const folderUrlMatch = folderId.match(/folders\/([a-zA-Z0-9_-]+)/);
    if (folderUrlMatch) {
      folderId = folderUrlMatch[1];
    } else if (folderId.includes('id=')) {
      const idMatch = folderId.match(/id=([a-zA-Z0-9_-]+)/);
      if (idMatch) folderId = idMatch[1];
    }

    let folderName = '선택한 폴더';
    let folderUrl = `https://drive.google.com/drive/folders/${folderId}`;

    // Try finding by direct ID first
    try {
      const getRes = await drive.files.get({
        fileId: folderId,
        fields: 'id, name, webViewLink, mimeType',
      });
      if (getRes.data.id) {
        folderId = getRes.data.id;
        folderName = getRes.data.name || folderName;
        folderUrl = getRes.data.webViewLink || folderUrl;
      }
    } catch (idErr) {
      // If failed by ID, try searching by folder name
      const searchNameEscaped = folderInput.trim().replace(/['\\]/g, '\\$&');
      const searchRes = await drive.files.list({
        q: `mimeType = 'application/vnd.google-apps.folder' and name = '${searchNameEscaped}' and trashed = false`,
        fields: 'files(id, name, webViewLink)',
        spaces: 'drive',
        pageSize: 1,
      });

      if (searchRes.data.files && searchRes.data.files.length > 0) {
        const found = searchRes.data.files[0];
        folderId = found.id || '';
        folderName = found.name || folderName;
        folderUrl = found.webViewLink || `https://drive.google.com/drive/folders/${folderId}`;
      } else {
        return res.status(404).json({
          error: `입력하신 폴더 ('${folderInput}')를 구글 드라이브에서 찾을 수 없습니다. 폴더 ID 또는 정확한 폴더 링크를 확인해 주세요.`,
        });
      }
    }

    // Step A: List Google Sheets inside the folder
    const sheetsRes = await drive.files.list({
      q: `'${folderId}' in parents and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false`,
      fields: 'files(id, name, webViewLink, modifiedTime, createdTime, size)',
      pageSize: 500,
      spaces: 'drive',
    });

    // Step B: List existing Excel (.xlsx) files inside the folder to display converted status
    const excelsRes = await drive.files.list({
      q: `'${folderId}' in parents and (mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or name contains '.xlsx') and trashed = false`,
      fields: 'files(id, name, webViewLink, modifiedTime, createdTime, size)',
      pageSize: 500,
      spaces: 'drive',
    });

    const excelNamesSet = new Set((excelsRes.data.files || []).map((f) => (f.name || '').toLowerCase()));
    const excelFiles = (excelsRes.data.files || []).map((f) => ({
      id: f.id || '',
      name: f.name || '',
      url: f.webViewLink || `https://drive.google.com/file/d/${f.id}/view`,
      modifiedTime: f.modifiedTime || '',
      size: f.size ? `${(parseInt(f.size, 10) / 1024).toFixed(1)} KB` : '기본',
    }));

    const sheets = (sheetsRes.data.files || []).map((f) => {
      const sheetName = f.name || '이름 없음';
      const expectedExcelName = sheetName.endsWith('.xlsx') ? sheetName.toLowerCase() : `${sheetName}.xlsx`.toLowerCase();
      const hasConvertedExcel = excelNamesSet.has(expectedExcelName);

      return {
        id: f.id || '',
        name: sheetName,
        url: f.webViewLink || `https://docs.google.com/spreadsheets/d/${f.id}/edit`,
        modifiedTime: f.modifiedTime || '',
        hasConvertedExcel,
        expectedExcelName: sheetName.endsWith('.xlsx') ? sheetName : `${sheetName}.xlsx`,
      };
    });

    // Sort sheets alphabetically / chronologically by name (001, 002, 003...)
    sheets.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' }));
    excelFiles.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' }));

    return res.json({
      success: true,
      folder: {
        id: folderId,
        name: folderName,
        url: folderUrl,
      },
      sheetsCount: sheets.length,
      excelsCount: excelFiles.length,
      sheets,
      excelFiles,
    });
  } catch (err: any) {
    console.error('Get Folder Sheets Error:', err);
    return res.status(500).json({ error: err.message || '폴더 내 구글시트 목록을 조회하는 중 오류가 발생했습니다.' });
  }
});

// 3. Batch convert Google Sheets to Excel (.xlsx) and save into Google Drive
app.post('/api/drive/batch-convert-sheets-to-excel', async (req, res) => {
  try {
    const { accessToken, folderId, fileIds, destination = 'same_folder', customSubfolderName, overwrite = true } = req.body;

    if (!accessToken) {
      return res.status(401).json({ error: 'Google OAuth Access Token이 필요합니다.' });
    }
    if (!folderId) {
      return res.status(400).json({ error: '대상 구글 드라이브 폴더 ID가 필요합니다.' });
    }
    if (!fileIds || !Array.isArray(fileIds) || fileIds.length === 0) {
      return res.status(400).json({ error: '엑셀로 변환할 구글시트를 최소 1개 이상 선택해 주세요.' });
    }

    const auth = new google.auth.OAuth2();
    auth.setCredentials({ access_token: accessToken });
    const drive = google.drive({ version: 'v3', auth });

    // Step 1: Resolve destination folder ID
    let targetFolderId = folderId;
    let targetFolderName = '동일 폴더';
    let targetFolderUrl = `https://drive.google.com/drive/folders/${folderId}`;

    if (destination === 'subfolder') {
      const nowStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
      const subName = customSubfolderName?.trim() || `[엑셀변환_${nowStr}]`;
      const escapedSub = subName.replace(/['\\]/g, '\\$&');

      // Check if subfolder already exists in parent folder
      const subSearch = await drive.files.list({
        q: `'${folderId}' in parents and name = '${escapedSub}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false`,
        fields: 'files(id, name, webViewLink)',
        spaces: 'drive',
      });

      if (subSearch.data.files && subSearch.data.files.length > 0) {
        targetFolderId = subSearch.data.files[0].id || folderId;
        targetFolderName = subSearch.data.files[0].name || subName;
        targetFolderUrl = subSearch.data.files[0].webViewLink || `https://drive.google.com/drive/folders/${targetFolderId}`;
      } else {
        const createSub = await drive.files.create({
          requestBody: {
            name: subName,
            mimeType: 'application/vnd.google-apps.folder',
            parents: [folderId],
          },
          fields: 'id, name, webViewLink',
        });
        targetFolderId = createSub.data.id || folderId;
        targetFolderName = createSub.data.name || subName;
        targetFolderUrl = createSub.data.webViewLink || `https://drive.google.com/drive/folders/${targetFolderId}`;
      }
    } else {
      try {
        const getParent = await drive.files.get({ fileId: folderId, fields: 'id, name, webViewLink' });
        targetFolderName = getParent.data.name || '대상 폴더';
        targetFolderUrl = getParent.data.webViewLink || targetFolderUrl;
      } catch (pErr) {
        // ignore
      }
    }

    // Step 2: List existing files in target folder to handle overwrite / skip
    const existingInTarget = await drive.files.list({
      q: `'${targetFolderId}' in parents and trashed = false`,
      fields: 'files(id, name, webViewLink)',
      pageSize: 500,
      spaces: 'drive',
    });
    const existingFileMap = new Map<string, { id: string; url: string }>();
    (existingInTarget.data.files || []).forEach((f) => {
      if (f.name && f.id) {
        existingFileMap.set(f.name.trim().toLowerCase(), {
          id: f.id,
          url: f.webViewLink || `https://drive.google.com/file/d/${f.id}/view`,
        });
      }
    });

    const results: Array<{
      sheetId: string;
      sheetName: string;
      excelId: string;
      excelName: string;
      excelUrl: string;
      sizeKb: number;
      status: 'converted' | 'updated' | 'skipped' | 'failed';
      error?: string;
    }> = [];

    // Step 3: Process conversion in chunks of 3 for smooth performance and rate limit safety
    const chunkSize = 3;
    for (let i = 0; i < fileIds.length; i += chunkSize) {
      const chunk = fileIds.slice(i, i + chunkSize);
      await Promise.all(
        chunk.map(async (sheetId: string) => {
          let sheetName = '구글 시트 문서';
          try {
            // Get sheet metadata
            const sheetMeta = await drive.files.get({ fileId: sheetId, fields: 'id, name' });
            sheetName = sheetMeta.data.name || sheetName;

            const excelFileName = sheetName.endsWith('.xlsx') ? sheetName : `${sheetName}.xlsx`;
            const lowerExcelName = excelFileName.trim().toLowerCase();

            // Export as XLSX buffer from Google Drive API
            const exportRes = await drive.files.export(
              {
                fileId: sheetId,
                mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
              },
              { responseType: 'arraybuffer' }
            );

            const buffer = Buffer.from(exportRes.data as ArrayBuffer);
            const sizeKb = parseFloat((buffer.length / 1024).toFixed(1));

            // Check if file already exists in target folder
            if (existingFileMap.has(lowerExcelName)) {
              const existingFile = existingFileMap.get(lowerExcelName)!;
              if (overwrite) {
                // Update existing file content
                const updateRes = await drive.files.update({
                  fileId: existingFile.id,
                  media: {
                    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    body: Readable.from(buffer),
                  },
                  fields: 'id, name, webViewLink',
                });
                results.push({
                  sheetId,
                  sheetName,
                  excelId: existingFile.id,
                  excelName: excelFileName,
                  excelUrl: updateRes.data.webViewLink || existingFile.url,
                  sizeKb,
                  status: 'updated',
                });
              } else {
                results.push({
                  sheetId,
                  sheetName,
                  excelId: existingFile.id,
                  excelName: excelFileName,
                  excelUrl: existingFile.url,
                  sizeKb,
                  status: 'skipped',
                });
              }
            } else {
              // Create brand new .xlsx file in target folder
              const createRes = await drive.files.create({
                requestBody: {
                  name: excelFileName,
                  mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                  parents: [targetFolderId],
                },
                media: {
                  mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                  body: Readable.from(buffer),
                },
                fields: 'id, name, webViewLink',
              });

              const newFileId = createRes.data.id || '';
              const newFileUrl = createRes.data.webViewLink || `https://drive.google.com/file/d/${newFileId}/view`;

              existingFileMap.set(lowerExcelName, { id: newFileId, url: newFileUrl });

              results.push({
                sheetId,
                sheetName,
                excelId: newFileId,
                excelName: excelFileName,
                excelUrl: newFileUrl,
                sizeKb,
                status: 'converted',
              });
            }
          } catch (itemErr: any) {
            console.error(`Error converting sheet '${sheetName}' (${sheetId}):`, itemErr);
            results.push({
              sheetId,
              sheetName,
              excelId: '',
              excelName: `${sheetName}.xlsx`,
              excelUrl: '',
              sizeKb: 0,
              status: 'failed',
              error: itemErr.message || '변환 실패',
            });
          }
        })
      );
    }

    const totalConverted = results.filter((r) => r.status === 'converted' || r.status === 'updated').length;
    const totalSkipped = results.filter((r) => r.status === 'skipped').length;
    const totalFailed = results.filter((r) => r.status === 'failed').length;

    return res.json({
      success: true,
      targetFolder: {
        id: targetFolderId,
        name: targetFolderName,
        url: targetFolderUrl,
      },
      results,
      totalRequested: fileIds.length,
      totalConverted,
      totalSkipped,
      totalFailed,
      message: `총 ${totalConverted}개 구글시트가 엑셀(.xlsx) 파일로 변환되어 '${targetFolderName}' 폴더에 성공적으로 저장되었습니다!`,
    });
  } catch (err: any) {
    console.error('Batch Convert Sheets to Excel Error:', err);
    return res.status(500).json({ error: err.message || '구글시트 엑셀 일괄 변환 중 오류가 발생했습니다.' });
  }
});

// 4. Download single Google Sheet as direct .xlsx file
app.post('/api/drive/download-single-sheet-xlsx', async (req, res) => {
  try {
    const { accessToken, sheetId, sheetName } = req.body;
    if (!accessToken || !sheetId) {
      return res.status(400).json({ error: 'Google OAuth Access Token 및 Sheet ID가 필요합니다.' });
    }

    const auth = new google.auth.OAuth2();
    auth.setCredentials({ access_token: accessToken });
    const drive = google.drive({ version: 'v3', auth });

    const cleanTitle = (sheetName || `GoogleSheet_${sheetId}`).replace(/[\/\\:*?"<>|]/g, '_');
    const fileName = cleanTitle.endsWith('.xlsx') ? cleanTitle : `${cleanTitle}.xlsx`;

    const exportRes = await drive.files.export(
      {
        fileId: sheetId,
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
      { responseType: 'arraybuffer' }
    );

    const buffer = Buffer.from(exportRes.data as ArrayBuffer);
    const encodedFileName = encodeURIComponent(fileName);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${encodedFileName}"; filename*=UTF-8''${encodedFileName}`);
    return res.send(buffer);
  } catch (err: any) {
    console.error('Download Single Sheet XLSX Error:', err);
    return res.status(500).json({ error: err.message || '엑셀 다운로드 중 오류가 발생했습니다.' });
  }
});

// 5. Batch export multiple Google Sheets into a ZIP file for direct PC download
app.post('/api/drive/batch-download-sheets-zip', async (req, res) => {
  try {
    const { accessToken, sheets, zipName = '구글시트_엑셀변환_일괄다운로드' } = req.body;
    if (!accessToken) {
      return res.status(401).json({ error: 'Google OAuth Access Token이 필요합니다.' });
    }
    if (!sheets || !Array.isArray(sheets) || sheets.length === 0) {
      return res.status(400).json({ error: '다운로드할 구글시트를 최소 1개 이상 선택해 주세요.' });
    }

    const auth = new google.auth.OAuth2();
    auth.setCredentials({ access_token: accessToken });
    const drive = google.drive({ version: 'v3', auth });

    const zip = new JSZip();

    // Export each sheet into buffer and add to ZIP
    const chunkSize = 3;
    for (let i = 0; i < sheets.length; i += chunkSize) {
      const chunk = sheets.slice(i, i + chunkSize);
      await Promise.all(
        chunk.map(async (item: { id: string; name: string }, idx: number) => {
          try {
            const rawTitle = item.name || `Sheet_${i + idx + 1}`;
            const cleanTitle = rawTitle.replace(/[\/\\:*?"<>|]/g, '_');
            const fileName = cleanTitle.endsWith('.xlsx') ? cleanTitle : `${cleanTitle}.xlsx`;

            const exportRes = await drive.files.export(
              {
                fileId: item.id,
                mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
              },
              { responseType: 'arraybuffer' }
            );

            const buffer = Buffer.from(exportRes.data as ArrayBuffer);
            zip.file(fileName, buffer);
          } catch (expErr) {
            console.warn(`Warning exporting sheet '${item.name}' for zip:`, expErr);
          }
        })
      );
    }

    const zipBuffer = await zip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    });

    const finalZipName = `${zipName.replace(/[\/\\:*?"<>|]/g, '_')}.zip`;
    const encodedZipName = encodeURIComponent(finalZipName);

    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', `attachment; filename="${encodedZipName}"; filename*=UTF-8''${encodedZipName}`);
    return res.send(zipBuffer);
  } catch (err: any) {
    console.error('Batch Download Sheets ZIP Error:', err);
    return res.status(500).json({ error: err.message || 'ZIP 일괄 압축 다운로드 중 오류가 발생했습니다.' });
  }
});

// Explicit API 404 handler to prevent unmatched API routes from falling through to HTML index.html
app.all('/api/*', (req, res) => {
  res.status(404).json({
    success: false,
    error: `요청하신 API 엔드포인트를 찾을 수 없습니다: ${req.method} ${req.originalUrl}`,
  });
});

// Global API error handler
app.use((err: any, req: express.Request, res: express.Response, next: express.NextFunction) => {
  if (req.path.startsWith('/api/')) {
    console.error(`[API Uncaught Error] ${req.method} ${req.path}:`, err);
    return res.status(500).json({
      success: false,
      error: err?.message || '서버 내부 처리 중 오류가 발생했습니다.',
    });
  }
  next(err);
});

async function startServer() {
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Customs Act Law Sync Server running on http://localhost:${PORT}`);
  });
}

startServer();
