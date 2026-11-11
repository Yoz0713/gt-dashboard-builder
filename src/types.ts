export interface SpreadsheetListItem {
  id: string;
  title: string;
  hasCustomerRecord: boolean;
}

export interface UserProfile {
  id: string;
  name: string;
  email: string;
  picture: string;
}

export interface SheetData {
  values: string[][];
}

export interface SheetInfo {
  properties: {
    sheetId: number;
    title: string;
    index: number;
  };
}

export interface SpreadsheetMetadata {
  properties: {
    title: string;
  };
  sheets: SheetInfo[];
}

export interface ColumnStats {
  name: string;
  values: number[];
  labels: string[];
}

export interface DataSummary {
  totalRows: number;
  nonEmptyRows: number;
  totalColumns: number;
}

export interface MonthlyReport {
  month: string;
  year: number;
  newCustomers: number;
  completedDeals: number;
  conversionRate: number;
  totalAmount: number;
  averageAmount: number;
}

export interface CustomerAnalysis {
  monthlyData: MonthlyReport[];
  totalCustomers: number;
  totalCompletedDeals: number;
  totalAmount: number;
  overallConversionRate: number;
  dateRange: {
    earliest: string;
    latest: string;
  };
  // New Analysis Fields
  ageAnalysis: { range: string; count: number }[];
  sourceAnalysis: { source: string; count: number; conversionRate: number }[];
  hearingLossAnalysis: { degree: string; count: number }[];
  weekdayAnalysis: { day: string; visits: number; deals: number; conversionRate: number }[];
  // Buckets for Comparison
  yearBuckets: { [key: string]: PeriodStats };
  quarterBuckets: { [key: string]: PeriodStats };
  monthBuckets: { [key: string]: PeriodStats };
}

export interface CompetitionRankingEntry {
  rank: number;
  storeName: string;
  totalReferrals: number;
  hearingLossCustomers: number;
  normalCustomers: number;
}

export interface CompetitionRankingSkippedSpreadsheet {
  spreadsheetId: string;
  spreadsheetTitle: string;
  reason: string;
}

export interface CompetitionSheetSource {
  spreadsheetId: string;
  spreadsheetTitle: string;
  sheetData: SheetData;
}

export interface CompetitionRankingAggregateResult {
  entries: CompetitionRankingEntry[];
  skippedSpreadsheets: CompetitionRankingSkippedSpreadsheet[];
}

export interface PeriodStats {
  period: string; // Label (e.g., "2024", "2024-Q1", "2024-01")
  newCustomers: number;
  completedDeals: number;
  totalAmount: number;
  conversionRate: number;
  averageOrderValue: number;
  ageDistribution: { [key: string]: number };
  sourceDistribution: { [key: string]: number };
  hearingLossDistribution: { [key: string]: number };
  salespersonPerformance: { [key: string]: { visits: number; deals: number; revenue: number } };
}

export interface DateRange {
  startYear: number;
  startMonth: number;
  endYear: number;
  endMonth: number;
} 

export interface ClinicPatientRecord {
  serviceDate: string;
  sortKey: number;
  name: string;
  birthDate: string;
  age: number | null;
  leftPTA: number | null;
  rightPTA: number | null;
  worsePTA: number | null;
  hearingDegree: string;
  isHearingLoss: boolean;
  isDealt: boolean;
  amount: number;
  audiologist: string;
}

export interface ClinicMonthlyPoint {
  month: string;
  label: string;
  referrals: number;
  hearingLoss: number;
  deals: number;
}

export interface ClinicFollowUpReport {
  clinic: string;
  totalReferrals: number;
  hearingLossCount: number;
  normalCount: number;
  dealtCount: number;
  conversionRate: number;
  hearingLossRate: number;
  totalAmount: number;
  averageAmount: number;
  firstReferralDate: string;
  lastReferralDate: string;
  activeMonths: number;
  averagePerMonth: number;
  peakMonthLabel: string;
  monthly: ClinicMonthlyPoint[];
  hearingDegreeDistribution: { degree: string; count: number }[];
  ageDistribution: { range: string; count: number }[];
  audiologistDistribution: { name: string; count: number }[];
  patients: ClinicPatientRecord[];
}

/** 待追蹤名單的單筆客戶（只收未成交者；PTA 閾值在顯示端即時套用）。 */
export interface FollowUpCustomerRecord {
  spreadsheetId: string;
  storeName: string;
  serviceDate: string;
  sortKey: number;
  name: string;
  phone: string;
  age: number | null;
  leftPTA: number | null;
  rightPTA: number | null;
  worsePTA: number | null;
  hearingDegree: string;
  sources: string[];
  referralUnit: string;
  audiologist: string;
  note: string;
}

export interface FollowUpDataset {
  /** 分析區間內所有未成交客戶，尚未套用 PTA 閾值。 */
  records: FollowUpCustomerRecord[];
  /** 分析區間內總來客數（含已成交）。 */
  totalCustomers: number;
  /** 各門市的分析區間總來客數（含已成交），供門市篩選時計算比例。 */
  storeTotals: { [storeName: string]: number };
  hasNoteColumn: boolean;
  hasPhoneColumn: boolean;
  /** 至少一份來源有左／右耳 PTA 欄位；全部沒有時畫面需提示。 */
  hasPTAColumn: boolean;
  skippedSpreadsheets: CompetitionRankingSkippedSpreadsheet[];
}

export type FollowUpSortKey = 'pta-desc' | 'date-desc' | 'date-asc';

export interface FollowUpFilters {
  ptaThreshold: number;
  store?: string;
  audiologist?: string;
  degree?: string;
  source?: string;
  keyword?: string;
  sortBy?: FollowUpSortKey;
}

export interface FollowUpSummary {
  followUpCount: number;
  undealtCount: number;
  totalCustomers: number;
  /** 待追蹤人數佔同期未成交客戶比例（%）。 */
  followUpRate: number;
  /** 較差耳 PTA > 55 dB（中重度以上）的人數。 */
  severeCount: number;
  averageWorsePTA: number | null;
  audiologistDistribution: { name: string; count: number }[];
  sourceDistribution: { name: string; count: number }[];
  storeDistribution: { name: string; count: number }[];
}
