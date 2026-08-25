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
