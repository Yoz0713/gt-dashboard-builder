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
}

export interface DateRange {
  startYear: number;
  startMonth: number;
  endYear: number;
  endMonth: number;
} 