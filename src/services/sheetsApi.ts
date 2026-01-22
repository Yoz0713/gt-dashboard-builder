import { SpreadsheetListItem, SheetData } from '../types';

const API_BASE = process.env.NODE_ENV === 'development' ? 'http://localhost:3000' : '';

/**
 * 取得服務帳戶有權限的試算表列表
 */
export const fetchSavedSpreadsheets = async (): Promise<SpreadsheetListItem[]> => {
    const response = await fetch(`${API_BASE}/api/sheets`);

    if (!response.ok) {
        const error = await response.json();
        throw new Error(error.message || 'Failed to fetch spreadsheets');
    }

    const data = await response.json();
    return data.spreadsheets;
};

/**
 * 從服務帳戶讀取指定試算表的「來客紀錄」資料
 */
export const fetchSpreadsheetData = async (spreadsheetId: string): Promise<SheetData> => {
    const response = await fetch(`${API_BASE}/api/sheets/${spreadsheetId}`);

    if (!response.ok) {
        const error = await response.json();
        throw new Error(error.message || 'Failed to fetch spreadsheet data');
    }

    const data = await response.json();
    return { values: data.values };
};
