import { SpreadsheetListItem, SheetData } from '../types';

const API_BASE = process.env.NODE_ENV === 'development' ? 'http://localhost:3000' : '';

const ensureJsonResponse = async (response: Response, endpoint: string) => {
    const contentType = response.headers.get('content-type') || '';

    if (contentType.includes('application/json')) {
        return;
    }

    const body = await response.text();
    const looksLikeHtml = body.trim().startsWith('<!DOCTYPE html') || body.trim().startsWith('<html');

    if (looksLikeHtml) {
        throw new Error(
            `Expected JSON from ${endpoint}, but received HTML. Local \`npm start\` does not serve Vercel \`/api\` routes; run the app behind Vercel dev or another API server.`
        );
    }
};

/**
 * 取得服務帳戶有權限的試算表列表
 */
export const fetchSavedSpreadsheets = async (): Promise<SpreadsheetListItem[]> => {
    const response = await fetch(`${API_BASE}/api/sheets`, {
        cache: 'no-store',
        headers: {
            'Cache-Control': 'no-cache',
            'Pragma': 'no-cache',
        },
    });

    await ensureJsonResponse(response, '/api/sheets');

    if (!response.ok) {
        let message = 'Failed to fetch spreadsheets';

        try {
            const error = await response.json();
            message = error.message || error.error || message;
        } catch {
            // 304/empty-body responses can fail JSON parsing; preserve a useful message.
            message = `${message} (${response.status})`;
        }

        throw new Error(message);
    }

    const data = await response.json();
    return data.spreadsheets;
};

/**
 * 從服務帳戶讀取指定試算表的「來客紀錄」資料
 */
export const fetchSpreadsheetData = async (spreadsheetId: string): Promise<SheetData> => {
    const response = await fetch(`${API_BASE}/api/sheets/${spreadsheetId}`, {
        cache: 'no-store',
        headers: {
            'Cache-Control': 'no-cache',
            'Pragma': 'no-cache',
        },
    });

    await ensureJsonResponse(response, `/api/sheets/${spreadsheetId}`);

    if (!response.ok) {
        let message = 'Failed to fetch spreadsheet data';

        try {
            const error = await response.json();
            message = error.message || error.error || message;
        } catch {
            message = `${message} (${response.status})`;
        }

        throw new Error(message);
    }

    const data = await response.json();
    return { values: data.values };
};
