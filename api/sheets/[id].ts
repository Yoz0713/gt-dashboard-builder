import type { VercelRequest, VercelResponse } from '@vercel/node';
import { getGoogleAuth, getSheetsClient } from '../_lib/googleAuth';

/**
 * GET /api/sheets/[id]
 * 讀取指定試算表的「來客紀錄」工作表資料
 */
export default async function handler(req: VercelRequest, res: VercelResponse) {
    // 只允許 GET 請求
    if (req.method !== 'GET') {
        return res.status(405).json({ error: 'Method not allowed' });
    }

    const { id } = req.query;

    if (!id || typeof id !== 'string') {
        return res.status(400).json({ error: 'Missing spreadsheet ID' });
    }

    try {
        const auth = getGoogleAuth();
        const sheets = getSheetsClient(auth);

        // 先取得試算表 metadata 以確認「來客紀錄」工作表存在
        const metadata = await sheets.spreadsheets.get({
            spreadsheetId: id,
            fields: 'properties.title,sheets.properties.title',
        });

        const sheetTitles = metadata.data.sheets?.map(s => s.properties?.title) || [];

        if (!sheetTitles.includes('來客紀錄')) {
            return res.status(404).json({
                error: 'Sheet not found',
                message: '此試算表沒有「來客紀錄」工作表',
            });
        }

        // 讀取「來客紀錄」工作表資料
        const dataResponse = await sheets.spreadsheets.values.get({
            spreadsheetId: id,
            range: '來客紀錄',
        });

        return res.status(200).json({
            spreadsheetTitle: metadata.data.properties?.title,
            sheetName: '來客紀錄',
            values: dataResponse.data.values || [],
        });
    } catch (error: any) {
        console.error('Error fetching sheet data:', error);

        if (error.code === 404) {
            return res.status(404).json({
                error: 'Spreadsheet not found',
                message: '找不到此試算表，請確認服務帳戶有權限存取',
            });
        }

        return res.status(500).json({
            error: 'Failed to fetch sheet data',
            message: error.message,
        });
    }
}
