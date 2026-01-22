import type { VercelRequest, VercelResponse } from '@vercel/node';
import { getGoogleAuth, getDriveClient, getSheetsClient } from './_lib/googleAuth';

interface SpreadsheetListItem {
    id: string;
    title: string;
    hasCustomerRecord: boolean;
}

/**
 * GET /api/sheets
 * 列出服務帳戶有權限存取的所有 Google 試算表
 * 並檢查每個試算表是否包含「來客紀錄」工作表
 */
export default async function handler(req: VercelRequest, res: VercelResponse) {
    // 只允許 GET 請求
    if (req.method !== 'GET') {
        return res.status(405).json({ error: 'Method not allowed' });
    }

    try {
        const auth = getGoogleAuth();
        const drive = getDriveClient(auth);
        const sheets = getSheetsClient(auth);

        // 使用 Drive API 列出所有試算表
        const driveResponse = await drive.files.list({
            q: "mimeType='application/vnd.google-apps.spreadsheet'",
            fields: 'files(id, name)',
            orderBy: 'modifiedTime desc',
        });

        const files = driveResponse.data.files || [];
        const spreadsheets: SpreadsheetListItem[] = [];

        // 檢查每個試算表是否包含「來客紀錄」工作表
        for (const file of files) {
            if (!file.id || !file.name) continue;

            try {
                const metadata = await sheets.spreadsheets.get({
                    spreadsheetId: file.id,
                    fields: 'sheets.properties.title',
                });

                const sheetTitles = metadata.data.sheets?.map(s => s.properties?.title) || [];
                const hasCustomerRecord = sheetTitles.includes('來客紀錄');

                // 只加入有「來客紀錄」工作表的試算表
                if (hasCustomerRecord) {
                    spreadsheets.push({
                        id: file.id,
                        title: file.name,
                        hasCustomerRecord: true,
                    });
                }
            } catch (err) {
                // 如果無法讀取試算表 metadata，跳過
                console.error(`Failed to read metadata for ${file.name}:`, err);
            }
        }

        return res.status(200).json({
            spreadsheets,
            total: spreadsheets.length,
        });
    } catch (error: any) {
        console.error('Error listing spreadsheets:', error);
        return res.status(500).json({
            error: 'Failed to list spreadsheets',
            message: error.message,
        });
    }
}
