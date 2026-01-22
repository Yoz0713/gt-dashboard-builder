const { google } = require('googleapis');
const { JWT } = require('google-auth-library');

/**
 * 建立 Google 服務帳戶認證客戶端
 * 使用環境變數中的憑證資訊
 */
function getGoogleAuth() {
    const clientEmail = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
    const privateKey = process.env.GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY?.replace(/\\n/g, '\n');

    if (!clientEmail || !privateKey) {
        throw new Error('Missing Google Service Account credentials in environment variables');
    }

    const auth = new JWT({
        email: clientEmail,
        key: privateKey,
        scopes: [
            'https://www.googleapis.com/auth/spreadsheets.readonly',
            'https://www.googleapis.com/auth/drive.readonly',
        ],
    });

    return auth;
}

/**
 * 取得已認證的 Google Sheets API 客戶端
 */
function getSheetsClient(auth) {
    return google.sheets({ version: 'v4', auth });
}

/**
 * 取得已認證的 Google Drive API 客戶端
 */
function getDriveClient(auth) {
    return google.drive({ version: 'v3', auth });
}

module.exports = { getGoogleAuth, getSheetsClient, getDriveClient };
