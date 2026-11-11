/**
 * 報告畫面共用的顯示用小工具（姓名遮罩、距今天數、門市名稱）。
 * 診所回訪報告與待追蹤客戶名單共用，避免兩邊規則漂移。
 */

/** 由 'YYYY-MM-DD' 算出距今天數；空字串或無法解析時回傳 null。 */
export const daysSinceDate = (isoDate: string): number | null => {
    if (!isoDate) return null;

    const [year, month, day] = isoDate.split('-').map((part) => parseInt(part, 10));
    if (!year || !month || !day) return null;

    const target = new Date(year, month - 1, day);
    if (isNaN(target.getTime())) return null;

    const now = new Date();
    const startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate());

    return Math.max(0, Math.round((startOfToday.getTime() - target.getTime()) / 86400000));
};

/** 由時間戳記算出距今天數；0 或無效值回傳 null。 */
export const daysSinceTimestamp = (timestamp: number): number | null => {
    if (!timestamp) return null;

    const target = new Date(timestamp);
    if (isNaN(target.getTime())) return null;

    const startOfTarget = new Date(target.getFullYear(), target.getMonth(), target.getDate());
    const now = new Date();
    const startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate());

    return Math.max(0, Math.round((startOfToday.getTime() - startOfTarget.getTime()) / 86400000));
};

/** 王大明 -> 王○明；兩個字時只遮後面一個字。 */
export const maskCustomerName = (name: string): string => {
    if (name.length <= 1) return name;
    if (name.length === 2) return `${name[0]}○`;
    return `${name[0]}${'○'.repeat(name.length - 2)}${name[name.length - 1]}`;
};

/** 1950/03/12 -> 1950/○○/○○；遮罩時保留年份，仍可辨識年齡層。 */
export const maskBirthDate = (birthDate: string): string => {
    if (!birthDate) return '';

    const year = birthDate.match(/\d{4}/);
    return year ? `${year[0]}/○○/○○` : '○○○';
};

/**
 * 0912-345-678 -> 0912-○○○-678：保留前 4 碼與末 3 碼，其餘數字遮掉。
 * 一格內有多支電話時逐支處理；少於 7 碼的數字（例如分機）不動。
 */
export const maskPhone = (phone: string): string =>
    (phone || '').replace(/\d[\d ()-]*\d/g, (segment) => {
        const digitCount = (segment.match(/\d/g) || []).length;
        if (digitCount < 7) return segment;

        let position = 0;
        return segment.replace(/\d/g, (digit) => {
            position++;
            return position <= 4 || position > digitCount - 3 ? digit : '○';
        });
    });

/**
 * 試算表名稱通常是「門市 + 工作表名」（例：湖口店來客紀錄），
 * 報告表頭只要門市本身，因此把固定的工作表名與前後分隔符去掉。
 */
export const formatStoreName = (spreadsheetTitle: string): string =>
    spreadsheetTitle
        .replace(/來客紀錄/g, '')
        .replace(/^[\s\-_–—|｜]+|[\s\-_–—|｜]+$/g, '')
        .trim();
