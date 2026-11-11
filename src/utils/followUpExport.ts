import type { Workbook } from 'exceljs';
import { FollowUpCustomerRecord } from '../types';
import { maskCustomerName, maskPhone } from './display';

/**
 * 待追蹤名單匯出 Excel。
 * - 第一列為表頭並凍結、套用自動篩選，打開檔案就能直接用 Excel 的下拉篩選／排序。
 * - 日期、年齡、PTA 寫成真正的日期與數字，Excel 的日期分組與數字篩選（大於、介於…）才能用。
 * - 電話一律寫成文字，避免 Excel 吃掉開頭的 0。
 */

export interface FollowUpExportOptions {
    maskNames: boolean;
    includePhone: boolean;
    includeNote: boolean;
}

export interface FollowUpExportMeta {
    /** 用於檔名 */
    storeLabel: string;
    generatedAt?: Date;
}

export interface FollowUpExportColumn {
    header: string;
    key: string;
    width: number;
    numFmt?: string;
}

export type FollowUpExportCell = string | number | Date | null;

export const FOLLOW_UP_SHEET_NAME = '待追蹤名單';

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/** 與畫面報告表頭相同的品牌藍（ARGB）。 */
const HEADER_FILL = 'FF008CD7';

/** 明細工作表統一格式 */
export const COLUMN_WIDTH = 15;
export const NOTE_COLUMN_WIDTH = 150;
export const PHONE_COLUMN_WIDTH = 32;
export const SOURCE_COLUMN_WIDTH = 26;
export const ROW_HEIGHT = 25;
/** 資料列交錯底色：第 1、3、5… 筆白底，第 2、4、6… 筆淡藍底。 */
export const ODD_ROW_FILL = 'FFFFFFFF';
export const EVEN_ROW_FILL = 'FFEAF4FB';

/**
 * 列高固定 25，儲存格內的換行會讓第二行被切掉看不到，
 * 因此匯出時把多行內容（多支電話、多行備註）合併成一行，以「；」分隔。
 */
const toSingleLine = (value: string) => value.replace(/\s*[\r\n]+\s*/g, '；').trim();

/**
 * Excel 的日期沒有時區：exceljs 以 UTC 換算序號，直接傳本地午夜的 Date 在台灣（UTC+8）會變成前一天。
 * 因此以本地的年月日組出 UTC 午夜。
 */
const toExcelDate = (timestamp: number): Date | null => {
    if (!timestamp) return null;
    const local = new Date(timestamp);
    if (isNaN(local.getTime())) return null;
    return new Date(Date.UTC(local.getFullYear(), local.getMonth(), local.getDate()));
};

/** 依目前畫面狀態（遮罩、是否有電話／備註欄）組出匯出用的欄位與資料列。純函式，方便測試。 */
export const buildFollowUpExportTable = (
    records: FollowUpCustomerRecord[],
    options: FollowUpExportOptions
): { columns: FollowUpExportColumn[]; rows: FollowUpExportCell[][] } => {
    const columns: FollowUpExportColumn[] = [
        { header: '門市', key: 'store', width: COLUMN_WIDTH },
        { header: '服務日期', key: 'serviceDate', width: COLUMN_WIDTH, numFmt: 'yyyy/mm/dd' },
        { header: '姓名', key: 'name', width: COLUMN_WIDTH },
        ...(options.includePhone ? [{ header: '電話', key: 'phone', width: PHONE_COLUMN_WIDTH }] : []),
        { header: '年齡', key: 'age', width: COLUMN_WIDTH },
        { header: '左耳 PTA', key: 'leftPTA', width: COLUMN_WIDTH },
        { header: '右耳 PTA', key: 'rightPTA', width: COLUMN_WIDTH },
        { header: '聽損程度', key: 'degree', width: COLUMN_WIDTH },
        { header: '顧客來源', key: 'sources', width: SOURCE_COLUMN_WIDTH },
        { header: '轉介單位', key: 'referralUnit', width: COLUMN_WIDTH },
        { header: '主聽力師', key: 'audiologist', width: COLUMN_WIDTH },
        ...(options.includeNote ? [{ header: '備註', key: 'note', width: NOTE_COLUMN_WIDTH }] : []),
    ];

    const rows = records.map((record) => {
        const values: { [key: string]: FollowUpExportCell } = {
            store: record.storeName,
            serviceDate: toExcelDate(record.sortKey) ?? record.serviceDate,
            name: options.maskNames ? maskCustomerName(record.name) : record.name,
            phone: toSingleLine(options.maskNames ? maskPhone(record.phone) : record.phone),
            age: record.age,
            leftPTA: record.leftPTA,
            rightPTA: record.rightPTA,
            degree: record.hearingDegree,
            sources: record.sources.join('、'),
            referralUnit: record.referralUnit,
            audiologist: record.audiologist,
            note: toSingleLine(record.note),
        };

        return columns.map((column) => values[column.key] ?? null);
    });

    return { columns, rows };
};

const loadExcelJS = async () => {
    // 只在按下匯出時才下載 exceljs，不拖慢主畫面載入。
    const excelModule: any = await import('exceljs');
    return (excelModule.default ?? excelModule) as typeof import('exceljs');
};

const columnLetter = (index: number) => {
    let letter = '';
    let n = index;
    while (n > 0) {
        const remainder = (n - 1) % 26;
        letter = String.fromCharCode(65 + remainder) + letter;
        n = Math.floor((n - 1) / 26);
    }
    return letter;
};

export const createFollowUpWorkbook = async (
    records: FollowUpCustomerRecord[],
    options: FollowUpExportOptions,
    meta: FollowUpExportMeta
): Promise<Workbook> => {
    const ExcelJS = await loadExcelJS();
    const { columns, rows } = buildFollowUpExportTable(records, options);
    const generatedAt = meta.generatedAt ?? new Date();

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'gt-dashboard-builder';
    workbook.created = generatedAt;

    const sheet = workbook.addWorksheet(FOLLOW_UP_SHEET_NAME, {
        views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }],
    });

    sheet.columns = columns.map((column) => ({
        header: column.header,
        key: column.key,
        width: column.width,
        style: column.numFmt ? { numFmt: column.numFmt } : {},
    }));

    rows.forEach((row) => sheet.addRow(row));

    // 統一格式：所有儲存格（含空白格）左右、上下置中，列高 25；資料列奇偶交錯底色。
    for (let rowNumber = 1; rowNumber <= rows.length + 1; rowNumber++) {
        const row = sheet.getRow(rowNumber);
        const isHeader = rowNumber === 1;
        // 第 2 列是第 1 筆資料（奇數筆）
        const fillColor = isHeader ? HEADER_FILL : (rowNumber % 2 === 0 ? ODD_ROW_FILL : EVEN_ROW_FILL);

        row.height = ROW_HEIGHT;

        for (let columnNumber = 1; columnNumber <= columns.length; columnNumber++) {
            const cell = row.getCell(columnNumber);
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } };
            if (isHeader) {
                cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            }
        }
    }

    // 自動篩選涵蓋表頭到最後一筆資料
    sheet.autoFilter = `A1:${columnLetter(columns.length)}${Math.max(1, rows.length + 1)}`;

    return workbook;
};

/** 把檔名中 Windows 不允許的字元換掉。 */
export const buildFollowUpFileName = (storeLabel: string, generatedAt: Date = new Date()) => {
    const stamp = `${generatedAt.getFullYear()}${String(generatedAt.getMonth() + 1).padStart(2, '0')}${String(generatedAt.getDate()).padStart(2, '0')}`;
    const safeStore = (storeLabel || '全部門市').replace(/[\\/:*?"<>|]/g, '_').slice(0, 60);
    return `待追蹤客戶名單_${safeStore}_${stamp}.xlsx`;
};

export const downloadFollowUpExcel = async (
    records: FollowUpCustomerRecord[],
    options: FollowUpExportOptions,
    meta: FollowUpExportMeta
): Promise<void> => {
    const generatedAt = meta.generatedAt ?? new Date();
    const workbook = await createFollowUpWorkbook(records, options, { ...meta, generatedAt });
    const buffer = await workbook.xlsx.writeBuffer();

    const url = URL.createObjectURL(new Blob([buffer], { type: XLSX_MIME }));
    const link = document.createElement('a');
    link.href = url;
    link.download = buildFollowUpFileName(meta.storeLabel, generatedAt);
    document.body.appendChild(link);
    link.click();
    link.remove();
    // 部分瀏覽器在 click 後才開始讀取 blob，延遲釋放。
    setTimeout(() => URL.revokeObjectURL(url), 1000);
};
