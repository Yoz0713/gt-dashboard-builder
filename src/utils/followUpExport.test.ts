/**
 * @jest-environment node
 */
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { FollowUpCustomerRecord } from '../types';
import {
    buildFollowUpExportTable,
    buildFollowUpFileName,
    COLUMN_WIDTH,
    createFollowUpWorkbook,
    EVEN_ROW_FILL,
    FOLLOW_UP_SHEET_NAME,
    NOTE_COLUMN_WIDTH,
    ODD_ROW_FILL,
    ROW_HEIGHT,
} from './followUpExport';

/** 欄寬預設 15，以下欄位單獨設定。 */
const EXPECTED_WIDTHS: { [header: string]: number } = { 備註: 150, 電話: 32, 顧客來源: 26 };

const createRecord = (overrides: Partial<FollowUpCustomerRecord> = {}): FollowUpCustomerRecord => ({
    spreadsheetId: 'hukou',
    storeName: '湖口店',
    serviceDate: '2025-01-05',
    sortKey: new Date(2025, 0, 5).getTime(),
    name: '王大明',
    phone: '0912-345-678',
    age: 70,
    leftPTA: 60,
    rightPTA: 30,
    worsePTA: 60,
    hearingDegree: '中重度',
    sources: ['診所轉介', '聽篩'],
    referralUnit: '康健診所',
    audiologist: '林小美',
    note: '家人陪同再來',
    ...overrides,
});

const records = [
    createRecord(),
    createRecord({
        name: '陳美麗',
        phone: '',
        serviceDate: '2025-02-10',
        sortKey: new Date(2025, 1, 10).getTime(),
        age: null,
        leftPTA: 45,
        rightPTA: null,
        worsePTA: 45,
        hearingDegree: '中度',
        sources: [],
        referralUnit: '',
        note: '',
    }),
];

const meta = {
    storeLabel: '湖口店',
    generatedAt: new Date(2025, 10, 11, 9, 30),
};

describe('buildFollowUpExportTable', () => {
    it('includes phone and note columns only when requested', () => {
        const withAll = buildFollowUpExportTable(records, { maskNames: false, includePhone: true, includeNote: true });
        expect(withAll.columns.map((column) => column.header)).toEqual([
            '門市', '服務日期', '姓名', '電話', '年齡', '左耳 PTA', '右耳 PTA',
            '聽損程度', '顧客來源', '轉介單位', '主聽力師', '備註',
        ]);
        expect(withAll.columns.map((column) => column.header)).not.toContain('距今天數');
        expect(withAll.columns.map((column) => column.header)).not.toContain('較差耳 PTA');

        const minimal = buildFollowUpExportTable(records, { maskNames: false, includePhone: false, includeNote: false });
        expect(minimal.columns.map((column) => column.header)).not.toContain('電話');
        expect(minimal.columns.map((column) => column.header)).not.toContain('備註');
        expect(minimal.rows[0]).toHaveLength(minimal.columns.length);
    });

    it('writes dates as UTC-midnight dates and numbers as numbers', () => {
        const { columns, rows } = buildFollowUpExportTable(records, { maskNames: false, includePhone: true, includeNote: true });
        const cell = (row: number, header: string) => rows[row][columns.findIndex((column) => column.header === header)];

        expect(cell(0, '服務日期')).toEqual(new Date(Date.UTC(2025, 0, 5)));
        expect(cell(0, '年齡')).toBe(70);
        expect(cell(0, '左耳 PTA')).toBe(60);
        expect(cell(0, '顧客來源')).toBe('診所轉介、聽篩');
        expect(cell(0, '電話')).toBe('0912-345-678');
        // 空值輸出為空白儲存格，而不是 0 或「—」
        expect(cell(1, '年齡')).toBeNull();
        expect(cell(1, '右耳 PTA')).toBeNull();
    });

    it('flattens multi-line phone / note cells into one line so a 25pt row shows everything', () => {
        const { columns, rows } = buildFollowUpExportTable(
            [createRecord({ phone: '0912345678\n03-5551234', note: '想先比較價格\r\n 下個月再聯絡 ' })],
            { maskNames: false, includePhone: true, includeNote: true }
        );

        expect(rows[0][columns.findIndex((column) => column.header === '電話')]).toBe('0912345678；03-5551234');
        expect(rows[0][columns.findIndex((column) => column.header === '備註')]).toBe('想先比較價格；下個月再聯絡');
    });

    it('uses width 15 for every column except note 150, phone 32 and customer source 26', () => {
        const { columns } = buildFollowUpExportTable(records, { maskNames: false, includePhone: true, includeNote: true });

        columns.forEach((column) => {
            expect(column.width).toBe(EXPECTED_WIDTHS[column.header] ?? COLUMN_WIDTH);
        });
        expect(COLUMN_WIDTH).toBe(15);
        expect(NOTE_COLUMN_WIDTH).toBe(150);
    });

    it('applies name / phone masking when the mask toggle is on', () => {
        const { columns, rows } = buildFollowUpExportTable(records, { maskNames: true, includePhone: true, includeNote: false });

        expect(rows[0][columns.findIndex((column) => column.header === '姓名')]).toBe('王○明');
        expect(rows[0][columns.findIndex((column) => column.header === '電話')]).toBe('0912-○○○-678');
    });
});

describe('createFollowUpWorkbook', () => {
    const roundTrip = async () => {
        const workbook = await createFollowUpWorkbook(records, { maskNames: false, includePhone: true, includeNote: true }, meta);
        const buffer = await workbook.xlsx.writeBuffer();
        return buffer as ArrayBuffer;
    };

    it('produces an xlsx whose first sheet has an auto filter over header + data and a frozen header row', async () => {
        const buffer = await roundTrip();
        const zip = await JSZip.loadAsync(buffer);
        const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');

        // 12 欄（A~L），表頭 + 2 筆資料
        expect(sheetXml).toContain('<autoFilter ref="A1:L3"');
        expect(sheetXml).toMatch(/<pane[^>]*ySplit="1"[^>]*state="frozen"/);
    });

    it('can be read back with the expected header and values, as a single worksheet', async () => {
        const buffer = await roundTrip();
        const loaded = new ExcelJS.Workbook();
        await loaded.xlsx.load(buffer);

        const sheet = loaded.getWorksheet(FOLLOW_UP_SHEET_NAME)!;
        expect(sheet.getRow(1).getCell(1).value).toBe('門市');
        expect(sheet.getRow(1).getCell(1).font?.bold).toBe(true);
        expect(sheet.rowCount).toBe(3);

        const row = sheet.getRow(2);
        expect(row.getCell(3).value).toBe('王大明');
        expect(row.getCell(4).value).toBe('0912-345-678'); // 電話保持文字
        expect(row.getCell(2).value).toEqual(new Date(Date.UTC(2025, 0, 5)));
        expect(row.getCell(2).numFmt).toBe('yyyy/mm/dd');
        expect(row.getCell(5).value).toBe(70);
        expect(row.getCell(6).value).toBe(60);
        expect(row.getCell(7).value).toBe(30);

        // 只有明細一個工作表，不再附「報表資訊」
        expect(loaded.worksheets.map((worksheet) => worksheet.name)).toEqual([FOLLOW_UP_SHEET_NAME]);
    });

    it('formats every cell centered, rows 25pt high, widths 15 / note 150, with alternating row fills', async () => {
        const threeRecords = [...records, createRecord({ name: '趙極重', note: '' })];
        const workbook = await createFollowUpWorkbook(threeRecords, { maskNames: false, includePhone: true, includeNote: true }, meta);
        const loaded = new ExcelJS.Workbook();
        await loaded.xlsx.load(await workbook.xlsx.writeBuffer() as ArrayBuffer);
        const sheet = loaded.getWorksheet(FOLLOW_UP_SHEET_NAME)!;
        const columnCount = 12;
        const fillOf = (row: number, column: number) =>
            (sheet.getRow(row).getCell(column).fill as { fgColor?: { argb?: string } }).fgColor?.argb;

        for (let column = 1; column <= columnCount; column++) {
            const header = sheet.getRow(1).getCell(column).value as string;
            expect(sheet.getColumn(column).width).toBe(EXPECTED_WIDTHS[header] ?? 15);
        }

        for (let row = 1; row <= 4; row++) {
            expect(sheet.getRow(row).height).toBe(ROW_HEIGHT);
            for (let column = 1; column <= columnCount; column++) {
                // 含空白儲存格（陳美麗沒有電話、轉介單位、備註）
                expect(sheet.getRow(row).getCell(column).alignment).toMatchObject({ horizontal: 'center', vertical: 'middle' });
            }
        }

        // 表頭品牌藍；第 1、3 筆白底，第 2 筆淡藍底，整列一致
        expect(fillOf(1, 1)).toBe('FF008CD7');
        for (let column = 1; column <= columnCount; column++) {
            expect(fillOf(2, column)).toBe(ODD_ROW_FILL);
            expect(fillOf(3, column)).toBe(EVEN_ROW_FILL);
            expect(fillOf(4, column)).toBe(ODD_ROW_FILL);
        }
        expect(ODD_ROW_FILL).not.toBe(EVEN_ROW_FILL);

        // 日期格式不因套用置中而遺失
        expect(sheet.getRow(2).getCell(2).numFmt).toBe('yyyy/mm/dd');
    });

    it('still writes a valid filter range for an empty list', async () => {
        const workbook = await createFollowUpWorkbook([], { maskNames: false, includePhone: false, includeNote: false }, meta);
        const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer() as ArrayBuffer);
        const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');

        expect(sheetXml).toContain('<autoFilter ref="A1:J1"');
    });
});

describe('buildFollowUpFileName', () => {
    it('includes the store and date and strips characters Windows rejects', () => {
        expect(buildFollowUpFileName('湖口店', new Date(2025, 10, 11))).toBe('待追蹤客戶名單_湖口店_20251111.xlsx');
        expect(buildFollowUpFileName('A/B:C', new Date(2025, 0, 2))).toBe('待追蹤客戶名單_A_B_C_20250102.xlsx');
        expect(buildFollowUpFileName('', new Date(2025, 0, 2))).toBe('待追蹤客戶名單_全部門市_20250102.xlsx');
    });
});
