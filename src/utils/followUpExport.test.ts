/**
 * @jest-environment node
 */
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { FollowUpCustomerRecord } from '../types';
import {
    buildFollowUpExportTable,
    buildFollowUpFileName,
    buildPhoneListTable,
    COLUMN_WIDTH,
    createFollowUpWorkbook,
    EVEN_ROW_FILL,
    FOLLOW_UP_SHEET_NAME,
    NOTE_COLUMN_WIDTH,
    ODD_ROW_FILL,
    partitionExportableRecords,
    PHONE_LIST_SHEET_NAME,
    ROW_HEIGHT,
    TEXT_NUM_FMT,
    toValidMobile,
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

// 依畫面排序：王（正確）、陳（沒電話）、林（市話）、趙（被吃掉開頭 0）、吳（亂碼）
const records = [
    createRecord(),
    createRecord({ name: '陳美麗', phone: '' }),
    createRecord({ name: '林市話', phone: '03-5551234' }),
    createRecord({
        name: '趙極重',
        phone: '933111222',
        serviceDate: '2025-04-01',
        sortKey: new Date(2025, 3, 1).getTime(),
        age: null,
        leftPTA: 95,
        rightPTA: null,
        worsePTA: 95,
        hearingDegree: '極重度',
        sources: [],
        referralUnit: '',
        note: '想先比較價格\r\n 下個月再聯絡 ',
    }),
    createRecord({ name: '吳亂碼', phone: '12345' }),
];

const options = { maskNames: false, includeNote: true };

const meta = {
    storeLabel: '湖口店',
    generatedAt: new Date(2025, 10, 11, 9, 30),
};

const column = (table: { columns: { header: string }[] }, header: string) =>
    table.columns.findIndex((item) => item.header === header);

describe('toValidMobile', () => {
    it.each([
        ['0912345678', '0912345678'],
        ['0912-345-678', '0912345678'],
        ['0912 345 678', '0912345678'],
        ['(0912)345678', '0912345678'],
        ['912345678', '0912345678'],
        ['+886 912 345 678', '0912345678'],
        ['+886-0912-345-678', '0912345678'],
        ['０９１２３４５６７８', '0912345678'],
        ['03-5551234\n0933111222', '0933111222'],
        ['0912345678、0933111222', '0912345678'],
        ['0912345678 0933111222', '0912345678'],
        ['0912-345-678（本人）', '0912345678'],
    ])('normalizes %p to %p', (raw, expected) => {
        expect(toValidMobile(raw)).toBe(expected);
    });

    it.each([
        [''],
        ['   '],
        ['03-5551234'],
        ['02-29123456'],
        ['12345'],
        ['091234567'],
        ['09123456789'],
        ['0812345678'],
        ['無'],
        ['#N/A'],
    ])('rejects %p', (raw) => {
        expect(toValidMobile(raw)).toBeNull();
    });
});

describe('partitionExportableRecords', () => {
    it('keeps customers with a valid mobile in on-screen order and reports the rest', () => {
        const { exportable, skipped } = partitionExportableRecords(records);

        expect(exportable.map(({ record, mobile }) => [record.name, mobile])).toEqual([
            ['王大明', '0912345678'],
            ['趙極重', '0933111222'],
        ]);
        expect(skipped.map((record) => record.name)).toEqual(['陳美麗', '林市話', '吳亂碼']);
    });
});

describe('buildFollowUpExportTable', () => {
    it('always includes a text-formatted phone column; note column only when requested', () => {
        const withNote = buildFollowUpExportTable(records, options);
        expect(withNote.columns.map((item) => item.header)).toEqual([
            '門市', '服務日期', '姓名', '電話', '年齡', '左耳 PTA', '右耳 PTA',
            '聽損程度', '顧客來源', '轉介單位', '主聽力師', '備註',
        ]);
        expect(withNote.columns[column(withNote, '電話')].numFmt).toBe(TEXT_NUM_FMT);

        const withoutNote = buildFollowUpExportTable(records, { ...options, includeNote: false });
        expect(withoutNote.columns.map((item) => item.header)).not.toContain('備註');
        expect(withoutNote.columns.map((item) => item.header)).toContain('電話');
        expect(withoutNote.rows[0]).toHaveLength(withoutNote.columns.length);
    });

    it('removes customers without a phone or with an invalid phone, and normalizes the rest to 09xxxxxxxx', () => {
        const table = buildFollowUpExportTable(records, options);

        expect(table.rows.map((row) => row[column(table, '姓名')])).toEqual(['王大明', '趙極重']);
        expect(table.rows.map((row) => row[column(table, '電話')])).toEqual(['0912345678', '0933111222']);
        table.rows.forEach((row) => expect(typeof row[column(table, '電話')]).toBe('string'));
    });

    it('writes dates as UTC-midnight dates, numbers as numbers and blanks as empty cells', () => {
        const table = buildFollowUpExportTable(records, options);
        const [wang, zhao] = table.rows;

        expect(wang[column(table, '服務日期')]).toEqual(new Date(Date.UTC(2025, 0, 5)));
        expect(wang[column(table, '年齡')]).toBe(70);
        expect(wang[column(table, '左耳 PTA')]).toBe(60);
        expect(wang[column(table, '顧客來源')]).toBe('診所轉介、聽篩');
        expect(zhao[column(table, '年齡')]).toBeNull();
        expect(zhao[column(table, '右耳 PTA')]).toBeNull();
        // 列高固定 25，多行備註合併成一行
        expect(zhao[column(table, '備註')]).toBe('想先比較價格；下個月再聯絡');
    });

    it('uses width 15 for every column except note 150, phone 32 and customer source 26', () => {
        const { columns } = buildFollowUpExportTable(records, options);

        columns.forEach((item) => expect(item.width).toBe(EXPECTED_WIDTHS[item.header] ?? COLUMN_WIDTH));
        expect(NOTE_COLUMN_WIDTH).toBe(150);
    });

    it('applies name / phone masking when the mask toggle is on', () => {
        const table = buildFollowUpExportTable(records, { ...options, maskNames: true });

        expect(table.rows[0][column(table, '姓名')]).toBe('王○明');
        expect(table.rows[0][column(table, '電話')]).toBe('0912○○○678');
    });
});

describe('buildPhoneListTable', () => {
    it('extracts only name and phone for the same customers in the same order', () => {
        const table = buildPhoneListTable(records, options);

        expect(table.columns.map((item) => [item.header, item.width, item.numFmt])).toEqual([
            ['姓名', 15, undefined],
            ['電話', 32, TEXT_NUM_FMT],
        ]);
        expect(table.rows).toEqual([
            ['王大明', '0912345678'],
            ['趙極重', '0933111222'],
        ]);
    });

    it('follows the mask toggle too', () => {
        expect(buildPhoneListTable(records, { ...options, maskNames: true }).rows).toEqual([
            ['王○明', '0912○○○678'],
            ['趙○重', '0933○○○222'],
        ]);
    });
});

describe('createFollowUpWorkbook', () => {
    const loadWorkbook = async (input: FollowUpCustomerRecord[] = records, exportOptions = options) => {
        const workbook = await createFollowUpWorkbook(input, exportOptions, meta);
        const buffer = await workbook.xlsx.writeBuffer() as ArrayBuffer;
        const loaded = new ExcelJS.Workbook();
        await loaded.xlsx.load(buffer);
        return { buffer, loaded };
    };

    const sheetXml = async (buffer: ArrayBuffer, index: number) => {
        const zip = await JSZip.loadAsync(buffer);
        return zip.file(`xl/worksheets/sheet${index}.xml`)!.async('string');
    };

    /** 兩個工作表必須套用一模一樣的格式。 */
    const expectUniformFormatting = (sheet: ExcelJS.Worksheet, columnCount: number, dataRowCount: number) => {
        const fillOf = (row: number, col: number) =>
            (sheet.getRow(row).getCell(col).fill as { fgColor?: { argb?: string } }).fgColor?.argb;

        for (let col = 1; col <= columnCount; col++) {
            const header = sheet.getRow(1).getCell(col).value as string;
            expect(sheet.getColumn(col).width).toBe(EXPECTED_WIDTHS[header] ?? COLUMN_WIDTH);
            expect(sheet.getRow(1).getCell(col).font?.bold).toBe(true);
            expect(fillOf(1, col)).toBe('FF008CD7');
        }

        for (let row = 1; row <= dataRowCount + 1; row++) {
            expect(sheet.getRow(row).height).toBe(ROW_HEIGHT);
            for (let col = 1; col <= columnCount; col++) {
                expect(sheet.getRow(row).getCell(col).alignment).toMatchObject({ horizontal: 'center', vertical: 'middle' });
                if (row > 1) {
                    // 第 1、3… 筆白底，第 2、4… 筆淡藍底
                    expect(fillOf(row, col)).toBe(row % 2 === 0 ? ODD_ROW_FILL : EVEN_ROW_FILL);
                }
            }
        }
    };

    it('contains exactly the 待追蹤名單 and 電話名單 sheets, both with frozen headers and auto filters', async () => {
        const { buffer, loaded } = await loadWorkbook();

        expect(loaded.worksheets.map((sheet) => sheet.name)).toEqual([FOLLOW_UP_SHEET_NAME, PHONE_LIST_SHEET_NAME]);

        const followUpXml = await sheetXml(buffer, 1);
        const phoneXml = await sheetXml(buffer, 2);
        // 待追蹤名單 12 欄（A~L）、電話名單 2 欄（A~B），表頭 + 2 筆有效電話
        expect(followUpXml).toContain('<autoFilter ref="A1:L3"');
        expect(phoneXml).toContain('<autoFilter ref="A1:B3"');
        expect(followUpXml).toMatch(/<pane[^>]*ySplit="1"[^>]*state="frozen"/);
        expect(phoneXml).toMatch(/<pane[^>]*ySplit="1"[^>]*state="frozen"/);
    });

    it('formats both sheets identically: widths, 25pt rows, centered cells, alternating fills', async () => {
        const { loaded } = await loadWorkbook([...records, createRecord({ name: '孫第三', phone: '0955000111', note: '' })]);

        expectUniformFormatting(loaded.getWorksheet(FOLLOW_UP_SHEET_NAME)!, 12, 3);
        expectUniformFormatting(loaded.getWorksheet(PHONE_LIST_SHEET_NAME)!, 2, 3);
    });

    it('stores phones as text in text format so Excel keeps the leading 0', async () => {
        const { loaded } = await loadWorkbook();
        const followUp = loaded.getWorksheet(FOLLOW_UP_SHEET_NAME)!;
        const phoneList = loaded.getWorksheet(PHONE_LIST_SHEET_NAME)!;

        [followUp.getRow(2).getCell(4), phoneList.getRow(2).getCell(2)].forEach((cell) => {
            expect(cell.value).toBe('0912345678');
            expect(cell.type).toBe(ExcelJS.ValueType.String);
            expect(cell.numFmt).toBe(TEXT_NUM_FMT);
        });

        expect(phoneList.getRow(1).getCell(1).value).toBe('姓名');
        expect(phoneList.getRow(1).getCell(2).value).toBe('電話');
        expect(phoneList.getRow(3).values).toEqual([undefined, '趙極重', '0933111222']);
        expect(phoneList.rowCount).toBe(3);
    });

    it('keeps the other column values and formats on the 待追蹤名單 sheet', async () => {
        const { loaded } = await loadWorkbook();
        const row = loaded.getWorksheet(FOLLOW_UP_SHEET_NAME)!.getRow(2);

        expect(row.getCell(3).value).toBe('王大明');
        expect(row.getCell(2).value).toEqual(new Date(Date.UTC(2025, 0, 5)));
        expect(row.getCell(2).numFmt).toBe('yyyy/mm/dd');
        expect(row.getCell(5).value).toBe(70);
        expect(row.getCell(6).value).toBe(60);
        expect(row.getCell(7).value).toBe(30);
    });

    it('still writes valid sheets when nobody has a valid phone', async () => {
        const { buffer } = await loadWorkbook([createRecord({ phone: '' })], { ...options, includeNote: false });

        expect(await sheetXml(buffer, 1)).toContain('<autoFilter ref="A1:K1"');
        expect(await sheetXml(buffer, 2)).toContain('<autoFilter ref="A1:B1"');
    });
});

describe('buildFollowUpFileName', () => {
    it('includes the store and date and strips characters Windows rejects', () => {
        expect(buildFollowUpFileName('湖口店', new Date(2025, 10, 11))).toBe('待追蹤客戶名單_湖口店_20251111.xlsx');
        expect(buildFollowUpFileName('A/B:C', new Date(2025, 0, 2))).toBe('待追蹤客戶名單_A_B_C_20250102.xlsx');
        expect(buildFollowUpFileName('', new Date(2025, 0, 2))).toBe('待追蹤客戶名單_全部門市_20250102.xlsx');
    });
});
