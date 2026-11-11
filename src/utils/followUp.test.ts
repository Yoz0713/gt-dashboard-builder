import { CompetitionSheetSource, DateRange, FollowUpCustomerRecord } from '../types';
import {
    buildFollowUpDataset,
    createDealtChecker,
    normalizePhone,
    selectFollowUpCustomers,
    summarizeFollowUpCustomers,
} from './analysis';
import { parseAge } from './parsers';
import { maskPhone } from './display';

describe('maskPhone', () => {
    it('keeps the first 4 and last 3 digits', () => {
        expect(maskPhone('0912-345-678')).toBe('0912-○○○-678');
        expect(maskPhone('0912345678')).toBe('0912○○○678');
        expect(maskPhone('03-5551234')).toBe('03-55○○234');
    });

    it('masks each number separately and leaves short numbers alone', () => {
        expect(maskPhone('0912345678\n035551234')).toBe('0912○○○678\n0355○○234');
        expect(maskPhone('分機 123')).toBe('分機 123');
        expect(maskPhone('')).toBe('');
    });
});

describe('createDealtChecker', () => {
    it('treats the shared dealt keywords and the detected status column as dealt', () => {
        const isDealt = createDealtChecker(['姓名', '成交狀態']);

        expect(isDealt({ 姓名: 'A', 成交狀態: '已成交' })).toBe(true);
        expect(isDealt({ 姓名: 'B', 成交狀態: '是' })).toBe(true);
        expect(isDealt({ 姓名: 'C', 成交狀態: '考慮中' })).toBe(false);
        expect(isDealt({ 姓名: 'D', 成交狀態: '' })).toBe(false);
    });

    it('also checks fixed keys such as 是否借機 even without a status column', () => {
        const isDealt = createDealtChecker(['姓名', '是否借機']);

        expect(isDealt({ 姓名: 'A', 是否借機: 'TRUE' })).toBe(true);
        expect(isDealt({ 姓名: 'B', 是否借機: '否' })).toBe(false);
    });
});

describe('follow-up customer list', () => {
    const dateRange: DateRange = { startYear: 2025, startMonth: 1, endYear: 2025, endMonth: 6 };

    const HEADERS = ['服務日期', '姓名', '年齡', '顧客生日 (西元/月/日)', '顧客來源', '診所名稱', '左耳 PTA', '右耳 PTA', '狀態', '主聽力師', '備註', '聯絡電話'];

    const hukouRows: string[][] = [
        HEADERS,
        // 手機號碼被試算表當成數字，開頭的 0 不見了
        ['2025-01-05', '王大明', '70', '1955-3-12', '診所轉介', '康健診所', '60', '30', '', '林小美', '家人陪同再來', '912345678'],
        ['2025-02-10', '陳美麗', '', '1950/7/1', '網路廣告', '', '45', '42', '考慮中', '張大偉', '', '03-5551234'],
        ['2025-03-01', '李成交', '68', '', '診所轉介', '康健診所', '75', '70', '成交', '林小美', ''],
        ['2025-03-15', '趙輕微', '50', '', '路過', '', '30', '35', '', '張大偉', ''],
        ['2025-04-20', '錢無測', '81', '', '路過', '', '', '', '', '林小美', ''],
        ['2025-05-02', '孫邊界', '77', '', '聽篩, 網路廣告', '', '40', '40', '', '林小美', ''],
        ['2024-12-31', '周區間外', '80', '', '路過', '', '90', '90', '', '林小美', ''],
        ['2025-07-01', '吳區間外', '80', '', '路過', '', '90', '90', '', '林小美', ''],
    ];

    const zhubeiRows: string[][] = [
        ['服務日期', '姓名', '年齡', '左耳 PTA', '右耳 PTA', '是否成交', '主聽力師', '轉介門市'],
        ['2025-06-01', '鄭重度', '85', '95', '80', '否', '黃小芬', '竹北門市'],
        ['2025-01-20', '馮中度', '60', '50', '58', '是', '黃小芬', '竹北門市'],
    ];

    const sources: CompetitionSheetSource[] = [
        { spreadsheetId: 'hukou', spreadsheetTitle: '湖口店來客紀錄', sheetData: { values: hukouRows } },
        { spreadsheetId: 'zhubei', spreadsheetTitle: '竹北店 - 來客紀錄', sheetData: { values: zhubeiRows } },
    ];

    const names = (records: FollowUpCustomerRecord[]) => records.map((record) => record.name);

    describe('buildFollowUpDataset', () => {
        it('keeps only undealt customers inside the date range and counts all in-range customers', () => {
            const dataset = buildFollowUpDataset(sources, dateRange);

            expect(names(dataset.records)).not.toContain('李成交');
            expect(names(dataset.records)).not.toContain('馮中度');
            expect(names(dataset.records)).not.toContain('周區間外');
            expect(names(dataset.records)).not.toContain('吳區間外');
            expect(dataset.records).toHaveLength(6);
            expect(dataset.totalCustomers).toBe(8);
            expect(dataset.storeTotals).toEqual({ 湖口店: 6, 竹北店: 2 });
        });

        it('parses the per-customer fields used by the detail table', () => {
            const dataset = buildFollowUpDataset(sources, dateRange);
            const find = (name: string) => dataset.records.find((record) => record.name === name);

            expect(find('王大明')).toMatchObject({
                storeName: '湖口店',
                spreadsheetId: 'hukou',
                serviceDate: '2025-01-05',
                phone: '0912345678',
                age: 70,
                leftPTA: 60,
                rightPTA: 30,
                worsePTA: 60,
                hearingDegree: '中重度',
                sources: ['診所轉介'],
                referralUnit: '康健診所',
                audiologist: '林小美',
                note: '家人陪同再來',
            });
            // 年齡欄空白時由生日推算
            expect(find('陳美麗')?.age).toBe(parseAge('', '1950/07/01'));
            expect(find('陳美麗')?.phone).toBe('03-5551234');
            expect(find('趙輕微')?.phone).toBe('');
            expect(find('陳美麗')?.referralUnit).toBe('');
            expect(find('孫邊界')?.sources).toEqual(['聽篩', '網路廣告']);
            expect(find('錢無測')).toMatchObject({ leftPTA: null, rightPTA: null, worsePTA: null, hearingDegree: '未知' });
            expect(find('鄭重度')).toMatchObject({
                storeName: '竹北店',
                referralUnit: '竹北門市',
                hearingDegree: '極重度',
                note: '',
            });
        });

        it('reports note / PTA column presence and skips unusable sheets', () => {
            const dataset = buildFollowUpDataset([
                ...sources,
                { spreadsheetId: 'broken', spreadsheetTitle: '壞掉的表', sheetData: { values: [['姓名', '左耳 PTA'], ['A', '80']] } },
                { spreadsheetId: 'empty', spreadsheetTitle: '空表', sheetData: { values: [['服務日期']] } },
            ], dateRange);

            expect(dataset.hasNoteColumn).toBe(true);
            expect(dataset.hasPTAColumn).toBe(true);
            expect(dataset.skippedSpreadsheets.map((sheet) => sheet.spreadsheetId)).toEqual(['broken', 'empty']);

            const noExtras = buildFollowUpDataset([
                { spreadsheetId: 'x', spreadsheetTitle: 'X', sheetData: { values: [['服務日期', '姓名'], ['2025-01-01', 'A']] } },
            ], dateRange);
            expect(noExtras.hasNoteColumn).toBe(false);
            expect(noExtras.hasPhoneColumn).toBe(false);
            expect(noExtras.hasPTAColumn).toBe(false);
            expect(dataset.hasPhoneColumn).toBe(true);
        });

        it('picks the phone column but not a note column whose name contains 電話', () => {
            const dataset = buildFollowUpDataset([{
                spreadsheetId: 'x',
                spreadsheetTitle: 'X',
                sheetData: {
                    values: [
                        ['服務日期', '姓名', '電話回訪紀錄', '手機', '左耳 PTA', '右耳 PTA'],
                        ['2025-01-01', 'A', '已致電，下週再聯絡', '0911-222-333', '60', '50'],
                    ],
                },
            }], dateRange);

            expect(dataset.records[0]).toMatchObject({ phone: '0911-222-333', note: '已致電，下週再聯絡' });
        });

        it('derives age from a Chinese-format birthday when the age cell is blank', () => {
            const dataset = buildFollowUpDataset([{
                spreadsheetId: 'x',
                spreadsheetTitle: 'X',
                sheetData: {
                    values: [
                        ['服務日期', '姓名', '年齡', '生日', '左耳 PTA', '右耳 PTA'],
                        ['2025-01-01', 'A', '', '1950年3月12日', '60', '50'],
                    ],
                },
            }], dateRange);

            expect(dataset.records[0].age).toBe(parseAge('', '1950/03/12'));
            expect(dataset.records[0].age).toBeGreaterThan(70);
        });

        it('restores the leading 0 only for 9-digit numbers starting with 9', () => {
            expect(normalizePhone('912345678')).toBe('0912345678');
            expect(normalizePhone(' 0912-345-678 ')).toBe('0912-345-678');
            expect(normalizePhone('35551234')).toBe('35551234');
            expect(normalizePhone('')).toBe('');
        });

        it('recomputes the in-range list when the date range changes, without refetching', () => {
            const firstQuarter = buildFollowUpDataset(sources, { startYear: 2025, startMonth: 1, endYear: 2025, endMonth: 3 });

            expect(names(firstQuarter.records).sort()).toEqual(['王大明', '趙輕微', '陳美麗'].sort());
            expect(firstQuarter.totalCustomers).toBe(5);
        });
    });

    describe('selectFollowUpCustomers', () => {
        const { records } = buildFollowUpDataset(sources, dateRange);

        it('requires worse-ear PTA strictly above the threshold and excludes customers without PTA', () => {
            const selected = names(selectFollowUpCustomers(records, { ptaThreshold: 40 }));

            expect(selected).toEqual(expect.arrayContaining(['王大明', '陳美麗', '鄭重度']));
            expect(selected).not.toContain('孫邊界'); // 40 / 40：等於閾值不算
            expect(selected).not.toContain('趙輕微');
            expect(selected).not.toContain('錢無測'); // 雙耳 PTA 空白
            expect(selected).toHaveLength(3);
        });

        it('never includes dealt customers even with very high PTA', () => {
            const selected = names(selectFollowUpCustomers(records, { ptaThreshold: 0 }));

            expect(selected).not.toContain('李成交');
            expect(selected).not.toContain('馮中度');
        });

        it('includes a customer when only one ear exceeds the threshold', () => {
            expect(names(selectFollowUpCustomers(records, { ptaThreshold: 55 }))).toEqual(['鄭重度', '王大明']);
        });

        it('tightens the list as the threshold rises', () => {
            expect(names(selectFollowUpCustomers(records, { ptaThreshold: 70 }))).toEqual(['鄭重度']);
            expect(selectFollowUpCustomers(records, { ptaThreshold: 95 })).toHaveLength(0);
        });

        it('applies store, audiologist, degree, source and keyword filters', () => {
            expect(names(selectFollowUpCustomers(records, { ptaThreshold: 40, store: '竹北店' }))).toEqual(['鄭重度']);
            expect(names(selectFollowUpCustomers(records, { ptaThreshold: 40, audiologist: '張大偉' }))).toEqual(['陳美麗']);
            expect(names(selectFollowUpCustomers(records, { ptaThreshold: 40, degree: '中度' }))).toEqual(['陳美麗']);
            expect(names(selectFollowUpCustomers(records, { ptaThreshold: 40, source: '網路廣告' }))).toEqual(['陳美麗']);
            expect(names(selectFollowUpCustomers(records, { ptaThreshold: 40, source: '康健診所' }))).toEqual(['王大明']);
            expect(names(selectFollowUpCustomers(records, { ptaThreshold: 40, keyword: ' 大明 ' }))).toEqual(['王大明']);
        });

        it('sorts by worse-ear PTA (default) or by visit date', () => {
            expect(names(selectFollowUpCustomers(records, { ptaThreshold: 40 }))).toEqual(['鄭重度', '王大明', '陳美麗']);
            expect(names(selectFollowUpCustomers(records, { ptaThreshold: 40, sortBy: 'date-desc' }))).toEqual(['鄭重度', '陳美麗', '王大明']);
            expect(names(selectFollowUpCustomers(records, { ptaThreshold: 40, sortBy: 'date-asc' }))).toEqual(['王大明', '陳美麗', '鄭重度']);
        });

        it('does not mutate the input records', () => {
            const before = names(records);
            selectFollowUpCustomers(records, { ptaThreshold: 40, sortBy: 'date-asc' });
            expect(names(records)).toEqual(before);
        });
    });

    describe('summarizeFollowUpCustomers', () => {
        const dataset = buildFollowUpDataset(sources, dateRange);

        it('summarizes counts, ratio against undealt customers, severity and distributions', () => {
            const selected = selectFollowUpCustomers(dataset.records, { ptaThreshold: 40 });
            const summary = summarizeFollowUpCustomers(selected, dataset);

            expect(summary.followUpCount).toBe(3);
            expect(summary.undealtCount).toBe(6);
            expect(summary.totalCustomers).toBe(8);
            expect(summary.followUpRate).toBeCloseTo(50);
            expect(summary.severeCount).toBe(2); // 60、95 > 55
            expect(summary.averageWorsePTA).toBeCloseTo((60 + 45 + 95) / 3, 1);
            expect(summary.audiologistDistribution).toHaveLength(3);
            expect(summary.audiologistDistribution.every((item) => item.count === 1)).toBe(true);
            expect(summary.sourceDistribution).toEqual(expect.arrayContaining([
                { name: '診所轉介', count: 1 },
                { name: '網路廣告', count: 1 },
                { name: '未填寫', count: 1 },
            ]));
            expect(summary.storeDistribution).toEqual([
                { name: '湖口店', count: 2 },
                { name: '竹北店', count: 1 },
            ]);
        });

        it('uses the selected store as the denominator scope', () => {
            const selected = selectFollowUpCustomers(dataset.records, { ptaThreshold: 40, store: '竹北店' });
            const summary = summarizeFollowUpCustomers(selected, dataset, '竹北店');

            expect(summary.undealtCount).toBe(1);
            expect(summary.totalCustomers).toBe(2);
            expect(summary.followUpRate).toBeCloseTo(100);
        });

        it('handles an empty selection', () => {
            const summary = summarizeFollowUpCustomers([], dataset);

            expect(summary.followUpCount).toBe(0);
            expect(summary.followUpRate).toBe(0);
            expect(summary.averageWorsePTA).toBeNull();
            expect(summary.audiologistDistribution).toEqual([]);
        });
    });
});
