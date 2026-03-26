import { CompetitionSheetSource, DateRange, SheetData } from '../types';
import { aggregateCompetitionRankings, analyzeData } from './analysis';

describe('competition ranking aggregation', () => {
    const dateRange: DateRange = {
        startYear: 2025,
        startMonth: 1,
        endYear: 2025,
        endMonth: 12,
    };

    const createCompetitionSource = (
        spreadsheetId: string,
        spreadsheetTitle: string,
        rows: string[][]
    ): CompetitionSheetSource => ({
        spreadsheetId,
        spreadsheetTitle,
        sheetData: { values: rows },
    });

    it('aggregates multiple spreadsheets into total referrals and PTA-based hearing-loss vs normal counts', () => {
        const sources: CompetitionSheetSource[] = [
            createCompetitionSource('sheet-a', '北區中心', [
                ['日期', '服務日期', '左耳 PTA', '右耳 PTA', '狀態', '轉介門市'],
                ['2025-01-01', '2025-01-01', '50', '10', '', 'Alpha'],
                ['2025-01-02', '2025-01-02', '15', '10', '成交', 'Alpha'],
                ['2025-01-03', '2025-01-03', '45', '10', '', 'Bravo'],
            ]),
            createCompetitionSource('sheet-b', '南區中心', [
                ['日期', '服務日期', '左耳 PTA', '右耳 PTA', '狀態', '轉介門市'],
                ['2025-01-04', '2025-01-04', '20', '20', '', 'Alpha'],
                ['2025-01-05', '2025-01-05', '15', '15', '', 'Bravo'],
                ['2025-01-06', '2025-01-06', '60', '10', '', 'Charlie'],
            ]),
        ];

        const result = aggregateCompetitionRankings(sources, dateRange, 40);

        expect(result.skippedSpreadsheets).toEqual([]);
        expect(result.entries).toEqual([
            { rank: 1, storeName: 'Alpha', totalReferrals: 3, hearingLossCustomers: 1, normalCustomers: 2 },
            { rank: 2, storeName: 'Bravo', totalReferrals: 2, hearingLossCustomers: 1, normalCustomers: 1 },
            { rank: 3, storeName: 'Charlie', totalReferrals: 1, hearingLossCustomers: 1, normalCustomers: 0 },
        ]);
    });

    it('sorts by total referrals descending and uses store name as the stable tie-breaker', () => {
        const sources: CompetitionSheetSource[] = [
            createCompetitionSource('sheet-a', '北區中心', [
                ['日期', '服務日期', '左耳 PTA', '右耳 PTA', '轉介門市'],
                ['2025-01-01', '2025-01-01', '50', '10', 'Bravo'],
                ['2025-01-02', '2025-01-02', '20', '20', 'Alpha'],
                ['2025-01-03', '2025-01-03', '20', '20', 'Alpha'],
                ['2025-01-04', '2025-01-04', '20', '20', 'Bravo'],
                ['2025-01-05', '2025-01-05', '20', '20', 'Delta'],
                ['2025-01-06', '2025-01-06', '20', '20', 'Delta'],
                ['2025-01-07', '2025-01-07', '20', '20', 'Delta'],
            ]),
        ];

        const result = aggregateCompetitionRankings(sources, dateRange, 40);

        expect(result.entries.map((entry) => entry.storeName)).toEqual(['Delta', 'Alpha', 'Bravo']);
        expect(result.entries.map((entry) => entry.rank)).toEqual([1, 2, 3]);
    });

    it('skips spreadsheets that are missing service date or store name columns without affecting other sheets', () => {
        const sources: CompetitionSheetSource[] = [
            createCompetitionSource('valid-sheet', '有效試算表', [
                ['日期', '服務日期', '左耳 PTA', '右耳 PTA', '轉介門市'],
                ['2025-01-01', '2025-01-01', '50', '10', 'Alpha'],
            ]),
            createCompetitionSource('invalid-sheet', '缺欄位試算表', [
                ['日期', '服務日期', '左耳 PTA', '右耳 PTA'],
                ['2025-01-01', '2025-01-01', '50', '10'],
            ]),
        ];

        const result = aggregateCompetitionRankings(sources, dateRange, 40);

        expect(result.entries).toEqual([
            { rank: 1, storeName: 'Alpha', totalReferrals: 1, hearingLossCustomers: 1, normalCustomers: 0 },
        ]);
        expect(result.skippedSpreadsheets).toEqual([
            {
                spreadsheetId: 'invalid-sheet',
                spreadsheetTitle: '缺欄位試算表',
                reason: '缺少服務日期或轉介門市欄位，無法納入競賽排行。',
            },
        ]);
    });

    it('counts rows with missing or unparseable PTA values as normal customers while preserving total referrals', () => {
        const sources: CompetitionSheetSource[] = [
            createCompetitionSource('sheet-a', 'PTA 缺失試算表', [
                ['日期', '服務日期', '轉介門市'],
                ['2025-01-01', '2025-01-01', 'Alpha'],
                ['2025-01-02', '2025-01-02', 'Alpha'],
            ]),
            createCompetitionSource('sheet-b', 'PTA 異常格式試算表', [
                ['日期', '服務日期', '左耳 PTA', '右耳 PTA', '轉介門市'],
                ['2025-01-03', '2025-01-03', 'abc', 'xyz', 'Alpha'],
            ]),
        ];

        const result = aggregateCompetitionRankings(sources, dateRange, 40);

        expect(result.entries).toEqual([
            { rank: 1, storeName: 'Alpha', totalReferrals: 3, hearingLossCustomers: 0, normalCustomers: 3 },
        ]);
        expect(result.skippedSpreadsheets).toEqual([]);
    });
});

describe('analyzeData overview integration', () => {
    const dateRange: DateRange = {
        startYear: 2025,
        startMonth: 1,
        endYear: 2025,
        endMonth: 12,
    };

    const sheetData: SheetData = {
        values: [
            ['日期', '服務日期', '顧客來源', '左耳 PTA', '右耳 PTA', '狀態', '成交金額', '欄位7', '欄位8', '欄位9', '欄位10', '門市名稱'],
            ['2025-01-01', '2025-01-01', '門市轉介', '50', '10', '', '', '', '', '', '', 'Alpha'],
            ['2025-01-02', '2025-01-02', '門市轉介', '15', '10', '成交', '0', '', '', '', '', 'Alpha'],
            ['2025-01-03', '2025-01-03', '門市轉介', '45', '10', '', '', '', '', '', '', 'Bravo'],
            ['2025-01-04', '2025-01-04', '網路廣告', '60', '20', '', '', '', '', '', '', ''],
        ],
    };

    it('keeps existing overview metrics available while exposing the updated competition ranking shape', () => {
        const result = analyzeData(sheetData, dateRange, 40);

        expect(result).not.toBeNull();
        expect(result?.customerAnalysis.totalCustomers).toBeGreaterThan(0);
        expect(Object.keys(result?.customerAnalysis.yearBuckets || {})).toContain('2025');
        expect(result?.competitionRankingAnalysis).toEqual([
            { rank: 1, storeName: 'Alpha', totalReferrals: 2, hearingLossCustomers: 1, normalCustomers: 1 },
            { rank: 2, storeName: 'Bravo', totalReferrals: 1, hearingLossCustomers: 1, normalCustomers: 0 },
        ]);
    });
});
