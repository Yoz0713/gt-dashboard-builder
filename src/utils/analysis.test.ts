import { analyzeData } from './analysis';
import { DateRange, SheetData } from '../types';

describe('analyzeData competition ranking', () => {
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
            ['2025-01-04', '2025-01-04', '門市轉介', '15', '10', '', '', '', '', '', '', 'Bravo'],
            ['2025-01-05', '2025-01-05', '門市轉介', '10', '10', '', '', '', '', '', '', 'Bravo'],
            ['2025-01-06', '2025-01-06', '門市轉介', '45', '10', '', '', '', '', '', '', 'Beta'],
            ['2025-01-07', '2025-01-07', '門市轉介', '10', '10', '', '', '', '', '', '', 'Beta'],
            ['2025-01-08', '2025-01-08', '門市轉介', '45', '10', '', '', '', '', '', '', 'Zeta'],
            ['2025-01-09', '2025-01-09', '門市轉介', '10', '10', '', '', '', '', '', '', 'Zeta'],
            ['2025-01-10', '2025-01-10', '門市轉介', '60', '20', '', '', '', '', '', '', ''],
            ['2025-01-11', '2025-01-11', '網路廣告', '60', '20', '', '', '', '', '', '', 'Gamma'],
            ['2025-01-12', '2025-01-12', '網路廣告', '60', '20', '', '', '', '', '', '', ''],
        ],
    };

    it('uses store-name presence to identify store referrals and separates potential vs non-potential customers', () => {
        const result = analyzeData(sheetData, dateRange, 40);

        expect(result).not.toBeNull();
        expect(result?.competitionRankingAnalysis).toEqual([
            { rank: 1, storeName: 'Alpha', potentialCustomers: 2, nonPotentialCustomers: 0 },
            { rank: 2, storeName: 'Bravo', potentialCustomers: 1, nonPotentialCustomers: 2 },
            { rank: 3, storeName: 'Beta', potentialCustomers: 1, nonPotentialCustomers: 1 },
            { rank: 4, storeName: 'Zeta', potentialCustomers: 1, nonPotentialCustomers: 1 },
            { rank: 5, storeName: 'Gamma', potentialCustomers: 1, nonPotentialCustomers: 0 },
        ]);

        expect(result?.competitionRankingAnalysis.find((entry) => entry.storeName === '未填寫門市')).toBeUndefined();
        expect(result?.competitionRankingAnalysis.find((entry) => entry.storeName === '')).toBeUndefined();
    });

    it('keeps existing comparison buckets and overview metrics available', () => {
        const result = analyzeData(sheetData, dateRange, 40);

        expect(result).not.toBeNull();
        expect(result?.customerAnalysis.totalCustomers).toBeGreaterThan(0);
        expect(Object.keys(result?.customerAnalysis.yearBuckets || {})).toContain('2025');
        expect(result?.competitionRankingAnalysis).toHaveLength(5);
    });
});
