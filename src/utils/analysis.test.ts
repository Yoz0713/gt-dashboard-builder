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

describe('clinic follow-up analysis', () => {
    const dateRange: DateRange = {
        startYear: 2025,
        startMonth: 1,
        endYear: 2025,
        endMonth: 12,
    };

    const clinicSheetData: SheetData = {
        values: [
            ['服務日期', '姓名', '年齡', '診所名稱', '左耳 PTA', '右耳 PTA', '狀態', '成交金額', '主聽力師'],
            ['2025-01-05', '王大明', '70', '康健診所', '55', '30', '成交', 'NT$120,000', '林小美'],
            ['2025-01-20', '陳小華', '45', '康健診所', '20', '25', '', '', '林小美'],
            ['2025-02-10', '李阿姨', '80', '康健診所', '65', '70', '成交', '90000', '張聽力師'],
            ['2025-03-01', '張先生', '60', '仁愛耳鼻喉科', '45', '10', '', '', '林小美'],
            ['2025-03-02', '趙小姐', '50', '#N/A', '50', '50', '', '', '林小美'],
            ['2025-03-03', '孫先生', '55', '', '30', '30', '', '', '林小美'],
        ],
    };

    const getReports = (sheetData: SheetData = clinicSheetData, ptaThreshold = 40) =>
        analyzeData(sheetData, dateRange, ptaThreshold)?.clinicFollowUpAnalysis ?? [];

    it('groups referrals by clinic name and sorts by total referrals descending', () => {
        const reports = getReports();

        expect(reports.map((report) => report.clinic)).toEqual(['康健診所', '仁愛耳鼻喉科']);
        expect(reports.map((report) => report.totalReferrals)).toEqual([3, 1]);
    });

    it('uses clinic name as the stable tie-breaker when total referrals match', () => {
        const reports = getReports({
            values: [
                ['服務日期', '診所名稱', '左耳 PTA', '右耳 PTA'],
                ['2025-01-01', 'Charlie', '50', '10'],
                ['2025-01-02', 'Alpha', '50', '10'],
                ['2025-01-03', 'Bravo', '50', '10'],
            ],
        });

        expect(reports.map((report) => report.clinic)).toEqual(['Alpha', 'Bravo', 'Charlie']);
    });

    it('classifies hearing loss by PTA threshold and deals by the shared dealt check', () => {
        const [kangJian] = getReports();

        expect(kangJian.hearingLossCount).toBe(2);
        expect(kangJian.normalCount).toBe(1);
        expect(kangJian.dealtCount).toBe(2);
        expect(kangJian.hearingLossRate).toBeCloseTo(66.67, 1);
        expect(kangJian.conversionRate).toBeCloseTo(66.67, 1);
        expect(kangJian.totalAmount).toBe(210000);
        expect(kangJian.averageAmount).toBe(105000);
    });

    it('re-classifies hearing loss when the PTA threshold changes', () => {
        const [kangJian] = getReports(clinicSheetData, 60);

        expect(kangJian.hearingLossCount).toBe(1);
        expect(kangJian.normalCount).toBe(2);
    });

    it('builds monthly buckets, first/last referral dates and the peak month', () => {
        const [kangJian] = getReports();

        expect(kangJian.monthly).toEqual([
            { month: '2025-01', label: '2025年1月', referrals: 2, hearingLoss: 1, deals: 1 },
            { month: '2025-02', label: '2025年2月', referrals: 1, hearingLoss: 1, deals: 1 },
        ]);
        expect(kangJian.firstReferralDate).toBe('2025-01-05');
        expect(kangJian.lastReferralDate).toBe('2025-02-10');
        expect(kangJian.activeMonths).toBe(2);
        expect(kangJian.averagePerMonth).toBeCloseTo(1.5, 5);
        expect(kangJian.peakMonthLabel).toBe('2025年1月');
    });

    it('lists patient records newest first with hearing degree and age distributions', () => {
        const [kangJian] = getReports();

        expect(kangJian.patients.map((patient) => patient.name)).toEqual(['李阿姨', '陳小華', '王大明']);
        expect(kangJian.patients[0]).toMatchObject({
            serviceDate: '2025-02-10',
            age: 80,
            leftPTA: 65,
            rightPTA: 70,
            worsePTA: 70,
            hearingDegree: '中重度',
            isHearingLoss: true,
            isDealt: true,
            amount: 90000,
            audiologist: '張聽力師',
        });
        expect(kangJian.hearingDegreeDistribution).toEqual([
            { degree: '正常', count: 1 },
            { degree: '中度', count: 1 },
            { degree: '中重度', count: 1 },
        ]);
        expect(kangJian.ageDistribution).toEqual([
            { range: '45-64歲', count: 1 },
            { range: '65-79歲', count: 1 },
            { range: '80歲以上', count: 1 },
        ]);
        expect(kangJian.audiologistDistribution).toEqual([
            { name: '林小美', count: 2 },
            { name: '張聽力師', count: 1 },
        ]);
    });

    it('excludes rows whose clinic name is empty or an invalid spreadsheet value', () => {
        const reports = getReports();

        expect(reports.map((report) => report.clinic)).not.toContain('#N/A');
        expect(reports.map((report) => report.clinic)).not.toContain('');
        expect(reports.reduce((sum, report) => sum + report.totalReferrals, 0)).toBe(4);
    });

    it('keeps rows with unparseable PTA in the total while not counting them as hearing loss', () => {
        const reports = getReports({
            values: [
                ['服務日期', '診所名稱', '左耳 PTA', '右耳 PTA'],
                ['2025-01-01', '康健診所', 'abc', 'xyz'],
                ['2025-01-02', '康健診所', '', ''],
                ['2025-01-03', '康健診所', '50', '10'],
            ],
        });

        expect(reports[0].totalReferrals).toBe(3);
        expect(reports[0].hearingLossCount).toBe(1);
        expect(reports[0].normalCount).toBe(2);
        expect(reports[0].hearingDegreeDistribution).toEqual([
            { degree: '中度', count: 1 },
            { degree: '未知', count: 2 },
        ]);
    });

    it('locates the clinic column by fuzzy header match', () => {
        const reports = getReports({
            values: [
                ['服務日期', '轉介診所', '左耳 PTA', '右耳 PTA'],
                ['2025-01-01', '康健診所', '50', '10'],
            ],
        });

        expect(reports.map((report) => report.clinic)).toEqual(['康健診所']);
    });

    it('returns an empty report list when the source sheet has no clinic column', () => {
        const reports = getReports({
            values: [
                ['服務日期', '門市名稱', '左耳 PTA', '右耳 PTA'],
                ['2025-01-01', 'Alpha', '50', '10'],
            ],
        });

        expect(reports).toEqual([]);
    });

    it('leaves the existing clinic Top 8 and competition ranking outputs untouched', () => {
        const result = analyzeData(clinicSheetData, dateRange, 40);

        // 既有 clinicAnalysis 走精確 key 且不過濾 #N/A，刻意與新報表的行為分開，維持 Top 8 圖表原樣。
        expect(result?.clinicAnalysis.map((clinic: { clinic: string }) => clinic.clinic)).toEqual([
            '康健診所',
            '仁愛耳鼻喉科',
            '#N/A',
        ]);
        // 這份樣本沒有轉介門市欄位，競賽排行維持空陣列。
        expect(result?.competitionRankingAnalysis).toEqual([]);
    });
});
