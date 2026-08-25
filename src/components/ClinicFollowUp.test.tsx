import React from 'react';
import { render, screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import {
    ClinicFollowUp,
    calculateReferralFee,
    daysSinceDate,
    formatStoreName,
    maskCustomerName,
} from './ClinicFollowUp';
import { ClinicFollowUpReport, ClinicPatientRecord, DateRange } from '../types';

// forwardRef：元件會把 ref 掛到圖表實例上，好在列印時主動 resize
jest.mock('react-chartjs-2', () => {
    const ReactModule = require('react');
    return {
        Bar: ReactModule.forwardRef(() => ReactModule.createElement('div', null, 'bar-chart')),
        Doughnut: ReactModule.forwardRef(() => ReactModule.createElement('div', null, 'doughnut-chart')),
    };
});

const dateRange: DateRange = {
    startYear: 2025,
    startMonth: 1,
    endYear: 2025,
    endMonth: 12,
};

const createPatient = (overrides: Partial<ClinicPatientRecord> = {}): ClinicPatientRecord => ({
    serviceDate: '2025-03-01',
    sortKey: new Date(2025, 2, 1).getTime(),
    name: '王大明',
    age: 70,
    leftPTA: 55,
    rightPTA: 30,
    worsePTA: 55,
    hearingDegree: '中度',
    isHearingLoss: true,
    isDealt: true,
    amount: 120000,
    audiologist: '林小美',
    ...overrides,
});

const createReport = (overrides: Partial<ClinicFollowUpReport> = {}): ClinicFollowUpReport => ({
    clinic: '康健診所',
    totalReferrals: 3,
    hearingLossCount: 2,
    normalCount: 1,
    dealtCount: 2,
    conversionRate: 66.6667,
    hearingLossRate: 66.6667,
    totalAmount: 200000,
    averageAmount: 100000,
    firstReferralDate: '2025-01-05',
    lastReferralDate: '2025-03-20',
    activeMonths: 3,
    averagePerMonth: 1,
    peakMonthLabel: '2025年2月',
    monthly: [{ month: '2025-02', label: '2025年2月', referrals: 2, hearingLoss: 1, deals: 1 }],
    hearingDegreeDistribution: [
        { degree: '正常', count: 1 },
        { degree: '中度', count: 2 },
    ],
    ageDistribution: [{ range: '65-79歲', count: 3 }],
    audiologistDistribution: [{ name: '林小美', count: 3 }],
    patients: [
        createPatient(),
        createPatient({ name: '陳小華', isDealt: false, isHearingLoss: false, amount: 0 }),
        createPatient({ name: '李阿姨', amount: 80000, serviceDate: '2025-02-14' }),
    ],
    ...overrides,
});

const secondReport = createReport({
    clinic: '仁愛耳鼻喉科',
    totalReferrals: 1,
    hearingLossCount: 1,
    normalCount: 0,
    dealtCount: 0,
    conversionRate: 0,
    hearingLossRate: 100,
    totalAmount: 0,
    averageAmount: 0,
    activeMonths: 2,
    lastReferralDate: '2025-01-10',
    patients: [createPatient({ name: '張先生', isDealt: false, amount: 0 })],
});

const renderComponent = (reports: ClinicFollowUpReport[]) =>
    render(
        <ClinicFollowUp
            reports={reports}
            dateRange={dateRange}
            spreadsheetTitle="竹北店來客紀錄"
            ptaThreshold={40}
        />
    );

describe('ClinicFollowUp', () => {
    it('selects the clinic with the most referrals by default', () => {
        renderComponent([createReport(), secondReport]);

        expect(screen.getByText('康健診所 回訪報告')).toBeInTheDocument();
        expect(screen.getByText('橫跨 3 個月')).toBeInTheDocument();
        expect(screen.getByText('目前選定：康健診所')).toBeInTheDocument();
    });

    it('switches the report when another clinic is selected', async () => {
        const user = userEvent.setup();
        renderComponent([createReport(), secondReport]);

        await user.click(screen.getByRole('button', { name: /仁愛耳鼻喉科/ }));

        expect(screen.getByText('仁愛耳鼻喉科 回訪報告')).toBeInTheDocument();
        expect(screen.getByText('橫跨 2 個月')).toBeInTheDocument();
        expect(screen.queryByText('康健診所 回訪報告')).not.toBeInTheDocument();
    });

    it('filters the clinic list by the search keyword', async () => {
        const user = userEvent.setup();
        renderComponent([createReport(), secondReport]);

        await user.type(screen.getByLabelText('搜尋診所名稱'), '仁愛');

        expect(screen.getByRole('button', { name: /仁愛耳鼻喉科/ })).toBeInTheDocument();
        expect(screen.queryByRole('button', { name: /康健診所/ })).not.toBeInTheDocument();
    });

    it('shows a message when the search matches no clinic', async () => {
        const user = userEvent.setup();
        renderComponent([createReport(), secondReport]);

        await user.type(screen.getByLabelText('搜尋診所名稱'), '不存在的診所');

        expect(screen.getByText('找不到符合「不存在的診所」的診所。')).toBeInTheDocument();
    });

    it('masks customer names in both the fee table and the detail table', async () => {
        const user = userEvent.setup();
        renderComponent([createReport()]);

        // 王大明與李阿姨已配戴，因此同時出現在轉介費用表與客戶明細表
        expect(screen.getAllByText('王大明')).toHaveLength(2);

        await user.click(screen.getByLabelText('姓名遮罩'));

        expect(screen.queryByText('王大明')).not.toBeInTheDocument();
        expect(screen.getAllByText('王○明')).toHaveLength(2);
        expect(screen.getByText('陳○華')).toBeInTheDocument();
    });

    it('builds talking points from the selected report', async () => {
        const user = userEvent.setup();
        renderComponent([createReport(), secondReport]);

        expect(screen.getByText(/本期共轉介 3 位個案，其中 2 位（67%）/)).toBeInTheDocument();
        expect(screen.getByText('已完成配戴 2 位，配戴率 67%。')).toBeInTheDocument();

        await user.click(screen.getByRole('button', { name: /仁愛耳鼻喉科/ }));

        expect(screen.getByText('本期尚未有個案完成配戴，建議一併討論轉介後的追蹤流程。')).toBeInTheDocument();
    });

    it('shows an explicit empty state when no clinic data is available', () => {
        renderComponent([]);

        expect(screen.getByText('此分析區間內查無「診所名稱」欄位資料。')).toBeInTheDocument();
        expect(screen.queryByText('轉介客戶明細')).not.toBeInTheDocument();
    });
});

describe('ClinicFollowUp referral fee block', () => {
    it('defaults to a 10% referral fee and totals only the dealt patients', () => {
        renderComponent([createReport()]);

        expect(screen.getByLabelText('回饋比例')).toHaveValue(10);
        expect(screen.getByText('轉介費用 (10%)')).toBeInTheDocument();

        // 王大明 120,000 -> 12,000；李阿姨 80,000 -> 8,000；陳小華未配戴不計入。
        // 每個金額各出現兩次：轉介費用表一次、客戶明細表一次。
        expect(screen.getAllByText('NT$12,000')).toHaveLength(2);
        expect(screen.getAllByText('NT$8,000')).toHaveLength(2);
        // 合計 20,000 出現在 KPI 卡與表尾
        expect(screen.getAllByText('NT$20,000')).toHaveLength(2);
        expect(screen.getAllByText('NT$200,000')).toHaveLength(2);
        expect(screen.getByText('來自 2 筆已配戴個案')).toBeInTheDocument();
    });

    it('recalculates the fee when the rate is changed', async () => {
        const user = userEvent.setup();
        renderComponent([createReport()]);

        const rateInput = screen.getByLabelText('回饋比例');
        await user.clear(rateInput);
        await user.type(rateInput, '15');

        expect(screen.getByText('轉介費用 (15%)')).toBeInTheDocument();
        expect(screen.getAllByText('NT$18,000')).toHaveLength(2);
        expect(screen.getAllByText('NT$12,000')).toHaveLength(2);
        expect(screen.getAllByText('NT$30,000')).toHaveLength(2);
    });

    it('warns when a dealt patient has no deal amount recorded', () => {
        renderComponent([
            createReport({
                patients: [createPatient({ amount: 0 }), createPatient({ name: '李阿姨', amount: 80000 })],
            }),
        ]);

        expect(
            screen.getByText(/有 1 筆已配戴個案未填寫成交金額/)
        ).toBeInTheDocument();
    });

    it('explains that there is no fee to calculate when nobody was fitted', () => {
        renderComponent([secondReport]);

        expect(screen.getByText('本期此診所尚無已配戴個案，因此沒有可計算的轉介費用。')).toBeInTheDocument();
    });
});

describe('calculateReferralFee', () => {
    it('takes the given percentage of the deal amount, rounded to the nearest dollar', () => {
        expect(calculateReferralFee(120000, 10)).toBe(12000);
        expect(calculateReferralFee(85500, 10)).toBe(8550);
        expect(calculateReferralFee(99999, 10)).toBe(10000);
    });

    it('returns zero for missing amounts or rates', () => {
        expect(calculateReferralFee(0, 10)).toBe(0);
        expect(calculateReferralFee(-100, 10)).toBe(0);
        expect(calculateReferralFee(120000, 0)).toBe(0);
    });
});

describe('ClinicFollowUp print output', () => {
    it('titles the exported report without the old wording', () => {
        renderComponent([createReport()]);

        expect(screen.getByText('大樹聽力中心診所轉介名單分析')).toBeInTheDocument();
        expect(screen.queryByText(/轉介成效回訪報告/)).not.toBeInTheDocument();
    });

    it('carries the clinic name and store into the exported header, without the period', () => {
        renderComponent([createReport()]);

        const header = screen.getByText('大樹聽力中心診所轉介名單分析').closest('div');

        expect(header).toHaveStyle({ backgroundColor: 'rgb(0, 140, 215)' });
        expect(header).toHaveTextContent('診所名稱：康健診所');
        // 試算表叫「竹北店來客紀錄」，表頭只帶門市
        expect(header).toHaveTextContent('服務門市：竹北店');
        expect(header).not.toHaveTextContent('來客紀錄');
        expect(header).not.toHaveTextContent('分析區間');
    });

    it('keeps the fee KPI tiles on screen only, leaving just the table in the export', () => {
        renderComponent([createReport()]);

        const tiles = screen.getByText('已配戴人數').closest('div')?.parentElement;

        expect(tiles).toHaveClass('print:hidden');
        // 費用明細表本身仍會匯出
        expect(screen.getByText('轉介費用 (10%)')).toBeInTheDocument();
    });

    it('hides the deal amount total from the export but keeps the fee total', () => {
        renderComponent([createReport()]);

        expect(screen.getByTestId('fee-total-amount')).toHaveClass('print:hidden');
        // 合計轉介費用不加 print:hidden，匯出時仍會顯示
        expect(screen.getAllByText('NT$20,000')).toHaveLength(2);
    });

    it('keeps the talking points out of the exported report', () => {
        renderComponent([createReport()]);

        expect(screen.getByText('回訪重點').closest('div')).toHaveClass('print:hidden');
    });

    it('exports the referral fee block when there is a fee to pay', () => {
        renderComponent([createReport()]);

        expect(screen.getByTestId('referral-fee-block')).not.toHaveClass('print:hidden');
    });

    it('drops the referral fee block from the export when no fee is owed', () => {
        renderComponent([secondReport]);

        expect(screen.getByTestId('referral-fee-block')).toHaveClass('print:hidden');
    });

    it('drops the referral fee block when every dealt patient has no amount', () => {
        renderComponent([createReport({ patients: [createPatient({ amount: 0 })] })]);

        expect(screen.getByTestId('referral-fee-block')).toHaveClass('print:hidden');
    });
});

describe('formatStoreName', () => {
    it('strips the worksheet name from the spreadsheet title', () => {
        expect(formatStoreName('竹北店來客紀錄')).toBe('竹北店');
        expect(formatStoreName('湖口店 來客紀錄')).toBe('湖口店');
        expect(formatStoreName('大樹聽力中心-湖口店-來客紀錄')).toBe('大樹聽力中心-湖口店');
    });

    it('leaves a title that has no worksheet name untouched', () => {
        expect(formatStoreName('竹北店')).toBe('竹北店');
    });

    it('returns an empty string when nothing but the worksheet name is left', () => {
        expect(formatStoreName('來客紀錄')).toBe('');
        expect(formatStoreName('')).toBe('');
    });
});

describe('maskCustomerName', () => {
    it('keeps the first and last characters and masks the middle', () => {
        expect(maskCustomerName('王大明')).toBe('王○明');
        expect(maskCustomerName('歐陽大明')).toBe('歐○○明');
    });

    it('masks only the trailing character for two-character names', () => {
        expect(maskCustomerName('林安')).toBe('林○');
    });

    it('leaves single-character and empty names untouched', () => {
        expect(maskCustomerName('陳')).toBe('陳');
        expect(maskCustomerName('')).toBe('');
    });
});

describe('daysSinceDate', () => {
    it('returns null for empty or unparseable dates', () => {
        expect(daysSinceDate('')).toBeNull();
        expect(daysSinceDate('not-a-date')).toBeNull();
    });

    it('counts whole days since the given date', () => {
        const now = new Date();
        const sevenDaysAgo = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 7);
        const iso = `${sevenDaysAgo.getFullYear()}-${String(sevenDaysAgo.getMonth() + 1).padStart(2, '0')}-${String(sevenDaysAgo.getDate()).padStart(2, '0')}`;

        expect(daysSinceDate(iso)).toBe(7);
    });

    it('never returns a negative value for future dates', () => {
        const now = new Date();
        const nextWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 7);
        const iso = `${nextWeek.getFullYear()}-${String(nextWeek.getMonth() + 1).padStart(2, '0')}-${String(nextWeek.getDate()).padStart(2, '0')}`;

        expect(daysSinceDate(iso)).toBe(0);
    });
});
