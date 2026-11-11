import React from 'react';
import { render, screen, within } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { FollowUpCustomers } from './FollowUpCustomers';
import { CompetitionSheetSource, DateRange, SpreadsheetListItem } from '../types';
import { downloadFollowUpExcel } from '../utils/followUpExport';

jest.mock('../utils/followUpExport', () => ({
    downloadFollowUpExcel: jest.fn(),
}));

const mockedDownloadFollowUpExcel = downloadFollowUpExcel as jest.MockedFunction<typeof downloadFollowUpExcel>;

const dateRange: DateRange = { startYear: 2025, startMonth: 1, endYear: 2025, endMonth: 12 };

const HEADERS = ['服務日期', '姓名', '年齡', '顧客來源', '左耳 PTA', '右耳 PTA', '狀態', '主聽力師', '備註', '手機'];

const hukou: CompetitionSheetSource = {
    spreadsheetId: 'hukou',
    spreadsheetTitle: '湖口店來客紀錄',
    sheetData: {
        values: [
            HEADERS,
            ['2025-01-05', '王大明', '70', '診所轉介', '60', '30', '', '林小美', '家人陪同再來', '0912-345-678'],
            ['2025-02-10', '陳美麗', '65', '網路廣告', '45', '42', '', '張大偉', ''],
            ['2025-03-01', '李成交', '68', '診所轉介', '80', '70', '成交', '林小美', ''],
            ['2025-04-01', '趙極重', '82', '路過', '95', '91', '', '張大偉', ''],
            ['2025-05-01', '孫輕微', '50', '路過', '30', '20', '', '林小美', ''],
        ],
    },
};

const zhubei: CompetitionSheetSource = {
    spreadsheetId: 'zhubei',
    spreadsheetTitle: '竹北店來客紀錄',
    sheetData: {
        values: [
            ['服務日期', '姓名', '左耳 PTA', '右耳 PTA', '是否成交', '主聽力師'],
            ['2025-06-01', '鄭竹北', '75', '50', '否', '黃小芬'],
        ],
    },
};

const savedSpreadsheets: SpreadsheetListItem[] = [
    { id: 'hukou', title: '湖口店來客紀錄', hasCustomerRecord: true },
    { id: 'zhubei', title: '竹北店來客紀錄', hasCustomerRecord: true },
];

type Props = React.ComponentProps<typeof FollowUpCustomers>;

const renderList = (overrides: Partial<Props> = {}) => {
    const props: Props = {
        sources: [hukou],
        currentSpreadsheetId: 'hukou',
        dateRange,
        ptaThreshold: 40,
        savedSpreadsheets: [],
        extraSpreadsheetIds: [],
        selectionLimit: 10,
        ...overrides,
    };

    return render(<FollowUpCustomers {...props} />);
};

const rowNames = () =>
    screen.queryAllByTestId('follow-up-row').map((row) => {
        // 第一欄（單店）為服務日期，第二欄為姓名
        const cells = within(row).getAllByRole('cell');
        return cells[1].textContent;
    });

const kpiValue = (testId: string) => within(screen.getByTestId(testId)).getAllByText(/.+/)[1].textContent;

describe('FollowUpCustomers', () => {
    it('lists only undealt customers whose worse-ear PTA exceeds the global threshold, sorted by PTA', () => {
        renderList();

        expect(rowNames()).toEqual(['趙極重', '王大明', '陳美麗']);
        expect(screen.queryByText('李成交')).not.toBeInTheDocument();
        expect(screen.queryByText('孫輕微')).not.toBeInTheDocument();
        expect(kpiValue('follow-up-kpi-count')).toBe('3');
        // 未成交 4 人中有 3 人達標
        expect(kpiValue('follow-up-kpi-rate')).toBe('75%');
        expect(kpiValue('follow-up-kpi-severe')).toBe('2');
    });

    it('updates the list and KPIs immediately when a quick threshold is chosen', async () => {
        const user = userEvent.setup();
        renderList();

        await user.click(screen.getByRole('button', { name: '55 dB' }));
        expect(rowNames()).toEqual(['趙極重', '王大明']);
        expect(kpiValue('follow-up-kpi-count')).toBe('2');
        expect(screen.getByRole('button', { name: '55 dB' })).toHaveAttribute('aria-pressed', 'true');

        await user.click(screen.getByRole('button', { name: '70 dB' }));
        expect(rowNames()).toEqual(['趙極重']);

        await user.click(screen.getByRole('button', { name: /重設為全域閾值/ }));
        expect(rowNames()).toHaveLength(3);
    });

    it('accepts a manually typed threshold', async () => {
        const user = userEvent.setup();
        renderList();

        const input = screen.getByLabelText(/追蹤 PTA 閾值/);
        await user.clear(input);
        await user.type(input, '90');

        expect(rowNames()).toEqual(['趙極重']);
    });

    it('follows changes to the global threshold', () => {
        const { rerender } = renderList();

        rerender(
            <FollowUpCustomers
                sources={[hukou]}
                currentSpreadsheetId="hukou"
                dateRange={dateRange}
                ptaThreshold={70}
                savedSpreadsheets={[]}
                extraSpreadsheetIds={[]}
                selectionLimit={10}
            />
        );

        expect(rowNames()).toEqual(['趙極重']);
    });

    it('filters by audiologist, degree, source and name keyword', async () => {
        const user = userEvent.setup();
        renderList();

        await user.selectOptions(screen.getByLabelText('主聽力師'), '張大偉');
        expect(rowNames()).toEqual(['趙極重', '陳美麗']);

        await user.selectOptions(screen.getByLabelText('聽損程度'), '中度');
        expect(rowNames()).toEqual(['陳美麗']);

        await user.click(screen.getByRole('button', { name: '清除篩選條件' }));
        expect(rowNames()).toHaveLength(3);

        await user.selectOptions(screen.getByLabelText('顧客來源／轉介'), '診所轉介');
        expect(rowNames()).toEqual(['王大明']);

        await user.click(screen.getByRole('button', { name: '清除篩選條件' }));
        await user.type(screen.getByLabelText('姓名關鍵字'), '美麗');
        expect(rowNames()).toEqual(['陳美麗']);
    });

    it('sorts by visit date in both directions', async () => {
        const user = userEvent.setup();
        renderList();

        await user.click(screen.getByRole('button', { name: /到店日期（近→遠）/ }));
        expect(rowNames()).toEqual(['趙極重', '陳美麗', '王大明']);

        await user.click(screen.getByRole('button', { name: /到店日期（久→近）/ }));
        expect(rowNames()).toEqual(['王大明', '陳美麗', '趙極重']);
    });

    it('masks customer names and phone numbers when the toggle is on', async () => {
        const user = userEvent.setup();
        renderList();

        expect(screen.getByText('0912-345-678')).toBeInTheDocument();

        await user.click(screen.getByLabelText('姓名／電話遮罩'));

        expect(rowNames()).toEqual(['趙○重', '王○明', '陳○麗']);
        expect(screen.queryByText('王大明')).not.toBeInTheDocument();
        expect(screen.getByText('0912-○○○-678')).toBeInTheDocument();
        expect(screen.queryByText('0912-345-678')).not.toBeInTheDocument();
    });

    it('shows age only, without the birthday', () => {
        renderList();

        expect(screen.getByRole('columnheader', { name: '年齡' })).toBeInTheDocument();
        expect(screen.queryByRole('columnheader', { name: /生日/ })).not.toBeInTheDocument();
        const wangRow = screen.getAllByTestId('follow-up-row').find((row) => within(row).queryByText('王大明'));
        expect(within(wangRow as HTMLElement).getByText('70 歲')).toBeInTheDocument();
    });

    it('shows the phone and note columns only when the sheet has them', () => {
        const { unmount } = renderList();
        expect(screen.getByRole('columnheader', { name: '電話' })).toBeInTheDocument();
        expect(screen.getByRole('columnheader', { name: '備註' })).toBeInTheDocument();
        expect(screen.getByText('家人陪同再來')).toBeInTheDocument();
        unmount();

        renderList({ sources: [zhubei], currentSpreadsheetId: 'zhubei' });
        expect(screen.queryByRole('columnheader', { name: '電話' })).not.toBeInTheDocument();
        expect(screen.queryByRole('columnheader', { name: '備註' })).not.toBeInTheDocument();
    });

    it('adds a store column and store filter when multiple spreadsheets are merged', async () => {
        const user = userEvent.setup();
        renderList({ sources: [hukou, zhubei], savedSpreadsheets, extraSpreadsheetIds: ['zhubei'] });

        expect(screen.getByRole('columnheader', { name: '門市' })).toBeInTheDocument();
        expect(screen.getAllByTestId('follow-up-row')).toHaveLength(4);

        await user.selectOptions(screen.getByLabelText('門市'), '竹北店');
        const rows = screen.getAllByTestId('follow-up-row');
        expect(rows).toHaveLength(1);
        expect(within(rows[0]).getByText('鄭竹北')).toBeInTheDocument();
        // 比例的分母改為竹北店的未成交客戶
        expect(kpiValue('follow-up-kpi-rate')).toBe('100%');
    });

    it('locks the current store in the merge panel and forwards toggles for others', async () => {
        const user = userEvent.setup();
        const onToggleSpreadsheet = jest.fn();
        renderList({ savedSpreadsheets, onToggleSpreadsheet });

        const currentCheckbox = screen.getByRole('checkbox', { name: /湖口店來客紀錄/ });
        expect(currentCheckbox).toBeChecked();
        expect(currentCheckbox).toBeDisabled();

        await user.click(screen.getByRole('checkbox', { name: /竹北店來客紀錄/ }));
        expect(onToggleSpreadsheet).toHaveBeenCalledWith('zhubei');
    });

    it('shows an empty state when nobody matches', async () => {
        const user = userEvent.setup();
        renderList();

        await user.type(screen.getByLabelText('姓名關鍵字'), '不存在');

        expect(screen.queryAllByTestId('follow-up-row')).toHaveLength(0);
        expect(screen.getByText(/目前沒有符合條件的待追蹤客戶/)).toBeInTheDocument();
    });

    it('warns when the sheet has no PTA columns', () => {
        renderList({
            sources: [{
                spreadsheetId: 'x',
                spreadsheetTitle: 'X',
                sheetData: { values: [['服務日期', '姓名'], ['2025-01-01', 'A']] },
            }],
        });

        expect(screen.getByText(/找不到「左耳 PTA」或「右耳 PTA」欄位/)).toBeInTheDocument();
    });

    it('hides controls when printing and shows a print-only report header', () => {
        renderList({ savedSpreadsheets });

        const printHeader = screen.getByText('大樹聽力中心未成交潛力客戶追蹤名單').closest('div');
        expect(printHeader).toHaveClass('hidden', 'print:block');
        expect(printHeader).toHaveTextContent('較差耳 PTA > 40 dB');
        expect(printHeader).toHaveTextContent('共 3 位');

        expect(screen.getByText('未成交潛力客戶追蹤').closest('section')).toHaveClass('print:hidden');
        expect(screen.getByText('合併分析範圍').closest('section')).toHaveClass('print:hidden');
    });

    it('exports the currently filtered, sorted list to Excel with the current mask state', async () => {
        const user = userEvent.setup();
        mockedDownloadFollowUpExcel.mockReset();
        mockedDownloadFollowUpExcel.mockResolvedValue(undefined);
        renderList();

        await user.click(screen.getByRole('button', { name: '55 dB' }));
        await user.click(screen.getByLabelText('姓名／電話遮罩'));
        await user.click(screen.getByRole('button', { name: /匯出 Excel/ }));

        expect(mockedDownloadFollowUpExcel).toHaveBeenCalledTimes(1);
        const [exportedRecords, options, meta] = mockedDownloadFollowUpExcel.mock.calls[0];
        expect(exportedRecords.map((record) => record.name)).toEqual(['趙極重', '王大明']);
        expect(options).toEqual({ maskNames: true, includePhone: true, includeNote: true });
        expect(meta).toEqual({ storeLabel: '湖口店' });
    });

    it('disables Excel export when the list is empty and shows export errors', async () => {
        const user = userEvent.setup();
        mockedDownloadFollowUpExcel.mockReset();
        mockedDownloadFollowUpExcel.mockRejectedValue(new Error('disk full'));
        renderList();

        await user.click(screen.getByRole('button', { name: /匯出 Excel/ }));
        expect(await screen.findByText('匯出 Excel 失敗：disk full')).toBeInTheDocument();

        await user.type(screen.getByLabelText('姓名關鍵字'), '不存在');
        expect(screen.getByRole('button', { name: /匯出 Excel/ })).toBeDisabled();
    });

    it('prints via window.print', async () => {
        const user = userEvent.setup();
        const printSpy = jest.spyOn(window, 'print').mockImplementation(() => undefined);
        renderList();

        await user.click(screen.getByRole('button', { name: /匯出待追蹤名單/ }));
        expect(printSpy).toHaveBeenCalled();

        printSpy.mockRestore();
    });
});
