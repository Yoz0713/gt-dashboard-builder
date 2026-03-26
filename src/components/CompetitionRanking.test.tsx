import React from 'react';
import { render, screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { CompetitionRanking } from './CompetitionRanking';
import { SpreadsheetListItem } from '../types';

describe('CompetitionRanking', () => {
    const savedSpreadsheets: SpreadsheetListItem[] = Array.from({ length: 11 }, (_, index) => ({
        id: `sheet-${index + 1}`,
        title: `試算表 ${index + 1}`,
        hasCustomerRecord: true,
    }));

    it('renders the competition-only spreadsheet selector above the ranking table', () => {
        render(
            <CompetitionRanking
                entries={[
                    { rank: 1, storeName: 'Alpha', totalReferrals: 3, hearingLossCustomers: 1, normalCustomers: 2 },
                ]}
                savedSpreadsheets={savedSpreadsheets.slice(0, 2)}
                selectedCompetitionSpreadsheetIds={['sheet-1']}
                competitionSelectionLimit={10}
                competitionSelectionMessage=""
            />
        );

        expect(screen.getByText(/Competition Only/i)).toBeInTheDocument();
        expect(screen.getByText(/這個選取區只影響競賽排行/i)).toBeInTheDocument();
        expect(screen.getByRole('heading', { name: '競賽排行資料範圍' })).toBeInTheDocument();
        expect(screen.getByRole('heading', { name: '競賽排行' })).toBeInTheDocument();
    });

    it('disables unchecked spreadsheets when the selection limit is reached and keeps checked ones interactive', async () => {
        const user = userEvent.setup();
        const onToggleCompetitionSpreadsheet = jest.fn();

        render(
            <CompetitionRanking
                entries={[]}
                savedSpreadsheets={savedSpreadsheets}
                selectedCompetitionSpreadsheetIds={['sheet-1', 'sheet-2', 'sheet-3', 'sheet-4', 'sheet-5', 'sheet-6', 'sheet-7', 'sheet-8', 'sheet-9', 'sheet-10']}
                onToggleCompetitionSpreadsheet={onToggleCompetitionSpreadsheet}
                competitionSelectionLimit={10}
                competitionSelectionMessage="競賽排行最多只能選擇 10 份試算表。"
            />
        );

        const checkedSheet = screen.getByRole('checkbox', { name: /試算表 1 納入此試算表的門市轉介資料做整合排行/i });
        const uncheckedSheet = screen.getByRole('checkbox', { name: /試算表 11 納入此試算表的門市轉介資料做整合排行/i });

        expect(checkedSheet).toBeEnabled();
        expect(uncheckedSheet).toBeDisabled();
        expect(screen.getByText(/最多只能選擇 10 份試算表/i)).toBeInTheDocument();

        await user.click(checkedSheet);

        expect(onToggleCompetitionSpreadsheet).toHaveBeenCalledWith('sheet-1');
    });
});
