import React from 'react';
import { render, screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import App from './App';
import { useGoogleSheetData } from './hooks/useGoogleSheetData';

jest.mock('./hooks/useGoogleSheetData');

jest.mock('./components/LoginView', () => ({
    LoginView: ({ onLogin }: { onLogin: () => void }) => <button onClick={onLogin}>login-view</button>,
}));

jest.mock('./components/Controls', () => ({
    Controls: () => <div>controls</div>,
}));

jest.mock('./components/DashboardStats', () => ({
    DashboardStats: () => <div>dashboard-stats</div>,
}));

jest.mock('./components/AnalysisCharts', () => ({
    AnalysisCharts: () => <div>analysis-charts</div>,
}));

jest.mock('./components/ComparisonView', () => ({
    ComparisonView: () => <div>comparison-view</div>,
}));

jest.mock('./components/CompetitionRanking', () => ({
    CompetitionRanking: () => <div>competition-ranking-view</div>,
}));

const mockedUseGoogleSheetData = useGoogleSheetData as jest.MockedFunction<typeof useGoogleSheetData>;

describe('App report tabs', () => {
    beforeEach(() => {
        mockedUseGoogleSheetData.mockReturnValue({
            user: { id: '1', name: 'Tester', email: 'test@example.com', picture: '' },
            accessToken: '',
            sheetData: { values: [['日期'], ['2025-01-01']] },
            availableSheets: [],
            selectedSheet: '',
            spreadsheetId: '',
            spreadsheetTitle: 'Test Sheet',
            loading: false,
            error: '',
            analysisResult: {
                customerAnalysis: {
                    monthlyData: [],
                    totalCustomers: 1,
                    totalCompletedDeals: 0,
                    totalAmount: 0,
                    overallConversionRate: 0,
                    dateRange: { earliest: '2025-01-01', latest: '2025-01-01' },
                    ageAnalysis: [],
                    sourceAnalysis: [],
                    hearingLossAnalysis: [],
                    weekdayAnalysis: [],
                    yearBuckets: {},
                    quarterBuckets: {},
                    monthBuckets: {},
                },
                salesmenAnalysis: {},
                clinicAnalysis: [],
                storeReferralAnalysis: [],
                hearingScreeningAnalysis: [],
                competitionRankingAnalysis: [],
            },
            savedSpreadsheets: [],
            setSpreadsheetId: jest.fn(),
            setSelectedSheet: jest.fn(),
            login: jest.fn(),
            logout: jest.fn(),
            loadSpreadsheetMetadata: jest.fn(),
            loadSheetData: jest.fn(),
            loadSavedSpreadsheets: jest.fn(),
            loadSavedSpreadsheet: jest.fn(),
            performAnalysis: jest.fn(),
        });
    });

    it('shows overview by default and allows switching to comparison and competition tabs', async () => {
        const user = userEvent.setup();
        render(<App />);

        expect(screen.getByText('dashboard-stats')).toBeInTheDocument();
        expect(screen.getByText('analysis-charts')).toBeInTheDocument();

        await user.click(screen.getByRole('button', { name: /區間比較/i }));
        expect(screen.getByText('comparison-view')).toBeInTheDocument();
        expect(screen.queryByText('dashboard-stats')).not.toBeInTheDocument();

        await user.click(screen.getByRole('button', { name: /競賽排行/i }));
        expect(screen.getByText('competition-ranking-view')).toBeInTheDocument();
        expect(screen.queryByText('comparison-view')).not.toBeInTheDocument();

        await user.click(screen.getByRole('button', { name: /總覽分析/i }));
        expect(screen.getByText('dashboard-stats')).toBeInTheDocument();
        expect(screen.getByText('analysis-charts')).toBeInTheDocument();
    });
});
