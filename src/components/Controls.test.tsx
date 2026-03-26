import React from 'react';
import { render, screen } from '@testing-library/react';
import { Controls } from './Controls';
import { DateRange, SpreadsheetListItem } from '../types';

describe('Controls', () => {
    const savedSpreadsheets: SpreadsheetListItem[] = [
        { id: 'sheet-1', title: '試算表 1', hasCustomerRecord: true },
        { id: 'sheet-2', title: '試算表 2', hasCustomerRecord: true },
    ];

    const dateRange: DateRange = {
        startYear: 2025,
        startMonth: 1,
        endYear: 2025,
        endMonth: 12,
    };

    it('keeps only the main report settings and does not render competition-only selection UI', () => {
        render(
            <Controls
                sheetUrl=""
                setSheetUrl={jest.fn()}
                handleSheetUrlSubmit={jest.fn((event) => event.preventDefault())}
                availableSheets={[]}
                selectedSheet=""
                handleSheetChange={jest.fn()}
                dateRange={dateRange}
                setDateRange={jest.fn()}
                spreadsheetTitle="主試算表"
                ptaThreshold={40}
                setPtaThreshold={jest.fn()}
                onGenerateAnalysis={jest.fn()}
                savedSpreadsheets={savedSpreadsheets}
                onSelectSavedSpreadsheet={jest.fn()}
                hasData
            />
        );

        expect(screen.getByText(/已儲存的試算表/i)).toBeInTheDocument();
        expect(screen.queryByText(/競賽排行資料範圍/i)).not.toBeInTheDocument();
        expect(screen.queryByText(/Competition Only/i)).not.toBeInTheDocument();
    });
});
