import React from 'react';
import { SheetInfo, DateRange } from '../types';

interface ControlsProps {
    sheetUrl: string;
    setSheetUrl: (url: string) => void;
    handleSheetUrlSubmit: (e: React.FormEvent) => void;
    availableSheets: SheetInfo[];
    selectedSheet: string;
    handleSheetChange: (e: React.ChangeEvent<HTMLSelectElement>) => void;
    dateRange: DateRange;
    setDateRange: React.Dispatch<React.SetStateAction<DateRange>>;
    spreadsheetTitle: string;
    ptaThreshold: number;
    setPtaThreshold: (threshold: number) => void;
    onGenerateAnalysis: () => void;
}

export const Controls: React.FC<ControlsProps> = ({
    sheetUrl,
    setSheetUrl,
    handleSheetUrlSubmit,
    availableSheets,
    selectedSheet,
    handleSheetChange,
    dateRange,
    setDateRange,
    spreadsheetTitle,
    ptaThreshold,
    setPtaThreshold,
    onGenerateAnalysis
}) => {
    return (
        <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-100 mb-8 print:hidden transition-all">
            <div className="flex items-center justify-between mb-4">
                <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                    <svg className="w-5 h-5 text-slate-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 6V4m0 2a2 2 0 100 4m0-4a2 2 0 110 4m-6 8a2 2 0 100-4m0 4a2 2 0 110-4m0 4v2m0-6V4m6 6v10m6-2a2 2 0 100-4m0 4a2 2 0 110-4m0 4v2m0-6V4" />
                    </svg>
                    資料分析設定
                </h3>
                <form onSubmit={handleSheetUrlSubmit} className="flex gap-2 max-w-lg w-full">
                    <input
                        type="text"
                        value={sheetUrl}
                        onChange={(e) => setSheetUrl(e.target.value)}
                        placeholder="輸入 Google Sheets 網址..."
                        className="flex-1 rounded-lg border-slate-200 text-sm focus:border-blue-500 focus:ring-blue-500 p-2 border"
                    />
                    <button
                        type="submit"
                        className="px-4 py-2 bg-slate-800 text-white rounded-lg hover:bg-slate-700 text-sm font-medium transition-colors"
                    >
                        讀取
                    </button>
                </form>
            </div>

            <hr className="border-slate-100 my-4" />

            {availableSheets.length > 0 ? (
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6 items-end">
                    <div>
                        <label className="block text-xs font-semibold text-slate-500 mb-1 uppercase tracking-wider">
                            選擇工作表
                        </label>
                        <select
                            value={selectedSheet}
                            onChange={handleSheetChange}
                            className="block w-full rounded-lg border-slate-200 text-sm focus:border-blue-500 focus:ring-blue-500 p-2.5 border bg-slate-50"
                        >
                            <option value="">請選擇...</option>
                            {availableSheets.map((sheet) => (
                                <option key={sheet.properties.sheetId} value={sheet.properties.title}>
                                    {sheet.properties.title}
                                </option>
                            ))}
                        </select>
                    </div>

                    <div className="md:col-span-2">
                        <label className="block text-xs font-semibold text-slate-500 mb-1 uppercase tracking-wider">分析區間 (起~迄)</label>
                        <div className="flex items-center gap-2 bg-slate-50 p-1.5 rounded-lg border border-slate-200">
                            <input
                                type="number"
                                className="w-20 rounded border-slate-200 p-1 text-center text-sm"
                                value={dateRange.startYear}
                                onChange={e => setDateRange(p => ({ ...p, startYear: parseInt(e.target.value) }))}
                            />
                            <span className="text-slate-400 text-xs">年</span>
                            <input
                                type="number"
                                className="w-16 rounded border-slate-200 p-1 text-center text-sm"
                                value={dateRange.startMonth}
                                min={1} max={12}
                                onChange={e => setDateRange(p => ({ ...p, startMonth: parseInt(e.target.value) }))}
                            />
                            <span className="text-slate-400 text-xs">月</span>
                            <span className="text-slate-300">➜</span>
                            <input
                                type="number"
                                className="w-20 rounded border-slate-200 p-1 text-center text-sm"
                                value={dateRange.endYear}
                                onChange={e => setDateRange(p => ({ ...p, endYear: parseInt(e.target.value) }))}
                            />
                            <span className="text-slate-400 text-xs">年</span>
                            <input
                                type="number"
                                className="w-16 rounded border-slate-200 p-1 text-center text-sm"
                                value={dateRange.endMonth}
                                min={1} max={12}
                                onChange={e => setDateRange(p => ({ ...p, endMonth: parseInt(e.target.value) }))}
                            />
                            <span className="text-slate-400 text-xs">月</span>
                        </div>
                    </div>

                    <div>
                        <label className="block text-xs font-semibold text-slate-500 mb-1 uppercase tracking-wider">目標門市</label>
                        <div className="w-full rounded-lg border border-slate-200 bg-slate-100 p-2.5 text-sm text-slate-700">
                            {spreadsheetTitle || '未設定'}
                        </div>
                    </div>

                    <div className="flex gap-2">
                        <div className="flex-1">
                            <label className="block text-xs font-semibold text-slate-500 mb-1 uppercase tracking-wider">PTA 閾值</label>
                            <input
                                type="number"
                                className="w-full rounded-lg border-slate-200 p-2.5 text-sm border bg-slate-50"
                                value={ptaThreshold}
                                onChange={e => setPtaThreshold(parseInt(e.target.value))}
                            />
                        </div>
                        <button
                            onClick={onGenerateAnalysis}
                            className="flex-[2] bg-blue-600 text-white rounded-lg hover:bg-blue-700 font-medium text-sm transition-all shadow-md hover:shadow-lg active:scale-95 flex items-center justify-center gap-2"
                        >
                            <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2" />
                            </svg>
                            生成報告
                        </button>
                    </div>
                </div>
            ) : (
                <div className="text-center py-8 text-slate-400">
                    <p>請先輸入 Google Sheets 網址以載入資料</p>
                </div>
            )}
        </div>
    );
};
