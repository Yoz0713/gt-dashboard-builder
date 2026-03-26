import React from 'react';
import {
    CompetitionRankingEntry,
    CompetitionRankingSkippedSpreadsheet,
    SpreadsheetListItem,
} from '../types';

interface CompetitionRankingProps {
    entries: CompetitionRankingEntry[];
    loading?: boolean;
    error?: string;
    selectedSpreadsheetCount?: number;
    skippedSpreadsheets?: CompetitionRankingSkippedSpreadsheet[];
    savedSpreadsheets?: SpreadsheetListItem[];
    selectedCompetitionSpreadsheetIds?: string[];
    onToggleCompetitionSpreadsheet?: (id: string) => void;
    competitionSelectionLimit?: number;
    competitionSelectionMessage?: string;
}

export const CompetitionRanking: React.FC<CompetitionRankingProps> = ({
    entries,
    loading = false,
    error = '',
    selectedSpreadsheetCount = 0,
    skippedSpreadsheets = [],
    savedSpreadsheets = [],
    selectedCompetitionSpreadsheetIds = [],
    onToggleCompetitionSpreadsheet,
    competitionSelectionLimit = 10,
    competitionSelectionMessage = '',
}) => {
    const isCompetitionSelectionAtLimit = selectedCompetitionSpreadsheetIds.length >= competitionSelectionLimit;

    if (loading) {
        return (
            <div className="space-y-4">
                <CompetitionSelectionPanel
                    savedSpreadsheets={savedSpreadsheets}
                    selectedCompetitionSpreadsheetIds={selectedCompetitionSpreadsheetIds}
                    onToggleCompetitionSpreadsheet={onToggleCompetitionSpreadsheet}
                    competitionSelectionLimit={competitionSelectionLimit}
                    competitionSelectionMessage={competitionSelectionMessage}
                    isCompetitionSelectionAtLimit={isCompetitionSelectionAtLimit}
                />
                <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-8 text-center text-slate-500">
                    競賽排行整合中，請稍候...
                </div>
            </div>
        );
    }

    if (error) {
        return (
            <div className="space-y-4">
                <CompetitionSelectionPanel
                    savedSpreadsheets={savedSpreadsheets}
                    selectedCompetitionSpreadsheetIds={selectedCompetitionSpreadsheetIds}
                    onToggleCompetitionSpreadsheet={onToggleCompetitionSpreadsheet}
                    competitionSelectionLimit={competitionSelectionLimit}
                    competitionSelectionMessage={competitionSelectionMessage}
                    isCompetitionSelectionAtLimit={isCompetitionSelectionAtLimit}
                />
                <div className="bg-red-50 rounded-xl border border-red-200 p-6 text-red-700">
                    {error}
                </div>
            </div>
        );
    }

    if (!entries.length) {
        return (
            <div className="space-y-4">
                <CompetitionSelectionPanel
                    savedSpreadsheets={savedSpreadsheets}
                    selectedCompetitionSpreadsheetIds={selectedCompetitionSpreadsheetIds}
                    onToggleCompetitionSpreadsheet={onToggleCompetitionSpreadsheet}
                    competitionSelectionLimit={competitionSelectionLimit}
                    competitionSelectionMessage={competitionSelectionMessage}
                    isCompetitionSelectionAtLimit={isCompetitionSelectionAtLimit}
                />
                {skippedSpreadsheets.length > 0 && (
                    <div className="bg-amber-50 rounded-xl border border-amber-200 p-4 text-sm text-amber-700">
                        <p className="font-semibold">部分試算表未納入競賽排行</p>
                        <ul className="mt-2 list-disc pl-5">
                            {skippedSpreadsheets.map((sheet) => (
                                <li key={sheet.spreadsheetId}>{sheet.spreadsheetTitle}: {sheet.reason}</li>
                            ))}
                        </ul>
                    </div>
                )}
                <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-8 text-center text-slate-500">
                    {selectedSpreadsheetCount > 0
                        ? '目前分析區間內沒有門市轉介排行資料。'
                        : '請先選擇要納入競賽排行的已儲存試算表。'}
                </div>
            </div>
        );
    }

    return (
        <div className="space-y-4">
            <CompetitionSelectionPanel
                savedSpreadsheets={savedSpreadsheets}
                selectedCompetitionSpreadsheetIds={selectedCompetitionSpreadsheetIds}
                onToggleCompetitionSpreadsheet={onToggleCompetitionSpreadsheet}
                competitionSelectionLimit={competitionSelectionLimit}
                competitionSelectionMessage={competitionSelectionMessage}
                isCompetitionSelectionAtLimit={isCompetitionSelectionAtLimit}
            />
            {skippedSpreadsheets.length > 0 && (
                <div className="bg-amber-50 rounded-xl border border-amber-200 p-4 text-sm text-amber-700">
                    <p className="font-semibold">部分試算表未納入競賽排行</p>
                    <ul className="mt-2 list-disc pl-5">
                        {skippedSpreadsheets.map((sheet) => (
                            <li key={sheet.spreadsheetId}>{sheet.spreadsheetTitle}: {sheet.reason}</li>
                        ))}
                    </ul>
                </div>
            )}

            <div className="bg-white rounded-xl shadow-sm border border-slate-100 overflow-hidden">
                <div className="p-6 border-b border-slate-100">
                    <h2 className="text-lg font-bold text-slate-800 border-l-4 border-orange-500 pl-3">競賽排行</h2>
                    <p className="text-sm text-slate-500 mt-2">
                        此排行僅供參考，尚未加權門市距離和客人的聽力狀況，實際排名請以競賽公告為準。
                    </p>
                </div>

                <div className="overflow-x-auto">
                    <table className="min-w-full divide-y divide-slate-100">
                        <thead className="bg-slate-50">
                            <tr>
                                <th className="px-6 py-3 text-left text-xs font-semibold text-slate-500 uppercase">名次</th>
                                <th className="px-6 py-3 text-left text-xs font-semibold text-slate-500 uppercase">轉介門市</th>
                                <th className="px-6 py-3 text-right text-xs font-semibold text-slate-500 uppercase">總轉介人次</th>
                                <th className="px-6 py-3 text-right text-xs font-semibold text-slate-500 uppercase">聽損客</th>
                                <th className="px-6 py-3 text-right text-xs font-semibold text-slate-500 uppercase">正常客</th>
                            </tr>
                        </thead>
                        <tbody className="bg-white divide-y divide-slate-100 text-sm">
                            {entries.map((entry) => (
                                <tr key={entry.storeName} className="hover:bg-slate-50 transition-colors">
                                    <td className="px-6 py-4 font-semibold text-slate-900">{entry.rank}</td>
                                    <td className="px-6 py-4 text-slate-700">{entry.storeName}</td>
                                    <td className="px-6 py-4 text-right font-semibold text-slate-900">{entry.totalReferrals}</td>
                                    <td className="px-6 py-4 text-right text-amber-600 font-semibold">{entry.hearingLossCustomers}</td>
                                    <td className="px-6 py-4 text-right text-emerald-600 font-semibold">{entry.normalCustomers}</td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    );
};

interface CompetitionSelectionPanelProps {
    savedSpreadsheets: SpreadsheetListItem[];
    selectedCompetitionSpreadsheetIds: string[];
    onToggleCompetitionSpreadsheet?: (id: string) => void;
    competitionSelectionLimit: number;
    competitionSelectionMessage: string;
    isCompetitionSelectionAtLimit: boolean;
}

const CompetitionSelectionPanel: React.FC<CompetitionSelectionPanelProps> = ({
    savedSpreadsheets,
    selectedCompetitionSpreadsheetIds,
    onToggleCompetitionSpreadsheet,
    competitionSelectionLimit,
    competitionSelectionMessage,
    isCompetitionSelectionAtLimit,
}) => {
    if (!savedSpreadsheets.length) {
        return null;
    }

    return (
        <section className="rounded-2xl border border-orange-200 bg-gradient-to-br from-orange-50 via-white to-amber-50 p-5 shadow-sm">
            <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
                <div>
                    <p className="text-xs font-semibold uppercase tracking-[0.2em] text-orange-600">Competition Only</p>
                    <h2 className="mt-2 text-xl font-bold text-slate-900">競賽排行資料範圍</h2>
                    <p className="mt-2 max-w-2xl text-sm text-slate-600">
                        這個選取區只影響競賽排行，不會改變總覽分析與區間比較的主試算表結果。
                    </p>
                </div>
                <div className="rounded-full border border-orange-200 bg-white px-4 py-2 text-sm font-semibold text-orange-700">
                    已選 {selectedCompetitionSpreadsheetIds.length} / {competitionSelectionLimit}
                </div>
            </div>

            <div className="mt-4 grid gap-3 md:grid-cols-2 xl:grid-cols-3">
                {savedSpreadsheets.map((sheet) => {
                    const isChecked = selectedCompetitionSpreadsheetIds.includes(sheet.id);
                    const isDisabled = !isChecked && isCompetitionSelectionAtLimit;

                    return (
                        <label
                            key={sheet.id}
                            className={`group flex items-start gap-3 rounded-xl border px-4 py-3 transition-all ${isChecked
                                ? 'border-orange-300 bg-white shadow-sm ring-2 ring-orange-100'
                                : 'border-slate-200 bg-white/80'
                                } ${isDisabled ? 'cursor-not-allowed opacity-60' : 'cursor-pointer hover:border-orange-200 hover:bg-white'}`}
                        >
                            <input
                                type="checkbox"
                                checked={isChecked}
                                disabled={isDisabled}
                                onChange={() => onToggleCompetitionSpreadsheet?.(sheet.id)}
                                className="mt-1 h-4 w-4 rounded border-slate-300 text-orange-600 focus:ring-orange-500"
                            />
                            <span className="min-w-0 flex-1">
                                <span className="block truncate text-sm font-semibold text-slate-800">{sheet.title}</span>
                                <span className="mt-1 block text-xs text-slate-500">
                                    納入此試算表的門市轉介資料做整合排行
                                </span>
                            </span>
                        </label>
                    );
                })}
            </div>

            {competitionSelectionMessage && (
                <p className="mt-4 text-sm font-medium text-amber-700">{competitionSelectionMessage}</p>
            )}
        </section>
    );
};
