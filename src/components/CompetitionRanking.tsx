import React from 'react';
import { CompetitionRankingEntry } from '../types';

interface CompetitionRankingProps {
    entries: CompetitionRankingEntry[];
}

export const CompetitionRanking: React.FC<CompetitionRankingProps> = ({ entries }) => {
    if (!entries.length) {
        return (
            <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-8 text-center text-slate-500">
                目前分析區間內沒有門市轉介排行資料。
            </div>
        );
    }

    return (
        <div className="bg-white rounded-xl shadow-sm border border-slate-100 overflow-hidden">
            <div className="p-6 border-b border-slate-100">
                <h2 className="text-lg font-bold text-slate-800 border-l-4 border-orange-500 pl-3">競賽排行</h2>
                <p className="text-sm text-slate-500 mt-2">
                    因會有距離加成，所以此排行不代表最後競賽的排名!
                </p>
            </div>

            <div className="overflow-x-auto">
                <table className="min-w-full divide-y divide-slate-100">
                    <thead className="bg-slate-50">
                        <tr>
                            <th className="px-6 py-3 text-left text-xs font-semibold text-slate-500 uppercase">名次</th>
                            <th className="px-6 py-3 text-left text-xs font-semibold text-slate-500 uppercase">轉介門市</th>
                            <th className="px-6 py-3 text-right text-xs font-semibold text-slate-500 uppercase">聽力正常</th>
                            <th className="px-6 py-3 text-right text-xs font-semibold text-slate-500 uppercase">聽力異常</th>
                        </tr>
                    </thead>
                    <tbody className="bg-white divide-y divide-slate-100 text-sm">
                        {entries.map((entry) => (
                            <tr key={entry.storeName} className="hover:bg-slate-50 transition-colors">
                                <td className="px-6 py-4 font-semibold text-slate-900">{entry.rank}</td>
                                <td className="px-6 py-4 text-slate-700">{entry.storeName}</td>
                                <td className="px-6 py-4 text-right text-emerald-600 font-semibold">{entry.potentialCustomers}</td>
                                <td className="px-6 py-4 text-right text-slate-600">{entry.nonPotentialCustomers}</td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>
        </div>
    );
};
