import React from 'react';
import { CustomerAnalysis } from '../types';

interface DashboardStatsProps {
    analysis: CustomerAnalysis;
}

const StatCard: React.FC<{
    title: string;
    value: string | number;
    subtext: string;
    colorClass: string;
    icon?: React.ReactNode;
}> = ({ title, value, subtext, colorClass, icon }) => (
    <div className={`bg-white rounded-xl shadow-sm border border-slate-100 p-6 flex flex-col transition-all hover:shadow-md ${colorClass}`}>
        <div className="flex justify-between items-start mb-4">
            <h3 className="text-slate-500 text-sm font-semibold uppercase tracking-wider">{title}</h3>
            {icon && <div className="p-2 rounded-lg bg-slate-50">{icon}</div>}
        </div>
        <div className="mt-auto">
            <p className="text-3xl font-bold text-slate-800 tracking-tight">{value}</p>
            <p className="text-xs text-slate-400 mt-1 font-medium">{subtext}</p>
        </div>
    </div>
);

export const DashboardStats: React.FC<DashboardStatsProps> = ({ analysis }) => {
    // Calculate Average Order Value
    const averageOrderValue = analysis.totalCompletedDeals > 0
        ? Math.round(analysis.totalAmount / analysis.totalCompletedDeals)
        : 0;

    return (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-5 gap-6 mb-8 print:grid-cols-5 print:gap-4">
            <StatCard
                title="總潛力客戶數"
                value={analysis.totalCustomers}
                subtext="符合PTA標準或已成交"
                colorClass="border-l-4 border-l-blue-500"
                icon={
                    <svg className="w-5 h-5 text-blue-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M17 20h5v-2a3 3 0 00-5.356-1.857M17 20H7m10 0v-2c0-.656-.126-1.283-.356-1.857M7 20H2v-2a3 3 0 015.356-1.857M7 20v-2c0-.656.126-1.283.356-1.857m0 0a5.002 5.002 0 019.288 0M15 7a3 3 0 11-6 0 3 3 0 016 0zm6 3a2 2 0 11-4 0 2 2 0 014 0z" />
                    </svg>
                }
            />

            <StatCard
                title="總成交數"
                value={analysis.totalCompletedDeals}
                subtext="已確認成交訂單"
                colorClass="border-l-4 border-l-green-500"
                icon={
                    <svg className="w-5 h-5 text-green-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                    </svg>
                }
            />

            <StatCard
                title="整體成交率"
                value={`${analysis.overallConversionRate.toFixed(1)}%`}
                subtext="成交數 / 潛力客戶"
                colorClass="border-l-4 border-l-purple-500"
                icon={
                    <svg className="w-5 h-5 text-purple-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 7h8m0 0v8m0-8l-8 8-4-4-6 6" />
                    </svg>
                }
            />

            <StatCard
                title="總營業額"
                value={`$${analysis.totalAmount.toLocaleString()}`}
                subtext="當期累積金額"
                colorClass="border-l-4 border-l-yellow-500"
                icon={
                    <svg className="w-5 h-5 text-yellow-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8c-1.657 0-3 .895-3 2s1.343 2 3 2 3 .895 3 2-1.343 2-3 2m0-8c1.11 0 2.08.402 2.599 1M12 8V7m0 1v8m0 0v1m0-1c-1.11 0-2.08-.402-2.599-1M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                    </svg>
                }
            />

            <StatCard
                title="平均客單價 (AOV)"
                value={`$${averageOrderValue.toLocaleString()}`}
                subtext="營業額 / 成交數"
                colorClass="border-l-4 border-l-pink-500"
                icon={
                    <svg className="w-5 h-5 text-pink-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M17 9V7a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2m2 4h10a2 2 0 002-2v-6a2 2 0 00-2-2H9a2 2 0 00-2 2v6a2 2 0 002 2zm7-5a2 2 0 11-4 0 2 2 0 014 0z" />
                    </svg>
                }
            />
        </div>
    );
};
