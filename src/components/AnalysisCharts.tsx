import React from 'react';
import {
    Chart as ChartJS,
    CategoryScale,
    LinearScale,
    BarElement,
    Title,
    Tooltip,
    Legend,
    ArcElement,
    PointElement,
    LineElement,
} from 'chart.js';
import { Bar, Doughnut } from 'react-chartjs-2';
import { AnalysisResult } from '../utils/analysis';

import ChartDataLabels from 'chartjs-plugin-datalabels';

// Register Chart.js components
ChartJS.register(
    CategoryScale,
    LinearScale,
    BarElement,
    Title,
    Tooltip,
    Legend,
    ArcElement,
    PointElement,
    LineElement,
    ChartDataLabels
);

// Common Chart Colors
const COLORS = {
    primary: '#0ea5e9', // Sky 500
    secondary: '#10b981', // Emerald 500
    accent: '#f59e0b', // Amber 500
    purple: '#8b5cf6', // Violet 500
    slate: '#64748b', // Slate 500
    red: '#ef4444', // Red 500
};

// Common Options
const commonOptions = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
        legend: { position: 'bottom' as const, labels: { usePointStyle: true, padding: 20 } },
        datalabels: {
            color: '#334155', // Slate 700
            font: { weight: 'bold' as const, size: 12 },
            anchor: 'end' as const,
            align: 'top' as const,
            offset: 4,
            formatter: (value: number) => value > 0 ? value : '', // Don't show 0
        },
    },
    scales: {
        y: { beginAtZero: true, grace: '10%' } // Add some space at top for labels
    }
};

interface AnalysisChartsProps {
    analysisResult: AnalysisResult;
}

export const AnalysisCharts: React.FC<AnalysisChartsProps> = ({ analysisResult }) => {
    const { customerAnalysis, salesmenAnalysis, clinicAnalysis, storeReferralAnalysis } = analysisResult;

    // --- 1. 月份報告圖表 (Bar + Line) ---
    const monthlyChartData = {
        labels: customerAnalysis.monthlyData.map(item => item.month),
        datasets: [
            {
                type: 'bar' as const,
                label: '來客數',
                data: customerAnalysis.monthlyData.map(item => item.newCustomers),
                backgroundColor: 'rgba(14, 165, 233, 0.7)', // Sky 500
                borderRadius: 4,
            },
            {
                type: 'bar' as const,
                label: '成交數',
                data: customerAnalysis.monthlyData.map(item => item.completedDeals),
                backgroundColor: 'rgba(16, 185, 129, 0.7)', // Emerald 500
                borderRadius: 4,
            },
        ],
    };

    // --- 2. 年齡分佈 (Bar) ---
    const ageChartData = {
        labels: customerAnalysis.ageAnalysis.map(a => a.range),
        datasets: [{
            label: '人數',
            data: customerAnalysis.ageAnalysis.map(a => a.count),
            backgroundColor: 'rgba(139, 92, 246, 0.6)', // Violet
            borderColor: COLORS.purple,
            borderWidth: 1,
            borderRadius: 4,
        }]
    };

    // --- 3. 來源分析 (Horizontal Bar) ---
    // 取前 8 個來源以免太擠
    const topSources = customerAnalysis.sourceAnalysis.slice(0, 8);
    const sourceChartData = {
        labels: topSources.map(s => s.source),
        datasets: [{
            label: '人數',
            data: topSources.map(s => s.count),
            backgroundColor: 'rgba(245, 158, 11, 0.6)', // Amber
            borderColor: COLORS.accent,
            borderWidth: 1,
            borderRadius: 4,
        }]
    };
    const sourceChartOptions = {
        ...commonOptions,
        indexAxis: 'y' as const,
        plugins: {
            ...commonOptions.plugins,
            datalabels: {
                ...commonOptions.plugins.datalabels,
                anchor: 'end' as const,
                align: 'end' as const, // Put label to the right of the bar
                offset: 4,
            }
        },
        scales: {
            x: { beginAtZero: true, grace: '10%' }
        }
    };

    // --- 4. 聽損程度 (Doughnut) ---
    const hearingLossData = {
        labels: customerAnalysis.hearingLossAnalysis.map(h => h.degree),
        datasets: [{
            data: customerAnalysis.hearingLossAnalysis.map(h => h.count),
            backgroundColor: [
                '#10b981', // 正常 - Green
                '#0ea5e9', // 輕度 - Blue
                '#f59e0b', // 中度 - Amber
                '#f97316', // 中重度 - Orange
                '#ef4444', // 重度 - Red
                '#7f1d1d', // 極重度 - Dark Red
                '#94a3b8', // 未知 - Slate
            ],
            borderWidth: 0,
        }]
    };

    // --- 5. 診所轉介 (Bar) ---
    const topClinics = clinicAnalysis.slice(0, 8); // Top 8
    const clinicChartData = {
        labels: topClinics.map((c: any) => c.clinic),
        datasets: [
            {
                label: '轉介數',
                data: topClinics.map((c: any) => c.total),
                backgroundColor: 'rgba(14, 165, 233, 0.6)',
                borderRadius: 4,
            },
            {
                label: '成交數',
                data: topClinics.map((c: any) => c.dealt),
                backgroundColor: 'rgba(16, 185, 129, 0.6)',
                borderRadius: 4,
            },
        ],
    };

    // --- 6. 門市轉介 (Bar) ---
    const topStores = storeReferralAnalysis.slice(0, 8); // Top 8
    const storeChartData = {
        labels: topStores.map((s: any) => s.store),
        datasets: [
            {
                label: '來客數',
                data: topStores.map((s: any) => s.total),
                backgroundColor: 'rgba(59, 130, 246, 0.6)', // Blue
                borderRadius: 4,
            },
            {
                label: '成交數',
                data: topStores.map((s: any) => s.dealt),
                backgroundColor: 'rgba(16, 185, 129, 0.6)', // Green
                borderRadius: 4,
            },
        ],
    };

    // --- 8. 平日/假日 (Bar) ---
    const weekdayChartData = {
        labels: customerAnalysis.weekdayAnalysis?.map(w => w.day) || [],
        datasets: [
            {
                label: '來客數',
                data: customerAnalysis.weekdayAnalysis?.map(w => w.visits) || [],
                backgroundColor: 'rgba(59, 130, 246, 0.7)',
                borderRadius: 4,
            },
            {
                label: '成交數',
                data: customerAnalysis.weekdayAnalysis?.map(w => w.deals) || [],
                backgroundColor: 'rgba(16, 185, 129, 0.7)',
                borderRadius: 4,
            }
        ]
    };

    return (
        <div className="space-y-6 print:space-y-4" >
            {/* 1. 營運趨勢 & 年齡分佈 */}
            < div className="grid grid-cols-1 lg:grid-cols-3 gap-6 print:grid-cols-2" >
                <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-6 lg:col-span-2">
                    <h2 className="text-lg font-bold text-slate-800 mb-4 border-l-4 border-sky-500 pl-3">月份營運趨勢</h2>
                    <div className="h-64">
                        <Bar data={monthlyChartData} options={commonOptions} />
                    </div>
                </div>

                <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-6">
                    <h2 className="text-lg font-bold text-slate-800 mb-4 border-l-4 border-violet-500 pl-3">客群年齡分佈</h2>
                    <div className="h-64">
                        <Bar data={ageChartData} options={commonOptions} />
                    </div>
                </div>
            </div >

            {/* 2. 來源 & 聽損 */}
            < div className="grid grid-cols-1 md:grid-cols-2 gap-6" >
                <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-6">
                    <h2 className="text-lg font-bold text-slate-800 mb-4 border-l-4 border-amber-500 pl-3">主要客源分析</h2>
                    <div className="h-64">
                        <Bar data={sourceChartData} options={sourceChartOptions} />
                    </div>
                </div>

                <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-6">
                    <h2 className="text-lg font-bold text-slate-800 mb-4 border-l-4 border-red-500 pl-3">聽損程度佔比</h2>
                    <div className="h-64 relative">
                        <Doughnut
                            data={hearingLossData}
                            options={{
                                ...commonOptions,
                                scales: {}, // Remove scales for doughnut
                                plugins: {
                                    ...commonOptions.plugins,
                                    datalabels: {
                                        color: '#fff',
                                        font: { weight: 'bold', size: 12 },
                                        formatter: (value, ctx) => {
                                            if (value === 0) return '';
                                            const total = ctx.chart.data.datasets[0].data.reduce((a: any, b: any) => a + b, 0);
                                            const percentageVal = (value / total) * 100;
                                            if (percentageVal < 5) return '';
                                            return `${value}\n(${percentageVal.toFixed(1)}%)`;
                                        },
                                        anchor: 'center',
                                        align: 'center',
                                        offset: 0
                                    }
                                }
                            }}
                        />
                    </div>
                </div>
            </div >

            {/* 3. 業務員 & 診所 (表格/圖表) */}
            < div className="grid grid-cols-1 lg:grid-cols-2 gap-6 print:break-before-page" >
                {/* 業務員表格 */}
                < div className="bg-white rounded-xl shadow-sm border border-slate-100 overflow-hidden" >
                    <div className="p-6 border-b border-slate-100">
                        <h2 className="text-lg font-bold text-slate-800 border-l-4 border-emerald-500 pl-3">業務員業績表</h2>
                    </div>


                    <div className="overflow-x-auto">
                        <table className="min-w-full divide-y divide-slate-100">
                            <thead className="bg-slate-50">
                                <tr>
                                    <th className="px-6 py-3 text-left text-xs font-semibold text-slate-500 uppercase">業務員</th>
                                    <th className="px-6 py-3 text-left text-xs font-semibold text-slate-500 uppercase">潛客</th>
                                    <th className="px-6 py-3 text-left text-xs font-semibold text-slate-500 uppercase">成交</th>
                                    <th className="px-6 py-3 text-left text-xs font-semibold text-slate-500 uppercase">轉化率</th>
                                    <th className="px-6 py-3 text-left text-xs font-semibold text-slate-500 uppercase">業績</th>
                                </tr>
                            </thead>
                            <tbody className="bg-white divide-y divide-slate-100 text-sm">
                                {Object.values(salesmenAnalysis).map((salesman: any) => (
                                    <tr key={salesman.業務員} className="hover:bg-slate-50 transition-colors">
                                        <td className="px-6 py-4 font-medium text-slate-900">{salesman.業務員}</td>
                                        <td className="px-6 py-4 text-slate-600">{salesman.潛力客戶數}</td>
                                        <td className="px-6 py-4 text-slate-600">{salesman.訂單數量}</td>
                                        <td className="px-6 py-4 font-semibold text-blue-600">
                                            {salesman.成交率.toFixed(1)}%
                                        </td>
                                        <td className="px-6 py-4 text-slate-600 font-mono">
                                            ${salesman.當季業績累積.toLocaleString()}
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                </div >

                {/* 診所圖表 */}
                < div className="bg-white rounded-xl shadow-sm border border-slate-100 p-6" >
                    <h2 className="text-lg font-bold text-slate-800 mb-4 border-l-4 border-sky-500 pl-3">熱門轉介診所 (Top 8)</h2>
                    <div className="h-[300px]">
                        <Bar data={clinicChartData} options={commonOptions} />
                    </div>
                </div >
            </div >

            {/* 4. 新增分析 (平日假日/門市轉介) */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 print:break-inside-avoid">
                <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-6">
                    <h2 className="text-lg font-bold text-slate-800 mb-4 border-l-4 border-blue-500 pl-3">每日來客趨勢</h2>
                    <div className="h-64">
                        <Bar data={weekdayChartData} options={commonOptions} />
                    </div>
                </div>

                <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-6">
                    <h2 className="text-lg font-bold text-slate-800 mb-4 border-l-4 border-indigo-500 pl-3">熱門門市轉介</h2>
                    <div className="h-64">
                        <Bar data={storeChartData} options={commonOptions} />
                    </div>
                </div>
            </div>
        </div >
    );
};
