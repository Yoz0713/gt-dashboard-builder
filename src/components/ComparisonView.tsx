import React, { useState, useEffect } from 'react';
import { CustomerAnalysis, PeriodStats } from '../types';
import {
    Chart as ChartJS,
    CategoryScale,
    LinearScale,
    BarElement,
    LineController,
    LineElement,
    PointElement,
    Title,
    Tooltip,
    Legend
} from 'chart.js';
import { Bar } from 'react-chartjs-2';
import ChartDataLabels from 'chartjs-plugin-datalabels';

ChartJS.register(
    CategoryScale,
    LinearScale,
    BarElement,
    LineController,
    LineElement,
    PointElement,
    Title,
    Tooltip,
    Legend,
    ChartDataLabels
);

interface ComparisonViewProps {
    analysis: CustomerAnalysis;
}

type Granularity = 'year' | 'quarter' | 'month';

export const ComparisonView: React.FC<ComparisonViewProps> = ({ analysis }) => {
    const [granularity, setGranularity] = useState<Granularity>('quarter');
    const [selectedPeriods, setSelectedPeriods] = useState<string[]>([]);

    // Get available buckets based on granularity
    const getBuckets = () => {
        switch (granularity) {
            case 'year': return analysis.yearBuckets || {};
            case 'quarter': return analysis.quarterBuckets || {};
            case 'month': return analysis.monthBuckets || {};
            default: return {};
        }
    };

    const buckets = getBuckets();
    const allPeriods = Object.keys(buckets).sort().reverse(); // Newest first

    // Auto-select top 3 periods when granularity changes
    useEffect(() => {
        setSelectedPeriods(allPeriods.slice(0, 3));
    }, [granularity]);

    const togglePeriod = (period: string) => {
        if (selectedPeriods.includes(period)) {
            setSelectedPeriods(selectedPeriods.filter(p => p !== period));
        } else {
            if (selectedPeriods.length < 5) {
                setSelectedPeriods([...selectedPeriods, period].sort().reverse());
            } else {
                alert('最多只能選擇 5 個區間進行比較');
            }
        }
    };

    // Prepare Comparison Data
    const constDisplayPeriods = [...selectedPeriods].sort();
    const comparisonData = constDisplayPeriods.map(period => buckets[period]).filter(Boolean);

    // --- Helper: Group Periods by Year for Better UX ---
    const getGroupedPeriods = () => {
        if (granularity === 'year') return { 'Year': allPeriods };

        const groups: { [year: string]: string[] } = {};
        allPeriods.forEach(p => {
            const year = p.split('-')[0];
            if (!groups[year]) groups[year] = [];
            groups[year].push(p);
        });
        return groups;
    };

    const periodGroups = getGroupedPeriods();

    // --- Chart Data Preparation Helpers ---

    const getGroupedBarData = (
        title: string,
        dataMapExtractor: (stat: PeriodStats) => { [key: string]: number },
        topN: number = 5
    ) => {
        // 1. Collect all unique keys
        const allKeysSet = new Set<string>();
        comparisonData.forEach(d => {
            Object.keys(dataMapExtractor(d) || {}).forEach(k => allKeysSet.add(k));
        });
        let allKeys = Array.from(allKeysSet);

        // 2. Sort Keys
        allKeys.sort();

        // If too many keys, take top N
        if (allKeys.length > topN) {
            const keyCounts: { [key: string]: number } = {};
            allKeys.forEach(key => {
                keyCounts[key] = comparisonData.reduce((sum, d) => sum + (dataMapExtractor(d)[key] || 0), 0);
            });
            allKeys = allKeys.sort((a, b) => keyCounts[b] - keyCounts[a]).slice(0, topN);
        }

        // 3. Build Datasets
        const datasets = comparisonData.map((periodData, index) => {
            const color = [
                'rgba(59, 130, 246, 0.7)', // Blue
                'rgba(16, 185, 129, 0.7)', // Green
                'rgba(245, 158, 11, 0.7)', // Amber
                'rgba(239, 68, 68, 0.7)',  // Red
                'rgba(139, 92, 246, 0.7)'  // Purple
            ][index % 5];

            return {
                label: periodData.period,
                data: allKeys.map(key => dataMapExtractor(periodData)[key] || 0),
                backgroundColor: color,
            };
        });

        return {
            labels: allKeys,
            datasets
        };
    };

    // --- Main Trend Chart Data ---
    const trendChartData = {
        labels: constDisplayPeriods,
        datasets: [
            {
                label: '營業額 (Revenue)',
                data: comparisonData.map(d => d.totalAmount),
                backgroundColor: 'rgba(245, 158, 11, 0.7)', // Amber
                yAxisID: 'y',
                order: 2
            },
            {
                label: '成交數 (Deals)',
                data: comparisonData.map(d => d.completedDeals),
                backgroundColor: 'rgba(16, 185, 129, 0.7)', // Emerald
                yAxisID: 'y1',
                type: 'line' as const,
                borderColor: '#10b981',
                borderWidth: 2,
                pointBackgroundColor: '#fff',
                order: 1
            }
        ]
    };

    const trendChartOptions = {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
            legend: { position: 'bottom' as const },
            datalabels: {
                display: true,
                align: 'top' as const,
                formatter: (value: number) => value > 0 ? value.toLocaleString() : ''
            }
        },
        scales: {
            y: {
                type: 'linear' as const,
                display: true,
                position: 'left' as const,
                title: { display: true, text: '營業額 ($)' }
            },
            y1: {
                type: 'linear' as const,
                display: true,
                position: 'right' as const,
                grid: { drawOnChartArea: false },
                title: { display: true, text: '成交數 (筆)' }
            }
        }
    };

    // --- Sales Performance Data Helper (By Revenue) ---
    const getSalesPerformanceData = () => {
        // 1. Collect all unique salespeople and calculate overall stats for sorting
        const overallStats: { [name: string]: number } = {};

        comparisonData.forEach(periodData => {
            const perf = periodData.salespersonPerformance || {};
            Object.keys(perf).forEach(name => {
                overallStats[name] = (overallStats[name] || 0) + (perf[name].revenue || 0);
            });
        });

        // 2. Sort by Total Revenue
        const sortedSalespeople = Object.keys(overallStats)
            .map(name => ({ name, revenue: overallStats[name] }))
            .sort((a, b) => b.revenue - a.revenue) // Descending
            .slice(0, 10); // Top 10

        const topNames = sortedSalespeople.map(s => s.name);

        // 3. Build Datasets
        const datasets = comparisonData.map((periodData, index) => {
            const color = [
                'rgba(59, 130, 246, 0.7)', // Blue
                'rgba(16, 185, 129, 0.7)', // Green
                'rgba(245, 158, 11, 0.7)', // Amber
                'rgba(239, 68, 68, 0.7)',  // Red
                'rgba(139, 92, 246, 0.7)'  // Purple
            ][index % 5];

            return {
                label: periodData.period,
                data: topNames.map(name => {
                    const perf = periodData.salespersonPerformance?.[name];
                    return perf ? perf.revenue || 0 : 0;
                }),
                backgroundColor: color,
            };
        });

        return {
            labels: topNames,
            datasets
        };
    };

    // --- Detail Charts ---
    const ageChartData = getGroupedBarData('年齡分佈', d => d.ageDistribution);
    const sourceChartData = getGroupedBarData('客源分佈', d => d.sourceDistribution);
    const hearingChartData = getGroupedBarData('聽損程度', d => d.hearingLossDistribution);
    const salesChartData = getSalesPerformanceData(); // Use new revenue helper

    const groupedChartOptions = {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
            legend: { position: 'bottom' as const },
            datalabels: {
                display: true,
                color: '#1e293b', // slate-800 for better visibility
                font: { weight: 'bold' as const },
                formatter: (v: number) => v > 0 ? v : '',
                anchor: 'end' as const,
                align: 'start' as const,
                offset: -4
            }
        },
        scales: {
            y: { beginAtZero: true }
        }
    };

    const salesChartOptions = {
        ...groupedChartOptions,
        scales: {
            y: {
                beginAtZero: true,
                title: { display: true, text: '總營收 ($)' }
            }
        },
        plugins: {
            ...groupedChartOptions.plugins,
            datalabels: {
                display: false, // Hide labels to prevent overlap
            },
            tooltip: {
                callbacks: {
                    // @ts-ignore
                    label: (context) => `${context.dataset.label}: $${(context.raw as number).toLocaleString()}`
                }
            }
        }
    };

    return (
        <div className="space-y-6 animate-fade-in pb-12">
            {/* Control Panel */}
            <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-4 sticky top-0 z-30 opacity-95 backdrop-blur-sm transition-all">
                <div className="flex flex-col sm:flex-row justify-between items-center mb-3 gap-3">
                    <div className="flex items-center gap-2">
                        <h2 className="text-lg font-bold text-slate-800">區間比較分析</h2>
                        <span className="text-xs text-slate-500 hidden sm:inline-block">| 多維度成效對比 (最多5個)</span>
                    </div>

                    {/* Granularity Switcher */}
                    <div className="flex bg-slate-100 p-0.5 rounded-lg scale-90 sm:scale-100 origin-right">
                        {(['year', 'quarter', 'month'] as Granularity[]).map((g) => (
                            <button
                                key={g}
                                onClick={() => setGranularity(g)}
                                className={`px-3 py-1.5 text-xs font-medium rounded-md transition-all ${granularity === g
                                    ? 'bg-white text-blue-600 shadow-sm'
                                    : 'text-slate-500 hover:text-slate-700'
                                    }`}
                            >
                                {g === 'year' ? '年度' : g === 'quarter' ? '季度' : '月度'}
                            </button>
                        ))}
                    </div>
                </div>

                {/* Period Selector (Grouped by Year) */}
                <div className="space-y-1.5 max-h-40 overflow-y-auto pr-1 custom-scrollbar">
                    {Object.keys(periodGroups).sort().reverse().map(year => (
                        <div key={year} className="flex flex-row items-center gap-2 border-b border-slate-50 pb-1.5 last:border-0 last:pb-0">
                            {granularity !== 'year' && (
                                <div className="w-12 text-xs font-bold text-slate-400 shrink-0 text-right">{year}</div>
                            )}
                            <div className="flex flex-wrap gap-1.5">
                                {periodGroups[year].map(period => (
                                    <button
                                        key={period}
                                        onClick={() => togglePeriod(period)}
                                        className={`px-2.5 py-1 text-[10px] sm:text-xs rounded-full border transition-all ${selectedPeriods.includes(period)
                                            ? 'bg-blue-50 border-blue-200 text-blue-700 font-medium shadow-sm ring-1 ring-blue-100'
                                            : 'bg-white border-slate-200 text-slate-600 hover:border-slate-300 hover:bg-slate-50'
                                            }`}
                                    >
                                        {granularity === 'year' ? period : period.split('-')[1]}
                                    </button>
                                ))}
                            </div>
                        </div>
                    ))}
                </div>
            </div>

            {/* 1. Main Trend Section */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                {/* Trend Chart */}
                <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-6">
                    <h3 className="text-lg font-bold text-slate-800 mb-4 border-l-4 border-amber-500 pl-3">營收與成交趨勢</h3>
                    <div className="h-80">
                        {/* @ts-ignore */}
                        <Bar data={trendChartData} options={trendChartOptions} />
                    </div>
                </div>

                {/* Detailed Table */}
                <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-6 overflow-hidden">
                    <h3 className="text-lg font-bold text-slate-800 mb-4 border-l-4 border-blue-500 pl-3">核心指標總表</h3>
                    <div className="overflow-x-auto h-80">
                        <table className="min-w-full divide-y divide-slate-100">
                            <thead className="bg-slate-50 sticky top-0 z-10">
                                <tr>
                                    <th className="px-4 py-3 text-left text-xs font-semibold text-slate-500 uppercase">區間</th>
                                    <th className="px-4 py-3 text-right text-xs font-semibold text-slate-500 uppercase">來客</th>
                                    <th className="px-4 py-3 text-right text-xs font-semibold text-slate-500 uppercase">成交</th>
                                    <th className="px-4 py-3 text-right text-xs font-semibold text-slate-500 uppercase">轉化率</th>
                                    <th className="px-4 py-3 text-right text-xs font-semibold text-slate-500 uppercase">客單</th>
                                    <th className="px-4 py-3 text-right text-xs font-semibold text-slate-500 uppercase">總營收</th>
                                </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-100">
                                {comparisonData.map((data) => (
                                    <tr key={data.period} className="hover:bg-slate-50 transition-colors">
                                        <td className="px-4 py-3 text-sm font-medium text-slate-900">{data.period}</td>
                                        <td className="px-4 py-3 text-sm text-right text-slate-600">{data.newCustomers}</td>
                                        <td className="px-4 py-3 text-sm text-right text-slate-600">{data.completedDeals}</td>
                                        <td className="px-4 py-3 text-sm text-right font-semibold text-blue-600">
                                            {data.conversionRate.toFixed(1)}%
                                        </td>
                                        <td className="px-4 py-3 text-sm text-right text-slate-600 font-mono">
                                            ${data.averageOrderValue.toLocaleString()}
                                        </td>
                                        <td className="px-4 py-3 text-sm text-right font-bold text-amber-600 font-mono">
                                            ${data.totalAmount.toLocaleString()}
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>

            {/* 2. Detailed Breakdown Section */}
            <h2 className="text-2xl font-bold text-slate-800 mt-2 mb-4 px-2">深度結構分析</h2>
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                {/* Age Distribution */}
                <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-6">
                    <h3 className="text-lg font-bold text-slate-800 mb-4 border-l-4 border-indigo-500 pl-3">年齡層分佈</h3>
                    <div className="h-64">
                        {/* @ts-ignore */}
                        <Bar data={ageChartData} options={groupedChartOptions} />
                    </div>
                </div>

                {/* Source Distribution */}
                <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-6">
                    <h3 className="text-lg font-bold text-slate-800 mb-4 border-l-4 border-emerald-500 pl-3">主要客源 (Top 5)</h3>
                    <div className="h-64">
                        {/* @ts-ignore */}
                        <Bar data={sourceChartData} options={groupedChartOptions} />
                    </div>
                </div>

                {/* Hearing Loss Distribution */}
                <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-6">
                    <h3 className="text-lg font-bold text-slate-800 mb-4 border-l-4 border-rose-500 pl-3">聽損程度分佈</h3>
                    <div className="h-64">
                        {/* @ts-ignore */}
                        <Bar data={hearingChartData} options={groupedChartOptions} />
                    </div>
                </div>

                {/* Salesperson Performance Distribution (New) */}
                <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-6">
                    <h3 className="text-lg font-bold text-slate-800 mb-4 border-l-4 border-cyan-500 pl-3">業務員業績 (營收 Top 10)</h3>
                    <div className="h-64">
                        {/* @ts-ignore */}
                        <Bar data={salesChartData} options={salesChartOptions} />
                    </div>
                </div>
            </div>
        </div>
    );
};
