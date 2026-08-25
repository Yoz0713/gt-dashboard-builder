import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
    Chart as ChartJS,
    CategoryScale,
    LinearScale,
    BarElement,
    Title,
    Tooltip,
    Legend,
    ArcElement,
} from 'chart.js';
import { Bar, Doughnut } from 'react-chartjs-2';
import ChartDataLabels from 'chartjs-plugin-datalabels';
import { ClinicFollowUpReport, DateRange } from '../types';
import { buildClinicTalkingPoints } from '../utils/analysis';

ChartJS.register(CategoryScale, LinearScale, BarElement, Title, Tooltip, Legend, ArcElement, ChartDataLabels);

const COLORS = {
    primary: '#0ea5e9', // Sky 500
    secondary: '#10b981', // Emerald 500
    accent: '#f59e0b', // Amber 500
    purple: '#8b5cf6', // Violet 500
    slate: '#64748b', // Slate 500
    red: '#ef4444', // Red 500
};

const commonOptions = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
        legend: { position: 'bottom' as const, labels: { usePointStyle: true, padding: 20 } },
        datalabels: {
            color: '#334155',
            font: { weight: 'bold' as const, size: 12 },
            anchor: 'end' as const,
            align: 'top' as const,
            offset: 4,
            formatter: (value: number) => (value > 0 ? value : ''),
        },
    },
    scales: {
        y: { beginAtZero: true, grace: '10%' },
    },
};

const doughnutOptions = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
        legend: { position: 'bottom' as const, labels: { usePointStyle: true, padding: 16 } },
        datalabels: {
            color: '#ffffff',
            font: { weight: 'bold' as const, size: 12 },
            formatter: (value: number, context: any) => {
                const total = (context.dataset.data as number[]).reduce((sum, item) => sum + item, 0);
                if (!total || !value) return '';
                return `${Math.round((value / total) * 100)}%`;
            },
        },
    },
};

const DEGREE_COLORS: { [degree: string]: string } = {
    正常: '#94a3b8',
    輕度: '#38bdf8',
    中度: '#f59e0b',
    中重度: '#fb923c',
    重度: '#ef4444',
    極重度: '#b91c1c',
    未知: '#cbd5e1',
};

const SORT_OPTIONS: { key: ClinicSortKey; label: string }[] = [
    { key: 'referrals', label: '轉介人次' },
    { key: 'recent', label: '最近轉介' },
    { key: 'conversion', label: '配戴率' },
];

type ClinicSortKey = 'referrals' | 'recent' | 'conversion';

const STALE_DAYS = 60;

/** 大樹聽力中心報告表頭主色 */
const BRAND_BLUE = 'rgb(0, 140, 215)';

const REPORT_TITLE = '大樹聽力中心診所轉介名單分析';

/** 轉介費用預設以成交金額的 10% 回饋給轉介診所。 */
export const DEFAULT_REFERRAL_FEE_RATE = 10;

export const calculateReferralFee = (amount: number, ratePercent: number): number => {
    if (!amount || amount <= 0 || !ratePercent || ratePercent <= 0) return 0;
    return Math.round((amount * ratePercent) / 100);
};

const formatCurrency = (amount: number) => `NT$${Math.round(amount).toLocaleString('en-US')}`;

/** 由 'YYYY-MM-DD' 算出距今天數；空字串或無法解析時回傳 null。 */
export const daysSinceDate = (isoDate: string): number | null => {
    if (!isoDate) return null;

    const [year, month, day] = isoDate.split('-').map((part) => parseInt(part, 10));
    if (!year || !month || !day) return null;

    const target = new Date(year, month - 1, day);
    if (isNaN(target.getTime())) return null;

    const now = new Date();
    const startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate());

    return Math.max(0, Math.round((startOfToday.getTime() - target.getTime()) / 86400000));
};

/** 王大明 -> 王○明；兩個字時只遮後面一個字。 */
export const maskCustomerName = (name: string): string => {
    if (name.length <= 1) return name;
    if (name.length === 2) return `${name[0]}○`;
    return `${name[0]}${'○'.repeat(name.length - 2)}${name[name.length - 1]}`;
};

/**
 * 試算表名稱通常是「門市 + 工作表名」（例：湖口店來客紀錄），
 * 報告表頭只要門市本身，因此把固定的工作表名與前後分隔符去掉。
 */
export const formatStoreName = (spreadsheetTitle: string): string =>
    spreadsheetTitle
        .replace(/來客紀錄/g, '')
        .replace(/^[\s\-_–—|｜]+|[\s\-_–—|｜]+$/g, '')
        .trim();

const formatDateRange = (dateRange: DateRange) =>
    `${dateRange.startYear}年${dateRange.startMonth}月 ~ ${dateRange.endYear}年${dateRange.endMonth}月`;

interface ClinicFollowUpProps {
    reports: ClinicFollowUpReport[];
    dateRange: DateRange;
    spreadsheetTitle?: string;
    ptaThreshold: number;
}

export const ClinicFollowUp: React.FC<ClinicFollowUpProps> = ({
    reports,
    dateRange,
    spreadsheetTitle = '',
    ptaThreshold,
}) => {
    const [selectedClinic, setSelectedClinic] = useState<string>('');
    const [search, setSearch] = useState<string>('');
    const [sortKey, setSortKey] = useState<ClinicSortKey>('referrals');
    const [maskNames, setMaskNames] = useState<boolean>(false);
    const [referralFeeRate, setReferralFeeRate] = useState<number>(DEFAULT_REFERRAL_FEE_RATE);

    const monthlyChartRef = useRef<any>(null);
    const degreeChartRef = useRef<any>(null);
    const ageChartRef = useRef<any>(null);

    // 列印時容器寬度會改變，但 Chart.js 不會自己重算 canvas 尺寸，需要主動 resize，
    // 否則圖表會以螢幕寬度輸出而被頁面邊界裁切。index.css 另有等比縮放的 CSS 安全網。
    useEffect(() => {
        const resizeCharts = () => {
            [monthlyChartRef, degreeChartRef, ageChartRef].forEach((ref) => {
                ref.current?.resize();
            });
        };

        const printMediaQuery = typeof window.matchMedia === 'function' ? window.matchMedia('print') : null;
        const handleMediaChange = (event: MediaQueryListEvent) => {
            if (event.matches) resizeCharts();
        };

        window.addEventListener('beforeprint', resizeCharts);
        window.addEventListener('afterprint', resizeCharts);
        printMediaQuery?.addEventListener?.('change', handleMediaChange);

        return () => {
            window.removeEventListener('beforeprint', resizeCharts);
            window.removeEventListener('afterprint', resizeCharts);
            printMediaQuery?.removeEventListener?.('change', handleMediaChange);
        };
    }, []);

    // 診所清單變動（換試算表 / 換區間）時，自動落回轉介人次最高的診所
    useEffect(() => {
        if (!reports.length) {
            setSelectedClinic('');
            return;
        }

        setSelectedClinic((previous) =>
            reports.some((report) => report.clinic === previous) ? previous : reports[0].clinic
        );
    }, [reports]);

    const visibleClinics = useMemo(() => {
        const keyword = search.trim().toLowerCase();
        const filtered = keyword
            ? reports.filter((report) => report.clinic.toLowerCase().includes(keyword))
            : [...reports];

        return filtered.sort((a, b) => {
            if (sortKey === 'recent') {
                return b.lastReferralDate.localeCompare(a.lastReferralDate);
            }
            if (sortKey === 'conversion') {
                return b.conversionRate - a.conversionRate;
            }
            return b.totalReferrals - a.totalReferrals;
        });
    }, [reports, search, sortKey]);

    const selectedReport = useMemo(
        () => reports.find((report) => report.clinic === selectedClinic) || reports[0] || null,
        [reports, selectedClinic]
    );

    if (!reports.length) {
        return (
            <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-8 text-center">
                <p className="text-slate-600 font-medium">此分析區間內查無「診所名稱」欄位資料。</p>
                <p className="mt-2 text-sm text-slate-500">
                    請確認來源工作表包含「診所名稱」欄位，且分析區間內有診所轉介紀錄。
                </p>
            </div>
        );
    }

    if (!selectedReport) {
        return null;
    }

    const daysSinceLast = daysSinceDate(selectedReport.lastReferralDate);
    const isStale = daysSinceLast !== null && daysSinceLast > STALE_DAYS;
    const talkingPoints = buildClinicTalkingPoints(selectedReport, ptaThreshold, daysSinceLast);

    const dealtPatients = selectedReport.patients.filter((patient) => patient.isDealt);
    const totalReferralFee = dealtPatients.reduce(
        (sum, patient) => sum + calculateReferralFee(patient.amount, referralFeeRate),
        0
    );
    const missingAmountCount = dealtPatients.filter((patient) => patient.amount <= 0).length;
    const dealtAmountTotal = dealtPatients.reduce((sum, patient) => sum + patient.amount, 0);
    const dealtAverageAmount = dealtPatients.length > 0 ? Math.round(dealtAmountTotal / dealtPatients.length) : 0;
    const hasReferralFee = totalReferralFee > 0;
    const storeName = formatStoreName(spreadsheetTitle);

    const monthlyChartData = {
        labels: selectedReport.monthly.map((point) => point.label),
        datasets: [
            {
                label: '轉介人次',
                data: selectedReport.monthly.map((point) => point.referrals),
                backgroundColor: 'rgba(14, 165, 233, 0.7)',
                borderRadius: 4,
            },
            {
                label: '聽損個案',
                data: selectedReport.monthly.map((point) => point.hearingLoss),
                backgroundColor: 'rgba(245, 158, 11, 0.7)',
                borderRadius: 4,
            },
            {
                label: '完成配戴',
                data: selectedReport.monthly.map((point) => point.deals),
                backgroundColor: 'rgba(16, 185, 129, 0.7)',
                borderRadius: 4,
            },
        ],
    };

    const degreeChartData = {
        labels: selectedReport.hearingDegreeDistribution.map((item) => item.degree),
        datasets: [
            {
                data: selectedReport.hearingDegreeDistribution.map((item) => item.count),
                backgroundColor: selectedReport.hearingDegreeDistribution.map(
                    (item) => DEGREE_COLORS[item.degree] || COLORS.slate
                ),
                borderWidth: 2,
                borderColor: '#ffffff',
            },
        ],
    };

    const ageChartData = {
        labels: selectedReport.ageDistribution.map((item) => item.range),
        datasets: [
            {
                label: '人數',
                data: selectedReport.ageDistribution.map((item) => item.count),
                backgroundColor: 'rgba(139, 92, 246, 0.6)',
                borderColor: COLORS.purple,
                borderWidth: 1,
                borderRadius: 4,
            },
        ],
    };

    return (
        <div className="space-y-6">
            {/* 診所選取（列印時隱藏） */}
            <section className="rounded-2xl border border-teal-200 bg-gradient-to-br from-teal-50 via-white to-sky-50 p-5 shadow-sm print:hidden">
                <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
                    <div>
                        <p className="text-xs font-semibold uppercase tracking-[0.2em] text-teal-600">Clinic Follow-up</p>
                        <h2 className="mt-2 text-xl font-bold text-slate-900">診所回訪對象</h2>
                        <p className="mt-2 max-w-2xl text-sm text-slate-600">
                            以下為 {formatDateRange(dateRange)} 內有轉介紀錄的診所，共 {reports.length} 家。選定一家後即可產出可列印的回訪報告。
                        </p>
                    </div>
                    <div className="rounded-full border border-teal-200 bg-white px-4 py-2 text-sm font-semibold text-teal-700">
                        目前選定：{selectedReport.clinic}
                    </div>
                </div>

                <div className="mt-4 flex flex-col gap-3 md:flex-row md:items-center">
                    <input
                        type="text"
                        value={search}
                        onChange={(event) => setSearch(event.target.value)}
                        placeholder="搜尋診所名稱..."
                        aria-label="搜尋診所名稱"
                        className="flex-1 rounded-lg border border-slate-200 bg-white p-2.5 text-sm focus:border-teal-500 focus:ring-teal-500"
                    />
                    <div className="flex items-center gap-1 rounded-lg border border-slate-200 bg-white p-1">
                        {SORT_OPTIONS.map((option) => (
                            <button
                                key={option.key}
                                onClick={() => setSortKey(option.key)}
                                className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-colors ${sortKey === option.key
                                    ? 'bg-teal-50 text-teal-700 ring-1 ring-teal-200'
                                    : 'text-slate-500 hover:text-slate-700 hover:bg-slate-50'
                                    }`}
                            >
                                {option.label}
                            </button>
                        ))}
                    </div>
                </div>

                {visibleClinics.length === 0 ? (
                    <p className="mt-4 text-sm text-slate-500">找不到符合「{search}」的診所。</p>
                ) : (
                    <div className="mt-4 grid gap-3 md:grid-cols-2 xl:grid-cols-3">
                        {visibleClinics.map((report) => {
                            const isSelected = report.clinic === selectedReport.clinic;
                            const clinicDaysSince = daysSinceDate(report.lastReferralDate);
                            const clinicIsStale = clinicDaysSince !== null && clinicDaysSince > STALE_DAYS;

                            return (
                                <button
                                    key={report.clinic}
                                    onClick={() => setSelectedClinic(report.clinic)}
                                    aria-pressed={isSelected}
                                    className={`text-left rounded-xl border px-4 py-3 transition-all ${isSelected
                                        ? 'border-teal-300 bg-white shadow-sm ring-2 ring-teal-100'
                                        : 'border-slate-200 bg-white/80 hover:border-teal-200 hover:bg-white'
                                        }`}
                                >
                                    <span className="flex items-start justify-between gap-2">
                                        <span className="min-w-0 flex-1 truncate text-sm font-semibold text-slate-800">
                                            {report.clinic}
                                        </span>
                                        {clinicDaysSince !== null && (
                                            <span
                                                className={`shrink-0 rounded-full px-2 py-0.5 text-[11px] font-semibold ${clinicIsStale
                                                    ? 'bg-amber-50 text-amber-700 border border-amber-200'
                                                    : 'bg-slate-100 text-slate-500'
                                                    }`}
                                            >
                                                {clinicDaysSince} 天前
                                            </span>
                                        )}
                                    </span>
                                    <span className="mt-2 flex items-center gap-3 text-xs text-slate-500">
                                        <span>轉介 <b className="text-slate-800">{report.totalReferrals}</b></span>
                                        <span>聽損 <b className="text-amber-600">{report.hearingLossCount}</b></span>
                                        <span>配戴 <b className="text-emerald-600">{report.dealtCount}</b></span>
                                    </span>
                                </button>
                            );
                        })}
                    </div>
                )}
            </section>

            {/* 列印用抬頭：App 的頁首與 tab 列在列印時皆隱藏，報告需自行帶出識別資訊 */}
            <div
                className="report-brand-header hidden print:block px-5 py-4 mb-4 text-white"
                style={{ backgroundColor: BRAND_BLUE }}
            >
                <h1 className="text-2xl font-bold">{REPORT_TITLE}</h1>
                <p className="mt-1 text-sm">
                    診所名稱：{selectedReport.clinic}
                    {storeName ? ` ｜ 服務門市：${storeName}` : ''}
                    {` ｜ 產出日期：${new Date().toLocaleDateString('zh-TW')}`}
                </p>
            </div>

            <div className="flex items-center justify-between print:hidden">
                <h2 className="text-lg font-bold text-slate-800 border-l-4 border-teal-500 pl-3">
                    {selectedReport.clinic} 回訪報告
                </h2>
                <button
                    onClick={() => window.print()}
                    className="bg-teal-600 text-white px-5 py-2.5 rounded-lg hover:bg-teal-700 transition-colors flex items-center gap-2 shadow-md hover:shadow-lg active:scale-95 text-sm font-medium"
                >
                    🖨️ 匯出此診所回訪報告 (PDF)
                </button>
            </div>

            {/* KPI */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 print:grid-cols-4">
                <StatTile
                    label="轉介總人次"
                    value={`${selectedReport.totalReferrals}`}
                    hint={`橫跨 ${selectedReport.activeMonths} 個月`}
                    accent="text-sky-600"
                />
                <StatTile
                    label={`聽損個案 (PTA > ${ptaThreshold})`}
                    value={`${selectedReport.hearingLossCount}`}
                    hint={`占 ${selectedReport.hearingLossRate.toFixed(0)}%`}
                    accent="text-amber-600"
                />
                <StatTile
                    label="完成配戴"
                    value={`${selectedReport.dealtCount}`}
                    hint={`配戴率 ${selectedReport.conversionRate.toFixed(0)}%`}
                    accent="text-emerald-600"
                />
                <StatTile
                    label="最近轉介"
                    value={daysSinceLast === null ? '—' : `${daysSinceLast} 天前`}
                    hint={selectedReport.lastReferralDate || '無可解析日期'}
                    accent={isStale ? 'text-red-600' : 'text-slate-800'}
                />
            </div>

            {/* 回訪重點：面談前的內部提示，不列入交給診所的報告 */}
            <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-6 print:hidden">
                <h3 className="text-lg font-bold text-slate-800 mb-4 border-l-4 border-teal-500 pl-3">回訪重點</h3>
                <ul className="space-y-2 text-sm text-slate-700 list-disc pl-5">
                    {talkingPoints.map((point) => (
                        <li key={point}>{point}</li>
                    ))}
                </ul>
            </div>

            {/* 轉介費用：沒有需要給付的費用時不列入匯出報告 */}
            <div
                data-testid="referral-fee-block"
                className={`bg-white rounded-xl shadow-sm border border-slate-100 p-6 ${hasReferralFee ? '' : 'print:hidden'}`}
            >
                <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between mb-4">
                    <div>
                        <h3 className="text-lg font-bold text-slate-800 border-l-4 border-emerald-500 pl-3">轉介費用</h3>
                        <p className="mt-2 text-sm text-slate-500">
                            依已完成配戴個案的成交金額，以約定比例計算應回饋給 {selectedReport.clinic} 的轉介費用。
                        </p>
                    </div>
                    <div className="flex items-center gap-2 text-sm text-slate-600 shrink-0">
                        <label htmlFor="referral-fee-rate" className="font-medium">回饋比例</label>
                        <input
                            id="referral-fee-rate"
                            type="number"
                            min={0}
                            max={100}
                            step={0.5}
                            value={referralFeeRate}
                            onChange={(event) => {
                                const parsed = parseFloat(event.target.value);
                                setReferralFeeRate(isNaN(parsed) ? 0 : Math.min(100, Math.max(0, parsed)));
                            }}
                            className="w-20 rounded-lg border border-slate-200 p-2 text-right text-sm print:hidden"
                        />
                        <span className="hidden print:inline font-semibold text-slate-800">{referralFeeRate}</span>
                        <span>%</span>
                    </div>
                </div>

                {dealtPatients.length === 0 ? (
                    <p className="text-sm text-slate-500">
                        本期此診所尚無已配戴個案，因此沒有可計算的轉介費用。
                    </p>
                ) : (
                    <>
                        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 print:hidden">
                            <StatTile
                                label="已配戴人數"
                                value={`${dealtPatients.length}`}
                                hint={`共 ${selectedReport.totalReferrals} 位轉介`}
                                accent="text-emerald-600"
                            />
                            <StatTile
                                label="成交金額合計"
                                value={formatCurrency(dealtAmountTotal)}
                                hint={`平均 ${formatCurrency(dealtAverageAmount)}／位`}
                                accent="text-slate-800"
                            />
                            <StatTile
                                label="回饋比例"
                                value={`${referralFeeRate}%`}
                                hint="依成交金額計算"
                                accent="text-slate-800"
                            />
                            <StatTile
                                label="轉介費用合計"
                                value={formatCurrency(totalReferralFee)}
                                hint={`來自 ${dealtPatients.length} 筆已配戴個案`}
                                accent="text-emerald-600"
                            />
                        </div>

                        {missingAmountCount > 0 && (
                            <p className="mt-4 rounded-lg border border-amber-200 bg-amber-50 p-3 text-sm text-amber-700">
                                有 {missingAmountCount} 筆已配戴個案未填寫成交金額，該筆轉介費用以 NT$0 計算，請先於來客紀錄補齊金額。
                            </p>
                        )}

                        <div className="mt-4 overflow-x-auto">
                            <table className="clinic-report-table min-w-full divide-y divide-slate-100">
                                <thead style={{ backgroundColor: BRAND_BLUE }}>
                                    <tr>
                                        <th className="px-4 py-3 text-left text-xs font-semibold text-white uppercase">服務日期</th>
                                        <th className="px-4 py-3 text-left text-xs font-semibold text-white uppercase">姓名</th>
                                        <th className="px-4 py-3 text-right text-xs font-semibold text-white uppercase">成交金額</th>
                                        <th className="px-4 py-3 text-right text-xs font-semibold text-white uppercase">
                                            轉介費用 ({referralFeeRate}%)
                                        </th>
                                    </tr>
                                </thead>
                                <tbody className="bg-white divide-y divide-slate-100 text-sm">
                                    {dealtPatients.map((patient, index) => (
                                        <tr key={`fee-${patient.serviceDate}-${patient.name}-${index}`} className="hover:bg-slate-50 transition-colors">
                                            <td className="px-4 py-3 text-slate-700 whitespace-nowrap">{patient.serviceDate || '—'}</td>
                                            <td className="px-4 py-3 text-slate-800 font-medium">
                                                {maskNames ? maskCustomerName(patient.name) : patient.name}
                                            </td>
                                            <td className="px-4 py-3 text-right text-slate-700">{formatCurrency(patient.amount)}</td>
                                            <td className="px-4 py-3 text-right font-semibold text-emerald-600">
                                                {formatCurrency(calculateReferralFee(patient.amount, referralFeeRate))}
                                            </td>
                                        </tr>
                                    ))}
                                </tbody>
                                <tfoot className="bg-slate-50">
                                    <tr>
                                        <td className="px-4 py-3 text-sm font-semibold text-slate-700" colSpan={2}>合計</td>
                                        <td className="px-4 py-3 text-right text-sm font-semibold text-slate-800">
                                            <span data-testid="fee-total-amount" className="print:hidden">
                                                {formatCurrency(dealtAmountTotal)}
                                            </span>
                                        </td>
                                        <td className="px-4 py-3 text-right text-sm font-bold text-emerald-700">
                                            {formatCurrency(totalReferralFee)}
                                        </td>
                                    </tr>
                                </tfoot>
                            </table>
                        </div>
                    </>
                )}
            </div>

            {/* 月度趨勢 */}
            <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-6">
                <h3 className="text-lg font-bold text-slate-800 mb-4 border-l-4 border-sky-500 pl-3">月度轉介趨勢</h3>
                {selectedReport.monthly.length === 0 ? (
                    <p className="text-sm text-slate-500">此診所的轉介資料沒有可解析的服務日期，無法產生月度趨勢。</p>
                ) : (
                    <div className="h-64">
                        <Bar ref={monthlyChartRef} data={monthlyChartData} options={commonOptions} />
                    </div>
                )}
            </div>

            {/* 分布 */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 print:grid-cols-2">
                <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-6">
                    <h3 className="text-lg font-bold text-slate-800 mb-4 border-l-4 border-amber-500 pl-3">聽損程度分布</h3>
                    <div className="h-64">
                        <Doughnut ref={degreeChartRef} data={degreeChartData} options={doughnutOptions} />
                    </div>
                </div>
                <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-6">
                    <h3 className="text-lg font-bold text-slate-800 mb-4 border-l-4 border-violet-500 pl-3">年齡分布</h3>
                    <div className="h-64">
                        <Bar ref={ageChartRef} data={ageChartData} options={commonOptions} />
                    </div>
                </div>
            </div>

            {/* 客戶明細 */}
            <div className="clinic-report-detail bg-white rounded-xl shadow-sm border border-slate-100 overflow-hidden print:break-before-page">
                <div className="p-6 border-b border-slate-100 flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
                    <div>
                        <h3 className="text-lg font-bold text-slate-800 border-l-4 border-slate-500 pl-3">轉介客戶明細</h3>
                        <p className="mt-2 text-sm text-slate-500">
                            共 {selectedReport.patients.length} 筆，依服務日期由近至遠排列。
                        </p>
                    </div>
                    <label className="flex items-center gap-2 text-sm text-slate-600 print:hidden">
                        <input
                            type="checkbox"
                            checked={maskNames}
                            onChange={(event) => setMaskNames(event.target.checked)}
                            className="h-4 w-4 rounded border-slate-300 text-teal-600 focus:ring-teal-500"
                        />
                        姓名遮罩
                    </label>
                </div>

                <div className="overflow-x-auto">
                    <table className="clinic-report-table min-w-full divide-y divide-slate-100">
                        <thead style={{ backgroundColor: BRAND_BLUE }}>
                            <tr>
                                <th className="px-4 py-3 text-left text-xs font-semibold text-white uppercase">服務日期</th>
                                <th className="px-4 py-3 text-left text-xs font-semibold text-white uppercase">姓名</th>
                                <th className="px-4 py-3 text-right text-xs font-semibold text-white uppercase">年齡</th>
                                <th className="px-4 py-3 text-right text-xs font-semibold text-white uppercase">左耳 PTA</th>
                                <th className="px-4 py-3 text-right text-xs font-semibold text-white uppercase">右耳 PTA</th>
                                <th className="px-4 py-3 text-left text-xs font-semibold text-white uppercase">聽損程度</th>
                                <th className="px-4 py-3 text-center text-xs font-semibold text-white uppercase">是否配戴</th>
                                <th className="px-4 py-3 text-right text-xs font-semibold text-white uppercase">成交金額</th>
                                <th className="px-4 py-3 text-right text-xs font-semibold text-white uppercase">轉介費用</th>
                                <th className="px-4 py-3 text-left text-xs font-semibold text-white uppercase">主聽力師</th>
                            </tr>
                        </thead>
                        <tbody className="bg-white divide-y divide-slate-100 text-sm">
                            {selectedReport.patients.map((patient, index) => (
                                <tr key={`${patient.serviceDate}-${patient.name}-${index}`} className="hover:bg-slate-50 transition-colors">
                                    <td className="px-4 py-3 text-slate-700 whitespace-nowrap">{patient.serviceDate || '—'}</td>
                                    <td className="px-4 py-3 text-slate-800 font-medium">
                                        {maskNames ? maskCustomerName(patient.name) : patient.name}
                                    </td>
                                    <td className="px-4 py-3 text-right text-slate-700">{patient.age ?? '—'}</td>
                                    <td className={`px-4 py-3 text-right ${patient.leftPTA !== null && patient.leftPTA > ptaThreshold ? 'text-amber-600 font-semibold' : 'text-slate-700'}`}>
                                        {patient.leftPTA ?? '—'}
                                    </td>
                                    <td className={`px-4 py-3 text-right ${patient.rightPTA !== null && patient.rightPTA > ptaThreshold ? 'text-amber-600 font-semibold' : 'text-slate-700'}`}>
                                        {patient.rightPTA ?? '—'}
                                    </td>
                                    <td className="px-4 py-3 text-slate-700">{patient.hearingDegree}</td>
                                    <td className="px-4 py-3 text-center">
                                        {patient.isDealt ? (
                                            <span className="text-emerald-600 font-semibold">已配戴</span>
                                        ) : (
                                            <span className="text-slate-400">—</span>
                                        )}
                                    </td>
                                    <td className="px-4 py-3 text-right text-slate-700">
                                        {patient.isDealt ? formatCurrency(patient.amount) : '—'}
                                    </td>
                                    <td className="px-4 py-3 text-right text-slate-700">
                                        {patient.isDealt
                                            ? formatCurrency(calculateReferralFee(patient.amount, referralFeeRate))
                                            : '—'}
                                    </td>
                                    <td className="px-4 py-3 text-slate-700">{patient.audiologist}</td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            </div>

            <p className="text-xs text-slate-400 text-center print:text-slate-500">
                本報告依 {formatDateRange(dateRange)} 之來客紀錄產出，聽損個案以 PTA &gt; {ptaThreshold} dB 為判定標準，
                轉介費用以已配戴個案成交金額之 {referralFeeRate}% 計算，實際金額以雙方合約約定為準。
            </p>
        </div>
    );
};

interface StatTileProps {
    label: string;
    value: string;
    hint: string;
    accent: string;
}

const StatTile: React.FC<StatTileProps> = ({ label, value, hint, accent }) => (
    <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-5">
        <p className="text-xs font-semibold text-slate-500 uppercase tracking-wider">{label}</p>
        <p className={`mt-2 text-3xl font-bold ${accent}`}>{value}</p>
        <p className="mt-1 text-xs text-slate-400">{hint}</p>
    </div>
);
