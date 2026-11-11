import React, { useEffect, useMemo, useState } from 'react';
import {
    CompetitionSheetSource,
    DateRange,
    FollowUpCustomerRecord,
    FollowUpSortKey,
    SpreadsheetListItem,
} from '../types';
import {
    buildFollowUpDataset,
    selectFollowUpCustomers,
    summarizeFollowUpCustomers,
    SEVERE_PTA_THRESHOLD,
} from '../utils/analysis';
import { daysSinceTimestamp, maskCustomerName, maskPhone } from '../utils/display';
import { downloadFollowUpExcel, partitionExportableRecords } from '../utils/followUpExport';

/** 大樹聽力中心報告表頭主色（與診所回訪報告一致） */
const BRAND_BLUE = 'rgb(0, 140, 215)';

const REPORT_TITLE = '大樹聽力中心未成交潛力客戶追蹤名單';

export const QUICK_PTA_THRESHOLDS = [40, 55, 70];

const SORT_OPTIONS: { key: FollowUpSortKey; label: string }[] = [
    { key: 'pta-desc', label: '較差耳 PTA（重→輕）' },
    { key: 'date-desc', label: '到店日期（近→遠）' },
    { key: 'date-asc', label: '到店日期（久→近）' },
];

const DEGREE_OPTIONS = ['輕度', '中度', '中重度', '重度', '極重度'];

const DEGREE_BADGE_CLASSES: { [degree: string]: string } = {
    正常: 'bg-slate-100 text-slate-600 border-slate-200',
    輕度: 'bg-sky-50 text-sky-700 border-sky-200',
    中度: 'bg-amber-50 text-amber-700 border-amber-200',
    中重度: 'bg-orange-50 text-orange-700 border-orange-200',
    重度: 'bg-red-50 text-red-700 border-red-200',
    極重度: 'bg-red-100 text-red-800 border-red-300',
    未知: 'bg-slate-50 text-slate-500 border-slate-200',
};

/** 超過這個天數仍未成交，提醒優先聯繫。 */
const LONG_WAIT_DAYS = 90;

const DISTRIBUTION_LIMIT = 5;

const formatDateRange = (dateRange: DateRange) =>
    `${dateRange.startYear}年${dateRange.startMonth}月 ~ ${dateRange.endYear}年${dateRange.endMonth}月`;

const uniqueSorted = (values: string[]) =>
    Array.from(new Set(values.filter(Boolean))).sort((a, b) => a.localeCompare(b));

/** 選項消失（例如取消勾選某門市）後，舊的篩選值不能繼續默默生效。 */
const keepIfAvailable = (value: string, options: string[]) => (value && options.includes(value) ? value : '');

interface FollowUpCustomersProps {
    sources: CompetitionSheetSource[];
    currentSpreadsheetId: string;
    dateRange: DateRange;
    /** 全域 PTA 閾值；本頁的追蹤閾值預設跟隨它。 */
    ptaThreshold: number;
    savedSpreadsheets: SpreadsheetListItem[];
    extraSpreadsheetIds: string[];
    onToggleSpreadsheet?: (spreadsheetId: string) => void;
    selectionLimit: number;
    selectionMessage?: string;
    loading?: boolean;
    error?: string;
}

export const FollowUpCustomers: React.FC<FollowUpCustomersProps> = ({
    sources,
    currentSpreadsheetId,
    dateRange,
    ptaThreshold,
    savedSpreadsheets,
    extraSpreadsheetIds,
    onToggleSpreadsheet,
    selectionLimit,
    selectionMessage = '',
    loading = false,
    error = '',
}) => {
    const [thresholdInput, setThresholdInput] = useState<string>(String(ptaThreshold));
    const [storeFilter, setStoreFilter] = useState<string>('');
    const [audiologistFilter, setAudiologistFilter] = useState<string>('');
    const [degreeFilter, setDegreeFilter] = useState<string>('');
    const [sourceFilter, setSourceFilter] = useState<string>('');
    const [keyword, setKeyword] = useState<string>('');
    const [sortBy, setSortBy] = useState<FollowUpSortKey>('pta-desc');
    const [maskNames, setMaskNames] = useState<boolean>(false);
    const [exporting, setExporting] = useState<boolean>(false);
    const [exportError, setExportError] = useState<string>('');

    // 全域閾值改變時，本頁閾值重新跟上。
    useEffect(() => {
        setThresholdInput(String(ptaThreshold));
    }, [ptaThreshold]);

    const parsedThreshold = parseFloat(thresholdInput);
    const trackingThreshold = Number.isFinite(parsedThreshold) ? parsedThreshold : ptaThreshold;

    const dataset = useMemo(() => buildFollowUpDataset(sources, dateRange), [sources, dateRange]);

    const storeOptions = useMemo(() => uniqueSorted(Object.keys(dataset.storeTotals)), [dataset]);
    const audiologistOptions = useMemo(
        () => uniqueSorted(dataset.records.map((record) => record.audiologist)),
        [dataset]
    );
    const sourceOptions = useMemo(
        () => uniqueSorted(dataset.records.flatMap((record) => record.sources)),
        [dataset]
    );
    const referralOptions = useMemo(
        () => uniqueSorted(dataset.records.map((record) => record.referralUnit)).filter(
            (unit) => !sourceOptions.includes(unit)
        ),
        [dataset, sourceOptions]
    );

    const isMultiStore = storeOptions.length > 1;
    const activeStore = isMultiStore ? keepIfAvailable(storeFilter, storeOptions) : '';
    const activeAudiologist = keepIfAvailable(audiologistFilter, audiologistOptions);
    const activeSource = keepIfAvailable(sourceFilter, [...sourceOptions, ...referralOptions]);
    const activeDegree = keepIfAvailable(degreeFilter, DEGREE_OPTIONS);

    const selected = useMemo(
        () => selectFollowUpCustomers(dataset.records, {
            ptaThreshold: trackingThreshold,
            store: activeStore,
            audiologist: activeAudiologist,
            degree: activeDegree,
            source: activeSource,
            keyword,
            sortBy,
        }),
        [dataset, trackingThreshold, activeStore, activeAudiologist, activeDegree, activeSource, keyword, sortBy]
    );

    // Excel 只匯出電話格式正確的客戶，先算好人數讓使用者知道有幾位不會匯出。
    const exportPartition = useMemo(() => partitionExportableRecords(selected), [selected]);
    const exportableCount = exportPartition.exportable.length;
    const skippedExportCount = exportPartition.skipped.length;

    const summary = useMemo(
        () => summarizeFollowUpCustomers(selected, dataset, activeStore),
        [selected, dataset, activeStore]
    );

    const hasActiveFilters = Boolean(activeStore || activeAudiologist || activeDegree || activeSource || keyword.trim());

    const clearFilters = () => {
        setStoreFilter('');
        setAudiologistFilter('');
        setDegreeFilter('');
        setSourceFilter('');
        setKeyword('');
    };

    const topAudiologist = summary.audiologistDistribution[0];

    const printedStoreLabel = activeStore || storeOptions.join('、') || '—';

    const printedFilterLabel = [
        activeAudiologist && `主聽力師：${activeAudiologist}`,
        activeDegree && `聽損程度：${activeDegree}`,
        activeSource && `來源：${activeSource}`,
        keyword.trim() && `姓名含「${keyword.trim()}」`,
    ].filter(Boolean).join(' ｜ ');

    // 匯出內容與畫面一致：目前的篩選、排序與遮罩狀態。
    const handleExportExcel = async () => {
        setExporting(true);
        setExportError('');
        try {
            await downloadFollowUpExcel(
                selected,
                {
                    maskNames,
                    includeNote: dataset.hasNoteColumn,
                },
                { storeLabel: printedStoreLabel }
            );
        } catch (err: any) {
            setExportError(`匯出 Excel 失敗：${err?.message || '未知錯誤'}`);
        } finally {
            setExporting(false);
        }
    };

    return (
        <div className="space-y-6">
            <SpreadsheetSelectionPanel
                savedSpreadsheets={savedSpreadsheets}
                currentSpreadsheetId={currentSpreadsheetId}
                extraSpreadsheetIds={extraSpreadsheetIds}
                onToggleSpreadsheet={onToggleSpreadsheet}
                selectionLimit={selectionLimit}
                selectionMessage={selectionMessage}
                loading={loading}
                error={error}
            />

            {dataset.skippedSpreadsheets.length > 0 && (
                <div className="rounded-xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-800 print:hidden">
                    <p className="font-semibold">以下試算表未納入待追蹤名單：</p>
                    <ul className="mt-2 list-disc pl-5 space-y-1">
                        {dataset.skippedSpreadsheets.map((sheet) => (
                            <li key={sheet.spreadsheetId}>{sheet.spreadsheetTitle}：{sheet.reason}</li>
                        ))}
                    </ul>
                </div>
            )}

            {/* 篩選控制（列印時隱藏） */}
            <section className="rounded-2xl border border-rose-200 bg-gradient-to-br from-rose-50 via-white to-amber-50 p-5 shadow-sm print:hidden">
                <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
                    <div>
                        <p className="text-xs font-semibold uppercase tracking-[0.2em] text-rose-600">Follow-up List</p>
                        <h2 className="mt-2 text-xl font-bold text-slate-900">未成交潛力客戶追蹤</h2>
                        <p className="mt-2 max-w-2xl text-sm text-slate-600">
                            列出 {formatDateRange(dateRange)} 內尚未成交、且較差耳 PTA 大於追蹤閾值的客戶，方便安排回訪優先順序。
                        </p>
                    </div>
                    <div className="flex shrink-0 flex-col items-stretch gap-2 sm:flex-row sm:items-center">
                        <button
                            onClick={handleExportExcel}
                            disabled={exporting || exportableCount === 0}
                            className="bg-emerald-600 text-white px-5 py-2.5 rounded-lg hover:bg-emerald-700 transition-colors flex items-center justify-center gap-2 shadow-md hover:shadow-lg active:scale-95 text-sm font-medium disabled:cursor-not-allowed disabled:opacity-50 disabled:active:scale-100"
                        >
                            {exporting ? '產生 Excel 中…' : '📊 匯出 Excel（含篩選）'}
                        </button>
                        <button
                            onClick={() => window.print()}
                            className="bg-rose-600 text-white px-5 py-2.5 rounded-lg hover:bg-rose-700 transition-colors flex items-center justify-center gap-2 shadow-md hover:shadow-lg active:scale-95 text-sm font-medium"
                        >
                            🖨️ 匯出待追蹤名單 (PDF)
                        </button>
                    </div>
                </div>
                {selected.length > 0 && (
                    <p data-testid="follow-up-export-hint" className="mt-3 text-xs text-slate-500">
                        {!dataset.hasPhoneColumn
                            ? '試算表沒有電話欄位，無法匯出 Excel。'
                            : skippedExportCount > 0
                                ? `Excel 只匯出手機號碼正確的 ${exportableCount} 位；另有 ${skippedExportCount} 位沒有電話或號碼格式不符，不會匯出。`
                                : `Excel 將匯出 ${exportableCount} 位（含「待追蹤名單」與「電話名單」兩個工作表）。`}
                    </p>
                )}
                {exportError && <p className="mt-3 text-sm font-medium text-red-600">{exportError}</p>}

                {/* 追蹤 PTA 閾值 */}
                <div className="mt-5 flex flex-col gap-3 lg:flex-row lg:items-center">
                    <label htmlFor="follow-up-pta-threshold" className="text-sm font-semibold text-slate-700 whitespace-nowrap">
                        追蹤 PTA 閾值（較差耳 &gt;）
                    </label>
                    <div className="flex flex-wrap items-center gap-2">
                        <div className="flex items-center gap-1">
                            <input
                                id="follow-up-pta-threshold"
                                type="number"
                                min={0}
                                max={120}
                                value={thresholdInput}
                                onChange={(event) => setThresholdInput(event.target.value)}
                                className="w-24 rounded-lg border border-slate-200 bg-white p-2 text-sm focus:border-rose-500 focus:ring-rose-500"
                            />
                            <span className="text-sm text-slate-500">dB</span>
                        </div>
                        <div className="flex items-center gap-1 rounded-lg border border-slate-200 bg-white p-1">
                            {QUICK_PTA_THRESHOLDS.map((value) => (
                                <button
                                    key={value}
                                    onClick={() => setThresholdInput(String(value))}
                                    aria-pressed={trackingThreshold === value}
                                    className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-colors ${trackingThreshold === value
                                        ? 'bg-rose-50 text-rose-700 ring-1 ring-rose-200'
                                        : 'text-slate-500 hover:text-slate-700 hover:bg-slate-50'
                                        }`}
                                >
                                    {value} dB
                                </button>
                            ))}
                        </div>
                        {trackingThreshold !== ptaThreshold && (
                            <button
                                onClick={() => setThresholdInput(String(ptaThreshold))}
                                className="text-xs font-medium text-slate-500 underline-offset-2 hover:text-rose-600 hover:underline"
                            >
                                重設為全域閾值（{ptaThreshold} dB）
                            </button>
                        )}
                    </div>
                </div>

                {/* 輔助篩選 */}
                <div className="mt-4 grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
                    {isMultiStore && (
                        <FilterSelect label="門市" value={activeStore} onChange={setStoreFilter}>
                            {storeOptions.map((store) => (
                                <option key={store} value={store}>{store}</option>
                            ))}
                        </FilterSelect>
                    )}
                    <FilterSelect label="主聽力師" value={activeAudiologist} onChange={setAudiologistFilter}>
                        {audiologistOptions.map((name) => (
                            <option key={name} value={name}>{name}</option>
                        ))}
                    </FilterSelect>
                    <FilterSelect label="聽損程度" value={activeDegree} onChange={setDegreeFilter}>
                        {DEGREE_OPTIONS.map((degree) => (
                            <option key={degree} value={degree}>{degree}</option>
                        ))}
                    </FilterSelect>
                    <FilterSelect label="顧客來源／轉介" value={activeSource} onChange={setSourceFilter}>
                        {sourceOptions.length > 0 && (
                            <optgroup label="顧客來源">
                                {sourceOptions.map((source) => (
                                    <option key={`source-${source}`} value={source}>{source}</option>
                                ))}
                            </optgroup>
                        )}
                        {referralOptions.length > 0 && (
                            <optgroup label="轉介診所／門市">
                                {referralOptions.map((unit) => (
                                    <option key={`referral-${unit}`} value={unit}>{unit}</option>
                                ))}
                            </optgroup>
                        )}
                    </FilterSelect>
                    <label className="flex flex-col gap-1 text-xs font-semibold text-slate-600">
                        姓名關鍵字
                        <input
                            type="text"
                            value={keyword}
                            onChange={(event) => setKeyword(event.target.value)}
                            placeholder="輸入姓名..."
                            className="rounded-lg border border-slate-200 bg-white p-2 text-sm font-normal focus:border-rose-500 focus:ring-rose-500"
                        />
                    </label>
                </div>

                <div className="mt-4 flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
                    <div className="flex flex-wrap items-center gap-2">
                        <span className="text-xs font-semibold text-slate-600">排序</span>
                        <div className="flex flex-wrap items-center gap-1 rounded-lg border border-slate-200 bg-white p-1">
                            {SORT_OPTIONS.map((option) => (
                                <button
                                    key={option.key}
                                    onClick={() => setSortBy(option.key)}
                                    aria-pressed={sortBy === option.key}
                                    className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-colors ${sortBy === option.key
                                        ? 'bg-rose-50 text-rose-700 ring-1 ring-rose-200'
                                        : 'text-slate-500 hover:text-slate-700 hover:bg-slate-50'
                                        }`}
                                >
                                    {option.label}
                                </button>
                            ))}
                        </div>
                    </div>
                    <div className="flex items-center gap-4">
                        {hasActiveFilters && (
                            <button
                                onClick={clearFilters}
                                className="text-xs font-medium text-slate-500 hover:text-rose-600"
                            >
                                清除篩選條件
                            </button>
                        )}
                        <label className="flex items-center gap-2 text-sm text-slate-600">
                            <input
                                type="checkbox"
                                checked={maskNames}
                                onChange={(event) => setMaskNames(event.target.checked)}
                                className="h-4 w-4 rounded border-slate-300 text-rose-600 focus:ring-rose-500"
                            />
                            姓名／電話遮罩
                        </label>
                    </div>
                </div>
            </section>

            {/* 列印用抬頭：App 頁首與 tab 列印時隱藏，名單需自帶識別資訊 */}
            <div
                className="report-brand-header hidden print:block px-5 py-4 mb-4 text-white"
                style={{ backgroundColor: BRAND_BLUE }}
            >
                <h1 className="text-2xl font-bold">{REPORT_TITLE}</h1>
                <p className="mt-1 text-sm">
                    門市：{printedStoreLabel}
                    {` ｜ 分析區間：${formatDateRange(dateRange)}`}
                    {` ｜ 較差耳 PTA > ${trackingThreshold} dB`}
                    {` ｜ 共 ${summary.followUpCount} 位`}
                    {` ｜ 產出日期：${new Date().toLocaleDateString('zh-TW')}`}
                </p>
                {printedFilterLabel && <p className="mt-1 text-xs">篩選條件：{printedFilterLabel}</p>}
            </div>

            {!dataset.hasPTAColumn && sources.length > 0 && (
                <div className="rounded-xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-800">
                    找不到「左耳 PTA」或「右耳 PTA」欄位，無法依聽損程度篩選待追蹤客戶。請確認試算表表頭。
                </div>
            )}

            {/* KPI */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 print:grid-cols-4">
                <StatTile
                    testId="follow-up-kpi-count"
                    label="待追蹤總人數"
                    value={`${summary.followUpCount}`}
                    hint={`未成交且較差耳 PTA > ${trackingThreshold} dB`}
                    accent="text-rose-600"
                />
                <StatTile
                    testId="follow-up-kpi-rate"
                    label="佔同期未成交客戶"
                    value={`${summary.followUpRate.toFixed(0)}%`}
                    hint={`未成交 ${summary.undealtCount} 人／總來客 ${summary.totalCustomers} 人`}
                    accent="text-amber-600"
                />
                <StatTile
                    testId="follow-up-kpi-severe"
                    label={`中重度以上 (PTA > ${SEVERE_PTA_THRESHOLD})`}
                    value={`${summary.severeCount}`}
                    hint={summary.averageWorsePTA === null ? '平均較差耳 PTA —' : `平均較差耳 PTA ${summary.averageWorsePTA} dB`}
                    accent="text-red-600"
                />
                <StatTile
                    testId="follow-up-kpi-audiologist"
                    label="待追蹤最多的聽力師"
                    value={topAudiologist ? topAudiologist.name : '—'}
                    hint={topAudiologist ? `${topAudiologist.count} 位待追蹤` : '目前沒有待追蹤客戶'}
                    accent="text-sky-600"
                />
            </div>

            {/* 分布摘要 */}
            {summary.followUpCount > 0 && (
                <div className={`grid grid-cols-1 gap-4 ${isMultiStore ? 'md:grid-cols-3 print:grid-cols-3' : 'md:grid-cols-2 print:grid-cols-2'}`}>
                    <DistributionCard title="各主聽力師待追蹤人數" items={summary.audiologistDistribution} total={summary.followUpCount} />
                    <DistributionCard title="各顧客來源待追蹤人數" items={summary.sourceDistribution} total={summary.followUpCount} />
                    {isMultiStore && (
                        <DistributionCard title="各門市待追蹤人數" items={summary.storeDistribution} total={summary.followUpCount} />
                    )}
                </div>
            )}

            {/* 明細 */}
            <div className="clinic-report-detail bg-white rounded-xl shadow-sm border border-slate-100 overflow-hidden">
                <div className="p-6 border-b border-slate-100 flex items-center justify-between">
                    <h3 className="text-lg font-bold text-slate-800 border-l-4 border-rose-500 pl-3">待追蹤客戶明細</h3>
                    <span className="text-sm text-slate-500">共 {selected.length} 位</span>
                </div>

                {selected.length === 0 ? (
                    <p className="p-8 text-center text-sm text-slate-500">
                        {sources.length === 0
                            ? '尚未載入任何試算表。'
                            : `目前沒有符合條件的待追蹤客戶（未成交且較差耳 PTA > ${trackingThreshold} dB）。`}
                    </p>
                ) : (
                    <div className="overflow-x-auto print:overflow-visible">
                        <table className="clinic-report-table min-w-full divide-y divide-slate-100">
                            <thead style={{ backgroundColor: BRAND_BLUE }}>
                                <tr>
                                    {isMultiStore && <HeaderCell left>門市</HeaderCell>}
                                    <HeaderCell left>服務日期</HeaderCell>
                                    <HeaderCell left>姓名</HeaderCell>
                                    {dataset.hasPhoneColumn && <HeaderCell left>電話</HeaderCell>}
                                    <HeaderCell align="right">年齡</HeaderCell>
                                    <HeaderCell align="right">左耳 PTA</HeaderCell>
                                    <HeaderCell align="right">右耳 PTA</HeaderCell>
                                    <HeaderCell>聽損程度</HeaderCell>
                                    <HeaderCell>來源／轉介單位</HeaderCell>
                                    <HeaderCell>主聽力師</HeaderCell>
                                    {dataset.hasNoteColumn && <HeaderCell left>備註</HeaderCell>}
                                </tr>
                            </thead>
                            <tbody className="bg-white divide-y divide-slate-100 text-sm">
                                {selected.map((record, index) => (
                                    <FollowUpRow
                                        key={`${record.spreadsheetId}-${record.serviceDate}-${record.name}-${index}`}
                                        record={record}
                                        threshold={trackingThreshold}
                                        maskNames={maskNames}
                                        showStore={isMultiStore}
                                        showPhone={dataset.hasPhoneColumn}
                                        showNote={dataset.hasNoteColumn}
                                    />
                                ))}
                            </tbody>
                        </table>
                    </div>
                )}
            </div>

            <p className="text-xs text-slate-400 text-center">
                本名單依 {formatDateRange(dateRange)} 之來客紀錄產出，僅列出尚未成交、且左耳或右耳 PTA 大於 {trackingThreshold} dB 之客戶；
                雙耳 PTA 皆未填寫者不列入。
            </p>
        </div>
    );
};

interface FollowUpRowProps {
    record: FollowUpCustomerRecord;
    threshold: number;
    maskNames: boolean;
    showStore: boolean;
    showPhone: boolean;
    showNote: boolean;
}

const FollowUpRow: React.FC<FollowUpRowProps> = ({ record, threshold, maskNames, showStore, showPhone, showNote }) => {
    const daysSince = daysSinceTimestamp(record.sortKey);
    const isLongWait = daysSince !== null && daysSince > LONG_WAIT_DAYS;
    const sourceLabel = record.sources.join('、');

    return (
        <tr data-testid="follow-up-row" className="hover:bg-slate-50 transition-colors">
            {showStore && (
                <td className="clinic-report-left px-4 py-3 text-slate-700 whitespace-nowrap">{record.storeName}</td>
            )}
            <td className="clinic-report-left px-4 py-3 text-slate-700 whitespace-nowrap">
                <span>{record.serviceDate || '—'}</span>
                {daysSince !== null && (
                    <span
                        className={`ml-2 rounded-full px-2 py-0.5 text-[11px] font-semibold ${isLongWait
                            ? 'bg-amber-50 text-amber-700 border border-amber-200'
                            : 'bg-slate-100 text-slate-500'
                            }`}
                    >
                        {daysSince} 天前
                    </span>
                )}
            </td>
            <td className="clinic-report-left px-4 py-3 text-slate-800 font-medium whitespace-nowrap">
                {maskNames ? maskCustomerName(record.name) : record.name}
            </td>
            {showPhone && (
                <td className="clinic-report-left px-4 py-3 text-slate-700 whitespace-pre-line tabular-nums">
                    {record.phone ? (maskNames ? maskPhone(record.phone) : record.phone) : '—'}
                </td>
            )}
            <td className="px-4 py-3 text-right text-slate-700 whitespace-nowrap">
                {record.age !== null ? `${record.age} 歲` : '—'}
            </td>
            <td className={`px-4 py-3 text-right whitespace-nowrap ${record.leftPTA !== null && record.leftPTA > threshold ? 'text-rose-600 font-semibold' : 'text-slate-700'}`}>
                {record.leftPTA ?? '—'}
            </td>
            <td className={`px-4 py-3 text-right whitespace-nowrap ${record.rightPTA !== null && record.rightPTA > threshold ? 'text-rose-600 font-semibold' : 'text-slate-700'}`}>
                {record.rightPTA ?? '—'}
            </td>
            <td className="px-4 py-3 whitespace-nowrap">
                <span className={`inline-block rounded-full border px-2 py-0.5 text-xs font-semibold ${DEGREE_BADGE_CLASSES[record.hearingDegree] || DEGREE_BADGE_CLASSES['未知']}`}>
                    {record.hearingDegree}
                </span>
            </td>
            <td className="px-4 py-3 text-slate-700 whitespace-nowrap">
                <span>{sourceLabel || '—'}</span>
                {record.referralUnit && record.referralUnit !== sourceLabel && (
                    <span className="block text-xs text-slate-400">{record.referralUnit}</span>
                )}
            </td>
            <td className="px-4 py-3 text-slate-700 whitespace-nowrap">{record.audiologist}</td>
            {showNote && (
                <td className="clinic-report-left clinic-report-wrap px-4 py-3 text-slate-600 min-w-[10rem] whitespace-pre-wrap">
                    {record.note || '—'}
                </td>
            )}
        </tr>
    );
};

const HeaderCell: React.FC<{ children: React.ReactNode; left?: boolean; align?: 'left' | 'right' }> = ({
    children,
    left = false,
    align = 'left',
}) => (
    <th
        className={`${left ? 'clinic-report-left ' : ''}px-4 py-3 ${align === 'right' ? 'text-right' : 'text-left'} text-xs font-semibold text-white uppercase whitespace-nowrap`}
    >
        {children}
    </th>
);

interface FilterSelectProps {
    label: string;
    value: string;
    onChange: (value: string) => void;
    children: React.ReactNode;
}

const FilterSelect: React.FC<FilterSelectProps> = ({ label, value, onChange, children }) => (
    <label className="flex flex-col gap-1 text-xs font-semibold text-slate-600">
        {label}
        <select
            value={value}
            onChange={(event) => onChange(event.target.value)}
            className="rounded-lg border border-slate-200 bg-white p-2 text-sm font-normal focus:border-rose-500 focus:ring-rose-500"
        >
            <option value="">全部</option>
            {children}
        </select>
    </label>
);

interface StatTileProps {
    testId: string;
    label: string;
    value: string;
    hint: string;
    accent: string;
}

const StatTile: React.FC<StatTileProps> = ({ testId, label, value, hint, accent }) => (
    <div data-testid={testId} className="bg-white rounded-xl shadow-sm border border-slate-100 p-5">
        <p className="text-xs font-semibold text-slate-500 uppercase tracking-wider">{label}</p>
        <p className={`mt-2 text-3xl font-bold truncate ${accent}`}>{value}</p>
        <p className="mt-1 text-xs text-slate-400">{hint}</p>
    </div>
);

interface DistributionCardProps {
    title: string;
    items: { name: string; count: number }[];
    total: number;
}

const DistributionCard: React.FC<DistributionCardProps> = ({ title, items, total }) => {
    const visibleItems = items.slice(0, DISTRIBUTION_LIMIT);
    const hiddenCount = items.length - visibleItems.length;

    return (
        <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-5">
            <h4 className="text-sm font-bold text-slate-700 mb-3">{title}</h4>
            <ul className="space-y-2">
                {visibleItems.map((item) => (
                    <li key={item.name} className="text-sm">
                        <div className="flex items-center justify-between gap-2">
                            <span className="truncate text-slate-700">{item.name}</span>
                            <span className="shrink-0 font-semibold text-slate-800">{item.count}</span>
                        </div>
                        <div className="mt-1 h-1.5 rounded-full bg-slate-100">
                            <div
                                className="h-1.5 rounded-full bg-rose-400"
                                style={{ width: `${total > 0 ? Math.min(100, (item.count / total) * 100) : 0}%` }}
                            />
                        </div>
                    </li>
                ))}
            </ul>
            {hiddenCount > 0 && <p className="mt-2 text-xs text-slate-400">另有 {hiddenCount} 項</p>}
        </div>
    );
};

interface SpreadsheetSelectionPanelProps {
    savedSpreadsheets: SpreadsheetListItem[];
    currentSpreadsheetId: string;
    extraSpreadsheetIds: string[];
    onToggleSpreadsheet?: (spreadsheetId: string) => void;
    selectionLimit: number;
    selectionMessage: string;
    loading: boolean;
    error: string;
}

const SpreadsheetSelectionPanel: React.FC<SpreadsheetSelectionPanelProps> = ({
    savedSpreadsheets,
    currentSpreadsheetId,
    extraSpreadsheetIds,
    onToggleSpreadsheet,
    selectionLimit,
    selectionMessage,
    loading,
    error,
}) => {
    if (!savedSpreadsheets.length) {
        return null;
    }

    const selectedCount = extraSpreadsheetIds.length + 1;
    const isAtLimit = selectedCount >= selectionLimit;
    const currentInList = savedSpreadsheets.some((sheet) => sheet.id === currentSpreadsheetId);

    return (
        <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm print:hidden">
            <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
                <div>
                    <h2 className="text-base font-bold text-slate-900">合併分析範圍</h2>
                    <p className="mt-1 max-w-2xl text-sm text-slate-600">
                        目前載入的門市固定納入{currentInList ? '' : '（不在已儲存清單中）'}；可再勾選其他已儲存的試算表，合併成一張跨門市待追蹤名單。
                        這裡的選擇只影響本頁。
                    </p>
                </div>
                <div className="rounded-full border border-slate-200 bg-slate-50 px-4 py-2 text-sm font-semibold text-slate-700 whitespace-nowrap">
                    已選 {selectedCount} / {selectionLimit}
                </div>
            </div>

            <div className="mt-4 grid gap-3 md:grid-cols-2 xl:grid-cols-3">
                {savedSpreadsheets.map((sheet) => {
                    const isCurrent = sheet.id === currentSpreadsheetId;
                    const isChecked = isCurrent || extraSpreadsheetIds.includes(sheet.id);
                    const isDisabled = isCurrent || (!isChecked && isAtLimit);

                    return (
                        <label
                            key={sheet.id}
                            className={`flex items-start gap-3 rounded-xl border px-4 py-3 transition-all ${isChecked
                                ? 'border-rose-300 bg-white shadow-sm ring-2 ring-rose-100'
                                : 'border-slate-200 bg-white/80'
                                } ${isDisabled ? 'cursor-not-allowed' : 'cursor-pointer hover:border-rose-200'} ${!isChecked && isDisabled ? 'opacity-60' : ''}`}
                        >
                            <input
                                type="checkbox"
                                checked={isChecked}
                                disabled={isDisabled}
                                onChange={() => onToggleSpreadsheet?.(sheet.id)}
                                className="mt-1 h-4 w-4 rounded border-slate-300 text-rose-600 focus:ring-rose-500"
                            />
                            <span className="min-w-0 flex-1">
                                <span className="block truncate text-sm font-semibold text-slate-800">{sheet.title}</span>
                                <span className="mt-1 block text-xs text-slate-500">
                                    {isCurrent ? '目前門市（固定納入）' : '合併此門市的未成交客戶'}
                                </span>
                            </span>
                        </label>
                    );
                })}
            </div>

            {loading && <p className="mt-4 text-sm text-slate-500">載入合併試算表中…</p>}
            {error && <p className="mt-4 text-sm font-medium text-red-600">{error}</p>}
            {selectionMessage && <p className="mt-4 text-sm font-medium text-amber-700">{selectionMessage}</p>}
        </section>
    );
};
