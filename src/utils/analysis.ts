import { SheetData, DateRange, CustomerAnalysis, PeriodStats, CompetitionRankingEntry } from '../types';
import { parseAmount, getEarPTA, getHearingLossDegree, parseAge } from './parsers';

export interface AnalysisResult {
    customerAnalysis: CustomerAnalysis;
    salesmenAnalysis: any;
    clinicAnalysis: any;
    storeReferralAnalysis: any;
    hearingScreeningAnalysis: any;
    competitionRankingAnalysis: CompetitionRankingEntry[];
}

export const analyzeData = (
    sheetData: SheetData,
    dateRange: DateRange,
    ptaThreshold: number
): AnalysisResult | null => {

    if (!sheetData.values || sheetData.values.length < 2) return null;

    const values = sheetData.values;
    const headers = values[0];

    // 尋找相關欄位
    const dateColumnIndex = headers.findIndex(header =>
        header.includes('初次到店') || header.includes('日期') || header.includes('Date')
    );
    const statusColumnIndex = headers.findIndex(header =>
        header.includes('成交') || header.includes('狀態') || header.includes('Status')
    );
    const amountColumnIndex = headers.findIndex(header =>
        header.includes('金額') || header.includes('Amount') || header.includes('價格')
    );

    if (dateColumnIndex === -1) {
        console.error('找不到日期相關欄位，請確認資料表包含初次到店日期');
        return null;
    }

    const monthlyData: {
        [key: string]: {
            month: string;
            year: number;
            newCustomers: number;
            completedDeals: number;
            conversionRate: number;
            totalAmount: number;
            averageAmount: number;
        }
    } = {};

    let earliestDate: Date | null = null;
    let latestDate: Date | null = null;

    // 存儲符合條件的資料行
    const filteredRows: any[] = [];

    // 處理每一行數據
    for (let i = 1; i < values.length; i++) {
        if (!values[i] || !values[i][dateColumnIndex]) continue;

        const dateStr = values[i][dateColumnIndex].trim();
        if (!dateStr) continue;

        // 解析日期
        let date: Date | null = null;
        const dateFormats = [
            /(\d{4})[/-](\d{1,2})[/-](\d{1,2})/,
            /(\d{1,2})[/-](\d{1,2})[/-](\d{4})/,
            /(\d{4})年(\d{1,2})月(\d{1,2})日?/,
        ];

        for (const format of dateFormats) {
            const match = dateStr.match(format);
            if (match) {
                if (format === dateFormats[0]) {
                    date = new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));
                } else if (format === dateFormats[1]) {
                    date = new Date(parseInt(match[3]), parseInt(match[1]) - 1, parseInt(match[2]));
                } else if (format === dateFormats[2]) {
                    date = new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));
                }
                break;
            }
        }

        if (!date) {
            date = new Date(dateStr);
            if (isNaN(date.getTime())) continue;
        }

        if (!date || isNaN(date.getTime())) continue;

        // 檢查是否在選擇的月份區間內
        const dataYear = date.getFullYear();
        const dataMonth = date.getMonth() + 1;

        const startDate = new Date(dateRange.startYear, dateRange.startMonth - 1, 1);
        const endDate = new Date(dateRange.endYear, dateRange.endMonth, 0);

        if (date < startDate || date > endDate) continue;

        // 將符合條件的資料行添加到篩選結果中，包含所有欄位的詳細信息
        const rowData: any = {
            基本信息: {
                原始行號: i + 1,
                日期字段: dateStr,
                解析日期: date.toLocaleDateString('zh-TW'),
                年月: `${dataYear}年${dataMonth}月`,
                狀態: statusColumnIndex !== -1 && values[i][statusColumnIndex] ? values[i][statusColumnIndex] : '無狀態欄位',
                金額: amountColumnIndex !== -1 && values[i][amountColumnIndex] ? values[i][amountColumnIndex] : '無金額欄位'
            },
            所有欄位數據: {}
        };

        // 添加所有欄位的詳細信息
        headers.forEach((header: string, index: number) => {
            const cellValue = values[i][index] || '';
            rowData.所有欄位數據[header] = {
                欄位名稱: header,
                欄位索引: index,
                內容: cellValue,
                是否為空: cellValue.trim() === '',
                字符長度: cellValue.length
            };
        });

        filteredRows.push(rowData);

        const monthKey = `${dataYear}-${String(dataMonth).padStart(2, '0')}`;

        if (!monthlyData[monthKey]) {
            const monthNames = ['1月', '2月', '3月', '4月', '5月', '6月',
                '7月', '8月', '9月', '10月', '11月', '12月'];
            monthlyData[monthKey] = {
                month: `${dataYear}年${monthNames[dataMonth - 1]}`,
                year: dataYear,
                newCustomers: 0,
                completedDeals: 0,
                conversionRate: 0,
                totalAmount: 0,
                averageAmount: 0,
            };
        }

        if (!earliestDate || date < earliestDate) earliestDate = date;
        if (!latestDate || date > latestDate) latestDate = date;
    }

    if (filteredRows.length === 0) return null;

    // 將符合條件的資料整合成物件陣列
    const customersArray = filteredRows.map((row) => {
        const customerObject: { [key: string]: string } = {};

        // 將每個表頭作為key，對應的資料作為value
        headers.forEach((header: string) => {
            const cellData = row.所有欄位數據[header];
            customerObject[header] = cellData ? cellData.內容 : '';
        });

        return customerObject;
    });

    // Helper to check if a customer is dealt
    const checkIsDealt = (customer: { [key: string]: string }) => {
        const keysToCheck = [
            '是否成交', '成交', '狀態', 'Status', 'IsDealt', 'Dealt',
            '是否借機', '是否有借機', '借機'
        ];

        // Also check the detected status column
        if (statusColumnIndex !== -1) {
            keysToCheck.push(headers[statusColumnIndex]);
        }

        return keysToCheck.some(key => {
            const val = (customer[key] || '').trim();
            return ['是', 'TRUE', 'True', 'true', '成交', '已成交', 'OK', 'Success'].includes(val);
        });
    };

    const invalidStoreNames = new Set(['', '#N/A', 'N/A', '#REF!']);
    const storeNameHeader = headers.find(header =>
        header.includes('轉介門市') ||
        header.includes('門市名稱') ||
        header.toLowerCase().includes('store')
    ) || headers[11];

    const getStoreReferralName = (customer: { [key: string]: string }) => {
        const rawStoreName = storeNameHeader ? (customer[storeNameHeader] || '').trim() : '';
        return invalidStoreNames.has(rawStoreName) ? '' : rawStoreName;
    };

    // 新增分析變數
    const ageGroups: { [key: string]: number } = {
        '0-18歲': 0, '19-44歲': 0, '45-64歲': 0, '65-79歲': 0, '80歲以上': 0, '未填寫': 0
    };

    const sourceStats: { [key: string]: { count: number; deals: number } } = {};
    const hearingLossStats: { [key: string]: number } = {};
    const weekdayStats: { [key: number]: { visits: number; deals: number } } = {
        0: { visits: 0, deals: 0 }, 1: { visits: 0, deals: 0 }, 2: { visits: 0, deals: 0 },
        3: { visits: 0, deals: 0 }, 4: { visits: 0, deals: 0 }, 5: { visits: 0, deals: 0 },
        6: { visits: 0, deals: 0 }
    };

    // 業務員分析
    const salesmenAnalysis: {
        [key: string]: {
            業務員: string,
            訂單數量: number,
            潛力客戶數: number,
            成交率: number,
            當季業績累積: number
        }
    } = {};

    // 計算整體數據
    let totalPotentialCustomers = 0; // 總來客數（基於PTA閾值）
    let totalOrders = 0; // 總訂單數（基於是否成交）
    let totalRevenue = 0; // 總營業額

    // 先初始化所有業務員的統計數據
    customersArray.forEach(customer => {
        const salesman = customer['主聽力師'] || customer['聽力師'] || '未知業務員';

        if (!salesmenAnalysis[salesman]) {
            salesmenAnalysis[salesman] = {
                業務員: salesman,
                訂單數量: 0,
                潛力客戶數: 0,
                成交率: 0,
                當季業績累積: 0
            };
        }
    });

    // 分析每個業務員的數據 & 累積新分析數據
    customersArray.forEach(customer => {
        const salesman = customer['主聽力師'] || customer['聽力師'] || '未知業務員';
        const leftEarPTA = getEarPTA(customer, '左');
        const rightEarPTA = getEarPTA(customer, '右');

        // 判斷是否成交
        const isDealt = checkIsDealt(customer);

        // 判斷是否為潛力客戶：左/右 PTA 高於閾值 **或** 已成交
        const isPotentialCustomer = leftEarPTA > ptaThreshold || rightEarPTA > ptaThreshold || isDealt;

        // 如果是潛力客戶
        if (isPotentialCustomer) {
            salesmenAnalysis[salesman].潛力客戶數++;
            totalPotentialCustomers++;

            // --- 年齡分析 ---
            const ageStr = customer['年齡'];
            const birthDateStr = customer['顧客生日 (西元/月/日)'] || customer['生日'] || customer['BirthDate'];
            const age = parseAge(ageStr, birthDateStr);

            if (isNaN(age)) {
                ageGroups['未填寫']++;
            } else if (age <= 18) {
                ageGroups['0-18歲']++;
            } else if (age <= 44) {
                ageGroups['19-44歲']++;
            } else if (age <= 64) {
                ageGroups['45-64歲']++;
            } else if (age <= 79) {
                ageGroups['65-79歲']++;
            } else {
                ageGroups['80歲以上']++;
            }

            // --- 來源分析 ---
            // 動態尋找來源欄位（可能叫「顧客來源」、「來源」、「Source」等）
            const sourceKey = Object.keys(customer).find(k =>
                k.includes('來源') || k.toLowerCase().includes('source') || k.includes('渠道')
            );
            const sourceRaw = sourceKey ? customer[sourceKey] : '';

            // 來源可能包含換行或逗號，拆分多重來源
            const sources = sourceRaw
                .split(/[\n,，、;；]/)
                .map((s: string) => s.trim())
                .filter((s: string) => s !== '' && s !== '#N/A' && s !== 'N/A');

            if (sources.length === 0) sources.push('未知');

            sources.forEach((source: string) => {
                if (!sourceStats[source]) sourceStats[source] = { count: 0, deals: 0 };
                sourceStats[source].count++;
                if (isDealt) sourceStats[source].deals++;
            });

            // --- 聽損程度分析 (取較差耳) ---
            const worsePTA = Math.max(isNaN(leftEarPTA) ? 0 : leftEarPTA, isNaN(rightEarPTA) ? 0 : rightEarPTA);

            // Only count degrees above the threshold setting
            if (worsePTA > ptaThreshold) {
                const degree = getHearingLossDegree(worsePTA);
                if (!hearingLossStats[degree]) hearingLossStats[degree] = 0;
                hearingLossStats[degree]++;
            }
        }

        if (isDealt) {
            salesmenAnalysis[salesman].訂單數量++;
            totalOrders++;
        }


        // --- 平日/假日分析 ---
        // 使用服務日期

        let dateObj: Date | null = null;
        const possibleDateCols = ['服務日期', '初次到店', '日期'];
        for (const col of possibleDateCols) {
            if (customer[col]) {
                dateObj = new Date(customer[col]);
                if (!isNaN(dateObj.getTime())) break;
            }
        }

        if (dateObj) {
            const day = dateObj.getDay(); // 0 (Sun) - 6 (Sat)
            weekdayStats[day].visits++;
            if (isDealt) weekdayStats[day].deals++;
        }

    });

    // 格式化新分析數據
    const ageAnalysis = Object.entries(ageGroups).map(([range, count]) => ({ range, count }));

    const sourceAnalysis = Object.entries(sourceStats)
        .map(([source, stats]) => ({
            source,
            count: stats.count,
            conversionRate: stats.count > 0 ? (stats.deals / stats.count) * 100 : 0
        }))
        .sort((a, b) => b.count - a.count);

    const hearingLossAnalysis = Object.entries(hearingLossStats)
        .map(([degree, count]) => ({ degree, count }))
        .sort((a, b) => b.count - a.count); // 或是可以照聽損程度排序，這裡先照數量

    // 計算每個業務員的當季累積業績（成交客戶的金額加總）
    Object.keys(salesmenAnalysis).forEach(salesman => {
        // 篩選出該業務員的所有客戶
        const salesmanCustomers = customersArray.filter(customer =>
            (customer['主聽力師'] || customer['聽力師'] || '未知業務員') === salesman
        );

        // 從該業務員的客戶中，篩選出成交的客戶並加總金額
        let totalAmount = 0;
        salesmanCustomers.forEach(customer => {
            const isDealt = checkIsDealt(customer);

            const rawAmount = customer['成交金額'] || customer['金額'] || customer['價格'] || customer['營業額'] || '';
            const dealAmount = parseAmount(rawAmount);
            if (isDealt && !isNaN(dealAmount) && dealAmount > 0) {
                totalAmount += dealAmount;
                totalRevenue += dealAmount;
            }
        });

        salesmenAnalysis[salesman].當季業績累積 = totalAmount;
    });

    // 計算成交率
    Object.keys(salesmenAnalysis).forEach(salesman => {
        const data = salesmenAnalysis[salesman];
        if (data.潛力客戶數 > 0) {
            data.成交率 = (data.訂單數量 / data.潛力客戶數) * 100;
        }
    });

    // 計算整體統計數據
    const overallConversionRate = totalPotentialCustomers > 0 ? (totalOrders / totalPotentialCustomers) * 100 : 0;

    // 生成月份統計數據 (基於customersArray)
    const customerMonthlyData: {
        [key: string]: {
            month: string;
            year: number;
            newCustomers: number;
            completedDeals: number;
            conversionRate: number;
            totalAmount: number;
            averageAmount: number;
        }
    } = {};

    // 分析customersArray中的每個客戶的月份數據
    customersArray.forEach(customer => {
        const serviceDate = customer['服務日期'] || customer['初次到店'] || '';
        if (!serviceDate) return;

        const date = new Date(serviceDate);
        if (isNaN(date.getTime())) return;

        const monthKey = `${date.getFullYear()}年${date.getMonth() + 1}月`;

        if (!customerMonthlyData[monthKey]) {
            customerMonthlyData[monthKey] = {
                month: monthKey,
                year: date.getFullYear(),
                newCustomers: 0,
                completedDeals: 0,
                conversionRate: 0,
                totalAmount: 0,
                averageAmount: 0
            };
        }

        // PTA數據判斷
        const leftEarPTA = getEarPTA(customer, '左');
        const rightEarPTA = getEarPTA(customer, '右');
        const isDealtClinic = checkIsDealt(customer);
        const isPotential = leftEarPTA > ptaThreshold || rightEarPTA > ptaThreshold || isDealtClinic;

        // 成交判斷
        const isDealt = checkIsDealt(customer);

        // 成交金額
        const rawAmountMonthly = customer['成交金額'] || customer['金額'] || customer['價格'] || customer['營業額'] || '';
        const dealAmount = parseAmount(rawAmountMonthly);

        if (isPotential) {
            customerMonthlyData[monthKey].newCustomers++;
        }

        if (isDealt) {
            customerMonthlyData[monthKey].completedDeals++;
            if (!isNaN(dealAmount) && dealAmount > 0) {
                customerMonthlyData[monthKey].totalAmount += dealAmount;
            }
        }
    });

    // 計算每月的成交率和平均金額
    Object.keys(customerMonthlyData).forEach(month => {
        const data = customerMonthlyData[month];
        if (data.newCustomers > 0) {
            data.conversionRate = (data.completedDeals / data.newCustomers) * 100;
        }
        if (data.completedDeals > 0) {
            data.averageAmount = data.totalAmount / data.completedDeals;
        }
    });

    // 轉換為陣列並排序
    const sortedMonthlyData = Object.values(customerMonthlyData).sort((a, b) => {
        if (a.year !== b.year) return a.year - b.year;

        const aMonth = parseInt(a.month.replace(/\d+年(\d+)月/, '$1'));
        const bMonth = parseInt(b.month.replace(/\d+年(\d+)月/, '$1'));

        return aMonth - bMonth;
    });

    /* ====================== 診所轉介分析 ====================== */
    const clinicSummary: {
        [key: string]: {
            clinic: string;
            total: number;
            potential: number;
            dealt: number;
            conversionRate: number;
            totalAmount: number;
        }
    } = {};

    customersArray.forEach(customer => {
        const clinicName = (customer['診所名稱'] || '').trim();
        if (!clinicName) return; // 無診所名稱視為非診所轉介

        if (!clinicSummary[clinicName]) {
            clinicSummary[clinicName] = {
                clinic: clinicName,
                total: 0,
                potential: 0,
                dealt: 0,
                conversionRate: 0,
                totalAmount: 0,
            };
        }

        clinicSummary[clinicName].total++;

        const leftPTA = getEarPTA(customer, '左');
        const rightPTA = getEarPTA(customer, '右');
        const isDealtClinic = customer['是否成交'] === '是' || customer['是否成交'] === 'TRUE' ||
            customer['是否借機'] === 'TRUE' || customer['成交'] === '是' ||
            customer['成交'] === 'TRUE' || customer['狀態'] === '成交' || customer['狀態'] === '已成交';
        const isPotential = leftPTA > ptaThreshold || rightPTA > ptaThreshold || isDealtClinic;

        if (isPotential) clinicSummary[clinicName].potential++;

        const isDealt = checkIsDealt(customer);

        if (isDealt) {
            clinicSummary[clinicName].dealt++;
            const dealAmt = parseAmount(customer['成交金額'] || customer['金額'] || customer['價格'] || customer['營業額']);
            if (!isNaN(dealAmt) && dealAmt > 0) {
                clinicSummary[clinicName].totalAmount += dealAmt;
            }
        }
    });

    Object.values(clinicSummary).forEach(c => {
        c.conversionRate = c.total > 0 ? (c.dealt / c.total) * 100 : 0;
    });

    // 轉陣列並排序 (依總轉介數)
    const clinicArray = Object.values(clinicSummary).sort((a, b) => b.total - a.total);

    /* ====================== 門市轉介分析 ====================== */
    const storeSummary: { [key: string]: { store: string; total: number; potential: number; dealt: number; totalAmount: number; conversionRate: number; } } = {};

    customersArray.forEach(customer => {
        const storeName = getStoreReferralName(customer);
        if (!storeName) return;

        if (!storeSummary[storeName]) {
            storeSummary[storeName] = {
                store: storeName,
                total: 0,
                potential: 0,
                dealt: 0,
                totalAmount: 0,
                conversionRate: 0
            };
        }
        storeSummary[storeName].total++;

        const leftPTA = getEarPTA(customer, '左');
        const rightPTA = getEarPTA(customer, '右');

        // 成交判斷
        const isDealt = checkIsDealt(customer);

        const isPotential = leftPTA > ptaThreshold || rightPTA > ptaThreshold || isDealt;

        if (isPotential) {
            storeSummary[storeName].potential++;
        }

        if (isDealt) {
            storeSummary[storeName].dealt++;
            const dealAmt = parseAmount(customer['成交金額'] || customer['金額'] || customer['價格'] || customer['營業額']);
            if (!isNaN(dealAmt) && dealAmt > 0) {
                storeSummary[storeName].totalAmount += dealAmt;
            }
        }
    });

    Object.values(storeSummary).forEach(s => {
        s.conversionRate = s.total > 0 ? (s.dealt / s.total) * 100 : 0;
    });

    const storeArray = Object.values(storeSummary).sort((a, b) => b.total - a.total);

    const competitionRankingAnalysis: CompetitionRankingEntry[] = Object.values(storeSummary)
        .map(store => ({
            rank: 0,
            storeName: store.store,
            potentialCustomers: store.potential,
            nonPotentialCustomers: Math.max(store.total - store.potential, 0),
        }))
        .sort((a, b) => {
            if (b.potentialCustomers !== a.potentialCustomers) {
                return b.potentialCustomers - a.potentialCustomers;
            }

            if (b.nonPotentialCustomers !== a.nonPotentialCustomers) {
                return b.nonPotentialCustomers - a.nonPotentialCustomers;
            }

            return a.storeName.localeCompare(b.storeName);
        })
        .map((entry, index) => ({
            ...entry,
            rank: index + 1,
        }));

    /* ====================== 聽篩活動來源（按月份）分析 ====================== */
    const hearingSummary: { [key: string]: { month: string; year: number; total: number; potential: number; dealt: number; conversionRate: number; totalAmount: number; names: string[]; } } = {};

    customersArray.forEach(customer => {
        // 嘗試尋找含「顧客來源」字樣的欄位 (可能包含換行/特殊符號)
        const sourceKey = Object.keys(customer).find(k => k.replace(/\s/g, '').includes('顧客來源'));
        const sourceVal = sourceKey ? customer[sourceKey] : '';
        if (!sourceVal || !sourceVal.toString().includes('聽篩')) return; // 僅統計含「聽篩」字樣

        const serviceDate = customer['服務日期'] || customer['初次到店'] || '';
        if (!serviceDate) return;
        const date = new Date(serviceDate);
        if (isNaN(date.getTime())) return;

        const monthKey = `${date.getFullYear()}年${date.getMonth() + 1}月`;
        if (!hearingSummary[monthKey]) {
            hearingSummary[monthKey] = { month: monthKey, year: date.getFullYear(), total: 0, potential: 0, dealt: 0, conversionRate: 0, totalAmount: 0, names: [] };
        }

        hearingSummary[monthKey].total++;

        // PTA & 成交
        const leftPTA = getEarPTA(customer, '左');
        const rightPTA = getEarPTA(customer, '右');
        const isDealtHS = checkIsDealt(customer);

        if (leftPTA > ptaThreshold || rightPTA > ptaThreshold || isDealtHS) {
            hearingSummary[monthKey].potential++;
        }

        if (isDealtHS) {
            hearingSummary[monthKey].dealt++;
            const dealAmt = parseAmount(customer['成交金額'] || customer['金額'] || customer['價格'] || customer['營業額']);
            if (!isNaN(dealAmt) && dealAmt > 0) {
                hearingSummary[monthKey].totalAmount += dealAmt;
            }
        }

        // 取得客戶姓名
        const nameKey = Object.keys(customer).find(k => k.replace(/\s/g, '').includes('姓名') || k.toLowerCase().includes('name'));
        const customerName = nameKey ? customer[nameKey] : '未命名';
        hearingSummary[monthKey].names.push(customerName);
    });

    // 計算成交率並排序
    Object.values(hearingSummary).forEach(item => {
        item.conversionRate = item.total > 0 ? (item.dealt / item.total) * 100 : 0;
    });

    const hearingArray = Object.values(hearingSummary).sort((a, b) => {
        if (a.year !== b.year) return a.year - b.year;
        const aMonth = parseInt(a.month.replace(/\d+年(\d+)月/, '$1'));
        const bMonth = parseInt(b.month.replace(/\d+年(\d+)月/, '$1'));
        return aMonth - bMonth;
    });

    // --- 比較分析數據彙整 (Comparison Buckets) ---
    const yearBuckets: { [key: string]: PeriodStats } = {};
    const quarterBuckets: { [key: string]: PeriodStats } = {};
    const monthBuckets: { [key: string]: PeriodStats } = {};

    const initBucket = (bucketMap: { [key: string]: PeriodStats }, key: string) => {
        if (!bucketMap[key]) {
            bucketMap[key] = {
                period: key,
                newCustomers: 0,
                completedDeals: 0,
                totalAmount: 0,
                conversionRate: 0,
                averageOrderValue: 0,
                ageDistribution: {},
                sourceDistribution: {},
                hearingLossDistribution: {},
                salespersonPerformance: {}
            };
        }
    };

    const updateBucket = (
        bucketMap: { [key: string]: PeriodStats },
        key: string,
        isPotential: boolean,
        isDealt: boolean,
        dealAmount: number,
        age: number,
        source: string,
        hearingDegree: string,
        salesperson: string
    ) => {
        initBucket(bucketMap, key);
        const bucket = bucketMap[key];

        if (isPotential) {
            bucket.newCustomers++;
            // Update Age Distribution
            if (age > 0) {
                const range = age < 20 ? '20歲以下' :
                    age >= 90 ? '90歲以上' :
                        `${Math.floor(age / 10) * 10}-${Math.floor(age / 10) * 10 + 9}歲`;
                bucket.ageDistribution[range] = (bucket.ageDistribution[range] || 0) + 1;
            }
            // Update Source Distribution
            if (source) {
                bucket.sourceDistribution[source] = (bucket.sourceDistribution[source] || 0) + 1;
            }
            if (hearingDegree && hearingDegree !== 'Check') {
                bucket.hearingLossDistribution[hearingDegree] = (bucket.hearingLossDistribution[hearingDegree] || 0) + 1;
            }

            // Update Salesperson Visits (Potential Customers)
            if (salesperson) {
                if (!bucket.salespersonPerformance[salesperson]) {
                    bucket.salespersonPerformance[salesperson] = { visits: 0, deals: 0, revenue: 0 };
                }
                bucket.salespersonPerformance[salesperson].visits++;
            }
        }

        if (isDealt) {
            bucket.completedDeals++;
            if (dealAmount > 0) bucket.totalAmount += dealAmount;

            // Update Salesperson Deals & Revenue
            if (salesperson) {
                // Ensure initialized
                if (!bucket.salespersonPerformance[salesperson]) {
                    bucket.salespersonPerformance[salesperson] = { visits: 1, deals: 0, revenue: 0 };
                }
                bucket.salespersonPerformance[salesperson].deals++;
                if (dealAmount > 0) {
                    bucket.salespersonPerformance[salesperson].revenue += dealAmount;
                }
            }
        }
    };

    // Calculate buckets using customersArray which has normalized data
    customersArray.forEach(customer => {
        const serviceDate = customer['服務日期'] || customer['初次到店'] || '';
        if (!serviceDate) return;
        const date = new Date(serviceDate);
        if (isNaN(date.getTime())) return;

        const leftPTA = getEarPTA(customer, '左');
        const rightPTA = getEarPTA(customer, '右');
        const worsePTA = Math.max(leftPTA, rightPTA);
        const isDealt = checkIsDealt(customer);
        const isPotential = leftPTA > ptaThreshold || rightPTA > ptaThreshold || isDealt;

        const dealRaw = customer['成交金額'] || customer['金額'] || customer['價格'] || customer['營業額'] || '';
        const dealAmount = parseAmount(dealRaw);
        const validDealAmount = (!isNaN(dealAmount) && dealAmount > 0) ? dealAmount : 0;

        const age = parseAge(customer['年齡'] || customer['Age'], customer['出生日期'] || customer['生日']);

        // Dynamic Source Key Finding (matching Overview logic)
        const sourceKey = Object.keys(customer).find(k =>
            k.includes('來源') || k.toLowerCase().includes('source') || k.includes('渠道')
        );
        const sourceVal = (sourceKey ? customer[sourceKey] : '').trim();

        let finalSource = sourceVal;
        const storeReferralName = getStoreReferralName(customer);
        if (storeReferralName) {
            finalSource = storeReferralName;
        }

        const hearingDegree = worsePTA > ptaThreshold ? getHearingLossDegree(worsePTA) : '';

        // Dynamic Salesperson Key Finding
        const salesKey = Object.keys(customer).find(k =>
            k.includes('選配師') || k.toLowerCase().includes('sales') || k.includes('業務') || k.includes('主聽力師')
        );
        let salesperson = (salesKey ? customer[salesKey] : '').trim();

        // Normalize Salesperson Name (Title Case)
        if (salesperson) {
            salesperson = salesperson.toLowerCase().replace(/(?:^|\s)\S/g, a => a.toUpperCase());
        }

        const year = date.getFullYear();
        const month = date.getMonth() + 1;
        const quarter = Math.ceil(month / 3);

        const yearKey = `${year}`;
        const quarterKey = `${year}-Q${quarter}`;
        const monthKey = `${year}-${String(month).padStart(2, '0')}`;

        // Pass extra data for distributions
        updateBucket(yearBuckets, yearKey, isPotential, isDealt, validDealAmount, age, finalSource, hearingDegree, salesperson);
        updateBucket(quarterBuckets, quarterKey, isPotential, isDealt, validDealAmount, age, finalSource, hearingDegree, salesperson);
        updateBucket(monthBuckets, monthKey, isPotential, isDealt, validDealAmount, age, finalSource, hearingDegree, salesperson);
    });

    // Calculate rates for buckets
    const finalizeBuckets = (bucketMap: { [key: string]: PeriodStats }) => {
        Object.values(bucketMap).forEach(stat => {
            stat.conversionRate = stat.newCustomers > 0 ? (stat.completedDeals / stat.newCustomers) * 100 : 0;
            stat.averageOrderValue = stat.completedDeals > 0 ? Math.round(stat.totalAmount / stat.completedDeals) : 0;
        });
    };

    finalizeBuckets(yearBuckets);
    finalizeBuckets(quarterBuckets);
    finalizeBuckets(monthBuckets);

    return {
        customerAnalysis: {
            monthlyData: sortedMonthlyData,
            totalCustomers: totalPotentialCustomers,
            totalCompletedDeals: totalOrders,
            totalAmount: totalRevenue,
            overallConversionRate: overallConversionRate,
            dateRange: {
                earliest: customersArray.length > 0 ? customersArray[0]['服務日期'] || customersArray[0]['初次到店'] || '' : '',
                latest: customersArray.length > 0 ? customersArray[customersArray.length - 1]['服務日期'] || customersArray[customersArray.length - 1]['初次到店'] || '' : ''
            },
            ageAnalysis,
            sourceAnalysis,
            hearingLossAnalysis,
            weekdayAnalysis: [
                { day: '週日', ...weekdayStats[0] },
                { day: '週一', ...weekdayStats[1] },
                { day: '週二', ...weekdayStats[2] },
                { day: '週三', ...weekdayStats[3] },
                { day: '週四', ...weekdayStats[4] },
                { day: '週五', ...weekdayStats[5] },
                { day: '週六', ...weekdayStats[6] },
            ].map(d => ({ ...d, conversionRate: d.visits > 0 ? (d.deals / d.visits) * 100 : 0 })),
            yearBuckets,
            quarterBuckets,
            monthBuckets
        },
        salesmenAnalysis: salesmenAnalysis,
        clinicAnalysis: clinicArray,
        storeReferralAnalysis: storeArray,
        hearingScreeningAnalysis: hearingArray,
        competitionRankingAnalysis,
    };
};
