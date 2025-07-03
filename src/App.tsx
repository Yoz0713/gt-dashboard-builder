import React, { useState, useCallback } from 'react';
import { useGoogleLogin, googleLogout } from '@react-oauth/google';
import { Chart as ChartJS, CategoryScale, LinearScale, BarElement, Title, Tooltip, Legend } from 'chart.js';
import { Bar } from 'react-chartjs-2';
import { UserProfile, SheetData, SheetInfo, SpreadsheetMetadata, CustomerAnalysis, DateRange } from './types';

// Register Chart.js components
ChartJS.register(CategoryScale, LinearScale, BarElement, Title, Tooltip, Legend);

// 自訂長條圖數值標籤插件（僅用於需要顯示數字的圖表）
const barLabelPlugin = {
  id: 'barLabelPlugin',
  afterDatasetsDraw: (chart: any) => {
    const { ctx } = chart;
    ctx.save();
    chart.data.datasets.forEach((dataset: any, datasetIndex: number) => {
      const meta = chart.getDatasetMeta(datasetIndex);
      meta.data.forEach((bar: any, index: number) => {
        const data = dataset.data[index];
        ctx.fillStyle = '#000';
        ctx.font = '10px sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'bottom';
        ctx.fillText(data, bar.x, bar.y - 4);
      });
    });
    ctx.restore();
  },
};

const App: React.FC = () => {
  const [user, setUser] = useState<UserProfile | null>(null);
  const [accessToken, setAccessToken] = useState<string>('');
  const [sheetUrl, setSheetUrl] = useState<string>('');
  const [spreadsheetId, setSpreadsheetId] = useState<string>('');
  const [sheetData, setSheetData] = useState<SheetData | null>(null);
  const [availableSheets, setAvailableSheets] = useState<SheetInfo[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string>('');
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string>('');
  const [customerAnalysis, setCustomerAnalysis] = useState<CustomerAnalysis | null>(null);
  const [dateRange, setDateRange] = useState<DateRange>({
    startYear: new Date().getFullYear(),
    startMonth: 1,
    endYear: new Date().getFullYear(),
    endMonth: 12,
  });
  const [ptaThreshold, setPtaThreshold] = useState(40); // PTA閾值，預設40dBHL
  const [selectedStore, setSelectedStore] = useState('桃園藝文店'); // 選擇的店家
  const [salesmenAnalysis, setSalesmenAnalysis] = useState<any>(null); // 業務員分析數據
  const [clinicAnalysis, setClinicAnalysis] = useState<any>(null); // 診所分析數據
  const [storeReferralAnalysis, setStoreReferralAnalysis] = useState<any>(null); // 門市轉介分析
  const [hearingScreeningAnalysis, setHearingScreeningAnalysis] = useState<any>(null); // 聽篩活動來源分析

  // 店家選項
  const storeOptions = [
    '桃園藝文店',
    '桃園龜山店', 
    '桃園內壢二店',
    '桃園環東店',
    '新竹湖口店',
    '北屯崇德店'
  ];

  // Google OAuth login
  const login = useGoogleLogin({
    onSuccess: (response) => {
      console.log('Login successful, access token received');
      setAccessToken(response.access_token);
      fetchUserProfile(response.access_token);
    },
    onError: (error) => {
      console.error('Login Failed:', error);
      setError('Login failed. Please make sure you are added as a test user in Google Cloud Console.');
    },
    scope: 'openid profile email https://www.googleapis.com/auth/spreadsheets.readonly',
  });

  // Fetch user profile
  const fetchUserProfile = async (token: string) => {
    try {
      const response = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
        headers: { Authorization: `Bearer ${token}` },
      });
      const profile = await response.json();
      setUser({
        id: profile.id,
        name: profile.name,
        email: profile.email,
        picture: profile.picture,
      });
    } catch (error) {
      console.error('Error fetching user profile:', error);
      setError('Failed to fetch user profile.');
    }
  };

  // Extract spreadsheet ID from URL
  const extractSpreadsheetId = (url: string): string => {
    const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    return match ? match[1] : '';
  };

  // Fetch spreadsheet metadata to get available sheets
  const fetchSpreadsheetMetadata = useCallback(async (id: string) => {
    if (!accessToken) {
      setError('No access token available. Please login again.');
      return;
    }

    try {
      setLoading(true);
      setError('');
      
      const response = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${id}?fields=sheets.properties`,
        {
          headers: { 
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json',
          },
        }
      );

      if (!response.ok) {
        const errorText = await response.text();
        console.error('API Error Response:', errorText);
        
        if (response.status === 403) {
          setError('Permission denied. Please make sure you have access to this spreadsheet and the Google Sheets API is enabled.');
        } else if (response.status === 404) {
          setError('Spreadsheet not found. Please check the URL and try again.');
        } else {
          setError(`Failed to fetch spreadsheet information: ${response.status} ${response.statusText}`);
        }
        return;
      }

      const metadata: SpreadsheetMetadata = await response.json();
      
      if (!metadata.sheets || metadata.sheets.length === 0) {
        setError('No sheets found in this spreadsheet.');
        return;
      }
      
      setAvailableSheets(metadata.sheets);
      setSelectedSheet(''); // Reset selection
    } catch (error) {
      console.error('Error fetching spreadsheet metadata:', error);
      if (error instanceof TypeError && error.message.includes('fetch')) {
        setError('Network error. Please check your internet connection and try again.');
      } else {
        setError(`Failed to fetch spreadsheet information: ${error instanceof Error ? error.message : 'Unknown error'}`);
      }
    } finally {
      setLoading(false);
    }
  }, [accessToken]);

  // Fetch sheet data
  const fetchSheetData = useCallback(async (id: string, sheetName: string) => {
    if (!accessToken || !sheetName) return;

    try {
      setLoading(true);
      setError('');
      
      const response = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${encodeURIComponent(sheetName)}`,
        {
          headers: { Authorization: `Bearer ${accessToken}` },
        }
      );

      if (!response.ok) {
        const errorText = await response.text();
        console.error('API Error Response:', errorText);
        throw new Error(`Failed to fetch sheet data: ${response.status} ${response.statusText}`);
      }

      const data: SheetData = await response.json();
      setSheetData(data);
    } catch (error) {
      console.error('Error fetching sheet data:', error);
      setError(`Failed to fetch sheet data: ${error instanceof Error ? error.message : 'Unknown error'}`);
    } finally {
      setLoading(false);
    }
  }, [accessToken]);

  // 客戶分析功能
  const performCustomerAnalysis = useCallback(() => {
    if (!sheetData?.values || sheetData.values.length < 2) return;

    // 解析金額字串（移除非數字與小數點符號，例如 NT$152,000.00）
    const parseAmount = (value: string | undefined): number => {
      if (!value) return NaN;
      const cleaned = value.toString().replace(/[^0-9.]/g, '');
      return parseFloat(cleaned);
    };

    // 將可能含逗號、文字的數字清理並轉成 number，失敗回傳 NaN
    const parseNumber = (value: string | undefined): number => {
      if (!value) return NaN;
      const cleaned = value.toString().replace(/[^0-9.]/g, '');
      return parseFloat(cleaned);
    };

    // 取得客戶物件中的左 / 右耳 PTA 數值
    const getEarPTA = (customer: { [key: string]: string }, earLabel: '左' | '右'): number => {
      const key = Object.keys(customer).find(k => k.includes(`${earLabel}耳`) && k.toUpperCase().includes('PTA'));
      if (!key) return NaN;
      return parseNumber(customer[key]);
    };

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
      setError('找不到日期相關欄位，請確認資料表包含初次到店日期');
      return;
    }

    const monthlyData: { [key: string]: {
      month: string;
      year: number;
      newCustomers: number;
      completedDeals: number;
      conversionRate: number;
      totalAmount: number;
      averageAmount: number;
    } } = {};
    
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

      // 原有的月份數據將被customersArray分析覆蓋，這裡只記錄初始結構

      if (!earliestDate || date < earliestDate) earliestDate = date;
      if (!latestDate || date > latestDate) latestDate = date;
    }

    // 舊的月份分析邏輯被移除，現在完全基於customersArray進行分析

    // 輸出篩選結果到控制台

    
          if (filteredRows.length > 0) {
        // 將符合條件的資料整合成物件陣列
        const customersArray = filteredRows.map((row, index) => {
          const customerObject: { [key: string]: string } = {};
          
          // 將每個表頭作為key，對應的資料作為value
          headers.forEach((header: string) => {
            const cellData = row.所有欄位數據[header];
            customerObject[header] = cellData ? cellData.內容 : '';
          });
          
          return customerObject;
        });
        
        // 輸出完整陣列
        console.log(customersArray);
      
      
     
           // 業務員分析
      console.log('\n=== 業務員銷售分析 ===');
      const salesmenAnalysis: { [key: string]: {
        業務員: string,
        訂單數量: number,
        潛力客戶數: number,
        成交率: number,
        當季業績累積: number
      } } = {};
      
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

      // 分析每個業務員的數據
      customersArray.forEach(customer => {
        const salesman = customer['主聽力師'] || customer['聽力師'] || '未知業務員';
        const leftEarPTA = getEarPTA(customer, '左');
        const rightEarPTA = getEarPTA(customer, '右');
        
        // 判斷是否成交（主要依靠"是否成交"欄位）先行計算，供潛力判斷使用
        const isDealt = customer['是否成交'] === '是' || customer['是否成交'] === 'TRUE' || 
                       customer['是否借機'] === 'TRUE' || customer['是否有借機'] === 'TRUE' ||
                       customer['是否借機'] === '是' ||
                       customer['成交'] === '是' || customer['成交'] === 'TRUE' ||
                       customer['狀態'] === '成交' || customer['狀態'] === '已成交';

        // 判斷是否為潛力客戶：左/右 PTA 高於閾值 **或** 已成交
        const isPotentialCustomer = leftEarPTA > ptaThreshold || rightEarPTA > ptaThreshold || isDealt;
        
        // 如果是潛力客戶
        if (isPotentialCustomer) {
          salesmenAnalysis[salesman].潛力客戶數++;
          totalPotentialCustomers++;
        }
        
        // 如果有成交
        if (isDealt) {
          salesmenAnalysis[salesman].訂單數量++;
          totalOrders++;
        }
      });

      // 計算每個業務員的當季累積業績（成交客戶的金額加總）
      Object.keys(salesmenAnalysis).forEach(salesman => {
        // 篩選出該業務員的所有客戶
        const salesmanCustomers = customersArray.filter(customer => 
          (customer['主聽力師'] || customer['聽力師'] || '未知業務員') === salesman
        );
        
        // 從該業務員的客戶中，篩選出成交的客戶並加總金額
        let totalAmount = 0;
        salesmanCustomers.forEach(customer => {
          const isDealt = customer['是否成交'] === '是' || customer['是否成交'] === 'TRUE' ||
                          customer['是否借機'] === 'TRUE' || customer['是否有借機'] === 'TRUE' ||
                          customer['是否借機'] === '是' ||
                          customer['成交'] === '是' || customer['成交'] === 'TRUE' ||
                          customer['狀態'] === '成交' || customer['狀態'] === '已成交';

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
      

      
      // 更新customerAnalysis以反映正確的數據
      setCustomerAnalysis(prev => ({
        ...prev!,
        totalCustomers: totalPotentialCustomers,
        totalCompletedDeals: totalOrders,
        totalAmount: totalRevenue,
        overallConversionRate: overallConversionRate
      }));
     
           // 儲存業務員分析數據供界面使用
      // 生成月份統計數據 (基於customersArray)
      const customerMonthlyData: { [key: string]: {
        month: string;
        year: number;
        newCustomers: number;
        completedDeals: number;
        conversionRate: number;
        totalAmount: number;
        averageAmount: number;
      } } = {};
      
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
        const isDealtClinic = customer['是否成交'] === '是' || customer['是否成交'] === 'TRUE' ||
                              customer['是否借機'] === 'TRUE' || customer['成交'] === '是' ||
                              customer['成交'] === 'TRUE' || customer['狀態'] === '成交' || customer['狀態'] === '已成交';
        const isPotential = leftEarPTA > ptaThreshold || rightEarPTA > ptaThreshold || isDealtClinic;
        
        // 成交判斷
        const isDealt = customer['是否成交'] === '是' || customer['是否成交'] === 'TRUE' || 
                       customer['是否借機'] === 'TRUE' || customer['是否有借機'] === 'TRUE' ||
                       customer['是否借機'] === 'TRUE' || customer['是否借機'] === '是';
        
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
      
      // 儲存業務員分析數據
      setSalesmenAnalysis(salesmenAnalysis);
      
      /* ====================== 診所轉介分析 ====================== */
      const clinicSummary: { [key: string]: {
        clinic: string;
        total: number;
        potential: number;
        dealt: number;
        conversionRate: number;
        totalAmount: number;
      }} = {};

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

        const isDealt = customer['是否成交'] === '是' || customer['是否成交'] === 'TRUE' ||
                       customer['是否借機'] === 'TRUE' || customer['成交'] === '是' ||
                       customer['成交'] === 'TRUE' || customer['狀態'] === '成交' || customer['狀態'] === '已成交';

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
      const clinicArray = Object.values(clinicSummary).sort((a,b)=>b.total - a.total);
      setClinicAnalysis(clinicArray);
       
      // 更新客戶分析數據
      setCustomerAnalysis({
        monthlyData: sortedMonthlyData,
        totalCustomers: totalPotentialCustomers,
        totalCompletedDeals: totalOrders,
        totalAmount: totalRevenue,
        overallConversionRate: overallConversionRate,
        dateRange: {
          earliest: customersArray.length > 0 ? customersArray[0]['服務日期'] || customersArray[0]['初次到店'] || '' : '',
          latest: customersArray.length > 0 ? customersArray[customersArray.length - 1]['服務日期'] || customersArray[customersArray.length - 1]['初次到店'] || '' : ''
        }
      });

      /* ====================== 門市轉介分析 ====================== */
      const storeSummary: { [key:string]: { store:string; total:number; potential:number; } } = {};

      customersArray.forEach(customer=>{
        // 可能的門市欄位名稱
        const storeKey = Object.keys(customer).find(k=>k.includes('門市') && k.includes('自帶'));
        if(!storeKey) return;
        const storeNameRaw = (customer[storeKey] || '').trim();
        if(!storeNameRaw || storeNameRaw === '#N/A') return;

        if(!storeSummary[storeNameRaw]){
          storeSummary[storeNameRaw] = { store: storeNameRaw, total:0, potential:0 };
        }
        storeSummary[storeNameRaw].total++;

        const leftPTA = getEarPTA(customer,'左');
        const rightPTA = getEarPTA(customer,'右');
        const isDealtStore = customer['是否成交'] === '是' || customer['是否成交'] === 'TRUE' ||
                             customer['是否借機'] === 'TRUE' || customer['成交'] === '是' ||
                             customer['成交'] === 'TRUE' || customer['狀態'] === '成交' || customer['狀態'] === '已成交';
        if(leftPTA > ptaThreshold || rightPTA > ptaThreshold || isDealtStore){
          storeSummary[storeNameRaw].potential++;
        }
      });

      const storeArray = Object.values(storeSummary).sort((a,b)=>b.total - a.total);
      setStoreReferralAnalysis(storeArray);

      /* ====================== 聽篩活動來源（按月份）分析 ====================== */
      const hearingSummary: { [key:string]: { month:string; year:number; total:number; potential:number; dealt:number; conversionRate:number; totalAmount:number; } } = {};

      customersArray.forEach(customer => {
        const sourceVal = (customer['顧客來源'] || customer['顧客來源↵(可複選)'] || '').toString();
        if (!sourceVal.includes('聽篩')) return; // 僅統計含「聽篩」字樣

        const serviceDate = customer['服務日期'] || customer['初次到店'] || '';
        if (!serviceDate) return;
        const date = new Date(serviceDate);
        if (isNaN(date.getTime())) return;

        const monthKey = `${date.getFullYear()}年${date.getMonth() + 1}月`;
        if (!hearingSummary[monthKey]) {
          hearingSummary[monthKey] = { month: monthKey, year: date.getFullYear(), total:0, potential:0, dealt:0, conversionRate:0, totalAmount:0 };
        }

        hearingSummary[monthKey].total++;

        // PTA & 成交
        const leftPTA = getEarPTA(customer,'左');
        const rightPTA = getEarPTA(customer,'右');
        const isDealtHS = customer['是否成交'] === '是' || customer['是否成交'] === 'TRUE' ||
                          customer['是否借機'] === 'TRUE' || customer['是否有借機'] === 'TRUE' ||
                          customer['是否借機'] === '是' || customer['成交'] === '是' || customer['成交'] === 'TRUE' ||
                          customer['狀態'] === '成交' || customer['狀態'] === '已成交';

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
      });

      // 計算成交率並排序
      Object.values(hearingSummary).forEach(item => {
        item.conversionRate = item.total > 0 ? (item.dealt / item.total) * 100 : 0;
      });

      const hearingArray = Object.values(hearingSummary).sort((a,b)=> {
        if (a.year !== b.year) return a.year - b.year;
        const aMonth = parseInt(a.month.replace(/\d+年(\d+)月/,'$1'));
        const bMonth = parseInt(b.month.replace(/\d+年(\d+)月/,'$1'));
        return aMonth - bMonth;
      });

      setHearingScreeningAnalysis(hearingArray);
    }
  }, [sheetData, dateRange, ptaThreshold]);

  // Handle sheet URL input
  const handleSheetUrlSubmit = useCallback((e: React.FormEvent) => {
    e.preventDefault();
    const id = extractSpreadsheetId(sheetUrl);
    if (id) {
      setSpreadsheetId(id);
      fetchSpreadsheetMetadata(id);
    } else {
      setError('Invalid Google Sheets URL. Please check the URL and try again.');
    }
  }, [sheetUrl, fetchSpreadsheetMetadata]);

  // Handle sheet selection change
  const handleSheetChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const newSheet = e.target.value;
    setSelectedSheet(newSheet);
    if (spreadsheetId && newSheet) {
      fetchSheetData(spreadsheetId, newSheet);
    }
  };

  // 生成分析報告
  const generateAnalysis = () => {
    if (selectedSheet && sheetData?.values) {
      performCustomerAnalysis();
    }
  };

  // 生成報表標題
  const generateReportTitle = () => {
    const startYear = dateRange.startYear;
    const startMonth = dateRange.startMonth;
    const endYear = dateRange.endYear;
    const endMonth = dateRange.endMonth;
    
    return {
      dateRange: `${startYear}年${startMonth}月~${endYear}年${endMonth}月`,
      storeAnalysis: `${selectedStore}數據分析`
    };
  };

  // Logout function
  const handleLogout = () => {
    googleLogout();
    setUser(null);
    setAccessToken('');
    setSheetData(null);
    setAvailableSheets([]);
    setSelectedSheet('');
    setSpreadsheetId('');
    setSheetUrl('');
    setCustomerAnalysis(null);
    setSalesmenAnalysis(null);
    setClinicAnalysis(null);
    setStoreReferralAnalysis(null);
    setSelectedStore('桃園藝文店');
    setError('');
  };

  // 月份報告圖表配置
  const monthlyChartData = customerAnalysis ? {
    labels: customerAnalysis.monthlyData.map(item => item.month),
    datasets: [
      {
        label: '來客數',
        data: customerAnalysis.monthlyData.map(item => item.newCustomers),
        backgroundColor: 'rgba(34, 197, 94, 0.5)',
        borderColor: 'rgba(34, 197, 94, 1)',
        borderWidth: 2,
      },
      {
        label: '成交數',
        data: customerAnalysis.monthlyData.map(item => item.completedDeals),
        backgroundColor: 'rgba(59, 130, 246, 0.5)',
        borderColor: 'rgba(59, 130, 246, 1)',
        borderWidth: 2,
      },
    ],
  } : null;

  const monthlyChartOptions = {
    responsive: true,
    plugins: {
      legend: {
        position: 'top' as const,
      },
      title: {
        display: true,
        text: '客戶分析報告',
      },
    },
    scales: {
      y: {
        beginAtZero: true,
        title: {
          display: true,
          text: '數量',
        },
      },
      x: {
        title: {
          display: true,
          text: '月份',
        },
      },
    },
  };

  // 診所轉介圖表資料
  const clinicChartData = clinicAnalysis ? {
    labels: clinicAnalysis.map((c: any) => c.clinic),
    datasets: [
      {
        label: '轉介數',
        data: clinicAnalysis.map((c: any) => c.total),
        backgroundColor: 'rgba(59,130,246,0.6)',
        borderColor: 'rgba(59,130,246,1)',
        borderWidth: 1,
      },
      {
        label: '潛力客戶',
        data: clinicAnalysis.map((c: any) => c.potential),
        backgroundColor: 'rgba(34,197,94,0.6)',
        borderColor: 'rgba(34,197,94,1)',
        borderWidth: 1,
      },
    ],
  } : null;

  const clinicChartOptions = {
    responsive: true,
    plugins: {
      legend: { position: 'top' as const },
    },
    scales: {
      y: { beginAtZero: true },
    },
  };

  // 門市轉介圖表資料
  const storeChartData = storeReferralAnalysis ? {
    labels: storeReferralAnalysis.map((s:any)=>s.store),
    datasets:[
      {
        label:'轉介數',
        data: storeReferralAnalysis.map((s:any)=>s.total),
        backgroundColor:'rgba(59,130,246,0.6)',
        borderColor:'rgba(59,130,246,1)',
        borderWidth:1,
      },
      {
        label:'潛力客戶',
        data: storeReferralAnalysis.map((s:any)=>s.potential),
        backgroundColor:'rgba(34,197,94,0.6)',
        borderColor:'rgba(34,197,94,1)',
        borderWidth:1,
      },
    ],
  }:null;

  const storeChartOptions = {
    responsive:true,
    plugins:{ legend:{ position:'top' as const }},
    scales:{ y:{ beginAtZero:true }},
  };

  // 聽篩活動圖表資料
  const hearingChartData = hearingScreeningAnalysis ? {
    labels: hearingScreeningAnalysis.map((h:any)=>h.month),
    datasets: [
      {
        label: '來客數',
        data: hearingScreeningAnalysis.map((h:any)=>h.total),
        backgroundColor: 'rgba(59,130,246,0.6)',
        borderColor: 'rgba(59,130,246,1)',
        borderWidth: 1,
      },
      {
        label: '潛力客戶',
        data: hearingScreeningAnalysis.map((h:any)=>h.potential),
        backgroundColor: 'rgba(34,197,94,0.6)',
        borderColor: 'rgba(34,197,94,1)',
        borderWidth: 1,
      },
      {
        label: '成交數',
        data: hearingScreeningAnalysis.map((h:any)=>h.dealt),
        backgroundColor: 'rgba(249,115,22,0.6)',
        borderColor: 'rgba(249,115,22,1)',
        borderWidth: 1,
      },
    ],
  }: null;

  const hearingChartOptions = {
    responsive:true,
    plugins:{ legend:{ position:'top' as const }},
    scales:{ y:{ beginAtZero:true }},
  };

  return (
    <div className="min-h-screen bg-gray-100 py-8">
      <div className="max-w-6xl mx-auto px-4">
        <h1 className="text-3xl font-bold text-center mb-8 text-gray-800 no-print">
          客戶分析統計系統
        </h1>

        {!user ? (
          <div className="card text-center">
            <h2 className="text-xl font-semibold mb-4">請登入Google帳號</h2>
            <p className="text-gray-600 mb-6">
              請先登入以存取您的Google試算表資料
            </p>
            <button
              onClick={() => login()}
              className="bg-blue-500 hover:bg-blue-600 text-white font-medium py-2 px-4 rounded-lg transition-colors"
            >
              使用Google登入
            </button>
          </div>
        ) : (
          <div className="space-y-6">
            {/* User Profile */}
            <div className="card before-report">
              <div className="flex items-center justify-between">
                <div className="flex items-center space-x-4">
                  <img
                    src={user.picture}
                    alt={user.name}
                    className="w-10 h-10 rounded-full"
                  />
                  <div>
                    <h3 className="font-semibold">{user.name}</h3>
                    <p className="text-gray-600 text-sm">{user.email}</p>
                  </div>
                </div>
                <button
                  onClick={handleLogout}
                  className="bg-red-500 hover:bg-red-600 text-white font-medium py-2 px-4 rounded-lg transition-colors"
                >
                  登出
                </button>
              </div>
            </div>

            {/* Sheet URL Input */}
            <div className="card before-report">
              <h3 className="text-lg font-semibold mb-4">輸入Google試算表網址</h3>
              <form onSubmit={handleSheetUrlSubmit} className="space-y-4">
                <input
                  type="url"
                  value={sheetUrl}
                  onChange={(e) => setSheetUrl(e.target.value)}
                  placeholder="https://docs.google.com/spreadsheets/d/..."
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                  required
                />
                <button
                  type="submit"
                  disabled={loading}
                  className="bg-green-500 hover:bg-green-600 text-white font-medium py-2 px-4 rounded-lg transition-colors disabled:opacity-50"
                >
                  {loading ? '載入中...' : '載入試算表'}
                </button>
              </form>
            </div>

            {/* Sheet Selector */}
            {availableSheets.length > 0 && (
              <div className="card before-report">
                <h3 className="text-lg font-semibold mb-4">選擇工作表</h3>
                <select
                  value={selectedSheet}
                  onChange={handleSheetChange}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                >
                  <option value="">請選擇工作表...</option>
                  {availableSheets.map((sheet) => (
                    <option key={sheet.properties.sheetId} value={sheet.properties.title}>
                      {sheet.properties.title}
                    </option>
                  ))}
                </select>
              </div>
            )}

            {/* Date Range Selector */}
            {sheetData && (
              <div className="card before-report">
                <h3 className="text-lg font-semibold mb-4">選擇分析月份區間</h3>
                <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">開始年份</label>
                    <select
                      value={dateRange.startYear}
                      onChange={(e) => setDateRange({...dateRange, startYear: parseInt(e.target.value)})}
                      className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                    >
                      {Array.from({length: 10}, (_, i) => new Date().getFullYear() - 5 + i).map(year => (
                        <option key={year} value={year}>{year}年</option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">開始月份</label>
                    <select
                      value={dateRange.startMonth}
                      onChange={(e) => setDateRange({...dateRange, startMonth: parseInt(e.target.value)})}
                      className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                    >
                      {Array.from({length: 12}, (_, i) => i + 1).map(month => (
                        <option key={month} value={month}>{month}月</option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">結束年份</label>
                    <select
                      value={dateRange.endYear}
                      onChange={(e) => setDateRange({...dateRange, endYear: parseInt(e.target.value)})}
                      className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                    >
                      {Array.from({length: 10}, (_, i) => new Date().getFullYear() - 5 + i).map(year => (
                        <option key={year} value={year}>{year}年</option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">結束月份</label>
                    <select
                      value={dateRange.endMonth}
                      onChange={(e) => setDateRange({...dateRange, endMonth: parseInt(e.target.value)})}
                      className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                    >
                      {Array.from({length: 12}, (_, i) => i + 1).map(month => (
                        <option key={month} value={month}>{month}月</option>
                      ))}
                    </select>
                  </div>
                </div>
                
                {/* PTA閾值選擇器和店家選擇器 */}
                <div className="mt-4 grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-2">
                      PTA閾值設定 (潛力客戶判斷標準)
                    </label>
                    <div className="flex items-center space-x-2">
                      <select
                        value={ptaThreshold}
                        onChange={(e) => setPtaThreshold(parseInt(e.target.value))}
                        className="px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                      >
                        {Array.from({length: 14}, (_, i) => 25 + (i * 5)).map(value => (
                          <option key={value} value={value}>{value} dBHL</option>
                        ))}
                      </select>
                      <span className="text-sm text-gray-600">
                        超過此數值視為潛力客戶
                      </span>
                    </div>
                  </div>
                  
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-2">
                      選擇分析店家
                    </label>
                    <select
                      value={selectedStore}
                      onChange={(e) => setSelectedStore(e.target.value)}
                      className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                    >
                      {storeOptions.map(store => (
                        <option key={store} value={store}>{store}</option>
                      ))}
                    </select>
                  </div>
                </div>
                
                <button
                  onClick={generateAnalysis}
                  disabled={loading || !selectedSheet}
                  className="mt-4 bg-purple-500 hover:bg-purple-600 text-white font-medium py-2 px-4 rounded-lg transition-colors disabled:opacity-50"
                >
                  生成分析報告
                </button>
              </div>
            )}

            {/* Error Message */}
            {error && (
              <div className="bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded-lg before-report">
                {error}
              </div>
            )}

            {/* 報表標題區塊 */}
            {customerAnalysis && (
              <div className="card report-title">
                <div className="text-center">
                  <h1 className="text-2xl font-bold text-gray-800 mb-2">
                    {generateReportTitle().dateRange}
                  </h1>
                  <h2 className="text-xl font-semibold text-blue-600">
                    {generateReportTitle().storeAnalysis}
                  </h2>
                </div>
              </div>
            )}

            {/* 銷售分析報告 */}
            {customerAnalysis && (
              <div className="print-area">
                <div className="card mb-6">
                  <h3 className="text-xl font-bold mb-6 text-center bg-indigo-100 py-3 rounded-lg">
                    店銷售分析（潛力分貝數：{ptaThreshold} dB HL）
                  </h3>
                  
                  <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                    {/* 左側：業務員業績圖表 */}
                    <div className="lg:col-span-2">
                      <div className="card-light">
                        <h4 className="text-lg font-semibold mb-4">業績員圖表</h4>
                        {/* 業務員業績圖表 */}
                        <div className="h-64 flex items-end space-x-6 bg-white p-4 rounded">
                          {salesmenAnalysis && (() => {
                            const salesmenArr = Object.values(salesmenAnalysis) as any[];
                            const maxOrder = Math.max(...salesmenArr.map(s => s.訂單數量));
                            const maxAmount = Math.max(...salesmenArr.map(s => s.當季業績累積));
                            return salesmenArr.map((salesman, index) => {
                              const orderHeight = maxOrder > 0 ? Math.max(32, (salesman.訂單數量 / maxOrder) * 160) : 32;
                              const amountHeight = maxAmount > 0 ? Math.max(32, (salesman.當季業績累積 / maxAmount) * 160) : 32;
                              return (
                                <div key={index} className="flex flex-col items-center space-y-2">
                                  {/* 條形組 */}
                                  <div className="flex space-x-2 items-end">
                                    {/* 訂單數量條 */}
                                    <div
                                      className="bg-blue-500 w-6 flex items-end justify-center text-white text-[10px] font-bold pb-1"
                                      style={{ height: `${orderHeight}px` }}
                                    >
                                      {salesman.訂單數量}
                                    </div>
                                    {/* 累積金額條 */}
                                    <div
                                      className="bg-green-500 w-6 flex items-end justify-center text-white text-[10px] font-bold pb-1"
                                      style={{ height: `${amountHeight}px` }}
                                    >
                                      {Math.round(salesman.當季業績累積 / 1000)}k
                                    </div>
                                  </div>
                                  <span className="text-sm whitespace-nowrap max-w-[60px] text-center">{salesman.業務員}</span>
                                </div>
                              );
                            });
                          })()}
                        </div>
                        <div className="flex justify-center mt-2 space-x-4">
                          <div className="flex items-center">
                            <div className="w-4 h-4 bg-blue-500 mr-2"></div>
                            <span className="text-sm">訂單數量</span>
                          </div>
                          <div className="flex items-center">
                            <div className="w-4 h-4 bg-green-500 mr-2"></div>
                            <span className="text-sm">累積金額 (k)</span>
                          </div>
                        </div>
                      </div>
                    </div>

                    {/* 右側：統計數據 */}
                    <div className="space-y-4">
                      {/* 業務數據 */}
                      <div className="card-light">
                        <h4 className="text-base font-semibold mb-3 text-blue-800">業務數據</h4>
                        <div className="grid grid-cols-2 gap-2 text-sm">
                          <div>
                            <div className="text-gray-600">潛力來客數</div>
                            <div className="font-bold text-lg text-blue-600">{customerAnalysis.totalCustomers}</div>
                          </div>
                          <div>
                            <div className="text-gray-600">訂單數</div>
                            <div className="font-bold text-lg text-blue-600">{customerAnalysis.totalCompletedDeals}</div>
                          </div>
                          <div>
                            <div className="text-gray-600">營業額</div>
                            <div className="font-bold text-lg text-blue-600">{customerAnalysis.totalAmount.toLocaleString()}</div>
                          </div>
                          <div>
                            <div className="text-gray-600">客單價</div>
                            <div className="font-bold text-lg text-blue-600">
                              {customerAnalysis.totalCompletedDeals > 0 
                                ? `NT$ ${(customerAnalysis.totalAmount / customerAnalysis.totalCompletedDeals).toFixed(1).replace(/\B(?=(\d{3})+(?!\d))/g, ',')}` 
                                : 'NT$ 0'}
                            </div>
                          </div>
                          <div>
                            <div className="text-gray-600">成交率</div>
                            <div className="font-bold text-lg text-blue-600">{customerAnalysis.overallConversionRate.toFixed(1)}%</div>
                          </div>
                        </div>
                      </div>

                      {/* 下排：業務員訂單分析區塊 */}
                    </div>
                  </div>
                </div>
              </div>
            )}

            {/* 業務員訂單分析區塊 */}
            {salesmenAnalysis && (
            <div className="mt-8">
              <div className="card-light">
                <h4 className="text-base font-semibold mb-3">業務員訂單分析</h4>
                <div className="overflow-x-auto">
                  <table className="w-full text-xs">
                    <thead>
                      <tr className="bg-gray-200">
                        <th className="p-2 text-left">業務</th>
                        <th className="p-2 text-center">訂單數量</th>
                        <th className="p-2 text-center">潛力客戶</th>
                        <th className="p-2 text-center">成交率</th>
                        <th className="p-2 text-center">期間業績累積</th>
                      </tr>
                    </thead>
                    <tbody className="text-xs">
                      {salesmenAnalysis && Object.values(salesmenAnalysis).map((salesman: any, index: number) => (
                        <tr key={index} className="border-b">
                          <td className="p-2">{salesman.業務員}</td>
                          <td className="p-2 text-center">{salesman.訂單數量}</td>
                          <td className="p-2 text-center">{salesman.潛力客戶數}</td>
                          <td className="p-2 text-center">
                            {salesman.潛力客戶數 > 0 ? `${salesman.成交率.toFixed(1)}%` : '#DIV/0!'}
                          </td>
                          <td className="p-2 text-center">{salesman.當季業績累積.toLocaleString()}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
            )}

            {/* 客戶分析報告 */}
            {customerAnalysis && (
              <div className="card">
                <h3 className="text-lg font-semibold mb-4">客戶分析報告</h3>
                
                {/* 基本統計 */}
                <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-6">
                  <div className="card-light">
                    <div className="text-2xl font-bold text-green-600">{customerAnalysis.totalCustomers}</div>
                    <div className="text-sm text-gray-600">總來客數</div>
                  </div>
                  <div className="card-light">
                    <div className="text-2xl font-bold text-blue-600">{customerAnalysis.totalCompletedDeals}</div>
                    <div className="text-sm text-gray-600">總成交數</div>
                  </div>
                  <div className="card-light">
                    <div className="text-2xl font-bold text-purple-600">{customerAnalysis.overallConversionRate.toFixed(1)}%</div>
                    <div className="text-sm text-gray-600">整體成交率</div>
                  </div>
                  <div className="card-light">
                    <div className="text-2xl font-bold text-yellow-600">NT$ {customerAnalysis.totalAmount.toLocaleString()}</div>
                    <div className="text-sm text-gray-600">總成交金額</div>
                  </div>
                </div>

                {/* 月份圖表 */}
                {monthlyChartData && (
                  <div className="mb-6">
                    <div className="chart-wrapper h-80 w-full">
                      <Bar data={monthlyChartData} options={monthlyChartOptions} />
                    </div>
                  </div>
                )}

                {/* 月份詳細數據表 */}
                <div className="overflow-x-auto">
                  <table className="min-w-full table-auto">
                    <thead>
                      <tr className="bg-gray-50">
                        <th className="px-4 py-2 text-left text-sm font-medium text-gray-700 border-b">月份</th>
                        <th className="px-4 py-2 text-left text-sm font-medium text-gray-700 border-b">來客數</th>
                        <th className="px-4 py-2 text-left text-sm font-medium text-gray-700 border-b">成交數</th>
                        <th className="px-4 py-2 text-left text-sm font-medium text-gray-700 border-b">成交率</th>
                        <th className="px-4 py-2 text-left text-sm font-medium text-gray-700 border-b">總金額</th>
                        <th className="px-4 py-2 text-left text-sm font-medium text-gray-700 border-b">客單價</th>
                      </tr>
                    </thead>
                    <tbody>
                      {customerAnalysis.monthlyData.map((item, index) => (
                        <tr key={index} className="hover:bg-gray-50">
                          <td className="px-4 py-2 text-sm text-gray-900 border-b">{item.month}</td>
                          <td className="px-4 py-2 text-sm text-gray-900 border-b">{item.newCustomers}</td>
                          <td className="px-4 py-2 text-sm text-gray-900 border-b">{item.completedDeals}</td>
                          <td className="px-4 py-2 text-sm text-gray-900 border-b">{item.conversionRate.toFixed(1)}%</td>
                          <td className="px-4 py-2 text-sm text-gray-900 border-b">NT$ {item.totalAmount.toLocaleString()}</td>
                          <td className="px-4 py-2 text-sm text-gray-900 border-b">NT$ {item.averageAmount.toLocaleString()}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* 診所轉介分析 */}
            {clinicAnalysis && clinicAnalysis.length > 0 && (
              <div className="card mt-8">
                <h3 className="text-lg font-semibold mb-4">診所與其它轉介分析</h3>
                {clinicChartData && (
                  <div className="chart-wrapper h-80 mb-6 w-full">
                    <Bar data={clinicChartData} options={clinicChartOptions} plugins={[barLabelPlugin]} />
                  </div>
                )}
                <div className="overflow-x-auto">
                  <table className="w-full text-xs">
                    <thead>
                      <tr className="bg-gray-200">
                        <th className="p-2 text-left">診所</th>
                        <th className="p-2 text-center">轉介數</th>
                        <th className="p-2 text-center">潛力客戶</th>
                        <th className="p-2 text-center">成交數</th>
                        <th className="p-2 text-center">成交率</th>
                        <th className="p-2 text-center">累積金額</th>
                      </tr>
                    </thead>
                    <tbody>
                      {clinicAnalysis.map((c:any,idx:number)=>(
                        <tr key={idx} className="border-b">
                          <td className="p-2">{c.clinic}</td>
                          <td className="p-2 text-center">{c.total}</td>
                          <td className="p-2 text-center">{c.potential}</td>
                          <td className="p-2 text-center">{c.dealt}</td>
                          <td className="p-2 text-center">{c.conversionRate.toFixed(1)}%</td>
                          <td className="p-2 text-center">{c.totalAmount.toLocaleString()}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* 門市轉介分析 */}
            {storeReferralAnalysis && storeReferralAnalysis.length > 0 && (
              <div className="card mt-8">
                <h3 className="text-lg font-semibold mb-4">門市轉介分析</h3>
                {storeChartData && (
                  <div className="chart-wrapper h-80 mb-6 w-full">
                    <Bar data={storeChartData} options={storeChartOptions} plugins={[barLabelPlugin]} />
                  </div>
                )}
                <div className="overflow-x-auto">
                  <table className="w-full text-xs">
                    <thead>
                      <tr className="bg-gray-200">
                        <th className="p-2 text-left">門市</th>
                        <th className="p-2 text-center">轉介數</th>
                        <th className="p-2 text-center">潛力客戶</th>
                      </tr>
                    </thead>
                    <tbody>
                      {storeReferralAnalysis.map((s:any,idx:number)=>(
                        <tr key={idx} className="border-b">
                          <td className="p-2">{s.store}</td>
                          <td className="p-2 text-center">{s.total}</td>
                          <td className="p-2 text-center">{s.potential}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* 聽篩活動分析 */}
            {hearingScreeningAnalysis && hearingScreeningAnalysis.length > 0 && (
              <div className="card mt-8">
                <h3 className="text-lg font-semibold mb-4">聽篩活動來源分析</h3>
                {hearingChartData && (
                  <div className="chart-wrapper h-80 mb-6 w-full">
                    <Bar data={hearingChartData} options={hearingChartOptions} plugins={[barLabelPlugin]} />
                  </div>
                )}
                <div className="overflow-x-auto">
                  <table className="w-full text-xs">
                    <thead>
                      <tr className="bg-gray-200">
                        <th className="p-2 text-left">來源</th>
                        <th className="p-2 text-center">來客數</th>
                        <th className="p-2 text-center">潛力客戶</th>
                        <th className="p-2 text-center">成交數</th>
                        <th className="p-2 text-center">成交率</th>
                        <th className="p-2 text-center">總金額</th>
                      </tr>
                    </thead>
                    <tbody>
                      {hearingScreeningAnalysis.map((h:any,idx:number)=>(
                        <tr key={idx} className="border-b">
                          <td className="p-2">{h.month}</td>
                          <td className="p-2 text-center">{h.total}</td>
                          <td className="p-2 text-center">{h.potential}</td>
                          <td className="p-2 text-center">{h.dealt}</td>
                          <td className="p-2 text-center">{h.conversionRate.toFixed(1)}%</td>
                          <td className="p-2 text-center">{h.totalAmount.toLocaleString()}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        )}
      </div>
      {/* 列印報告按鈕（固定右下，列印時隱藏） */}
      {customerAnalysis && (
        <button
          onClick={() => window.print()}
          className="no-print fixed bottom-4 right-4 bg-purple-600 hover:bg-purple-700 text-white font-medium px-4 py-2 rounded-lg shadow-lg"
        >
          列印報告
        </button>
      )}
    </div>
  );
};

export default App; 