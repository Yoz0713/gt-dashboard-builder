import React, { useState, useEffect } from 'react';
import { useGoogleSheetData } from './hooks/useGoogleSheetData';
import { LoginView } from './components/LoginView';
import { Controls } from './components/Controls';
import { DashboardStats } from './components/DashboardStats';
import { AnalysisCharts } from './components/AnalysisCharts';
import { ComparisonView } from './components/ComparisonView';
import { CompetitionRanking } from './components/CompetitionRanking';
import { ClinicFollowUp } from './components/ClinicFollowUp';
import { DateRange } from './types';

const App: React.FC = () => {
  const {
    user,
    availableSheets,
    selectedSheet,
    loadSheetData,
    loading,
    error,
    analysisResult,
    login,
    logout,
    loadSpreadsheetMetadata,
    loadSavedSpreadsheet,
    savedSpreadsheets,
    performAnalysis,
    spreadsheetTitle,
    sheetData,
    selectedCompetitionSpreadsheetIds,
    competitionRankingEntries,
    competitionSkippedSpreadsheets,
    competitionLoading,
    competitionError,
    competitionSelectionMessage,
    competitionSelectionLimit,
    toggleCompetitionSpreadsheetSelection,
    loadCompetitionRanking,
  } = useGoogleSheetData();

  // UI Local State for controls
  const [sheetUrl, setSheetUrl] = useState<string>('');
  const [dateRange, setDateRange] = useState<DateRange>({
    startYear: new Date().getFullYear(),
    startMonth: 1,
    endYear: new Date().getFullYear(),
    endMonth: 12,
  });
  const [ptaThreshold, setPtaThreshold] = useState(40);
  const [activeTab, setActiveTab] = useState<'overview' | 'comparison' | 'competition' | 'clinic'>('overview');

  // Effect handlers
  const handleSheetUrlSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const match = sheetUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    const id = match ? match[1] : '';
    if (id) {
      loadSpreadsheetMetadata(id);
    } else {
      console.error('Invalid URL');
    }
  };

  const handleSheetChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const newSheet = e.target.value;
    const match = sheetUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    const id = match ? match[1] : '';
    if (id && newSheet) {
      loadSheetData(id, newSheet);
    }
  };

  const handleGenerateAnalysis = () => {
    performAnalysis(dateRange, ptaThreshold);
  };

  // 自動觸發分析: 當 sheetData 載入完成時
  useEffect(() => {
    if (sheetData) {
      performAnalysis(dateRange, ptaThreshold);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sheetData]);

  useEffect(() => {
    if (analysisResult) {
      loadCompetitionRanking(dateRange, ptaThreshold);
    }
  }, [analysisResult, dateRange, ptaThreshold, selectedCompetitionSpreadsheetIds, loadCompetitionRanking]);

  if (!user) {
    return <LoginView onLogin={login} />;
  }

  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-900 print:bg-white">
      <header className="bg-white border-b border-slate-200 px-8 py-4 mb-8 flex justify-between items-center shadow-sm print:hidden">
        <div className="flex items-center gap-4">
          <div className="bg-blue-600 text-white p-2 rounded-lg">
            <svg className="w-6 h-6" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z" />
            </svg>
          </div>
          <h1 className="text-xl font-bold text-slate-800 tracking-tight">Google Sheets Dashboard</h1>
          {user && (
            <div className="flex items-center gap-2 bg-slate-100 px-3 py-1 rounded-full border border-slate-200">
              <span className="text-xs text-slate-500 uppercase font-semibold">User</span>
              <span className="text-sm font-medium text-slate-700">{user.name}</span>
            </div>
          )}
        </div>
        {user && (
          <button
            onClick={logout}
            className="px-4 py-2 text-sm text-slate-600 hover:text-red-600 hover:bg-red-50 rounded-lg transition-colors font-medium"
          >
            Sign out
          </button>
        )}
      </header>

      <main className="max-w-7xl mx-auto px-6 pb-12 print:px-0 print:pb-0 print:max-w-none">
        <Controls
          sheetUrl={sheetUrl}
          setSheetUrl={setSheetUrl}
          handleSheetUrlSubmit={handleSheetUrlSubmit}
          availableSheets={availableSheets}
          selectedSheet={selectedSheet}
          handleSheetChange={handleSheetChange}
          dateRange={dateRange}
          setDateRange={setDateRange}
          spreadsheetTitle={spreadsheetTitle}
          ptaThreshold={ptaThreshold}
          setPtaThreshold={setPtaThreshold}
          onGenerateAnalysis={handleGenerateAnalysis}
          savedSpreadsheets={savedSpreadsheets}
          onSelectSavedSpreadsheet={loadSavedSpreadsheet}
          hasData={!!analysisResult || !!availableSheets.length || !!sheetData}
        />

        {error && (
          <div className="mb-8 p-4 bg-red-50 border-l-4 border-red-500 text-red-700 rounded-r-lg shadow-sm print:hidden">
            <div className="flex items-center gap-2">
              <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
              </svg>
              <p className="font-medium">{error}</p>
            </div>
          </div>
        )}

        {loading && (
          <div className="flex flex-col justify-center items-center h-64 bg-white rounded-xl shadow-sm border border-slate-100">
            <div className="animate-spin rounded-full h-10 w-10 border-b-2 border-blue-600 mb-4"></div>
            <p className="text-slate-500 font-medium">資料分析中，請稍候...</p>
          </div>
        )}

        {!loading && analysisResult && (
          <div className="space-y-8 animate-fade-in print:space-y-4">
            {/* Header & Tabs */}
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center mb-6 gap-4 print:hidden">
              <div className="flex items-center gap-4">
                <h2 className="text-2xl font-bold text-slate-800 tracking-tight">
                  {dateRange.startYear}年{dateRange.startMonth}月 ~ {dateRange.endYear}年{dateRange.endMonth}月
                </h2>
              </div>

              {/* Tab Switcher */}
              <div className="bg-white p-1 rounded-xl shadow-sm border border-slate-200 flex flex-wrap gap-1">
                <button
                  onClick={() => setActiveTab('overview')}
                  className={`px-6 py-2.5 text-sm font-semibold rounded-lg transition-all flex items-center gap-2 ${activeTab === 'overview'
                    ? 'bg-blue-50 text-blue-600 ring-1 ring-blue-200 shadow-sm'
                    : 'text-slate-500 hover:text-slate-700 hover:bg-slate-50'
                    }`}
                >
                  <span>📈</span> 總覽分析
                </button>
                <button
                  onClick={() => setActiveTab('comparison')}
                  className={`px-6 py-2.5 text-sm font-semibold rounded-lg transition-all flex items-center gap-2 ${activeTab === 'comparison'
                    ? 'bg-blue-50 text-blue-600 ring-1 ring-blue-200 shadow-sm'
                    : 'text-slate-500 hover:text-slate-700 hover:bg-slate-50'
                    }`}
                >
                  <span>⚖️</span> 區間比較
                </button>
                <button
                  onClick={() => setActiveTab('competition')}
                  className={`px-6 py-2.5 text-sm font-semibold rounded-lg transition-all flex items-center gap-2 ${activeTab === 'competition'
                    ? 'bg-blue-50 text-blue-600 ring-1 ring-blue-200 shadow-sm'
                    : 'text-slate-500 hover:text-slate-700 hover:bg-slate-50'
                    }`}
                >
                  <span>🏆</span> 競賽排行
                </button>
                <button
                  onClick={() => setActiveTab('clinic')}
                  className={`px-6 py-2.5 text-sm font-semibold rounded-lg transition-all flex items-center gap-2 ${activeTab === 'clinic'
                    ? 'bg-blue-50 text-blue-600 ring-1 ring-blue-200 shadow-sm'
                    : 'text-slate-500 hover:text-slate-700 hover:bg-slate-50'
                    }`}
                >
                  <span>🏥</span> 診所資料回訪分析
                </button>
              </div>

              <button
                onClick={() => window.print()}
                className="bg-slate-900 text-white px-5 py-2.5 rounded-lg hover:bg-slate-800 transition-colors flex items-center gap-2 shadow-lg hover:shadow-xl active:scale-95"
              >
                <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z" />
                </svg>
                匯出 PDF 報告
              </button>
            </div>

            {/* Content Views */}
            <div className="animate-fade-in">
              {activeTab === 'overview' && (
                <>
                  <DashboardStats analysis={analysisResult.customerAnalysis} />
                  <AnalysisCharts analysisResult={analysisResult} />
                </>
              )}

              {activeTab === 'comparison' && (
                <ComparisonView analysis={analysisResult.customerAnalysis} />
              )}

              {activeTab === 'competition' && (
                <CompetitionRanking
                  entries={competitionRankingEntries}
                  loading={competitionLoading}
                  error={competitionError}
                  selectedSpreadsheetCount={selectedCompetitionSpreadsheetIds.length}
                  skippedSpreadsheets={competitionSkippedSpreadsheets}
                  savedSpreadsheets={savedSpreadsheets}
                  selectedCompetitionSpreadsheetIds={selectedCompetitionSpreadsheetIds}
                  onToggleCompetitionSpreadsheet={toggleCompetitionSpreadsheetSelection}
                  competitionSelectionLimit={competitionSelectionLimit}
                  competitionSelectionMessage={competitionSelectionMessage}
                />
              )}

              {activeTab === 'clinic' && (
                <ClinicFollowUp
                  reports={analysisResult.clinicFollowUpAnalysis}
                  dateRange={dateRange}
                  spreadsheetTitle={spreadsheetTitle}
                  ptaThreshold={ptaThreshold}
                />
              )}
            </div>
          </div>
        )}
      </main>
    </div>
  );
};

export default App;
