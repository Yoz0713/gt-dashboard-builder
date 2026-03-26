import { useState, useCallback, useEffect } from 'react';
import { useGoogleLogin, googleLogout } from '@react-oauth/google';
import {
    UserProfile,
    SheetData,
    SheetInfo,
    DateRange,
    SpreadsheetListItem,
    CompetitionRankingEntry,
    CompetitionRankingSkippedSpreadsheet,
} from '../types';
import { fetchUserProfile, fetchSpreadsheetMetadata, fetchSheetData as apiFetchSheetData } from '../services/googleSheets';
import { fetchSavedSpreadsheets, fetchSpreadsheetData, fetchMultipleSpreadsheets } from '../services/sheetsApi';
import { aggregateCompetitionRankings, analyzeData, AnalysisResult } from '../utils/analysis';

const COMPETITION_SELECTION_LIMIT = 10;

export const useGoogleSheetData = () => {
    const [user, setUser] = useState<UserProfile | null>(null);
    const [accessToken, setAccessToken] = useState<string>('');
    const [sheetData, setSheetData] = useState<SheetData | null>(null);
    const [availableSheets, setAvailableSheets] = useState<SheetInfo[]>([]);
    const [selectedSheet, setSelectedSheet] = useState<string>('');
    const [spreadsheetId, setSpreadsheetId] = useState<string>('');
    const [spreadsheetTitle, setSpreadsheetTitle] = useState<string>('');
    const [loading, setLoading] = useState<boolean>(false);
    const [error, setError] = useState<string>('');
    const [analysisResult, setAnalysisResult] = useState<AnalysisResult | null>(null);
    const [savedSpreadsheets, setSavedSpreadsheets] = useState<SpreadsheetListItem[]>([]);
    const [selectedCompetitionSpreadsheetIds, setSelectedCompetitionSpreadsheetIds] = useState<string[]>([]);
    const [competitionRankingEntries, setCompetitionRankingEntries] = useState<CompetitionRankingEntry[]>([]);
    const [competitionSkippedSpreadsheets, setCompetitionSkippedSpreadsheets] = useState<CompetitionRankingSkippedSpreadsheet[]>([]);
    const [competitionLoading, setCompetitionLoading] = useState<boolean>(false);
    const [competitionError, setCompetitionError] = useState<string>('');
    const [competitionSelectionMessage, setCompetitionSelectionMessage] = useState<string>('');

    // Google Login
    const login = useGoogleLogin({
        onSuccess: (response) => {
            console.log('Login successful, access token received');
            setAccessToken(response.access_token);
            handleFetchUserProfile(response.access_token);
        },
        onError: (error) => {
            console.error('Login Failed:', error);
            setError('Login failed. Please make sure you are added as a test user in Google Cloud Console.');
        },
        scope: 'openid profile email https://www.googleapis.com/auth/spreadsheets.readonly',
    });

    const handleFetchUserProfile = async (token: string) => {
        try {
            const profile = await fetchUserProfile(token);
            setUser(profile);
        } catch (err) {
            console.error('Error fetching user profile:', err);
            setError('Failed to fetch user profile.');
        }
    };

    const logout = useCallback(() => {
        googleLogout();
        setUser(null);
        setAccessToken('');
        setSheetData(null);
        setAvailableSheets([]);
        setSelectedSheet('');
        setSpreadsheetId('');
        setSpreadsheetTitle('');
        setAnalysisResult(null);
        setSavedSpreadsheets([]);
        setSelectedCompetitionSpreadsheetIds([]);
        setCompetitionRankingEntries([]);
        setCompetitionSkippedSpreadsheets([]);
        setCompetitionLoading(false);
        setCompetitionError('');
        setCompetitionSelectionMessage('');
        setError('');
    }, []);

    // Moved loadSheetData UP so it can be called by loadSpreadsheetMetadata
    const loadSheetData = useCallback(async (id: string, sheetName: string) => {
        if (!accessToken || !sheetName) return;
        try {
            setLoading(true);
            setError('');
            setSelectedSheet(sheetName); // Update selected sheet state
            const data = await apiFetchSheetData(id, sheetName, accessToken);
            setSheetData(data);
        } catch (err: any) {
            setError(err.message || 'Failed to fetch sheet data');
        } finally {
            setLoading(false);
        }
    }, [accessToken]);

    const loadSpreadsheetMetadata = useCallback(async (id: string) => {
        if (!accessToken) {
            setError('No access token available. Please login again.');
            return;
        }
        try {
            setLoading(true);
            setError('');
            setSpreadsheetId(id);
            const metadata = await fetchSpreadsheetMetadata(id, accessToken);

            if (!metadata.sheets || metadata.sheets.length === 0) {
                setError('No sheets found in this spreadsheet.');
                return;
            }
            setAvailableSheets(metadata.sheets);
            setSpreadsheetTitle(metadata.properties.title);

            // Auto-select '來客紀錄' if it exists
            const defaultSheet = metadata.sheets.find(s => s.properties.title === '來客紀錄');
            if (defaultSheet) {
                const sheetName = defaultSheet.properties.title;
                setSelectedSheet(sheetName);
                // Trigger load immediately
                loadSheetData(id, sheetName);
            } else {
                setSelectedSheet('');
            }
        } catch (err: any) {
            setError(err.message || 'Failed to fetch spreadsheet metadata');
        } finally {
            setLoading(false);
        }
    }, [accessToken, loadSheetData]);

    // 從服務帳戶載入已儲存的試算表列表
    const loadSavedSpreadsheets = useCallback(async () => {
        try {
            setLoading(true);
            setError('');
            const spreadsheets = await fetchSavedSpreadsheets();
            setSavedSpreadsheets(spreadsheets);
        } catch (err: any) {
            setError(err.message || 'Failed to fetch saved spreadsheets');
        } finally {
            setLoading(false);
        }
    }, []);

    // 從服務帳戶載入選定的試算表資料
    const loadSavedSpreadsheet = useCallback(async (id: string, title: string) => {
        try {
            setLoading(true);
            setError('');
            setSpreadsheetId(id);
            setSpreadsheetTitle(title);
            setSelectedSheet('來客紀錄');
            const data = await fetchSpreadsheetData(id);
            setSheetData(data);
        } catch (err: any) {
            setError(err.message || 'Failed to load spreadsheet');
        } finally {
            setLoading(false);
        }
    }, []);

    // 登入後自動載入已儲存的試算表列表
    useEffect(() => {
        if (user) {
            loadSavedSpreadsheets();
        }
    }, [user, loadSavedSpreadsheets]);

    useEffect(() => {
        const savedSpreadsheetIds = new Set(savedSpreadsheets.map((sheet) => sheet.id));

        setSelectedCompetitionSpreadsheetIds((previousIds) => {
            const filteredIds = previousIds.filter((id) => savedSpreadsheetIds.has(id));

            if (filteredIds.length > 0) {
                return filteredIds;
            }

            if (spreadsheetId && savedSpreadsheetIds.has(spreadsheetId)) {
                return [spreadsheetId];
            }

            return filteredIds;
        });
    }, [savedSpreadsheets, spreadsheetId]);

    const performAnalysis = useCallback((dateRange: DateRange, ptaThreshold: number) => {
        if (sheetData) {
            const result = analyzeData(sheetData, dateRange, ptaThreshold);
            setAnalysisResult(result);
        }
    }, [sheetData]);

    const toggleCompetitionSpreadsheetSelection = useCallback((spreadsheetIdToToggle: string) => {
        setCompetitionSelectionMessage('');

        setSelectedCompetitionSpreadsheetIds((previousIds) => {
            if (previousIds.includes(spreadsheetIdToToggle)) {
                return previousIds.filter((id) => id !== spreadsheetIdToToggle);
            }

            if (previousIds.length >= COMPETITION_SELECTION_LIMIT) {
                setCompetitionSelectionMessage(`競賽排行最多只能選擇 ${COMPETITION_SELECTION_LIMIT} 份試算表。`);
                return previousIds;
            }

            return [...previousIds, spreadsheetIdToToggle];
        });
    }, []);

    const loadCompetitionRanking = useCallback(async (dateRange: DateRange, ptaThreshold: number) => {
        if (selectedCompetitionSpreadsheetIds.length === 0) {
            setCompetitionRankingEntries([]);
            setCompetitionSkippedSpreadsheets([]);
            setCompetitionError('');
            setCompetitionLoading(false);
            return;
        }

        try {
            setCompetitionLoading(true);
            setCompetitionError('');

            const selectedSpreadsheets = savedSpreadsheets.filter((sheet) =>
                selectedCompetitionSpreadsheetIds.includes(sheet.id)
            );

            const sheets = await fetchMultipleSpreadsheets(selectedCompetitionSpreadsheetIds);
            const competitionSources = sheets.map(({ spreadsheetId: id, sheetData: selectedSheetData }) => ({
                spreadsheetId: id,
                spreadsheetTitle: selectedSpreadsheets.find((sheet) => sheet.id === id)?.title || id,
                sheetData: selectedSheetData,
            }));

            const result = aggregateCompetitionRankings(competitionSources, dateRange, ptaThreshold);
            setCompetitionRankingEntries(result.entries);
            setCompetitionSkippedSpreadsheets(result.skippedSpreadsheets);
        } catch (err: any) {
            setCompetitionError(err.message || 'Failed to load competition ranking');
            setCompetitionRankingEntries([]);
            setCompetitionSkippedSpreadsheets([]);
        } finally {
            setCompetitionLoading(false);
        }
    }, [savedSpreadsheets, selectedCompetitionSpreadsheetIds]);

    return {
        user,
        accessToken,
        sheetData,
        availableSheets,
        selectedSheet,
        spreadsheetId,
        spreadsheetTitle,
        loading,
        error,
        analysisResult,
        savedSpreadsheets,
        selectedCompetitionSpreadsheetIds,
        competitionRankingEntries,
        competitionSkippedSpreadsheets,
        competitionLoading,
        competitionError,
        competitionSelectionMessage,
        competitionSelectionLimit: COMPETITION_SELECTION_LIMIT,
        setSpreadsheetId,
        setSelectedSheet,
        login,
        logout,
        loadSpreadsheetMetadata,
        loadSheetData,
        loadSavedSpreadsheets,
        loadSavedSpreadsheet,
        performAnalysis,
        toggleCompetitionSpreadsheetSelection,
        loadCompetitionRanking,
    };
};
