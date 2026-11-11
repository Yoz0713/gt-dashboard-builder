import { useState, useCallback, useEffect, useMemo } from 'react';
import { useGoogleLogin, googleLogout } from '@react-oauth/google';
import {
    UserProfile,
    SheetData,
    SheetInfo,
    DateRange,
    SpreadsheetListItem,
    CompetitionRankingEntry,
    CompetitionRankingSkippedSpreadsheet,
    CompetitionSheetSource,
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
    // 待追蹤名單：目前載入的門市一律納入，這裡只記「額外合併」的已儲存試算表。
    const [followUpExtraSpreadsheetIds, setFollowUpExtraSpreadsheetIds] = useState<string[]>([]);
    // 已抓回的原始列快取；日期區間或閾值變動只在記憶體內重算，不重抓。
    const [followUpSheetCache, setFollowUpSheetCache] = useState<Record<string, SheetData>>({});
    const [followUpLoading, setFollowUpLoading] = useState<boolean>(false);
    const [followUpError, setFollowUpError] = useState<string>('');
    const [followUpSelectionMessage, setFollowUpSelectionMessage] = useState<string>('');

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
        setFollowUpExtraSpreadsheetIds([]);
        setFollowUpSheetCache({});
        setFollowUpLoading(false);
        setFollowUpError('');
        setFollowUpSelectionMessage('');
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

    // 已不在清單中的試算表、或剛好變成目前門市的試算表，從額外合併名單移除。
    useEffect(() => {
        const savedSpreadsheetIds = new Set(savedSpreadsheets.map((sheet) => sheet.id));

        setFollowUpExtraSpreadsheetIds((previousIds) => {
            const filteredIds = previousIds.filter((id) => savedSpreadsheetIds.has(id) && id !== spreadsheetId);
            return filteredIds.length === previousIds.length ? previousIds : filteredIds;
        });
    }, [savedSpreadsheets, spreadsheetId]);

    const toggleFollowUpSpreadsheetSelection = useCallback((spreadsheetIdToToggle: string) => {
        // 目前門市固定納入，不能取消。
        if (spreadsheetIdToToggle === spreadsheetId) return;

        setFollowUpSelectionMessage('');

        setFollowUpExtraSpreadsheetIds((previousIds) => {
            if (previousIds.includes(spreadsheetIdToToggle)) {
                return previousIds.filter((id) => id !== spreadsheetIdToToggle);
            }

            // 上限含目前門市本身。
            if (previousIds.length + 1 >= COMPETITION_SELECTION_LIMIT) {
                setFollowUpSelectionMessage(`待追蹤名單最多合併 ${COMPETITION_SELECTION_LIMIT} 份試算表（含目前門市）。`);
                return previousIds;
            }

            return [...previousIds, spreadsheetIdToToggle];
        });
    }, [spreadsheetId]);

    // 只抓快取中還沒有的試算表。
    useEffect(() => {
        const missingIds = followUpExtraSpreadsheetIds.filter((id) => !followUpSheetCache[id]);
        if (missingIds.length === 0) return;

        let cancelled = false;

        const loadMissing = async () => {
            try {
                setFollowUpLoading(true);
                setFollowUpError('');
                const sheets = await fetchMultipleSpreadsheets(missingIds);
                if (cancelled) return;

                setFollowUpSheetCache((previousCache) => {
                    const nextCache = { ...previousCache };
                    sheets.forEach(({ spreadsheetId: id, sheetData: fetchedSheetData }) => {
                        nextCache[id] = fetchedSheetData;
                    });
                    return nextCache;
                });
            } catch (err: any) {
                if (cancelled) return;
                setFollowUpError(err.message || 'Failed to load spreadsheets for follow-up list');
                // 抓取失敗的試算表取消勾選，避免 effect 反覆重試。
                setFollowUpExtraSpreadsheetIds((previousIds) => previousIds.filter((id) => !missingIds.includes(id)));
            } finally {
                if (!cancelled) setFollowUpLoading(false);
            }
        };

        loadMissing();

        return () => {
            cancelled = true;
        };
    }, [followUpExtraSpreadsheetIds, followUpSheetCache]);

    const followUpSources = useMemo<CompetitionSheetSource[]>(() => {
        const sources: CompetitionSheetSource[] = [];

        if (sheetData) {
            sources.push({
                spreadsheetId: spreadsheetId || 'current',
                spreadsheetTitle: spreadsheetTitle || '目前門市',
                sheetData,
            });
        }

        followUpExtraSpreadsheetIds.forEach((id) => {
            const cachedSheetData = followUpSheetCache[id];
            if (!cachedSheetData || id === spreadsheetId) return;

            sources.push({
                spreadsheetId: id,
                spreadsheetTitle: savedSpreadsheets.find((sheet) => sheet.id === id)?.title || id,
                sheetData: cachedSheetData,
            });
        });

        return sources;
    }, [sheetData, spreadsheetId, spreadsheetTitle, followUpExtraSpreadsheetIds, followUpSheetCache, savedSpreadsheets]);

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
        followUpSources,
        followUpExtraSpreadsheetIds,
        followUpLoading,
        followUpError,
        followUpSelectionMessage,
        toggleFollowUpSpreadsheetSelection,
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
