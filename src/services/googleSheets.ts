import { SheetData, SpreadsheetMetadata, UserProfile } from '../types';

export const fetchUserProfile = async (token: string): Promise<UserProfile> => {
    const response = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
        headers: { Authorization: `Bearer ${token}` },
    });

    if (!response.ok) {
        throw new Error('Failed to fetch user profile');
    }

    const profile = await response.json();
    return {
        id: profile.id,
        name: profile.name,
        email: profile.email,
        picture: profile.picture,
    };
};

export const fetchSpreadsheetMetadata = async (id: string, token: string): Promise<SpreadsheetMetadata> => {
    const response = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${id}?fields=properties.title,sheets.properties`,
        {
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json',
            },
        }
    );

    if (!response.ok) {
        const errorText = await response.text();
        console.error('API Error Response:', errorText);

        if (response.status === 403) {
            throw new Error('Permission denied. Please make sure you have access to this spreadsheet and the Google Sheets API is enabled.');
        } else if (response.status === 404) {
            throw new Error('Spreadsheet not found. Please check the URL and try again.');
        } else {
            throw new Error(`Failed to fetch spreadsheet information: ${response.status} ${response.statusText}`);
        }
    }

    return await response.json();
};

export const fetchSheetData = async (id: string, sheetName: string, token: string): Promise<SheetData> => {
    const response = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${encodeURIComponent(sheetName)}`,
        {
            headers: { Authorization: `Bearer ${token}` },
        }
    );

    if (!response.ok) {
        const errorText = await response.text();
        console.error('API Error Response:', errorText);
        throw new Error(`Failed to fetch sheet data: ${response.status} ${response.statusText}`);
    }

    return await response.json();
};
