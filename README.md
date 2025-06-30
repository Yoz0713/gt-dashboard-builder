# Google Sheets Dashboard Builder

A React TypeScript application that allows users to authenticate with Google OAuth and create interactive dashboards from their Google Sheets data.

## Features

- **Google OAuth Authentication**: Secure login with Google accounts
- **Google Sheets Integration**: Read data from any accessible Google Sheet
- **Sheet Selection**: Choose from multiple sheets within a spreadsheet
- **Data Visualization**: 
  - Visual summary showing total rows, non-empty rows, and column count
  - Interactive bar chart for Column B data using Chart.js
- **Responsive Design**: Built with TailwindCSS for modern, mobile-friendly UI

## Prerequisites

1. Node.js (version 16 or higher)
2. Google Cloud Console project with OAuth2 credentials
3. Google Sheets API enabled

## Google Cloud Setup

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the following APIs:
   - Google Sheets API
   - Google OAuth2 API
4. Create OAuth2 credentials:
   - Go to "Credentials" → "Create Credentials" → "OAuth client ID"
   - Choose "Web application"
   - Add your domain to authorized origins (e.g., `http://localhost:3000` for development)
5. Copy the Client ID

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd gt-dashboard-builder
```

2. Install dependencies:
```bash
npm install
```

3. Create a `.env` file in the root directory:
```bash
REACT_APP_GOOGLE_CLIENT_ID=your-google-client-id-here.apps.googleusercontent.com
```

4. Start the development server:
```bash
npm start
```

5. Open [http://localhost:3000](http://localhost:3000) in your browser

## Usage

1. **Sign In**: Click "Sign in with Google" and authenticate with your Google account
2. **Enter Sheet URL**: Paste a Google Sheets URL (make sure you have read access to the sheet)
3. **Select Sheet**: Choose which sheet/tab you want to analyze from the dropdown
4. **View Dashboard**: 
   - See data summary with row and column counts
   - View bar chart visualization of Column B data
   - Browse the full data table

## Project Structure

```
src/
├── App.tsx          # Main application component
├── types.ts         # TypeScript type definitions
├── index.tsx        # React entry point
└── index.css        # TailwindCSS styles
```

## Dependencies

- **React 18**: Modern React with functional components
- **TypeScript**: Type-safe development
- **@react-oauth/google**: Google OAuth integration
- **Chart.js + react-chartjs-2**: Data visualization
- **TailwindCSS**: Utility-first CSS framework

## API Permissions

The application requests the following Google API scopes:
- `openid profile email`: Basic user profile information
- `https://www.googleapis.com/auth/spreadsheets.readonly`: Read-only access to Google Sheets

## Troubleshooting

### OAuth Errors
- Ensure your Google Cloud project has the correct APIs enabled
- Check that your Client ID is correctly set in the `.env` file
- Verify authorized origins include your domain

### Sheet Access Errors
- Make sure you have read access to the Google Sheet
- Check if the sheet URL is correctly formatted
- Ensure the sheet contains data in the expected format

### Chart Not Displaying
- Verify that Column B contains numeric data
- Check browser console for any JavaScript errors
- Ensure Chart.js dependencies are properly installed

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

This project is open source and available under the MIT License. 