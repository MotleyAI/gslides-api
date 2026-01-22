# Creating a credentials.json to use with this package

## 1. Go to the Google Cloud Console Credentials page

Open this link:

👉 [Google Cloud Console Credentials](https://console.cloud.google.com/apis/credentials)

## 2. Create or Select a Google Cloud Project

- Click the project dropdown at the top
- Select an existing project or click "New Project" to create one
- After creation, make sure it's selected

## 3. Enable APIs (VERY important!)

You must enable both:

- **Google Slides API**
- **Google Sheets API**

To do this:

1. Go to **APIs & Services > Library**
2. Search for and enable each API:
   - Search "Google Slides API" → Click → Enable
   - Search "Google Sheets API" → Click → Enable

## 4. Create OAuth 2.0 Credentials

1. Go to the Credentials page
2. Click **Create Credentials** → Choose **OAuth client ID**
3. If prompted, configure the OAuth consent screen:
   - Choose **External**
   - Fill in the basic info (App name, support email, etc.)
   - Add scopes:
     - `https://www.googleapis.com/auth/presentations`
     - `https://www.googleapis.com/auth/spreadsheets`
   - Add your Google account as a test user (under "Test Users")
4. After the consent screen is done, return and:
   - Select Application type: **Desktop app**
   - Name it (e.g., "SlidesSheetsApp")
   - Click **Create**

## 5. Download the Credentials File

After creating the OAuth client ID:

1. Click **Download JSON**
2. Save it as `credentials.json` in the directory that you will pass to the `initialize_credentials` function.
