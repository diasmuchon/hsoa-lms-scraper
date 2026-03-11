# HSOA LMS Gradebook Checker

Automatically checks grades for students from HSOA LMS and updates Google Sheets.

## Setup

### 1. Google Sheets Preparation

Create a sheet named "STUDENTS" with student IDs in column A starting from row 2:

| A | B | C |
|---|---|---|
| Student ID | | |
| 2023379315 | | |
| 2023379316 | | |
| ... | | |

### 2. GitHub Secrets

Set these secrets in your repository:

- `HSOA_USERNAME`: Your HSOA LMS username
- `HSOA_PASSWORD`: Your HSOA LMS password  
- `GOOGLE_CREDENTIALS_JSON`: Your Google service account JSON
- `GOOGLE_SPREADSHEET_ID`: Your Google Sheet ID

### 3. Service Account Setup

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a service account
3. Download JSON credentials
4. Share your Google Sheet with the service account email

## How It Works

1. Reads student IDs from "STUDENTS" sheet column A2:A
2. Logs into HSOA LMS
3. Fetches grades for each student
4. Updates "GRADES" sheet with new records
5. Runs automatically twice daily (6 AM/6 PM Oman time)

## Manual Run

You can manually trigger the workflow from GitHub Actions tab.
