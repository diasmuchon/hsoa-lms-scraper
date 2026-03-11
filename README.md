# HSOA Gradebook Checker

Automatically scrape gradebook data from HSOA LMS and upload to Google Sheets.

## Features

- ✅ Logs into HSOA LMS automatically
- ✅ Navigates to User Management and searches for students
- ✅ Opens student profiles and extracts gradebook data
- ✅ Uploads to Google Sheets in specified format
- ✅ Runs on GitHub Actions (headless)
- ✅ Can be triggered from Google Sheets menu

## Setup Instructions

### 1. GitHub Repository Setup

1. Fork this repository or create a new one
2. Add the following files:
   - `gradebook_checker.py`
   - `requirements.txt`
   - `.github/workflows/gradebook_checker.yml`
   - `README.md`

### 2. GitHub Secrets Configuration

Go to your repository Settings → Secrets and variables → Actions → New repository secret:

Add these secrets:
- `HSOA_USERNAME`: Your HSOA login username
- `HSOA_PASSWORD`: Your HSOA login password
- `GOOGLE_CREDENTIALS_JSON`: Google Service Account JSON (see below)
- `GOOGLE_SPREADSHEET_ID`: Your Google Sheets ID (from URL)
- `STUDENT_IDS`: (Optional) Default student IDs to check
- `GITHUB_TOKEN`: (Optional) For Apps Script integration

### 3. Google Service Account Setup

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select existing
3. Enable Google Sheets API
4. Create Service Account:
   - IAM & Admin → Service Accounts → Create Service Account
   - Grant "Editor" role
   - Create key → JSON → Download
5. Add Service Account to your Google Sheet:
   - Open your Google Sheet
   - Share → Add the service account email as Editor

### 4. Google Apps Script Setup

1. Open your Google Sheet
2. Extensions → Apps Script
3. Paste the `google_apps_script.gs` code
4. Update configuration:
   - `GITHUB_TOKEN`: Your GitHub Personal Access Token
   - `GITHUB_REPO`: Your GitHub username/repository
5. Save and reload the Google Sheet
6. You should see "HSOA Tools" in the menu

## Usage

### From Google Sheets
1. Open your Google Sheet
2. Click "HSOA Tools" in the menu
3. Select "Check Gradebook" or "Select Students to Check"
4. Results will appear in the "GRADES_PROGRESS" sheet

### From GitHub Actions
1. Go to your repository → Actions
2. Select "HSOA Gradebook Checker"
3. Click "Run workflow"
4. Enter student IDs (optional)
5. Click "Run workflow"

### From Command Line
```bash
# Install dependencies
pip install -r requirements.txt

# Run with environment variables
export HSOA_USERNAME="your_username"
export HSOA_PASSWORD="your_password"
export GOOGLE_CREDENTIALS_JSON='{"your": "json"}'
export GOOGLE_SPREADSHEET_ID="your_spreadsheet_id"
export STUDENT_IDS="2023379315,2023379316"

python gradebook_checker.py --upload-sheets
