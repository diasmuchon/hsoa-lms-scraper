import os
import json
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from playwright.sync_api import sync_playwright

# Environment variables injected by GitHub Actions
STUDENT_ID = os.environ.get("STUDENT_ID")
SHEET_ROW = int(os.environ.get("SHEET_ROW"))
LMS_USER = os.environ.get("LMS_USER")
LMS_PASS = os.environ.get("LMS_PASS")
GCP_CREDS_JSON = os.environ.get("GCP_CREDS_JSON")
SPREADSHEET_ID = '1uEOB5-MwhlsbD8IrLkhMD6PPx0vBeWKck4HmpB4Lsro'

def update_google_sheet(row, course_list):
    creds_dict = json.loads(GCP_CREDS_JSON)
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SPREADSHEET_ID).worksheet("STUDENTS")
    
    # Format the array into a newline-separated string
    subjects_text = "\n".join(course_list)
    sheet.update_cell(row, 10, subjects_text) # Column J is index 10
    print(f"Successfully updated row {row} with {len(course_list)} courses.")

def run():
    course_list = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        
        print("Logging into LMS...")
        page.goto("https://hsoa.ordolms.com/login")
        page.fill("input[name='username']", LMS_USER)
        page.fill("input[name='password']", LMS_PASS)
        page.click("button[type='submit']")
        
        # Wait for the dashboard to load by looking for the User Management icon text
        page.wait_for_selector("text=User Management", timeout=15000)
        
        print(f"Navigating to User Management for ID: {STUDENT_ID}")
        page.goto("https://hsoa.ordolms.com/home/userManagement")
        
        # Filter by student ID
        page.wait_for_selector("input[placeholder='Ex. Pedro Perez']")
        page.fill("input[placeholder='Ex. Pedro Perez']", STUDENT_ID)
        time.sleep(3) # Wait for Angular to filter the list
        
        # Click the settings gear icon for the student
        page.locator("mat-icon", has_text="settings").first.click()
        
        print("Loading courses table...")
        page.wait_for_selector("mat-select", timeout=15000)
        page.locator("mat-select").click()
        page.locator("mat-option", has_text="30").click()
        time.sleep(4) # Wait for the table rows to populate
        
        # Extract the table rows
        rows = page.query_selector_all("tbody[role='rowgroup'] tr.mat-row")
        for row in rows:
            cells = row.query_selector_all("td")
            if len(cells) >= 8:
                code = cells[0].inner_text().strip()
                name = cells[1].inner_text().strip()
                grade = cells[2].inner_text().strip()
                percentage = cells[6].inner_text().strip()
                
                # If percentage is 100%, consider it 1 credit, else 0
                credits = "1" if percentage == "100%" else "0"
                formatted_line = f"{name} ({code})|{percentage}|{grade}|{credits}"
                course_list.append(formatted_line)
                
        browser.close()
    
    if course_list:
        update_google_sheet(SHEET_ROW, course_list)
    else:
        print("No courses found or scraping failed.")

if __name__ == '__main__':
    run()
