#!/usr/bin/env python3
"""
Grades Progress Tracker for HS Online Academy.
Extracts student grade information and uploads to Google Sheets.

Environment Variables Required:
  HSOA_USERNAME            - Login username
  HSOA_PASSWORD            - Login password
  GOOGLE_CREDENTIALS_JSON  - Google service account JSON (as string)
  GOOGLE_SPREADSHEET_ID    - Google Sheets spreadsheet ID
"""

import argparse
import json
import logging
import os
import sys
import time
from dataclasses import dataclass
from queue import Queue
from threading import Thread

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

try:
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    GOOGLE_SHEETS_AVAILABLE = True
except ImportError:
    GOOGLE_SHEETS_AVAILABLE = False

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)

# ============================================================
# CONFIGURATION
# ============================================================
@dataclass
class Config:
    username: str = ""
    password: str = ""
    google_credentials_json: str = ""
    google_spreadsheet_id: str = ""
    google_sheet_name: str = "GRADES_PROGRESS"
    students_sheet_name: str = "STUDENTS"
    login_url: str = "https://hsoa.ordolms.com/"
    user_management_url: str = "https://hsoa.ordolms.com/home/userManagement"
    headless_mode: bool = True
    max_workers: int = 1
    page_load_timeout_seconds: int = 15
    implicit_wait_seconds: int = 3
    short_wait_seconds: int = 5

# ============================================================
# GOOGLE SHEETS
# ============================================================
def get_google_sheets_service(cfg: Config):
    if not GOOGLE_SHEETS_AVAILABLE:
        log.warning("Google Sheets libraries not installed.")
        return None
    if not cfg.google_credentials_json:
        log.warning("GOOGLE_CREDENTIALS_JSON environment variable not set.")
        return None
    try:
        info = json.loads(cfg.google_credentials_json)
        SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
        credentials = Credentials.from_service_account_info(info, scopes=SCOPES)
        service = build("sheets", "v4", credentials=credentials)
        return service
    except Exception as e:
        log.warning("Error connecting to Google Sheets: %s", e)
        return None

def fetch_active_students_from_sheet(cfg: Config) -> list:
    service = get_google_sheets_service(cfg)
    if not service:
        log.error("Could not connect to Google Sheets to fetch student data.")
        return []
    
    try:
        # Read student data from columns A, B, C, and E starting from row 2
        range_name = f"{cfg.students_sheet_name}!A2:E"
        sheet = service.spreadsheets()
        result = sheet.values().get(
            spreadsheetId=cfg.google_spreadsheet_id,
            range=range_name
        ).execute()
        
        values = result.get('values', [])
        students = []
        
        for row in values:
            # Ensure row has enough columns
            if len(row) >= 5:
                student_id = row[0].strip() if row[0] else ""
                first_name = row[1].strip() if row[1] else ""
                last_name = row[2].strip() if row[2] else ""
                status = row[4].strip().lower() if row[4] else ""
                
                # Only include students marked as "active"
                if student_id and status == "active":
                    students.append({
                        "id": student_id,
                        "first_name": first_name,
                        "last_name": last_name,
                        "full_name": f"{first_name} {last_name}".strip()
                    })
                
        log.info("Fetched %d active students from Google Sheets", len(students))
        return students
    except Exception as e:
        log.error("Error fetching student data from Google Sheets: %s", e)
        return []

def upload_to_google_sheets(cfg: Config, results: list) -> bool:
    service = get_google_sheets_service(cfg)
    if not service:
        log.warning("Could not connect to Google Sheets.")
        return False
    try:
        # Prepare header
        header = ["Student ID", "Student Name", "Course Code", "Course Name", "Assigned Grade", "Status", "Percentage"]
        rows = [header]
        
        # Process each student's data
        for result in results:
            student_id = result["student_id"]
            student_name = result["student_name"]
            
            if result["courses"]:
                for course in result["courses"]:
                    rows.append([
                        student_id,
                        student_name,
                        course["code"],
                        course["name"],
                        course["assigned_grade"],
                        course["status"],
                        course["percentage"]
                    ])
            else:
                # Add a row even if no courses found
                rows.append([student_id, student_name, "No courses found", "", "", "", ""])
        
        # Clear existing data in the sheet
        spreadsheet_id = cfg.google_spreadsheet_id
        service.spreadsheets().values().clear(
            spreadsheetId=spreadsheet_id,
            range=f"{cfg.google_sheet_name}!A:G",
        ).execute()
        
        # Upload new data
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{cfg.google_sheet_name}!A1",
            valueInputOption="RAW",
            body={"values": rows},
        ).execute()
        
        log.info("Uploaded %d records to Google Sheets.", len(rows) - 1)
        return True
    except Exception as e:
        log.warning("Error uploading to Google Sheets: %s", e)
        return False

# ============================================================
# SELENIUM HELPERS
# ============================================================
def js_click(driver: webdriver.Chrome, element):
    driver.execute_script("arguments[0].click();", element)

def safe_click(driver: webdriver.Chrome, element):
    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});")
        time.sleep(0.2)
    except Exception:
        pass
    try:
        element.click()
    except Exception:
        js_click(driver, element)

# ============================================================
# CHROMEDRIVER SETUP
# ============================================================
def setup_chrome_driver(cfg: Config, worker_id: int = 0):
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)
    if cfg.headless_mode:
        chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-features=VizDisplayCompositor")

    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.set_page_load_timeout(cfg.page_load_timeout_seconds)
        driver.implicitly_wait(cfg.implicit_wait_seconds)
        return driver
    except Exception as e:
        log.error("Error setting up ChromeDriver (worker %d): %s", worker_id, e)
        return None

# ============================================================
# SELENIUM ACTIONS
# ============================================================
def login_to_hsoa(driver: webdriver.Chrome, cfg: Config) -> bool:
    try:
        driver.get(cfg.login_url)
        time.sleep(2)
        if "home" in driver.current_url or "dashboard" in driver.current_url:
            return True
        username_field = WebDriverWait(driver, cfg.short_wait_seconds).until(
            EC.presence_of_element_located((By.NAME, "username"))
        )
        username_field.clear()
        username_field.send_keys(cfg.username)
        password_field = driver.find_element(By.NAME, "password")
        password_field.clear()
        password_field.send_keys(cfg.password)
        submit_button = driver.find_element(By.CSS_SELECTOR, 'button[type="submit"]')
        safe_click(driver, submit_button)
        time.sleep(3)
        return "login" not in driver.current_url.lower()
    except Exception as e:
        log.error("Login error: %s", e)
        return False

def navigate_to_user_management(driver: webdriver.Chrome, cfg: Config) -> bool:
    try:
        # Click on User Management link in sidebar
        user_management_link = WebDriverWait(driver, cfg.short_wait_seconds).until(
            EC.element_to_be_clickable((By.XPATH, '//span[contains(text(), "User Management")]'))
        )
        safe_click(driver, user_management_link)
        time.sleep(2)
        return True
    except Exception as e:
        log.error("Navigation to User Management error: %s", e)
        return False

def search_for_student(driver: webdriver.Chrome, student_id: str, cfg: Config) -> bool:
    try:
        # Find and fill the search input
        search_input = WebDriverWait(driver, cfg.short_wait_seconds).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[data-placeholder*="Pedro"]'))
        )
        search_input.clear()
        search_input.send_keys(student_id)
        time.sleep(1)
        return True
    except Exception as e:
        log.error("Search for student error: %s", e)
        return False

def open_student_profile(driver: webdriver.Chrome, cfg: Config) -> bool:
    try:
        # Click on the settings icon to open student profile
        settings_icon = WebDriverWait(driver, cfg.short_wait_seconds).until(
            EC.element_to_be_clickable((By.XPATH, '//mat-icon[contains(text(), "settings")]'))
        )
        safe_click(driver, settings_icon)
        time.sleep(2)
        return True
    except Exception as e:
        log.error("Open student profile error: %s", e)
        return False

def switch_to_new_window(driver: webdriver.Chrome) -> bool:
    try:
        # Switch to the newly opened window/tab
        main_window = driver.current_window_handle
        new_window = next(
            (w for w in driver.window_handles if w != main_window), None
        )
        if not new_window:
            return False
        
        driver.switch_to.window(new_window)
        time.sleep(1)
        return True
    except Exception as e:
        log.error("Switch to new window error: %s", e)
        return False

def navigate_to_gradebook(driver: webdriver.Chrome, cfg: Config) -> bool:
    try:
        # Click on the Gradebook tab
        gradebook_tab = WebDriverWait(driver, cfg.short_wait_seconds).until(
            EC.element_to_be_clickable((By.XPATH, '//div[contains(text(), "Gradebook")]'))
        )
        safe_click(driver, gradebook_tab)
        time.sleep(2)
        return True
    except Exception as e:
        log.error("Navigate to Gradebook error: %s", e)
        return False

def change_items_per_page(driver: webdriver.Chrome, cfg: Config) -> bool:
    try:
        # Change items per page to show all courses
        items_selector = WebDriverWait(driver, cfg.short_wait_seconds).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'mat-select[aria-label*="Items per page"]'))
        )
        safe_click(driver, items_selector)
        time.sleep(0.5)
        
        # Select "30" items per page
        options = driver.find_elements(By.CSS_SELECTOR, "mat-option")
        for option in options:
            text = option.text.strip()
            if text == "30":
                safe_click(driver, option)
                time.sleep(1)
                return True
        
        return False
    except Exception as e:
        log.error("Change items per page error: %s", e)
        return False

def extract_student_name(driver: webdriver.Chrome) -> str:
    try:
        # Extract student name from the profile page
        name_element = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '.student-name, h1, .profile-header'))
        )
        return name_element.text.strip()
    except Exception:
        return "Unknown Student"

def extract_course_data(driver: webdriver.Chrome, cfg: Config) -> list:
    courses = []
    try:
        # Wait for course table to load
        WebDriverWait(driver, cfg.short_wait_seconds).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "tbody tr"))
        )
        time.sleep(1)
        
        # Find all course rows
        rows = driver.find_elements(By.CSS_SELECTOR, "tbody tr")
        
        for row in rows:
            try:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) >= 7:  # Ensure we have enough columns
                    course_code = cells[0].text.strip()
                    course_name = cells[1].text.strip()
                    assigned_grade = cells[2].text.strip()
                    status = cells[4].text.strip()
                    percentage = cells[6].text.strip()
                    
                    # Only include courses with data
                    if course_code or course_name:
                        courses.append({
                            "code": course_code,
                            "name": course_name,
                            "assigned_grade": assigned_grade,
                            "status": status,
                            "percentage": percentage
                        })
            except Exception as e:
                log.warning("Error extracting course data from row: %s", e)
                continue
                
        return courses
    except Exception as e:
        log.error("Error extracting course data: %s", e)
        return courses

def process_student(
    driver: webdriver.Chrome,
    student_id: str,
    cfg: Config,
) -> dict:
    result = {
        "student_id": student_id,
        "student_name": "",
        "courses": [],
        "success": False,
    }
    
    try:
        # Navigate to user management
        if not navigate_to_user_management(driver, cfg):
            return result
            
        # Search for student
        if not search_for_student(driver, student_id, cfg):
            return result
            
        # Open student profile
        if not open_student_profile(driver, cfg):
            return result
            
        # Switch to new window
        if not switch_to_new_window(driver):
            return result
            
        # Get student name
        result["student_name"] = extract_student_name(driver)
        
        # Navigate to gradebook
        if not navigate_to_gradebook(driver, cfg):
            # Close window and return if gradebook navigation fails
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
            return result
            
        # Change items per page
        change_items_per_page(driver, cfg)
        
        # Extract course data
        result["courses"] = extract_course_data(driver, cfg)
        result["success"] = True
        
        # Close the student profile window
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        time.sleep(0.5)
        
        return result
    except Exception as e:
        log.error("Error processing %s: %s", student_id, e)
        try:
            # Ensure we close any open windows
            while len(driver.window_handles) > 1:
                driver.switch_to.window(driver.window_handles[-1])
                driver.close()
            driver.switch_to.window(driver.window_handles[0])
        except Exception:
            pass
        return result

# ============================================================
# WORKER
# ============================================================
def worker_process_students(
    worker_id: int,
    students: list,  # Now a list of student dictionaries
    cfg: Config,
    results_queue: Queue,
) -> None:
    log.info("[Worker %d] Starting...", worker_id)
    driver = setup_chrome_driver(cfg, worker_id)
    if not driver:
        log.error("[Worker %d] Failed to start browser.", worker_id)
        return
    
    try:
        # Login to HS Online Academy
        if not login_to_hsoa(driver, cfg):
            log.warning("[Worker %d] Login may have failed, continuing...", worker_id)
        
        # Process each student
        for student in students:
            student_id = student["id"]
            full_name = student["full_name"]
            
            log.info("[Worker %d] Processing: %s (%s)", worker_id, full_name, student_id)
            
            # Process the student through the LMS
            result = process_student(driver, student_id, cfg)
            
            # Override the student name with the one from our sheet
            result["student_name"] = full_name
            result["student_id"] = student_id  # Ensure correct ID is used
            
            results_queue.put(result)
            log.info(
                "[Worker %d] Done: %s - Found %d courses",
                worker_id,
                full_name,
                len(result["courses"]),
            )
    except Exception as e:
        log.error("[Worker %d] Error: %s", worker_id, e)
    finally:
        driver.quit()
        log.info("[Worker %d] Finished.", worker_id)

# ============================================================
# CLI
# ============================================================
def parse_args():
    parser = argparse.ArgumentParser(
        description="Extract student grades from HS Online Academy and upload to Google Sheets."
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=1,
        help="Number of parallel browser workers (default: 1).",
    )
    return parser.parse_args()

def build_config(args) -> Config:
    username = os.environ.get("HSOA_USERNAME", "")
    password = os.environ.get("HSOA_PASSWORD", "")
    if not username or not password:
        log.error(
            "HSOA_USERNAME and HSOA_PASSWORD environment variables must be set."
        )
        sys.exit(1)

    return Config(
        username=username,
        password=password,
        google_credentials_json=os.environ.get("GOOGLE_CREDENTIALS_JSON", ""),
        google_spreadsheet_id=os.environ.get("GOOGLE_SPREADSHEET_ID", ""),
        google_sheet_name=os.environ.get("GOOGLE_SHEET_NAME", "GRADES_PROGRESS"),
        students_sheet_name=os.environ.get("STUDENTS_SHEET_NAME", "STUDENTS"),
        max_workers=args.workers,
    )

def distribute(items: list, n: int) -> list:
    """Split items into n roughly equal chunks."""
    chunks = [[] for _ in range(n)]
    for i, item in enumerate(items):
        chunks[i % n].append(item)
    return chunks

def main():
    args = parse_args()
    cfg = build_config(args)
    
    # Fetch active students from Google Sheets
    students = fetch_active_students_from_sheet(cfg)
    if not students:
        log.error("No active students found in Google Sheets.")
        sys.exit(1)

    log.info(
        "Processing %d active student(s) with %d worker(s).",
        len(students),
        cfg.max_workers,
    )

    results_queue: Queue = Queue()
    num_workers = min(cfg.max_workers, len(students))
    chunks = distribute(students, num_workers)

    threads = []
    for worker_id, chunk in enumerate(chunks):
        t = Thread(
            target=worker_process_students,
            args=(worker_id, chunk, cfg, results_queue),
        )
        t.start()
        threads.append(t)

    for t in threads:
        t.join()

    all_results = []
    while not results_queue.empty():
        result = results_queue.get()
        all_results.append(result)

    log.info("Processing complete. Uploading to Google Sheets...")
    
    if cfg.google_spreadsheet_id:
        if upload_to_google_sheets(cfg, all_results):
            log.info("Successfully uploaded data to Google Sheets.")
        else:
            log.error("Failed to upload data to Google Sheets.")
    else:
        log.warning("GOOGLE_SPREADSHEET_ID not set; skipping Google Sheets upload.")

    success_count = sum(1 for r in all_results if r["success"])
    log.info(
        "Done. %d/%d students processed successfully.", success_count, len(students)
    )

if __name__ == "__main__":
    main()
