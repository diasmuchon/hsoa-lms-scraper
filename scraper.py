#!/usr/bin/env python3
"""
HSOA Gradebook Checker - CLI version for GitHub Actions.

Scrapes student gradebook data from HSOA LMS and uploads to Google Sheets.

Environment Variables:
  HSOA_USERNAME            - Login username
  HSOA_PASSWORD            - Login password
  GOOGLE_CREDENTIALS_JSON  - Google service account JSON (as string)
  GOOGLE_SPREADSHEET_ID    - Google Sheets spreadsheet ID
  GOOGLE_SHEET_NAME        - Sheet name (default: GRADES_PROGRESS)
  STUDENT_IDS              - Comma-separated list of student IDs to process
"""

import argparse
import csv
import json
import logging
import os
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import List, Dict, Optional, Tuple

from selenium import webdriver
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementNotInteractableException,
    StaleElementReferenceException,
    ElementClickInterceptedException
)
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager

try:
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    GOOGLE_SHEETS_AVAILABLE = True
except ImportError:
    GOOGLE_SHEETS_AVAILABLE = False

# ============================================================
# LOGGING CONFIGURATION
# ============================================================
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
    output_csv_file: Path = Path("gradebook_report.csv")
    google_credentials_json: str = ""
    google_spreadsheet_id: str = ""
    google_sheet_name: str = "GRADES_PROGRESS"
    login_url: str = "https://hsoa.ordolms.com/"
    user_management_url: str = "https://hsoa.ordolms.com/home/userManagement"
    headless_mode: bool = True
    page_load_timeout_seconds: int = 30
    implicit_wait_seconds: int = 5
    short_wait_seconds: int = 5
    long_wait_seconds: int = 10


# ============================================================
# GOOGLE SHEETS FUNCTIONS
# ============================================================
def get_google_sheets_service(cfg: Config):
    """Initialize Google Sheets API service."""
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


def upload_to_google_sheets(cfg: Config, results: List[Dict]) -> bool:
    """Upload gradebook results to Google Sheets."""
    service = get_google_sheets_service(cfg)
    if not service:
        log.warning("Could not connect to Google Sheets.")
        return False
    
    try:
        # Build header row
        header = [
            "Student ID",
            "Student Name",
            "Course Code",
            "Course Name",
            "Assigned Grade",
            "Status",
            "Percentage",
            "Last Updated"
        ]
        
        rows = [header]
        
        # Build data rows - one row per course per student
        for result in results:
            student_id = result.get("student_id", "")
            student_name = result.get("student_name", "")
            courses = result.get("courses", [])
            
            if not courses:
                # Add a row even if no courses found
                rows.append([
                    student_id,
                    student_name,
                    "No courses found",
                    "",
                    "",
                    "",
                    "",
                    time.strftime("%Y-%m-%d %H:%M:%S")
                ])
            else:
                for course in courses:
                    rows.append([
                        student_id,
                        student_name,
                        course.get("code", ""),
                        course.get("name", ""),
                        course.get("assigned_grade", ""),
                        course.get("status", ""),
                        course.get("percentage", ""),
                        time.strftime("%Y-%m-%d %H:%M:%S")
                    ])

        spreadsheet_id = cfg.google_spreadsheet_id
        sheet_name = cfg.google_sheet_name
        
        # Clear existing data
        try:
            service.spreadsheets().values().clear(
                spreadsheetId=spreadsheet_id,
                range=f"{sheet_name}!A:H",
            ).execute()
            log.info("Cleared existing data from Google Sheets.")
        except Exception as e:
            log.warning("Could not clear sheet, may be empty: %s", e)
        
        # Write new data
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_name}!A1",
            valueInputOption="RAW",
            body={"values": rows},
        ).execute()
        
        log.info("Successfully uploaded %d rows to Google Sheets.", len(rows) - 1)
        return True
        
    except Exception as e:
        log.error("Error uploading to Google Sheets: %s", e)
        return False


# ============================================================
# SELENIUM HELPER FUNCTIONS
# ============================================================
def js_click(driver: webdriver.Chrome, element):
    """Click element using JavaScript."""
    driver.execute_script("arguments[0].click();", element)


def safe_click(driver: webdriver.Chrome, element, use_js: bool = True):
    """Safely click an element with scrolling and fallback."""
    try:
        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", 
            element
        )
        time.sleep(0.3)
    except Exception:
        pass
    
    try:
        if use_js:
            js_click(driver, element)
        else:
            element.click()
    except (ElementClickInterceptedException, ElementNotInteractableException):
        js_click(driver, element)
    except Exception as e:
        log.warning("Click failed, trying JavaScript: %s", e)
        js_click(driver, element)


def wait_for_element(driver: webdriver.Chrome, by: By, value: str, 
                     timeout: int = 10, clickable: bool = False):
    """Wait for element to be present or clickable."""
    try:
        if clickable:
            return WebDriverWait(driver, timeout).until(
                EC.element_to_be_clickable((by, value)))
        return WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, value)))
    except TimeoutException:
        return None


def wait_for_elements(driver: webdriver.Chrome, by: By, value: str, 
                      timeout: int = 10, min_count: int = 1):
    """Wait for elements to be present and return them."""
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: len(d.find_elements(by, value)) >= min_count
        )
        return driver.find_elements(by, value)
    except TimeoutException:
        return []


# ============================================================
# CHROMEDRIVER SETUP
# ============================================================
def setup_chrome_driver(cfg: Config) -> Optional[webdriver.Chrome]:
    """Set up Chrome WebDriver with appropriate options."""
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--start-maximized")
    
    # Add stealth options
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    chrome_options.add_experimental_option("useAutomationExtension", False)
    
    if cfg.headless_mode:
        chrome_options.add_argument("--headless=new")
    
    chrome_options.add_argument("--disable-features=VizDisplayCompositor")
    chrome_options.add_argument("--log-level=3")  # Suppress logging

    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.set_page_load_timeout(cfg.page_load_timeout_seconds)
        driver.implicitly_wait(cfg.implicit_wait_seconds)
        
        # Add stealth script to avoid detection
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        return driver
    except Exception as e:
        log.error("Error setting up ChromeDriver: %s", e)
        return None


# ============================================================
# LOGIN FUNCTIONS
# ============================================================
def login_to_hsoa(driver: webdriver.Chrome, cfg: Config) -> bool:
    """Login to HSOA LMS."""
    try:
        log.info("Navigating to login page...")
        driver.get(cfg.login_url)
        time.sleep(3)
        
        # Check if already logged in
        if "home" in driver.current_url or "dashboard" in driver.current_url:
            log.info("Already logged in.")
            return True
        
        # Wait for login form
        time.sleep(2)
        
        # Find username field
        username_field = wait_for_element(
            driver, By.XPATH, '//input[@name="username" or @placeholder="Username"]', 
            cfg.short_wait_seconds
        )
        if not username_field:
            # Try other possible selectors
            username_field = wait_for_element(driver, By.CSS_SELECTOR, 
                                            'input[type="text"], input[type="email"]', 
                                            cfg.short_wait_seconds)
        
        if not username_field:
            log.error("Could not find username field")
            # Try to take screenshot for debugging
            try:
                driver.save_screenshot("login_error.png")
                log.info("Saved screenshot: login_error.png")
            except:
                pass
            return False
        
        username_field.clear()
        username_field.send_keys(cfg.username)
        log.info("Entered username")
        time.sleep(0.5)
        
        # Find password field
        password_field = driver.find_element(By.XPATH, '//input[@type="password"]')
        password_field.clear()
        password_field.send_keys(cfg.password)
        log.info("Entered password")
        time.sleep(0.5)
        
        # Find and click submit button
        submit_buttons = driver.find_elements(By.XPATH, 
                                            '//button[@type="submit" or contains(text(), "Login") or contains(text(), "Sign In")]')
        if submit_buttons:
            safe_click(driver, submit_buttons[0])
        else:
            # Try form submission
            password_field.submit()
        
        log.info("Clicked login button")
        time.sleep(5)
        
        # Check for login errors
        error_elements = driver.find_elements(By.XPATH, 
                                            '//div[contains(@class, "error") or contains(@class, "alert")]')
        if error_elements:
            for error in error_elements[:3]:  # Check first 3 error elements
                if error.is_displayed():
                    log.error("Login error message: %s", error.text[:100])
        
        # Verify login success
        if "login" in driver.current_url.lower():
            log.error("Login failed - still on login page")
            return False
        
        log.info("Login successful - current URL: %s", driver.current_url)
        return True
        
    except Exception as e:
        log.error("Login error: %s", e)
        import traceback
        log.error(traceback.format_exc())
        return False


# ============================================================
# NAVIGATION FUNCTIONS
# ============================================================
def navigate_to_user_management(driver: webdriver.Chrome, cfg: Config) -> bool:
    """Navigate to User Management page."""
    try:
        log.info("Navigating to User Management...")
        
        # Try direct URL first
        driver.get(cfg.user_management_url)
        time.sleep(3)
        
        # Verify we're on the right page
        time.sleep(2)
        current_url = driver.current_url
        if "userManagement" in current_url or "user-management" in current_url.lower():
            log.info("Successfully navigated to User Management")
            return True
        
        # Try finding the User Management button/menu item
        user_mgmt_selectors = [
            '//span[contains(text(), "User Management")]',
            '//button[contains(text(), "User Management")]',
            '//a[contains(text(), "User Management")]',
            '//div[contains(text(), "User Management")]',
            '//mat-icon[contains(text(), "manage_accounts")]',
            '//*[contains(@class, "user-management")]',
        ]
        
        for selector in user_mgmt_selectors:
            try:
                element = wait_for_element(driver, By.XPATH, selector, 3, clickable=True)
                if element:
                    safe_click(driver, element)
                    time.sleep(3)
                    return True
            except:
                continue
        
        log.warning("Could not find User Management navigation element")
        return False
        
    except Exception as e:
        log.error("Error navigating to User Management: %s", e)
        return False


# ============================================================
# STUDENT SEARCH AND PROFILE FUNCTIONS
# ============================================================
def search_student(driver: webdriver.Chrome, student_id: str, cfg: Config) -> Tuple[bool, str]:
    """Search for a student by ID and return (success, student_name)."""
    try:
        log.info("Searching for student: %s", student_id)
        time.sleep(2)
        
        # Find search input - try multiple selectors
        search_selectors = [
            '//input[contains(@data-placeholder, "Pedro") or contains(@placeholder, "search")]',
            '//input[@type="search"]',
            '//input[contains(@class, "mat-input")]',
            '//input[contains(@id, "search")]',
        ]
        
        search_input = None
        for selector in search_selectors:
            search_input = wait_for_element(driver, By.XPATH, selector, 5)
            if search_input:
                break
        
        if not search_input:
            log.error("Could not find search input")
            return False, ""
        
        # Clear and enter student ID
        try:
            search_input.clear()
            time.sleep(0.5)
            search_input.send_keys(student_id)
            log.info("Entered student ID in search")
            time.sleep(2)  # Wait for search results
        except:
            # Try JavaScript approach
            driver.execute_script("arguments[0].value = '';", search_input)
            driver.execute_script("arguments[0].value = arguments[1];", search_input, student_id)
            driver.execute_script("arguments[0].dispatchEvent(new Event('input'));", search_input)
            time.sleep(2)
        
        # Wait for table to update
        time.sleep(2)
        
        # Try to find the student row and get name
        student_name = ""
        try:
            # Look for table rows
            rows = wait_for_elements(driver, By.XPATH, '//tbody//tr[contains(@class, "mat-row")]', 5)
            
            for row in rows:
                try:
                    # Check if this row contains the student ID
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if len(cells) >= 3:  # Should have at least ID, first name, last name
                        # Look for student ID in any cell
                        row_text = row.text
                        if student_id in row_text:
                            # Try to extract name from cells
                            first_name_cell = row.find_element(By.XPATH, './/td[contains(@class, "firstName") or contains(text(), "Cyrus")]')
                            last_name_cell = row.find_element(By.XPATH, './/td[contains(@class, "lastName") or contains(text(), "Lemmon")]')
                            
                            if first_name_cell and last_name_cell:
                                student_name = f"{first_name_cell.text.strip()} {last_name_cell.text.strip()}"
                            elif len(cells) >= 3:
                                # Assume first name is cell 1, last name is cell 2 (adjust based on actual structure)
                                student_name = f"{cells[1].text.strip()} {cells[2].text.strip()}"
                            break
                except:
                    continue
                    
            if not student_name:
                log.warning("Could not extract student name from table")
        except Exception as e:
            log.warning("Could not get student name from table: %s", e)
        
        return True, student_name
        
    except Exception as e:
        log.error("Error searching for student: %s", e)
        return False, ""


def open_student_profile(driver: webdriver.Chrome, student_id: str, cfg: Config) -> bool:
    """Click the settings icon to open student profile."""
    try:
        log.info("Looking for settings icon for student: %s", student_id)
        time.sleep(2)
        
        # Try multiple selectors for the settings/overview link
        settings_selectors = [
            f'//a[contains(@href, "{student_id}")]',
            '//a[@mattooltip="overview"]',
            '//a[contains(@href, "overview")]',
            '//mat-icon[contains(text(), "settings")]',
            '//button[contains(@class, "mat-icon-button")]//mat-icon[contains(text(), "settings")]',
        ]
        
        settings_element = None
        for selector in settings_selectors:
            try:
                elements = driver.find_elements(By.XPATH, selector)
                for element in elements:
                    if element.is_displayed():
                        settings_element = element
                        break
                if settings_element:
                    break
            except:
                continue
        
        if not settings_element:
            log.error("Could not find settings icon/link")
            return False
        
        # Get the href if it's an anchor tag
        try:
            if settings_element.tag_name == 'a':
                href = settings_element.get_attribute('href')
                if href:
                    log.info("Navigating directly to student profile: %s", href)
                    driver.get(href)
                    time.sleep(3)
                    return True
        except:
            pass
        
        # Otherwise click the element
        log.info("Clicking settings icon")
        safe_click(driver, settings_element)
        time.sleep(3)
        
        # Check if new window/tab opened
        if len(driver.window_handles) > 1:
            driver.switch_to.window(driver.window_handles[-1])
            log.info("Switched to new window/tab")
        
        return True
        
    except Exception as e:
        log.error("Error opening student profile: %s", e)
        return False


# ============================================================
# GRADEBOOK FUNCTIONS
# ============================================================
def navigate_to_gradebook_tab(driver: webdriver.Chrome, cfg: Config) -> bool:
    """Click on the Gradebook tab in student profile."""
    try:
        log.info("Looking for Gradebook tab...")
        time.sleep(2)
        
        # Try multiple selectors for Gradebook tab
        gradebook_selectors = [
            '//div[contains(@class, "mat-tab-label-content") and contains(text(), "Gradebook")]',
            '//div[contains(text(), "Gradebook")]',
            '//span[contains(text(), "Gradebook")]',
            '//*[contains(text(), "Gradebook") and (@role="tab" or contains(@class, "tab"))]',
        ]
        
        for selector in gradebook_selectors:
            try:
                gradebook_tab = wait_for_element(driver, By.XPATH, selector, 5, clickable=True)
                if gradebook_tab and gradebook_tab.is_displayed():
                    log.info("Found Gradebook tab, clicking...")
                    safe_click(driver, gradebook_tab)
                    time.sleep(3)
                    
                    # Verify we're on Gradebook tab
                    active_tabs = driver.find_elements(By.XPATH, 
                                                      '//div[contains(@class, "mat-tab-label-active")]')
                    for tab in active_tabs:
                        if "gradebook" in tab.text.lower():
                            log.info("Successfully switched to Gradebook tab")
                            return True
                    
                    # Alternative verification
                    gradebook_content = driver.find_elements(By.XPATH, 
                                                           '//*[contains(text(), "Course Code") or contains(text(), "ENG092ER")]')
                    if gradebook_content:
                        log.info("Gradebook content detected")
                        return True
            except:
                continue
        
        log.error("Could not find or click Gradebook tab")
        return False
        
    except Exception as e:
        log.error("Error navigating to Gradebook tab: %s", e)
        return False


def set_items_per_page(driver: webdriver.Chrome, items: int = 30, cfg: Config = None) -> bool:
    """Set the items per page dropdown to specified value."""
    try:
        log.info(f"Setting items per page to {items}...")
        time.sleep(2)
        
        # Find the items per page selector
        items_selectors = [
            '//mat-select[contains(@aria-label, "Items per page")]',
            '//mat-select[@aria-label="Items per page:"]',
            '//mat-select[contains(@id, "mat-select")]',
            '//div[contains(@class, "mat-select")]',
        ]
        
        items_selector = None
        for selector in items_selectors:
            items_selector = wait_for_element(driver, By.XPATH, selector, 5, clickable=True)
            if items_selector:
                break
        
        if not items_selector:
            log.warning("Could not find items per page selector")
            return False
        
        # Click to open dropdown
        safe_click(driver, items_selector)
        time.sleep(1)
        
        # Find and click the option
        option_selectors = [
            f'//mat-option[contains(@class, "mat-option") and .//span[contains(text(), "{items}")]]',
            f'//mat-option[.//span[text()="{items}"]]',
            f'//div[contains(@class, "mat-option-text") and contains(text(), "{items}")]',
        ]
        
        for selector in option_selectors:
            try:
                option = wait_for_element(driver, By.XPATH, selector, 3, clickable=True)
                if option:
                    safe_click(driver, option)
                    time.sleep(2)  # Wait for page to reload
                    log.info(f"Successfully set items per page to {items}")
                    return True
            except:
                continue
        
        log.warning(f"Could not find option for {items} items per page")
        return False
        
    except Exception as e:
        log.error("Error setting items per page: %s", e)
        return False


def extract_gradebook_data(driver: webdriver.Chrome, cfg: Config) -> List[Dict]:
    """Extract gradebook data from the current page."""
    courses = []
    try:
        log.info("Extracting gradebook data...")
        time.sleep(2)
        
        # Wait for table to load
        table = wait_for_element(driver, By.XPATH, '//table[contains(@class, "mat-table")]', 10)
        if not table:
            log.warning("No gradebook table found")
            return courses
        
        # Find all course rows
        rows = wait_for_elements(driver, By.XPATH, 
                               '//tbody//tr[contains(@class, "mat-row")]', 5, min_count=1)
        
        if not rows:
            log.warning("No course rows found in gradebook")
            return courses
        
        log.info(f"Found {len(rows)} course rows")
        
        for row in rows:
            try:
                # Extract data from each cell
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) < 7:  # Need at least 7 cells for required data
                    continue
                
                # Course Code (first cell)
                code = cells[0].text.strip() if cells[0].text.strip() else "N/A"
                
                # Course Name (second cell, remove checkmark if present)
                name = cells[1].text.strip()
                name = re.sub(r'^✅\s*', '', name)  # Remove leading checkmark
                
                # Assigned Grade (third cell)
                assigned_grade = cells[2].text.strip() if cells[2].text.strip() else "N/A"
                
                # Status (fifth cell, get color and text)
                status = "Unknown"
                try:
                    status_cell = cells[4]
                    status_text = status_cell.text.strip().lower()
                    
                    # Check for status indicators
                    if "completed" in status_text:
                        status = "Completed"
                    elif "in progress" in status_text:
                        status = "In Progress"
                    elif "pending" in status_text:
                        status = "Pending"
                    else:
                        # Try to get from paragraph text
                        p_tags = status_cell.find_elements(By.TAG_NAME, "p")
                        for p in p_tags:
                            p_text = p.text.strip().lower()
                            if "completed" in p_text:
                                status = "Completed"
                                break
                            elif "in progress" in p_text:
                                status = "In Progress"
                                break
                except:
                    status = "Unknown"
                
                # Percentage (seventh cell)
                percentage = cells[6].text.strip() if len(cells) > 6 else "0%"
                percentage = percentage.replace("%", "").strip()
                
                course_data = {
                    "code": code,
                    "name": name,
                    "assigned_grade": assigned_grade,
                    "status": status,
                    "percentage": percentage
                }
                
                courses.append(course_data)
                log.debug(f"Extracted course: {code} - {name} - Grade: {assigned_grade} - Status: {status} - {percentage}%")
                
            except Exception as e:
                log.warning(f"Error extracting course data from row: {e}")
                continue
        
        log.info(f"Successfully extracted {len(courses)} courses")
        return courses
        
    except Exception as e:
        log.error("Error extracting gradebook data: %s", e)
        return courses


# ============================================================
# MAIN PROCESSING FUNCTION
# ============================================================
def process_student(driver: webdriver.Chrome, student_id: str, cfg: Config) -> Dict:
    """Process a single student's gradebook data."""
    result = {
        "student_id": student_id,
        "student_name": "",
        "courses": [],
        "success": False,
        "error": None
    }
    
    try:
        log.info(f"=== Processing student: {student_id} ===")
        
        # Step 1: Search for student
        search_success, student_name = search_student(driver, student_id, cfg)
        if not search_success:
            result["error"] = "Student search failed"
            return result
        
        result["student_name"] = student_name or f"Student {student_id}"
        
        # Step 2: Open student profile
        if not open_student_profile(driver, student_id, cfg):
            result["error"] = "Could not open student profile"
            return result
        
        # Step 3: Navigate to Gradebook tab
        if not navigate_to_gradebook_tab(driver, cfg):
            result["error"] = "Could not navigate to Gradebook tab"
            # Try to close profile and go back
            try:
                if len(driver.window_handles) > 1:
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
            except:
                pass
            return result
        
        # Step 4: Set items per page to 30
        set_items_per_page(driver, 30, cfg)
        time.sleep(3)  # Wait for data to reload
        
        # Step 5: Extract gradebook data
        courses = extract_gradebook_data(driver, cfg)
        result["courses"] = courses
        
        # Step 6: Close profile tab/window and return to main window
        try:
            if len(driver.window_handles) > 1:
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
            else:
                # Go back to user management
                driver.back()
                time.sleep(2)
        except:
            pass
        
        result["success"] = True
        log.info(f"Successfully processed student {student_id}: {len(courses)} courses found")
        
    except Exception as e:
        result["error"] = str(e)
        log.error(f"Error processing student {student_id}: {e}")
        
        # Try to recover and return to main window
        try:
            if len(driver.window_handles) > 1:
                for i in range(len(driver.window_handles) - 1, 0, -1):
                    driver.switch_to.window(driver.window_handles[i])
                    driver.close()
                driver.switch_to.window(driver.window_handles[0])
            else:
                driver.get(cfg.user_management_url)
        except:
            pass
    
    return result


# ============================================================
# CSV OUTPUT FUNCTIONS
# ============================================================
def ensure_csv_header(path: Path) -> None:
    """Ensure CSV file has proper header."""
    if path.exists():
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([
            "Student ID", "Student Name", "Course Code", "Course Name",
            "Assigned Grade", "Status", "Percentage", "Timestamp"
        ])


def write_result_to_csv(path: Path, result: Dict) -> None:
    """Write a single student's results to CSV."""
    with open(path, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        
        if not result["courses"]:
            writer.writerow([
                result["student_id"],
                result["student_name"],
                "No courses found",
                "",
                "",
                "",
                "",
                timestamp
            ])
        else:
            for course in result["courses"]:
                writer.writerow([
                    result["student_id"],
                    result["student_name"],
                    course.get("code", ""),
                    course.get("name", ""),
                    course.get("assigned_grade", ""),
                    course.get("status", ""),
                    course.get("percentage", ""),
                    timestamp
                ])


# ============================================================
# COMMAND LINE INTERFACE
# ============================================================
def parse_args():
    parser = argparse.ArgumentParser(
        description="Check gradebook records on HSOA LMS and export to CSV/Google Sheets."
    )
    parser.add_argument(
        "--students",
        required=False,
        help="Comma-separated student IDs, or path to a file with one ID per line.",
    )
    parser.add_argument(
        "--output",
        default="gradebook_report.csv",
        help="Output CSV file path (default: gradebook_report.csv).",
    )
    parser.add_argument(
        "--upload-sheets",
        action="store_true",
        help="Upload results to Google Sheets after processing.",
    )
    parser.add_argument(
        "--no-headless",
        action="store_true",
        help="Run browser in non-headless mode (visible).",
    )
    return parser.parse_args()


def load_student_ids(students_arg: str) -> list:
    """Return list of student IDs from a comma-separated string or file path."""
    if not students_arg:
        # Try to get from environment variable
        env_students = os.environ.get("STUDENT_IDS", "")
        if env_students:
            students_arg = env_students
        else:
            return []
    
    # Check if it's a file
    if os.path.exists(students_arg) and os.path.isfile(students_arg):
        try:
            with open(students_arg, "r", encoding="utf-8") as f:
                ids = [line.strip() for line in f if line.strip()]
            return ids
        except Exception as e:
            log.warning(f"Could not read student IDs from file {students_arg}: {e}")
    
    # Split on commas
    ids = [sid.strip() for sid in students_arg.split(",") if sid.strip()]
    return ids


def build_config(args) -> Config:
    """Build configuration from environment variables and arguments."""
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
        output_csv_file=Path(args.output),
        google_credentials_json=os.environ.get("GOOGLE_CREDENTIALS_JSON", ""),
        google_spreadsheet_id=os.environ.get("GOOGLE_SPREADSHEET_ID", ""),
        google_sheet_name=os.environ.get("GOOGLE_SHEET_NAME", "GRADES_PROGRESS"),
        headless_mode=not args.no_headless,
    )


def main():
    """Main function."""
    args = parse_args()
    student_ids = load_student_ids(args.students)
    
    if not student_ids:
        log.error("No student IDs provided. Use --students argument or set STUDENT_IDS environment variable.")
        sys.exit(1)
    
    cfg = build_config(args)
    output_path = cfg.output_csv_file
    
    log.info(f"Processing {len(student_ids)} student(s).")
    log.info(f"Output will be saved to: {output_path}")
    
    # Ensure CSV file has header
    ensure_csv_header(output_path)
    
    # Setup Chrome driver
    driver = setup_chrome_driver(cfg)
    if not driver:
        log.error("Failed to set up Chrome driver.")
        sys.exit(1)
    
    all_results = []
    success_count = 0
    
    try:
        # Step 1: Login
        log.info("Attempting to login...")
        if not login_to_hsoa(driver, cfg):
            log.error("Login failed. Please check credentials.")
            driver.quit()
            sys.exit(1)
        
        # Step 2: Navigate to User Management
        if not navigate_to_user_management(driver, cfg):
            log.error("Failed to navigate to User Management.")
            driver.quit()
            sys.exit(1)
        
        # Step 3: Process each student
        for student_id in student_ids:
            result = process_student(driver, student_id, cfg)
            all_results.append(result)
            
            # Write to CSV
            write_result_to_csv(output_path, result)
            
            if result["success"]:
                success_count += 1
                log.info(f"✓ Student {student_id}: {len(result['courses'])} courses")
            else:
                log.warning(f"✗ Student {student_id} failed: {result.get('error', 'Unknown error')}")
            
            # Small delay between students
            time.sleep(2)
        
        log.info(f"Processing complete. {success_count}/{len(student_ids)} students processed successfully.")
        
        # Step 4: Upload to Google Sheets if requested
        if args.upload_sheets:
            if not cfg.google_spreadsheet_id:
                log.warning("GOOGLE_SPREADSHEET_ID not set; skipping Google Sheets upload.")
            else:
                log.info("Uploading results to Google Sheets...")
                if upload_to_google_sheets(cfg, all_results):
                    log.info("✓ Successfully uploaded to Google Sheets")
                else:
                    log.warning("✗ Failed to upload to Google Sheets")
        
        log.info(f"Results saved to: {output_path}")
        
    except KeyboardInterrupt:
        log.info("Process interrupted by user.")
    except Exception as e:
        log.error(f"Unexpected error: {e}")
        import traceback
        log.error(traceback.format_exc())
    finally:
        driver.quit()
        log.info("Browser closed.")


if __name__ == "__main__":
    main()
