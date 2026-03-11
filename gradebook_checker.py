#!/usr/bin/env python3
import os
import sys
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

# Google Sheets setup
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
SPREADSHEET_ID = os.environ['GOOGLE_SPREADSHEET_ID']

def get_student_ids_from_sheet():
    """Get student IDs from STUDENTS sheet column A2:A"""
    try:
        # Authenticate with Google Sheets
        credentials = Credentials.from_service_account_info(
            eval(os.environ['GOOGLE_CREDENTIALS_JSON']), 
            scopes=SCOPES
        )
        gc = gspread.authorize(credentials)
        
        # Open the spreadsheet and worksheet
        spreadsheet = gc.open_by_key(SPREADSHEET_ID)
        students_sheet = spreadsheet.worksheet('STUDENTS')
        
        # Get all values from column A starting from row 2
        student_ids = students_sheet.col_values(1)[1:]  # Skip header row
        
        # Filter out empty strings and None values
        student_ids = [str(sid).strip() for sid in student_ids if str(sid).strip()]
        
        print(f"Retrieved {len(student_ids)} student IDs from sheet: {student_ids}")
        return student_ids
        
    except Exception as e:
        print(f"Error reading student IDs from sheet: {e}")
        return []

def setup_driver():
    """Setup Chrome driver with options"""
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--window-size=1920,1080')
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

def login_to_hsoa(driver, username, password):
    """Login to HSOA LMS"""
    try:
        driver.get("https://lms.hsoa.edu.om/login/index.php")
        
        # Wait for and fill login form
        username_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "username"))
        )
        password_field = driver.find_element(By.ID, "password")
        login_button = driver.find_element(By.ID, "loginbtn")
        
        username_field.send_keys(username)
        password_field.send_keys(password)
        login_button.click()
        
        # Wait for login to complete
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "usermenu"))
        )
        print("Login successful")
        return True
        
    except Exception as e:
        print(f"Login failed: {e}")
        return False

def get_student_grades(driver, student_id):
    """Get grades for a specific student"""
    try:
        # Navigate to gradebook
        driver.get(f"https://lms.hsoa.edu.om/grade/report/overview/index.php?id={student_id}")
        
        # Wait for grades to load
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "table"))
        )
        
        # Extract grade data
        tables = driver.find_elements(By.TAG_NAME, "table")
        grades_data = []
        
        for table in tables:
            try:
                rows = table.find_elements(By.TAG_NAME, "tr")
                for row in rows[1:]:  # Skip header row
                    cols = row.find_elements(By.TAG_NAME, "td")
                    if len(cols) >= 2:
                        course = cols[0].text.strip()
                        grade = cols[1].text.strip()
                        if course and grade:
                            grades_data.append({
                                'student_id': student_id,
                                'course': course,
                                'grade': grade,
                                'timestamp': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
                            })
            except:
                continue
        
        return grades_data
        
    except Exception as e:
        print(f"Error getting grades for student {student_id}: {e}")
        return []

def update_google_sheet(grades_data):
    """Update Google Sheet with new grades"""
    try:
        credentials = Credentials.from_service_account_info(
            eval(os.environ['GOOGLE_CREDENTIALS_JSON']), 
            scopes=SCOPES
        )
        gc = gspread.authorize(credentials)
        spreadsheet = gc.open_by_key(SPREADSHEET_ID)
        
        # Use GRADES sheet or create if doesn't exist
        try:
            grades_sheet = spreadsheet.worksheet('GRADES')
        except:
            grades_sheet = spreadsheet.add_worksheet(title='GRADES', rows="1000", cols="10")
            # Add headers
            grades_sheet.append_row(['Student ID', 'Course', 'Grade', 'Timestamp'])
        
        # Clear old data but keep headers
        if len(grades_data) > 0:
            # Get existing data to avoid duplicates
            existing_data = grades_sheet.get_all_records()
            existing_records = set()
            for record in existing_data:
                existing_records.add((str(record.get('Student ID', '')), 
                                   str(record.get('Course', '')), 
                                   str(record.get('Grade', ''))))
            
            # Add only new records
            new_rows = []
            for grade in grades_data:
                record_key = (str(grade['student_id']), str(grade['course']), str(grade['grade']))
                if record_key not in existing_records:
                    new_rows.append([grade['student_id'], grade['course'], grade['grade'], grade['timestamp']])
            
            if new_rows:
                grades_sheet.append_rows(new_rows)
                print(f"Added {len(new_rows)} new grade records to sheet")
            else:
                print("No new grade records to add")
        
    except Exception as e:
        print(f"Error updating Google Sheet: {e}")

def main():
    """Main function"""
    print("Starting HSOA LMS Gradebook Checker...")
    
    # Get credentials from environment
    username = os.environ['HSOA_USERNAME']
    password = os.environ['HSOA_PASSWORD']
    
    # Get student IDs from Google Sheet
    student_ids = get_student_ids_from_sheet()
    
    if not student_ids:
        print("No student IDs found. Exiting.")
        return
    
    print(f"Processing {len(student_ids)} students")
    
    # Setup driver
    driver = setup_driver()
    all_grades = []
    
    try:
        # Login
        if not login_to_hsoa(driver, username, password):
            return
        
        # Get grades for each student
        for i, student_id in enumerate(student_ids, 1):
            print(f"Processing student {i}/{len(student_ids)}: {student_id}")
            grades = get_student_grades(driver, student_id)
            all_grades.extend(grades)
            
            # Small delay to avoid overloading the server
            time.sleep(2)
        
        # Update Google Sheet
        if all_grades:
            update_google_sheet(all_grades)
            print(f"Successfully processed {len(all_grades)} grade records")
        else:
            print("No grades were retrieved")
            
    except Exception as e:
        print(f"Error during processing: {e}")
    finally:
        driver.quit()
        print("Gradebook checker completed")

if __name__ == "__main__":
    main()
