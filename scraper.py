import os
import json
import time
from playwright.sync_api import sync_playwright
import gspread
from oauth2client.service_account import ServiceAccountCredentials

def get_lms_courses(student_id):
    """
    Scrape the LMS website to get student courses.
    Extracts Name, Assigned Grade, Percentage, and Total Accumulated
    """
    courses = []
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        
        try:
            # Step 1: Login to LMS
            page.goto("https://hsoa.ordolms.com/login")
            page.fill("input#mat-input-0", "admindias")  # Username field
            page.fill("input#mat-input-1", "Little0809@")  # Password field
            page.click("button[type='submit']")
            
            # Step 2: Navigate to User Management
            page.wait_for_selector("text=manage_accounts", timeout=20000)
            page.click("mat-icon:text('manage_accounts')")
            # Wait for navigation
            page.wait_for_url("**/userManagement**")
            
            # Step 3: Filter by student ID
            page.fill("#mat-input-3", student_id)  # Filter input
            time.sleep(4)  # Wait for table to load
            
            # Step 4: Click settings for the student
            # Find the student row by ID and click the settings icon
            page.click("table tr td:text('{}') + td + td + td + td mat-icon:text('settings')".format(student_id))
            
            # Step 5: Wait and update dropdown to show 30 entries
            page.wait_for_selector(".mat-select", timeout=30000)
            page.click(".mat-select")  # Click the dropdown
            page.click("mat-option:text('30')")  # Select 30 from dropdown
            time.sleep(4)  # Wait for table to update
            
            # Step 6: Extract table data
            # Rows with all required data (8 data points)
            rows = page.locator("tbody tr:has(td)")
            courses_data = []
            
            for i in range(rows.count()):
                if i < 30:  # Safety check to prevent infinite looping
                    course_row = rows.nth(i)
                    
                    # Extract specific columns from the row based on your HTML reference
                    # [Name, Grade Level, Status, Grade, Credits]
                    try:
                        name = course_row.locator("td:nth-child(2)").inner_text()
                        assigned_grade = course_row.locator("td:nth-child(3)").inner_text()
                        percentage = course_row.locator("td:nth-child(7)").inner_text()
                        accumulated = course_row.locator("td:nth-child(8)").inner_text()
                        
                        course_data = {
                            'name': name.strip(),
                            'assigned_grade': assigned_grade.strip(),
                            'percentage': percentage.strip(),
                            'accumulated': accumulated.strip()
                        }
                        courses_data.append(course_data)
                    except Exception as e:
                        # Skip malformed rows
                        continue
                        
            # Create subject entries in Google Sheets format: Name|Status|Grade|Credits
            # For LMS data, percentage is mapped to status, 
            # and if percentage is 100% consider 1 credit
            subjects_list = []
            for course in courses_data:
                name = course["name"]
                percentage = course["percentage"]
                
                # Calculate credit based on percentage
                # 100% completion equals 1 credit
                credit_value = "1.00" if percentage.strip() == "100%" else "0.00"
                
                # Add the course as a formatted subject line
                subject_line = f"{name}|{percentage}|{course['assigned_grade']}|{credit_value}"
                subjects_list.append(subject_line)
                
            # Calculate total credits and save to Google Sheets
            creds_json = os.environ["GCP_CREDS_JSON"]
            creds_dict = json.loads(creds_json)
            
            scope = ["https://spreadsheets.google.com/feeds", 
                     "https://www.googleapis.com/auth/drive"]
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
            client = gspread.authorize(creds)
            
            # Open the student portal spreadsheet
            spreadsheet = client.open_by_key("1uEOB5-MwhlsbD8IrLkhMD6PPx0vBeWKck4HmpB4Lsro")
            sheet = spreadsheet.worksheet("STUDENTS")
            
            # Find the row corresponding to the student ID
            all_cells = sheet.findall(student_id)
            if not all_cells:
                return False, f"No student found with ID: {student_id}"
                
            # Get the first cell match which is likely the correct row
            target_row_idx = -1
            for cell in all_cells:
                if cell.col == 1:  # Check if it's in the ID column (Column A)
                    target_row_idx = cell.row
                    break
                    
            if target_row_idx == -1:
                # If there was no match in column A, just use the first ID found
                target_row_idx = all_cells[0].row
            
            if target_row_idx < 2:  # First row is header, so start from row 2
                return False, "Invalid student row"
                
            # Create the subject entries string
            subjects_text = "\n".join(subjects_list)
            
            # Update the "Subjects" column (column J = Col 10)
            sheet.update_acell(f'J{target_row_idx}', subjects_text)
            
            # Calculate total credits for HSOA finished courses (100% completion)
            hsoa_finished_credits = sum(1.0 for course in courses_data if course["percentage"].strip() == "100%")
            
            # Update HSOA finished credits column with calculated value
            transferred_col = None
            all_headers = sheet.row_values(1)  # First row has headers
            for i, header in enumerate(all_headers):
                if header.lower().strip() == "transferred credits":
                    transferred_col = i + 1  # Column numbers are 1-indexed
                    break
                    
            if transferred_col:
                transferred_val = sheet.cell(target_row_idx, transferred_col).value
                try:
                    transferred_credits = float(transferred_val or 0)
                except ValueError:
                    transferred_credits = 0.0
            else:
                transferred_credits = 0.0
                
            total_credits = transferred_credits + hsoa_finished_credits
            # Create a "Total Credits" column if needed
            all_headers = sheet.row_values(1)
            total_cred_col = None
            for i, header in enumerate(all_headers):
                if header.lower().strip() == "total credits":
                    total_cred_col = i + 1  # Column numbers are 1-indexed
                    break
                    
            if not total_cred_col:
                # Add new column header if it doesn't exist
                sheet.update_acell(f'{chr(ord("A")+len(all_headers))}1', "Total Credits")
                total_cred_col = len(all_headers) + 1
            
            if total_cred_col:
                sheet.update_acell(f'{chr(ord("A")+total_cred_col-1)}{target_row_idx}', total_credits)
                
        except Exception as e:
            print(f"Error during processing: {str(e)}")
        finally:
            browser.close()
            
    return True, f"Updated data for student {student_id}. Found {len(courses_data)} courses"

# Execute the function with passed environment variables
student_id = os.environ.get("STUDENT_ID", "")
success, message = get_lms_courses(student_id)
print(message)
