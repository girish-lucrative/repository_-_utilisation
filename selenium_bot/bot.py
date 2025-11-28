from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (StaleElementReferenceException, 
                                      TimeoutException,
                                      NoSuchElementException)
from datetime import datetime
from datetime import timedelta
from selenium.webdriver.common.action_chains import ActionChains
import time
import os
import logging
from flask import Flask, render_template, request, redirect, url_for, flash
import pandas as pd
import winsound
from openpyxl import Workbook
import os

class CertificateBot:
    def __init__(self, username, password, excel_data, download_folder, process_type):
        self.username = username
        self.password = password
        self.excel_data = excel_data
        self.download_folder = download_folder
        self.process_type = process_type  # 'epcg', 'adv', or 'both'
        self.driver = None
        self.current_index = 0

    def process_all_certificates(self):
        try:
            if not self.excel_data:
                return {"success": False, "message": "No data found"}
            
            # Process based on selected type
            while self.current_index < len(self.excel_data):
                full_login = self.current_index == 0
                result = self._process_certificate(full_login=full_login)
                if not result.get('success'):
                    return result
            
            return {"success": True, "message": "All processes completed successfully"}
                        
        except Exception as e:
            return {"success": False, "message": f"Error processing certificates: {e}"}
        finally:
            self.close_browser()

    def _process_certificate(self, full_login=False):
        try:
            row = self.excel_data[self.current_index]

            if full_login:
                self.start_browser()
                login_result = self.login()
                if not login_result.get('success'):
                    return login_result
            
            # Process based on selected type
            if self.process_type in ['epcg', 'both']:
                epcg_result = self.process_epcg(row)
                if not epcg_result.get('success'):
                    return epcg_result
            
            if self.process_type in ['adv', 'both']:
                adv_result = self.process_adv(row)
                if not adv_result.get('success'):
                    return adv_result
            
            self.current_index += 1
            return {"success": True, "message": f"Processed certificate data"}
            
        except Exception as e:
            return {"success": False, "message": f"Error processing certificate: {e}"}

    def process_epcg(self, row):
        """Process EPCG certificate"""
        try:
            print("Starting EPCG process...")
            wait = WebDriverWait(self.driver, 50)
            
            # Navigate to certificate section
            print("Navigating to EPCG certificate section...")
            time.sleep(1)
            
            # Your existing EPCG processing code here
            my_dashboard = wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="200928"]/a'))
            )
            my_dashboard.click()
            print("✅ Clicked on my dashboard")
            
            # Continue with your existing EPCG code...
            # (Include all the EPCG-specific code from your original fill_certificate method)

            repo = wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="98000023"]/a'))
            )
            repo.click()
            print("✅ Clicked on repositories")
            time.sleep(1)
            self.driver.execute_script("document.body.style.zoom='80%'")
            time.sleep(0.5)
            self.driver.execute_script("document.body.style.zoom='60%'")
            time.sleep(0.5)

            Bill_repo = wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="content"]/div[3]/div/div[2]/div[1]/div/a'))
            )
            Bill_repo.click()
            print("✅ Clicked on Bill repositories")
            
            for i in range(2):
                
                wait = WebDriverWait(self.driver, 100)
                dropdown = wait.until(EC.element_to_be_clickable((By.ID, "txt_selectBill")))
                dropdown.click()
                time.sleep(1)
                if i==0:
                
                   # Select by visible text using XPath
                   option = wait.until(EC.element_to_be_clickable((By.XPATH, "//option[text()='Shipping Bill']")))
                   option.click()
                   time.sleep(2)
                   print("select shipping bill")
                else:
                    # Select by visible text using XPath
                   option = wait.until(EC.element_to_be_clickable((By.XPATH, "//option[text()='Bill of Entry']")))
                   option.click()
                   time.sleep(2)
                   print("select Bill of Entry")

                for idx, row in enumerate(self.excel_data):
                    print(row)
                
                    # --- Fix: handle pandas Timestamp directly ---
                    shipping_date = row.get("EPCG Shipping Bill Date")
                    if pd.isna(shipping_date):
                        print(f"Skipping Row {idx + 1}: No Shipping Bill Date found")
                        continue
                
                    # Ensure it's a datetime object
                    if isinstance(shipping_date, pd.Timestamp):
                        shipping_date = shipping_date.to_pydatetime()
                
                    shipping_bill_date = shipping_date.strftime("%d/%m/%Y")
                    print("Formatted Shipping Bill Date:", shipping_bill_date)
                
                    # --- Wait for date field and fill it ---
                    wait.until(EC.presence_of_element_located((By.ID, "fromDateOfSelectedBil")))
                    print("shown shipping bill date")
                
                    self.driver.execute_script(f"""
                        var fromDate = document.getElementById('fromDateOfSelectedBil');
                        fromDate.value = '{shipping_bill_date}';
                        fromDate.dispatchEvent(new Event('change'));
                    """)
                
                    time.sleep(1)
                    
                
                    # --- Fill Authorisation Number ---
                    auth_no = str(row.get("EPCG Authorisation Number", "")).strip()
                    if i==0:
                        self.driver.find_element(By.ID, "authorisationNo").clear()
                        self.driver.find_element(By.ID, "authorisationNo").send_keys(auth_no)
                    else:
                        self.driver.find_element(By.ID, "boeLicenseNumber").clear()
                        self.driver.find_element(By.ID, "boeLicenseNumber").send_keys(auth_no)
                    
                    print("add authorisation number")
                
                    # --- Click Search ---
                    search = wait.until(EC.element_to_be_clickable((By.ID, 'repSearchBtn')))
                    search.click()
                    print("✅ Clicked search")
                    if i==0:

                        # --- Wait for table to appear ---
                        wait.until(EC.presence_of_element_located((By.ID, "billRepositoryTable")))
                        print("table shown")
                        
                        all_rows = []
                        a=1
                        while True:
                            
                            # Wait for the table body to load
                            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#billRepositoryTable tbody tr")))
                        
                            rows = self.driver.find_elements(By.CSS_SELECTOR, "#billRepositoryTable tbody tr")
                            print(f"Found {len(rows)} rows on this page - {a}")
                            
                        
                            for r in rows:
                                cols = r.find_elements(By.TAG_NAME, "td")
                                if len(cols) > 1:  # skip "No data available" row
                                    data = [c.text.strip() for c in cols]
                                    all_rows.append(data)
                        
                            # Try to find "Next" button and check if it is enabled
                            try:
                                self.driver.execute_script("window.scrollBy(0, -300);")
                                time.sleep(0.5)
                                self.driver.execute_script("window.scrollBy(0, 300);")
                                time.sleep(0.5)
                                wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="billRepositoryTable_next"]/a')))
                                next_btn = self.driver.find_element(By.XPATH, '//*[@id="billRepositoryTable_next"]/a')
                                time.sleep(1)
                                next_class = next_btn.get_attribute("class")
                                if "disabled" in next_class:
                                    print("✅ No more pages. Extraction complete.")
                                    break
                                else:
                                    self.driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
                                    next_btn.click()
                                    time.sleep(2)
                                    a+=1  # Wait for next page data to load
                            except Exception as e:
                                print("⚠️ Pagination ended or not found:", e)
                                if len(rows)==10:
                                    popup_click = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="CancelOk"]')))
                                    popup_click.click()
                                    print("popup handled")
                                    time.sleep(1)
                                    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="billRepositoryTable_next"]/a')))
                                    
                                    next_btn = self.driver.find_element(By.XPATH, '//*[@id="billRepositoryTable_next"]/a')
                                    next_class = next_btn.get_attribute("class")
                                    if "disabled" in next_class:
                                        print("✅ No more pages. Extraction complete.")
                                        break
                                    else:
                                        self.driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
                                        next_btn.click()
                                        time.sleep(2)
                                        a+=1
                                    continue
                                else:
                                    print("all data done")
                                    break
                        
                        # --- Extract headers ---
                        headers = [h.text.strip() for h in self.driver.find_elements(By.CSS_SELECTOR, "#billRepositoryTable thead th")]
                        
                        # --- Save to Excel ---
                        # --- Save to Local Downloads Folder ---
                        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"EPCG_shipping_bill_Data_{timestamp}.xlsx"
                        file_path = os.path.join(downloads_path, filename)                    
    
                        df = pd.DataFrame(all_rows, columns=headers)
                        df.to_excel(file_path, index=False)                    
    
                        print(f"💾 Data saved to: {file_path}")
                    else:
                                                    # --- Wait for Bill of Entry table to load ---
                        wait.until(EC.presence_of_element_located((By.ID, "billOfEntryTable")))
                        
                        all_rows = []
                        b=1
                        while True:
                            # Wait for "Processing..." to disappear
                            try:
                                wait.until_not(EC.visibility_of_element_located((By.ID, "billOfEntryTable_processing")))
                            except:
                                pass
                        
                            # Wait for table rows
                            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#billOfEntryTable tbody tr")))
                        
                            rows = self.driver.find_elements(By.CSS_SELECTOR, "#billOfEntryTable tbody tr")
                            print(f"Found {len(rows)} rows on this page {b}")
                        
                            for r in rows:
                                cols = r.find_elements(By.TAG_NAME, "td")
                                if len(cols) > 1:  # skip "No data available"
                                    data = [c.text.strip() for c in cols]
                                    all_rows.append(data)
                        
                            # --- Handle Pagination ---
                            try:
                                self.driver.execute_script("window.scrollBy(0, -300);")
                                time.sleep(0.5)
                                self.driver.execute_script("window.scrollBy(0, 300);")
                                time.sleep(0.5)
                                wait.until(EC.presence_of_element_located((By.ID, "billOfEntryTable_next")))
                                next_btn = self.driver.find_element(By.ID, "billOfEntryTable_next")
                                time.sleep(1)
                                next_class = next_btn.get_attribute("class")
                                
                                if "disabled" in next_class:
                                    print("✅ No more pages. Extraction complete.")
                                    break
                                else:
                                    self.driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
                                    next_btn.click()
                                    b+=1
                                    time.sleep(2)  # allow table to refresh
                            except Exception as e:
                                print("⚠️ Pagination ended or not found:", e)
                                if len(rows)==10:
                                    popup_click = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="CancelOk"]')))
                                    popup_click.click()
                                    print("popup handled")
                                    wait.until(EC.presence_of_element_located((By.ID, "billOfEntryTable_next")))
                                    next_btn = self.driver.find_element(By.ID, "billOfEntryTable_next")
                                    time.sleep(1)
                                    next_class = next_btn.get_attribute("class")
                                    
                                    if "disabled" in next_class:
                                        print("✅ No more pages. Extraction complete.")
                                        break
                                    else:
                                        self.driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
                                        next_btn.click()
                                        b+=1
                                        time.sleep(2)
                                    continue
                                else:
                                    print("all data done")
                                    break
                        
                        # --- Extract Table Headers ---
                        headers = [h.text.strip() for h in self.driver.find_elements(By.CSS_SELECTOR, "#billOfEntryTable thead th")]
                        
                        # --- Save to Local Downloads Folder ---
                        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"EPCG_BillOfEntry_Data_{timestamp}.xlsx"
                        file_path = os.path.join(downloads_path, filename)
                        
                        df = pd.DataFrame(all_rows, columns=headers)
                        df.to_excel(file_path, index=False)
                        
                        print(f"💾 Data saved to: {file_path}")
            time.sleep(1)
            self.driver.refresh()  
            time.sleep(1)
            AI = wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="element-to-hide"]/div[20]/div/div/a[1]/span'))
            )
            AI.click()
            print("✅ Clicked on AI")
            repo = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='EPCG']"))
            )
            repo.click()
            print("✅ Clicked on EPCG")
            repo = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//h5[normalize-space()='Closure of EPCG/Issuance of Post Export Scrip']"))
            )
            repo.click()
            print("✅ Clicked on closure of EPCG")
            time.sleep(1)
            repo = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[@id='btnNewApp']"))
            )
            repo.click()
            print("Start fresh application")
            wait = WebDriverWait(self.driver, 100)
            dropdown = wait.until(EC.element_to_be_clickable((By.ID, "applicationFor")))
            dropdown.click()
            option = wait.until(EC.element_to_be_clickable((By.XPATH, "//option[text()='REDEMPTION']")))
            option.click()
            time.sleep(1)
            auth_closure = wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="custom-accordion"]/div[2]/div[1]/a'))
            )
            auth_closure.click()
            print("Auth closure")
            
            
            
            # EPCG Authorization Number to match
            target_auth_no = row.get("EPCG Authorisation Number")
            
            # Wait for the table to be visible
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.dataTables_scrollBody")))
            
            found = False
            page = 1
            
            while True:
                try:
                    # Find the cell containing the EPCG Authorization Number
                    cell_xpath = f"//div[@class='dataTables_scrollBody']//table//td[normalize-space()='{target_auth_no}']"
                    target_cell = self.driver.find_element(By.XPATH, cell_xpath)
            
                    # Scroll the cell into view
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", target_cell)
                    time.sleep(1)
            
                    # Get the parent row and click its radio button
                    parent_row = target_cell.find_element(By.XPATH, "./ancestor::tr")
                    radio_button = parent_row.find_element(By.CSS_SELECTOR, "input[type='radio']")
                    self.driver.execute_script("arguments[0].click();", radio_button)
            
                    print(f"✅ Selected EPCG Authorization Number: {target_auth_no}")
                    found = True
                    break
            
                except Exception:
                    # Check for next page if not found
                    try:
                        self.driver.execute_script("window.scrollBy(0, -100);")
                        self.driver.execute_script("window.scrollBy(0, 100);")
                        wait.until(EC.presence_of_element_located((By.ID, "epcgauthTbl_next")))
                        next_button = self.driver.find_element(By.ID, "epcgauthTbl_next")
                        if "disabled" in next_button.get_attribute("class"):
                            break  # No more pages left
                        self.driver.execute_script("arguments[0].click();", next_button)
                        page += 1
                        print(f"🔁 Searching on page {page}...")
                        time.sleep(2)
                    except:
                        break
            
            if not found:
                print(f"❌ Target EPCG Authorization Number {target_auth_no} not found in any page.")
            
            prd_val = wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="eoPending"]/div/div/div[2]/div/div/div[1]/label'))
            )
            prd_val.click()
            nxt = wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="btnNext"]'))
            )
            nxt.click()
            time.sleep(1)
            self.driver.execute_script("document.body.style.zoom='80%'")
            time.sleep(0.5)
            self.driver.execute_script("document.body.style.zoom='60%'")
            time.sleep(0.5)
            
            # ii = wait.until(
            #     EC.element_to_be_clickable((By.XPATH, '//*[@id="closureAuthorisationDetails"]/div[2]/div[1]/a'))
            # )
            # ii.click()
            # print("click import ")
            time.sleep(1)
            self.driver.execute_script("window.scrollBy(0, 300);")
            ja = wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="closureAuthorisationDetails"]/div[4]/div[1]/a'))
            )
            ja.click()
            print("click juri address ")
            jac = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//textarea[@id='jurisdictionalAddress']"))
            )
            jac.clear()  # optional, clears existing text
            jac.send_keys("-")
            print("Sent '-' to jurisdictionalAddress field")
            save_next = wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="btnNext"]'))
            )
            save_next.click()
            print("click save and next  ")
            time.sleep(2)
            # self.driver.execute_script("document.body.style.zoom='80%'")
            # time.sleep(0.5)
            # self.driver.execute_script("document.body.style.zoom='60%'")
            # time.sleep(0.5)
            
            # Scroll to the bottom of the page
            wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="custom-accordion"]/div/div[1]/div[1]/a')))
            print("wait")
            
            self.driver.execute_script("window.scrollBy(0, 500);")
                            
            try:
                print("=== STARTING TABLE EXTRACTION ===")
                print("Setting entries per page to 100 for import table...")
                try:
                    # Find the dropdown for entries per page
                    entries_dropdown = self.driver.find_element(By.NAME, "closureImportItemTbl_length")
                    select = Select(entries_dropdown)
                    
                    # Select 100 entries
                    select.select_by_value("100")
                    print("✓ Selected 100 entries per page for import table")
                    
                    # Wait for table to reload with more entries
                    time.sleep(3)
                    print("Waiting for import table to reload with 100 entries...")
                    
                    # Wait for the table to refresh
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#closureImportItemTbl tbody tr")))
                    time.sleep(2)
                    
                except Exception as e:
                    print(f"Could not set entries to 100 for import table: {e}")
                
                time.sleep(5)
                print("Page loaded, waiting for table...")
                
                # Wait for table to be present
                table = wait.until(EC.presence_of_element_located((By.ID, "closureImportItemTbl")))
                print("✓ Main table found")
                
                # PRECISE HEADER EXTRACTION - Based on your HTML structure
                print("Extracting headers precisely...")
                
                # Method 1: Extract from the specific header structure in your HTML
                headers = []
                
                # Get the main header rows
                header_rows = self.driver.find_elements(By.CSS_SELECTOR, "#closureImportItemTbl thead tr")
                print(f"Found {len(header_rows)} header rows")
                
                # Extract from first header row (main headers)
                first_row_headers = header_rows[0].find_elements(By.TAG_NAME, "th")
                print(f"First row has {len(first_row_headers)} headers")
                
                for i, header in enumerate(first_row_headers):
                    text = header.text.strip()
                    colspan = header.get_attribute("colspan")
                    rowspan = header.get_attribute("rowspan")
                    
                    print(f"Header {i+1}: '{text}' (colspan: {colspan}, rowspan: {rowspan})")
                    
                    if text:
                        if colspan and int(colspan) > 1:
                            # This is a group header - we'll handle the subheaders separately
                            print(f"  -> Group header spanning {colspan} columns: {text}")
                        elif rowspan and int(rowspan) > 1:
                            # This header spans multiple rows
                            headers.append(text)
                        else:
                            headers.append(text)
                
                # Extract from second header row (sub-headers)
                if len(header_rows) > 1:
                    second_row_headers = header_rows[1].find_elements(By.TAG_NAME, "th")
                    print(f"Second row has {len(second_row_headers)} sub-headers")
                    
                    for i, header in enumerate(second_row_headers):
                        text = header.text.strip()
                        if text:
                            headers.append(text)
                            print(f"Sub-header {i+1}: '{text}'")
                
                print(f"Headers after two-row extraction: {headers}")
                print(f"Total headers: {len(headers)}")
                
                # Method 2: Direct extraction from all visible th elements (alternative approach)
                if len(headers) < 10:  # If we don't have enough headers, try alternative
                    print("\nTrying alternative header extraction...")
                    headers = []
                    
                    # Get all th elements and filter properly
                    all_ths = self.driver.find_elements(By.CSS_SELECTOR, "#closureImportItemTbl thead th")
                    
                    for i, th in enumerate(all_ths):
                        text = th.text.strip()
                        if text and len(text) > 1:  # Filter out empty and very short texts
                            # Skip group headers that don't represent actual data columns
                            if not any(phrase in text for phrase in ["Details of", "Details Of"]):
                                headers.append(text)
                                print(f"Header {i+1}: '{text}'")
                    
                    print(f"Alternative headers: {headers}")
                
                # Method 3: Extract based on data column count and known header structure
                if not headers or len(headers) < 10:
                    print("\nUsing known header structure from HTML...")
                    # Based on your HTML structure, these are the expected headers:
                    expected_headers = [
                        "SNo.", "SNo. Of Item", "ITC(HS) Code", "Description of Capital goods to be Imported",
                        "Quantity Imported", "Unit of measure", "Quantity Invalidated", "Unit of measure (Invalidated)",
                        "Whether Capital goods is restricted for import", "Installation Certificate No.", 
                        "Installation Certificate Date", "Bill of Entry No./GST Invoice No./Invoice No.",
                        "Bill of Entry Date/GST Invoice Date/Invoice Date", "Supply Date", "Invoice Serial No.",
                        "CIF value of imports/deemed imports (INR)", "Duty saved amount (INR)", "Duty saved value (in USD)",
                        "Edit/ Delete"
                    ]
                    
                    # Check how many columns we actually have in data
                    try:
                        sample_row = self.driver.find_element(By.CSS_SELECTOR, "#closureImportItemTbl tbody tr")
                        actual_columns = len(sample_row.find_elements(By.TAG_NAME, "td"))
                        print(f"Actual data columns: {actual_columns}")
                        
                        if actual_columns <= len(expected_headers):
                            headers = expected_headers[:actual_columns]
                            print(f"Using {actual_columns} headers from expected list")
                        else:
                            headers = expected_headers
                            print("Using all expected headers")
                    except:
                        headers = expected_headers[:19]  # Use first 19 as default
                        print("Using default expected headers")
                
                print(f"\nFINAL HEADERS ({len(headers)}):")
                for i, header in enumerate(headers):
                    print(f"  {i+1}. {header}")
                
                # DATA EXTRACTION
                print("\n=== STARTING DATA EXTRACTION ===")
                all_data = []
                page_count = 1
                
                while True:
                    print(f"--- Processing Page {page_count} ---")
                    
                    # Wait for table
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#closureImportItemTbl tbody")))
                    time.sleep(2)
                    
                    # Get current page rows
                    current_rows = self.driver.find_elements(By.CSS_SELECTOR, "#closureImportItemTbl tbody tr")
                    print(f"Found {len(current_rows)} rows on page {page_count}")
                    
                    if not current_rows:
                        print("No rows found, stopping...")
                        break
                    
                    page_data = []
                    for row in current_rows:
                        cells = row.find_elements(By.TAG_NAME, "td")
                        row_data = [cell.text.strip() for cell in cells]
                        
                        if row_data and any(cell for cell in row_data if cell):  # Only add non-empty rows
                            page_data.append(row_data)
                    
                    all_data.extend(page_data)
                    print(f"Added {len(page_data)} rows from page {page_count}")
                    
                    # Check for next page
                    try:
                        self.driver.execute_script("window.scrollBy(0, -300);")
                        time.sleep(0.5)
                        self.driver.execute_script("window.scrollBy(0, 300);")
                        time.sleep(0.5)
                        wait.until(EC.presence_of_element_located((By.ID, "closureImportItemTbl_next")))
                        time.sleep(1)
                        next_button = self.driver.find_element(By.ID, "closureImportItemTbl_next")
                        if "disabled" in next_button.get_attribute("class"):
                            print("Last page reached")
                            break
                        else:
                            print("Going to next page...")
                            self.driver.execute_script("arguments[0].click();", next_button)
                            time.sleep(3)
                            page_count += 1
                            
                            if page_count > 20:
                                print("Safety limit reached")
                                break
                                
                    except NoSuchElementException:
                        print("No next button found")
                        break
                
                print(f"\nExtraction complete: {len(all_data)} total rows from {page_count} pages")
                
                # CREATE DATAFRAME AND SAVE
                if all_data:
                    # Ensure headers match data columns
                    if len(headers) != len(all_data[0]):
                        print(f"Adjusting headers: {len(headers)} -> {len(all_data[0])}")
                        if len(headers) > len(all_data[0]):
                            headers = headers[:len(all_data[0])]
                        else:
                            # Use the headers we have and add generic for missing ones
                            base_headers = headers.copy()
                            for i in range(len(headers), len(all_data[0])):
                                base_headers.append(f"Column_{i+1}")
                            headers = base_headers
                    
                    df = pd.DataFrame(all_data, columns=headers)
                    
                    # Save to Downloads
                    downloads_path = os.path.expanduser("~/Downloads")
                    timestamp = time.strftime("%Y%m%d_%H%M%S")
                    filename = os.path.join(downloads_path, f"EPCG_itemOfImport_{timestamp}.xlsx")
                    
                    df.to_excel(filename, index=False, engine='openpyxl')
                    print(f"✓ SUCCESS: File saved to: {filename}")
                    
                    print("\nFirst 2 rows with actual headers:")
                    print(df.head(2))
                    
                else:
                    print("✗ No data extracted")
            
            except Exception as e:
                print(f"✗ ERROR: {e}")
                import traceback
                traceback.print_exc()
    
            # Items of Export 
            le = wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="custom-accordion"]/div/div[2]/div[1]/a'))
            )
            le.click()
            try:
                print("=== STARTING EXPORT TABLE EXTRACTION ===")
                print("Setting entries per page to 100...")
                try:
                    # Find the dropdown for entries per page
                    entries_dropdown = self.driver.find_element(By.NAME, "closureExportItemTbl_length")
                    select = Select(entries_dropdown)
                    
                    # Select 100 entries
                    select.select_by_value("100")
                    print("✓ Selected 100 entries per page")
                    
                    # Wait for table to reload with more entries
                    time.sleep(3)
                    print("Waiting for table to reload with 100 entries...")
                    
                    # Wait for the table to refresh
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#closureExportItemTbl tbody tr")))
                    time.sleep(2)
                    
                except Exception as e:
                    print(f"Could not set entries to 100: {e}")
            
                # Wait for export table to be present
                table = wait.until(EC.presence_of_element_located((By.ID, "closureExportItemTbl")))
                print("✓ Export table found")
                
                # EXTRACT ACTUAL HEADERS FOR EXPORT TABLE
                print("Extracting actual headers for export table...")
                
                headers = []
                
                # Get all header rows from export table
                header_rows = self.driver.find_elements(By.CSS_SELECTOR, "#closureExportItemTbl thead tr")
                print(f"Found {len(header_rows)} header rows")
                
                # Process first row headers (main headers)
                if len(header_rows) >= 1:
                    first_row_headers = header_rows[0].find_elements(By.TAG_NAME, "th")
                    print(f"First row has {len(first_row_headers)} headers")
                    
                    for i, header in enumerate(first_row_headers):
                        text = header.text.strip()
                        colspan = header.get_attribute("colspan")
                        rowspan = header.get_attribute("rowspan")
                        
                        print(f"Header {i+1}: '{text}' (colspan: {colspan}, rowspan: {rowspan})")
                        
                        if text:
                            # Skip group headers that span multiple columns
                            if colspan and int(colspan) > 1:
                                print(f"  -> Group header: {text} (spans {colspan} columns)")
                                # Don't add group headers as individual columns
                            elif rowspan and int(rowspan) > 1:
                                # This is a regular header spanning multiple rows
                                headers.append(text)
                            else:
                                headers.append(text)
                
                # Process second row headers (sub-headers)
                if len(header_rows) >= 2:
                    second_row_headers = header_rows[1].find_elements(By.TAG_NAME, "th")
                    print(f"Second row has {len(second_row_headers)} sub-headers")
                    
                    for i, header in enumerate(second_row_headers):
                        text = header.text.strip()
                        if text:
                            headers.append(text)
                            print(f"Sub-header {i+1}: '{text}'")
                
                print(f"Headers after extraction: {headers}")
                print(f"Total headers: {len(headers)}")
                
                # If we don't have enough headers, use known structure from HTML
                if len(headers) < 15:
                    print("\nUsing known header structure from HTML...")
                    expected_headers = [
                        "SNo.", "Select", "SNo. Of Item", "ITC(HS) Code/Service Code", 
                        "Description of the Item", "ITC (HS) Code of the alternate Product", 
                        "Description of the Alternate Product Item", "Type of Export", "EO Block Period",
                        "Shipping Bill No. / Bill of Export", "Port code of registration", 
                        "Shipping Bill Date", "Invoice No.", "Invoice Date", "Invoice Serial No.",
                        "FOB Value/FOR Value (in FC)", "FOB Value/ FOR value (in USD normalized)", 
                        "FOB Value/FOR Value (in INR)", "Foreign Currency", "Exchange rate of FC to INR",
                        "ECGC Claimed?", "Edit/ Delete"
                    ]
                    
                    # Check actual data columns
                    try:
                        sample_row = self.driver.find_element(By.CSS_SELECTOR, "#closureExportItemTbl tbody tr")
                        actual_columns = len(sample_row.find_elements(By.TAG_NAME, "td"))
                        print(f"Actual data columns: {actual_columns}")
                        
                        if actual_columns <= len(expected_headers):
                            headers = expected_headers[:actual_columns]
                        else:
                            headers = expected_headers
                    except:
                        headers = expected_headers
                
                print(f"\nFINAL HEADERS ({len(headers)}):")
                for i, header in enumerate(headers):
                    print(f"  {i+1}. {header}")
                
                # DATA EXTRACTION
                print("\n=== STARTING DATA EXTRACTION ===")
                all_data = []
                page_count = 1
                
                while True:
                    print(f"--- Processing Page {page_count} ---")
                    
                    # Wait for table
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#closureExportItemTbl tbody")))
                    time.sleep(2)
                    
                    # Get current page rows
                    current_rows = self.driver.find_elements(By.CSS_SELECTOR, "#closureExportItemTbl tbody tr")
                    print(f"Found {len(current_rows)} rows on page {page_count}")
                    
                    if not current_rows:
                        print("No rows found, stopping...")
                        break
                    
                    page_data = []
                    for row_idx, row in enumerate(current_rows):
                        cells = row.find_elements(By.TAG_NAME, "td")
                        row_data = []
                        
                        for cell in cells:
                            # Check if cell has links (like Shipping Bill numbers)
                            links = cell.find_elements(By.TAG_NAME, "a")
                            if links:
                                cell_text = links[0].text.strip()
                            else:
                                cell_text = cell.text.strip()
                            row_data.append(cell_text)
                        
                        if row_data and any(cell for cell in row_data if cell):  # Only add non-empty rows
                            page_data.append(row_data)
                            
                            # Show first row sample
                            if row_idx == 0 and page_count == 1:
                                print(f"First row sample: {row_data[:5]}...")  # Show first 5 cells
                    
                    all_data.extend(page_data)
                    print(f"Added {len(page_data)} rows from page {page_count}")
                    
                    # Check for next page
                    try:
                        self.driver.execute_script("window.scrollBy(0, -100);")
                        self.driver.execute_script("window.scrollBy(0, 100);")
                        wait.until(EC.presence_of_element_located((By.ID, "closureExportItemTbl_next")))
                        next_button = self.driver.find_element(By.ID, "closureExportItemTbl_next")
                        if "disabled" in next_button.get_attribute("class"):
                            print("Last page reached")
                            break
                        else:
                            print("Going to next page...")
                            # Scroll to button and click
                            self.driver.execute_script("arguments[0].scrollIntoView();", next_button)
                            time.sleep(1)
                            self.driver.execute_script("arguments[0].click();", next_button)
                            
                            # Wait for page to load
                            time.sleep(3)
                            page_count += 1
                            
                            # # Safety limit
                            # if page_count > 50:
                            #     print("Safety limit reached - stopping after 50 pages")
                            #     break
                                
                    except NoSuchElementException:
                        print("No next button found")
                        break
                    except Exception as e:
                        print(f"Error clicking next button: {e}")
                        break
                
                print(f"\nExtraction complete: {len(all_data)} total rows from {page_count} pages")
                
                # CREATE DATAFRAME AND SAVE
                if all_data:
                    # Ensure headers match data columns
                    if len(headers) != len(all_data[0]):
                        print(f"Adjusting headers: {len(headers)} -> {len(all_data[0])}")
                        if len(headers) > len(all_data[0]):
                            headers = headers[:len(all_data[0])]
                        else:
                            # Use the headers we have and add generic for missing ones
                            base_headers = headers.copy()
                            for i in range(len(headers), len(all_data[0])):
                                base_headers.append(f"Column_{i+1}")
                            headers = base_headers
                    
                    df = pd.DataFrame(all_data, columns=headers)
                    
                    # Save to Downloads
                    downloads_path = os.path.expanduser("~/Downloads")
                    timestamp = time.strftime("%Y%m%d_%H%M%S")
                    filename = os.path.join(downloads_path, f"EPCG_ItemsOfExport_{timestamp}.xlsx")
                    
                    df.to_excel(filename, index=False, engine='openpyxl')
                    print(f"✓ SUCCESS: Export table data saved to: {filename}")
                    
                    # print("\nFirst 3 rows with actual headers:")
                    # print(df.head(3))
                    
                    # # Show column information
                    # print(f"\nDataFrame shape: {df.shape}")
                    # print("Column names:")
                    # for col in df.columns:
                    #     print(f"  - {col}")
                    
                else:
                    print("✗ No data extracted from export table")
            
            except Exception as e:
                print(f"✗ ERROR: {e}")
                import traceback
                traceback.print_exc()
            time.sleep(1)
            save_next3 = wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="btnNext"]'))
            )
            save_next3.click()
            time.sleep(1)
            try:
                alt_ok = wait.until(
                    EC.element_to_be_clickable((By.XPATH, '//button[contains(@class, "bootbox-accept")]'))
                )
                alt_ok.click()
                time.sleep(1)
            
            except Exception as e:
                print("Retrying OK click…", e)
                alt_ok = wait.until(
                    EC.element_to_be_clickable((By.XPATH, '//button[contains(@class, "bootbox-accept")]'))
                )
                alt_ok.click()
                time.sleep(1)
            
            wait = WebDriverWait(self.driver, 100)
            
            print("\n=== STARTING DATA EXTRACTION ===")
            
            all_data = []
            page_count = 1
            
            
            # ======================================================
            #  FIXED COLUMN HEADERS (YOUR REQUIRED EXACT ORDER)
            # ======================================================
            final_headers = [
                "SI.NO.",
                "Select",
                "Shipping Bill/Invoice No.",
                "Port of registration",
                "Shipping Bill/Invoice Date",
                "EDI/Non-EDI SB",
                "ECGC Reimbursed?",
                "eBRC/FIRC Number",
                "Date of Realisation",
                "Realised Value (in FC)",
                "Foreign Currency",
                "Exchange rate of FC to INR",
                "Realised Value (in INR)",
                "Realised Value (in USD)",
                "Bank Name",
                "IFSC Code",
                "BRC Type",
                "Edit/ Delete"
            ]
            
            print("\nUsing fixed correct headers:")
            print(final_headers)
            
            
            
            # ======================================================
            #                 START PAGINATION
            # ======================================================
            while True:
                print(f"\n--- Processing Page {page_count} ---")
            
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#sbBankRealiseTbl tbody")))
                time.sleep(2)
            
                current_rows = self.driver.find_elements(By.CSS_SELECTOR, "#sbBankRealiseTbl tbody tr")
                print(f"Found {len(current_rows)} rows on page {page_count}")
            
                if not current_rows:
                    print("No rows found, stopping...")
                    break
            
                page_data = []
            
                for row_idx, row in enumerate(current_rows):
            
                    # Capture ALL <td> including hidden (display:none)
                    cells = row.find_elements(By.TAG_NAME, "td")
                    row_data = []
            
                    for cell in cells:
                        links = cell.find_elements(By.TAG_NAME, "a")
                        if links:
                            cell_text = links[0].text.strip()
                        else:
                            cell_text = cell.text.strip()
                        row_data.append(cell_text)
            
                    # add only non-empty rows
                    if any(cell.strip() for cell in row_data):
                        page_data.append(row_data)
            
                        if row_idx == 0 and page_count == 1:
                            print("Sample row:", row_data)
            
                all_data.extend(page_data)
                print(f"Added {len(page_data)} rows from page {page_count}")
            
            
            
                # ======================================================
                #       NEXT PAGE BUTTON CLICK
                # ======================================================
                try:
                    next_button = self.driver.find_element(By.ID, "sbBankRealiseTbl_next")
            
                    if "disabled" in next_button.get_attribute("class"):
                        print("Last page reached.")
                        break
            
                    print("Going to next page...")
                    self.driver.execute_script("arguments[0].scrollIntoView();", next_button)
                    time.sleep(1)
                    self.driver.execute_script("arguments[0].click();", next_button)
            
                    time.sleep(3)
                    page_count += 1
            
                    if page_count > 50:
                        print("Safety limit reached… stopping after 50 pages.")
                        break
            
                except Exception as e:
                    print(f"Error clicking next button: {e}")
                    break
            
            
            
            # ======================================================
            #             DATA EXTRACTION FINISHED
            # ======================================================
            print(f"\nExtraction complete: {len(all_data)} total rows from {page_count} pages")
            
            
            
            # ======================================================
            #             BUILD DATAFRAME & SAVE
            # ======================================================
            if all_data:
                # Ensure correct number of columns
                row_len = len(all_data[0])
                header_len = len(final_headers)
            
                if row_len != header_len:
                    print(f"\n⚠ Column count mismatch! Data has {row_len}, headers have {header_len}")
            
                    if header_len > row_len:
                        final_headers = final_headers[:row_len]
                    else:
                        for i in range(header_len, row_len):
                            final_headers.append(f"Column_{i+1}")
            
                # Create DataFrame
                df = pd.DataFrame(all_data, columns=final_headers)
            
                # Save
                downloads_path = os.path.expanduser("~/Downloads")
                timestamp = time.strftime("%Y%m%d_%H%M%S")
                filename = os.path.join(downloads_path, f"EPCG_RealisationDetails_{timestamp}.xlsx")
            
                df.to_excel(filename, index=False, engine='openpyxl')
            
                print(f"\n✓ SUCCESS: Bank realization table data saved to: {filename}")
            
            time.sleep(1)
            
            return {"success": True, "message": "EPCG process completed"}
            
        except Exception as e:
            return {"success": False, "message": f"Error in EPCG process: {e}"}

    def process_adv(self, row):
        """Process ADV certificate"""
        try:
            print("Starting ADV process...")
            wait = WebDriverWait(self.driver, 50)
            
            # Navigate to ADV section
            print("Navigating to ADV certificate section...")
            time.sleep(1)
            
            # Your existing ADV processing code here
            my_dashboard = wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="200928"]/a'))
            )
            my_dashboard.click()
            print("✅ Clicked on my dashboard")
            
            # Continue with your existing ADV code...
            # (Include all the ADV-specific code from your original fill_certificate method)
            repo = wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="98000023"]/a'))
            )

            repo.click()
            print("✅ Clicked on repositories")

            Bill_repo = wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="content"]/div[3]/div/div[2]/div[1]/div/a'))
            )
            Bill_repo.click()
            print("✅ Clicked on Bill repositories")
            
            for i in range(2):
                print(i)
                
                wait = WebDriverWait(self.driver, 15)
                dropdown = wait.until(EC.element_to_be_clickable((By.ID, "txt_selectBill")))
                dropdown.click()
                time.sleep(1)
                if i==0:
                
                   # Select by visible text using XPath
                   option = wait.until(EC.element_to_be_clickable((By.XPATH, "//option[text()='Shipping Bill']")))
                   option.click()
                   time.sleep(2)
                   print("select shipping bill")
                else:
                    # Select by visible text using XPath
                   option = wait.until(EC.element_to_be_clickable((By.XPATH, "//option[text()='Bill of Entry']")))
                   option.click()
                   time.sleep(2)
                   print("select Bill of Entry")

                for idx, row in enumerate(self.excel_data):
                    print(row)
                
                    # --- Fix: handle pandas Timestamp directly ---
                    shipping_date = row.get("ADV Shipping Bill Date")
                    if pd.isna(shipping_date):
                        print(f"Skipping Row {idx + 1}: No Shipping Bill Date found")
                        continue
                
                    # Ensure it's a datetime object
                    if isinstance(shipping_date, pd.Timestamp):
                        shipping_date = shipping_date.to_pydatetime()
                
                    shipping_bill_date = shipping_date.strftime("%d/%m/%Y")
                    print("Formatted Shipping Bill Date:", shipping_bill_date)
                
                    # --- Wait for date field and fill it ---
                    wait.until(EC.presence_of_element_located((By.ID, "fromDateOfSelectedBil")))
                    print("shown shipping bill date")
                
                    self.driver.execute_script(f"""
                        var fromDate = document.getElementById('fromDateOfSelectedBil');
                        fromDate.value = '{shipping_bill_date}';
                        fromDate.dispatchEvent(new Event('change'));
                    """)
                
                    time.sleep(1)
                    
                
                    # --- Fill Authorisation Number ---
                    auth_no = str(row.get("ADV Authorisation Number", "")).strip()
                    if i==0:
                        self.driver.find_element(By.ID, "authorisationNo").clear()
                        self.driver.find_element(By.ID, "authorisationNo").send_keys(auth_no)
                    else:
                        self.driver.find_element(By.ID, "boeLicenseNumber").clear()
                        self.driver.find_element(By.ID, "boeLicenseNumber").send_keys(auth_no)
                    
                    print("add authorisation number")
                    time.sleep(1)
                
                    # --- Click Search ---
                    search = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="repSearchBtn"]')))
                    
                    search.click()
                    print("✅ Clicked search")
                    if i==0:

                        # --- Wait for table to appear ---
                        wait.until(EC.presence_of_element_located((By.ID, "billRepositoryTable")))
                        print("table shown")
                        
                        all_rows = []
                        
                        while True:
                            # Wait for the table body to load
                            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#billRepositoryTable tbody tr")))
                        
                            rows = self.driver.find_elements(By.CSS_SELECTOR, "#billRepositoryTable tbody tr")
                            print(f"Found {len(rows)} rows on this page")
                        
                            for r in rows:
                                cols = r.find_elements(By.TAG_NAME, "td")
                                if len(cols) > 1:  # skip "No data available" row
                                    data = [c.text.strip() for c in cols]
                                    all_rows.append(data)
                        
                            # Try to find "Next" button and check if it is enabled
                            try:
                                self.driver.execute_script("window.scrollBy(0, -100);")
                                self.driver.execute_script("window.scrollBy(0, 100);")
                                wait.until(EC.presence_of_element_located((By.ID, "billRepositoryTable_next")))
                                next_btn = self.driver.find_element(By.ID, "billRepositoryTable_next")
                                next_class = next_btn.get_attribute("class")
                                if "disabled" in next_class:
                                    print("✅ No more pages. Extraction complete.")
                                    break
                                else:
                                    self.driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
                                    next_btn.click()
                                    time.sleep(2)  # Wait for next page data to load
                            except Exception as e:
                                print("⚠️ Pagination ended or not found:", e)
                                break
                        
                        # --- Extract headers ---
                        headers = [h.text.strip() for h in self.driver.find_elements(By.CSS_SELECTOR, "#billRepositoryTable thead th")]
                        
                        # --- Save to Excel ---
                        # --- Save to Local Downloads Folder ---
                        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"ADV_shipping_bill_Data_{timestamp}.xlsx"
                        file_path = os.path.join(downloads_path, filename)                    
    
                        df = pd.DataFrame(all_rows, columns=headers)
                        df.to_excel(file_path, index=False)                    
    
                        print(f"💾 Data saved to: {file_path}")
                    else:
                                                    # --- Wait for Bill of Entry table to load ---
                        wait.until(EC.presence_of_element_located((By.ID, "billOfEntryTable")))
                        
                        all_rows = []
                        
                        while True:
                            # Wait for "Processing..." to disappear
                            try:
                                wait.until_not(EC.visibility_of_element_located((By.ID, "billOfEntryTable_processing")))
                            except:
                                pass
                        
                            # Wait for table rows
                            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#billOfEntryTable tbody tr")))
                        
                            rows = self.driver.find_elements(By.CSS_SELECTOR, "#billOfEntryTable tbody tr")
                            print(f"Found {len(rows)} rows on this page")
                        
                            for r in rows:
                                cols = r.find_elements(By.TAG_NAME, "td")
                                if len(cols) > 1:  # skip "No data available"
                                    data = [c.text.strip() for c in cols]
                                    all_rows.append(data)
                        
                            # --- Handle Pagination ---
                            try:
                                self.driver.execute_script("window.scrollBy(0, -100);")
                                self.driver.execute_script("window.scrollBy(0, 100);")
                                wait.until(EC.presence_of_element_located((By.ID, "billOfEntryTable_next")))
                                next_btn = self.driver.find_element(By.ID, "billOfEntryTable_next")
                                next_class = next_btn.get_attribute("class")
                                if "disabled" in next_class:
                                    print("✅ No more pages. Extraction complete.")
                                    break
                                else:
                                    self.driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
                                    next_btn.click()
                                    time.sleep(2)  # allow table to refresh
                            except Exception as e:
                                print("⚠️ Pagination ended or not found:", e)
                                break
                        
                        # --- Extract Table Headers ---
                        headers = [h.text.strip() for h in self.driver.find_elements(By.CSS_SELECTOR, "#billOfEntryTable thead th")]
                        
                        # --- Save to Local Downloads Folder ---
                        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"ADV_BillOfEntry_Data_{timestamp}.xlsx"
                        file_path = os.path.join(downloads_path, filename)
                        
                        df = pd.DataFrame(all_rows, columns=headers)
                        df.to_excel(file_path, index=False)
                        
                        print(f"💾 Data saved to: {file_path}")
            time.sleep(1)
            self.driver.refresh()
            time.sleep(1)
            self.driver.execute_script("document.body.style.zoom='80%'")
            time.sleep(0.5)
            self.driver.execute_script("document.body.style.zoom='60%'")
            time.sleep(0.5)
            adv_auth = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Advanced')]"))
            )
            adv_auth.click()
            closure_adv_auth = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//h5[normalize-space()='Closure of Advance Authorisation']"))
            )
            closure_adv_auth.click()
            new_app = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[@id='btnNewApp']"))
            )
            new_app.click()
            target_auth_no = str(row.get("ADV Authorisation Number", "")).strip()
            xpath = f"//table[@id='authorizationTable']//tr[td[3][normalize-space()='{target_auth_no}']]//input[@type='radio']"
            
            wait = WebDriverWait(self.driver, 10)
            radio = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            self.driver.execute_script("arguments[0].scrollIntoView(true);", radio)
            radio.click()
            prd_val = wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="withOutAmendment"]'))
            )
            prd_val.click()
            nxt = wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="btnNext"]'))
            )
            nxt.click()
            time.sleep(1)
            
            h_cert = wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="caCeSelect"]/div/div/div/div[2]/label'))
            )
            # now apply zoom
            
            h_cert.click()
            time.sleep(2)
            self.driver.execute_script("window.scrollBy(0, 600);")
            self.driver.execute_script("document.body.click();")
            # time.sleep(1)
            self.driver.execute_script("document.body.style.zoom='80%'")
            time.sleep(0.5)
            self.driver.execute_script("document.body.style.zoom='60%'")
            time.sleep(0.5)
            # self.driver.execute_script("document.body.style.zoom='100%'")
            # time.sleep(0.5)
            # advance licence at CIP ad endorsed 
            time.sleep(0.5)
            wait = WebDriverWait(self.driver, 15)
            
            # --- Locate the main table ---
            table = wait.until(
                EC.presence_of_element_located((By.ID, "exportEndorsedTbl"))
            )
            
            all_rows = []
            time.sleep(0.5)
            self.driver.execute_script("document.body.click();")
            
            # --- Extract all pages ---
            while True:
                # Extract headers only once
                headers = [th.text.strip() for th in table.find_elements(By.TAG_NAME, "th")]
            
                # Extract body rows
                body_rows = table.find_elements(By.TAG_NAME, "tr")[1:]
                for r in body_rows:
                    cells = [td.text.strip() for td in r.find_elements(By.TAG_NAME, "td")]
                    if cells:
                        all_rows.append(cells)
            
                # --- Pagination: click next page ---
                try:
                    next_btn = wait.until(
                        EC.presence_of_element_located((By.ID, "exportEndorsedTbl_next"))
                    )
            
                    # If disabled → last page
                    if "disabled" in next_btn.get_attribute("class"):
                        break
            
                    # Scroll small up/down to avoid overlay issues
                    self.driver.execute_script("window.scrollBy(0, -150);")
                    self.driver.execute_script("window.scrollBy(0, 150);")
            
                    next_btn.click()
                    time.sleep(1.2)
            
                    # Re-fetch table after pagination change
                    table = self.driver.find_element(By.ID, "exportEndorsedTbl")
            
                except Exception as e:
                    break
            
            # --- Save in Downloads folder ---
            downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"EXPORT_ENDORSED_DATA_{timestamp}.xlsx"
            file_path = os.path.join(downloads_path, filename)
            
            df = pd.DataFrame(all_rows, columns=headers)
            df.to_excel(file_path, index=False)
            
            print(f"💾 Export Endorsed Table Saved: {file_path}")
            # self.driver.execute_script("document.body.style.zoom='50%'")
            time.sleep(0.5)
            wait = WebDriverWait(self.driver, 15)
            
            # self.driver.execute_script("window.scrollBy(0, 1000);")
            # ---- Locate main table ----
            table = wait.until(
                EC.presence_of_element_located((By.ID, "exportShippingGstBillTbl"))
            )
            # t_click = wait.until(
            #     EC.element_to_be_clickable((By.XPATH, '//*[@id="exportShippingGstBillTbl"]'))
            # )
            # t_click.click()
            # self.driver.execute_script("document.body.style.zoom='50%'")
            
            
            all_rows = []  
            # self.driver.execute_script("document.body.style.zoom='50%'")
            
            time.sleep(0.5) 
            self.driver.execute_script("document.body.click();")
            # ---- Extract all pages ----
            while True:
            
                # extract headers (only first page)
                headers = [th.text.strip() for th in table.find_elements(By.TAG_NAME, "th")]                
                # extract rows
                body_rows = table.find_elements(By.TAG_NAME, "tr")[1:]
                for r in body_rows:
                    cells = [td.text.strip() for td in r.find_elements(By.TAG_NAME, "td")]
                    if cells:
                        all_rows.append(cells)
            
                # find next page button
                try:
                    self.driver.execute_script("window.scrollBy(0, -100);")
                    self.driver.execute_script("window.scrollBy(0, 100);")
                    wait.until(EC.presence_of_element_located((By.ID, "exportShippingGstBillTbl_next")))
                    next_btn = self.driver.find_element(By.ID, "exportShippingGstBillTbl_next")
                    # if disabled → last page
                    if "disabled" in next_btn.get_attribute("class"):
                        break
            
                    next_btn.click()
                    time.sleep(1.2)
            
                    table = self.driver.find_element(By.ID, "exportShippingGstBillTbl")
            
                except:
                    break
            
            # --- Save to Local Downloads Folder ---
            downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"ADV_AS_per_shipping_bill_{timestamp}.xlsx"
            file_path = os.path.join(downloads_path, filename)
            
            df = pd.DataFrame(all_rows, columns=headers)
            df.to_excel(file_path, index=False)
            
            print(f"💾 Data saved to: {file_path}")
            # reset zoom
            # self.driver.execute_script("document.body.style.zoom='100%'")
            # AS per importer exporter data
            wait = WebDriverWait(self.driver, 20)
            table_id = "exporterImporterTbl"
            
            # ------------ HEADERS ------------
            header_xpath = f"//table[@id='{table_id}']//thead//th"
            wait.until(EC.presence_of_all_elements_located((By.XPATH, header_xpath)))
            
            importer_headers = [
                h.get_attribute("innerText").strip()
                for h in self.driver.find_elements(By.XPATH, header_xpath)
            ]
            
            # ------------ DATA EXTRACTION ------------
            importer_all_rows = []
            
            while True:
            
                row_xpath = f"//table[@id='{table_id}']//tbody//tr[not(contains(@class,'child'))]"
                wait.until(EC.presence_of_all_elements_located((By.XPATH, row_xpath)))
            
                rows = self.driver.find_elements(By.XPATH, row_xpath)
            
                for r in rows:
                    cols = r.find_elements(By.XPATH, "./td")
                    importer_row = [c.get_attribute("innerText").strip() for c in cols]
                    importer_all_rows.append(importer_row)
                self.driver.execute_script("window.scrollBy(0, -100);")
                self.driver.execute_script("window.scrollBy(0, 100);")
                wait.until(EC.presence_of_element_located((By.XPATH, f"//li[@id='{table_id}_next']")))
                next_btn = self.driver.find_element(By.XPATH, f"//li[@id='{table_id}_next']")
            
                if "disabled" in next_btn.get_attribute("class"):
                    break
            
                self.driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
                next_btn.click()
                time.sleep(1)
            
            # ------------ SAVE FILE ------------
            downloads = os.path.join(os.path.expanduser("~"), "Downloads")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = os.path.join(downloads, f"export_as_per_importer{timestamp}.xlsx")
            
            df = pd.DataFrame(importer_all_rows, columns=importer_headers)
            df.to_excel(file_path, index=False)
            
            print("Saved:", file_path)
            self.driver.execute_script("window.scrollBy(0, 500);")
            
            print("\n=== STARTING EXPORT REALIZATION EXTRACTION ===")
            
            wait.until(EC.presence_of_element_located((By.ID, "exportRealizationTbl")))
            wait.until(EC.visibility_of_element_located((By.ID, "exportRealizationTbl")))
            
            all_rows = []
            page = 1
            
            # ------------------
            # GET HEADERS ONCE
            # ------------------
            rel_table = self.driver.find_element(By.ID, "exportRealizationTbl")
            header_elems = rel_table.find_elements(By.CSS_SELECTOR, "thead th")
            
            rel_headers = [th.text.strip() for th in header_elems]
            print("Headers:", rel_headers)
            
            
            # ------------------
            # PAGINATION LOOP
            # ------------------
            while True:
                print(f"\n--- Extracting Page {page} ---")
            
                # re-fetch table for safe reference
                table = self.driver.find_element(By.ID, "exportRealizationTbl")
            
                row_elems = table.find_elements(By.CSS_SELECTOR, "tbody tr")
            
                if not row_elems:
                    time.sleep(1)
                    table = self.driver.find_element(By.ID, "exportRealizationTbl")
                    row_elems = table.find_elements(By.CSS_SELECTOR, "tbody tr")
            
                # extract rows
                for tr in row_elems:
                    cells = tr.find_elements(By.TAG_NAME, "td")
                    all_rows.append([td.text.strip() for td in cells])
            
                print(f"Collected {len(row_elems)} rows from page {page}")
            
                # ------------------
                # CLICK NEXT
                # ------------------
                try:
                    next_btn = self.driver.find_element(By.ID, "exportRealizationTbl_next")
            
                    # stop if disabled
                    if "disabled" in next_btn.get_attribute("class"):
                        print("\nReached last page. Stopping.")
                        break
            
                    # click next page
                    self.driver.execute_script("arguments[0].scrollIntoView();", next_btn)
                    time.sleep(0.5)
                    self.driver.execute_script("arguments[0].click();", next_btn)
            
                    time.sleep(2)
                    page += 1
            
                except Exception as e:
                    print("Next button error:", e)
                    break
            
            
            # ------------------------
            # SAVE FINAL EXCEL FILE
            # ------------------------
            downloads = os.path.join(os.path.expanduser("~"), "Downloads")
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            file_path = os.path.join(downloads, f"export_realization_{timestamp}.xlsx")
            
            df = pd.DataFrame(all_rows, columns=rel_headers)
            df.to_excel(file_path, index=False)
            
            print("\n✓ DONE! Extracted", len(all_rows), "rows")
            print("File saved:", file_path)
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            save_prc = wait.until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="btnNext"]'))
            )
            save_prc.click()
            time.sleep(2)
            all_rows = []
            # -----------------------------------------------------
            # Extract table headers
            # -----------------------------------------------------
            header_elems = wait.until(
                EC.presence_of_all_elements_located(
                    (By.CSS_SELECTOR, "#exportItemsTable thead th")
                )
            )
            
            headers = [h.text.strip() for h in header_elems]
            
            # -----------------------------------------------------
            # Loop through all pages
            # -----------------------------------------------------
            while True:
                # wait for rows
                time.sleep(1)
                rows = self.driver.find_elements(By.CSS_SELECTOR, "#exportItemsTable tbody tr")
            
                for r in rows:
                    cols = [c.text.strip() for c in r.find_elements(By.TAG_NAME, "td")]
                    all_rows.append(cols)
            
                # Check for next button
                try:
                    self.driver.execute_script("window.scrollBy(0, -100);")
                    self.driver.execute_script("window.scrollBy(0, 100);")
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,
                        "#exportItemsTable_paginate .paginate_button.next")))
                    next_btn = self.driver.find_element(
                        By.CSS_SELECTOR,
                        "#exportItemsTable_paginate .paginate_button.next"
                    )
                    if "disabled" in next_btn.get_attribute("class"):
                        break  # no more pages
                    next_btn.click()
                except:
                    break
            
            # -----------------------------------------------------
            # Save to Excel in Downloads folder
            # -----------------------------------------------------
            downloads = os.path.join(os.path.expanduser("~"), "Downloads")
            file_path = os.path.join(downloads, f"input_As_endorsed_auth{timestamp}.xlsx")
            
            df = pd.DataFrame(all_rows, columns=headers)
            df.to_excel(file_path, index=False)
            
            print("Saved:", file_path)
            time.sleep(2)
            # input bill of entries
            self.driver.execute_script("window.scrollBy(0, 500);")
            time.sleep(1)
            b_all_rows = []
            # ----------------------------------------------------
            # Extract headers (these are always visible)
            # ----------------------------------------------------
            b_headers_elements = wait.until(
                EC.presence_of_all_elements_located(
                    (By.CSS_SELECTOR, "#inputItemsGSTTable thead th")
                )
            )
            b_headers = [h.get_attribute("innerText").strip() for h in b_headers_elements]
            
            
            # ----------------------------------------------------
            # Extract all pages
            # ----------------------------------------------------
            while True:
                time.sleep(1)
            
                # all TRs
                b_rows = self.driver.find_elements(By.CSS_SELECTOR, "#inputItemsGSTTable tbody tr")
            
                for r in b_rows:
                    # ✅ extract all TDs including hidden using JS
                    cols = self.driver.execute_script(
                        "return Array.from(arguments[0].querySelectorAll('td')).map(td => td.innerText.trim());",
                        r
                    )
                    b_all_rows.append(cols)
            
                # pagination: strict next button
                try:
                    self.driver.execute_script("window.scrollBy(0, -100);")
                    self.driver.execute_script("window.scrollBy(0, 100);")
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,
                        "#inputItemsGSTTable_paginate .paginate_button.next a")))
                    next_btn = self.driver.find_element(
                        By.CSS_SELECTOR,
                        "#inputItemsGSTTable_paginate .paginate_button.next a"
                    )
            
                    parent_li = next_btn.find_element(By.XPATH, "..")
                    if "disabled" in parent_li.get_attribute("class"):
                        break
            
                    next_btn.click()
            
                except Exception:
                    break
            
            
            
            # ----------------------------------------------------
            # Save Excel to Downloads
            # ----------------------------------------------------
            downloads = os.path.join(os.path.expanduser("~"), "Downloads")
            b_file_path = os.path.join(downloads, f"bill_of_entry{timestamp}.xlsx")
            
            df = pd.DataFrame(b_all_rows, columns=b_headers)
            df.to_excel(b_file_path, index=False)
            
            print("Saved:", b_file_path)
            # input as per importer/exporter
            self.driver.execute_script("window.scrollBy(0, 500);")
            time.sleep(2)
            c_all_rows = []
            # ----------------------------------------------------
            # Extract table headers
            # ----------------------------------------------------
            c_headers_el = wait.until(
                EC.presence_of_all_elements_located(
                    (By.CSS_SELECTOR, "#ImporterExporterTable thead th")
                )
            )
            c_headers = [h.text.strip() for h in c_headers_el]
            
            # ----------------------------------------------------
            # Extract all pages
            # ----------------------------------------------------
            while True:
                time.sleep(1)
            
                c_rows = self.driver.find_elements(By.CSS_SELECTOR, "#ImporterExporterTable tbody tr")
            
                for r in c_rows:
                    c_cols = [c.text.strip() for c in r.find_elements(By.TAG_NAME, "td")]
                    c_all_rows.append(c_cols)
            
                try:
                    self.driver.execute_script("window.scrollBy(0, -100);")
                    self.driver.execute_script("window.scrollBy(0, 100);")
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,
                        "#ImporterExporterTable_paginate .paginate_button.next")))
                    next_btn = self.driver.find_element(
                        By.CSS_SELECTOR,
                        "#ImporterExporterTable_paginate .paginate_button.next"
                    )
            
                    if "disabled" in next_btn.get_attribute("class"):
                        break  # no more pages
            
                    next_btn.click()
                except:
                    break
            
            # ----------------------------------------------------
            # Save Excel to Downloads folder
            # ----------------------------------------------------
            downloads = os.path.join(os.path.expanduser("~"), "Downloads")
            c_file_path = os.path.join(downloads, f"input_importer_exporter{timestamp}.xlsx")
            
            df = pd.DataFrame(c_all_rows, columns=c_headers)
            df.to_excel(c_file_path, index=False)
            
            print("Saved:", c_file_path)
        
            return {"success": True, "message": "ADV process completed"}
            
        except Exception as e:
            return {"success": False, "message": f"Error in ADV process: {e}"}

    def start_browser(self):
        options = webdriver.ChromeOptions()
        prefs = {
            "download.default_directory": os.path.expanduser("~/Downloads"),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True
        }
        options.add_experimental_option("prefs", prefs)
        options.add_experimental_option("detach", True)
        options.add_argument("--force-device-scale-factor=0.98")
        options.add_argument("--high-dpi-support=0.98")
        options.add_argument("--disable-gpu")
        self.driver = webdriver.Chrome(options=options)
        self.driver.maximize_window()

    def login(self):
        try:
            self.driver.get("https://www.dgft.gov.in/CP/")
            
            # Click skip button if present
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "skip"))
                ).click()
                print("Clicked skip button")
            except TimeoutException:
                print("Skip button not found, continuing...")

            time.sleep(2)
            
            # Switch to new window if available
            if len(self.driver.window_handles) > 1:
                self.driver.switch_to.window(self.driver.window_handles[-1])
            
            # Wait for login form
            WebDriverWait(self.driver, 30).until(
                EC.presence_of_element_located((By.ID, "username"))
            )

            # Fill credentials
            self.driver.find_element(By.ID, "username").clear()
            self.driver.find_element(By.ID, "username").send_keys(self.username)
            
            self.driver.find_element(By.ID, "password").clear()
            self.driver.find_element(By.ID, "password").send_keys(self.password)

            print("⚠️ Please solve the CAPTCHA manually and click 'Login'...")
            print("⚠️ Waiting for manual CAPTCHA solution...")
            
            # Wait for user to manually solve CAPTCHA and login
            WebDriverWait(self.driver, 300).until(
                EC.url_changes(self.driver.current_url)
            )
            
            print("✅ Login successful!")
            return {"success": True, "message": "Login successful"}
            
        except Exception as e:
            return {"success": False, "message": f"Login failed: {e}"}

    # def close_browser(self):
    #     if self.driver:
    #         self.driver.quit()