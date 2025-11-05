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

class CertificateBot:
    def __init__(self, username, password, excel_data, download_folder):
        self.username = username
        self.password = password
        self.excel_data = excel_data
        # self.required_files_folder = required_files_folder
        self.download_folder = download_folder
        self.driver = None
        self.current_index = 0
        
        
        # self.required_docs = os.path.abspath(required_files_folder)

    def process_all_certificates(self):
        try:
            if not self.excel_data:
                return {"success": False, "message": "No data found in Excel file"}
            
            # Process all certificates
            while self.current_index < len(self.excel_data):
                
                full_login = self.current_index == 0
                result = self._process_certificate(full_login=full_login)
                if not result.get('success'):
                    return result
            
                        
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
              
            
            fill_result = self.fill_certificate()
            if not fill_result.get('success'):
                return fill_result
            
            self.current_index += 1
            return {"success": True, "message": f"Processed  {row}"}
            
        except Exception as e:
            return {"success": False, "message": f"Error processing certificate: {e}"}

    def start_browser(self):
        options = webdriver.ChromeOptions()
        prefs = {
            "download.default_directory": self.download_folder,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True
        }
        options.add_experimental_option("prefs", prefs)
        options.add_experimental_option("detach", True)
        # Do NOT add headless option, keep browser visible
        # options.add_argument("--headless")  # (do not use)

        # profile_dir = os.path.join(self.download_folder, "chrome_profile")
        # if not os.path.exists(profile_dir):
        #     os.makedirs(profile_dir)
        # options.add_argument(f"--user-data-dir={profile_dir}")
        
        
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
            
    def fill_certificate(self):
        try:
            print(1)
            time.sleep(1)
            row = self.excel_data[self.current_index]
            if self.current_index==0:

                wait = WebDriverWait(self.driver, 50)
            
                # Navigate to certificate section
                print("Navigating to certificate section...")
                time.sleep(1)
                
                # Try to find and click on Advanced Authorization
                my_dashboard = wait.until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="200928"]/a'))
                )
                my_dashboard.click()
                print("✅ Clicked on my dashboard")
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
                        shipping_date = row.get("Shipping Bill Date")
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
                        auth_no = str(row.get("Authorisation Number", "")).strip()
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
                            filename = f"EPCG_shipping_bill_Data_{timestamp}.xlsx"
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
                            filename = f"EPCG_BillOfEntry_Data_{timestamp}.xlsx"
                            file_path = os.path.join(downloads_path, filename)
                            
                            df = pd.DataFrame(all_rows, columns=headers)
                            df.to_excel(file_path, index=False)
                            
                            print(f"💾 Data saved to: {file_path}")

                time.sleep(1)
                self.driver.refresh()  
                time.sleep(1)
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

                wait = WebDriverWait(self.driver, 15)
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
                target_auth_no = row.get("Authorisation Number")
                

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


                
                

            else:
                time.sleep(2)
    
            
        

        except Exception as e:
            return {"success": False, "message": f"Error in fill_certificate: {e}"}

    # def close_browser(self):
    #     if self.driver:
    #         self.driver.quit()
