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
        options.add_argument("--force-device-scale-factor=0.98")   # 50% zoom
        options.add_argument("--high-dpi-support=0.98")
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
                
                # # Try to find and click on Advanced Authorization
                # my_dashboard = wait.until(
                #     EC.element_to_be_clickable((By.XPATH, '//*[@id="200928"]/a'))
                # )
                # my_dashboard.click()
                # print("✅ Clicked on my dashboard")
                # repo = wait.until(
                #     EC.element_to_be_clickable((By.XPATH, '//*[@id="98000023"]/a'))
                # )
                # repo.click()
                # print("✅ Clicked on repositories")
    
                # Bill_repo = wait.until(
                #     EC.element_to_be_clickable((By.XPATH, '//*[@id="content"]/div[3]/div/div[2]/div[1]/div/a'))
                # )
                # Bill_repo.click()
                # print("✅ Clicked on Bill repositories")
                
                # for i in range(2):
                    
                #     wait = WebDriverWait(self.driver, 15)
                #     dropdown = wait.until(EC.element_to_be_clickable((By.ID, "txt_selectBill")))
                #     dropdown.click()
                #     time.sleep(1)

                #     if i==0:
                    
                #        # Select by visible text using XPath
                #        option = wait.until(EC.element_to_be_clickable((By.XPATH, "//option[text()='Shipping Bill']")))
                #        option.click()
                #        time.sleep(2)
                #        print("select shipping bill")

                #     else:
                #         # Select by visible text using XPath
                #        option = wait.until(EC.element_to_be_clickable((By.XPATH, "//option[text()='Bill of Entry']")))
                #        option.click()
                #        time.sleep(2)
                #        print("select Bill of Entry")
    
                #     for idx, row in enumerate(self.excel_data):
                #         print(row)
                    
                #         # --- Fix: handle pandas Timestamp directly ---
                #         shipping_date = row.get("Shipping Bill Date")
                #         if pd.isna(shipping_date):
                #             print(f"Skipping Row {idx + 1}: No Shipping Bill Date found")
                #             continue
                    
                #         # Ensure it's a datetime object
                #         if isinstance(shipping_date, pd.Timestamp):
                #             shipping_date = shipping_date.to_pydatetime()
                    
                #         shipping_bill_date = shipping_date.strftime("%d/%m/%Y")
                #         print("Formatted Shipping Bill Date:", shipping_bill_date)
                    
                #         # --- Wait for date field and fill it ---
                #         wait.until(EC.presence_of_element_located((By.ID, "fromDateOfSelectedBil")))
                #         print("shown shipping bill date")
                    
                #         self.driver.execute_script(f"""
                #             var fromDate = document.getElementById('fromDateOfSelectedBil');
                #             fromDate.value = '{shipping_bill_date}';
                #             fromDate.dispatchEvent(new Event('change'));
                #         """)
                    
                #         time.sleep(1)

                        

                    
                #         # --- Fill Authorisation Number ---
                #         auth_no = str(row.get("Authorisation Number", "")).strip()
                #         if i==0:
                #             self.driver.find_element(By.ID, "authorisationNo").clear()
                #             self.driver.find_element(By.ID, "authorisationNo").send_keys(auth_no)
                #         else:
                #             self.driver.find_element(By.ID, "boeLicenseNumber").clear()
                #             self.driver.find_element(By.ID, "boeLicenseNumber").send_keys(auth_no)
                        

                #         print("add authorisation number")
                    
                #         # --- Click Search ---
                #         search = wait.until(EC.element_to_be_clickable((By.ID, 'repSearchBtn')))
                #         search.click()
                #         print("✅ Clicked search")

                #         if i==0:
    
                #             # --- Wait for table to appear ---
                #             wait.until(EC.presence_of_element_located((By.ID, "billRepositoryTable")))
                #             print("table shown")
                            
                #             all_rows = []
                            
                #             while True:
                #                 # Wait for the table body to load
                #                 wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#billRepositoryTable tbody tr")))
                            
                #                 rows = self.driver.find_elements(By.CSS_SELECTOR, "#billRepositoryTable tbody tr")
                #                 print(f"Found {len(rows)} rows on this page")
                            
                #                 for r in rows:
                #                     cols = r.find_elements(By.TAG_NAME, "td")
                #                     if len(cols) > 1:  # skip "No data available" row
                #                         data = [c.text.strip() for c in cols]
                #                         all_rows.append(data)
                            
                #                 # Try to find "Next" button and check if it is enabled
                #                 try:
                #                     next_btn = self.driver.find_element(By.ID, "billRepositoryTable_next")
                #                     next_class = next_btn.get_attribute("class")
                #                     if "disabled" in next_class:
                #                         print("✅ No more pages. Extraction complete.")
                #                         break
                #                     else:
                #                         self.driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
                #                         next_btn.click()
                #                         time.sleep(2)  # Wait for next page data to load
                #                 except Exception as e:
                #                     print("⚠️ Pagination ended or not found:", e)
                #                     break
                            
                #             # --- Extract headers ---
                #             headers = [h.text.strip() for h in self.driver.find_elements(By.CSS_SELECTOR, "#billRepositoryTable thead th")]
                            
                #             # --- Save to Excel ---
                #             # --- Save to Local Downloads Folder ---
                #             downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
                #             timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                #             filename = f"EPCG_shipping_bill_Data_{timestamp}.xlsx"
                #             file_path = os.path.join(downloads_path, filename)                    
        
                #             df = pd.DataFrame(all_rows, columns=headers)
                #             df.to_excel(file_path, index=False)                    
        
                #             print(f"💾 Data saved to: {file_path}")

                #         else:
                #                                         # --- Wait for Bill of Entry table to load ---
                #             wait.until(EC.presence_of_element_located((By.ID, "billOfEntryTable")))
                            
                #             all_rows = []
                            
                #             while True:
                #                 # Wait for "Processing..." to disappear
                #                 try:
                #                     wait.until_not(EC.visibility_of_element_located((By.ID, "billOfEntryTable_processing")))
                #                 except:
                #                     pass
                            
                #                 # Wait for table rows
                #                 wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#billOfEntryTable tbody tr")))
                            
                #                 rows = self.driver.find_elements(By.CSS_SELECTOR, "#billOfEntryTable tbody tr")
                #                 print(f"Found {len(rows)} rows on this page")
                            
                #                 for r in rows:
                #                     cols = r.find_elements(By.TAG_NAME, "td")
                #                     if len(cols) > 1:  # skip "No data available"
                #                         data = [c.text.strip() for c in cols]
                #                         all_rows.append(data)
                            
                #                 # --- Handle Pagination ---
                #                 try:
                #                     next_btn = self.driver.find_element(By.ID, "billOfEntryTable_next")
                #                     next_class = next_btn.get_attribute("class")
                #                     if "disabled" in next_class:
                #                         print("✅ No more pages. Extraction complete.")
                #                         break
                #                     else:
                #                         self.driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
                #                         next_btn.click()
                #                         time.sleep(2)  # allow table to refresh
                #                 except Exception as e:
                #                     print("⚠️ Pagination ended or not found:", e)
                #                     break
                            
                #             # --- Extract Table Headers ---
                #             headers = [h.text.strip() for h in self.driver.find_elements(By.CSS_SELECTOR, "#billOfEntryTable thead th")]
                            
                #             # --- Save to Local Downloads Folder ---
                #             downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
                #             timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                #             filename = f"EPCG_BillOfEntry_Data_{timestamp}.xlsx"
                #             file_path = os.path.join(downloads_path, filename)
                            
                #             df = pd.DataFrame(all_rows, columns=headers)
                #             df.to_excel(file_path, index=False)
                            
                #             print(f"💾 Data saved to: {file_path}")

                # time.sleep(1)
                # self.driver.refresh()  
                # time.sleep(1)
                # repo = wait.until(
                #     EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='EPCG']"))
                # )
                # repo.click()
                # print("✅ Clicked on EPCG")

                # repo = wait.until(
                #     EC.element_to_be_clickable((By.XPATH, "//h5[normalize-space()='Closure of EPCG/Issuance of Post Export Scrip']"))
                # )
                # repo.click()
                # print("✅ Clicked on closure of EPCG")
                # time.sleep(1)

                # repo = wait.until(
                #     EC.element_to_be_clickable((By.XPATH, "//button[@id='btnNewApp']"))
                # )
                # repo.click()
                # print("Start fresh application")

                # wait = WebDriverWait(self.driver, 15)
                # dropdown = wait.until(EC.element_to_be_clickable((By.ID, "applicationFor")))
                # dropdown.click()
                # option = wait.until(EC.element_to_be_clickable((By.XPATH, "//option[text()='REDEMPTION']")))
                # option.click()
                # time.sleep(1)

                # auth_closure = wait.until(
                #     EC.element_to_be_clickable((By.XPATH, '//*[@id="custom-accordion"]/div[2]/div[1]/a'))
                # )
                # auth_closure.click()
                # print("Auth closure")
                

                
                
                # # EPCG Authorization Number to match
                # target_auth_no = row.get("EPCG Authorisation Number")
                

                # # Wait for the table to be visible
                # wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.dataTables_scrollBody")))
                
                # found = False
                # page = 1
                
                # while True:
                #     try:
                #         # Find the cell containing the EPCG Authorization Number
                #         cell_xpath = f"//div[@class='dataTables_scrollBody']//table//td[normalize-space()='{target_auth_no}']"
                #         target_cell = self.driver.find_element(By.XPATH, cell_xpath)
                
                #         # Scroll the cell into view
                #         self.driver.execute_script("arguments[0].scrollIntoView(true);", target_cell)
                #         time.sleep(1)
                
                #         # Get the parent row and click its radio button
                #         parent_row = target_cell.find_element(By.XPATH, "./ancestor::tr")
                #         radio_button = parent_row.find_element(By.CSS_SELECTOR, "input[type='radio']")
                #         self.driver.execute_script("arguments[0].click();", radio_button)
                
                #         print(f"✅ Selected EPCG Authorization Number: {target_auth_no}")
                #         found = True
                #         break
                
                #     except Exception:
                #         # Check for next page if not found
                #         try:
                #             next_button = self.driver.find_element(By.ID, "epcgauthTbl_next")
                #             if "disabled" in next_button.get_attribute("class"):
                #                 break  # No more pages left
                #             self.driver.execute_script("arguments[0].click();", next_button)
                #             page += 1
                #             print(f"🔁 Searching on page {page}...")
                #             time.sleep(2)
                #         except:
                #             break
                
                # if not found:
                #     print(f"❌ Target EPCG Authorization Number {target_auth_no} not found in any page.")
                
                # prd_val = wait.until(
                #     EC.element_to_be_clickable((By.XPATH, '//*[@id="eoPending"]/div/div/div[2]/div/div/div[1]/label'))
                # )
                # prd_val.click()
                # nxt = wait.until(
                #     EC.element_to_be_clickable((By.XPATH, '//*[@id="btnNext"]'))
                # )
                # nxt.click()
                # time.sleep(1)
                # self.driver.execute_script("window.scrollBy(0, 300);")
                # ii = wait.until(
                #     EC.element_to_be_clickable((By.XPATH, '//*[@id="closureAuthorisationDetails"]/div[2]/div[1]/a'))
                # )
                # ii.click()
                # print("item of import open")
                # time.sleep(1)
                # self.driver.execute_script("window.scrollBy(0, -300);")
                # ii = wait.until(
                #     EC.element_to_be_clickable((By.XPATH, '//*[@id="closureAuthorisationDetails"]/div[2]/div[1]/a'))
                # )
                # ii.click()
                # print("item of import close")
                # time.sleep(1)
                

                # time.sleep(1)
                # self.driver.execute_script("window.scrollBy(0, 300);")
                # ie = wait.until(
                #     EC.element_to_be_clickable((By.XPATH, '//*[@id="closureAuthorisationDetails"]/div[3]/div[1]/a'))
                # )
                # ie.click()
                # print("items of export")

                # # --- Wait for the DataTable wrapper to load ---
                # wait.until(EC.presence_of_element_located((By.ID, "authExportItemTbl_wrapper")))
                
                # # --- Extract headers (check both visible and cloned header tables) ---
                # header_elements = self.driver.find_elements(By.XPATH,
                #     "(//div[@id='authExportItemTbl_wrapper']//table//thead//th | //table[contains(@id,'authExportItemTbl')]//thead//th)"
                # )
                # headers = [h.text.strip() for h in header_elements if h.text.strip()]
                # if not headers:
                #     print("⚠️ No visible headers found. Retrying...")
                #     time.sleep(2)
                #     header_elements = self.driver.find_elements(By.XPATH,
                #         "(//div[@id='authExportItemTbl_wrapper']//table//thead//th | //table[contains(@id,'authExportItemTbl')]//thead//th)"
                #     )
                #     headers = [h.text.strip() for h in header_elements if h.text.strip()]
                
                # print("✅ Headers found:", headers)
                
                # # --- Prepare data storage ---
                # all_data = []
                
                # # --- Pagination loop ---
                # while True:
                #     # Wait for rows to appear
                #     rows = wait.until(EC.presence_of_all_elements_located(
                #         (By.XPATH, "//table[contains(@id,'authExportItemTbl')]//tbody/tr")
                #     ))
                
                #     for row in rows:
                #         cols = [col.text.strip() for col in row.find_elements(By.XPATH, ".//td")]
                #         if any(cols):  # avoid blank rows
                #             all_data.append(cols)
                
                #     # --- Locate and check Next button ---
                #     try:
                #         next_btn = self.driver.find_element(By.ID, "authExportItemTbl_next")
                #         if "disabled" in next_btn.get_attribute("class"):
                #             print("🟢 Reached last page.")
                #             break
                #         self.driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
                #         time.sleep(0.5)
                #         self.driver.execute_script("arguments[0].click();", next_btn)
                #         print("➡️ Clicked Next Page... waiting to load new data...")
                #         time.sleep(2.5)
                #     except Exception as e:
                #         print("⚠️ Pagination ended or Next not found:", e)
                #         break
                
                # # --- Convert to DataFrame ---
                # df = pd.DataFrame(all_data, columns=headers if headers else None)
                
                # # --- Save to local Downloads folder ---
                # downloads_path = os.path.join(os.path.expanduser("~"), "Downloads", "EPCG_EXPORT_Data.xlsx")
                # df.to_excel(downloads_path, index=False)
                
                # print(f"\n✅ Data extracted successfully with headers and saved to:\n{downloads_path}")

                # time.sleep(1)
                # self.driver.refresh()


                # my_dashboard = wait.until(
                #     EC.element_to_be_clickable((By.XPATH, '//*[@id="200928"]/a'))
                # )
                # my_dashboard.click()
                # print("✅ Clicked on my dashboard")
                # repo = wait.until(
                #     EC.element_to_be_clickable((By.XPATH, '//*[@id="98000023"]/a'))
                # )
                # repo.click()
                # print("✅ Clicked on repositories")
    
                # Bill_repo = wait.until(
                #     EC.element_to_be_clickable((By.XPATH, '//*[@id="content"]/div[3]/div/div[2]/div[1]/div/a'))
                # )
                # Bill_repo.click()
                # print("✅ Clicked on Bill repositories")
                
                # for i in range(2):
                    
                #     wait = WebDriverWait(self.driver, 15)
                #     dropdown = wait.until(EC.element_to_be_clickable((By.ID, "txt_selectBill")))
                #     dropdown.click()
                #     time.sleep(1)

                #     if i==0:
                    
                #        # Select by visible text using XPath
                #        option = wait.until(EC.element_to_be_clickable((By.XPATH, "//option[text()='Shipping Bill']")))
                #        option.click()
                #        time.sleep(2)
                #        print("select shipping bill")

                #     else:
                #         # Select by visible text using XPath
                #        option = wait.until(EC.element_to_be_clickable((By.XPATH, "//option[text()='Bill of Entry']")))
                #        option.click()
                #        time.sleep(2)
                #        print("select Bill of Entry")
    
                #     for idx, row in enumerate(self.excel_data):
                #         print(row)
                    
                #         # --- Fix: handle pandas Timestamp directly ---
                #         shipping_date = row.get("ADV Shipping Bill Date")
                #         if pd.isna(shipping_date):
                #             print(f"Skipping Row {idx + 1}: No Shipping Bill Date found")
                #             continue
                    
                #         # Ensure it's a datetime object
                #         if isinstance(shipping_date, pd.Timestamp):
                #             shipping_date = shipping_date.to_pydatetime()
                    
                #         shipping_bill_date = shipping_date.strftime("%d/%m/%Y")
                #         print("Formatted Shipping Bill Date:", shipping_bill_date)
                    
                #         # --- Wait for date field and fill it ---
                #         wait.until(EC.presence_of_element_located((By.ID, "fromDateOfSelectedBil")))
                #         print("shown shipping bill date")
                    
                #         self.driver.execute_script(f"""
                #             var fromDate = document.getElementById('fromDateOfSelectedBil');
                #             fromDate.value = '{shipping_bill_date}';
                #             fromDate.dispatchEvent(new Event('change'));
                #         """)
                    
                #         time.sleep(1)

                        

                    
                #         # --- Fill Authorisation Number ---
                #         auth_no = str(row.get("ADV Authorisation Number", "")).strip()
                #         if i==0:
                #             self.driver.find_element(By.ID, "authorisationNo").clear()
                #             self.driver.find_element(By.ID, "authorisationNo").send_keys(auth_no)
                #         else:
                #             self.driver.find_element(By.ID, "boeLicenseNumber").clear()
                #             self.driver.find_element(By.ID, "boeLicenseNumber").send_keys(auth_no)
                        

                #         print("add authorisation number")
                    
                #         # --- Click Search ---
                #         search = wait.until(EC.element_to_be_clickable((By.ID, 'repSearchBtn')))
                #         search.click()
                #         print("✅ Clicked search")

                #         if i==0:
    
                #             # --- Wait for table to appear ---
                #             wait.until(EC.presence_of_element_located((By.ID, "billRepositoryTable")))
                #             print("table shown")
                            
                #             all_rows = []
                            
                #             while True:
                #                 # Wait for the table body to load
                #                 wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#billRepositoryTable tbody tr")))
                            
                #                 rows = self.driver.find_elements(By.CSS_SELECTOR, "#billRepositoryTable tbody tr")
                #                 print(f"Found {len(rows)} rows on this page")
                            
                #                 for r in rows:
                #                     cols = r.find_elements(By.TAG_NAME, "td")
                #                     if len(cols) > 1:  # skip "No data available" row
                #                         data = [c.text.strip() for c in cols]
                #                         all_rows.append(data)
                            
                #                 # Try to find "Next" button and check if it is enabled
                #                 try:
                #                     next_btn = self.driver.find_element(By.ID, "billRepositoryTable_next")
                #                     next_class = next_btn.get_attribute("class")
                #                     if "disabled" in next_class:
                #                         print("✅ No more pages. Extraction complete.")
                #                         break
                #                     else:
                #                         self.driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
                #                         next_btn.click()
                #                         time.sleep(2)  # Wait for next page data to load
                #                 except Exception as e:
                #                     print("⚠️ Pagination ended or not found:", e)
                #                     break
                            
                #             # --- Extract headers ---
                #             headers = [h.text.strip() for h in self.driver.find_elements(By.CSS_SELECTOR, "#billRepositoryTable thead th")]
                            
                #             # --- Save to Excel ---
                #             # --- Save to Local Downloads Folder ---
                #             downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
                #             timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                #             filename = f"ADV_shipping_bill_Data_{timestamp}.xlsx"
                #             file_path = os.path.join(downloads_path, filename)                    
        
                #             df = pd.DataFrame(all_rows, columns=headers)
                #             df.to_excel(file_path, index=False)                    
        
                #             print(f"💾 Data saved to: {file_path}")

                #         else:
                #                                         # --- Wait for Bill of Entry table to load ---
                #             wait.until(EC.presence_of_element_located((By.ID, "billOfEntryTable")))
                            
                #             all_rows = []
                            
                #             while True:
                #                 # Wait for "Processing..." to disappear
                #                 try:
                #                     wait.until_not(EC.visibility_of_element_located((By.ID, "billOfEntryTable_processing")))
                #                 except:
                #                     pass
                            
                #                 # Wait for table rows
                #                 wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#billOfEntryTable tbody tr")))
                            
                #                 rows = self.driver.find_elements(By.CSS_SELECTOR, "#billOfEntryTable tbody tr")
                #                 print(f"Found {len(rows)} rows on this page")
                            
                #                 for r in rows:
                #                     cols = r.find_elements(By.TAG_NAME, "td")
                #                     if len(cols) > 1:  # skip "No data available"
                #                         data = [c.text.strip() for c in cols]
                #                         all_rows.append(data)
                            
                #                 # --- Handle Pagination ---
                #                 try:
                #                     next_btn = self.driver.find_element(By.ID, "billOfEntryTable_next")
                #                     next_class = next_btn.get_attribute("class")
                #                     if "disabled" in next_class:
                #                         print("✅ No more pages. Extraction complete.")
                #                         break
                #                     else:
                #                         self.driver.execute_script("arguments[0].scrollIntoView(true);", next_btn)
                #                         next_btn.click()
                #                         time.sleep(2)  # allow table to refresh
                #                 except Exception as e:
                #                     print("⚠️ Pagination ended or not found:", e)
                #                     break
                            
                #             # --- Extract Table Headers ---
                #             headers = [h.text.strip() for h in self.driver.find_elements(By.CSS_SELECTOR, "#billOfEntryTable thead th")]
                            
                #             # --- Save to Local Downloads Folder ---
                #             downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
                #             timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                #             filename = f"ADV_BillOfEntry_Data_{timestamp}.xlsx"
                #             file_path = os.path.join(downloads_path, filename)
                            
                #             df = pd.DataFrame(all_rows, columns=headers)
                #             df.to_excel(file_path, index=False)
                            
                #             print(f"💾 Data saved to: {file_path}")

                # time.sleep(1)

                time.sleep(1)
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

                time.sleep(1)
                self.driver.execute_script("document.body.style.zoom='80%'")
                time.sleep(0.5)
                self.driver.execute_script("document.body.style.zoom='60%'")
                time.sleep(0.5)
                # self.driver.execute_script("document.body.style.zoom='100%'")
                # time.sleep(0.5)
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

                # Export Realization details

                # ensure table is fully drawn
                wait.until(EC.presence_of_element_located((By.ID, "exportRealizationTbl")))
                wait.until(EC.visibility_of_element_located((By.ID, "exportRealizationTbl")))
                
                # re-fetch table each time to avoid stale DOM
                rel_table = self.driver.find_element(By.ID, "exportRealizationTbl")
                
                # -----------------------------
                # HEADERS
                # -----------------------------
                rel_headers = []
                header_elems = rel_table.find_elements(By.CSS_SELECTOR, "thead th")
                
                if not header_elems:
                    time.sleep(1)
                    rel_table = self.driver.find_element(By.ID, "exportRealizationTbl")
                    header_elems = rel_table.find_elements(By.CSS_SELECTOR, "thead th")
                
                for th in header_elems:
                    rel_headers.append(th.text.strip())
                
                # -----------------------------
                # ROWS
                # -----------------------------
                rel_all_rows = []
                row_elems = rel_table.find_elements(By.CSS_SELECTOR, "tbody tr")
                
                if not row_elems:
                    time.sleep(1)
                    rel_table = self.driver.find_element(By.ID, "exportRealizationTbl")
                    row_elems = rel_table.find_elements(By.CSS_SELECTOR, "tbody tr")
                
                for tr in row_elems:
                    cells = tr.find_elements(By.TAG_NAME, "td")
                    rel_all_rows.append([td.text.strip() for td in cells])
                
                # -----------------------------
                # SAVE EXCEL
                # -----------------------------
                downloads = os.path.join(os.path.expanduser("~"), "Downloads")
                file_path = os.path.join(downloads, "export_realization.xlsx")
                
                df = pd.DataFrame(rel_all_rows, columns=rel_headers)
                df.to_excel(file_path, index=False)
                
                print("DONE")
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
                file_path = os.path.join(downloads, "input_As_endorsed_auth.xlsx")
                
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
                b_file_path = os.path.join(downloads, "bill_of_entry.xlsx")
                
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
                c_file_path = os.path.join(downloads, "input_importer_exporter.xlsx")
                
                df = pd.DataFrame(c_all_rows, columns=c_headers)
                df.to_excel(c_file_path, index=False)
                
                print("Saved:", c_file_path)


     
            else:
                time.sleep(2)
    
            
        

        except Exception as e:
            return {"success": False, "message": f"Error in fill_certificate: {e}"}

    # def close_browser(self):
    #     if self.driver:
    #         self.driver.quit()