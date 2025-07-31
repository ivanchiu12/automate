#!/usr/bin/env python3
"""
CRM Auto Login Script
Automates login to the CRM system at http://192.168.1.152/crm/eware.dll/go
"""

from __future__ import annotations
import time
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import json

class CRMAutoLogin:
    def __init__(self, headless: bool = False, invoice_number: str = None, invoice_numbers: list = None, return_json: bool = False, no_interactive: bool = False, web_output: bool = False):
        self.url = "http://192.168.1.152/crm/eware.dll/go"
        self.username = "ivan.chiu"
        self.password = "12345678"
        self.headless = headless
        self.invoice_number = invoice_number
        self.invoice_numbers = invoice_numbers or ([invoice_number] if invoice_number else [])
        self.return_json = return_json
        self.no_interactive = no_interactive
        self.web_output = web_output
        
    def setup_browser(self):
        """Setup Selenium Chrome browser"""
        chrome_options = Options()
        if self.headless:
            chrome_options.add_argument('--headless')
            chrome_options.add_argument('--disable-gpu')
            chrome_options.add_argument('--window-size=1920,1080')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        try:
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            print("✅ Selenium Chrome browser initialized successfully", file=sys.stderr)
            return True
        except Exception as e:
            print(f"❌ Failed to initialize Selenium Chrome browser: {e}", file=sys.stderr)
            return False
    
    def open_website(self):
        """Open the CRM website"""
        try:
            print(f"🌐 Opening website: {self.url}", file=sys.stderr)
            self.driver.get(self.url)
            print("✅ Website opened successfully", file=sys.stderr)
            return True
        except Exception as e:
            print(f"❌ Failed to open website: {e}", file=sys.stderr)
            return False
    
    def login(self):
        """Perform login with credentials"""
        try:
            print("🔐 Starting login process...", file=sys.stderr)
            time.sleep(2)
            username_field = self.driver.find_element(By.NAME, "EWARE_USERID")
            if not username_field:
                print("❌ Username field not found", file=sys.stderr)
                return False
            print("✅ Found username field", file=sys.stderr)
            username_field.clear()
            username_field.send_keys(self.username)
            print(f"✅ Entered username: {self.username}", file=sys.stderr)
            password_field = self.driver.find_element(By.NAME, "PASSWORD")
            if not password_field:
                print("❌ Password field not found", file=sys.stderr)
                return False
            print("✅ Found password field", file=sys.stderr)
            password_field.clear()
            password_field.send_keys(self.password)
            print("✅ Entered password", file=sys.stderr)
            login_button = self.driver.find_element(By.CLASS_NAME, "Logonbutton")
            if not login_button:
                print("❌ Login button not found", file=sys.stderr)
                return False
            print("✅ Found login button", file=sys.stderr)
            login_button.click()
            print("✅ Clicked login button", file=sys.stderr)
            time.sleep(3)
            logon_buttons = self.driver.find_elements(By.CLASS_NAME, "Logonbutton")
            if len(logon_buttons) < 3:
                print(f"❌ Found only {len(logon_buttons)} Logonbutton(s), need at least 3 for the third one", file=sys.stderr)
                return False
            print(f"✅ Found {len(logon_buttons)} Logonbutton elements", file=sys.stderr)
            third_button = logon_buttons[2]
            inner_link = third_button.find_element(By.TAG_NAME, "a")
            if inner_link:
                inner_link.click()
            else:
                third_button.click()
            print("✅ Clicked the third Logonbutton", file=sys.stderr)
            time.sleep(3)
            print("✅ 3-second wait completed", file=sys.stderr)
            # --- Switch to left navigation frame (EWARE_MENU) and click the Find button ---
            self.driver.switch_to.default_content()
            try:
                self.driver.switch_to.frame("EWARE_MENU")
            except Exception:
                print("⚠️ Could not switch to frame 'EWARE_MENU' – trying alternative method", file=sys.stderr)
                menu_frame = WebDriverWait(self.driver, 10).until(
                    EC.frame_to_be_available_and_switch_to_it((By.NAME, "EWARE_MENU"))
                )
            try:
                find_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "Find"))
                )
                print("✅ Found Find button", file=sys.stderr)
                find_button.click()
                print("✅ Clicked Find button", file=sys.stderr)
            except Exception as e:
                print(f"❌ Find button not found or not clickable: {e}", file=sys.stderr)
                return False

            # --- Switch to top frame (EWARE_TOP) to access dropdown ---
            self.driver.switch_to.default_content()
            try:
                self.driver.switch_to.frame("EWARE_TOP")
            except Exception:
                print("⚠️ Could not switch to frame 'EWARE_TOP' – trying alternative method", file=sys.stderr)
                WebDriverWait(self.driver, 10).until(
                    EC.frame_to_be_available_and_switch_to_it((By.NAME, "EWARE_TOP"))
                )

            select_elem = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "SELECTMenuOption"))
            )
            print("✅ Found SELECTMenuOption dropdown", file=sys.stderr)
            
            # Get all options and print them for debugging
            options = select_elem.find_elements(By.TAG_NAME, "option")
            print(f"Found {len(options)} options in dropdown:", file=sys.stderr)
            for i, option in enumerate(options):
                print(f"  {i}: '{option.text}'", file=sys.stderr)
            
            # Select Opportunities option
            found = False
            for option in options:
                if option.text.strip().lower() == 'opportunities':
                    option.click()
                    found = True
                    print("✅ Selected 'Opportunities' option", file=sys.stderr)
                    break
            if not found:
                print("❌ 'Opportunities' option not found in dropdown", file=sys.stderr)
                return False
            time.sleep(3)  # Wait for page to load after selecting Opportunities
            
            # Switch back to main content frame for invoice input
            self.driver.switch_to.default_content()
            try:
                self.driver.switch_to.frame("EWARE_MID")
            except Exception:
                print("⚠️ Could not switch to frame 'EWARE_MID' – trying alternative method", file=sys.stderr)
                WebDriverWait(self.driver, 10).until(
                    EC.frame_to_be_available_and_switch_to_it((By.NAME, "EWARE_MID"))
                )
            return True
        except Exception as e:
            print(f"❌ Login failed: {e}", file=sys.stderr)
            return False
    
    def keep_browser_open(self):
        """Keep the browser open for user interaction"""
        try:
            print("\n🎉 Login completed!", file=sys.stderr)
            print("The browser will remain open for you to use the CRM system.", file=sys.stderr)
            print("Press Ctrl+C to close the browser and exit the script.", file=sys.stderr)
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("\n👋 Closing browser...", file=sys.stderr)
            self.close()
    
    def close(self):
        """Close the browser"""
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            print("✅ Browser closed", file=sys.stderr)
    
    def search_invoice(self, invoice_number):
        """Search for a single invoice and return results"""
        print(f"🔍 Searching for invoice: {invoice_number}", file=sys.stderr)
        
        # Make sure we're in the correct frame for invoice input
        self.driver.switch_to.default_content()
        try:
            self.driver.switch_to.frame("EWARE_MID")
        except Exception:
            WebDriverWait(self.driver, 10).until(
                EC.frame_to_be_available_and_switch_to_it((By.NAME, "EWARE_MID"))
            )
        
        invoice_field = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.ID, "oppo_afwinvno"))
        )
        if not invoice_field:
            print("❌ Invoice input field not found", file=sys.stderr)
            return []
        
        invoice_field.clear()
        invoice_field.send_keys(invoice_number)
        print(f"✅ Entered invoice number: {invoice_number}", file=sys.stderr)
        time.sleep(1)
        
        # Try different search button selectors
        search_button = None
        try:
            search_button = self.driver.find_element(By.CSS_SELECTOR, "a.ButtonItem[href*='EntryForm.submit']")
        except:
            try:
                search_button = self.driver.find_element(By.CSS_SELECTOR, "a.ButtonItem img[src*='Search.gif']")
            except:
                try:
                    search_button = self.driver.find_element(By.NAME, "Find")
                except:
                    search_button = None
        
        if search_button:
            search_button.click()
            print("✅ Clicked Search/Find button", file=sys.stderr)
            time.sleep(5)
        else:
            print("⚠️ Search/Find button not found - you may need to adjust selector", file=sys.stderr)
            return []
        
        # Parse results
        html = self.driver.page_source
        records = self._parse_results(html)
        print(f"✅ Found {len(records)} records for invoice {invoice_number}", file=sys.stderr)
        
        # Add source invoice to each record
        for record in records:
            record['_source_invoice'] = invoice_number
            
        return records

    def run(self):
        """Main execution method"""
        print("🚀 Starting CRM Auto Login...", file=sys.stderr)
        print("=" * 50, file=sys.stderr)
        if not self.setup_browser():
            return False, []
        try:
            if not self.open_website():
                return False, []
            if not self.login():
                print("❌ Login process failed", file=sys.stderr)
                return False, []
            
            all_records = []
            
            if self.invoice_numbers:
                for i, invoice in enumerate(self.invoice_numbers):
                    print(f"📄 Processing invoice {i+1}/{len(self.invoice_numbers)}: {invoice}", file=sys.stderr)
                    records = self.search_invoice(invoice)
                    # Add record index for identification
                    for record in records:
                        record['_record_index'] = i + 1
                    all_records.extend(records)
                    
                    # Small delay between searches
                    if i < len(self.invoice_numbers) - 1:
                        time.sleep(2)
            elif self.invoice_number:
                # Backward compatibility for single invoice
                records = self.search_invoice(self.invoice_number)
                for record in records:
                    record['_record_index'] = 1
                all_records = records
            
            # Print results
            if self.return_json or self.web_output:
                print(json.dumps(all_records, ensure_ascii=False))
            else:
                # Print as readable table
                if all_records:
                    headers = [h for h in all_records[0].keys() if not h.startswith('_')]
                    row_format = " | ".join(["{:>15}"] * len(headers))
                    print("\n" + row_format.format(*headers))
                    print("-" * (18 * len(headers)))
                    for rec in all_records:
                        print(row_format.format(*[str(rec[h]) for h in headers if not h.startswith('_')]))
                else:
                    print("No records found.")
            
            if not self.return_json and not self.no_interactive and not self.web_output:
                self.keep_browser_open()
            
            return True, all_records
        except Exception as e:
            print(f"❌ Unexpected error: {e}", file=sys.stderr)
            return False, []
        finally:
            self.close()

    def _parse_results(self, html: str):
        """Parse CRM table rows into list of dicts, matching the actual table columns dynamically and robustly."""
        soup = BeautifulSoup(html, 'html.parser')
        
        # Try multiple approaches to find the correct table
        table = None
        
        # First try: Look for table with CONTENT class
        tables = soup.find_all('table', class_='CONTENT')
        print(f"Found {len(tables)} tables with class 'CONTENT'", file=sys.stderr)
        
        if len(tables) > 0:
            # If multiple tables, try to find the one with data rows
            for i, t in enumerate(tables):
                rows = t.find_all('tr')
                data_rows = [r for r in rows if r.find('td', class_='ROW1') or r.find('td', class_='ROW2')]
                print(f"Table {i}: {len(rows)} total rows, {len(data_rows)} data rows", file=sys.stderr)
                if len(data_rows) > 0:
                    table = t
                    print(f"Selected table {i} with {len(data_rows)} data rows", file=sys.stderr)
                    break
        
        # If no table found, try looking for any table with ROW1/ROW2 classes
        if not table:
            all_tables = soup.find_all('table')
            print(f"Searching through {len(all_tables)} total tables", file=sys.stderr)
            for i, t in enumerate(all_tables):
                rows = t.find_all('tr')
                data_rows = [r for r in rows if r.find('td', class_='ROW1') or r.find('td', class_='ROW2')]
                if len(data_rows) > 0:
                    table = t
                    print(f"Found table {i} with {len(data_rows)} data rows", file=sys.stderr)
                    break
        
        if not table:
            print("❌ No table with data rows found", file=sys.stderr)
            return []
            
        rows = table.find_all('tr')
        if not rows or len(rows) < 2:
            print("❌ Table has insufficient rows", file=sys.stderr)
            return []
            
        # Find the first row with GRIDHEAD class for headers
        header_row = None
        for row in rows:
            if row.find('td', class_='GRIDHEAD'):
                header_row = row
                break
        if not header_row:
            header_row = rows[0]
            
        header_cells = header_row.find_all(['td', 'th'])
        headers = []
        for cell in header_cells:
            a = cell.find('a')
            if a and a.get_text(strip=True):
                text = a.get_text(separator=' ', strip=True)
            else:
                text = cell.get_text(separator=' ', strip=True)
            text = text.replace('\xa0', '').strip()
            headers.append(text)
            
        # Remove leading empty headers
        while headers and headers[0] == '':
            headers.pop(0)
            
        print(f"Found headers: {headers}", file=sys.stderr)
        
        results = []
        for row in rows:
            # Only process data rows (ROW1/ROW2)
            if not (row.find('td', class_='ROW1') or row.find('td', class_='ROW2')):
                continue
            cells = row.find_all('td')
            if not cells:
                continue
            # Use only the last N headers for mapping
            n = len(cells)
            used_headers = headers[-n:] if len(headers) >= n else [f'col{i}' for i in range(n)]
            record = {}
            for i, cell in enumerate(cells):
                a = cell.find('a')
                if a and a.get_text(strip=True):
                    text = a.get_text(separator=' ', strip=True)
                else:
                    text = cell.get_text(separator=' ', strip=True)
                text = text.replace('\xa0', '').strip()
                record[used_headers[i] if i < len(used_headers) else f'col{i}'] = text
            results.append(record)
            
        print(f"Parsed {len(results)} records", file=sys.stderr)
        return results

def main():
    import argparse
    import subprocess
    parser = argparse.ArgumentParser(description="CRM Auto Login using Selenium")
    parser.add_argument("--headless", action="store_true", help="Run browser in headless mode")
    parser.add_argument("--invoice", help="Single invoice number to input after navigation")
    parser.add_argument("--invoices", nargs='+', help="Multiple invoice numbers to search")
    parser.add_argument("--image", default="bank.png", help="Path to bank payment advice image or PDF (default: bank.png)")
    parser.add_argument("--json", action="store_true", help="Return search table as JSON and exit (no interactive browser)")
    parser.add_argument("--no-interactive", action="store_true", help="Do not keep browser open after parsing results")
    parser.add_argument("--web-output", action="store_true", help="Output JSON for web integration (implies --json)")
    args = parser.parse_args()
    
    invoice_numbers = args.invoices or []
    if args.invoice:
        invoice_numbers = [args.invoice]
    
    if not invoice_numbers and not args.json:
        print("🔎 Running imagedetect.py to extract invoice numbers from file...", file=sys.stderr)
        try:
            # Import here to avoid circular imports
            from imagedetect import extract_invoice
            detected_invoices, _ = extract_invoice(args.image)
            if detected_invoices:
                invoice_numbers = detected_invoices
                print(f"✅ Detected {len(invoice_numbers)} invoice number(s): {', '.join(invoice_numbers)}", file=sys.stderr)
            else:
                print("⚠️ Could not detect any invoice numbers from file.", file=sys.stderr)
        except Exception as e:
            print(f"❌ Error running imagedetect.py: {e}", file=sys.stderr)
    
    if args.web_output:
        args.json = True  # web_output implies JSON output
    
    auto_login = CRMAutoLogin(
        headless=args.headless, 
        invoice_numbers=invoice_numbers,
        return_json=args.json, 
        no_interactive=args.no_interactive, 
        web_output=args.web_output
    )
    success, records = auto_login.run()
    if not success:
        print("\n❌ Script execution failed", file=sys.stderr)
        sys.exit(1)
    else:
        if not args.json:
            print("\n✅ Script execution completed", file=sys.stderr)
        sys.exit(0)

if __name__ == "__main__":
    main() 