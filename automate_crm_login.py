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
    def __init__(self, headless: bool = False, invoice_number: str = None, return_json: bool = False):
        self.url = "http://192.168.1.152/crm/eware.dll/go"
        self.username = "ivan.chiu"
        self.password = "12345678"
        self.headless = headless
        self.invoice_number = invoice_number
        self.return_json = return_json
        
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

            # --- Switch to main content frame (EWARE_MID) to access dropdown ---
            self.driver.switch_to.default_content()
            try:
                self.driver.switch_to.frame("EWARE_MID")
            except Exception:
                WebDriverWait(self.driver, 10).until(
                    EC.frame_to_be_available_and_switch_to_it((By.NAME, "EWARE_MID"))
                )

            select_elem = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "SELECTMenuOption"))
            )
            if not select_elem:
                print("❌ Dropdown SELECTMenuOption not found", file=sys.stderr)
                return False
            print("✅ Found SELECTMenuOption dropdown", file=sys.stderr)
            options = select_elem.find_elements(By.TAG_NAME, "option")
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
            time.sleep(2)
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
    
    def run(self):
        """Main execution method"""
        print("🚀 Starting CRM Auto Login...", file=sys.stderr)
        print("=" * 50, file=sys.stderr)
        if not self.setup_browser():
            return False
        try:
            if not self.open_website():
                return False
            if not self.login():
                print("❌ Login process failed", file=sys.stderr)
                return False
            if self.invoice_number:
                invoice_field = self.driver.find_element(By.ID, "oppo_afwinvno")
                if not invoice_field:
                    print("❌ Invoice input field not found", file=sys.stderr)
                else:
                    invoice_field.clear()
                    invoice_field.send_keys(self.invoice_number)
                    print(f"✅ Entered invoice number: {self.invoice_number}", file=sys.stderr)
                    time.sleep(1)
                    search_button = self.driver.find_element(By.CSS_SELECTOR, "a.ButtonItem[href*='EntryForm.submit']")
                    if not search_button:
                        search_button = self.driver.find_element(By.CSS_SELECTOR, "a.ButtonItem img[src*='Search.gif']") or self.driver.find_element(By.NAME, "txt:Find")
                    if search_button:
                        search_button.click()
                        print("✅ Clicked Search/Find button", file=sys.stderr)
                        time.sleep(5)
                    else:
                        print("⚠️ Search/Find button not found - you may need to adjust selector", file=sys.stderr)
                time.sleep(2)
            # Always parse and print results after search
            time.sleep(2)
            # After clicking the search/find button and waiting:
            self.driver.switch_to.frame('EWARE_MID')  # Use the correct frame name
            find_button = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "Find"))
            )
            html = self.driver.page_source
            records = self._parse_results(html)
            self.driver.switch_to.default_content() # Switch back if you need to interact with the main page
            if self.return_json:
                print(json.dumps(records, ensure_ascii=False))
            else:
                # Print as readable table
                if records:
                    headers = list(records[0].keys())
                    row_format = " | ".join(["{:>15}"] * len(headers))
                    print("\n" + row_format.format(*headers))
                    print("-" * (18 * len(headers)))
                    for rec in records:
                        print(row_format.format(*[str(rec[h]) for h in headers]))
                else:
                    print("No records found.")
            if not self.return_json:
                self.keep_browser_open()
        except Exception as e:
            print(f"❌ Unexpected error: {e}", file=sys.stderr)
            return False
        finally:
            self.close()
        return True

    def _parse_results(self, html: str):
        """Parse CRM table rows into list of dicts, matching the actual table columns dynamically and robustly."""
        soup = BeautifulSoup(html, 'html.parser')
        table = soup.find('table', class_='CONTENT')
        # Parse table
        if not table:
            return []
        rows = table.find_all('tr')
        if not rows or len(rows) < 2:
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
            used_headers = headers[-n:]
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
        return results

def main():
    import argparse
    import subprocess
    parser = argparse.ArgumentParser(description="CRM Auto Login using Selenium")
    parser.add_argument("--headless", action="store_true", help="Run browser in headless mode")
    parser.add_argument("--invoice", help="Invoice number to input after navigation (overrides auto-detect)")
    parser.add_argument("--image", default="bank.png", help="Path to bank payment advice image (default: bank.png)")
    parser.add_argument("--json", action="store_true", help="Return search table as JSON and exit (no interactive browser)")
    args = parser.parse_args()
    invoice_number = args.invoice
    if not invoice_number and not args.json:
        print("🔎 Running imagedetect.py to extract invoice number from image...", file=sys.stderr)
        try:
            result = subprocess.run(
                [sys.executable, "imagedetect.py", args.image],
                capture_output=True, text=True, check=True
            )
            output = result.stdout
            invoice_number = None
            for line in output.splitlines():
                if line.lower().startswith("invoice:"):
                    invoice_number = line.split(":", 1)[-1].strip()
                    break
            if invoice_number:
                print(f"✅ Detected invoice number: {invoice_number}", file=sys.stderr)
            else:
                print("⚠️ Could not detect invoice number from image.", file=sys.stderr)
        except Exception as e:
            print(f"❌ Error running imagedetect.py: {e}", file=sys.stderr)
            invoice_number = None
    auto_login = CRMAutoLogin(
        headless=args.headless, invoice_number=invoice_number, return_json=args.json
    )
    success = auto_login.run()
    if not success:
        print("\n❌ Script execution failed", file=sys.stderr)
        sys.exit(1)
    else:
        if not args.json:
            print("\n✅ Script execution completed", file=sys.stderr)
        sys.exit(0)

if __name__ == "__main__":
    main() 