#!/usr/bin/env python3
"""
CRM Auto Login Script
Automates login to the CRM system at http://192.168.1.152/crm/eware.dll/go
"""

from __future__ import annotations
import time
import sys
from DrissionPage import ChromiumPage, ChromiumOptions, Chromium
from DrissionPage.errors import BrowserConnectError

class CRMAutoLogin:
    def __init__(self, headless: bool = False, invoice_number: str = None):
        self.url = "http://192.168.1.152/crm/eware.dll/go"
        self.username = "ivan.chiu"
        self.password = "12345678"
        self.page = None
        self.headless = headless
        self.invoice_number = invoice_number
        
    def setup_browser(self):
        """Setup DrissionPage browser"""
        co = ChromiumOptions()
        co.headless(self.headless)
        
        try:
            self.page = ChromiumPage(co)
            print("✅ DrissionPage browser initialized successfully")
            return True
        except BrowserConnectError:
            # If no debuggable instance is running, launch one ourselves
            print("[DrissionPage] 未检测到已开启的调试端口，正在自动启动新的浏览器实例…")
            try:
                co = ChromiumOptions()
                co.headless(self.headless)
                # On macOS you can override Chrome path if needed:
                # co.set_browser_path("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome")
                
                browser = Chromium(co)
                self.page = browser.new_tab()
                print("✅ DrissionPage browser launched successfully")
                return True
            except Exception as e:
                print(f"❌ Failed to initialize DrissionPage browser: {e}")
                return False
        except Exception as e:
            print(f"❌ Failed to initialize DrissionPage browser: {e}")
            return False
    
    def open_website(self):
        """Open the CRM website"""
        try:
            print(f"🌐 Opening website: {self.url}")
            self.page.get(self.url)
            print("✅ Website opened successfully")
            return True
        except Exception as e:
            print(f"❌ Failed to open website: {e}")
            return False
    
    def login(self):
        """Perform login with credentials"""
        try:
            print("🔐 Starting login process...")
            
            # Wait for the page to load
            time.sleep(2)
            
            # Find username field by name attribute
            username_field = self.page.ele('@name=EWARE_USERID')
            if not username_field:
                print("❌ Username field not found")
                return False
            print("✅ Found username field")
            
            # Clear and enter username
            username_field.clear()
            username_field.input(self.username)
            print(f"✅ Entered username: {self.username}")
            
            # Find password field by name attribute
            password_field = self.page.ele('@name=PASSWORD')
            if not password_field:
                print("❌ Password field not found")
                return False
            print("✅ Found password field")
            
            # Clear and enter password
            password_field.clear()
            password_field.input(self.password)
            print("✅ Entered password")
            
            # Find and click the login button using the class name
            login_button = self.page.ele('.Logonbutton')
            if not login_button:
                print("❌ Login button not found")
                return False
            print("✅ Found login button")
            
            # Click the login button
            login_button.click()
            print("✅ Clicked login button")
            
            # Wait a moment for the login to process and next page to load
            time.sleep(3)
            
            # Find all Logonbutton elements
            logon_buttons = self.page.eles('.Logonbutton')
            if len(logon_buttons) < 3:
                print(f"❌ Found only {len(logon_buttons)} Logonbutton(s), need at least 3 for the third one")
                return False
            print(f"✅ Found {len(logon_buttons)} Logonbutton elements")
            
            # Select the third one (index 2)
            third_button = logon_buttons[2]
            
            # Click the inner <a> if present, else the <td>
            inner_link = third_button.ele('tag:a')
            if inner_link:
                inner_link.click()
            else:
                third_button.click()
            print("✅ Clicked the third Logonbutton")
            
            # Wait a bit after clicking
            time.sleep(3)
            
            print("✅ 3-second wait completed")
            
            # Find and click the Find button by ID
            find_button = self.page.ele('@id=Find')
            if not find_button:
                # Fallback to finding by text or class
                find_button = self.page.ele('txt:Find') or self.page.ele('.MENUITEM')
            if not find_button:
                print("❌ Find button not found")
                return False
            print("✅ Found Find button")
            
            find_button.click()
            print("✅ Clicked Find button")
            
            # Optional short wait after clicking
            time.sleep(2)
            
            # Select 'Opportunities' from the dropdown
            select_elem = self.page.ele('@id=SELECTMenuOption')
            if not select_elem:
                print("❌ Dropdown SELECTMenuOption not found")
                return False
            print("✅ Found SELECTMenuOption dropdown")
            
            # Find the 'Opportunities' option and select it
            options = select_elem.eles('tag:option')
            found = False
            for option in options:
                if option.text.strip().lower() == 'opportunities':
                    option.click()
                    found = True
                    print("✅ Selected 'Opportunities' option")
                    break
            if not found:
                print("❌ 'Opportunities' option not found in dropdown")
                return False
            
            # Wait a bit for the page to load after selection
            time.sleep(2)
            
            return True
            
        except Exception as e:
            print(f"❌ Login failed: {e}")
            return False
    
    def keep_browser_open(self):
        """Keep the browser open for user interaction"""
        try:
            print("\n🎉 Login completed!")
            print("The browser will remain open for you to use the CRM system.")
            print("Press Ctrl+C to close the browser and exit the script.")
            
            # Keep the script running until user interrupts
            while True:
                time.sleep(1)
                
        except KeyboardInterrupt:
            print("\n👋 Closing browser...")
            self.close()
    
    def close(self):
        """Close the browser"""
        if self.page:
            try:
                # Try to close the current page/tab
                self.page.close()
            except Exception:
                pass
            
            # Try to gracefully quit the browser
            browser = getattr(self.page, "browser", None)
            if browser:
                for method in ("quit", "close", "kill"):
                    if hasattr(browser, method):
                        try:
                            getattr(browser, method)()
                            break 
                        except Exception:
                            continue
            print("✅ Browser closed")
    
    def run(self):
        """Main execution method"""
        print("🚀 Starting CRM Auto Login...")
        print("=" * 50)
        
        # Setup browser
        if not self.setup_browser():
            return False
        
        try:
            # Open website
            if not self.open_website():
                return False
            
            # Perform login and navigation
            if not self.login():
                print("❌ Login process failed")
                return False
            
            # If invoice number provided, input it
            if self.invoice_number:
                invoice_field = self.page.ele('@id=oppo_afwinvno')
                if not invoice_field:
                    print("❌ Invoice input field not found")
                else:
                    invoice_field.clear()
                    invoice_field.input(self.invoice_number)
                    print(f"✅ Entered invoice number: {self.invoice_number}")
                    time.sleep(1)
                    # Click the Search/Find button
                    search_button = self.page.ele("css:a.ButtonItem[href*='EntryForm.submit']")
                    if not search_button:
                        # Fallback: look for anchor containing image Search.gif or text 'Find'
                        search_button = self.page.ele("css:a.ButtonItem img[src*='Search.gif']") or \
                                         self.page.ele("txt:Find")
                    if search_button:
                        search_button.click()
                        print("✅ Clicked Search/Find button")
                    else:
                        print("⚠️ Search/Find button not found - you may need to adjust selector")
                time.sleep(2)  # Short wait after input and search
            
            # Keep browser open for user interaction
            self.keep_browser_open()
            
        except Exception as e:
            print(f"❌ Unexpected error: {e}")
            return False
        finally:
            self.close()
        
        return True

def main():
    """Main function"""
    import argparse
    import subprocess
    import sys
    import json
    
    parser = argparse.ArgumentParser(description="CRM Auto Login using DrissionPage")
    parser.add_argument("--headless", action="store_true", help="Run browser in headless mode")
    parser.add_argument("--invoice", help="Invoice number to input after navigation (overrides auto-detect)")
    parser.add_argument("--image", default="bank.png", help="Path to bank payment advice image (default: bank.png)")
    args = parser.parse_args()
    
    invoice_number = args.invoice
    
    # If no invoice provided, run imagedetect.py to extract it
    if not invoice_number:
        print("🔎 Running imagedetect.py to extract invoice number from image...")
        try:
            result = subprocess.run([
                sys.executable, "imagedetect.py", args.image
            ], capture_output=True, text=True, check=True)
            output = result.stdout
            # Try to find the line with 'Invoice:'
            invoice_number = None
            for line in output.splitlines():
                if line.lower().startswith("invoice:"):
                    invoice_number = line.split(":", 1)[-1].strip()
                    break
            if invoice_number:
                print(f"✅ Detected invoice number: {invoice_number}")
            else:
                print("⚠️ Could not detect invoice number from image. Proceeding without auto-input.")
        except Exception as e:
            print(f"❌ Error running imagedetect.py: {e}")
            invoice_number = None
    
    auto_login = CRMAutoLogin(headless=args.headless, invoice_number=invoice_number)
    success = auto_login.run()
    
    if not success:
        print("\n❌ Script execution failed")
        sys.exit(1)
    else:
        print("\n✅ Script execution completed")
        sys.exit(0)

if __name__ == "__main__":
    main() 