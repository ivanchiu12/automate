#!/usr/bin/env python3
"""
Integrated Flask web server with OCR and CRM automation capabilities.
Combines image processing, text extraction, and CRM automation in one application.
"""

from flask import Flask, render_template, request, redirect, url_for, flash
import os
import uuid
import sys
import json
import time
import argparse
import cv2
import requests
import base64
from matplotlib import pyplot as plt
from dotenv import load_dotenv
from openai import OpenAI
import fitz  # PyMuPDF
from pdf2image import convert_from_path
from PIL import Image
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'tiff', 'bmp', 'gif', 'pdf'}

app = Flask(__name__)
app.secret_key = 'supersecret'  # change in production
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# ===============================
# IMAGE PROCESSING & OCR FUNCTIONS
# ===============================

def preprocess_image(image):
    """Preprocess image for better OCR accuracy"""
    # Convert to grayscale
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    
    # Apply thresholding to binarize
    _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    
    # Denoise
    denoised = cv2.fastNlMeansDenoising(thresh, None, 20, 7, 21)
    
    return denoised

def extract_text_with_api(image):
    """Extract text using Google Cloud Vision API"""
    load_dotenv()
    api_key = os.getenv('GOOGLE_API_KEY')
    if not api_key:
        print("❌ GOOGLE_API_KEY not found in .env file")
        print("💡 Please get an API key from: https://console.cloud.google.com/apis/credentials")
        return ""

    # Convert OpenCV image to base64
    _, buffer = cv2.imencode('.png', image)
    image_base64 = base64.b64encode(buffer).decode('utf-8')

    # Prepare the request for Google Cloud Vision API
    url = f"https://vision.googleapis.com/v1/images:annotate?key={api_key}"
    headers = {
        'Content-Type': 'application/json'
    }
    data = {
        "requests": [
            {
                "image": {
                    "content": image_base64
                },
                "features": [
                    {
                        "type": "TEXT_DETECTION",
                        "maxResults": 1
                    }
                ]
            }
        ]
    }

    try:
        response = requests.post(url, headers=headers, json=data, timeout=30)
        response.raise_for_status()

        result = response.json()

        # Check if Vision API was successful
        if 'responses' in result and len(result['responses']) > 0:
            response_data = result['responses'][0]
            if 'fullTextAnnotation' in response_data:
                parsed_text = response_data['fullTextAnnotation']['text']
                return parsed_text
            elif 'error' in response_data:
                error_msg = response_data['error']['message']
                print(f"❌ Vision API error: {error_msg}")
                return ""
            else:
                print("❌ No text detected in the image")
                return ""
        else:
            print("❌ Invalid response from Vision API")
            return ""

    except requests.exceptions.RequestException as e:
        print(f"❌ Error calling Vision API: {e}")
        return ""
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        return ""

def is_pdf_file(file_path):
    """Check if the file is a PDF"""
    return file_path.lower().endswith('.pdf')

def pdf_to_images(pdf_path):
    """Convert PDF pages to images"""
    try:
        # Convert PDF to images (one image per page)
        images = convert_from_path(pdf_path, dpi=300)
        return images
    except Exception as e:
        print(f"❌ Error converting PDF to images: {e}")
        print("💡 Trying alternative method using PyMuPDF...")
        return pdf_to_images_alternative(pdf_path)

def pdf_to_images_alternative(pdf_path):
    """Alternative PDF to images conversion using PyMuPDF"""
    try:
        import fitz  # PyMuPDF
        doc = fitz.open(pdf_path)
        images = []
        
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            # Render page to image
            mat = fitz.Matrix(2.0, 2.0)  # 2x zoom for better quality
            pix = page.get_pixmap(matrix=mat)
            
            # Convert to PIL Image
            img_data = pix.tobytes("ppm")
            from io import BytesIO
            img = Image.open(BytesIO(img_data))
            images.append(img)
        
        doc.close()
        print(f"✅ Successfully converted {len(images)} pages using PyMuPDF")
        return images
    except Exception as e:
        print(f"❌ Alternative PDF conversion also failed: {e}")
        return []

def extract_text_from_pdf_pages(pdf_path):
    """Extract text from all pages of a PDF, separated by ------"""
    images = pdf_to_images(pdf_path)
    if not images:
        return ""
    
    all_text = []
    for i, image in enumerate(images):
        print(f"📄 Processing page {i+1}/{len(images)}...")
        
        # Convert PIL image to opencv format
        cv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        
        # Preprocess and extract text
        preprocessed = preprocess_image(cv_image)
        page_text = extract_text_with_api(preprocessed)
        
        # Print OCR results for this page
        print(f"\n🔍 OCR Results for Page {i+1}:")
        print("=" * 50)
        print(page_text)
        print("=" * 50)
        
        all_text.append(f"PAGE {i+1}:\n{page_text}")
    
    return "\n------\n".join(all_text)

def parse_bank_info(text):
    """Parse key information using xAI Grok API - handles both single page and multi-page text"""
    load_dotenv()
    api_key = os.getenv('XAI_API_KEY')
    if not api_key:
        print("❌ XAI_API_KEY not found in .env")
        return []
    
    client = OpenAI(
        api_key=api_key,
        base_url="https://api.x.ai/v1",
    )
    
    # Check if this is multi-page text (contains ------)
    if "------" in text:
        prompt = f"""Extract information from this multi-page bank payment advice text. Each page is separated by "------".
                For each page, extract all available information:
                - date: The payment date (if available)
                - amount: The payment amount with currency (if available)
                - payee: The recipient name (if available)
                - payer: The sender name (if available)
                - reference: Any reference number (if available)
                - invoice: The invoice number (e.g., 25-AVS-RES-00109-RN) (if available)
                - page_number: The page number this information came from

                Return as a JSON array of objects. Process ALL pages, even if they don't contain complete payment information or invoice numbers.

                Text:
                {text}

                Respond only with valid JSON array.
                """
    else:
        prompt = f"""Extract the following information from this bank payment advice text as a JSON object:
                - date: The payment date (if available)
                - amount: The payment amount with currency (if available)
                - payee: The recipient name (if available)
                - payer: The sender name (if available)
                - reference: Any reference number (if available)
                - invoice: The invoice number (e.g., 25-AVS-RES-00109-RN) (if available)

                Text:
                {text}

                Respond only with valid JSON.
                """
    
    try:
        response = client.chat.completions.create(
            model="grok-4-0709",
            messages=[
                {"role": "system", "content": "You are a helpful assistant that extracts structured data from text."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1,
        )
        extracted_json = response.choices[0].message.content.strip()
        
        # Remove any markdown code block formatting if present
        if extracted_json.startswith("```json"):
            extracted_json = extracted_json[7:]
        if extracted_json.endswith("```"):
            extracted_json = extracted_json[:-3]
        extracted_json = extracted_json.strip()
        
        parsed_data = json.loads(extracted_json)
        
        # Ensure we always return a list
        if isinstance(parsed_data, dict):
            return [parsed_data]
        elif isinstance(parsed_data, list):
            # Return all records, even if they don't have invoice numbers
            return parsed_data
        else:
            return []
            
    except Exception as e:
        print(f"❌ Error calling xAI API: {e}")
        return []

def extract_invoice(file_path: str) -> tuple[list[str], list[dict]]:
    """Extract invoice numbers and parsed info from image or PDF. Returns (invoice_numbers, parsed_info_list)"""
    if is_pdf_file(file_path):
        print(f"📄 Processing PDF file: {file_path}")
        extracted_text = extract_text_from_pdf_pages(file_path)
        if not extracted_text:
            print(f"❌ Failed to extract text from PDF: {file_path}")
            return [], []
        
        parsed_info_list = parse_bank_info(extracted_text)
        # Include all invoice numbers, including None values for pages without invoices
        invoice_numbers = [info.get('invoice') for info in parsed_info_list]
        return invoice_numbers, parsed_info_list
    else:
        print(f"🖼️ Processing image file: {file_path}")
        image = cv2.imread(file_path)
        if image is None:
            print(f"❌ Failed to load image: {file_path}")
            return [], []

        preprocessed = preprocess_image(image)
        extracted_text = extract_text_with_api(preprocessed)
        parsed_info_list = parse_bank_info(extracted_text)
        # Include all invoice numbers, including None values for pages without invoices
        invoice_numbers = [info.get('invoice') for info in parsed_info_list]
        return invoice_numbers, parsed_info_list

# ===============================
# CRM AUTOMATION FUNCTIONS
# ===============================

class CRMAutoLogin:
    def __init__(self, headless: bool = False, invoice_number: str = None, invoice_numbers: list = None, return_json: bool = False, no_interactive: bool = False, web_output: bool = False):
        self.url = "http://192.168.1.152/crm/eware.dll/go"
        self.username = "ivan.chiu"
        self.password = "25207090"
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
            print("✅ Selenium Chrome browser initialized successfully")
            return True
        except Exception as e:
            print(f"❌ Failed to initialize Selenium Chrome browser: {e}")
            return False
    
    def open_website(self):
        """Open the CRM website"""
        try:
            print(f"🌐 Opening website: {self.url}")
            self.driver.get(self.url)
            print("✅ Website opened successfully")
            return True
        except Exception as e:
            print(f"❌ Failed to open website: {e}")
            return False
    
    def login(self, max_retries=3):
        """Perform login with credentials"""
        retry_count = 0

        while retry_count < max_retries:
            try:
                print("🔐 Starting login process...")
                time.sleep(2)
                
                # First check if we're already logged in by looking for Logonbutton elements
                logon_buttons = self.driver.find_elements(By.CLASS_NAME, "Logonbutton")
                print(f"🔍 Found {len(logon_buttons)} Logonbutton elements on page load")
                
                if len(logon_buttons) == 0:
                    print("✅ No login buttons found - already logged in!")
                    return self._handle_already_logged_in()
                
                username_field = self.driver.find_element(By.NAME, "EWARE_USERID")
                if not username_field:
                    print("❌ Username field not found")
                    return False
                print("✅ Found username field")
                username_field.clear()
                username_field.send_keys(self.username)
                print(f"✅ Entered username: {self.username}")
                password_field = self.driver.find_element(By.NAME, "PASSWORD")
                if not password_field:
                    print("❌ Password field not found")
                    return False
                print("✅ Found password field")
                password_field.clear()
                password_field.send_keys(self.password)
                print("✅ Entered password")
                login_button = self.driver.find_element(By.CLASS_NAME, "Logonbutton")
                if not login_button:
                    print("❌ Login button not found")
                    return False
                print("✅ Found login button")
                login_button.click()
                print("✅ Clicked login button")

                # Wait and check for Logonbutton elements with retry logic
                logon_buttons_found = False
                for attempt in range(5):  # Try up to 5 times with increasing wait
                    time.sleep(3 + attempt)  # Start with 3s, then 4s, 5s, 6s, 7s
                    logon_buttons = self.driver.find_elements(By.CLASS_NAME, "Logonbutton")
                    print(f"🔄 Attempt {attempt + 1}: Found {len(logon_buttons)} Logonbutton elements")

                    if len(logon_buttons) >= 3:
                        logon_buttons_found = True
                        print(f"✅ Successfully found {len(logon_buttons)} Logonbutton elements")
                        break
                    elif attempt < 4:  # Don't print this on the last attempt
                        print(f"⏳ Waiting longer for page to load... ({3 + attempt + 1}s)")

                if not logon_buttons_found:
                    if retry_count < max_retries - 1:
                        print(f"❌ Found only {len(logon_buttons)} Logonbutton(s), need at least 3. Retrying... ({retry_count + 1}/{max_retries})")
                        retry_count += 1
                        # Refresh the page and try again
                        self.driver.refresh()
                        time.sleep(2)
                        continue
                    else:
                        print(f"❌ Found only {len(logon_buttons)} Logonbutton(s), need at least 3 for the third one. All {max_retries} retries failed.")
                        return False

                third_button = logon_buttons[2]
                inner_link = third_button.find_element(By.TAG_NAME, "a")
                if inner_link:
                    inner_link.click()
                else:
                    third_button.click()
                print("✅ Clicked the third Logonbutton")
                time.sleep(3)
                print("✅ 3-second wait completed")

                # If we get here, login was successful
                # --- Switch to left navigation frame (EWARE_MENU) and click the Find button ---
                self.driver.switch_to.default_content()
                try:
                    self.driver.switch_to.frame("EWARE_MENU")
                except Exception:
                    print("⚠️ Could not switch to frame 'EWARE_MENU' – trying alternative method")
                    menu_frame = WebDriverWait(self.driver, 10).until(
                        EC.frame_to_be_available_and_switch_to_it((By.NAME, "EWARE_MENU"))
                    )
                try:
                    find_button = WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "Find"))
                    )
                    print("✅ Found Find button")
                    find_button.click()
                    print("✅ Clicked Find button")
                except Exception as e:
                    print(f"❌ Find button not found or not clickable: {e}")
                    return False

                # --- Switch to top frame (EWARE_TOP) to access dropdown ---
                self.driver.switch_to.default_content()
                try:
                    self.driver.switch_to.frame("EWARE_TOP")
                except Exception:
                    print("⚠️ Could not switch to frame 'EWARE_TOP' – trying alternative method")
                    WebDriverWait(self.driver, 10).until(
                        EC.frame_to_be_available_and_switch_to_it((By.NAME, "EWARE_TOP"))
                    )

                select_elem = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.ID, "SELECTMenuOption"))
                )
                print("✅ Found SELECTMenuOption dropdown")

                # Get all options and print them for debugging
                options = select_elem.find_elements(By.TAG_NAME, "option")
                print(f"Found {len(options)} options in dropdown:")
                for i, option in enumerate(options):
                    print(f"  {i}: '{option.text}'")

                # Select Opportunities option
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
                time.sleep(3)  # Wait for page to load after selecting Opportunities

                # Switch back to main content frame for invoice input
                self.driver.switch_to.default_content()
                try:
                    self.driver.switch_to.frame("EWARE_MID")
                except Exception:
                    print("⚠️ Could not switch to frame 'EWARE_MID' – trying alternative method")
                    WebDriverWait(self.driver, 10).until(
                        EC.frame_to_be_available_and_switch_to_it((By.NAME, "EWARE_MID"))
                    )
                return True

            except Exception as e:
                if retry_count < max_retries - 1:
                    print(f"❌ Login attempt {retry_count + 1} failed: {e}")
                    print(f"🔄 Retrying login... ({retry_count + 2}/{max_retries})")
                    retry_count += 1
                    # Refresh the page and try again
                    self.driver.refresh()
                    time.sleep(2)
                    continue
                else:
                    print(f"❌ Login failed after {max_retries} attempts: {e}")
                    return False

        return False

    def _handle_already_logged_in(self):
        """Handle the case when user is already logged in"""
        try:
            print("🔄 Setting up CRM navigation since already logged in...")
            
            # Switch to left navigation frame (EWARE_MENU) and click the Find button
            self.driver.switch_to.default_content()
            try:
                self.driver.switch_to.frame("EWARE_MENU")
            except Exception:
                print("⚠️ Could not switch to frame 'EWARE_MENU' – trying alternative method")
                menu_frame = WebDriverWait(self.driver, 10).until(
                    EC.frame_to_be_available_and_switch_to_it((By.NAME, "EWARE_MENU"))
                )
            try:
                find_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "Find"))
                )
                print("✅ Found Find button")
                find_button.click()
                print("✅ Clicked Find button")
            except Exception as e:
                print(f"❌ Find button not found or not clickable: {e}")
                return False

            # Switch to top frame (EWARE_TOP) to access dropdown
            self.driver.switch_to.default_content()
            try:
                self.driver.switch_to.frame("EWARE_TOP")
            except Exception:
                print("⚠️ Could not switch to frame 'EWARE_TOP' – trying alternative method")
                WebDriverWait(self.driver, 10).until(
                    EC.frame_to_be_available_and_switch_to_it((By.NAME, "EWARE_TOP"))
                )

            select_elem = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "SELECTMenuOption"))
            )
            print("✅ Found SELECTMenuOption dropdown")

            # Get all options and print them for debugging
            options = select_elem.find_elements(By.TAG_NAME, "option")
            print(f"Found {len(options)} options in dropdown:")
            for i, option in enumerate(options):
                print(f"  {i}: '{option.text}'")

            # Select Opportunities option
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
            time.sleep(3)  # Wait for page to load after selecting Opportunities

            # Switch back to main content frame for invoice input
            self.driver.switch_to.default_content()
            try:
                self.driver.switch_to.frame("EWARE_MID")
            except Exception:
                print("⚠️ Could not switch to frame 'EWARE_MID' – trying alternative method")
                WebDriverWait(self.driver, 10).until(
                    EC.frame_to_be_available_and_switch_to_it((By.NAME, "EWARE_MID"))
                )
            
            print("✅ Successfully set up CRM navigation for already logged in session")
            return True
            
        except Exception as e:
            print(f"❌ Failed to set up CRM navigation: {e}")
            return False
    
    def close(self):
        """Close the browser"""
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            print("✅ Browser closed")
    
    def search_invoice(self, invoice_number):
        """Search for a single invoice and return results"""
        # Check if invoice_number is None or empty - if so, search by fee instead
        if not invoice_number or invoice_number.strip() == '':
            return self.search_by_fee()
        
        print(f"🔍 Searching for invoice: {invoice_number}")
        
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
            print("❌ Invoice input field not found")
            return []
        
        invoice_field.clear()
        invoice_field.send_keys(invoice_number)
        print(f"✅ Entered invoice number: {invoice_number}")
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
            print("✅ Clicked Search/Find button")
            time.sleep(5)
        else:
            print("⚠️ Search/Find button not found - you may need to adjust selector")
            return []
        
        # Parse results
        html = self.driver.page_source
        records = self._parse_results(html)
        print(f"✅ Found {len(records)} records for invoice {invoice_number}")
        
        # Add source invoice to each record
        for record in records:
            record['_source_invoice'] = invoice_number
            
        return records

    def search_by_fee(self, fee_amount=""):
        """Search by fee when no invoice number is available"""
        print(f"🔍 Searching by fee (no invoice number available)")
        
        # Make sure we're in the correct frame for fee input
        self.driver.switch_to.default_content()
        try:
            self.driver.switch_to.frame("EWARE_MID")
        except Exception:
            WebDriverWait(self.driver, 10).until(
                EC.frame_to_be_available_and_switch_to_it((By.NAME, "EWARE_MID"))
            )
        
        # Find the fee input field
        try:
            fee_field = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "oppo_dwnetlicfee"))
            )
        except Exception as e:
            print(f"❌ Fee input field not found: {e}")
            return []
        
        # Clear and enter fee amount (if provided, otherwise leave empty for general search)
        fee_field.clear()
        if fee_amount:
            fee_field.send_keys(str(fee_amount))
            print(f"✅ Entered fee amount: {fee_amount}")
        else:
            print("✅ Fee field cleared for general search")
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
            print("✅ Clicked Search/Find button for fee search")
            time.sleep(5)
        else:
            print("⚠️ Search/Find button not found - you may need to adjust selector")
            return []
        
        # Parse results
        html = self.driver.page_source
        records = self._parse_results(html)
        print(f"✅ Found {len(records)} records for fee search")
        
        # Add source info to each record
        for record in records:
            record['_source_invoice'] = f"fee_search_{fee_amount}" if fee_amount else "fee_search_general"
            
        return records

    def run(self):
        """Main execution method"""
        print("🚀 Starting CRM Auto Login...")
        print("=" * 50)
        if not self.setup_browser():
            return False, []
        try:
            if not self.open_website():
                return False, []
            if not self.login():
                print("❌ Login process failed")
                return False, []
            
            all_records = []
            
            if self.invoice_numbers:
                for i, invoice in enumerate(self.invoice_numbers):
                    print(f"📄 Processing invoice {i+1}/{len(self.invoice_numbers)}: {invoice}")
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
            
            return True, all_records
        except Exception as e:
            print(f"❌ Unexpected error: {e}")
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
        print(f"Found {len(tables)} tables with class 'CONTENT'")
        
        if len(tables) > 0:
            # If multiple tables, try to find the one with data rows
            for i, t in enumerate(tables):
                rows = t.find_all('tr')
                data_rows = [r for r in rows if r.find('td', class_='ROW1') or r.find('td', class_='ROW2')]
                print(f"Table {i}: {len(rows)} total rows, {len(data_rows)} data rows")
                if len(data_rows) > 0:
                    table = t
                    print(f"Selected table {i} with {len(data_rows)} data rows")
                    break
        
        # If no table found, try looking for any table with ROW1/ROW2 classes
        if not table:
            all_tables = soup.find_all('table')
            print(f"Searching through {len(all_tables)} total tables")
            for i, t in enumerate(all_tables):
                rows = t.find_all('tr')
                data_rows = [r for r in rows if r.find('td', class_='ROW1') or r.find('td', class_='ROW2')]
                if len(data_rows) > 0:
                    table = t
                    print(f"Found table {i} with {len(data_rows)} data rows")
                    break
        
        if not table:
            print("❌ No table with data rows found")
            return []
            
        rows = table.find_all('tr')
        if not rows or len(rows) < 2:
            print("❌ Table has insufficient rows")
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
            
        print(f"Found headers: {headers}")
        
        results = []
        seen_records = set()  # To track duplicates
        
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
            has_data = False  # Track if this record has any meaningful data
            
            for i, cell in enumerate(cells):
                a = cell.find('a')
                if a and a.get_text(strip=True):
                    text = a.get_text(separator=' ', strip=True)
                else:
                    text = cell.get_text(separator=' ', strip=True)
                text = text.replace('\xa0', '').strip()
                record[used_headers[i] if i < len(used_headers) else f'col{i}'] = text
                
                # Check if this cell has meaningful data (not empty or just whitespace)
                if text and text.strip():
                    has_data = True
            
            # Only add records that have actual data
            if has_data:
                # Create a unique key for this record to avoid duplicates
                record_key = tuple(sorted(record.items()))
                if record_key not in seen_records:
                    seen_records.add(record_key)
                    results.append(record)
            
        print(f"Parsed {len(results)} records")
        return results

# ===============================
# WEB APPLICATION ROUTES
# ===============================

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part', 'warning')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No selected file', 'warning')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = f"{uuid.uuid4().hex}_{file.filename}"
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            
            # Extract invoices (can be multiple from PDF)
            invoice_numbers, parsed_info_list = extract_invoice(file_path)
            if not parsed_info_list:
                flash('No information extracted from the file.', 'danger')
                return render_template('result.html', records=parsed_info_list, all_crm_rows=[], logs="")
            else:
                # Filter out None invoice numbers for CRM processing
                valid_invoice_numbers = [inv for inv in invoice_numbers if inv is not None]
                flash(f'Extracted {len(parsed_info_list)} page(s) with {len(valid_invoice_numbers)} invoice(s). Launching CRM automation...', 'success')
                
                # Initialize CRM automation directly (no subprocess)
                try:
                    print("🚀 Starting integrated CRM automation...")
                    crm = CRMAutoLogin(
                        headless=True,
                        invoice_numbers=invoice_numbers,  # Include all invoice numbers (including None)
                        return_json=True,
                        no_interactive=True,
                        web_output=True
                    )
                    
                    success, all_crm_rows = crm.run()
                    
                    if success:
                        if all_crm_rows:
                            flash(f'CRM automation completed! Found {len(all_crm_rows)} total records.', 'success')
                        else:
                            flash('CRM automation completed but no records found.', 'warning')
                    else:
                        flash('CRM automation failed. Check logs for details.', 'danger')
                        all_crm_rows = []
                        
                    # Since we're running directly, logs are printed to console
                    combined_logs = "CRM automation completed. Check terminal for detailed logs."
                    
                except Exception as e:
                    flash(f'CRM automation error: {str(e)}', 'danger')
                    all_crm_rows = []
                    combined_logs = f"ERROR: {str(e)}"

                flash(f'Processing completed! Total CRM records found: {len(all_crm_rows)}', 'info')
                return render_template('result.html', records=parsed_info_list, all_crm_rows=all_crm_rows, logs=combined_logs)
        else:
            flash('File type not allowed. Please upload an image.', 'danger')
            return redirect(request.url)
    return render_template('index.html')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5002, debug=True) 