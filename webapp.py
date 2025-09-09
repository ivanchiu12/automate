#!/usr/bin/env python3
"""
Integrated Flask web server with OCR and CRM automation capabilities.
Combines image processing, text extraction, and CRM automation in one application.
"""

from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify
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

def calculate_text_dimensions(text, font_size, font_name='helv'):
    """Calculate approximate text dimensions for given text and font size"""
    try:
        # Rough estimation based on font metrics
        # These are approximations for common fonts
        char_width_ratio = {
            'helv': 0.6,
            'times': 0.55,
            'cour': 0.6,
            'cjk': 0.8,
            'china-ss': 0.8,
            'china-ts': 0.8
        }
        
        # Get character width multiplier
        char_width = char_width_ratio.get(font_name, 0.6) * font_size
        line_height = font_size * 1.2  # Standard line height
        
        lines = text.split('\n')
        max_line_width = max(len(line) for line in lines) if lines else 0
        text_height = len(lines) * line_height
        text_width = max_line_width * char_width
        
        return text_width, text_height
    except Exception as e:
        print(f"⚠️ Error calculating text dimensions: {e}")
        # Return conservative estimates
        return 300, 200

def find_available_spaces_on_page(doc, page_num, required_width=None, required_height=None):
    """Find all available spaces on a PDF page, sorted by preference"""
    try:
        page = doc.load_page(page_num)
        rect = page.rect
        
        # Get existing text blocks to identify occupied areas
        text_blocks = page.get_text("dict")["blocks"]
        occupied_areas = []
        
        for block in text_blocks:
            if "bbox" in block:
                occupied_areas.append(fitz.Rect(block["bbox"]))
        
        # Define potential positions with different sizes
        # Each position has a priority score (lower = better)
        potential_positions = [
            # Top right corner - various sizes
            {'rect': fitz.Rect(rect.width * 0.65, 30, rect.width - 10, 300), 'priority': 1},
            {'rect': fitz.Rect(rect.width * 0.7, 30, rect.width - 10, 250), 'priority': 2},
            {'rect': fitz.Rect(rect.width * 0.75, 30, rect.width - 10, 200), 'priority': 3},
            
            # Bottom right corner - various sizes
            {'rect': fitz.Rect(rect.width * 0.65, rect.height - 250, rect.width - 10, rect.height - 10), 'priority': 4},
            {'rect': fitz.Rect(rect.width * 0.7, rect.height - 200, rect.width - 10, rect.height - 10), 'priority': 5},
            {'rect': fitz.Rect(rect.width * 0.75, rect.height - 150, rect.width - 10, rect.height - 10), 'priority': 6},
            
            # Right margin - full height
            {'rect': fitz.Rect(rect.width * 0.75, 100, rect.width - 10, rect.height - 100), 'priority': 7},
            {'rect': fitz.Rect(rect.width * 0.8, 50, rect.width - 10, rect.height - 50), 'priority': 8},
            
            # Bottom left corner
            {'rect': fitz.Rect(10, rect.height - 250, rect.width * 0.35, rect.height - 10), 'priority': 9},
            {'rect': fitz.Rect(10, rect.height - 200, rect.width * 0.3, rect.height - 10), 'priority': 10},
            
            # Top left corner
            {'rect': fitz.Rect(10, 30, rect.width * 0.35, 300), 'priority': 11},
            {'rect': fitz.Rect(10, 30, rect.width * 0.3, 250), 'priority': 12},
            
            # Full width bottom strip
            {'rect': fitz.Rect(10, rect.height - 200, rect.width - 10, rect.height - 10), 'priority': 13},
            {'rect': fitz.Rect(10, rect.height - 150, rect.width - 10, rect.height - 10), 'priority': 14},
            
            # Left margin
            {'rect': fitz.Rect(10, 100, rect.width * 0.25, rect.height - 100), 'priority': 15},
        ]
        
        # Filter positions based on size requirements
        if required_width or required_height:
            filtered_positions = []
            for pos in potential_positions:
                pos_rect = pos['rect']
                width_ok = not required_width or pos_rect.width >= required_width
                height_ok = not required_height or pos_rect.height >= required_height
                if width_ok and height_ok:
                    filtered_positions.append(pos)
            potential_positions = filtered_positions
        
        # Check for overlaps and return available positions
        available_positions = []
        for pos in potential_positions:
            pos_rect = pos['rect']
            overlap_found = False
            
            for occupied in occupied_areas:
                if pos_rect.intersects(occupied):
                    intersection = pos_rect & occupied
                    overlap_ratio = intersection.get_area() / pos_rect.get_area()
                    if overlap_ratio > 0.3:  # If more than 30% overlap
                        overlap_found = True
                        break
            
            if not overlap_found:
                available_positions.append(pos)
        
        # Sort by priority (lower priority number = better position)
        available_positions.sort(key=lambda x: x['priority'])
        
        return [pos['rect'] for pos in available_positions]
        
    except Exception as e:
        print(f"❌ Error finding available spaces on page {page_num}: {e}")
        # Return a default position
        return [fitz.Rect(400, 400, 850, 750)]

def find_blank_space_on_page(doc, page_num, required_width=None, required_height=None):
    """Find a suitable blank space on a PDF page to insert text"""
    available_spaces = find_available_spaces_on_page(doc, page_num, required_width, required_height)
    
    if available_spaces:
        return available_spaces[0]  # Return the best available space
    else:
        # Fallback to a default position if no space found
        page = doc.load_page(page_num)
        rect = page.rect
        return fitz.Rect(rect.width * 0.65, rect.height - 200, rect.width - 10, rect.height - 10)

def format_single_record_text(result, record_num, max_fields=15):
    """Format a single CRM record for display"""
    record_text = f"Record {record_num}:\n"
    
    # Get all non-internal fields (not starting with _) and display them
    displayed_fields = 0
    
    for key, value in result.items():
        if (not key.startswith('_') and 
            value and 
            str(value).strip() and 
            str(value).strip() not in ['', 'None', '0'] and 
            displayed_fields < max_fields):
            
            # Clean and format the value
            clean_value = str(value).strip()
            # Truncate very long values but keep more characters
            if len(clean_value) > 40:
                clean_value = clean_value[:37] + "..."
            
            # Use shorter field names for space efficiency
            short_key = key
            if len(key) > 12:
                short_key = key[:9] + "..."
            
            record_text += f"• {short_key}: {clean_value}\n"
            displayed_fields += 1
    
    # Add source invoice info
    if result.get('_original_invoice_string'):
        orig_inv = str(result['_original_invoice_string'])
        if len(orig_inv) > 30:
            orig_inv = orig_inv[:27] + "..."
        record_text += f"• Source Inv: {orig_inv}\n"
    
    return record_text

def format_extracted_info_text(parsed_info):
    """Format extracted page info for display"""
    if not parsed_info:
        return ""
    
    info_text = "Extracted Info:\n"
    if parsed_info.get('date'):
        info_text += f"• Date: {parsed_info['date']}\n"
    if parsed_info.get('amount'):
        info_text += f"• Amount: {parsed_info['amount']}\n"
    if parsed_info.get('payee'):
        payee = str(parsed_info['payee'])[:25] + "..." if len(str(parsed_info['payee'])) > 25 else str(parsed_info['payee'])
        info_text += f"• Payee: {payee}\n"
    if parsed_info.get('reference'):
        ref = str(parsed_info['reference'])[:20] + "..." if len(str(parsed_info['reference'])) > 20 else str(parsed_info['reference'])
        info_text += f"• Ref: {ref}\n"
    
    return info_text

def insert_text_with_auto_resize(page, text, available_rect, font_size, font_type, text_color=(0, 0, 0)):
    """Insert text with automatic resizing if it doesn't fit"""
    # Determine fonts to try based on user selection
    if font_type == 'auto':
        fonts_to_try = ["cjk", "china-ss", "china-ts", "helv", "times", "cour"]
    else:
        fonts_to_try = [font_type]
        fallback_fonts = ["cjk", "china-ss", "china-ts", "helv", "times", "cour"]
        for fallback in fallback_fonts:
            if fallback != font_type and fallback not in fonts_to_try:
                fonts_to_try.append(fallback)
    
    # Try different font sizes starting from the requested size
    font_sizes_to_try = [font_size]
    if font_size > 4:
        font_sizes_to_try.extend([font_size - 1, font_size - 2])
    if font_size > 6:
        font_sizes_to_try.append(font_size - 3)
    
    # Add a white background rectangle with black border
    bg_rect = fitz.Rect(available_rect.x0 - 5, available_rect.y0 - 5, 
                      available_rect.x1 + 5, available_rect.y1 + 5)
    page.draw_rect(bg_rect, color=(0, 0, 0), fill=(1, 1, 1), width=2)
    
    # Create inner text area
    inner_rect = fitz.Rect(available_rect.x0 + 5, available_rect.y0 + 5, 
                         available_rect.x1 - 5, available_rect.y1 - 5)
    
    for current_font_size in font_sizes_to_try:
        for font_name in fonts_to_try:
            try:
                # Try to insert text with current font and size
                inserted = page.insert_textbox(inner_rect, text, 
                                             fontsize=current_font_size, 
                                             color=text_color,
                                             fontname=font_name,
                                             align=0)
                
                if inserted >= 0:
                    print(f"✅ Text inserted successfully with font: {font_name}, size: {current_font_size}")
                    return True, font_name, current_font_size
                    
            except Exception as e:
                continue
    
    # If textbox method fails, try alternative point-based insertion
    print("⚠️ Textbox insertion failed, trying point-based method...")
    for current_font_size in font_sizes_to_try:
        for font_name in fonts_to_try:
            try:
                point = fitz.Point(inner_rect.x0, inner_rect.y0 + 15)
                page.insert_text(point, text, fontsize=current_font_size, 
                               color=text_color, fontname=font_name)
                print(f"✅ Text inserted with alternative method - font: {font_name}, size: {current_font_size}")
                return True, font_name, current_font_size
            except Exception as e:
                continue
    
    print("❌ All text insertion methods failed")
    return False, None, None

def create_annotated_pdf(original_pdf_path, parsed_info_list, all_crm_rows, output_path, invoice_to_page_mapping=None, font_size=6, font_type='auto'):
    """Create an annotated PDF with CRM results overlaid on the original pages with auto-resizing and multi-area support"""
    try:
        print(f"📖 Opening original PDF: {original_pdf_path}")
        if not os.path.exists(original_pdf_path):
            print(f"❌ Original PDF file does not exist: {original_pdf_path}")
            return False
            
        doc = fitz.open(original_pdf_path)
        print(f"✅ Successfully opened PDF with {len(doc)} pages")
        
        # Group CRM results by their page index
        crm_by_page = {}
        for row in all_crm_rows:
            page_idx = row.get('_page_index', 0)
            if page_idx not in crm_by_page:
                crm_by_page[page_idx] = []
            crm_by_page[page_idx].append(row)
        
        # Process each page
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            
            # Check if we have CRM results for this page
            if page_num in crm_by_page:
                crm_results = crm_by_page[page_num]
                print(f"📄 Processing page {page_num + 1} with {len(crm_results)} CRM records")
                
                # Get extracted page info
                parsed_info = parsed_info_list[page_num] if page_num < len(parsed_info_list) else {}
                
                # If we have multiple records, try to fit them in separate areas
                if len(crm_results) > 1:
                    print(f"🔄 Multiple records detected, attempting to separate into different areas...")
                    
                    # Create individual texts for each record
                    record_texts = []
                    for i, result in enumerate(crm_results):
                        header = f"CRM Results (Page {page_num + 1}) - Part {i + 1}:\n" + "=" * 30 + "\n"
                        record_text = header + format_single_record_text(result, i + 1)
                        record_texts.append(record_text)
                    
                    # Add extracted info to the last record
                    if parsed_info:
                        extracted_text = format_extracted_info_text(parsed_info)
                        if extracted_text:
                            record_texts[-1] += "\n" + extracted_text
                    
                    # Try to place each record in a different area
                    available_spaces = find_available_spaces_on_page(doc, page_num)
                    placed_successfully = True
                    
                    for i, record_text in enumerate(record_texts):
                        if i < len(available_spaces):
                            # Calculate required dimensions for this text
                            req_width, req_height = calculate_text_dimensions(record_text, font_size, font_type)
                            
                            # Use the available space
                            text_rect = available_spaces[i]
                            
                            # Check if the space is large enough, if not find a better one
                            if text_rect.width < req_width or text_rect.height < req_height:
                                better_spaces = find_available_spaces_on_page(doc, page_num, req_width, req_height)
                                if better_spaces:
                                    text_rect = better_spaces[0]
                            
                            print(f"📝 Placing record {i + 1} in area: {text_rect}")
                            success, font_used, size_used = insert_text_with_auto_resize(
                                page, record_text, text_rect, font_size, font_type
                            )
                            
                            if not success:
                                print(f"⚠️ Failed to place record {i + 1} separately")
                                placed_successfully = False
                                break
                        else:
                            print(f"⚠️ No available space for record {i + 1}")
                            placed_successfully = False
                            break
                    
                    if placed_successfully:
                        print(f"✅ Successfully placed all {len(crm_results)} records in separate areas")
                        continue
                    else:
                        print(f"⚠️ Could not place all records separately, falling back to combined approach")
                
                # Fallback: Combine all records in a single area (original approach but with auto-resize)
                print(f"📝 Combining all records into single area with auto-resize...")
                
                # Format combined text
                result_text = f"CRM Results (Page {page_num + 1}):\n" + "=" * 25 + "\n"
                
                for i, result in enumerate(crm_results):
                    result_text += format_single_record_text(result, i + 1) + "\n"
                
                # Add extracted page info
                if parsed_info:
                    extracted_text = format_extracted_info_text(parsed_info)
                    if extracted_text:
                        result_text += extracted_text
                
                # Calculate required dimensions
                req_width, req_height = calculate_text_dimensions(result_text, font_size, font_type)
                
                # Find best available space
                text_rect = find_blank_space_on_page(doc, page_num, req_width, req_height)
                
                print(f"📝 Inserting combined text on page {page_num + 1}:")
                print(f"   Text length: {len(result_text)} characters")
                print(f"   Required dimensions: {req_width:.1f} x {req_height:.1f}")
                print(f"   Available space: {text_rect.width:.1f} x {text_rect.height:.1f}")
                
                # Insert text with auto-resize
                success, font_used, size_used = insert_text_with_auto_resize(
                    page, result_text, text_rect, font_size, font_type
                )
                
                if success:
                    print(f"✅ Added CRM results to page {page_num + 1} ({len(crm_results)} records) with font: {font_used}, size: {size_used}")
                else:
                    print(f"❌ Failed to add CRM results to page {page_num + 1}")
                
            else:
                print(f"ℹ️ No CRM results for page {page_num + 1}")
        
        # Save the annotated PDF with error handling
        try:
            print(f"💾 Saving annotated PDF to: {output_path}")
            doc.save(output_path, garbage=4, deflate=True)
            doc.close()
            
            # Verify the saved file
            if os.path.exists(output_path):
                file_size = os.path.getsize(output_path)
                print(f"✅ Annotated PDF saved successfully:")
                print(f"   Path: {output_path}")
                print(f"   Size: {file_size} bytes")
                
                # Quick PDF validation
                try:
                    with open(output_path, 'rb') as f:
                        header = f.read(4)
                        if header == b'%PDF':
                            print("✅ PDF file validation passed")
                            return True
                        else:
                            print(f"❌ Invalid PDF header: {header}")
                            return False
                except Exception as e:
                    print(f"❌ Error validating PDF: {e}")
                    return False
            else:
                print(f"❌ PDF file was not created at: {output_path}")
                return False
                
        except Exception as save_error:
            print(f"❌ Error saving PDF: {save_error}")
            if doc:
                doc.close()
            return False
        
    except Exception as e:
        print(f"❌ Error creating annotated PDF: {e}")
        return False

def parse_bank_info(text):
    """Parse key information using xAI Grok API - handles both single page and multi-page text"""
    load_dotenv()
    api_key = os.getenv('XAI_API_KEY')
    if not api_key:
        print("❌ XAI_API_KEY not found in .env")
        return []

    print(f"✅ Grok API key loaded successfully")
    client = OpenAI(
        api_key=api_key,
        base_url="https://api.x.ai/v1",
    )
    
    # Check if this is multi-page text (contains ------)
    if "------" in text:
        # Count pages to ensure proper mapping
        pages = text.split("------")
        page_count = len(pages)
        
        prompt = f"""Extract information from this multi-page bank payment advice text. Each page is separated by "------".
                
                CRITICAL REQUIREMENTS:
                1. You MUST return exactly {page_count} records - one for each page
                2. Process pages in order (PAGE 1, PAGE 2, etc.)
                3. If a page has multiple invoice numbers, put ALL of them in the invoice field separated by commas
                4. Even if a page appears empty or has no payment info, still create a record for it
                
                For each page, extract all available information:
                - date: The payment date (if available, otherwise null)
                - amount: The payment amount with currency (if available, otherwise null)
                - payee: The recipient name (if available, otherwise null)
                - payer: The sender name (if available, otherwise null)
                - reference: Any reference number (if available, otherwise null)
                - invoice: ALL invoice numbers found on this page (e.g., "25-AVS-RES-00109-RN" or "25-AVS-RET-00161-RN,25-AVS-RET-00162-RN" if multiple), otherwise null
                - page_number: The page number (1, 2, 3, etc.)

                Return exactly {page_count} objects as a JSON array. Each page gets exactly one record.

                Text:
                {text}

                Respond only with valid JSON array containing exactly {page_count} objects.
                """
    else:
        prompt = f"""Extract the following information from this bank payment advice text as a JSON object:
                - date: The payment date (if available, otherwise null)
                - amount: The payment amount with currency (if available, otherwise null)
                - payee: The recipient name (if available, otherwise null)
                - payer: The sender name (if available, otherwise null)
                - reference: Any reference number (if available, otherwise null)
                - invoice: ALL invoice numbers found (e.g., "25-AVS-RES-00109-RN" or "25-AVS-RET-00161-RN,25-AVS-RET-00162-RN" if multiple), otherwise null

                Text:
                {text}

                Respond only with valid JSON.
                """
    
    try:
        print(f"🚀 Making API call to Grok (xAI) API...")
        print(f"   🤖 Model: grok-code-fast-1")
        print(f"   📝 Prompt length: {len(prompt)} characters")
        print(f"   🌐 Endpoint: https://api.x.ai/v1/chat/completions")
        print(f"   ⏰ Starting request at: {time.strftime('%Y-%m-%d %H:%M:%S')}")

        response = client.chat.completions.create(
            model="grok-code-fast-1",
            messages=[
                {"role": "system", "content": "You are a helpful assistant that extracts structured data from text."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1,
        )

        print(f"✅ Grok API request completed successfully!")
        print(f"   📊 Response received at: {time.strftime('%Y-%m-%d %H:%M:%S')}")
        extracted_json = response.choices[0].message.content.strip()
        
        # Print Grok results to terminal
        print("\n" + "="*60)
        print("🤖 GROK LLM RESPONSE:")
        print("="*60)
        print(extracted_json)
        print("="*60 + "\n")
        
        # Remove any markdown code block formatting if present
        if extracted_json.startswith("```json"):
            extracted_json = extracted_json[7:]
        if extracted_json.endswith("```"):
            extracted_json = extracted_json[:-3]
        extracted_json = extracted_json.strip()
        
        parsed_data = json.loads(extracted_json)
        
        # Print parsed data for verification
        print("📋 PARSED DATA:")
        print(json.dumps(parsed_data, indent=2, ensure_ascii=False))
        print("="*60 + "\n")
        
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

def split_invoice_numbers(invoice_string):
    """Split comma-separated invoice numbers into individual invoices"""
    if not invoice_string:
        return [None]
    
    # Split by comma and clean up each invoice number
    invoices = []
    for inv in invoice_string.split(','):
        inv = inv.strip()
        if inv:
            invoices.append(inv)
    
    return invoices if invoices else [None]

def ensure_all_pages_represented(all_crm_rows, parsed_info_list, invoice_to_page_mapping):
    """Ensure all pages are represented in CRM results, even if no results found"""
    if not parsed_info_list:
        return all_crm_rows
    
    # Define all expected columns in consistent order
    expected_columns = [
        'Page', 'Source', 'Opened', 'Parent Company Name', 'Premises Name', 
        'Premises Address', 'Salutation', 'Contact Name', 'Title', 'Business E-Mail', 
        'Phone Number', 'Licence Category', 'Expected Closed', 'Stage', 
        'Payment Received Date', 'Invoice No', 'Net Licence Fee', 'Territory', 'Assigned To'
    ]
    
    # Ensure all existing rows have all columns
    for row in all_crm_rows:
        for col in expected_columns:
            if col not in row:
                row[col] = ''
    
    # Create a set of pages that have results
    pages_with_results = set()
    for row in all_crm_rows:
        page_index = row.get('_page_index')
        if page_index is not None:
            pages_with_results.add(page_index)
    
    # Add empty rows for pages without results
    for page_idx in range(len(parsed_info_list)):
        if page_idx not in pages_with_results:
            # Find the corresponding invoice mapping for this page
            invoice_for_page = None
            for mapping in invoice_to_page_mapping:
                if mapping['page_index'] == page_idx:
                    invoice_for_page = mapping['invoice']
                    break
            
            # Create an empty editable row for this page with all expected columns
            empty_row = {
                '_page_index': page_idx,
                '_invoice_index': page_idx,
                '_page_info': parsed_info_list[page_idx] if page_idx < len(parsed_info_list) else {},
                '_original_invoice_string': invoice_for_page,
                '_source_invoice': invoice_for_page or f"Page {page_idx + 1}",
            }
            
            # Add all expected columns
            for col in expected_columns:
                if col == 'Page':
                    empty_row[col] = f"Page {page_idx + 1}"
                elif col == 'Source':
                    empty_row[col] = invoice_for_page or "No Invoice"
                else:
                    empty_row[col] = ''
                    
            all_crm_rows.append(empty_row)
    
    # Sort by page index to maintain order
    all_crm_rows.sort(key=lambda x: x.get('_page_index', 0))
    
    return all_crm_rows

def extract_invoice(file_path: str) -> tuple[list[str], list[dict], list[dict]]:
    """Extract invoice numbers and parsed info from image or PDF. 
    Returns (all_invoice_numbers, parsed_info_list, invoice_to_page_mapping)"""
    if is_pdf_file(file_path):
        print(f"📄 Processing PDF file: {file_path}")
        extracted_text = extract_text_from_pdf_pages(file_path)
        if not extracted_text:
            print(f"❌ Failed to extract text from PDF: {file_path}")
            return [], [], []

        print(f"🚀 Sending PDF extracted text to Grok API for parsing...")
        print(f"   📊 Text length: {len(extracted_text)} characters")
        print(f"   📄 Processing PDF: {file_path}")
        parsed_info_list = parse_bank_info(extracted_text)
    else:
        print(f"🖼️ Processing image file: {file_path}")
        image = cv2.imread(file_path)
        if image is None:
            print(f"❌ Failed to load image: {file_path}")
            return [], [], []

        preprocessed = preprocess_image(image)
        extracted_text = extract_text_with_api(preprocessed)

        print(f"🚀 Sending image extracted text to Grok API for parsing...")
        print(f"   📊 Text length: {len(extracted_text)} characters")
        print(f"   🖼️ Processing image: {file_path}")
        parsed_info_list = parse_bank_info(extracted_text)
    
    # Process each page's invoices and create mapping
    all_invoice_numbers = []
    invoice_to_page_mapping = []
    
    for page_idx, info in enumerate(parsed_info_list):
        page_invoice_string = info.get('invoice')
        page_invoices = split_invoice_numbers(page_invoice_string)
        
        # Add each invoice with its page mapping
        for invoice in page_invoices:
            all_invoice_numbers.append(invoice)
            invoice_to_page_mapping.append({
                'invoice': invoice,
                'page_index': page_idx,
                'page_info': info,
                'original_invoice_string': page_invoice_string
            })
    
    print(f"📋 Extracted {len(all_invoice_numbers)} total invoices from {len(parsed_info_list)} pages")
    for i, mapping in enumerate(invoice_to_page_mapping):
        print(f"  Invoice {i+1}: {mapping['invoice']} (Page {mapping['page_index']+1})")
    
    return all_invoice_numbers, parsed_info_list, invoice_to_page_mapping

# ===============================
# CRM AUTOMATION FUNCTIONS
# ===============================

class CRMAutoLogin:
    def __init__(self, headless: bool = False, invoice_number: str = None, invoice_numbers: list = None, 
                 invoice_to_page_mapping: list = None, return_json: bool = False, no_interactive: bool = False, web_output: bool = False):
        self.url = "http://192.168.1.152/crm/eware.dll/go"
        self.username = "ivan.chiu"
        self.password = "25207090"
        self.headless = headless
        self.invoice_number = invoice_number
        self.invoice_numbers = invoice_numbers or ([invoice_number] if invoice_number else [])
        self.invoice_to_page_mapping = invoice_to_page_mapping or []
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
        
        # Add source invoice to each record (only invoice number, no page info)
        for record in records:
            record['_source_invoice'] = invoice_number
            record['Source'] = invoice_number  # Only the invoice number for display
            
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
        source_value = f"fee_search_{fee_amount}" if fee_amount else "fee_search_general"
        for record in records:
            record['_source_invoice'] = source_value
            record['Source'] = source_value  # Also add as 'Source' for display
            
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
                    
                    # Add mapping information to each record
                    for record in records:
                        record['_invoice_index'] = i
                        # Find corresponding mapping info
                        if i < len(self.invoice_to_page_mapping):
                            mapping = self.invoice_to_page_mapping[i]
                            record['_page_index'] = mapping['page_index']
                            record['_page_info'] = mapping['page_info']
                            record['_original_invoice_string'] = mapping['original_invoice_string']
                            # Add separate Page column
                            record['Page'] = f"Page {mapping['page_index'] + 1}"
                        else:
                            # Fallback for backward compatibility
                            record['_page_index'] = i
                            record['_page_info'] = {}
                            record['_original_invoice_string'] = invoice
                            record['Page'] = f"Page {i + 1}"
                    
                    all_records.extend(records)
                    
                    # Small delay between searches
                    if i < len(self.invoice_numbers) - 1:
                        time.sleep(2)
            elif self.invoice_number:
                # Backward compatibility for single invoice
                records = self.search_invoice(self.invoice_number)
                for record in records:
                    record['_invoice_index'] = 0
                    record['_page_index'] = 0
                    record['_page_info'] = {}
                    record['_original_invoice_string'] = self.invoice_number
                    record['Page'] = "Page 1"
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
            
            # Extract invoices (can be multiple from PDF with proper page mapping)
            all_invoice_numbers, parsed_info_list, invoice_to_page_mapping = extract_invoice(file_path)
            if not parsed_info_list:
                flash('No information extracted from the file.', 'danger')
                return render_template('result.html', records=parsed_info_list, all_crm_rows=[], logs="")
            else:
                # Filter out None invoice numbers for CRM processing
                valid_invoices = [inv for inv in all_invoice_numbers if inv is not None]
                flash(f'Extracted {len(parsed_info_list)} page(s) with {len(valid_invoices)} invoice(s). Launching CRM automation...', 'success')
                
                # Initialize CRM automation directly (no subprocess)
                try:
                    print("🚀 Starting integrated CRM automation...")
                    crm = CRMAutoLogin(
                        headless=True,
                        invoice_numbers=all_invoice_numbers,  # Include all invoice numbers (including None)
                        invoice_to_page_mapping=invoice_to_page_mapping,  # Pass the mapping
                        return_json=True,
                        no_interactive=True,
                        web_output=True
                    )
                    
                    success, all_crm_rows = crm.run()
                    
                    # Ensure all pages are represented, even if no results found
                    all_crm_rows = ensure_all_pages_represented(all_crm_rows, parsed_info_list, invoice_to_page_mapping)
                    
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

                # Generate annotated PDF if original file was a PDF and we have results
                annotated_pdf_filename = None
                if is_pdf_file(file_path) and (all_crm_rows or parsed_info_list):
                    try:
                        annotated_pdf_filename = f"annotated_{uuid.uuid4().hex}_{filename}"
                        annotated_pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], annotated_pdf_filename)
                        
                        success = create_annotated_pdf(file_path, parsed_info_list, all_crm_rows, annotated_pdf_path, invoice_to_page_mapping, 6, 'auto')
                        if success:
                            flash('Annotated PDF generated successfully!', 'success')
                        else:
                            annotated_pdf_filename = None
                            flash('Failed to generate annotated PDF.', 'warning')
                    except Exception as e:
                        print(f"❌ Error generating annotated PDF: {e}")
                        annotated_pdf_filename = None
                        flash(f'Error generating annotated PDF: {str(e)}', 'warning')

                flash(f'Processing completed! Total CRM records found: {len(all_crm_rows)}', 'info')
                return render_template('result.html', 
                                     records=parsed_info_list, 
                                     all_crm_rows=all_crm_rows, 
                                     logs=combined_logs,
                                     annotated_pdf=annotated_pdf_filename)
        else:
            flash('File type not allowed. Please upload an image.', 'danger')
            return redirect(request.url)
    return render_template('index.html')


@app.route('/serve_pdf/<filename>')
def serve_pdf(filename):
    """Serve generated PDF files for download"""
    try:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        if os.path.exists(file_path) and ('annotated_' in filename or 'updated_annotated_' in filename):
            # Extract the original filename for download
            if filename.startswith('updated_annotated_'):
                # For updated PDFs: updated_annotated_{uuid}_{original_filename}
                parts = filename.split('_', 3)
                if len(parts) >= 4:
                    download_name = f"updated_results_{parts[3]}"
                else:
                    download_name = f"updated_results_{filename}"
            elif filename.startswith('annotated_'):
                # For original annotated PDFs: annotated_{uuid}_{original_filename}
                parts = filename.split('_', 2)
                if len(parts) >= 3:
                    download_name = f"results_{parts[2]}"
                else:
                    download_name = f"results_{filename}"
            else:
                download_name = filename
            
            return send_file(file_path, as_attachment=True, download_name=download_name)
        else:
            return jsonify({'error': 'File not found'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/generate_annotated_pdf', methods=['POST'])
def generate_annotated_pdf():
    """Generate annotated PDF with user's edited data"""
    try:
        data = request.get_json()
        table_data = data.get('tableData', [])
        font_size = data.get('fontSize', 6)  # Default to 6px if not provided
        font_type = data.get('fontType', 'auto')  # Default to auto if not provided
        
        if not table_data:
            return jsonify({'success': False, 'error': 'No table data provided'}), 400
        
        # We need to find the original PDF file to annotate
        # For now, we'll use the most recent PDF file in uploads
        pdf_files = [f for f in os.listdir(app.config['UPLOAD_FOLDER']) 
                    if f.lower().endswith('.pdf') and not f.startswith('annotated_')]
        
        if not pdf_files:
            return jsonify({'success': False, 'error': 'No original PDF file found'}), 400
        
        # Use the most recent PDF file
        pdf_files.sort(key=lambda f: os.path.getmtime(os.path.join(app.config['UPLOAD_FOLDER'], f)), reverse=True)
        original_pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_files[0])
        
        # Create filename for the new annotated PDF
        base_filename = os.path.splitext(pdf_files[0])[0]  # Remove extension
        annotated_filename = f"updated_annotated_{uuid.uuid4().hex}_{base_filename}.pdf"
        annotated_pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], annotated_filename)
        
        print(f"📝 Creating updated annotated PDF:")
        print(f"   Original PDF: {pdf_files[0]}")
        print(f"   New filename: {annotated_filename}")
        print(f"   Output path: {annotated_pdf_path}")
        
        # Group table data by page index
        crm_by_page = {}
        parsed_info_list = []
        
        for row in table_data:
            page_idx = row.get('_page_index', 0)
            if page_idx not in crm_by_page:
                crm_by_page[page_idx] = []
            crm_by_page[page_idx].append(row)
            
            # Create parsed info entry if not exists
            while len(parsed_info_list) <= page_idx:
                parsed_info_list.append({})
        
        # Convert grouped data back to list format for the PDF function
        all_crm_rows = []
        for page_idx in sorted(crm_by_page.keys()):
            all_crm_rows.extend(crm_by_page[page_idx])
        
        # Create invoice to page mapping (simplified)
        invoice_to_page_mapping = []
        for page_idx in range(len(parsed_info_list)):
            invoice_to_page_mapping.append({
                'page_index': page_idx,
                'invoice': f"Page {page_idx + 1}",
                'page_info': parsed_info_list[page_idx] if page_idx < len(parsed_info_list) else {},
                'original_invoice_string': f"Page {page_idx + 1}"
            })
        
        # Generate the annotated PDF
        print(f"🚀 Calling create_annotated_pdf function...")
        print(f"   📝 Using font size: {font_size}px")
        print(f"   🔤 Using font type: {font_type}")
        success = create_annotated_pdf(
            original_pdf_path, 
            parsed_info_list, 
            all_crm_rows, 
            annotated_pdf_path, 
            invoice_to_page_mapping,
            font_size,
            font_type
        )
        
        # Verify the file was created and is valid
        if success and os.path.exists(annotated_pdf_path):
            file_size = os.path.getsize(annotated_pdf_path)
            print(f"✅ PDF created successfully:")
            print(f"   File path: {annotated_pdf_path}")
            print(f"   File size: {file_size} bytes")
            
            # Basic PDF validation - check if it starts with PDF header
            try:
                with open(annotated_pdf_path, 'rb') as f:
                    header = f.read(4)
                    if header == b'%PDF':
                        print("✅ PDF header validation passed")
                    else:
                        print(f"⚠️ PDF header validation failed: {header}")
            except Exception as e:
                print(f"⚠️ PDF validation error: {e}")
            
            return jsonify({
                'success': True, 
                'filename': annotated_filename,
                'message': 'Annotated PDF generated successfully with your edits!'
            })
        else:
            error_msg = 'Failed to generate annotated PDF'
            if not success:
                error_msg += ' - PDF creation function returned False'
            if not os.path.exists(annotated_pdf_path):
                error_msg += f' - Output file does not exist: {annotated_pdf_path}'
            
            print(f"❌ {error_msg}")
            return jsonify({'success': False, 'error': error_msg}), 500
            
    except Exception as e:
        print(f"❌ Error in generate_annotated_pdf route: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5002, debug=True) 