import argparse
import cv2
import requests
import base64
from matplotlib import pyplot as plt
import os
from dotenv import load_dotenv
from openai import OpenAI
import fitz  # PyMuPDF
from pdf2image import convert_from_path
from PIL import Image
import numpy as np
import json
from io import BytesIO

def parse_arguments():
    parser = argparse.ArgumentParser(description="Scan and extract info from bank payment advice images")
    parser.add_argument("image_path", nargs="?", default="bank.png", help="Path to the bank image file (default: bank.png)")
    parser.add_argument("--display", action="store_true", help="Display images with bounding boxes")
    return parser.parse_args()

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

# Backward compatibility function
def extract_single_invoice(file_path: str) -> tuple[str | None, dict]:
    """Extract first invoice number and parsed info for backward compatibility"""
    invoice_numbers, parsed_info_list = extract_invoice(file_path)
    if invoice_numbers and parsed_info_list:
        return invoice_numbers[0], parsed_info_list[0]
    return None, {}

def main():
    args = parse_arguments()
    
    # Process file (image or PDF)
    invoice_numbers, parsed_info_list = extract_invoice(args.image_path)
    
    if is_pdf_file(args.image_path):
        print(f"\n📄 Processing PDF: {args.image_path}")
        
        if parsed_info_list:
            print(f"\n🔑 Extracted {len(parsed_info_list)} records from PDF:")
            for i, parsed_info in enumerate(parsed_info_list, 1):
                print(f"\n--- Record {i} ---")
                for key, value in parsed_info.items():
                    if value:  # Only print non-empty values
                        print(f"{key.capitalize()}: {value}")
        else:
            print("No records extracted from PDF")
    else:
        # Single image processing
        image = cv2.imread(args.image_path)
        if image is None:
            print(f"❌ Failed to load image: {args.image_path}")
            return
        
        # Preprocess and extract text for display
        preprocessed = preprocess_image(image)
        extracted_text = extract_text_with_api(preprocessed)
        print("\n📝 Extracted Text:")
        print(extracted_text)
        
        # Show parsed info
        print("\n🔑 Parsed Information (via Grok LLM):")
        if parsed_info_list:
            for key, value in parsed_info_list[0].items():
                print(f"{key.capitalize()}: {value}")
        else:
            print("No information extracted")
        
        if args.display:
            # Display the preprocessed image
            image_rgb = cv2.cvtColor(preprocessed, cv2.COLOR_GRAY2RGB)
            plt.imshow(image_rgb, cmap='gray')
            plt.title("Preprocessed Image for OCR")
            plt.show()

if __name__ == "__main__":
    main()