#!/usr/bin/env python3
"""
Bank Payment Advice Scanner
Uses OpenCV and Tesseract OCR to extract key information from bank slips
"""

import argparse
import re
import cv2
import pytesseract
from matplotlib import pyplot as plt
import os
from dotenv import load_dotenv
from openai import OpenAI

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

def extract_text(image):
    """Extract text using Tesseract with custom config"""
    custom_config = r'--oem 3 --psm 6'  # Assume structured text
    return pytesseract.image_to_string(image, config=custom_config, lang='eng')

def parse_bank_info(text):
    """Parse key information using xAI Grok API"""
    load_dotenv()
    api_key = os.getenv('XAI_API_KEY')
    if not api_key:
        print("❌ XAI_API_KEY not found in .env")
        return {}
    
    client = OpenAI(
        api_key=api_key,
        base_url="https://api.x.ai/v1",
    )
    
    prompt = f"""Extract the following information from this bank payment advice text as a JSON object:
    - date: The payment date
    - amount: The payment amount with currency
    - payee: The recipient name
    - payer: The sender name
    - reference: Any reference number
    - invoice: The invoice number (e.g., 25-AVS-RES-00109-RN)
    
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
        import json
        return json.loads(extracted_json)
    except Exception as e:
        print(f"❌ Error calling xAI API: {e}")
        return {}

def main():
    args = parse_arguments()
    
    # Load image
    image = cv2.imread(args.image_path)
    if image is None:
        print(f"❌ Failed to load image: {args.image_path}")
        return
    
    # Preprocess
    preprocessed = preprocess_image(image)
    
    # Extract text
    extracted_text = extract_text(preprocessed)
    print("\n📝 Extracted Text:")
    print(extracted_text)
    
    # Parse key info using LLM
    parsed_info = parse_bank_info(extracted_text)
    print("\n🔑 Parsed Information (via Grok LLM):")
    if parsed_info:
        for key, value in parsed_info.items():
            print(f"{key.capitalize()}: {value}")
        if 'invoice' not in parsed_info:
            print("Invoice: Not found")
    else:
        print("No information extracted")
    
    if args.display:
        # For visualization (optional)
        image_rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
        data = pytesseract.image_to_data(preprocessed, output_type=pytesseract.Output.DICT, config='--oem 3 --psm 6')
        n_boxes = len(data['level'])
        for i in range(n_boxes):
            if int(data['conf'][i]) > 60:
                (x, y, w, h) = (data['left'][i], data['top'][i], data['width'][i], data['height'][i])
                cv2.rectangle(image_rgb, (x, y), (x + w, y + h), (255, 0, 0), 2)
        
        plt.imshow(image_rgb)
        plt.title("Image with Bounding Boxes")
        plt.show()

if __name__ == "__main__":
    main()