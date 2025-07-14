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
    """Parse key information from extracted text"""
    info = {}
    
    # Extract date (e.g., 18 Jun 2021)
    date_match = re.search(r'Date:\s*(\d{1,2} [A-Za-z]{3} \d{4})', text)
    info['date'] = date_match.group(1) if date_match else 'Not found'
    
    # Extract amount (e.g., HKD 13,000.00)
    amount_match = re.search(r'(HKD|USD)? ?[\d,]+.\d{2}', text)
    info['amount'] = amount_match.group(0) if amount_match else 'Not found'
    
    # Extract payee (e.g., Osim)
    payee_match = re.search(r'Pay (?:the order of|to) (.*)', text, re.IGNORECASE)
    info['payee'] = payee_match.group(1).strip() if payee_match else 'Not found'
    
    # Extract reference or other fields as needed
    ref_match = re.search(r'Ref\. No\. (\S+)', text)
    info['reference'] = ref_match.group(1) if ref_match else 'Not found'
    
    return info

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
    
    # Parse key info
    parsed_info = parse_bank_info(extracted_text)
    print("\n🔑 Parsed Information:")
    for key, value in parsed_info.items():
        print(f"{key.capitalize()}: {value}")
    
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