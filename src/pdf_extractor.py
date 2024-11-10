import os
import pandas as pd
import pdfplumber
from pathlib import Path
import logging
import re
from datetime import datetime

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('pdf_extractor.log'),
        logging.StreamHandler()
    ]
)

def extract_contractor_name(text):
    """Extract contractor name using multiple patterns"""
    contractor_patterns = [
        r"(Xpert'?s LLC)",
        r"(Ceres Environmental Services,\s*Inc\.?)",
        r"(Wright Tree Service of Puerto Rico,\s*LLC)",
    ]
    
    for pattern in contractor_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return None

def extract_pdf_data(pdf_path):
    """Extract specific fields from the PDF"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]
            text = page.extract_text()
            
            # Print full text for debugging
            logging.info("Extracted text:")
            logging.info(text)
            
            # Dictionary to store extracted values
            data = {}
            
            # Updated patterns to match the actual PDF format
            patterns = {
                'Task Order #': r'Task Order Number:?\s*([\w-]+)',
                'Total Amount': r'Task Order Total Amount:?\s*\$?([\d,]+\.\d{2})',
                'Feeder ID': r'(?:Feeder|Feeder ID:?)\s*(\d{4}-\d{2})',
                'Feeder Total Miles': r'Length:?\s*([\d.]+)\s*(?:overhead miles|Overhead miles)',
                '# of Work Orders': r'Work Orders:?\s*(\d+)(?:\s*WO locations)?',
                'Task Order Start Date': r'Start Date:?\s*(\d{2}/\d{2}/(?:\d{2}|\d{4}))',
                'Task Order End Date': r'End Date:?\s*(\d{2}/\d{2}/(?:\d{2}|\d{4}))'
            }
            
            # Extract contractor name
            contractor_name = extract_contractor_name(text)
            if contractor_name:
                data['Contractor'] = contractor_name
                logging.info(f"Extracted Contractor: {contractor_name}")
            else:
                logging.warning("Could not find contractor name in PDF")
            
            # Extract other fields
            for field, pattern in patterns.items():
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    value = match.group(1).strip()
                    # Convert numeric values
                    if field == 'Total Amount':
                        value = float(value.replace(',', ''))
                    elif field == 'Feeder Total Miles':
                        value = float(value)
                    elif field == '# of Work Orders':
                        value = int(value)
                    elif 'Date' in field:
                        try:
                            if len(value.split('/')[-1]) == 2:
                                date_obj = datetime.strptime(value, '%m/%d/%y')
                            else:
                                date_obj = datetime.strptime(value, '%m/%d/%Y')
                            value = date_obj.strftime('%m/%d/%Y')
                        except ValueError as e:
                            logging.error(f"Error parsing date {value}: {str(e)}")
                    
                    data[field] = value
                    logging.info(f"Extracted {field}: {value}")
                else:
                    logging.warning(f"Could not find {field} in PDF")
                    
            return data
            
    except Exception as e:
        logging.error(f"Error processing PDF: {str(e)}")
        return None

def update_excel(data, excel_path):
    """Update or create Excel file with extracted data"""
    try:
        # Create new DataFrame if file doesn't exist
        if not os.path.exists(excel_path):
            df = pd.DataFrame(columns=[
                'Task Order #', 'Contractor', 'Total Amount', 'Feeder ID', 
                'Feeder Total Miles', '# of Work Orders',
                'Task Order Start Date', 'Task Order End Date'
            ])
        else:
            df = pd.read_excel(excel_path)
        
        # Check if task order already exists
        task_order = data.get('Task Order #')
        if task_order and not df['Task Order #'].astype(str).eq(task_order).any():
            # Add new row
            new_row = pd.DataFrame([data])
            df = pd.concat([df, new_row], ignore_index=True)
            
            # Sort by start date
            df['Task Order Start Date'] = pd.to_datetime(df['Task Order Start Date'])
            df = df.sort_values('Task Order Start Date', ascending=False)
            
            # Save to Excel
            df.to_excel(excel_path, index=False)
            logging.info(f"Successfully updated Excel file with task order: {task_order}")
        else:
            logging.info(f"Task order {task_order} already exists in Excel file")
        
    except Exception as e:
        logging.error(f"Error updating Excel file: {str(e)}")

def main():
    # Set up paths
    input_folder = Path("data/input")
    output_folder = Path("data/output")
    output_folder.mkdir(parents=True, exist_ok=True)
    excel_path = output_folder / "task_orders.xlsx"
    
    # Process all PDFs in input folder
    for pdf_file in input_folder.glob("*.pdf"):
        logging.info(f"Processing {pdf_file}")
        data = extract_pdf_data(pdf_file)
        
        if data:
            update_excel(data, excel_path)
        else:
            logging.error(f"Failed to extract data from {pdf_file}")

if __name__ == "__main__":
    main()