#!/usr/bin/env python3

import json
import sys
import re
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Optional, Tuple
import PyPDF2
import logging


def get_downloads_folder() -> Path:
    """Return the user's Downloads folder (works on Windows, macOS, and Linux)."""
    return Path.home() / "Downloads"


def save_result_to_downloads(result: Dict, source_name: Optional[str] = None) -> str:
    """
    Save parsed JSON result to a text file in the user's Downloads folder.
    Returns the path of the saved file.
    """
    downloads = get_downloads_folder()
    downloads.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if source_name:
        safe_name = re.sub(r'[^\w\-.]', '_', Path(source_name).stem)[:50]
        filename = f"sobey_parser_{safe_name}_{timestamp}.txt"
    else:
        filename = f"sobey_parser_result_{timestamp}.txt"
    out_path = downloads / filename
    result_with_path = {**result, "saved_to": str(out_path)}
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(json.dumps(result_with_path, indent=2, ensure_ascii=False))
    return str(out_path)

# Setup logging (INFO only for internal use; console is silenced in main())
logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)


class LineItemData:
    """Helper class to store line item data"""
    def __init__(self):
        self.vendor_no = ""
        self.vendor_name = ""
        self.cubes = ""
        self.weight = ""
        self.pieces = ""
        self.po = ""
        self.match_end = 0


class SobeyTemplate1PdfParser:
    """PDF Parser for Sobey Template 1 and Template 2"""
    
    # Template 1 regex: Single pickup/delivery date, Stop 2 for ship-to
    REGEX_TEMPLATE1 = r"(\d+)\s+-\s+(.*)\s*Cube\s*:\s*([0-9,]+)\s*Weight\s*:\s*([0-9,]+)\s*Pieces\s*:\s*([0-9,]+)\s*[A-Z]{3}-PO-\d{2}-\d{4}-(\d+).*\s+Ref\s*Number:.*\s*Pallet\s*Count:(.*)\s*(.*)"
    
    # Template 2 regex - FIXED to be more precise
    # Must match: vendor number, vendor name (non-greedy), then Cube/Weight/Pieces, then PO
    REGEX_TEMPLATE2 = r"(\d{6})\s+-\s+([^C]+?)\s+Cube\s*:\s*([0-9,]+)\s+Weight\s*:\s*([0-9,]+)\s+Pieces\s*:\s*([0-9,]+)\s+[A-Z]{3}-PO-\d{2}-\d{4}-(\d+)"
    
    def __init__(self):
        pass
    
    def convert_date_format(self, date_str: str) -> str:
        """
        Convert date from 'Oct 20, 2025 11:59:00 PM' to '20/10/2025'
        """
        try:
            # Remove any newlines or extra whitespace
            date_str = re.sub(r'\s+', ' ', date_str).strip()
            
            # Parse input format: "Oct 20, 2025 11:59:00 PM"
            date_obj = datetime.strptime(date_str, "%b %d, %Y %I:%M:%S %p")
            
            # Output format: "20/10/2025"
            return date_obj.strftime("%d/%m/%Y")
        except Exception as e:
            log.error(f"Error converting date format: {date_str}, {e}")
            return date_str  # Return original if conversion fails
    
    def extract_vendor_name(self, full_vendor_string: str) -> str:
        """
        Extracts the base vendor name from the full vendor string
        Example: "Smucker Foods of Canada - 24 - TRA St. Johns - Mount Pearl DC"
        Returns: "Smucker Foods of Canada"
        Example: "237772 - Agropur Industrial Div"
        Returns: "Agropur Industrial Div"
        """
        if not full_vendor_string or not full_vendor_string.strip():
            return ""
        
        # Clean up the string
        full_vendor_string = re.sub(r'\s+', ' ', full_vendor_string).strip()
        
        # Case 1: String starts with vendor number
        # Pattern: "237772 - Agropur Industrial Div - 2 - TRA Location"
        pattern_with_number = re.compile(r'^\d+\s*-\s*(.*?)\s*-\s*\d+')
        match_with_number = pattern_with_number.search(full_vendor_string)
        
        if match_with_number:
            vendor_name = match_with_number.group(1).strip()
            log.info(f"Extracted vendor name (with number prefix): '{vendor_name}' from: '{full_vendor_string}'")
            return vendor_name
        
        # Case 2: String does NOT start with vendor number
        # Pattern: "Smucker Foods of Canada - 24 - TRA St. Johns"
        pattern_no_number = re.compile(r'^(.*?)\s*-\s*\d+\s*-')
        match_no_number = pattern_no_number.search(full_vendor_string)
        
        if match_no_number:
            vendor_name = match_no_number.group(1).strip()
            log.info(f"Extracted vendor name (no number prefix): '{vendor_name}' from: '{full_vendor_string}'")
            return vendor_name
        
        # Case 3: Simple format with number at start, no stop number after
        # Pattern: "237772 - Agropur Industrial Div"
        simple_pattern = re.compile(r'^\d+\s*-\s*(.+?)$')
        simple_match = simple_pattern.search(full_vendor_string)
        
        if simple_match:
            vendor_name = simple_match.group(1).strip()
            log.info(f"Extracted vendor name (simple): '{vendor_name}' from: '{full_vendor_string}'")
            return vendor_name
        
        # If no match, return the full string cleaned up
        log.warning(f"Could not extract vendor name from: '{full_vendor_string}', returning full string")
        return full_vendor_string
    
    def extract_shipment_type(self, text: str, po_number: str) -> str:
        """
        Extracts the shipment type from the PO number
        Example: "AMS-PO-10-2025-4610227518 (GROC)" returns "GROC"
        """
        # More flexible pattern - handle extra spaces
        pattern = re.compile(rf'[A-Z]{{3}}-PO-\d{{2}}-\d{{4}}-{po_number}\s*\(([^)]+)\)', re.IGNORECASE)
        match = pattern.search(text)
        if match:
            return match.group(1).strip()
        return ""
    
    def detect_template(self, text: str) -> str:
        """Detect which template is being used"""
        # Look for "Pickup :" or "Pickup:" with date pattern
        template2_pattern = re.compile(r'Pickup\s*:\s*\w{3}\s+\d{1,2},\s+\d{4}\s+\d{1,2}:\d{2}:\d{2}')
        
        if template2_pattern.search(text):
            log.info("Detected Template 2 (Multiple pickup/delivery dates per line item)")
            return "Template-2"
        else:
            log.info("Detected Template 1 (Single pickup/delivery date)")
            return "Template-1"
    
    def extract_stop_destinations(self, text: str) -> List[str]:
        """Extract stop destinations from the PDF text"""
        destinations = []
        stop_pattern = re.compile(r'(?s)Stop:\s*(\d+)\s*Destination:\s*(.*?)\s*Stop Location Memo:', re.IGNORECASE)
        
        for match in stop_pattern.finditer(text):
            stop_number = int(match.group(1))
            destination = match.group(2).strip()
            
            # Add to list indexed by stop number (Stop 1 at index 0, etc.)
            while len(destinations) < stop_number:
                destinations.append("")
            destinations[stop_number - 1] = destination
        
        return destinations
    
    def extract_text_from_pdf(self, pdf_path: str) -> str:
        """Extract text from PDF file"""
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                text = ""
                for page in pdf_reader.pages:
                    text += page.extract_text()
                return text
        except Exception as e:
            log.error(f"Error extracting text from PDF: {e}")
            raise
    
    def process_template1(self, text: str, stop_destinations: List[str]) -> List[Dict]:
        """Process Template 1 PDF"""
        records = []
        pattern = re.compile(self.REGEX_TEMPLATE1)
        
        # Extract pickup and delivery dates for Template-1 (format: "Pickup On : DD/MM/YYYY")
        pickup_on_pattern = re.compile(r'Pickup\s+On\s*:\s*(\d{2}/\d{2}/\d{4})', re.IGNORECASE)
        deliver_on_pattern = re.compile(r'Deliver\s+On\s*:\s*(\d{2}/\d{2}/\d{4})', re.IGNORECASE)
        
        pickup_date_match = pickup_on_pattern.search(text)
        deliver_date_match = deliver_on_pattern.search(text)
        
        pickup_date = pickup_date_match.group(1) if pickup_date_match else ""
        del_date = deliver_date_match.group(1) if deliver_date_match else ""
        
        log.info(f"Template-1: Found pickup_date={pickup_date}, del_date={del_date}")
        
        # Extract ship_from from Stop 1 (index 0)
        ship_from = ""
        if len(stop_destinations) > 0:
            ship_from = stop_destinations[0]  # Stop 1 is at index 0
            log.info(f"Template-1: Using Stop 1 as ship_from: {ship_from}")
        
        # Extract ship_to from Stop 2 (index 1)
        ship_to = ""
        if len(stop_destinations) > 1:
            ship_to = stop_destinations[1]  # Stop 2 is at index 1
            log.info(f"Template-1: Using Stop 2 as ship_to: {ship_to}")
        
        for match in pattern.finditer(text):
            # Extract description from Pallet Count (group 7)
            description = match.group(7).strip() if match.group(7) else ''
            
            # Remove trailing pipe character if present
            if description:
                description = description.rstrip('|').strip()
            
            # If description is still empty, try to extract it directly from text near the PO
            if not description:
                po_number = match.group(6)
                # Look for Pallet Count near the PO number
                po_context_pattern = re.compile(
                    rf'{re.escape(po_number)}[\s\S]*?Pallet\s+Count:\s*([^\n|]+?)(?=\s*[|]|\s*Cube|\s*Weight|$)',
                    re.IGNORECASE | re.MULTILINE
                )
                po_context_match = po_context_pattern.search(text)
                if po_context_match:
                    description = po_context_match.group(1).strip()
                    log.info(f"Template-1: Extracted description from context: {description}")
            
            record = {
                'template': 'Template-1',
                'pickup_date': pickup_date,
                'del_date': del_date,
                'ship_from': ship_from,
                'ship_to': ship_to,
                'vendor_no': match.group(1),
                'vendor_name': self.extract_vendor_name(match.group(2)),
                'cubes': match.group(3),
                'weight': match.group(4),
                'pieces': match.group(5),
                'po': match.group(6),
                'shipment_type': self.extract_shipment_type(text, match.group(6)),
                'pallets': '',
                'description': description
            }
            
            records.append(record)
        
        return records
    
    def process_template2(self, text: str, stop_destinations: List[str]) -> List[Dict]:
        """Process Template 2 PDF"""
        records = []
        
        # FIXED: Extract line items with better regex that handles line breaks
        line_items = []
        
        # Pattern that allows vendor name to span multiple lines
        # Uses [\s\S]*? to match any character including newlines (non-greedy)
        # Stops when it encounters "Cube" (with optional whitespace before it)
        shipment_pattern = re.compile(
            r'(\d{6})\s+-\s+([\s\S]+?)(?=\s*Cube\s*:)\s*Cube\s*:\s*([0-9,]+)\s+Weight\s*:\s*([0-9,]+)\s+Pieces\s*:\s*([0-9,]+)\s+([A-Z]{3}-PO-\d{2}-\d{4}-(\d+))',
            re.MULTILINE
        )
        
        for match in shipment_pattern.finditer(text):
            item = LineItemData()
            item.vendor_no = match.group(1)
            raw_vendor_name = match.group(2)
            # Clean up vendor name - collapse all whitespace including newlines into single spaces
            item.vendor_name = re.sub(r'\s+', ' ', raw_vendor_name).strip()
            item.cubes = match.group(3)
            item.weight = match.group(4)
            item.pieces = match.group(5)
            item.po = match.group(7)  # PO number from group 7
            item.match_end = match.end()
            
            log.info(f"Captured line item: VendorNo={item.vendor_no}, VendorName={item.vendor_name}, PO={item.po}")
            line_items.append(item)
        
        log.info(f"Found {len(line_items)} line items in Template-2")
        
        # Extract pallet count data (product descriptions)
        pallet_pattern = re.compile(r'Pallet\s+Count:\s*([^\n]+?)(?=\s*(?:Pickup|Delivery|$))', re.MULTILINE | re.DOTALL)
        pallet_data = []
        for match in pallet_pattern.finditer(text):
            pallet = match.group(1).strip()
            if pallet:
                log.info(f"Captured pallet data: {pallet}")
                pallet_data.append(pallet)
        
        # Extract pickup dates - handle both "Pickup :" and "Pickup:"
        pickup_pattern = re.compile(r'Pickup\s*:\s*([A-Za-z]{3}\s+\d{1,2},\s+\d{4}\s+\d{1,2}:\d{2}:\d{2}\s*(?:AM|PM))', re.IGNORECASE)
        
        # Extract delivery dates - handle "Deliver :", "Delivery :", and variations
        deliver_pattern = re.compile(r'Deliver(?:y)?\s*:\s*([A-Za-z]{3}\s+\d{1,2},\s+\d{4}[\s\S]*?\d{1,2}:\d{2}:\d{2}\s*(?:AM|PM))', re.IGNORECASE)
        
        pickup_dates = [match.group(1) for match in pickup_pattern.finditer(text)]
        delivery_dates = [re.sub(r'\s+', ' ', match.group(1)).strip() for match in deliver_pattern.finditer(text)]
        
        log.info(f"Found {len(pickup_dates)} pickup dates and {len(delivery_dates)} delivery dates")
        
        # Extract ship_from from Stop 1 (index 0) for Template-2
        ship_from = ""
        if len(stop_destinations) > 0:
            ship_from = stop_destinations[0]  # Stop 1 is at index 0
            log.info(f"Template-2: Using Stop 1 as ship_from: {ship_from}")
        
        # Process each line item
        for idx, item in enumerate(line_items):
            # Get dates for this line item
            pickup_date = pickup_dates[idx] if idx < len(pickup_dates) else ""
            delivery_date = delivery_dates[idx] if idx < len(delivery_dates) else ""
            
            # Convert date formats
            formatted_delivery_date = self.convert_date_format(delivery_date) if delivery_date else ""
            formatted_pickup_date = self.convert_date_format(pickup_date) if pickup_date else ""
            
            # Extract ship_to based on stop number in vendor name
            ship_to = ""
            vendor_stop_pattern = re.compile(r'-\s*(\d+)\s*-\s*(.+)$')
            vendor_stop_match = vendor_stop_pattern.search(item.vendor_name)
            
            if vendor_stop_match:
                vendor_stop_num = vendor_stop_match.group(1).strip()
                log.info(f"Extracted stop number '{vendor_stop_num}' from vendor name: {item.vendor_name}")
                
                # Find matching stop destination
                for stop_idx in range(len(stop_destinations)):
                    stop_dest = stop_destinations[stop_idx]
                    stop_num_pattern = re.compile(r'^(\d+),')
                    stop_num_match = stop_num_pattern.search(stop_dest)
                    if stop_num_match:
                        dest_stop_num = stop_num_match.group(1).strip()
                        if dest_stop_num == vendor_stop_num:
                            ship_to = stop_dest
                            log.info(f"Matched vendor stop '{vendor_stop_num}' with destination stop: {stop_dest}")
                            break
                
                if not ship_to:
                    log.warning(f"No matching destination found for stop number: {vendor_stop_num}")
            else:
                log.warning(f"Could not extract stop number from vendor name: {item.vendor_name}")
            
            # Get description from pallet data
            description = pallet_data[idx] if idx < len(pallet_data) else ""
            
            # Extract clean vendor name
            clean_vendor_name = self.extract_vendor_name(item.vendor_name)
            log.info(f"Processing line item {idx}: Full vendor='{item.vendor_name}' -> Clean vendor='{clean_vendor_name}'")
            
            # Extract shipment type from PO
            shipment_type = self.extract_shipment_type(text, item.po)
            
            record = {
                'template': 'Template-1',  # Hardcoded as per original Java code
                'del_date': formatted_delivery_date,
                'pickup_date': formatted_pickup_date,
                'ship_from': ship_from,
                'ship_to': ship_to,
                'vendor_no': item.vendor_no,
                'vendor_name': clean_vendor_name,
                'cubes': item.cubes,
                'weight': item.weight,
                'pieces': item.pieces,
                'po': item.po,
                'shipment_type': shipment_type,
                'pallets': '',
                'description': description
            }
            
            records.append(record)
        
        return records
    
    def parse_pdf(self, pdf_path: str) -> Dict:
        """Main PDF parsing function"""
        try:
            # Check if file exists
            pdf_file = Path(pdf_path)
            if not pdf_file.exists():
                return {
                    "error": f"PDF file not found at path: {pdf_path}",
                    "capability": "parse_pdf"
                }
            
            # Extract text from PDF
            text = self.extract_text_from_pdf(str(pdf_file))
            
            # Detect template type
            template_type = self.detect_template(text)
            
            # Extract stop destinations
            stop_destinations = self.extract_stop_destinations(text)
            
            # Process based on template type
            if template_type == "Template-1":
                records = self.process_template1(text, stop_destinations)
            else:  # Template-2
                records = self.process_template2(text, stop_destinations)
            
            return {
                "result": {
                    "template_type": template_type,
                    "records_count": len(records),
                    "records": records,
                    "file_name": pdf_file.name,
                    "processed_at": datetime.now().isoformat()
                },
                "capability": "parse_pdf"
            }
            
        except Exception as e:
            log.error(f"Error processing PDF: {e}", exc_info=True)
            return {
                "error": str(e),
                "capability": "parse_pdf"
            }
    
    def parse_directory(self, directory_path: str) -> Dict:
        """Parse all PDF files in a directory"""
        try:
            dir_path = Path(directory_path)
            if not dir_path.exists() or not dir_path.is_dir():
                return {
                    "error": f"Directory not found: {directory_path}",
                    "capability": "parse_directory"
                }
            
            pdf_files = list(dir_path.glob("*.pdf"))
            if not pdf_files:
                return {
                    "error": f"No PDF files found in directory: {directory_path}",
                    "capability": "parse_directory"
                }
            
            results = []
            for pdf_file in pdf_files:
                result = self.parse_pdf(str(pdf_file))
                if "result" in result:
                    results.append(result["result"])
                else:
                    results.append({"file": str(pdf_file), "error": result.get("error", "Unknown error")})
            
            return {
                "result": {
                    "total_files": len(pdf_files),
                    "results": results,
                    "processed_at": datetime.now().isoformat()
                },
                "capability": "parse_directory"
            }
            
        except Exception as e:
            log.error(f"Error processing directory: {e}", exc_info=True)
            return {
                "error": str(e),
                "capability": "parse_directory"
            }


def main():
    """Main entry point - reads JSON from stdin; saves full result to file, prints only save location."""
    # Silence all logging to the terminal (full output is only in the saved text file)
    logging.getLogger().setLevel(logging.CRITICAL)
    for h in list(logging.root.handlers):
        logging.root.removeHandler(h)

    try:
        input_data = json.load(sys.stdin)

        capability = input_data.get("capability")
        args = input_data.get("args", {})

        parser = SobeyTemplate1PdfParser()

        if capability == "parse_pdf":
            pdf_path = args.get("pdf_path")
            if not pdf_path:
                result = {
                    "error": "Missing required parameter: pdf_path",
                    "capability": "parse_pdf"
                }
            else:
                result = parser.parse_pdf(pdf_path)
            source_for_name = pdf_path if isinstance(pdf_path, str) and not result.get("error") else None
            saved_path = save_result_to_downloads(result, source_for_name)
            print(json.dumps({"capability": "parse_pdf", "saved_to": saved_path}, indent=2))

        elif capability == "parse_directory":
            directory_path = args.get("directory_path")
            if not directory_path:
                result = {
                    "error": "Missing required parameter: directory_path",
                    "capability": "parse_directory"
                }
            else:
                result = parser.parse_directory(directory_path)
            saved_path = save_result_to_downloads(result, None)
            print(json.dumps({"capability": "parse_directory", "saved_to": saved_path}, indent=2))

        else:
            print(json.dumps({"error": f"Unknown capability: {capability}", "capability": capability}, indent=2))

    except Exception as e:
        err_result = {
            "error": str(e),
            "capability": "unknown"
        }
        saved_path = save_result_to_downloads(err_result, None)
        print(json.dumps({"capability": "unknown", "saved_to": saved_path}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()