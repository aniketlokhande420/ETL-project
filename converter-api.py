from flask import Flask, request, send_file, jsonify
import gdown
import os
import xml.etree.ElementTree as ET
from openpyxl import Workbook
import re
from io import BytesIO

app = Flask(__name__)

# Step 1: Download XML from Google Drive using gdown
def download_file_from_gdrive(url, output):
    gdown.download(url, output, quiet=False)

# Function to convert Google Drive URL to the required download URL format
def convert_drive_url_to_download_url(xml_url):
    # Extract file ID using regex
    match = re.search(r'd/([a-zA-Z0-9_-]+)', xml_url)
    if match:
        file_id = match.group(1)
        download_url = f"https://drive.google.com/uc?id={file_id}"
        return download_url
    else:
        raise ValueError("Invalid Google Drive URL format")

# Step 2: Parse the XML
def parse_xml(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()
    return root

# Step 3: Extract Data and Organize into a Structured Format
def extract_transactions(root):
    transactions = []
    
    for voucher in root.findall(".//VOUCHER"):
        vch_no = voucher.findtext('VOUCHERNUMBER')
        date = voucher.findtext('DATE')
        debtor = voucher.findtext('PARTYLEDGERNAME')
        amount = voucher.findtext('AMOUNT')
        particulars = debtor  
        transaction_type = 'Parent'

        transactions.append({
            'Date': date,
            'Transaction Type': transaction_type,
            'Vch No.': vch_no,
            'Ref No': 'NA',
            'Ref Type': 'NA',
            'Ref Date': 'NA',
            'Debtor': debtor,
            'Ref Amount': 'NA',
            'Amount': amount,
            'Particulars': particulars
        })

        for ledger_entry in voucher.findall(".//ALLLEDGERENTRIES.LIST"):
            ref_no = ledger_entry.findtext('BILLALLOCATIONS.LIST/NAME', 'NA')
            ref_type = ledger_entry.findtext('BILLALLOCATIONS.LIST/BILLTYPE', 'NA')
            ref_amount = ledger_entry.findtext('BILLALLOCATIONS.LIST/AMOUNT', 'NA')
            ref_date = ledger_entry.findtext('BILLALLOCATIONS.LIST/DUEDATEOFPYMT', 'NA')
            ledger_name = ledger_entry.findtext('LEDGERNAME')
            
            transactions.append({
                'Date': date,
                'Transaction Type': 'Child',
                'Vch No.': vch_no,
                'Ref No': ref_no,
                'Ref Type': ref_type,
                'Ref Date': ref_date,
                'Debtor': ledger_name,
                'Ref Amount': ref_amount,
                'Amount': 'NA',
                'Particulars': ledger_name
            })
        
        for ledger_entry in voucher.findall(".//ALLLEDGERENTRIES.LIST"):
            ledger_name = ledger_entry.findtext('LEDGERNAME')
            amount = ledger_entry.findtext('AMOUNT')
            
            transactions.append({
                'Date': date,
                'Transaction Type': 'Other',
                'Vch No.': vch_no,
                'Ref No': 'NA',
                'Ref Type': 'NA',
                'Ref Date': 'NA',
                'Debtor': ledger_name,
                'Ref Amount': 'NA',
                'Amount': amount,
                'Particulars': ledger_name
            })
    
    return transactions

# Step 4: Write Data to Excel (in-memory)
def write_to_xlsx(transactions):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    
    headers = ['Date', 'Transaction Type', 'Vch No.', 'Ref No', 'Ref Type', 'Ref Date', 'Debtor', 'Ref Amount', 'Amount', 'Particulars']
    ws.append(headers)
    
    for transaction in transactions:
        ws.append([
            transaction['Date'],
            transaction['Transaction Type'],
            transaction['Vch No.'],
            transaction['Ref No'],
            transaction['Ref Type'],
            transaction['Ref Date'],
            transaction['Debtor'],
            transaction['Ref Amount'],
            transaction['Amount'],
            transaction['Particulars']
        ])

    # Save the workbook in-memory using BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)  # Rewind the buffer
    return output

@app.route('/convert', methods=['POST'])
def convert_xml_to_xlsx():
    try:
        # Get the Google Drive URL from the request
        data = request.get_json()
        xml_url = data.get('xml_url')
        if not xml_url:
            return jsonify({"error": "No file URL provided"}), 400
        
        # Convert the provided URL to the download URL
        try:
            download_url = convert_drive_url_to_download_url(xml_url)
        except ValueError as ve:
            return jsonify({"error": str(ve)}), 400
        
        # Step 1: Download the file
        xml_file = 'downloaded_input.xml'
        download_file_from_gdrive(download_url, xml_file)

        # Step 2: Parse the XML
        root = parse_xml(xml_file)

        # Step 3: Extract transactions
        transactions = extract_transactions(root)

        # Step 4: Write to XLSX in-memory
        output_file = write_to_xlsx(transactions)

        # Step 5: Send the XLSX file as a response (from memory)
        return send_file(output_file, as_attachment=True, download_name='output.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        return jsonify({"error": str(e)}), 500
    
    finally:
        if os.path.exists(xml_file):
            os.remove(xml_file)

if __name__ == '__main__':
    app.run(debug=True)
