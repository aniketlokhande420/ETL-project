import xml.etree.ElementTree as ET
from openpyxl import Workbook

# Step 1: Parse the XML
def parse_xml(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()
    return root

# Step 2: Extract Data and Organize into a Structured Format
def extract_transactions(root):
    transactions = []
    
    # Process Vouchers (Parent Transactions)
    for voucher in root.findall(".//VOUCHER"):
        vch_no = voucher.findtext('VOUCHERNUMBER')
        date = voucher.findtext('DATE')
        debtor = voucher.findtext('PARTYLEDGERNAME')
        amount = voucher.findtext('AMOUNT')
        particulars = debtor  # For Parent, Particulars is the same as debtor
        transaction_type = 'Parent'

        # Add the Parent transaction
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
        
        # Process Child Transactions (Reference Entries)
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
        
        # Process Other Transactions
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

# Step 3: Write Data to Excel
def write_to_xlsx(transactions, output_file):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    
    # Write header
    headers = ['Date', 'Transaction Type', 'Vch No.', 'Ref No', 'Ref Type', 'Ref Date', 'Debtor', 'Ref Amount', 'Amount', 'Particulars']
    ws.append(headers)
    
    # Write transaction rows
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
    
    # Save the workbook
    wb.save(output_file)

# Main function to run the conversion
if __name__ == '__main__':
    xml_file = 'Input.xml'  # Path to the input XML file
    output_file = 'Output.xlsx'  # Path to the output XLSX file

    # Step 1: Parse XML
    root = parse_xml(xml_file)

    # Step 2: Extract Transactions
    transactions = extract_transactions(root)

    # Step 3: Write Data to XLSX
    write_to_xlsx(transactions, output_file)

    print(f"Data has been successfully written to {output_file}")

