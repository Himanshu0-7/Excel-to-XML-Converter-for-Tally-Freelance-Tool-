import html
import tkinter as tk
import pandas as pd
from tkinter import scrolledtext, ttk
from openpyxl import Workbook
from tkinter import filedialog, messagebox, scrolledtext
from openpyxl.worksheet.datavalidation import DataValidation
from tkinter import font


# Global variable to hold XML content
xml_content = ""
xml_master = ''
xml_purchase = ''
xml_Sales = ''
xml_sales_master = ''
xml_purchase_master = ''
def safe_float(value):
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0.0

def convert_to_voucher_xml(file_path):
    global xml_content
    try:
        # Load Excel
        df = pd.read_excel(file_path)

        # Create Tally XML structure
        xml_data = [
            '<ENVELOPE>',
            '  <HEADER>',
            '    <TALLYREQUEST>Import Data</TALLYREQUEST>',
            '  </HEADER>',
            '  <BODY>',
            '    <IMPORTDATA>',
            '      <REQUESTDESC>',
            '        <REPORTNAME>Vouchers</REPORTNAME>',
            '        <STATICVARIABLES>',
            '          <SVCURRENTCOMPANY></SVCURRENTCOMPANY>',
            '        </STATICVARIABLES>',
            '      </REQUESTDESC>',
            '      <REQUESTDATA>'
        ]

        # Loop through Excel rows -> Create vouchers
        df.columns = df.columns.str.replace(" ", "", regex=False).str.lower()
        for _, row in df.iterrows():
            date_value = row.get("date",'')
            if pd.notna(date_value):
                try:
                    date_value = pd.to_datetime(date_value,dayfirst=True, errors="coerce")
                    date = date_value.strftime("%Y%m%d") if pd.notna(date_value) else ""
                except Exception:
                    date = ""
            else:
                date = ""

            party = str(row.get('ledgername', "")).strip().lower()
            Ref = str(row.get('refno./chequeno.', '')).strip()
            Description = str(row.get('description', '')).strip()
            Remark = str(row.get("remarks", '') ).strip() 
            Debit_val = row.get("debit", 0)
            Credit_val = row.get("credit", 0)
            Amount = str(row.get('amount','')).strip().lower()
            Dr_Cr = str(row.get("dr/cr",'')).strip().lower()
            Debit = safe_float(Debit_val) if pd.notna(Debit_val) else 0.0
            Credit = safe_float(Credit_val) if pd.notna(Credit_val) else 0.0
            Bank_name = str(row.get('bankname', '')).strip()


            if Amount.endswith('(dr)'):
                num = safe_float(Amount.replace('(dr)','').strip())
                Debit += num
            elif Amount.endswith('(cr)'):
                num = safe_float(Amount.replace('(cr)','').strip())
                Credit += num
            elif Dr_Cr in ('dr', 'cr') and Amount:  # fallback to Dr / Cr column
                num = safe_float(Amount)
                if Dr_Cr == 'dr':
                    Debit += num
                else:
                    Credit += num
            # print(Debit)
            party_lower=party
            if party_lower == 'cash':
                        if Debit > 0.0:
                            vch_typ = 'Contra'
                            xml_data.extend([
                    '        <TALLYMESSAGE xmlns:UDF="TallyUDF">',
                    f'          <VOUCHER VCHTYPE="{vch_typ}" ACTION="Create" OBJVIEW="Accounting Voucher View">',
                    f'          <VOUCHERTYPENAME>{vch_typ}</VOUCHERTYPENAME>',
                    f'            <DATE>{date}</DATE>',
                    f'            <PARTYLEDGER></PARTYLEDGER>',
                    f'            <Narration>{html.escape(Description)+html.escape(Ref)+html.escape(Remark)}</Narration>',


                    f'            <ALLLEDGERENTRIES.LIST>',
                    f'              <LEDGERNAME>{Bank_name}</LEDGERNAME>',
                    f'              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>',
                    f'              <AMOUNT>{Debit:.2f}</AMOUNT>',
                    f'            </ALLLEDGERENTRIES.LIST>',


                    f'            <ALLLEDGERENTRIES.LIST>',
                    f'              <LEDGERNAME>{html.escape(party)}</LEDGERNAME>',
                    f'              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>',
                    f'              <AMOUNT>-{Debit:.2f}</AMOUNT>',
                    f'            </ALLLEDGERENTRIES.LIST>',
                    '          </VOUCHER>',
                    '        </TALLYMESSAGE>'
                    ])  # build Contra XML here

            

                        elif Credit > 0.0:
                            vch_typ = 'Contra'
                            xml_data.extend([
                    '        <TALLYMESSAGE xmlns:UDF="TallyUDF">',
                    f'          <VOUCHER VCHTYPE="{vch_typ}" ACTION="Create" OBJVIEW="Accounting Voucher View">',
                    f'          <VOUCHERTYPENAME>{vch_typ}</VOUCHERTYPENAME>',
                    f'            <DATE>{date}</DATE>',
                    f'            <PARTYLEDGER>Party</PARTYLEDGER>',
                    f'            <Narration>{html.escape(Description) + html.escape(Remark)}</Narration>',


                    f'            <ALLLEDGERENTRIES.LIST>',
                    f'              <LEDGERNAME>{Bank_name}</LEDGERNAME>',
                    f'              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>',
                    f'              <AMOUNT>-{Credit:.2f}</AMOUNT>',
                    f'            </ALLLEDGERENTRIES.LIST>',


                    f'            <ALLLEDGERENTRIES.LIST>',
                    f'              <LEDGERNAME>{html.escape(party)}</LEDGERNAME>',
                    f'              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>',
                    f'              <AMOUNT>{Credit:.2f}</AMOUNT>',
                    f'            </ALLLEDGERENTRIES.LIST>',
                    '          </VOUCHER>',
                    '        </TALLYMESSAGE>'
                ])  # build Contra XML here
        
            else:
                if  Credit > 0.0:
                    vch_typ = 'Receipt'
                    xml_data.extend([
                    '        <TALLYMESSAGE xmlns:UDF="TallyUDF">',
                    f'          <VOUCHER VCHTYPE="{vch_typ}" ACTION="Create" OBJVIEW="Accounting Voucher View">',
                    f'          <VOUCHERTYPENAME>{vch_typ}</VOUCHERTYPENAME>',
                    f'            <DATE>{date}</DATE>',
                    f'            <PARTYLEDGER>Party</PARTYLEDGER>',
                    f'            <Narration>{html.escape(Description) + html.escape(Remark)}</Narration>',


                    f'            <ALLLEDGERENTRIES.LIST>',
                    f'              <LEDGERNAME>{Bank_name}</LEDGERNAME>',
                    f'              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>',
                    f'              <AMOUNT>-{Credit:.2f}</AMOUNT>',
                    f'            </ALLLEDGERENTRIES.LIST>',


                    f'            <ALLLEDGERENTRIES.LIST>',
                    f'              <LEDGERNAME>{html.escape(party)}</LEDGERNAME>',
                    f'              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>',
                    f'              <AMOUNT>{Credit:.2f}</AMOUNT>',
                    f'            </ALLLEDGERENTRIES.LIST>',

                    '          </VOUCHER>',
                    '        </TALLYMESSAGE>'
                ])  # build Receipt XML
                    
                elif Debit > 0.0:
                 vch_typ = 'Payment'
                 xml_data.extend([
                    '        <TALLYMESSAGE xmlns:UDF="TallyUDF">',
                    f'          <VOUCHER VCHTYPE="{vch_typ}" ACTION="Create" OBJVIEW="Accounting Voucher View">',
                    f'          <VOUCHERTYPENAME>{vch_typ}</VOUCHERTYPENAME>',
                    f'            <DATE>{date}</DATE>',
                    f'            <PARTYLEDGER>Party</PARTYLEDGER>',
                    f'            <Narration>{html.escape(Description)+html.escape(Remark)}</Narration>',


                    f'            <ALLLEDGERENTRIES.LIST>',
                    f'              <LEDGERNAME>{Bank_name}</LEDGERNAME>',
                    f'              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>',
                    f'              <AMOUNT>{Debit:.2f}</AMOUNT>',
                    f'            </ALLLEDGERENTRIES.LIST>',


                    f'            <ALLLEDGERENTRIES.LIST>',
                    f'              <LEDGERNAME>{html.escape(party)}</LEDGERNAME>',
                    f'              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>',
                    f'              <AMOUNT>-{Debit:.2f}</AMOUNT>',
                    f'            </ALLLEDGERENTRIES.LIST>',

                    '          </VOUCHER>',
                    '        </TALLYMESSAGE>'
                ])  # build Payment XML


        xml_data.extend([
            '      </REQUESTDATA>',
            '    </IMPORTDATA>',
            '  </BODY>',
            '</ENVELOPE>'
        ])

        xml_content = "\n".join(xml_data)

        # Show in text box
        text_box1.delete("1.0", tk.END)
        text_box1.insert(tk.END, xml_content)

        messagebox.showinfo("Success", "File converted to Tally XML! Now you can save it.")

    except Exception as e:
        messagebox.showerror("Error", str(e))

def convert_to_master_xml(file_path):
    global xml_master
    try:
        # Load Excel
        df = pd.read_excel(file_path)

        # Create Tally XML structure
        xml_master_data = [
           '<ENVELOPE>',
            '  <HEADER>',
            '    <TALLYREQUEST>Import Data</TALLYREQUEST>',
            '  </HEADER>',
            '  <BODY>',
            '    <IMPORTDATA>',
            '      <REQUESTDESC>',
            '        <REPORTNAME>All Masters</REPORTNAME>',
            '      </REQUESTDESC>',
            '      <REQUESTDATA>'
        ]

        for _, row in df.iterrows():
            group_mapping = {
    # ------------------ Cash / Bank ------------------
    "cash": "Cash-in-Hand",
    "cash in hand": "Cash-in-Hand",
    "petty cash": "Cash-in-Hand",
    "hand cash": "Cash-in-Hand",
    "bank": "Bank Accounts",
    "bank account": "Bank Accounts",
    "bank accounts": "Bank Accounts",
    "saving account": "Bank Accounts",
    "current account": "Bank Accounts",
    "bank od": "Bank OD A/c",
    "od account": "Bank OD A/c",
    "overdraft": "Bank OD A/c",
    
    # ------------------ Expenses ------------------
    "personal expense": "Indirect Expenses",
    "persional expense": "Indirect Expenses",   # common spelling mistake
    "indirect expense": "Indirect Expenses",
    "indirect expenses": "Indirect Expenses",
    "general expense": "Indirect Expenses",
    "office expense": "Indirect Expenses",
    "admin expense": "Indirect Expenses",
    "administrative expense": "Indirect Expenses",
    "direct expense": "Direct Expenses",
    "direct expenses": "Direct Expenses",
    "purchase expense": "Direct Expenses",
    
    # ------------------ Incomes ------------------
    "direct income": "Direct Incomes",
    "direct incomes": "Direct Incomes",
    "indirect income": "Indirect Incomes",
    "indirect incomes": "Indirect Incomes",
    "misc income": "Indirect Incomes",
    "interest income": "Indirect Incomes",
    "commission income": "Indirect Incomes",
    
    # ------------------ Sales & Purchase ------------------
    "sales": "Sales Accounts",
    "sale": "Sales Accounts",
    "sales account": "Sales Accounts",
    "purchase": "Purchase Accounts",
    "purchases": "Purchase Accounts",
    "purchase account": "Purchase Accounts",
    
    # ------------------ Sundry ------------------
    "sundry creditor": "Sundry Creditors",
    "sundry creditors": "Sundry Creditors",
    "creditor": "Sundry Creditors",
    "creditors": "Sundry Creditors",
    "sundry debtor": "Sundry Debtors",
    "sundry debtors": "Sundry Debtors",
    "debtor": "Sundry Debtors",
    "debtors": "Sundry Debtors",
    "customer": "Sundry Debtors",
    "vendor": "Sundry Creditors",
    "supplier": "Sundry Creditors",
    
    # ------------------ Loans & Capital ------------------
    "secured loan": "Secured Loans",
    "secured loans": "Secured Loans",
    "unsecured loan": "Unsecured Loans",
    "unsecured loans": "Unsecured Loans",
    "loan": "Loans (Liability)",
    "loans": "Loans (Liability)",
    "capital": "Capital Account",
    "capital account": "Capital Account",
    "proprietor capital": "Capital Account",
    "partner capital": "Capital Account",
    
    # ------------------ Misc ------------------
    "suspense": "Suspense A/c",
    "duties & taxes": "Duties &amp; Taxes",
    "duties and taxes": "Duties &amp; Taxes",
    "tax": "Duties &amp; Taxes",
    "gst": "Duties &amp; Taxes",
    "igst": "Duties &amp; Taxes",
    "cgst": "Duties &amp; Taxes",
    "sgst": "Duties &amp; Taxes",
    "vat": "Duties &amp; Taxes",
    "tds": "Duties &amp; Taxes",
    "provision": "Provisions",
    "reserves": "Reserves &amp; Surplus",
    "reserve": "Reserves &amp; Surplus",
    "investment": "Investments",
    "investments": "Investments",
    "fixed asset": "Fixed Assets",
    "fixed assets": "Fixed Assets",
    "furniture": "Fixed Assets",
    "machinery": "Fixed Assets",
    "vehicle": "Fixed Assets",
    "loan & advance": "Loans &amp; Advances (Asset)",
    "loans & advances": "Loans &amp; Advances (Asset)",
    "advance": "Loans &amp; Advances (Asset)",
    "deposit": "Deposits (Asset)",
    "deposits": "Deposits (Asset)",
    "fd": "Deposits (Asset)",
    "branch": "Branch / Divisions",
    "division": "Branch / Divisions",
}
            party = str(row.get('Ledger Name', "")).strip()
            underGroup_raw = str(row.get('Under Group', '')).strip().lower()
            underGroup = group_mapping.get(underGroup_raw, underGroup_raw.title())
            xml_master_data.extend([
            '<TALLYMESSAGE xmlns:UDF="TallyUDF">',
            f'       <LEDGER NAME="{html.escape(party)}" ACTION="Create/Alter">',
            f'           <NAME>{html.escape(party)}</NAME>',
            f'            <PARENT>{html.escape(underGroup)}</PARENT>',
            f'            <ISBILLWISEON>No</ISBILLWISEON>',
            f'            <ISCOSTCENTRESON>No</ISCOSTCENTRESON>',
            ])
            xml_master_data.extend([
            '        </LEDGER>',
            '    </TALLYMESSAGE>',
        ])
                  
        xml_master_data.extend([
            '      </REQUESTDATA>',
            '    </IMPORTDATA>',
            '  </BODY>',
            '</ENVELOPE>'
        ])
        xml_master = '\n'.join(xml_master_data)
        text_box2.delete("1.0", tk.END)
        text_box2.insert(tk.END, xml_master)
    except Exception as e:
        messagebox.showerror("Error", str(e))


def convert_to_Sales_master_xml(file_path):
    global xml_sales_master
    try:
        # Load Excel
        df = pd.read_excel(file_path)

        # Create Tally XML structure

        xml_sales_master_data = [
           '<ENVELOPE>',
            '  <HEADER>',
            '    <TALLYREQUEST>Import Data</TALLYREQUEST>',
            '  </HEADER>',
            '  <BODY>',
            '    <IMPORTDATA>',
            '      <REQUESTDESC>',
            '        <REPORTNAME>All Masters</REPORTNAME>',
            '      </REQUESTDESC>',
            '      <REQUESTDATA>'
        ]
        seen = set()
        for _, row in df.iterrows():
            party = str(row.get('Party Name', "")).strip()
            Sales = str(row.get('Sales Name', "")).strip()
            cgst_name = str(row.get('Cgst Ledger Name', '')).strip().lower()
            sgst_name = str(row.get('Sgst Ledger Name', '')).strip().lower()
            key = (party, Sales,cgst_name, sgst_name)
    
    # Skip if already seen
            
            
            if key in seen:
                 continue
            xml_sales_master_data.extend([
            '<TALLYMESSAGE xmlns:UDF="TallyUDF">',
            f'       <LEDGER NAME="{html.escape(party)}" ACTION="Create/Alter">',
            f'           <NAME>{html.escape(party)}</NAME>',
            f'            <PARENT>Sundry Debtors</PARENT>',
            f'            <ISGSTAPPLICABLE>Yes</ISGSTAPPLICABLE>',
            '        </LEDGER>',
            '</TALLYMESSAGE>',
            ])
            xml_sales_master_data.extend([
            '<TALLYMESSAGE xmlns:UDF="TallyUDF">',
            f'       <LEDGER NAME="{html.escape(Sales)}" ACTION="Create/Alter">',
            f'           <NAME>{html.escape(Sales)}</NAME>',
            f'            <PARENT>Sales Accounts</PARENT>',
            f'            <ISGSTAPPLICABLE>Yes</ISGSTAPPLICABLE>',
            '        </LEDGER>',
            '</TALLYMESSAGE>',
            ])
            xml_sales_master_data.extend([
            '<TALLYMESSAGE xmlns:UDF="TallyUDF">',
            f'       <LEDGER NAME="{html.escape(cgst_name)}" ACTION="Create/Alter">',
            f'           <NAME>{html.escape(cgst_name)}</NAME>',
            f'            <PARENT>{html.escape('Duties & Taxes')}</PARENT>',
            f'            <ISGSTAPPLICABLE>Yes</ISGSTAPPLICABLE>',
            '        </LEDGER>',
            '</TALLYMESSAGE>',
            ])
            xml_sales_master_data.extend([
            '<TALLYMESSAGE xmlns:UDF="TallyUDF">',
            f'       <LEDGER NAME="{html.escape(sgst_name)}" ACTION="Create/Alter">',
            f'           <NAME>{html.escape(sgst_name)}</NAME>',
            f'            <PARENT>{html.escape('Duties & Taxes')}</PARENT>',
            f'            <ISGSTAPPLICABLE>Yes</ISGSTAPPLICABLE>',
            '        </LEDGER>',
            '</TALLYMESSAGE>',
            ])
            seen.add(key)

        xml_sales_master_data.extend([
            '      </REQUESTDATA>',
            '    </IMPORTDATA>',
            '  </BODY>',
            '</ENVELOPE>'
        ])

        xml_sales_master = '\n'.join(xml_sales_master_data)
        text_box2.delete("1.0", tk.END)
        text_box2.insert(tk.END, xml_sales_master)

    except Exception as e:
        messagebox.showerror("Error", str(e))

def convert_to_purchase_master_xml(file_path):
    global xml_purchase_master
    try:
        # Load Excel
        df = pd.read_excel(file_path)

        # Create Tally XML structure

        xml_purchase_master_data = [
           '<ENVELOPE>',
            '  <HEADER>',
            '    <TALLYREQUEST>Import Data</TALLYREQUEST>',
            '  </HEADER>',
            '  <BODY>',
            '    <IMPORTDATA>',
            '      <REQUESTDESC>',
            '        <REPORTNAME>All Masters</REPORTNAME>',
            '      </REQUESTDESC>',
            '      <REQUESTDATA>'
        ]
        seen = set()
        for _, row in df.iterrows():
            party = str(row.get('Party Name', "")).strip()
            purchase = str(row.get('Purchase Name', "")).strip()
            cgst_name = str(row.get('Cgst Ledger Name', '')).strip().lower()
            sgst_name = str(row.get('Sgst Ledger Name', '')).strip().lower()
            key = (party, purchase,cgst_name, sgst_name)
    
    # Skip if already seen
            
            
            if key in seen:
                 continue
            xml_purchase_master_data.extend([
            '<TALLYMESSAGE xmlns:UDF="TallyUDF">',
            f'       <LEDGER NAME="{html.escape(party)}" ACTION="Create/Alter">',
            f'           <NAME>{html.escape(party)}</NAME>',
            f'            <PARENT>Sundry Debtors</PARENT>',
            f'            <ISGSTAPPLICABLE>Yes</ISGSTAPPLICABLE>',
            '        </LEDGER>',
            '</TALLYMESSAGE>',
            ])
            xml_purchase_master_data.extend([
            '<TALLYMESSAGE xmlns:UDF="TallyUDF">',
            f'       <LEDGER NAME="{html.escape(purchase)}" ACTION="Create/Alter">',
            f'           <NAME>{html.escape(purchase)}</NAME>',
            f'            <PARENT>Sales Accounts</PARENT>',
            f'            <ISGSTAPPLICABLE>Yes</ISGSTAPPLICABLE>',
            '        </LEDGER>',
            '</TALLYMESSAGE>',
            ])
            xml_purchase_master_data.extend([
            '<TALLYMESSAGE xmlns:UDF="TallyUDF">',
            f'       <LEDGER NAME="{html.escape(cgst_name)}" ACTION="Create/Alter">',
            f'           <NAME>{html.escape(cgst_name)}</NAME>',
            f'            <PARENT>{html.escape('Duties & Taxes')}</PARENT>',
            f'            <ISGSTAPPLICABLE>Yes</ISGSTAPPLICABLE>',
            '        </LEDGER>',
            '</TALLYMESSAGE>',
            ])
            xml_purchase_master_data.extend([
            '<TALLYMESSAGE xmlns:UDF="TallyUDF">',
            f'       <LEDGER NAME="{html.escape(sgst_name)}" ACTION="Create/Alter">',
            f'           <NAME>{html.escape(sgst_name)}</NAME>',
            f'            <PARENT>{html.escape('Duties & Taxes')}</PARENT>',
            f'            <ISGSTAPPLICABLE>Yes</ISGSTAPPLICABLE>',
            '        </LEDGER>',
            '</TALLYMESSAGE>',
            ])
            seen.add(key)

        xml_purchase_master_data.extend([
            '      </REQUESTDATA>',
            '    </IMPORTDATA>',
            '  </BODY>',
            '</ENVELOPE>'
        ])

        xml_purchase_master = '\n'.join(xml_purchase_master_data)
        text_box2.delete("1.0", tk.END)
        text_box2.insert(tk.END, xml_purchase_master)

    except Exception as e:
        messagebox.showerror("Error", str(e))


def convert_to_Sales_xml(file__path):
    global xml_Sales
    try:
        # Load Excel
        df = pd.read_excel(file__path)

        # Create Tally XML structure
        xml_Sales_data = [
            '<ENVELOPE>',
            '  <HEADER>',
            '    <TALLYREQUEST>Import Data</TALLYREQUEST>',
            '  </HEADER>',
            '  <BODY>',
            '    <IMPORTDATA>',
            '      <REQUESTDESC>',
            '        <REPORTNAME>Vouchers</REPORTNAME>',
            '        <STATICVARIABLES>',
            '          <SVCURRENTCOMPANY>Don</SVCURRENTCOMPANY>',
            '        </STATICVARIABLES>',
            '      </REQUESTDESC>',
            '      <REQUESTDATA>'
        ]

        for _, row in df.iterrows():
            group_mapping = {
    # Cash / Bank
            "cash": "Cash-in-Hand",
            "cash in hand": "Cash-in-Hand",
            "petty cash": "Cash-in-Hand",
            "bank": "Bank Accounts",
            "bank account": "Bank Accounts",
            "bank od": "Bank OD A/c",
            "od account": "Bank OD A/c",
        
            # Expenses
            "personal expense": "Indirect Expenses",
            "persional expense": "Indirect Expenses",  # common spelling mistake
            "indirect expense": "Indirect Expenses",
            "indirect expenses": "Indirect Expenses",
            "direct expense": "Direct Expenses",
            "direct expenses": "Direct Expenses",
        
            # Incomes
            "direct income": "Direct Incomes",
            "direct incomes": "Direct Incomes",
            "indirect income": "Indirect Incomes",
            "indirect incomes": "Indirect Incomes",
        
            # Sales & Purchase
            "sales": "Sales Accounts",
            "sales account": "Sales Accounts",
            "purchase": "Purchase Accounts",
            "purchase account": "Purchase Accounts",
        
            # Sundry
            "sundry creditor": "Sundry Creditors",
            "sundry creditors": "Sundry Creditors",
            "creditor": "Sundry Creditors",
            "creditors": "Sundry Creditors",
            "sundry debtor": "Sundry Debtors",
            "sundry debtors": "Sundry Debtors",
            "debtor": "Sundry Debtors",
            "debtors": "Sundry Debtors",
        
            # Loans & Capital
            "secured loan": "Secured Loans",
            "secured loans": "Secured Loans",
            "unsecured loan": "Unsecured Loans",
            "unsecured loans": "Unsecured Loans",
            "capital": "Capital Account",
            "capital account": "Capital Account",
        
            # Misc
            "suspense": "Suspense A/c",
            "duties & taxes": "Duties & Taxes",
            "tax": "Duties & Taxes",
            "gst": "Duties & Taxes",
            "tds": "Duties & Taxes",
            "provision": "Provisions",
            "reserves": "Reserves & Surplus",
            "investment": "Investments",
            "fixed asset": "Fixed Assets",
            "fixed assets": "Fixed Assets",
            "loan & advance": "Loans & Advances (Asset)",
            "loans & advances": "Loans & Advances (Asset)",
            "deposit": "Deposits (Asset)",
            "deposits": "Deposits (Asset)",
            "branch": "Branch / Divisions",
}
            # underGroup_raw = str(row.get('Under Group', '')).strip().lower()
            # underGroup = group_mapping.get(underGroup_raw, underGroup_raw.title())

            date_value = row.get("Date",'')
            if pd.notna(date_value):
                try:
                    date_value = pd.to_datetime(date_value,dayfirst=True, errors="coerce")
                    date = date_value.strftime("%Y%m%d") if pd.notna(date_value) else ""
                except Exception:
                    date = ""
            else:
                date = ""
            party = str(row.get('Party Name', "")).strip()
            Sales = str(row.get('Sales Name', "")).strip()
            cgst_name = str(row.get('Cgst Ledger Name', '')).strip().lower()
            sgst_name = str(row.get('Sgst Ledger Name', '')).strip().lower()
            amount_val = row.get("Amount", 0)
            sales_amount_val = row.get("Sales Amount", 0)
            cgst_amount_val = row.get("Cgst Amount", 0)
            sgst_amount_val = row.get("Sgst Amount", 0)
            amount = safe_float(amount_val) if pd.notna(amount_val) else 0.0
            sales_amount = safe_float(sales_amount_val) if pd.notna(sales_amount_val) else 0.0
            cgst_amount = safe_float(cgst_amount_val) if pd.notna(cgst_amount_val) else 0.0
            sgst_amount = safe_float(sgst_amount_val) if pd.notna(sgst_amount_val) else 0.0
            voucher_no = str(row.get('Voucher No', '')).strip()
            ref_no =  str(row.get('Reference No', '')).strip()

        
            xml_Sales_data.extend([
            '<TALLYMESSAGE xmlns:UDF="TallyUDF">',
            f'<VOUCHER VCHTYPE="{'Sales'}" ACTION="Create" OBJVIEW="Accounting Voucher View">',
            f'            <VOUCHERTYPENAME>{'Sales'}</VOUCHERTYPENAME>',
            f'            <DATE>{date}</DATE>',
            f'            <VOUCHERNUMBER>{voucher_no}</VOUCHERNUMBER>',
            f'            <PARTYLEDGERNAME>{html.escape(party)}</PARTYLEDGERNAME>',
            f'            <REFERENCE>{ref_no}</REFERENCE>',
            f'            <ISINVOICE>Yes</ISINVOICE>',

            '           <ALLLEDGERENTRIES.LIST>',
            f'               <LEDGERNAME>{html.escape(party)}</LEDGERNAME>',
            f'               <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>',
            f'               <AMOUNT>-{amount}</AMOUNT>',
            '           </ALLLEDGERENTRIES.LIST>',

            '           <ALLLEDGERENTRIES.LIST>',
            f'               <LEDGERNAME>{html.escape(Sales)}</LEDGERNAME>',
            f'               <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>',
            f'               <AMOUNT>{sales_amount}</AMOUNT>',
            '           </ALLLEDGERENTRIES.LIST>',

            '            <ALLLEDGERENTRIES.LIST>',
            f'                <LEDGERNAME>{html.escape(cgst_name)}</LEDGERNAME>',
            f'                <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>',
            f'                <AMOUNT>{cgst_amount}</AMOUNT>',
            '            </ALLLEDGERENTRIES.LIST>',
    
            '            <ALLLEDGERENTRIES.LIST>',
            f'                <LEDGERNAME>{html.escape(sgst_name)}</LEDGERNAME>',
            f'                <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>',
            f'                <AMOUNT>{sgst_amount}</AMOUNT>',
            '            </ALLLEDGERENTRIES.LIST>',
    
            # '           <ALLINVENTORYENTRIES.LIST>'
            # f'               <STOCKITEMNAME></STOCKITEMNAME>'
            # f'               <ISDEEMEDPOSITIVE></ISDEEMEDPOSITIVE>'
            # f'               <RATE></RATE>'
            # f'               <AMOUNT></AMOUNT>'
            # f'               <ACTUALQTY></ACTUALQTY>'
            # f'               <BILLEDQTY></BILLEDQTY>'
            # '           </ALLINVENTORYENTRIES.LIST>'

            '        </VOUCHER>',
            '        </TALLYMESSAGE>',

           
            ])     
        xml_Sales_data.extend([
            '      </REQUESTDATA>',
            '    </IMPORTDATA>',
            '  </BODY>',
            '</ENVELOPE>'
        ])
        xml_Sales = '\n'.join(xml_Sales_data)
        text_box1.delete("1.0", tk.END)
        text_box1.insert(tk.END, xml_Sales)

    except Exception as e:
        messagebox.showerror("Error", str(e))




def convert_to_purchase_xml(file__path):
    global xml_purchase
    try:
        # Load Excel
        df = pd.read_excel(file__path)

        # Create Tally XML structure
        xml_purchase_data = [
            '<ENVELOPE>',
            '  <HEADER>',
            '    <TALLYREQUEST>Import Data</TALLYREQUEST>',
            '  </HEADER>',
            '  <BODY>',
            '    <IMPORTDATA>',
            '      <REQUESTDESC>',
            '        <REPORTNAME>Vouchers</REPORTNAME>',
            '        <STATICVARIABLES>',
            '          <SVCURRENTCOMPANY></SVCURRENTCOMPANY>',
            '        </STATICVARIABLES>',
            '      </REQUESTDESC>',
            '      <REQUESTDATA>'
        ]

        for _, row in df.iterrows():
            group_mapping = {
    # Cash / Bank
            "cash": "Cash-in-Hand",
            "cash in hand": "Cash-in-Hand",
            "petty cash": "Cash-in-Hand",
            "bank": "Bank Accounts",
            "bank account": "Bank Accounts",
            "bank od": "Bank OD A/c",
            "od account": "Bank OD A/c",
        
            # Expenses
            "personal expense": "Indirect Expenses",
            "persional expense": "Indirect Expenses",  # common spelling mistake
            "indirect expense": "Indirect Expenses",
            "indirect expenses": "Indirect Expenses",
            "direct expense": "Direct Expenses",
            "direct expenses": "Direct Expenses",
        
            # Incomes
            "direct income": "Direct Incomes",
            "direct incomes": "Direct Incomes",
            "indirect income": "Indirect Incomes",
            "indirect incomes": "Indirect Incomes",
        
            # Sales & Purchase
            "sales": "Sales Accounts",
            "sales account": "Sales Accounts",
            "purchase": "Purchase Accounts",
            "purchase account": "Purchase Accounts",
        
            # Sundry
            "sundry creditor": "Sundry Creditors",
            "sundry creditors": "Sundry Creditors",
            "creditor": "Sundry Creditors",
            "creditors": "Sundry Creditors",
            "sundry debtor": "Sundry Debtors",
            "sundry debtors": "Sundry Debtors",
            "debtor": "Sundry Debtors",
            "debtors": "Sundry Debtors",
        
            # Loans & Capital
            "secured loan": "Secured Loans",
            "secured loans": "Secured Loans",
            "unsecured loan": "Unsecured Loans",
            "unsecured loans": "Unsecured Loans",
            "capital": "Capital Account",
            "capital account": "Capital Account",
        
            # Misc
            "suspense": "Suspense A/c",
            "duties & taxes": "Duties & Taxes",
            "tax": "Duties & Taxes",
            "gst": "Duties & Taxes",
            "tds": "Duties & Taxes",
            "provision": "Provisions",
            "reserves": "Reserves & Surplus",
            "investment": "Investments",
            "fixed asset": "Fixed Assets",
            "fixed assets": "Fixed Assets",
            "loan & advance": "Loans & Advances (Asset)",
            "loans & advances": "Loans & Advances (Asset)",
            "deposit": "Deposits (Asset)",
            "deposits": "Deposits (Asset)",
            "branch": "Branch / Divisions",
}
            # underGroup_raw = str(row.get('Under Group', '')).strip().lower()
            # underGroup = group_mapping.get(underGroup_raw, underGroup_raw.title())

            date_value = row.get("Date",'')
            if pd.notna(date_value):
                try:
                    date_value = pd.to_datetime(date_value,dayfirst=True, errors="coerce")
                    date = date_value.strftime("%Y%m%d") if pd.notna(date_value) else ""
                except Exception:
                    date = ""
            else:
                date = ""
            party = str(row.get('Party Name', "")).strip()
            purchase = str(row.get('Purchase Name', "")).strip()
            cgst_name = str(row.get('Cgst Ledger Name', '')).strip().lower()
            sgst_name = str(row.get('Sgst Ledger Name', '')).strip().lower()
            amount_val = row.get("Amount", 0)
            purchase_amount_val = row.get("Purchase Amount", 0)
            cgst_amount_val = row.get("Cgst Amount", 0)
            sgst_amount_val = row.get("Sgst Amount", 0)
            amount = safe_float(amount_val) if pd.notna(amount_val) else 0.0
            sales_amount = safe_float(purchase_amount_val) if pd.notna(purchase_amount_val) else 0.0
            cgst_amount = safe_float(cgst_amount_val) if pd.notna(cgst_amount_val) else 0.0
            sgst_amount = safe_float(sgst_amount_val) if pd.notna(sgst_amount_val) else 0.0
            voucher_no = str(row.get('Voucher No', '')).strip()
            ref_no =  str(row.get('Reference No', '')).strip()

        
            xml_purchase_data.extend([
            '<TALLYMESSAGE xmlns:UDF="TallyUDF">',
            f'<VOUCHER VCHTYPE="{'Purchase'}" ACTION="Create" OBJVIEW="Accounting Voucher View">',
            f'            <VOUCHERTYPENAME>{'Purchase'}</VOUCHERTYPENAME>',
            f'            <DATE>{date}</DATE>',
            f'            <VOUCHERNUMBER>{voucher_no}</VOUCHERNUMBER>',
            f'            <PARTYLEDGERNAME>{html.escape(party)}</PARTYLEDGERNAME>',
            f'            <REFERENCE>{ref_no}</REFERENCE>',
            f'            <ISINVOICE>Yes</ISINVOICE>',

            '           <ALLLEDGERENTRIES.LIST>',
            f'               <LEDGERNAME>{html.escape(party)}</LEDGERNAME>',
            f'               <ISDEEMEDPOSITIVE>NO</ISDEEMEDPOSITIVE>',
            f'               <AMOUNT>{amount}</AMOUNT>',
            '           </ALLLEDGERENTRIES.LIST>',

            '           <ALLLEDGERENTRIES.LIST>',
            f'               <LEDGERNAME>{html.escape(purchase)}</LEDGERNAME>',
            f'               <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>',
            f'               <AMOUNT>-{sales_amount}</AMOUNT>',
            '           </ALLLEDGERENTRIES.LIST>',

            '            <ALLLEDGERENTRIES.LIST>',
            f'                <LEDGERNAME>{html.escape(cgst_name)}</LEDGERNAME>',
            f'                <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>',
            f'                <AMOUNT>-{cgst_amount}</AMOUNT>',
            '            </ALLLEDGERENTRIES.LIST>',
    
            '            <ALLLEDGERENTRIES.LIST>',
            f'                <LEDGERNAME>{html.escape(sgst_name)}</LEDGERNAME>',
            f'                <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>',
            f'                <AMOUNT>-{sgst_amount}</AMOUNT>',
            '            </ALLLEDGERENTRIES.LIST>',
    
            # '           <ALLINVENTORYENTRIES.LIST>'
            # f'               <STOCKITEMNAME></STOCKITEMNAME>'
            # f'               <ISDEEMEDPOSITIVE></ISDEEMEDPOSITIVE>'
            # f'               <RATE></RATE>'
            # f'               <AMOUNT></AMOUNT>'
            # f'               <ACTUALQTY></ACTUALQTY>'
            # f'               <BILLEDQTY></BILLEDQTY>'
            # '           </ALLINVENTORYENTRIES.LIST>'

            '        </VOUCHER>',
            '        </TALLYMESSAGE>',

           
            ])     
        xml_purchase_data.extend([
            '      </REQUESTDATA>',
            '    </IMPORTDATA>',
            '  </BODY>',
            '</ENVELOPE>'
        ])
        xml_purchase = '\n'.join(xml_purchase_data)
        text_box1.delete("1.0", tk.END)
        text_box1.insert(tk.END, xml_purchase)

    except Exception as e:
        messagebox.showerror("Error", str(e))

def upload_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file_path:
        convert_to_voucher_xml(file_path)
        convert_to_master_xml(file_path)

def upload_Sales_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file_path:
        # convert_to_voucher_xml(file_path)
        convert_to_Sales_xml(file_path)
        convert_to_Sales_master_xml(file_path)\
        
def upload_Purchase_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file_path:
        # convert_to_voucher_xml(file_path)
        convert_to_purchase_xml(file_path)
        convert_to_purchase_master_xml(file_path)


def save_voucher_file():
    global xml_content
    if not xml_content:
        messagebox.showwarning("Warning", "No XML content to save!")
        return

    output_path = filedialog.asksaveasfilename(
        defaultextension=".xml",
        filetypes=[("XML files", "*.xml")]
    )
    if output_path:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(xml_content)
        messagebox.showinfo("Saved", f"File saved at:\n{output_path}")

def save_master_file():
    global xml_master
    if not xml_master:
        messagebox.showwarning("Warning", "No XML content to save!")
        return

    output_path = filedialog.asksaveasfilename(
        defaultextension=".xml",
        filetypes=[("XML files", "*.xml")]
    )
    if output_path:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(xml_master)
        messagebox.showinfo("Saved", f"File saved at:\n{output_path}")

def save_Sales_file():
    global xml_Sales
    if not xml_Sales:
        messagebox.showwarning("Warning", "No XML content to save!")
        return

    output_path = filedialog.asksaveasfilename(
        defaultextension=".xml",
        filetypes=[("XML files", "*.xml")]
    )
    if output_path:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(xml_Sales)
        messagebox.showinfo("Saved", f"File saved at:\n{output_path}")

def save_Purchase_file():
    global xml_purchase
    if not xml_purchase:
        messagebox.showwarning("Warning", "No XML content to save!")
        return

    output_path = filedialog.asksaveasfilename(
        defaultextension=".xml",
        filetypes=[("XML files", "*.xml")]
    )
    if output_path:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(xml_purchase)
        messagebox.showinfo("Saved", f"File saved at:\n{output_path}")


def download_sample_file():
    try:
        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx *.xls")],
            title="Save Sample Excel File"
        )
        if not output_path:
            return  # User cancelled

        # Create new Excel workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Sample Data"

        # Add headers
        headers = [
            "Date", "Ref No./Cheque No.", "Description",
            "Debit", "Credit", "Bank Name","Ledger Name", "Under Group"
        ]
        ws.append(headers)

        # ✅ Create hidden sheet for group list
        ws_groups = wb.create_sheet(title="Groups")
        under_group_options = [
            "Branch/Divisions", "Capital Account", "Current Assets", "Current Liabilities",
            "Direct Expenses", "Direct Incomes", "Fixed Assets", "Indirect Expenses",
            "Indirect Incomes", "Investments", "Miscellaneous Expenses (Asset)",
            "Purchase Accounts", "Sales Accounts", "Suspense Account", "Provisions",
            "Bank Accounts", "Cash-in-Hand", "Deposits (Asset)", "Duties & Taxes",
            "Loans & Advances (Asset)", "Reserves & Surplus", "Secured Loans",
            "Stock-in-Hand", "Sundry Creditors", "Sundry Debtors", "Unsecured Loans",
            "Retained Earnings"
        ]

        # Write options in column A of Groups sheet
        for i, group in enumerate(under_group_options, start=1):
            ws_groups[f"A{i}"] = group

        # Hide the sheet
        ws_groups.sheet_state = "hidden"

        # Create Data Validation referencing the Groups sheet range
        last_row = len(under_group_options)
        dv = DataValidation(
            type="list",
            formula1=f"=Groups!$A$1:$A${last_row}",
            allow_blank=True
        )
        ws.add_data_validation(dv)

        # Apply validation to column H
        dv.add("H2:H1048576")

        # Save file
        wb.save(output_path)
        messagebox.showinfo("Success", f"Sample file saved at:\n{output_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))



# GUI setup
root = tk.Tk()
root.title("Excel to Tally XML Converter")
# root.geometry("1920x1080")
root.configure(bg="black")
root.state('zoomed')  # dark background

# Fonts (Poppins)
title_font = font.Font(family="Poppins", size=20, weight="bold")
label_font = font.Font(family="Poppins", size=14)
button_font = font.Font(family="Poppins", size=12, weight="bold")
text_font = font.Font(family="Poppins", size=10)


def Search_code_clipboard():
    vba_code = """
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Dim searchValue As String
    Dim cell As Range
    Dim rowContent As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim ValidationValues As Range
    Dim MatchCell As Range
    Dim results As Collection
    Dim tempList As String
    Dim SearchCell As Range
    
    Set ws = Me
    
    ' ----- Search Bar Functionality -----
    On Error Resume Next
    Set SearchCell = ws.Range("search")
    On Error GoTo 0
    
    If Not SearchCell Is Nothing Then
        If Not Intersect(Target, SearchCell) Is Nothing Then
            Application.EnableEvents = False
            Application.ScreenUpdating = False
            
            searchValue = Trim(SearchCell.Value)
            
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            lastCol = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
            
            ws.Rows.Hidden = False
            
            If searchValue <> "" Then
                For Each cell In ws.Range("A2:A" & lastRow)
                    rowContent = Join(Application.Transpose(Application.Transpose(ws.Range(ws.Cells(cell.Row, 1), ws.Cells(cell.Row, lastCol)).Value)), " ")
                    If InStr(1, rowContent, searchValue, vbTextCompare) = 0 Then
                        cell.EntireRow.Hidden = True
                    End If
                Next cell
            End If
            
            Application.ScreenUpdating = True
            Application.EnableEvents = True
        End If
    End If
    
    ' ----- Autocomplete Dropdown in "Under Group" Column -----
    If Target.Column = 8 And Target.Row > 1 Then
        On Error GoTo exitHandler
        Application.EnableEvents = False
        
        Set ValidationValues = Sheets("Groups").Range("A1:A100")
        Set results = New Collection
        
        If Len(Target.Value) > 0 Then
            For Each MatchCell In ValidationValues
                If LCase(MatchCell.Value) Like "*" & LCase(Target.Value) & "*" Then
                    results.Add MatchCell.Value
                End If
            Next MatchCell
        End If
        
        If results.Count > 0 Then
            tempList = ""
            For Each MatchCell In results
                tempList = tempList & MatchCell & ","
            Next MatchCell
            tempList = Left(tempList, Len(tempList) - 1)
            
            With Target.Validation
                .Delete
            End With
            ws.Range(Target.Address).Validation.Add Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, Formula1:=tempList
        Else
            On Error Resume Next
            Target.Validation.Delete
            On Error GoTo 0
        End If
    End If

exitHandler:
    Application.EnableEvents = True
End Sub
"""
    root.clipboard_clear()
    root.clipboard_append(vba_code)
    root.update()  # keep it on the clipboard
    messagebox.showinfo("Copied!", "VBA code copied to clipboard!")


# Style setup for ttk buttons
style = ttk.Style(root)
style.configure("TButton",
                font=button_font,
                padding=8,
                relief="flat",
                background="white",
                foreground="black")
style.map("TButton",
          background=[("active", "white")],
          foreground=[("active", "black")])

# Title and instruction
title_label = tk.Label(root, text="Excel to Tally XML Converter", font=title_font, fg="white", bg='black')
title_label.pack(pady=20)

label = tk.Label(root, text="Upload Bank Statement (Excel)", font=label_font, fg="#edf2f4", bg="black")
label.pack(pady=10)

# Horizontal Buttons Frame
btn_frame = tk.Frame(root, bg='black')
btn_frame.pack(pady=20)

sample_btn = ttk.Button(btn_frame, text="XML Converter", command=download_sample_file)
sample_btn.grid(row=0, column=0, padx=10, pady=5)

sales_upload_btn = ttk.Button(btn_frame, text="Sales XML File", command=upload_Sales_file)
sales_upload_btn.grid(row=0, column=1, padx=10, pady=5)

sales_save_btn = ttk.Button(btn_frame, text="Save Sales XML", command=save_Sales_file)
sales_save_btn.grid(row=0, column=2, padx=10, pady=5)

purchase_upload_btn = ttk.Button(btn_frame, text="Purchase XML File", command=upload_Purchase_file)
purchase_upload_btn.grid(row=0, column=3, padx=10, pady=5)

purchase_save_btn = ttk.Button(btn_frame, text="Save Purchase XML", command=save_Purchase_file)
purchase_save_btn.grid(row=0, column=4, padx=10, pady=5)

voucher_master_btn = ttk.Button(btn_frame, text="Voucher & Master File", command=upload_file)
voucher_master_btn.grid(row=0, column=5, padx=10, pady=5)

download_btn = ttk.Button(btn_frame, text="Download Sample  File", command=download_sample_file)
download_btn.grid(row=0, column=6, padx=10, pady=5)

copy_btn = ttk.Button(btn_frame, text="Copy Search Code", command=Search_code_clipboard)
copy_btn.grid(row=1, column=6, padx=10, pady=5)

# Frame with rounded look
frame = tk.Frame(root, bg="black", relief="raised", bd=4)
frame.pack(pady=20, padx=10, fill="both", expand=True)

# Voucher Section
label1 = tk.Label(frame, text="Voucher XML", font=label_font, bg="black", fg="white")
label1.grid(row=0, column=0, padx=5, pady=(5,0))

save_btn1 = ttk.Button(frame, text="Save Voucher XML", command=save_voucher_file)
save_btn1.grid(row=1, column=0, pady=5)

text_box1 = scrolledtext.ScrolledText(frame, wrap=tk.WORD, width=100, height=30, font=text_font, bg="#f8f9fa")
text_box1.grid(row=2, column=0, padx=5, pady=5)

# Master Section
label2 = tk.Label(frame, text="Master XML", font=label_font, bg="black", fg="white")
label2.grid(row=0, column=1, padx=5, pady=(5,0))

save_btn2 = ttk.Button(frame, text="Save Master XML", command=save_master_file)
save_btn2.grid(row=1, column=1, pady=5)

text_box2 = scrolledtext.ScrolledText(frame, wrap=tk.WORD, width=100, height=30, font=text_font, bg="#f8f9fa")
text_box2.grid(row=2, column=1, padx=5, pady=5)

# Grid weights
frame.columnconfigure(0, weight=1)
frame.columnconfigure(1, weight=1)
frame.rowconfigure(2, weight=1)

root.mainloop()

# text_box.config(state=tk.NORMAL)        # Enable editing temporarily
# text_box.delete("1.0", tk.END)          # Clear previous content
# text_box.insert(tk.END, xml_content)    # Insert new content
# print(xml_content)
# text_box.config(state=tk.DISABLED)
# def copy_text(event=None):
#     root.clipboard_clear()
#     root.clipboard_append(text_box.get("sel.first", "sel.last"))
            # if Credit > 0.0:
            #     if party.lower() != 'cash':
            #         vch_typ = 'Receipt'
            #     xml_data.extend
            # elif Debit > 0.0:
            #     if party.lower() != 'cash':
            #         vch_typ = 'Payment'
            #     # messagebox.showinfo('enter in debit')
            #     xml_data.extend
            # elif Credit >0.0:
            #     if party.lower() == 'cash':
            #         vch_typ = 'Receipt'
            #     xml_data.extend
            # elif Debit>0.0:
            #     if party.lower() == 'cash':
            #         vch_typ = 'Contra'
            #     # messagebox.showinfo('enter in debit')
            #     xml_data.extend
            #     party_lower = party.strip().lower()

# text_box.bind("<Control-c>", copy_text)

        #     if Bank_name:
        #         xml_master_data.extend([
        #     f'            <BANKDETAILS.LIST>',
        #     f'                <BANKNAME>{Bank_name}</BANKNAME>',
        #     f'                <ACCOUNTNUMBER>{Account_no}</ACCOUNTNUMBER>',
        #     f'                <IFSCODE>{Ifsc_code}</IFSCODE>',
        #     f'            </BANKDETAILS.LIST>',
        #  ])