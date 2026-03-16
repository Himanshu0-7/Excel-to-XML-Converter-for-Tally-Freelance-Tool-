# Excel to XML Converter for Tally

A Python-based tool that converts Excel bank statement data into XML format compatible with Tally for easy import. This tool helps automate the process of entering bank transactions in Tally by transforming structured Excel data into Tally-compatible XML vouchers.

## Features

* Upload Excel bank statement files
* Convert Excel transaction data to Tally-compatible XML
* Simple and interactive UI built with Streamlit
* Automatic data parsing using Pandas
* Batch processing of multiple transactions
* Reduces manual accounting data entry

## Tech Stack

* Python
* Streamlit
* Pandas
* XML Processing (ElementTree)

## Installation

1. Clone the repository

```bash
git clone https://github.com/your-username/tally-excel-to-xml-converter.git
cd tally-excel-to-xml-converter
```

2. Install dependencies

```bash
pip install -r requirements.txt
```

3. Run the application

```bash
streamlit run app.py
```

## How It Works

1. Upload the Excel bank statement file.
2. The application reads the Excel data using Pandas.
3. Data is processed and mapped to Tally XML voucher structure.
4. XML file is generated and ready for import into Tally.

## Use Case

Some versions of Tally do not support direct bank statement imports. This tool converts Excel-based transaction data into the XML format required by Tally, allowing accountants to import transactions automatically instead of entering them manually.

## Example Workflow

Excel Bank Statement → Python Processing → XML Generation → Import into Tally

## License

This project is available under the MIT License.
