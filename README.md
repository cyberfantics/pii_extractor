# Advanced PII Extractor

Welcome to the Advanced PII Extractor! This Python script scans various file formats for Personally Identifiable Information (PII) such as email addresses, phone numbers, Social Security Numbers (SSNs), credit card numbers, and IP addresses.

## Author
**Syed Mansoor ul Hassan Bukhari**  
[GitHub Profile](https://github.com/cyberfantics)  </br>
[LinkedIn](https://www.linkedin.com/in/mansoor-bukhari-77549a264/)

## Repository
[GitHub Repository](https://github.com/cyberfantics/pii_extractor.git)

## Description
The `advanced_pii_extractor.py` script parses and scans the following file types for PII:
- `.docx` files
- `.txt`, `.doc`, `.csv`, `.log`, and `.html` files
- `.xlsx` files
- `.pdf` files

## Features
- Identifies and extracts email addresses, phone numbers, SSNs, credit card numbers, and IP addresses.
- Saves matches to separate files (`email_matches.txt`, `phone_matches.txt`, etc.).
- Supports parsing from ZIP archives for `.docx` files and text extraction from PDFs.

## Usage
1. Ensure all dependencies are installed:
   ```bash
   pip install openpyxl PyPDF2
   ```
2. Download And Run:
   ```bash
   git clone https://github.com/cyberfantics/pii_extractor.git
   cd pii_extractor.git
   python pii_extractor.py
   ```

## Example Output

**Matches are saved in respective files:**
  1. email_matches.txt
  2. phone_matches.txt
  3. ssn_matches.txt
  4. credit_card_matches.txt
  5. ip_matches.txt
Each file contains the file path and matched PII entries.

## License
This project is licensed under the MIT License.
