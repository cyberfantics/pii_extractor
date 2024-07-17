# -*- coding: utf-8 -*-
"""
Created on Wed Jul 17 19:23:45 2024

@author: Mansoor
"""

import os
import re
from zipfile import ZipFile
import openpyxl
import PyPDF2

# Regular expressions for PII
email_regex = r'[a-z0-9]+[\._]?[a-z0-9]+[@]\w+[.]\w{2,3}'
phone_regex = r'(\d{3}-\d{3}-\d{4})'
ssn_regex = r'\d{3}-\d{2}-\d{4}'
cc_regex = r'(\d{4}-\d{4}-\d{4}-\d{4})'
ip_regex = r'(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})'
regexes = {
    'email': email_regex,
    'phone': phone_regex,
    'ssn': ssn_regex,
    'credit_card': cc_regex,
    'ip': ip_regex
}

# Function to find PII in data
def findPII(data):
    matches = {}
    for key, regex in regexes.items():
        found = re.findall(regex, data)
        if found:
            print(f"Found matches for {key}: {found}")
            if key not in matches:
                matches[key] = []
            matches[key].extend(found)
    return matches

# Function to print matches and save to specific files
def printMatches(filedir, matches):
    if matches:
        for key, values in matches.items():
            output_file_path = f"{key}_matches.txt"
            with open(output_file_path, "a") as output_file:
                output_file.write(f"{filedir}\n")
                for match in values:
                    output_file.write(f"{match}\n")
                print(f"Matches found in {filedir} for {key}: {values}")

# Function to parse .docx files
def parseDocx(root, docs):
    for doc in docs:
        filedir = os.path.join(root, doc)
        print(f"Parsing DOCX file: {filedir}")
        try:
            with ZipFile(filedir, "r") as zip:
                data = zip.read("word/document.xml").decode("utf-8")
                matches = findPII(data)
                printMatches(filedir, matches)
        except Exception as e:
            print(f"Error parsing DOCX file {filedir}: {e}")

# Function to parse text files
def parseText(root, txts):
    for txt in txts:
        filedir = os.path.join(root, txt)
        print(f"Parsing text file: {filedir}")
        try:
            with open(filedir, "r", encoding='utf-8', errors='ignore') as f:
                data = f.read()
                matches = findPII(data)
                printMatches(filedir, matches)
        except Exception as e:
            print(f"Error parsing text file {filedir}: {e}")

# Function to parse .xlsx files
def parseXlsx(root, xlsxs):
    for xlsx in xlsxs:
        filedir = os.path.join(root, xlsx)
        print(f"Parsing XLSX file: {filedir}")
        try:
            wb = openpyxl.load_workbook(filedir, read_only=True)
            for sheet in wb.worksheets:
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.value:
                            matches = findPII(str(cell.value))
                            printMatches(filedir, matches)
        except Exception as e:
            print(f"Error parsing XLSX file {filedir}: {e}")

# Function to parse .pdf files
def parsePdf(root, pdfs):
    for pdf in pdfs:
        filedir = os.path.join(root, pdf)
        print(f"Parsing PDF file: {filedir}")
        try:
            with open(filedir, "rb") as file:
                reader = PyPDF2.PdfFileReader(file)
                for page_num in range(reader.numPages):
                    page = reader.getPage(page_num)
                    data = page.extract_text()
                    if data:
                        matches = findPII(data)
                        printMatches(filedir, matches)
        except Exception as e:
            print(f"Error parsing PDF file {filedir}: {e}")

# Function to find files in directory
def findFiles(directory):
    txt_ext = [".txt", ".doc", ".csv", ".log", ".html"]
    # Clear previous content in the output files
    for key in regexes.keys():
        with open(f"{key}_matches.txt", "w") as output_file:
            pass
    for root, dirs, files in os.walk(directory):
        parseDocx(root, [f for f in files if f.endswith(".docx")])
        parseXlsx(root, [f for f in files if f.endswith(".xlsx")])
        parsePdf(root, [f for f in files if f.endswith(".pdf")])
        print(f'Searching in {root}')
        for ext in txt_ext:
            parseText(root, [f for f in files if f.endswith(ext)])
    print("Matches saved to respective files based on type")

# Directory to search - explicitly setting to Desktop
desktop_directory = os.path.join(os.environ['USERPROFILE'], 'Downloads')

# Execute the function
findFiles(desktop_directory)
