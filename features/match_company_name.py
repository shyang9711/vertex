import fitz  # PyMuPDF
import os
import re
from colorama import init, Fore, Style
from tkinter import Tk, filedialog

init(autoreset=True)

CORPORATE_SUFFIXES = {
    "CORP": "CORPORATION",
    "INC": "INCORPORATED",
    "CO": "COMPANY",
    "LTD": "LIMITED",
    "PROF": "PROFESSIONAL",
}

def normalize_company_name(name):
    if not name:
        return ""
    name = name.upper()
    name = re.sub(r"[^\w\s]", "", name)  # Remove punctuation
    words = name.split()
    normalized = [CORPORATE_SUFFIXES.get(word, word) for word in words]
    return " ".join(normalized)

def extract_company_name(pdf_path, file_type):
    if not os.path.exists(pdf_path):
        print(Fore.RED + f"File not found: {pdf_path}")
        return None, None

    try:
        with fitz.open(pdf_path) as doc:
            if file_type == "paycheck":
                for page in doc:
                    lines = page.get_text().splitlines()
                    for i, line in enumerate(lines):
                        if "MEMO: Pay Period" in line:
                            name = lines[i + 2].strip() if i + 2 < len(lines) else ""
                            dba = lines[i + 4].strip() if (i + 4 < len(lines) and "DBA" in lines[i + 4].strip())else ""
                            return name, dba

            elif file_type == "payroll":
                lines = doc[0].get_text().splitlines()
                first_line = lines[0]
                for line in lines[:10]:
                    if "INC" in line or "LLC" in line or "CORP" in line:
                        return line.strip(), None
                    else:
                        return first_line, None
    except Exception as e:
        print(Fore.RED + f"Error reading {pdf_path}: {e}")

    return None, None

def ask_for_file(prompt_title):
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title=prompt_title, filetypes=[("PDF files", "*.pdf")])
    root.destroy()
    return file_path

def compare_corporate_names():
    print(Style.BRIGHT + "\nPlease select the PDF files.\n")

    paycheck_pdf = ask_for_file("Select the PAYCHECK PDF file")
    if not paycheck_pdf:
        print(Fore.RED + "No paycheck file selected.")
        return

    payroll_pdf = ask_for_file("Select the PAYROLL DETAILS PDF file")
    if not payroll_pdf:
        print(Fore.RED + "No payroll details file selected.")
        return

    print(Style.BRIGHT + f"\nChecking corporate names between:\n- {paycheck_pdf}\n- {payroll_pdf}\n")

    paycheck_name, paycheck_dba = extract_company_name(paycheck_pdf, "paycheck")
    payroll_name, _ = extract_company_name(payroll_pdf, "payroll")

    # Show extracted results
    if paycheck_name:
        print(Fore.CYAN + f"[Paycheck PDF] Name: {paycheck_name}")
    if paycheck_dba:
        print(Fore.CYAN + f"[Paycheck PDF] DBA : {paycheck_dba}")
    if payroll_name:
        print(Fore.CYAN + f"[Payroll PDF]  Name: {payroll_name}")

    # Normalize for comparison
    norm_paycheck_name = normalize_company_name(paycheck_name)
    norm_paycheck_dba = normalize_company_name(paycheck_dba)
    norm_payroll_name = normalize_company_name(payroll_name)

    match_found = False

    # Exact name match
    if norm_paycheck_name == norm_payroll_name:
        print(Fore.GREEN + Style.BRIGHT + "\n✅ Corporate names MATCH (normalized).")
        match_found = True

    # DBA match
    elif norm_paycheck_dba and norm_paycheck_dba == norm_payroll_name:
        print(Fore.GREEN + Style.BRIGHT + "\n✅ Corporate DBA MATCHES Payroll Company (normalized).")
        match_found = True

    if not match_found:
        print(Fore.RED + Style.BRIGHT + "\n❌ Corporate names DO NOT match.")
        print(Fore.MAGENTA + f"Normalized Paycheck Name: {norm_paycheck_name}")
        print(Fore.MAGENTA + f"Normalized Paycheck DBA : {norm_paycheck_dba}")
        print(Fore.MAGENTA + f"Normalized Payroll Name : {norm_payroll_name}")

if __name__ == "__main__":
    compare_corporate_names()
