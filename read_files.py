import pdfplumber
import pandas as pd
import re
from openpyxl import load_workbook
import arabic_reshaper
from bidi.algorithm import get_display

# new structure
def read_pdf(pdf_path: str) -> list:
    """read all context from pdf file.

    arguments:
        pdf_file (str) : path of the file.

    Returns:
        type: return a list content all context.
    """
    # extracing contexts from pdf_file
    context = []
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            text = page.extract_text()
            if text:
                context.extend(text.splitlines())
    return context

def convert_to_excel(date: dict, output_file: str) -> None:
    """convert all data to excel.
    
    Arguments:
        data (dict) : data as dict {key: value}
        output_file (str) : file name to save data

    Returns:
        type: save file 
    """

    df = pd.DataFrame(date, index=False)
    df.to_excel(output_file)
    print("Data is saved ...")
    return
     
def filters(text: str) -> None:
    """filter all text

    Arguments:
        text (str) : text to find the match

    Returns:
        type: return filtered text

    """
    # Values
    contract_no = None
    tenancy_start_date = None
    tenancy_end_date = None
    company_name = None  # r"Company+\s+Name+\/+Founder)+(\s+.)"
    natinal_address = None
    due_date = None
    end_of_payments = None
    amount = None # r"^\d+\.\d+\s+\d{4}-\d{2}-\d{2}\s+\d{4}-\d{2}-\d{2}.*\d{4}-\d{2}-\d{2}\s+\d{4}-\d{2}-\d{2}\s+\d+\s*$"
    # extracting pattern

    return 
    
# import PyPDF2
# import os

# def read_pdf_lines(pdf_path):
#     lines = []
#     with open(pdf_path, 'rb') as file:
#         reader = PyPDF2.PdfReader(file)
#         for page in reader.pages:
#             text = page.extract_text()
#             if text:
#                 lines.extend(text.splitlines())
#     return lines

# # Example usage:
# pdf_file = r'C://Users//amric//Desktop//Git//Discord//dictionary.pdf'  # Replace with your PDF file path
# for line in read_pdf_lines(pdf_file):
#     print(line)

# import pdfplumber
# import pandas as pd
# import re
# from openpyxl import load_workbook
# import arabic_reshaper
# from bidi.algorithm import get_display

# # دالة لترتيب النص العربي
# def reshape_arabic(text):
#     try:
#         reshaped = arabic_reshaper.reshape(text)
#         return get_display(reshaped)
#     except:
#         return text

# # استخراج آخر 3 أسطر من جدول الدفعات
# def extract_last_3_payment_lines(text):
#     # النمط: يبدأ بمبلغ ثم تواريخ ثم رقم في النهاية
#     pattern = r"^\d+\.\d+\s+\d{4}-\d{2}-\d{2}\s+\d{4}-\d{2}-\d{2}.*\d{4}-\d{2}-\d{2}\s+\d{4}-\d{2}-\d{2}\s+\d+\s*$"
#     lines = text.splitlines()
#     payment_lines = [line for line in lines if re.match(pattern, line)]
#     return payment_lines[-3:] if len(payment_lines) >= 3 else payment_lines


# def read_pdf_tenant_info(pdf_path):
#     records = []

#     with pdfplumber.open(pdf_path) as pdf:
#         for page_number, page in enumerate(pdf.pages):
#             text = page.extract_text()
#             if not text:
#                 print(f"⚠️ Page {page_number+1} has no text.")
#                 continue

#             print(f"📄 Page {page_number+1} text found.")
        
#             row = {}

#             # Contract No
#             match_contract = re.search(r"Contract\s*No\.?\s*[:\-]?\s*(.+)", text, re.IGNORECASE)
#             if match_contract:
#                 row["Contract No"] = reshape_arabic(match_contract.group(1).strip())
#                 print(f"🔹 Contract No: {row['Contract No']}")

#             # Tenancy Start Date
#             match_start = re.search(r"Tenancy\s*Start\s*Date\s*[:\-]?\s*([\d\/\-]+)", text, re.IGNORECASE)
#             if match_start:
#                 row["Tenancy Start Date"] = match_start.group(1).strip()
#                 print(f"🔹 Tenancy Start Date: {row['Tenancy Start Date']}")

#             # Tenancy End Date
#             match_end = re.search(r"Tenancy\s*End\s*Date\s*[:\-]?\s*([\d\/\-]+)", text, re.IGNORECASE)
#             if match_end:
#                 row["Tenancy End Date"] = match_end.group(1).strip()
#                 print(f"🔹 Tenancy End Date: {row['Tenancy End Date']}")

#             # Company name / Founder
#             match_tenant = re.search(r"Company\s*name\s*/\s*Founder\s*[:\-]?\s*(.+)", text, re.IGNORECASE)
#             if match_tenant:
#                 row["Company name/Founder"] = reshape_arabic(match_tenant.group(1).strip())
#                 print(f"🔹 Company name/Founder: {row['Company name/Founder']}")

#             # National Address
#             match_address = re.search(r"National\s*Address\s*[:\-]?\s*(.+)", text, re.IGNORECASE)
#             if match_address:
#                 row["National Address"] = reshape_arabic(match_address.group(1).strip())
#                 print(f"🔹 National Address: {row['National Address']}")
            
#             match_address = re.search(r"National\s*Address\s*[:\-]?\s*(.+)", text, re.IGNORECASE)
#             if match_address:
#                 row["National Address"] = reshape_arabic(match_address.group(1).strip())
#                 print(f"🔹 National Address: {row['National Address']}")
#             match_payment = re.search(r"Payment\s*Method\s*[:\-]?\s*(.+)", text, re.IGNORECASE)

#             # Payments - آخر 3 دفعات
#             if match_payment:
#                 last_3_payments = extract_last_3_payment_lines(text)
#                 row["Payment"] = {k:v for k,v in enumerate(last_3_payments, start=1)}
#                 print(f"🔹 Last Payments: {row['Payment'][1]}")
#                 print(f"🔹 Last Payments: {row['Payment'][2]}")
#                 print(f"🔹 Last Payments: {row['Payment'][3]}")

#             else:
#                 row["National Address"] = ""  # إذا ما وجدنا العنوان

#             if row:
#                 records.append(row)

#     return records

# def write(output_data, excel_file):
#     if output_data:
#         df = pd.DataFrame(output_data, columns=[
#             "Contract No", "Tenancy Start Date", "Tenancy End Date", "Company name/Founder", "National Address", "Last Payments"
#         ])
#         df.to_excel(excel_file, index=False)

#         # توسيع الأعمدة تلقائياً
#         wb = load_workbook(excel_file)
#         ws = wb.active
#         for col in ws.columns:
#             max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
#             ws.column_dimensions[col[0].column_letter].width = max_length + 5
#         wb.save(excel_file)

#         print(f"✅ تم استخراج البيانات وحفظها في {excel_file}")
#     else:
#         print("❌ لم يتم العثور على بيانات.")


# # ======== المسارات ========
# # pdf_file = r"C:\Users\ream8\Desktop\Project\10988496532.pdf"
# # excel_file = r"C:\Users\ream8\Desktop\Project\10988496532_output.xlsx"
# pdf_file = r"C:\Users\amric\Desktop\Git\Discord\10988496532.pdf"
# excel_file = r"C:\Users\amric\Desktop\Git\Discord\10988496532_output.xlsx"
# # ======== التنفيذ ========
# output_data = read_pdf_tenant_info(pdf_file)
# # write(output_data, excel_file)

import pdfplumber
import pandas as pd
import re
from openpyxl import load_workbook
import arabic_reshaper
from bidi.algorithm import get_display

def reshape_arabic(text):
    if not text:
        return ""
    try:
        reshaped = arabic_reshaper.reshape(text)
        return get_display(reshaped)
    except Exception:
        return text

def read_pdf_tenant_info(pdf_path):
    records = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages):
            text = page.extract_text()
            if not text:
                continue

            text = reshape_arabic(text)
            row = {
                "Contract No": "",
                "Tenancy Start Date": "",
                "Tenancy End Date": "",
                "Company name/Founder": "",
                "National Address": "",
                "Due Date(AD)": "",
                "End of payment deadline(AD)": "",
                "Amount": ""
            }

            # Contract No → قراءة الرقم الكامل فقط
            match_contract = re.search(r"Contract\s*No\.?\s*[:\-]?\s*([\d\s\/\-]+)", text, re.IGNORECASE)
            if match_contract:
                row["Contract No"] = match_contract.group(1).strip()

            # Tenancy Start Date
            match_start = re.search(
                r"Tenancy\s*Start\s*Date\s*[:\-]?\s*([\d]{2,4}[\/\-][\d]{1,2}[\/\-][\d]{2,4})",
                text, re.IGNORECASE
            )
            if match_start:
                row["Tenancy Start Date"] = match_start.group(1).strip()

            # Tenancy End Date
            match_end = re.search(
                r"Tenancy\s*End\s*Date\s*[:\-]?\s*([\d]{2,4}[\/\-][\d]{1,2}[\/\-][\d]{2,4})",
                text, re.IGNORECASE
            )
            if match_end:
                row["Tenancy End Date"] = match_end.group(1).strip()

            # Company name/Founder
            match_tenant = re.search(
                r"Company\s*[\r\n]+\s*name/Founder\s*[:\-]?\s*(.+)",
                text, re.IGNORECASE | re.DOTALL
            )
            if match_tenant:
                tenant_text = match_tenant.group(1).strip()
                tenant_lines = tenant_text.split("\n")[:3]
                tenant_cleaned = " ".join(line.strip() for line in tenant_lines if line.strip())
                row["Company name/Founder"] = tenant_cleaned

            # National Address → فقط السطر الأول بدون نص إضافي
            match_address = re.search(r"National\s*Address\s*[:\-]?\s*(.+)", text, re.IGNORECASE | re.DOTALL)
            if match_address:
                address_text = match_address.group(1).strip()
                row["National Address"] = address_text.split("\n")[0].strip()

            # قراءة جدول الدفعات من الصفحة الثالثة
            if page_number == 2:
                tables = page.extract_tables()
                if len(tables) >= 2:
                    payment_table = tables[1]
                    headers = [reshape_arabic(str(h)) for h in payment_table[0]]

                    # إيجاد مواقع الأعمدة المطلوبة
                    due_date_idx = next((i for i, h in enumerate(headers) if "Due Date(AD)" in h), None)
                    end_deadline_idx = next((i for i, h in enumerate(headers) if "End of payment deadline(AD)" in h), None)
                    amount_idx = next((i for i, h in enumerate(headers) if "Amount" in h), None)

                    if due_date_idx is not None and end_deadline_idx is not None and amount_idx is not None:
                        for payment_row in payment_table[1:]:
                            payment_row = [reshape_arabic(str(cell)) if cell else "" for cell in payment_row]
                            row_copy = row.copy()
                            row_copy["Due Date(AD)"] = payment_row[due_date_idx]
                            row_copy["End of payment deadline(AD)"] = payment_row[end_deadline_idx]
                            row_copy["Amount"] = payment_row[amount_idx]
                            records.append(row_copy)
                continue


            if any(value for value in row.values()):
                records.append(row)

        return records


pdf_file = r"C:\Users\amric\Desktop\Git\Discord\10988496532.pdf"
excel_file = r"C:\Users\amric\Desktop\Git\Discord\10988496532_output.xlsx"

output_data = read_pdf_tenant_info(pdf_file)

if output_data:
    df = pd.DataFrame(output_data, columns=[
        "Contract No", "Tenancy Start Date", "Tenancy End Date",
        "Company name/Founder", "National Address",
        "Due Date(AD)", "End of payment deadline(AD)", "Amount"
    ])
    df.to_excel(excel_file, index=False)

    wb = load_workbook(excel_file)
    ws = wb.active
    for col in ws.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_length + 5
    wb.save(excel_file)

    print(f"✅ تم استخراج البيانات وحفظها في {excel_file}")
else:
    print("❌ لم يتم العثور على بيانات.")

# if __name__ == "__main__":
#     pdf_file = r"C:\Users\amric\Desktop\Git\Discord\10988496532.pdf"
#     excel_file = r"C:\Users\amric\Desktop\Git\Discord\10988496532_output.xlsx"
#     output_data = read_pdf_tenant_info(pdf_file)
#     write_to_excel(output_data, excel_file)
