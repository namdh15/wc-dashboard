import re
import pandas as pd
from docx import Document
from tqdm import tqdm
import zipfile
from lxml import etree

from utils import read_docx_full_text

# docx_file = "TTTD-KH-SME_Template.docx"
docx_file = "Ly lich khoa hoc (HVCH).doc"
excel_file = "Tonghop.xlsx"

# regex lấy mã SME
# pattern = re.compile(r"SME(?:\.[A-Z0-9]+)+")
# pattern = re.compile(r"SME(?:\.[A-Z0-9]*\d)+")
pattern = re.compile(r"SME(?:\.[A-Z0-9]+)+")

# --------- lấy mã từ docx ----------
full_text = read_docx_full_text(docx_file)

print(f"Full text length: {len(full_text)} characters")
# with open("full_text_debug_func.txt", "w", encoding="utf-8") as f:
#     f.write(full_text)

# docx_codes = pattern.findall(full_text)
# print(f'Total matches found: {len(docx_codes)}')
# docx_codes = set(docx_codes)
# print(f"\nTổng mã tìm thấy trong DOCX: {len(docx_codes)}")
# # print("Mã trong DOCX:")
# for code in docx_codes:
#     print(code)

### DEBUG
# for i, element in enumerate(doc.element.body.iter()):
#     if element.text:
#         text = element.text

#         matches = pattern.findall(text)
#         regex_count = len(matches)

#         sme_count = text.count("SME")

#         if regex_count != sme_count:
#             print("\n--- Debug Info ---")
#             print(f"\nElement index: {i} - Text: \n{text}")
#             print(f"SME count in text: {sme_count}")
#             print(f"Regex matches: {regex_count}")
#             print("Matches:", matches)
#             print("Text:", text)
####


# # --------- lấy mã từ excel ----------
# df = pd.read_excel(excel_file)

# # giả sử cột chứa mã tên là CODE
# excel_codes = set(df["CODE"].astype(str))

# # --------- so sánh ----------
# missing_in_excel = docx_codes - excel_codes
# unused_in_docx = excel_codes - docx_codes

# print("Mã có trong DOCX nhưng thiếu trong Excel:")
# print(missing_in_excel)

# print("\nMã có trong Excel nhưng không dùng trong DOCX:")
# print(unused_in_docx)
