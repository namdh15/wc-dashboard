import zipfile
import re
from lxml import etree
from docx import Document
from tqdm import tqdm

def read_docx_full_text(docx_path):
    texts = []

    # with zipfile.ZipFile(docx_path) as z:
    #     for name in z.namelist():
    #         if not name.endswith(".xml"):
    #             continue

    #         xml = z.read(name)
    #         root = etree.fromstring(xml)

    #         for node in root.xpath("//w:t", namespaces={
    #             "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    #         }):
    #             if node.text:
    #                 texts.append(node.text)

    # full_text = " ".join(texts)

    # # -------- normalize ----------
    # full_text = re.sub(r"\s+", " ", full_text)

    # # fix: S ME -> SME
    # full_text = re.sub(r"S\s+ME", "SME", full_text)

    # # fix khoảng trắng quanh dấu .
    # full_text = re.sub(r"\s*\.\s*", ".", full_text)

    # # fix khoảng trắng giữa số
    # full_text = re.sub(r"(\d)\s+(\d)", r"\1\2", full_text)

    # return full_text


    doc = Document(docx_path)
    # full_text = " ".join([element.text for element in doc.element.body.iter() if element.text]).replace("SME", "\nSME")
    # texts = full_text.splitlines("\n")
    texts = [element.text for element in doc.element.body.iter() if element.text]

    previous_text = ""
    full_text = ""
    for text in tqdm(texts, desc="Processing DOCX"):
        if previous_text:
            if previous_text in text:
                previous_text = text
            else:
                full_text += previous_text + " "
                previous_text = text
        else:
            previous_text = text
    else:
        if previous_text:
            full_text += previous_text + " "
    # full_text = full_text.replace("SME", "\nSME")
    # # Write full text to a file for debugging
    # with open("full_text_debug_duplicate.txt", "w", encoding="utf-8") as f:
    #     f.write(full_text)

    return full_text