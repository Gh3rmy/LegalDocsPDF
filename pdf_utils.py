# pdf_utils.py
import fitz

def remove_selected_pages(input_pdf, output_pdf, pages_to_remove):
    doc = fitz.open(input_pdf)
    pages_to_keep = [i for i in range(len(doc)) if i not in pages_to_remove]
    new_doc = fitz.open()
    for i in pages_to_keep:
        new_doc.insert_pdf(doc, from_page=i, to_page=i)
    new_doc.save(output_pdf)
    new_doc.close()
    doc.close()