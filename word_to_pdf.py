import sys
import os
import comtypes.client
import time
import shutil


def docx_to_pdf(in_file, out_file):
    """Convert .docs to .pdf"""
    pdf_key_format = 17
    file_in = os.path.abspath(in_file)
    file_out = os.path.abspath(out_file)
    worddoc = comtypes.client.CreateObject('Word.Application')
    time.sleep(2)
    doc = worddoc.Documents.Open(file_in)
    doc.SaveAs(file_out, FileFormat=pdf_key_format)
    doc.Close()
    worddoc.Quit


def move_pdf_to_folder():
    """move the pdf files in the pdf folder"""
    destination_file = os.getcwd() + "\pdf_files"
    for file in os.listdir():
        if file.endswith(".pdf"):
            try:
                shutil.move(file, destination_file)
            except shutil.Error:
                pass
    for file in os.listdir():
        if file.endswith('.pdf'):
            os.unlink(file)
