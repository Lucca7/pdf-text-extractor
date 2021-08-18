# Description: this program extracts text from a PDF file and dumps it into a newly created DOCX file.
# It is done character per character and maintains bold and/or italic formatting.
#
# How to use: provide the full path of the directory where the PDF file is located and
# the full name of the PDF file in the respective fields of the user interface, then click "Extract".
# A new DOCX file will be created in the same directory with the same name.
#
# Last updated: 2021-08-17
#
# Author: Lucca Gonçalves

from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LTChar
from docx import Document
from PySimpleGUI import PySimpleGUI as sg
import os.path

# # The code belows formats strings to make sure they only contain valid xml characters
# def valid_xml_char_ordinal(c):
#     codepoint = ord(c)
#     # Conditions ordered by presumed frequency
#     return (
#         0x20 <= codepoint <= 0xD7FF or
#         codepoint in (0x9, 0xA, 0xD) or
#         0xE000 <= codepoint <= 0xFFFD or
#         0x10000 <= codepoint <= 0x10FFFF
#         )

# The code below extracts text from a PDF file and writes it to a DOCX file
# It is done character by character and it mantains bold and/or italic formatting
def extract_text_to_document(document, complete_path):
    for page_layout in extract_pages(complete_path):
        for element in page_layout:
            if isinstance(element, LTTextContainer):
                prg = document.add_paragraph('')
                for text_line in element:
                    for character in text_line:
                        if isinstance(character, LTChar):
                            run = prg.add_run(character.get_text())
                            if "bold" in character.fontname:
                                run.bold = True
                            if "Bold" in character.fontname:
                                run.bold = True
                            if "italic" in character.fontname:
                                run.italic = True
                            if "Italic" in character.fontname:
                                run.italic = True
                            #print(character.fontname)
                            #print(character.size)
    return

# Setting the layout of the GUI
sg.theme("Reddit")
layout = [
    [sg.Text("Path:"), sg.Input(key="path")],
    [sg.Text("File:"), sg.Input(key="file")],
    [sg.Button("Extract")]
]

# Generating the GUI
window = sg.Window("Extract text from PDF to DOCX", layout)

# Reading the events that happen in the GUI
while True:

    events, values = window.read()

    # In case user closes windows
    if events == sg.WINDOW_CLOSED:
        break

    # Start text extraction after user has given the input and clicked "Extract"
    if events == "Extract":

        pdf_path = values['path']
        pdf_file = values['file']
        complete_path = os.path.join(pdf_path, pdf_file)

        document = Document()
        extract_text_to_document(document, complete_path)
        document.save(complete_path.replace("pdf", "docx"))
        break