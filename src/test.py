import win32com.client as win32
from win32com.client import constants
import pyautogui
import os 

word = win32.Dispatch("Word.Application")
# word.Visible = True
# doc = word.ActiveDocument
# print(doc.Name)
# selection = word.Selection

# def export_to_pdf(doc, output_path):
#     """
#     Export the given Word document to a PDF file.

#     :param doc: The Word document to export.
#     :param output_path: The path where the PDF file will be saved.
#     """
#     doc.SaveAs(output_path, FileFormat=constants.wdFormatPDF)

# # Example usage
# output_pdf_path = os.path.join(os.getcwd(), "output.pdf")

# export_to_pdf(doc, output_pdf_path)

# Rest of the code...
# a = constants.__dict__

# selection.GoTo(What=constants.wdGoToPage, Which=constants.wdGoToAbsolute, Count=2)
# first_page_number = selection.Information(constants.wdActiveEndAdjustedPageNumber)

# print(type(a))

# OddHeaderRange = doc.Sections(1).Headers(constants.wdHeaderFooterPrimary).Range
# EvenHeaderRange = doc.Sections(1).Headers(constants.wdHeaderFooterEvenPages).Range
# doc.Sections(1).Headers(constants.wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
# doc.Sections(1).Headers(1).PageNumbers.StartingNumber = 1

# wdGoToPage = 1  # GoTo page
# wdGoToLast = 2  # GoTo the last instance

# # Navigate to the last page
# selection.GoTo(What=wdGoToPage, Which=wdGoToLast)
# last_page_number = selection.Information(constants.wdActiveEndAdjustedPageNumber)
# print(last_page_number)
# # doc.Sections(1).Headers(1).PageNumbers.Add(2, True)

# # print(OddHeaderRange.Text)
# # print(EvenHeaderRange.Text)

# p = ['Unit 0 \t1 - 30\r', 'Unit 1\t\t31 - 41\r', 'unit 2\t42 - 54\r', 'unit 3\t55 - 61\r', 'Unit 4\t62 - 73\r', 'unit 5\t74 - 91\r', 'unit 6\t91 - 102\r', 'Unit 7\t102 - 121\r', 'Unit 8\t122 - 140\r', 'Unit 9\t141 - 159\r', 'Unit 10\t160 - 185\r', 'unit 11\t186 - 196\r', 'unit 12\t197 - 205\r', 'unit 13\t206 - 217\r', 'Unit 14\t218 - 230\r', 'unit 15\t231 - 242\r', 'unit 16\t243 - 249\r', 'Unit 17\t250 - 258\r', 'Unit 18 \t259 - 270\r', 'Unit 19\t271 - 277\r', '\r', '\r']
# text = "unit"
# def find_and_replace_font(text, word=word):
#     find = word.Selection.Find
#     find.ClearFormatting()
#     find.Replacement.ClearFormatting()
#     find.Text = text
#     find.Replacement.Font.Size = 8.5
#     find.Replacement.Text = ""
#     find.Forward = True
#     find.Wrap = 1  # wdFindContinue
#     find.Format = True
#     find.MatchCase = False
#     find.MatchWholeWord = False
#             # Replace all instances
#     while find.Execute:
#         pass 

# # # find_and_replace_font(fnd)

# # # List all the styles in doc 
# # styles = doc.Styles
# # a = [a for a in styles]
# # print(a)
# active_window = word.ActiveWindow

# wdPaneNone = 0
# wdNormalView = 1
# wdOutlineView = 2
# wdPrintView = 3
# wdSeekCurrentPageHeader = 9

# # Check if the window is split and close the second pane if so
# if active_window.View.SplitSpecial != wdPaneNone:
#     active_window.Panes(2).Close()

# # Change view type if it's Normal or Outline to Print Layout view
# if (active_window.ActivePane.View.Type == wdNormalView or
#         active_window.ActivePane.View.Type == wdOutlineView):
#     active_window.ActivePane.View.Type = wdPrintView

# pyautogui.hotkey('ctrl', 'home')
# # Set the SeekView to Current Page Header
# active_window.ActivePane.View.SeekView = wdSeekCurrentPageHeader
# active_window.ActivePane.View.SeekView = 0
# selection = word.Selection

# # active_window.ActivePane.View.SeekView = 0
# pyautogui.press('esc')
# pyautogui.press('esc')

# p = []

# for i in range(doc.Paragraphs.Count):
#     print(doc.Paragraphs(i+1).Range.Text)

#     p.append(doc.Paragraphs(i+1).Range.Text)

# print(p)
from spire.doc import *
from spire.doc.common import *

# Create word document
document = Document()

document.LoadFromFile(r"D:\03 Solutions- Kamal\+2 Solution File\NEB Solution Health and Physical Education 12\Solution to Model and Exam Questions (Health - XII).doc")
# Load a doc or docx file

output_path = os.path.join(os.getcwd(), "output/ToPDF.pdf")
document.SaveToFile(output_path, FileFormat.PDF)
document.Close()
