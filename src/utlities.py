import logging
import pandas as pd
import time
import os
import pyautogui
from src.shared_objects import WordApp
from win32com.client import constants
from src.logger import setup_logging
from tkinter import filedialog

logging = setup_logging()


def adjust_header(active_window, selection):
    """
    Adjusts the header of the active window in Word.

    This function is equivalent to the following manual steps:

    1. If the window is split, close the second pane.
    2. If the view is Normal or Outline, change it to Print Layout.
    3. Move the cursor to the beginning of the document.
    4. Set the SeekView to Current Page Header.
    5. Select the first paragraph.
    6. Set the font size to 1.
    7. Set the space after the paragraph to 0.
    8. Set the space before the paragraph to 0.
    9. Set the SeekView back to 0.

    :param active_window: The active window in Word.
    :param selection: The Selection object in Word.
    :return: None
    """

    logging.info("Adjusting header of the active window.")
    wdPaneNone = 0
    wdNormalView = 1
    wdOutlineView = 2
    wdPrintView = 3
    wdSeekCurrentPageHeader = 9

    # Check if the window is split and close the second pane if so
    if active_window.View.SplitSpecial != wdPaneNone:
        logging.info("Window is split. Closing the second pane.")
        active_window.Panes(2).Close()

    # Change view type if it's Normal or Outline to Print Layout view
    if (active_window.ActivePane.View.Type == wdNormalView or
            active_window.ActivePane.View.Type == wdOutlineView):
        logging.info("Changing view to Print Layout.")
        active_window.ActivePane.View.Type = wdPrintView

    pyautogui.hotkey('ctrl', 'home')
    # Set the SeekView to Current Page Header
    logging.info("Setting SeekView to Current Page Header.")
    active_window.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    selection.Paragraphs(1).Range.Select()
    selection.Range.Font.Size = 1
    selection.Range.ParagraphFormat.SpaceAfter = 0
    selection.Range.ParagraphFormat.SpaceBefore = 0
    active_window.ActivePane.View.SeekView = 0
    logging.info("Header adjustment completed.")
    return None

def makePdf(tree, dirr):
    # Get selected files from the Treeview
    selected_items = tree.selection()
    selected_files = [tree.item(item, 'values')[0] for item in selected_items]

    # Initialize Word Application
    word = WordApp()
    word.Visible = False  # Run in the background
    output_dir = filedialog.askdirectory(title="Select Output Directory")
    for file in selected_files:
        try:
            # Open the Word document
            filepath = os.path.join(dirr, file)
            doc = word.open_document(filepath)
            
            output_filename = os.path.join(output_dir, os.path.splitext(os.path.basename(filepath))[0] + ".pdf")
            doc.SaveAs(output_filename, FileFormat=17)  # 17 is the file format for PDF
            doc.Close()


            logging.info(f"Successfully converted: {file} to {output_dir}")
            doc.Close()
        except Exception as e:
            logging.error(f"Error converting {file} to PDF: {e}")

def rou(v):
    return round(v * 20) / 20

def collect_data(progress_callback=None):
    logging.info("Collecting data from the active Word document.")
    word = WordApp()
    doc = word.get_active_document()
    para_number = 0
    data = []
    ranges = []

    total_paragraphs = len(doc.Paragraphs)
    for i, para in enumerate(doc.Paragraphs):
        para_range = para.Range
        para_range_start = para_range.Start
        para_range_end = para_range.End
        para_number += 1
        para_text = para_range.Text.strip()
        Fl_Indent = round(para.FirstLineIndent / 28.35, 2)
        Hanging = round(para.LeftIndent / 28.35, 2)
        char_count = len(para_text)
        word_count = len(para_range.Words)
        tab_count = para_text.count("\t")
        eq_count = para_text.count("=")
        is_italic = "Yes" if para_range.Font.Italic else "No"
        is_bold = "Yes" if para_range.Font.Bold else "No"
        font_name = para_range.Font.Name
        para_style = para.Style.NameLocal
        font_size = para_range.Font.Size
        numbering_type = para.Range.ListFormat.ListType
        list_value = para.Range.ListFormat.ListValue if numbering_type != 0 else "0"
        Within_Table = False
        if para_range.Tables.Count > 0:
            Within_Table = True
        data.append(
            [
                para_number,
                para_style,
                Fl_Indent,
                Hanging,
                font_name,
                font_size,
                char_count,
                word_count,
                tab_count,
                eq_count,
                is_italic,
                is_bold,
                para_text,
                numbering_type,
                list_value,
                Within_Table
            ])
        ranges.append([para_range_start, para_range_end])

        if progress_callback:
            progress_callback((i + 1) / total_paragraphs)

    df = pd.DataFrame(
        data,
        columns=[
            "Paragraph Number",
            "Paragraph Style",
            "First Line Indent (cm)",
            "Hanging Indent (cm)",
            "Font Name",
            "Font Size",
            "Character Count",
            "Word Count",
            "Tab Count",
            "Equals Sign Count",
            "Is Italic",
            "Is Bold",
            "Text",
            "Numbering Type",
            "List Value",
            "Within Table"
        ],
    )

    df["Paragraph Range"] = ranges

    logging.info("Data collection completed.")
    return df, doc
