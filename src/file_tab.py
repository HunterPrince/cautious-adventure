import win32com
import pyautogui
from src import logger
from src.utlities import adjust_header, makePdf
import win32com.client as win32
from win32com.client import constants
import gc
import os
import re
import tkinter as tnktr
import customtkinter as ctk
from tkinter import ttk, messagebox
from src.shared_objects import WordApp
from src.logger import setup_logging

logger = setup_logging()

class DocumentHandler:
    def __init__(self, word_app_instance=None):
        """
        Initializes the DocumentHandler instance.

        If word_app_instance is None, creates a new instance of WordApp.
        Otherwise, uses the provided instance of WordApp.

        Parameters
        ----------
        word_app_instance : WordApp, optional
            The instance of WordApp to use. The default is None.
        """
        self.word_app_instance = word_app_instance or WordApp()
        self.doc = None
        
    def open_document(self, filepath):
        """
        Opens the document and returns a reference to the document.
        """
        try:
            self.doc = self.word_app_instance.open_document(filepath)
            return self.doc
        except Exception as e:
            logger.error(f"Failed to open document {filepath}: {e}")
            return None

    def extract_properties(self, doc):
        """Extracts properties like pages, sections, and headers (even/odd)."""
        try:
            self.doc.Activate()
            selection = self.word_app_instance.word_app.Selection
            active_window = self.word_app_instance.word_app.ActiveWindow
            adjust_header(active_window, selection)
            pages = doc.ComputeStatistics(2)  # 2 = wdStatisticPages
            sections = doc.Sections.Count

            # Extract even and odd headers from the first section
            section = doc.Sections(1)
            first_page_number = selection.Information(constants.wdActiveEndAdjustedPageNumber)
            selection.EndKey(constants.wdStory)
            last_page_number = selection.Information(constants.wdActiveEndAdjustedPageNumber)
            logger.info(f"Last page number: {last_page_number}")

            odd_page_header = section.Headers(win32.constants.wdHeaderFooterPrimary).Range.Text.strip()
            even_page_header = section.Headers(win32.constants.wdHeaderFooterEvenPages).Range.Text.strip()

            if not odd_page_header and not even_page_header:
                return {
                    'pages': pages,
                    'sections': sections,
                    'headers': 'Not declared',
                    'Starting_Page': first_page_number,
                    'Ending_Page': last_page_number
                }
            else:
                return {
                    'pages': pages,
                    'sections': sections,
                    'headers': {
                        'even': even_page_header,
                        'odd': odd_page_header,
                    },
                    'Starting_Page': first_page_number,
                    'Ending_Page': last_page_number
                }
        except Exception as e:
            logger.error(f"Error extracting document properties: {e}")
            return {}

    def close_document(self, save_changes=False):
        """Close the document after the actions are performed."""
        try:
            self.word_app_instance.close_document(self.doc)
        except Exception as e:
            logger.error(f"Error closing document: {e}")

    def edit_header_footer(self, doc, header_text=None, footer_text=None):
        """Edits the header and footer of the document."""
        try:
            if header_text:
                for section in doc.Sections:
                    for header in section.Headers:
                        header.Range.Text = header_text
            if footer_text:
                for section in doc.Sections:
                    for footer in section.Footers:
                        footer.Range.Text = footer_text
            logger.info(f"Successfully updated header/footer in document {doc.Name}")
        except Exception as e:
            logger.error(f"Error updating header/footer: {e}")

def count_wordapp_instances():
    return sum(1 for obj in gc.get_objects() if isinstance(obj, WordApp))

def count_document_instances():
    return sum(1 for obj in gc.get_objects() if isinstance(obj, WordApp))

def list_doc_files(directory):
    return [f for f in os.listdir(directory) if f.endswith('.doc') and not f.startswith('~') or f.endswith('.docx')]

def custom_sort_key(filename):
    match = re.search(r'Unit (\d+)|unit (\d+)', filename)
    if match:
        return int(match.group(1) or match.group(2)), filename
    return float('inf'), filename

def setup_file_treeview(parent_frame):
    """
    Create a Treeview widget with columns for File, Pages, Sections, Headers, Starting Page, Ending Page, and Footers.
    The Treeview is packed to fill the entire parent frame with both directions and buttons are added below the Treeview.

    :param parent_frame: The frame to pack the Treeview widget into.
    :type parent_frame: ttk.Frame
    :return: The created Treeview widget.
    :rtype: ttk.Treeview
    """
    tree = ttk.Treeview(parent_frame, columns=('File', 'Pages', 'Sections', 'Headers', 'Starting Page', 'Ending Page', 'Footers'), show='headings')
    tree.heading('File', text='File')
    tree.heading('Pages', text='Pages')
    tree.heading('Sections', text='Sections')
    tree.heading('Headers', text='Headers')
    tree.heading('Starting Page', text='Starting Page')
    tree.heading('Ending Page', text='Ending Page')
    tree.heading('Footers', text='Footers')

    # Define column properties
    tree.column('File', width=200, anchor='w')
    tree.column('Pages', width=50, anchor='center')
    tree.column('Sections', width=70, anchor='center')
    tree.column('Headers', width=100, anchor='center')
    tree.column('Starting Page', width=100, anchor='center')
    tree.column('Ending Page', width=100, anchor='center')
    tree.column('Footers', width=100, anchor='center')

    tree.pack(fill=tnktr.BOTH, expand=True)

    # Add Buttons Below the TreeView
    btn_frame = ctk.CTkFrame(parent_frame)
    btn_frame.pack(fill=tnktr.X, padx=10, pady=5)

    # Button to Select Source File
    btn_select_source = ctk.CTkButton(btn_frame, text="Select Source File", command=lambda: select_source_file(tree))
    btn_select_source.pack(side=tnktr.LEFT, padx=5)
    # btn_select_source.bind("<Enter>", lambda e: show_status("Select a source file"))
    # btn_select_source.bind("<Leave>", lambda e: show_status(""))

    # Button to make pdf of selected files
    btn_make_pdf = ctk.CTkButton(btn_frame, text="Make PDF", command=lambda: makePdf(tree, dirr))
    btn_make_pdf.pack(side=tnktr.RIGHT, padx=5)
    # btn_make_pdf.bind("<Enter>", lambda e: show_status("Make PDF of selected files"))
    # btn_make_pdf.bind("<Leave>", lambda e: show_status(""))

    # Button to Select and Copy Headers to Destination Files
    btn_copy_headers = ctk.CTkButton(btn_frame, text="Select Destination and Copy Headers", command=lambda: select_destination_and_copy_headers(tree))
    btn_copy_headers.pack(side=tnktr.LEFT, padx=5)
    # btn_copy_headers.bind("<Enter>", lambda e: show_status("Select destination files and copy headers"))
    # btn_copy_headers.bind("<Leave>", lambda e: show_status(""))

    continue_page_number = ctk.CTkButton(btn_frame, text="Continue Page Numbers", command=lambda: continue_page_numbers(tree))
    continue_page_number.pack(side=tnktr.RIGHT, padx=5)
    # continue_page_number.bind("<Enter>", lambda e: show_status("Continue page numbers across documents"))
    # continue_page_number.bind("<Leave>", lambda e: show_status(""))

    return tree

def select_source_file(tree):
    """
    Function to select the source file from the TreeView and store it in a global variable.
    
    Parameters:
    tree (ttk.Treeview): The Treeview widget with document entries.
    """
    global selected_source_file
    selected_items = tree.selection()

    if len(selected_items) != 1:
        messagebox.showerror("Error", "Please select exactly one source file.")
        return

    selected_source_file = tree.item(selected_items[0], 'values')[0]  # Get the file name of the selected item

    messagebox.showinfo("Source Selected", f"Source file '{selected_source_file}' selected.")
    show_status(f"Source file '{selected_source_file}' selected.")

def select_destination_and_copy_headers(tree):
    """
    Function to select destination files from the TreeView and copy headers from the selected source file.
    Parameters:
    tree (ttk.Treeview): The Treeview widget with document entries.
    """
    global selected_source_file

    # Ensure a source file has been selected
    if not selected_source_file:
        messagebox.showerror("Error", "Please select a source file first.")
        return
    selected_items = tree.selection()
    if len(selected_items) == 0:
        messagebox.showerror("Error", "Please select at least one destination file.")
        return
    destination_files = [tree.item(item, 'values')[0] for item in selected_items]

    logger.info(f"Destination files: {destination_files}")
    copy_headers_to_files(selected_source_file, destination_files, dirr)

def copy_headers_to_files(source_file, destination_files, directory):
    """
    Function to copy headers (even and odd) from source file to destination files.
    Parameters:
    source_file (str): The file name of the source document.
    destination_files (list): A list of file names of the destination documents.
    """
    try:
        # Open the source document
        word_instance = WordApp()
        doc_handler = DocumentHandler(word_instance)
        source_doc = doc_handler.open_document(os.path.join(directory, source_file))
        
        if not source_doc:
            messagebox.showerror("Error", f"Failed to open source file '{source_file}'.")
            return

        # Extract odd and even headers from the source document
        source_even_header_range = source_doc.Sections(1).Headers(win32.constants.wdHeaderFooterEvenPages).Range
        source_odd_header_range = source_doc.Sections(1).Headers(win32.constants.wdHeaderFooterPrimary).Range     
        source_even_header_range.Select()
        source_odd_header_range.Select()
        for dest_file in destination_files:
            dest_doc = doc_handler.open_document(os.path.join(directory, dest_file))
            if not dest_doc:
                messagebox.showerror("Error", f"Failed to open destination file '{dest_file}'. Skipping...")
                continue
            # Insert headers into the destination document
            insert_headers(dest_doc, source_even_header_range, source_odd_header_range)
            messagebox.showinfo("Success", f"Copied headers to '{dest_file}'.")

            # Close destination document after copying headers
            doc_handler.close_document(dest_doc, True)

        # Close source document after extracting headers
        doc_handler.close_document(source_doc, False)

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def insert_headers(doc, source_even_header_range, source_odd_header_range):
    """
    Function to insert odd and even headers into a document with rich formatting.
    
    Parameters:
    doc: The Word document object.
    source_even_header_range: The range of the even header in the source document.
    source_odd_header_range: The range of the odd header in the source document.
    """
    try:
        # Get the primary (odd page) header range
        odd_header_range = doc.Sections(1).Headers(win32.constants.wdHeaderFooterPrimary).Range
        even_header_range = doc.Sections(1).Headers(win32.constants.wdHeaderFooterEvenPages).Range

        # Copy and paste headers
        source_odd_header_range.Copy()
        odd_header_range.Select()
        odd_header_range.Text = ""
        odd_header_range.Paste()
        if odd_header_range.Characters.Last.Text in ["\n", "\r"]:
            odd_header_range.Characters.Last.Delete()

        source_even_header_range.Copy()
        even_header_range.Select()
        even_header_range.Text = ""
        even_header_range.Paste()
        if even_header_range.Characters.Last.Text in ["\n", "\r"]:
            even_header_range.Characters.Last.Delete()
    except Exception as e:
        logger.error(f"Error inserting headers: {e}")

def continue_page_numbers(tree):
    """
    Function to continue page numbering from one document to another.
    
    Parameters:
    tree (ttk.Treeview): The Treeview widget with document entries.
    """
    starting_page_number = 1
    selected_items = tree.selection()
    selected_items = [tree.item(item, 'values')[0] for item in selected_items]
    if len(selected_items) <= 1:
        messagebox.showerror("Error", "Please select more than two files for continuing page numbers")
        return

    word_instance = WordApp()
    doc_handler = DocumentHandler(word_instance)
    selected_items = [f for f in selected_items if f != ""]
    new_page_numbers = {}

    for s in selected_items:
        doc = doc_handler.open_document(os.path.join(dirr, s))
        if not doc:
            continue

        doc.Sections(1).Headers(constants.wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
        doc.Sections(1).Headers(constants.wdHeaderFooterPrimary).PageNumbers.StartingNumber = starting_page_number
        selection = doc_handler.word_app_instance.word_app.Selection
        active_window = word_instance.word_app.ActiveWindow
        adjust_header(active_window, selection)
        selection.GoTo(What=constants.wdGoToPage, Which=constants.wdGoToLast)
        last_page_number = selection.Information(constants.wdActiveEndAdjustedPageNumber)
        total_pages = doc.ComputeStatistics(2)
        new_page_numbers[s] = {'Starting_Page': starting_page_number, 'Ending_Page': last_page_number, 'Total_Pages': total_pages}
        starting_page_number = last_page_number + 1 if last_page_number < total_pages else starting_page_number + total_pages

        doc_handler.close_document(doc, True)

    # Update the view for the file tab page
    for key, value in new_page_numbers.items():
        matching_item = None
        for item in tree.get_children():
            current_values = tree.item(item, 'values')
            if current_values[0] == key:
                matching_item = item
                break

        if matching_item:
            updated_values = list(tree.item(matching_item, 'values'))
            updated_values[4] = value['Starting_Page']
            updated_values[5] = value['Ending_Page']
            tree.item(matching_item, values=tuple(updated_values))

def populate_tree(tree, directory):
    """
    Populates the treeview with document files and their properties.
    
    Parameters:
    tree (ttk.Treeview): The treeview to populate.
    directory (str): The directory to scan for files.
    
    Returns:
    None
    """
    global dirr
    dirr = directory
    process_documents(directory, tree)

def process_documents(directory, tree):
    """Main function to process multiple documents and insert data into the TreeView."""
    files = list_doc_files(directory)
    files = [f for f in files if not f.startswith('~')]
    files = sorted(files, key=custom_sort_key)

    tree.delete(*tree.get_children())  # Clear previous entries

    word_instance = WordApp()
    doc_handler = DocumentHandler(word_instance)

    for file in files:
        filepath = os.path.join(directory, file)
        if not os.path.exists(filepath):
            logger.warning(f"File not found: {filepath}")
            continue

        doc = doc_handler.open_document(filepath)
        if doc:
            properties = doc_handler.extract_properties(doc)
            logger.info(properties)

            file_id = tree.insert('', 'end', values=(
                file,
                properties.get('pages', 'N/A'),
                properties.get('sections', 'N/A'),
                "Not declared" if properties.get('headers') == 'Not declared' else "Headers",
                properties.get('Starting_Page', 'N/A'),
                properties.get('Ending_Page', 'N/A')
            ))

            headers = properties.get('headers', {})
            if headers != 'Not declared':
                even_header = headers.get('even', 'N/A')
                odd_header = headers.get('odd', 'N/A')
                if even_header or odd_header:
                    tree.insert(file_id, 'end', values=('', '', '', f'Even Header: {even_header}', '', '', ''))
                    tree.insert(file_id, 'end', values=('', '', '', f'Odd Header: {odd_header}', '', '', ''))

            logger.info(f"Active instances of WordApp: {count_wordapp_instances()}")
            logger.info(f"Active instances of WordApp: {count_document_instances()}")
            doc_handler.close_document(doc)

# def show_status(message):
#     """Updates the status bar with the given message."""
#     status_bar.config(text=message)

# # Add a status bar

# status_bar = ctk.CTkLabel(root, text="Ready", anchor='w')
# status_bar.pack(side='bottom', fill='x')
