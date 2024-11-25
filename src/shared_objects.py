import pythoncom
import win32com.client as win32
import os
from pathlib import Path
from src.logger import setup_logging

class WordApp:
    """
    A class to interact with Microsoft Word through the win32com library.

    This class initializes a Word application instance, opens or creates a document,
    and provides various methods to interact with the document.

    Parameters
    ----------
    filepath : str, optional
        The path to the Word document to open. If not provided, no document is opened.
    visible : bool, optional
        Whether to make the Word application visible. Default is True.
    display_alerts : bool, optional
        Whether to display Word alerts. Default is True.
    """
    def __init__(self, filepath=None, visible=True, display_alerts=True):
        

        pythoncom.CoInitialize()
        self.word_app = win32.Dispatch("Word.Application")
        self.word_app.Visible = visible
        self.word_app.DisplayAlerts = display_alerts
        self.doc = None
        self.file_name = None
        self.logger = setup_logging()
        if filepath:
            self.open_document(filepath)
        else:
            self._initialize_document()

        self.logger.info("WordApp initialized")

    def _initialize_document(self):
        """Initializes a new document or connects to the active document."""
        if self.word_app.Documents.Count < 1:
            self.doc = self.word_app.Documents.Add()
        else:
            self.doc = self.word_app.ActiveDocument
        if hasattr(self.doc, 'Name'):
            self.file_name = self.doc.Name
        else:
            raise AttributeError("The document object does not have a 'Name' attribute.")

    def open_document(self, filepath):
        """Opens a Word document."""
        try:
            formatted_path = Path(filepath).resolve()
            self.doc = self.word_app.Documents.Open(str(formatted_path))
            self.file_name = os.path.basename(filepath)
            self.logger.info(f"Document opened: {self.file_name}")
            return self.doc
        except Exception as e:
            self.logger.error(f"Failed to open document: {e}")
            raise

    def save_document(self, filepath=None):
        """Saves the Word document."""
        try:
            if filepath:
                self.doc.SaveAs(Path(filepath).resolve())
                self.file_name = os.path.basename(filepath)
            else:
                self.doc.Save()
            self.logger.info(f"Document saved: {self.file_name}")
        except Exception as e:
            self.logger.error(f"Failed to save document: {e}")
            raise

    def close_document(self):
        """Closes the Word document."""
        try:
            if self.doc:
                self.doc.Close()
                self.logger.info(f"Document closed: {self.file_name}")
                self.doc = None
        except Exception as e:
            self.logger.error(f"Failed to close document: {e}")
            raise

    def quit_word(self):
        """Quits the Word application."""
        try:
            if self.word_app:
                self.word_app.Quit()
                self.logger.info("Word application quit")
        except Exception as e:
            self.logger.error(f"Failed to quit Word application: {e}")
            raise

    def get_active_document(self):
        """Returns the active document in the Word application."""
        try:
            return self.word_app.ActiveDocument
        except Exception as e:
            self.logger.error(f"No active document found: {e}")
            return None

    def get_word_app(self):
        """Returns the Word application instance."""
        return self.word_app

    def get_document(self):
        """Returns the current document."""
        return self.doc

    def insert_text(self, text, position=None):
        """Inserts text into the document at the specified position."""
        try:
            if position:
                self.doc.Range(position, position).Text = text
            else:
                self.doc.Content.Text += text
            self.logger.info("Text inserted into document")
        except Exception as e:
            self.logger.error(f"Failed to insert text: {e}")
            raise

    def format_text(self, start, end, font_name=None, font_size=None, bold=None, italic=None):
        """Formats text in the document."""
        try:
            rng = self.doc.Range(start, end)
            if font_name:
                rng.Font.Name = font_name
            if font_size:
                rng.Font.Size = font_size
            if bold is not None:
                rng.Font.Bold = bold
            if italic is not None:
                rng.Font.Italic = italic
            self.logger.info("Text formatted in document")
        except Exception as e:
            self.logger.error(f"Failed to format text: {e}")
            raise

    def __enter__(self):
        """Enter the runtime context related to this object."""
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        """Exit the runtime context related to this object."""
        self.quit_word()
        pythoncom.CoUninitialize()
        if exc_type:
            self.logger.error(f"Exception: {exc_type}, {exc_value}")
            raise


