import json
import tkinter
import time
import queue
import threading
import logging
from tkinter import filedialog, ttk
import customtkinter as ctk
from src import logger
from src.Data_Analysis import DataFrameTab
from src.word_tab import WordTab
from src.file_tab import setup_file_treeview, populate_tree
from src.logger import setup_logging

# Constants
LOG_FILE_PATH = "./data/log/word_app.log"
INITIAL_THEME = "Dark"
INITIAL_COLOR_THEME = "green"
WINDOW_TITLE = "Word AddIn"
WINDOW_GEOMETRY = "1000x600"
QUEUE_PROCESS_INTERVAL = 100

# Configure logging
logger = setup_logging(LOG_FILE_PATH)

class App(ctk.CTk):
    def __init__(self):
        """
        Constructor for the App class.

        Sets up the application window, sets the initial theme, and adds the
        necessary UI elements (toggle theme button, Word Editing tab, and File
        Handling tab).

        :return: None
        """
        super().__init__()
        self.queue = queue.Queue()
        self.setup_ui()
        self.configure_ttk_style()
        self.after(QUEUE_PROCESS_INTERVAL, self.process_queue)

    def setup_ui(self):
        """Sets up the user interface."""
        # Set initial theme for the application
        ctk.set_appearance_mode(INITIAL_THEME)
        ctk.set_default_color_theme(INITIAL_COLOR_THEME)

        # Configure window
        self.title(WINDOW_TITLE)
        self.geometry(WINDOW_GEOMETRY)

        # Create Notebook (tabbed interface)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill='both', expand=True)

        # Add tabs
        self.word_tab_frame = WordEditingTab(self.notebook, self)
        self.notebook.add(self.word_tab_frame, text="Word Editing")

        self.file_handling_tab_frame = FileHandlingTab(self.notebook, self)
        self.notebook.add(self.file_handling_tab_frame, text="File Handling")

        self.data_analysis_tab_frame = DataFrameTab(self.notebook, self, self.queue)
        self.notebook.add(self.data_analysis_tab_frame, text="Data Analysis")

        # Add status bar
        self.status_bar = ctk.CTkLabel(self, text="Ready", anchor='w')
        self.status_bar.pack(side='bottom', fill='x')

    def process_queue(self):
        """Processes tasks in the queue."""
        while not self.queue.empty():
            try:
                task = self.queue.get_nowait()
                # Handle the task (example: updating the UI or logging)
                logger.info(f"Processing task from queue: {task}")
                self.queue.task_done()
            except queue.Empty:
                break
            except Exception as e:
                logger.error(f"Error processing queue: {e}")
        self.after(QUEUE_PROCESS_INTERVAL, self.process_queue)

    def configure_ttk_style(self, theme=INITIAL_THEME):
        """
        Configures the ttk style for the application.

        Configures the theme, notebook, tabs, treeview, and headings for the application.
        The theme can be set to either "Dark" or "Light" to change the overall appearance
        of the application.

        :param theme: The theme to use for the application. Default is "Dark".
        :type theme: str
        """
        style = ttk.Style()
        style.theme_use('default')

        if theme == "Dark":
            # Dark mode styles
            notebook_bg = '#333333'
            tab_bg = '#444444'
            tab_selected_bg = '#555555'
            tree_bg = "#2d2d2d"
            tree_fg = "white"
            heading_bg = "#444444"
            heading_fg = "white"
        else:
            # Light mode styles
            notebook_bg = '#ffffff'
            tab_bg = '#dddddd'
            tab_selected_bg = '#cccccc'
            tree_bg = "#f9f9f9"
            tree_fg = "black"
            heading_bg = "#e0e0e0"
            heading_fg = "black"

        # Configure Notebook tab style
        style.configure('TNotebook', background=notebook_bg)
        style.configure('TNotebook.Tab', background=tab_bg, foreground=tree_fg)
        style.map('TNotebook.Tab', background=[('selected', tab_selected_bg)])

        # Configure Treeview style
        style.configure("Treeview",
                        background=tree_bg,
                        foreground=tree_fg,
                        fieldbackground=tree_bg,
                        font=("Arial", 10))
        style.map("Treeview", background=[("selected", tab_bg)])

        style.configure("Treeview.Heading",
                        background=heading_bg,
                        foreground=heading_fg,
                        font=("Arial", 10))

    def toggle_theme(self):
        """Toggles the theme between Light and Dark."""
        current_theme = ctk.get_appearance_mode()
        new_theme = "Light" if current_theme == "Dark" else "Dark"
        ctk.set_appearance_mode(new_theme)
        self.configure_ttk_style(theme=new_theme)


class FileHandlingTab(ctk.CTkFrame):
    def __init__(self, notebook, root):
        """
        Initialize the FileHandlingTab frame with a treeview and a load button.

        This constructor sets up a treeview widget for displaying file information
        and a button to load files into the treeview. The button initiates a
        threaded file loading process.

        :param notebook: The parent notebook widget to which this frame belongs.
        :type notebook: ctk.CTk
        :param root: The root window or main application instance.
        :type root: ctk.CTk
        """
        super().__init__(notebook)
        self.setup_ui()

    def setup_ui(self):
        """Sets up the user interface."""
        self.tree = setup_file_treeview(self)
        self.load_button = ctk.CTkButton(self, text="Load Files", command=self.load_files_threaded)
        self.load_button.pack(pady=10)

    def load_files_threaded(self):
        """Starts a thread for loading files."""
        threading.Thread(target=self.load_files).start()

    def load_files(self):
        """Loads files into the treeview."""
        try:
            directory = filedialog.askdirectory()
            if directory:
                self.after(0, lambda: populate_tree(self.tree, directory))
        except Exception as e:
            logger.error(f"Error loading files: {e}")


class WordEditingTab(ctk.CTkFrame):
    def __init__(self, notebook, root):
        """
        Initialize the WordEditingTab frame with a treeview and control panel.

        The constructor sets up a treeview widget for displaying and interacting with
        word document data. It also creates a control panel with various buttons and 
        input fields for user interaction, such as search, refresh, navigation, and 
        styling options.

        :param notebook: The parent notebook widget to which this frame belongs.
        :type notebook: ctk.CTk
        :param root: The root window or main application instance.
        :type root: ctk.CTk
        """
        super().__init__(notebook)
        self.setup_ui(root)

    def setup_ui(self, root):
        """Sets up the user interface."""
        # Initialize Treeview
        self.tree = setup_treeview(self, root)

        # Compact control panel for search, refresh, go to, process, and styling
        self.controls_frame = ctk.CTkFrame(self)
        self.controls_frame.pack(fill='x', padx=10, pady=10,)

        # Add Regex label and entry
        self.regex_label = ctk.CTkLabel(self.controls_frame, text="Regex")
        self.regex_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.regex_entry = ctk.CTkEntry(self.controls_frame, width=100)
        self.regex_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        try:
            self.word_tab = WordTab(self.tree)
        except Exception as e:
            logger.error(f"Error initializing WordTab: {e}")
            self.word_tab = WordTab(self.tree)

        # Add Search and Refresh buttons
        self.search_button = ctk.CTkButton(self.controls_frame, text="Search", command=self.refresh_data)
        self.search_button.grid(row=0, column=2, padx=5, pady=5)
        self.search_button.bind("<Enter>", lambda e: self.show_status("Search for matches"))
        self.search_button.bind("<Leave>", lambda e: self.show_status(""))

        self.refresh_button = ctk.CTkButton(self.controls_frame, text="Refresh", command=self.refresh_data)
        self.refresh_button.grid(row=0, column=3, padx=5, pady=5)
        self.refresh_button.bind("<Enter>", lambda e: self.show_status("Refresh the data"))
        self.refresh_button.bind("<Leave>", lambda e: self.show_status(""))

        # Add Go To and Process buttons
        self.goto_button = ctk.CTkButton(self.controls_frame, text="Go To", command=self.to_paragraph)
        self.goto_button.grid(row=0, column=4, padx=5, pady=5)
        self.goto_button.bind("<Enter>", lambda e: self.show_status("Go to the selected paragraph"))
        self.goto_button.bind("<Leave>", lambda e: self.show_status(""))

        self.process_button = ctk.CTkButton(self.controls_frame, text="Process", command=self.process_paragraph)
        self.process_button.grid(row=0, column=5, padx=5, pady=5)
        self.process_button.bind("<Enter>", lambda e: self.show_status("Process the selected paragraph"))
        self.process_button.bind("<Leave>", lambda e: self.show_status(""))

        # Style and Font Entries
        ctk.CTkLabel(self.controls_frame, text="Style to Apply:").grid(row=1, column=0, sticky="w")
        self.style_entry = ctk.CTkEntry(self.controls_frame)
        self.style_entry.grid(row=1, column=1, columnspan=2, padx=5, pady=5)

        ctk.CTkLabel(self.controls_frame, text="Font:").grid(row=1, column=4, sticky="w")
        self.font_entry = ctk.CTkEntry(self.controls_frame)
        self.font_entry.grid(row=1, column=5, columnspan=2, padx=5, pady=5)

    def refresh_data(self):
        """Refreshes the data in the TreeView based on the regex pattern."""
        try:
            # Clear previous match positions
            self.word_tab.match_positions.clear()

            # Get regex pattern from the input field
            regex_pattern = self.regex_entry.get().strip()

            # Call check_word_selection to populate the TreeView with matches
            self.word_tab.check_word_selection(
                regex_entry=regex_pattern,
                update_treeview=self.update_treeview
            )
        except Exception as e:
            logger.error(f"Error during refresh: {e}")

    def update_treeview(self, df):
        """Updates the TreeView with new data."""
        logger.info("Updating TreeView...")
        logger.info(f"The df is: {df}")

        # Clear existing data
        self.tree.delete(*self.tree.get_children())

        # Insert new data from DataFrame
        for idx, row in df.iterrows():
            self.tree.insert("", "end", values=row.tolist(), iid=idx)

    def to_paragraph(self):
        """Defines the action when the 'Go To' button is clicked."""
        logger.info("Go To Paragraph is clicked")
        self.word_tab.goto_paragraph(self.tree)

    def process_paragraph(self):
        """Defines the action when the 'Process' button is clicked."""
        self.word_tab.prompt_for_value_and_replace(self.tree, self.word_tab)

    def show_status(self, message):
        """Updates the status bar with the given message."""
        self.status_bar.config(text=message)


def setup_treeview(parent_frame, root):
    """
    Creates a Treeview widget with columns for Match, Style, Paragraph, First Line Indent, and Left Indent.
    Each column is given a heading and a width in pixels, and the Treeview is packed to fill the entire
    parent frame.

    :param parent_frame: The frame to pack the Treeview widget into.
    :type parent_frame: ttk.Frame
    :param root: The root window or main application instance.
    :type root: ctk.CTk
    :return: The created Treeview widget.
    :rtype: ttk.Treeview
    """
    tree = ttk.Treeview(parent_frame, columns=('Match', 'Style', 'Paragraph', 'First Line Indent', 'Left Indent'), show='headings')
    tree.heading('Match', text='Match')
    tree.heading('Style', text='Style')
    tree.heading('Paragraph', text='Paragraph')
    tree.heading('First Line Indent', text='First Line Indent')
    tree.heading('Left Indent', text='Left Indent')

    # Define column properties
    tree.column('Match', width=150, anchor='w')
    tree.column('Style', width=100, anchor='w')
    tree.column('Paragraph', width=250, anchor='w')
    tree.column('First Line Indent', width=120, anchor='w')
    tree.column('Left Indent', width=100, anchor='w')

    # Bind double-click event to the on_double_click function
    tree.bind("<Double-1>", lambda event: WordTab.on_double_click(event, tree, root))
    # Pack the Treeview
    tree.pack(fill=tkinter.BOTH, expand=True)
    return tree

if __name__ == "__main__":
    app = App()
    app.mainloop()
