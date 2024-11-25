from tkinter.tix import COLUMN
from numpy import column_stack
from src.shared_objects import WordApp
import pandas as pd
from customtkinter import CTkButton, CTkComboBox
import customtkinter as ctk
from tkinter import ttk, messagebox
import threading
import json
from src.utlities import collect_data
from src.logger import setup_logging

# Setup logging
logger = setup_logging()

class DataFrameTab(ctk.CTkFrame):
    def __init__(self, notebook, root, queue):
        super().__init__(notebook)
        self.queue = queue
        self.df = None
        self.filtered_df = None
        self.original_df = None
        self.filter_history = []
        self.history_file = r"Data\filter_history.json"
        self.load_filter_history()
        self.word_app = WordApp()
        self.doc = self.word_app.get_active_document()
        self.drag_data = {"item": None, "index": None}
        self.setup_ui()
        self.setup_treeview()
        self.setup_scrollbars()

    def setup_ui(self):
        self.sidebar_frame = ctk.CTkFrame(self)
        self.sidebar_frame.pack(side="bottom", fill="y", padx=10, pady=10)
        self.button_frame = ctk.CTkFrame(self.sidebar_frame)
        self.button_frame.grid(row=0, column=8, columnspan=1, padx=5, pady=5, sticky="ew")

        self.filter_label = ctk.CTkLabel(self.button_frame, text="Filter Expression:")
        self.filter_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.filter_entry = ctk.CTkEntry(self.button_frame, width=150)
        self.filter_entry.grid(row=0, column=1,columnspan=2, padx=5, pady=5, sticky="w")
        self.filter_entry.bind("<Return>", self.apply_filter)

        self.filter_history_dropdown = CTkComboBox(self.button_frame, values=self.filter_history, width=150)
        self.filter_history_dropdown.grid(row=0, column=2,columnspan=2, padx=5, pady=5, sticky="w")
        self.filter_history_dropdown.bind("<<ComboboxSelected>>", self.load_filter_from_history)


        self.import_button = CTkButton( self.button_frame, text="Import Data", command=self.import_from_csv)
        self.import_button.grid(row=0, column=7, padx=5, pady=5, sticky="w")

        self.export_button = CTkButton(  self.button_frame, text="Export Data", command=self.export_to_csv)
        self.export_button.grid(row=0, column=6, padx=5, pady=5, sticky="w")

        self.sort_label = ctk.CTkLabel( self.button_frame, text="Sort by:")
        self.sort_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.sort_dropdown = CTkComboBox( self.button_frame, values=['Font Name', 'Paragraph Style', "Paragraph Number", "Within Table", "Font Size", "IsBold", "Character Count"], width=150)
        self.sort_dropdown.grid(row=1, column=1,columnspan=2, padx=5, pady=5, sticky="w")

        self.sort_button = CTkButton( self.button_frame, text="Sort Data", command=self.sort_data)
        self.sort_button.grid(row=1, column=4, padx=5, pady=5, sticky="ew")

        self.search_label = ctk.CTkLabel( self.button_frame, text="Search:")
        self.search_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.search_entry = ctk.CTkEntry( self.button_frame, width=150)
        self.search_entry.grid(row=2, column=1,columnspan=2, padx=5, pady=5, sticky="w")
        self.search_entry.bind("<Return>", self.search_data)

        self.goto_button = CTkButton( self.button_frame, text="Go To Paragraph", command=lambda: self.goto_paragraph(self.treeview))
        self.goto_button.grid(row=0, column=4, padx=5, pady=5, sticky="ew")

        self.modify_button = CTkButton( self.button_frame, text="Modify Selected", command=self.modify_selected_paragraphs)
        self.modify_button.grid(row=2, column=4, padx=5, pady=5, sticky="ew")

        self.revert_button = CTkButton( self.button_frame, text="Revert DataFrame", command=self.revert_to_original)
        self.revert_button.grid(row=0, column=5, padx=5, pady=5, sticky="ew")

        self.load_file_button = CTkButton( self.button_frame, text="Load Files", command=lambda: self.run_in_thread(self.load_file_for_data_frame))
        self.load_file_button.grid(row=1, column=5, columnspan=1, padx=5, pady=5, sticky="ew")

        # self.progress_bar = ctk.CTkProgressBar( self.button_frame)
        # self.progress_bar.grid(row=10, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

        self.button_frame.grid_columnconfigure(1, weight=1)
        for i in range(11):
             self.button_frame.grid_rowconfigure(i, weight=1)
        self.button_frame.grid_rowconfigure(11, weight=4)


    def setup_treeview(self):
        self.treeview_frame = ctk.CTkFrame(self)
        self.treeview_frame.pack(pady=10, padx=10, fill="both", expand=True)

        self.treeview = ttk.Treeview(self.treeview_frame, show="headings", height=15)
        self.treeview.pack(side="left", fill="both", expand=True)
        self.treeview.bind("<Double-1>", self.on_double_click)
        self.treeview.bind("<ButtonPress-1>", self.on_treeview_press)
        self.treeview.bind("<B1-Motion>", self.on_treeview_motion)
        self.treeview.bind("<ButtonRelease-1>", self.on_treeview_release)

    def setup_scrollbars(self):
        self.vertical_scrollbar = ctk.CTkScrollbar(self.treeview_frame, orientation="vertical", command=self.treeview.yview)
        self.vertical_scrollbar.pack(side="right", fill="y")

        self.horizontal_scrollbar = ctk.CTkScrollbar(self, orientation="horizontal", command=self.treeview.xview)
        self.horizontal_scrollbar.pack(fill="x")

        self.treeview.configure(yscrollcommand=self.vertical_scrollbar.set, xscrollcommand=self.horizontal_scrollbar.set)

    def on_treeview_press(self, event):
        item = self.treeview.identify_row(event.y)
        if item:
            self.drag_data["item"] = item
            self.drag_data["index"] = self.treeview.index(item)

    def on_treeview_motion(self, event):
        if self.drag_data["item"]:
            self.treeview.move(self.drag_data["item"], "", self.treeview.index(self.treeview.identify_row(event.y)))
    # def update_word_document_order(self):
    #     word = WordApp()
    #     doc = word.get_active_document()
    #     for index, row in self.df.iterrows():
    #         para_range_data = row["Paragraph Range"]
    #         para_range_start = para_range_data[0]
    #         para_range_end = para_range_data[1]
    #         para_range = doc.Range(Start=para_range_start, End=para_range_end)
    #         para_range.Select()
    def on_treeview_release(self, event):
        if self.drag_data["item"]:
            new_index = self.treeview.index(self.drag_data["item"])
            self.update_dataframe_order(self.drag_data["index"], new_index)
            self.drag_data["item"] = None
            self.drag_data["index"] = None

    def update_dataframe_order(self, old_index, new_index):
        if old_index != new_index:
            row = self.df.iloc[old_index]
            self.df.drop(old_index, inplace=True)
            self.df = pd.concat([self.df.iloc[:new_index], pd.DataFrame([row]), self.df.iloc[new_index:]]).reset_index(drop=True)
            self.update_word_document_order()

    def setup_scrollbars(self):
        self.vertical_scrollbar = ctk.CTkScrollbar(self.treeview_frame, orientation="vertical", command=self.treeview.yview)
        self.vertical_scrollbar.pack(side="right", fill="y")

        self.horizontal_scrollbar = ctk.CTkScrollbar(self, orientation="horizontal", command=self.treeview.xview)
        self.horizontal_scrollbar.pack(fill="x")

        self.treeview.configure(yscrollcommand=self.vertical_scrollbar.set, xscrollcommand=self.horizontal_scrollbar.set)

    def on_double_click(self, event):
        item = self.treeview.selection()[0]
        column = self.treeview.identify_column(event.x)
        row = self.treeview.identify_row(event.y)
        col_index = int(column.replace('#', '')) - 1

        if col_index in [1, 2, 3, 4]:  # Allow editing for specific columns
            self.edit_cell(row, col_index)

    def edit_cell(self, row, col_index):
        x, y, width, height = self.treeview.bbox(row, f'#{col_index + 1}')
        entry = ctk.CTkEntry(self.treeview, width=width, height=height)  # Pass width and height here
        entry.place(x=x, y=y)

        def on_focus_out(event):
            new_value = entry.get()
            self.treeview.set(row, f'#{col_index + 1}', new_value)
            self.update_dataframe(row, col_index, new_value)
            entry.destroy()

        entry.bind("<FocusOut>", on_focus_out)

    def update_dataframe(self, row, col_index, new_value):
        para_number = self.treeview.item(row, "values")[0]
        column_name = self.treeview["columns"][col_index]
        self.df.loc[self.df["Paragraph Number"] == int(para_number), column_name] = new_value
        self.update_word_document(para_number, column_name, new_value)

    def update_word_document(self, para_number, column_name, new_value):
        doc = self.doc 
        para_range_data = self.df.loc[self.df["Paragraph Number"] == int(para_number), "Paragraph Range"].values[0]
        para_range_start = para_range_data[0]
        para_range_end = para_range_data[1]
        para_range = doc.Range(Start=para_range_start, End=para_range_end)
        para_range.Select()

        try:
            if column_name == "Font Name":
                para_range.Font.Name = new_value
            elif column_name == "Font Size":
                para_range.Font.Size = float(new_value)
            elif column_name == "Paragraph Style":
                para_range.Style = new_value
            elif column_name == "Hanging Indent":
                para_range.ParagraphFormat.LeftIndent = float(new_value) * 28.346
            elif column_name == "First Line Indent":
                para_range.ParagraphFormat.FirstLineIndent = float(new_value) * 28.346
        except Exception as e:
            logger.error(f"Error updating Word document for paragraph {para_number} with {column_name}: {e}")

        logger.info(f"Updated Word document for paragraph {para_number} with {column_name}: {new_value}")



    def search_data(self, event=None):
        search_term = self.search_entry.get()
        if self.df is not None and search_term:
            search_results = self.df[self.df.apply(lambda row: row.astype(str).str.contains(search_term, case=False).any(), axis=1)]
            self.filtered_df = search_results
            self.update_display()
        else:
            messagebox.showwarning("No Data", "No data available to search.")

    def apply_filter(self, event=None):
        if self.df is not None:
            filter_expression = self.filter_entry.get()
            try:
                self.filtered_df = self.df.query(filter_expression)
                self.update_display()
                if filter_expression not in self.filter_history:
                    self.filter_history.append(filter_expression)
                    self.filter_history_dropdown["values"] = self.filter_history
            except Exception as e:
                self.treeview.delete(*self.treeview.get_children())
                self.treeview.insert("", "end", values=["Error in filter expression:", str(e)])
                logger.error(f"Error in filter expression: {e}")
        else:
            messagebox.showwarning("No Data", "No data available to filter.")

    def sort_data(self, sort_column=None):
        if self.df is not None:
            if not sort_column:
                sort_column = self.sort_dropdown.get()
            if sort_column:
                try:
                    self.filtered_df = self.filtered_df.sort_values(by=sort_column)
                    self.update_display()
                except Exception as e:
                    messagebox.showerror("Error", f"Error sorting by {sort_column}: {e}")
                    logger.error(f"Error sorting by {sort_column}: {e}")
            else:
                messagebox.showwarning("No Column Selected", "Please select a column to sort by.")
        else:
            messagebox.showwarning("No Data", "No data available to sort.")

    def update_display(self):
        for col in self.treeview.get_children():
            self.treeview.delete(col)

        if self.filtered_df is not None:
            display_columns = [col for col in self.filtered_df.columns if col != "Paragraph Range"]
            self.treeview["columns"] = display_columns
            for col in display_columns:
                self.treeview.heading(col, text=col, command=lambda _col=col: self.sort_data(_col))
                self.treeview.column(col, width=100, anchor="w", stretch=True)

            for _, row in self.filtered_df.iterrows():
                self.treeview.insert("", "end", values=row[display_columns].tolist())

            self.sort_dropdown["values"] = display_columns
        else:
            self.treeview["columns"] = []
            self.treeview.insert("", "end", values=["No data loaded."])

    def process_queue(self):
        while not self.queue.empty():
            task = self.queue.get()
            logger.info(f"Processing task from queue: {task}")
            self.queue.task_done()

    def load_filter_from_history(self, event):
        selected_filter = self.filter_history_dropdown.get()
        self.filter_entry.delete(0, "end")
        self.filter_entry.insert(0, selected_filter)

    def save_filter_history(self):
        with open(self.history_file, "w") as f:
            json.dump(self.filter_history, f)
        logger.info("Filter history saved.")

    def load_filter_history(self):
        try:
            with open(self.history_file, "r") as f:
                content = f.read().strip()
                if content:
                    self.filter_history = json.loads(content)
                else:
                    self.filter_history = []
        except FileNotFoundError:
            self.filter_history = []
        except json.JSONDecodeError:
            self.filter_history = []
        logger.info("Filter history loaded.")

    def load_file_for_data_frame(self):
        data, doc = collect_data()
        self.df = pd.DataFrame(data)
        self.original_df = self.df.copy()
        self.filtered_df = self.df
        self.update_display()
        doc = self.doc 
        logger.info("File loaded and DataFrame created.")
        self.update_progress_bar(1.0)  # Set progress bar to 100% when done

    def update_progress_bar(self, value):
        self.progress_bar.set(value)
        self.update_idletasks()

    def run_in_thread(self, target_function):
        def wrapper():
            self.update_progress_bar(0)  # Initialize progress bar to 0%
            target_function()
            self.update_progress_bar(1.0)  # Set progress bar to 100% when done

        thread = threading.Thread(target=wrapper, daemon=True)
        thread.start()
        logger.info(f"Started thread for {target_function.__name__}")

    def revert_to_original(self):
        if self.original_df is not None:
            self.filtered_df = self.original_df.copy()
            self.update_display()
            logger.info("Reverted to original data.")

    def prompt_update_active_document(self):
        user_response = messagebox.askyesno(
            "Update Active Document",
            "Changes have been made. Do you want to update the original document?",
        )
        if user_response:
            self.run_in_thread(self.update_active_document)
        

    def export_to_csv(self):
        if self.df is not None:
            try:
                file_path = r"Data\exported_data.csv"
                self.df.to_csv(file_path, index=False)
                messagebox.showinfo("Export Successful", f"Data exported to {file_path}")
                logger.info(f"Data exported to {file_path}")
            except Exception as e:
                messagebox.showerror("Export Failed", f"Failed to export data: {e}")
                logger.error(f"Failed to export data: {e}")
        else:
            messagebox.showwarning("No Data", "No data available to export.")

    def import_from_csv(self):
        try:
            file_path = r"Data\exported_data.csv"
            self.df= pd.read_csv(file_path)
            user_response = messagebox.askyesno("Update View", "Do you want to update the view with the imported data?")
            if user_response:
                self.filtered_df = self.df
                self.update_display()
                messagebox.showinfo("Import Successful", f"Data imported from {file_path}")
                logger.info(f"Data imported from {file_path}")
            else:
                logger.info("User chose not to update the view.")
        except Exception as e:
            messagebox.showerror("Import Failed", f"Failed to import data: {e}")
            logger.error(f"Failed to import data: {e}")
    
        # ... existing code ...
    def update_active_document(self):
        self.load_file_for_data_frame()
        messagebox.showinfo("Document Updated", "The document has been updated successfully.")
        logger.info("Active document updated.")

    def run_in_thread(self, target_function):
        thread = threading.Thread(target=target_function, daemon=True)
        thread.start()
        logger.info(f"Started thread for {target_function.__name__}")

    def copy_selected_cell(self, event=None):
        selected_item = self.treeview.selection()
        if selected_item:
            cell_values = self.treeview.item(selected_item)["values"]
            if cell_values:
                col_idx = self.treeview.identify_column(event.x)[1:]
                row_idx = int(selected_item[0]) - 1
                content = cell_values[int(col_idx) - 1]
                self.clipboard_clear()
                self.clipboard_append(content)
                self.update_idletasks()
                logger.info(f"Copied cell content: {content}")

    def modify_selected_paragraphs(self):
        doc = self.doc 
        modify_window = ctk.CTkToplevel(self)
        modify_window.title("Modify Paragraph Attributes")
        modify_window.geometry("400x250")

        inputs = {}
        attributes = ["Font Name", "Font Size", "Paragraph Style", "Hanging Indent", "First Line Indent"]
        for i, attr in enumerate(attributes):
            label = ctk.CTkLabel(modify_window, text=attr)
            label.grid(row=i, column=0, padx=10, pady=5, sticky="w")
            entry = ctk.CTkEntry(modify_window, placeholder_text=f"Enter {attr} (or leave empty)")
            entry.grid(row=i, column=1, padx=10, pady=5)
            inputs[attr] = entry

        def submit_changes():
            selected_items = self.treeview.selection()
            for item in selected_items:
                values = self.treeview.item(item, "values")
                para_number = values[0]

                try:
                    para_range_data = self.df.loc[self.df["Paragraph Number"] == int(para_number), "Paragraph Range"].values[0]
                    para_range_start = para_range_data[0]
                    para_range_end = para_range_data[1]
                    para_range = doc.Range(Start=para_range_start, End=para_range_end)
                    para_range.Select()
                    if inputs["Font Name"].get():
                        para_range.Font.Name = inputs["Font Name"].get()
                    if inputs["Font Size"].get():
                        para_range.Font.Size = float(inputs["Font Size"].get())
                    if inputs["Paragraph Style"].get():
                        para_range.Style = inputs["Paragraph Style"].get()
                    if inputs["Hanging Indent"].get():
                        para_range.ParagraphFormat.LeftIndent = float(float(inputs["Hanging Indent"].get()) * 28.346)
                    if inputs["First Line Indent"].get():
                        para_range.ParagraphFormat.FirstLineIndent = float(float(inputs["First Line Indent"].get()) * 28.346)

                    # Update DataFrame with new changes
                    self.df.loc[self.df["Paragraph Number"] == int(para_number), "Font Name"] = para_range.Font.Name
                    self.df.loc[self.df["Paragraph Number"] == int(para_number), "Font Size"] = para_range.Font.Size
                    self.df.loc[self.df["Paragraph Number"] == int(para_number), "Paragraph Style"] = para_range.Style
                    self.df.loc[self.df["Paragraph Number"] == int(para_number), "Hanging Indent"] = para_range.ParagraphFormat.LeftIndent / 28.346
                    self.df.loc[self.df["Paragraph Number"] == int(para_number), "First Line Indent"] = para_range.ParagraphFormat.FirstLineIndent / 28.346

                    logger.info(f"Applied changes to paragraph {para_number}")

                except Exception as e:
                    logger.error(f"Error modifying paragraph {para_number}: {e}")

            self.update_display()
            modify_window.destroy()

        submit_button = ctk.CTkButton(modify_window, text="Submit", command=submit_changes)
        submit_button.grid(row=len(attributes), columnspan=2, pady=10)

    def goto_paragraph(self, tree):
        doc = self.doc 
        try:
            selected_items = tree.selection()
            if not selected_items:
                messagebox.showinfo("Info", "No row selected.")
                return

            selected_item = selected_items[0]
            selected_values = tree.item(selected_item, "values")
            para_number = selected_values[0]

            if not para_number.isdigit():
                messagebox.showerror("Error", f"Invalid Paragraph Number: {para_number}")
                return

            para_number = int(para_number)
            para_range_data = self.df.loc[self.df["Paragraph Number"] == para_number, "Paragraph Range"].values
            if not len(para_range_data):
                messagebox.showerror("Error", "No matching paragraph range found in the DataFrame.")
                return

            para_range = para_range_data[0]
            para_range_start = para_range[0]
            para_range_end = para_range[1]
            doc.Range(Start=para_range_start, End=para_range_end).Select()
            logger.info(f"Moved to paragraph {para_number}")

        except IndexError:
            messagebox.showerror("Error", "Invalid selection. No matching range found.")
            logger.error("Invalid selection. No matching range found.")
        except Exception as e:
            logger.error(f"Error in goto_paragraph: {e}")

