from tkinter import filedialog, simpledialog
import pywintypes
import re
import pandas as pd
from src.shared_objects import WordApp
import tkinter as tk
from tkinter import ttk, messagebox
from src.logger import setup_logging

logger = setup_logging('word_tab.log')

class WordTab:
    """
    Initialize the WordTab instance.
    Opens a new Word document if none are open, otherwise uses the active document.
    """
    def __init__(self, tree):
        self.word_app_instance = WordApp()
        if self.word_app_instance.word_app.Documents.Count < 1:
            self.doc = self.word_app_instance.word_app.Documents.Add()
        elif self.word_app_instance.file_name != self.word_app_instance.word_app.ActiveDocument.Name:
            self.doc = self.word_app_instance.word_app.ActiveDocument
        self.tree = tree
        self.ranges = []
        self.match_positions = []
        self.initialize_word_document()

    def check_word_selection(self, regex_entry, update_treeview):
        try:
            selection = self.word_app_instance.word_app.Selection
            selected_text = selection.Text.strip()
            logger.info(f"The selected text is: {selected_text}")
            logger.info(f"The regex entry is: {regex_entry}")
            regex_pattern = regex_entry

            if regex_pattern:  # If regex is provided, use regex
                pattern = regex_pattern
            elif selected_text:  # If no regex, use the selected text
                selected_text = re.escape(selected_text)
                pattern = fr"{selected_text}"
                logger.info(f"The pattern is: {pattern}")
            else:
                messagebox.showinfo("Info", "No regex pattern or selected text provided.")
                return

            match_data = []
            for para in self.doc.Paragraphs:
                para_text = para.Range.Text.replace("\r", "\n")
                # Find all matches in the paragraph's text
                matches = re.finditer(pattern, para_text)
                for match in matches:
                    match_start = match.start()
                    match_end = match.end()
                    # Get the start and end positions in the document
                    para_start_pos = para.Range.Start
                    para_end_pos = para.Range.End
                    # Calculate the start and end positions of the match relative to the document
                    match_range_start = para_start_pos + match_start
                    match_range_end = para_start_pos + match_end

                    # Append the range to self.ranges
                    self.ranges.append((match_range_start, match_range_end))

                    # Select the matched range in the Word document
                    match_range = self.doc.Range(Start=match_range_start, End=match_range_end)
                    match_range.Select()
                    paragraph_range = self.doc.Range(Start=para_start_pos, End=para_end_pos)
                    paragraph_range.Select()
                    paragraph = paragraph_range.Paragraphs(1).Range.Text
                    logger.info(f"The paragraph is: {paragraph}")
                    style = paragraph_range.Paragraphs(1).Style.NameLocal
                    logger.info(f"The style is: {style}")
                    first_line_indent = round(paragraph_range.ParagraphFormat.FirstLineIndent * 0.0352778, 2)
                    left_indent = round(paragraph_range.ParagraphFormat.LeftIndent * 0.0352778, 2)
                    right_indent = round(paragraph_range.ParagraphFormat.RightIndent * 0.0352778, 2)
                    logger.info(f"The first line indent is: {first_line_indent}, left indent is: {left_indent}, right indent is: {right_indent}")
                    match_data.append({
                        'Match': paragraph,
                        'Style': style,
                        'Paragraph': match.group(),
                        'First Line Indent': first_line_indent,
                        'Left Indent': left_indent,
                        'Right Indent': right_indent
                    })
                    logger.info(f"Match data: {match_data}")
                    self.match_positions.append((match_range_start, match_range_end, match.group()))
            df = pd.DataFrame(match_data)
            update_treeview(df)

        except pywintypes.com_error as e:
            if e.args[0] == -2147417848:  # Check for the specific error code
                self.doc = self.word_app_instance.word_app.ActiveDocument

        except Exception as e:
            logger.error(f"Error in check_word_selection: {e}")
            messagebox.showinfo("Error", f"Error in check_word_selection: {e}")

    def goto_paragraph(self, tree):
        try:
            # Get the selected row
            selected_items = tree.selection()
            if not selected_items:  # Check if any row is selected
                messagebox.showinfo("Info", "No row selected.")
                return
            # Get the first item in the selection list
            selected_index = int(selected_items[0])  # Treeview selection returns a tuple
            # Ensure that the selected index is within bounds of match_positions
            if selected_index >= len(self.match_positions):
                messagebox.showerror("Error", "Selected index out of range.")
                return

            # Get the start and end position of the match
            start_pos, end_pos, matched_string = self.match_positions[selected_index]
            logger.info(f"Matched text: {matched_string}")

            # Select the range in the Word document
            new_range = self.doc.Range(Start=start_pos, End=end_pos)
            new_range.Select()
            if new_range.Text != matched_string:
                paragraph_text = self.doc.Range(Start=start_pos, End=end_pos).Paragraphs(1).Range.Text
                selection = self.word_app_instance.word_app.Selection
                new_start_pos = selection.Paragraphs(1).Range.Start
                if matched_string in paragraph_text:
                    s = paragraph_text.find(matched_string)
                    new_range = selection.Paragraphs(1).Range
                    new_range.Start = new_start_pos + s + 1
                    new_range.End = new_start_pos + s + len(matched_string) + 1
                    new_range.Select()
            else:
                new_range.Select()
                pass
        except IndexError:
            messagebox.showerror("Error", "No valid row selected.")
        except Exception as e:
            logger.error(f"Error in goto_paragraph: {e}")

    def update_indents_and_tabs(self, selected_index, left_indent_entry, first_line_indent_entry, right_indent_entry, tab_entry, new_text_entry=None):
        """
        Updates the selected paragraph's indents, tab stops, and text content in the Word document.

        Args:
        - selected_index: Index of the selected match/paragraph in the document.
        - left_indent_entry, first_line_indent_entry, right_indent_entry, tab_entry: Entry widgets for indents and tabs.
        - new_text_entry: Optional Entry widget for updating the paragraph text.
        """
        try:
            # Get the start and end positions of the selected paragraph/match
            start_pos, end_pos, _ = self.match_positions[selected_index]

            # Retrieve indent and tab values from the Entry widgets
            left_indent = float(left_indent_entry.get())
            first_line_indent = float(first_line_indent_entry.get())
            right_indent = float(right_indent_entry.get())
            tabs_input = [float(tab) for tab in tab_entry.get().split(',')]

            # Get the range of the selected paragraph in the Word document
            selection_range = self.doc.Range(Start=start_pos, End=end_pos)

            # Update paragraph indents
            paragraph_format = selection_range.ParagraphFormat
            paragraph_format.LeftIndent = left_indent
            paragraph_format.FirstLineIndent = first_line_indent
            paragraph_format.RightIndent = right_indent

            # Update paragraph tab stops
            paragraph_format.TabStops.ClearAll()  # Clear existing tab stops
            for tab in tabs_input:
                paragraph_format.TabStops.Add(Position=tab)  # Add new tab stops

            # Optionally update the text content of the paragraph
            if new_text_entry:
                new_text = new_text_entry.get()
                selection_range.Text = new_text  # Replace the text in the selected range with new text

        except Exception as e:
            logger.error(f"Error applying indents, tabs, or updating text: {e}")

    def on_double_click(event, tree, word_tab):
        """Handle double-click event on the Treeview to allow in-place editing."""
        # Get selected item
        item_id = tree.selection()[0]
        column = tree.identify_column(event.x)
        column_index = int(column.replace('#', '')) - 1  # Convert column index from 1-based to 0-based
        row = tree.identify_row(event.y)

        # Get item values
        values = tree.item(item_id, 'values')

        # Create an entry widget for editing
        entry_widget = tk.Entry(tree)
        entry_widget.insert(0, values[column_index])  # Insert the current value into the entry field

        # Place entry widget over the selected cell
        entry_widget.place(x=tree.bbox(item_id, column)[0], y=tree.bbox(item_id, column)[1],
                           width=tree.bbox(item_id, column)[2], height=tree.bbox(item_id, column)[3])

        # Focus the entry widget and bind the event to update the value on Enter key press
        entry_widget.focus()

        # Bind the "Enter" key event to call the update function
        entry_widget.bind("<Return>", lambda event: word_tab.on_enter(event, tree, item_id, column_index, entry_widget, word_tab))

        # Bind the focus-out event to remove the entry widget if the user clicks outside
        entry_widget.bind("<FocusOut>", lambda event: word_tab.on_focus_out(event, entry_widget))

    def on_focus_out(self, event, entry_widget):
        """Destroy the entry widget when it loses focus."""
        entry_widget.destroy()

    def on_enter(self, event, tree, item_id, column_index, entry_widget, word_tab):
        """When Enter key is pressed, update the value, call the appropriate update method, and remove the entry widget."""

        # Get the current values of the selected row
        values = tree.item(item_id, 'values')

        # Update the Treeview item with the new value
        new_value = entry_widget.get()  # Value from the entry widget
        new_values = list(values)
        new_values[column_index] = new_value  # Update the specific column
        tree.item(item_id, values=tuple(new_values))  # Set the updated values

        # Identify what was changed (text, style, or indents)
        if column_index == 0:  # Assume column 0 is for text
            word_tab.update_text(new_value, item_id)
        elif column_index == 1:  # Assume column 1 is for style
            word_tab.update_style(new_value, item_id, word_tab)

        entry_widget.destroy()

    def update_text(self, new_text, item_id):
        """Update the text in the Word document for the selected paragraph."""
        try:
            selected_index = int(item_id)
            start_pos, end_pos, _ = self.match_positions[selected_index]

            # Select the range in the Word document
            selection_range = self.doc.Range(Start=start_pos, End=end_pos)
            new_s_pos = selection_range.Paragraphs(1).Range.Start
            new_e_pos = selection_range.Paragraphs(1).Range.End
            paragraph_range = self.doc.Range(Start=new_s_pos, End=new_e_pos - 1)
            # find the difference between two paragraph range and the new text
            paragraph_range.Text = new_text[:len(new_text) - 1]
            # Update the text

        except Exception as e:
            logger.error(f"Error updating text: {e}")

    def update_style(self, new_style, item_id, word_tab):
        """Update the style of the selected paragraph."""
        try:
            selected_index = int(item_id)
            start_pos, end_pos, _ = word_tab.match_positions[selected_index]

            # Select the range in the Word document
            selection_range = word_tab.doc.Range(Start=start_pos, End=end_pos)

            # Update the style
            selection_range.Select()
            selection_range.Style = word_tab.doc.Styles(new_style)

        except Exception as e:
            logger.error(f"Error updating style: {e}")

    def on_row_select(self, event, tree):
        """Handle row selection based on mouse click."""
        # Clear current selection
        for item in tree.selection():
            tree.selection_remove(item)

        # Select the item clicked
        row_id = tree.identify_row(event.y)
        if row_id:
            tree.selection_set(row_id)
        else:
            tree.selection_clear()

    def replace_values_in_selected_rows(self, tree, new_value, word_tab):
        """Replace values in the second column of all selected rows with the new value and update styles."""
        selected_items = tree.selection()
        if selected_items:
            for item_id in selected_items:
                # Get current values of the row (as a list for easy manipulation)
                current_values = list(tree.item(item_id, 'values'))
                # Update the second column (index 1) with the new value
                current_values[1] = new_value
                # Set the updated values back into the tree item
                tree.item(item_id, values=tuple(current_values))
                # Call the function to update the style in the Word document
                self.update_style(new_value, item_id, word_tab)

    def handle_entry_change(self, event, tree, entry_widget, word_tab):
        """Handle the change in the entry widget and update the specific row and column."""
        # Get the selected item ID
        item_id = tree.selection()[0]

        # Identify the column based on the event's x position
        column = tree.identify_column(event.x)
        column_index = int(column.replace('#', '')) - 1  # Convert column index from 1-based to 0-based

        # Get the value from the entry widget
        new_value = entry_widget.get()

        # Update only the specific column in the selected row
        values = list(tree.item(item_id, 'values'))  # Get current values as a list
        values[column_index] = new_value  # Update the specific column
        tree.item(item_id, values=tuple(values))  # Update the item with new values

        # Call the function to update the style in the Word document
        self.update_style(new_value, item_id, word_tab)

    def prompt_for_value_and_replace(self, tree, word_tab):
        """Prompt the user for a value and replace the second column in the selected rows, and update styles."""
        new_value = simpledialog.askstring("Input", "Enter a value to replace the second column in the selected rows with:")
        if new_value is not None:
            self.replace_values_in_selected_rows(tree, new_value, word_tab)

    def initialize_word_document(self):
        """Try to initialize the Word document for the WordTab instance.

        This method tries to access the active Word document and assigns it to the
        `doc` attribute. If no active document is found, it raises an exception and
        displays an error message.

        :raises Exception: If unable to access the active Word document.
        """
        try:
            self.doc = self.word_app_instance.word_app.ActiveDocument
            if self.doc is None:
                raise Exception("No active document found.")
        except Exception as e:
            logger.error(f"Error initializing Word document: {e}")
            messagebox.showerror("Error", "Unable to access the active Word document.")
        except pywintypes.com_error as e:
            if e.args[0] == -2147417848:  # Check for the specific error code
                self.doc = self.word_app_instance.word_app.ActiveDocument
        except Exception as e:
            logger.error(f"Error: {e}")
            messagebox.showinfo("Error", f"Error in check_word_selection: {e}")

