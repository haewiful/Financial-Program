'''
import tkinter as tk
from tkinter import messagebox, filedialog
import tkinter.ttk as ttk # For the Treeview widget
import os
import sys

# Import custom classes from other files
try:
    from excel_generator import ExcelGenerator
    from word_report import WordReportGenerator
except ImportError as e:
    print(f"Error importing modules. Ensure all three files are present. Detail: {e}")
    sys.exit(1)

DATA_HEADERS = ("부서", "항목", "입금", "출금")

class MainApplication(tk.Tk):
    """The main GUI class for the file generation program using tkinter."""

    def __init__(self):
        super().__init__()
        self.title("Document Automation Tool (macOS)")
        
        # --- 1. Define Window Size ---
        self.minsize(width=400, height=250)
        self.resizable(False, False)
        
        # --- 2. Calculate Screen Center Coordinates ---
        self.update_idletasks()
        window_width = self.winfo_reqwidth()
        window_height = self.winfo_reqheight()

        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        center_x = int((screen_width / 2) - (window_width / 2))
        center_y = int((screen_height / 2) - (window_height / 2))
        
        # --- 3. Set Window Position ---
        self.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        
        self.current_excel_generator = None 
        self.preview_tree = None # Placeholder for the Treeview widget

        self.create_widgets()

    def create_widgets(self):
        """Sets up the visual components (buttons and labels) in the main window."""
        
        title_label = tk.Label(
            self, 
            text="Document Automation Dashboard", 
            font=('Arial', 16, 'bold'),
            pady=10
        )
        title_label.pack()

        button_frame = tk.Frame(self)
        button_frame.pack(pady=20)

        # TODO when clicked multiple times, don't create a new window but just focus on the previous one
        excel_button = tk.Button(
            button_frame, 
            text="1. Create & Edit New Excel File", 
            command=self.open_excel_window,
            width=35,
            height=2,
            bg='#4CAF50', 
            fg='black'
        )
        excel_button.pack(pady=10)

        word_button = tk.Button(
            button_frame, 
            text="2. Generate Word Report from Excel", 
            command=self.generate_word_report_gui,
            width=35,
            height=2,
            bg='#2196F3', 
            fg='black'
        )
        word_button.pack(pady=10)

    # ------------------------------------------------------------------
    # --- Action 1: Create Excel File (Opens a secondary window) ---
    # ------------------------------------------------------------------

    def open_excel_window(self):
        """Opens a new top-level window for Excel data entry and live preview."""
        
        self.current_excel_generator = ExcelGenerator(DATA_HEADERS)

        excel_win = tk.Toplevel(self)
        excel_win.title("Excel Data Entry & Live Preview")
        
        entry_widgets = {}
        fields = DATA_HEADERS
        
        # --- Frame for Input Fields ---
        input_frame = tk.Frame(excel_win)
        input_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)
        
        for i, field_name in enumerate(fields):
            tk.Label(input_frame, text=f"{field_name}:", anchor="w").grid(row=0, column=i*2, padx=5, sticky="w")
            entry = tk.Entry(input_frame, width=15)
            entry.grid(row=0, column=i*2 + 1, padx=5)
            entry_widgets[field_name] = entry

        # --- Add Row Button ---
        add_button = tk.Button(input_frame, 
                               text="Add Row", 
                               command=lambda: self.add_row_gui(entry_widgets))
        add_button.grid(row=0, column=len(fields)*2, padx=10)
        
        # --- Live Preview Table (ttk.Treeview) ---
        table_frame = tk.Frame(excel_win)
        table_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.preview_tree.bind('<Button-1>', self.on_treeview_click)

        columns = ("#", "dept", "entry", "deposit", "withdrawal")
        self.preview_tree = ttk.Treeview(
            table_frame, 
            columns=columns, 
            show='headings',
            height=10
        )
        
        # Configure scrollbar
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.preview_tree.yview)
        vsb.pack(side='right', fill='y')
        self.preview_tree.configure(yscrollcommand=vsb.set)

        # Define column headings and widths
        self.preview_tree.heading("#", text="IDX")
        self.preview_tree.column("#", width=50, anchor='center')
        
        self.preview_tree.heading("dept", text="부서")
        self.preview_tree.column("dept", width=100)
        
        self.preview_tree.heading("entry", text="항목")
        self.preview_tree.column("entry", width=100, anchor='center')
        
        self.preview_tree.heading("deposit", text="입금")
        self.preview_tree.column("deposit", width=100, anchor='center')

        self.preview_tree.heading("withdrawal", text="출금")
        self.preview_tree.column("withdrawal", width=100, anchor='center')
        
        self.preview_tree.pack(fill='both', expand=True)

        # --- Footer Frame ---
        footer_frame = tk.Frame(excel_win)
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=5)
        
        # --- Save Button ---
        save_button = tk.Button(footer_frame, 
                                text="Save Excel File", 
                                command=lambda: self.save_excel_file_gui(excel_win),
                                bg='#FF9800', fg='black')
        save_button.pack(pady=5)

        # --- Calculate Size and Center the New Window ---
        excel_win.update_idletasks()
        
        # Get the required size (this is now the size dictated by the components)
        win_width = excel_win.winfo_reqwidth()
        win_height = excel_win.winfo_reqheight()
        # Get screen size from the parent window
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        center_x = int((screen_width / 2) - (win_width / 2))
        center_y = int((screen_height / 2) - (win_height / 2))
        
        # Set the final geometry (size and position)
        excel_win.geometry(f"{win_width}x{win_height}+{center_x}+{center_y}")


    def add_row_gui(self, entry_widgets):
        """Calls the backend to add a row and updates the Treeview preview."""
        dept = entry_widgets["부서"].get()
        entry = entry_widgets["항목"].get()
        deposit = entry_widgets["입금"].get()
        withdrawal = entry_widgets["출금"].get()
        
        if dept and entry and (deposit or withdrawal):
            try:
                # 1. Call the generator's method
                self.current_excel_generator.add_data_row(dept, entry, deposit, withdrawal)
                
                # 2. Update the Treeview
                self.update_treeview_preview()
                
                # 3. Clear fields
                for entry in entry_widgets.values():
                    entry.delete(0, tk.END)
                
            except ValueError as e:
                messagebox.showerror("Input Error", str(e))
            except Exception as e:
                messagebox.showerror("Error", f"An unexpected error occurred: {e}")
        else:
            messagebox.showerror("Error", "All fields must be filled.")

    def update_treeview_preview(self):
        """Refreshes the Treeview with the current data from the ExcelGenerator."""
        
        # 1. Clear all existing items in the Treeview
        for dept in self.preview_tree.get_children():
            self.preview_tree.delete(dept)

        # 2. Iterate through the generator's data and insert into the Treeview
        sheet = self.current_excel_generator.sheet
        data_row_index = 1 # User-facing index

        # Start from row 2 (skipping the header row 1)
        for row in sheet.iter_rows(min_row=2, values_only=True): 
            
            # The Treeview insert takes the row index as the first value for display
            display_values = [data_row_index] + list(row)
            
            self.preview_tree.insert('', tk.END, values=display_values)
            data_row_index += 1
            
    def save_excel_file_gui(self, window_to_close):
        """Prompts for a file name and saves the generated Excel file."""
        if self.current_excel_generator is None:
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save Excel File As"
        )

        if file_path:
            if self.current_excel_generator.save_file(file_path):
                messagebox.showinfo("Success", f"File saved successfully to:\n{os.path.basename(file_path)}")
                window_to_close.destroy()
            else:
                messagebox.showerror("Error", "Failed to save the Excel file. Check permissions.")

    # ------------------------------------------------------------------
    # --- Action 2: Generate Word Report ---
    # ------------------------------------------------------------------

    def generate_word_report_gui(self):
        """Prompts for an Excel file and generates the Word report."""
        
        excel_file_path = filedialog.askopenfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Select Excel File for Report"
        )
        
        if not excel_file_path:
            return # User cancelled selection

        report_maker = WordReportGenerator()
        
        try:
            # generate_report now returns the save path
            saved_doc_path = report_maker.generate_report(excel_file_path)
            
            messagebox.showinfo(
                "Success", 
                f"Word Report generated successfully!\nSaved as: {os.path.basename(saved_doc_path)}"
            )
        except (FileNotFoundError, ValueError) as e:
            messagebox.showerror("Generation Error", str(e))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate Word Report: {e}")
            
# --- Main Execution Block ---
if __name__ == "__main__":
    app = MainApplication()
    app.mainloop()

'''


import tkinter as tk
from tkinter import messagebox, filedialog
import tkinter.ttk as ttk # For the Treeview widget
import os
import sys

# Import custom classes from other files
try:
    # Ensure ExcelGenerator and WordReportGenerator methods match the new 4-column structure!
    from excel_generator import ExcelGenerator
    from word_report import WordReportGenerator
except ImportError as e:
    print(f"Error importing modules. Ensure all three files are present. Detail: {e}")
    sys.exit(1)

# GLOBAL CONFIGURATION VARIABLE
DATA_HEADERS = ("부서", "항목", "입금", "출금")


class MainApplication(tk.Tk):
    """The main GUI class for the file generation program using tkinter."""

    def __init__(self):
        super().__init__()
        self.title("Document Automation Tool (macOS)")
        
        # --- 1. Define Window Size ---
        self.minsize(width=400, height=250)
        self.resizable(False, False)
        
        # --- 2. Calculate Screen Center Coordinates ---
        self.update_idletasks()
        window_width = self.winfo_reqwidth()
        window_height = self.winfo_reqheight()

        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        center_x = int((screen_width / 2) - (window_width / 2))
        center_y = int((screen_height / 2) - (window_height / 2))
        
        # --- 3. Set Window Position ---
        self.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        
        # State variables
        self.current_excel_generator = None
        self.preview_tree = None 
        # 🟢 State variable to track the Excel editing window
        self.excel_toplevel_window = None 

        self.create_widgets()

    def create_widgets(self):
        """Sets up the visual components (buttons and labels) in the main window."""
        
        title_label = tk.Label(
            self, 
            text="Document Automation Dashboard", 
            font=('Arial', 16, 'bold'),
            pady=10,
            fg='black'
        )
        title_label.pack()

        button_frame = tk.Frame(self)
        button_frame.pack(pady=20)

        excel_button = tk.Button(
            button_frame, 
            text="1. Create & Edit New Excel File", 
            command=self.open_excel_window,
            width=35,
            height=2,
            bg='#4CAF50', 
            fg='black'
        )
        excel_button.pack(pady=10)

        word_button = tk.Button(
            button_frame, 
            text="2. Generate Word Report from Excel", 
            command=self.generate_word_report_gui,
            width=35,
            height=2,
            bg='#2196F3', 
            fg='black'
        )
        word_button.pack(pady=10)

    # ------------------------------------------------------------------
    # --- Action 1: Create Excel File (Opens a secondary window) ---
    # ------------------------------------------------------------------

    def open_excel_window(self):
        """Opens a new top-level window for Excel data entry and live preview."""
        
        # 🟢 FIX 1: Check if window already exists and focus it
        if self.excel_toplevel_window and self.excel_toplevel_window.winfo_exists():
            self.excel_toplevel_window.lift() # Bring to front
            return
            
        self.current_excel_generator = ExcelGenerator(DATA_HEADERS)

        excel_win = tk.Toplevel(self)
        self.excel_toplevel_window = excel_win # Store reference
        excel_win.title("Excel Data Entry & Live Preview")
        # 🟢 Set behavior on close to destroy reference
        excel_win.protocol("WM_DELETE_WINDOW", lambda: self.close_excel_window(excel_win))

        
        entry_widgets = {}
        fields = DATA_HEADERS
        
        # --- Frame for Input Fields ---
        input_frame = tk.Frame(excel_win)
        input_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)
        
        current_col = 0
        for field_name in fields:
            tk.Label(input_frame, text=f"{field_name}:", anchor="w").grid(row=0, column=current_col, padx=5, sticky="w")
            current_col += 1
            entry = tk.Entry(input_frame, width=15)
            entry.grid(row=0, column=current_col, padx=(0, 15)) # Increased padx to prevent cutting off Add Row
            entry_widgets[field_name] = entry
            current_col += 1

        # --- Add Row Button ---
        add_button = tk.Button(input_frame, 
                               text="Add Row", 
                               command=lambda: self.add_row_gui(entry_widgets))
        add_button.grid(row=0, column=current_col, padx=10)
        
        # --- Live Preview Table (ttk.Treeview) ---
        table_frame = tk.Frame(excel_win)
        table_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=5)

        columns = ("#", "dept", "entry", "deposit", "withdrawal")
        
        # 🟢 FIX 2: Define self.preview_tree BEFORE binding events
        self.preview_tree = ttk.Treeview(
            table_frame, 
            columns=columns, 
            show='headings',
            height=10
        )
        
        # 🟢 FIX 3: Bind the click event AFTER defining the Treeview
        self.preview_tree.bind('<Button-1>', self.on_treeview_click)
        
        # Configure scrollbar
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.preview_tree.yview)
        vsb.pack(side='right', fill='y')
        self.preview_tree.configure(yscrollcommand=vsb.set)

        # Define column headings and widths
        self.preview_tree.heading("#", text="IDX")
        self.preview_tree.column("#", width=50, anchor='center')
        
        # Adjusted Treeview widths for better fit
        self.preview_tree.heading("dept", text=DATA_HEADERS[0]); self.preview_tree.column("dept", width=150)
        self.preview_tree.heading("entry", text=DATA_HEADERS[1]); self.preview_tree.column("entry", width=120, anchor='center')
        self.preview_tree.heading("deposit", text=DATA_HEADERS[2]); self.preview_tree.column("deposit", width=120, anchor='center')
        self.preview_tree.heading("withdrawal", text=DATA_HEADERS[3]); self.preview_tree.column("withdrawal", width=120, anchor='center')
        
        self.preview_tree.pack(fill='both', expand=True)

        # --- Footer Frame ---
        footer_frame = tk.Frame(excel_win)
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=5)
        
        # --- Calculate Size and Center the New Window ---
        excel_win.update_idletasks()
        win_width = excel_win.winfo_reqwidth()
        win_height = excel_win.winfo_reqheight()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        center_x = int((screen_width / 2) - (win_width / 2))
        center_y = int((screen_height / 2) - (win_height / 2))
        
        excel_win.geometry(f"{win_width}x{win_height}+{center_x}+{center_y}")
        
        # --- Save Button ---
        save_button = tk.Button(footer_frame, 
                                 text="Save Excel File", 
                                 command=lambda: self.save_excel_file_gui(excel_win),
                                 bg='#FF9800', fg='black')
        save_button.pack(pady=5) # Reduced pady to fit better in footer frame

    def close_excel_window(self, window):
        """Handles the window close event to clear the window reference."""
        # Reset the reference when the window is closed
        self.excel_toplevel_window = None
        window.destroy()

    def add_row_gui(self, entry_widgets):
        """Calls the backend to add a row and updates the Treeview preview."""
        dept = entry_widgets[DATA_HEADERS[0]].get()
        entry = entry_widgets[DATA_HEADERS[1]].get()
        deposit = entry_widgets[DATA_HEADERS[2]].get()
        withdrawal = entry_widgets[DATA_HEADERS[3]].get()
        
        if dept and entry and (deposit or withdrawal):
            try:
                self.current_excel_generator.add_data_row(dept, entry, deposit, withdrawal)
                
                self.update_treeview_preview()
                
                for widget in entry_widgets.values():
                    widget.delete(0, tk.END)
                
            except ValueError as e:
                messagebox.showerror("Input Error", str(e))
            except Exception as e:
                messagebox.showerror("Error", f"An unexpected error occurred: {e}")
        else:
            messagebox.showerror("Error", "All fields must be filled.")

    # --- Treeview Editing Methods ---
    
    def on_treeview_click(self, event):
        """Handler for Treeview clicks to start cell editing."""
        if not self.preview_tree:
            return

        item = self.preview_tree.identify_row(event.y)
        column_id = self.preview_tree.identify_column(event.x)
        
        # Check if we clicked a valid item and an editable column
        # Editable columns start at #2 (dept)
        if not item or column_id == '#0' or column_id == '#1': 
            return
        
        # Get the 1-based index of the column (1 for IDX, 2 for dept, etc.)
        column_index = int(column_id.replace('#', '')) - 1
        
        # Get the user-facing row index from the 'values' list (first element)
        user_idx_str = self.preview_tree.item(item, 'values')[0]
        user_row_index = int(user_idx_str)
        
        # Get the bounding box of the clicked cell
        bbox = self.preview_tree.bbox(item, column_id)
        if not bbox:
            return

        self.start_cell_editor(item, column_index, user_row_index, bbox)
        
    def start_cell_editor(self, item, column_index, user_row_index, bbox):
        """Creates a temporary entry widget for cell editing."""
        
        # Column index 1 is 'IDX', so subtract 1 to get the index for DATA_HEADERS
        data_header_index = column_index - 1 
        col_name = DATA_HEADERS[data_header_index]
        
        # Get current cell value (column_index is the list index in 'values')
        current_value = self.preview_tree.item(item, 'values')[column_index]
        
        # Create a temporary Entry widget
        editor = tk.Entry(self.preview_tree, bd=0, bg='white', highlightthickness=1, highlightcolor="blue")
        editor.insert(0, current_value)
        editor.select_range(0, tk.END)
        editor.focus()
        
        # Position the editor over the cell (bbox = (x, y, width, height))
        editor.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])

        def on_editor_confirm(event):
            """Saves the new value when Enter is pressed."""
            new_value = editor.get()
            
            try:
                # Call the Excel Generator to update the backend data
                # Requires 'update_data_cell' in excel_generator.py
                self.current_excel_generator.update_data_cell(
                    user_row_index, 
                    col_name, 
                    new_value
                )
                
                # Update the Treeview and clean up
                self.update_treeview_preview()
                editor.destroy()
                
            except (ValueError, Exception) as e:
                messagebox.showerror("Validation Error", f"Update failed for '{col_name}'.\nDetails: {str(e)}")
                editor.destroy()

        def on_editor_lose_focus():
            """Destroys the editor if the user clicks away or presses Escape."""
            editor.destroy()

        # Bind events
        editor.bind('<Return>', on_editor_confirm) # Enter key
        editor.bind('<Escape>', lambda e: editor.destroy()) # Escape key

        # Bind FocusOut slightly later to allow the initial click to register
        self.after(100, lambda: editor.bind('<FocusOut>', lambda e: on_editor_lose_focus()))
        
    # --- Other Methods ---

    def update_treeview_preview(self):
        """Refreshes the Treeview with the current data from the ExcelGenerator."""
        
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)

        sheet = self.current_excel_generator.sheet
        data_row_index = 1 

        for row in sheet.iter_rows(min_row=2, values_only=True): 
            display_values = [data_row_index] + list(row)
            self.preview_tree.insert('', tk.END, values=display_values)
            data_row_index += 1
            
    def save_excel_file_gui(self, window_to_close):
        """Prompts for a file name and saves the generated Excel file."""
        if self.current_excel_generator is None:
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save Excel File As"
        )

        if file_path:
            if self.current_excel_generator.save_file(file_path):
                messagebox.showinfo("Success", f"File saved successfully to:\n{os.path.basename(file_path)}")
                self.close_excel_window(window_to_close) # Use the clean close method
            else:
                messagebox.showerror("Error", "Failed to save the Excel file. Check permissions.")

    def generate_word_report_gui(self):
        """Prompts for an Excel file and generates the Word report."""
        
        excel_file_path = filedialog.askopenfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Select Excel File for Report"
        )
        
        if not excel_file_path:
            return # User cancelled selection

        report_maker = WordReportGenerator()
        
        try:
            saved_doc_path = report_maker.generate_report(excel_file_path)
            
            messagebox.showinfo(
                "Success", 
                f"Word Report generated successfully!\nSaved as: {os.path.basename(saved_doc_path)}"
            )
        except (FileNotFoundError, ValueError) as e:
            messagebox.showerror("Generation Error", str(e))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate Word Report: {e}")
            
# --- Main Execution Block ---
if __name__ == "__main__":
    app = MainApplication()
    app.mainloop()