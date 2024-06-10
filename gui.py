import os.path
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from whatsapp_message import start_whatsapp_messaging
from make_logs import get_logger


class ExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to Whatsapp Bot")
        self.root.geometry("600x400")
        self.root.configure(bg="gray")

        self.file_path = ""
        self.selected_sheets = []
        
        # Logger
        self.start_logger()

        # Widgets
        self.create_widgets()

    def start_logger(self):
        self.logger = get_logger()

    def create_widgets(self):
        # Button to select Excel file
        self.select_file_button = tk.Button(self.root, text="Select Excel File", command=self.select_file,
                                            bg="#4795ed", fg="white")
        self.select_file_button.pack(pady=10)

        # Label for displaying selected file
        self.file_label = tk.Label(self.root, text="No file selected", wraplength=600, fg="white", bg="gray")
        self.file_label.pack(pady=10)

        # Listbox for sheet names with multiple selection enabled
        self.sheet_listbox = tk.Listbox(self.root, selectmode=tk.MULTIPLE, height=10, bg="white")
        self.sheet_listbox.pack(pady=10, fill=tk.BOTH, expand=True)

        # Button to confirm sheet selection
        self.confirm_button = tk.Button(self.root, text="Confirm Selection", command=self.confirm_selection,
                                        bg="#4795ed", fg="white")
        self.confirm_button.pack(pady=10)

        # Text widget for displaying logs
        self.log_text = tk.Text(self.root, height=10, bg="black", fg="#5ba0f0")
        # self.log_text.pack(pady=10, fill=tk.BOTH, expand=True)

    def log(self, message):
        self.log_text.insert(tk.END, message + '\n')
        self.log_text.see(tk.END)  # Scroll to the end

    def select_file(self):
        # Open file dialog to select an Excel file
        self.file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")])

        if self.file_path:
            self.file_label.config(text=f"Selected File : {self.file_path}")
            self.log(f"Selected file: {self.file_path}")
            self.logger.log(f"Selected file: {self.file_path}")
            self.load_sheets()

    def load_sheets(self):
        try:
            # Read the Excel file to get sheet names
            excel_file = pd.ExcelFile(self.file_path)
            sheet_names = excel_file.sheet_names
            self.sheet_listbox.delete(0, tk.END)  # Clear previous entries
            for sheet in sheet_names:
                self.sheet_listbox.insert(tk.END, sheet)
            # self.log(f"Loaded sheets: {', '.join(sheet_names)} \n")
        except Exception as e:
            self.show_message("Error", f"Failed to read Excel file: {e}")
            self.log(f"Error: Failed to read Excel file: {e}")
            self.logger.log(f"Error: Failed to read Excel file: {e}")


    def on_messaging_complete(self, title, message):
        self.log("\n<<< MESSAGING PROCESS COMPLETE. CLICK X button TO CLOSE THE APP >>>")
        self.show_message(title, message)

    def confirm_selection(self):
        # Get selected items from the listbox
        selected_indices = self.sheet_listbox.curselection()
        self.selected_sheets = [self.sheet_listbox.get(i) for i in selected_indices]
        if not self.selected_sheets:
            self.show_message("Warning", "Please select at least one sheet.")
            self.log("Warning: No sheets selected.")
            self.logger.log("Warning: No sheets selected.")
            return

        # Remove buttons and listbox
        self.select_file_button.pack_forget()
        self.sheet_listbox.pack_forget()
        self.confirm_button.pack_forget()

        # Pack the log_text widget to display logs
        self.log_text.pack(pady=10, fill=tk.BOTH, expand=True)
        self.log(f"Selected Sheets : {self.selected_sheets}")
        self.logger.log(f"Selected Sheets : {self.selected_sheets}")


        # Call the external function and pass the log callback
        # new thread used to ensure logging is not blocked by whatsapp messaging thread
        threading.Thread(target=start_whatsapp_messaging,
                         args=(self.file_path, self.selected_sheets, self.log, self.on_messaging_complete)).start()

    def show_message(self, title, message):
        # Create a new top-level window
        messagebox_window = tk.Toplevel(self.root)
        messagebox_window.title(title)
        messagebox_window.geometry("300x150")
        messagebox_window.configure(bg="green")

        # Create a label for the message
        message_label = tk.Label(messagebox_window, text=message, fg="white", bg="green", wraplength=250)
        message_label.pack(pady=20, padx=10)

        # Create an OK button to close the message box
        ok_button = tk.Button(messagebox_window, text="OK", command=messagebox_window.destroy, bg="#4795ed", fg="white")
        ok_button.pack(pady=10)


# Create the main window
root = tk.Tk()
app = ExcelApp(root)
root.mainloop()
