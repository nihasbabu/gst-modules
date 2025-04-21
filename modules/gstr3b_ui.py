# gstr3b_ui.py

import tkinter as tk
from tkinter import filedialog, messagebox
import os

from telemetry import send_event
from gstr3b_processor import process_gstr3b

class GSTR3BProcessorUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Process GSTR3B Processing")
        self.root.geometry("500x480")
        self.json_files = []  # List of file paths
        self.template_file = None

        # Title
        tk.Label(root, text="Process GSTR3B Processing", font=("Arial", 16, "bold")).pack(pady=10)

        # Main Frame for List Box
        main_frame = tk.Frame(root)
        main_frame.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)

        # GSTR3B JSON Section
        tk.Label(main_frame, text="GSTR3B JSON", font=("Arial", 10, "bold")).pack()
        self.json_list = tk.Listbox(main_frame, height=15, width=60, selectmode=tk.EXTENDED)
        self.json_list.pack(pady=0, fill=tk.Y, expand=True)

        # Add/Remove Buttons
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(pady=5)
        tk.Button(btn_frame, text="+ Add", command=self.add_json_file).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="- Remove", command=self.delete_json_file).pack(side=tk.LEFT, padx=5)

        # Template File Section
        template_frame = tk.Frame(root)
        template_frame.pack(pady=5)
        tk.Label(template_frame, text="Template Excel File (Optional):").pack(side=tk.LEFT)
        self.template_label = tk.Label(template_frame, text="No file selected")
        self.template_label.pack(side=tk.LEFT, padx=5)
        tk.Button(template_frame, text="Select", command=self.select_template).pack(side=tk.LEFT, padx=2)
        tk.Button(template_frame, text="Clear", command=self.clear_template).pack(side=tk.LEFT, padx=2)

        # Process Button
        self.process_btn = tk.Button(
            root,
            text="Process GSTR3B",
            font=("Arial", 12),
            command=self.process_files,
            state=tk.DISABLED,
            bg="light grey"
        )
        self.process_btn.pack(pady=10)

        self.update_process_button()

    def add_json_file(self):
        files = filedialog.askopenfilenames(filetypes=[("JSON Files", "*.json")])
        financial_months = ["04","05","06","07","08","09","10","11","12","01","02","03"]

        for file in files:
            if file not in self.json_files:
                self.json_files.append(file)

        def sort_key(path):
            name = os.path.basename(path).lower()
            if name.endswith('.json'):
                name = name[:-5]
            parts = name.split('_')
            if len(parts) < 2:
                return (len(financial_months), 9999)
            last = parts[-1]
            if last == 'nf' and len(parts) > 2:
                last = parts[-2]
            if len(last) >= 6 and last[:6].isdigit():
                m, y = last[:2], last[2:6]
                if m in financial_months:
                    return (financial_months.index(m), int(y))
            return (len(financial_months), 9999)

        self.json_files.sort(key=sort_key)
        self.json_list.delete(0, tk.END)
        for file in self.json_files:
            self.json_list.insert(tk.END, os.path.basename(file))
        self.update_process_button()

    def delete_json_file(self):
        selected = self.json_list.curselection()
        if not selected:
            return
        for idx in reversed(selected):
            self.json_files.pop(idx)
            self.json_list.delete(idx)
        self.update_process_button()

    def select_template(self):
        file = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        if file:
            self.template_file = file
            self.template_label.config(text=os.path.basename(file))

    def clear_template(self):
        self.template_file = None
        self.template_label.config(text="No file selected")

    def update_process_button(self):
        if self.json_files:
            self.process_btn.config(state=tk.NORMAL, bg="light green")
        else:
            self.process_btn.config(state=tk.DISABLED, bg="light grey")

    def process_files(self):
        if not self.json_files:
            messagebox.showerror("Error", "No JSON files selected for processing.")
            return

        save_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            title="Save GSTR-3B Summary Report As"
        )
        if not save_file:
            return

        self.process_btn.config(text="Processing...", state=tk.DISABLED, bg="light grey")
        self.root.update()

        try:
            result = process_gstr3b(self.json_files, self.template_file, save_file)

            # Telemetry: success event
            send_event("gstr3b_complete", {
                "input_files": self.json_files,
                "output_file": save_file,
                "message": result
            })

            messagebox.showinfo("Success", result)
            self.json_files.clear()
            self.json_list.delete(0, tk.END)
            self.process_btn.config(text="Process GSTR3B")
            self.update_process_button()

        except Exception as e:
            # Telemetry: error event
            send_event("error", {
                "module": "gstr3b_ui",
                "error": str(e),
                "input_files": self.json_files
            })

            self.show_error_with_copy("Error", f"An error occurred:\n{str(e)}")
            self.process_btn.config(text="Process GSTR3B")
            self.update_process_button()

    def show_error_with_copy(self, title, message):
        print(f"ERROR: {message}")
        error_window = tk.Toplevel(self.root)
        error_window.title(title)
        error_window.geometry("400x150")
        error_window.resizable(False, False)
        tk.Label(error_window, text=message, wraplength=380, justify="left").pack(pady=10)

        def copy_error():
            import pyperclip
            pyperclip.copy(message)
            copy_button.config(text="Copied!")
            error_window.after(2000, lambda: copy_button.config(text="Copy Error"))

        copy_button = tk.Button(error_window, text="Copy Error", command=copy_error)
        copy_button.pack(side="left", padx=10, pady=10)
        tk.Button(error_window, text="OK", command=error_window.destroy).pack(side="right", padx=10, pady=10)
        error_window.transient(self.root)
        error_window.grab_set()
        self.root.wait_window(error_window)


if __name__ == "__main__":
    root = tk.Tk()
    app = GSTR3BProcessorUI(root)
    root.mainloop()
