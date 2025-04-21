import tkinter as tk
from tkinter import filedialog, messagebox
import os
import datetime

# ←— NEW IMPORTS for telemetry
from telemetry import send_event
from license_util import get_machine_guid

from gstr1_processor import process_gstr1, parse_filename, get_tax_period, parse_large_filename  # Import required functions

class GSTR1ProcessorUI:
    def __init__(self, root):
        self.root = root
        self.root.title("GSTR1 Processing")
        self.root.geometry("500x480")
        self.small_files = []
        self.large_files = []
        self.template_file = None
        self.excluded_sections_by_month = {}
        self.base_height = 450

        tk.Label(root, text="GSTR1 Processing", font=("Arial", 16, "bold")).pack(pady=5)
        main_frame = tk.Frame(root)
        main_frame.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)

        # <500 JSON Section
        left_frame = tk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tk.Label(left_frame, text="GSTR1 JSON (<500)", font=("Arial", 10, "bold")).pack()
        self.small_listbox = tk.Listbox(left_frame, height=13, width=38, selectmode=tk.MULTIPLE)
        self.small_listbox.pack(pady=2, fill=tk.Y, expand=True)
        self.small_listbox.bind("<Button-1>", self.single_click_small)
        self.small_listbox.bind("<Shift-Button-1>", self.shift_click_small)
        self.small_listbox.bind("<Control-Button-1>", self.ctrl_click_small)
        btn_frame = tk.Frame(left_frame)
        btn_frame.pack(pady=5)
        tk.Button(btn_frame, text="+ Add", command=self.add_small_file).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="- Remove", command=self.delete_small_file).pack(side=tk.LEFT, padx=5)

        # >500 JSON Section
        right_frame = tk.Frame(main_frame)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tk.Label(right_frame, text="GSTR1 JSON Zip (>500)", font=("Arial", 10, "bold")).pack()
        self.large_listbox = tk.Listbox(right_frame, height=13, width=38, selectmode=tk.MULTIPLE)
        self.large_listbox.pack(pady=2, fill=tk.Y, expand=True)
        self.large_listbox.bind("<Button-1>", self.single_click_large)
        self.large_listbox.bind("<Shift-Button-1>", self.shift_click_large)
        self.large_listbox.bind("<Control-Button-1>", self.ctrl_click_large)
        large_btn_frame = tk.Frame(right_frame)
        large_btn_frame.pack(pady=5)
        tk.Button(large_btn_frame, text="+ Add", command=self.add_large_file).pack(side=tk.LEFT, padx=5)
        tk.Button(large_btn_frame, text="- Remove", command=self.delete_large_file).pack(side=tk.LEFT, padx=5)

        # Warning Frame
        self.warning_frame = tk.Frame(root, borderwidth=1, relief="solid")
        self.warning_title = tk.Label(self.warning_frame, text="Warning !", fg="red", font=("Arial", 10, "underline"))
        self.warning_text = tk.Label(self.warning_frame, text="", fg="red", justify=tk.LEFT, wraplength=450)
        self.ignore_var = tk.BooleanVar()
        self.ignore_check = tk.Checkbutton(self.warning_frame, text="Ignore All Warnings", variable=self.ignore_var, command=self.update_process_button)

        # Template File Section
        template_frame = tk.Frame(root)
        template_frame.pack(pady=5)
        tk.Label(template_frame, text="Template Excel File (Optional):").pack(side=tk.LEFT)
        self.template_label = tk.Label(template_frame, text="No file selected")
        self.template_label.pack(side=tk.LEFT, padx=5)
        tk.Button(template_frame, text="Select", command=self.select_template).pack(side=tk.LEFT, padx=2)
        tk.Button(template_frame, text="Clear", command=self.clear_template).pack(side=tk.LEFT, padx=2)

        # Process Button
        self.process_btn = tk.Button(root, text="Process GSTR1", font=("Arial", 12), command=self.process_files, state=tk.DISABLED, bg="light grey")
        self.process_btn.pack(pady=10)

    def single_click_small(self, event):
        index = self.small_listbox.nearest(event.y)
        self.small_listbox.selection_clear(0, tk.END)
        self.small_listbox.selection_set(index)
        return "break"

    def shift_click_small(self, event):
        index = self.small_listbox.nearest(event.y)
        if not self.small_listbox.curselection():
            self.small_listbox.selection_set(index)
        else:
            anchor = self.small_listbox.curselection()[0]
            start, end = min(anchor, index), max(anchor, index)
            self.small_listbox.selection_clear(0, tk.END)
            self.small_listbox.selection_set(start, end)
        return "break"

    def ctrl_click_small(self, event):
        index = self.small_listbox.nearest(event.y)
        if index in self.small_listbox.curselection():
            self.small_listbox.selection_clear(index)
        else:
            self.small_listbox.selection_set(index, last=None)
        return "break"

    def single_click_large(self, event):
        index = self.large_listbox.nearest(event.y)
        self.large_listbox.selection_clear(0, tk.END)
        self.large_listbox.selection_set(index)
        return "break"

    def shift_click_large(self, event):
        index = self.large_listbox.nearest(event.y)
        if not self.large_listbox.curselection():
            self.large_listbox.selection_set(index)
        else:
            anchor = self.large_listbox.curselection()[0]
            start, end = min(anchor, index), max(anchor, index)
            self.large_listbox.selection_clear(0, tk.END)
            self.large_listbox.selection_set(start, end)
        return "break"

    def ctrl_click_large(self, event):
        index = self.large_listbox.nearest(event.y)
        if index in self.large_listbox.curselection():
            self.large_listbox.selection_clear(index)
        else:
            self.large_listbox.selection_set(index, last=None)
        return "break"

    def add_small_file(self):
        files = filedialog.askopenfilenames(filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")])
        financial_order = ["April", "May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March"]
        for file in files:
            month, excluded = parse_filename(file)
            if not month or (file, month) in self.small_files:
                continue
            self.small_files.append((file, month))
            if excluded:
                self.excluded_sections_by_month[month] = excluded
        self.small_files.sort(key=lambda x: financial_order.index(get_tax_period(x[1])) if get_tax_period(x[1]) in financial_order else 999)
        self.small_listbox.delete(0, tk.END)
        for file, _ in self.small_files:
            self.small_listbox.insert(tk.END, os.path.basename(file))
        self.update_process_button()

    def delete_small_file(self):
        selections = self.small_listbox.curselection()
        if selections:
            for index in reversed(selections):
                file, month = self.small_files.pop(index)
                self.small_listbox.delete(index)
                if month in self.excluded_sections_by_month and not any(m == month for _, m in self.small_files):
                    del self.excluded_sections_by_month[month]
            self.update_process_button()

    def add_large_file(self):
        files = filedialog.askopenfilenames(filetypes=[("ZIP Files", "*.zip"), ("All Files", "*.*")])
        financial_order = ["April", "May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March"]
        for file in files:
            month = parse_large_filename(file)  # Extracts "012025" from "GSTR1_Full_012025.zip"
            if not month or (file, month) in self.large_files:
                continue
            self.large_files.append((file, month))
        self.large_files.sort(key=lambda x: financial_order.index(get_tax_period(x[1])) if get_tax_period(x[1]) in financial_order else 999)
        self.large_listbox.delete(0, tk.END)
        for file, _ in self.large_files:
            self.large_listbox.insert(tk.END, os.path.basename(file))
        self.update_process_button()

    def delete_large_file(self):
        selections = self.large_listbox.curselection()
        if selections:
            for index in reversed(selections):
                self.large_files.pop(index)
                self.large_listbox.delete(index)
            self.update_process_button()

    def select_template(self):
        file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])
        if file:
            self.template_file = file
            self.template_label.config(text=os.path.basename(file))

    def clear_template(self):
        self.template_file = None
        self.template_label.config(text="No file selected")

    def update_process_button(self):
        warnings = []
        required_months = set(self.excluded_sections_by_month.keys())
        selected_large_months = {month for _, month in self.large_files}
        selected_small_months = {month for _, month in self.small_files}
        missing_months = required_months - selected_large_months
        if missing_months:
            warnings.append(f"'>500' JSON file for {', '.join(sorted(missing_months))} not selected")

        small_month_counts = {}
        for _, month in self.small_files:
            small_month_counts[month] = small_month_counts.get(month, 0) + 1
        duplicate_small = [month for month, count in small_month_counts.items() if count > 1]
        if duplicate_small:
            warnings.append(f"Multiple '<500' JSON files selected for {', '.join(sorted(duplicate_small))}")

        large_month_counts = {}
        for _, month in self.large_files:
            large_month_counts[month] = large_month_counts.get(month, 0) + 1
        duplicate_large = [month for month, count in large_month_counts.items() if count > 1]
        if duplicate_large:
            warnings.append(f"Multiple >500 JSON files selected for {', '.join(sorted(duplicate_large))}")

        if self.large_files:
            missing_small_months = selected_large_months - selected_small_months
            if missing_small_months or not self.small_files:
                months_str = ', '.join(sorted(missing_small_months)) if missing_small_months else 'any month'
                warnings.append(f"No <500 JSON file selected for month {months_str}. Only >500 data will be processed.")

        has_files = bool(self.small_files or self.large_files)
        if warnings:
            if not self.warning_frame.winfo_ismapped():
                self.warning_frame.pack(pady=5, padx=10, fill=tk.X)
            self.warning_title.pack(pady=2)
            self.warning_text.config(text="\n".join(warnings))
            self.warning_text.pack(pady=2)
            self.ignore_check.pack(pady=2)
            self.root.geometry(f"500x{self.base_height + 100}")
            if self.ignore_var.get() and has_files:
                self.process_btn.config(state=tk.NORMAL, bg="light green")
            else:
                self.process_btn.config(state=tk.DISABLED, bg="light grey")
        else:
            if self.warning_frame.winfo_ismapped():
                self.warning_frame.pack_forget()
                self.root.geometry(f"500x{self.base_height}")
            if has_files:
                self.process_btn.config(state=tk.NORMAL, bg="light green")
            else:
                self.process_btn.config(state=tk.DISABLED, bg="light grey")

    def process_files(self):
        if not self.small_files and not self.large_files:
            messagebox.showerror("Error", "No files selected for processing.")
            return

        save_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            title="Save GSTR1 Report As"
        )
        if not save_file:
            return

        self.process_btn.config(text="Processing...", state=tk.DISABLED, bg="light grey")
        self.root.update()

        # Prepare lists for telemetry
        small_file_paths = [f for f, _ in self.small_files]
        large_file_paths = [f for f, _ in self.large_files]

        try:
            # If you want a record count, have process_gstr1 return it:
            # record_count = process_gstr1(...)
            process_gstr1(
                small_file_paths,
                {month: (file, month) for file, month in self.large_files},
                self.excluded_sections_by_month,
                self.template_file,
                save_file,
                ignore_warnings=self.ignore_var.get()
            )

            # ←— TELEMETRY: Success event
            send_event("gstr1_complete", {
                "input_small_files": small_file_paths,
                "input_large_files": large_file_paths,
                "output_file": save_file,
                # "records_processed": record_count  # if you capture this
            })

            messagebox.showinfo("Success", f"GSTR1 report saved successfully at:\n{save_file}")
            self.small_files.clear()
            self.large_files.clear()
            self.excluded_sections_by_month.clear()
            self.small_listbox.delete(0, tk.END)
            self.large_listbox.delete(0, tk.END)
            self.ignore_var.set(False)
            self.process_btn.config(text="Process GSTR1")
            self.update_process_button()
        except Exception as e:
            # ←— TELEMETRY: Error event
            send_event("error", {
                "module": "gstr1_ui",
                "error": str(e),
                "input_small_files": small_file_paths,
                "input_large_files": large_file_paths
            })

            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
            self.process_btn.config(text="Process GSTR1")
            self.update_process_button()

if __name__ == "__main__":
    root = tk.Tk()
    app = GSTR1ProcessorUI(root)
    root.mainloop()