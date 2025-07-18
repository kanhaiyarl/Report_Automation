import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from datetime import datetime
import os

try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

class ToolTip(object):
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        # Only show tooltip on mouse hover, not on focus/tab
        widget.bind("<Enter>", self.show_tip)
        widget.bind("<Leave>", self.hide_tip)
        # Do NOT bind to <FocusIn> or <FocusOut>

    def show_tip(self, event=None):
        if self.tipwindow or not self.text:
            return
        # Only show on mouse events
        if event is not None and event.type != tk.EventType.Enter:
            return
        x, y, cx, cy = self.widget.bbox("insert") if self.widget.winfo_ismapped() else (0,0,0,0)
        x = x + self.widget.winfo_rootx() + 25
        y = y + self.widget.winfo_rooty() + 20
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(tw, text=self.text, justify=tk.LEFT,
                         background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                         font=("tahoma", "9", "normal"))
        label.pack(ipadx=6, ipady=2)

    def hide_tip(self, event=None):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

class ExcelMergeApp:
    def toggle_file2_state(self):
        if self.append_replace_var.get():
            self.file2_entry.config(state='normal')
            self.file2_btn.config(state='normal')
        else:
            self.file2_entry.config(state='disabled')
            self.file2_btn.config(state='disabled')

    def __init__(self, master):
        self.master = master
        master.title("Excel Sheet Merge Tool")
        master.geometry("570x600")
        master.configure(bg="#f4f6fb")
        master.resizable(False, False)

        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.datetime_str = tk.StringVar()

        label_font = ("Segoe UI", 11)
        entry_font = ("Segoe UI", 11)
        button_font = ("Segoe UI", 11, "bold")
        heading_font = ("Segoe UI", 16, "bold")
        section_font = ("Segoe UI", 12, "bold")

        # Logo area
        logo_frame = tk.Frame(master, bg="#f4f6fb")
        logo_frame.pack(pady=(10, 0))
        if PIL_AVAILABLE:
            try:
                logo_img = Image.open(os.path.join(os.path.dirname(__file__), "excel_logo.png"))
                logo_img = logo_img.resize((40, 40), Image.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(logo_img)
                tk.Label(logo_frame, image=self.logo_photo, bg="#f4f6fb").pack()
            except Exception:
                tk.Label(logo_frame, text="📊", font=("Segoe UI Emoji", 32), bg="#f4f6fb").pack()
        else:
            tk.Label(logo_frame, text="📊", font=("Segoe UI Emoji", 32), bg="#f4f6fb").pack()

        # Title
        tk.Label(master, text="Excel Sheet Merge Tool", font=heading_font, bg="#f4f6fb", fg="#2b3a67").pack(pady=(0, 3))
        tk.Label(master, text="Easily merge and update Excel files based on date-time.", font=("Segoe UI", 10), bg="#f4f6fb", fg="#4a4a4a").pack(pady=(0, 10))

        # Frame for file selectors
        file_frame = tk.Frame(master, bg="#f4f6fb", highlightbackground="#d1d9e6", highlightthickness=1, bd=0)
        file_frame.pack(pady=5, fill=tk.X, padx=30)

        # File 1
        tk.Label(file_frame, text="File 1 (where changes happen):", font=section_font, bg="#f4f6fb", anchor="w").grid(row=0, column=0, sticky="w", pady=(0,2))
        file1_entry = tk.Entry(file_frame, textvariable=self.file1_path, width=36, font=entry_font, relief="groove", bd=2)
        file1_entry.grid(row=1, column=0, padx=(0,8), sticky="w")
        file1_btn = tk.Button(file_frame, text="Browse", command=self.browse_file1, font=button_font, bg="#5b9bd5", fg="white", activebackground="#3d6fa5", relief="flat", cursor="hand2")
        file1_btn.grid(row=1, column=1, sticky="w")
        ToolTip(file1_btn, "Select the Excel file you want to update.")

        # File 2
        tk.Label(file_frame, text="File 2 (to append/replace):", font=section_font, bg="#f4f6fb", anchor="w").grid(row=2, column=0, sticky="w", pady=(12,2))
        self.file2_entry = tk.Entry(file_frame, textvariable=self.file2_path, width=36, font=entry_font, relief="groove", bd=2)
        self.file2_entry.grid(row=3, column=0, padx=(0,8), sticky="w")
        self.file2_btn = tk.Button(file_frame, text="Browse", command=self.browse_file2, font=button_font, bg="#5b9bd5", fg="white", activebackground="#3d6fa5", relief="flat", cursor="hand2")
        self.file2_btn.grid(row=3, column=1, sticky="w")
        ToolTip(self.file2_btn, "Select the Excel file to take new data from.")

        # Date-time input
        dt_frame = tk.Frame(master, bg="#f4f6fb")
        dt_frame.pack(pady=(18, 0), fill=tk.X, padx=30)
        tk.Label(dt_frame, text="Enter Date-Time (YYYY-MM-DD HH:MM):", font=label_font, bg="#f4f6fb").grid(row=0, column=0, sticky="w")
        dt_entry = tk.Entry(dt_frame, textvariable=self.datetime_str, width=20, font=entry_font, relief="groove", bd=2)
        dt_entry.grid(row=0, column=1, padx=(12,0))
        ToolTip(dt_entry, "Type the date and time (e.g. 2025-07-18 10:15)")

        # Operation checkboxes
        option_frame = tk.Frame(master, bg="#f4f6fb")
        option_frame.pack(pady=(10,0), fill=tk.X, padx=30)
        self.remove_rows_var = tk.BooleanVar(value=True)
        self.append_replace_var = tk.BooleanVar(value=True)
        remove_cb = tk.Checkbutton(option_frame, text="Remove rows from File 1 before date-time", variable=self.remove_rows_var, bg="#f4f6fb", font=label_font, anchor="w")
        append_cb = tk.Checkbutton(option_frame, text="Append/replace from File 2 on/after date-time", variable=self.append_replace_var, bg="#f4f6fb", font=label_font, anchor="w", command=self.toggle_file2_state)
        remove_cb.grid(row=0, column=0, sticky="w")
        append_cb.grid(row=1, column=0, sticky="w")
        ToolTip(remove_cb, "Remove rows in File 1 where 'datetime' is before the given date-time.")
        ToolTip(append_cb, "Append or replace rows in File 1 with rows from File 2 on or after the date-time.")
        self.toggle_file2_state()

        # Buttons frame
        btns_frame = tk.Frame(master, bg="#f4f6fb")
        btns_frame.pack(pady=25)

        process_btn = tk.Button(btns_frame, text="Process", command=self.process_preview, font=button_font, bg="#4b7bec", fg="white", activebackground="#3867d6", relief="flat", cursor="hand2")
        process_btn.grid(row=0, column=0, padx=10)
        ToolTip(process_btn, "Preview what will happen (row counts, etc) without saving.")

        execute_btn = tk.Button(btns_frame, text="Execute", command=self.process_files, font=button_font, bg="#2b3a67", fg="white", activebackground="#1a2240", relief="flat", cursor="hand2")
        execute_btn.grid(row=0, column=1, padx=10)
        ToolTip(execute_btn, "Run the process and save the updated Excel file.")

        # Footer
        footer = tk.Label(master, text="© 2025 Excel Merge Tool | Help: Use standard Excel files with 'id' and 'datetime' columns.", font=("Segoe UI", 9), bg="#e8eaf6", fg="#4a4a4a", bd=1, relief="flat")
        footer.pack(side=tk.BOTTOM, fill=tk.X, pady=(15,0))

    def browse_file1(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.file1_path.set(path)

    def browse_file2(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.file2_path.set(path)

    def get_processed_dataframe(self):
        file1 = self.file1_path.get()
        file2 = self.file2_path.get()
        dt_str = self.datetime_str.get()
        do_remove = self.remove_rows_var.get()
        do_append = self.append_replace_var.get()

        if not file1 or not file2 or not dt_str:
            return None, "Please select both files and enter a date-time."
        if not (do_remove or do_append):
            return None, "Please select at least one operation to perform."
        try:
            dt = datetime.strptime(dt_str, "%Y-%m-%d %H:%M")
        except ValueError:
            return None, "Invalid date-time format. Use YYYY-MM-DD HH:MM"
        try:
            df1 = pd.read_excel(file1)
            df2 = pd.read_excel(file2)
        except Exception as e:
            return None, f"Failed to read Excel files: {e}"

        # Handle datetime columns with various possible names
        dt_col1 = None
        dt_col2 = None
        for col in df1.columns:
            if col.strip().lower() in ['datetime', 'end date', 'enddate']:
                dt_col1 = col
                break
        for col in df2.columns:
            if col.strip().lower() in ['datetime', 'endezeit', 'end date', 'enddate']:
                dt_col2 = col
                break
        if not dt_col1 or not dt_col2:
            return None, "File 1 must have a 'End date' column and File 2 must have an 'Endezeit' column (or variants)."
        df1 = df1.rename(columns={dt_col1: 'datetime'})
        df2 = df2.rename(columns={dt_col2: 'datetime'})
        df1['datetime'] = pd.to_datetime(df1['datetime'], errors='coerce')
        df2['datetime'] = pd.to_datetime(df2['datetime'], errors='coerce')
        if df1['datetime'].isnull().all() or df2['datetime'].isnull().all():
            return None, "Failed to parse date-times in one or both files. Please check the date columns."

        # Assume there is a unique key column to identify duplicates, e.g., 'id'.
        if 'id' not in df1.columns or 'id' not in df2.columns:
            return None, "Both files must have an 'id' column (unique row identifier)."

        result_df = df1.copy()
        if do_remove:
            result_df = result_df[result_df['datetime'] >= dt].copy()
        if do_append:
            df2_filtered = df2[df2['datetime'] >= dt].copy()
            # Remove rows from result_df that have the same 'id' as in df2_filtered
            result_df = result_df[~result_df['id'].isin(df2_filtered['id'])]
            # Append df2_filtered rows
            result_df = pd.concat([result_df, df2_filtered], ignore_index=True)
        result_df = result_df.sort_values(by='datetime')
        return result_df, None

    def process_preview(self):
        result_df, err = self.get_processed_dataframe()
        if err:
            messagebox.showerror("Preview Error", err)
            return
        msg = f"Preview:\nRows in result: {len(result_df)}\nColumns: {list(result_df.columns)}"
        messagebox.showinfo("Process Preview", msg)

    def process_files(self):
        result_df, err = self.get_processed_dataframe()
        if err:
            messagebox.showerror("Error", err)
            return
        # Save the result
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="Save updated File 1 as...")
        if save_path:
            try:
                result_df.to_excel(save_path, index=False)
                messagebox.showinfo("Success", f"File saved to {save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMergeApp(root)
    root.mainloop()
