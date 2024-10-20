import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import csv
from datetime import datetime
import ttkbootstrap as tb
from ttkbootstrap.constants import *
import winsound

class CSVMatcherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV Matcher")
        self.style = tb.Style(theme="cosmo")
        self.create_widgets()
        self.csv_data = []
        self.original_csv_data = []  # Store original CSV data for the total count
        self.csv_file_path = ""
        self.cartno_usage = {str(i): 0 for i in range(1, 21)}

    def create_widgets(self):
        entry_frame = ttk.Frame(self.root)
        entry_frame.grid(row=0, column=0, padx=10, pady=10)

        self.name_label = ttk.Label(entry_frame, text="Name:")
        self.name_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.name_entry = ttk.Entry(entry_frame)
        self.name_entry.grid(row=0, column=1, padx=5, pady=5)
        self.name_entry.bind('<Return>', lambda event: self.on_enter(event, self.lotno_entry))

        self.lotno_label = ttk.Label(entry_frame, text="LotNo:")
        self.lotno_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.lotno_entry = ttk.Entry(entry_frame)
        self.lotno_entry.grid(row=1, column=1, padx=5, pady=5)
        self.lotno_entry.bind('<Return>', lambda event: self.on_enter(event, self.splitno_entry))

        self.splitno_label = ttk.Label(entry_frame, text="SplitNo:")
        self.splitno_label.grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.splitno_entry = ttk.Entry(entry_frame)
        self.splitno_entry.grid(row=2, column=1, padx=5, pady=5)
        self.splitno_entry.bind('<Return>', lambda event: self.on_enter(event, self.overlabel_entry))

        self.overlabel_label = ttk.Label(entry_frame, text="Overlabel:")
        self.overlabel_label.grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.overlabel_entry = ttk.Entry(entry_frame)
        self.overlabel_entry.grid(row=3, column=1, padx=5, pady=5)
        self.overlabel_entry.bind('<Return>', lambda event: self.on_enter(event, self.frozen_entry))

        self.frozen_label = ttk.Label(entry_frame, text="Frozen:")
        self.frozen_label.grid(row=4, column=0, padx=5, pady=5, sticky="e")
        self.frozen_entry = ttk.Entry(entry_frame)
        self.frozen_entry.grid(row=4, column=1, padx=5, pady=5)
        self.frozen_entry.bind('<Return>', lambda event: self.on_enter(event, self.holecheck_entry))

        self.holecheck_label = ttk.Label(entry_frame, text="Holecheck:")
        self.holecheck_label.grid(row=5, column=0, padx=5, pady=5, sticky="e")
        self.holecheck_entry = ttk.Entry(entry_frame)
        self.holecheck_entry.grid(row=5, column=1, padx=5, pady=5)
        self.holecheck_entry.bind('<Return>', lambda event: self.on_enter(event, self.labelcheck_entry))

        self.labelcheck_label = ttk.Label(entry_frame, text="Labelcheck:")
        self.labelcheck_label.grid(row=6, column=0, padx=5, pady=5, sticky="e")
        self.labelcheck_entry = ttk.Entry(entry_frame)
        self.labelcheck_entry.grid(row=6, column=1, padx=5, pady=5)
        self.labelcheck_entry.bind('<Return>', lambda event: self.on_enter(event, self.cartno_entry))

        self.cartno_label = ttk.Label(entry_frame, text="Cartno (1-20):")
        self.cartno_label.grid(row=7, column=0, padx=5, pady=5, sticky="e")
        self.cartno_entry = ttk.Entry(entry_frame)
        self.cartno_entry.grid(row=7, column=1, padx=5, pady=5)

        self.total_count_label = ttk.Label(entry_frame, text="Total Rows: 0")
        self.total_count_label.grid(row=8, column=0, padx=5, pady=5, sticky="w")
        self.matched_count_label = ttk.Label(entry_frame, text="Matched Rows: 0")
        self.matched_count_label.grid(row=8, column=1, padx=5, pady=5, sticky="w")

        self.load_button = ttk.Button(entry_frame, text="Load CSV", command=self.load_csv)
        self.load_button.grid(row=9, column=0, padx=5, pady=5)

        self.match_button = ttk.Button(entry_frame, text="Match Data", command=self.match_data)
        self.match_button.grid(row=9, column=1, padx=5, pady=5)

        tree_frame = ttk.Frame(self.root)
        tree_frame.grid(row=1, column=0, padx=10, pady=10)
        self.tree = ttk.Treeview(tree_frame, show='headings')
        self.tree.pack(fill='both', expand=True)

    def on_enter(self, event, next_entry):
        next_entry.focus()

    def reset_entries(self):
        self.name_entry.delete(0, tk.END)
        self.lotno_entry.delete(0, tk.END)
        self.splitno_entry.delete(0, tk.END)
        self.overlabel_entry.delete(0, tk.END)
        self.frozen_entry.delete(0, tk.END)
        self.holecheck_entry.delete(0, tk.END)
        self.labelcheck_entry.delete(0, tk.END)
        self.cartno_entry.delete(0, tk.END)

    def load_csv(self):
        self.csv_file_path = filedialog.askopenfilename()
        if not self.csv_file_path:
            return

        with open(self.csv_file_path, newline='') as csvfile:
            reader = csv.reader(csvfile)
            headers = next(reader)
            self.tree["columns"] = headers
            for header in headers:
                self.tree.heading(header, text=header)
                self.tree.column(header, width=100)

            self.csv_data = [row for row in reader]
            self.original_csv_data = list(self.csv_data)  # Store original CSV data
            print(f"CSV data loaded: {len(self.csv_data)} rows")
            self.update_treeview(self.csv_data)
            self.total_count_label.config(text=f"Total Rows: {len(self.original_csv_data)}")



    def update_treeview(self, data):
        self.tree.delete(*self.tree.get_children())
        for row in data:
            self.tree.insert("", "end", values=row)
        self.total_count_label.config(text=f"Total Rows: {len(data)}")

    def show_info(self, title, message):
        messagebox_id = tk.Toplevel(self.root)
        tk.Message(messagebox_id, text=message, padx=20, pady=20).pack()
        self.root.after(3000, messagebox_id.destroy)

    def match_data(self):
        print("match_data called")

        name = self.name_entry.get()
        lotno = self.lotno_entry.get()
        splitno = self.splitno_entry.get()
        overlabel = self.overlabel_entry.get()
        frozen = self.frozen_entry.get()
        holecheck = self.holecheck_entry.get()
        labelcheck = self.labelcheck_entry.get()
        cartno = self.cartno_entry.get()

        print(f"Entries: Name={name}, LotNo={lotno}, SplitNo={splitno}, Overlabel={overlabel}, Frozen={frozen}")

        if not (name or lotno or splitno or overlabel or frozen):
            messagebox.showwarning("Input Error", "Please enter at least one value to match.")
            return

        if not cartno.isdigit() or not (1 <= int(cartno) <= 20):
            messagebox.showwarning("Input Error", "Cartno must be a number between 1 and 20.")
            return

        if self.cartno_usage[cartno] >= 20:
            messagebox.showwarning("Input Error", f"Cartno {cartno} has been used 20 times. Please choose another number.")
            return

        matched_rows = []
        for row in self.csv_data:
            print(f"Checking row: {row}")
            if len(row) >= 5 and name in row[0] and lotno in row[1] and splitno in row[2] and overlabel in row[3] and frozen in row[4]:
                matched_rows.append(row)

        print(f"Matched rows: {matched_rows}")
        self.matched_count_label.config(text=f"Matched Rows: {len(matched_rows)}")
        print(f"Matched Rows Count: {len(matched_rows)}")

        if matched_rows:
            winsound.Beep(1000, 500)
            self.save_results(matched_rows, "OK", holecheck, labelcheck, cartno)
            self.delete_matched_data(matched_rows)
            self.cartno_usage[cartno] += 1
            self.show_info("Match Found", "OK")
        else:
            self.save_results([], "NG", holecheck, labelcheck, cartno)
            self.show_info("No Match", "NG")

        self.reset_entries()

    def save_results(self, matched_rows, result, holecheck, labelcheck, cartno):
        today_date = datetime.now().strftime("%Y-%m-%d")
        file_name = f"{today_date}_matchingresult.txt"
        with open(file_name, "a", newline='') as file:
            writer = csv.writer(file)
            if matched_rows:
                for row in matched_rows:
                    writer.writerow(row + [result, holecheck, labelcheck, cartno])
            else:
                writer.writerow([self.name_entry.get(), self.lotno_entry.get(), self.splitno_entry.get(), self.overlabel_entry.get(), self.frozen_entry.get(), result, holecheck, labelcheck, cartno])

    def delete_matched_data(self, matched_rows):
        remaining_data = [row for row in self.csv_data if row not in matched_rows]
        print(f"Remaining data after deletion: {remaining_data}")

        with open(self.csv_file_path, "w", newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(["Name", "LotNo", "SplitNo", "Overlabel", "Frozen"])
            writer.writerows(remaining_data)

        self.csv_data = remaining_data
        self.update_treeview(self.csv_data)
        print(f"CSV data after update: {self.csv_data}")

if __name__ == "__main__":
    root = tb.Window(themename="cosmo")
    app = CSVMatcherApp(root)
    root.mainloop()


