import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from processor import clean_csv_to_xlsx, process_zerobounce


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Cleaner")
        self.root.geometry("600x450")
        self.root.resizable(False, False)

        self.file_path = tk.StringVar()

        self._build_ui()

    def _build_ui(self):
        padding = {"padx": 10, "pady": 5}

        # File selection
        file_frame = ttk.LabelFrame(self.root, text="Select File", padding=10)
        file_frame.pack(fill="x", **padding)

        ttk.Entry(file_frame, textvariable=self.file_path, state="readonly").pack(
            side="left", fill="x", expand=True, padx=(0, 5)
        )
        ttk.Button(file_frame, text="Browse", command=self._browse).pack(side="right")

        # Actions
        action_frame = ttk.LabelFrame(self.root, text="Actions", padding=10)
        action_frame.pack(fill="x", **padding)

        ttk.Button(
            action_frame, text="Clean Columns → XLSX", command=self._run_clean
        ).pack(side="left", fill="x", expand=True, padx=(0, 5))

        ttk.Button(
            action_frame, text="Process ZeroBounce", command=self._run_zerobounce
        ).pack(side="left", fill="x", expand=True)

        # Log
        log_frame = ttk.LabelFrame(self.root, text="Log", padding=10)
        log_frame.pack(fill="both", expand=True, **padding)

        self.log_text = tk.Text(log_frame, state="disabled", wrap="word", height=12)
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.log_text.pack(side="left", fill="both", expand=True)

    def _browse(self):
        path = filedialog.askopenfilename(
            title="Select CSV file",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if path:
            self.file_path.set(path)

    def _log(self, message):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _set_buttons(self, state):
        for child in self.root.winfo_children():
            if isinstance(child, ttk.LabelFrame) and child.cget("text") == "Actions":
                for btn in child.winfo_children():
                    btn.configure(state=state)

    def _run_clean(self):
        path = self.file_path.get()
        if not path:
            messagebox.showwarning("No file", "Please select a CSV file first.")
            return
        self._set_buttons("disabled")
        self._log("Starting column cleanup...")
        thread = threading.Thread(target=self._clean_worker, args=(path,), daemon=True)
        thread.start()

    def _run_zerobounce(self):
        path = self.file_path.get()
        if not path:
            messagebox.showwarning("No file", "Please select a CSV file first.")
            return
        self._set_buttons("disabled")
        self._log("Starting ZeroBounce processing...")
        thread = threading.Thread(
            target=self._zerobounce_worker, args=(path,), daemon=True
        )
        thread.start()

    def _clean_worker(self, path):
        try:
            output = clean_csv_to_xlsx(path, log=self._log)
            if output:
                self._log(f"Done! Output: {output}")
        except Exception as e:
            self._log(f"Error: {e}")
        finally:
            self.root.after(0, lambda: self._set_buttons("normal"))

    def _zerobounce_worker(self, path):
        try:
            result = process_zerobounce(path, log=self._log)
            if result:
                self._log("All files saved successfully.")
        except Exception as e:
            self._log(f"Error: {e}")
        finally:
            self.root.after(0, lambda: self._set_buttons("normal"))


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
