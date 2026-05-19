import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import json
import os
from processor import clean_csv_to_xlsx, process_zerobounce, CONFIG_PATH


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
        ).pack(side="left", fill="x", expand=True, padx=(0, 5))

        ttk.Button(
            action_frame, text="⚙ Settings", command=self._open_settings
        ).pack(side="right")

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

    def _open_settings(self):
        SettingsWindow(self.root)


class SettingsWindow:
    def __init__(self, parent):
        self.win = tk.Toplevel(parent)
        self.win.title("Column Settings")
        self.win.geometry("550x400")
        self.win.resizable(False, False)
        self.win.transient(parent)
        self.win.grab_set()

        self._load_config()
        self._build_ui()
        self._populate_tree()

    def _load_config(self):
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            self.config = json.load(f)

    def _save_config(self):
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(self.config, f, indent=2)

    def _build_ui(self):
        # Treeview
        tree_frame = ttk.Frame(self.win, padding=10)
        tree_frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(
            tree_frame, columns=("source", "target"), show="headings", height=12
        )
        self.tree.heading("source", text="Source Column")
        self.tree.heading("target", text="Output Column")
        self.tree.column("source", width=240)
        self.tree.column("target", width=240)

        scrollbar = ttk.Scrollbar(tree_frame, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)

        # Reorder buttons
        reorder_frame = ttk.Frame(self.win, padding=(10, 0, 10, 5))
        reorder_frame.pack(fill="x")

        ttk.Button(reorder_frame, text="↑ Move Up", command=self._move_up).pack(
            side="left", padx=(0, 5)
        )
        ttk.Button(reorder_frame, text="↓ Move Down", command=self._move_down).pack(
            side="left"
        )
        ttk.Button(reorder_frame, text="✕ Remove", command=self._remove).pack(
            side="right"
        )

        # Add column
        add_frame = ttk.LabelFrame(self.win, text="Add Column", padding=10)
        add_frame.pack(fill="x", padx=10, pady=(0, 10))

        ttk.Label(add_frame, text="Source:").pack(side="left")
        self.source_entry = ttk.Entry(add_frame, width=20)
        self.source_entry.pack(side="left", padx=(2, 10))

        ttk.Label(add_frame, text="Target:").pack(side="left")
        self.target_entry = ttk.Entry(add_frame, width=20)
        self.target_entry.pack(side="left", padx=(2, 10))

        ttk.Button(add_frame, text="Add", command=self._add_column).pack(side="right")

    def _populate_tree(self):
        self.tree.delete(*self.tree.get_children())
        columns = self.config["columns"]
        order = self.config["order"]

        for target_name in order:
            sources = [s for s, t in columns.items() if t == target_name]
            for src in sources:
                self.tree.insert("", "end", values=(src, target_name))

    def _get_selected(self):
        sel = self.tree.selection()
        if not sel:
            return None, None, None
        item = sel[0]
        values = self.tree.item(item, "values")
        return item, values[0], values[1]

    def _select_source(self, source_name):
        for child in self.tree.get_children():
            if self.tree.item(child, "values")[0] == source_name:
                self.tree.selection_set(child)
                self.tree.see(child)
                return

    def _move_up(self):
        item, src, target = self._get_selected()
        if item is None:
            return

        order = self.config["order"]
        idx = order.index(target)
        if idx == 0:
            return

        order[idx], order[idx - 1] = order[idx - 1], order[idx]
        self._save_config()
        self._populate_tree()
        self._select_source(src)

    def _move_down(self):
        item, src, target = self._get_selected()
        if item is None:
            return

        order = self.config["order"]
        idx = order.index(target)
        if idx >= len(order) - 1:
            return

        order[idx], order[idx + 1] = order[idx + 1], order[idx]
        self._save_config()
        self._populate_tree()
        self._select_source(src)

    def _remove(self):
        item, src, target = self._get_selected()
        if item is None:
            return

        columns = self.config["columns"]
        if src in columns:
            del columns[src]

        # Remove from order only if no other source maps to this target
        if not any(t == target for s, t in columns.items() if s != src):
            if target in self.config["order"]:
                self.config["order"].remove(target)

        self._save_config()
        self._populate_tree()

    def _add_column(self):
        source = self.source_entry.get().strip()
        target = self.target_entry.get().strip()

        if not source or not target:
            messagebox.showwarning("Missing info", "Both source and target are required.")
            return

        columns = self.config["columns"]
        if source in columns:
            messagebox.showwarning("Duplicate", f"'{source}' already exists.")
            return

        columns[source] = target

        if target not in self.config["order"]:
            self.config["order"].append(target)

        self._save_config()
        self._populate_tree()

        self.source_entry.delete(0, "end")
        self.target_entry.delete(0, "end")


if __name__ == "__main__":
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            windll.user32.SetProcessDPIAware()
        except Exception:
            pass
    root = tk.Tk()
    app = App(root)
    root.mainloop()
