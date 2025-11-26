# launch_gui.py
"""
Small GUI launcher for the Excel module.
Double-click this file or run `python launch_gui.py`.
"""
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import os
import traceback

# Set a default sample path (edit if needed)
SAMPLE_PATH = r"C:\Users\Pc\PyCharmMiscProject\sample_sales.xlsx"

def import_session():
    from excel_module.commands import Session
    return Session

class LauncherApp:
    def __init__(self, root):
        self.root = root
        root.title("Excel Automation Launcher")
        root.geometry("700x480")
        root.resizable(False, False)
        frm = tk.Frame(root, padx=12, pady=12)
        frm.pack(fill=tk.BOTH, expand=True)

        # File picker
        row1 = tk.Frame(frm); row1.pack(fill=tk.X, pady=(0,8))
        tk.Label(row1, text="Workbook:", width=10).pack(side=tk.LEFT)
        self.path_var = tk.StringVar()
        self.path_entry = tk.Entry(row1, textvariable=self.path_var, width=70); self.path_entry.pack(side=tk.LEFT, padx=(4,4))
        tk.Button(row1, text="Choose file", command=self.choose_file).pack(side=tk.LEFT, padx=(4,0))
        tk.Button(row1, text="Use sample", command=self.use_sample).pack(side=tk.LEFT, padx=(4,0))

        # Options
        row2 = tk.Frame(frm); row2.pack(fill=tk.X, pady=(0,8))
        self.visible_var = tk.BooleanVar(value=True)
        tk.Checkbutton(row2, text="Show Excel while running (debug)", variable=self.visible_var).pack(side=tk.LEFT)
        self.open_var = tk.BooleanVar(value=True)
        tk.Checkbutton(row2, text="Open workbook after run", variable=self.open_var).pack(side=tk.LEFT, padx=(8,0))

        # Buttons
        row3 = tk.Frame(frm); row3.pack(fill=tk.X, pady=(0,8))
        self.run_btn = tk.Button(row3, text="Run", width=14, command=self.start_run); self.run_btn.pack(side=tk.LEFT)
        tk.Button(row3, text="Clear Log", width=12, command=self.clear_log).pack(side=tk.LEFT, padx=(8,0))
        tk.Button(row3, text="Quit", width=12, command=self.on_quit).pack(side=tk.LEFT, padx=(8,0))

        # Log area
        tk.Label(frm, text="Status / Log:").pack(anchor="w")
        self.log = scrolledtext.ScrolledText(frm, height=18, state="disabled", wrap=tk.WORD); self.log.pack(fill=tk.BOTH, expand=True)

        if os.path.exists(SAMPLE_PATH):
            self.path_var.set(SAMPLE_PATH)
        self._running = False

    def choose_file(self):
        p = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx *.xls")])
        if p: self.path_var.set(p)

    def use_sample(self):
        if os.path.exists(SAMPLE_PATH):
            self.path_var.set(SAMPLE_PATH)
        else:
            messagebox.showwarning("Sample not found", SAMPLE_PATH)

    def clear_log(self):
        self.log.configure(state="normal"); self.log.delete("1.0", tk.END); self.log.configure(state="disabled")

    def on_quit(self):
        if self._running and not messagebox.askyesno("Quit", "A run is in progress. Quit anyway?"):
            return
        self.root.quit()

    def _log(self, *msgs):
        text = " ".join(str(m) for m in msgs)
        self.log.configure(state="normal"); self.log.insert(tk.END, text + "\n"); self.log.see(tk.END); self.log.configure(state="disabled")

    def start_run(self):
        if self._running:
            messagebox.showwarning("Already running", "A run is in progress.")
            return
        path = self.path_var.get().strip()
        if not path:
            messagebox.showwarning("No file", "Please choose an Excel file first.")
            return
        if not os.path.exists(path):
            messagebox.showerror("File not found", path)
            return
        threading.Thread(target=self._run_pipeline, args=(path, self.visible_var.get(), self.open_var.get()), daemon=True).start()

    def _run_pipeline(self, path, visible, open_in_excel):
        self._running = True; self.run_btn.configure(state="disabled")
        try:
            self._log("START:", path)
            try:
                Session = import_session()
            except Exception as e:
                self._log("Import failed:", str(e)); messagebox.showerror("Import error", str(e)); return
            sess = Session(path=path)
            try:
                self._log("Loading file...")
                sess.load(path)
                self._log("Cleaning data...")
                sess.clean()
                self._log("Opening Excel and writing cleaned data...")
                sess.write_to_excel(sheet_name="SourceData", visible=visible)
                self._log("Building dashboard (this may take a few seconds)...")
                sess.create_dashboard(dashboard_name="AutoDashboard", visible=visible, open_in_excel=open_in_excel)
                self._log("DONE: Dashboard created")
                messagebox.showinfo("Success", f"Dashboard created for:\n{path}")
            except Exception as e:
                tb = traceback.format_exc()
                self._log("ERROR:", str(e))
                self._log(tb)
                messagebox.showerror("Error", str(e))
            finally:
                try: sess.close()
                except: pass
        finally:
            self._running = False; self.run_btn.configure(state="normal")

if __name__ == "__main__":
    root = tk.Tk(); app = LauncherApp(root); root.mainloop()
