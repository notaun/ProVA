# launch_gui.py (place next to excel_module/)
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import os
import traceback

# Set SAMPLE_PATH to your sample file if you want a one-click test:
SAMPLE_PATH = r"C:\Users\Pc\PyCharmMiscProject\sample_sales.xlsx"

def import_session():
    try:
        from excel_module.commands import Session
        return Session
    except Exception as e:
        raise RuntimeError("Import failed: ensure excel_module is on PYTHONPATH and deps installed.\n" + str(e))


class LauncherApp:
    def __init__(self, root):
        self.root = root
        root.title("Excel Automation Launcher")
        root.geometry("640x420")
        root.resizable(False, False)
        frm = tk.Frame(root, padx=12, pady=12)
        frm.pack(fill=tk.BOTH, expand=True)

        # file row
        row1 = tk.Frame(frm); row1.pack(fill=tk.X, pady=(0,8))
        tk.Label(row1, text="Workbook:", width=10).pack(side=tk.LEFT)
        self.path_var = tk.StringVar()
        self.path_entry = tk.Entry(row1, textvariable=self.path_var, width=60); self.path_entry.pack(side=tk.LEFT, padx=(4,4))
        tk.Button(row1, text="Choose file", command=self.choose_file).pack(side=tk.LEFT, padx=(4,0))
        tk.Button(row1, text="Use sample", command=self.use_sample).pack(side=tk.LEFT, padx=(4,0))

        # options
        row2 = tk.Frame(frm); row2.pack(fill=tk.X, pady=(0,8))
        self.visible_var = tk.BooleanVar(value=True)
        tk.Checkbutton(row2, text="Show Excel while running (debug)", variable=self.visible_var).pack(side=tk.LEFT)
        self.open_var = tk.BooleanVar(value=True)
        tk.Checkbutton(row2, text="Open workbook after run", variable=self.open_var).pack(side=tk.LEFT, padx=(8,0))

        # buttons
        row3 = tk.Frame(frm); row3.pack(fill=tk.X, pady=(0,8))
        self.run_btn = tk.Button(row3, text="Run", width=12, command=self.start_run); self.run_btn.pack(side=tk.LEFT)
        tk.Button(row3, text="Clear Log", width=12, command=self.clear_log).pack(side=tk.LEFT, padx=(8,0))
        tk.Button(row3, text="Quit", width=12, command=self.on_quit).pack(side=tk.LEFT, padx=(8,0))

        # log
        tk.Label(frm, text="Status / Log:").pack(anchor="w")
        self.log = scrolledtext.ScrolledText(frm, height=14, state="disabled", wrap=tk.WORD); self.log.pack(fill=tk.BOTH, expand=True)

        if os.path.exists(SAMPLE_PATH):
            self.path_var.set(SAMPLE_PATH)
        self._running = False

    def choose_file(self):
        p = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx *.xls")])
        if p: self.path_var.set(p)
    def use_sample(self):
        if os.path.exists(SAMPLE_PATH): self.path_var.set(SAMPLE_PATH)
        else: messagebox.showwarning("Sample not found", SAMPLE_PATH)
    def clear_log(self):
        self.log.configure(state="normal"); self.log.delete("1.0", tk.END); self.log.configure(state="disabled")
    def on_quit(self):
        if self._running and not messagebox.askyesno("Quit", "A run is in progress. Quit anyway?"): return
        self.root.quit()

    def _log(self, *msgs):
        text = " ".join(str(m) for m in msgs)
        self.log.configure(state="normal"); self.log.insert(tk.END, text + "\n"); self.log.see(tk.END); self.log.configure(state="disabled")

    def start_run(self):
        if self._running: messagebox.showwarning("Already running", "A run is in progress."); return
        path = self.path_var.get().strip()
        if not path: messagebox.showwarning("No file", "Please choose an Excel file first."); return
        if not os.path.exists(path): messagebox.showerror("File not found", path); return
        threading.Thread(target=self._run_pipeline, args=(path, self.visible_var.get(), self.open_var.get()), daemon=True).start()

    def _run_pipeline(self, path, visible, open_in_excel):
        self._running = True; self.run_btn.configure(state="disabled")
        try:
            self._log(f"START: {path}")
            try:
                Session = import_session()
            except Exception as e:
                self._log("Import error: " + str(e)); messagebox.showerror("Import error", str(e)); return
            sess = Session(path=path)
            try:
                self._log("Loading..."); sess.load(path)
                self._log("Cleaning..."); sess.clean()
                self._log("Writing SourceData..."); sess.write_to_excel(sheet_name="SourceData", visible=visible)
                self._log("Creating dashboard..."); sess.create_dashboard(dashboard_name="AutoDashboard", visible=visible, open_in_excel=open_in_excel)
                self._log("DONE"); messagebox.showinfo("Success", f"Dashboard created for:\n{path}")
            except Exception as e:
                tb = traceback.format_exc(); self._log("ERROR: " + str(e)); self._log(tb); messagebox.showerror("Error", str(e))
            finally:
                try: sess.close()
                except: pass
        finally:
            self._running = False; self.run_btn.configure(state="normal")

if __name__ == "__main__":
    root = tk.Tk(); app = LauncherApp(root); root.mainloop()
