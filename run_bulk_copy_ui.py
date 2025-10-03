import os
import sys
import threading
import subprocess
import queue
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Ensure our own stdout/stderr are UTF-8-capable (for any print in this GUI)
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

def which_python():
    return sys.executable or "python"

def find_project_python(project_root: str | None) -> str | None:
    """
    Try to find venv python inside project_root/.venv.
    Windows: .venv\Scripts\python.exe
    Unix:    .venv/bin/python or python3
    """
    if not project_root:
        return None
    venv = Path(project_root) / ".venv"
    if os.name == "nt":
        cand = venv / "Scripts" / "python.exe"
        if cand.exists():
            return str(cand)
    else:
        for name in ("python", "python3"):
            cand = venv / "bin" / name
            if cand.exists():
                return str(cand)
    return None

def venv_bin_dir(project_root: str | None) -> str | None:
    if not project_root:
        return None
    venv = Path(project_root) / ".venv"
    if os.name == "nt":
        p = venv / "Scripts"
    else:
        p = venv / "bin"
    return str(p) if p.exists() else None

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Bulk Copy SharePoint Graph - UI")
        self.geometry("920x600")
        self.resizable(True, True)

        self.proc = None
        self.q = queue.Queue()

        self.var_script = tk.StringVar(value=str(Path("bulk_copy_sharepoint_graph.py").resolve()))
        self.var_mode = tk.StringVar(value="masiva")
        self.var_excel = tk.StringVar(value="")
        self.var_sheet = tk.StringVar(value="Hoja1")
        self.var_same_file = tk.StringVar(value="masivo.pdf")
        self.var_src_dir = tk.StringVar(value="detracciones")
        self.var_ext = tk.StringVar(value=".pdf")
        self.var_dry = tk.BooleanVar(value=False)
        self.var_use_venv = tk.BooleanVar(value=True)  # NEW: run inside project .venv if available

        self._build_form()
        self._build_log()
        self._toggle_mode_fields()
        self.after(80, self._drain_queue)

    # ---------------- UI ----------------
    def _build_form(self):
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill="x")

        ttk.Label(frm, text="Script:").grid(row=0, column=0, sticky="w")
        e_script = ttk.Entry(frm, textvariable=self.var_script, width=80)
        e_script.grid(row=0, column=1, sticky="we", padx=6)
        ttk.Button(frm, text="Buscar...", command=self._pick_script).grid(row=0, column=2, padx=4)

        ttk.Label(frm, text="Modo:").grid(row=1, column=0, sticky="w", pady=(8,0))
        cmb = ttk.Combobox(frm, textvariable=self.var_mode, values=["masiva","detracciones"], width=18, state="readonly")
        cmb.grid(row=1, column=1, sticky="w", padx=6, pady=(8,0))
        cmb.bind("<<ComboboxSelected>>", lambda e: self._toggle_mode_fields())

        ttk.Label(frm, text="Excel:").grid(row=2, column=0, sticky="w", pady=(8,0))
        e_excel = ttk.Entry(frm, textvariable=self.var_excel, width=80)
        e_excel.grid(row=2, column=1, sticky="we", padx=6, pady=(8,0))
        ttk.Button(frm, text="Buscar...", command=self._pick_excel).grid(row=2, column=2, padx=4, pady=(8,0))

        ttk.Label(frm, text="Hoja (sheet):").grid(row=3, column=0, sticky="w", pady=(8,0))
        e_sheet = ttk.Entry(frm, textvariable=self.var_sheet, width=24)
        e_sheet.grid(row=3, column=1, sticky="w", padx=6, pady=(8,0))

        self.row_masiva = 4
        ttk.Label(frm, text="Archivo único (same-file):").grid(row=self.row_masiva, column=0, sticky="w", pady=(8,0))
        self.e_same = ttk.Entry(frm, textvariable=self.var_same_file, width=80)
        self.e_same.grid(row=self.row_masiva, column=1, sticky="we", padx=6, pady=(8,0))
        self.b_same = ttk.Button(frm, text="Buscar...", command=self._pick_same_file)
        self.b_same.grid(row=self.row_masiva, column=2, padx=4, pady=(8,0))

        self.row_det = 5
        ttk.Label(frm, text="Directorio origen (src-dir):").grid(row=self.row_det, column=0, sticky="w", pady=(8,0))
        self.e_src = ttk.Entry(frm, textvariable=self.var_src_dir, width=80)
        self.e_src.grid(row=self.row_det, column=1, sticky="we", padx=6, pady=(8,0))
        self.b_src = ttk.Button(frm, text="Buscar...", command=self._pick_src_dir)
        self.b_src.grid(row=self.row_det, column=2, padx=4, pady=(8,0))

        ttk.Label(frm, text="Extensión (ext):").grid(row=self.row_det+1, column=0, sticky="w", pady=(8,0))
        self.e_ext = ttk.Entry(frm, textvariable=self.var_ext, width=24)
        self.e_ext.grid(row=self.row_det+1, column=1, sticky="w", padx=6, pady=(8,0))

        ttk.Checkbutton(frm, text="Dry run (simular sin subir)", variable=self.var_dry).grid(row=self.row_det+2, column=1, sticky="w", pady=(6,0))

        # NEW: venv toggle
        ttk.Checkbutton(frm, text="Usar .venv del proyecto (si existe)", variable=self.var_use_venv).grid(row=self.row_det+3, column=1, sticky="w", pady=(6,0))

        # Barra de acciones
        bar = ttk.Frame(self, padding=(10,4))
        bar.pack(fill="x")

        self.btn_run = ttk.Button(bar, text="▶ Ejecutar Copia", command=self._run)
        self.btn_run.pack(side="left")

        self.btn_run_extractor = ttk.Button(bar, text="🧾 Ejecutar partir detracciones", command=self._run_extractor)
        self.btn_run_extractor.pack(side="left", padx=6)

        self.btn_stop = ttk.Button(bar, text="■ Detener", command=self._stop, state="disabled")
        self.btn_stop.pack(side="left", padx=6)

        for i in range(3):
            frm.columnconfigure(i, weight=1)

    def _build_log(self):
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Salida del proceso:").pack(anchor="w")
        self.txt = tk.Text(frm, height=20, wrap="word", state="disabled")
        self.txt.pack(fill="both", expand=True)
        self.scroll = ttk.Scrollbar(self.txt, command=self.txt.yview)
        self.txt.configure(yscrollcommand=self.scroll.set)

    # ------------- Helpers UI -------------
    def _toggle_mode_fields(self):
        m = self.var_mode.get()
        state_m = "normal" if m == "masiva" else "disabled"
        self.e_same.configure(state=state_m)
        self.b_same.configure(state=state_m)
        state_d = "normal" if m == "detracciones" else "disabled"
        self.e_src.configure(state=state_d)
        self.b_src.configure(state=state_d)
        self.e_ext.configure(state=state_d)

    def _pick_script(self):
        p = filedialog.askopenfilename(title="Seleccionar script", filetypes=[("Python", "*.py"),("Todos", "*.*")])
        if p: self.var_script.set(p)

    def _pick_excel(self):
        p = filedialog.askopenfilename(title="Seleccionar Excel", filetypes=[("Excel", "*.xlsx;*.xls"),("Todos", "*.*")])
        if p: self.var_excel.set(p)

    def _pick_same_file(self):
        p = filedialog.askopenfilename(title="Seleccionar archivo único")
        if p: self.var_same_file.set(p)

    def _pick_src_dir(self):
        p = filedialog.askdirectory(title="Seleccionar carpeta de origen")
        if p: self.var_src_dir.set(p)

    def _append_log(self, text):
        self.txt.configure(state="normal")
        self.txt.insert("end", text)
        self.txt.see("end")
        self.txt.configure(state="disabled")

    def _drain_queue(self):
        try:
            while True:
                line = self.q.get_nowait()
                self._append_log(line)
        except queue.Empty:
            pass
        if self.proc and self.proc.poll() is not None:
            self._append_log("\n--- Proceso finalizado ---\n")
            self.btn_run.configure(state="normal")
            self.btn_run_extractor.configure(state="normal")
            self.btn_stop.configure(state="disabled")
            self.proc = None
        self.after(80, self._drain_queue)

    # ------------- Runner core -------------
    def _runner(self, args, cwd, venv_dir):
        try:
            # Log which interpreter and venv
            self.q.put(f"\nEjecutando: {' '.join(args)}\n")
            if venv_dir:
                self.q.put(f"Usando .venv: {venv_dir}\n\n")
            else:
                self.q.put("Sin .venv (usando Python del sistema)\n\n")

            env_utf8 = dict(os.environ)
            env_utf8['PYTHONIOENCODING'] = 'utf-8'
            env_utf8['PYTHONUTF8'] = '1'
            env_utf8['PYTHONUNBUFFERED'] = '1'
            # If we have a venv, emulate activation for subprocess
            if venv_dir:
                env_utf8['VIRTUAL_ENV'] = str(Path(venv_dir).parent)  # path to .venv
                # Prepend venv bin/Scripts to PATH
                env_utf8['PATH'] = str(venv_dir) + os.pathsep + env_utf8.get('PATH', '')

            self.proc = subprocess.Popen(
                args,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding='utf-8',
                errors='replace',
                cwd=cwd,
                env=env_utf8
            )
            for line in self.proc.stdout:
                self.q.put(line)
            self.proc.wait()
        except Exception as e:
            self.q.put(f"\n[ERROR] {e}\n")

    def _choose_python(self, project_root: str | None):
        """
        Return (python_exe, venv_bin_path or None)
        """
        venv_python = None
        venv_dir = None
        if self.var_use_venv.get():
            venv_python = find_project_python(project_root)
            venv_dir = venv_bin_dir(project_root) if venv_python else None
        return (venv_python or which_python(), venv_dir)

    # ------------- Run bulk_copy... -------------
    def _run(self):
        script = self.var_script.get().strip()
        if not script or not Path(script).exists():
            messagebox.showerror("Error", "Selecciona la ruta del script bulk_copy_sharepoint_graph.py")
            return
        excel = self.var_excel.get().strip()
        if not excel or not Path(excel).exists():
            messagebox.showerror("Error", "Selecciona un archivo Excel válido")
            return
        sheet = self.var_sheet.get().strip() or None
        mode = self.var_mode.get()
        project_root = os.path.dirname(script) or None

        py, venv_dir = self._choose_python(project_root)
        args = [py, script, "--mode", mode, "--excel", excel]
        if sheet:
            args += ["--sheet", sheet]
        if self.var_dry.get():
            args += ["--dry"]

        if mode == "masiva":
            sf = self.var_same_file.get().strip()
            if not sf or not Path(sf).exists():
                messagebox.showerror("Error", "Selecciona el archivo único (same-file)")
                return
            args += ["--same-file", sf]
        else:
            src = self.var_src_dir.get().strip()
            if not src or not Path(src).exists():
                messagebox.showerror("Error", "Selecciona el directorio de origen (src-dir)")
                return
            ext = self.var_ext.get().strip() or ".pdf"
            if not ext.startswith("."):
                ext = "." + ext
            args += ["--src-dir", src, "--ext", ext]

        self._prepare_and_start(args, cwd=project_root, venv_dir=venv_dir)

    # ------------- Run extraer_detracciones.py -------------
    def _run_extractor(self):
        script = self.var_script.get().strip()
        if not script or not Path(script).exists():
            messagebox.showerror("Error", "Selecciona la ruta del script bulk_copy_sharepoint_graph.py para detectar la raíz del proyecto")
            return
        project_root = os.path.dirname(script) or None
        extractor = os.path.join(project_root, "extraer_detracciones.py")
        if not os.path.exists(extractor):
            messagebox.showerror("Error", f"No se encontró extraer_detracciones.py en: {project_root}")
            return

        py, venv_dir = self._choose_python(project_root)
        args = [py, extractor]
        self._prepare_and_start(args, cwd=project_root, venv_dir=venv_dir)

    def _prepare_and_start(self, args, cwd=None, venv_dir=None):
        # Disable run buttons, enable stop
        self.btn_run.configure(state="disabled")
        self.btn_run_extractor.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.txt.configure(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.configure(state="disabled")

        t = threading.Thread(target=self._runner, args=(args, cwd, venv_dir), daemon=True)
        t.start()

    # ------------- Stop -------------
    def _stop(self):
        if self.proc and self.proc.poll() is None:
            try:
                self.proc.terminate()
            except Exception:
                pass
            self._append_log("\n--- Señal de detención enviada ---\n")

if __name__ == "__main__":
    App().mainloop()
