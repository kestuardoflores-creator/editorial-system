"""
installer_gui.py — Sistema de Gestion Editorial
GUI de instalacion para Windows. Llamado por installer.bat.
"""

import sys
import os
import json
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

# ── Auto-install dependencies ──────────────────────────────────────────────────
def _pip(pkg):
    subprocess.check_call(
        [sys.executable, "-m", "pip", "install", "--quiet", pkg],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )

try:
    import requests
except ImportError:
    _pip("requests")
    import requests

try:
    import openpyxl
except ImportError:
    _pip("openpyxl")

# ── Paths ──────────────────────────────────────────────────────────────────────
INSTALLER_DIR = Path(__file__).parent
ROOT_DIR      = INSTALLER_DIR.parent

GITHUB_REPO   = "kestuardoflores-creator/editorial-system"
GITHUB_BRANCH = "main"
GITHUB_RAW    = f"https://raw.githubusercontent.com/{GITHUB_REPO}/{GITHUB_BRANCH}"
GITHUB_API    = f"https://api.github.com/repos/{GITHUB_REPO}/contents"

NORMATIVA_LABELS = {
    "apa7":          "APA 7ma Edicion",
    "appa_eep_2021": "APPA EEP 2021 - Facultad de Ingenieria USAC",
    "ieee":          "IEEE",
    "vancouver":     "Vancouver",
}

# ── Colors ─────────────────────────────────────────────────────────────────────
C_BG      = "#F0F2F5"
C_WHITE   = "#FFFFFF"
C_ACCENT  = "#1558D6"
C_ACCENT2 = "#E8F0FE"
C_TEXT    = "#202124"
C_MUTED   = "#5F6368"
C_BORDER  = "#DADCE0"
C_OK      = "#1E8E3E"
C_WARN    = "#F29900"
C_ERR     = "#D93025"
C_LOG_BG  = "#1E1E2E"
C_LOG_FG  = "#CDD6F4"
C_LOG_OK  = "#A6E3A1"
C_LOG_ERR = "#F38BA8"
C_LOG_INF = "#89B4FA"
C_LOG_WRN = "#FAB387"


# ─────────────────────────────────────────────────────────────────────────────
# INSTALLER LOGIC (no colorama, no input() — pure functions)
# ─────────────────────────────────────────────────────────────────────────────

def fetch_normativas():
    try:
        r = requests.get(f"{GITHUB_API}/config", timeout=10)
        if r.status_code == 200:
            return [f["name"].replace(".json", "")
                    for f in r.json() if f["name"].endswith(".json")]
    except Exception:
        pass
    return ["apa7", "appa_eep_2021"]


def download_file(url, dest_path):
    r = requests.get(url, timeout=20)
    r.raise_for_status()
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    dest_path.write_bytes(r.content)


def run_installation(project_name, project_dir, normativa_id, log_cb):
    """
    Full installation logic. log_cb(msg, tag) sends lines to the GUI log.
    Tags: "ok", "err", "warn", "info"
    """
    try:
        # 1. Create folder structure
        log_cb("Creando estructura de carpetas...", "info")
        for folder in ["markdowns", "word", "assets/images", "assets/data", "config"]:
            (project_dir / folder).mkdir(parents=True, exist_ok=True)
        log_cb("  Carpetas creadas.", "ok")

        # 2. Download normativa JSON
        log_cb(f"Descargando normativa {normativa_id}...", "info")
        json_url  = f"{GITHUB_RAW}/config/{normativa_id}.json"
        json_dest = project_dir / "config" / f"{normativa_id}.json"
        download_file(json_url, json_dest)
        log_cb(f"  config/{normativa_id}.json descargado.", "ok")

        # 3. Convert JSON -> Excel
        log_cb("Convirtiendo normativa a Excel...", "info")
        try:
            assembler_dir = ROOT_DIR / "assembler"
            sys.path.insert(0, str(assembler_dir))
            from norm_excel import json_to_excel
            xlsx_dest = project_dir / "config" / f"{normativa_id}.xlsx"
            json_to_excel(json_dest, xlsx_dest)
            log_cb(f"  config/{normativa_id}.xlsx creado (editable en Excel).", "ok")
        except Exception as e:
            log_cb(f"  No se pudo crear Excel: {e}", "warn")

        # 4. Download examples
        log_cb(f"Descargando ejemplos para '{normativa_id}'...", "info")
        try:
            r = requests.get(f"{GITHUB_API}/examples/{normativa_id}", timeout=10)
            if r.status_code == 200:
                files = r.json()
                for fi in files:
                    url  = f"{GITHUB_RAW}/examples/{normativa_id}/{fi['name']}"
                    dest = project_dir / "markdowns" / fi["name"]
                    download_file(url, dest)
                    log_cb(f"  markdowns/{fi['name']}", "ok")
            else:
                log_cb("  No se encontraron ejemplos en GitHub.", "warn")
        except Exception as e:
            log_cb(f"  Error descargando ejemplos: {e}", "warn")

        # 5. Download assembler files
        log_cb("Descargando archivos del ensamblador...", "info")
        asm_dir = project_dir.parent / "assembler" \
                  if (project_dir.parent / "assembler").exists() \
                  else ROOT_DIR / "assembler"
        asm_dir.mkdir(exist_ok=True)

        for fname in ["assembler.py", "watcher.py", "requirements.txt", "norm_excel.py"]:
            dest = asm_dir / fname
            if not dest.exists():
                try:
                    download_file(f"{GITHUB_RAW}/assembler/{fname}", dest)
                    log_cb(f"  assembler/{fname}", "ok")
                except Exception as e:
                    log_cb(f"  No se pudo descargar {fname}: {e}", "warn")

        # 6. Install pip dependencies
        req_file = asm_dir / "requirements.txt"
        if req_file.exists():
            log_cb("Instalando dependencias Python...", "info")
            try:
                subprocess.check_call(
                    [sys.executable, "-m", "pip", "install", "-r", str(req_file), "--quiet"],
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                )
                log_cb("  Dependencias instaladas.", "ok")
            except Exception as e:
                log_cb(f"  Advertencia pip: {e}", "warn")

        # 7. Save project config
        cfg_path = project_dir / "config" / "project.json"
        cfg_path.write_text(
            json.dumps({"project_name": project_name, "normativa": normativa_id},
                       indent=2, ensure_ascii=False),
            encoding="utf-8"
        )

        log_cb("", "info")
        log_cb("Instalacion completada exitosamente.", "ok")
        return True

    except Exception as e:
        log_cb(f"Error inesperado: {e}", "err")
        return False


# ─────────────────────────────────────────────────────────────────────────────
# GUI
# ─────────────────────────────────────────────────────────────────────────────

class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Sistema de Gestion Editorial — Instalador")
        self.geometry("640x620")
        self.minsize(640, 620)
        self.configure(bg=C_BG)
        self._normativas = []
        self._install_done = False
        self._build_ui()
        self.after(300, self._load_normativas)

    # ── UI ─────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Header ────────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=C_ACCENT, height=64)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(
            hdr, text="Sistema de Gestion Editorial",
            font=("Segoe UI", 15, "bold"), bg=C_ACCENT, fg=C_WHITE
        ).pack(side="left", padx=20, pady=16)
        tk.Label(
            hdr, text="v1.0",
            font=("Segoe UI", 10), bg=C_ACCENT, fg="#A8C7FA"
        ).pack(side="left", pady=22)

        # ── Card ──────────────────────────────────────────────────────────────
        card = tk.Frame(self, bg=C_WHITE, relief="flat",
                        highlightbackground=C_BORDER, highlightthickness=1)
        card.pack(fill="x", padx=20, pady=(16, 8))

        pad = {"padx": 20, "pady": (0, 12)}

        # Step indicators
        steps = tk.Frame(card, bg=C_WHITE)
        steps.pack(fill="x", padx=20, pady=(16, 12))
        for i, label in enumerate(["1  Nombre", "2  Ubicacion", "3  Normativa", "4  Instalar"], 1):
            col = C_ACCENT if i == 1 else C_MUTED
            tk.Label(steps, text=label, font=("Segoe UI", 9, "bold"),
                     bg=C_WHITE, fg=col).pack(side="left", padx=(0, 24))
        self._step_labels = steps.winfo_children()

        tk.Frame(card, bg=C_BORDER, height=1).pack(fill="x", padx=20, pady=(0, 16))

        # Project name
        self._field_label(card, "Nombre del proyecto")
        self._name_var = tk.StringVar(value="Mi_Tesis")
        name_entry = tk.Entry(
            card, textvariable=self._name_var,
            font=("Segoe UI", 11), relief="flat",
            bg=C_BG, fg=C_TEXT, insertbackground=C_TEXT,
            highlightbackground=C_BORDER, highlightthickness=1,
        )
        name_entry.pack(fill="x", **pad)
        self._name_var.trace_add("write", self._sync_location)

        # Location
        self._field_label(card, "Ubicacion del proyecto")
        loc_row = tk.Frame(card, bg=C_WHITE)
        loc_row.pack(fill="x", padx=20, pady=(0, 12))
        self._loc_var = tk.StringVar(value=str(Path.home() / "Documents" / "Mi_Tesis"))
        tk.Entry(
            loc_row, textvariable=self._loc_var,
            font=("Segoe UI", 11), relief="flat",
            bg=C_BG, fg=C_TEXT, insertbackground=C_TEXT,
            highlightbackground=C_BORDER, highlightthickness=1,
        ).pack(side="left", fill="x", expand=True)
        tk.Button(
            loc_row, text="Examinar", font=("Segoe UI", 9),
            bg=C_ACCENT2, fg=C_ACCENT, relief="flat",
            padx=10, pady=4, cursor="hand2",
            command=self._browse,
        ).pack(side="left", padx=(8, 0))

        # Normativa
        self._field_label(card, "Normativa")
        self._norm_var = tk.StringVar(value="Cargando normativas...")
        self._norm_cb = ttk.Combobox(
            card, textvariable=self._norm_var,
            font=("Segoe UI", 11), state="disabled",
        )
        self._norm_cb.pack(fill="x", **pad)

        # Install button
        self._btn = tk.Button(
            card,
            text="Instalar proyecto",
            font=("Segoe UI", 12, "bold"),
            bg=C_ACCENT, fg=C_WHITE,
            activebackground="#1248C0",
            activeforeground=C_WHITE,
            relief="flat", padx=0, pady=10,
            state="disabled",
            cursor="hand2",
            command=self._start_install,
        )
        self._btn.pack(fill="x", padx=20, pady=(4, 20))

        # ── Log ───────────────────────────────────────────────────────────────
        log_hdr = tk.Frame(self, bg=C_BG)
        log_hdr.pack(fill="x", padx=20, pady=(0, 4))
        tk.Label(log_hdr, text="Progreso de instalacion",
                 font=("Segoe UI", 9, "bold"), bg=C_BG, fg=C_MUTED
                 ).pack(side="left")

        log_frame = tk.Frame(self, bg=C_LOG_BG,
                             highlightbackground=C_BORDER, highlightthickness=1)
        log_frame.pack(fill="both", expand=True, padx=20, pady=(0, 16))

        self._log = tk.Text(
            log_frame, state="disabled",
            font=("Consolas", 9), bg=C_LOG_BG, fg=C_LOG_FG,
            relief="flat", wrap="word", padx=10, pady=8,
            insertbackground=C_LOG_FG,
        )
        scrollbar = tk.Scrollbar(log_frame, command=self._log.yview,
                                 bg=C_LOG_BG, troughcolor=C_LOG_BG,
                                 relief="flat", bd=0)
        self._log.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self._log.pack(fill="both", expand=True)

        self._log.tag_config("ok",   foreground=C_LOG_OK)
        self._log.tag_config("err",  foreground=C_LOG_ERR)
        self._log.tag_config("warn", foreground=C_LOG_WRN)
        self._log.tag_config("info", foreground=C_LOG_INF)
        self._log.tag_config("dim",  foreground="#6C7086")

    def _field_label(self, parent, text):
        tk.Label(parent, text=text, font=("Segoe UI", 9, "bold"),
                 bg=C_WHITE, fg=C_MUTED, anchor="w"
                 ).pack(fill="x", padx=20, pady=(0, 4))

    # ── Helpers ────────────────────────────────────────────────────────────────
    def _browse(self):
        path = filedialog.askdirectory(title="Selecciona ubicacion del proyecto")
        if path:
            self._loc_var.set(path)

    def _sync_location(self, *_):
        name = self._name_var.get().strip().replace(" ", "_")
        if "Documents" in self._loc_var.get():
            self._loc_var.set(str(Path.home() / "Documents" / name))

    def _log_write(self, msg, tag="info"):
        self._log.config(state="normal")
        prefix = {"ok": "  +  ", "err": "  x  ",
                  "warn": "  !  ", "info": "  >  "}.get(tag, "     ")
        self._log.insert("end", prefix + msg + "\n", tag)
        self._log.see("end")
        self._log.config(state="disabled")
        self.update_idletasks()

    # ── Load normativas ────────────────────────────────────────────────────────
    def _load_normativas(self):
        self._log_write("Consultando normativas en GitHub...", "info")
        threading.Thread(target=self._fetch, daemon=True).start()

    def _fetch(self):
        names  = fetch_normativas()
        labels = [NORMATIVA_LABELS.get(n, n.upper()) for n in names]
        self._normativas = list(zip(names, labels))
        self.after(0, lambda: self._set_normativas(labels))

    def _set_normativas(self, labels):
        self._norm_cb["values"] = labels
        self._norm_cb["state"]  = "readonly"
        if labels:
            self._norm_cb.current(0)
        self._btn["state"] = "normal"
        self._log_write(f"{len(labels)} normativas disponibles.", "ok")

    # ── Install ────────────────────────────────────────────────────────────────
    def _start_install(self):
        name     = self._name_var.get().strip().replace(" ", "_")
        location = self._loc_var.get().strip()
        label    = self._norm_var.get()
        norm_id  = next((n for n, l in self._normativas if l == label), "apa7")

        if not name:
            messagebox.showwarning("Falta dato", "Escribe el nombre del proyecto.")
            return
        if not location:
            messagebox.showwarning("Falta dato", "Selecciona la ubicacion.")
            return

        project_dir = Path(location)

        self._btn["state"] = "disabled"
        self._log_write("", "dim")
        self._log_write(f"Proyecto  : {name}", "info")
        self._log_write(f"Ubicacion : {location}", "info")
        self._log_write(f"Normativa : {label}", "info")
        self._log_write("", "dim")

        threading.Thread(
            target=self._run,
            args=(name, project_dir, norm_id),
            daemon=True,
        ).start()

    def _run(self, name, project_dir, norm_id):
        def cb(msg, tag):
            self.after(0, lambda m=msg, t=tag: self._log_write(m, t))

        success = run_installation(name, project_dir, norm_id, cb)
        self.after(0, lambda: self._on_done(success, project_dir))

    def _on_done(self, success, project_dir):
        if success:
            self._btn.config(
                text="Instalado correctamente",
                bg=C_OK, state="disabled"
            )
            messagebox.showinfo(
                "Instalacion completa",
                f"Tu proyecto fue creado en:\n{project_dir}\n\n"
                "Siguiente paso:\n"
                "  python assembler/assembler.py"
            )
        else:
            self._btn.config(state="normal")
            messagebox.showerror(
                "Error", "La instalacion tuvo errores.\nRevisa el log."
            )


# ── Entry point ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    App().mainloop()
