from __future__ import annotations
import io, json, os, shutil, sys, threading, tkinter as tk
from contextlib import redirect_stdout, redirect_stderr
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

if getattr(sys, "frozen", False):
    APP_DIR       = Path(sys._MEIPASS) / "app"   # type: ignore[attr-defined]
    PORTABLE_ROOT = Path(sys.executable).parent
else:
    APP_DIR       = Path(__file__).resolve().parent
    PORTABLE_ROOT = APP_DIR.parent

sys.path.insert(0, str(APP_DIR / "assembler"))

USER_CONFIG          = PORTABLE_ROOT / "app" / "config"
USER_EXAMPLES        = PORTABLE_ROOT / "app" / "examples"
PROJECTS_DIR         = PORTABLE_ROOT / "projects"
EXTRA_PROJECTS_FILE  = PORTABLE_ROOT / "extra_projects.json"
BUNDLED_CONFIG   = APP_DIR / "config"
BUNDLED_EXAMPLES = APP_DIR / "examples"


def ensure_user_data():
    for d in (USER_CONFIG, USER_EXAMPLES, PROJECTS_DIR):
        d.mkdir(parents=True, exist_ok=True)
    if BUNDLED_CONFIG.exists() and not any(USER_CONFIG.glob("*.json")):
        for src in BUNDLED_CONFIG.glob("*.json"):
            (USER_CONFIG / src.name).write_bytes(src.read_bytes())
    if BUNDLED_EXAMPLES.exists() and not any(USER_EXAMPLES.iterdir()):
        shutil.copytree(BUNDLED_EXAMPLES, USER_EXAMPLES, dirs_exist_ok=True)

ensure_user_data()

BG, WHITE, ACCENT, ACCENT2 = "#F0F2F5", "#FFFFFF", "#1558D6", "#E8F0FE"
TEXT, MUTED, BORDER = "#202124", "#5F6368", "#DADCE0"
LOG_BG, LOG_FG = "#1E1E2E", "#CDD6F4"
LOG_COLORS = {"ok": "#A6E3A1", "err": "#F38BA8", "info": "#89B4FA", "warn": "#FAB387"}
LOG_PFX    = {"ok": "[OK] ", "err": "[XX] ", "warn": "[!!] ", "info": "[..] "}

NORMATIVA_LABELS = {
    "apa7":          "APA 7 - Ciencias sociales y humanidades",
    "appa_eep_2021": "APPA EEP 2021 - Facultad de Ingenieria USAC",
    "ieee":          "IEEE - Ingenieria y ciencias exactas",
    "vancouver":     "Vancouver - Ciencias de la salud",
}


def list_normativas():
    return sorted(p.stem for p in USER_CONFIG.glob("*.json") if p.stem != "project")


def load_extra_projects():
    if not EXTRA_PROJECTS_FILE.exists():
        return []
    try:
        return [Path(p) for p in json.loads(EXTRA_PROJECTS_FILE.read_text(encoding="utf-8"))]
    except Exception:
        return []


def save_extra_projects(paths):
    EXTRA_PROJECTS_FILE.write_text(
        json.dumps([str(p) for p in paths], indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def list_projects():
    seen, result = set(), []
    if PROJECTS_DIR.exists():
        for p in PROJECTS_DIR.iterdir():
            if p.is_dir() and (p / "config" / "project.json").exists():
                seen.add(p.resolve()); result.append(p)
    for p in load_extra_projects():
        if p.exists() and (p / "config" / "project.json").exists() and p.resolve() not in seen:
            seen.add(p.resolve()); result.append(p)
    return sorted(result, key=lambda p: p.stat().st_mtime, reverse=True)


def create_project(name, location, normativa, log):
    try:
        for sub in ("markdowns", "word", "assets/images", "assets/data", "config"):
            (location / sub).mkdir(parents=True, exist_ok=True)

        src = USER_CONFIG / f"{normativa}.json"
        if not src.exists():
            src = BUNDLED_CONFIG / f"{normativa}.json"
        if not src.exists():
            log(f"No se encontró normativa: {normativa}.json", "err"); return False

        dest = location / "config" / f"{normativa}.json"
        dest.write_bytes(src.read_bytes())
        log(f"Normativa copiada: {normativa}.json", "ok")

        try:
            from norm_excel import json_to_excel, create_indice_excel
            json_to_excel(dest, location / "config" / f"{normativa}.xlsx")
            log(f"Excel editable: {normativa}.xlsx", "ok")
            create_indice_excel(location)
            log("Índice editable: config/indice.xlsx", "ok")
        except Exception as e:
            log(f"Sin Excel editable: {e}", "warn")

        ex = USER_EXAMPLES / normativa
        if not ex.exists():
            ex = BUNDLED_EXAMPLES / normativa
        if ex.exists():
            for f in ex.iterdir():
                if f.is_file():
                    (location / "markdowns" / f.name).write_bytes(f.read_bytes())
                    log(f"Ejemplo: {f.name}", "ok")

        (location / "config" / "project.json").write_text(
            json.dumps({"project_name": name, "normativa": normativa}, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

        (location / ".obsidian").mkdir(exist_ok=True)
        (location / ".obsidian" / "app.json").write_text(
            json.dumps({"useMarkdownLinks": True, "attachmentFolderPath": "assets/images", "newLinkFormat": "relative"}, indent=2),
            encoding="utf-8",
        )
        log("Obsidian config: .obsidian/app.json", "ok")

        (location / ".vscode").mkdir(exist_ok=True)
        (location / ".vscode" / "settings.json").write_text(
            json.dumps({"markdown.copyFiles.destination": {"**/*": "assets/images/"}}, indent=2),
            encoding="utf-8",
        )
        log("VS Code config: .vscode/settings.json", "ok")

        log("Proyecto creado.", "ok")
        return True
    except Exception as e:
        log(f"Error: {e}", "err"); return False


def open_folder(path):
    if sys.platform == "win32":
        os.startfile(str(path))  # type: ignore[attr-defined]
    else:
        import subprocess; subprocess.Popen(["xdg-open", str(path)])


_observer = None


def start_watcher(project_dir, normativa, log):
    global _observer
    if _observer:
        log("Ya hay un watcher corriendo.", "warn"); return False
    try:
        import watcher as wmod
        wmod._set_project_root(project_dir)
        from watchdog.observers import Observer

        stream = _LogStream(log)
        with redirect_stdout(stream), redirect_stderr(stream):
            for md in sorted(wmod.MARKDOWNS_DIR.glob("*.md")):
                wmod.convert_md_to_docx(md, normativa)
        stream.flush()

        obs = Observer()
        obs.schedule(wmod.MarkdownHandler(normativa), str(wmod.MARKDOWNS_DIR), recursive=False)
        obs.schedule(wmod.WordHandler(), str(wmod.WORD_DIR), recursive=False)
        obs.start()
        _observer = obs
        log(f"Watcher activo: {project_dir.name}", "ok")
        return True
    except Exception as e:
        log(f"Error watcher: {e}", "err"); return False


def stop_watcher(log):
    global _observer
    if not _observer:
        log("Sin watcher activo.", "info"); return
    try:
        _observer.stop(); _observer.join(timeout=2)
        log("Watcher detenido.", "ok")
    except Exception as e:
        log(f"Error deteniendo watcher: {e}", "warn")
    _observer = None


def run_assembler(project_dir, normativa, log):
    try:
        import importlib, assembler
        importlib.reload(assembler)
        assembler._set_project_root(project_dir)
        stream = _LogStream(log)
        with redirect_stdout(stream), redirect_stderr(stream):
            assembler.assemble(normativa)
        stream.flush()
        log("Compilación completada.", "ok")
    except Exception as e:
        import traceback
        log(f"Error: {e}", "err")
        for line in traceback.format_exc().splitlines():
            log(line, "err")


class _LogStream(io.TextIOBase):
    def __init__(self, cb):
        self._cb, self._buf = cb, ""

    def write(self, s):
        self._buf += s
        while "\n" in self._buf:
            line, self._buf = self._buf.split("\n", 1)
            if line := line.rstrip():
                tag = "ok" if "✅" in line else "err" if "❌" in line else "warn" if "⚠" in line else "info"
                self._cb(line, tag)
        return len(s)

    def flush(self):
        if self._buf.strip():
            self._cb(self._buf.strip(), "info"); self._buf = ""


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Sistema de Gestion Editorial")
        self.geometry("780x680"); self.minsize(720, 600)
        self.configure(bg=BG)
        self._normativas: list[tuple[str, str]] = []
        self._build_ui()
        self._load_normativas()
        self._refresh_projects()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.log(f"Listo. Carpeta: {PORTABLE_ROOT}", "ok")

    def _build_ui(self):
        hdr = tk.Frame(self, bg=ACCENT, height=64)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text="  Sistema de Gestion Editorial",
                 font=("Segoe UI", 15, "bold"), bg=ACCENT, fg=WHITE).pack(side="left", padx=20, pady=16)
        tk.Label(hdr, text="v1.0", font=("Segoe UI", 10),
                 bg=ACCENT, fg="#A8C7FA").pack(side="right", padx=20)

        sty = ttk.Style(self)
        try: sty.theme_use("clam")
        except Exception: pass
        sty.configure("TNotebook", background=BG, borderwidth=0)
        sty.configure("TNotebook.Tab", padding=(20, 10), font=("Segoe UI", 10, "bold"),
                       background=BG, foreground=MUTED)
        sty.map("TNotebook.Tab", background=[("selected", WHITE)], foreground=[("selected", ACCENT)])

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=20, pady=(16, 8))
        t1, t2 = tk.Frame(nb, bg=WHITE), tk.Frame(nb, bg=WHITE)
        nb.add(t1, text="  Nuevo proyecto  "); nb.add(t2, text="  Proyectos existentes  ")
        self._build_tab_new(t1); self._build_tab_existing(t2)

        tk.Label(self, text="Registro", font=("Segoe UI", 9, "bold"),
                 bg=BG, fg=MUTED).pack(anchor="w", padx=20)
        lf = tk.Frame(self, bg=LOG_BG, highlightbackground=BORDER, highlightthickness=1)
        lf.pack(fill="both", padx=20, pady=(0, 16))
        lf.configure(height=180); lf.pack_propagate(False)
        self._log_w = tk.Text(lf, state="disabled", font=("Monospace", 9),
                              bg=LOG_BG, fg=LOG_FG, relief="flat", wrap="word", padx=10, pady=8)
        sb = tk.Scrollbar(lf, command=self._log_w.yview, bg=LOG_BG, relief="flat")
        self._log_w.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y"); self._log_w.pack(fill="both", expand=True)
        for tag, fg in LOG_COLORS.items():
            self._log_w.tag_config(tag, foreground=fg)

    def _btn(self, parent, text, cmd, primary=False):
        kw = dict(font=("Segoe UI", 10, "bold"), relief="flat", padx=14, pady=8, cursor="hand2")
        kw.update((dict(bg=ACCENT, fg=WHITE, activebackground="#1248C0", activeforeground=WHITE))
                  if primary else dict(bg=ACCENT2, fg=ACCENT))
        return tk.Button(parent, text=text, command=cmd, **kw)

    def _entry(self, parent, var):
        return tk.Entry(parent, textvariable=var, font=("Segoe UI", 11), relief="flat",
                        bg=BG, fg=TEXT, highlightbackground=BORDER, highlightthickness=1)

    def _build_tab_new(self, p):
        tk.Label(p, text="Crear un proyecto nuevo", font=("Segoe UI", 13, "bold"),
                 bg=WHITE, fg=TEXT).pack(anchor="w", padx=24, pady=(20, 4))
        tk.Label(p, text="Genera la carpeta con capítulos, ejemplos y normativa.",
                 font=("Segoe UI", 9), bg=WHITE, fg=MUTED).pack(fill="x", padx=24, pady=(0, 16))

        for lbl in ("Nombre del proyecto", "Ubicación", "Normativa"):
            tk.Label(p, text=lbl, font=("Segoe UI", 9, "bold"), bg=WHITE, fg=MUTED, anchor="w"
                     ).pack(fill="x", padx=24, pady=(0, 4))
            if lbl == "Nombre del proyecto":
                self._name_var = tk.StringVar(value="Mi_Tesis")
                self._entry(p, self._name_var).pack(fill="x", padx=24, pady=(0, 12))
                self._name_var.trace_add("write", self._sync_location)
            elif lbl == "Ubicación":
                row = tk.Frame(p, bg=WHITE); row.pack(fill="x", padx=24, pady=(0, 12))
                self._loc_var = tk.StringVar(value=str(PROJECTS_DIR / "Mi_Tesis"))
                self._entry(row, self._loc_var).pack(side="left", fill="x", expand=True)
                tk.Button(row, text="Examinar", font=("Segoe UI", 9), bg=ACCENT2, fg=ACCENT,
                          relief="flat", padx=12, pady=4, cursor="hand2",
                          command=self._browse).pack(side="left", padx=(8, 0))
            else:
                self._norm_var = tk.StringVar()
                self._norm_cb = ttk.Combobox(p, textvariable=self._norm_var,
                                             font=("Segoe UI", 11), state="readonly")
                self._norm_cb.pack(fill="x", padx=24, pady=(0, 12))

        self._btn(p, "Crear proyecto", self._on_create, primary=True).pack(fill="x", padx=24, pady=(8, 20))

    def _build_tab_existing(self, p):
        tk.Label(p, text="Tus proyectos", font=("Segoe UI", 13, "bold"),
                 bg=WHITE, fg=TEXT).pack(anchor="w", padx=24, pady=(20, 4))
        tk.Label(p, text="Selecciona un proyecto y elige una acción.",
                 font=("Segoe UI", 9), bg=WHITE, fg=MUTED).pack(fill="x", padx=24, pady=(0, 16))

        lf = tk.Frame(p, bg=WHITE, highlightbackground=BORDER, highlightthickness=1)
        lf.pack(fill="both", expand=True, padx=24, pady=(0, 12))
        self._proj_lb = tk.Listbox(lf, font=("Segoe UI", 10), bg=WHITE, fg=TEXT,
                                   selectbackground=ACCENT2, selectforeground=ACCENT,
                                   relief="flat", borderwidth=0, activestyle="none")
        sb = tk.Scrollbar(lf, command=self._proj_lb.yview)
        self._proj_lb.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y"); self._proj_lb.pack(fill="both", expand=True, padx=8, pady=8)

        acts = tk.Frame(p, bg=WHITE); acts.pack(fill="x", padx=24, pady=(0, 20))
        self._btn(acts, "Abrir carpeta", self._on_open).pack(side="left", padx=(0, 8))
        self._btn_w = self._btn(acts, "Iniciar watcher", self._on_toggle_watcher)
        self._btn_w.pack(side="left", padx=(0, 8))
        self._btn(acts, "Compilar Word", self._on_compile, primary=True).pack(side="left", padx=(0, 8))
        tk.Button(acts, text="Refrescar", font=("Segoe UI", 9), bg=WHITE, fg=MUTED,
                  relief="flat", padx=10, pady=8, cursor="hand2",
                  command=self._refresh_projects).pack(side="right")
        tk.Button(acts, text="Agregar", font=("Segoe UI", 9), bg=WHITE, fg=MUTED,
                  relief="flat", padx=10, pady=8, cursor="hand2",
                  command=self._on_add_project).pack(side="right", padx=(0, 4))

    def log(self, msg, tag="info"):
        self._log_w.config(state="normal")
        self._log_w.insert("end", LOG_PFX.get(tag, "  ") + msg + "\n", tag)
        self._log_w.see("end"); self._log_w.config(state="disabled")
        self.update_idletasks()

    def _load_normativas(self):
        names = list_normativas()
        self._normativas = [(n, NORMATIVA_LABELS.get(n, n.upper())) for n in names]
        labels = [l for _, l in self._normativas]
        self._norm_cb["values"] = labels
        if labels:
            self._norm_cb.current(0); self.log(f"{len(labels)} normativas.", "ok")
        else:
            self._norm_cb.set("(sin normativas)"); self.log("Sin normativas.", "warn")

    def _refresh_projects(self):
        self._proj_lb.delete(0, "end")
        for p in list_projects():
            try: norm = json.loads((p / "config" / "project.json").read_text(encoding="utf-8")).get("normativa", "?")
            except Exception: norm = "?"
            self._proj_lb.insert("end", f"  {p.name}  [{norm}]  ({p})")
        if not self._proj_lb.size():
            self._proj_lb.insert("end", "  (sin proyectos)")
            self._proj_lb.itemconfigure(0, foreground=MUTED)

    def _selected(self):
        sel = self._proj_lb.curselection()
        if not sel:
            messagebox.showinfo("Selección", "Elige un proyecto."); return None
        projects = list_projects()
        return projects[sel[0]] if sel[0] < len(projects) else None

    def _normativa_of(self, p):
        try: return json.loads((p / "config" / "project.json").read_text(encoding="utf-8")).get("normativa", "apa7")
        except Exception: return "apa7"

    def _browse(self):
        path = filedialog.askdirectory(initialdir=str(PROJECTS_DIR))
        if path: self._loc_var.set(path)

    def _sync_location(self, *_):
        name = self._name_var.get().strip().replace(" ", "_") or "Mi_Tesis"
        if str(PROJECTS_DIR) in self._loc_var.get():
            self._loc_var.set(str(PROJECTS_DIR / name))

    def _on_create(self):
        name = self._name_var.get().strip().replace(" ", "_")
        if not name: messagebox.showwarning("Falta dato", "Escribe el nombre."); return
        location = Path(self._loc_var.get().strip())
        if not self._normativas: messagebox.showerror("Error", "Sin normativas."); return
        norm_id = next((n for n, l in self._normativas if l == self._norm_var.get()), self._normativas[0][0])
        if location.exists() and any(location.iterdir()):
            if not messagebox.askyesno("Carpeta no vacía", f"{location}\n¿Continuar?"): return
        self.log(f"Creando '{name}'...", "info")
        threading.Thread(target=self._create_bg, args=(name, location, norm_id), daemon=True).start()

    def _create_bg(self, name, location, norm_id):
        cb = lambda m, t: self.after(0, self.log, m, t)
        ok = create_project(name, location, norm_id, cb)
        self.after(0, self._refresh_projects)
        if ok: self.after(0, messagebox.showinfo, "Listo", f"Proyecto en:\n{location}")

    def _on_open(self):
        p = self._selected()
        if p: open_folder(p)

    def _on_toggle_watcher(self):
        if _observer:
            stop_watcher(self.log); self._btn_w.config(text="Iniciar watcher"); return
        p = self._selected()
        if not p: return
        norm = self._normativa_of(p)
        cb = lambda m, t: self.after(0, self.log, m, t)
        def run():
            if start_watcher(p, norm, cb):
                self.after(0, lambda: self._btn_w.config(text="Detener watcher"))
        threading.Thread(target=run, daemon=True).start()

    def _on_compile(self):
        p = self._selected()
        if not p: return
        cb = lambda m, t: self.after(0, self.log, m, t)
        threading.Thread(target=run_assembler, args=(p, self._normativa_of(p), cb), daemon=True).start()

    def _on_add_project(self):
        path = filedialog.askdirectory(title="Seleccionar carpeta del proyecto")
        if not path:
            return
        p = Path(path)
        if not (p / "config" / "project.json").exists():
            messagebox.showerror("No válido", "No se encontró config/project.json en esa carpeta.")
            return
        extras = load_extra_projects()
        if p not in extras:
            extras.append(p)
            save_extra_projects(extras)
        self._refresh_projects()
        self.log(f"Proyecto agregado: {p.name}", "ok")

    def _on_close(self):
        if _observer: stop_watcher(lambda m, t: None)
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
