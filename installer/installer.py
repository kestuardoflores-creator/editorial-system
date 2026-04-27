"""
installer.py — Sistema de Gestión Editorial
Instalador interactivo para Windows.

Uso:
    python installer/installer.py
"""

import os
import sys
import json
import subprocess
from pathlib import Path

# ── Auto-install minimal deps ──────────────────────────────────────────────────
def _pip(pkg):
    subprocess.check_call([sys.executable, "-m", "pip", "install",
                           "--quiet", pkg])

try:
    import requests
    from colorama import Fore, Style, init
except ImportError:
    print("Preparando instalador...")
    _pip("requests")
    _pip("colorama")
    import requests
    from colorama import Fore, Style, init

init(autoreset=True)
if os.name == "nt":
    os.system("chcp 65001 > nul")

# ── GitHub config ──────────────────────────────────────────────────────────────
GITHUB_REPO    = "kestuardoflores-creator/editorial-system"
GITHUB_BRANCH  = "main"
GITHUB_RAW     = f"https://raw.githubusercontent.com/{GITHUB_REPO}/{GITHUB_BRANCH}"
GITHUB_API     = f"https://api.github.com/repos/{GITHUB_REPO}/contents"

# ── Local paths ────────────────────────────────────────────────────────────────
INSTALLER_DIR = Path(__file__).parent
ROOT_DIR      = INSTALLER_DIR.parent

# ─────────────────────────────────────────────────────────────────────────────
# UI HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def clear():
    os.system("cls" if os.name == "nt" else "clear")

def banner():
    print(Fore.CYAN + """
  +--------------------------------------------------+
  |      Sistema de Gestion Editorial v1.0           |
  |      Markdown + Python + Word                    |
  +--------------------------------------------------+
""" + Style.RESET_ALL)

def section(title):
    print(Fore.YELLOW + f"\n  -- {title} " + "-" * (44 - len(title)) + Style.RESET_ALL)

def ok(msg):
    print(Fore.GREEN + f"  ✔  {msg}" + Style.RESET_ALL)

def info(msg):
    print(Fore.WHITE + f"  →  {msg}" + Style.RESET_ALL)

def warn(msg):
    print(Fore.YELLOW + f"  ⚠  {msg}" + Style.RESET_ALL)

def err(msg):
    print(Fore.RED + f"  ✘  {msg}" + Style.RESET_ALL)

def ask(prompt, default=None):
    if default:
        label = Fore.CYAN + f"  {prompt} [{default}]: " + Style.RESET_ALL
    else:
        label = Fore.CYAN + f"  {prompt}: " + Style.RESET_ALL
    val = input(label).strip()
    return val if val else default

def choose(prompt, options):
    """Present a numbered menu and return the chosen option."""
    print(Fore.CYAN + f"\n  {prompt}\n" + Style.RESET_ALL)
    for i, opt in enumerate(options, 1):
        print(Fore.WHITE + f"    [{i}] {opt['label']}" + Style.RESET_ALL)
        if opt.get("desc"):
            print(Fore.WHITE + Style.DIM + f"        {opt['desc']}" + Style.RESET_ALL)
    print()

    while True:
        raw = input(Fore.CYAN + "  Tu elección: " + Style.RESET_ALL).strip()
        if raw.isdigit() and 1 <= int(raw) <= len(options):
            return options[int(raw) - 1]
        warn("Ingresa un número válido.")

def progress(msg):
    print(Fore.WHITE + Style.DIM + f"  ...  {msg}" + Style.RESET_ALL, end="\r")

# ─────────────────────────────────────────────────────────────────────────────
# GITHUB HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def fetch_available_normativas():
    """List .json files in config/ from GitHub to show available normatives."""
    try:
        r = requests.get(f"{GITHUB_API}/config", timeout=10)
        if r.status_code == 200:
            files = r.json()
            return [f["name"].replace(".json", "") for f in files
                    if f["name"].endswith(".json")]
    except Exception:
        pass
    # Fallback if GitHub is unavailable
    return ["apa7"]


def download_file(url, dest_path):
    """Download a single file from GitHub raw URL."""
    r = requests.get(url, timeout=15)
    r.raise_for_status()
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    dest_path.write_bytes(r.content)


def download_normativa_config(normativa, dest_dir):
    """Download config/{normativa}.json from GitHub, then convert to Excel."""
    url       = f"{GITHUB_RAW}/config/{normativa}.json"
    json_dest = dest_dir / "config" / f"{normativa}.json"
    xlsx_dest = dest_dir / "config" / f"{normativa}.xlsx"

    progress(f"Descargando normativa {normativa}.json...")
    try:
        download_file(url, json_dest)
        ok(f"Normativa descargada: config/{normativa}.json")
    except Exception as e:
        err(f"No se pudo descargar la normativa: {e}")
        return False

    # Convert JSON → Excel for user-friendly editing
    progress("Convirtiendo normativa a Excel...")
    try:
        sys.path.insert(0, str(ROOT_DIR / "assembler"))
        from norm_excel import json_to_excel
        json_to_excel(json_dest, xlsx_dest)
        ok(f"Excel editable creado: config/{normativa}.xlsx")
        info("Puedes editar los estilos abriendo ese archivo en Excel.")
    except Exception as e:
        warn(f"No se pudo crear el Excel: {e}")

    return True


def download_examples(normativa, dest_dir):
    """Download all example files for the selected normative."""
    try:
        r = requests.get(f"{GITHUB_API}/examples/{normativa}", timeout=10)
        if r.status_code != 200:
            warn(f"No se encontraron ejemplos para '{normativa}' en GitHub.")
            return

        files = r.json()
        info(f"Descargando {len(files)} ejemplos para '{normativa}'...")

        for file_info in files:
            filename = file_info["name"]
            url      = f"{GITHUB_RAW}/examples/{normativa}/{filename}"
            dest     = dest_dir / "markdowns" / filename
            progress(f"Descargando {filename}...")
            try:
                download_file(url, dest)
                ok(f"Ejemplo descargado: markdowns/{filename}")
            except Exception as e:
                warn(f"No se pudo descargar {filename}: {e}")

    except Exception as e:
        warn(f"Error al descargar ejemplos: {e}")


def download_assembler_files():
    """Download assembler scripts from GitHub into local assembler/ folder."""
    assembler_dir = ROOT_DIR / "assembler"
    assembler_dir.mkdir(exist_ok=True)

    files = ["assembler.py", "watcher.py", "requirements.txt"]
    for f in files:
        url  = f"{GITHUB_RAW}/assembler/{f}"
        dest = assembler_dir / f
        if not dest.exists():
            progress(f"Descargando assembler/{f}...")
            try:
                download_file(url, dest)
                ok(f"assembler/{f}")
            except Exception as e:
                warn(f"No se pudo descargar assembler/{f}: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# PROJECT STRUCTURE
# ─────────────────────────────────────────────────────────────────────────────

def create_project_structure(project_dir):
    """Create all required folders for a new project."""
    folders = [
        "markdowns",
        "word",
        "assets/images",
        "assets/data",
        "config",
    ]
    for folder in folders:
        (project_dir / folder).mkdir(parents=True, exist_ok=True)
    ok("Estructura de carpetas creada")


def install_pip_requirements():
    """Install Python dependencies from requirements.txt."""
    req_file = ROOT_DIR / "assembler" / "requirements.txt"
    if not req_file.exists():
        warn("No se encontró requirements.txt. Instala las dependencias manualmente.")
        return
    info("Instalando dependencias de Python...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install",
                               "-r", str(req_file), "--quiet"])
        ok("Dependencias instaladas correctamente")
    except subprocess.CalledProcessError:
        warn("Algunas dependencias no se pudieron instalar. Ejecuta:")
        warn(f"  pip install -r {req_file}")


def save_project_config(project_dir, project_name, normativa):
    """Save project metadata for the assembler to read."""
    config = {
        "project_name": project_name,
        "normativa": normativa,
    }
    path = project_dir / "config" / "project.json"
    path.write_text(json.dumps(config, indent=2, ensure_ascii=False), encoding="utf-8")
    ok("Configuración del proyecto guardada")


# ─────────────────────────────────────────────────────────────────────────────
# NORMATIVA DISPLAY NAMES
# ─────────────────────────────────────────────────────────────────────────────

NORMATIVA_INFO = {
    "apa7": {
        "label": "APA 7ª Edición",
        "desc":  "American Psychological Association, 7ma edición. Ciencias sociales, psicología, educación."
    },
    "ieee": {
        "label": "IEEE",
        "desc":  "Institute of Electrical and Electronics Engineers. Ingeniería, tecnología, ciencias exactas."
    },
    "vancouver": {
        "label": "Vancouver",
        "desc":  "Estilo numérico para ciencias de la salud y medicina."
    },
}

def build_normativa_options(names):
    options = []
    for n in names:
        info_data = NORMATIVA_INFO.get(n, {"label": n.upper(), "desc": ""})
        options.append({"id": n, "label": info_data["label"], "desc": info_data["desc"]})
    return options


# ─────────────────────────────────────────────────────────────────────────────
# MAIN INSTALLER
# ─────────────────────────────────────────────────────────────────────────────

def main():
    clear()
    banner()

    print(Fore.WHITE + """
  Bienvenido al instalador del Sistema de Gestión Editorial.
  Este asistente configurará tu proyecto en pocos pasos.
""" + Style.RESET_ALL)

    # ── Step 1: Project name ───────────────────────────────────────────────────
    section("PASO 1 — Nombre del Proyecto")
    project_name = ask("Nombre del proyecto", default="Mi_Tesis")
    project_name = project_name.replace(" ", "_")

    # ── Step 2: Location ───────────────────────────────────────────────────────
    section("PASO 2 — Ubicación")
    default_loc = str(Path.home() / "Documents" / project_name)
    location    = ask("¿Dónde crear el proyecto?", default=default_loc)
    project_dir = Path(location)

    if project_dir.exists() and any(project_dir.iterdir()):
        warn(f"La carpeta ya existe y no está vacía: {project_dir}")
        cont = ask("¿Continuar de todas formas? (s/n)", default="s")
        if cont.lower() != "s":
            print(Fore.YELLOW + "\n  Instalación cancelada.\n")
            sys.exit(0)

    # ── Step 3: Normativa ──────────────────────────────────────────────────────
    section("PASO 3 — Normativa")
    info("Consultando normativas disponibles en GitHub...")

    available = fetch_available_normativas()
    options   = build_normativa_options(available)

    if not options:
        err("No se pudo obtener la lista de normativas.")
        err("Verifica tu conexión a internet.")
        sys.exit(1)

    chosen = choose("Selecciona la normativa de tu documento:", options)
    normativa_id = chosen["id"]
    ok(f"Normativa seleccionada: {chosen['label']}")

    # ── Step 4: Confirm ────────────────────────────────────────────────────────
    section("PASO 4 — Confirmación")
    print(Fore.WHITE + f"""
  Proyecto  : {Fore.CYAN}{project_name}{Fore.WHITE}
  Ubicación : {Fore.CYAN}{project_dir}{Fore.WHITE}
  Normativa : {Fore.CYAN}{chosen['label']}{Fore.WHITE}
""" + Style.RESET_ALL)

    confirm = ask("¿Confirmar instalación? (s/n)", default="s")
    if confirm.lower() != "s":
        print(Fore.YELLOW + "\n  Instalación cancelada.\n")
        sys.exit(0)

    # ── Step 5: Build project ──────────────────────────────────────────────────
    section("PASO 5 — Creando proyecto")
    project_dir.mkdir(parents=True, exist_ok=True)
    create_project_structure(project_dir)

    # ── Step 6: Download from GitHub ──────────────────────────────────────────
    section("PASO 6 — Descargando archivos")

    download_normativa_config(normativa_id, project_dir)
    download_examples(normativa_id, project_dir)
    download_assembler_files()

    # ── Step 7: Save config ────────────────────────────────────────────────────
    save_project_config(project_dir, project_name, normativa_id)

    # ── Step 8: Install dependencies ───────────────────────────────────────────
    section("PASO 7 — Instalando dependencias")
    install_pip_requirements()

    # ── Done ───────────────────────────────────────────────────────────────────
    print(Fore.GREEN + """
  +--------------------------------------------------+
  |         Instalacion completada!                  |
  +--------------------------------------------------+
""" + Style.RESET_ALL)

    print(Fore.WHITE + f"""  Tu proyecto está listo en:
  {Fore.CYAN}{project_dir}{Fore.WHITE}

  Próximos pasos:

    1. Escribe tus capítulos en:
       {Fore.CYAN}markdowns/{Fore.WHITE}

    2. Inicia la sincronización automática:
       {Fore.CYAN}python assembler/watcher.py{Fore.WHITE}

    3. Compila el documento final:
       {Fore.CYAN}python assembler/assembler.py --normativa {normativa_id}{Fore.WHITE}

    4. Tu documento Word estará en:
       {Fore.CYAN}word/Tesis_Final.docx{Fore.WHITE}

  Revisa los ejemplos en {Fore.CYAN}markdowns/{Fore.WHITE} para aprender la sintaxis.
""" + Style.RESET_ALL)


if __name__ == "__main__":
    main()
