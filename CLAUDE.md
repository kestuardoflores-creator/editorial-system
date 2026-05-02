# CLAUDE.md — Sistema de Gestión Editorial

## What this project is

Desktop GUI app: compiles Markdown files into a formatted Word document following an academic normative (APA 7, APPA EEP 2021, etc.). Markdown is always the master source.

---

## How to run

```bash
bash run.sh          # creates .venv, installs deps, launches GUI
```

Manual (dev):
```bash
source .venv/bin/activate
python app/launcher.py
```

No tests, no linter config. Not a git repo yet.

---

## Project structure

```
EditorialSystem/
├── run.sh                        ← entry point
├── app/
│   ├── launcher.py               ← tkinter GUI (main window)
│   ├── assembler/
│   │   ├── assembler.py          ← Markdown → Word compiler
│   │   ├── watcher.py            ← bidirectional MD ↔ DOCX sync (watchdog)
│   │   ├── norm_excel.py         ← converts normative JSON ↔ Excel
│   │   └── requirements.txt      ← python-docx, openpyxl, watchdog, bibtexparser, mammoth, Pillow
│   ├── config/
│   │   ├── apa7.json             ← APA 7 normative (bundled)
│   │   └── appa_eep_2021.json    ← APPA EEP 2021 normative (bundled)
│   └── examples/
│       └── apa7/                 ← example MD files, portada.json, referencias.bib
├── projects/                     ← user projects created by the GUI (auto-created)
├── Anteproyecto_v2.md            ← original design spec
├── GUIA_CONTENIDO.md             ← user guide for writing content
└── repo_dump_v2.md               ← legacy full-code snapshot (not needed with Claude Code)
```

---

## Architecture

### launcher.py (388 lines)
tkinter GUI with two tabs: **New project** and **Existing projects**. Runs assembler and watcher in background threads. On project creation, copies the chosen normative JSON to `project/config/`, generates an editable `normativa.xlsx` and `indice.xlsx` via `norm_excel`, and copies example files.

On project creation, also generates `.obsidian/app.json` (disables Wikilinks, sets `assets/images` as attachment folder) and `.vscode/settings.json` (auto-saves pasted images to `assets/images/`).

Key path globals:
- `APP_DIR` — bundled app code (`app/`)
- `PORTABLE_ROOT` — root directory (where `run.sh` lives)
- `PROJECTS_DIR` — `portable_root/projects/`
- `USER_CONFIG` — `portable_root/app/config/` (user can edit)

### assembler.py (455 lines)
Called as `assembler.assemble(normativa)` after `_set_project_root(project_dir)`.

**Pipeline:**
1. Loads normative from `config/normativa.xlsx` (preferred) or `config/normativa.json`
2. Loads `markdowns/referencias.bib` and `markdowns/portada.json`
3. Reads all `markdowns/*.md` in alphabetical order
4. Parses each MD file → `Elem` objects (heading / paragraph / callout)
5. Applies styles from normative JSON to each element
6. Builds TOC, figure list, table list
7. Appends References section (only cited entries first, then uncited)
8. Saves to `word/Tesis_Final.docx`

**Callouts parsed:** `FIG_TIT`, `TABLA_TIT`, `ECUACION`, and any custom tag defined in the normative.

**TOC numbering:** configured via `config/indice.xlsx` → `Numeracion` sheet (auto-created at project creation).

**Normative priority:** `config/normativa.xlsx` > `config/normativa.json`

### watcher.py (137 lines)
Called as a thread from launcher. Uses watchdog `Observer`:
- `MarkdownHandler` → `.md` change → `convert_md_to_docx()` (individual chapter preview, no TOC)
- `WordHandler` → `.docx` change → `convert_docx_to_md()` via mammoth
- 2-second debounce prevents infinite loops
- Skips DOCX→MD conversion if filename contains "Final" (protects `Tesis_Final.docx`)

### norm_excel.py (225 lines)
- `json_to_excel(json_path, xlsx_path)` — creates editable Excel from normative JSON
- `excel_to_dict(xlsx_path)` — reads Excel back to dict (used by assembler)
- `create_indice_excel(project_dir)` — creates `config/indice.xlsx` with `Numeracion` sheet for TOC number formatting

---

## Normative JSON schema

Each normative is a JSON file with root fields and an `estilos` list:

```json
{
  "normativa": "APA 7",
  "version": "1.1.0",
  "inicio_capitulo": "IMPAR",
  "chars_por_pagina": 2500,
  "margenes_cm": {"top": 2.54, "bottom": 2.54, "left": 2.54, "right": 2.54},
  "estilos": [
    {
      "ID_Etiqueta": "TEXTO_APA",
      "Fuente": "Times New Roman",
      "Tamano": 12,
      "Interlineado": 2.0,
      "Negrita": false,
      "Italica": false,
      "Alineacion": "JUSTIFY",
      "Sangria_1era": 36,
      "Color_Texto": "#000000",
      "Espaciado_Antes": 0,
      "Espaciado_Despues": 0,
      "Es_Numerable": false
    }
  ]
}
```

`inicio_capitulo` values: `IMPAR`, `PAR`, `NUEVA`, `CONTINUO`

---

## Current status (2026-05-02)

### Implemented
- [x] tkinter GUI — new project, open, watcher toggle, compile
- [x] Markdown → Word assembler (headings, paragraphs, figures, tables, equations, citations)
- [x] Bidirectional watcher (MD↔DOCX live sync)
- [x] Normative editor (JSON ↔ Excel)
- [x] APA 7 normative (`apa7.json`)
- [x] APPA EEP 2021 normative (`appa_eep_2021.json`)
- [x] BibTeX references (article, book, inproceedings)
- [x] TOC, figure list, table list generation
- [x] Cover page from `portada.json`
- [x] Configurable TOC numbering via `indice.xlsx`
- [x] Obsidian + VS Code editor configs generated per project (`.obsidian/app.json`, `.vscode/settings.json`)

### Not yet done
- [ ] `ieee.json` — referenced in launcher `NORMATIVA_LABELS` but file missing
- [ ] `vancouver.json` — same
- [ ] `EditorialSystem.spec` — PyInstaller recipe
- [ ] `.github/workflows/build.yml` — CI build for Linux + Windows
- [ ] Git repository (no `.git` yet)
- [ ] Test suite

---

## Known design constraints

- `portada.json` and `referencias.bib` must be in `markdowns/`, not `config/`
- MD files compiled in alphabetical order — use numeric prefixes (`01_`, `02_`)
- `referencias.bib` must be named exactly that
- `fmt_ref_apa7()` only formats APA 7 — all normatives currently use the same reference formatter
- Equation callout (`ECUACION`) renders as styled text only, no LaTeX rendering
