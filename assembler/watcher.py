"""
watcher.py — Sistema de Gestión Editorial
Sincronización bidireccional MD ↔ DOCX en segundo plano.

Uso:
    python assembler/watcher.py
    python assembler/watcher.py --normativa ieee

Ctrl+C para detener.
"""

import argparse
import json
import os
import sys
import time
from pathlib import Path
from datetime import datetime

try:
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler
    from docx import Document
    from docx.shared import Pt
    import mammoth
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install",
                           "watchdog", "python-docx", "mammoth"])
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler
    from docx import Document
    from docx.shared import Pt
    import mammoth

# ── Paths ──────────────────────────────────────────────────────────────────────
ROOT          = Path(__file__).parent.parent
MARKDOWNS_DIR = ROOT / "markdowns"
WORD_DIR      = ROOT / "word"
CONFIG_DIR    = ROOT / "config"

WORD_DIR.mkdir(exist_ok=True)

# Tracks recently modified files to prevent echo loops
_recently_changed = {}
DEBOUNCE_SECONDS  = 2.0


def is_debounced(path):
    """Return True if the file was modified by us recently (loop prevention)."""
    key = str(path)
    last = _recently_changed.get(key, 0)
    return (time.time() - last) < DEBOUNCE_SECONDS


def mark_changed(path):
    _recently_changed[str(path)] = time.time()


# ─────────────────────────────────────────────────────────────────────────────
# MD → DOCX (single file conversion using the assembler engine)
# ─────────────────────────────────────────────────────────────────────────────

def convert_md_to_docx(md_path, normativa_name):
    """Convert a single .md file to a preview .docx in the word/ folder."""
    try:
        # Import the assembler's parser and style engine
        sys.path.insert(0, str(Path(__file__).parent))
        from assembler import (parse_markdown, add_styled_paragraph,
                               apply_style, apply_run_style, ALIGN_MAP)
        import assembler as asm

        norm_path = CONFIG_DIR / f"{normativa_name}.json"
        with open(norm_path, encoding="utf-8") as f:
            norm = json.load(f)
        style_map = {s["ID_Etiqueta"]: s for s in norm["estilos"]}
        margenes  = norm.get("margenes_cm", {"top": 2.54, "bottom": 2.54,
                                              "left": 2.54, "right": 2.54})

        from docx.shared import Cm
        doc     = Document()
        section = doc.sections[0]
        asm.set_margins(section, margenes)

        # Clear default paragraph
        for p in doc.paragraphs:
            p._element.getparent().remove(p._element)

        elements = parse_markdown(md_path)
        numbering = asm.NumberingEngine()

        for elem in elements:
            if elem.kind == "heading":
                level = elem.kwargs["level"]
                text  = elem.kwargs["text"]
                label = asm.HEADING_MAP.get(level, "TEXTO_APA")
                s     = style_map.get(label, style_map.get("TEXTO_APA", {}))
                if level == 1:
                    numbering.next_chapter()
                add_styled_paragraph(doc, text, s)

            elif elem.kind == "paragraph":
                s = style_map.get("TEXTO_APA", {})
                add_styled_paragraph(doc, elem.kwargs["text"], s)

            elif elem.kind == "callout":
                tag     = elem.kwargs["tag"]
                caption = elem.kwargs["caption"]
                s       = style_map.get(tag, style_map.get("TEXTO_APA", {}))
                prefix  = numbering.build_prefix(s) if s.get("Es_Numerable") else ""
                text    = f"{prefix} {caption}".strip() if prefix else caption
                add_styled_paragraph(doc, text, s)

        out_name = md_path.stem + ".docx"
        out_path = WORD_DIR / out_name
        mark_changed(out_path)
        doc.save(str(out_path))
        _log(f"MD→DOCX: {md_path.name} → {out_path.name}")

    except Exception as e:
        _log(f"❌ Error convirtiendo {md_path.name}: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# DOCX → MD (using mammoth)
# ─────────────────────────────────────────────────────────────────────────────

def convert_docx_to_md(docx_path):
    """Convert a .docx back to .md (Markdown is always the master source)."""
    # Skip the final compiled document
    if "Final" in docx_path.name or "Tesis_Final" in docx_path.name:
        return

    try:
        with open(docx_path, "rb") as f:
            result = mammoth.convert_to_markdown(f)

        md_text = result.value
        out_name = docx_path.stem + ".md"
        out_path = MARKDOWNS_DIR / out_name
        mark_changed(out_path)

        out_path.write_text(md_text, encoding="utf-8")
        _log(f"DOCX→MD: {docx_path.name} → {out_path.name}")

        if result.messages:
            for msg in result.messages:
                _log(f"   ⚠️  {msg}")

    except Exception as e:
        _log(f"❌ Error convirtiendo {docx_path.name}: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# EVENT HANDLERS
# ─────────────────────────────────────────────────────────────────────────────

class MarkdownHandler(FileSystemEventHandler):
    def __init__(self, normativa):
        self.normativa = normativa

    def on_modified(self, event):
        if event.is_directory:
            return
        p = Path(event.src_path)
        if p.suffix.lower() == ".md" and not is_debounced(p):
            _log(f"📝 Cambio detectado: {p.name}")
            time.sleep(0.3)  # small delay to ensure file write is complete
            convert_md_to_docx(p, self.normativa)

    def on_created(self, event):
        self.on_modified(event)


class WordHandler(FileSystemEventHandler):
    def on_modified(self, event):
        if event.is_directory:
            return
        p = Path(event.src_path)
        if p.suffix.lower() == ".docx" and not is_debounced(p):
            _log(f"📄 Cambio en Word detectado: {p.name} → convirtiendo a Markdown")
            time.sleep(0.5)
            convert_docx_to_md(p)

    def on_created(self, event):
        self.on_modified(event)


# ─────────────────────────────────────────────────────────────────────────────
# LOGGER
# ─────────────────────────────────────────────────────────────────────────────

def _log(msg):
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Watcher bidireccional MD ↔ DOCX")
    parser.add_argument("--normativa", default="apa7",
                        help="Normativa para convertir MD→DOCX. Ej: apa7, ieee")
    args = parser.parse_args()

    print("\n" + "="*60)
    print("  Sistema de Gestión Editorial — Watcher")
    print("="*60)
    print(f"\n  Normativa   : {args.normativa}")
    print(f"  Escuchando  : {MARKDOWNS_DIR}")
    print(f"  Escuchando  : {WORD_DIR}")
    print("\n  Ctrl+C para detener\n")

    # Initial conversion of all .md files
    _log("Conversión inicial de todos los archivos .md...")
    for md_file in sorted(MARKDOWNS_DIR.glob("*.md")):
        convert_md_to_docx(md_file, args.normativa)

    # Setup observers
    observer = Observer()
    observer.schedule(MarkdownHandler(args.normativa), str(MARKDOWNS_DIR), recursive=False)
    observer.schedule(WordHandler(), str(WORD_DIR), recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        print("\n\n  Watcher detenido.")

    observer.join()


if __name__ == "__main__":
    main()
