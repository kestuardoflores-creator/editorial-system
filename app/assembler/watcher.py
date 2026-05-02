import argparse, json, sys, time
from datetime import datetime
from pathlib import Path

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from docx import Document
import mammoth

ROOT = MARKDOWNS_DIR = WORD_DIR = CONFIG_DIR = Path(__file__).parent.parent
sys.path.insert(0, str(Path(__file__).parent))


def _set_project_root(project_dir):
    global ROOT, MARKDOWNS_DIR, WORD_DIR, CONFIG_DIR
    ROOT = Path(project_dir)
    MARKDOWNS_DIR, WORD_DIR, CONFIG_DIR = ROOT / "markdowns", ROOT / "word", ROOT / "config"
    WORD_DIR.mkdir(exist_ok=True)


_recently = {}
DEBOUNCE  = 2.0


def _debounced(path):
    return (time.time() - _recently.get(str(path), 0)) < DEBOUNCE

def _mark(path):
    _recently[str(path)] = time.time()

def _log(msg):
    print(f"[{datetime.now():%H:%M:%S}] {msg}")


def convert_md_to_docx(md_path, normativa):
    try:
        from assembler import (parse_markdown, add_styled_paragraph,
                               HEADING_MAP, NumberingEngine, set_margins)
        norm = json.loads((CONFIG_DIR / f"{normativa}.json").read_text(encoding="utf-8"))
        sm   = {s["ID_Etiqueta"]: s for s in norm["estilos"]}

        doc = Document()
        set_margins(doc.sections[0], norm.get("margenes_cm", {}))
        for p in doc.paragraphs: p._element.getparent().remove(p._element)

        num = NumberingEngine()
        for elem in parse_markdown(md_path):
            if elem.kind == "heading":
                lv, text = elem.kw["level"], elem.kw["text"]
                s = sm.get(HEADING_MAP.get(lv, "TEXTO_APA"), sm.get("TEXTO_APA", {}))
                if lv == 1: num.next_chapter()
                add_styled_paragraph(doc, text, s)
            elif elem.kind == "paragraph":
                add_styled_paragraph(doc, elem.kw["text"], sm.get("TEXTO_APA", {}))
            elif elem.kind == "callout":
                s   = sm.get(elem.kw["tag"], sm.get("TEXTO_APA", {}))
                pre = num.build_prefix(s) if s.get("Es_Numerable") else ""
                add_styled_paragraph(doc, f"{pre} {elem.kw['caption']}".strip(), s)

        out = WORD_DIR / (md_path.stem + ".docx")
        _mark(out); doc.save(str(out))
        _log(f"MD→DOCX: {md_path.name} → {out.name}")
    except Exception as e:
        _log(f"❌ {md_path.name}: {e}")


def convert_docx_to_md(docx_path):
    if "Final" in docx_path.name: return
    try:
        with open(docx_path, "rb") as f:
            result = mammoth.convert_to_markdown(f)
        out = MARKDOWNS_DIR / (docx_path.stem + ".md")
        _mark(out); out.write_text(result.value, encoding="utf-8")
        _log(f"DOCX→MD: {docx_path.name} → {out.name}")
        for msg in result.messages: _log(f"  ⚠️ {msg}")
    except Exception as e:
        _log(f"❌ {docx_path.name}: {e}")


class MarkdownHandler(FileSystemEventHandler):
    def __init__(self, normativa): self.normativa = normativa

    def on_modified(self, event):
        if event.is_directory: return
        p = Path(event.src_path)
        if p.suffix.lower() == ".md" and not _debounced(p):
            _log(f"📝 Cambio: {p.name}"); time.sleep(0.3)
            convert_md_to_docx(p, self.normativa)

    on_created = on_modified


class WordHandler(FileSystemEventHandler):
    def on_modified(self, event):
        if event.is_directory: return
        p = Path(event.src_path)
        if p.suffix.lower() == ".docx" and not _debounced(p):
            _log(f"📄 Cambio: {p.name}"); time.sleep(0.5)
            convert_docx_to_md(p)

    on_created = on_modified


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--normativa", default=None)
    ap.add_argument("--project", default=None)
    args = ap.parse_args()
    if args.project: _set_project_root(Path(args.project).resolve())

    normativa = args.normativa
    if not normativa:
        pj = CONFIG_DIR / "project.json"
        if pj.exists():
            try: normativa = json.loads(pj.read_text(encoding="utf-8")).get("normativa")
            except Exception: pass
    normativa = normativa or "apa7"

    print(f"\nWatcher | normativa: {normativa} | proyecto: {ROOT}")
    print("Ctrl+C para detener\n")

    for md in sorted(MARKDOWNS_DIR.glob("*.md")):
        convert_md_to_docx(md, normativa)

    obs = Observer()
    obs.schedule(MarkdownHandler(normativa), str(MARKDOWNS_DIR), recursive=False)
    obs.schedule(WordHandler(), str(WORD_DIR), recursive=False)
    obs.start()
    try:
        while True: time.sleep(1)
    except KeyboardInterrupt:
        obs.stop(); print("\nWatcher detenido.")
    obs.join()


if __name__ == "__main__":
    main()
