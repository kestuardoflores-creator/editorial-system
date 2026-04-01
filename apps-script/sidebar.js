// ── sidebar.js ────────────────────────────────────────────────────────────────
// Server-side entry point for the Editorial sidebar.
// Handles menu creation, sidebar launch, label application, and sync.
// ─────────────────────────────────────────────────────────────────────────────

const PROP_SHEET_ID  = 'EDITORIAL_SHEET_ID';
const SHEET_ESTILOS  = 'Configuracion_Estilos';
const SHEET_REGISTRO = 'Registro_Elementos';
const NR_PREFIX      = 'editorial_';   // NamedRange name prefix

// ── Menu ──────────────────────────────────────────────────────────────────────

function onOpen() {
  DocumentApp.getUi()
    .createMenu('📚 Editorial')
    .addItem('Abrir panel de etiquetado', 'showSidebar')
    .addSeparator()
    .addItem('Sincronizar estilos', 'syncAllStyles')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService
    .createHtmlOutputFromFile('sidebar')
    .setTitle('Editorial — Etiquetado')
    .setWidth(300);
  DocumentApp.getUi().showSidebar(html);
}

// ── Sheet connection ──────────────────────────────────────────────────────────

// Called once by the Colab installer after creating the sheet.
function setSheetId(sheetId) {
  PropertiesService.getScriptProperties()
    .setProperty(PROP_SHEET_ID, sheetId);
}

function _getSheetId() {
  const id = PropertiesService.getScriptProperties()
    .getProperty(PROP_SHEET_ID);
  if (!id) throw new Error('Sheet ID no configurado. Ejecuta el installer primero.');
  return id;
}

// ── Styles (called by sidebar on load) ───────────────────────────────────────

// Returns the list of styles for the sidebar dropdown.
function getStyles() {
  const ws = SpreadsheetApp
    .openById(_getSheetId())
    .getSheetByName(SHEET_ESTILOS);

  const [header, ...rows] = ws.getDataRange().getValues();

  return rows
    .filter(row => row[0])          // skip empty rows
    .map(row => {
      const obj = {};
      header.forEach((col, i) => { obj[col] = row[i]; });
      return obj;
    });
}

// Returns the label ID of the currently selected paragraph, or null.
function getCurrentLabel() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  if (!selection) return null;

  const elements = selection.getRangeElements();
  if (!elements.length) return null;

  const targetEl = _resolveParaParagraph(elements[0].getElement());
  if (!targetEl) return null;

  for (const nr of doc.getNamedRanges()) {
    if (!nr.getName().startsWith(NR_PREFIX)) continue;
    for (const re of nr.getRange().getRangeElements()) {
      if (_resolveParaParagraph(re.getElement()) === targetEl) {
        // Name format: editorial_LABELID_UUID
        return nr.getName().split('_')[1];
      }
    }
  }
  return null;
}

// ── Apply label (main action) ─────────────────────────────────────────────────

function applyLabel(labelId) {
  const doc       = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();

  if (!selection) {
    return { ok: false, msg: 'No hay texto seleccionado.' };
  }

  const elements = selection.getRangeElements();
  if (!elements.length) {
    return { ok: false, msg: 'Selección vacía.' };
  }

  // Resolve to paragraph element
  const paragraph = _resolveParaParagraph(elements[0].getElement());
  if (!paragraph) {
    return { ok: false, msg: 'Selecciona dentro de un párrafo válido.' };
  }

  // Load style
  const styles = getStyles();
  const style  = styles.find(s => s['ID_Etiqueta'] === labelId);
  if (!style) {
    return { ok: false, msg: `Etiqueta "${labelId}" no encontrada en el Sheet.` };
  }

  // AS-05: If paragraph already has a label, remove old NamedRange + mark HUERFANO
  _removeExistingLabel(doc, paragraph);

  // AS-02: Apply formatting
  applyFormat(paragraph.asParagraph(), style);

  // AS-03: Apply Color_Marcado background
  paragraph.asParagraph().setBackgroundColor(style['Color_Marcado'] || null);

  // AS-04: Create new NamedRange
  const rangeId = NR_PREFIX + labelId + '_' +
    Utilities.getUuid().replace(/-/g, '').substring(0, 8);
  const range = doc.newRange().addElement(paragraph).build();
  doc.addNamedRange(rangeId, range);

  // AS-04: Register in Registro_Elementos
  const refText = paragraph.asParagraph().getText().substring(0, 80);
  _upsertRegistro(labelId, rangeId, refText, 'OK');

  return { ok: true, msg: `✅ "${style['Nombre_Visible']}" aplicado.` };
}

// ── Sync: re-apply styles to every labeled paragraph ─────────────────────────

function syncAllStyles() {
  const doc     = DocumentApp.getActiveDocument();
  const styles  = getStyles();
  const styleMap = {};
  styles.forEach(s => { styleMap[s['ID_Etiqueta']] = s; });

  let count = 0;
  for (const nr of doc.getNamedRanges()) {
    if (!nr.getName().startsWith(NR_PREFIX)) continue;

    // Name format: editorial_LABELID_UUID  →  split by '_', index 1
    const labelId = nr.getName().split('_')[1];
    const style   = styleMap[labelId];
    if (!style) continue;

    for (const re of nr.getRange().getRangeElements()) {
      const para = _resolveParaParagraph(re.getElement());
      if (!para) continue;
      applyFormat(para.asParagraph(), style);
      para.asParagraph().setBackgroundColor(style['Color_Marcado'] || null);
      count++;
    }
  }

  DocumentApp.getUi().alert(`✅ ${count} párrafo(s) sincronizados con los estilos actuales.`);
}

// ── Internal helpers ──────────────────────────────────────────────────────────

// Resolve any element to its containing Paragraph, or null.
function _resolveParaParagraph(el) {
  if (!el) return null;
  if (el.getType() === DocumentApp.ElementType.PARAGRAPH) return el;
  const parent = el.getParent ? el.getParent() : null;
  if (parent && parent.getType() === DocumentApp.ElementType.PARAGRAPH) return parent;
  return null;
}

// Remove any existing editorial NamedRange on the given paragraph.
// Marks the old row as HUERFANO in Registro_Elementos.
function _removeExistingLabel(doc, paragraph) {
  for (const nr of doc.getNamedRanges()) {
    if (!nr.getName().startsWith(NR_PREFIX)) continue;
    for (const re of nr.getRange().getRangeElements()) {
      if (_resolveParaParagraph(re.getElement()) === paragraph) {
        _markRegistroEstado(nr.getName(), 'HUERFANO');
        nr.remove();
        return;
      }
    }
  }
}

// Insert or update a row in Registro_Elementos.
function _upsertRegistro(labelId, rangeId, refText, estado) {
  const ws    = SpreadsheetApp.openById(_getSheetId())
    .getSheetByName(SHEET_REGISTRO);
  const docId = DocumentApp.getActiveDocument().getId();
  const data  = ws.getDataRange().getValues();

  // Check if rangeId row already exists (update in place)
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === rangeId) {
      ws.getRange(i + 1, 1, 1, 7)
        .setValues([[docId, labelId, rangeId, refText, '', '', estado]]);
      return;
    }
  }

  // New row
  ws.appendRow([docId, labelId, rangeId, refText, '', '', estado]);
}

// Update only the Estado column for a given rangeId.
function _markRegistroEstado(rangeId, estado) {
  const ws   = SpreadsheetApp.openById(_getSheetId())
    .getSheetByName(SHEET_REGISTRO);
  const data = ws.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === rangeId) {
      ws.getRange(i + 1, 7).setValue(estado);
      return;
    }
  }
}
