// ── format-engine.js ──────────────────────────────────────────────────────────
// Applies visual formatting to a Google Docs Paragraph based on a style object
// read from the Configuracion_Estilos sheet.
//
// Exposed function:  applyFormat(paragraph, style)
// Called by:         sidebar.js → applyLabel() and syncAllStyles()
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Applies all formatting properties from a style row to a Paragraph element.
 *
 * @param {GoogleAppsScript.Document.Paragraph} paragraph
 * @param {Object} style  — row from Configuracion_Estilos as a key-value object
 */
function applyFormat(paragraph, style) {

  // ── Text-level attributes ─────────────────────────────────────────────────
  const font   = String(style['Fuente']      || 'Times New Roman');
  const size   = parseFloat(style['Tamano'])  || 12;
  const bold   = _parseBool(style['Negrita']);
  const italic = _parseBool(style['Italica']);
  const color  = String(style['Color_Texto'] || '#000000');

  const text = paragraph.editAsText();
  text.setFontFamily(font);
  text.setFontSize(size);
  text.setBold(bold);
  text.setItalic(italic);
  text.setForegroundColor(color);

  // ── Paragraph-level attributes ────────────────────────────────────────────
  const align      = _parseAlignment(style['Alineacion']);
  const spacing    = parseFloat(style['Interlineado'])     || 2.0;
  const indent     = parseFloat(style['Sangria_1era'])     || 0;
  const spaceBefore = parseFloat(style['Espaciado_Antes']) || 0;
  const spaceAfter  = parseFloat(style['Espaciado_Despues']) || 0;

  paragraph.setAlignment(align);
  paragraph.setLineSpacing(spacing);
  paragraph.setIndentFirstLine(indent);
  paragraph.setSpacingBefore(spaceBefore);
  paragraph.setSpacingAfter(spaceAfter);
}

// ── Private helpers ───────────────────────────────────────────────────────────

/**
 * Maps alignment string from the sheet to the Apps Script enum.
 * Defaults to LEFT for any unknown value.
 */
function _parseAlignment(alignStr) {
  const map = {
    'CENTER':  DocumentApp.HorizontalAlignment.CENTER,
    'LEFT':    DocumentApp.HorizontalAlignment.LEFT,
    'RIGHT':   DocumentApp.HorizontalAlignment.RIGHT,
    'JUSTIFY': DocumentApp.HorizontalAlignment.JUSTIFY,
  };
  return map[String(alignStr || '').toUpperCase()]
    || DocumentApp.HorizontalAlignment.LEFT;
}

/**
 * Parses boolean values that may come from the Sheet as TRUE/FALSE strings,
 * actual booleans, or 1/0.
 */
function _parseBool(value) {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number')  return value !== 0;
  return String(value).trim().toUpperCase() === 'TRUE';
}
