# Guía para Contribuir una Nueva Normativa

Este documento explica cómo crear un archivo de configuración para una normativa nueva (IEEE, Vancouver, ICONTEC, APA 6, variante institucional, etc.) y agregarlo al sistema.

---

## 1. Estructura general del archivo JSON

Cada normativa vive en un único archivo dentro de `config/`:

```
config/
  apa7.json
  appa_eep_2021.json
  mi_normativa.json     ← el tuyo irá aquí
```

El archivo tiene dos secciones: **campos raíz** (configuración global) y **estilos** (uno por elemento tipográfico).

---

## 2. Campos raíz

```json
{
  "normativa":        "Nombre legible de la normativa",
  "version":          "1.0.0",
  "inicio_capitulo":  "IMPAR",
  "chars_por_pagina": 2500,
  "margenes_cm": {
    "top":    2.54,
    "bottom": 2.54,
    "left":   2.54,
    "right":  2.54
  },
  "estilos": [ ... ]
}
```

### `inicio_capitulo` — salto entre capítulos

| Valor | Comportamiento |
|---|---|
| `IMPAR` | Cada capítulo empieza en página impar (estándar para impresión a doble cara) |
| `PAR` | Cada capítulo empieza en página par |
| `NUEVA` | Cada capítulo empieza en cualquier página nueva |
| `CONTINUO` | Sin salto — el capítulo sigue inmediatamente después del anterior |

### `chars_por_pagina`

Estimación del número de caracteres por página. El sistema usa este valor para calcular números de página aproximados en la Tabla de Contenido. Un valor típico es:

- Interlineado doble, márgenes 2.54 cm → `2500`
- Interlineado 1.5, márgenes 2.54 cm → `3200`
- Interlineado sencillo, márgenes 2.54 cm → `4000`

### Márgenes con espejo (encuadernación)

Si tu normativa usa márgenes espejados (margen interior/exterior para impresión doble cara), usa este esquema:

```json
"margenes_cm": {
  "top":    4.00,
  "bottom": 2.50,
  "inner":  4.00,
  "outer":  2.50,
  "espejo": true
}
```

---

## 3. Estructura de un estilo

Cada objeto dentro de `"estilos"` describe cómo se ve un elemento del documento:

```json
{
  "ID_Etiqueta":      "H1_MINORM",
  "Nombre_Visible":   "Encabezado Nivel 1",
  "Fuente":           "Times New Roman",
  "Tamano":           12,
  "Interlineado":     2.0,
  "Negrita":          true,
  "Italica":          false,
  "Sangria_1era":     0,
  "Alineacion":       "LEFT",
  "Color_Texto":      "#000000",
  "Espaciado_Antes":  0,
  "Espaciado_Despues": 0,
  "Es_Numerable":     false,
  "Prefijo_Texto":    "",
  "Separador_Num":    ".",
  "Formato_Prefijo":  "CONTINUO",
  "Formato_Numero":   "ARABIC",
  "Posicion_Pagina":  "BREAK_TEXT",
  "Alineacion_Objeto":"CENTER"
}
```

---

## 4. Referencia de campos por estilo

### Tipografía

| Campo | Tipo | Descripción | Ejemplos |
|---|---|---|---|
| `ID_Etiqueta` | string | Identificador único, en mayúsculas con guion bajo. **No uses espacios ni caracteres especiales.** | `H1_IEEE`, `FIG_TIT`, `TEXTO_VAN` |
| `Nombre_Visible` | string | Nombre que ve el usuario en el sidebar | `Encabezado Nivel 1` |
| `Fuente` | string | Familia tipográfica exacta como aparece en Word | `Arial`, `Times New Roman`, `Calibri` |
| `Tamano` | number | Tamaño en puntos | `10`, `11`, `12`, `14` |
| `Interlineado` | number | Factor de interlineado | `1.0`, `1.5`, `2.0` |
| `Negrita` | boolean | | `true`, `false` |
| `Italica` | boolean | | `true`, `false` |

### Espaciado y alineación

| Campo | Tipo | Descripción | Ejemplos |
|---|---|---|---|
| `Sangria_1era` | number (pt) | Sangría de primera línea en puntos. **Positivo** = sangría normal. **Negativo** = sangría francesa (colgante). | `36` (1.27 cm), `-36` (francesa), `0` |
| `Alineacion` | string | | `LEFT`, `CENTER`, `RIGHT`, `JUSTIFY` |
| `Color_Texto` | string | Código hexadecimal | `#000000` (negro), `#404040` |
| `Espaciado_Antes` | number (pt) | Espacio antes del párrafo en puntos | `0`, `6`, `12`, `24` |
| `Espaciado_Despues` | number (pt) | Espacio después del párrafo en puntos | `0`, `6`, `12` |

> **Conversión rápida cm → pt:** 1 cm ≈ 28.35 pt. Ejemplos comunes:
> - 1.27 cm = 36 pt
> - 2.00 cm = 57 pt
> - 2.50 cm = 71 pt
> - 3.00 cm = 85 pt

### Numeración automática

Estos campos solo importan cuando `Es_Numerable` es `true`.

| Campo | Tipo | Descripción | Ejemplos |
|---|---|---|---|
| `Es_Numerable` | boolean | ¿Este elemento recibe número automático? | `true`, `false` |
| `Prefijo_Texto` | string | Palabra antes del número | `Figura`, `Tabla`, `Ecuación`, `` (vacío) |
| `Separador_Num` | string | Carácter entre capítulo y elemento (solo aplica en `CAPITULO_ELEMENTO`) | `.`, ` `, `-` |
| `Formato_Prefijo` | string | Convención de numeración — ver tabla abajo | `CONTINUO`, `CAPITULO_ELEMENTO` |
| `Formato_Numero` | string | Sistema numérico | `ARABIC`, `ROMAN_UPPER`, `ROMAN_LOWER` |

#### `Formato_Prefijo`

| Valor | Resultado | Cuándo usarlo |
|---|---|---|
| `CONTINUO` | Figura 1, Figura 2, … Figura 22 | Todo el documento es una secuencia continua |
| `CAPITULO_ELEMENTO` | Figura 3.1, Figura 3.2 | La numeración reinicia en cada capítulo |

> **Nota importante:** `Formato_Prefijo` es configurable por elemento. Si tu normativa dice que las figuras son continuas pero las ecuaciones reinician por capítulo, simplemente usa valores distintos para cada estilo. El usuario también puede cambiarlo directamente en el Excel sin tocar el JSON.

#### `Formato_Numero`

| Valor | Resultado |
|---|---|
| `ARABIC` | 1, 2, 3, 4 |
| `ROMAN_UPPER` | I, II, III, IV |
| `ROMAN_LOWER` | i, ii, iii, iv |

### Posición del objeto

| Campo | Descripción | Valores |
|---|---|---|
| `Posicion_Pagina` | Cómo el título interactúa con el objeto | `BREAK_TEXT` (separa el texto), `INLINE` (fluye con el texto) |
| `Alineacion_Objeto` | Alineación de la imagen o tabla | `CENTER`, `LEFT`, `RIGHT` |

---

## 5. Estilos obligatorios

Tu normativa **debe incluir al menos estos IDs** para que el sistema funcione correctamente. Los demás son opcionales.

| ID obligatorio | Rol en el sistema |
|---|---|
| `TIT_CAP` | Título de capítulo (Heading 1 en Markdown) |
| `TEXTO_[SUFIJO]` | Párrafo normal sin etiqueta |
| `FIG_TIT` | Título de figura |
| `TABLA_TIT` | Título de tabla |
| `TOC_TITULO` | Encabezado de la Tabla de Contenido |
| `TOC_ENTRADA` | Líneas de la Tabla de Contenido |
| `IDX_TITULO` | Encabezado de Lista de Figuras / Tablas |
| `IDX_ENTRADA` | Líneas de las listas |
| `PORTADA_TITULO` | Título en la portada |
| `PORTADA_AUTOR` | Autor en la portada |
| `PORTADA_INFO` | Información institucional en la portada |
| `REFERENCIA` | Entradas de la lista de referencias |

---

## 6. Convención para el `ID_Etiqueta`

Para evitar conflictos entre normativas, usa un sufijo que identifique la tuya:

| Patrón | Ejemplo |
|---|---|
| `H1_[SIGLA]` | `H1_IEEE`, `H1_VAN`, `H1_ICONTEC` |
| `TEXTO_[SIGLA]` | `TEXTO_IEEE`, `TEXTO_APA` |
| `FIG_TIT` | Este es universal — siempre se llama igual |
| `TABLA_TIT` | Ídem |

Los IDs `FIG_TIT`, `TABLA_TIT`, `ECUACION`, `TOC_TITULO`, `TOC_ENTRADA`, `IDX_TITULO`, `IDX_ENTRADA`, `REFERENCIA`, `PORTADA_TITULO`, `PORTADA_AUTOR`, `PORTADA_INFO` son **reservados** y el ensamblador los usa por nombre. No los cambies.

---

## 7. Ejemplo mínimo completo — IEEE

```json
{
  "normativa": "IEEE",
  "version": "1.0.0",
  "inicio_capitulo": "NUEVA",
  "chars_por_pagina": 3800,
  "margenes_cm": {
    "top": 1.90, "bottom": 2.54, "left": 1.90, "right": 1.90
  },
  "estilos": [
    {
      "ID_Etiqueta": "TIT_CAP",
      "Nombre_Visible": "Título de Sección",
      "Fuente": "Times New Roman",
      "Tamano": 10,
      "Interlineado": 1.0,
      "Negrita": true,
      "Italica": false,
      "Sangria_1era": 0,
      "Alineacion": "CENTER",
      "Color_Texto": "#000000",
      "Espaciado_Antes": 12,
      "Espaciado_Despues": 6,
      "Es_Numerable": false,
      "Prefijo_Texto": "",
      "Separador_Num": ".",
      "Formato_Prefijo": "CONTINUO",
      "Formato_Numero": "ARABIC",
      "Posicion_Pagina": "BREAK_TEXT",
      "Alineacion_Objeto": "CENTER"
    },
    {
      "ID_Etiqueta": "TEXTO_IEEE",
      "Nombre_Visible": "Párrafo Normal",
      "Fuente": "Times New Roman",
      "Tamano": 10,
      "Interlineado": 1.0,
      "Negrita": false,
      "Italica": false,
      "Sangria_1era": 14,
      "Alineacion": "JUSTIFY",
      "Color_Texto": "#000000",
      "Espaciado_Antes": 0,
      "Espaciado_Despues": 0,
      "Es_Numerable": false,
      "Prefijo_Texto": "",
      "Separador_Num": ".",
      "Formato_Prefijo": "CONTINUO",
      "Formato_Numero": "ARABIC",
      "Posicion_Pagina": "BREAK_TEXT",
      "Alineacion_Objeto": "CENTER"
    },
    {
      "ID_Etiqueta": "FIG_TIT",
      "Nombre_Visible": "Título de Figura",
      "Fuente": "Times New Roman",
      "Tamano": 8,
      "Interlineado": 1.0,
      "Negrita": false,
      "Italica": false,
      "Sangria_1era": 0,
      "Alineacion": "CENTER",
      "Color_Texto": "#000000",
      "Espaciado_Antes": 6,
      "Espaciado_Despues": 6,
      "Es_Numerable": true,
      "Prefijo_Texto": "Fig.",
      "Separador_Num": ".",
      "Formato_Prefijo": "CONTINUO",
      "Formato_Numero": "ARABIC",
      "Posicion_Pagina": "INLINE",
      "Alineacion_Objeto": "CENTER"
    },
    {
      "ID_Etiqueta": "TABLA_TIT",
      "Nombre_Visible": "Título de Tabla",
      "Fuente": "Times New Roman",
      "Tamano": 8,
      "Interlineado": 1.0,
      "Negrita": true,
      "Italica": false,
      "Sangria_1era": 0,
      "Alineacion": "CENTER",
      "Color_Texto": "#000000",
      "Espaciado_Antes": 6,
      "Espaciado_Despues": 6,
      "Es_Numerable": true,
      "Prefijo_Texto": "TABLE",
      "Separador_Num": " ",
      "Formato_Prefijo": "CONTINUO",
      "Formato_Numero": "ROMAN_UPPER",
      "Posicion_Pagina": "INLINE",
      "Alineacion_Objeto": "CENTER"
    }
  ]
}
```

---

## 8. Cómo contribuir al repositorio

1. **Fork** del repositorio en GitHub.
2. Crea tu archivo en `config/mi_normativa.json`.
3. Crea la carpeta `examples/mi_normativa/` con al menos:
   - `ejemplo_citas.md`
   - `ejemplo_figuras.md`
   - `portada.json`
   - `referencias.bib`
4. Verifica que el archivo JSON sea válido:
   ```powershell
   python -c "import json; json.load(open('config/mi_normativa.json'))"
   ```
5. Verifica que el Excel se genere correctamente:
   ```powershell
   python assembler/norm_excel.py mi_normativa
   ```
6. Abre el Pull Request con una descripción que incluya:
   - Nombre completo de la normativa
   - Institución o estándar de origen
   - País o región
   - Fuente del documento oficial (si está disponible públicamente)

---

## 9. Herramienta de validación

Puedes convertir tu JSON a Excel en cualquier momento para verificar visualmente que los valores son correctos:

```powershell
python assembler/norm_excel.py mi_normativa
```

Esto genera `config/mi_normativa.xlsx` que puedes abrir en Excel y revisar columna por columna antes de hacer el Pull Request.

Para agregar este comando al script `norm_excel.py`, el punto de entrada CLI ya está incluido al final del archivo.
