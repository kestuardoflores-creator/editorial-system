# Sistema de Gestión Editorial

Sistema local para redactar, formatear y compilar tesis e informes técnicos académicos.

**Stack: Markdown + Python + Word. Sin dependencias de internet.**

---

## Estructura del Proyecto

```
Tu_Proyecto/
├── markdowns/               ← escribe aquí (Obsidian o cualquier editor)
│   ├── 01_Introduccion.md
│   ├── 02_Marco_Teorico.md
│   ├── portada.json         ← datos de portada
│   └── referencias.bib      ← base de datos de citas
├── word/                    ← generado automáticamente
│   ├── 01_Introduccion.docx
│   └── Tesis_Final.docx
├── assets/
│   ├── images/              ← imágenes PNG, JPG
│   └── data/                ← tablas y gráficos Excel
├── config/
│   ├── apa7.json            ← normativa APA 7
│   └── ieee.json            ← normativa IEEE (contribución)
├── assembler/
│   ├── assembler.py         ← compilador del documento final
│   ├── watcher.py           ← sincronización MD ↔ DOCX
│   └── requirements.txt
└── examples/                ← ejemplos de sintaxis
```

---

## Inicio Rápido

### 1. Instalar dependencias

```powershell
pip install -r assembler/requirements.txt
```

### 2. Escribir el contenido

Escribe en `markdowns/` usando Obsidian o Notepad++.

### 3. Sincronización automática (watcher)

```powershell
python assembler/watcher.py
```

Detecta cambios en `.md` → genera `.docx` de vista previa en `word/`.
Si editas el `.docx` → convierte de vuelta a `.md` automáticamente.

### 4. Compilar el documento final

```powershell
python assembler/assembler.py --normativa apa7
```

Genera `word/Tesis_Final.docx` con:
- Portada
- Tabla de Contenido
- Capítulos con estilos normativos
- Numeración automática de figuras y tablas
- Lista de Figuras y Lista de Tablas
- Referencias en formato APA 7

---

## Sintaxis de Etiquetado

### Encabezados (implícitos)

```markdown
# Título del Capítulo           → TIT_CAP (salto a página impar)
## Marco Conceptual             → H1_APA  (negrita, centrado)
### Diseño de la Investigación  → H2_APA  (negrita, izquierda)
#### Sub-sección                → H3_APA  (negrita cursiva)
```

### Párrafo normal

```markdown
El texto normal no necesita etiqueta especial.
El sistema lo formatea automáticamente como TEXTO_APA.
```

### Figuras e imágenes

```markdown
> [!FIG_TIT src="assets/images/mi_figura.png"]
> Descripción de la figura que aparece debajo del número
```

### Gráficos desde Excel

```markdown
> [!FIG_TIT src="assets/data/resultados.xlsx" sheet="Grafico1"]
> Tendencia de ventas Q1–Q4 por región
```

### Tablas desde Excel

```markdown
> [!TABLA_TIT src="assets/data/datos.xlsx" sheet="Tabla1"]
> Comparación de resultados por grupo experimental
```

### Ecuaciones

```markdown
> [!ECUACION]
> Y = β₀ + β₁X₁ + β₂X₂ + ε
```

### Citas bibliográficas

```markdown
Como señala [@smith2020], la metodología cuantitativa permite...
Varios autores coinciden [@garcia2021; @lopez2019].
```

---

## Portada (`markdowns/portada.json`)

```json
{
  "titulo": "Título de la Tesis",
  "autor": "Apellidos, Nombre",
  "institucion": "Universidad Nacional...",
  "facultad": "Facultad de...",
  "programa": "Escuela Profesional de...",
  "ciudad": "Lima",
  "anio": "2024"
}
```

---

## Referencias (`markdowns/referencias.bib`)

```bibtex
@article{smith2020,
  author  = {Smith, John},
  title   = {Título del artículo},
  journal = {Journal Name},
  year    = {2020},
  volume  = {12},
  number  = {3},
  pages   = {45--67},
  doi     = {10.1000/example}
}

@book{garcia2021,
  author    = {Garcia, María},
  title     = {Título del libro},
  publisher = {Editorial Universitaria},
  year      = {2021}
}
```

---

## Contribuir una Nueva Normativa

1. Fork del repositorio.
2. Crear `config/mi_normativa.json` siguiendo la misma estructura de `apa7.json`.
3. Pull Request.

El sistema lo detecta automáticamente sin modificar ningún otro archivo.

---

## Librerías Python

| Librería | Uso |
|---|---|
| `python-docx` | Crear y editar archivos Word |
| `openpyxl` | Leer tablas y gráficos de Excel |
| `watchdog` | Detectar cambios en archivos |
| `bibtexparser` | Procesar archivos `.bib` |
| `mammoth` | Convertir DOCX → Markdown |
| `Pillow` | Insertar imágenes en Word |

---

## Licencia

MIT
