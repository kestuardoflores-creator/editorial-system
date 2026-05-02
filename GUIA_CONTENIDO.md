# Guía de contenido — Sistema de Gestión Editorial

Este documento explica qué archivos debe crear para que el ensamblador produzca un documento Word correcto.

---

## Estructura de un proyecto

```
projects/NombreProyecto/
├── markdowns/          ← tus archivos de contenido van aquí
│   ├── 01_portada.json         ← datos de portada (no es .md)
│   ├── 02_referencias.bib      ← referencias BibTeX (no es .md)
│   ├── 03_introduccion.md
│   ├── 04_marco_teorico.md
│   ├── 05_metodologia.md
│   └── 06_resultados.md
├── assets/
│   ├── images/         ← imágenes PNG / JPG para figuras
│   └── data/           ← archivos Excel para tablas
├── config/             ← generado automáticamente al crear el proyecto
│   ├── normativa.json  ← copia de la normativa elegida (fuente de respaldo)
│   ├── normativa.xlsx  ← versión editable en Excel (tiene prioridad sobre el JSON)
│   └── indice.xlsx     ← controla la numeración de la tabla de contenido
├── .obsidian/          ← config Obsidian (generado automáticamente)
│   └── app.json        ← desactiva Wikilinks, usa assets/images como carpeta de adjuntos
├── .vscode/            ← config VS Code (generado automáticamente)
│   └── settings.json   ← pegar imagen la guarda en assets/images/
└── word/               ← aquí aparece el .docx final
```

> Ver `editor_setup.md` para instrucciones de configuración de cada editor.

> Los `.md` se ensamblan en **orden alfabético** por nombre de archivo.
> Usa prefijos numéricos (`01_`, `02_`) para controlar el orden.

---

## portada.json

Debe estar en `markdowns/portada.json` con estos campos:

```json
{
  "titulo":      "Título completo de la tesis",
  "autor":       "Apellido Paterno Apellido Materno, Nombre",
  "institucion": "Universidad Nacional",
  "facultad":    "Facultad de Ingeniería",
  "programa":    "Escuela de Estudios de Postgrado",
  "ciudad":      "Ciudad de Guatemala",
  "anio":        "2024"
}
```

---

## referencias.bib

El nombre del archivo debe ser exactamente `referencias.bib`. El ensamblador soporta `article`, `book`, `inproceedings`.

```bibtex
@article{clave2024,
  author  = {Apellido, Nombre and Otro, Autor},
  title   = {Título del artículo},
  journal = {Nombre de la Revista},
  year    = {2024},
  volume  = {10},
  number  = {2},
  pages   = {100--120},
  doi     = {10.xxxx/xxxxx}
}

@book{libro2023,
  author    = {Autor, Nombre},
  title     = {Título del libro},
  publisher = {Editorial},
  year      = {2023}
}

@inproceedings{congreso2022,
  author    = {Autor, Nombre},
  title     = {Título de la ponencia},
  booktitle = {Nombre del Congreso},
  year      = {2022},
  pages     = {50--60}
}
```

---

## Archivos Markdown — sintaxis soportada

### Encabezados → niveles de la normativa

| Markdown | Estilo aplicado | Uso |
|---|---|---|
| `# Título` | TIT_CAP | Título de capítulo — genera salto de página |
| `## Subtítulo` | H1_APA | Sección nivel 1 |
| `### Subtítulo` | H2_APA | Sección nivel 2 |
| `#### Subtítulo` | H3_APA | Sección nivel 3 |
| `##### Subtítulo` | H4_APA | Sección nivel 4 |
| `###### Subtítulo` | H5_APA | Sección nivel 5 |

> Cada archivo `.md` debe empezar con un solo `# Título`. No saltes niveles.

### Párrafos normales

Escribe texto libre. Las líneas continuas se unen en un solo párrafo.
Una línea en blanco separa párrafos.

```markdown
# Introducción

Este es el primer párrafo del capítulo. Puede ser tan largo
como necesites, las líneas se unen automáticamente.

Este es el segundo párrafo, separado por una línea en blanco.
```

### Citas en el texto — `[@clave]`

```markdown
Como señala [@smith2020], el método cuantitativo es el más usado.

Varios autores coinciden en esto [@garcia2021; @lopez2019].
```

El ensamblador convierte `[@smith2020]` → `(Smith, 2020)` automáticamente.
La clave debe existir en `referencias.bib`.

### Figuras — callout `[!FIG_TIT]`

```markdown
> [!FIG_TIT src="assets/images/mi_grafico.png"]
> Descripción breve de la figura
```

- `src` debe ser la ruta **relativa a la carpeta del proyecto**.
- La línea `>` siguiente es el pie de figura.
- El número se asigna automáticamente (`Figura 1.1`, `Figura 1.2`, etc.).

### Tablas desde Excel — callout `[!TABLA_TIT]`

```markdown
> [!TABLA_TIT src="assets/data/datos.xlsx" sheet="Hoja1"]
> Descripción de la tabla
```

- `src` ruta relativa al proyecto.
- `sheet` nombre exacto de la hoja en Excel.
- El número se asigna automáticamente.

### Ecuaciones — callout `[!ECUACION]`

```markdown
> [!ECUACION]
> Y = β₀ + β₁X₁ + β₂X₂ + ε
```

La ecuación se renderiza como texto con el estilo `ECUACION` de la normativa.
No hay renderizado LaTeX; escribe la ecuación directamente en texto plano o con caracteres Unicode.

---

## Tabla de Contenido y numeración de secciones

La numeración en la tabla de contenido se configura en `config/indice.xlsx`, hoja `Numeracion`.

| Nivel | Estilo | Separador | Prefijo |
|---|---|---|---|
| 1 | ARABIC | . | Capítulo |
| 2 | ARABIC | . | |
| 3 | ARABIC | . | |

**Estilos disponibles:** `ARABIC`, `ROMAN_UPPER`, `ROMAN_LOWER`, `LETRA_UPPER`, `LETRA_LOWER`, `NINGUNA`

Ejemplo de resultado con prefijo `Capítulo` en nivel 1:
```
Capítulo 1.  Introducción
    1.1  Antecedentes
    1.1.1  Contexto nacional
```

---

## Editar la normativa

Al crear un proyecto se genera `config/normativa.xlsx` — un Excel editable con todos los estilos. El ensamblador lo usa en lugar del JSON si existe.

Para modificar un estilo: abre el Excel, edita la fila correspondiente y guarda. No es necesario tocar el JSON.

---

## Ejemplo de capítulo completo

```markdown
# Marco Teórico

Este capítulo presenta los fundamentos conceptuales de la investigación.

## Antecedentes

Diversos estudios han abordado este tema [@smith2020; @garcia2021].
El enfoque cuantitativo ha demostrado ser el más efectivo [@lopez2019].

## Variables del Estudio

Las variables se operacionalizan en la siguiente tabla.

> [!TABLA_TIT src="assets/data/variables.xlsx" sheet="Variables"]
> Operacionalización de variables del estudio

## Modelo Matemático

El modelo de regresión utilizado es:

> [!ECUACION]
> Ŷ = β₀ + β₁X₁ + β₂X₂

Donde Ŷ es el valor estimado de la variable dependiente.

## Diagrama del Marco Conceptual

> [!FIG_TIT src="assets/images/marco_conceptual.png"]
> Diagrama del marco conceptual de la investigación
```

---

## Reglas importantes

1. **Un `# Título` por archivo** — cada archivo `.md` debe empezar con un solo `#`.
2. **Nunca saltes niveles** — no pongas `###` si no hay `##` antes.
3. **Las claves de citas deben coincidir** con las de `referencias.bib`.
4. **Las imágenes deben existir** en `assets/images/` antes de compilar.
5. **El `portada.json` va en `markdowns/`**, no en `config/`.
6. **El archivo BibTeX se llama exactamente `referencias.bib`**.
7. **Orden de archivos = orden alfabético** — usa prefijos numéricos `01_`, `02_`.
8. **`config/normativa.xlsx` tiene prioridad** sobre `config/normativa.json` — edita el Excel, no el JSON.
