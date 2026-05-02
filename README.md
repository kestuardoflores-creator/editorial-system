# Sistema de Gestión Editorial

Compila archivos Markdown en documentos Word con formato académico profesional (APA 7, APPA EEP 2021, IEEE, Vancouver).

---

## Instalación

### Linux

```bash
git clone https://github.com/kestuardoflores-creator/editorial-system
cd editorial-system
bash run.sh
```

### Windows

**Opción A — Ejecutable (recomendado, no requiere Python)**
1. Descargar `EditorialSystem.zip` desde [Releases](https://github.com/kestuardoflores-creator/editorial-system/releases)
2. Descomprimir y ejecutar `EditorialSystem.exe`

**Opción B — Desde código fuente (requiere Python 3.11+)**
1. Descargar o clonar el repositorio
2. Doble clic en `run.bat`

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
├── .vscode/            ← config VS Code (generado automáticamente)
└── word/               ← aquí aparece el .docx final
```

> Los `.md` se ensamblan en **orden alfabético**. Usa prefijos numéricos (`01_`, `02_`) para controlar el orden.

---

## portada.json

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

El archivo debe llamarse exactamente `referencias.bib`. Soporta `article`, `book`, `inproceedings`.

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
```

---

## Sintaxis Markdown

### Encabezados

| Markdown | Estilo | Uso |
|---|---|---|
| `# Título` | TIT_CAP | Título de capítulo — genera salto de página |
| `## Subtítulo` | H1_APA | Sección nivel 1 |
| `### Subtítulo` | H2_APA | Sección nivel 2 |
| `#### Subtítulo` | H3_APA | Sección nivel 3 |
| `##### Subtítulo` | H4_APA | Sección nivel 4 |
| `###### Subtítulo` | H5_APA | Sección nivel 5 |

> Cada archivo `.md` debe empezar con un solo `# Título`. No saltes niveles.

### Citas

```markdown
Como señala [@smith2020], el método es el más usado.
Varios autores coinciden [@garcia2021; @lopez2019].
```

### Figuras

```markdown
> [!FIG_TIT src="assets/images/mi_grafico.png"]
> Descripción breve de la figura
```

### Tablas desde Excel

```markdown
> [!TABLA_TIT src="assets/data/datos.xlsx" sheet="Hoja1"]
> Descripción de la tabla
```

### Ecuaciones

```markdown
> [!ECUACION]
> Y = β₀ + β₁X₁ + β₂X₂ + ε
```

---

## Tabla de Contenido y numeración

Configurada en `config/indice.xlsx`, hoja `Numeracion`.

| Nivel | Estilo | Separador | Prefijo |
|---|---|---|---|
| 1 | ARABIC | . | Capítulo |
| 2 | ARABIC | . | |

**Estilos disponibles:** `ARABIC`, `ROMAN_UPPER`, `ROMAN_LOWER`, `LETRA_UPPER`, `LETRA_LOWER`, `NINGUNA`

---

## Editar la normativa

Al crear un proyecto se genera `config/normativa.xlsx`. Edita ese archivo para cambiar fuentes, tamaños y espaciados. El ensamblador lo usa en lugar del JSON si existe.

---

## Ejemplo de capítulo

```markdown
# Marco Teórico

Este capítulo presenta los fundamentos conceptuales.

## Antecedentes

Diversos estudios han abordado este tema [@smith2020; @garcia2021].

## Variables del Estudio

> [!TABLA_TIT src="assets/data/variables.xlsx" sheet="Variables"]
> Operacionalización de variables del estudio

## Modelo Matemático

> [!ECUACION]
> Ŷ = β₀ + β₁X₁ + β₂X₂

## Diagrama

> [!FIG_TIT src="assets/images/marco_conceptual.png"]
> Diagrama del marco conceptual
```

---

## Reglas importantes

1. **Un `# Título` por archivo** — cada `.md` empieza con un solo `#`.
2. **Nunca saltes niveles** — no uses `###` si no hay `##` antes.
3. **Las claves de citas deben coincidir** con las de `referencias.bib`.
4. **Las imágenes deben existir** en `assets/images/` antes de compilar.
5. **`portada.json` va en `markdowns/`**, no en `config/`.
6. **El archivo BibTeX se llama exactamente `referencias.bib`**.
7. **Orden de archivos = orden alfabético** — usa prefijos `01_`, `02_`.
8. **`config/normativa.xlsx` tiene prioridad** sobre `config/normativa.json`.
