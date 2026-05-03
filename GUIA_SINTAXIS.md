# Guía de sintaxis — Sistema Editorial

Referencia rápida de cómo escribir cada elemento de una tesis en Markdown.

---

## Estructura de archivos

```
markdowns/
├── 01_Introduccion.md
├── 02_Marco_Teorico.md    ← prefijo numérico define el orden de compilación
├── ...
└── referencias.bib        ← fuentes bibliográficas (obligatorio si hay citas)

assets/
└── images/                ← imágenes del documento
```

---

## Títulos y encabezados

```markdown
# Título del capítulo          → TIT_CAP   (negrita, centrado)
## Sección                     → H1_APA    (negrita, centrado)
### Subsección                 → H2_APA    (negrita, izquierda)
#### Sub-subsección            → H3_APA    (negrita+cursiva, izquierda)
##### Nivel 5                  → H4_APA    (cursiva, izquierda con sangría)
###### Nivel 6                 → H5_APA    (cursiva, izquierda con sangría más)
```

---

## Párrafo de texto

```markdown
Texto normal del párrafo. Los espacios en blanco entre párrafos
son ignorados — el interlineado y espaciado los define la normativa.

Siguiente párrafo aquí.
```

---

## Citas bibliográficas

```markdown
Según [@autor2024], el tema es relevante.

Varios autores [@autor1; @autor2] coinciden en que...
```

Las claves deben coincidir con los IDs en `referencias.bib`.

---

## Figura con título y nota

```markdown
> [!FIG_TIT]
> Título descriptivo de la figura

![](../assets/images/nombre_imagen.png)
> [!NOTA_FIG]
> *Nota.* Autor (año). Descripción. Fuente o URL.
```

- `FIG_TIT` genera automáticamente el número de figura (`Figura 1`, `Figura 2`…)
- La imagen va entre `FIG_TIT` y `NOTA_FIG`
- La nota es opcional; si no hay nota, omitir el bloque `NOTA_FIG`

---

## Imagen sin título (ilustrativa)

```markdown
![Descripción alternativa](../assets/images/nombre_imagen.png)
```

Se inserta centrada, sin número ni nota. Usar solo para imágenes decorativas.

---

## Tabla

```markdown
> [!TABLA_TIT src="assets/tablas/datos.xlsx" sheet="Hoja1"]
> Título descriptivo de la tabla
```

- `src` ruta relativa al raíz del proyecto
- `sheet` nombre de la hoja (si se omite, usa la primera hoja)
- `TABLA_TIT` genera automáticamente el número de tabla

---

## Ecuación

```markdown
> [!ECUACION]
> E = mc²
```

Se renderiza como texto con estilo de ecuación (no LaTeX).

---

## Cualquier etiqueta personalizada

```markdown
> [!NOMBRE_ETIQUETA]
> Contenido del elemento
```

`NOMBRE_ETIQUETA` debe existir como `ID_Etiqueta` en la normativa activa.

---

## Páginas previas al índice

Portada, resumen, dedicatoria, agradecimientos, etc. **no se generan desde Markdown** — su formato varía según universidad y se editan directamente en Word.

El ensamblador comienza desde el primer archivo `.md` numerado. Las páginas previas al índice deben agregarse manualmente al documento final.

---

## Referencias — `referencias.bib`

### Artículo de revista
```bibtex
@article{autor2024,
  author  = {Apellido, Nombre},
  title   = {Título del artículo},
  journal = {Nombre de la Revista},
  year    = {2024},
  volume  = {10},
  number  = {2},
  pages   = {100--115},
  doi     = {10.1000/xyz}
}
```

### Libro
```bibtex
@book{autor2020,
  author    = {Apellido, Nombre},
  title     = {Título del libro},
  publisher = {Editorial},
  year      = {2020}
}
```

### Conferencia / actas
```bibtex
@inproceedings{autor2023,
  author    = {Apellido, Nombre},
  title     = {Título del trabajo},
  booktitle = {Nombre del Congreso},
  year      = {2023},
  pages     = {45--50}
}
```

---

## Comentarios (ignorados por el ensamblador)

```markdown
<!-- Este texto no aparece en el Word -->
<!-- etiqueta implícita: H1_APA -->
```

Útil para anotaciones personales o etiquetas de referencia visual en Obsidian.

---

## Reglas generales

| Regla | Detalle |
|---|---|
| Orden de archivos | Prefijo numérico (`01_`, `02_`…) define el orden |
| Imágenes soportadas | `.png`, `.jpg`, `.jpeg`, `.webp` (WebP se convierte a PNG automáticamente) |
| Nombre de referencias | `referencias.bib` exacto, en `markdowns/` |
| Captions de callouts | Cada línea de contenido debe iniciar con `> ` |
| Etiquetas | Siempre en mayúsculas: `FIG_TIT`, `NOTA_FIG`, `TABLA_TIT`, `ECUACION` |
