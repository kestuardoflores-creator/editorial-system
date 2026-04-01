# 📚 Sistema de Gestión Editorial

Automatiza la redacción, etiquetado y compilación de documentos académicos (tesis, informes técnicos) usando **Google Docs + Google Sheets + Google Colab**.

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/kestuardoflores-creator/editorial-system/blob/main/installer/installer.ipynb)

---

## ¿Qué hace este sistema?

- **Etiqueta** cada párrafo con su tipo (H1, párrafo, figura, tabla…) desde un sidebar en Google Docs.
- **Aplica formato APA 7** automáticamente desde una base de datos central en Google Sheets.
- **Ensambla** todos los capítulos en un único documento final, con numeración automática e índices.

## Inicio rápido

### Pre-requisito (una sola vez)
Habilita la **Apps Script API**:
1. Ve a → https://console.cloud.google.com/apis/library/script.googleapis.com
2. Haz clic en **Habilitar**.

### Instalación
1. Haz clic en el badge **Open in Colab** de arriba.
2. Ejecuta las celdas en orden.
3. Tu proyecto quedará creado en Google Drive automáticamente.

> **Idempotente:** puedes ejecutar el installer varias veces. Solo actualiza los estilos APA 7; no elimina ni duplica tus archivos.

---

## Estructura del repositorio

```
editorial-system/
├── installer/
│   └── installer.ipynb       ← Installer (1 clic en Colab)
├── assembler/
│   └── assembler.ipynb       ← Ensamblador (próximamente)
├── apps-script/
│   ├── sidebar.js            ← Lógica del sidebar
│   ├── format-engine.js      ← Motor de formato
│   └── sidebar.html          ← UI del sidebar
├── config/
│   └── default-styles.json   ← Estilos APA 7 por defecto
├── docs/
│   └── Anteproyecto.md       ← Diseño técnico completo
└── examples/
    └── test-case/
        └── demo.ipynb        ← Demo con 2 capítulos (próximamente)
```

---

## Estilos incluidos (APA 7)

| Etiqueta | Descripción |
|---|---|
| `H1_APA` | Encabezado Nivel 1 — Centrado, negrita |
| `H2_APA` | Encabezado Nivel 2 — Izquierda, negrita |
| `H3_APA` | Encabezado Nivel 3 — Izquierda, negrita cursiva |
| `H4_APA` | Encabezado Nivel 4 — Indentado, negrita |
| `H5_APA` | Encabezado Nivel 5 — Indentado, negrita cursiva |
| `TIT_CAP` | Título de capítulo |
| `TEXTO_APA` | Párrafo normal |
| `FIG_TIT` | Título de figura |
| `TABLA_TIT` | Título de tabla |

Puedes agregar tus propios estilos editando `Configuracion_Estilos` en el Google Sheet del proyecto.

---

## FAQ

**¿Puedo usar otra normativa además de APA 7?**
Sí. Edita la hoja `Configuracion_Estilos` o contribuye un nuevo archivo en `config/`.

**¿Qué pasa si ejecuto el installer dos veces?**
Nada malo. Solo se actualizan los estilos APA 7. Tus documentos y datos no se tocan.

**¿Necesito instalar algo localmente?**
No. Todo corre en Google Colab y Google Drive.
