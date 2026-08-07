# BioSize Clinical

Aplicación Streamlit para cálculo trazable del tamaño de muestra en investigación clínica.

## Funciones principales

- Importación de XLSX, CSV y SQLite mediante pandas y SQLAlchemy.
- Lectura automática de protocolos en **JSON, YAML, PDF y Word** (`.docx`, `.docm` y `.doc`).
- Extracción de parámetros desde párrafos o tablas: diseño, error alfa, poder estadístico, efecto esperado, variabilidad, pérdidas y parámetros específicos.
- Reconocimiento de porcentajes (`5%`, `80%`), decimales con coma (`0,05`) y denominaciones clínicas como DMCI, margen de no inferioridad, prevalencia, sensibilidad, especificidad, OR, RR y hazard ratio.
- Panel de revisión con los parámetros detectados y vista previa del texto extraído.
- Diseños descriptivos, comparativos, superioridad, no inferioridad, equivalencia, OR, RR, precisión diagnóstica y supervivencia.
- Ajuste por población finita, pérdidas esperadas y corrección de continuidad.
- Curva dinámica de potencia o, cuando corresponde, curva de precisión.
- Informe Word editable con justificación matemática y párrafo CONSORT.
- Checklist operativo CONSORT 2010 en Word y PDF.

## Requisitos para PDF y Word

- Los PDF deben contener **texto seleccionable**. Un PDF formado únicamente por imágenes necesita OCR previo.
- El formato Word recomendado es `.docx`.
- Los archivos `.doc` antiguos se leen mediante `antiword`/`catdoc` cuando están disponibles; de lo contrario se aplica una recuperación heurística. Para máxima confiabilidad, conviértalos a `.docx`.
- Para lograr una detección inequívoca, conserve etiquetas como:

```text
Tipo de diseño: medias independientes
Error alfa: 0,05
Poder estadístico: 0,80
Efecto esperado: 5
Variabilidad estimada: 10
Pérdidas esperadas: 10%
```

La app incluye plantillas descargables en JSON, YAML, Word y PDF.

## Instalación local

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Nota metodológica

Las fórmulas son herramientas de planificación. Diseños por conglomerados, multiplicidad, análisis intermedios, riesgos no proporcionales, requisitos regulatorios o modelos complejos requieren validación bioestadística específica.

La app genera el checklist CONSORT 2010 solicitado. Para nuevos ensayos, conviene revisar también la actualización CONSORT vigente.
