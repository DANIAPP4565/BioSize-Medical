# BioSize Clinical

Aplicación Streamlit para cálculo trazable del tamaño de muestra en investigación clínica.

## Funciones principales

- Importación de XLSX, CSV y SQLite mediante pandas y SQLAlchemy.
- Importación automática de protocolos JSON/YAML.
- Diseños descriptivos, comparativos, superioridad, no inferioridad, equivalencia, OR, RR, precisión diagnóstica y supervivencia.
- Ajuste por población finita, pérdidas esperadas y corrección de continuidad.
- Curva dinámica de potencia o, cuando corresponde, curva de precisión.
- Informe Word editable con justificación matemática y párrafo CONSORT.
- Checklist operativo CONSORT 2010 en Word y PDF.

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
