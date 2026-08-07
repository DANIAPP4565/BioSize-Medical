# BioSize Clinical v1.2

Aplicación Streamlit para cálculo trazable del tamaño de muestra en investigación clínica.

## Funciones principales

- Importación de XLSX, CSV y SQLite mediante pandas y SQLAlchemy.
- Lectura automática de protocolos en **JSON, YAML, PDF y Word** (`.docx`, `.docm` y `.doc`).
- Extracción de diseño, error alfa, poder, efecto esperado, variabilidad, pérdidas y parámetros específicos.
- Identificación asistida del **outcome primario** desde párrafos, tablas o estructuras JSON/YAML.
- Detección de nombre, definición operacional, tipo, unidad, momento de evaluación, columna esperada y valor del evento.
- Clasificación del outcome como continuo, binario, binario apareado, tiempo hasta evento, diagnóstico, ordinal o conteo.
- Verificación de coherencia entre el tipo de outcome y el diseño estadístico seleccionado.
- Vinculación semiautomática del outcome con columnas del dataset mediante similitud semántica, compatibilidad de tipo y nombres preferidos.
- Presentación de alternativas de columna y **nivel de confianza** de la correspondencia.
- Confirmación obligatoria del investigador antes de ejecutar el cálculo.
- Estimación de desviación estándar o proporción directamente desde la columna confirmada.
- Soporte para outcomes primarios/coprimarios estructurados: cálculo por outcome y recomendación del mayor N cuando existen parámetros completos.
- Diseños descriptivos, comparativos, superioridad, no inferioridad, equivalencia, OR, RR, precisión diagnóstica y supervivencia.
- Ajuste por población finita, pérdidas esperadas y corrección de continuidad.
- Curva dinámica de potencia o precisión.
- Informe Word editable con definición del outcome, columna vinculada, confianza, justificación matemática y párrafo CONSORT.
- Checklist operativo CONSORT 2010 en Word y PDF.

## Campos recomendados para el outcome

Los protocolos pueden incluir:

```text
Desenlace primario: Presión arterial sistólica a las 12 semanas
Definición del desenlace: Cambio desde el valor basal hasta la semana 12
Tipo de outcome: continuo
Unidad del outcome: mmHg
Momento de evaluación: 12 semanas
Columna del dataset: PAS_12semanas
Valor del evento: 1
Desenlaces secundarios: Presión arterial diastólica y control tensional
```

En JSON/YAML se utilizan las claves:

```yaml
outcome_primario: Presión arterial sistólica a las 12 semanas
definicion_outcome: Cambio desde el valor basal hasta la semana 12
tipo_outcome: continuo
unidad_outcome: mmHg
momento_evaluacion: 12 semanas
columna_dataset_outcome: PAS_12semanas
valor_evento: 1
outcome_secundarios: Presión arterial diastólica
```

## Outcomes coprimarios estructurados

Para calcular varios outcomes primarios o coprimarios, utilice una lista `outcomes`. Cada outcome debe aportar su propio diseño, efecto y variabilidad:

```json
{
  "tipo_diseno": "medias_independientes",
  "error_alfa": 0.05,
  "poder_estadistico": 0.80,
  "efecto_esperado": 5,
  "variabilidad_estimada": 10,
  "outcomes": [
    {
      "nombre": "Presión arterial sistólica a las 12 semanas",
      "rol": "primario",
      "tipo": "continuo",
      "tipo_diseno": "medias_independientes",
      "efecto_esperado": 5,
      "variabilidad_estimada": 10
    },
    {
      "nombre": "Presión arterial diastólica a las 12 semanas",
      "rol": "coprimario",
      "tipo": "continuo",
      "tipo_diseno": "medias_independientes",
      "efecto_esperado": 3,
      "variabilidad_estimada": 10
    }
  ]
}
```

La app calcula cada escenario y adopta el mayor N. Cuando faltan parámetros para un outcome coprimario, lo informa y no inventa valores.

## Requisitos para PDF y Word

- Los PDF deben contener **texto seleccionable**. Un PDF formado únicamente por imágenes necesita OCR previo.
- El formato Word recomendado es `.docx`.
- Los archivos `.doc` antiguos se leen mediante `antiword` o `catdoc` cuando están disponibles; de lo contrario se aplica recuperación heurística.
- La detección automática es una propuesta. El cálculo queda bloqueado hasta que el investigador confirme el outcome y su vinculación con los datos.

## Instalación local

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Archivos de ejemplo

- `protocolo_ejemplo.json`
- `protocolo_ejemplo.yaml`
- `protocolo_ejemplo.docx`
- `protocolo_ejemplo.pdf`
- `protocolo_coprimarios_ejemplo.json`

## Nota metodológica

Las fórmulas son herramientas de planificación. La selección y definición del outcome primario debe realizarse antes del análisis. Los diseños con multiplicidad, outcomes coprimarios, conglomerados, análisis intermedios, riesgos no proporcionales o requisitos regulatorios requieren validación bioestadística específica.
