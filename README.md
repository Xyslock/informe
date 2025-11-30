# 📄 Generador de Informe de Afectaciones Críticas

Este proyecto permite generar **informes automatizados en formato .docx** sobre afectaciones críticas de servicios, a partir de un archivo CSV de incidentes.  
Fue desarrollado usando **Python**, **Streamlit**, **Pandas**, **Matplotlib** y **python-docx**.

El objetivo es proporcionar una herramienta eficiente para consolidar, analizar y presentar incidentes críticos de forma clara, estandarizada y con gráficos automáticos.

---

## 🚀 Características principales

- Interfaz web construida con **Streamlit**.  
- Generación automática de un informe profesional en **Word (.docx)**.  
- Portada con logo institucional.  
- Tablas detalladas de incidentes.  
- Cálculo de estadísticas clave.  
- Gráficos automáticos:
  - Top 5 servicios con mayor número de incidentes.
  - Distribución por tipo de afectación.
- Top 5 incidentes de mayor duración.
- Soporte para múltiples codificaciones CSV.
- Encabezados, pie de página y numeración automática.

---

## 🗂 Estructura del proyecto

```
📁 proyecto/
│
├── app.py
├── generar_informe.py
├── incidentes.csv (opcional)
├── img/
│   └── logo.png
└── README.md
```

---

## 🧩 Requisitos

### Python 3.9 o superior  
Instalar dependencias:

```bash
pip install streamlit pandas python-docx matplotlib
```

---

## ▶️ Ejecución del proyecto

```bash
streamlit run app.py
```

Luego abre el enlace local generado por Streamlit.

---

## 📥 Uso de la aplicación

1. Selecciona el **rango de fechas** a analizar.  
2. (Opcional) Carga un archivo **CSV** con incidentes.  
3. Haz clic en **"Generar Informe"**.  
4. Descarga el archivo **.docx** generado.

---

## 📝 Formato del informe generado

Incluye:

- Portada profesional  
- Objetivo  
- Alcance  
- Metodología  
- Gráficos  
- Top 5 incidentes más largos  
- Tabla de incidentes  
- Totales finales

---

## 📊 Gráficos incluidos

- **Top 5 servicios con más incidentes**  
- **Distribución por tipo de afectación**

---

## 🧠 Lógica general del procesamiento

1. Lectura del CSV con manejo de varias codificaciones.  
2. Limpieza y normalización de datos.  
3. Filtrado por fecha.  
4. Cálculo de estadísticas.  
5. Generación del documento Word.  
6. Inserción de gráficos e imágenes.  
7. Exportación del archivo final.

---

## 🏢 Créditos
Desarrollado para automatizar la generación de informes dentro del marco operativo de **Axity**.
