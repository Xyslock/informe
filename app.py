# app.py
import streamlit as st
from datetime import datetime, timedelta
import os
from generar_informe import crear_informe
# 👇 Esto es nuevo: forzamos reconstrucción
st.write("")

st.set_page_config(
    page_title="Informe de Afectaciones Críticas",
    page_icon="📄",
    layout="centered"
)

st.title("📄 Generador de Informe de Afectaciones Críticas")

# Fechas por defecto: última semana
hoy = datetime.today()
ultima_semana = hoy - timedelta(days=7)

col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("📅 Fecha de inicio", value=ultima_semana)
with col2:
    end_date = st.date_input("📅 Fecha de fin", value=hoy)

# Opción para subir CSV
st.markdown("---")
uploaded_file = st.file_uploader("📂 Subir archivo de incidentes (CSV, opcional)", type="csv")

if start_date > end_date:
    st.error("❌ La fecha de inicio no puede ser mayor que la fecha de fin.")
else:
    if st.button("🚀 Generar Informe"):
        with st.spinner("Generando informe..."):
            start_str = start_date.strftime("%d/%m/%Y")
            end_str = end_date.strftime("%d/%m/%Y")

            csv_path = 'incidentes.csv'
            if uploaded_file is not None:
                with open(csv_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

            try:
                doc_bytes = crear_informe(start_str, end_str, csv_path=csv_path)

                st.success("✅ ¡Informe generado con éxito!")

                st.download_button(
                    label="📥 Descargar Informe (.docx)",
                    data=doc_bytes,
                    file_name="Informe_Afectaciones_Criticas.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

                # Limpiar archivo CSV si se subió
                if uploaded_file is not None and os.path.exists(csv_path):
                    os.remove(csv_path)

            except Exception as e:
                st.error(f"❌ Error al generar el informe: {str(e)}")

                st.exception(e)
