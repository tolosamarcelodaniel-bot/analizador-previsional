import streamlit as st
import pandas as pd
import re
from pypdf import PdfReader

st.set_page_config(page_title="Analizador Previsional PDF", layout="wide")

st.title("🛠️ Herramienta de Análisis Previsional (Lectura de PDFs)")
st.markdown("Cargue su archivo PDF de historia laboral para extraer texto, mapear campos y calcular las reglas del mejor cargo.")

# 1. CARGA DE ARCHIVO (Ahora acepta PDF)
archivo_subido = st.file_uploader("📂 Arrastre o seleccione el archivo PDF de origen (.pdf)", type=["pdf"])

if archivo_subido is not None:
    st.success("✅ PDF cargado correctamente.")
    
    # Extraer texto del PDF
    with st.spinner("Leyendo y extrayendo texto del PDF..."):
        lector_pdf = PdfReader(archivo_subido)
        texto_completo = ""
        for pagina in lector_pdf.pages:
            texto_completo += pagina.extract_text() + "\n"
            
    st.subheader("📋 Vista Previa del Texto Extraído")
    st.text_area("Texto detectado en el documento (primeros 2000 caracteres):", texto_completo[:2000], height=200)

    st.markdown("---")
    st.subheader("🎛️ Configuración del Análisis Previsional")
    
    col1, col2 = st.columns(2)
    with col1:
        cuil_usuario = st.text_input("Introduzca el CUIL/Documento a buscar en el PDF:", value="20-34567890-9")
        cargo_buscar = st.text_input("Escriba el nombre exacto del Cargo de Mayor Jerarquía:", value="Director")
    with col2:
        meses_manuales = st.number_input("Meses acumulados encontrados para este cargo:", min_value=0, value=36, step=1)

    # 2. MOTOR LÓGICO DE REGLAS PREVISIONALES
    if st.button("🚀 Calcular Viabilidad de Mejor Cargo"):
        st.subheader("📊 Resultado del Análisis Condicional")
        
        # Buscar si el cargo existe en el texto extraído
        existe_cargo = cargo_buscar.lower() in texto_completo.lower()
        
        if existe_cargo:
            st.info(f"🔍 El cargo **'{cargo_buscar}'** fue localizado dentro del documento PDF.")
            
            # Evaluación de Reglas en Cascada según tu requerimiento
            if meses_manuales >= 36:
                st.success(f"🏆 **Regla 1 Cumplida:** El cargo califica como Mejor Cargo por Continuidad (Posee {meses_manuales} meses continuos, mínimo requerido: 36).")
            elif meses_manuales >= 60:
                st.warning(f"⚖️ **Regla 2 Cumplida:** El cargo califica como Mejor Cargo por Alternancia (Posee {meses_manuales} meses alternados, mínimo requerido: 60).")
            else:
                st.error(f"❌ **No Califica:** El cargo fue hallado pero solo acumula {meses_manuales} meses (No alcanza los 36 continuos ni los 60 alternados).")
                
            # Crear un reporte limpio para descargar
            reporte_datos = {
                "CUIL/Documento": [cuil_usuario],
                "Cargo Evaluado": [cargo_buscar],
                "Meses Registrados": [meses_manuales],
                "Condición Legal": ["Continuidad (Regla 1)" if meses_manuales >= 36 else ("Alternancia (Regla 2)" if meses_manuales >= 60 else "No Califica")]
            }
            df_reporte = pd.DataFrame(reporte_datos)
            
            # Botón de descarga
            @st.cache_data
            def convertir_excel(df):
                import io
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Resultado_PDF')
                return output.getvalue()
                
            st.download_button(
                label="📥 Descargar Dictamen en Excel",
                data=convertir_excel(df_reporte),
                file_name=f"Dictamen_Previsional_{cuil_usuario}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error(f"❌ El cargo '{cargo_buscar}' no pudo ser encontrado textualmente en el PDF cargado. Revise la ortografía en la vista previa.")
