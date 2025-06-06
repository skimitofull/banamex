import streamlit as st
import pandas as pd
import numpy as np
from fpdf import FPDF
import io
from datetime import datetime

def parse_banamex_excel_individual_rows(df):
    """
    Parsea el Excel de Banamex manteniendo CADA FILA INDIVIDUAL
    tal como aparece en el Excel original
    """
    df = df.copy()
    df.columns = ['FECHA', 'CONCEPTO', 'RETIROS', 'DEPOSITOS', 'SALDO']
    
    # Limpiar y procesar cada fila individualmente
    filas_procesadas = []
    
    for idx, row in df.iterrows():
        fila_procesada = {
            'FECHA': str(row['FECHA']).strip() if pd.notna(row['FECHA']) and str(row['FECHA']).strip() != 'nan' else '',
            'CONCEPTO': str(row['CONCEPTO']).strip() if pd.notna(row['CONCEPTO']) and str(row['CONCEPTO']).strip() != 'nan' else '',
            'RETIROS': float(row['RETIROS']) if pd.notna(row['RETIROS']) and str(row['RETIROS']).strip() != 'nan' else None,
            'DEPOSITOS': float(row['DEPOSITOS']) if pd.notna(row['DEPOSITOS']) and str(row['DEPOSITOS']).strip() != 'nan' else None,
            'SALDO': float(row['SALDO']) if pd.notna(row['SALDO']) and str(row['SALDO']).strip() != 'nan' else None
        }
        
        # Solo agregar filas que no sean completamente vacías
        if any([fila_procesada['FECHA'], fila_procesada['CONCEPTO'], 
                fila_procesada['RETIROS'], fila_procesada['DEPOSITOS'], fila_procesada['SALDO']]):
            filas_procesadas.append(fila_procesada)
    
    return pd.DataFrame(filas_procesadas)

class BanamexEstadoCuentaPDFOriginal(FPDF):
    def __init__(self, cliente="", numero_cliente="", periodo=""):
        super().__init__(orientation='P', unit='mm', format='A4')
        self.cliente = cliente
        self.numero_cliente = numero_cliente
        self.periodo = periodo
        self.set_auto_page_break(auto=True, margin=20)
        self.page_num = 1
        
    def header(self):
        # Título principal
        self.set_font('Arial', 'B', 14)
        self.cell(0, 8, f'ESTADO DE CUENTA AL {self.periodo.upper()}', 0, 1, 'C')
        self.ln(2)
        
        # Información del cliente
        self.set_font('Arial', 'B', 10)
        self.cell(40, 6, 'CLIENTE:', 0, 0, 'L')
        self.cell(0, 6, f'Página: {self.page_num} de 29', 0, 1, 'R')
        
        self.set_font('Arial', '', 10)
        self.cell(0, 6, self.numero_cliente, 0, 1, 'L')
        
        self.set_font('Arial', 'B', 10)
        self.cell(0, 6, self.cliente, 0, 1, 'L')
        self.ln(3)
        
        # Información adicional en páginas > 1
        if self.page_num > 1:
            self.set_font('Arial', '', 8)
            self.cell(0, 4, 'Centro de Atención Telefónica', 0, 1, 'L')
            self.cell(0, 4, 'Ciudad de México: 55 1226 2639', 0, 1, 'L')
            self.cell(0, 4, 'Resto del país: 800 021 2345', 0, 1, 'L')
            self.ln(2)
        
        # Título de la tabla
        self.set_font('Arial', 'B', 11)
        self.cell(0, 6, 'DETALLE DE OPERACIONES', 0, 1, 'L')
        self.ln(1)
        
        # Encabezados de la tabla
        self.set_font('Arial', 'B', 9)
        headers = ['FECHA', 'CONCEPTO', 'RETIROS', 'DEPOSITOS', 'SALDO']
        widths = [20, 95, 25, 25, 25]
        
        for header, width in zip(headers, widths):
            self.cell(width, 6, header, 1, 0, 'C')
        self.ln()
    
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', '', 8)
        self.cell(0, 10, '000191.B41EJDA029.OD.0121.01', 0, 0, 'L')
    
    def add_fila_individual(self, fecha, concepto, retiros, depositos, saldo):
        """
        Agrega UNA SOLA FILA a la tabla, exactamente como aparece en el Excel
        """
        widths = [20, 95, 25, 25, 25]
        aligns = ['C', 'L', 'R', 'R', 'R']
        
        # Verificar si necesitamos nueva página
        if self.get_y() + 6 > self.page_break_trigger:
            self.add_page()
            self.page_num += 1
        
        # Preparar valores para mostrar
        valores = [
            fecha if fecha else '',
            concepto if concepto else '',
            f'{retiros:,.2f}' if retiros is not None and retiros != 0 else '',
            f'{depositos:,.2f}' if depositos is not None and depositos != 0 else '',
            f'{saldo:,.2f}' if saldo is not None and saldo != 0 else ''
        ]
        
        # Configurar fuente para el contenido
        self.set_font('Arial', '', 8)
        
        # Agregar cada celda
        for i, (valor, width, align) in enumerate(zip(valores, widths, aligns)):
            self.cell(width, 6, valor, 1, 0, align)
        
        self.ln()

# Streamlit App
st.set_page_config(
    page_title="Conversor Banamex Excel → PDF", 
    layout="wide", 
    page_icon="🏦"
)

st.title("🏦 Conversor Estado de Cuenta Banamex")
st.markdown("**Excel → PDF con formato IDÉNTICO al original (fila por fila)**")
st.markdown("---")

# Sidebar con información
with st.sidebar:
    st.header("📋 Información del Cliente")
    cliente = st.text_input("Nombre del Cliente", "PATRICIA IÑIGUEZ FLORES")
    numero_cliente = st.text_input("Número de Cliente", "61900627")
    periodo = st.text_input("Período", "21 DE ENERO DE 2025")
    
    st.markdown("---")
    st.markdown("### 📖 Instrucciones")
    st.markdown("""
    1. Sube el archivo Excel exportado del PDF
    2. Verifica los datos procesados
    3. Ajusta la información del cliente
    4. Genera el PDF IDÉNTICO al original
    """)

# Área principal
col1, col2 = st.columns([2, 1])

with col1:
    st.header("📤 Cargar Archivo Excel")
    uploaded_file = st.file_uploader(
        "Selecciona el archivo Excel del estado de cuenta",
        type=['xlsx', 'xls'],
        help="Archivo Excel exportado directamente del PDF de Banamex"
    )

with col2:
    if uploaded_file:
        st.success("✅ Archivo cargado")
        st.info(f"📄 {uploaded_file.name}")

if uploaded_file is not None:
    try:
        # Procesar el archivo
        df_original = pd.read_excel(uploaded_file)
        df_filas = parse_banamex_excel_individual_rows(df_original)
        
        st.success(f"✅ Procesadas {len(df_filas)} filas exitosamente")
        
        # Mostrar estadísticas
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("📊 Total Filas", len(df_filas))
        
        with col2:
            total_retiros = df_filas['RETIROS'].sum() if 'RETIROS' in df_filas.columns else 0
            st.metric("💸 Total Retiros", f"${total_retiros:,.2f}")
        
        with col3:
            total_depositos = df_filas['DEPOSITOS'].sum() if 'DEPOSITOS' in df_filas.columns else 0
            st.metric("💰 Total Depósitos", f"${total_depositos:,.2f}")
        
        with col4:
            # Buscar el último saldo no nulo
            saldos_validos = df_filas[df_filas['SALDO'].notna()]['SALDO']
            saldo_final = saldos_validos.iloc[-1] if len(saldos_validos) > 0 else 0
            st.metric("🏦 Saldo Final", f"${saldo_final:,.2f}")
        
        # Mostrar vista previa de los datos
        st.header("👀 Vista Previa de Filas (Formato Original)")
        st.dataframe(
            df_filas.head(15),
            use_container_width=True,
            column_config={
                "RETIROS": st.column_config.NumberColumn(format="$%.2f"),
                "DEPOSITOS": st.column_config.NumberColumn(format="$%.2f"),
                "SALDO": st.column_config.NumberColumn(format="$%.2f")
            }
        )
        
        if len(df_filas) > 15:
            st.info(f"Mostrando las primeras 15 de {len(df_filas)} filas")
        
        # Botón para generar PDF
        st.markdown("---")
        if st.button("🔄 Generar PDF Estado de Cuenta (Formato Original)", type="primary", use_container_width=True):
            with st.spinner("Generando PDF IDÉNTICO al formato Banamex original..."):
                # Crear PDF
                pdf = BanamexEstadoCuentaPDFOriginal(
                    cliente=cliente,
                    numero_cliente=numero_cliente,
                    periodo=periodo
                )
                
                pdf.add_page()
                
                # Agregar CADA FILA INDIVIDUAL
                for _, row in df_filas.iterrows():
                    pdf.add_fila_individual(
                        fecha=row['FECHA'],
                        concepto=row['CONCEPTO'],
                        retiros=row['RETIROS'],
                        depositos=row['DEPOSITOS'],
                        saldo=row['SALDO']
                    )
                
                # Generar PDF en memoria
                buffer = io.BytesIO()
                pdf.output(buffer)
                pdf_bytes = buffer.getvalue()
                
                st.success("✅ PDF generado exitosamente con formato IDÉNTICO!")
                
                # Botón de descarga
                st.download_button(
                    label="📥 Descargar Estado de Cuenta PDF (Formato Original)",
                    data=pdf_bytes,
                    file_name=f"estado_cuenta_original_{numero_cliente}_{datetime.now().strftime('%Y%m%d')}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
    
    except Exception as e:
        st.error(f"❌ Error al procesar el archivo: {str(e)}")
        st.info("Verifica que el archivo sea un Excel válido exportado del PDF de Banamex")

else:
    st.info("👆 Sube un archivo Excel para comenzar")
    
    # Mostrar ejemplo de estructura esperada
    with st.expander("📋 Ver estructura esperada del Excel"):
        st.markdown("""
        **El Excel debe tener esta estructura (FILA POR FILA):**
        
        | FECHA | CONCEPTO | RETIROS | DEPOSITOS | SALDO |
        |-------|----------|---------|-----------|-------|
        | 22 DIC | SALDO ANTERIOR | | | 44230.27 |
        | 23 DIC | DEPOSITO POR DEVOLUCION DE | | | |
        | | MERCANCIA | | | |
        | | 75445504354481086801511 | | | |
        | | SUC 0342 | | | |
        | | CAJA 0093 AUT 02132404 HORA 06:46 | | 208.86 | 44439.13 |
        
        **Cada fila del Excel = Una fila en el PDF**
        - No se agrupan conceptos
        - Cada línea se respeta individualmente
        - Formato idéntico al PDF original de Banamex
        """)
