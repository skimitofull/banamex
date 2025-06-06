import streamlit as st
import pandas as pd
import numpy as np
from fpdf import FPDF
import textwrap
import io
from datetime import datetime

def parse_banamex_excel(df):
    """
    Parsea el Excel de Banamex manteniendo la estructura original
    donde cada movimiento tiene:
    - Una fila con fecha
    - Varias filas de concepto
    - Última fila con monto y saldo
    """
    df = df.copy()
    df.columns = ['FECHA', 'CONCEPTO', 'RETIROS', 'DEPOSITOS', 'SALDO']
    
    movimientos = []
    i = 0
    
    while i < len(df):
        fecha = df.iloc[i]['FECHA']
        
        if pd.notna(fecha) and str(fecha).strip() not in ['FECHA', 'nan']:
            concepto_completo = ""
            fecha_movimiento = str(fecha).strip()
            
            concepto_inicial = df.iloc[i]['CONCEPTO']
            if pd.notna(concepto_inicial):
                concepto_completo = str(concepto_inicial).strip()
            
            j = i + 1
            monto_retiro = None
            monto_deposito = None
            saldo_final = None
            
            while j < len(df):
                fila_actual = df.iloc[j]
                
                if pd.notna(fila_actual['FECHA']):
                    break
                
                concepto_fila = fila_actual['CONCEPTO']
                if pd.notna(concepto_fila):
                    concepto_texto = str(concepto_fila).strip()
                    if concepto_texto and concepto_texto != 'nan':
                        concepto_completo += " " + concepto_texto
                
                retiro = fila_actual['RETIROS']
                deposito = fila_actual['DEPOSITOS']
                saldo = fila_actual['SALDO']
                
                if pd.notna(retiro) and str(retiro).strip() != 'nan':
                    monto_retiro = float(retiro)
                
                if pd.notna(deposito) and str(deposito).strip() != 'nan':
                    monto_deposito = float(deposito)
                
                if pd.notna(saldo) and str(saldo).strip() != 'nan':
                    saldo_final = float(saldo)
                
                if (monto_retiro is not None or monto_deposito is not None) and saldo_final is not None:
                    break
                
                j += 1
            
            if concepto_completo.strip() and concepto_completo.strip() != 'nan':
                movimiento = {
                    'Fecha': fecha_movimiento,
                    'Concepto': concepto_completo.strip(),
                    'Retiros': monto_retiro if monto_retiro else 0.0,
                    'Depositos': monto_deposito if monto_deposito else 0.0,
                    'Saldo': saldo_final if saldo_final else 0.0
                }
                movimientos.append(movimiento)
            
            i = j
        else:
            i += 1
    
    return pd.DataFrame(movimientos)

class BanamexEstadoCuentaPDF(FPDF):
    def __init__(self, cliente="", numero_cliente="", periodo=""):
        super().__init__(orientation='P', unit='mm', format='A4')
        self.cliente = cliente
        self.numero_cliente = numero_cliente
        self.periodo = periodo
        self.set_auto_page_break(auto=True, margin=15)
        self.page_num = 1
        
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 8, f'ESTADO DE CUENTA AL {self.periodo.upper()}', 0, 1, 'C')
        
        self.set_font('Arial', '', 10)
        self.cell(0, 6, f'CLIENTE: {self.numero_cliente}', 0, 0, 'L')
        self.cell(0, 6, f'Página: {self.page_num}', 0, 1, 'R')
        
        self.set_font('Arial', 'B', 10)
        self.cell(0, 6, self.cliente, 0, 1, 'L')
        self.ln(3)
        
        self.set_font('Arial', 'B', 11)
        self.cell(0, 8, 'DETALLE DE OPERACIONES', 0, 1, 'L')
        
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
    
    def add_movimiento(self, fecha, concepto, retiros, depositos, saldo):
        widths = [20, 95, 25, 25, 25]
        
        self.set_font('Arial', '', 8)
        max_chars_per_line = 45
        concepto_lines = textwrap.wrap(concepto, max_chars_per_line)
        if not concepto_lines:
            concepto_lines = ['']
        
        row_height = 4
        total_height = len(concepto_lines) * row_height
        
        if self.get_y() + total_height > self.page_break_trigger:
            self.add_page()
            self.page_num += 1
        
        start_x = self.get_x()
        start_y = self.get_y()
        
        # Fecha
        self.set_xy(start_x, start_y)
        self.cell(widths[0], total_height, fecha, 1, 0, 'C')
        
        # Concepto multilínea
        self.set_xy(start_x + widths[0], start_y)
        for i, line in enumerate(concepto_lines):
            if i == 0:
                self.cell(widths[1], row_height, line, 'LRT', 0, 'L')
            elif i == len(concepto_lines) - 1:
                self.set_xy(start_x + widths[0], start_y + i * row_height)
                self.cell(widths[1], row_height, line, 'LRB', 0, 'L')
            else:
                self.set_xy(start_x + widths[0], start_y + i * row_height)
                self.cell(widths[1], row_height, line, 'LR', 0, 'L')
        
        # Retiros
        retiros_text = f'{retiros:,.2f}' if retiros > 0 else ''
        self.set_xy(start_x + widths[0] + widths[1], start_y)
        self.cell(widths[2], total_height, retiros_text, 1, 0, 'R')
        
        # Depósitos
        depositos_text = f'{depositos:,.2f}' if depositos > 0 else ''
        self.set_xy(start_x + widths[0] + widths[1] + widths[2], start_y)
        self.cell(widths[3], total_height, depositos_text, 1, 0, 'R')
        
        # Saldo
        saldo_text = f'{saldo:,.2f}' if saldo != 0 else ''
        self.set_xy(start_x + widths[0] + widths[1] + widths[2] + widths[3], start_y)
        self.cell(widths[4], total_height, saldo_text, 1, 0, 'R')
        
        self.set_xy(start_x, start_y + total_height)

# Streamlit App
st.set_page_config(
    page_title="Conversor Banamex Excel → PDF", 
    layout="wide", 
    page_icon="🏦"
)

st.title("🏦 Conversor Estado de Cuenta Banamex")
st.markdown("**Excel → PDF con formato idéntico al original**")
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
    4. Genera el PDF idéntico al original
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
        df_movimientos = parse_banamex_excel(df_original)
        
        st.success(f"✅ Procesados {len(df_movimientos)} movimientos exitosamente")
        
        # Mostrar estadísticas
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("📊 Total Movimientos", len(df_movimientos))
        
        with col2:
            total_retiros = df_movimientos['Retiros'].sum()
            st.metric("💸 Total Retiros", f"${total_retiros:,.2f}")
        
        with col3:
            total_depositos = df_movimientos['Depositos'].sum()
            st.metric("💰 Total Depósitos", f"${total_depositos:,.2f}")
        
        with col4:
            saldo_final = df_movimientos['Saldo'].iloc[-1] if len(df_movimientos) > 0 else 0
            st.metric("🏦 Saldo Final", f"${saldo_final:,.2f}")
        
        # Mostrar vista previa de los datos
        st.header("👀 Vista Previa de Movimientos")
        st.dataframe(
            df_movimientos.head(10),
            use_container_width=True,
            column_config={
                "Retiros": st.column_config.NumberColumn(format="$%.2f"),
                "Depositos": st.column_config.NumberColumn(format="$%.2f"),
                "Saldo": st.column_config.NumberColumn(format="$%.2f")
            }
        )
        
        if len(df_movimientos) > 10:
            st.info(f"Mostrando los primeros 10 de {len(df_movimientos)} movimientos")
        
        # Botón para generar PDF
        st.markdown("---")
        if st.button("🔄 Generar PDF Estado de Cuenta", type="primary", use_container_width=True):
            with st.spinner("Generando PDF idéntico al formato Banamex..."):
                # Crear PDF
                pdf = BanamexEstadoCuentaPDF(
                    cliente=cliente,
                    numero_cliente=numero_cliente,
                    periodo=periodo
                )
                
                pdf.add_page()
                
                # Agregar todos los movimientos
                for _, row in df_movimientos.iterrows():
                    pdf.add_movimiento(
                        fecha=row['Fecha'],
                        concepto=row['Concepto'],
                        retiros=row['Retiros'],
                        depositos=row['Depositos'],
                        saldo=row['Saldo']
                    )
                
                # Generar PDF en memoria
                pdf_bytes = bytes(pdf.output(dest='S').encode('latin-1'))
                
                st.success("✅ PDF generado exitosamente!")
                
                # Botón de descarga
                st.download_button(
                    label="📥 Descargar Estado de Cuenta PDF",
                    data=pdf_bytes,
                    file_name=f"estado_cuenta_{numero_cliente}_{datetime.now().strftime('%Y%m%d')}.pdf",
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
        **El Excel debe tener esta estructura:**
        
        | FECHA | CONCEPTO | RETIROS | DEPOSITOS | SALDO |
        |-------|----------|---------|-----------|-------|
        | 22 DIC | SALDO ANTERIOR | | | 3000 |
        | 23 DIC | DEPOSITO POR DEVOLUCION DE | | | |
        | | MERCANCIA | | | |
        | | 75445504354481090854912 | | | |
        | | SUC 0342 | | | |
        | | CAJA 0093 AUT 02132404 HORA 06:46 | | 2000 | 5000 |
        
        **Cada movimiento tiene:**
        - Una fila con fecha
        - Varias filas con el concepto completo
        - La última fila con el monto y saldo final
        """)
