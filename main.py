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
        
        # Agregar todas las filas (incluso las aparentemente vacías)
        filas_procesadas.append(fila_procesada)
    
    return pd.DataFrame(filas_procesadas)

class BanamexEstadoCuentaPDFExacto(FPDF):
    def __init__(self, cliente="", numero_cliente="", periodo=""):
        # Tamaño exacto: 187.33 mm x 279.4 mm = 531 pt x 792 pt
        super().__init__(orientation='P', unit='pt', format=(531, 792))
        self.cliente = cliente
        self.numero_cliente = numero_cliente
        self.periodo = periodo
        self.set_auto_page_break(auto=False)  # Control manual de páginas
        self.page_num = 1
        self.filas_en_pagina = 0
        self.max_filas_por_pagina = 52  # Máximo 52 filas incluyendo encabezado
        self.y_inicial_datos = 104.73901  # Posición inicial de los datos
        self.altura_fila = 12.3037  # Altura calculada: (733.54797 - 104.73901) / 51
        self.fila_global = 0  # Contador global para alternado de colores
        
        # Posiciones horizontales exactas
        self.x_fecha = 14.361
        self.x_concepto = 57.378  # Calculado desde FECHA + ancho
        self.x_retiros = 369.129
        self.x_depositos = 440.0  # Calculado desde RETIROS + ancho
        self.x_saldo = 510.87  # Calculado desde DEPOSITOS + ancho
        
        # Anchos de columnas exactos
        self.w_fecha = 43.017  # x_concepto - x_fecha
        self.w_concepto = 311.751  # x_retiros - x_concepto
        self.w_retiros = 70.871  # x_depositos - x_retiros
        self.w_depositos = 70.87  # x_saldo - x_depositos
        self.w_saldo = 70.87  # Ancho restante
        
    def header(self):
        # Título principal
        self.set_font('Helvetica', 'B', 14)
        self.set_xy(0, 20)
        self.cell(531, 14, f'ESTADO DE CUENTA AL {self.periodo.upper()}', 0, 1, 'C')
        
        # Información del cliente
        self.set_font('Helvetica', 'B', 10)
        self.set_xy(14.361, 45)
        self.cell(200, 10, 'CLIENTE:', 0, 0, 'L')
        self.set_xy(350, 45)
        self.cell(0, 10, f'Página: {self.page_num} de 29', 0, 0, 'R')
        
        self.set_font('Helvetica', '', 10)
        self.set_xy(14.361, 58)
        self.cell(0, 10, self.numero_cliente, 0, 0, 'L')
        
        self.set_font('Helvetica', 'B', 10)
        self.set_xy(14.361, 71)
        self.cell(0, 10, self.cliente, 0, 0, 'L')
        
        # Información adicional en páginas > 1
        if self.page_num > 1:
            self.set_font('Helvetica', '', 8)
            self.set_xy(14.361, 45)
            self.cell(0, 4, 'Centro de Atención Telefónica', 0, 0, 'L')
            self.set_xy(14.361, 52)
            self.cell(0, 4, 'Ciudad de México: 55 1226 2639', 0, 0, 'L')
            self.set_xy(14.361, 59)
            self.cell(0, 4, 'Resto del país: 800 021 2345', 0, 0, 'L')
        
        # Título de la tabla
        self.set_font('Helvetica', 'B', 11)
        self.set_xy(14.361, 80)
        self.cell(0, 11, 'DETALLE DE OPERACIONES', 0, 0, 'L')
        
        # Encabezados de la tabla
        self.set_font('Helvetica', 'B', 9)
        self.set_fill_color(255, 255, 255)  # Fondo blanco para encabezado
        self.set_draw_color(0, 0, 0)  # Líneas negras
        
        headers = ['FECHA', 'CONCEPTO', 'RETIROS', 'DEPOSITOS', 'SALDO']
        x_positions = [self.x_fecha, self.x_concepto, self.x_retiros, self.x_depositos, self.x_saldo]
        widths = [self.w_fecha, self.w_concepto, self.w_retiros, self.w_depositos, self.w_saldo]
        
        y_header = 92.448
        
        # Dibujar encabezados
        for i, (header, x_pos, width) in enumerate(zip(headers, x_positions, widths)):
            self.set_xy(x_pos, y_header)
            self.cell(width, self.altura_fila, header, 0, 0, 'C', True)
            
            # Líneas verticales (excepto después de la última columna)
            if i < len(headers) - 1:
                x_line = x_pos + width
                self.line(x_line, y_header, x_line, y_header + self.altura_fila)
        
        # Línea horizontal debajo del encabezado
        self.line(self.x_fecha, y_header + self.altura_fila, 
                 self.x_saldo + self.w_saldo, y_header + self.altura_fila)
        
        # Resetear contador de filas (encabezado cuenta como 1)
        self.filas_en_pagina = 1
    
    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', '', 8)
        self.set_x(14.361)
        self.cell(0, 10, '000191.B41EJDA029.OD.0121.01', 0, 0, 'L')
    
    def add_fila_individual(self, fecha, concepto, retiros, depositos, saldo):
        """
        Agrega UNA SOLA FILA con formato exacto: filas alternadas y líneas verticales
        """
        # Verificar si necesitamos nueva página
        if self.filas_en_pagina >= self.max_filas_por_pagina:
            self.add_page()
            self.page_num += 1
            self.filas_en_pagina = 1  # Reset después del encabezado
        
        # Alternar colores de fondo basado en fila global
        if self.fila_global % 2 == 0:
            self.set_fill_color(255, 255, 255)  # Blanco
        else:
            self.set_fill_color(191, 191, 191)  # Gris #bfbfbf
        
        self.set_draw_color(0, 0, 0)  # Líneas negras
        
        # Preparar valores para mostrar
        valores = [
            fecha if fecha else '',
            concepto if concepto else '',
            f'{retiros:,.2f}' if retiros is not None and retiros != 0 else '',
            f'{depositos:,.2f}' if depositos is not None and depositos != 0 else '',
            f'{saldo:,.2f}' if saldo is not None and saldo != 0 else ''
        ]
        
        # Configurar fuente para el contenido
        self.set_font('Helvetica', '', 9)
        
        # Calcular posición Y de la fila actual
        y_actual = self.y_inicial_datos + (self.filas_en_pagina - 1) * self.altura_fila
        
        # Posiciones y anchos
        x_positions = [self.x_fecha, self.x_concepto, self.x_retiros, self.x_depositos, self.x_saldo]
        widths = [self.w_fecha, self.w_concepto, self.w_retiros, self.w_depositos, self.w_saldo]
        aligns = ['C', 'L', 'R', 'R', 'R']
        
        # Agregar cada celda con fondo alternado
        for i, (valor, x_pos, width, align) in enumerate(zip(valores, x_positions, widths, aligns)):
            self.set_xy(x_pos, y_actual)
            self.cell(width, self.altura_fila, valor, 0, 0, align, True)
            
            # Líneas verticales (excepto después de la última columna)
            if i < len(valores) - 1:
                x_line = x_pos + width
                self.line(x_line, y_actual, x_line, y_actual + self.altura_fila)
        
        self.filas_en_pagina += 1
        self.fila_global += 1  # Incrementar contador global para alternado

# Streamlit App
st.set_page_config(
    page_title="Conversor Banamex Excel → PDF", 
    layout="wide", 
    page_icon="🏦"
)

st.title("🏦 Conversor Estado de Cuenta Banamex")
st.markdown("**Excel → PDF con formato EXACTO (187.33mm x 279.4mm, Helvetica, filas alternadas)**")
st.markdown("---")

# Sidebar con información
with st.sidebar:
    st.header("📋 Información del Cliente")
    cliente = st.text_input("Nombre del Cliente", "PATRICIA IÑIGUEZ FLORES")
    numero_cliente = st.text_input("Número de Cliente", "61900627")
    periodo = st.text_input("Período", "21 DE ENERO DE 2025")
    
    st.markdown("---")
    st.markdown("### 📖 Especificaciones Exactas")
    st.markdown("""
    ✅ **Tamaño:** 187.33mm x 279.4mm (531pt x 792pt)  
    ✅ **Fuente:** Helvetica tamaño 9  
    ✅ **Filas alternadas:** Blanca/gris #bfbfbf (global)  
    ✅ **Máximo:** 52 filas por página  
    ✅ **Líneas verticales negras** entre columnas  
    ✅ **Posiciones exactas en puntos (pts)**  
    ✅ **Sin empalme de filas**
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
        st.header("👀 Vista Previa (Formato Exacto)")
        st.dataframe(
            df_filas.head(20),
            use_container_width=True,
            column_config={
                "RETIROS": st.column_config.NumberColumn(format="$%.2f"),
                "DEPOSITOS": st.column_config.NumberColumn(format="$%.2f"),
                "SALDO": st.column_config.NumberColumn(format="$%.2f")
            }
        )
        
        if len(df_filas) > 20:
            st.info(f"Mostrando las primeras 20 de {len(df_filas)} filas")
        
        # Botón para generar PDF
        st.markdown("---")
        if st.button("🔄 Generar PDF EXACTO (187.33mm x 279.4mm)", type="primary", use_container_width=True):
            with st.spinner("Generando PDF con formato EXACTO de Banamex..."):
                # Crear PDF
                pdf = BanamexEstadoCuentaPDFExacto(
                    cliente=cliente,
                    numero_cliente=numero_cliente,
                    periodo=periodo
                )
                
                pdf.add_page()
                
                # Agregar CADA FILA INDIVIDUAL con formato exacto
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
                
                st.success("✅ PDF generado con formato EXACTO!")
                
                # Botón de descarga
                st.download_button(
                    label="📥 Descargar PDF EXACTO (187.33mm x 279.4mm)",
                    data=pdf_bytes,
                    file_name=f"estado_cuenta_exacto_{numero_cliente}_{datetime.now().strftime('%Y%m%d')}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
    
    except Exception as e:
        st.error(f"❌ Error al procesar el archivo: {str(e)}")
        st.info("Verifica que el archivo sea un Excel válido exportado del PDF de Banamex")

else:
    st.info("👆 Sube un archivo Excel para comenzar")
    
    # Mostrar ejemplo de estructura esperada
    with st.expander("📋 Formato EXACTO implementado"):
        st.markdown("""
        **✅ Características implementadas:**
        
        - **Tamaño exacto:** 187.33mm x 279.4mm (531pt x 792pt)
        - **Fuente:** Helvetica tamaño 9
        - **Filas alternadas:** Blanca y gris (#bfbfbf) - alternado global
        - **Máximo 52 filas** por página (incluyendo encabezado)
        - **Líneas verticales negras** entre columnas (excepto la última)
        - **Línea horizontal negra** debajo del encabezado
        - **Posiciones exactas en puntos (pts)** según el original
        - **Sin empalme de filas** - altura calculada exactamente
        - **Cada fila del Excel = Una fila en el PDF**
        
        **El resultado será IDÉNTICO al PDF original de Banamex** 🎯
        """)
