import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
from datetime import datetime
import numpy as np

MM_TO_PT = 2.83465
PAGE_W_PT = 187.33 * MM_TO_PT
PAGE_H_PT = 279.40 * MM_TO_PT

# Medidas de las líneas verticales (en puntos)
LINE_WIDTH_PT = 0.75
LINE_COLOR = 0  # Negro (0 para escala de grises)

# Posiciones X de las líneas (en puntos)
X_LINE_POS_PT = [
    20.11 * MM_TO_PT,   # Primera línea
    91.12 * MM_TO_PT,   # Segunda línea
    115.68 * MM_TO_PT,  # Tercera línea
    142.35 * MM_TO_PT   # Cuarta línea
]

# Posición Y de inicio de las líneas (en puntos)
Y_LINE_START_PT = 31.77 * MM_TO_PT

# Longitud de las líneas (en puntos)
LINE_LENGTH_PT = 228.88 * MM_TO_PT

# Posiciones X de las columnas (en puntos)
X_COLS_PT = [5.07 * MM_TO_PT, 20.47 * MM_TO_PT, 105.12 * MM_TO_PT, 131.46 * MM_TO_PT, 153.27 * MM_TO_PT]
X_BAND_L_PT = 4.97 * MM_TO_PT
X_BAND_R_PT = (187.33 - 18.42) * MM_TO_PT
COL_W_PT = [X_COLS_PT[i+1] - X_COLS_PT[i] for i in range(4)]
COL_W_PT.append(X_BAND_R_PT - X_COLS_PT[-1])

Y_DATA_1_PT = 104.73901
Y_HEADER_PT = 92.448
BOTTOM_MG_PT = 18.16 * MM_TO_PT
ROWS_PAGE = 51
ROW_H_PT = (PAGE_H_PT - BOTTOM_MG_PT - Y_DATA_1_PT) / (ROWS_PAGE - 1)

def clean_cell(val):
    """Limpia celdas de texto (FECHA, CONCEPTO)"""
    if val is None:
        return ''
    if isinstance(val, float) and np.isnan(val):
        return ''
    sval = str(val).strip()
    if sval.lower() in ['nan', 'none', 'null', '']:
        return ''
    return sval

def monto_cell(val):
    """Limpia y formatea celdas de montos (RETIROS, DEPOSITOS, SALDO)"""
    if val is None:
        return ''
    if isinstance(val, float) and np.isnan(val):
        return ''
    if isinstance(val, str) and val.strip().lower() in ['nan', 'none', 'null', '']:
        return ''
    try:
        fval = float(val)
        if np.isnan(fval):
            return ''
        return f'{fval:,.2f}'
    except:
        return ''

def parse_excel(df):
    """Parsea el Excel y toma solo las primeras 5 columnas"""
    df = df.iloc[:, :5].copy()
    df.columns = ['FECHA', 'CONCEPTO', 'RETIROS', 'DEPOSITOS', 'SALDO']
    parsed = []
    for _, r in df.iterrows():
        parsed.append({
            'FECHA': r['FECHA'],
            'CONCEPTO': r['CONCEPTO'],
            'RETIROS': r['RETIROS'],
            'DEPOSITOS': r['DEPOSITOS'],
            'SALDO': r['SALDO'],
        })
    return pd.DataFrame(parsed)

class BanamexPDF(FPDF):
    def __init__(self, cliente, num_cte, periodo):
        super().__init__(unit='pt', format=(PAGE_W_PT, PAGE_H_PT))
        self.cliente = cliente
        self.num_cte = num_cte
        self.periodo = periodo
        self.set_auto_page_break(False)
        self.page_no_global = 0
        self.row_global = 0

    def header(self):
        self.page_no_global += 1
        self.set_font('Helvetica', 'B', 14)
        self.cell(0, 18, f'ESTADO DE CUENTA AL {self.periodo.upper()}', 0, 1, 'C')
        self.set_font('Helvetica', 'B', 10)
        self.set_x(X_BAND_L_PT)
        self.cell(200, 10, 'CLIENTE:', 0, 0)
        self.set_x(PAGE_W_PT-120)
        self.cell(100, 10, f'Página: {self.page_no_global}', 0, 1, 'R')
        self.set_font('Helvetica', '', 10)
        self.set_x(X_BAND_L_PT)
        self.cell(0, 10, self.num_cte, 0, 1)
        self.set_font('Helvetica', 'B', 10)
        self.set_x(X_BAND_L_PT)
        self.cell(0, 10, self.cliente, 0, 1)
        self.set_font('Helvetica', 'B', 9)
        self.set_y(Y_HEADER_PT)
        headers = ['FECHA','CONCEPTO','RETIROS','DEPÓSITOS','SALDO']
        for i, h in enumerate(headers):
            self.set_x(X_COLS_PT[i])
            self.cell(COL_W_PT[i], ROW_H_PT, h, 0, 0, 'C', True)
            # Dibujar las líneas verticales
        self.set_line_width(LINE_WIDTH_PT)
        self.set_draw_color(LINE_COLOR)
        for x in X_LINE_POS_PT:
            self.line(x, Y_LINE_START_PT, x, Y_LINE_START_PT + LINE_LENGTH_PT)
        self.line(X_BAND_L_PT, Y_HEADER_PT+ROW_H_PT,
                  X_BAND_R_PT, Y_HEADER_PT+ROW_H_PT)
        self.rows_in_page = 0
        self.set_y(Y_DATA_1_PT)

    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', '', 8)
        self.set_x(X_BAND_L_PT)
        self.cell(0, 10, '000191.B41EJDA029.OD.0121.01', 0, 0, 'L')

    def add_row(self, fecha, concepto, retiros, depositos, saldo):
        if self.rows_in_page >= ROWS_PAGE:
            self.add_page()
        if self.row_global % 2 == 0:
            self.set_fill_color(255, 255, 255)
        else:
            self.set_fill_color(191, 191, 191)

        # AQUÍ ESTÁ LA CLAVE: usar funciones diferentes para texto y montos
        vals = [
            clean_cell(fecha),      # Para FECHA (texto)
            clean_cell(concepto),   # Para CONCEPTO (texto)
            monto_cell(retiros),    # Para RETIROS (monto)
            monto_cell(depositos),  # Para DEPOSITOS (monto)
            monto_cell(saldo)       # Para SALDO (monto)
        ]
        
        aligns = ['C', 'L', 'R', 'R', 'R']
        y = Y_DATA_1_PT + self.rows_in_page * ROW_H_PT
        self.set_font('Helvetica', '', 9)
        
        for i, val in enumerate(vals):
            self.set_xy(X_COLS_PT[i], y)
            self.cell(COL_W_PT[i], ROW_H_PT, val, 0, 0, aligns[i], True)
            if i < 4:
                self.line(X_COLS_PT[i+1], y, X_COLS_PT[i+1], y+ROW_H_PT)
        
        self.rows_in_page += 1
        self.row_global += 1

# INTERFAZ DE STREAMLIT
st.set_page_config(page_title='Banamex Excel → PDF', layout='wide', page_icon='🏦')
st.title('🏦 Conversor Estado de Cuenta Banamex – SIN NAN ✅')

with st.sidebar:
    st.header('📋 Datos del cliente')
    cliente = st.text_input('Nombre del Cliente', 'PATRICIA IÑIGUEZ FLORES')
    numcte = st.text_input('Número de Cliente', '61900627')
    periodo = st.text_input('Período', '21 DE ENERO DE 2025')
    
    st.markdown('---')
    st.markdown('### ✅ Características')
    st.markdown('''
    * **Página:** 187.33 mm × 279.4 mm
    * **Fuente:** Helvetica 9 pt
    * **Filtro Anti-NAN:** ✅ ACTIVO
    * **Alternado:** Blanco / Gris #BFBFBF
    * **Líneas:** Negras entre columnas
    * **Formato:** Idéntico al original Banamex
    ''')

st.markdown('### 📁 Subir archivo Excel')
excel_file = st.file_uploader(
    'Selecciona tu archivo Excel con los movimientos bancarios', 
    type=['xlsx', 'xls'],
    help='El archivo debe tener las columnas: FECHA, CONCEPTO, RETIROS, DEPOSITOS, SALDO'
)

if excel_file:
    try:
        # Leer y procesar el archivo
        df_raw = pd.read_excel(excel_file)
        df = parse_excel(df_raw)
        
        st.success(f'✅ Archivo procesado correctamente: **{len(df)} filas**')
        
        # Mostrar preview
        st.markdown('### 👀 Vista previa de los datos')
        st.dataframe(df.head(15), use_container_width=True)
        
        # Estadísticas
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric('Total Filas', len(df))
        with col2:
            retiros_count = df['RETIROS'].notna().sum()
            st.metric('Retiros', retiros_count)
        with col3:
            depositos_count = df['DEPOSITOS'].notna().sum()
            st.metric('Depósitos', depositos_count)
        with col4:
            saldos_count = df['SALDO'].notna().sum()
            st.metric('Saldos', saldos_count)
        
        st.markdown('---')
        
        # Botón para generar PDF
        if st.button('🚀 Generar PDF Estado de Cuenta', type='primary', use_container_width=True):
            with st.spinner('Generando PDF con formato Banamex...'):
                try:
                    pdf = BanamexPDF(cliente, numcte, periodo)
                    pdf.add_page()
                    
                    # Agregar cada fila al PDF
                    for _, r in df.iterrows():
                        pdf.add_row(r['FECHA'], r['CONCEPTO'], r['RETIROS'], r['DEPOSITOS'], r['SALDO'])
                    
                    # Generar el archivo
                    buf = io.BytesIO()
                    pdf.output(buf)
                    
                    st.success('✅ PDF generado exitosamente!')
                    
                    # Botón de descarga
                    st.download_button(
                        label='📥 Descargar Estado de Cuenta PDF',
                        data=buf.getvalue(),
                        file_name=f'Banamex_{numcte}_{datetime.now():%Y%m%d_%H%M%S}.pdf',
                        mime='application/pdf',
                        use_container_width=True
                    )
                    
                except Exception as e:
                    st.error(f'❌ Error al generar el PDF: {str(e)}')
                    
    except Exception as e:
        st.error(f'❌ Error al procesar el archivo: {str(e)}')
        st.info('💡 Asegúrate de que el archivo Excel tenga el formato correcto.')

else:
    st.info('👆 Sube tu archivo Excel para comenzar')
    st.markdown('### 📋 Formato esperado del archivo Excel')
    st.markdown('''
    El archivo debe tener estas columnas en orden:
    1. **FECHA** - Fecha del movimiento (puede estar vacía en algunas filas)
    2. **CONCEPTO** - Descripción del movimiento
    3. **RETIROS** - Monto de retiros (puede estar vacío)
    4. **DEPOSITOS** - Monto de depósitos (puede estar vacío)
    5. **SALDO** - Saldo después del movimiento (puede estar vacío)
    ''')
