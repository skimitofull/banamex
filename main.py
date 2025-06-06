import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
from datetime import datetime
import numpy as np

MM_TO_PT = 2.83465
PAGE_W_PT = 187.33 * MM_TO_PT
PAGE_H_PT = 279.40 * MM_TO_PT

# Líneas verticales (en mm desde el borde izquierdo)
X_LINE_MM = [20.11, 91.12, 115.68, 142.35]
X_LINE_PT = [x * MM_TO_PT for x in X_LINE_MM]
Y_LINE_START_PT = 31.77 * MM_TO_PT
LINE_LENGTH_PT = 228.88 * MM_TO_PT
LINE_WIDTH_PT = 0.75

# Columnas (en puntos)
X_COLS_PT = [
    14.37,          # FECHA - 14.37 pt desde el borde izquierdo
    X_LINE_PT[0],   # CONCEPTO - alineada con la primera línea
    X_LINE_PT[1],   # RETIROS - alineada con la segunda línea
    X_LINE_PT[2],   # DEPOSITOS - alineada con la tercera línea
    X_LINE_PT[3]    # SALDO - alineada con la cuarta línea
]

X_BAND_R_PT = (187.33 - 18.42) * MM_TO_PT
COL_W_PT = [
    X_COLS_PT[1] - X_COLS_PT[0],
    X_COLS_PT[2] - X_COLS_PT[1],
    X_COLS_PT[3] - X_COLS_PT[2],
    X_COLS_PT[4] - X_COLS_PT[3],
    X_BAND_R_PT - X_COLS_PT[4]
]

Y_DATA_1_PT = 104.73901
Y_HEADER_PT = 92.448
BOTTOM_MG_PT = 18.16 * MM_TO_PT
ROWS_PAGE = 51
ROW_H_PT = (PAGE_H_PT - BOTTOM_MG_PT - Y_DATA_1_PT) / (ROWS_PAGE - 1)

def clean_cell(val):
    if val is None:
        return ''
    if isinstance(val, float) and np.isnan(val):
        return ''
    sval = str(val).strip()
    if sval.lower() in ['nan', 'none', 'null', '']:
        return ''
    return sval

def monto_cell(val):
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
        self.is_first_page = True

    def header(self):
        # No header en la primera hoja
        if self.is_first_page:
            self.is_first_page = False
            self.rows_in_page = 0
            self.set_y(Y_DATA_1_PT)
            return

        # Header normal en las siguientes hojas
        self.page_no_global += 1
        self.set_font('Helvetica', 'B', 9)
        self.set_y(Y_HEADER_PT)
        headers = ['FECHA','CONCEPTO','RETIROS','DEPÓSITOS','SALDO']
        for i, h in enumerate(headers):
            self.set_x(X_COLS_PT[i])
            self.cell(COL_W_PT[i], ROW_H_PT, h, 0, 0, 'C', True)
        self.rows_in_page = 0
        self.set_y(Y_DATA_1_PT)

    def footer(self):
        pass  # Sin footer

    def add_row(self, fecha, concepto, retiros, depositos, saldo):
        if self.rows_in_page >= ROWS_PAGE:
            self.add_page()
        if self.row_global % 2 == 0:
            self.set_fill_color(255, 255, 255)
        else:
            self.set_fill_color(191, 191, 191)
        vals = [
            clean_cell(fecha),
            clean_cell(concepto),
            monto_cell(retiros),
            monto_cell(depositos),
            monto_cell(saldo)
        ]
        aligns = ['C', 'L', 'R', 'R', 'R']
        y = Y_DATA_1_PT + self.rows_in_page * ROW_H_PT
        self.set_font('Helvetica', '', 9)
        for i, val in enumerate(vals):
            self.set_xy(X_COLS_PT[i], y)
            self.cell(COL_W_PT[i], ROW_H_PT, val, 0, 0, aligns[i], True)
        self.rows_in_page += 1
        self.row_global += 1

    def draw_vertical_lines(self):
        self.set_line_width(LINE_WIDTH_PT)
        self.set_draw_color(0)
        for x in X_LINE_PT:
            self.line(x, Y_LINE_START_PT, x, Y_LINE_START_PT + LINE_LENGTH_PT)

    def _endpage(self):
        # Sobrescribe el método para dibujar las líneas al final de cada página
        super()._endpage()
        self.draw_vertical_lines()

# INTERFAZ DE STREAMLIT
st.set_page_config(page_title='Banamex Excel → PDF', layout='wide', page_icon='🏦')
st.title('🏦 Conversor Estado de Cuenta Banamex – Formato Final')

with st.sidebar:
    st.header('📋 Datos del cliente')
    cliente = st.text_input('Nombre del Cliente', 'PATRICIA IÑIGUEZ FLORES')
    numcte = st.text_input('Número de Cliente', '61900627')
    periodo = st.text_input('Período', '21 DE ENERO DE 2025')

st.markdown('### 📁 Subir archivo Excel')
excel_file = st.file_uploader(
    'Selecciona tu archivo Excel con los movimientos bancarios',
    type=['xlsx', 'xls'],
    help='El archivo debe tener las columnas: FECHA, CONCEPTO, RETIROS, DEPOSITOS, SALDO'
)

if excel_file:
    try:
        df_raw = pd.read_excel(excel_file)
        df = parse_excel(df_raw)
        st.success(f'✅ Archivo procesado correctamente: **{len(df)} filas**')
        st.markdown('### 👀 Vista previa de los datos')
        st.dataframe(df.head(15), use_container_width=True)
        if st.button('🚀 Generar PDF Estado de Cuenta', type='primary', use_container_width=True):
            with st.spinner('Generando PDF con formato Banamex...'):
                try:
                    pdf = BanamexPDF(cliente, numcte, periodo)
                    pdf.add_page()
                    for _, r in df.iterrows():
                        pdf.add_row(r['FECHA'], r['CONCEPTO'], r['RETIROS'], r['DEPOSITOS'], r['SALDO'])
                    buf = io.BytesIO()
                    pdf.output(buf)
                    st.success('✅ PDF generado exitosamente!')
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
