import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
from datetime import datetime
import numpy as np

# Constantes de conversión
MM_TO_PT = 2.83465
PAGE_W_PT = 187.33 * MM_TO_PT
PAGE_H_PT = 279.40 * MM_TO_PT

# Líneas verticales (en mm desde el borde izquierdo)
X_LINE_MM = [20.00, 91.13, 115.78, 142.45]
X_LINE_PT = [x * MM_TO_PT for x in X_LINE_MM]

# Columnas (en puntos)
X_COLS_PT = [
    5.11 * MM_TO_PT,   # FECHA - alineada a 5.11 mm
    X_LINE_PT[0],      # CONCEPTO
    X_LINE_PT[1],      # RETIROS
    X_LINE_PT[2],      # DEPOSITOS
    X_LINE_PT[3]       # SALDO
]

# Ancho derecho para cálculo del ancho de la última columna
X_BAND_R_PT = (187.33 - 18.42) * MM_TO_PT
COL_W_PT = [
    X_COLS_PT[1] - X_COLS_PT[0],  # Ancho de "FECHA"
    X_COLS_PT[2] - X_COLS_PT[1],  # Ancho de "CONCEPTO"
    X_COLS_PT[3] - X_COLS_PT[2],  # Ancho de "RETIROS"
    X_COLS_PT[4] - X_COLS_PT[3],  # Ancho de "DEPÓSITOS"
    X_BAND_R_PT - X_COLS_PT[4]    # Ancho de "SALDO"
]

# Posiciones verticales
Y_HEADER_PT = 50.0 * MM_TO_PT     # Altura del encabezado en páginas siguientes
Y_DATA_1_PT = 106.73901            # Ajuste fino: alinea fecha con franjas grises
BOTTOM_MG_PT = 18.16 * MM_TO_PT   # Margen inferior

# Altura de fila base
ROW_H_PT = 12  # Ajuste fino para mejorar espaciado visual

# Pie de página
FOOTER_TEXT = "Centro de Atención Telefónica Ciudad de México: 55 1226 2639 Resto del país: 800 021 2345"

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

def split_text(pdf, text, max_width, font_size=9):
    pdf.set_font('Helvetica', '', font_size)
    words = text.split(' ')
    lines = []
    current_line = ''

    for word in words:
        test_line = current_line + ' ' + word.strip() if current_line else word.strip()
        if pdf.get_string_width(test_line) <= max_width:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line)
            current_line = word.strip()
    if current_line:
        lines.append(current_line)
    return lines

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
        if self.is_first_page:
            self.is_first_page = False
            self.rows_in_page = 0
            self.set_y(Y_DATA_1_PT)
        else:
            self.page_no_global += 1
            self.set_font('Helvetica', 'B', 9)
            self.set_y(Y_HEADER_PT)
            headers = ['FECHA','CONCEPTO','RETIROS','DEPÓSITOS','SALDO']
            for i, h in enumerate(headers):
                self.set_x(X_COLS_PT[i])
                self.cell(COL_W_PT[i], ROW_H_PT, h, 0, 0, 'C', True)
            # Encabezado adicional
            self.set_font('Helvetica', '', 8)
            self.set_xy(0, Y_HEADER_PT - 12)
            self.cell(PAGE_W_PT, 10, f'ESTADO DE CUENTA AL {self.periodo}', 0, 0, 'C')
            self.set_xy(0, Y_HEADER_PT - 6)
            self.cell(PAGE_W_PT, 10, f'CLIENTE: {self.num_cte}     Página: {self.page_no_global + 1} de ???     {self.cliente}', 0, 0, 'C')
            self.rows_in_page = 0
            self.set_y(Y_DATA_1_PT)

    def footer(self):
        self.set_y(-20)
        self.set_font('Helvetica', '', 7)
        self.cell(0, 10, FOOTER_TEXT, 0, 0, 'C')

    def add_row(self, fecha, concepto, retiros, depositos, saldo):
        if self.rows_in_page >= 51:
            self.add_page()

        # Convertir valores a cadenas limpias
        fecha_str = clean_cell(fecha)
        concepto_str = clean_cell(concepto)
        retiros_str = monto_cell(retiros)
        depositos_str = monto_cell(depositos)
        saldo_str = monto_cell(saldo)

        # Dividir concepto en líneas
        concept_lines = split_text(self, concepto_str, COL_W_PT[1] - 6)
        row_height = ROW_H_PT * len(concept_lines)

        y = Y_DATA_1_PT + self.rows_in_page * ROW_H_PT

        # Dibuja franja gris/blanco
        if self.row_global % 2 == 0:
            self.set_fill_color(255, 255, 255)
        else:
            self.set_fill_color(191, 191, 191)
        self.rect(X_COLS_PT[0], y, X_BAND_R_PT - X_COLS_PT[0], row_height, style='F')

        # Dibuja líneas encima de la celda
        self.set_line_width(1.0)
        self.set_draw_color(0)
        for x in X_LINE_PT:
            self.line(x, y, x, y + row_height)

        # Escribe los valores
        self.set_font('Helvetica', '', 9)

        # FECHA - centrado horizontal y vertical
        self.set_xy(X_COLS_PT[0], y + (row_height / 3) - 3)
        self.cell(COL_W_PT[0], ROW_H_PT, fecha_str, 0, 0, 'C', False)

        # CONCEPTO (con múltiples líneas) - alineado a la izquierda y centrado verticalmente
        line_y_start = y + ((row_height - (len(concept_lines) * ROW_H_PT)) / 2)
        for i, line in enumerate(concept_lines):
            self.set_xy(X_COLS_PT[1], line_y_start + i * ROW_H_PT + 3)
            self.cell(COL_W_PT[1], ROW_H_PT, line, 0, 0, 'L', False)

        # RETIROS - alineado a la derecha y centrado verticalmente
        self.set_xy(X_COLS_PT[2], y + (row_height / 2) - 5)
        self.cell(COL_W_PT[2], ROW_H_PT, retiros_str, 0, 0, 'R', False)

        # DEPÓSITOS - alineado a la derecha y centrado verticalmente
        self.set_xy(X_COLS_PT[3], y + (row_height / 2) - 5)
        self.cell(COL_W_PT[3], ROW_H_PT, depositos_str, 0, 0, 'R', False)

        # SALDO - alineado a la derecha y centrado verticalmente
        self.set_xy(X_COLS_PT[4], y + (row_height / 2) - 5)
        self.cell(COL_W_PT[4], ROW_H_PT, saldo_str, 0, 0, 'R', False)

        # Avanzar filas usadas
        self.rows_in_page += len(concept_lines)
        self.row_global += 1


# INTERFAZ STREAMLIT
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
