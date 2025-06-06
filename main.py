import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
from datetime import datetime
import numpy as np

MM_TO_PT = 2.83465
PAGE_W_PT = 187.33 * MM_TO_PT
PAGE_H_PT = 279.40 * MM_TO_PT

X_COLS_PT = [m * MM_TO_PT for m in (5.07, 20.47, 105.12, 131.46, 153.27)]
X_BAND_L_PT = 4.97 * MM_TO_PT
X_BAND_R_PT = (187.33 - 18.42) * MM_TO_PT
COL_W_PT = [X_COLS_PT[i+1] - X_COLS_PT[i] for i in range(4)]
COL_W_PT.append(X_BAND_R_PT - X_COLS_PT[-1])

Y_DATA_1_PT = 104.73901
Y_HEADER_PT = 92.448
BOTTOM_MG_PT = 18.16 * MM_TO_PT
ROWS_PAGE = 51
ROW_H_PT = (PAGE_H_PT - BOTTOM_MG_PT - Y_DATA_1_PT) / (ROWS_PAGE - 1)

def limpiar_nan(val):
    """Devuelve '' si el valor es None, NaN, 'nan', 'None', 'null', o string vacío."""
    if val is None:
        return ''
    if isinstance(val, float) and np.isnan(val):
        return ''
    sval = str(val).strip()
    if sval.lower() in ['nan', 'none', 'null', '']:
        return ''
    return sval

def monto_str(val):
    """Formatea el monto solo si es número válido, si no, regresa vacío."""
    limpio = limpiar_nan(val)
    if limpio == '':
        return ''
    try:
        return f'{float(limpio):,.2f}'
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
            if i < 4:
                self.line(X_COLS_PT[i+1], Y_HEADER_PT,
                          X_COLS_PT[i+1], Y_HEADER_PT + ROW_H_PT)
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

        vals = [
            limpiar_nan(fecha),
            limpiar_nan(concepto),
            monto_str(retiros),
            monto_str(depositos),
            monto_str(saldo)
        ]
        aligns = ['C', 'L', 'R', 'R', 'R']
        y = Y_DATA_1_PT + self.rows_in_page * ROW_H_PT
        self.set_font('Helvetica', '', 9)
        for i, val in enumerate(vals):
            # FILTRO FINAL: elimina cualquier 'nan' que se haya colado
            if isinstance(val, str):
                val = val.replace('nan', '').replace('NaN', '').replace('None', '').replace('null', '')
            self.set_xy(X_COLS_PT[i], y)
            self.cell(COL_W_PT[i], ROW_H_PT, val, 0, 0, aligns[i], True)
            if i < 4:
                self.line(X_COLS_PT[i+1], y, X_COLS_PT[i+1], y+ROW_H_PT)
        self.rows_in_page += 1
        self.row_global += 1

st.set_page_config(page_title='Banamex Excel → PDF', layout='wide', page_icon='🏦')
st.title('🏦 Conversor Estado de Cuenta Banamex – SIN NAN')

with st.sidebar:
    st.header('Datos del cliente')
    cliente = st.text_input('Nombre', 'PATRICIA IÑIGUEZ FLORES')
    numcte = st.text_input('Número de Cliente', '61900627')
    periodo = st.text_input('Período', '21 DE ENERO DE 2025')
    st.markdown('''
* **Ancho x Alto página:** 187.33 mm × 279.4 mm
* **Fuente:** Helvetica 9 pt
* **FILTRO ANTI-NAN:** ✅ ACTIVO
* **Alternado global blanco / gris #bfbfbf**
* **Líneas negras en columnas (sin la última)**
''')

excel_file = st.file_uploader('Sube tu archivo Excel', type=['xlsx', 'xls'])

if excel_file:
    df_raw = pd.read_excel(excel_file)
    df = parse_excel(df_raw)
    st.success(f'✅ Archivo leído. Filas: {len(df)} - FILTRO ANTI-NAN APLICADO')
    st.dataframe(df.head(15), use_container_width=True)

    if st.button('🚀 Generar PDF SIN NAN'):
        with st.spinner('Generando PDF con FILTRO ANTI-NAN...'):
            pdf = BanamexPDF(cliente, numcte, periodo)
            pdf.add_page()
            for _, r in df.iterrows():
                pdf.add_row(r['FECHA'], r['CONCEPTO'], r['RETIROS'], r['DEPOSITOS'], r['SALDO'])
            buf = io.BytesIO()
            pdf.output(buf)
            st.success('✅ PDF generado SIN NAN!')
            st.download_button(
                '📥 Descargar PDF LIMPIO',
                data=buf.getvalue(),
                file_name=f'Banamex_LIMPIO_{numcte}_{datetime.now():%Y%m%d}.pdf',
                mime='application/pdf'
            )
else:
    st.info('👉 Sube el Excel para convertirlo')
