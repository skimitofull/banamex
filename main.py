import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
from datetime import datetime

MM_TO_PT = 2.83465
PAGE_W_PT = 187.33 * MM_TO_PT
PAGE_H_PT = 279.40 * MM_TO_PT

# ─── POSICIONES HORIZONTALES EXACTAS ──────────────────────────────────────
X_COLS_PT = [m * MM_TO_PT for m in (5.07, 20.47, 105.12, 131.46, 153.27)]
X_BAND_L_PT = 4.97 * MM_TO_PT
X_BAND_R_PT = (187.33 - 18.42) * MM_TO_PT
COL_W_PT = [X_COLS_PT[i+1] - X_COLS_PT[i] for i in range(4)]
COL_W_PT.append(X_BAND_R_PT - X_COLS_PT[-1])

# ─── POSICIONES VERTICALES ────────────────────────────────────────────────
Y_DATA_1_PT = 104.73901
Y_HEADER_PT = 92.448
BOTTOM_MG_PT = 18.16 * MM_TO_PT
ROWS_PAGE = 51
ROW_H_PT = (PAGE_H_PT - BOTTOM_MG_PT - Y_DATA_1_PT) / (ROWS_PAGE - 1)

# ─── FUNCIÓN ULTRA AGRESIVA PARA BANEAR "nan" ─────────────────────────────
def banear_nan_completamente(value):
    """BANEA COMPLETAMENTE cualquier rastro de 'nan' o valores vacíos"""
    # Si es None o NaN, devolver cadena vacía
    if value is None or pd.isna(value):
        return ''
    
    # Convertir a string y limpiar
    str_value = str(value).strip()
    
    # Lista de valores prohibidos (en minúsculas para comparación)
    valores_prohibidos = ['nan', 'none', 'null', 'na', '<na>', 'n/a']
    
    # Si el valor está en la lista prohibida, devolver vacío
    if str_value.lower() in valores_prohibidos:
        return ''
    
    # Si el string está vacío después de strip, devolver vacío
    if not str_value:
        return ''
    
    return str_value

def limpiar_numero(value):
    """Limpia valores numéricos y devuelve None si no es válido"""
    if value is None or pd.isna(value):
        return None
    
    try:
        # Intentar convertir a float
        num_val = float(value)
        # Si es 0, devolver None (para que aparezca vacío)
        return num_val if num_val != 0 else None
    except (ValueError, TypeError):
        return None

# ─── PARSER EXCEL CON LIMPIEZA ULTRA AGRESIVA ────────────────────────────
def parse_excel(df):
    df = df.copy()
    df.columns = ['FECHA', 'CONCEPTO', 'RETIROS', 'DEPOSITOS', 'SALDO']
    parsed = []

    for _, r in df.iterrows():
        # Aplicar limpieza ultra agresiva a cada campo
        fecha_limpia = banear_nan_completamente(r['FECHA'])
        concepto_limpio = banear_nan_completamente(r['CONCEPTO'])
        retiros_limpio = limpiar_numero(r['RETIROS'])
        depositos_limpio = limpiar_numero(r['DEPOSITOS'])
        saldo_limpio = limpiar_numero(r['SALDO'])

        parsed.append({
            'FECHA': fecha_limpia,
            'CONCEPTO': concepto_limpio,
            'RETIROS': retiros_limpio,
            'DEPOSITOS': depositos_limpio,
            'SALDO': saldo_limpio,
        })

    return pd.DataFrame(parsed)

# ─── CLASE PDF CON FILTRO ANTI-NAN ───────────────────────────────────────
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

        # Encabezados de la tabla
        self.set_font('Helvetica', 'B', 9)
        self.set_y(Y_HEADER_PT)
        headers = ['FECHA','CONCEPTO','RETIROS','DEPÓSITOS','SALDO']
        for i, h in enumerate(headers):
            self.set_x(X_COLS_PT[i])
            self.cell(COL_W_PT[i], ROW_H_PT, h, 0, 0, 'C', True)
            if i < 4:
                self.line(X_COLS_PT[i+1], Y_HEADER_PT,
                          X_COLS_PT[i+1], Y_HEADER_PT + ROW_H_PT)

        # línea horizontal
        self.line(X_BAND_L_PT, Y_HEADER_PT+ROW_H_PT,
                  X_BAND_R_PT, Y_HEADER_PT+ROW_H_PT)

        self.rows_in_page = 0
        self.set_y(Y_DATA_1_PT)

    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', '', 8)
        self.set_x(X_BAND_L_PT)
        self.cell(0, 10, '000191.B41EJDA029.OD.0121.01', 0, 0, 'L')

    def filtro_anti_nan(self, texto):
        """FILTRO FINAL: Si detecta 'nan' en el texto, lo reemplaza por vacío"""
        if not texto:
            return ''
        
        texto_str = str(texto).strip()
        
        # Lista de palabras prohibidas
        prohibidas = ['nan', 'none', 'null', 'na', '<na>', 'n/a']
        
        # Si el texto completo es una palabra prohibida, devolver vacío
        if texto_str.lower() in prohibidas:
            return ''
        
        # Reemplazar cualquier ocurrencia de palabras prohibidas
        for prohibida in prohibidas:
            texto_str = texto_str.replace(prohibida, '')
            texto_str = texto_str.replace(prohibida.upper(), '')
            texto_str = texto_str.replace(prohibida.capitalize(), '')
        
        return texto_str.strip()

    def add_row(self, fecha, concepto, retiros, depositos, saldo):
        if self.rows_in_page >= ROWS_PAGE:
            self.add_page()

        # alternancia blanco / gris (global)
        if self.row_global % 2 == 0:
            self.set_fill_color(255, 255, 255)
        else:
            self.set_fill_color(191, 191, 191)

        # Preparar valores con FILTRO ANTI-NAN FINAL
        vals = [
            self.filtro_anti_nan(fecha),
            self.filtro_anti_nan(concepto),
            self.filtro_anti_nan(f'{retiros:,.2f}' if retiros is not None else ''),
            self.filtro_anti_nan(f'{depositos:,.2f}' if depositos is not None else ''),
            self.filtro_anti_nan(f'{saldo:,.2f}' if saldo is not None else '')
        ]

        aligns = ['C', 'L', 'R', 'R', 'R']
        y = Y_DATA_1_PT + self.rows_in_page * ROW_H_PT
        self.set_font('Helvetica', '', 9)

        for i, val in enumerate(vals):
            self.set_xy(X_COLS_PT[i], y)
            # APLICAR FILTRO ANTI-NAN UNA VEZ MÁS ANTES DE ESCRIBIR
            texto_final = self.filtro_anti_nan(val)
            self.cell(COL_W_PT[i], ROW_H_PT, texto_final, 0, 0, aligns[i], True)
            if i < 4:
                self.line(X_COLS_PT[i+1], y, X_COLS_PT[i+1], y+ROW_H_PT)

        self.rows_in_page += 1
        self.row_global += 1

# ─── STREAMLIT ────────────────────────────────────────────────────────────
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

            # Procesar filas con filtro anti-nan
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
