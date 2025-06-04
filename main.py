import streamlit as st
import pdfplumber
import pandas as pd
import io

st.set_page_config(page_title="Conversor PDF a Excel - Estado de Cuenta", layout="centered")

st.title("🧾 Conversión de Estado de Cuenta PDF a Excel")

uploaded_file = st.file_uploader("Sube tu archivo PDF del estado de cuenta Banamex", type=["pdf"])

if uploaded_file:
    st.success("Archivo cargado correctamente. Procesando...")
    with pdfplumber.open(uploaded_file) as pdf:
        all_text = ""
        for page in pdf.pages:
            all_text += page.extract_text() + "\n"

    # Procesamiento básico por líneas
    lines = all_text.split("\n")
    data = []
    current_date = ""
    current_concept = []

    for line in lines:
        if line.strip()[:6].upper().count(" ") == 1 and line[:2].isdigit():
            if current_date and current_concept:
                data.append(current_concept)
            current_date = line.strip()[:6]
            current_concept = [current_date + " " + line.strip()[6:]]
        elif any(char.isdigit() for char in line) and line.strip().count(".") >= 2:
            parts = line.strip().split()
            try:
                retiro = float(parts[-2].replace(",", ""))
                saldo = float(parts[-1].replace(",", ""))
                deposito = ""
                concepto = " ".join(current_concept)
                data.append([current_date, concepto.strip(), retiro, deposito, saldo])
                current_concept = []
            except:
                continue
        else:
            if current_concept is not None:
                current_concept.append(line.strip())

    # Solo conservar los registros válidos
    data_valid = [item for item in data if isinstance(item, list) and len(item) == 5]

    df = pd.DataFrame(data_valid, columns=["Fecha", "Concepto", "Retiros", "Depositos", "Saldo"])

    # Descargar como Excel
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, engine='openpyxl')
    buffer.seek(0)

    st.success("Conversión completada. Descarga tu archivo:")
    st.download_button("📥 Descargar Excel", buffer, file_name="estado_cuenta_convertido.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")