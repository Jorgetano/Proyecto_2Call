import streamlit as st
from docx import Document
from datetime import datetime
from io import BytesIO

# Función para reemplazar los marcadores de posición en el documento
def replace_placeholder(doc, placeholder, replacement):
    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            for run in paragraph.runs:
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, replacement)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_placeholder(cell, placeholder, replacement)

# Función para generar el documento de Word
def create_document(data, template_path, output_path):
    doc = Document(template_path)
    for placeholder, replacement in data.items():
        replace_placeholder(doc, placeholder, replacement)
    
    doc.save(output_path)
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# Título de la aplicación
st.title("Formulario Único Desconocimiento de Transacciones")

# Ingresar datos personales
st.header("I. Antecedentes personales")
nombre = st.text_input("Nombre Tarjetahabiente (Cardholder name)")
tc = st.text_input("N° tarjeta (Cardholder number) 4 últimos dígitos")
tt_tarjeta = st.selectbox("Tipo de tarjeta (Card type)", ["Titular", "Adicional", "Ambas"])
direccion = st.text_input("Dirección (Address)")
correo = st.text_input("E-Mail")
telefono = st.text_input("Celular (Cel Phone)")
rut = st.text_input("RUT")

# Fecha de reclamo
fecha_actual = datetime.now().strftime("%d/%m/%Y")

# Detalle de transacciones reclamadas
st.header("III. Detalle Transacciones Reclamadas (Transaction Details)")
moneda = st.selectbox("Moneda para la suma total", ["Pesos", "Dólares"])
if moneda == "Dólares":
    monto_total_label = "USD"
else:
    monto_total_label = "$"
    
num_transacciones = st.number_input("Cantidad de transacciones reclamadas", min_value=1, max_value=15, step=1)
transacciones = []
monto_total = 0.0  # Variable para sumar los montos


for i in range(num_transacciones):
    fecha = st.text_input(f"Fecha (dd/mm/aa) - Transacción {i+1}")
    nombre_comercio = st.text_input(f"Nombre del Comercio - Transacción {i+1}")
    monto = st.number_input(f"Monto - Transacción {i+1}", min_value=0.0, format="%.2f")
    transacciones.append((fecha, nombre_comercio, monto))
    monto_total += monto  # Sumar el monto a la variable de suma total

# Observaciones
st.header("IV. Observaciones")
observaciones = st.text_area("Observaciones")

# Botón para generar el documento
if st.button("Generar Documento"):
    data = {
        "{{Nombre}}": nombre,
        "{{Tc}}": tc,
        "Titular": tt_tarjeta,
        "{{Direccin}}": direccion,
        "Direccion": direccion,
        "Fecha actual": fecha_actual,
        "{{Correo}}": correo,
        "{{Telefono}}": telefono,
        "{{Rut}}": rut,
        "Numero TRX": str(num_transacciones),
        "{{Run}}": f"{monto_total_label} {monto_total:.2f}",  # Concatenar la moneda y el monto total
        "{{Observaciones}}": observaciones
    }   
    for a, (fecha, nombre_comercio, monto) in enumerate(transacciones, start=1):
        data[f"Fecha{a}"] = fecha
        data[f"NombreComercio{a}"] = nombre_comercio
        data[f"Monto{a}"] = f"{monto_total_label} {monto:.2f}"

    # Ruta de salida del documento
    output_path = "Formulario_unico_desconocimiento_ult_4.docx"

    # Ruta del template
    template_path = "template.docx"

    # Crear el documento y obtener los bytes
    doc_file = create_document(data, template_path, output_path)
    
    # Descargar el documento generado
    st.download_button(label="Descargar Documento", data=doc_file, file_name=output_path)
