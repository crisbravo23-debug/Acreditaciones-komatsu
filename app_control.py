
import streamlit as st
import pandas as pd
import os
import webbrowser
import base64
from datetime import datetime
import qrcode
from io import BytesIO
import zipfile
import tempfile

st.set_page_config(layout="wide")
# ==============================
# 🔥 LOGO BASE64
# ==============================
def cargar_logo(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()

logo_base64 = cargar_logo("komatsu.png")

# ==============================
# 🎨 ESTILOS
# ==============================
st.markdown(f"""
<style>

.stApp{{
    background:
    radial-gradient(circle at center,
    #0d2d63 0%,
    #071b3b 35%,
    #030b1c 100%);
}}

.card-pro{{
    background:
    linear-gradient(135deg,
        rgba(8,25,55,0.95),
        rgba(4,15,35,0.95)
    );
    border:2px solid #2490ff;
    border-radius:24px;
    padding:40px;
    margin-bottom:25px;

    box-shadow:
    0 0 25px rgba(36,144,255,.45),
    inset 0 0 40px rgba(36,144,255,.08);
}}

.card-title{{
    color:white;
    font-size:24px;
    font-weight:700;
    margin-bottom:25px;
}}

.nombre{{
    color:white;
    font-size:34px;
    font-weight:700;
    margin-bottom:15px;
}}

.info{{
    color:#d7e8ff;
    font-size:18px;
    margin-bottom:12px;
}}

.card-pro img{{
    border-radius:18px;
    border:2px solid #2490ff;
    box-shadow:0 0 20px rgba(36,144,255,.45);
}}

</style>

<img src="data:image/png;base64,{logo_base64}" style="position:fixed; top:80px; right:60px; width:200px; opacity:0.15;">

""", unsafe_allow_html=True)
st.markdown("""
<style>

.card{
    position: relative;
}

.card::before{
    content:"";
    position:absolute;
    top:0;
    left:0;
    width:100%;
    height:5px;

    background: linear-gradient(
        90deg,
        #00BFFF,
        #1E90FF,
        #00BFFF
    );

    border-radius:25px 25px 0 0;
}

.card:hover {
    transform: scale(1.01);

    box-shadow:
        0 0 25px #4DB2FF,
        0 0 50px rgba(77,178,255,0.8),
        0 0 80px rgba(77,178,255,0.5);
}

</style>
""", unsafe_allow_html=True)

st.markdown("""
<style>

[data-testid="stMetric"]{
    background: rgba(8,25,55,0.95);
    border: 2px solid #2490FF;
    border-radius: 18px;
    padding: 20px;
    text-align:center;

    box-shadow:
        0 0 15px rgba(36,144,255,.4);
}

/* BOTONES */
.stButton > button{
    background:#0F3A7A;
    color:white;
    border:2px solid #2490FF;
    border-radius:12px;

    box-shadow:
        0 0 10px rgba(36,144,255,.5),
        0 0 20px rgba(36,144,255,.3);
}

</style>
""", unsafe_allow_html=True)

col_logo, col_title = st.columns([1,5])

with col_logo:
    st.image("komatsu.png", width=240)

with col_title:
    st.title("ACREDITACIONES RADOMIRO TOMIC")
st.markdown("""
<style>

/* ✅ TEXTO GLOBAL */
html, body {
    color: white !important;
}

/* ✅ TODAS LAS ETIQUETAS DE STREAMLIT */
* {
    color: white !important;
}

/* ✅ SELECTBOX */
[data-baseweb="select"] * {
    color: white !important;
}

/* ✅ METRICAS */
[data-testid="stMetricValue"] {
    color: white !important;
    font-size: 34px !important;
    font-weight: bold !important;
}

[data-testid="stMetricLabel"] {
    color: #a8d4ff !important;
}

/* ✅ TEXTO DENTRO DE METRIC */
[data-testid="stMetric"] * {
    color: white !important;
}

/* ✅ TITULOS */
h1, h2, h3, h4 {
    color: white !important;
}

/* ✅ INFO (TUS CLASES) */
.info {
    color: #d7e8ff !important;
}

.nombre {
    color: white !important;
}

/* ✅ TEXTOS GENERALES */
p, span, label {
    color: #e6f0ff !important;
}

</style>
""", unsafe_allow_html=True)

ARCHIVO = "CONTROL_TOTAL.xlsx"
ARCHIVO_MATRIZ = "Certificados.xlsx"

BASE_PATH = r"C:\Users\u1305913\OneDrive - Komatsu Ltd\TUTOR DE CAPACITACIÓN\CERTIFICADOS LEGALES 2026"

# ==============================
# CARGAR MATRIZ
# ==============================
df_matriz = pd.read_excel(ARCHIVO_MATRIZ, sheet_name="MATRIZ")
df_matriz.columns = df_matriz.columns.str.strip().str.upper()
df_matriz["NOMBRE COMPLETO"] = df_matriz["NOMBRE COMPLETO"].str.strip().str.upper()

# ==============================
# CARGAR DATOS CONTROL
# ==============================
df = pd.read_excel(ARCHIVO)
df.columns = df.columns.str.upper()

# ==============================
# FUNCIONES
# ==============================
def estado(fecha):
    if pd.isna(fecha):
        return "—"

    fecha = pd.to_datetime(fecha)
    hoy = datetime.now()

    if fecha < hoy:
        return "🔴 VENCIDO"

    dias_restantes = (fecha - hoy).days

    if dias_restantes <= 45:
        return "🟡 POR VENCER"

    return "🟢 VIGENTE"

def dato(df, col):
    if col in df.columns and not df.empty:
        val = df[col].iloc[0]
        return val if pd.notna(val) else "Sin dato"
    return "Sin dato"
	
def obtener_foto(rut):

    carpeta = "fotos"

    if pd.isna(rut):
        return None

    rut = str(rut).strip()

    extensiones = [".jpg", ".jpeg", ".png"]

    for ext in extensiones:

        ruta = os.path.join(carpeta, rut + ext)

        if os.path.exists(ruta):
            return ruta

    return None

def mostrar_pdf(pdf_file):
    with open(pdf_file, "rb") as f:
        base64_pdf = base64.b64encode(f.read()).decode("utf-8")

    pdf_display = f"""
    <iframe
        src="data:application/pdf;base64,{base64_pdf}"
        width="100%"
        height="800"
        type="application/pdf">
    </iframe>
    """

    st.markdown(pdf_display, unsafe_allow_html=True)

import qrcode
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO

def generar_qr_personalizado(url, nombre):

    # ✅ crear QR
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=2,
    )

    qr.add_data(url)
    qr.make(fit=True)

    qr_img = qr.make_image(fill_color="#0F3A7A", back_color="white").convert("RGB")

    ancho, alto = qr_img.size

    # ✅ crear imagen más grande (para agregar nombre abajo)
    nueva_altura = alto + 80
    img_final = Image.new("RGB", (ancho, nueva_altura), "white")

    img_final.paste(qr_img, (0, 0))

    draw = ImageDraw.Draw(img_final)

    # ✅ fuentes
    try:
        font_centro = ImageFont.truetype("arial.ttf", 200)
        font_nombre = ImageFont.truetype("arial.ttf", 200)
    except:
        font_centro = ImageFont.load_default()
        font_nombre = ImageFont.load_default()

    # ✅ TEXTO CENTRO (KOMATSU RT)
    texto = "KOMATSU RT"

    bbox = draw.textbbox((0, 0), texto, font=font_centro)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]

    x = (ancho - tw) // 2
    y = (alto - th) // 2

    draw.rectangle(
        [x-10, y-10, x+tw+10, y+th+10],
        fill="white"
    )

    draw.text((x, y), texto, fill="#0F3A7A", font=font_centro)

    # ✅ NOMBRE ABAJO
    nombre = nombre.upper()

    bbox2 = draw.textbbox((0, 0), nombre, font=font_nombre)
    nw = bbox2[2] - bbox2[0]
    nh = bbox2[3] - bbox2[1]

    x_nombre = (ancho - nw) // 2
    y_nombre = alto + 20

    draw.text((x_nombre, y_nombre), nombre, fill="black", font=font_nombre)

    # ✅ guardar
    buffer = BytesIO()
    img_final.save(buffer, format="PNG")
    buffer.seek(0)

    return buffer
	
# ==============================
# SELECTOR
# ==============================

query_params = st.query_params
rut_url = query_params.get("rut", None)

lista_trabajadores = sorted(df_matriz["NOMBRE COMPLETO"].unique())

trabajador_qr = None

if rut_url:
    fila = df_matriz[df_matriz["RUT"].astype(str).str.strip() == str(rut_url)]
    if not fila.empty:
        trabajador_qr = fila["NOMBRE COMPLETO"].iloc[0]

if trabajador_qr:
    trabajador = st.selectbox(
        "Seleccionar trabajador",
        lista_trabajadores,
        index=lista_trabajadores.index(trabajador_qr)
    )
else:
    trabajador = st.selectbox("Seleccionar trabajador", lista_trabajadores)

trabajador_clean = trabajador.strip().upper()

df_persona = df_matriz[df_matriz["NOMBRE COMPLETO"] == trabajador_clean]
df_trab = df[df["TRABAJADOR"].str.upper() == trabajador_clean]

rut_trabajador = str(dato(df_persona, "RUT")).strip()

# ==============================
# 📊 RESUMEN
# ==============================
with st.container(border=True):

	df_cert = df_trab[df_trab["TIPO"] == "CERTIFICACION"]
	df_op = df_trab[df_trab["TIPO"].str.upper().str.contains("OPERACIÓN_EQUIPOS")]
	df_form = df_trab[df_trab["TIPO"] == "FORMACION"]

	c1, c2, c3, c4 = st.columns(4)

	c1.metric("📄 Certificados", len(df_cert))
	c2.metric("🚜 Operación", len(df_op))
	c3.metric("🎓 Formación", len(df_form))
	c4.metric("📊 Total", len(df_trab))

# ==============================
# PERFIL
# ==============================
with st.container(border=True):

	st.subheader("👤 Perfil del Trabajador")


with st.container(border=True):

    col1, col2, col3 = st.columns([1.5, 4, 1.5])

    # 🔹 FOTO
    with col1:
        foto = obtener_foto(rut_trabajador)
        if foto:
            st.image(foto, width=200)

    # 🔹 DATOS
    with col2:
        st.write(f"## {trabajador}")
        st.write(f"📇 Rut: {dato(df_persona,'RUT')}")
        st.write(f"💼 Cargo: {dato(df_persona,'CARGO')}")
        st.write(f"🏢 Empresa: {dato(df_persona,'EMPRESA')}")
        st.write(f"🪪 Licencia: {dato(df_persona,'LICENCIA')}")
        st.write(f"📧 Correo: {dato(df_persona,'E-MAIL')}")

    # 🔹 QR A LA DERECHA
    with col3:
        BASE_URL = "https://acreditaciones-komatsu-i3nrnuq22ipz2qgisfz7rb.streamlit.app/"
        url_trabajador = f"{BASE_URL}?rut={rut_trabajador}"

        qr_img = generar_qr_personalizado(
    url_trabajador,
    trabajador   # 👈 nombre del trabajador
)

        st.image(qr_img, width=180)
        
st.download_button(
    label="⬇️ Descargar QR",
    data=qr_img,
    file_name=f"QR_{rut_trabajador}.png",
    mime="image/png"
)
# ==============================
# ACREDITACIONES 
# ==============================

# 🔹 cargar hoja
xls = pd.ExcelFile(ARCHIVO_MATRIZ)

# 🔍 detectar hoja automáticamente
sheet_acred = None
for hoja in xls.sheet_names:
    if "ACREDIT" in hoja.upper():
        sheet_acred = hoja
        break

# ✅ validar
if sheet_acred is None:
    st.error("❌ No se encontró hoja de ACREDITACIONES")
else:
    df_acred = pd.read_excel(ARCHIVO_MATRIZ, sheet_name=sheet_acred)
df_acred.columns = df_acred.columns.str.strip().str.upper()

# 🔹 obtener rut y nombre
rut_trabajador = str(dato(df_persona, "RUT")).strip()
nombre_trabajador = trabajador_clean

# 🔹 buscar por RUT (columna A)
df_acred_trab = df_acred[
    df_acred.iloc[:, 0].astype(str).str.strip() == rut_trabajador
]

# 🔹 si no encuentra, buscar por nombre (columna B)
if df_acred_trab.empty:
    df_acred_trab = df_acred[
        df_acred.iloc[:, 1].astype(str).str.upper().str.strip() == nombre_trabajador
    ]

# 🔹 tomar columnas desde D en adelante
df_acred_filtrado = df_acred_trab.iloc[:, 3:]

# 🔹 transformar columnas a filas
acreditaciones = []

if not df_acred_filtrado.empty:
    for col in df_acred_filtrado.columns:

        valor = df_acred_filtrado[col].iloc[0]

        if pd.notna(valor):
            acreditaciones.append({
                "CURSO": col,
                "VENCIMIENTO": valor
            })

df_acreditaciones = pd.DataFrame(acreditaciones)

# ==============================
# MOSTRAR EN APP
# ==============================
with st.container(border=True):

    st.subheader("📋 Acreditaciones")

with st.container(border=True):

    col1, col2, col3 = st.columns([2,1,2])
    col1.markdown("**ACREDITACIONES**")
    col2.markdown("**VENCIMIENTO**")
    col3.markdown("**ESTADO**")

with st.container(border=True):

    for i, row in df_acreditaciones.iterrows():

        col1, col2, col3 = st.columns([2,1,2])

        fecha = pd.to_datetime(row["VENCIMIENTO"])

        col1.write(row["CURSO"])
        col2.write(fecha.strftime("%d-%m-%Y"))
        col3.write(estado(fecha))

# ==============================
# ✅ CERTIFICADOS
# ==============================
with st.container(border=True):

    st.subheader("📄 Certificados Legales")

archivo_cert ="Certificados legales.xlsx"

xls = pd.ExcelFile(archivo_cert)

certificados = []

for hoja in xls.sheet_names:

    try:
        df_hoja = pd.read_excel(archivo_cert, sheet_name=hoja)

        # validar que tenga al menos 4 columnas
        if df_hoja.shape[1] < 4:
            continue

        # buscar por RUT en columna A (index 0)
        df_trab_hoja = df_hoja[
            df_hoja.iloc[:, 0].astype(str).str.strip() == str(rut_trabajador)
        ]

        if not df_trab_hoja.empty:

            fecha = df_trab_hoja.iloc[0, 3]  # columna D

            if pd.notna(fecha):
                certificados.append({
                    "CURSO": hoja,
                    "VENCIMIENTO": fecha
                })

    except:
        continue

df_cert_excel = pd.DataFrame(certificados)
st.subheader("📦 Descargar QR de todos los trabajadores")

if st.button("Generar QR masivo"):

    BASE_URL = "https://acreditaciones-komatsu-q32swdvybciqekug6s9y8c.streamlit.app"

    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w") as zip_file:

        for _, row in df_matriz.iterrows():

            rut = str(row["RUT"]).strip().replace(".", "").upper()
            nombre = row["NOMBRE COMPLETO"]

            url = f"{BASE_URL}?rut={rut}"

            qr = generar_qr_personalizado(url, nombre)

            nombre_archivo = f"{nombre}_{rut}.png".replace(" ", "_")

            zip_file.writestr(nombre_archivo, qr.getvalue())

    zip_buffer.seek(0)

    st.download_button(
        label="⬇️ Descargar todos los QR",
        data=zip_buffer,
        file_name="QR_Trabajadores_Komatsu.zip",
        mime="application/zip"
    )

# ==============================
# ✅ MOSTRAR
# ==============================
with st.container(border=True):
    if df_cert_excel.empty:
        st.warning("⚠️ Sin certificados encontrados")
    else:
        col1, col2, col3 = st.columns([4,2,2])

        col1.markdown("**CURSO**")
        col2.markdown("**VENCIMIENTO**")
        col3.markdown("**ESTADO**")

with st.container(border=True):

    for _, row in df_cert_excel.iterrows():

        col1, col2, col3 = st.columns([4,2,2])

        fecha = pd.to_datetime(row["VENCIMIENTO"], errors="coerce")

        if pd.isna(fecha):
            continue

        col1.write(row["CURSO"])
        col2.write(fecha.strftime("%d-%m-%Y"))
        col3.write(estado(fecha))

# ==============================
# 2. OPERACION EQUIPO (TODOS)
# ==============================
with st.container(border=True):

    st.subheader("📄 Operación de equipos")

archivo_cert ="Operación equipos.xlsx"

xls = pd.ExcelFile(archivo_cert)

certificados = []

for hoja in xls.sheet_names:

    try:
        df_hoja = pd.read_excel(archivo_cert, sheet_name=hoja)

        # validar que tenga al menos 4 columnas
        if df_hoja.shape[1] < 4:
            continue

        # buscar por RUT en columna A (index 0)
        df_trab_hoja = df_hoja[
            df_hoja.iloc[:, 0].astype(str).str.strip() == str(rut_trabajador)
        ]

        if not df_trab_hoja.empty:

            fecha = df_trab_hoja.iloc[0, 3]  # columna D

            if pd.notna(fecha):
                certificados.append({
                    "CURSO": hoja,
                    "VENCIMIENTO": fecha
                })

    except:
        continue

df_cert_excel = pd.DataFrame(certificados)

# ==============================
# ✅ MOSTRAR
# ==============================
with st.container(border=True):
    if df_cert_excel.empty:
        st.warning("⚠️ Sin certificados encontrados")
    else:
        col1, col2, col3 = st.columns([4,2,2])

        col1.markdown("**CURSO**")
        col2.markdown("**VENCIMIENTO**")
        col3.markdown("**ESTADO**")

with st.container(border=True):

    for _, row in df_cert_excel.iterrows():

        col1, col2, col3 = st.columns([4,2,2])

        fecha = pd.to_datetime(row["VENCIMIENTO"], errors="coerce")

        if pd.isna(fecha):
            continue

        col1.write(row["CURSO"])
        col2.write(fecha.strftime("%d-%m-%Y"))
        col3.write(estado(fecha))

# ==============================
# 3. FORMACION
# ==============================
with st.container(border=True):

    st.subheader("🎓 Formación CFK")

with st.container(border=True):

    col1, col2, col3 = st.columns([5,2,2])
    col1.markdown("**CURSO**")
    col2.markdown("**ESTADO**")

with st.container(border=True):

    for i, row in df_trab[df_trab["TIPO"] == "FORMACION"].iterrows():

        col1, col2, col3 = st.columns([5, 2, 2])

        col1.write(row["CURSO"])
		
		st.subheader("📦 Descargar QR de todos los trabajadores")

if st.button("Generar QR masivo"):
    
    BASE_URL = "https://acreditaciones-komatsu-q32swdvybciqekug6s9y8c.streamlit.app/"

    # ✅ crear archivo ZIP temporal
    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w") as zip_file:

        for _, row in df_matriz.iterrows():

            rut = str(row["RUT"]).strip().replace(".", "").upper()
            nombre = row["NOMBRE COMPLETO"]

            # ✅ crear URL
            url = f"{BASE_URL}?rut={rut}"

            # ✅ generar QR personalizado
            qr = generar_qr_personalizado(url, nombre)

            # ✅ nombre archivo
            nombre_archivo = f"{nombre}.png".replace(" ", "_")

            # ✅ guardar en el ZIP
            zip_file.writestr(nombre_archivo, qr.getvalue())

    zip_buffer.seek(0)

    # ✅ descargar ZIP
    st.download_button(
        label="⬇️ Descargar todos los QR",
        data=zip_buffer,
        file_name="QR_Trabajadores_Komatsu.zip",
        mime="application/zip"
    )
        col2.write("✅ FORMACIÓN COMPLETADA")



        
