import streamlit as st
import pandas as pd
from datetime import datetime
import os
import webbrowser
import base64

# ==============================
# CONFIG
# ==============================
st.set_page_config(layout="wide")

st.markdown("""
<h1 style='
color:white;
font-size:52px;
font-weight:800;
margin-bottom:20px;
'>
📊 ACREDITACIONES KOMATSU RT
</h1>
""", unsafe_allow_html=True)

archivo = "Certificados.xlsx"

# ==============================
# LOGO BASE64
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

/* ✅ LOGO SIN BORDE */
.logo-corner{{
    position: fixed;
    top: 100px;
    right: 80px;
    width: 220px;

    opacity: 0.18;
    filter: brightness(1.5);

    z-index: 0;
    pointer-events: none;

    border: none !important;
    box-shadow: none !important;
}}

/* TARJETAS */
.card-pro{{
    background:
    linear-gradient(135deg,
        rgba(8,25,55,0.95),
        rgba(4,15,35,0.95)
    );
    border:2px solid #2490ff;
    border-radius:24px;
    padding:28px;
    margin-bottom:22px;

    box-shadow:
    0 0 25px rgba(36,144,255,.45),
    inset 0 0 40px rgba(36,144,255,.08);
}}

.card-title{{
    color:white;
    font-size:22px;
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

.linea{{
    border-bottom:1px solid rgba(255,255,255,.15);
    margin:12px 0;
}}

/* BOTONES */
.stButton button{{
    width:100%;
    border-radius:12px;
    border:1px solid #2490ff;
    background:#0f2248;
    color:white;
    font-weight:600;
}}

.stButton button:hover{{
    background:#16346d;
    border:1px solid #58b1ff;
}}

/* BADGES */
.badge{{
    display:inline-block;
    min-width:120px;
    text-align:center;
    padding:8px;
    border-radius:10px;
    font-weight:700;
    color:white;
}}

.vigente{{background:#178f4b;}}
.vencer{{background:#d39a12;}}
.vencido{{background:#c93939;}}
.info_badge{{background:#246bcb;}}

/* FOTO SOLO EN PERFIL */
.card-pro img{{
    border-radius:18px;
    border:2px solid #2490ff;
    box-shadow:0 0 20px rgba(36,144,255,.45);
}}

</style>

<!-- ✅ LOGO -->
<img src="data:image/png;base64,{logo_base64}" class="logo-corner">

""", unsafe_allow_html=True)

# ==============================
# DATOS
# ==============================
df_matriz = pd.read_excel(archivo, sheet_name="MATRIZ")
df_matriz.columns = df_matriz.columns.str.strip().str.upper()

hojas = pd.read_excel(archivo, sheet_name=None)

lista = []
for hoja, df_tmp in hojas.items():

    if hoja.upper() == "MATRIZ":
        continue

    df_tmp.columns = df_tmp.columns.str.strip().str.upper()

    for _, fila in df_tmp.iterrows():
        if pd.isna(fila.get("RUT")):
            continue

        lista.append({
            "RUT": str(fila["RUT"]).strip(),
            "Curso": hoja,
            "Vencimiento": fila.get("VENCIMIENTO")
        })

df_cursos = pd.DataFrame(lista)

data = df_cursos.merge(df_matriz, on="RUT", how="left")
data["Vencimiento"] = pd.to_datetime(data["Vencimiento"], errors="coerce")

def semaforo(f):
    if pd.isna(f):
        return "SIN INFO"
    dias = (f - datetime.now()).days
    if f < datetime.now():
        return "VENCIDO"
    elif dias <= 30:
        return "POR VENCER"
    return "VIGENTE"

data["EstadoCalc"] = data["Vencimiento"].apply(semaforo)

trabajador = st.selectbox(
    "Seleccionar trabajador",
    sorted(data["NOMBRE COMPLETO"].dropna().unique())
)

df = data[data["NOMBRE COMPLETO"] == trabajador]

# ==============================
# RUTA
# ==============================
BASE_PATH = r"C:\\Users\\u1305913\\OneDrive - Komatsu Ltd\\TUTOR DE CAPACITACIÓN\\CERTIFICADOS LEGALES 2026"

def obtener_foto(nombre):
    ruta = os.path.join(BASE_PATH, nombre)
    if os.path.exists(ruta):
        for f in os.listdir(ruta):
            if f.lower().endswith(".jpg"):
                return os.path.join(ruta, f)
    return None

def buscar_pdf(nombre, curso):
    ruta = os.path.join(BASE_PATH, nombre)
    if not os.path.exists(ruta):
        return None

    for carpeta in os.listdir(ruta):
        if curso.lower().replace(" ","") in carpeta.lower().replace(" ",""):
            sub = os.path.join(ruta, carpeta)
            if os.path.isdir(sub):
                for f in os.listdir(sub):
                    if f.lower().endswith(".pdf"):
                        return os.path.join(sub, f)
    return None

# ==============================
# PERFIL
# ==============================
st.markdown('<div class="card-pro">', unsafe_allow_html=True)
st.markdown('<div class="card-title">👤 1. Perfil del Trabajador</div>', unsafe_allow_html=True)

col1, col2 = st.columns([1.2,6])

foto = obtener_foto(trabajador)

with col1:
    if foto:
        st.image(foto, width=220)

with col2:
    st.markdown(f'<div class="nombre">{trabajador}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="info">📇 RUT: {df["RUT"].iloc[0]}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="info">💼 Cargo: {df["CARGO"].iloc[0]}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="info">🏢 Empresa: {df["EMPRESA"].iloc[0]}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="info">🪪 Licencia: {df["LICENCIA"].iloc[0]}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="info">📧 Correo: {df["E-MAIL"].iloc[0]}</div>', unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# ==============================
# CERTIFICADOS
# ==============================
st.markdown('<div class="card-pro">', unsafe_allow_html=True)
st.markdown('<div class="card-title">📄 2. Estado de Documentos y Certificados</div>', unsafe_allow_html=True)

h1,h2,h3,h4 = st.columns([4,2,2,2])
h1.write("Documento")
h2.write("Fecha")
h3.write("Estado")
h4.write("Acción")

st.markdown('<div class="linea"></div>', unsafe_allow_html=True)

for i,row in df.iterrows():

    pdf = buscar_pdf(trabajador,row["Curso"])

    c1,c2,c3,c4 = st.columns([4,2,2,2])

    c1.write(row["Curso"])

    fecha = row["Vencimiento"].strftime("%d-%b-%Y") if pd.notna(row["Vencimiento"]) else "Sin info"
    c2.write(fecha)

    if row["EstadoCalc"]=="VIGENTE":
        c3.markdown('<span class="badge vigente">VIGENTE</span>', unsafe_allow_html=True)
    elif row["EstadoCalc"]=="POR VENCER":
        c3.markdown('<span class="badge vencer">POR VENCER</span>', unsafe_allow_html=True)
    elif row["EstadoCalc"]=="VENCIDO":
        c3.markdown('<span class="badge vencido">VENCIDO</span>', unsafe_allow_html=True)
    else:
        c3.markdown('<span class="badge info_badge">SIN INFO</span>', unsafe_allow_html=True)

    if pdf:
        if c4.button("📄 Ver Archivo", key=i):
            webbrowser.open_new(pdf)
    else:
        c4.write("—")

    st.markdown('<div class="linea"></div>', unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# ==============================
# RESUMEN
# ==============================
st.markdown('<div class="card-pro">', unsafe_allow_html=True)
st.markdown('<div class="card-title">📊 3. Resumen de Certificaciones</div>', unsafe_allow_html=True)

c1,c2,c3,c4 = st.columns(4)

c1.metric("Total", len(df))
c2.metric("Vigentes", len(df[df["EstadoCalc"]=="VIGENTE"]))
c3.metric("Por vencer", len(df[df["EstadoCalc"]=="POR VENCER"]))
c4.metric("Vencidos", len(df[df["EstadoCalc"]=="VENCIDO"]))

st.markdown('</div>', unsafe_allow_html=True)




