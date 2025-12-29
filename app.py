import streamlit as st
import pandas as pd
from pathlib import Path
from PIL import Image

from parser import parse_liquidacion_pdf
from exporters import ventas_xlsx_bytes, cpns_xlsx_bytes, gastos_xlsx_bytes

APP_TITLE = "IA liquidaciones agropecuarias"

ASSETS_DIR = Path(__file__).parent / "assets"
LOGO_PATH = ASSETS_DIR / "logo_aie.png"

st.set_page_config(
    page_title=APP_TITLE,
    page_icon=Image.open(LOGO_PATH) if LOGO_PATH.exists() else None,
    layout="wide",
)

c1, c2 = st.columns([1, 6])
with c1:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), use_container_width=True)
with c2:
    st.title(APP_TITLE)

pdf_files = st.file_uploader(
    "Subí una o más liquidaciones (PDF)",
    type=["pdf"],
    accept_multiple_files=True
)

def fmt_amount(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    try:
        v = float(x)
    except Exception:
        return str(x)
    # 1.000,00
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_aliq(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    try:
        v = float(x)
    except Exception:
        return str(x)
    # 10,500
    return f"{v:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".")

if pdf_files:
    liqs = []
    preview_rows = []

    with st.spinner("Procesando PDFs..."):
        for uf in pdf_files:
            data = uf.read()
            l = parse_liquidacion_pdf(data, filename=uf.name)
            liqs.append(l)

            preview_rows.append({
                "Archivo": uf.name,
                "Fecha": l.fecha,
                "COE": l.coe,
                "CUIT Comprador": l.comprador.cuit,
                "Grano": l.grano,
                "Campaña": l.campaña,
                "Kg": l.kilos,
                "Precio/Kg": l.precio,
                "Neto": l.neto,
                "Alic IVA": l.alic_iva,
                "IVA": l.iva,
                "Ret IVA": l.ret_iva,   # <- ya viene del cuadro RETENCIONES (monto)
                "Total": l.total,
            })

    st.subheader("Vista previa")
    df = pd.DataFrame(preview_rows)

    df_show = df.copy()
    for col in ["Kg", "Precio/Kg", "Neto", "IVA", "Ret IVA", "Total"]:
        if col in df_show.columns:
            df_show[col] = df_show[col].apply(fmt_amount)
    if "Alic IVA" in df_show.columns:
        df_show["Alic IVA"] = df_show["Alic IVA"].apply(fmt_aliq)

    st.dataframe(df_show, use_container_width=True, hide_index=True)

    col1, col2, col3 = st.columns(3)

    with col1:
        st.download_button(
            "Descargar Ventas",
            data=ventas_xlsx_bytes(liqs),
            file_name="ventas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col2:
        st.download_button(
            "Descargar CPNs",
            data=cpns_xlsx_bytes(liqs),
            file_name="cpns.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col3:
        st.download_button(
            "Descargar Gastos",
            data=gastos_xlsx_bytes(liqs),
            file_name="gastos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
