# app.py
import streamlit as st
import pandas as pd
from pathlib import Path
from PIL import Image

from parser import parse_liquidacion_pdf
from exporters import (
    build_ventas_rows,
    build_cpns_rows,
    build_gastos_rows,
    df_to_xlsx_bytes,
)

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
        return x
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_aliq(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    try:
        v = float(x)
    except Exception:
        return x
    return f"{v:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".")

if pdf_files:
    docs = []
    preview_rows = []

    with st.spinner("Procesando PDFs..."):
        for uf in pdf_files:
            data = uf.read()
            doc = parse_liquidacion_pdf(data, filename=uf.name)
            docs.append(doc)

            preview_rows.append({
                "Archivo": uf.name,
                "Fecha": doc.fecha,
                "Localidad": doc.localidad,
                "COE": doc.coe,
                "CUIT Comprador": doc.acopio.cuit,
                "Acopio/Comprador": doc.acopio.razon_social,
                "Grano": doc.grano,
                "Campaña": doc.campaña,
                "Kg": doc.kilos,
                "Precio/Kg": doc.precio,
                "Neto": doc.neto,
                "Alic IVA": doc.alic_iva,
                "IVA": doc.iva,
                "Percep IVA": doc.percep_iva,
                "Ret IVA": doc.ret_iva,
                "Ret Gan": doc.ret_gan,
                "Total": doc.total,
            })

    st.subheader("Vista previa")
    df = pd.DataFrame(preview_rows)

    df_show = df.copy()
    for col in ["Kg", "Precio/Kg", "Neto", "IVA", "Percep IVA", "Ret IVA", "Ret Gan", "Total"]:
        if col in df_show.columns:
            df_show[col] = df_show[col].apply(fmt_amount)
    if "Alic IVA" in df_show.columns:
        df_show["Alic IVA"] = df_show["Alic IVA"].apply(fmt_aliq)

    st.dataframe(df_show, use_container_width=True, hide_index=True)

    col1, col2, col3 = st.columns(3)

    with col1:
        dfv = build_ventas_rows(docs)
        out = df_to_xlsx_bytes(dfv, sheet_name="Ventas")
        st.download_button(
            "Descargar Ventas",
            data=out,
            file_name="ventas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col2:
        dfc = build_cpns_rows(docs)
        out = df_to_xlsx_bytes(dfc, sheet_name="CPNs")
        st.download_button(
            "Descargar CPNs",
            data=out,
            file_name="cpns.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col3:
        dfg = build_gastos_rows(docs)
        out = df_to_xlsx_bytes(dfg, sheet_name="Gastos")
        st.download_button(
            "Descargar Gastos",
            data=out,
            file_name="gastos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
