from __future__ import annotations

from pathlib import Path
import pandas as pd
import streamlit as st
from PIL import Image

from parser import parse_liquidacion_pdf
from exporters import build_excel_ventas, build_excel_cpns, build_excel_gastos

APP_TITLE = "IA liquidaciones agropecuarias"

ASSETS_DIR = Path(__file__).parent / "assets"
LOGO_PATH = ASSETS_DIR / "logo_aie.png"


def fmt_amount(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    try:
        v = float(x)
        return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return x


def fmt_aliq(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    try:
        v = float(x)
        return f"{v:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return x


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
    accept_multiple_files=True,
)

if pdf_files:
    parsed = []
    preview_rows = []

    with st.spinner("Procesando PDFs..."):
        for uf in pdf_files:
            data = uf.read()
            doc = parse_liquidacion_pdf(data, filename=uf.name)
            parsed.append(doc)

            preview_rows.append({
                "Archivo": uf.name,
                "Fecha": doc.fecha,
                "Localidad": doc.localidad,
                "COE": doc.coe,
                "Tipo": doc.tipo_cbte,
                "Acopio/Comprador": (doc.comprador.razon_social or "").strip(),
                "CUIT Comprador": doc.comprador.cuit,
                "Grano": doc.grano,
                "Campaña": doc.campaña or "",
                "Kg": doc.kilos,
                "Precio/Kg": doc.precio,
                "Neto": doc.neto,
                "Alic IVA": doc.alic_iva,
                "IVA": doc.iva,
                "Percep IVA": getattr(doc, "perc_iva", 0.0),
                "Ret IVA": doc.ret_iva,
                "Ret Gan": doc.ret_gan,
                "Total": doc.total,
            })

    st.subheader("Vista previa")
    df = pd.DataFrame(preview_rows)

    df_show = df.copy()
    for col in ["Kg", "Precio/Kg", "Neto", "IVA", "Percep IVA", "Ret IVA", "Ret Gan", "Total"]:
        df_show[col] = df_show[col].apply(fmt_amount)
    df_show["Alic IVA"] = df_show["Alic IVA"].apply(fmt_aliq)

    st.dataframe(df_show, use_container_width=True, hide_index=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        out = build_excel_ventas(parsed)
        st.download_button(
            "Descargar Ventas",
            data=out.getvalue(),
            file_name="ventas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col2:
        out = build_excel_cpns(parsed)
        st.download_button(
            "Descargar CPNs",
            data=out.getvalue(),
            file_name="cpns.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col3:
        out = build_excel_gastos(parsed)
        st.download_button(
            "Descargar Gastos",
            data=out.getvalue(),
            file_name="gastos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
