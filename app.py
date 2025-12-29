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
    """1.000.000,00 (AR)"""
    if x is None:
        return ""
    try:
        if isinstance(x, str) and not x.strip():
            return ""
        v = float(x)
        return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return x


def fmt_aliq(x):
    """10,500 (3 decimales AR)"""
    if x is None:
        return ""
    try:
        if isinstance(x, str) and not x.strip():
            return ""
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
    parsed_docs = []
    preview_rows = []

    with st.spinner("Procesando PDFs..."):
        for uf in pdf_files:
            try:
                data = uf.read()
                doc = parse_liquidacion_pdf(data, filename=uf.name)
                parsed_docs.append(doc)

                preview_rows.append(
                    {
                        "Archivo": uf.name,
                        "Fecha": getattr(doc, "fecha", ""),
                        "Localidad": getattr(doc, "localidad", ""),
                        "COE": getattr(doc, "coe", ""),
                        "Tipo": getattr(doc, "tipo_comprobante", ""),
                        "Acopio/Comprador": getattr(doc, "comprador_rs", ""),
                        "CUIT Comprador": getattr(doc, "comprador_cuit", ""),
                        "Grano": getattr(doc, "grano", ""),
                        "Campaña": getattr(doc, "campaña", ""),
                        "Kg": getattr(doc, "kilos", None),
                        "Precio/Kg": getattr(doc, "precio_kg", None),
                        "Subtotal": getattr(doc, "subtotal", None),
                        "Alic IVA": getattr(doc, "alicuota_iva", None),
                        "IVA": getattr(doc, "iva", None),
                        "Percep IVA": getattr(doc, "perc_iva", None),
                        "Ret IVA": getattr(doc, "ret_iva", None),
                        "Ret Gan": getattr(doc, "ret_gan", None),
                        "Total": getattr(doc, "total", None),
                    }
                )
            except Exception as e:
                st.error(f"Error procesando {uf.name}: {e}")

    if preview_rows:
        st.subheader("Vista previa")
        df = pd.DataFrame(preview_rows)

        df_show = df.copy()

        for col in ["Kg", "Precio/Kg", "Subtotal", "IVA", "Percep IVA", "Ret IVA", "Ret Gan", "Total"]:
            if col in df_show.columns:
                df_show[col] = df_show[col].apply(fmt_amount)

        if "Alic IVA" in df_show.columns:
            df_show["Alic IVA"] = df_show["Alic IVA"].apply(fmt_aliq)

        st.dataframe(df_show, use_container_width=True, hide_index=True)

        b1, b2, b3 = st.columns(3)

        with b1:
            out = build_excel_ventas(parsed_docs)
            st.download_button(
                "Descargar Ventas",
                data=out.getvalue(),
                file_name="ventas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with b2:
            out = build_excel_cpns(parsed_docs)
            st.download_button(
                "Descargar CPNs",
                data=out.getvalue(),
                file_name="cpns.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with b3:
            out = build_excel_gastos(parsed_docs)
            st.download_button(
                "Descargar Gastos",
                data=out.getvalue(),
                file_name="gastos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
