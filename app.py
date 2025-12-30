# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path
from PIL import Image

import pandas as pd
import streamlit as st

from parser import parse_liquidacion_pdf
from exporters import (
    build_ventas_rows,
    build_cpns_rows,
    build_gastos_rows,
    df_to_xlsx_bytes,
)

APP_TITLE = "IA Liquidaciones Agropecuarias"

HERE = Path(__file__).parent

def first_existing(*paths: Path):
    for p in paths:
        if p and p.exists():
            return p
    return None

LOGO_PATH = first_existing(
    HERE / "assets" / "logo_aie.png",
    HERE / "logo_aie.png",
)

FAVICON_PATH = first_existing(
    HERE / "assets" / "favicon-aie.ico",
    HERE / "favicon-aie.ico",
    HERE / "assets" / "favicon-aie.png",
    HERE / "favicon-aie.png",
)

def load_favicon(path: Path) -> Image.Image:
    img = Image.open(path).convert("RGBA")
    # Normaliza tamaño para que el navegador lo tome como favicon
    img = img.resize((32, 32))
    return img

# IMPORTANTE: esto debe ser el primer st.* del archivo
st.set_page_config(
    page_title=APP_TITLE,
    page_icon=load_favicon(FAVICON_PATH) if FAVICON_PATH else None,
    layout="wide",
)

# Header como tu referencia, pero con logo MÁS GRANDE y menos “gap”
h1, h2 = st.columns([0.55, 12])  # achica la columna del logo para que no quede lejos del título
with h1:
    if LOGO_PATH:
        st.image(str(LOGO_PATH), width=120)  # subí/bajá 110–150 si querés
with h2:
    st.markdown(f"<h1 style='margin:0; padding-top:6px;'>{APP_TITLE}</h1>", unsafe_allow_html=True)

# --- lo demás queda EXACTAMENTE igual ---
files = st.file_uploader("Subí una o más liquidaciones (PDF)", type=["pdf"], accept_multiple_files=True)


def _fmt_monto(x):
    try:
        return f"{float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return x

def _fmt_alic(x):
    try:
        return f"{float(x):,.3f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return x

if files:
    liqs = []
    for f in files:
        liqs.append(parse_liquidacion_pdf(f.getvalue(), f.name))

    preview = pd.DataFrame([{
        "Archivo": l.filename,
        "CUIT Comprador": l.comprador.cuit,
        "Grano": l.grano,
        "Campaña": l.campaña or "",
        "Kg": l.kilos,
        "Precio/Kg": l.precio,
        "Neto": l.neto,
        "Alic IVA": l.alic_iva,
        "IVA": l.iva,
        "Ret IVA": l.ret_iva,
        "Ret Gan": l.ret_gan,  # siempre 0 por parser
        "Total": l.total,
    } for l in liqs])

    st.subheader("Vista previa")
    st.dataframe(
        preview.style.format({
            "Kg": _fmt_monto,
            "Precio/Kg": _fmt_monto,
            "Neto": _fmt_monto,
            "Alic IVA": _fmt_alic,
            "IVA": _fmt_monto,
            "Ret IVA": _fmt_monto,
            "Ret Gan": _fmt_monto,
            "Total": _fmt_monto,
        }),
        use_container_width=True
    )

    ventas_df = build_ventas_rows(liqs)
    cpns_df = build_cpns_rows(liqs)
    gastos_df = build_gastos_rows(liqs)

    c1, c2, c3 = st.columns(3)

    with c1:
        st.download_button(
            "Descargar Ventas",
            data=df_to_xlsx_bytes(ventas_df, "Ventas"),
            file_name="ventas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with c2:
        st.download_button(
            "Descargar CPNs",
            data=df_to_xlsx_bytes(cpns_df, "CPNs"),
            file_name="cpns.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with c3:
        st.download_button(
            "Descargar Gastos",
            data=df_to_xlsx_bytes(gastos_df, "Gastos"),
            file_name="gastos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

# ----------------- Footer fijo -----------------
st.markdown(
    """
    <style>
      .aie-footer{
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background: #ffffff;
        color: #6b7280;
        text-align: center;
        padding: 6px 0;
        font-size: 12px;
        border-top: 1px solid #e5e7eb;
        z-index: 999;
      }
      .block-container{ padding-bottom: 48px; }
    </style>
    <div class="aie-footer">Herramienta para uso interno AIE San Justo | Developer Alfonso Alderete</div>
    """,
    unsafe_allow_html=True
)
