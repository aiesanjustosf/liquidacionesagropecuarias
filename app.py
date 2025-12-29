# -*- coding: utf-8 -*-
from __future__ import annotations

import pandas as pd
import streamlit as st

from parser import parse_liquidacion_pdf
from exporters import (
    build_ventas_rows,
    build_cpns_rows,
    build_gastos_rows,
    df_to_xlsx_bytes,
)

st.set_page_config(page_title="IA Liquidaciones Agropecuarias", layout="wide")

st.title("IA Liquidaciones Agropecuarias")
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

    # Vista previa (la grilla estaba bien: mantenemos formato visible)
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
        )

    with c2:
        st.download_button(
            "Descargar CPNs",
            data=df_to_xlsx_bytes(cpns_df, "CPNs"),
            file_name="cpns.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with c3:
        st.download_button(
            "Descargar Gastos",
            data=df_to_xlsx_bytes(gastos_df, "Gastos"),
            file_name="gastos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
