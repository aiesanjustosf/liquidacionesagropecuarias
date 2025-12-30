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

# Nuevo: control de duplicados por COE
skip_duplicates = st.checkbox("Omitir duplicados (mismo COE)", value=True)

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

def _fmt_int(x):
    try:
        return f"{int(x):,}".replace(",", ".")
    except Exception:
        return x

if files:
    liqs = []
    seen_by_coe = {}   # coe -> filename (primera aparición)
    dup_rows = []      # para mostrar detalle

    for f in files:
        liq = parse_liquidacion_pdf(f.getvalue(), f.name)
        coe = (liq.coe or "").strip()

        if coe:
            if coe in seen_by_coe:
                dup_rows.append({
                    "COE": coe,
                    "Archivo duplicado": f.name,
                    "Ya cargado": seen_by_coe[coe],
                })
                if skip_duplicates:
                    continue
            else:
                seen_by_coe[coe] = f.name

        liqs.append(liq)

    if dup_rows:
        st.warning("Se detectaron liquidaciones duplicadas por COE.")
        st.dataframe(pd.DataFrame(dup_rows), use_container_width=True, hide_index=True)

    if not liqs:
        st.error("No quedaron liquidaciones para procesar (todas eran duplicadas por COE).")
        st.stop()

    # Base numérica (para resumen y grilla)
    base = pd.DataFrame([{
        "Archivo": l.filename,
        "COE": l.coe,
        "CUIT Comprador": l.comprador.cuit,
        "Grano": l.grano or "SIN DATO",
        "Campaña": l.campaña or "",
        "Kg": float(l.kilos or 0.0),
        "Precio/Kg": float(l.precio or 0.0),
        "Neto": float(l.neto or 0.0),
        "Alic IVA": float(l.alic_iva or 0.0),
        "IVA": float(l.iva or 0.0),
        "Ret IVA": float(l.ret_iva or 0.0),
        "Ret Gan": float(l.ret_gan or 0.0),
        "Total": float(l.total or 0.0),
    } for l in liqs])

    # Vista previa (mantenemos formato visible)
    st.subheader("Vista previa")
    st.dataframe(
        base.drop(columns=["COE"]).style.format({
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

    # Nuevo: Resumen por tipo de grano (segunda grilla)
    st.subheader("Resumen por grano")
    resumen = (
        base.groupby("Grano", as_index=False)
            .agg(
                **{
                    "Cant. Liquidaciones": ("Archivo", "count"),
                    "Kg": ("Kg", "sum"),
                    "Neto": ("Neto", "sum"),
                    "IVA": ("IVA", "sum"),
                    "Ret IVA": ("Ret IVA", "sum"),
                    "Total": ("Total", "sum"),
                }
            )
            .sort_values("Grano")
    )

    # Total general (fila final)
    total_row = {
        "Grano": "TOTAL",
        "Cant. Liquidaciones": int(resumen["Cant. Liquidaciones"].sum()),
        "Kg": float(resumen["Kg"].sum()),
        "Neto": float(resumen["Neto"].sum()),
        "IVA": float(resumen["IVA"].sum()),
        "Ret IVA": float(resumen["Ret IVA"].sum()),
        "Total": float(resumen["Total"].sum()),
    }
    resumen = pd.concat([resumen, pd.DataFrame([total_row])], ignore_index=True)

    st.dataframe(
        resumen.style.format({
            "Cant. Liquidaciones": _fmt_int,
            "Kg": _fmt_monto,
            "Neto": _fmt_monto,
            "IVA": _fmt_monto,
            "Ret IVA": _fmt_monto,
            "Total": _fmt_monto,
        }),
        use_container_width=True,
        hide_index=True
    )

    # Exportaciones (sin cambios)
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
